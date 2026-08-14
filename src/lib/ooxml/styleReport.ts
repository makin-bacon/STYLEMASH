import type {
  FormattingSignature,
  ParsedDocx,
  RunRef,
  StyleEntity,
  StyleEntityVariant,
  StyleOrigin,
  UserStyleRecord,
} from '../../types/ooxml'
import { NS } from './constants'
import { wChildren } from './domUtils'
import { signatureToKey } from './signature'
import {
  buildResolutionContext,
  buildStylesMap,
  getDocDefaultsRPr,
  resolveRunFormatting,
} from './styleResolution'
import { buildThemeColorMap } from './themeColor'

const SAMPLE_TEXT_MAX_LENGTH = 80

function getRunText(runEl: Element): string {
  let text = ''
  for (const t of wChildren(runEl, 't')) text += t.textContent ?? ''
  return text
}

function truncate(text: string, maxLength: number): string {
  return text.length > maxLength ? `${text.slice(0, maxLength)}…` : text
}

function originToKey(origin: StyleOrigin): string {
  return origin.kind === 'direct' ? 'direct' : `${origin.kind}:${origin.styleId}`
}

interface EntityDraft {
  signature: FormattingSignature
  sampleText: string
  variants: Map<string, StyleEntityVariant>
}

/** Walks every run in word/document.xml's body, resolves each one's
 * effective visual formatting, and groups them by that signature into the
 * Style Report - then further splits each group into variants by origin
 * (direct formatting vs. a specific named style), so text that reached the
 * same look via different paths stays independently selectable. Headers,
 * footers, and footnotes/endnotes are out of scope for v1 (they live in
 * separate zip parts we don't parse); runs with no visible text (e.g.
 * drawing-only or page-break-only runs) are skipped.
 *
 * Recomputed wholesale after every merge/edit rather than patched
 * incrementally - simpler, and cheap even for large documents. */
export function buildStyleReport(parsedDocx: ParsedDocx): StyleEntity[] {
  const stylesMap = buildStylesMap(parsedDocx.stylesXml)
  const docDefaultsRPr = getDocDefaultsRPr(parsedDocx.stylesXml)
  const themeColors = buildThemeColorMap(parsedDocx.themeXml)
  const ctx = buildResolutionContext(stylesMap, docDefaultsRPr, themeColors)

  const drafts = new Map<string, EntityDraft>()

  const paragraphs = parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'p')
  for (let i = 0; i < paragraphs.length; i++) {
    const paragraphEl = paragraphs[i]
    // Descendant (not direct-child) lookup so runs wrapped in w:hyperlink,
    // w:ins/w:del (tracked changes), or w:sdt content are picked up too,
    // without special-casing each wrapper element.
    const runs = paragraphEl.getElementsByTagNameNS(NS.w, 'r')
    for (let j = 0; j < runs.length; j++) {
      const runEl = runs[j]
      const text = getRunText(runEl)
      if (text.length === 0) continue

      const { signature, origin } = resolveRunFormatting(runEl, paragraphEl, ctx)
      const sigKey = signatureToKey(signature)

      let draft = drafts.get(sigKey)
      if (!draft) {
        draft = { signature, sampleText: '', variants: new Map() }
        drafts.set(sigKey, draft)
      }
      if (!draft.sampleText) draft.sampleText = truncate(text, SAMPLE_TEXT_MAX_LENGTH)

      const originKey = originToKey(origin)
      let variant = draft.variants.get(originKey)
      if (!variant) {
        variant = { id: `${sigKey}::${originKey}`, origin, occurrenceCount: 0, sampleText: '', runRefs: [] }
        draft.variants.set(originKey, variant)
      }
      variant.occurrenceCount += 1
      if (!variant.sampleText) variant.sampleText = truncate(text, SAMPLE_TEXT_MAX_LENGTH)
      variant.runRefs.push({ runElement: runEl, paragraphElement: paragraphEl, origin })
    }
  }

  const entities: StyleEntity[] = Array.from(drafts.entries()).map(([sigKey, draft]) => {
    const variants = Array.from(draft.variants.values()).sort(
      (a, b) => b.occurrenceCount - a.occurrenceCount,
    )
    const occurrenceCount = variants.reduce((sum, v) => sum + v.occurrenceCount, 0)
    return { id: sigKey, signature: draft.signature, occurrenceCount, sampleText: draft.sampleText, variants }
  })

  // Most-common formatting first - the most useful order for cleanup triage.
  return entities.sort((a, b) => b.occurrenceCount - a.occurrenceCount)
}

/** Finds a single variant by its id across the whole report - used to look
 * up what the raw-XML editor or an "edit style" action should act on. */
export function findVariantById(styleReport: StyleEntity[], variantId: string): StyleEntityVariant | null {
  for (const entity of styleReport) {
    const variant = entity.variants.find((v) => v.id === variantId)
    if (variant) return variant
  }
  return null
}

/** Flattens every RunRef belonging to the given variant ids, in report
 * order - the shape mergeStyles()/applyXmlFragmentToRunRefs() operate on,
 * decoupling those mutators from the report's grouping shape entirely. */
export function collectRunRefsForVariantIds(styleReport: StyleEntity[], variantIds: Set<string>): RunRef[] {
  const refs: RunRef[] = []
  for (const entity of styleReport) {
    for (const variant of entity.variants) {
      if (variantIds.has(variant.id)) refs.push(...variant.runRefs)
    }
  }
  return refs
}

/** Counts how many runs currently reference `styleId`, by re-scanning the
 * latest Style Report rather than hand-maintaining a running total on the
 * UserStyleRecord - keeps the "User-Created Styles" panel's occurrence
 * counts always in sync with the real document state. */
export function countOccurrencesForStyleId(styleReport: StyleEntity[], styleId: string): number {
  let total = 0
  for (const entity of styleReport) {
    for (const variant of entity.variants) {
      if (variant.origin.kind === 'named-character' && variant.origin.styleId === styleId) {
        total += variant.occurrenceCount
      }
    }
  }
  return total
}

/** How far through "merge everything into a User-Created style" the
 * document currently is - drives the Style Report's progress bar. Counted
 * at the variant level (the same granularity as selection/merging itself):
 * a variant counts as "merged" once its look is controlled by a named
 * character style that's tracked as a UserStyleRecord; every other variant
 * (direct formatting, or still under one of the document's own original
 * named styles) counts as remaining. */
export function computeMergeProgress(
  styleReport: StyleEntity[],
  userStyles: UserStyleRecord[],
): { total: number; merged: number; remaining: number } {
  const userStyleIds = new Set(userStyles.map((r) => r.styleId))
  let total = 0
  let merged = 0
  for (const entity of styleReport) {
    for (const variant of entity.variants) {
      total += 1
      if (variant.origin.kind === 'named-character' && userStyleIds.has(variant.origin.styleId)) {
        merged += 1
      }
    }
  }
  return { total, merged, remaining: total - merged }
}
