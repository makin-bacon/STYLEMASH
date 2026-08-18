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
import { buildParagraphMarkers } from './numbering'
import { signatureToKey } from './signature'
import {
  buildResolutionContext,
  buildStylesMap,
  getDocDefaultsRPr,
  resolveRunFormatting,
} from './styleResolution'
import { buildThemeColorMap } from './themeColor'

const SAMPLE_TEXT_MAX_LENGTH = 80

/** Concatenates every `<w:t>` child's text content within a run - the same
 * "visible text" extraction used to group/sample the Style Report, exported
 * so consumers like DocumentPreviewPanel can render the same text runs. */
export function getRunText(runEl: Element): string {
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
 * (direct formatting vs. a specific named style) AND by list membership, so
 * text that reached the same look via different paths - or that looks
 * identical to a list item purely by coincidence (very common: list
 * paragraphs frequently carry no character formatting of their own) - stays
 * independently selectable, and each variant's sample/marker stays
 * representative of every occurrence it actually contains rather than being
 * silently outvoted by whichever text happened to appear first in the
 * document. Headers, footers, and footnotes/endnotes are out of scope for
 * v1 (they live in separate zip parts we don't parse); runs with no visible
 * text (e.g. drawing-only or page-break-only runs) are skipped.
 *
 * Recomputed wholesale after every merge/edit rather than patched
 * incrementally - simpler, and cheap even for large documents. */
export function buildStyleReport(parsedDocx: ParsedDocx): StyleEntity[] {
  const stylesMap = buildStylesMap(parsedDocx.stylesXml)
  const docDefaultsRPr = getDocDefaultsRPr(parsedDocx.stylesXml)
  const themeColors = buildThemeColorMap(parsedDocx.themeXml)
  const ctx = buildResolutionContext(stylesMap, docDefaultsRPr, themeColors)
  const paragraphMarkers = buildParagraphMarkers(parsedDocx)

  const drafts = new Map<string, EntityDraft>()

  const paragraphs = parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'p')
  for (let i = 0; i < paragraphs.length; i++) {
    const paragraphEl = paragraphs[i]
    const isListItem = paragraphMarkers.has(paragraphEl)
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
      const variantKey = isListItem ? `${originKey}::list` : originKey
      let variant = draft.variants.get(variantKey)
      if (!variant) {
        variant = { id: `${sigKey}::${variantKey}`, origin, occurrenceCount: 0, sampleText: '', runRefs: [] }
        draft.variants.set(variantKey, variant)
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

/** True for the two origin kinds that mean "this run/paragraph's look comes
 * from a named style StyleMash could have created" - named-character (via
 * mergeStyles()) or named-paragraph (via mergeParagraphStyle(), e.g. a
 * bullet/numbered list style). Direct formatting never counts. */
function isNamedStyleOrigin(origin: StyleOrigin): origin is Extract<StyleOrigin, { styleId: string }> {
  return origin.kind === 'named-character' || origin.kind === 'named-paragraph'
}

/** Counts how many runs currently reference `styleId`, by re-scanning the
 * latest Style Report rather than hand-maintaining a running total on the
 * UserStyleRecord - keeps the "User-Created Styles" panel's occurrence
 * counts always in sync with the real document state. Matches both
 * character styles (w:rStyle) and paragraph styles (w:pStyle) - a
 * UserStyleRecord can be either kind (see UserStyleRecord.kind). */
export function countOccurrencesForStyleId(styleReport: StyleEntity[], styleId: string): number {
  let total = 0
  for (const entity of styleReport) {
    for (const variant of entity.variants) {
      if (isNamedStyleOrigin(variant.origin) && variant.origin.styleId === styleId) {
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
 * style (character or paragraph) that's tracked as a UserStyleRecord; every
 * other variant (direct formatting, or still under one of the document's
 * own original named styles) counts as remaining. */
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
      if (isNamedStyleOrigin(variant.origin) && userStyleIds.has(variant.origin.styleId)) {
        merged += 1
      }
    }
  }
  return { total, merged, remaining: total - merged }
}

/** The Style Report entries actually worth showing in StyleReportPanel:
 * variants already merged into a tracked UserStyleRecord (the same
 * "merged" test computeMergeProgress uses) are dropped - once an
 * occurrence's look is controlled by a User-Created style, re-selecting it
 * in the report would be a no-op (it's already exactly that style's look),
 * so leaving it visible is just clutter once the whole point of merging is
 * done. An entity left with zero variants is dropped entirely; one with
 * some-but-not-all variants merged gets a recomputed occurrenceCount so the
 * "N total across M sources" text stays consistent with what's still shown
 * (entities/variant arrays are returned as-is, unfiltered, when nothing
 * about them needed to change - avoids reallocating the common case). */
export function filterUnmergedEntities(styleReport: StyleEntity[], userStyles: UserStyleRecord[]): StyleEntity[] {
  const userStyleIds = new Set(userStyles.map((r) => r.styleId))
  const result: StyleEntity[] = []
  for (const entity of styleReport) {
    const variants = entity.variants.filter(
      (v) => !(isNamedStyleOrigin(v.origin) && userStyleIds.has(v.origin.styleId)),
    )
    if (variants.length === 0) continue
    if (variants.length === entity.variants.length) {
      result.push(entity)
    } else {
      const occurrenceCount = variants.reduce((sum, v) => sum + v.occurrenceCount, 0)
      result.push({ ...entity, variants, occurrenceCount })
    }
  }
  return result
}
