import type { ParsedDocx, StyleEntity, UserStyleRecord } from '../../types/ooxml'
import { mergeStyles, removeStyleById } from './mergeStyles'
import { countOccurrencesForStyleId, buildStyleReport } from './styleReport'
import { buildResolutionContext, buildStylesMap, getDocDefaultsRPr, resolveStyleRPr } from './styleResolution'
import { trackedChildrenToSignature } from './signature'
import { buildThemeColorMap, resolveColorElement } from './themeColor'

/** Materializes every named style Document B actually uses in its own body
 * text as a real character-type <w:style> in `targetDocx` (Document A)'s
 * stylesXml - via the existing mergeStyles() "+ New Style" code path
 * (empty sourceRunRefs), so a materialized style is byte-for-byte the same
 * shape, and behaves identically as a reuse target, as any manually-created
 * UserStyleRecord. Mutates targetDocx.stylesXml in place; referenceDocx is
 * read-only.
 *
 * "Used" means: is the origin styleId of some Style Report variant in
 * Document B whose look isn't fully masked by direct formatting (origin
 * kind 'named-character' or 'named-paragraph') - reuses the app's existing
 * notion of "where a run's formatting came from" rather than a bespoke
 * scan. A style's materialized signature is its own fully-cascaded look
 * (through Document B's own basedOn chain to Document B's own docDefaults
 * and theme) - not any particular Style Report entity's signature, which is
 * paragraph-context-dependent for character styles (see resolveRunFormatting). */
export function materializeReferenceDocStyles(
  targetDocx: ParsedDocx,
  referenceDocx: ParsedDocx,
): UserStyleRecord[] {
  const styleReportB = buildStyleReport(referenceDocx)
  const stylesMapB = buildStylesMap(referenceDocx.stylesXml)
  const themeColorsB = buildThemeColorMap(referenceDocx.themeXml)
  const ctxB = buildResolutionContext(stylesMapB, getDocDefaultsRPr(referenceDocx.stylesXml), themeColorsB)

  const usedStyleIds = new Set<string>()
  for (const entity of styleReportB) {
    for (const variant of entity.variants) {
      if (variant.origin.kind === 'named-character' || variant.origin.kind === 'named-paragraph') {
        usedStyleIds.add(variant.origin.styleId)
      }
    }
  }
  // Set iteration order == insertion order == styleReportB's own
  // most-common-first order, so the resulting records land in a sensible
  // order in the User-Created Styles panel for free.

  const records: UserStyleRecord[] = []
  for (const bStyleId of usedStyleIds) {
    const bStyle = stylesMapB.get(bStyleId)
    if (!bStyle) continue // defensive; origin always points at a real StyleDef

    const tracked = resolveStyleRPr(bStyleId, ctxB.stylesMap, ctxB.docDefaultsTracked, ctxB.styleRPrCache)
    const signature = trackedChildrenToSignature(tracked, (colorEl) =>
      resolveColorElement(colorEl, themeColorsB),
    )

    const newStyleId = mergeStyles(targetDocx, [], signature, bStyle.name)

    records.push({
      styleId: newStyleId,
      name: bStyle.name,
      targetSignature: signature,
      createdAt: Date.now(),
      fromReferenceDoc: true,
    })
  }
  return records
}

/** The other half of the "remove Document B" lifecycle: for every
 * `fromReferenceDoc` record, keeps it (stripped of the flag) if it's
 * actually in use, or fully removes it (record + <w:style> definition) if
 * it never got merged into - a complete undo. Every materialized style is
 * type="character" and only ever applied via w:rStyle (never w:pStyle), so
 * countOccurrencesForStyleId is a complete usage count for these ids, not
 * an approximation. Pure - the caller is expected to bump its own
 * `parsedDocx` wrapper since `stylesXml` is mutated in place. */
export function reconcileUserStylesOnReferenceDocRemoval(
  userStyles: UserStyleRecord[],
  styleReport: StyleEntity[],
  stylesXml: XMLDocument,
): UserStyleRecord[] {
  const kept: UserStyleRecord[] = []
  for (const record of userStyles) {
    if (!record.fromReferenceDoc) {
      kept.push(record)
      continue
    }
    const occurrences = countOccurrencesForStyleId(styleReport, record.styleId)
    if (occurrences > 0) {
      const { fromReferenceDoc: _drop, ...rest } = record
      kept.push(rest)
    } else {
      removeStyleById(stylesXml, record.styleId)
      // record dropped entirely
    }
  }
  return kept
}
