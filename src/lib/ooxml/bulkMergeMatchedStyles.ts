import type { ParsedDocx, RunRef, StyleEntity, UserStyleRecord } from '../../types/ooxml'
import { mergeStyles } from './mergeStyles'

/** Every currently-selectable Style Report variant id whose visible origin
 * is a named style (character or paragraph) sharing its name with a style
 * materialized from Document B (see referenceDocStyles.ts) - i.e. "this
 * text in Document A is under a style Document B also defines by name."
 * Direct-formatting variants never match, since they have no style name to
 * compare. */
export function findVariantIdsMatchingReferenceStyleNames(
  styleReport: StyleEntity[],
  userStyles: UserStyleRecord[],
): Set<string> {
  const referenceNames = new Set(userStyles.filter((r) => r.fromReferenceDoc).map((r) => r.name))
  const matched = new Set<string>()
  if (referenceNames.size === 0) return matched

  for (const entity of styleReport) {
    for (const variant of entity.variants) {
      if (
        (variant.origin.kind === 'named-character' || variant.origin.kind === 'named-paragraph') &&
        referenceNames.has(variant.origin.styleName)
      ) {
        matched.add(variant.id)
      }
    }
  }
  return matched
}

/** Bulk-merges every selected variant whose origin style name matches a
 * Document-B-derived UserStyleRecord into that specific record's style -
 * one mergeStyles() call per distinct matched name, so a selection spanning
 * several different Document B style names ends up correctly split across
 * their respective targets in a single action, rather than folded into one
 * (which is what the single-target MergeDialog flow would do instead).
 * Selected variants that don't match any Document-B-derived name (direct
 * formatting, or a name Document B doesn't define) are left untouched - not
 * an error, just skipped. Mutates `parsedDocx` in place, same as
 * mergeStyles() itself. */
export function bulkMergeVariantsIntoMatchingReferenceStyles(
  parsedDocx: ParsedDocx,
  styleReport: StyleEntity[],
  selectedVariantIds: Set<string>,
  userStyles: UserStyleRecord[],
): void {
  const referenceRecordsByName = new Map(
    userStyles.filter((r) => r.fromReferenceDoc).map((r) => [r.name, r] as const),
  )
  if (referenceRecordsByName.size === 0) return

  const buckets = new Map<string, { record: UserStyleRecord; runRefs: RunRef[] }>()
  for (const entity of styleReport) {
    for (const variant of entity.variants) {
      if (!selectedVariantIds.has(variant.id)) continue
      if (variant.origin.kind !== 'named-character' && variant.origin.kind !== 'named-paragraph') continue
      const record = referenceRecordsByName.get(variant.origin.styleName)
      if (!record) continue

      let bucket = buckets.get(record.styleId)
      if (!bucket) {
        bucket = { record, runRefs: [] }
        buckets.set(record.styleId, bucket)
      }
      bucket.runRefs.push(...variant.runRefs)
    }
  }

  for (const { record, runRefs } of buckets.values()) {
    mergeStyles(parsedDocx, runRefs, record.targetSignature, record.name, record.styleId)
  }
}
