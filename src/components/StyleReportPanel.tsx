import type { StyleEntity, StyleEntityVariant } from '../types/ooxml'
import type { ParagraphMarker } from '../lib/ooxml/numbering'
import { signatureToCss } from '../lib/signatureToCss'
import { describeSignature } from '../lib/styleDescriptions'
import { InfoTooltip } from './InfoTooltip'
import { StyleVariantRow } from './StyleVariantRow'

/** A variant's list marker is taken from its first occurrence's paragraph -
 * the same paragraph that produced `variant.sampleText` (see
 * styleReport.ts#buildStyleReport: the run that sets sampleText is always
 * runRefs[0], since empty-text runs are skipped before either is touched).
 * So the marker shown always matches the sample text it's prefixed to. */
function markerFor(variant: StyleEntityVariant, paragraphMarkers: Map<Element, ParagraphMarker>) {
  const paragraphEl = variant.runRefs[0]?.paragraphElement
  return paragraphEl ? paragraphMarkers.get(paragraphEl) : undefined
}

interface StyleReportPanelProps {
  styleReport: StyleEntity[]
  selectedIds: Set<string>
  /** Resolved list marker per paragraph, shared with DocumentPreviewPanel -
   * lets a numbered/bulleted entry's sample text keep looking like a list
   * item here instead of a plain paragraph. */
  paragraphMarkers: Map<Element, ParagraphMarker>
  onToggleSelect: (variantId: string) => void
  onMergeSelected: () => void
  /** True once at least one style has been imported from Document B - the
   * bulk-match controls below are hidden entirely otherwise. */
  hasReferenceStyles: boolean
  bulkMergeError: string | null
  onSelectMatchingReferenceStyles: () => void
  onBulkMergeMatched: () => void
  /** How many Style Report entries are already merged into a User-Created
   * style vs. still outstanding - see styleReport.ts#computeMergeProgress. */
  mergeProgress: { total: number; merged: number; remaining: number }
  /** Same action as the "Save your work locally" button - offered again
   * here once every entry is matched, so there's a next step right where
   * the (now-empty-looking) list used to be. */
  onSave: () => void
  /** Drives this panel's own "Mash a different file" header button, plus
   * the equivalent link in the empty state below. */
  onRipAnotherFile: () => void
}

/** Left-hand panel: every distinct text style/appearance found in the
 * uploaded document, most common first. An entry with only one variant
 * renders as a single flat row; an entry that mixes e.g. style-derived text
 * with direct-override text of the same look (or list text that merely
 * happens to share formatting with non-list text - very common when a
 * document's lists use manual/direct formatting rather than a named style)
 * gets a shared header plus one selectable sub-row per variant, so they can
 * be merged independently rather than being silently lumped together. */
export function StyleReportPanel({
  styleReport,
  selectedIds,
  paragraphMarkers,
  onToggleSelect,
  onMergeSelected,
  hasReferenceStyles,
  bulkMergeError,
  onSelectMatchingReferenceStyles,
  onBulkMergeMatched,
  mergeProgress,
  onSave,
  onRipAnotherFile,
}: StyleReportPanelProps) {
  const percentMerged =
    mergeProgress.total === 0 ? 100 : Math.round((mergeProgress.merged / mergeProgress.total) * 100)
  const allMatched = mergeProgress.total > 0 && mergeProgress.remaining === 0
  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-hidden rounded-lg border border-slate-200 bg-white">
      <div className="flex items-start justify-between gap-2 border-b border-slate-200 bg-slate-800 px-4 py-4">
        <h2 className="flex items-center gap-1.5 text-sm font-semibold text-slate-200">
          Current styles <span className="font-normal text-slate-400">({styleReport.length})</span>
          <InfoTooltip text="Select items from this list and you'll see them in your document. Now either merge these with a newly created style or one from an uploaded reference Word file in the &quot;New&quot; styles list." />
        </h2>
        <button
          type="button"
          onClick={onRipAnotherFile}
          className="shrink-0 rounded-md border border-indigo-200 px-2 py-1 text-xs font-medium text-slate-300 hover:bg-slate-600"
        >
          Mash a different file
        </button>
      </div>

      <ul className="min-h-0 flex-1 overflow-y-auto">
        {allMatched ? (
          <li className="flex h-full items-center justify-center px-4 py-6 text-center text-sm text-slate-400">
            <p>
              <button
                type="button"
                onClick={onSave}
                className="text-xs font-medium text-indigo-600 hover:text-indigo-800 hover:underline"
              >
                Save your work
              </button>{' '}
              <br />or{' '}
              <button
                type="button"
                onClick={onRipAnotherFile}
                className="text-xs font-medium text-indigo-600 hover:text-indigo-800 hover:underline"
              >
                mash another file
              </button>
            </p>
          </li>
        ) : (
          <>
            {styleReport.length === 0 && (
              <li className="px-4 py-6 text-center text-sm text-slate-400">
                No formatted text found in this document.
              </li>
            )}
            {styleReport.map((entity) =>
              entity.variants.length === 1 ? (
                <StyleVariantRow
                  key={entity.variants[0].id}
                  signature={entity.signature}
                  variant={entity.variants[0]}
                  listMarker={markerFor(entity.variants[0], paragraphMarkers)}
                  selected={selectedIds.has(entity.variants[0].id)}
                  onToggleSelect={() => onToggleSelect(entity.variants[0].id)}
                />
              ) : (
                <li key={entity.id} className="border-b border-slate-200 last:border-b-0">
                  <div className="bg-slate-50 px-4 py-1.5">
                    <p className="truncate text-sm" style={signatureToCss(entity.signature)}>
                      {/* Marker and text both come from variants[0] (the most common
                          variant) rather than entity.sampleText - keeping them from the
                          same occurrence so the marker shown always matches the text
                          it's prefixed to (entity.sampleText is just "whichever text
                          appeared first in the document" across every variant, which
                          since list vs. non-list is its own variant dimension could
                          easily belong to a different variant than variants[0]). */}
                      {markerFor(entity.variants[0], paragraphMarkers)?.text && (
                        <span className="mr-1 text-slate-400">
                          {markerFor(entity.variants[0], paragraphMarkers)?.text}
                        </span>
                      )}
                      {entity.variants[0].sampleText || '(no visible text)'}
                    </p>
                    <p className="mt-0.5 truncate text-xs text-slate-500">
                      {describeSignature(entity.signature)} · {entity.occurrenceCount} total across{' '}
                      {entity.variants.length} sources
                    </p>
                  </div>
                  <ul>
                    {entity.variants.map((variant) => (
                      <StyleVariantRow
                        key={variant.id}
                        signature={entity.signature}
                        variant={variant}
                        listMarker={markerFor(variant, paragraphMarkers)}
                        indented
                        selected={selectedIds.has(variant.id)}
                        onToggleSelect={() => onToggleSelect(variant.id)}
                      />
                    ))}
                  </ul>
                </li>
              ),
            )}
          </>
        )}
      </ul>

      {/* Bulk-match panel disabled for now - needs more work and was adding
          stray padding under the page footer. Re-enable by restoring the
          `hasReferenceStyles &&` condition below. */}
      {false && hasReferenceStyles && (
        <div className="border-t border-slate-200 bg-indigo-50/50 px-4 py-2">
          <label className="flex items-center gap-2 text-xs text-slate-600">
            <input
              type="checkbox"
              onChange={(e) => {
                if (e.target.checked) onSelectMatchingReferenceStyles()
              }}
              className="accent-indigo-600"
            />
            Select styles matching Document B
          </label>
          <button
            type="button"
            disabled={selectedIds.size === 0}
            onClick={onBulkMergeMatched}
            className="mt-1.5 w-full rounded-md border border-indigo-300 bg-white px-3 py-1.5 text-xs font-medium text-indigo-700 enabled:hover:bg-indigo-100 disabled:cursor-not-allowed disabled:border-slate-200 disabled:text-slate-400"
          >
            Merge matched styles into Document B
          </button>
          {bulkMergeError && <p className="mt-1 text-xs text-red-600">{bulkMergeError}</p>}
        </div>
      )}

      <div className="border-t border-slate-200 px-4 py-2">
        <div className="pt-2 mb-4">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>
              {mergeProgress.merged} of {mergeProgress.total} merged into a User-Created style
            </span>
            <span>{mergeProgress.remaining} left</span>
          </div>
          <div className="mt-1 h-1.5 w-full overflow-hidden rounded-full bg-slate-200">
            <div
              className="h-full rounded-full bg-indigo-500 transition-[width]"
              style={{ width: `${percentMerged}%` }}
            />
          </div>
        </div>

        <button
          type="button"
          disabled={selectedIds.size === 0}
          onClick={onMergeSelected}
          className="w-full rounded-md bg-indigo-600 px-3 py-1.5 text-xs font-medium text-white enabled:hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300"
        >
          Mash it {selectedIds.size > 0 ? `(${selectedIds.size})` : ''}
        </button>
      </div>
    </div>
  )
}
