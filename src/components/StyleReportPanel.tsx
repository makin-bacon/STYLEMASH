import type { StyleEntity } from '../types/ooxml'
import { signatureToCss } from '../lib/signatureToCss'
import { describeSignature } from '../lib/styleDescriptions'
import { StyleVariantRow } from './StyleVariantRow'

interface StyleReportPanelProps {
  styleReport: StyleEntity[]
  selectedIds: Set<string>
  onToggleSelect: (variantId: string) => void
  onEditXml: (variantId: string) => void
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
  /** Same action as AppHeader's "Rip a different file" button. */
  onRipAnotherFile: () => void
}

/** Left-hand panel: every distinct text style/appearance found in the
 * uploaded document, most common first. An entry with only one variant
 * renders as a single flat row; an entry that mixes e.g. style-derived text
 * with direct-override text of the same look gets a shared header plus one
 * selectable sub-row per variant, so they can be merged independently. */
export function StyleReportPanel({
  styleReport,
  selectedIds,
  onToggleSelect,
  onEditXml,
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
      <div className="border-b border-slate-200 bg-slate-50 px-4 py-2">
        <h2 className="text-sm font-semibold text-slate-700">
          Style Report <span className="font-normal text-slate-400">({styleReport.length})</span>
        </h2>
        <p className="text-xs text-slate-500">
          Every distinct text appearance found in the document - select entries to merge them into
          one style.
        </p>
      </div>

      <ul className="min-h-0 flex-1 overflow-y-auto">
        {allMatched ? (
          <li className="flex h-full items-center justify-center px-4 py-6 text-center text-sm text-slate-400">
            <p>
              <button
                type="button"
                onClick={onSave}
                className="font-medium text-indigo-600 hover:text-indigo-800 hover:underline"
              >
                Save your work
              </button>{' '}
              <br />or{' '}
              <button
                type="button"
                onClick={onRipAnotherFile}
                className="font-medium text-indigo-600 hover:text-indigo-800 hover:underline"
              >
                rip another file
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
                  selected={selectedIds.has(entity.variants[0].id)}
                  onToggleSelect={() => onToggleSelect(entity.variants[0].id)}
                  onEditXml={() => onEditXml(entity.variants[0].id)}
                />
              ) : (
                <li key={entity.id} className="border-b border-slate-200 last:border-b-0">
                  <div className="bg-slate-50 px-4 py-1.5">
                    <p className="truncate text-sm" style={signatureToCss(entity.signature)}>
                      {entity.sampleText || '(no visible text)'}
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
                        indented
                        selected={selectedIds.has(variant.id)}
                        onToggleSelect={() => onToggleSelect(variant.id)}
                        onEditXml={() => onEditXml(variant.id)}
                      />
                    ))}
                  </ul>
                </li>
              ),
            )}
          </>
        )}
      </ul>

      {hasReferenceStyles && (
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
          Do it {selectedIds.size > 0 ? `(${selectedIds.size})` : ''}
        </button>
      </div>
    </div>
  )
}
