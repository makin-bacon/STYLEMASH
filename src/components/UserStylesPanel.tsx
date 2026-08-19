import type { ReferenceDocState } from '../hooks/useDocxWorkspace'
import type { StyleEntity, UserStyleRecord } from '../types/ooxml'
import type { ParagraphMarker } from '../lib/ooxml/numbering'
import { countOccurrencesForStyleId } from '../lib/ooxml/styleReport'
import { signatureToCss } from '../lib/signatureToCss'
import { AttachReferenceDocButton } from './AttachReferenceDocButton'
import { InfoTooltip } from './InfoTooltip'

/** Finds a representative list marker for a User-Created style, the same
 * way StyleReportPanel does for a Style Report variant - looked up by
 * styleId (rather than taking a variant directly) since a UserStyleRecord
 * doesn't have one, just the styleId every variant it controls shares.
 * Matches both character (w:rStyle) and paragraph (w:pStyle) origins, since
 * a UserStyleRecord can be either kind. Returns the first matching variant's
 * marker, so a style that's been merged from list-item occurrences keeps
 * showing that it's a list style here too, not just in the Style
 * Report/preview. */
function markerForStyleId(
  styleReport: StyleEntity[],
  styleId: string,
  paragraphMarkers: Map<Element, ParagraphMarker>,
): ParagraphMarker | undefined {
  for (const entity of styleReport) {
    for (const variant of entity.variants) {
      if (
        (variant.origin.kind !== 'named-character' && variant.origin.kind !== 'named-paragraph') ||
        variant.origin.styleId !== styleId
      ) {
        continue
      }
      const paragraphEl = variant.runRefs[0]?.paragraphElement
      const marker = paragraphEl && paragraphMarkers.get(paragraphEl)
      if (marker) return marker
    }
  }
  return undefined
}

interface UserStylesPanelProps {
  userStyles: UserStyleRecord[]
  styleReport: StyleEntity[]
  /** Resolved list marker per paragraph, shared with StyleReportPanel and
   * DocumentPreviewPanel - see markerForStyleId above. */
  paragraphMarkers: Map<Element, ParagraphMarker>
  onEditStyle: (styleId: string) => void
  onCreateNewStyle: () => void
  /** Populates the list with StyleMash's bundled starter style set (see
   * defaultStyles.ts) - a name collision with an existing style redefines
   * its look rather than duplicating it. User-initiated only; never runs
   * automatically. */
  onAddDefaultStyles: () => void
  /** The single style currently picked as a merge target (row click, not
   * "Edit") - null when none is. */
  selectedTargetStyleId: string | null
  onToggleSelectTarget: (styleId: string) => void
  /** Count of Style Report entries currently selected on the other panel -
   * drives the inline "Merge N selected here" action on the target row. */
  pendingSelectionCount: number
  onMergeSelectedIntoTarget: () => void
  /** Surfaced here (not just in MergeDialog) since MERGE_SELECTED_INTO_TARGET
   * has no dialog of its own to show it in. */
  mergeError: string | null
  /** Drives the footer's "Attach Document B (optional)" button while
   * nothing's attached, and its "Remove Document B" button once one is -
   * same footer slot either way, just swapping which button occupies it. */
  referenceDoc: ReferenceDocState
  onAttachReferenceDoc: (file: File) => void
  onRemoveReferenceDoc: () => void
  /** Wipes every User-Created style (record + <w:style> definition) in one
   * go. Pushes its own undo snapshot (see useDocxWorkspace), so an accidental
   * click is recoverable via the Undo button. */
  onClearUserStyles: () => void
}

/** Right-hand panel: the named styles StyleMash has created via merges
 * this session. Occurrence counts are always re-derived from the latest
 * Style Report (via countOccurrencesForStyleId) rather than hand-maintained,
 * so they can never drift out of sync with the actual document state.
 * "+ New Style" defines a style from scratch (0 occurrences until you merge
 * Style Report entries into it later via Edit, or target it directly from
 * the merge dialog when merging a fresh selection).
 *
 * Clicking a row (anywhere but "Edit") picks it as a merge target rather
 * than opening the edit dialog - the same "click to select" pattern as the
 * Style Report's rows, so selecting entries there and a target here is a
 * direct two-list gesture that ends with "Merge N selected here", no dialog
 * detour needed. "Edit" is still how you redefine a style's own look. */
export function UserStylesPanel({
  userStyles,
  styleReport,
  paragraphMarkers,
  onEditStyle,
  onCreateNewStyle,
  onAddDefaultStyles,
  selectedTargetStyleId,
  onToggleSelectTarget,
  pendingSelectionCount,
  onMergeSelectedIntoTarget,
  mergeError,
  referenceDoc,
  onAttachReferenceDoc,
  onRemoveReferenceDoc,
  onClearUserStyles,
}: UserStylesPanelProps) {
  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-hidden rounded-lg border border-slate-200 bg-white">
      <div className="flex items-start justify-between gap-2 border-b border-slate-200 bg-slate-800 px-4 py-4">
        <div>
          <h2 className="flex items-center gap-1.5 text-sm font-semibold text-slate-200">
            New Styles <span className="font-normal text-slate-400">({userStyles.length})</span>
            <InfoTooltip text="Select entries in the Style Report, then click a style here to merge them. To generate styles, hit the &quot;+ New Style&quot; button or upload a reference document." />
          </h2>
        </div>
        <div className="flex shrink-0 items-center gap-2">
          <button
            type="button"
            onClick={onAddDefaultStyles}
            className="rounded-md border border-indigo-200 px-2 py-1 text-xs font-medium text-slate-300 hover:bg-slate-600"
          >
            + Defaults
          </button>
          <button
            type="button"
            onClick={onCreateNewStyle}
            className="rounded-md border border-indigo-200 px-2 py-1 text-xs font-medium text-slate-300 hover:bg-slate-600"
          >
            + New Style
          </button>
        </div>
      </div>

      <ul className="min-h-0 flex-1 overflow-y-auto">
        {userStyles.length === 0 && (
          <li className="flex h-full items-center justify-center px-4 py-6 text-center text-sm text-slate-400">
            <p className="w-3/4">
              Select entries in the Style Report and click "Do it", or click "+ New Style"
              to define one from scratch.
            </p>
          </li>
        )}
        {userStyles.map((record) => {
          const occurrences = countOccurrencesForStyleId(styleReport, record.styleId)
          const markerText =
            markerForStyleId(styleReport, record.styleId, paragraphMarkers)?.text || record.listPreviewText
          const isTarget = selectedTargetStyleId === record.styleId
          return (
            <li
              key={record.styleId}
              onClick={() => onToggleSelectTarget(record.styleId)}
              className={`flex cursor-pointer items-start gap-3 border-b border-l-4 border-slate-200 px-4 py-3 transition-colors last:border-b-0 ${
                isTarget
                  ? 'border-l-indigo-500 bg-indigo-200 hover:bg-indigo-300 active:bg-indigo-400'
                  : 'border-l-transparent hover:border-l-indigo-300 hover:bg-slate-50 active:bg-slate-100'
              }`}
            >
              <div className="min-w-0 flex-1">
                <p className="truncate text-base font-medium" style={signatureToCss(record.targetSignature)}>
                  {markerText && <span className="mr-1 text-slate-400">{markerText}</span>}
                  {record.name}
                </p>
                <p className="mt-1 truncate text-xs text-slate-400">styleId: {record.styleId}</p>
                {isTarget && pendingSelectionCount > 0 && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation()
                      onMergeSelectedIntoTarget()
                    }}
                    className="mt-2 rounded-md bg-indigo-600 px-2 py-1 text-xs font-medium text-white hover:bg-indigo-700"
                  >
                    Merge {pendingSelectionCount} selected here
                  </button>
                )}
                {isTarget && mergeError && <p className="mt-1 text-xs text-red-600">{mergeError}</p>}
              </div>
              <div className="flex shrink-0 flex-col items-end gap-1">
                {record.fromReferenceDoc && (
                  <span className="rounded-full bg-indigo-50 px-2 py-0.5 text-xs font-medium text-indigo-600">
                    from Document B
                  </span>
                )}
                {record.kind === 'paragraph' && (
                  <span className="rounded-full bg-purple-50 px-2 py-0.5 text-xs font-medium text-purple-600">
                    {record.listFormat === 'bullet'
                      ? 'Bulleted list'
                      : record.listFormat === 'decimal'
                        ? 'Numbered list'
                        : 'Paragraph style'}
                  </span>
                )}
                <span className="rounded-full bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-600">
                  {occurrences}×
                </span>
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation()
                    onEditStyle(record.styleId)
                  }}
                  className="text-xs font-medium text-indigo-600 hover:text-indigo-800 hover:underline"
                >
                  Edit
                </button>
              </div>
            </li>
          )
        })}
      </ul>

      {/* Always rendered (never conditionally mounted) so this row's height
          never changes as Document B is attached/removed - same reasoning
          as DocumentPreviewPanel's own footer row. "Clear list" always
          occupies the left half; the right half still swaps between
          Attach/Remove Document B depending on referenceDoc.status. */}
      <div className="flex gap-2 border-t border-slate-200 px-4 py-2">
        <button
          type="button"
          disabled={userStyles.length === 0}
          onClick={onClearUserStyles}
          className="flex-1 rounded-md border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 enabled:hover:bg-slate-50 disabled:cursor-not-allowed disabled:text-slate-300"
        >
          Clear list
        </button>
        <div className="flex-1">
          {referenceDoc.status === 'loaded' ? (
            <button
              type="button"
              onClick={onRemoveReferenceDoc}
              className="w-full rounded-md border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
            >
              Remove Document B
            </button>
          ) : (
            <AttachReferenceDocButton
              status={referenceDoc.status}
              errorMessage={referenceDoc.errorMessage}
              onAttach={onAttachReferenceDoc}
            />
          )}
        </div>
      </div>
    </div>
  )
}
