import type { ReactNode } from 'react'
import type { StyleEntity, UserStyleRecord } from '../types/ooxml'
import { countOccurrencesForStyleId } from '../lib/ooxml/styleReport'
import { signatureToCss } from '../lib/signatureToCss'

interface UserStylesPanelProps {
  userStyles: UserStyleRecord[]
  styleReport: StyleEntity[]
  onEditStyle: (styleId: string) => void
  onCreateNewStyle: () => void
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
  /** Footer slot for SaveButton, kept as a prop (rather than hardcoded here)
   * for the same reason AppHeader takes it as children - this component
   * stays a dumb layout shell. */
  children?: ReactNode
}

/** Right-hand panel: the named styles StyleRipper has created via merges
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
  onEditStyle,
  onCreateNewStyle,
  selectedTargetStyleId,
  onToggleSelectTarget,
  pendingSelectionCount,
  onMergeSelectedIntoTarget,
  mergeError,
  children,
}: UserStylesPanelProps) {
  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-hidden rounded-lg border border-slate-200 bg-white">
      <div className="flex items-start justify-between gap-2 border-b border-slate-200 bg-slate-50 px-4 py-2">
        <div>
          <h2 className="text-sm font-semibold text-slate-700">
            User-Created Styles <span className="font-normal text-slate-400">({userStyles.length})</span>
          </h2>
          <p className="text-xs text-slate-500">
            Named styles created by merging Style Report entries. Select entries in the Style Report,
            then click a style here to merge them into it.
          </p>
        </div>
        <button
          type="button"
          onClick={onCreateNewStyle}
          className="shrink-0 rounded-md border border-indigo-200 px-2 py-1 text-xs font-medium text-indigo-600 hover:bg-indigo-50"
        >
          + New Style
        </button>
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
          const isTarget = selectedTargetStyleId === record.styleId
          return (
            <li
              key={record.styleId}
              onClick={() => onToggleSelectTarget(record.styleId)}
              className={`flex cursor-pointer items-start gap-3 border-b border-l-4 border-slate-200 px-4 py-3 transition-colors last:border-b-0 ${
                isTarget
                  ? 'border-l-indigo-500 bg-indigo-50 hover:bg-indigo-100 active:bg-indigo-200'
                  : 'border-l-transparent hover:border-l-indigo-300 hover:bg-slate-50 active:bg-slate-100'
              }`}
            >
              <div className="min-w-0 flex-1">
                <p className="truncate text-base font-medium" style={signatureToCss(record.targetSignature)}>
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

      {children && <div className="border-t border-slate-200 px-4 py-2">{children}</div>}
    </div>
  )
}
