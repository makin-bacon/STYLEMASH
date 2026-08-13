import type { StyleEntity, UserStyleRecord } from '../types/ooxml'
import { countOccurrencesForStyleId } from '../lib/ooxml/styleReport'
import { signatureToCss } from '../lib/signatureToCss'

interface UserStylesPanelProps {
  userStyles: UserStyleRecord[]
  styleReport: StyleEntity[]
  onEditStyle: (styleId: string) => void
  onCreateNewStyle: () => void
}

/** Right-hand panel: the named styles StyleRipper has created via merges
 * this session. Occurrence counts are always re-derived from the latest
 * Style Report (via countOccurrencesForStyleId) rather than hand-maintained,
 * so they can never drift out of sync with the actual document state.
 * "+ New Style" defines a style from scratch (0 occurrences until you merge
 * Style Report entries into it later via Edit, or target it directly from
 * the merge dialog when merging a fresh selection). */
export function UserStylesPanel({ userStyles, styleReport, onEditStyle, onCreateNewStyle }: UserStylesPanelProps) {
  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-hidden rounded-lg border border-slate-200 bg-white">
      <div className="flex items-start justify-between gap-2 border-b border-slate-200 bg-slate-50 px-4 py-2">
        <div>
          <h2 className="text-sm font-semibold text-slate-700">
            User-Created Styles <span className="font-normal text-slate-400">({userStyles.length})</span>
          </h2>
          <p className="text-xs text-slate-500">
            Named styles created by merging Style Report entries. These are written into the
            document's styles.xml on save.
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
          <li className="px-4 py-6 text-center text-sm text-slate-400">
            Select entries in the Style Report and click "Merge Selected", or click "+ New Style"
            to define one from scratch.
          </li>
        )}
        {userStyles.map((record) => {
          const occurrences = countOccurrencesForStyleId(styleReport, record.styleId)
          return (
            <li
              key={record.styleId}
              className="flex items-start gap-3 border-b border-slate-200 px-4 py-3 last:border-b-0"
            >
              <div className="min-w-0 flex-1">
                <p className="truncate text-base font-medium" style={signatureToCss(record.targetSignature)}>
                  {record.name}
                </p>
                <p className="mt-1 truncate text-xs text-slate-400">styleId: {record.styleId}</p>
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
                  onClick={() => onEditStyle(record.styleId)}
                  className="text-xs font-medium text-indigo-600 hover:text-indigo-800 hover:underline"
                >
                  Edit
                </button>
              </div>
            </li>
          )
        })}
      </ul>
    </div>
  )
}
