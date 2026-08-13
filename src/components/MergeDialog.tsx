import { useState } from 'react'
import type { FormattingSignature, UserStyleRecord } from '../types/ooxml'
import { signatureToCss } from '../lib/signatureToCss'

const DEFAULT_SIGNATURE: FormattingSignature = {
  fontFamily: 'Calibri',
  fontSizeHalfPt: 24, // 12pt
  colorValue: 'auto',
  bold: false,
  italic: false,
  underline: null,
  strike: false,
}

const UNDERLINE_OPTIONS = [
  { value: '', label: 'None' },
  { value: 'single', label: 'Single' },
  { value: 'double', label: 'Double' },
  { value: 'thick', label: 'Thick' },
  { value: 'wave', label: 'Wave' },
  { value: 'dotted', label: 'Dotted' },
  { value: 'dash', label: 'Dashed' },
]

/** Flat, editable-field shape derived from a FormattingSignature - used both
 * for this dialog's initial state and to re-prefill every field at once when
 * the user switches the "merge into" target in the dropdown below. */
function fieldsFromSignature(sig: FormattingSignature) {
  return {
    fontFamily: sig.fontFamily ?? 'Calibri',
    fontSizePt: String((sig.fontSizeHalfPt ?? 24) / 2),
    colorAuto: sig.colorValue === 'auto',
    colorHex: sig.colorValue === 'auto' ? '#000000' : `#${sig.colorValue}`,
    bold: sig.bold,
    italic: sig.italic,
    strike: sig.strike,
    underline: sig.underline ?? '',
  }
}

interface MergeDialogProps {
  /** Total occurrences currently selected in the Style Report (0 for the
   * "+ New Style" flow, which just defines a style with nothing merged into
   * it yet). */
  selectedCount: number
  /** Look of the first selected variant's entity, used as the starting
   * point for the editable fields - null when nothing is selected. */
  baselineSignature: FormattingSignature | null
  /** Every style created so far this session, offered as merge targets. */
  userStyles: UserStyleRecord[]
  /** Set when opened via UserStylesPanel's "Edit" button: locks the dialog
   * to redefining this one specific style (no target dropdown). */
  reuseRecord: UserStyleRecord | null
  error: string | null
  onConfirm: (targetProps: FormattingSignature, name: string, targetStyleId?: string) => void
  onCancel: () => void
}

/** Modal for creating a style, redefining an existing one, or merging a
 * Style Report selection into either. The parent remounts this component
 * with a fresh `key` each time it opens, so this internal draft state never
 * needs to be reset manually. */
export function MergeDialog({
  selectedCount,
  baselineSignature,
  userStyles,
  reuseRecord,
  error,
  onConfirm,
  onCancel,
}: MergeDialogProps) {
  // Whether merging a fresh selection can additionally target an existing
  // style instead of always creating a new one. Not offered when there's
  // nothing selected (that's "+ New Style", a from-scratch definition) or
  // when reuseRecord already locks the target (the "Edit" flow).
  const canTargetExisting = !reuseRecord && selectedCount > 0 && userStyles.length > 0

  const [targetStyleId, setTargetStyleId] = useState(reuseRecord?.styleId ?? '')
  const baseline = reuseRecord?.targetSignature ?? baselineSignature ?? DEFAULT_SIGNATURE
  const [name, setName] = useState(reuseRecord?.name ?? 'Custom Style')
  const [fields, setFields] = useState(fieldsFromSignature(baseline))

  const handleTargetChange = (id: string) => {
    setTargetStyleId(id)
    const record = id ? userStyles.find((r) => r.styleId === id) : undefined
    if (record) {
      setName(record.name)
      setFields(fieldsFromSignature(record.targetSignature))
    }
  }

  const draftSignature: FormattingSignature = {
    fontFamily: fields.fontFamily.trim() || null,
    fontSizeHalfPt: Math.round((Number.parseFloat(fields.fontSizePt) || 12) * 2),
    colorValue: fields.colorAuto ? 'auto' : fields.colorHex.replace('#', '').toUpperCase(),
    bold: fields.bold,
    italic: fields.italic,
    underline: fields.underline || null,
    strike: fields.strike,
  }

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    if (!name.trim()) return
    onConfirm(draftSignature, name.trim(), reuseRecord?.styleId ?? targetStyleId ?? undefined)
  }

  const heading = reuseRecord
    ? `Edit "${reuseRecord.name}"`
    : selectedCount > 0
      ? 'Merge into a style'
      : 'Create a new style'

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4">
      <div className="w-full max-w-md rounded-lg bg-white p-5 shadow-xl">
        <h2 className="text-lg font-semibold text-slate-800">{heading}</h2>

        {!reuseRecord && selectedCount > 0 && (
          <p className="mt-1 text-xs text-slate-500">
            Merging {selectedCount} selected occurrence{selectedCount === 1 ? '' : 's'} into one named
            style.
          </p>
        )}
        {!reuseRecord && selectedCount === 0 && (
          <p className="mt-1 text-xs text-slate-500">
            Defines a style with no text merged into it yet - use "Edit" on it later, or select it
            as a target here after selecting Style Report entries.
          </p>
        )}

        <form onSubmit={handleSubmit} className="mt-4 space-y-3">
          {canTargetExisting && (
            <label className="block text-sm">
              <span className="text-slate-600">Merge into</span>
              <select
                value={targetStyleId}
                onChange={(e) => handleTargetChange(e.target.value)}
                className="mt-1 w-full rounded-md border border-slate-300 px-2 py-1.5 text-sm focus:border-indigo-500 focus:outline-none"
              >
                <option value="">A new style</option>
                {userStyles.map((record) => (
                  <option key={record.styleId} value={record.styleId}>
                    {record.name}
                    {record.fromReferenceDoc ? ' (Document B)' : ''}
                  </option>
                ))}
              </select>
            </label>
          )}

          <label className="block text-sm">
            <span className="text-slate-600">Style name</span>
            <input
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              required
              className="mt-1 w-full rounded-md border border-slate-300 px-2 py-1.5 text-sm focus:border-indigo-500 focus:outline-none"
            />
          </label>

          <div className="grid grid-cols-2 gap-3">
            <label className="block text-sm">
              <span className="text-slate-600">Font</span>
              <input
                type="text"
                value={fields.fontFamily}
                onChange={(e) => setFields({ ...fields, fontFamily: e.target.value })}
                className="mt-1 w-full rounded-md border border-slate-300 px-2 py-1.5 text-sm focus:border-indigo-500 focus:outline-none"
              />
            </label>
            <label className="block text-sm">
              <span className="text-slate-600">Size (pt)</span>
              <input
                type="number"
                min="1"
                step="0.5"
                value={fields.fontSizePt}
                onChange={(e) => setFields({ ...fields, fontSizePt: e.target.value })}
                className="mt-1 w-full rounded-md border border-slate-300 px-2 py-1.5 text-sm focus:border-indigo-500 focus:outline-none"
              />
            </label>
          </div>

          <div className="flex items-center gap-3">
            <label className="flex items-center gap-2 text-sm text-slate-600">
              <input
                type="checkbox"
                checked={fields.colorAuto}
                onChange={(e) => setFields({ ...fields, colorAuto: e.target.checked })}
                className="accent-indigo-600"
              />
              Automatic color
            </label>
            {!fields.colorAuto && (
              <input
                type="color"
                value={fields.colorHex}
                onChange={(e) => setFields({ ...fields, colorHex: e.target.value })}
                className="h-7 w-10 cursor-pointer rounded border border-slate-300"
              />
            )}
          </div>

          <div className="flex flex-wrap items-center gap-4 text-sm text-slate-600">
            <label className="flex items-center gap-1.5">
              <input
                type="checkbox"
                checked={fields.bold}
                onChange={(e) => setFields({ ...fields, bold: e.target.checked })}
                className="accent-indigo-600"
              />
              Bold
            </label>
            <label className="flex items-center gap-1.5">
              <input
                type="checkbox"
                checked={fields.italic}
                onChange={(e) => setFields({ ...fields, italic: e.target.checked })}
                className="accent-indigo-600"
              />
              Italic
            </label>
            <label className="flex items-center gap-1.5">
              <input
                type="checkbox"
                checked={fields.strike}
                onChange={(e) => setFields({ ...fields, strike: e.target.checked })}
                className="accent-indigo-600"
              />
              Strikethrough
            </label>
            <label className="flex items-center gap-1.5">
              Underline
              <select
                value={fields.underline}
                onChange={(e) => setFields({ ...fields, underline: e.target.value })}
                className="rounded-md border border-slate-300 px-1.5 py-1 text-sm"
              >
                {UNDERLINE_OPTIONS.map((opt) => (
                  <option key={opt.value} value={opt.value}>
                    {opt.label}
                  </option>
                ))}
              </select>
            </label>
          </div>

          <div className="rounded-md border border-slate-200 bg-slate-50 px-3 py-2">
            <p className="text-xs text-slate-400">Preview</p>
            <p style={signatureToCss(draftSignature)}>The quick brown fox jumps over the lazy dog</p>
          </div>

          {error && <p className="text-sm text-red-600">{error}</p>}

          <div className="flex justify-end gap-2 pt-2">
            <button
              type="button"
              onClick={onCancel}
              className="rounded-md px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-100"
            >
              Cancel
            </button>
            <button
              type="submit"
              className="rounded-md bg-indigo-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-indigo-700"
            >
              {reuseRecord || targetStyleId ? 'Save changes' : 'Create style'}
            </button>
          </div>
        </form>
      </div>
    </div>
  )
}
