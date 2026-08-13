import { useState } from 'react'
import type { StyleEntityVariant } from '../types/ooxml'
import { serializeFirstRunRPr } from '../lib/ooxml/xmlFragmentEdit'

interface XmlEditorModalProps {
  variant: StyleEntityVariant
  error: string | null
  onApply: (fragmentText: string) => void
  onCancel: () => void
}

/** The "simple interface" raw-XML editor: a plain <textarea> showing the
 * variant's underlying <w:rPr>, editable and re-applied to every run in that
 * variant on save. Deliberately not a syntax-highlighting code editor (e.g.
 * CodeMirror) - these fragments are small enough that a textarea plus clear
 * validation errors covers the need without the added dependency. */
export function XmlEditorModal({ variant, error, onApply, onCancel }: XmlEditorModalProps) {
  const [text, setText] = useState(() => serializeFirstRunRPr(variant.runRefs))

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4">
      <div className="w-full max-w-lg rounded-lg bg-white p-5 shadow-xl">
        <h2 className="text-lg font-semibold text-slate-800">Edit raw XML</h2>
        <p className="mt-1 text-xs text-slate-500">
          Editing the &lt;w:rPr&gt; run properties for all {variant.occurrenceCount} occurrence
          {variant.occurrenceCount === 1 ? '' : 's'} of this variant.
        </p>

        <textarea
          value={text}
          onChange={(e) => setText(e.target.value)}
          spellCheck={false}
          rows={12}
          className="mt-3 w-full rounded-md border border-slate-300 bg-slate-50 p-2 font-mono text-xs text-slate-800 focus:border-indigo-500 focus:outline-none"
        />

        {error && <p className="mt-2 text-sm text-red-600">{error}</p>}

        <div className="mt-4 flex justify-end gap-2">
          <button
            type="button"
            onClick={onCancel}
            className="rounded-md px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-100"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={() => onApply(text)}
            className="rounded-md bg-indigo-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-indigo-700"
          >
            Apply
          </button>
        </div>
      </div>
    </div>
  )
}
