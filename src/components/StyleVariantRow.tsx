import type { FormattingSignature, StyleEntityVariant } from '../types/ooxml'
import { signatureToCss } from '../lib/signatureToCss'
import { describeOrigin, describeSignature } from '../lib/styleDescriptions'

interface StyleVariantRowProps {
  signature: FormattingSignature
  variant: StyleEntityVariant
  selected: boolean
  onToggleSelect: () => void
  onEditXml: () => void
  /** True when rendered as a sub-row nested under a shared entity header
   * (i.e. this signature has more than one variant) - just adds indent. */
  indented?: boolean
}

/** One selectable row: every run that shares both a resolved visual
 * signature AND the same origin (a specific named style, or pure direct
 * formatting). This is the unit of selection for merging and the target of
 * "Edit XML" - so a style-derived instance and a direct-override instance
 * that happen to look identical can be cleaned up independently. */
export function StyleVariantRow({
  signature,
  variant,
  selected,
  onToggleSelect,
  onEditXml,
  indented,
}: StyleVariantRowProps) {
  return (
    <li
      className={`flex items-start gap-3 border-b border-slate-200 py-3 last:border-b-0 ${
        indented ? 'pl-9 pr-4' : 'px-4'
      } ${selected ? 'bg-indigo-50' : 'bg-white'}`}
    >
      <input
        type="checkbox"
        checked={selected}
        onChange={onToggleSelect}
        className="mt-1 size-4 shrink-0 accent-indigo-600"
        aria-label={`Select: ${describeOrigin(variant.origin)}`}
      />

      <div className="min-w-0 flex-1">
        <p className="truncate text-base" style={signatureToCss(signature)}>
          {variant.sampleText || '(no visible text)'}
        </p>
        {!indented && <p className="mt-1 truncate text-xs text-slate-500">{describeSignature(signature)}</p>}
        <p className="mt-0.5 truncate text-xs text-slate-400">{describeOrigin(variant.origin)}</p>
      </div>

      <div className="flex shrink-0 flex-col items-end gap-1">
        <span className="rounded-full bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-600">
          {variant.occurrenceCount}×
        </span>
        <button
          type="button"
          onClick={onEditXml}
          className="text-xs font-medium text-indigo-600 hover:text-indigo-800 hover:underline"
        >
          Edit XML
        </button>
      </div>
    </li>
  )
}
