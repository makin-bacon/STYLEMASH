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
      onClick={onToggleSelect}
      className={`flex cursor-pointer items-start gap-3 border-b border-l-4 border-slate-200 py-3 transition-colors last:border-b-0 ${
        indented ? 'pl-8 pr-4' : 'px-4'
      } ${
        selected
          ? 'border-l-indigo-500 bg-indigo-50 hover:bg-indigo-100 active:bg-indigo-200'
          : 'border-l-transparent bg-white hover:bg-slate-50 active:bg-slate-100'
      }`}
    >
      {/* Visually hidden, not removed - the row's highlight color is the
          visible selected-state indicator, but this keeps the row
          keyboard-focusable/toggleable and announced correctly by screen
          readers. */}
      <input
        type="checkbox"
        checked={selected}
        onChange={onToggleSelect}
        onClick={(e) => e.stopPropagation()}
        className="sr-only"
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
          onClick={(e) => {
            e.stopPropagation()
            onEditXml()
          }}
          className="text-xs font-medium text-indigo-600 hover:text-indigo-800 hover:underline"
        >
          Edit XML
        </button>
      </div>
    </li>
  )
}
