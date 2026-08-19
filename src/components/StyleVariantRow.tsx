import type { FormattingSignature, StyleEntityVariant } from '../types/ooxml'
import type { ParagraphMarker } from '../lib/ooxml/numbering'
import { signatureToCss } from '../lib/signatureToCss'
import { describeOrigin, describeSignature } from '../lib/styleDescriptions'

interface StyleVariantRowProps {
  signature: FormattingSignature
  variant: StyleEntityVariant
  selected: boolean
  onToggleSelect: () => void
  /** This variant's list marker (bullet/number/letter), if its sample text
   * comes from a list paragraph - shown ahead of the sample text so a list
   * entry doesn't read as an ordinary paragraph while merging. */
  listMarker?: ParagraphMarker
  /** True when rendered as a sub-row nested under a shared entity header
   * (i.e. this signature has more than one variant) - just adds indent. */
  indented?: boolean
}

/** One selectable row: every run that shares both a resolved visual
 * signature AND the same origin (a specific named style, or pure direct
 * formatting) AND the same list membership. This is the unit of selection
 * for merging - so e.g. a bulleted-list instance and a plain-paragraph
 * instance that happen to share identical character formatting (very
 * common when a document's lists use manually-applied direct formatting
 * rather than a named style) stay independently selectable, instead of
 * being silently folded into one "Normal text" row that would merge the
 * list text right along with it. */
export function StyleVariantRow({
  signature,
  variant,
  selected,
  onToggleSelect,
  listMarker,
  indented,
}: StyleVariantRowProps) {
  return (
    <li
      onClick={onToggleSelect}
      className={`flex cursor-pointer items-start gap-3 border-b border-l-4 border-slate-200 py-3 transition-colors last:border-b-0 ${
        indented ? 'pl-8 pr-4' : 'px-4'
      } ${
        selected
          ? 'border-l-indigo-500 bg-indigo-200 hover:bg-indigo-300 active:bg-indigo-400'
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
          {listMarker?.text && <span className="mr-1 text-slate-400">{listMarker.text}</span>}
          {variant.sampleText || '(no visible text)'}
        </p>
        {!indented && <p className="mt-1 truncate text-xs text-slate-500">{describeSignature(signature)}</p>}
        <p className="mt-0.5 truncate text-xs text-slate-400">{describeOrigin(variant.origin)}</p>
      </div>

      <div className="flex shrink-0 flex-col items-end gap-1">
        <span className="rounded-full bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-600">
          {variant.occurrenceCount}×
        </span>
      </div>
    </li>
  )
}
