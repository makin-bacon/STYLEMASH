interface InfoTooltipProps {
  text: string
  className?: string
}

/** Small "i" affordance that surfaces `text` as a native browser tooltip on
 * hover/focus - used to move a panel header's explanatory paragraph out of
 * permanently-visible space without losing it entirely. A native title
 * tooltip (not a custom-positioned popup) is deliberate here: zero extra
 * markup/state, and consistent behavior everywhere it's dropped in. */
export function InfoTooltip({ text, className = '' }: InfoTooltipProps) {
  return (
    <span
      tabIndex={0}
      title={text}
      aria-label={text}
      className={`inline-flex h-4 w-4 shrink-0 cursor-help items-center justify-center rounded-full border border-current text-[10px] font-semibold leading-none opacity-60 hover:opacity-100 ${className}`}
    >
      i
    </span>
  )
}
