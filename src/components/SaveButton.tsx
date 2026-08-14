interface SaveButtonProps {
  disabled: boolean
  isSaving: boolean
  onSave: () => void
}

/** Always-available save action. Downloads the (possibly edited) document
 * locally in its original format, filename suffixed with "-RIPPED" - never
 * uploads anything anywhere, since StyleRipper does all processing
 * client-side in the browser. Rendered at the bottom of UserStylesPanel,
 * full-width to match that panel's "Do it"-style footer buttons. */
export function SaveButton({ disabled, isSaving, onSave }: SaveButtonProps) {
  return (
    <button
      type="button"
      disabled={disabled || isSaving}
      onClick={onSave}
      className="w-full rounded-md bg-emerald-600 px-3 py-1.5 text-xs font-medium text-white enabled:hover:bg-emerald-700 disabled:cursor-not-allowed disabled:bg-slate-300"
    >
      {isSaving ? 'Saving…' : 'Save your work locally'}
    </button>
  )
}
