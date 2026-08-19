interface UndoButtonProps {
  disabled: boolean
  onUndo: () => void
}

/** Reverts the most recent style-merging action - folding Style Report
 * selections into a style ("Mash it"/"+ New Style"), "Merge N selected
 * here", the Document B bulk match, or "Clear list" - by popping one
 * snapshot off useDocxWorkspace's undo stack. Disabled once that stack is
 * empty. Rendered in DocumentPreviewPanel's header, immediately left of
 * SaveButton. */
export function UndoButton({ disabled, onUndo }: UndoButtonProps) {
  return (
    <button
      type="button"
      disabled={disabled}
      onClick={onUndo}
      className="rounded-md bg-orange-500 px-3 py-1.5 text-xs font-medium text-white enabled:hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-slate-300"
    >
      Undo
    </button>
  )
}
