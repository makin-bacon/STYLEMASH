/** Persistent bottom-of-page privacy note - shown regardless of workspace
 * state (upload screen or loaded workspace), so it's always visible rather
 * than only while the dropzone is on screen. */
export function AppFooter() {
  return (
    <footer className="border-t border-slate-200 bg-gray-800 px-6 py-2 text-center text-xs text-slate-200">
      Everything happens locally in your browser - files are never uploaded anywhere.
    </footer>
  )
}
