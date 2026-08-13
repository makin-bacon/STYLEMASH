import type { ReactNode } from 'react'

interface AppHeaderProps {
  filename: string | null
  onLoadDifferentFile: () => void
  children?: ReactNode
}

/** Top bar: app name, currently-loaded filename, and a way back to the
 * upload screen. `children` is where App.tsx slots in the SaveButton, kept
 * as a prop rather than hardcoded here so this component stays a dumb
 * layout shell. */
export function AppHeader({ filename, onLoadDifferentFile, children }: AppHeaderProps) {
  return (
    <header className="flex items-center justify-between border-b border-slate-200 bg-gray-800 px-6 py-3">
      <div>
        <h1 className="text-lg font-bold text-slate-200">StyleRipper</h1>
        {filename && <p className="text-xs text-slate-500">CURRENTLY RIPPING: {filename}</p>}
      </div>

      <div className="flex items-center gap-3">
        {filename && (
          <button
            type="button"
            onClick={onLoadDifferentFile}
            className="rounded-md bg-slate-200 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-300"
          >
            Rip a different file
          </button>
        )}
        {children}
      </div>
    </header>
  )
}
