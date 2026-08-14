import { useState } from 'react'
import { HelpModal } from './HelpModal'

interface AppHeaderProps {
  filename: string | null
  onLoadDifferentFile: () => void
}

/** Top bar: app name, currently-loaded filename, and a way back to the
 * upload screen. The Help button/modal is self-contained (no workspace
 * state involved), so its open/closed state lives locally here rather than
 * being lifted into App.tsx. */
export function AppHeader({ filename, onLoadDifferentFile }: AppHeaderProps) {
  const [isHelpOpen, setIsHelpOpen] = useState(false)

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
        <button
          type="button"
          onClick={() => setIsHelpOpen(true)}
          className="rounded-md border border-slate-500 px-3 py-1.5 text-xs font-medium text-slate-200 hover:bg-slate-700"
        >
          Help
        </button>
      </div>

      {isHelpOpen && <HelpModal onClose={() => setIsHelpOpen(false)} />}
    </header>
  )
}
