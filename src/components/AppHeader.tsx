import { faFileWord } from '@fortawesome/free-regular-svg-icons'
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome'
import { useState } from 'react'
import { HelpModal } from './HelpModal'

interface AppHeaderProps {
  filename: string | null
}

/** Top bar: app name, currently-loaded filename, and Help. The Help
 * button/modal is self-contained (no workspace state involved), so its
 * open/closed state lives locally here rather than being lifted into
 * App.tsx. The "load a different file" action lives on StyleReportPanel's
 * header instead of here - see StyleReportPanel's "Mash a different file"
 * button. */
export function AppHeader({ filename }: AppHeaderProps) {
  const [isHelpOpen, setIsHelpOpen] = useState(false)

  return (
    <header className="flex items-center justify-between border-b border-slate-200 bg-slate-800 px-6 py-3">
      <div>
        <h1 className="flex items-center gap-2 text-xl font-bold text-slate-200">
          <FontAwesomeIcon icon={faFileWord} className="text-base" aria-hidden="true" />
          StyleMash
        </h1>
        {filename && <p className="text-xs text-slate-500">CURRENTLY MASHING: {filename}</p>}
      </div>

      <div className="flex items-center gap-3">
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
