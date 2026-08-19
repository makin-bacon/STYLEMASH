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
          {/* StyleMash's own mark (see /favicon.svg) - inlined as JSX rather
              than an <img> so its fill="currentColor" keeps inheriting this
              heading's text color, same as the FontAwesome glyph it
              replaces did. */}
          <svg viewBox="0 0 640 640" className="h-4 w-4" aria-hidden="true">
            <path
              fill="currentColor"
              d="M40 200C40 89.5 129.5 0 240 0L580 0c11 0 20 9 20 20s-9 20-20 20l-60 0 0 580c0 11-9 20-20 20s-20-9-20-20l0-580-80 0 0 580c0 11-9 20-20 20s-20-9-20-20l0-220-25.8 0c-7.2-14.5-16-27.9-26.1-40l51.9 0 0-320-120 0c-88.4 0-160 71.6-160 160 0 33.8 10.5 65.1 28.3 90.9-13.9 3.1-27 7.5-39.2 13.1-18.5-30.3-29.1-65.9-29.1-104zM0 487c0-84.5 68.5-153 153-153S306 402.5 306 487 237.5 640 153 640 0 571.5 0 487zM153 368C87.3 368 34 421.3 34 487s53.3 119 119 119 119-53.3 119-119-53.3-119-119-119zm37.3 58c5.5-7.6 16.2-9.3 23.7-3.7s9.3 16.2 3.7 23.7l-68 93.5c-2.9 4-7.5 6.6-12.4 6.9s-9.8-1.4-13.4-4.9L81.5 499c-6.6-6.6-6.6-17.4 0-24s17.4-6.6 24 0L134 503.4 190.3 426z"
            />
          </svg>
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
