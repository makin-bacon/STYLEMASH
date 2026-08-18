import { useRef, useState } from 'react'
import type { ReferenceDocState } from '../hooks/useDocxWorkspace'
import { hasDocxExtension, looksLikeZip } from '../lib/ooxml/fileValidation'

interface AttachReferenceDocButtonProps {
  status: ReferenceDocState['status']
  errorMessage: string | null
  onAttach: (file: File) => void
}

/** Trigger for attaching Document B, rendered in UserStylesPanel's footer.
 * Only rendered while nothing is attached yet; once Document B is loaded,
 * that same footer slot shows a "Remove Document B" button instead (see
 * UserStylesPanel), and DocumentPreviewPanel's own footer takes over
 * showing Document B's info plus "Merge content into Document B…". Owns its
 * own file input + the same extension/zip-magic-byte validation every
 * upload surface in this app uses (see fileValidation.ts). */
export function AttachReferenceDocButton({ status, errorMessage, onAttach }: AttachReferenceDocButtonProps) {
  const [localError, setLocalError] = useState<string | null>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  const handleFile = async (file: File | undefined | null) => {
    if (!file) return
    setLocalError(null)

    if (!hasDocxExtension(file.name)) {
      setLocalError('Only .docx and .dotx files are supported.')
      return
    }
    if (!(await looksLikeZip(file))) {
      setLocalError("This file doesn't look like a valid Office Open XML package.")
      return
    }
    onAttach(file)
  }

  const message = localError ?? errorMessage

  return (
    <div>
      <button
        type="button"
        onClick={() => inputRef.current?.click()}
        disabled={status === 'loading'}
        className="w-full rounded-md bg-indigo-600 px-3 py-1.5 text-xs font-medium text-white enabled:hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300"
      >
        {status === 'loading' ? 'Reading reference document…' : 'Attach Document B (optional)'}
      </button>
      <input
        ref={inputRef}
        type="file"
        accept=".docx,.dotx"
        className="hidden"
        onChange={(e) => void handleFile(e.target.files?.[0])}
      />
      {message && <p className="mt-1 text-xs text-red-600">{message}</p>}
    </div>
  )
}
