import { useRef, useState } from 'react'
import type { ReferenceDocState } from '../hooks/useDocxWorkspace'
import { hasDocxExtension, looksLikeZip } from '../lib/ooxml/fileValidation'

interface ReferenceDocPanelProps {
  referenceDoc: ReferenceDocState
  importedStyleCount: number
  isMergingContent: boolean
  onAttach: (file: File) => void
  onRemove: () => void
  onOpenContentMerge: () => void
}

/** Slim bar for attaching, viewing, and removing Document B - a second,
 * separate document used purely as a text style reference (see
 * referenceDocStyles.ts). Rendered between AppHeader and the main
 * Style-Report/User-Styles grid. Mirrors DropzoneUpload's own extension +
 * zip-magic-byte validation via the shared fileValidation.ts helpers. */
export function ReferenceDocPanel({
  referenceDoc,
  importedStyleCount,
  isMergingContent,
  onAttach,
  onRemove,
  onOpenContentMerge,
}: ReferenceDocPanelProps) {
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

  const errorMessage = localError ?? referenceDoc.errorMessage

  return (
    <div className="flex items-center justify-between gap-3 border-b border-slate-200 bg-white px-6 py-2">
      <div className="min-w-0 flex-1">
        <p className="text-xs font-semibold uppercase tracking-wide text-slate-400">
          Style reference (Document B – Optional)
        </p>
        {referenceDoc.status === 'empty' && (
          <p className="mt-0.5 text-sm text-slate-500">
            Attach a second document to use its styles as merge targets.
          </p>
        )}
        {referenceDoc.status === 'loading' && (
          <p className="mt-0.5 text-sm text-slate-500">Reading reference document…</p>
        )}
        {referenceDoc.status === 'loaded' && (
          <p className="mt-0.5 truncate text-sm text-slate-700">
            <span className="font-medium">{referenceDoc.parsedDocx?.originalFilename}</span>
            <span className="text-slate-400"> · {importedStyleCount} style{importedStyleCount === 1 ? '' : 's'} imported</span>
          </p>
        )}
        {errorMessage && <p className="mt-0.5 text-sm text-red-600">{errorMessage}</p>}
      </div>

      <div className="flex shrink-0 items-center gap-2">
        {referenceDoc.status === 'loaded' ? (
          <>
            <button
              type="button"
              onClick={onOpenContentMerge}
              disabled={isMergingContent}
              className="rounded-md bg-indigo-600 px-3 py-1.5 text-xs font-medium text-white enabled:hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300"
            >
              Merge content into Document B…
            </button>
            <button
              type="button"
              onClick={onRemove}
              className="rounded-md border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
            >
              Remove
            </button>
          </>
        ) : (
          <>
            <button
              type="button"
              onClick={() => inputRef.current?.click()}
              disabled={referenceDoc.status === 'loading'}
              className="rounded-md border border-indigo-200 px-3 py-1.5 text-xs font-medium text-indigo-600 hover:bg-indigo-50 disabled:cursor-not-allowed disabled:opacity-50"
            >
              Attach Document B
            </button>
            <input
              ref={inputRef}
              type="file"
              accept=".docx,.dotx"
              className="hidden"
              onChange={(e) => void handleFile(e.target.files?.[0])}
            />
          </>
        )}
      </div>
    </div>
  )
}
