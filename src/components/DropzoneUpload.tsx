import { useRef, useState } from 'react'
import type { DragEvent } from 'react'
import { hasDocxExtension, looksLikeZip } from '../lib/ooxml/fileValidation'

interface DropzoneUploadProps {
  onFileAccepted: (file: File) => void
  errorMessage: string | null
}

/** Drag-and-drop + click-to-browse upload target. Uses plain HTML5 drag
 * events rather than the react-dropzone package - a single-file .docx/.dotx
 * upload is simple enough that the extra dependency isn't warranted. Does a
 * cheap ZIP magic-number sniff before handing off to JSZip/DOMParser, so a
 * renamed non-docx file (or a real legacy .doc) fails fast with a clear
 * message instead of a raw parser exception. */
export function DropzoneUpload({ onFileAccepted, errorMessage }: DropzoneUploadProps) {
  const [isDragActive, setIsDragActive] = useState(false)
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
      setLocalError(
        "This file doesn't look like a valid Office Open XML package - it may be corrupted, or an older (pre-2007) .doc file, which isn't supported.",
      )
      return
    }

    onFileAccepted(file)
  }

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    setIsDragActive(false)
    void handleFile(e.dataTransfer.files[0])
  }

  return (
    <div className="flex flex-1 items-center justify-center p-8">
      <div
        onDragOver={(e) => {
          e.preventDefault()
          setIsDragActive(true)
        }}
        onDragLeave={() => setIsDragActive(false)}
        onDrop={handleDrop}
        onClick={() => inputRef.current?.click()}
        role="button"
        tabIndex={0}
        className={`flex w-full max-w-lg cursor-pointer flex-col items-center gap-3 rounded-xl border-2 border-dashed p-12 text-center transition-colors ${
          isDragActive ? 'border-indigo-500 bg-indigo-50' : 'border-slate-300 bg-white hover:border-slate-400'
        }`}
      >
        <p className="text-lg font-semibold text-slate-700">Drop a Word document here</p>
        <p className="text-sm text-slate-500">or click to browse - .docx / .dotx only</p>
        <input
          ref={inputRef}
          type="file"
          accept=".docx,.dotx"
          className="hidden"
          onChange={(e) => void handleFile(e.target.files?.[0])}
        />
        {(localError || errorMessage) && (
          <p className="mt-2 text-sm text-red-600">{localError || errorMessage}</p>
        )}
      </div>
    </div>
  )
}
