/** Triggers a native browser download of `blob` as `filename`. Deliberately
 * not using the `file-saver` package - this is the same handful of lines it
 * runs internally on any modern browser; its extra code paths only exist
 * for legacy IE/Edge, which this local dev-server tool doesn't need. */
export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = filename
  document.body.appendChild(anchor)
  anchor.click()
  document.body.removeChild(anchor)
  URL.revokeObjectURL(url)
}
