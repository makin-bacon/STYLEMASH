// The first 4 bytes of any ZIP file (docx/dotx are ZIP containers) - "PK\x03\x04".
const ZIP_MAGIC_BYTES = [0x50, 0x4b, 0x03, 0x04]

export function hasDocxExtension(filename: string): boolean {
  const lower = filename.toLowerCase()
  return lower.endsWith('.docx') || lower.endsWith('.dotx')
}

/** Cheap ZIP magic-number sniff so a renamed non-docx file (or a real
 * legacy .doc) fails fast with a clear message instead of a raw parser
 * exception - shared by every upload surface (main document, Document B). */
export async function looksLikeZip(file: File): Promise<boolean> {
  const header = new Uint8Array(await file.slice(0, 4).arrayBuffer())
  return ZIP_MAGIC_BYTES.every((byte, i) => header[i] === byte)
}
