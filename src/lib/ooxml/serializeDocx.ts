import type { ParsedDocx } from '../../types/ooxml'
import { DOCX_PATHS } from './constants'

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

export const MIME_TYPES: Record<ParsedDocx['originalExtension'], string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
}

/** XMLSerializer's declaration output isn't guaranteed to match what Word
 * originally wrote (encoding casing, standalone attr, etc). To avoid any
 * risk of Word flagging the file for repair, strip whatever it emits and
 * prepend the canonical Word-style declaration instead. */
export function serializePart(doc: XMLDocument): string {
  const serialized = new XMLSerializer().serializeToString(doc)
  const withoutDeclaration = serialized.replace(/^<\?xml[^>]*\?>\s*/i, '')
  return XML_DECLARATION + withoutDeclaration
}

export function buildRippedFilename(
  originalFilename: string,
  originalExtension: ParsedDocx['originalExtension'],
): string {
  const dotIndex = originalFilename.lastIndexOf('.')
  const baseName = dotIndex === -1 ? originalFilename : originalFilename.slice(0, dotIndex)
  return `${baseName}-RIPPED.${originalExtension}`
}

/** Writes the (possibly mutated) document.xml/styles.xml back into the
 * original zip and produces a downloadable Blob. theme1.xml and every other
 * zip entry are never rewritten - StyleRipper only ever touches the content
 * of two parts that already existed in the uploaded file, so no changes to
 * [Content_Types].xml or _rels are ever needed. */
export async function serializeDocx(
  parsedDocx: ParsedDocx,
): Promise<{ blob: Blob; filename: string }> {
  parsedDocx.zip.file(DOCX_PATHS.document, serializePart(parsedDocx.documentXml))
  parsedDocx.zip.file(DOCX_PATHS.styles, serializePart(parsedDocx.stylesXml))

  const blob = await parsedDocx.zip.generateAsync({
    type: 'blob',
    mimeType: MIME_TYPES[parsedDocx.originalExtension],
  })

  return {
    blob,
    filename: buildRippedFilename(parsedDocx.originalFilename, parsedDocx.originalExtension),
  }
}
