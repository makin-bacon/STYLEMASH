import type { ParsedDocx } from '../../types/ooxml'
import { DOCX_PATHS, NS } from './constants'

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const NUMBERING_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
const NUMBERING_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'

/** A brand-new word/numbering.xml part (created at runtime by
 * mergeParagraphStyle() for the document's first-ever list style - see
 * numbering.ts#ensureNumberingXml) isn't valid OOXML on its own: the
 * package's [Content_Types].xml must declare its content type, and
 * word/document.xml.rels must relate document.xml to it, or Word won't even
 * look at the part. A document that already had a numbering.xml (i.e. had
 * lists before StyleMash touched it) already has both - this is a no-op
 * for those, detected by checking each file's raw text for the part name
 * rather than unconditionally appending, so re-saving never accumulates
 * duplicate entries. */
async function ensureNumberingPartRegistered(zip: ParsedDocx['zip']): Promise<void> {
  const contentTypesEntry = zip.file(DOCX_PATHS.contentTypes)
  if (contentTypesEntry) {
    const text = await contentTypesEntry.async('text')
    if (!text.includes('/word/numbering.xml')) {
      const doc = new DOMParser().parseFromString(text, 'application/xml')
      const override = doc.createElementNS(NS.contentTypes, 'Override')
      override.setAttribute('PartName', '/word/numbering.xml')
      override.setAttribute('ContentType', NUMBERING_CONTENT_TYPE)
      doc.documentElement.appendChild(override)
      zip.file(DOCX_PATHS.contentTypes, serializePart(doc))
    }
  }

  const relsEntry = zip.file(DOCX_PATHS.documentRels)
  if (relsEntry) {
    const text = await relsEntry.async('text')
    if (!text.includes('numbering.xml')) {
      const doc = new DOMParser().parseFromString(text, 'application/xml')
      const existingIds = new Set(
        Array.from(doc.getElementsByTagNameNS(NS.relationships, 'Relationship')).map((el) =>
          el.getAttribute('Id'),
        ),
      )
      let n = 1
      while (existingIds.has(`rId${n}`)) n++
      const rel = doc.createElementNS(NS.relationships, 'Relationship')
      rel.setAttribute('Id', `rId${n}`)
      rel.setAttribute('Type', NUMBERING_RELATIONSHIP_TYPE)
      rel.setAttribute('Target', 'numbering.xml')
      doc.documentElement.appendChild(rel)
      zip.file(DOCX_PATHS.documentRels, serializePart(doc))
    }
  } else {
    // Extremely rare/malformed input (a document.xml with no rels part at
    // all) - write a minimal one from scratch rather than silently dropping
    // the numbering relationship.
    const doc = new DOMParser().parseFromString(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.relationships}"></Relationships>`,
      'application/xml',
    )
    const rel = doc.createElementNS(NS.relationships, 'Relationship')
    rel.setAttribute('Id', 'rId1')
    rel.setAttribute('Type', NUMBERING_RELATIONSHIP_TYPE)
    rel.setAttribute('Target', 'numbering.xml')
    doc.documentElement.appendChild(rel)
    zip.file(DOCX_PATHS.documentRels, serializePart(doc))
  }
}

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
 * zip entry are never rewritten - StyleMash only ever touches the content
 * of parts that already existed in the uploaded file, with one exception:
 * numberingXml, which mergeParagraphStyle() may have created from nothing
 * for a document's first-ever list style (see numbering.ts#ensureNumberingXml).
 * When that's happened, [Content_Types].xml/document.xml.rels need a new
 * entry too, or the part StyleMash just wrote wouldn't actually be part of
 * the package - see ensureNumberingPartRegistered. */
export async function serializeDocx(
  parsedDocx: ParsedDocx,
): Promise<{ blob: Blob; filename: string }> {
  parsedDocx.zip.file(DOCX_PATHS.document, serializePart(parsedDocx.documentXml))
  parsedDocx.zip.file(DOCX_PATHS.styles, serializePart(parsedDocx.stylesXml))

  if (parsedDocx.numberingXml) {
    parsedDocx.zip.file(DOCX_PATHS.numbering, serializePart(parsedDocx.numberingXml))
    await ensureNumberingPartRegistered(parsedDocx.zip)
  }

  const blob = await parsedDocx.zip.generateAsync({
    type: 'blob',
    mimeType: MIME_TYPES[parsedDocx.originalExtension],
  })

  return {
    blob,
    filename: buildRippedFilename(parsedDocx.originalFilename, parsedDocx.originalExtension),
  }
}
