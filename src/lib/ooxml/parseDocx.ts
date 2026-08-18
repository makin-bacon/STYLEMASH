import JSZip from 'jszip'
import type { ParsedDocx } from '../../types/ooxml'
import { DOCX_PATHS } from './constants'
import { hasParseError } from './domUtils'

/** A minimal, valid <w:styles> document to fall back to for the rare
 * malformed .docx that omits word/styles.xml entirely, so resolution code
 * never has to null-check the styles document itself. */
const EMPTY_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`

export class DocxParseError extends Error {}

function parseXml(text: string, partName: string): XMLDocument {
  const doc = new DOMParser().parseFromString(text, 'application/xml')
  if (hasParseError(doc)) {
    throw new DocxParseError(`"${partName}" is not valid XML - the file may be corrupted.`)
  }
  return doc
}

/** Loads a .docx/.dotx file into an in-memory ParsedDocx. Everything happens
 * client-side via JSZip + DOMParser - the file never leaves the browser. */
export async function parseDocx(file: File): Promise<ParsedDocx> {
  const lowerName = file.name.toLowerCase()
  const originalExtension: 'docx' | 'dotx' = lowerName.endsWith('.dotx') ? 'dotx' : 'docx'
  if (!lowerName.endsWith('.docx') && !lowerName.endsWith('.dotx')) {
    throw new DocxParseError('StyleMash only supports .docx and .dotx files (Office Open XML).')
  }

  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(file)
  } catch {
    throw new DocxParseError(
      'Could not open this file as a .docx/.dotx package - it may be corrupted or an older (pre-2007) .doc file, which uses a different, unsupported format.',
    )
  }

  const documentEntry = zip.file(DOCX_PATHS.document)
  if (!documentEntry) {
    throw new DocxParseError(
      `This file is missing ${DOCX_PATHS.document} and isn't a valid Word document.`,
    )
  }
  const documentXml = parseXml(await documentEntry.async('text'), DOCX_PATHS.document)

  const stylesEntry = zip.file(DOCX_PATHS.styles)
  const stylesXml = parseXml(
    stylesEntry ? await stylesEntry.async('text') : EMPTY_STYLES_XML,
    DOCX_PATHS.styles,
  )

  const themeEntry = zip.file(DOCX_PATHS.theme)
  const themeXml = themeEntry ? parseXml(await themeEntry.async('text'), DOCX_PATHS.theme) : null

  const numberingEntry = zip.file(DOCX_PATHS.numbering)
  const numberingXml = numberingEntry
    ? parseXml(await numberingEntry.async('text'), DOCX_PATHS.numbering)
    : null

  return {
    zip,
    documentXml,
    stylesXml,
    themeXml,
    numberingXml,
    originalFilename: file.name,
    originalExtension,
  }
}
