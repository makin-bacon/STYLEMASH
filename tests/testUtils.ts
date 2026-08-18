import JSZip from 'jszip'
import type { ParsedDocx } from '../src/types/ooxml'

const EMPTY_STYLES = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`

export function parseXmlString(xml: string): XMLDocument {
  return new DOMParser().parseFromString(xml, 'application/xml')
}

/** Builds a minimal ParsedDocx from raw XML strings for unit tests - no real
 * .docx file needed since parseDocx() itself isn't what's under test here. */
export function makeParsedDocx(opts: {
  documentXml: string
  stylesXml?: string
  themeXml?: string
  numberingXml?: string
}): ParsedDocx {
  return {
    zip: new JSZip(),
    documentXml: parseXmlString(opts.documentXml),
    stylesXml: parseXmlString(opts.stylesXml ?? EMPTY_STYLES),
    themeXml: opts.themeXml ? parseXmlString(opts.themeXml) : null,
    numberingXml: opts.numberingXml ? parseXmlString(opts.numberingXml) : null,
    originalFilename: 'test.docx',
    originalExtension: 'docx',
  }
}
