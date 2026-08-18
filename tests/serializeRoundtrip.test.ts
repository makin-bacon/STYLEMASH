import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import type { FormattingSignature } from '../src/types/ooxml'
import { mergeParagraphStyle } from '../src/lib/ooxml/mergeStyles'
import { buildParagraphMarkers } from '../src/lib/ooxml/numbering'
import { parseDocx } from '../src/lib/ooxml/parseDocx'
import { serializeDocx } from '../src/lib/ooxml/serializeDocx'
import { buildStyleReport } from '../src/lib/ooxml/styleReport'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const NEUTRAL_SIGNATURE: FormattingSignature = {
  fontFamily: 'Arial',
  fontSizeHalfPt: 24,
  colorValue: 'auto',
  bold: false,
  italic: false,
  underline: null,
  strike: false,
}

/** Builds a minimal-but-valid .docx File in-memory (no checked-in binary
 * fixture needed) - just enough parts for parseDocx to accept it. */
async function buildMinimalDocx(bodyXml: string): Promise<File> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>${bodyXml}</w:body></w:document>`,
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}></w:styles>`,
  )

  const blob = await zip.generateAsync({ type: 'blob' })
  return new File([blob], 'fixture.docx', {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  })
}

describe('parseDocx -> serializeDocx round trip', () => {
  it('preserves run count and text when serialized without modification', async () => {
    const file = await buildMinimalDocx('<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>')
    const parsedDocx = await parseDocx(file)
    const { blob, filename } = await serializeDocx(parsedDocx)

    expect(filename).toBe('fixture-RIPPED.docx')

    const reopenedZip = await JSZip.loadAsync(blob)
    const reDocumentXml = await reopenedZip.file('word/document.xml')!.async('text')

    expect(reDocumentXml).toContain('Hello world')
    expect(reDocumentXml.match(/<w:r>/g)?.length).toBe(1)
    // Regression guard for the XMLSerializer declaration gotcha (serializeDocx.ts).
    expect(reDocumentXml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(
      true,
    )
  })

  it('registers a brand-new word/numbering.xml part (Content_Types + rels) when a list style creates one', async () => {
    const file = await buildMinimalDocx('<w:p><w:r><w:t>First item</w:t></w:r></w:p>')
    const parsedDocx = await parseDocx(file)
    expect(parsedDocx.numberingXml).toBeNull() // fixture ships no numbering.xml at all

    const allRunRefs = buildStyleReport(parsedDocx).flatMap((e) => e.variants.flatMap((v) => v.runRefs))
    mergeParagraphStyle(parsedDocx, allRunRefs, NEUTRAL_SIGNATURE, 'List Style', 'decimal')
    expect(parsedDocx.numberingXml).not.toBeNull() // created in memory by the merge

    const { blob } = await serializeDocx(parsedDocx)
    const reopenedZip = await JSZip.loadAsync(blob)

    const numberingXml = await reopenedZip.file('word/numbering.xml')?.async('text')
    expect(numberingXml).toContain('abstractNum')
    expect(numberingXml).toContain('<w:numFmt w:val="decimal"')

    const contentTypes = await reopenedZip.file('[Content_Types].xml')!.async('text')
    expect(contentTypes).toContain('/word/numbering.xml')

    const rels = await reopenedZip.file('word/_rels/document.xml.rels')!.async('text')
    expect(rels).toContain('numbering.xml')
    expect(rels).toContain('relationships/numbering')

    // Full fidelity check: re-parsing the saved file from scratch resolves
    // the same "1." marker a fresh upload of this file would show.
    const reopenedFile = new File([blob], 'fixture-RIPPED.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    })
    const reparsed = await parseDocx(reopenedFile)
    const markers = buildParagraphMarkers(reparsed)
    const reparsedParagraph = reparsed.documentXml.getElementsByTagNameNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'p',
    )[0]
    expect(markers.get(reparsedParagraph)?.text).toBe('1.')
  })

  it('does not duplicate the Content_Types/rels entries for a document that already has numbering.xml', async () => {
    // A fixture that already ships a complete, properly-registered
    // numbering.xml - i.e. a real document that already had lists before
    // StyleMash ever touched it - to verify ensureNumberingPartRegistered
    // is a true no-op here rather than appending a redundant second entry.
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`,
    )
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
    )
    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`,
    )
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`)
    zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles ${W}></w:styles>`)
    zip.file('word/numbering.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering ${W}></w:numbering>`)

    const blob0 = await zip.generateAsync({ type: 'blob' })
    const file = new File([blob0], 'fixture.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    })
    const parsedDocx = await parseDocx(file)
    expect(parsedDocx.numberingXml).not.toBeNull()

    const { blob } = await serializeDocx(parsedDocx)
    const reopenedZip = await JSZip.loadAsync(blob)
    const contentTypes = await reopenedZip.file('[Content_Types].xml')!.async('text')
    const rels = await reopenedZip.file('word/_rels/document.xml.rels')!.async('text')

    expect((contentTypes.match(/numbering\.xml/g) ?? []).length).toBe(1)
    expect((rels.match(/<Relationship /g) ?? []).length).toBe(1)
  })

  it('rejects files that are not valid ZIP/OOXML packages', async () => {
    const notADocx = new File(['plain text, not a zip'], 'notes.docx', { type: 'text/plain' })
    await expect(parseDocx(notADocx)).rejects.toThrow()
  })

  it('rejects unsupported extensions', async () => {
    const legacyDoc = new File(['irrelevant'], 'old.doc', { type: 'application/msword' })
    await expect(parseDocx(legacyDoc)).rejects.toThrow()
  })
})
