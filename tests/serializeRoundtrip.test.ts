import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/lib/ooxml/parseDocx'
import { serializeDocx } from '../src/lib/ooxml/serializeDocx'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

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

  it('rejects files that are not valid ZIP/OOXML packages', async () => {
    const notADocx = new File(['plain text, not a zip'], 'notes.docx', { type: 'text/plain' })
    await expect(parseDocx(notADocx)).rejects.toThrow()
  })

  it('rejects unsupported extensions', async () => {
    const legacyDoc = new File(['irrelevant'], 'old.doc', { type: 'application/msword' })
    await expect(parseDocx(legacyDoc)).rejects.toThrow()
  })
})
