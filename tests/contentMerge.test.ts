import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { NS } from '../src/lib/ooxml/constants'
import {
  buildContentMergedDocx,
  buildContentMergedFilename,
  transplantContent,
} from '../src/lib/ooxml/contentMerge'
import { wChild } from '../src/lib/ooxml/domUtils'
import { parseDocx } from '../src/lib/ooxml/parseDocx'
import { makeParsedDocx } from './testUtils'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

function bodyOf(documentXml: XMLDocument): Element {
  return documentXml.getElementsByTagNameNS(NS.w, 'body')[0]
}

function bodyText(documentXml: XMLDocument): string {
  return Array.from(documentXml.getElementsByTagNameNS(NS.w, 't'))
    .map((t) => t.textContent ?? '')
    .join('')
}

describe('transplantContent', () => {
  it('replaces the body with A\'s content while keeping B\'s sectPr as the last child, unchanged', () => {
    const sourceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:t>From A</w:t></w:r></w:p>
        <w:sectPr><w:pgMar w:top="999"/></w:sectPr>
      </w:body></w:document>`,
    })
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:t>From B</w:t></w:r></w:p>
        <w:sectPr><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>
      </w:body></w:document>`,
    })

    const { documentXml } = transplantContent(sourceDocx, referenceDocx, { keepOriginalFormatting: true })

    expect(bodyText(documentXml)).toBe('From A')
    expect(bodyText(documentXml)).not.toContain('From B')

    const body = bodyOf(documentXml)
    const sectPr = wChild(body, 'sectPr')!
    expect(body.lastElementChild).toBe(sectPr)
    const pgMar = sectPr.getElementsByTagNameNS(NS.w, 'pgMar')[0]
    expect(pgMar.getAttributeNS(NS.w, 'top')).toBe('1440') // B's page setup, not A's
    expect(pgMar.getAttributeNS(NS.w, 'left')).toBe('1800')
  })

  it('strips a mid-document section break from a transplanted paragraph but keeps its other pPr children', () => {
    const sourceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p>
          <w:pPr>
            <w:pStyle w:val="Heading1"/>
            <w:sectPr><w:pgMar w:top="500"/></w:sectPr>
          </w:pPr>
          <w:r><w:t>Section break paragraph</w:t></w:r>
        </w:p>
        <w:sectPr><w:pgMar w:top="999"/></w:sectPr>
      </w:body></w:document>`,
    })
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>B</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    })

    const { documentXml } = transplantContent(sourceDocx, referenceDocx, { keepOriginalFormatting: true })

    const p = bodyOf(documentXml).getElementsByTagNameNS(NS.w, 'p')[0]
    const pPr = wChild(p, 'pPr')!
    expect(wChild(pPr, 'sectPr')).toBeNull()
    expect(wChild(pPr, 'pStyle')!.getAttributeNS(NS.w, 'val')).toBe('Heading1')
  })

  it('keep-original-formatting: copies an A-only style definition into the output under its own id', () => {
    const sourceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:rStyle w:val="OnlyInA"/></w:rPr><w:t>x</w:t></w:r></w:p>
        <w:sectPr/>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}><w:style w:type="character" w:styleId="OnlyInA"><w:name w:val="OnlyInA"/><w:rPr><w:b/></w:rPr></w:style></w:styles>`,
    })
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>B</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    })

    const { documentXml, stylesXml } = transplantContent(sourceDocx, referenceDocx, {
      keepOriginalFormatting: true,
    })

    const run = documentXml.getElementsByTagNameNS(NS.w, 'r')[0]
    expect(wChild(wChild(run, 'rPr'), 'rStyle')!.getAttributeNS(NS.w, 'val')).toBe('OnlyInA')

    const copied = Array.from(stylesXml.getElementsByTagNameNS(NS.w, 'style')).find(
      (el) => el.getAttributeNS(NS.w, 'styleId') === 'OnlyInA',
    )
    expect(copied).toBeDefined()
    expect(copied!.getElementsByTagNameNS(NS.w, 'b')).toHaveLength(1)
  })

  it('keep-original-formatting: never overwrites a style id Document B already defines, even if the look differs', () => {
    const sourceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
        <w:sectPr/>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style></w:styles>`,
    })
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>B</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
      stylesXml: `<w:styles ${W}><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:i/></w:rPr></w:style></w:styles>`,
    })

    const { stylesXml } = transplantContent(sourceDocx, referenceDocx, { keepOriginalFormatting: true })

    const normalStyles = Array.from(stylesXml.getElementsByTagNameNS(NS.w, 'style')).filter(
      (el) => el.getAttributeNS(NS.w, 'styleId') === 'Normal',
    )
    expect(normalStyles).toHaveLength(1) // not duplicated
    expect(normalStyles[0].getElementsByTagNameNS(NS.w, 'i')).toHaveLength(1) // B's version survives
    expect(normalStyles[0].getElementsByTagNameNS(NS.w, 'b')).toHaveLength(0) // A's version was NOT copied in
  })

  it('snap mode: repoints a same-named style reference at Document B\'s own id instead of copying A\'s', () => {
    const sourceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:pPr><w:pStyle w:val="A-Heading1"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
        <w:sectPr/>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}><w:style w:type="paragraph" w:styleId="A-Heading1"><w:name w:val="Heading 1"/><w:rPr><w:b/></w:rPr></w:style></w:styles>`,
    })
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>B</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
      stylesXml: `<w:styles ${W}><w:style w:type="paragraph" w:styleId="B-Heading1"><w:name w:val="Heading 1"/><w:rPr><w:i/></w:rPr></w:style></w:styles>`,
    })

    const { documentXml, stylesXml } = transplantContent(sourceDocx, referenceDocx, {
      keepOriginalFormatting: false,
    })

    const p = documentXml.getElementsByTagNameNS(NS.w, 'p')[0]
    const pStyle = wChild(wChild(p, 'pPr'), 'pStyle')!
    expect(pStyle.getAttributeNS(NS.w, 'val')).toBe('B-Heading1')

    // A's own "A-Heading1" definition was NOT copied in - only B's own id is used.
    const copiedFromA = Array.from(stylesXml.getElementsByTagNameNS(NS.w, 'style')).find(
      (el) => el.getAttributeNS(NS.w, 'styleId') === 'A-Heading1',
    )
    expect(copiedFromA).toBeUndefined()
  })

  it('snap mode falls back to gap-fill when no name match exists in Document B', () => {
    const sourceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:rStyle w:val="OnlyInA"/></w:rPr><w:t>x</w:t></w:r></w:p>
        <w:sectPr/>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}><w:style w:type="character" w:styleId="OnlyInA"><w:name w:val="OnlyInA"/><w:rPr><w:b/></w:rPr></w:style></w:styles>`,
    })
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>B</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    })

    const { documentXml, stylesXml } = transplantContent(sourceDocx, referenceDocx, {
      keepOriginalFormatting: false,
    })

    const run = documentXml.getElementsByTagNameNS(NS.w, 'r')[0]
    expect(wChild(wChild(run, 'rPr'), 'rStyle')!.getAttributeNS(NS.w, 'val')).toBe('OnlyInA')
    expect(
      Array.from(stylesXml.getElementsByTagNameNS(NS.w, 'style')).some(
        (el) => el.getAttributeNS(NS.w, 'styleId') === 'OnlyInA',
      ),
    ).toBe(true)
  })

  it('gap-fills a copied style\'s own basedOn ancestor recursively, and tolerates a circular chain', () => {
    const sourceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:rStyle w:val="Custom"/></w:rPr><w:t>x</w:t></w:r></w:p>
        <w:sectPr/>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="character" w:styleId="Custom"><w:name w:val="Custom"/><w:basedOn w:val="Emph"/><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Emph"><w:name w:val="Emph"/><w:basedOn w:val="Custom"/><w:rPr><w:i/></w:rPr></w:style>
      </w:styles>`,
    })
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>B</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    })

    // Should not infinite-loop despite the circular Custom<->Emph basedOn chain.
    const { stylesXml } = transplantContent(sourceDocx, referenceDocx, { keepOriginalFormatting: true })

    const ids = Array.from(stylesXml.getElementsByTagNameNS(NS.w, 'style')).map((el) =>
      el.getAttributeNS(NS.w, 'styleId'),
    )
    expect(ids).toContain('Custom')
    expect(ids).toContain('Emph')
  })
})

describe('buildContentMergedFilename', () => {
  it('appends -MERGED before the extension', () => {
    expect(buildContentMergedFilename('Report.docx', 'docx')).toBe('Report-MERGED.docx')
  })
})

describe('buildContentMergedDocx (end-to-end)', () => {
  async function buildMinimalDocx(bodyXml: string, extraParts: Record<string, string> = {}): Promise<File> {
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
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}><w:body>${bodyXml}</w:body></w:document>`,
    )
    zip.file(
      'word/styles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles ${W}></w:styles>`,
    )
    for (const [path, content] of Object.entries(extraParts)) {
      zip.file(path, content)
    }
    const blob = await zip.generateAsync({ type: 'blob' })
    return new File([blob], 'fixture.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    })
  }

  it('produces Document B with Document A\'s text, leaves every other Document B part byte-identical, and never mutates the source zip across repeated calls', async () => {
    const fileA = await buildMinimalDocx('<w:p><w:r><w:t>Hello from A</w:t></w:r></w:p>')
    const numberingMarker = '<w:numbering xmlns:w="fake">MARKER</w:numbering>'
    const fileB = await buildMinimalDocx(
      '<w:p><w:r><w:t>Hello from B</w:t></w:r></w:p><w:sectPr><w:pgMar w:top="1440"/></w:sectPr>',
      { 'word/numbering.xml': numberingMarker },
    )

    const sourceDocx = await parseDocx(fileA)
    const referenceDocx = await parseDocx(fileB)

    const { blob, filename } = await buildContentMergedDocx(sourceDocx, referenceDocx, {
      keepOriginalFormatting: true,
    })
    expect(filename).toBe('fixture-MERGED.docx')

    const reopened = await JSZip.loadAsync(blob)
    const outDocumentXml = await reopened.file('word/document.xml')!.async('text')
    expect(outDocumentXml).toContain('Hello from A')
    expect(outDocumentXml).not.toContain('Hello from B')
    expect(outDocumentXml).toContain('w:top="1440"') // B's sectPr survived

    const outNumbering = await reopened.file('word/numbering.xml')!.async('text')
    expect(outNumbering).toBe(numberingMarker) // untouched, byte-identical

    // Regression guard: calling it again with the same ParsedDocx pair must
    // not have mutated referenceDocx.zip as a side effect of the first call
    // (see contentMerge.ts's cloneZipForOutput - avoiding JSZip#clone()'s
    // shared-.files-object footgun).
    const beforeSecondCall = await referenceDocx.zip.file('word/document.xml')!.async('text')
    expect(beforeSecondCall).toContain('Hello from B')

    const second = await buildContentMergedDocx(sourceDocx, referenceDocx, { keepOriginalFormatting: true })
    const reopenedSecond = await JSZip.loadAsync(second.blob)
    const outDocumentXmlSecond = await reopenedSecond.file('word/document.xml')!.async('text')
    expect(outDocumentXmlSecond).toContain('Hello from A')

    const afterSecondCall = await referenceDocx.zip.file('word/document.xml')!.async('text')
    expect(afterSecondCall).toContain('Hello from B') // still untouched
  })
})
