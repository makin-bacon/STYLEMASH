import { describe, expect, it } from 'vitest'
import type { UserStyleRecord } from '../src/types/ooxml'
import { NS } from '../src/lib/ooxml/constants'
import { wAttr } from '../src/lib/ooxml/domUtils'
import {
  materializeReferenceDocStyles,
  reconcileUserStylesOnReferenceDocRemoval,
} from '../src/lib/ooxml/referenceDocStyles'
import { buildStyleReport } from '../src/lib/ooxml/styleReport'
import { makeParsedDocx } from './testUtils'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

function styleCount(stylesXml: XMLDocument): number {
  return stylesXml.getElementsByTagNameNS(NS.w, 'style').length
}

describe('materializeReferenceDocStyles', () => {
  it('materializes every defined paragraph/character style, used in body text or not, always as character styles when they carry no list numbering', () => {
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:rStyle w:val="Emph"/></w:rPr><w:t>one</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="character" w:styleId="Emph"><w:name w:val="Emph"/><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:rPr><w:i/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Unused"><w:name w:val="Unused"/><w:rPr><w:strike/></w:rPr></w:style>
      </w:styles>`,
    })
    const targetDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`,
    })

    const records = materializeReferenceDocStyles(targetDocx, referenceDocx, [])

    // All three defined styles show up, including "Unused" which never
    // appears in Document B's own body text.
    expect(records).toHaveLength(3)
    expect(records.every((r) => r.fromReferenceDoc)).toBe(true)
    expect(records.map((r) => r.name).sort()).toEqual(['Emph', 'Heading 1', 'Unused'])

    const emphRecord = records.find((r) => r.name === 'Emph')!
    const headingRecord = records.find((r) => r.name === 'Heading 1')!
    const unusedRecord = records.find((r) => r.name === 'Unused')!
    expect(emphRecord.targetSignature.bold).toBe(true)
    expect(headingRecord.targetSignature.italic).toBe(true)
    expect(unusedRecord.targetSignature.strike).toBe(true)

    // Materialized styles are always character-type, even the one sourced
    // from Document B's paragraph style.
    const styleEls = Array.from(targetDocx.stylesXml.getElementsByTagNameNS(NS.w, 'style'))
    expect(styleEls).toHaveLength(3)
    for (const el of styleEls) {
      expect(wAttr(el, 'type')).toBe('character')
    }
  })

  it('materializes a Document B paragraph style that carries its own list numbering as a matching paragraph-kind list style', () => {
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Bulleted item</w:t></w:r></w:p>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="paragraph" w:styleId="ListBullet">
          <w:name w:val="List Bullet"/>
          <w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>
        </w:style>
      </w:styles>`,
      numberingXml: `<w:numbering ${W}>
        <w:abstractNum w:abstractNumId="0">
          <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      </w:numbering>`,
    })
    const targetDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`,
    })
    expect(targetDocx.numberingXml).toBeNull()

    const records = materializeReferenceDocStyles(targetDocx, referenceDocx, [])

    expect(records).toHaveLength(1)
    expect(records[0].kind).toBe('paragraph')
    expect(records[0].listFormat).toBe('bullet')
    expect(records[0].listPreviewText).toBe('•')

    const styleEls = Array.from(targetDocx.stylesXml.getElementsByTagNameNS(NS.w, 'style'))
    expect(styleEls).toHaveLength(1)
    expect(wAttr(styleEls[0], 'type')).toBe('paragraph')

    // A brand-new numbering.xml was created in Document A with its own
    // bullet definition - not a copy of Document B's numId=1 (which
    // wouldn't exist as a real part in Document A's package).
    expect(targetDocx.numberingXml).not.toBeNull()
    const abstractEls = targetDocx.numberingXml!.getElementsByTagNameNS(NS.w, 'abstractNum')
    expect(abstractEls).toHaveLength(1)
    expect(wAttr(abstractEls[0].getElementsByTagNameNS(NS.w, 'lvl')[0], 'ilvl')).toBe('0')
  })

  it('captures a multilevel heading style\'s true depth in listPreviewText (Heading 3 one level deeper than Heading 2)', () => {
    // A single multilevel numbering definition shared by three heading
    // styles, each pinned to a different level - the standard Word
    // "Heading 1/2/3 -> 1./1.1./1.1.1." outline convention.
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Intro</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Background</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Details</w:t></w:r></w:p>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="paragraph" w:styleId="Heading1">
          <w:name w:val="Heading 1"/>
          <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
        </w:style>
        <w:style w:type="paragraph" w:styleId="Heading2">
          <w:name w:val="Heading 2"/>
          <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>
        </w:style>
        <w:style w:type="paragraph" w:styleId="Heading3">
          <w:name w:val="Heading 3"/>
          <w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr>
        </w:style>
      </w:styles>`,
      numberingXml: `<w:numbering ${W}>
        <w:abstractNum w:abstractNumId="0">
          <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
          <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl>
          <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3."/></w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      </w:numbering>`,
    })
    const targetDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`,
    })

    const records = materializeReferenceDocStyles(targetDocx, referenceDocx, [])
    const byName = (name: string) => records.find((r) => r.name === name)!

    expect(byName('Heading 1').listPreviewText).toBe('1.')
    expect(byName('Heading 2').listPreviewText).toBe('1.1.')
    expect(byName('Heading 3').listPreviewText).toBe('1.1.1.')
  })

  it('avoids clobbering a pre-existing style of the same name in Document A', () => {
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:rStyle w:val="Emph"/></w:rPr><w:t>one</w:t></w:r></w:p>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="character" w:styleId="Emph"><w:name w:val="Emph"/><w:rPr><w:b/></w:rPr></w:style>
      </w:styles>`,
    })
    const targetDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="character" w:styleId="Emph"><w:name w:val="Emph"/><w:rPr><w:i/></w:rPr></w:style>
      </w:styles>`,
    })

    const records = materializeReferenceDocStyles(targetDocx, referenceDocx, [])

    expect(records).toHaveLength(1)
    expect(records[0].styleId).not.toBe('Emph') // collision -> suffixed id via generateUniqueStyleId

    const stylesRoot = targetDocx.stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
    const styleEls = Array.from(stylesRoot.getElementsByTagNameNS(NS.w, 'style'))
    expect(styleEls).toHaveLength(2)

    const originalEmph = styleEls.find((el) => wAttr(el, 'styleId') === 'Emph')!
    expect(originalEmph.getElementsByTagNameNS(NS.w, 'i')).toHaveLength(1)
    expect(originalEmph.getElementsByTagNameNS(NS.w, 'b')).toHaveLength(0) // untouched by the materialize step
  })

  it('resolves a materialized style\'s full basedOn/docDefaults cascade, not just its own direct rPr', () => {
    const referenceDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:rStyle w:val="Emph"/></w:rPr><w:t>one</w:t></w:r></w:p>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="character" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Emph"><w:name w:val="Emph"/><w:basedOn w:val="Base"/><w:rPr><w:b/></w:rPr></w:style>
      </w:styles>`,
    })
    const targetDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`,
    })

    const records = materializeReferenceDocStyles(targetDocx, referenceDocx, [])

    // "Base" is never directly applied to a run in Document B (Emph only
    // reaches it via basedOn), but it's still a defined style, so it's
    // materialized too, alongside Emph.
    expect(records).toHaveLength(2)
    const emphRecord = records.find((r) => r.name === 'Emph')!
    const baseRecord = records.find((r) => r.name === 'Base')!

    expect(emphRecord.targetSignature.fontFamily).toBe('Georgia')
    expect(emphRecord.targetSignature.fontSizeHalfPt).toBe(20)
    expect(emphRecord.targetSignature.underline).toBe('single')
    expect(emphRecord.targetSignature.bold).toBe(true)

    // Base's own cascade: docDefaults' font/size plus its own direct
    // underline, but none of Emph's bold (the cascade only flows
    // basedOn -> derived style, never the other way).
    expect(baseRecord.targetSignature.fontFamily).toBe('Georgia')
    expect(baseRecord.targetSignature.fontSizeHalfPt).toBe(20)
    expect(baseRecord.targetSignature.underline).toBe('single')
    expect(baseRecord.targetSignature.bold).toBe(false)
  })

  it('replaces an existing user-created style of the same name in place, rather than adding a duplicate', () => {
    // Simulates "attach Document B1 (defines 'Heading 1', italic), remove
    // it, attach Document B2 (also defines 'Heading 1', but bold)" -
    // targetDocx already carries the style B1 left behind (as
    // materializeReferenceDocStyles itself would have written it), plus a
    // run already merged into it.
    const targetDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:rStyle w:val="Heading1"/></w:rPr><w:t>Already merged</w:t></w:r></w:p>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="character" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:rPr><w:i/></w:rPr></w:style>
      </w:styles>`,
    })
    const existingUserStyles: UserStyleRecord[] = [
      {
        styleId: 'Heading1',
        name: 'Heading 1',
        targetSignature: { fontFamily: null, fontSizeHalfPt: null, colorValue: 'auto', bold: false, italic: true, underline: null, strike: false },
        kind: 'character',
        listFormat: 'none',
        createdAt: 111,
        fromReferenceDoc: true,
      },
      {
        styleId: 'Manual1',
        name: 'Manual Style',
        targetSignature: { fontFamily: null, fontSizeHalfPt: null, colorValue: 'auto', bold: false, italic: false, underline: null, strike: false },
        kind: 'character',
        listFormat: 'none',
        createdAt: 222,
      },
    ]

    const referenceDocxB2 = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="character" w:styleId="H1"><w:name w:val="Heading 1"/><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="NewOne"><w:name w:val="New One"/><w:rPr><w:strike/></w:rPr></w:style>
      </w:styles>`,
    })

    const result = materializeReferenceDocStyles(targetDocx, referenceDocxB2, existingUserStyles)

    // No duplicate "Heading 1" - the old record's styleId/position/createdAt
    // are kept, just redefined to B2's look; "Manual Style" is untouched;
    // "New One" is appended as a genuinely new entry.
    expect(result.map((r) => r.name)).toEqual(['Heading 1', 'Manual Style', 'New One'])
    const heading = result.find((r) => r.name === 'Heading 1')!
    expect(heading.styleId).toBe('Heading1')
    expect(heading.createdAt).toBe(111)
    expect(heading.targetSignature.bold).toBe(true)
    expect(heading.targetSignature.italic).toBe(false)
    const manual = result.find((r) => r.name === 'Manual Style')!
    expect(manual).toEqual(existingUserStyles[1])

    // Exactly one <w:style w:styleId="Heading1"> in Document A, redefined in
    // place - not a second style of the same name under a new id.
    const stylesRoot = targetDocx.stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
    const heading1Els = Array.from(stylesRoot.getElementsByTagNameNS(NS.w, 'style')).filter(
      (el) => wAttr(el, 'styleId') === 'Heading1',
    )
    expect(heading1Els).toHaveLength(1)
    expect(heading1Els[0].getElementsByTagNameNS(NS.w, 'b')).toHaveLength(1)
    expect(heading1Els[0].getElementsByTagNameNS(NS.w, 'i')).toHaveLength(0)

    // The run merged into "Heading1" before the swap still points at the
    // same id, so it automatically picks up B2's bold look.
    const styleReport = buildStyleReport(targetDocx)
    const rStyleEls = targetDocx.documentXml.getElementsByTagNameNS(NS.w, 'rStyle')
    expect(rStyleEls).toHaveLength(1)
    expect(wAttr(rStyleEls[0], 'val')).toBe('Heading1')
    expect(styleReport.some((e) => e.signature.bold === true)).toBe(true)
  })
})

describe('reconcileUserStylesOnReferenceDocRemoval', () => {
  it('drops unused imported styles entirely but demotes used ones to manual', () => {
    const parsedDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:rStyle w:val="Used1"/></w:rPr><w:t>one</w:t></w:r></w:p>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="character" w:styleId="Used1"><w:name w:val="Used1"/><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Unused1"><w:name w:val="Unused1"/><w:rPr><w:i/></w:rPr></w:style>
      </w:styles>`,
    })
    const styleReport = buildStyleReport(parsedDocx)

    const records: UserStyleRecord[] = [
      {
        styleId: 'Used1',
        name: 'Used1',
        targetSignature: { fontFamily: null, fontSizeHalfPt: null, colorValue: 'auto', bold: true, italic: false, underline: null, strike: false },
        createdAt: 1,
        fromReferenceDoc: true,
      },
      {
        styleId: 'Unused1',
        name: 'Unused1',
        targetSignature: { fontFamily: null, fontSizeHalfPt: null, colorValue: 'auto', bold: false, italic: true, underline: null, strike: false },
        createdAt: 2,
        fromReferenceDoc: true,
      },
      {
        styleId: 'Manual1',
        name: 'Manual1',
        targetSignature: { fontFamily: null, fontSizeHalfPt: null, colorValue: 'auto', bold: false, italic: false, underline: null, strike: false },
        createdAt: 3,
      },
    ]

    const result = reconcileUserStylesOnReferenceDocRemoval(records, styleReport, parsedDocx.stylesXml)

    expect(result.map((r) => r.styleId).sort()).toEqual(['Manual1', 'Used1'])
    const usedResult = result.find((r) => r.styleId === 'Used1')!
    expect(usedResult.fromReferenceDoc).toBeUndefined()
    const manualResult = result.find((r) => r.styleId === 'Manual1')!
    expect(manualResult.fromReferenceDoc).toBeUndefined()

    expect(styleCount(parsedDocx.stylesXml)).toBe(1) // Unused1's <w:style> removed, Used1's kept
    expect(parsedDocx.stylesXml.getElementsByTagNameNS(NS.w, 'style')[0].getAttributeNS(NS.w, 'styleId')).toBe(
      'Used1',
    )
  })
})
