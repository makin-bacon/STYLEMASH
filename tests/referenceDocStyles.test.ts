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
  it('materializes only the styles actually used in body text, always as character styles', () => {
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

    const records = materializeReferenceDocStyles(targetDocx, referenceDocx)

    expect(records).toHaveLength(2)
    expect(records.every((r) => r.fromReferenceDoc)).toBe(true)
    expect(records.map((r) => r.name).sort()).toEqual(['Emph', 'Heading 1'])

    const emphRecord = records.find((r) => r.name === 'Emph')!
    const headingRecord = records.find((r) => r.name === 'Heading 1')!
    expect(emphRecord.targetSignature.bold).toBe(true)
    expect(headingRecord.targetSignature.italic).toBe(true)

    // Materialized styles are always character-type, even the one sourced
    // from Document B's paragraph style.
    const styleEls = Array.from(targetDocx.stylesXml.getElementsByTagNameNS(NS.w, 'style'))
    expect(styleEls).toHaveLength(2)
    for (const el of styleEls) {
      expect(wAttr(el, 'type')).toBe('character')
    }
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

    const records = materializeReferenceDocStyles(targetDocx, referenceDocx)

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

    const records = materializeReferenceDocStyles(targetDocx, referenceDocx)

    expect(records).toHaveLength(1)
    expect(records[0].targetSignature.fontFamily).toBe('Georgia')
    expect(records[0].targetSignature.fontSizeHalfPt).toBe(20)
    expect(records[0].targetSignature.underline).toBe('single')
    expect(records[0].targetSignature.bold).toBe(true)
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
