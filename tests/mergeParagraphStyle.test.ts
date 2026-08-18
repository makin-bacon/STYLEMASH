import { describe, expect, it } from 'vitest'
import type { FormattingSignature } from '../src/types/ooxml'
import { NS } from '../src/lib/ooxml/constants'
import { wChild } from '../src/lib/ooxml/domUtils'
import { mergeParagraphStyle } from '../src/lib/ooxml/mergeStyles'
import { buildParagraphMarkers } from '../src/lib/ooxml/numbering'
import { buildStyleReport, collectRunRefsForVariantIds, countOccurrencesForStyleId } from '../src/lib/ooxml/styleReport'
import { makeParsedDocx } from './testUtils'
import type { RunRef, StyleEntity } from '../src/types/ooxml'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

function allRunRefs(report: StyleEntity[]): RunRef[] {
  return report.flatMap((entity) => entity.variants.flatMap((variant) => variant.runRefs))
}

const NEUTRAL_SIGNATURE: FormattingSignature = {
  fontFamily: 'Arial',
  fontSizeHalfPt: 24,
  colorValue: 'auto',
  bold: false,
  italic: false,
  underline: null,
  strike: false,
}

describe('mergeParagraphStyle', () => {
  it('creates a schema-ordered paragraph style with its own bullet list and applies it via pStyle', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>First item</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Second item</w:t></w:r></w:p>
    </w:body></w:document>`

    const parsedDocx = makeParsedDocx({ documentXml })
    const initialReport = buildStyleReport(parsedDocx)

    const styleId = mergeParagraphStyle(
      parsedDocx,
      allRunRefs(initialReport),
      { ...NEUTRAL_SIGNATURE, bold: true },
      'Bulleted Style',
      'bullet',
    )

    // 1. The new style is a paragraph style with pPr (numPr) before rPr.
    const stylesRoot = parsedDocx.stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
    const styleEls = Array.from(stylesRoot.getElementsByTagNameNS(NS.w, 'style'))
    expect(styleEls).toHaveLength(1)
    expect(styleEls[0].getAttributeNS(NS.w, 'type')).toBe('paragraph')
    expect(Array.from(styleEls[0].children).map((c) => c.localName)).toEqual(['name', 'pPr', 'rPr'])

    const pPr = wChild(styleEls[0], 'pPr')!
    const numPr = wChild(pPr, 'numPr')!
    const numId = wChild(numPr, 'numId')!.getAttributeNS(NS.w, 'val')
    expect(numId).toBeTruthy()

    // 2. numbering.xml was created with a matching abstractNum/num pair.
    expect(parsedDocx.numberingXml).not.toBeNull()
    const numberingRoot = parsedDocx.numberingXml!.getElementsByTagNameNS(NS.w, 'numbering')[0]
    const numEls = Array.from(numberingRoot.getElementsByTagNameNS(NS.w, 'num'))
    expect(numEls).toHaveLength(1)
    expect(numEls[0].getAttributeNS(NS.w, 'numId')).toBe(numId)
    const abstractNumId = wChild(numEls[0], 'abstractNumId')!.getAttributeNS(NS.w, 'val')
    const abstractEls = Array.from(numberingRoot.getElementsByTagNameNS(NS.w, 'abstractNum'))
    expect(abstractEls).toHaveLength(1)
    expect(abstractEls[0].getAttributeNS(NS.w, 'abstractNumId')).toBe(abstractNumId)
    const lvl = wChild(abstractEls[0], 'lvl')!
    expect(wChild(lvl, 'numFmt')!.getAttributeNS(NS.w, 'val')).toBe('bullet')

    // 3. Every affected paragraph points at the new style, and every
    //    affected run had its direct formatting stripped (the paragraph
    //    style's rPr now determines the look).
    const paragraphs = Array.from(parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'p'))
    for (const paragraphEl of paragraphs) {
      const pStyle = wChild(wChild(paragraphEl, 'pPr'), 'pStyle')
      expect(pStyle!.getAttributeNS(NS.w, 'val')).toBe(styleId)
      const runRPr = wChild(paragraphEl.getElementsByTagNameNS(NS.w, 'r')[0], 'rPr')
      expect(wChild(runRPr, 'b')).toBeNull()
    }

    // 4. Both paragraphs resolve to sequential bullet markers.
    const markers = buildParagraphMarkers(parsedDocx)
    expect(markers.get(paragraphs[0])?.text).toBe('•')
    expect(markers.get(paragraphs[1])?.text).toBe('•')

    // 5. The Style Report attributes both runs to the new named-paragraph
    //    style, and countOccurrencesForStyleId (paragraph-style-aware) sees them.
    const finalReport = buildStyleReport(parsedDocx)
    expect(finalReport).toHaveLength(1)
    expect(finalReport[0].variants[0].origin).toEqual({
      kind: 'named-paragraph',
      styleId,
      styleName: 'Bulleted Style',
    })
    expect(countOccurrencesForStyleId(finalReport, styleId)).toBe(2)
  })

  it('creates a numbered (decimal) list with a distinct numId per style', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:r><w:t>Item A</w:t></w:r></w:p>
    </w:body></w:document>`
    const parsedDocx = makeParsedDocx({ documentXml })

    const firstId = mergeParagraphStyle(parsedDocx, [], NEUTRAL_SIGNATURE, 'Numbered A', 'decimal')
    const secondId = mergeParagraphStyle(parsedDocx, [], NEUTRAL_SIGNATURE, 'Numbered B', 'decimal')
    expect(firstId).not.toBe(secondId)

    const numberingRoot = parsedDocx.numberingXml!.getElementsByTagNameNS(NS.w, 'numbering')[0]
    const numEls = Array.from(numberingRoot.getElementsByTagNameNS(NS.w, 'num'))
    expect(numEls).toHaveLength(2)
    // Every abstractNum must precede every num per CT_Numbering's schema order.
    const children = Array.from(numberingRoot.children).map((c) => c.localName)
    const lastAbstractIndex = children.lastIndexOf('abstractNum')
    const firstNumIndex = children.indexOf('num')
    expect(lastAbstractIndex).toBeLessThan(firstNumIndex)
  })

  it('removes a paragraph\'s own direct numPr so the style\'s numbering (or lack thereof) takes over', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="99"/></w:numPr></w:pPr><w:r><w:t>Old list item</w:t></w:r></w:p>
    </w:body></w:document>`
    const parsedDocx = makeParsedDocx({ documentXml })
    const runRefs = allRunRefs(buildStyleReport(parsedDocx))

    const styleId = mergeParagraphStyle(parsedDocx, runRefs, NEUTRAL_SIGNATURE, 'Plain Style', 'none')

    const paragraphEl = parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'p')[0]
    const pPr = wChild(paragraphEl, 'pPr')!
    expect(wChild(pPr, 'numPr')).toBeNull()
    expect(wChild(pPr, 'pStyle')!.getAttributeNS(NS.w, 'val')).toBe(styleId)
  })

  it('reuses an existing paragraph style when reuseExistingStyleId is given, switching list format in place', () => {
    const documentXml = `<w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`
    const parsedDocx = makeParsedDocx({ documentXml })

    const firstId = mergeParagraphStyle(
      parsedDocx,
      allRunRefs(buildStyleReport(parsedDocx)),
      NEUTRAL_SIGNATURE,
      'My Paragraph Style',
      'none',
    )
    expect(parsedDocx.numberingXml).toBeNull() // no list requested yet

    const secondId = mergeParagraphStyle(parsedDocx, [], NEUTRAL_SIGNATURE, 'My Paragraph Style', 'decimal', firstId)
    expect(secondId).toBe(firstId)

    const stylesRoot = parsedDocx.stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
    expect(Array.from(stylesRoot.getElementsByTagNameNS(NS.w, 'style'))).toHaveLength(1)
    expect(parsedDocx.numberingXml).not.toBeNull()

    const paragraphEl = parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'p')[0]
    expect(wChild(wChild(paragraphEl, 'pPr'), 'pStyle')!.getAttributeNS(NS.w, 'val')).toBe(firstId)
  })

  it('deduplicates paragraphs: a paragraph with multiple selected runs only gets pStyle set once', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:p>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Run one</w:t></w:r>
        <w:r><w:rPr><w:i/></w:rPr><w:t>Run two</w:t></w:r>
      </w:p>
    </w:body></w:document>`
    const parsedDocx = makeParsedDocx({ documentXml })
    const runRefs = allRunRefs(buildStyleReport(parsedDocx))
    expect(runRefs).toHaveLength(2)

    mergeParagraphStyle(parsedDocx, runRefs, NEUTRAL_SIGNATURE, 'Shared Paragraph Style', 'none')

    const paragraphEl = parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'p')[0]
    const pPrs = Array.from(paragraphEl.getElementsByTagNameNS(NS.w, 'pPr')).filter(
      (el) => el.parentNode === paragraphEl,
    )
    expect(pPrs).toHaveLength(1)
  })
})

describe('collectRunRefsForVariantIds with paragraph-style variants', () => {
  it('still finds runs after they regroup under a named-paragraph origin', () => {
    const documentXml = `<w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`
    const parsedDocx = makeParsedDocx({ documentXml })
    mergeParagraphStyle(parsedDocx, allRunRefs(buildStyleReport(parsedDocx)), NEUTRAL_SIGNATURE, 'Style', 'bullet')

    const report = buildStyleReport(parsedDocx)
    const variantId = report[0].variants[0].id
    const refs = collectRunRefsForVariantIds(report, new Set([variantId]))
    expect(refs).toHaveLength(1)
  })
})
