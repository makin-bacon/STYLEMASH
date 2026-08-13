import { describe, expect, it } from 'vitest'
import type { FormattingSignature } from '../src/types/ooxml'
import { NS } from '../src/lib/ooxml/constants'
import { wChild } from '../src/lib/ooxml/domUtils'
import { mergeStyles } from '../src/lib/ooxml/mergeStyles'
import { buildStyleReport } from '../src/lib/ooxml/styleReport'
import { makeParsedDocx } from './testUtils'
import type { RunRef, StyleEntity } from '../src/types/ooxml'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

/** Flattens every RunRef across every variant of every entity - the merge
 * dialog normally does this via collectRunRefsForVariantIds() against a
 * specific selection, but these tests just want "everything". */
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

describe('mergeStyles', () => {
  it('creates a schema-ordered character style and repoints matching runs to it', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:p>
        <w:r><w:rPr><w:b/><w:vertAlign w:val="superscript"/></w:rPr><w:t>One</w:t></w:r>
        <w:r><w:rPr><w:i/></w:rPr><w:t>Two</w:t></w:r>
      </w:p>
    </w:body></w:document>`

    const parsedDocx = makeParsedDocx({ documentXml })
    const initialReport = buildStyleReport(parsedDocx)
    expect(initialReport).toHaveLength(2) // a bold entity and an italic entity

    const styleId = mergeStyles(
      parsedDocx,
      allRunRefs(initialReport),
      { fontFamily: 'Calibri', fontSizeHalfPt: 24, colorValue: 'FF0000', bold: true, italic: false, underline: null, strike: false },
      'Merged Style',
    )

    // 1. The new style exists, is a character style, and its rPr children
    //    are in CT_RPr schema order.
    const stylesRoot = parsedDocx.stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
    const styleEls = Array.from(stylesRoot.getElementsByTagNameNS(NS.w, 'style'))
    expect(styleEls).toHaveLength(1)
    expect(styleEls[0].getAttributeNS(NS.w, 'type')).toBe('character')
    expect(styleEls[0].getAttributeNS(NS.w, 'styleId')).toBe(styleId)

    const rPr = wChild(styleEls[0], 'rPr')!
    expect(Array.from(rPr.children).map((c) => c.localName)).toEqual([
      'rFonts',
      'b',
      'color',
      'sz',
      'szCs',
    ])

    // 2. Every affected run references the new style and had its tracked
    //    direct-formatting properties stripped...
    const runs = Array.from(parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'r'))
    for (const run of runs) {
      const runRPr = wChild(run, 'rPr')!
      expect(wChild(runRPr, 'rStyle')!.getAttributeNS(NS.w, 'val')).toBe(styleId)
      expect(wChild(runRPr, 'b')).toBeNull()
      expect(wChild(runRPr, 'i')).toBeNull()
    }

    // 3. ...but an untracked property (vertAlign) on the first run survives.
    expect(wChild(wChild(runs[0], 'rPr')!, 'vertAlign')).not.toBeNull()

    // 4. Re-running the Style Report collapses both merged runs into one
    //    entity, correctly attributed to the new named-character style.
    const finalReport = buildStyleReport(parsedDocx)
    expect(finalReport).toHaveLength(1)
    expect(finalReport[0].occurrenceCount).toBe(2)
    expect(finalReport[0].variants).toHaveLength(1)
    expect(finalReport[0].variants[0].origin).toEqual({
      kind: 'named-character',
      styleId,
      styleName: 'Merged Style',
    })
  })

  it('reuses an existing style when reuseExistingStyleId is given, redefining it in place', () => {
    const documentXml = `<w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`
    const parsedDocx = makeParsedDocx({ documentXml })

    const firstStyleId = mergeStyles(
      parsedDocx,
      allRunRefs(buildStyleReport(parsedDocx)),
      NEUTRAL_SIGNATURE,
      'My Style',
    )

    // Redefine the same style's look without merging any newly-selected entities.
    const secondStyleId = mergeStyles(
      parsedDocx,
      [],
      { fontFamily: 'Georgia', fontSizeHalfPt: 28, colorValue: 'auto', bold: true, italic: false, underline: null, strike: false },
      'My Style',
      firstStyleId,
    )

    expect(secondStyleId).toBe(firstStyleId)
    const stylesRoot = parsedDocx.stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
    expect(Array.from(stylesRoot.getElementsByTagNameNS(NS.w, 'style'))).toHaveLength(1)

    const report = buildStyleReport(parsedDocx)
    expect(report[0].signature.fontFamily).toBe('Georgia')
    expect(report[0].signature.bold).toBe(true)
  })

  it('generates unique styleIds when names collide', () => {
    const documentXml = `<w:document ${W}><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p></w:body></w:document>`
    const parsedDocx = makeParsedDocx({ documentXml })

    const id1 = mergeStyles(parsedDocx, [], NEUTRAL_SIGNATURE, 'Custom Style')
    const id2 = mergeStyles(parsedDocx, [], NEUTRAL_SIGNATURE, 'Custom Style')

    expect(id1).not.toBe(id2)
  })
})
