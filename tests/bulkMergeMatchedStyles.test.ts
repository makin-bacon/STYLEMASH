import { describe, expect, it } from 'vitest'
import type { UserStyleRecord } from '../src/types/ooxml'
import {
  bulkMergeVariantsIntoMatchingReferenceStyles,
  findVariantIdsMatchingReferenceStyleNames,
} from '../src/lib/ooxml/bulkMergeMatchedStyles'
import { NS } from '../src/lib/ooxml/constants'
import { wChild } from '../src/lib/ooxml/domUtils'
import { buildStyleReport } from '../src/lib/ooxml/styleReport'
import { makeParsedDocx } from './testUtils'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

function referenceRecord(name: string, styleId: string, bold: boolean): UserStyleRecord {
  return {
    styleId,
    name,
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: null,
      colorValue: 'auto',
      bold,
      italic: false,
      underline: null,
      strike: false,
    },
    createdAt: 1,
    fromReferenceDoc: true,
  }
}

describe('findVariantIdsMatchingReferenceStyleNames', () => {
  it('matches only named-style variants whose style name is a Document-B-derived record', () => {
    const parsedDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:rStyle w:val="Match1"/></w:rPr><w:t>a</w:t></w:r></w:p>
        <w:p><w:r><w:rPr><w:rStyle w:val="NoMatch"/></w:rPr><w:t>b</w:t></w:r></w:p>
        <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>c</w:t></w:r></w:p>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="character" w:styleId="Match1"><w:name w:val="Emph"/><w:rPr/></w:style>
        <w:style w:type="character" w:styleId="NoMatch"><w:name w:val="Other"/><w:rPr/></w:style>
      </w:styles>`,
    })
    const styleReport = buildStyleReport(parsedDocx)
    const userStyles = [referenceRecord('Emph', 'EmphInA', true)]

    const matched = findVariantIdsMatchingReferenceStyleNames(styleReport, userStyles)

    expect(matched.size).toBe(1)
    const matchedVariant = styleReport
      .flatMap((e) => e.variants)
      .find((v) => matched.has(v.id))!
    expect(matchedVariant.origin).toMatchObject({ kind: 'named-character', styleName: 'Emph' })
  })

  it('returns an empty set when there are no Document-B-derived records', () => {
    const parsedDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>`,
    })
    const matched = findVariantIdsMatchingReferenceStyleNames(buildStyleReport(parsedDocx), [])
    expect(matched.size).toBe(0)
  })
})

describe('bulkMergeVariantsIntoMatchingReferenceStyles', () => {
  it('splits a selection spanning two matched names across their respective Document B targets', () => {
    const parsedDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:rStyle w:val="HeadingA"/></w:rPr><w:t>heading text</w:t></w:r></w:p>
        <w:p><w:r><w:rPr><w:rStyle w:val="EmphA"/></w:rPr><w:t>emph text</w:t></w:r></w:p>
      </w:body></w:document>`,
      stylesXml: `<w:styles ${W}>
        <w:style w:type="character" w:styleId="HeadingA"><w:name w:val="Heading 1"/><w:rPr/></w:style>
        <w:style w:type="character" w:styleId="EmphA"><w:name w:val="Emph"/><w:rPr/></w:style>
        <w:style w:type="character" w:styleId="HeadingB-in-A"><w:name w:val="Heading 1"/><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="EmphB-in-A"><w:name w:val="Emph"/><w:rPr><w:i/></w:rPr></w:style>
      </w:styles>`,
    })
    const styleReport = buildStyleReport(parsedDocx)
    const selectedVariantIds = findVariantIdsMatchingReferenceStyleNames(styleReport, [
      referenceRecord('Heading 1', 'HeadingB-in-A', true),
      referenceRecord('Emph', 'EmphB-in-A', false),
    ])
    expect(selectedVariantIds.size).toBe(2)

    const userStyles = [
      referenceRecord('Heading 1', 'HeadingB-in-A', true),
      referenceRecord('Emph', 'EmphB-in-A', false),
    ]
    userStyles[1].targetSignature.italic = true

    bulkMergeVariantsIntoMatchingReferenceStyles(parsedDocx, styleReport, selectedVariantIds, userStyles)

    const runs = Array.from(parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'r'))
    const headingRun = runs.find((r) => r.textContent?.includes('heading'))!
    const emphRun = runs.find((r) => r.textContent?.includes('emph'))!

    expect(wChild(wChild(headingRun, 'rPr'), 'rStyle')!.getAttributeNS(NS.w, 'val')).toBe('HeadingB-in-A')
    expect(wChild(wChild(emphRun, 'rPr'), 'rStyle')!.getAttributeNS(NS.w, 'val')).toBe('EmphB-in-A')

    // No new <w:style> elements were created - both targets already existed.
    const styleCount = parsedDocx.stylesXml.getElementsByTagNameNS(NS.w, 'style').length
    expect(styleCount).toBe(4)
  })

  it('leaves unmatched selected variants (direct formatting, non-matching names) untouched', () => {
    const parsedDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body>
        <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>direct</w:t></w:r></w:p>
      </w:body></w:document>`,
    })
    const styleReport = buildStyleReport(parsedDocx)
    const allVariantIds = new Set(styleReport.flatMap((e) => e.variants).map((v) => v.id))

    bulkMergeVariantsIntoMatchingReferenceStyles(parsedDocx, styleReport, allVariantIds, [
      referenceRecord('Emph', 'EmphInA', true),
    ])

    // Nothing to merge into (no named-style match) - run untouched.
    const run = parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'r')[0]
    expect(wChild(run, 'rPr')!.getElementsByTagNameNS(NS.w, 'rStyle')).toHaveLength(0)
    expect(wChild(wChild(run, 'rPr'), 'b')).not.toBeNull()
  })

  it('is a no-op when there are no Document-B-derived records', () => {
    const parsedDocx = makeParsedDocx({
      documentXml: `<w:document ${W}><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>`,
    })
    const styleReport = buildStyleReport(parsedDocx)
    const allVariantIds = new Set(styleReport.flatMap((e) => e.variants).map((v) => v.id))

    bulkMergeVariantsIntoMatchingReferenceStyles(parsedDocx, styleReport, allVariantIds, [])

    expect(parsedDocx.stylesXml.getElementsByTagNameNS(NS.w, 'style')).toHaveLength(0)
  })
})
