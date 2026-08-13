import { describe, expect, it } from 'vitest'
import {
  buildStyleReport,
  collectRunRefsForVariantIds,
  findVariantById,
} from '../src/lib/ooxml/styleReport'
import { makeParsedDocx } from './testUtils'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

describe('buildStyleReport', () => {
  it('groups runs by resolved visual signature, split into variants by origin', () => {
    const stylesXml = `<w:styles ${W}>
      <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
        <w:name w:val="Normal"/>
      </w:style>
      <w:style w:type="character" w:styleId="Emph">
        <w:name w:val="Emphasis"/>
        <w:rPr><w:b/></w:rPr>
      </w:style>
    </w:styles>`

    // Two runs that look identical (bold) - one via a named character style,
    // one via direct formatting - should collapse into ONE entity, but stay
    // split into two selectable variants (one per origin).
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:r><w:rPr><w:rStyle w:val="Emph"/></w:rPr><w:t>Bold via style</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold direct</w:t></w:r></w:p>
      <w:p><w:r><w:t>Plain text</w:t></w:r></w:p>
    </w:body></w:document>`

    const report = buildStyleReport(makeParsedDocx({ documentXml, stylesXml }))

    expect(report).toHaveLength(2) // bold entity + plain entity

    const boldEntity = report.find((e) => e.signature.bold)!
    expect(boldEntity.occurrenceCount).toBe(2)
    expect(boldEntity.variants).toHaveLength(2) // one 'named-character' variant, one 'direct'
    expect(boldEntity.variants.map((v) => v.origin.kind).sort()).toEqual(['direct', 'named-character'])
    expect(boldEntity.variants.every((v) => v.occurrenceCount === 1)).toBe(true)

    const plainEntity = report.find((e) => !e.signature.bold)!
    expect(plainEntity.occurrenceCount).toBe(1)
    expect(plainEntity.variants).toHaveLength(1)
    expect(plainEntity.sampleText).toBe('Plain text')
  })

  it('sorts entities by occurrence count, most common first', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Rare</w:t></w:r></w:p>
      <w:p>
        <w:r><w:t>Common 1</w:t></w:r>
        <w:r><w:t>Common 2</w:t></w:r>
        <w:r><w:t>Common 3</w:t></w:r>
      </w:p>
    </w:body></w:document>`

    const report = buildStyleReport(makeParsedDocx({ documentXml }))
    expect(report[0].occurrenceCount).toBe(3)
    expect(report[1].occurrenceCount).toBe(1)
  })

  it('skips runs with no visible text (e.g. tab-only runs)', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:r><w:tab/></w:r><w:r><w:t>Visible</w:t></w:r></w:p>
    </w:body></w:document>`

    const report = buildStyleReport(makeParsedDocx({ documentXml }))
    expect(report).toHaveLength(1)
    expect(report[0].occurrenceCount).toBe(1)
  })

  it('traverses runs nested inside table cells', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:tbl>
        <w:tr><w:tc><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    </w:body></w:document>`

    const report = buildStyleReport(makeParsedDocx({ documentXml }))
    expect(report).toHaveLength(1)
    expect(report[0].sampleText).toBe('Cell text')
  })

  it('findVariantById and collectRunRefsForVariantIds locate the right runs across entities', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:p>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>
        <w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r>
        <w:r><w:t>Plain</w:t></w:r>
      </w:p>
    </w:body></w:document>`

    const report = buildStyleReport(makeParsedDocx({ documentXml }))
    const boldVariant = report.find((e) => e.signature.bold)!.variants[0]
    const italicVariant = report.find((e) => e.signature.italic)!.variants[0]

    expect(findVariantById(report, boldVariant.id)).toBe(boldVariant)
    expect(findVariantById(report, 'not-a-real-id')).toBeNull()

    const runRefs = collectRunRefsForVariantIds(report, new Set([boldVariant.id, italicVariant.id]))
    expect(runRefs).toHaveLength(2)
    expect(runRefs).toEqual(expect.arrayContaining([...boldVariant.runRefs, ...italicVariant.runRefs]))
  })
})
