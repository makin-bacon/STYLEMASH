import { describe, expect, it } from 'vitest'
import { NS } from '../src/lib/ooxml/constants'
import { wChild } from '../src/lib/ooxml/domUtils'
import { buildStyleReport } from '../src/lib/ooxml/styleReport'
import { applyXmlFragmentToRunRefs, parseXmlFragment, XmlFragmentError } from '../src/lib/ooxml/xmlFragmentEdit'
import { makeParsedDocx } from './testUtils'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

describe('parseXmlFragment', () => {
  it('parses a valid <w:rPr> fragment', () => {
    const el = parseXmlFragment('<w:rPr><w:b/></w:rPr>', 'rPr')
    expect(el.localName).toBe('rPr')
    expect(el.namespaceURI).toBe(NS.w)
  })

  it('throws on malformed XML', () => {
    expect(() => parseXmlFragment('<w:rPr><w:b></w:rPr>', 'rPr')).toThrow(XmlFragmentError)
  })

  it('throws when more than one root element is given', () => {
    expect(() => parseXmlFragment('<w:rPr/><w:rPr/>', 'rPr')).toThrow(XmlFragmentError)
  })

  it('throws when the wrong element is given', () => {
    expect(() => parseXmlFragment('<w:pPr/>', 'rPr')).toThrow(XmlFragmentError)
  })
})

describe('applyXmlFragmentToRunRefs', () => {
  it('applies the edited rPr to every run in the variant, and the report regroups afterward', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:p>
        <w:r><w:rPr><w:b/></w:rPr><w:t>One</w:t></w:r>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Two</w:t></w:r>
      </w:p>
    </w:body></w:document>`
    const parsedDocx = makeParsedDocx({ documentXml })
    const [entity] = buildStyleReport(parsedDocx)
    expect(entity.occurrenceCount).toBe(2)
    expect(entity.variants).toHaveLength(1)

    applyXmlFragmentToRunRefs(
      parsedDocx,
      entity.variants[0].runRefs,
      '<w:rPr><w:i/><w:color w:val="00FF00"/></w:rPr>',
    )

    const runs = Array.from(parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'r'))
    for (const run of runs) {
      const rPr = wChild(run, 'rPr')!
      expect(wChild(rPr, 'b')).toBeNull() // old rPr replaced wholesale, not merged
      expect(wChild(rPr, 'i')).not.toBeNull()
      expect(wChild(rPr, 'color')!.getAttributeNS(NS.w, 'val')).toBe('00FF00')
    }

    const report = buildStyleReport(parsedDocx)
    expect(report).toHaveLength(1)
    expect(report[0].occurrenceCount).toBe(2)
    expect(report[0].signature.italic).toBe(true)
  })
})
