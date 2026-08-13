import { describe, expect, it } from 'vitest'
import { NS } from '../src/lib/ooxml/constants'
import {
  buildResolutionContext,
  buildStylesMap,
  getDocDefaultsRPr,
  resolveRunFormatting,
} from '../src/lib/ooxml/styleResolution'
import { buildThemeColorMap } from '../src/lib/ooxml/themeColor'
import { parseXmlString } from './testUtils'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

function resolveFirstRun(documentXmlStr: string, stylesXmlStr: string, themeXmlStr?: string) {
  const documentXml = parseXmlString(documentXmlStr)
  const stylesXml = parseXmlString(stylesXmlStr)
  const themeXml = themeXmlStr ? parseXmlString(themeXmlStr) : null

  const ctx = buildResolutionContext(
    buildStylesMap(stylesXml),
    getDocDefaultsRPr(stylesXml),
    buildThemeColorMap(themeXml),
  )

  const paragraphEl = documentXml.getElementsByTagNameNS(NS.w, 'p')[0]
  const runEl = documentXml.getElementsByTagNameNS(NS.w, 'r')[0]
  return resolveRunFormatting(runEl, paragraphEl, ctx)
}

describe('resolveRunFormatting', () => {
  it('falls back to docDefaults when nothing else is set', () => {
    const stylesXml = `<w:styles ${W}>
      <w:docDefaults>
        <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
      </w:docDefaults>
    </w:styles>`
    const documentXml = `<w:document ${W}><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>`

    const { signature, origin } = resolveFirstRun(documentXml, stylesXml)
    expect(signature.fontFamily).toBe('Calibri')
    expect(signature.fontSizeHalfPt).toBe(22)
    expect(origin.kind).toBe('direct') // no pStyle anywhere -> nothing named to attribute it to
  })

  it('resolves a 2-level basedOn chain, with the child style overriding the base font', () => {
    const stylesXml = `<w:styles ${W}>
      <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Base">
        <w:name w:val="Base"/>
        <w:rPr><w:rFonts w:ascii="Georgia"/><w:color w:val="112233"/></w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Child">
        <w:name w:val="Child"/>
        <w:basedOn w:val="Base"/>
        <w:rPr><w:rFonts w:ascii="Verdana"/></w:rPr>
      </w:style>
    </w:styles>`
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>
    </w:body></w:document>`

    const { signature, origin } = resolveFirstRun(documentXml, stylesXml)
    expect(signature.fontFamily).toBe('Verdana') // overridden by Child
    expect(signature.colorValue).toBe('112233') // inherited from Base
    expect(signature.fontSizeHalfPt).toBe(20) // inherited from docDefaults
    expect(origin).toEqual({ kind: 'named-paragraph', styleId: 'Child', styleName: 'Child' })
  })

  it('overlays a character style (w:rStyle) on top of the paragraph style chain', () => {
    const stylesXml = `<w:styles ${W}>
      <w:style w:type="paragraph" w:styleId="Para">
        <w:name w:val="Para"/>
        <w:rPr><w:b/><w:rFonts w:ascii="Arial"/></w:rPr>
      </w:style>
      <w:style w:type="character" w:styleId="Emph">
        <w:name w:val="Emphasis"/>
        <w:rPr><w:color w:val="ff0000"/></w:rPr>
      </w:style>
    </w:styles>`
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:pPr><w:pStyle w:val="Para"/></w:pPr>
        <w:r><w:rPr><w:rStyle w:val="Emph"/></w:rPr><w:t>Hi</w:t></w:r>
      </w:p>
    </w:body></w:document>`

    const { signature, origin } = resolveFirstRun(documentXml, stylesXml)
    expect(signature.bold).toBe(true) // inherited from the paragraph style
    expect(signature.colorValue).toBe('FF0000') // from the character style, normalized uppercase
    expect(origin).toEqual({ kind: 'named-character', styleId: 'Emph', styleName: 'Emphasis' })
  })

  it('direct run formatting wins over everything and is attributed as "direct"', () => {
    const stylesXml = `<w:styles ${W}>
      <w:style w:type="paragraph" w:styleId="Para">
        <w:name w:val="Para"/>
        <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr>
      </w:style>
    </w:styles>`
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:pPr><w:pStyle w:val="Para"/></w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Comic Sans MS"/></w:rPr><w:t>Hi</w:t></w:r>
      </w:p>
    </w:body></w:document>`

    const { signature, origin } = resolveFirstRun(documentXml, stylesXml)
    expect(signature.fontFamily).toBe('Comic Sans MS')
    expect(signature.fontSizeHalfPt).toBe(20) // still inherited - direct rPr didn't touch size
    expect(origin).toEqual({ kind: 'direct' })
  })

  it('resolves a theme color from theme1.xml when @w:val is absent/auto', () => {
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:r><w:rPr><w:color w:val="auto" w:themeColor="accent1"/></w:rPr><w:t>Hi</w:t></w:r></w:p>
    </w:body></w:document>`
    const themeXml = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
          <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        </a:clrScheme>
      </a:themeElements>
    </a:theme>`

    const { signature } = resolveFirstRun(documentXml, `<w:styles ${W}></w:styles>`, themeXml)
    expect(signature.colorValue).toBe('4472C4')
  })

  it('treats a bare toggle element as true and @w:val="0" as false', () => {
    const documentXml = parseXmlString(`<w:document ${W}><w:body>
      <w:p>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>
        <w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>NotBold</w:t></w:r>
      </w:p>
    </w:body></w:document>`)
    const stylesXml = parseXmlString(`<w:styles ${W}></w:styles>`)
    const ctx = buildResolutionContext(buildStylesMap(stylesXml), getDocDefaultsRPr(stylesXml), new Map())
    const paragraphEl = documentXml.getElementsByTagNameNS(NS.w, 'p')[0]
    const runs = documentXml.getElementsByTagNameNS(NS.w, 'r')

    expect(resolveRunFormatting(runs[0], paragraphEl, ctx).signature.bold).toBe(true)
    expect(resolveRunFormatting(runs[1], paragraphEl, ctx).signature.bold).toBe(false)
  })

  it('guards against a circular w:basedOn chain instead of infinite-looping', () => {
    const stylesXml = `<w:styles ${W}>
      <w:style w:type="paragraph" w:styleId="A">
        <w:name w:val="A"/>
        <w:basedOn w:val="B"/>
        <w:rPr><w:rFonts w:ascii="Arial"/></w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="B">
        <w:name w:val="B"/>
        <w:basedOn w:val="A"/>
      </w:style>
    </w:styles>`
    const documentXml = `<w:document ${W}><w:body>
      <w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>
    </w:body></w:document>`

    expect(() => resolveFirstRun(documentXml, stylesXml)).not.toThrow()
  })
})
