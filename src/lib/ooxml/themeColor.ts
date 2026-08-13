import { NS } from './constants'
import { wAttr } from './domUtils'

// Maps OOXML w:themeColor values to DrawingML clrScheme child local names.
// See ECMA-376 Part 1 §17.3.2.6 (color) and §20.1.4.1.14 (clrScheme).
const THEME_COLOR_TO_SCHEME_TAG: Record<string, string> = {
  dk1: 'dk1',
  lt1: 'lt1',
  dk2: 'dk2',
  lt2: 'lt2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hyperlink: 'hlink',
  followedHyperlink: 'folHlink',
}

/** Builds a themeColor -> '#RRGGBB' lookup from theme1.xml's <a:clrScheme>.
 * Returns an empty map if the theme part is missing or malformed - callers
 * fall back to 'auto' in that case. */
export function buildThemeColorMap(themeXml: XMLDocument | null): Map<string, string> {
  const map = new Map<string, string>()
  if (!themeXml) return map

  const clrSchemes = themeXml.getElementsByTagNameNS(NS.a, 'clrScheme')
  const clrScheme = clrSchemes[0]
  if (!clrScheme) return map

  for (const [themeKey, schemeTag] of Object.entries(THEME_COLOR_TO_SCHEME_TAG)) {
    const schemeEl = findDirectChildByLocalName(clrScheme, schemeTag)
    if (!schemeEl) continue
    // Each scheme color element wraps exactly one of <a:srgbClr val="RRGGBB"/>
    // or <a:sysClr val="..." lastClr="RRGGBB"/>.
    const srgb = schemeEl.getElementsByTagNameNS(NS.a, 'srgbClr')[0]
    const sys = schemeEl.getElementsByTagNameNS(NS.a, 'sysClr')[0]
    const hex = srgb?.getAttribute('val') ?? sys?.getAttribute('lastClr')
    if (hex) map.set(themeKey, hex.toUpperCase())
  }
  return map
}

function findDirectChildByLocalName(parent: Element, localName: string): Element | null {
  for (const child of Array.from(parent.children)) {
    if (child.localName === localName) return child
  }
  return null
}

/** Resolves a <w:color> element to a '#RRGGBB'-less hex string or 'auto'.
 *
 * v1 scope: if @w:val is a concrete hex, it's used as-is (Word always writes
 * the computed RGB alongside a w:themeColor hint, so this covers the vast
 * majority of real documents). Only when @w:val is absent/"auto" do we fall
 * back to the theme scheme lookup. w:themeTint/w:themeShade percentage
 * lightening/darkening is NOT applied - a documented, cosmetically-minor
 * v1 limitation. */
export function resolveColorElement(
  colorEl: Element | null,
  themeColors: Map<string, string>,
): string {
  if (!colorEl) return 'auto'

  const val = wAttr(colorEl, 'val')
  if (val && val.toLowerCase() !== 'auto') return val.toUpperCase()

  const themeColor = wAttr(colorEl, 'themeColor')
  if (themeColor && themeColors.has(themeColor)) {
    return themeColors.get(themeColor)!
  }

  return 'auto'
}
