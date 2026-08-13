import type { FormattingSignature } from '../../types/ooxml'
import { TRACKED_RPR_LOCAL_NAMES } from './constants'
import { createWEl, insertRPrChildInOrder, removeWChild, setWAttr } from './domUtils'

/** Writes every tracked attribute of `sig` into `rPr` as explicit children
 * (creating or replacing each one), so the element fully and unambiguously
 * determines all 7 tracked properties - used when building a merged style's
 * <w:rPr>, where we never want an omitted child to silently mean "inherit". */
export function writeSignatureIntoRPr(rPr: Element, sig: FormattingSignature): void {
  const doc = rPr.ownerDocument

  const rFonts = createWEl(doc, 'rFonts')
  const font = sig.fontFamily ?? 'Calibri'
  setWAttr(rFonts, 'ascii', font)
  setWAttr(rFonts, 'hAnsi', font)
  setWAttr(rFonts, 'cs', font)
  insertRPrChildInOrder(rPr, 'rFonts', rFonts)

  writeToggle(rPr, 'b', sig.bold)
  writeToggle(rPr, 'i', sig.italic)
  writeToggle(rPr, 'strike', sig.strike)

  const color = createWEl(doc, 'color')
  setWAttr(color, 'val', sig.colorValue)
  insertRPrChildInOrder(rPr, 'color', color)

  const sizeHalfPt = String(sig.fontSizeHalfPt ?? 24) // 24 half-points = 12pt, Word's common default
  const sz = createWEl(doc, 'sz')
  setWAttr(sz, 'val', sizeHalfPt)
  insertRPrChildInOrder(rPr, 'sz', sz)
  const szCs = createWEl(doc, 'szCs')
  setWAttr(szCs, 'val', sizeHalfPt)
  insertRPrChildInOrder(rPr, 'szCs', szCs)

  if (sig.underline) {
    const u = createWEl(doc, 'u')
    setWAttr(u, 'val', sig.underline)
    insertRPrChildInOrder(rPr, 'u', u)
  } else {
    removeWChild(rPr, 'u')
  }
}

function writeToggle(rPr: Element, localName: string, value: boolean): void {
  if (value) {
    insertRPrChildInOrder(rPr, localName, createWEl(rPr.ownerDocument, localName))
  } else {
    removeWChild(rPr, localName)
  }
}

/** Removes every tracked-property child from `rPr` - used after pointing a
 * run at a merged style, since the style now fully determines those
 * properties and leaving the old direct overrides in place would contradict
 * (or in some viewers, take precedence over) the new w:rStyle reference. */
export function stripTrackedProps(rPr: Element): void {
  for (const localName of TRACKED_RPR_LOCAL_NAMES) {
    removeWChild(rPr, localName)
  }
}
