import type { FormattingSignature } from '../../types/ooxml'
import { wAttr } from './domUtils'

/** Reads a <w:b>/<w:i>/<w:strike>-style toggle element. Per OOXML, the bare
 * element (no @w:val) means "on"; @w:val of "0"/"false"/"off" means "off".
 * We do NOT implement the spec's full nested-style XOR toggle semantics
 * (§17.3.2.1) - this is a documented v1 simplification since it only
 * matters for rare, deeply-nested style configurations. */
export function readToggle(el: Element | undefined | null): boolean {
  if (!el) return false
  const val = wAttr(el, 'val')
  if (val === null) return true
  return !(val === '0' || val === 'false' || val === 'off')
}

/** Converts the cascade-resolved set of tracked <w:rPr> children into the
 * flat FormattingSignature used for grouping and display. `resolveColor`
 * is injected so this module doesn't need to know about theme.xml. */
export function trackedChildrenToSignature(
  tracked: Map<string, Element>,
  resolveColor: (colorEl: Element | null) => string,
): FormattingSignature {
  const rFonts = tracked.get('rFonts') ?? null
  const sz = tracked.get('sz') ?? null
  const u = tracked.get('u') ?? null
  const underlineVal = u ? wAttr(u, 'val') : null

  return {
    fontFamily: rFonts ? (wAttr(rFonts, 'ascii') ?? wAttr(rFonts, 'hAnsi') ?? null) : null,
    fontSizeHalfPt: sz ? parseIntOrNull(wAttr(sz, 'val')) : null,
    colorValue: resolveColor(tracked.get('color') ?? null),
    bold: readToggle(tracked.get('b')),
    italic: readToggle(tracked.get('i')),
    underline: underlineVal && underlineVal !== 'none' ? underlineVal : null,
    strike: readToggle(tracked.get('strike')),
  }
}

function parseIntOrNull(val: string | null): number | null {
  if (val === null) return null
  const n = Number.parseInt(val, 10)
  return Number.isNaN(n) ? null : n
}

/** Stable string key for grouping runs by signature (Map/object key safe). */
export function signatureToKey(sig: FormattingSignature): string {
  return [
    sig.fontFamily ?? '∅',
    sig.fontSizeHalfPt ?? '∅',
    sig.colorValue,
    sig.bold,
    sig.italic,
    sig.underline ?? '∅',
    sig.strike,
  ].join('|')
}
