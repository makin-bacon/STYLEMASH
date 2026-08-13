import type { FormattingSignature, StyleOrigin } from '../types/ooxml'

export function describeSignature(signature: FormattingSignature): string {
  const parts = [
    signature.fontFamily ?? 'default font',
    signature.fontSizeHalfPt ? `${signature.fontSizeHalfPt / 2}pt` : null,
    signature.colorValue !== 'auto' ? `#${signature.colorValue}` : null,
    signature.bold ? 'bold' : null,
    signature.italic ? 'italic' : null,
    signature.underline ? `${signature.underline} underline` : null,
    signature.strike ? 'strikethrough' : null,
  ]
  return parts.filter(Boolean).join(' · ')
}

export function describeOrigin(origin: StyleOrigin): string {
  if (origin.kind === 'direct') return 'Direct formatting (not part of a defined style)'
  return `Via "${origin.styleName}" style`
}
