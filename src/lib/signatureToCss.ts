import type { CSSProperties } from 'react'
import type { FormattingSignature } from '../types/ooxml'

/** Cheap, high-value live preview: renders a resolved FormattingSignature as
 * inline CSS so the Style Report / Merge Dialog can show roughly what the
 * text actually looks like, without needing a real Word rendering engine. */
export function signatureToCss(sig: FormattingSignature): CSSProperties {
  const decorations: string[] = []
  if (sig.underline) decorations.push('underline')
  if (sig.strike) decorations.push('line-through')

  return {
    fontFamily: sig.fontFamily ?? undefined,
    fontSize: sig.fontSizeHalfPt ? `${sig.fontSizeHalfPt / 2}pt` : undefined,
    color: sig.colorValue === 'auto' ? undefined : `#${sig.colorValue}`,
    fontWeight: sig.bold ? 'bold' : 'normal',
    fontStyle: sig.italic ? 'italic' : 'normal',
    textDecorationLine: decorations.length > 0 ? decorations.join(' ') : 'none',
  }
}
