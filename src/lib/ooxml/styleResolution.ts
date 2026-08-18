import type { StyleDef, StyleOrigin, FormattingSignature } from '../../types/ooxml'
import { NS, TRACKED_RPR_LOCAL_NAMES } from './constants'
import { wAttr, wChild, wChildren } from './domUtils'
import { trackedChildrenToSignature } from './signature'
import { resolveColorElement } from './themeColor'

type TrackedMap = Map<string, Element>

/** Indexes every <w:style> in styles.xml by its styleId. Only paragraph and
 * character styles are relevant to run formatting (table/numbering styles
 * are read here for completeness but never consulted by resolveRunFormatting). */
export function buildStylesMap(stylesXml: XMLDocument): Map<string, StyleDef> {
  const map = new Map<string, StyleDef>()
  const stylesRoot = stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
  if (!stylesRoot) return map

  for (const styleEl of wChildren(stylesRoot, 'style')) {
    const id = wAttr(styleEl, 'styleId')
    const type = wAttr(styleEl, 'type')
    if (!id || (type !== 'paragraph' && type !== 'character' && type !== 'table' && type !== 'numbering')) {
      continue
    }
    map.set(id, {
      id,
      type,
      name: wAttr(wChild(styleEl, 'name'), 'val') ?? id,
      basedOnId: wAttr(wChild(styleEl, 'basedOn'), 'val'),
      rPrElement: wChild(styleEl, 'rPr'),
      pPrElement: wChild(styleEl, 'pPr'),
      isDefault: wAttr(styleEl, 'default') === '1' || wAttr(styleEl, 'default') === 'true',
    })
  }
  return map
}

export function getDocDefaultsRPr(stylesXml: XMLDocument): Element | null {
  const stylesRoot = stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
  const docDefaults = stylesRoot ? wChild(stylesRoot, 'docDefaults') : null
  const rPrDefault = docDefaults ? wChild(docDefaults, 'rPrDefault') : null
  return rPrDefault ? wChild(rPrDefault, 'rPr') : null
}

export function getDefaultParagraphStyleId(stylesMap: Map<string, StyleDef>): string | null {
  for (const style of stylesMap.values()) {
    if (style.type === 'paragraph' && style.isDefault) return style.id
  }
  return null
}

function extractTracked(rPr: Element | null): TrackedMap {
  const map: TrackedMap = new Map()
  if (!rPr) return map
  for (const localName of TRACKED_RPR_LOCAL_NAMES) {
    const child = wChild(rPr, localName)
    if (child) map.set(localName, child)
  }
  return map
}

function overlay(base: TrackedMap, additions: TrackedMap): TrackedMap {
  const result = new Map(base)
  for (const [k, v] of additions) result.set(k, v)
  return result
}

/** Resolves a style's fully-cascaded tracked <w:rPr> children by walking its
 * w:basedOn chain back to docDefaults. Memoized per styleId (many runs share
 * a style) and guarded against circular basedOn chains, which do occur in
 * malformed real-world documents. */
export function resolveStyleRPr(
  styleId: string | null,
  stylesMap: Map<string, StyleDef>,
  docDefaultsTracked: TrackedMap,
  cache: Map<string, TrackedMap>,
  visiting: Set<string> = new Set(),
): TrackedMap {
  if (!styleId) return docDefaultsTracked
  const cached = cache.get(styleId)
  if (cached) return cached

  if (visiting.has(styleId)) {
    // Circular basedOn chain - stop recursing and fall back to docDefaults
    // rather than infinite-looping on a malformed document.
    return docDefaultsTracked
  }
  visiting.add(styleId)

  const style = stylesMap.get(styleId)
  if (!style) {
    visiting.delete(styleId)
    return docDefaultsTracked
  }

  const base = style.basedOnId
    ? resolveStyleRPr(style.basedOnId, stylesMap, docDefaultsTracked, cache, visiting)
    : docDefaultsTracked

  const resolved = overlay(base, extractTracked(style.rPrElement))
  cache.set(styleId, resolved)
  visiting.delete(styleId)
  return resolved
}

export interface ResolutionContext {
  stylesMap: Map<string, StyleDef>
  docDefaultsTracked: TrackedMap
  defaultParagraphStyleId: string | null
  themeColors: Map<string, string>
  styleRPrCache: Map<string, TrackedMap>
}

export function buildResolutionContext(
  stylesMap: Map<string, StyleDef>,
  docDefaultsRPr: Element | null,
  themeColors: Map<string, string>,
): ResolutionContext {
  return {
    stylesMap,
    docDefaultsTracked: extractTracked(docDefaultsRPr),
    defaultParagraphStyleId: getDefaultParagraphStyleId(stylesMap),
    themeColors,
    styleRPrCache: new Map(),
  }
}

/** Resolves one run's effective formatting by cascading, in increasing
 * precedence: docDefaults -> the run's paragraph style chain -> the run's
 * character style chain (w:rStyle, if any) -> the run's own direct <w:rPr>.
 *
 * v1 scope notes (see plan doc): paragraph-mark run properties
 * (w:pPr/w:rPr) are ignored (only affects the invisible pilcrow); toggle
 * properties use last-writer-wins rather than the spec's full XOR/toggle
 * semantics. */
export function resolveRunFormatting(
  runEl: Element,
  paragraphEl: Element,
  ctx: ResolutionContext,
): { signature: FormattingSignature; origin: StyleOrigin } {
  const pPr = wChild(paragraphEl, 'pPr')
  const pStyleId = wAttr(wChild(pPr, 'pStyle'), 'val') ?? ctx.defaultParagraphStyleId
  const paragraphCascade = resolveStyleRPr(
    pStyleId,
    ctx.stylesMap,
    ctx.docDefaultsTracked,
    ctx.styleRPrCache,
  )

  const directRPr = wChild(runEl, 'rPr')
  const rStyleId = wAttr(wChild(directRPr, 'rStyle'), 'val')
  const directTracked = extractTracked(directRPr)

  let cascade = paragraphCascade

  if (rStyleId) {
    const charCascade = resolveStyleRPr(
      rStyleId,
      ctx.stylesMap,
      ctx.docDefaultsTracked,
      ctx.styleRPrCache,
    )
    cascade = overlay(cascade, charCascade)
  }

  cascade = overlay(cascade, directTracked)

  // Origin reflects the most specific source of the run's *visible*
  // formatting, for the Style Report's "N via 'Emphasis', M direct"
  // breakdown: an explicit direct override beats a character style, which
  // beats the paragraph style chain (which resolves to at least "Normal"
  // for every run, so most unstyled text is correctly attributed there
  // rather than lumped in with genuinely hand-formatted text).
  let origin: StyleOrigin
  if (directTracked.size > 0) {
    origin = { kind: 'direct' }
  } else if (rStyleId) {
    origin = {
      kind: 'named-character',
      styleId: rStyleId,
      styleName: ctx.stylesMap.get(rStyleId)?.name ?? rStyleId,
    }
  } else if (pStyleId) {
    origin = {
      kind: 'named-paragraph',
      styleId: pStyleId,
      styleName: ctx.stylesMap.get(pStyleId)?.name ?? pStyleId,
    }
  } else {
    origin = { kind: 'direct' }
  }

  const signature = trackedChildrenToSignature(cascade, (colorEl) =>
    resolveColorElement(colorEl, ctx.themeColors),
  )

  return { signature, origin }
}
