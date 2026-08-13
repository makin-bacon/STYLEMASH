import type { FormattingSignature, ParsedDocx, RunRef } from '../../types/ooxml'
import { NS } from './constants'
import { createWEl, insertRPrChildInOrder, setWAttr, wAttr, wChild, wChildren } from './domUtils'
import { stripTrackedProps, writeSignatureIntoRPr } from './rPrHelpers'

function collectExistingStyleIds(stylesXml: XMLDocument): Set<string> {
  const ids = new Set<string>()
  const stylesRoot = stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
  if (!stylesRoot) return ids
  for (const styleEl of wChildren(stylesRoot, 'style')) {
    const id = wAttr(styleEl, 'styleId')
    if (id) ids.add(id)
  }
  return ids
}

function slugify(name: string): string {
  const cleaned = name.replace(/[^a-zA-Z0-9]/g, '')
  const base = cleaned.length > 0 ? cleaned : 'Style'
  return /^[0-9]/.test(base) ? `Style${base}` : base
}

/** Generates a styleId guaranteed not to collide with any existing
 * @w:styleId in this document, by slugifying `name` and appending -1, -2...
 * on collision. */
export function generateUniqueStyleId(stylesXml: XMLDocument, name: string): string {
  const existing = collectExistingStyleIds(stylesXml)
  const base = slugify(name)
  if (!existing.has(base)) return base
  let n = 1
  while (existing.has(`${base}${n}`)) n++
  return `${base}${n}`
}

export function findStyleElementById(stylesRoot: Element, styleId: string): Element | null {
  for (const styleEl of wChildren(stylesRoot, 'style')) {
    if (wAttr(styleEl, 'styleId') === styleId) return styleEl
  }
  return null
}

/** Removes a <w:style> by id, if present. Used by the "remove Document B"
 * cleanup (see referenceDocStyles.ts) to fully undo a materialized style
 * that was never actually merged into. Returns whether it removed one. */
export function removeStyleById(stylesXml: XMLDocument, styleId: string): boolean {
  const stylesRoot = stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
  const el = stylesRoot ? findStyleElementById(stylesRoot, styleId) : null
  if (!el || !stylesRoot) return false
  stylesRoot.removeChild(el)
  return true
}

/** Builds a new character-type <w:style>. CT_Style's schema order for the
 * children we use is name, basedOn, rPr - since this element's shape is
 * entirely under our control we hardcode that order directly rather than
 * going through the general rPr-child-ordering helper (that helper only
 * concerns itself with CT_RPr's children). */
function buildStyleElement(
  doc: XMLDocument,
  styleId: string,
  name: string,
  targetProps: FormattingSignature,
  hasDefaultParagraphFont: boolean,
): Element {
  const styleEl = createWEl(doc, 'style')
  setWAttr(styleEl, 'type', 'character')
  setWAttr(styleEl, 'styleId', styleId)

  const nameEl = createWEl(doc, 'name')
  setWAttr(nameEl, 'val', name)
  styleEl.appendChild(nameEl)

  if (hasDefaultParagraphFont) {
    const basedOnEl = createWEl(doc, 'basedOn')
    setWAttr(basedOnEl, 'val', 'DefaultParagraphFont')
    styleEl.appendChild(basedOnEl)
  }

  const rPr = createWEl(doc, 'rPr')
  writeSignatureIntoRPr(rPr, targetProps)
  styleEl.appendChild(rPr)

  return styleEl
}

function updateStyleNameAndRPr(
  styleEl: Element,
  name: string,
  targetProps: FormattingSignature,
): void {
  let nameEl = wChild(styleEl, 'name')
  if (!nameEl) {
    nameEl = createWEl(styleEl.ownerDocument, 'name')
    styleEl.insertBefore(nameEl, styleEl.firstChild)
  }
  setWAttr(nameEl, 'val', name)

  let rPr = wChild(styleEl, 'rPr')
  if (!rPr) {
    rPr = createWEl(styleEl.ownerDocument, 'rPr')
    styleEl.appendChild(rPr)
  }
  writeSignatureIntoRPr(rPr, targetProps)
}

function ensureRPrFirstChild(runEl: Element): Element {
  let rPr = wChild(runEl, 'rPr')
  if (!rPr) {
    rPr = createWEl(runEl.ownerDocument, 'rPr')
    runEl.insertBefore(rPr, runEl.firstChild)
  }
  return rPr
}

/** Points a run at `styleId` and strips the tracked-property children that
 * the style now supplies, per the merge's subsumption rule: everything the
 * app tracks (font/size/color/b/i/u/strike) is fully taken over by the
 * style; anything else on the run's rPr (vertAlign, spacing, lang, ...) is
 * left untouched. */
function applyStyleToRun(runEl: Element, styleId: string): void {
  const rPr = ensureRPrFirstChild(runEl)
  const rStyleEl = createWEl(runEl.ownerDocument, 'rStyle')
  setWAttr(rStyleEl, 'val', styleId)
  // 'rStyle' is first in RPR_CHILD_ORDER, so this also guarantees it lands
  // as rPr's first child, as CT_RPr requires.
  insertRPrChildInOrder(rPr, 'rStyle', rStyleEl)
  stripTrackedProps(rPr)
}

/**
 * Creates (or redefines/reuses) a single named character style and applies
 * it to every run in `sourceRunRefs` via w:rStyle - even runs that started
 * as pure direct/inherited formatting. `sourceRunRefs` is a flat list (not
 * grouped by Style Report entity/variant) so this function stays decoupled
 * from how the report groups things - callers typically build it with
 * collectRunRefsForVariantIds(). Passing an empty array is valid: it just
 * creates/redefines the style definition itself without touching any runs
 * (e.g. the "+ New Style" flow, or reusing an existing style purely to
 * tweak its look).
 *
 * Always creates/updates a character style (never a paragraph style),
 * because all 7 tracked attributes are fundamentally run-level concepts;
 * repointing w:pStyle instead would risk silently altering out-of-scope
 * paragraph properties (alignment, spacing, numbering).
 *
 * Pass `reuseExistingStyleId` to redefine/extend an already-created style
 * (e.g. editing an existing UserStyleRecord, or targeting an existing style
 * from the merge dialog) instead of creating a new one.
 *
 * Mutates parsedDocx.documentXml/stylesXml in place. Caller is expected to
 * recompute buildStyleReport() afterward - the merged runs will naturally
 * regroup into one entity/variant.
 */
export function mergeStyles(
  parsedDocx: ParsedDocx,
  sourceRunRefs: RunRef[],
  targetProps: FormattingSignature,
  name: string,
  reuseExistingStyleId?: string,
): string {
  const { stylesXml } = parsedDocx
  const stylesRoot = stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
  if (!stylesRoot) {
    throw new Error('This document\'s styles.xml is missing a <w:styles> root - cannot merge.')
  }

  let styleId: string

  if (reuseExistingStyleId) {
    const existingStyleEl = findStyleElementById(stylesRoot, reuseExistingStyleId)
    if (!existingStyleEl) {
      throw new Error(`Cannot reuse style "${reuseExistingStyleId}" - it no longer exists.`)
    }
    styleId = reuseExistingStyleId
    updateStyleNameAndRPr(existingStyleEl, name, targetProps)
  } else {
    styleId = generateUniqueStyleId(stylesXml, name)
    const hasDefaultParagraphFont = collectExistingStyleIds(stylesXml).has('DefaultParagraphFont')
    const newStyleEl = buildStyleElement(stylesXml, styleId, name, targetProps, hasDefaultParagraphFont)
    // Appending at the end of <w:styles> is always schema-safe: repeated
    // w:style elements have no relative-order constraint among themselves.
    stylesRoot.appendChild(newStyleEl)
  }

  for (const runRef of sourceRunRefs) {
    applyStyleToRun(runRef.runElement, styleId)
  }

  return styleId
}
