import { NS, RPR_CHILD_ORDER } from './constants'

// Thin, namespace-aware wrappers around the raw DOM API. Word's XML always
// uses the "w:" prefix, but we look elements/attributes up by namespace URI
// (not prefix) since a technically-valid docx could use a different prefix.

export function wChild(parent: Element | null, localName: string): Element | null {
  if (!parent) return null
  const found = parent.getElementsByTagNameNS(NS.w, localName)
  // getElementsByTagNameNS returns descendants, not just direct children;
  // filter to direct children since e.g. w:rPr and w:rPr-inside-w:pPr can
  // both contain unrelated same-named descendants in edge cases.
  for (let i = 0; i < found.length; i++) {
    if (found[i].parentNode === parent) return found[i]
  }
  return null
}

export function wChildren(parent: Element, localName: string): Element[] {
  const out: Element[] = []
  const found = parent.getElementsByTagNameNS(NS.w, localName)
  for (let i = 0; i < found.length; i++) {
    if (found[i].parentNode === parent) out.push(found[i])
  }
  return out
}

export function wAttr(el: Element | null, localName: string): string | null {
  if (!el) return null
  return el.getAttributeNS(NS.w, localName) ?? el.getAttribute(`w:${localName}`)
}

export function createWEl(doc: Document, localName: string): Element {
  return doc.createElementNS(NS.w, `w:${localName}`)
}

export function setWAttr(el: Element, localName: string, value: string): void {
  el.setAttributeNS(NS.w, `w:${localName}`, value)
}

/**
 * Inserts `newChild` into `rPr` at the position required by the CT_RPr
 * schema's fixed child order (RPR_CHILD_ORDER), replacing an existing child
 * of the same local name if one is present. This is the single choke point
 * every code path (merge, style creation, XML-fragment edits) must use when
 * touching <w:rPr> - getting the order wrong is what causes Word to flag a
 * file as needing repair on open.
 */
export function insertRPrChildInOrder(rPr: Element, localName: string, newChild: Element): void {
  const existing = wChild(rPr, localName)
  if (existing) {
    rPr.replaceChild(newChild, existing)
    return
  }

  const targetIndex = RPR_CHILD_ORDER.indexOf(localName as (typeof RPR_CHILD_ORDER)[number])
  if (targetIndex === -1) {
    // Unknown/unsupported element name - safest fallback is to append.
    rPr.appendChild(newChild)
    return
  }

  let insertBefore: Element | null = null
  for (const child of Array.from(rPr.children)) {
    const childIndex = RPR_CHILD_ORDER.indexOf(child.localName as (typeof RPR_CHILD_ORDER)[number])
    if (childIndex === -1) continue // unknown child, ignore for ordering purposes
    if (childIndex > targetIndex) {
      insertBefore = child
      break
    }
  }
  rPr.insertBefore(newChild, insertBefore)
}

export function removeWChild(parent: Element, localName: string): void {
  const existing = wChild(parent, localName)
  if (existing) parent.removeChild(existing)
}

/** True if `doc` failed to parse (DOMParser embeds a <parsererror> instead of throwing). */
export function hasParseError(doc: Document): boolean {
  return doc.getElementsByTagName('parsererror').length > 0
}
