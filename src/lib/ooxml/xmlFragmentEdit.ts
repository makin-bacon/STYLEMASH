import type { ParsedDocx, RunRef } from '../../types/ooxml'
import { NS } from './constants'
import { hasParseError, wChild } from './domUtils'

export class XmlFragmentError extends Error {}

/** Returns the first run's <w:rPr> (or an empty placeholder if it has none)
 * serialized for display in the "simple interface" XML editor - just a
 * <textarea>, per the user's own framing; the fragments involved are small
 * enough that syntax highlighting isn't worth the added dependency. */
export function serializeFirstRunRPr(runRefs: RunRef[]): string {
  const firstRun = runRefs[0]?.runElement
  const rPr = firstRun ? wChild(firstRun, 'rPr') : null
  if (!rPr) return '<w:rPr/>'
  return prettyPrint(new XMLSerializer().serializeToString(rPr))
}

// Browsers ship no XML pretty-printer; this is a cheap, well-known trick
// that's good enough for the small fragments involved here.
function prettyPrint(xml: string): string {
  return xml.replace(/></g, '>\n<')
}

/** Parses and validates a user-edited XML fragment, e.g. `<w:rPr>...</w:rPr>`.
 * The fragment can't be parsed standalone because its `w:` prefix needs a
 * namespace declaration, which normally lives on the real document's root -
 * so it's wrapped in a throwaway root that declares it. Throws
 * XmlFragmentError with a message suitable for showing directly to the user. */
export function parseXmlFragment(text: string, expectedTag: 'rPr' | 'pPr' = 'rPr'): Element {
  const wrapped = `<w:root xmlns:w="${NS.w}">${text}</w:root>`
  const doc = new DOMParser().parseFromString(wrapped, 'application/xml')
  if (hasParseError(doc)) {
    throw new XmlFragmentError('That XML could not be parsed - check for unclosed tags or typos.')
  }

  const children = Array.from(doc.documentElement.children)
  if (children.length !== 1) {
    throw new XmlFragmentError(`Expected exactly one <w:${expectedTag}> element.`)
  }
  const [child] = children
  if (child.localName !== expectedTag || child.namespaceURI !== NS.w) {
    throw new XmlFragmentError(`Expected a <w:${expectedTag}> element, found <${child.tagName}>.`)
  }
  return child
}

/** Applies a hand-edited <w:rPr> fragment to every run in `runRefs`,
 * replacing each run's existing <w:rPr> wholesale (or inserting one if it
 * had none). Caller should recompute buildStyleReport() afterward - the
 * edited runs' signature has likely changed, so they need to regroup. */
export function applyXmlFragmentToRunRefs(
  parsedDocx: ParsedDocx,
  runRefs: RunRef[],
  fragmentXmlText: string,
): void {
  const parsedFragment = parseXmlFragment(fragmentXmlText, 'rPr')
  // Parsed in a separate Document (the wrapper doc) - must be imported into
  // the live document once, then cloned per target run.
  const imported = parsedDocx.documentXml.importNode(parsedFragment, true)

  for (const runRef of runRefs) {
    const clone = imported.cloneNode(true) as Element
    const existingRPr = wChild(runRef.runElement, 'rPr')
    if (existingRPr) {
      runRef.runElement.replaceChild(clone, existingRPr)
    } else {
      runRef.runElement.insertBefore(clone, runRef.runElement.firstChild)
    }
  }
}
