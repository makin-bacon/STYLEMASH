import JSZip from 'jszip'
import type { ParsedDocx } from '../../types/ooxml'
import { DOCX_PATHS, NS } from './constants'
import { removeWChild, setWAttr, wAttr, wChild } from './domUtils'
import { findStyleElementById } from './mergeStyles'
import { MIME_TYPES, serializePart } from './serializeDocx'
import { buildStylesMap } from './styleResolution'

export interface ContentMergeOptions {
  /** true (default) = leave every leftover Document-A-origin w:pStyle/w:rStyle
   * reference pointing at a (gap-filled, if needed) copy of A's own style
   * definition - falling back to Document B's definition only when B
   * already defines that exact styleId (B is never overwritten). false =
   * "snap to Document B's style of the same name" - rewrite the reference
   * to Document B's own styleId for a same-type, same-name style, when one
   * exists, before falling back to the same gap-fill behavior. */
  keepOriginalFormatting: boolean
}

/** Clones Document B's documentXml/stylesXml and replaces the clone's body
 * content with Document A's, while preserving Document B's own page setup
 * (its trailing <w:sectPr>) untouched. Pure - mutates neither input
 * document's live state, so this can be re-run freely (e.g. toggling
 * `keepOriginalFormatting` and re-merging). */
export function transplantContent(
  sourceDocx: ParsedDocx,
  referenceDocx: ParsedDocx,
  options: ContentMergeOptions,
): { documentXml: XMLDocument; stylesXml: XMLDocument } {
  const documentXml = referenceDocx.documentXml.cloneNode(true) as XMLDocument
  const stylesXml = referenceDocx.stylesXml.cloneNode(true) as XMLDocument

  const body = documentXml.getElementsByTagNameNS(NS.w, 'body')[0]
  if (!body) throw new Error("Document B is missing a <w:body> - can't merge content into it.")
  const sourceBody = sourceDocx.documentXml.getElementsByTagNameNS(NS.w, 'body')[0]
  if (!sourceBody) throw new Error('Document A has no body content to merge.')

  // 1. Detach (don't discard) B's own trailing sectPr - it holds page setup
  //    (margins/headers/footers/section props) and, per CT_Body, must end
  //    up as body's LAST child again once we're done.
  const trailingSectPr = wChild(body, 'sectPr')
  if (trailingSectPr) body.removeChild(trailingSectPr)

  // 2. Clear everything else B's clone currently has as body content.
  while (body.firstChild) body.removeChild(body.firstChild)

  // 3. Import every direct child of A's body, in order.
  for (const sourceChild of Array.from(sourceBody.children)) {
    // Never transplant A's OWN trailing sectPr - A's page setup must not
    // survive; only B's (re-appended in step 4) should.
    if (sourceChild.namespaceURI === NS.w && sourceChild.localName === 'sectPr') continue

    const imported = documentXml.importNode(sourceChild, true) as Element
    if (imported.namespaceURI === NS.w && imported.localName === 'p') {
      // Strip any embedded mid-document section break (w:pPr/w:sectPr) -
      // that's A's own page setup for one of A's internal sections, which
      // must not leak into the output either. The rest of that paragraph's
      // own pPr (pStyle, alignment, ...) is left untouched.
      const pPr = wChild(imported, 'pPr')
      if (pPr) removeWChild(pPr, 'sectPr')
    }
    body.appendChild(imported)
  }

  // 4. Re-append B's sectPr as body's true last child (CT_Body requires: zero
  //    or more block-level elements, then optionally exactly one trailing sectPr).
  if (trailingSectPr) body.appendChild(trailingSectPr)

  reconcileStyleReferences(sourceDocx, referenceDocx, stylesXml, body, options)

  return { documentXml, stylesXml }
}

function reconcileStyleReferences(
  sourceDocx: ParsedDocx,
  referenceDocx: ParsedDocx,
  clonedStylesXml: XMLDocument,
  body: Element,
  options: ContentMergeOptions,
): void {
  const stylesMapA = buildStylesMap(sourceDocx.stylesXml)
  const stylesMapB = buildStylesMap(referenceDocx.stylesXml) // live B, read-only

  // key: `${'paragraph'|'character'}::${name}` - type-scoped so a same-named
  // paragraph style and character style in B can't be confused for a match.
  const nameToIdB = new Map<string, string>()
  for (const style of stylesMapB.values()) {
    if (style.type === 'paragraph' || style.type === 'character') {
      nameToIdB.set(`${style.type}::${style.name}`, style.id)
    }
  }

  const aStylesRoot = sourceDocx.stylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
  const clonedStylesRoot = clonedStylesXml.getElementsByTagNameNS(NS.w, 'styles')[0]
  if (!aStylesRoot || !clonedStylesRoot) return // defensive; parseDocx always provides a <w:styles> root

  const resolvedIds = new Set<string>() // ids already confirmed present (B's own, or already gap-filled)

  function ensureStyleChainDefined(styleId: string, visiting: Set<string>): void {
    if (resolvedIds.has(styleId) || visiting.has(styleId)) return
    if (findStyleElementById(clonedStylesRoot!, styleId)) {
      resolvedIds.add(styleId)
      return // B already defines this id - NEVER overwrite (B is canonical)
    }
    const aStyleEl = findStyleElementById(aStylesRoot!, styleId)
    if (!aStyleEl) {
      resolvedIds.add(styleId)
      return // dangling in A too - nothing to copy
    }

    visiting.add(styleId)
    const basedOnId = wAttr(wChild(aStyleEl, 'basedOn'), 'val')
    if (basedOnId) ensureStyleChainDefined(basedOnId, visiting) // recursively gap-fill ancestors too
    visiting.delete(styleId)

    clonedStylesRoot!.appendChild(clonedStylesXml.importNode(aStyleEl, true))
    resolvedIds.add(styleId)
  }

  function reconcileOneReference(el: Element, kind: 'paragraph' | 'character'): void {
    const refId = wAttr(el, 'val')
    if (!refId) return
    const aStyle = stylesMapA.get(refId)
    if (!aStyle) return // dangling reference already in A - not ours to fix

    if (!options.keepOriginalFormatting) {
      const bMatchId = nameToIdB.get(`${kind}::${aStyle.name}`)
      if (bMatchId) {
        setWAttr(el, 'val', bMatchId) // "snap" - repoint at B's own same-name style
        return
      }
      // no name match in B - fall through to gap-fill under A's own id
    }
    ensureStyleChainDefined(refId, new Set())
  }

  for (const el of Array.from(body.getElementsByTagNameNS(NS.w, 'pStyle'))) reconcileOneReference(el, 'paragraph')
  for (const el of Array.from(body.getElementsByTagNameNS(NS.w, 'rStyle'))) reconcileOneReference(el, 'character')
}

/** JSZip#clone() shares the same `.files` object by reference with the
 * original (shallow copy), so writing a new part into the "clone" would
 * mutate the original zip's file table too - silently corrupting Document
 * B's live, reused-elsewhere ParsedDocx.zip. This does a real shallow copy
 * of just the *dictionary* instead, so writing new entries into `out` never
 * touches `zip`; the untouched JSZipObject values are safely shared by
 * reference since they're only ever read. */
function cloneZipForOutput(zip: JSZip): JSZip {
  const out = new JSZip()
  out.files = { ...zip.files }
  return out
}

export function buildContentMergedFilename(
  originalFilename: string,
  extension: ParsedDocx['originalExtension'],
): string {
  const dotIndex = originalFilename.lastIndexOf('.')
  const baseName = dotIndex === -1 ? originalFilename : originalFilename.slice(0, dotIndex)
  return `${baseName}-MERGED.${extension}`
}

/** Produces a downloadable Document B, with Document A's body content
 * merged in and Document B's page setup untouched. Every zip part other
 * than word/document.xml and word/styles.xml (theme, headers/footers,
 * settings, numbering, media, _rels, [Content_Types].xml) is a shared,
 * untouched reference from Document B's own zip - this is what actually
 * guarantees "Document B's page setup is unchanged", not any special-casing
 * of sectPr alone.
 *
 * Known v1 scope gaps (mirrors the rest of the app's documented limits):
 * w:numId/numbering.xml references aren't reconciled (a transplanted
 * numbered/bulleted paragraph may render with wrong/missing list
 * formatting); hyperlink/image relationship ids (w:hyperlink/@r:id,
 * w:drawing r:embed) aren't reconciled either, since only Document B's
 * _rels/media survive; w:themeColor on direct-formatting runs resolves
 * against Document B's theme, not Document A's. */
export async function buildContentMergedDocx(
  sourceDocx: ParsedDocx,
  referenceDocx: ParsedDocx,
  options: ContentMergeOptions,
): Promise<{ blob: Blob; filename: string }> {
  const { documentXml, stylesXml } = transplantContent(sourceDocx, referenceDocx, options)

  const outputZip = cloneZipForOutput(referenceDocx.zip)
  outputZip.file(DOCX_PATHS.document, serializePart(documentXml))
  outputZip.file(DOCX_PATHS.styles, serializePart(stylesXml))

  const blob = await outputZip.generateAsync({
    type: 'blob',
    mimeType: MIME_TYPES[referenceDocx.originalExtension],
  })

  return {
    blob,
    filename: buildContentMergedFilename(referenceDocx.originalFilename, referenceDocx.originalExtension),
  }
}
