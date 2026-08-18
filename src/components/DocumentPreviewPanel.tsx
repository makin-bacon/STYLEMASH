import type { CSSProperties } from 'react'
import { useEffect, useMemo, useRef, useState } from 'react'
import type { ReferenceDocState } from '../hooks/useDocxWorkspace'
import type { FormattingSignature, ParsedDocx, StyleEntity } from '../types/ooxml'
import { NS } from '../lib/ooxml/constants'
import type { ParagraphMarker } from '../lib/ooxml/numbering'
import { collectRunRefsForVariantIds, getRunText } from '../lib/ooxml/styleReport'
import { signatureToCss } from '../lib/signatureToCss'
import { InfoTooltip } from './InfoTooltip'
import { SaveButton } from './SaveButton'

interface DocumentPreviewPanelProps {
  parsedDocx: ParsedDocx | null
  styleReport: StyleEntity[]
  /** Same selection Style Report rows drive for merging - reused here so
   * checking a row both stages it for "Do it" and highlights every place it
   * occurs in the live preview, with no separate UI to keep in sync. */
  selectedVariantIds: Set<string>
  /** Resolved list marker ("1.", "b)", "•"...) per paragraph, from the same
   * pass StyleReportPanel uses - keeps a numbered/bulleted paragraph looking
   * like a list here too, instead of silently dropping its marker. */
  paragraphMarkers: Map<Element, ParagraphMarker>
  /** Document B state - this panel's footer shows "Merge content into
   * Document B…" once loaded; attaching/removing Document B itself happens
   * via UserStylesPanel's footer (Attach when empty, Remove once loaded). */
  referenceDoc: ReferenceDocState
  isMergingContent: boolean
  onOpenContentMerge: () => void
  isSaving: boolean
  onSave: () => void
}

interface PreviewRun {
  key: string
  text: string
  runElement: Element
  css: CSSProperties
}

interface PreviewParagraph {
  key: string
  paragraphElement: Element
  runs: PreviewRun[]
}

const FLASH_DURATION_MS = 1400

/** signatureToCss() renders the document's actual font name verbatim, which
 * is fine for the Style Report's tiny sample lines but looks broken here
 * where a whole document's worth of text is rendered: most custom/corporate
 * fonts aren't installed in a browser, so the browser silently falls back to
 * its generic default anyway. Layering standard web-safe fallbacks after the
 * real name gets closer to how the font would actually look, and keeps the
 * preview rendering consistently instead of occasionally substituting an
 * unrelated system font. */
function previewCss(signature: FormattingSignature): CSSProperties {
  const family = signature.fontFamily?.trim()
  return {
    ...signatureToCss(signature),
    fontFamily: family
      ? `"${family}", Calibri, "Segoe UI", Arial, sans-serif`
      : 'Calibri, "Segoe UI", Arial, sans-serif',
  }
}

/** Left-hand panel: a read-only, best-effort rendering of word/document.xml's
 * visible text, styled per-run from the same live signatures the Style
 * Report is built from - so merges (which mutate the document's Elements in
 * place) show up here immediately, without a separate render pipeline to
 * keep in sync. Table layout/images/page geometry are intentionally not
 * reproduced; this is a formatting preview, not a document renderer. */
export function DocumentPreviewPanel({
  parsedDocx,
  styleReport,
  selectedVariantIds,
  paragraphMarkers,
  referenceDoc,
  isMergingContent,
  onOpenContentMerge,
  isSaving,
  onSave,
}: DocumentPreviewPanelProps) {
  const containerRef = useRef<HTMLDivElement>(null)
  const runNodesRef = useRef(new Map<Element, HTMLElement>())

  // Per-run resolved CSS, keyed by the same live Element instances RunRef
  // points at - built once per styleReport recompute rather than per-render.
  const runCss = useMemo(() => {
    const map = new Map<Element, CSSProperties>()
    for (const entity of styleReport) {
      const css = previewCss(entity.signature)
      for (const variant of entity.variants) {
        for (const ref of variant.runRefs) map.set(ref.runElement, css)
      }
    }
    return map
  }, [styleReport])

  const paragraphs = useMemo<PreviewParagraph[]>(() => {
    if (!parsedDocx) return []
    const paragraphEls = parsedDocx.documentXml.getElementsByTagNameNS(NS.w, 'p')
    const paras: PreviewParagraph[] = []
    for (let i = 0; i < paragraphEls.length; i++) {
      const paragraphEl = paragraphEls[i]
      const runEls = paragraphEl.getElementsByTagNameNS(NS.w, 'r')
      const runs: PreviewRun[] = []
      for (let j = 0; j < runEls.length; j++) {
        const runEl = runEls[j]
        const text = getRunText(runEl)
        if (text.length === 0) continue
        runs.push({ key: `${i}-${j}`, text, runElement: runEl, css: runCss.get(runEl) ?? {} })
      }
      paras.push({ key: `p-${i}`, paragraphElement: paragraphEl, runs })
    }
    return paras
  }, [parsedDocx, runCss])

  const highlightedRunElements = useMemo(() => {
    if (selectedVariantIds.size === 0) return new Set<Element>()
    return new Set(collectRunRefsForVariantIds(styleReport, selectedVariantIds).map((r) => r.runElement))
  }, [styleReport, selectedVariantIds])

  // Briefly flashes the runs that were just folded into a style, so a
  // successful merge reads as a visible change here rather than just a
  // silent re-style. Detected rather than event-driven: a merge is exactly
  // the transition from "some variants selected" to "none selected" that
  // also mutated the document (parsedDocx's wrapper is a fresh object on
  // every mutation - see useDocxWorkspace's reducer notes) - CLEAR_SELECTION
  // clears selection without that, so it doesn't trigger a false flash.
  const [flashedRunElements, setFlashedRunElements] = useState<Set<Element>>(new Set())
  const prevSelectedIdsRef = useRef(selectedVariantIds)
  const prevStyleReportRef = useRef(styleReport)
  const prevParsedDocxRef = useRef(parsedDocx)

  useEffect(() => {
    const prevSelected = prevSelectedIdsRef.current
    const prevReport = prevStyleReportRef.current
    const docChanged = prevParsedDocxRef.current !== parsedDocx
    prevSelectedIdsRef.current = selectedVariantIds
    prevStyleReportRef.current = styleReport
    prevParsedDocxRef.current = parsedDocx

    if (!docChanged || prevSelected.size === 0 || selectedVariantIds.size > 0) return
    const elements = new Set(collectRunRefsForVariantIds(prevReport, prevSelected).map((r) => r.runElement))
    if (elements.size === 0) return
    setFlashedRunElements(elements)
    const timer = setTimeout(() => setFlashedRunElements(new Set()), FLASH_DURATION_MS)
    return () => clearTimeout(timer)
  }, [selectedVariantIds, styleReport, parsedDocx])

  // Scrolls the first newly-selected occurrence into view, so picking a
  // Style Report row that's off-screen doesn't leave the user hunting for it.
  const prevScrolledIdsRef = useRef<Set<string>>(new Set())
  useEffect(() => {
    const newlyAdded = [...selectedVariantIds].filter((id) => !prevScrolledIdsRef.current.has(id))
    prevScrolledIdsRef.current = selectedVariantIds
    if (newlyAdded.length === 0) return

    for (const entity of styleReport) {
      for (const variant of entity.variants) {
        if (!newlyAdded.includes(variant.id) || variant.runRefs.length === 0) continue
        const node = runNodesRef.current.get(variant.runRefs[0].runElement)
        node?.scrollIntoView({ behavior: 'smooth', block: 'center' })
        return
      }
    }
  }, [selectedVariantIds, styleReport])

  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-hidden rounded-lg border border-slate-200 bg-white">
      <div className="flex items-start justify-between gap-2 border-b border-slate-200 bg-slate-50 px-4 py-4">
        <h2 className="flex items-center gap-1.5 text-sm font-semibold text-slate-700">
          Document Preview
          <InfoTooltip
            text={`${parsedDocx?.originalFilename ?? 'Live preview'} — This is a "style only" preview of your document. It will not display your page flow correctly but that's OK, that's not what this tool is for. To merge your style with approved styles, use the panels to the left.`}
          />
        </h2>
        <SaveButton disabled={!parsedDocx} isSaving={isSaving} onSave={onSave} />
      </div>

      <div ref={containerRef} className="min-h-0 flex-1 overflow-y-auto px-4 py-3">
        {paragraphs.length === 0 && (
          <p className="flex h-full items-center justify-center text-center text-sm text-slate-400">
            No visible text found in this document.
          </p>
        )}
        {paragraphs.map((para) => {
          const marker = paragraphMarkers.get(para.paragraphElement)
          return (
            <p
              key={para.key}
              className="mb-2 flex min-h-[1.25em] gap-2 text-sm leading-relaxed text-slate-800"
              style={marker ? { paddingLeft: `${marker.ilvl * 1.25}em` } : undefined}
            >
              {marker?.text && <span className="shrink-0 select-none text-slate-500">{marker.text}</span>}
              <span>
                {para.runs.length === 0
                  ? '\u00A0'
                  : para.runs.map((run) => {
                      const isFlashed = flashedRunElements.has(run.runElement)
                      const isHighlighted = highlightedRunElements.has(run.runElement)
                      return (
                        <span
                          key={run.key}
                          ref={(node) => {
                            if (node) runNodesRef.current.set(run.runElement, node)
                            else runNodesRef.current.delete(run.runElement)
                          }}
                          style={run.css}
                          className={`rounded-sm transition-colors duration-700 ${
                            isFlashed
                              ? 'bg-emerald-200 ring-2 ring-emerald-400'
                              : isHighlighted
                                ? 'bg-indigo-200 ring-2 ring-indigo-400'
                                : 'bg-transparent'
                          }`}
                        >
                          {run.text}
                        </span>
                      )
                    })}
              </span>
            </p>
          )
        })}
      </div>

      {/* Always rendered (never conditionally mounted) so this row's height
          never changes as Document B is attached/removed - when there's no
          "Merge content..." button to show, an invisible one of the same
          size holds its place instead of collapsing the row, so the preview
          area's height above it never visibly resizes on that interaction. */}
      <div className="flex shrink-0 flex-wrap items-center justify-end gap-2 border-t border-slate-200 bg-slate-50 px-4 py-2">
        <button
          type="button"
          onClick={onOpenContentMerge}
          disabled={isMergingContent || referenceDoc.status !== 'loaded'}
          aria-hidden={referenceDoc.status !== 'loaded'}
          tabIndex={referenceDoc.status === 'loaded' ? 0 : -1}
          className={`rounded-md bg-indigo-600 px-3 py-1.5 text-xs font-medium text-white enabled:hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300 ${
            referenceDoc.status === 'loaded' ? '' : 'invisible'
          }`}
        >
          Merge content into Document B…
        </button>
      </div>
    </div>
  )
}
