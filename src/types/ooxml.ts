import type JSZip from 'jszip'

/** A parsed, in-memory representation of an uploaded .docx/.dotx file.
 *
 * `documentXml` and `stylesXml` are live XMLDocuments that get mutated in
 * place for the whole session (merges and XML-fragment edits write directly
 * into them). They are never replaced wholesale, so Element references
 * handed out elsewhere (see RunRef) stay valid across operations. */
export interface ParsedDocx {
  zip: JSZip
  documentXml: XMLDocument
  stylesXml: XMLDocument
  /** Read-only; only consulted to resolve w:themeColor references. Never rewritten. */
  themeXml: XMLDocument | null
  /** Read-only; only consulted to resolve list numbering (see numbering.ts)
   * for the Document Preview and Style Report's list markers. Never rewritten -
   * StyleMash doesn't edit numbering definitions. Null when the docx has
   * no word/numbering.xml part (i.e. no lists at all). */
  numberingXml: XMLDocument | null
  originalFilename: string
  originalExtension: 'docx' | 'dotx'
}

/** The resolved, "how it actually looks" formatting of a run - the 7
 * attributes StyleMash tracks, groups by, and can merge. This is the
 * canonical grouping key for the Style Report: two runs with an identical
 * signature are considered "the same style" regardless of how each one
 * arrived at it (named style vs. direct formatting). */
export interface FormattingSignature {
  fontFamily: string | null
  fontSizeHalfPt: number | null
  /** '#RRGGBB' (always uppercase, no leading '#'... see signature.ts) or 'auto'. */
  colorValue: string
  bold: boolean
  italic: boolean
  /** w:u/@w:val, e.g. "single" | "double" | "wave" ... or null for "no underline". */
  underline: string | null
  strike: boolean
}

/** Where a run's effective formatting actually came from - used only for
 * the Style Report's display breakdown, not for grouping. */
export type StyleOrigin =
  | { kind: 'direct' }
  | { kind: 'named-character'; styleId: string; styleName: string }
  | { kind: 'named-paragraph'; styleId: string; styleName: string }

/** A live pointer back into the parsed document DOM. Merges and XML-fragment
 * edits mutate these exact elements - never re-queried by text or position. */
export interface RunRef {
  runElement: Element
  paragraphElement: Element
  origin: StyleOrigin
}

/** A selectable sub-row within a Style Report entry: every run that shares
 * both the entry's resolved visual signature AND the same origin (e.g. "all
 * the runs that look this way because of direct formatting" vs "...because
 * of the 'Heading 1' style"). Selecting/merging happens at this level so a
 * style-derived instance and a direct-override instance that happen to look
 * identical can be cleaned up independently. */
export interface StyleEntityVariant {
  /** Stable id: `${entity.id}::${originKind}[:styleId]` - the selection key. */
  id: string
  origin: StyleOrigin
  occurrenceCount: number
  /** First non-empty run text found for this variant, truncated for display. */
  sampleText: string
  runRefs: RunRef[]
}

/** One row in the Style Report: every run sharing a resolved FormattingSignature,
 * broken down into origin-based variants (see StyleEntityVariant). */
export interface StyleEntity {
  /** Stable key derived from the signature (see signature.ts#signatureToKey). */
  id: string
  signature: FormattingSignature
  /** Sum of every variant's occurrenceCount. */
  occurrenceCount: number
  /** First non-empty run text found across all variants, truncated for display. */
  sampleText: string
  /** Sorted most-common first. Almost always length 1; >1 when the same
   * look is reached via more than one path (e.g. a style plus, separately,
   * direct formatting elsewhere in the document). */
  variants: StyleEntityVariant[]
}

/** Whether a UserStyleRecord is a <w:style w:type="character"> (applied to
 * runs via w:rStyle - the only kind StyleMash could create before list
 * styles) or a <w:style w:type="paragraph"> (applied to whole paragraphs via
 * w:pStyle, which is what a bullet/numbered list - or any other
 * paragraph-level style - requires). */
export type UserStyleKind = 'character' | 'paragraph'

/** The list numbering a paragraph style can be created with. Only the two
 * kinds a user actually picks in MergeDialog - 'none' means an ordinary
 * paragraph style with no list attached. This is deliberately narrower than
 * numbering.ts's own numFmt vocabulary (which also resolves
 * lowerLetter/upperRoman/etc. for *reading* markers from documents that
 * already have them) - StyleMash only ever *creates* plain bullets or
 * decimal numbering. */
export type ListFormat = 'none' | 'bullet' | 'decimal'

/** A style StyleMash created (or is reusing) via a merge, shown in the
 * "User-Created Styles" panel. */
export interface UserStyleRecord {
  styleId: string
  name: string
  targetSignature: FormattingSignature
  /** 'character' (the original, only kind) or 'paragraph' (see
   * UserStyleKind) - drives which of mergeStyles()/mergeParagraphStyle() a
   * merge into this record uses, and how UserStylesPanel/MergeDialog present it. */
  kind: UserStyleKind
  /** Only meaningful when kind === 'paragraph' - 'none' for every character
   * style (character styles have no concept of list numbering). */
  listFormat: ListFormat
  /** Precomputed preview marker text for this style's list, e.g. "1.1." for
   * a Document B heading style nested two levels deep in a multilevel list
   * (see numbering.ts#buildStylePreviewMarker) - shown by UserStylesPanel
   * instead of re-deriving a marker from `listFormat` alone, since
   * `listFormat` only tracks the flat bullet/decimal shape StyleMash's own
   * createListNumId() can actually create, not a materialized style's true
   * depth. Undefined when listFormat is 'none'; a document occurrence's own
   * live-resolved marker still takes precedence over this once one exists. */
  listPreviewText?: string
  createdAt: number
  /** True when this record was materialized from a style Document B
   * actually uses, at the moment Document B was attached (see
   * referenceDocStyles.ts#materializeReferenceDocStyles). Drives the "from
   * Document B" badge and the removal-cleanup rule
   * (reconcileUserStylesOnReferenceDocRemoval): a record that still carries
   * this flag and has zero occurrences is deleted outright when Document B
   * is detached; one with occurrences > 0 just loses the flag and becomes
   * an ordinary user-owned style. */
  fromReferenceDoc?: true
}

/** A resolved entry from styles.xml - paragraph or character styles only;
 * table/numbering styles are read but never contribute to run formatting. */
export interface StyleDef {
  id: string
  type: 'paragraph' | 'character' | 'table' | 'numbering'
  name: string
  basedOnId: string | null
  rPrElement: Element | null
  /** Only populated for lookups that need paragraph-level properties (e.g.
   * numbering.ts resolving an inherited <w:numPr>) - resolveRunFormatting
   * never reads this. */
  pPrElement: Element | null
  isDefault: boolean
}
