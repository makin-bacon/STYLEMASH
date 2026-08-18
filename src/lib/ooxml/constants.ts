// XML namespace URIs used across Word (WordprocessingML) and DrawingML (theme) parts.
export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  // DrawingML "main" namespace - only used to read theme1.xml's <a:clrScheme>
  // when resolving w:themeColor references.
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  // OPC package-level namespaces - only used by serializeDocx.ts to register
  // a brand-new word/numbering.xml part (see ensureNumberingPartRegistered),
  // never touched when a document already ships one of its own.
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
  relationships: 'http://schemas.openxmlformats.org/package/2006/relationships',
} as const

// The zip paths we read/write. Headers, footers, and footnotes/endnotes are
// intentionally out of scope for v1 - see plan doc for rationale.
export const DOCX_PATHS = {
  document: 'word/document.xml',
  styles: 'word/styles.xml',
  theme: 'word/theme/theme1.xml',
  numbering: 'word/numbering.xml',
  contentTypes: '[Content_Types].xml',
  documentRels: 'word/_rels/document.xml.rels',
} as const

// Local (unprefixed) names of the <w:rPr> children StyleMash fully owns:
// the Style Report groups text by these, and a merge fully subsumes them
// (creates a style that sets all of them, then strips them from the runs
// it applies to). Everything else on a run's <w:rPr> is left untouched.
export const TRACKED_RPR_LOCAL_NAMES = [
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'strike',
  'color',
  'sz',
  'szCs',
  'u',
] as const

// The fixed child-element order required by the CT_RPr schema (ECMA-376 /
// ISO-29500 §17.3.2). Word (and most validators) will show a "needs repair"
// prompt if <w:rPr> children appear out of this order, so every code path
// that inserts an rPr child must insert at the schema-correct position
// rather than appending blindly. See rPrHelpers.ts.
export const RPR_CHILD_ORDER = [
  'rStyle',
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
  'specVanish',
  'oMath',
] as const
