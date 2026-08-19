import type { FormattingSignature, ListFormat, ParsedDocx, UserStyleKind, UserStyleRecord } from '../../types/ooxml'
import { mergeParagraphStyle, mergeStyles } from './mergeStyles'

interface DefaultStyleDefinition {
  name: string
  targetSignature: FormattingSignature
  kind: UserStyleKind
  listFormat: ListFormat
  listPreviewText?: string
}

/** StyleMash's bundled starter style set - the 17 styles actually applied
 * somewhere in a one-off reference document (TEST-DOC/STYLE-REFERENCE-FILE.docx,
 * examined and then deleted; see project history) rather than every style
 * merely *defined* in it. Each signature is that document's own fully-cascaded
 * look (through its basedOn chain to its own docDefaults/theme), captured once
 * so "+ Defaults" needs no file of its own at runtime. */
export const DEFAULT_STYLES: DefaultStyleDefinition[] = [
  {
    name: 'heading 1',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 40,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1',
  },
  {
    name: 'heading 2',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 32,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.1',
  },
  {
    name: 'heading 3',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 28,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.1.1',
  },
  {
    name: 'heading 4',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 24,
      colorValue: '0F4761',
      bold: false,
      italic: true,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.1.1.1',
  },
  {
    name: 'Normal Bold',
    targetSignature: {
      fontFamily: 'Arial',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'List Bullet',
    targetSignature: {
      fontFamily: 'Arial',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'bullet',
    listPreviewText: '•',
  },
  {
    name: 'List Number',
    targetSignature: {
      fontFamily: 'Arial',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.',
  },
  {
    name: 'caption',
    targetSignature: {
      fontFamily: 'Arial',
      fontSizeHalfPt: 20,
      colorValue: 'auto',
      bold: false,
      italic: true,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'Hyperlink',
    targetSignature: {
      fontFamily: 'Arial',
      fontSizeHalfPt: 24,
      colorValue: '467886',
      bold: true,
      italic: false,
      underline: 'single',
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'Heading 1 No Numbering',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 40,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'Heading 2 No Numbering',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 32,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'Heading 3 No Numbering',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 28,
      colorValue: '0F4761',
      bold: true,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'Heading 4 No Numbering',
    targetSignature: {
      fontFamily: null,
      fontSizeHalfPt: 24,
      colorValue: '0F4761',
      bold: false,
      italic: true,
      underline: null,
      strike: false,
    },
    kind: 'character',
    listFormat: 'none',
  },
  {
    name: 'List Bullet 2',
    targetSignature: {
      fontFamily: 'Arial',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'bullet',
    listPreviewText: '•',
  },
  {
    name: 'List Number 2',
    targetSignature: {
      fontFamily: 'Arial',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.',
  },
  {
    name: 'List Number 3',
    targetSignature: {
      fontFamily: 'Arial',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'decimal',
    listPreviewText: '1.',
  },
  {
    name: 'List Bullet 3',
    targetSignature: {
      fontFamily: 'Arial',
      fontSizeHalfPt: 24,
      colorValue: 'auto',
      bold: false,
      italic: false,
      underline: null,
      strike: false,
    },
    kind: 'paragraph',
    listFormat: 'bullet',
    listPreviewText: '•',
  },
]

/** Materializes DEFAULT_STYLES into `targetDocx` via the same "+ New Style"
 * code path (mergeStyles()/mergeParagraphStyle() with no source runs) every
 * other User-Created style goes through - so a default is byte-for-byte the
 * same shape, and behaves identically as a merge target, as any
 * manually-created or Document-B-imported style. A name collision with an
 * existing UserStyleRecord redefines that record's own styleId (newest
 * wins) instead of creating a duplicate - the same rule referenceDocStyles.ts
 * applies for Document B, so re-clicking "+ Defaults" after editing a
 * default's look just resets it rather than piling up a second entry.
 * Mutates targetDocx.stylesXml (and numberingXml, for a list default) in
 * place. Returns the full replacement for `existingUserStyles`. */
export function addDefaultStyles(
  targetDocx: ParsedDocx,
  existingUserStyles: UserStyleRecord[],
): UserStyleRecord[] {
  const existingByName = new Map(existingUserStyles.map((r) => [r.name, r]))
  const replacements = new Map<string, UserStyleRecord>()
  const brandNew: UserStyleRecord[] = []

  for (const def of DEFAULT_STYLES) {
    const collision = existingByName.get(def.name)
    const newStyleId =
      def.kind === 'character'
        ? mergeStyles(targetDocx, [], def.targetSignature, def.name, collision?.styleId)
        : mergeParagraphStyle(
            targetDocx,
            [],
            def.targetSignature,
            def.name,
            def.listFormat,
            collision?.styleId,
          )

    const record: UserStyleRecord = {
      styleId: newStyleId,
      name: def.name,
      targetSignature: def.targetSignature,
      kind: def.kind,
      listFormat: def.listFormat,
      listPreviewText: def.listPreviewText,
      createdAt: collision?.createdAt ?? Date.now(),
    }

    if (collision) {
      replacements.set(collision.styleId, record)
    } else {
      brandNew.push(record)
    }
  }

  const merged = existingUserStyles.map((r) => replacements.get(r.styleId) ?? r)
  return [...merged, ...brandNew]
}
