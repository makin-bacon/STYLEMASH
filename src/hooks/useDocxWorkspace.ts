import { useCallback, useMemo, useReducer } from 'react'
import type {
  FormattingSignature,
  ListFormat,
  ParsedDocx,
  StyleEntity,
  UserStyleKind,
  UserStyleRecord,
} from '../types/ooxml'
import { downloadBlob } from '../lib/download'
import {
  bulkMergeVariantsIntoMatchingReferenceStyles,
  findVariantIdsMatchingReferenceStyleNames,
} from '../lib/ooxml/bulkMergeMatchedStyles'
import { buildContentMergedDocx, type ContentMergeOptions } from '../lib/ooxml/contentMerge'
import { addDefaultStyles } from '../lib/ooxml/defaultStyles'
import { mergeParagraphStyle, mergeStyles, removeStyleById } from '../lib/ooxml/mergeStyles'
import { parseDocx } from '../lib/ooxml/parseDocx'
import {
  materializeReferenceDocStyles,
  reconcileUserStylesOnReferenceDocRemoval,
} from '../lib/ooxml/referenceDocStyles'
import { serializeDocx, serializePart } from '../lib/ooxml/serializeDocx'
import { buildStyleReport, collectRunRefsForVariantIds, findVariantById } from '../lib/ooxml/styleReport'
import { applyXmlFragmentToRunRefs } from '../lib/ooxml/xmlFragmentEdit'

/** A pre-mutation snapshot of everything a merge (or a full "Clear list")
 * touches, pushed onto WorkspaceState.undoStack right before the mutating
 * call - documentXml/stylesXml/numberingXml are mutated in place (see the
 * reducer's top-of-file note), so the only way to snapshot them is to
 * serialize to text and re-parse on restore, rather than keeping a second
 * live reference to the same, about-to-be-mutated Elements. */
interface UndoSnapshot {
  documentXml: string
  stylesXml: string
  numberingXml: string | null
  userStyles: UserStyleRecord[]
}

/** Keyed by the exact `parsedDocx` object a snapshot was taken from, so a
 * second call for the *same* pre-dispatch state (React 18 StrictMode
 * deliberately invokes a reducer twice per dispatch in dev, to surface
 * impure reducers - see the top-of-file note on this one's mutate-in-place
 * escape hatch) returns the pristine snapshot captured on the first call
 * instead of re-serializing documentXml/stylesXml *after* that first call's
 * mergeStyles()/etc. already mutated them. Both invocations receive the
 * identical `state` object from React, so `state.parsedDocx` is the same
 * reference either way - only the live Elements inside it have (or haven't
 * yet) been mutated. Entries fall out of the map for GC on their own, since
 * every mutating action replaces `parsedDocx` with a new wrapper object
 * afterward. */
const pristineSnapshotCache = new WeakMap<ParsedDocx, UndoSnapshot>()

function snapshotForUndo(parsedDocx: ParsedDocx, userStyles: UserStyleRecord[]): UndoSnapshot {
  const cached = pristineSnapshotCache.get(parsedDocx)
  if (cached) return cached
  const snapshot: UndoSnapshot = {
    documentXml: serializePart(parsedDocx.documentXml),
    stylesXml: serializePart(parsedDocx.stylesXml),
    numberingXml: parsedDocx.numberingXml ? serializePart(parsedDocx.numberingXml) : null,
    userStyles,
  }
  pristineSnapshotCache.set(parsedDocx, snapshot)
  return snapshot
}

function restoreSnapshot(parsedDocx: ParsedDocx, snapshot: UndoSnapshot): ParsedDocx {
  const parser = new DOMParser()
  return {
    ...parsedDocx,
    documentXml: parser.parseFromString(snapshot.documentXml, 'application/xml'),
    stylesXml: parser.parseFromString(snapshot.stylesXml, 'application/xml'),
    numberingXml: snapshot.numberingXml
      ? parser.parseFromString(snapshot.numberingXml, 'application/xml')
      : null,
  }
}

export interface ReferenceDocState {
  status: 'empty' | 'loading' | 'loaded' | 'error'
  errorMessage: string | null
  parsedDocx: ParsedDocx | null
}

const initialReferenceDocState: ReferenceDocState = {
  status: 'empty',
  errorMessage: null,
  parsedDocx: null,
}

export interface WorkspaceState {
  status: 'empty' | 'loading' | 'loaded' | 'error'
  errorMessage: string | null
  parsedDocx: ParsedDocx | null
  styleReport: StyleEntity[]
  userStyles: UserStyleRecord[]
  /** Selection operates at the variant level (see StyleEntityVariant) so a
   * style-derived instance and a direct-override instance of the same
   * visual look can be selected/merged independently. */
  selectedVariantIds: Set<string>
  /** The single User-Created style currently picked as a merge target (a
   * click on its row in UserStylesPanel, distinct from that row's "Edit"
   * button) - lets selectedVariantIds be folded into it directly via
   * MERGE_SELECTED_INTO_TARGET, without going through MergeDialog. */
  selectedTargetStyleId: string | null
  activeEditVariantId: string | null
  mergeDialogOpen: boolean
  /** Set when the merge dialog was opened to edit/extend an existing
   * UserStyleRecord (via UserStylesPanel's "Edit" button) rather than to
   * create a brand-new style from a fresh Style Report selection. */
  mergeDialogReuseStyleId: string | null
  mergeError: string | null
  xmlEditorError: string | null
  isSaving: boolean
  /** Document B: a second, separate document attached purely as a text
   * style reference (see referenceDocStyles.ts/contentMerge.ts). Entirely
   * independent of the merge/XML-editor modal state above. */
  referenceDoc: ReferenceDocState
  contentMergeDialogOpen: boolean
  isMergingContent: boolean
  contentMergeError: string | null
  /** Error from the last BULK_MERGE_MATCHED_TO_REFERENCE - surfaced inline
   * near the Style Report's bulk-merge control rather than in a modal,
   * since that action has no modal of its own. */
  bulkMergeError: string | null
  /** One entry per undoable action (CONFIRM_MERGE, MERGE_SELECTED_INTO_TARGET,
   * BULK_MERGE_MATCHED_TO_REFERENCE, CLEAR_USER_STYLES), oldest first - UNDO
   * pops the last one and restores it. Deliberately doesn't cover style-only
   * actions with no document effect until merged into (ADD_DEFAULT_STYLES,
   * attaching/removing Document B) - see snapshotForUndo's call sites. */
  undoStack: UndoSnapshot[]
}

const initialState: WorkspaceState = {
  status: 'empty',
  errorMessage: null,
  parsedDocx: null,
  styleReport: [],
  userStyles: [],
  selectedVariantIds: new Set(),
  selectedTargetStyleId: null,
  activeEditVariantId: null,
  mergeDialogOpen: false,
  mergeDialogReuseStyleId: null,
  mergeError: null,
  xmlEditorError: null,
  isSaving: false,
  referenceDoc: initialReferenceDocState,
  contentMergeDialogOpen: false,
  isMergingContent: false,
  contentMergeError: null,
  bulkMergeError: null,
  undoStack: [],
}

type Action =
  | { type: 'LOADING_STARTED' }
  | { type: 'FILE_LOADED'; parsedDocx: ParsedDocx; styleReport: StyleEntity[] }
  | { type: 'PARSE_ERROR'; message: string }
  | { type: 'TOGGLE_SELECT_VARIANT'; variantId: string }
  | { type: 'CLEAR_SELECTION' }
  | { type: 'OPEN_MERGE_DIALOG'; reuseExistingStyleId?: string }
  | { type: 'CLOSE_MODALS' }
  | {
      type: 'CONFIRM_MERGE'
      targetProps: FormattingSignature
      name: string
      kind: UserStyleKind
      listFormat: ListFormat
      reuseExistingStyleId?: string
    }
  | { type: 'OPEN_XML_EDITOR'; variantId: string }
  | { type: 'APPLY_XML_EDIT'; fragmentText: string }
  | { type: 'SAVING_STARTED' }
  | { type: 'SAVING_FINISHED' }
  | { type: 'REFERENCE_DOC_LOADING_STARTED' }
  | { type: 'REFERENCE_DOC_LOADED'; parsedDocx: ParsedDocx; userStyles: UserStyleRecord[] }
  | { type: 'REFERENCE_DOC_LOAD_ERROR'; message: string }
  | { type: 'REMOVE_REFERENCE_DOC' }
  | { type: 'OPEN_CONTENT_MERGE_DIALOG' }
  | { type: 'CONTENT_MERGE_STARTED' }
  | { type: 'CONTENT_MERGE_FINISHED' }
  | { type: 'CONTENT_MERGE_ERROR'; message: string }
  | { type: 'SELECT_VARIANTS'; variantIds: string[] }
  | { type: 'BULK_MERGE_MATCHED_TO_REFERENCE' }
  | { type: 'RESET' }
  | { type: 'TOGGLE_SELECT_TARGET_STYLE'; styleId: string }
  | { type: 'MERGE_SELECTED_INTO_TARGET' }
  | { type: 'ADD_DEFAULT_STYLES' }
  | { type: 'CLEAR_USER_STYLES' }
  | { type: 'UNDO' }

/** Note on the reducer's relationship to immutability: `parsedDocx`'s inner
 * XMLDocuments (documentXml/stylesXml) are mutated in place by mergeStyles()
 * and applyXmlFragmentToRunRefs() rather than treated as immutable data - a
 * deliberate escape hatch. Re-parsing the whole XML tree on every micro-edit
 * would be wasteful and would invalidate the very live Element references
 * (RunRef) the next operation needs. The reducer still returns a *new*
 * `parsedDocx` wrapper object on every mutation purely so React's shallow
 * comparison re-renders, even though the Document instances inside it are
 * the same mutated objects. */
function reducer(state: WorkspaceState, action: Action): WorkspaceState {
  switch (action.type) {
    case 'LOADING_STARTED':
      return { ...initialState, status: 'loading' }

    case 'FILE_LOADED':
      return {
        ...initialState,
        status: 'loaded',
        parsedDocx: action.parsedDocx,
        styleReport: action.styleReport,
      }

    case 'PARSE_ERROR':
      return { ...initialState, status: 'error', errorMessage: action.message }

    case 'TOGGLE_SELECT_VARIANT': {
      const next = new Set(state.selectedVariantIds)
      if (next.has(action.variantId)) next.delete(action.variantId)
      else next.add(action.variantId)
      return { ...state, selectedVariantIds: next }
    }

    case 'CLEAR_SELECTION':
      return { ...state, selectedVariantIds: new Set() }

    case 'OPEN_MERGE_DIALOG':
      return {
        ...state,
        mergeDialogOpen: true,
        mergeError: null,
        mergeDialogReuseStyleId: action.reuseExistingStyleId ?? null,
      }

    case 'CLOSE_MODALS':
      return {
        ...state,
        mergeDialogOpen: false,
        mergeDialogReuseStyleId: null,
        activeEditVariantId: null,
        mergeError: null,
        xmlEditorError: null,
        contentMergeDialogOpen: false,
        contentMergeError: null,
      }

    case 'CONFIRM_MERGE': {
      if (!state.parsedDocx) return state
      // Empty selection is valid: it's how "+ New Style" and "redefine this
      // style's look" (no newly-selected variants to fold in) both work -
      // mergeStyles() just creates/redefines the style definition itself.
      const sourceRunRefs = collectRunRefsForVariantIds(state.styleReport, state.selectedVariantIds)
      const undoSnapshot = snapshotForUndo(state.parsedDocx, state.userStyles)

      try {
        const styleId =
          action.kind === 'paragraph'
            ? mergeParagraphStyle(
                state.parsedDocx,
                sourceRunRefs,
                action.targetProps,
                action.name,
                action.listFormat,
                action.reuseExistingStyleId,
              )
            : mergeStyles(
                state.parsedDocx,
                sourceRunRefs,
                action.targetProps,
                action.name,
                action.reuseExistingStyleId,
              )
        const styleReport = buildStyleReport(state.parsedDocx)

        const existingIndex = state.userStyles.findIndex((r) => r.styleId === styleId)
        const listFormat = action.kind === 'paragraph' ? action.listFormat : 'none'
        const record: UserStyleRecord = {
          styleId,
          name: action.name,
          targetSignature: action.targetProps,
          kind: action.kind,
          listFormat,
          // Manually-created lists are always single-level (createListNumId
          // only ever creates one), so this is just the ilvl-0 marker - the
          // richer multilevel preview (e.g. "1.1.") is exclusive to styles
          // materialized from Document B (see referenceDocStyles.ts).
          listPreviewText: listFormat === 'bullet' ? '•' : listFormat === 'decimal' ? '1.' : undefined,
          createdAt: existingIndex === -1 ? Date.now() : state.userStyles[existingIndex].createdAt,
        }
        const userStyles =
          existingIndex === -1
            ? [...state.userStyles, record]
            : state.userStyles.map((r, i) => (i === existingIndex ? record : r))

        return {
          ...state,
          parsedDocx: { ...state.parsedDocx },
          styleReport,
          userStyles,
          selectedVariantIds: new Set(),
          selectedTargetStyleId: null,
          mergeDialogOpen: false,
          mergeDialogReuseStyleId: null,
          mergeError: null,
          undoStack: [...state.undoStack, undoSnapshot],
        }
      } catch (err) {
        return { ...state, mergeError: err instanceof Error ? err.message : 'Merge failed.' }
      }
    }

    case 'OPEN_XML_EDITOR':
      return { ...state, activeEditVariantId: action.variantId, xmlEditorError: null }

    case 'APPLY_XML_EDIT': {
      if (!state.parsedDocx || !state.activeEditVariantId) return state
      const variant = findVariantById(state.styleReport, state.activeEditVariantId)
      if (!variant) return state

      try {
        applyXmlFragmentToRunRefs(state.parsedDocx, variant.runRefs, action.fragmentText)
        const styleReport = buildStyleReport(state.parsedDocx)
        return {
          ...state,
          parsedDocx: { ...state.parsedDocx },
          styleReport,
          activeEditVariantId: null,
          xmlEditorError: null,
        }
      } catch (err) {
        return { ...state, xmlEditorError: err instanceof Error ? err.message : 'Could not apply XML.' }
      }
    }

    case 'SAVING_STARTED':
      return { ...state, isSaving: true }

    case 'SAVING_FINISHED':
      return { ...state, isSaving: false }

    case 'REFERENCE_DOC_LOADING_STARTED':
      return { ...state, referenceDoc: { status: 'loading', errorMessage: null, parsedDocx: null } }

    case 'REFERENCE_DOC_LOADED':
      return {
        ...state,
        referenceDoc: { status: 'loaded', errorMessage: null, parsedDocx: action.parsedDocx },
        // Already the full replacement list (materializeReferenceDocStyles
        // folds a duplicate-named style into its existing record - same
        // styleId, redefined look - rather than appending a second one; see
        // referenceDocStyles.ts), so this assigns directly instead of
        // spreading onto state.userStyles.
        userStyles: action.userStyles,
        // parsedDocx (A) was mutated in place by materializeReferenceDocStyles.
        // styleReport DOES need recomputing here (unlike a plain "new style
        // with no occurrences yet" used to be able to skip it): a
        // duplicate-name collision redefines an *existing* <w:style>'s
        // rPr/pPr in place, and any run in A already merged into that style
        // needs its Style Report signature (and therefore its Document
        // Preview look) refreshed to match the newest file's definition.
        parsedDocx: state.parsedDocx ? { ...state.parsedDocx } : null,
        styleReport: state.parsedDocx ? buildStyleReport(state.parsedDocx) : state.styleReport,
      }

    case 'REFERENCE_DOC_LOAD_ERROR':
      return { ...state, referenceDoc: { status: 'error', errorMessage: action.message, parsedDocx: null } }

    case 'REMOVE_REFERENCE_DOC': {
      if (!state.parsedDocx) return { ...state, referenceDoc: initialReferenceDocState }
      const userStyles = reconcileUserStylesOnReferenceDocRemoval(
        state.userStyles,
        state.styleReport,
        state.parsedDocx.stylesXml,
      )
      return {
        ...state,
        parsedDocx: { ...state.parsedDocx },
        userStyles,
        referenceDoc: initialReferenceDocState,
        // styleReport unchanged for the same reason as REFERENCE_DOC_LOADED:
        // only removing zero-occurrence style *definitions*.
        // Clear a dangling target selection if the style it pointed to was
        // one of the just-removed zero-occurrence ones.
        selectedTargetStyleId: userStyles.some((r) => r.styleId === state.selectedTargetStyleId)
          ? state.selectedTargetStyleId
          : null,
      }
    }

    case 'OPEN_CONTENT_MERGE_DIALOG':
      return { ...state, contentMergeDialogOpen: true, contentMergeError: null }

    case 'CONTENT_MERGE_STARTED':
      return { ...state, isMergingContent: true, contentMergeError: null }

    case 'CONTENT_MERGE_FINISHED':
      return { ...state, isMergingContent: false, contentMergeDialogOpen: false }

    case 'CONTENT_MERGE_ERROR':
      return { ...state, isMergingContent: false, contentMergeError: action.message }

    case 'SELECT_VARIANTS': {
      const next = new Set(state.selectedVariantIds)
      for (const id of action.variantIds) next.add(id)
      return { ...state, selectedVariantIds: next, bulkMergeError: null }
    }

    case 'BULK_MERGE_MATCHED_TO_REFERENCE': {
      if (!state.parsedDocx) return state
      const undoSnapshot = snapshotForUndo(state.parsedDocx, state.userStyles)
      try {
        bulkMergeVariantsIntoMatchingReferenceStyles(
          state.parsedDocx,
          state.styleReport,
          state.selectedVariantIds,
          state.userStyles,
        )
        const styleReport = buildStyleReport(state.parsedDocx)
        return {
          ...state,
          parsedDocx: { ...state.parsedDocx },
          styleReport,
          selectedVariantIds: new Set(),
          bulkMergeError: null,
          undoStack: [...state.undoStack, undoSnapshot],
        }
      } catch (err) {
        return { ...state, bulkMergeError: err instanceof Error ? err.message : 'Bulk merge failed.' }
      }
    }

    case 'RESET':
      return initialState

    case 'TOGGLE_SELECT_TARGET_STYLE':
      return {
        ...state,
        selectedTargetStyleId: state.selectedTargetStyleId === action.styleId ? null : action.styleId,
      }

    case 'MERGE_SELECTED_INTO_TARGET': {
      if (!state.parsedDocx || !state.selectedTargetStyleId) return state
      const targetRecord = state.userStyles.find((r) => r.styleId === state.selectedTargetStyleId)
      if (!targetRecord) return state

      const sourceRunRefs = collectRunRefsForVariantIds(state.styleReport, state.selectedVariantIds)
      const undoSnapshot = snapshotForUndo(state.parsedDocx, state.userStyles)
      try {
        if (targetRecord.kind === 'paragraph') {
          mergeParagraphStyle(
            state.parsedDocx,
            sourceRunRefs,
            targetRecord.targetSignature,
            targetRecord.name,
            targetRecord.listFormat,
            targetRecord.styleId,
          )
        } else {
          mergeStyles(
            state.parsedDocx,
            sourceRunRefs,
            targetRecord.targetSignature,
            targetRecord.name,
            targetRecord.styleId,
          )
        }
        const styleReport = buildStyleReport(state.parsedDocx)
        return {
          ...state,
          parsedDocx: { ...state.parsedDocx },
          styleReport,
          selectedVariantIds: new Set(),
          selectedTargetStyleId: null,
          mergeError: null,
          undoStack: [...state.undoStack, undoSnapshot],
        }
      } catch (err) {
        return { ...state, mergeError: err instanceof Error ? err.message : 'Merge failed.' }
      }
    }

    case 'ADD_DEFAULT_STYLES': {
      if (!state.parsedDocx) return state
      const userStyles = addDefaultStyles(state.parsedDocx, state.userStyles)
      return {
        ...state,
        parsedDocx: { ...state.parsedDocx },
        userStyles,
        // A name collision with an existing style redefines its rPr/pPr in
        // place (see addDefaultStyles) - same reasoning as REFERENCE_DOC_LOADED
        // for recomputing this: any run already merged into that style needs
        // its signature refreshed to match the newest look.
        styleReport: buildStyleReport(state.parsedDocx),
      }
    }

    case 'CLEAR_USER_STYLES': {
      if (!state.parsedDocx || state.userStyles.length === 0) return state
      const undoSnapshot = snapshotForUndo(state.parsedDocx, state.userStyles)
      for (const record of state.userStyles) {
        removeStyleById(state.parsedDocx.stylesXml, record.styleId)
      }
      return {
        ...state,
        parsedDocx: { ...state.parsedDocx },
        styleReport: buildStyleReport(state.parsedDocx),
        userStyles: [],
        selectedTargetStyleId: null,
        undoStack: [...state.undoStack, undoSnapshot],
      }
    }

    case 'UNDO': {
      if (!state.parsedDocx || state.undoStack.length === 0) return state
      const snapshot = state.undoStack[state.undoStack.length - 1]
      const parsedDocx = restoreSnapshot(state.parsedDocx, snapshot)
      return {
        ...state,
        parsedDocx,
        styleReport: buildStyleReport(parsedDocx),
        userStyles: snapshot.userStyles,
        undoStack: state.undoStack.slice(0, -1),
        selectedVariantIds: new Set(),
        selectedTargetStyleId: null,
        mergeError: null,
        bulkMergeError: null,
      }
    }

    default:
      return state
  }
}

/** Owns the entire StyleMash workspace: the parsed document, the derived
 * Style Report, user-created (merged) styles, and all selection/modal UI
 * state. A single reducer (rather than several useStates) because a merge
 * is one atomic transaction that must update selection, the report, and the
 * user-styles list together. */
export function useDocxWorkspace() {
  const [state, dispatch] = useReducer(reducer, initialState)

  const loadFile = useCallback(async (file: File) => {
    dispatch({ type: 'LOADING_STARTED' })
    try {
      const parsedDocx = await parseDocx(file)
      const styleReport = buildStyleReport(parsedDocx)
      dispatch({ type: 'FILE_LOADED', parsedDocx, styleReport })
    } catch (err) {
      dispatch({
        type: 'PARSE_ERROR',
        message: err instanceof Error ? err.message : 'Failed to load this file.',
      })
    }
  }, [])

  const toggleSelectVariant = useCallback((variantId: string) => {
    dispatch({ type: 'TOGGLE_SELECT_VARIANT', variantId })
  }, [])

  const clearSelection = useCallback(() => dispatch({ type: 'CLEAR_SELECTION' }), [])
  const openMergeDialog = useCallback((reuseExistingStyleId?: string) => {
    dispatch({ type: 'OPEN_MERGE_DIALOG', reuseExistingStyleId })
  }, [])
  const closeModals = useCallback(() => dispatch({ type: 'CLOSE_MODALS' }), [])

  const confirmMerge = useCallback(
    (
      targetProps: FormattingSignature,
      name: string,
      kind: UserStyleKind,
      listFormat: ListFormat,
      reuseExistingStyleId?: string,
    ) => {
      dispatch({ type: 'CONFIRM_MERGE', targetProps, name, kind, listFormat, reuseExistingStyleId })
    },
    [],
  )

  const openXmlEditor = useCallback((variantId: string) => {
    dispatch({ type: 'OPEN_XML_EDITOR', variantId })
  }, [])

  const applyXmlEdit = useCallback((fragmentText: string) => {
    dispatch({ type: 'APPLY_XML_EDIT', fragmentText })
  }, [])

  const save = useCallback(async () => {
    if (!state.parsedDocx) return
    dispatch({ type: 'SAVING_STARTED' })
    try {
      const { blob, filename } = await serializeDocx(state.parsedDocx)
      downloadBlob(blob, filename)
    } finally {
      dispatch({ type: 'SAVING_FINISHED' })
    }
  }, [state.parsedDocx])

  const loadReferenceDoc = useCallback(
    async (file: File) => {
      if (!state.parsedDocx) return
      dispatch({ type: 'REFERENCE_DOC_LOADING_STARTED' })
      try {
        const referenceParsedDocx = await parseDocx(file)
        const userStyles = materializeReferenceDocStyles(state.parsedDocx, referenceParsedDocx, state.userStyles)
        dispatch({ type: 'REFERENCE_DOC_LOADED', parsedDocx: referenceParsedDocx, userStyles })
      } catch (err) {
        dispatch({
          type: 'REFERENCE_DOC_LOAD_ERROR',
          message: err instanceof Error ? err.message : 'Failed to load this file.',
        })
      }
    },
    [state.parsedDocx, state.userStyles],
  )

  const removeReferenceDoc = useCallback(() => dispatch({ type: 'REMOVE_REFERENCE_DOC' }), [])
  const openContentMergeDialog = useCallback(() => dispatch({ type: 'OPEN_CONTENT_MERGE_DIALOG' }), [])

  const mergeContentIntoReferenceDoc = useCallback(
    async (keepOriginalFormatting: boolean) => {
      if (!state.parsedDocx || !state.referenceDoc.parsedDocx) return
      dispatch({ type: 'CONTENT_MERGE_STARTED' })
      try {
        const options: ContentMergeOptions = { keepOriginalFormatting }
        const { blob, filename } = await buildContentMergedDocx(
          state.parsedDocx,
          state.referenceDoc.parsedDocx,
          options,
        )
        downloadBlob(blob, filename)
        dispatch({ type: 'CONTENT_MERGE_FINISHED' })
      } catch (err) {
        dispatch({
          type: 'CONTENT_MERGE_ERROR',
          message: err instanceof Error ? err.message : 'Could not merge content into Document B.',
        })
      }
    },
    [state.parsedDocx, state.referenceDoc.parsedDocx],
  )

  const selectVariantsMatchingReferenceStyles = useCallback(() => {
    const variantIds = Array.from(
      findVariantIdsMatchingReferenceStyleNames(state.styleReport, state.userStyles),
    )
    dispatch({ type: 'SELECT_VARIANTS', variantIds })
  }, [state.styleReport, state.userStyles])

  const bulkMergeMatchedToReference = useCallback(() => {
    dispatch({ type: 'BULK_MERGE_MATCHED_TO_REFERENCE' })
  }, [])

  /** Returns to the landing screen in-app (no browser reload) - used by
   * StyleReportPanel's "Mash a different file" button, so the persistent
   * header/footer never flicker/remount along the way. */
  const reset = useCallback(() => dispatch({ type: 'RESET' }), [])

  const toggleSelectTargetStyle = useCallback((styleId: string) => {
    dispatch({ type: 'TOGGLE_SELECT_TARGET_STYLE', styleId })
  }, [])

  const mergeSelectedIntoTarget = useCallback(() => dispatch({ type: 'MERGE_SELECTED_INTO_TARGET' }), [])

  const addDefaultStylesAction = useCallback(() => dispatch({ type: 'ADD_DEFAULT_STYLES' }), [])
  const clearUserStyles = useCallback(() => dispatch({ type: 'CLEAR_USER_STYLES' }), [])
  const undo = useCallback(() => dispatch({ type: 'UNDO' }), [])

  // Summary of the current selection, for MergeDialog's prefill/messaging -
  // exposed as derived totals rather than raw entities/variants so the
  // dialog stays decoupled from the report's grouping shape.
  const selection = useMemo(() => {
    let totalOccurrences = 0
    let baselineSignature: FormattingSignature | null = null
    for (const entity of state.styleReport) {
      for (const variant of entity.variants) {
        if (state.selectedVariantIds.has(variant.id)) {
          totalOccurrences += variant.occurrenceCount
          if (!baselineSignature) baselineSignature = entity.signature
        }
      }
    }
    return { totalOccurrences, baselineSignature }
  }, [state.styleReport, state.selectedVariantIds])

  const activeEditVariant = useMemo(
    () => (state.activeEditVariantId ? findVariantById(state.styleReport, state.activeEditVariantId) : null),
    [state.styleReport, state.activeEditVariantId],
  )

  return {
    state,
    selection,
    activeEditVariant,
    actions: {
      loadFile,
      toggleSelectVariant,
      clearSelection,
      openMergeDialog,
      closeModals,
      confirmMerge,
      openXmlEditor,
      applyXmlEdit,
      save,
      loadReferenceDoc,
      removeReferenceDoc,
      openContentMergeDialog,
      mergeContentIntoReferenceDoc,
      selectVariantsMatchingReferenceStyles,
      bulkMergeMatchedToReference,
      reset,
      toggleSelectTargetStyle,
      mergeSelectedIntoTarget,
      addDefaultStyles: addDefaultStylesAction,
      clearUserStyles,
      undo,
    },
  }
}
