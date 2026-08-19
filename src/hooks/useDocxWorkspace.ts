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
import { mergeParagraphStyle, mergeStyles } from '../lib/ooxml/mergeStyles'
import { parseDocx } from '../lib/ooxml/parseDocx'
import {
  materializeReferenceDocStyles,
  reconcileUserStylesOnReferenceDocRemoval,
} from '../lib/ooxml/referenceDocStyles'
import { serializeDocx } from '../lib/ooxml/serializeDocx'
import { buildStyleReport, collectRunRefsForVariantIds, findVariantById } from '../lib/ooxml/styleReport'
import { applyXmlFragmentToRunRefs } from '../lib/ooxml/xmlFragmentEdit'

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
        }
      } catch (err) {
        return { ...state, mergeError: err instanceof Error ? err.message : 'Merge failed.' }
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
    },
  }
}
