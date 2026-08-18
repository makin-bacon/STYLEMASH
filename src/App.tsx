import { useMemo } from 'react'
import { AppFooter } from './components/AppFooter'
import { AppHeader } from './components/AppHeader'
import { ContentMergeDialog } from './components/ContentMergeDialog'
import { DocumentPreviewPanel } from './components/DocumentPreviewPanel'
import { DropzoneUpload } from './components/DropzoneUpload'
import { MergeDialog } from './components/MergeDialog'
import { StyleReportPanel } from './components/StyleReportPanel'
import { UserStylesPanel } from './components/UserStylesPanel'
import { XmlEditorModal } from './components/XmlEditorModal'
import { useDocxWorkspace } from './hooks/useDocxWorkspace'
import { buildParagraphMarkers } from './lib/ooxml/numbering'
import { computeMergeProgress, filterUnmergedEntities } from './lib/ooxml/styleReport'

/** Top-level app shell. Owns the single useDocxWorkspace instance and
 * switches between the upload screen and the two-panel workspace based on
 * workspace.state.status. All actual OOXML logic lives in src/lib/ooxml/ -
 * this component is purely about wiring state to the presentational pieces.
 *
 * AppHeader/AppFooter are rendered exactly once, outside the status switch,
 * so they stay mounted (and never replay the page-transition animation)
 * across a file upload - only the body content between them swaps, keyed by
 * view so *that* remount/fade is scoped to just the content that's actually
 * changing. */
function App() {
  const { state, selection, activeEditVariant, actions } = useDocxWorkspace()

  const reuseRecord = state.mergeDialogReuseStyleId
    ? (state.userStyles.find((r) => r.styleId === state.mergeDialogReuseStyleId) ?? null)
    : null
  const importedStyleCount = state.userStyles.filter((r) => r.fromReferenceDoc).length
  const isLoaded = state.status === 'loaded'
  const mergeProgress = computeMergeProgress(state.styleReport, state.userStyles)
  // StyleReportPanel only ever shows what's still outstanding - an entry
  // already folded into a User-Created style is done, so leaving it
  // visible would just be clutter (and re-selecting it would be a no-op).
  // Every other consumer of styleReport (DocumentPreviewPanel,
  // UserStylesPanel, mergeProgress above) keeps using the full,
  // unfiltered state.styleReport.
  const unmergedStyleReport = filterUnmergedEntities(state.styleReport, state.userStyles)
  // Shared with both panels below so a list paragraph's marker (bullet,
  // "1.", "b)"...) reads the same way in the live preview and in the Style
  // Report's sample text, rather than each panel resolving numbering itself.
  const paragraphMarkers = useMemo(
    () => (state.parsedDocx ? buildParagraphMarkers(state.parsedDocx) : new Map()),
    [state.parsedDocx],
  )

  return (
    <div className="flex h-full flex-col bg-slate-50">
      <AppHeader filename={isLoaded ? (state.parsedDocx?.originalFilename ?? null) : null} />

      {isLoaded ? (
        <div key="workspace" className="page-transition flex min-h-0 flex-1 flex-col">
          {/* CSS grid with 4 equal-fr columns: Style Report and User-Created
              Styles each take 1 column (25%), Document Preview spans 2
              (50%, unchanged from before this became a 3-column layout) -
              grid's `minmax(0, 1fr)` tracks divide width precisely, which a
              flex row's percentage-basis children can't do as cleanly
              alongside `gap`. h-full min-h-0 on every column is a stretch
              safety net on top of grid's own align-items:stretch default,
              same defensive reasoning as the flex row this replaces: never
              rely on a single implicit mechanism for "fill the row's
              height" when an explicit one is one class away. */}
          <main className="grid min-h-0 flex-1 grid-cols-4 gap-4 p-4">
            <div className="col-span-1 flex h-full min-h-0 min-w-0 flex-col">
              <StyleReportPanel
                styleReport={unmergedStyleReport}
                selectedIds={state.selectedVariantIds}
                paragraphMarkers={paragraphMarkers}
                onToggleSelect={actions.toggleSelectVariant}
                onMergeSelected={() => actions.openMergeDialog()}
                hasReferenceStyles={importedStyleCount > 0}
                bulkMergeError={state.bulkMergeError}
                onSelectMatchingReferenceStyles={actions.selectVariantsMatchingReferenceStyles}
                onBulkMergeMatched={actions.bulkMergeMatchedToReference}
                mergeProgress={mergeProgress}
                onSave={actions.save}
                onRipAnotherFile={actions.reset}
              />
            </div>
            <div className="col-span-1 flex h-full min-h-0 min-w-0 flex-col">
              <UserStylesPanel
                userStyles={state.userStyles}
                styleReport={state.styleReport}
                paragraphMarkers={paragraphMarkers}
                onEditStyle={actions.openMergeDialog}
                onCreateNewStyle={() => actions.openMergeDialog()}
                selectedTargetStyleId={state.selectedTargetStyleId}
                onToggleSelectTarget={actions.toggleSelectTargetStyle}
                pendingSelectionCount={state.selectedVariantIds.size}
                onMergeSelectedIntoTarget={actions.mergeSelectedIntoTarget}
                mergeError={state.mergeError}
                referenceDoc={state.referenceDoc}
                onAttachReferenceDoc={actions.loadReferenceDoc}
                onRemoveReferenceDoc={actions.removeReferenceDoc}
              />
            </div>
            <div className="col-span-2 flex h-full min-h-0 min-w-0 flex-col">
              <DocumentPreviewPanel
                parsedDocx={state.parsedDocx}
                styleReport={state.styleReport}
                selectedVariantIds={state.selectedVariantIds}
                paragraphMarkers={paragraphMarkers}
                referenceDoc={state.referenceDoc}
                isMergingContent={state.isMergingContent}
                onOpenContentMerge={actions.openContentMergeDialog}
                isSaving={state.isSaving}
                onSave={actions.save}
              />
            </div>
          </main>

          {state.mergeDialogOpen && (
            <MergeDialog
              key={reuseRecord?.styleId ?? 'new'}
              selectedCount={selection.totalOccurrences}
              baselineSignature={selection.baselineSignature}
              userStyles={state.userStyles}
              reuseRecord={reuseRecord}
              error={state.mergeError}
              onConfirm={(targetProps, name, kind, listFormat, targetStyleId) =>
                actions.confirmMerge(targetProps, name, kind, listFormat, reuseRecord?.styleId ?? targetStyleId)
              }
              onCancel={actions.closeModals}
            />
          )}

          {activeEditVariant && (
            <XmlEditorModal
              key={activeEditVariant.id}
              variant={activeEditVariant}
              error={state.xmlEditorError}
              onApply={actions.applyXmlEdit}
              onCancel={actions.closeModals}
            />
          )}

          {state.contentMergeDialogOpen && state.parsedDocx && state.referenceDoc.parsedDocx && (
            <ContentMergeDialog
              sourceFilename={state.parsedDocx.originalFilename}
              referenceFilename={state.referenceDoc.parsedDocx.originalFilename}
              isMerging={state.isMergingContent}
              error={state.contentMergeError}
              onConfirm={(keepOriginalFormatting) => actions.mergeContentIntoReferenceDoc(keepOriginalFormatting)}
              onCancel={actions.closeModals}
            />
          )}
        </div>
      ) : (
        <div key="landing" className="page-transition flex min-h-0 flex-1 flex-col">
          {state.status === 'loading' ? (
            <div className="flex flex-1 items-center justify-center text-slate-500">Reading document…</div>
          ) : (
            <DropzoneUpload onFileAccepted={actions.loadFile} errorMessage={state.errorMessage} />
          )}
        </div>
      )}

      <AppFooter />
    </div>
  )
}

export default App
