import { AppFooter } from './components/AppFooter'
import { AppHeader } from './components/AppHeader'
import { ContentMergeDialog } from './components/ContentMergeDialog'
import { DropzoneUpload } from './components/DropzoneUpload'
import { MergeDialog } from './components/MergeDialog'
import { ReferenceDocPanel } from './components/ReferenceDocPanel'
import { SaveButton } from './components/SaveButton'
import { StyleReportPanel } from './components/StyleReportPanel'
import { UserStylesPanel } from './components/UserStylesPanel'
import { XmlEditorModal } from './components/XmlEditorModal'
import { useDocxWorkspace } from './hooks/useDocxWorkspace'

/** Top-level app shell. Owns the single useDocxWorkspace instance and
 * switches between the upload screen and the two-panel workspace based on
 * workspace.state.status. All actual OOXML logic lives in src/lib/ooxml/ -
 * this component is purely about wiring state to the presentational pieces. */
function App() {
  const { state, selection, activeEditVariant, actions } = useDocxWorkspace()

  const reuseRecord = state.mergeDialogReuseStyleId
    ? (state.userStyles.find((r) => r.styleId === state.mergeDialogReuseStyleId) ?? null)
    : null
  const importedStyleCount = state.userStyles.filter((r) => r.fromReferenceDoc).length

  if (state.status === 'empty' || state.status === 'error' || state.status === 'loading') {
    return (
      <div className="flex h-full flex-col bg-slate-50">
        <AppHeader filename={null} onLoadDifferentFile={() => {}} />
        {state.status === 'loading' ? (
          <div className="flex flex-1 items-center justify-center text-slate-500">
            Reading document…
          </div>
        ) : (
          <DropzoneUpload onFileAccepted={actions.loadFile} errorMessage={state.errorMessage} />
        )}
        <AppFooter />
      </div>
    )
  }

  return (
    <div className="flex h-full flex-col bg-slate-50">
      <AppHeader
        filename={state.parsedDocx?.originalFilename ?? null}
        onLoadDifferentFile={() => window.location.reload()}
      >
        <SaveButton disabled={!state.parsedDocx} isSaving={state.isSaving} onSave={actions.save} />
      </AppHeader>

      <ReferenceDocPanel
        referenceDoc={state.referenceDoc}
        importedStyleCount={importedStyleCount}
        isMergingContent={state.isMergingContent}
        onAttach={actions.loadReferenceDoc}
        onRemove={actions.removeReferenceDoc}
        onOpenContentMerge={actions.openContentMergeDialog}
      />

      <main className="grid min-h-0 flex-1 grid-cols-2 gap-4 p-4">
        <StyleReportPanel
          styleReport={state.styleReport}
          selectedIds={state.selectedVariantIds}
          onToggleSelect={actions.toggleSelectVariant}
          onEditXml={actions.openXmlEditor}
          onMergeSelected={() => actions.openMergeDialog()}
          hasReferenceStyles={importedStyleCount > 0}
          bulkMergeError={state.bulkMergeError}
          onSelectMatchingReferenceStyles={actions.selectVariantsMatchingReferenceStyles}
          onBulkMergeMatched={actions.bulkMergeMatchedToReference}
        />
        <UserStylesPanel
          userStyles={state.userStyles}
          styleReport={state.styleReport}
          onEditStyle={actions.openMergeDialog}
          onCreateNewStyle={() => actions.openMergeDialog()}
        />
      </main>

      {state.mergeDialogOpen && (
        <MergeDialog
          key={reuseRecord?.styleId ?? 'new'}
          selectedCount={selection.totalOccurrences}
          baselineSignature={selection.baselineSignature}
          userStyles={state.userStyles}
          reuseRecord={reuseRecord}
          error={state.mergeError}
          onConfirm={(targetProps, name, targetStyleId) =>
            actions.confirmMerge(targetProps, name, reuseRecord?.styleId ?? targetStyleId)
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

      <AppFooter />
    </div>
  )
}

export default App
