/**
 * Stable Electron host entry for products embedding GenOffice Docs in a
 * WebContentsView. Unlike index.ts this module has no standalone app startup,
 * single-instance, updater, or BrowserWindow side effects.
 */
export {
  configureDocsRuntime,
  createDocsView,
  docsQueryDirty,
  markDocsNewBlank,
  queueDocsAiContent,
  registerDocsIpc,
  requestDocsClose,
  setActiveDocsResolver,
  setDocsFileSavedHook,
  setDocsShellWindow,
  setDocsThesisPanelHook,
  teardownDocsRenderer,
} from './docs-main'
export type { DocsThesisPanelAction } from './docs-main'
export { buildThesisDocx, describeThesisBlocks } from '@genoffice/docx-engine'
