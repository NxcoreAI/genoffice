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
  teardownDocsRenderer,
} from './docs-main'
