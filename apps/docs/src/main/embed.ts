/**
 * Stable Electron host entry for products embedding GenOffice Docs in a
 * WebContentsView. Unlike index.ts this module has no standalone app startup,
 * single-instance, updater, or BrowserWindow side effects.
 */
export {
  applyAgentDocTools,
  configureDocsRuntime,
  createDocsView,
  describeAgentDoc,
  docsQueryDirty,
  markDocsNewBlank,
  queueDocsAiContent,
  registerDocsIpc,
  requestDocsClose,
  setActiveDocsResolver,
  setDocsChromeHidden,
  setDocsFileSavedHook,
  setDocsMenuGate,
  setDocsShellWindow,
  setDocsThesisPanelHook,
  teardownDocsRenderer,
  waitForAgentReady,
} from './docs-main'
export type {
  DocsAgentDescribeResult,
  DocsAgentToolCall,
  DocsAgentToolResult,
  DocsThesisPanelAction,
} from './docs-main'
export {
  buildThesisDocx,
  buildThesisDocxOnTemplate,
  describeThesisBlocks,
} from '@genoffice/docx-engine'
