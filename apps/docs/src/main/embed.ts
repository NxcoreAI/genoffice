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
  exportDocsPdf,
  markDocsNewBlank,
  pushDocsAskQueue,
  queueDocsAiContent,
  registerDocsIpc,
  requestDocsClose,
  setActiveDocsResolver,
  setDocsAgentAskHook,
  setDocsAskQueueActionHook,
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
export type { DocsAskQueueAction, DocsAskQueueItem } from '../shared/ipc'
export {
  buildThesisDocx,
  buildThesisDocxOnTemplate,
  describeThesisBlocks,
} from '@genoffice/docx-engine'
