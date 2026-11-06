/**
 * Stable Electron host entry for products embedding GenOffice Markdown in a
 * WebContentsView (EverRoom). Unlike index.ts this module has no standalone
 * app startup or BrowserWindow side effects — it re-exports the host surface
 * of markdown-main only.
 */
export {
  configureMarkdownRuntime,
  createMarkdownView,
  markdownFileRenamed,
  markdownFilePath,
  markdownIsDirty,
  pushMarkdownAskQueue,
  registerMarkdownIpc,
  requestMarkdownClose,
  requestMarkdownSave,
  sendMarkdownExportRequest,
  sendMarkdownPrintRequest,
  setMarkdownAgentAskHook,
  setMarkdownAskQueueActionHook,
  setMarkdownChromeHidden,
  setMarkdownDocxExportedHook,
  setMarkdownFileSavedHook,
  teardownMarkdownRenderer,
  waitForMarkdownAgentReady,
} from './markdown-main'
export type {
  ExportFormat,
  MdAgentAskOp,
  MdAgentAskResult,
  MdAskQueueAction,
  MdAskQueueItem,
  MdComment,
  MdCommentsDoc,
  SaveMode,
} from '../shared/ipc'
