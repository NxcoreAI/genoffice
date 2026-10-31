/** Stable Electron host entry for embedding GenOffice Sheets in a WebContentsView. */
export {
  applyAgentWorkbookTools,
  configureSheetsRuntime,
  createSheetsView,
  describeAgentWorkbook,
  hasActiveQueuedWorkbook,
  queueWorkbookForView,
  registerSheetsIpc,
  requestSheetsClose,
  setActiveSheetsWebContents,
  setSheetsFileSavedHook,
  setSheetsShellWindow,
  sheetsPendingEditCount,
  stopSheetsSidecar,
  teardownSheetsRenderer,
  waitForAgentReady,
} from './sheets-main'
export type {
  SheetsAgentDescribeResult,
  SheetsAgentToolCall,
  SheetsAgentToolResult,
} from './sheets-main'
