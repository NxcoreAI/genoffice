/** Stable Electron host entry for embedding GenOffice Sheets in a WebContentsView. */
export {
  applyAgentWorkbookTools,
  configureSheetsRuntime,
  createSheetsView,
  describeAgentWorkbook,
  hasActiveQueuedWorkbook,
  queueWorkbookForView,
  pushSheetsAskQueue,
  registerSheetsIpc,
  requestSheetsClose,
  setActiveSheetsWebContents,
  setSheetsAgentAskHook,
  setSheetsAskQueueActionHook,
  setSheetsFileSavedHook,
  setSheetsShellWindow,
  sheetsPendingEditCount,
  stopSheetsSidecar,
  teardownSheetsRenderer,
  waitForAgentReady,
} from './sheets-main'
export type { SheetsAskQueueAction, SheetsAskQueueItem } from '../shared/desktop-api'
export type {
  SheetsAgentDescribeResult,
  SheetsAgentToolCall,
  SheetsAgentToolResult,
} from './sheets-main'
