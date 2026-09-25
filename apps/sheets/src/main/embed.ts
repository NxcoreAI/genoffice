/** Stable Electron host entry for embedding GenOffice Sheets in a WebContentsView. */
export {
  configureSheetsRuntime,
  createSheetsView,
  hasActiveQueuedWorkbook,
  queueWorkbookForView,
  registerSheetsIpc,
  requestSheetsClose,
  setActiveSheetsWebContents,
  setSheetsFileSavedHook,
  setSheetsShellWindow,
  sheetsPendingEditCount,
  stopSheetsSidecar,
} from './sheets-main'
