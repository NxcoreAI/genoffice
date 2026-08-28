/**
 * Stable Electron host entry for products embedding GenOffice PDF in a
 * WebContentsView. Unlike index.ts this module has no standalone app startup,
 * single-instance, updater, or BrowserWindow side effects, and it never
 * exports the GenOffice AI/standalone surface (tree-shaken in embed builds).
 */
export {
  configurePdfRuntime,
  createPdfView,
  pdfIsDirty,
  requestPdfClose,
} from './pdf-main'
