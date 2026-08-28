/**
 * Stable Electron host entry for products embedding GenOffice Slides in a
 * WebContentsView. Unlike index.ts this module has no standalone app startup,
 * single-instance, updater, or BrowserWindow side effects, and it never
 * exports the GenOffice AI/standalone surface (tree-shaken in embed builds).
 */
export {
  configureSlidesRuntime,
  createSlidesView,
  registerSlidesIpc,
  requestSlidesClose,
  setActiveSlidesWebContents,
  setSlidesShellWindow,
  slidesIsDirty,
} from './slides-main'
