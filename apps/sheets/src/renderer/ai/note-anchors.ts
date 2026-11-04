/**
 * Viewport geometry for host-note anchors on the sheet grid.
 *
 * The host pushes A1 ranges; the editor has to place a dashed box + numbered
 * badge over each of them and keep them attached while the user scrolls, zooms
 * or switches sheets. Univer's facade exposes scene-space cell rects
 * (`FRange.getCellRect`), so the mapping to viewport pixels mirrors
 * shape-draw.ts: subtract the main viewport's scroll in scene space, scale by
 * the sheet zoom, then offset by the render surface's DOM position.
 *
 * Anything transient (skeleton not ready, non-active sheet, range off-screen)
 * degrades to null — the caller simply skips that anchor.
 */
import { IRenderManagerService, SHEET_VIEWPORT_KEY } from '@univerjs/engine-render'

import { parseRange } from '../../domain/cell-address'
import type { UniverRuntime } from '../univer-state'

export interface NoteAnchorRect {
  readonly left: number
  readonly top: number
  readonly right: number
  readonly bottom: number
}

const GRID_HOST_ID = 'univer-container'
/** Pixels of overlap required before an anchor is worth drawing. */
const MIN_VISIBLE_PX = 4

/** The largest canvas inside the grid host is the sheet's render surface. */
function renderSurface(): DOMRect | null {
  const host = document.getElementById(GRID_HOST_ID)
  if (!host) return null
  let surface: DOMRect | null = null
  for (const canvas of host.querySelectorAll('canvas')) {
    const rect = canvas.getBoundingClientRect()
    if (!surface || rect.width * rect.height > surface.width * surface.height) surface = rect
  }
  return surface
}

/**
 * Viewport rect (client coordinates) of an A1 range on the active worksheet,
 * clipped to the visible grid. Null when the range is on another sheet, the
 * sheet has no layout yet, or the range is scrolled completely out of view.
 */
export function noteRangeRect(
  runtime: UniverRuntime | null,
  sheetId: string | undefined,
  rangeA1: string | undefined,
): NoteAnchorRect | null {
  if (!runtime || !sheetId || !rangeA1) return null
  try {
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const worksheet = workbook?.getActiveSheet()
    if (!workbook || !worksheet || worksheet.getSheetId() !== sheetId) return null
    const bounds = parseRange(rangeA1.replace(/\$/g, ''))
    const surface = renderSurface()
    if (!surface) return null
    const render = runtime.univer
      .__getInjector()
      .get(IRenderManagerService)
      .getRenderById(workbook.getId())
    const viewport = render?.scene?.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)
    const scrollX = viewport?.viewportScrollX ?? 0
    const scrollY = viewport?.viewportScrollY ?? 0
    const zoom = worksheet.getZoom() || 1
    const start = worksheet.getRange(bounds.startRow, bounds.startColumn, 1, 1).getCellRect()
    const end = worksheet.getRange(bounds.endRow, bounds.endColumn, 1, 1).getCellRect()
    const left = surface.left + (Math.min(start.left, end.left) - scrollX) * zoom
    const top = surface.top + (Math.min(start.top, end.top) - scrollY) * zoom
    const right = surface.left + (Math.max(start.right, end.right) - scrollX) * zoom
    const bottom = surface.top + (Math.max(start.bottom, end.bottom) - scrollY) * zoom
    const clipped = {
      left: Math.max(left, surface.left),
      top: Math.max(top, surface.top),
      right: Math.min(right, surface.right),
      bottom: Math.min(bottom, surface.bottom),
    }
    if (
      clipped.right - clipped.left < MIN_VISIBLE_PX ||
      clipped.bottom - clipped.top < MIN_VISIBLE_PX
    ) {
      return null
    }
    return clipped
  } catch {
    // Swapped skeletons / lazy-loading mid-frame: retry on the next scroll tick.
    return null
  }
}
