import { applyImageZOrder } from '@genoffice/docx-engine'

/**
 * Stacking-order helper for imported (Word-authored) DrawingML shapes and text
 * boxes. Editor-created shapes carry `imageZOrder` and are re-encoded when the
 * save path rebuilds their anchor; imported shapes keep their ORIGINAL anchor
 * bytes, so a reorder must patch only `wp:anchor@relativeHeight` and leave
 * every other anchor attribute/element untouched (the same targeted-patch idea
 * as image z-order, see applyImageZOrder in the docx engine).
 *
 * Word writes relativeHeight as 251658240 + rank; the agent-facing value is the
 * raw integer so it round-trips exactly what the file contains.
 */

export const RELATIVE_HEIGHT_BASE = 251658240

export function hasDrawingAnchor(xml: string): boolean {
  return /<wp:anchor[\s>]/.test(xml)
}

/** Current raw wp:anchor@relativeHeight of the fragment, or null when it has no anchor. */
export function readShapeRelativeHeight(xml: string): number | null {
  const attr = /<wp:anchor[^>]*?\brelativeHeight="(\d+)"/.exec(xml)?.[1]
  const value = attr === undefined ? NaN : Number(attr)
  return Number.isFinite(value) ? value : null
}

/** Patch ONLY relativeHeight (raw Word value); no-op for fragments without an anchor. */
export function patchShapeRelativeHeight(xml: string, relativeHeight: number): string {
  if (!hasDrawingAnchor(xml)) return xml
  const raw = Math.max(0, Math.round(relativeHeight))
  return applyImageZOrder(xml, raw - RELATIVE_HEIGHT_BASE)
}

/** Raw relativeHeight for an agent z-order value (z rank → Word raw). */
export function relativeHeightForRank(rank: number): number {
  return RELATIVE_HEIGHT_BASE + Math.round(rank)
}