/**
 * Host-driven image insert helpers for the Word agent (`insert_image_base64`).
 *
 * The host gateway owns "where the image comes from": it resolves `path`/`url`
 * into bytes and forwards raw base64 to the renderer together with an
 * `afterIndex` (block anchor) and a `maxWidthPx` cap. These helpers keep the
 * two decisions that used to be dropped — insertion position and rendered
 * width — pure, so they are testable without an editor or image decoding.
 */
import type { Node as PMNode } from '@tiptap/pm/model'

/** Width range of the tool schema (px); mirrors insert_image_base64's inputSchema. */
export const IMAGE_WIDTH_MIN_PX = 16
export const IMAGE_WIDTH_MAX_PX = 1200

/**
 * ProseMirror doc position to insert an image block at for a top-level block
 * index: right after that block. `-1`, a non-integer, or any out-of-range
 * value means "end of document" (the tool contract for `afterIndex`), so the
 * image never lands inside the anchored block or at the cursor.
 */
export function agentImageInsertPos(doc: PMNode, afterIndex: number): number {
  let pos = doc.content.size
  if (Number.isInteger(afterIndex) && afterIndex >= 0 && afterIndex < doc.childCount) {
    doc.forEach((node, offset, index) => {
      if (index === afterIndex) pos = offset + node.nodeSize
    })
  }
  return pos
}

/**
 * The width an agent image may use in whole px: bounded by the tool schema
 * (16..1200) and by the renderer's own hard cap (`cap`, e.g. the text column
 * width). Non-finite input falls back to the cap.
 */
export function clampImageWidth(maxWidthPx: number, cap: number): number {
  const capPx =
    Number.isFinite(cap) && cap > 0 ? Math.min(cap, IMAGE_WIDTH_MAX_PX) : IMAGE_WIDTH_MAX_PX
  const requested = Number.isFinite(maxWidthPx) ? maxWidthPx : capPx
  const schemaBounded = Math.min(IMAGE_WIDTH_MAX_PX, Math.max(IMAGE_WIDTH_MIN_PX, requested))
  return Math.floor(Math.min(capPx, schemaBounded))
}

/**
 * Scale natural image dimensions down to fit `capPx`; never upscale
 * (scale = min(1, capPx / natural.width)). Sizes come back in whole px, the
 * same values written into the docProtected image attrs.
 */
export function scaledImageSize(
  natural: { width: number; height: number },
  capPx: number,
): { width: number; height: number } {
  const cap = Number.isFinite(capPx) && capPx > 0 ? capPx : IMAGE_WIDTH_MAX_PX
  const scale = natural.width > 0 ? Math.min(1, cap / natural.width) : 1
  return {
    width: Math.round(natural.width * scale),
    height: Math.round(natural.height * scale),
  }
}
