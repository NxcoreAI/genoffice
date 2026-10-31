import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { imageSizeOf } from '../components/ribbon-tabs'
import { removeBackground, sampleBackgroundColors, type PixelImage } from '../cutout'

/**
 * Headless picture operations shared by the ribbon's Picture Format dialogs and
 * the host agent's `object.image.*` operations (single source of truth):
 * replacing picture bytes on an image block, pixel cropping, and the local
 * tolerance-cutout background removal.
 *
 * The editor's image blocks (docProtected, blockType "image") save through two
 * paths: imported pictures (docxIndex != null) swap bytes in place via the
 * `imageReplace` save patch so the drawing XML (wrap/position/geometry)
 * survives; editor-created pictures carry a pending `genImage` payload instead.
 * Both are set here exactly like the ribbon's replace/crop/cutout handlers.
 */

const DATA_URL_RE = /^data:(image\/(?:png|jpeg|gif));base64,([A-Za-z0-9+/=]+)$/s
const BASE64_STRICT_RE = /^[A-Za-z0-9+/]+={0,2}$/

export interface ImageBlockTarget {
  node: PmNode
  /** doc position of the block node */
  pos: number
  imported: boolean
}

export function imageBlockTargetAt(
  doc: PmNode,
  blockIndex: number,
): ImageBlockTarget | { error: string } {
  if (!Number.isInteger(blockIndex) || blockIndex < 0 || blockIndex >= doc.childCount) {
    return { error: `blockIndex must be an integer 0..${doc.childCount - 1}` }
  }
  let pos = 0
  for (let i = 0; i < blockIndex; i++) pos += doc.child(i).nodeSize
  const node = doc.child(blockIndex)
  if (node.type.name !== 'docProtected' || node.attrs?.blockType !== 'image') {
    return { error: `block ${blockIndex} is a ${node.type.name}, not an image block` }
  }
  return { node, pos, imported: node.attrs?.docxIndex !== null && node.attrs?.docxIndex !== undefined }
}

export interface ImageDataPayload {
  dataUrl?: string
  base64?: string
  mime?: string
}

export type ParsedImageBytes =
  | { ok: true; dataUrl: string; mime: string; base64: string }
  | { ok: false; error: string }

/** Renderer contract for replace: either a ready data URL, or base64 + mime (gateway resolves path/url upstream). */
export function parseImagePayload(payload: ImageDataPayload): ParsedImageBytes {
  if (typeof payload.dataUrl === 'string' && payload.dataUrl) {
    const match = DATA_URL_RE.exec(payload.dataUrl)
    if (!match) return { ok: false, error: 'dataUrl must be a data:image/png|jpeg|gif;base64,… URL' }
    return { ok: true, dataUrl: payload.dataUrl, mime: match[1], base64: match[2] }
  }
  const mime = String(payload.mime ?? '').trim()
  if (!/^image\/(png|jpeg|gif)$/.test(mime)) {
    return { ok: false, error: 'mime must be image/png, image/jpeg or image/gif (the gateway resolves path/url to base64+mime)' }
  }
  const base64 = String(payload.base64 ?? '').replace(/\s+/g, '')
  if (!base64 || !BASE64_STRICT_RE.test(base64) || base64.length % 4 !== 0) {
    return { ok: false, error: 'base64 must be valid standard base64 image bytes (no data: prefix)' }
  }
  return { ok: true, dataUrl: `data:${mime};base64,${base64}`, mime, base64 }
}

/** Natural pixel size of a picture data URL, null when it cannot be decoded. */
export async function pictureNaturalSize(
  dataUrl: string,
): Promise<{ width: number; height: number } | null> {
  try {
    const size = await imageSizeOf(dataUrl)
    return size.width > 0 && size.height > 0 ? size : null
  } catch {
    return null
  }
}

export interface PictureApplyResult {
  ok: boolean
  error?: string
  widthPx?: number
  heightPx?: number
  imported?: boolean
}

/**
 * Replace the bytes of the image block at `blockIndex` (headless twin of the
 * ribbon's applyPictureBytes): display width is kept, height follows the new
 * aspect ratio, authored crop/fill windows are dropped, and imported pictures
 * record `imageReplace` while editor-created ones record `genImage`.
 */
export async function applyPictureBytesToImageBlock(
  editor: Editor,
  blockIndex: number,
  dataUrl: string,
): Promise<PictureApplyResult> {
  const target = imageBlockTargetAt(editor.state.doc, blockIndex)
  if ('error' in target) return { ok: false, error: target.error }
  return applyPictureBytesToNode(editor, target.node, target.pos, dataUrl, target.imported)
}

/** Same replace path for the ribbon, which knows the selected node position (not the block index). */
export async function applyPictureBytesAtPos(
  editor: Editor,
  pos: number,
  dataUrl: string,
): Promise<PictureApplyResult> {
  const node = editor.state.doc.nodeAt(pos)
  if (!node || node.type.name !== 'docProtected' || node.attrs?.blockType !== 'image') {
    return { ok: false, error: 'the selected node is not an image block' }
  }
  const imported = node.attrs?.docxIndex !== null && node.attrs?.docxIndex !== undefined
  return applyPictureBytesToNode(editor, node, pos, dataUrl, imported)
}

async function applyPictureBytesToNode(
  editor: Editor,
  node: PmNode,
  pos: number,
  dataUrl: string,
  imported: boolean,
): Promise<PictureApplyResult> {
  const match = DATA_URL_RE.exec(dataUrl)
  if (!match) return { ok: false, error: 'the picture bytes are not a supported png/jpeg/gif data URL' }
  let natural: { width: number; height: number }
  try {
    natural = await imageSizeOf(dataUrl)
  } catch {
    return { ok: false, error: 'the picture bytes could not be decoded (corrupt image?)' }
  }
  if (!natural.width || !natural.height) {
    return { ok: false, error: 'the picture has no readable dimensions' }
  }
  const attrs = node.attrs
  const currentW = Number(attrs.imageWidthPx) || Math.min(natural.width, 620)
  const widthPx = Math.max(1, Math.round(currentW))
  const heightPx = Math.max(1, Math.round((currentW * natural.height) / natural.width))
  const next: Record<string, unknown> = {
    ...attrs,
    imageDataUrl: dataUrl,
    imageWidthPx: widthPx,
    imageHeightPx: heightPx,
    // the new bytes are the full picture (crop/cutout bake destructively);
    // the replace pipeline strips a:srcRect on save, so a Word-authored crop or
    // fill window must not keep clipping the new image until reload
    imageCrop: null,
    imageFillRect: null,
    ...(imported
      ? { imageReplace: { base64: match[2], mime: match[1] } }
      : { genImage: { base64: match[2], mime: match[1], widthPx, heightPx } }),
  }
  editor.view.dispatch(editor.state.tr.setNodeMarkup(pos, undefined, next))
  return { ok: true, widthPx, heightPx, imported }
}

// ---- pixel crop ----

export interface CropRectPx {
  x: number
  y: number
  w: number
  h: number
}

export interface NormalizedCrop {
  x: number
  y: number
  w: number
  h: number
}

/** Validate/clamp a pixel crop rect against the picture's natural size. */
export function normalizeCropRect(
  rect: unknown,
  natural: { width: number; height: number },
): { ok: true; rect: NormalizedCrop } | { ok: false; error: string } {
  if (!rect || typeof rect !== 'object') {
    return { ok: false, error: 'rect must be { x, y, w, h } in pixels against the image natural size' }
  }
  const raw = rect as Record<string, unknown>
  const values = ['x', 'y', 'w', 'h'].map((key) => raw[key])
  if (values.some((value) => typeof value !== 'number' || !Number.isFinite(value))) {
    return { ok: false, error: 'rect.x/y/w/h must be finite numbers (pixels)' }
  }
  const x = Math.round(raw.x as number)
  const y = Math.round(raw.y as number)
  const w = Math.round(raw.w as number)
  const h = Math.round(raw.h as number)
  if (w <= 0 || h <= 0) return { ok: false, error: 'rect.w/rect.h must be positive' }
  if (x < 0 || y < 0 || x + w > natural.width || y + h > natural.height) {
    return {
      ok: false,
      error:
        `rect is outside the image bounds (image is ${natural.width}×${natural.height}px; `
        + `x+w must be ≤ ${natural.width}, y+h ≤ ${natural.height})`,
    }
  }
  return { ok: true, rect: { x, y, w, h } }
}

function loadImage(dataUrl: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.onerror = () => reject(new Error('image decode failed'))
    img.src = dataUrl
  })
}

export interface CropResult {
  ok: boolean
  error?: string
  dataUrl?: string
  widthPx?: number
  heightPx?: number
  rect?: NormalizedCrop
}

/** Crop the picture pixels with a canvas (the CropDialog apply path, dialog-free). */
export async function cropImageDataUrl(dataUrl: string, rect: unknown): Promise<CropResult> {
  let img: HTMLImageElement
  try {
    img = await loadImage(dataUrl)
  } catch {
    return { ok: false, error: 'the current picture could not be decoded' }
  }
  const natural = { width: img.naturalWidth, height: img.naturalHeight }
  if (!natural.width || !natural.height) return { ok: false, error: 'the picture has no readable dimensions' }
  const normalized = normalizeCropRect(rect, natural)
  if (!normalized.ok) return { ok: false, error: normalized.error }
  const { x, y, w, h } = normalized.rect
  try {
    const canvas = document.createElement('canvas')
    canvas.width = w
    canvas.height = h
    const ctx = canvas.getContext('2d')
    if (!ctx) return { ok: false, error: 'canvas rendering is unavailable in this environment' }
    ctx.drawImage(img, x, y, w, h, 0, 0, w, h)
    // jpeg has no alpha channel; keep jpeg to avoid size bloat. Others (png/gif)
    // all output png to preserve transparency
    const isJpeg = /^data:image\/jpeg/.test(dataUrl)
    const cropped = isJpeg ? canvas.toDataURL('image/jpeg', 0.92) : canvas.toDataURL('image/png')
    return { ok: true, dataUrl: cropped, widthPx: w, heightPx: h, rect: normalized.rect }
  } catch {
    return { ok: false, error: 'cropping failed (canvas error)' }
  }
}

// ---- local tolerance-cutout background removal ----

export interface RemoveBackgroundResult {
  ok: boolean
  error?: string
  dataUrl?: string
  removedPct?: number
}

/** Dialog-free local cutout: full-resolution border sampling + flood fill → PNG with alpha. */
export async function removeImageBackgroundDataUrl(
  dataUrl: string,
  tolerance = 25,
): Promise<RemoveBackgroundResult> {
  let img: HTMLImageElement
  try {
    img = await loadImage(dataUrl)
  } catch {
    return { ok: false, error: 'the current picture could not be decoded' }
  }
  const width = img.naturalWidth
  const height = img.naturalHeight
  if (!width || !height) return { ok: false, error: 'the picture has no readable dimensions' }
  try {
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const ctx = canvas.getContext('2d')
    if (!ctx) return { ok: false, error: 'canvas rendering is unavailable in this environment' }
    ctx.drawImage(img, 0, 0)
    const image = ctx.getImageData(0, 0, width, height)
    const full: PixelImage = { data: image.data, width, height }
    const bgColors = sampleBackgroundColors(full)
    const result = removeBackground(full, tolerance, bgColors)
    const out = document.createElement('canvas')
    out.width = width
    out.height = height
    const outCtx = out.getContext('2d')
    if (!outCtx) return { ok: false, error: 'canvas rendering is unavailable in this environment' }
    outCtx.putImageData(new ImageData(result.data, width, height), 0, 0)
    return {
      ok: true,
      dataUrl: out.toDataURL('image/png'),
      removedPct: Math.round((result.removedCount / (width * height)) * 100),
    }
  } catch {
    return { ok: false, error: 'background removal failed (canvas error)' }
  }
}