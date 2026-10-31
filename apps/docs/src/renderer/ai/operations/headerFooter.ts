import type { HfImage } from '@genoffice/docx-engine'
import { parseImagePayload, pictureNaturalSize } from '../../editor/picture-ops'
import type { OperationContext, OperationDef, OperationOutcome } from './types'
import { failure, intOf, planned, success } from './shared'

/**
 * A3 header/footer image ops. Header/footer parts own their rels and media, so
 * the engine save path was extended to embed NEW images (media part + part rels
 * + drawing paragraph) and to drop original image paragraphs; these ops only
 * append/remove the HfImage display models in App state (which the silent save
 * then persists). Image bytes come from the gateway as base64+mime (path/url
 * resolved upstream) or a ready dataUrl.
 */

const HF_KINDS = ['header', 'footer'] as const
const IMAGE_ALIGNS = ['left', 'center', 'right'] as const

function requireEditable(ctx: OperationContext): string | null {
  if (!ctx.editor.isEditable) {
    return 'the document is read-only or protected; header/footer images cannot be modified'
  }
  return null
}

function hfImageList(ctx: OperationContext, kind: 'header' | 'footer'): HfImage[] {
  const state = ctx.app?.getHeaderFooterImages()
  if (!state) return []
  return kind === 'header' ? state.header : state.footer
}

interface HfImageDraft {
  image: HfImage
  widthPx: number
  heightPx: number
  floating: boolean
  behind: boolean
}

/** Validate the image source + layout fields and decode the natural size (no App mutation). */
async function buildHfImageDraft(
  input: Record<string, unknown>,
): Promise<{ ok: true; draft: HfImageDraft } | { ok: false; error: string }> {
  const parsed = parseImagePayload({
    dataUrl: typeof input.dataUrl === 'string' ? input.dataUrl : undefined,
    base64: typeof input.base64 === 'string' ? input.base64 : undefined,
    mime: typeof input.mime === 'string' ? input.mime : undefined,
  })
  if (!parsed.ok) return { ok: false, error: parsed.error }
  let widthPx: number | null = null
  if (input.widthPx !== undefined) {
    const parsedWidth = intOf(input.widthPx)
    if (parsedWidth === null || parsedWidth < 8 || parsedWidth > 2000) {
      return { ok: false, error: 'widthPx must be an integer 8..2000' }
    }
    widthPx = parsedWidth
  }
  const align = input.align === undefined ? 'left' : String(input.align)
  if (!(IMAGE_ALIGNS as readonly string[]).includes(align)) {
    return { ok: false, error: 'align must be left/center/right' }
  }
  if (input.floating !== undefined && typeof input.floating !== 'boolean') {
    return { ok: false, error: 'floating must be a boolean' }
  }
  if (input.behind !== undefined && typeof input.behind !== 'boolean') {
    return { ok: false, error: 'behind must be a boolean (picture watermark)' }
  }
  const floating = input.floating === true || input.behind === true
  const behind = input.behind === true
  const natural = await pictureNaturalSize(parsed.dataUrl)
  if (!natural) return { ok: false, error: 'the picture bytes could not be decoded (corrupt image?)' }
  const displayWidth = Math.max(8, Math.round(widthPx ?? Math.min(natural.width, 620)))
  const displayHeight = Math.max(1, Math.round((displayWidth * natural.height) / natural.width))
  const image: HfImage = {
    dataUrl: parsed.dataUrl,
    widthPx: displayWidth,
    heightPx: displayHeight,
    ...(align ? { align: align as HfImage['align'] } : {}),
    ...(floating
      ? {
          floating: true,
          ...(behind ? { behind: true } : {}),
          posH: align as NonNullable<HfImage['posH']>,
          posV: 'top' as NonNullable<HfImage['posV']>,
        }
      : {}),
  }
  return { ok: true, draft: { image, widthPx: displayWidth, heightPx: displayHeight, floating, behind } }
}

function hfKindOf(input: Record<string, unknown>): 'header' | 'footer' | { error: string } {
  const kind = String(input.kind ?? '')
  if (!(HF_KINDS as readonly string[]).includes(kind)) {
    return { error: 'kind must be "header" or "footer"' }
  }
  return kind as 'header' | 'footer'
}

function hfInsertImageOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): Promise<OperationOutcome> | OperationOutcome {
  const editError = requireEditable(ctx)
  if (editError) return failure(editError)
  const kind = hfKindOf(input)
  if (typeof kind === 'object') return failure(kind.error)
  const app = ctx.app
  if (!app) return failure('headerFooter.insertImage needs the app context (host call path)')
  return (async () => {
    const built = await buildHfImageDraft(input)
    if (!built.ok) return failure(built.error)
    const { draft } = built
    const where = `the ${kind}`
    if (ctx.dryRun) {
      return planned(
        `would insert a ${draft.widthPx}×${draft.heightPx}px ${draft.floating ? 'floating ' : ''}image into ${where}`
          + (draft.behind ? ' behind the body text' : ''),
        { kind, widthPx: draft.widthPx, heightPx: draft.heightPx },
      )
    }
    const { index } = app.addHeaderFooterImage(kind, draft.image)
    return success(
      `Inserted a ${draft.widthPx}×${draft.heightPx}px ${draft.floating ? 'floating ' : ''}image into ${where}`
        + `${draft.behind ? ' (behind the body text)' : ''}; ${kind} index ${index}.`,
      `insert ${kind} image`,
      { kind, index, widthPx: draft.widthPx, heightPx: draft.heightPx, floating: draft.floating, behind: draft.behind },
    )
  })()
}

function hfReplaceImageOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): Promise<OperationOutcome> | OperationOutcome {
  const editError = requireEditable(ctx)
  if (editError) return failure(editError)
  const kind = hfKindOf(input)
  if (typeof kind === 'object') return failure(kind.error)
  const app = ctx.app
  if (!app) return failure('headerFooter.replaceImage needs the app context (host call path)')
  return (async () => {
    const list = hfImageList(ctx, kind)
    // resolve the slot first (default: exactly one image; none = insert)
    let targetIndex = -1
    if (input.index === undefined) {
      if (list.length > 1) {
        return failure(`the ${kind} has ${list.length} images; pass index (see documentState.headerFooter.images)`)
      }
      targetIndex = list.length === 1 ? 0 : -1
    } else {
      const parsedIndex = intOf(input.index)
      if (parsedIndex === null || parsedIndex < 0) {
        return failure('index must be a non-negative integer (see documentState.headerFooter.images)')
      }
      if (parsedIndex >= list.length) {
        return failure(`${kind} has ${list.length} image(s); index ${parsedIndex} is out of range`)
      }
      targetIndex = parsedIndex
    }
    // validate the new source BEFORE any mutation, so a bad image never half-applies
    const built = await buildHfImageDraft(input)
    if (!built.ok) return failure(built.error)
    const { draft } = built
    if (ctx.dryRun) {
      return planned(
        targetIndex >= 0
          ? `would replace the ${kind} image at index ${targetIndex} with a ${draft.widthPx}×${draft.heightPx}px image`
          : `the ${kind} has no image; would insert a ${draft.widthPx}×${draft.heightPx}px image`,
        { kind, replaced: targetIndex >= 0 ? targetIndex : null, widthPx: draft.widthPx, heightPx: draft.heightPx },
      )
    }
    if (targetIndex >= 0 && !app.removeHeaderFooterImage(kind, targetIndex)) {
      return failure(`could not remove the ${kind} image at index ${targetIndex}`)
    }
    const { index } = app.addHeaderFooterImage(kind, draft.image)
    return success(
      targetIndex >= 0
        ? `Replaced the ${kind} image at index ${targetIndex} with a ${draft.widthPx}×${draft.heightPx}px image`
        : `The ${kind} had no image; inserted a ${draft.widthPx}×${draft.heightPx}px image`,
      `replace ${kind} image`,
      {
        kind,
        replaced: targetIndex >= 0 ? targetIndex : null,
        index,
        widthPx: draft.widthPx,
        heightPx: draft.heightPx,
        floating: draft.floating,
        behind: draft.behind,
      },
    )
  })()
}

function hfRemoveImageOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const editError = requireEditable(ctx)
  if (editError) return failure(editError)
  const kind = String(input.kind ?? '')
  if (!(HF_KINDS as readonly string[]).includes(kind)) {
    return failure('kind must be "header" or "footer"')
  }
  const app = ctx.app
  if (!app) return failure('headerFooter.removeImage needs the app context (host call path)')
  const index = intOf(input.index)
  if (index === null || index < 0) return failure('index must be a non-negative integer (see documentState.headerFooter.images)')
  const list = hfImageList(ctx, kind as 'header' | 'footer')
  if (index >= list.length) {
    return failure(`${kind} has ${list.length} image(s); index ${index} is out of range`)
  }
  const target = list[index]
  if (ctx.dryRun) {
    return planned(`would remove the ${kind} image at index ${index}`, { kind, index })
  }
  if (!app.removeHeaderFooterImage(kind as 'header' | 'footer', index)) {
    return failure(`could not remove the ${kind} image at index ${index}`)
  }
  return success(
    `Removed the ${kind} image at index ${index}`
      + ` (${target.widthPx ?? '?'}×${target.heightPx ?? '?'}px); the next save rebuilds the part.`,
    `remove ${kind} image`,
    { kind, index },
  )
}

export const headerFooterOperations: OperationDef[] = [
  {
    name: 'headerFooter.insertImage',
    group: 'headerFooter',
    description:
      'headerFooter.insertImage { kind: header|footer, base64?, mime?, dataUrl?, align?, widthPx?, floating?, behind? } — '
      + 'embed a picture in the default page header/footer. The gateway resolves path/url to base64+mime and forwards '
      + '{base64, mime}; a dataUrl is also accepted. widthPx keeps the aspect ratio (default: natural width capped at '
      + '620); floating/behind create a positioned (optionally behind-text) picture',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: HF_KINDS },
        base64: { type: 'string', description: 'standard base64 bytes (no data: prefix); gateway-resolved' },
        mime: { type: 'string', enum: ['image/png', 'image/jpeg', 'image/gif'] },
        dataUrl: { type: 'string', description: 'alternative to base64+mime: a data:image/...;base64,… URL' },
        align: { type: 'string', enum: IMAGE_ALIGNS },
        widthPx: { type: 'integer', minimum: 8, maximum: 2000, description: 'display width in px (aspect kept)' },
        floating: { type: 'boolean', description: 'positioned (wp:anchor) instead of inline' },
        behind: { type: 'boolean', description: 'behind the body text (picture watermark)' },
      },
      required: ['kind'],
    },
    handler: hfInsertImageOp,
  },
  {
    name: 'headerFooter.removeImage',
    group: 'headerFooter',
    description:
      'headerFooter.removeImage { kind, index } — remove the header/footer image at the index from '
      + 'documentState.headerFooter.images; removal is paragraph-level (a paragraph holding several images goes whole)',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: HF_KINDS },
        index: { type: 'integer', minimum: 0 },
      },
      required: ['kind', 'index'],
    },
    handler: hfRemoveImageOp,
  },
  {
    name: 'headerFooter.replaceImage',
    group: 'headerFooter',
    description:
      'headerFooter.replaceImage { kind: header|footer, index?, base64?, mime?, dataUrl?, align?, widthPx?, floating?, '
      + 'behind? } — swap the header/footer picture in one step (remove the indexed image + insert the new bytes). '
      + 'index defaults to the only image; when the part has NO image it inserts instead. The gateway resolves '
      + 'fileId/path/url to base64+mime (dataUrl also accepted); source and index are validated before any mutation',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: HF_KINDS },
        index: { type: 'integer', minimum: 0, description: 'image index from documentState.headerFooter.images' },
        base64: { type: 'string', description: 'standard base64 bytes (no data: prefix); gateway-resolved' },
        mime: { type: 'string', enum: ['image/png', 'image/jpeg', 'image/gif'] },
        dataUrl: { type: 'string', description: 'alternative to base64+mime: a data:image/...;base64,… URL' },
        align: { type: 'string', enum: IMAGE_ALIGNS },
        widthPx: { type: 'integer', minimum: 8, maximum: 2000, description: 'display width in px (aspect kept)' },
        floating: { type: 'boolean' },
        behind: { type: 'boolean', description: 'behind the body text (picture watermark)' },
      },
      required: ['kind'],
    },
    handler: hfReplaceImageOp,
  },
]
