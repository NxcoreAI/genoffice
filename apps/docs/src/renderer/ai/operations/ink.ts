import { nextInkId, type InkAnnotation, type InkPoint } from '../../editor/ink'
import type { OperationContext, OperationDef, OperationOutcome } from './types'
import { checkBlockIndex, failure, intOf, planned, success } from './shared'

/**
 * A3 ink op: programmatic freehand strokes on the page overlay. Reuses the
 * editor InkAnnotation model (anchor paragraph + px points relative to that
 * paragraph) and the existing save-time rasterization (buildInkImages): every
 * stroke is baked into a transparent PNG anchor on save, with the vector points
 * kept in wp:docPr descr for re-editing.
 *
 * Coordinates are normalized 0..1 to the page CONTENT box (page minus margins).
 * They are converted to the anchor paragraph's top-left; placement is exact for
 * blocks at the top of a page (pagination-accurate offsets are not exposed to
 * the host yet). pageIndex is accepted for intent/validation only.
 */

const MAX_STROKES = 50
const MAX_POINTS_PER_STROKE = 2000
const DEFAULT_INK_COLOR = 'C00000'
const DEFAULT_INK_WIDTH = 2
const TWIPS_PER_PX = 15

function requireEditable(ctx: OperationContext): string | null {
  if (!ctx.editor.isEditable) {
    return 'the document is read-only or protected; ink annotations cannot be added'
  }
  return null
}

/** docxIndex of the anchor block (or the nearest earlier/next saved block). */
function anchorDocxIndex(ctx: OperationContext, blockIndex: number): number | null {
  const doc = ctx.editor.state.doc
  for (let i = blockIndex; i >= 0; i--) {
    const index = doc.child(i).attrs?.docxIndex
    if (index !== null && index !== undefined) return Number(index)
  }
  for (let i = blockIndex + 1; i < doc.childCount; i++) {
    const index = doc.child(i).attrs?.docxIndex
    if (index !== null && index !== undefined) return Number(index)
  }
  return null
}

function inkAddOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const editError = requireEditable(ctx)
  if (editError) return failure(editError)
  const app = ctx.app
  if (!app) return failure('object.ink.add needs the app context (host call path)')
  const rawStrokes = input.strokes
  if (!Array.isArray(rawStrokes) || rawStrokes.length === 0) {
    return failure('strokes must be a non-empty array')
  }
  if (rawStrokes.length > MAX_STROKES) return failure(`at most ${MAX_STROKES} strokes per call`)
  const strokes: Array<{ color: string; width: number; points: InkPoint[] }> = []
  for (let i = 0; i < rawStrokes.length; i++) {
    const raw = rawStrokes[i]
    if (!raw || typeof raw !== 'object') return failure(`stroke #${i} must be an object`)
    const entry = raw as Record<string, unknown>
    let color = DEFAULT_INK_COLOR
    if (entry.color !== undefined) {
      color = String(entry.color).replace(/^#/, '').toUpperCase()
      if (!/^[0-9A-F]{6}$/.test(color)) return failure(`stroke #${i}: color must be 6-digit hex without #`)
    }
    let width = DEFAULT_INK_WIDTH
    if (entry.widthPx !== undefined) {
      const parsedWidth = Number(entry.widthPx)
      if (!Number.isFinite(parsedWidth) || parsedWidth < 0.5 || parsedWidth > 20) {
        return failure(`stroke #${i}: widthPx must be 0.5..20`)
      }
      width = parsedWidth
    }
    const rawPoints = entry.points
    if (!Array.isArray(rawPoints) || rawPoints.length < 2) {
      return failure(`stroke #${i}: points must be an array with at least 2 points`)
    }
    if (rawPoints.length > MAX_POINTS_PER_STROKE) {
      return failure(`stroke #${i}: at most ${MAX_POINTS_PER_STROKE} points`)
    }
    const points: InkPoint[] = []
    for (const point of rawPoints) {
      if (!point || typeof point !== 'object') return failure(`stroke #${i}: each point must be {x,y}`)
      const p = point as Record<string, unknown>
      const x = Number(p.x)
      const y = Number(p.y)
      if (!Number.isFinite(x) || !Number.isFinite(y) || x < 0 || x > 1 || y < 0 || y > 1) {
        return failure(`stroke #${i}: point coordinates must be normalized 0..1`)
      }
      points.push({ x, y })
    }
    strokes.push({ color, width, points })
  }

  const section = ctx.app?.section ?? null
  if (!section) return failure('page geometry is unavailable (empty document?)')
  const contentWidthPx = Math.max(1, (section.pageWidth - section.marginLeft - section.marginRight) / TWIPS_PER_PX)
  const contentHeightPx = Math.max(1, (section.pageHeight - section.marginTop - section.marginBottom) / TWIPS_PER_PX)

  let blockIndex = 0
  if (input.blockIndex !== undefined) {
    const parsed = checkBlockIndex(ctx.editor, input.blockIndex)
    if (typeof parsed === 'string') return failure(parsed)
    blockIndex = parsed
  }
  const anchorIndex = anchorDocxIndex(ctx, blockIndex)
  if (anchorIndex === null) {
    return failure('the anchor paragraph is not a saved block (save the document first or pass a saved blockIndex)')
  }
  let pageIndex: number | null = null
  if (input.pageIndex !== undefined) {
    const parsed = intOf(input.pageIndex)
    if (parsed === null || parsed < 0 || parsed > 10_000) {
      return failure('pageIndex must be a non-negative integer')
    }
    pageIndex = parsed
  }
  const toPx = (stroke: { points: InkPoint[] }): InkPoint[] =>
    stroke.points.map((point) => ({
      x: Math.round(point.x * contentWidthPx * 100) / 100,
      y: Math.round(point.y * contentHeightPx * 100) / 100,
    }))
  if (ctx.dryRun) {
    return planned(
      `would add ${strokes.length} ink stroke(s) anchored to block ${blockIndex}`
        + (pageIndex !== null ? ` (page ${pageIndex})` : ''),
      { blockIndex, anchorIndex, strokes: strokes.length },
    )
  }
  for (const stroke of strokes) {
    const annotation: InkAnnotation = {
      id: nextInkId(),
      anchorIndex,
      stroke: {
        tool: 'pen',
        color: stroke.color,
        width: stroke.width,
        points: toPx(stroke),
      },
    }
    app.addInk(annotation)
  }
  return success(
    `Added ${strokes.length} ink stroke(s) anchored to block ${blockIndex} (page content box `
      + `${Math.round(contentWidthPx)}×${Math.round(contentHeightPx)}px).`,
    'add ink strokes',
    { blockIndex, anchorIndex, strokes: strokes.length, ...(pageIndex !== null ? { pageIndex } : {}) },
  )
}

const MARK_POSITIONS = ['bottom', 'top', 'strike'] as const

/**
 * A4 ergonomics: one op for "draw a red line under the last paragraph".
 * The stroke is computed from the anchor block's RENDERED rect (top/bottom/
 * middle) so the model never guesses page coordinates; it reuses the same
 * InkAnnotation state + save-time rasterization as object.ink.add.
 */
function inkMarkOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const editError = requireEditable(ctx)
  if (editError) return failure(editError)
  const app = ctx.app
  if (!app) return failure('object.ink.mark needs the app context (host call path)')
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const position = String(input.position ?? '')
  if (!(MARK_POSITIONS as readonly string[]).includes(position)) {
    return failure('position must be "bottom", "top" or "strike"')
  }
  let color = DEFAULT_INK_COLOR
  if (input.color !== undefined) {
    color = String(input.color).replace(/^#/, '').toUpperCase()
    if (!/^[0-9A-F]{6}$/.test(color)) return failure('color must be 6-digit hex without #')
  }
  let width = 2
  if (input.widthPx !== undefined) {
    const parsed = Number(input.widthPx)
    if (!Number.isFinite(parsed) || parsed < 1 || parsed > 12) {
      return failure('widthPx must be 1..12')
    }
    width = parsed
  }
  const fraction = (value: unknown, fallback: number, label: string): number | { error: string } => {
    if (value === undefined) return fallback
    const parsed = Number(value)
    if (!Number.isFinite(parsed) || parsed < 0 || parsed > 1) {
      return { error: `${label} must be a number 0..1` }
    }
    return parsed
  }
  const from = fraction(input.from, 0.02, 'from')
  if (typeof from === 'object') return failure(from.error)
  const to = fraction(input.to, 0.98, 'to')
  if (typeof to === 'object') return failure(to.error)
  if (from >= to) return failure('from must be less than to (fractions of the text-column width)')

  const metrics = app.inkMeasureBlock(blockIndex)
  if (!metrics) {
    return failure(
      `block ${blockIndex} has no rendered position (not a saved/anchorable block, or the editor is not laid out); `
        + 'save the document first',
    )
  }
  const gap = Math.max(2, width / 2)
  const absoluteY =
    position === 'bottom'
      ? metrics.rect.y + metrics.rect.h + gap
      : position === 'top'
        ? Math.max(0, metrics.rect.y - gap)
        : metrics.rect.y + metrics.rect.h / 2
  const toPoint = (f: number): InkPoint => ({
    x: Math.round((metrics.rect.x - metrics.anchor.x + f * metrics.rect.w) * 100) / 100,
    y: Math.round((absoluteY - metrics.anchor.y) * 100) / 100,
  })
  const points = [toPoint(from), toPoint(to)]
  if (ctx.dryRun) {
    return planned(
      `would draw a ${width}px ${color} line on the ${position} edge of block ${blockIndex}`,
      { blockIndex, position, from, to },
    )
  }
  const annotation: InkAnnotation = {
    id: nextInkId(),
    anchorIndex: metrics.anchorIndex,
    stroke: { tool: 'pen', color, width, points },
  }
  app.addInk(annotation)
  return success(
    `Drew a ${width}px line at the ${position} edge of block ${blockIndex} `
      + `(from ${Math.round(from * 100)}% to ${Math.round(to * 100)}% of the text width).`,
    'mark ink line',
    { blockIndex, anchorIndex: metrics.anchorIndex, position, from, to },
  )
}

export const inkOperations: OperationDef[] = [
  {
    name: 'object.ink.add',
    group: 'object',
    description:
      'object.ink.add { strokes: [{ color?, widthPx?, points: [{x,y}] }], blockIndex?, pageIndex? } — draw freehand '
      + 'pen strokes on the page overlay. Points are normalized 0..1 of the page CONTENT box (converted to the anchor '
      + 'paragraph top-left); 1..50 strokes, 2..2000 points each; persisted as rasterized anchored PNGs (editable '
      + 'vector data kept in the file)',
    inputSchema: {
      type: 'object',
      properties: {
        strokes: {
          type: 'array',
          minItems: 1,
          maxItems: MAX_STROKES,
          items: {
            type: 'object',
            properties: {
              color: { type: 'string', description: '6-digit hex without # (default C00000)' },
              widthPx: { type: 'number', minimum: 0.5, maximum: 20, description: 'stroke width in px (default 2)' },
              points: {
                type: 'array',
                minItems: 2,
                maxItems: MAX_POINTS_PER_STROKE,
                items: {
                  type: 'object',
                  properties: { x: { type: 'number' }, y: { type: 'number' } },
                  required: ['x', 'y'],
                },
              },
            },
            required: ['points'],
          },
        },
        blockIndex: { type: 'integer', minimum: 0, description: 'anchor paragraph block (default 0)' },
        pageIndex: { type: 'integer', minimum: 0, description: 'intended page (validated; placement offsets are approximate)' },
      },
      required: ['strokes'],
    },
    handler: inkAddOp,
  },
  {
    name: 'object.ink.mark',
    group: 'object',
    description:
      'object.ink.mark { blockIndex, position: bottom|top|strike, color?, widthPx?, from?, to? } — draw a straight line '
      + 'on the rendered block: bottom/top edge (just outside) or strike through the middle; from/to are 0..1 fractions '
      + 'of the text-column width (default 0.02..0.98). Use this for "underline that paragraph" / "draw a red line under '
      + 'the last paragraph" — no page coordinates needed',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0 },
        position: { type: 'string', enum: MARK_POSITIONS },
        color: { type: 'string', description: '6-digit hex without # (default C00000)' },
        widthPx: { type: 'number', minimum: 1, maximum: 12, description: 'stroke width in px (default 2)' },
        from: { type: 'number', minimum: 0, maximum: 1, description: 'start fraction of the text width (default 0.02)' },
        to: { type: 'number', minimum: 0, maximum: 1, description: 'end fraction (default 0.98)' },
      },
      required: ['blockIndex', 'position'],
    },
    handler: inkMarkOp,
  },
]
