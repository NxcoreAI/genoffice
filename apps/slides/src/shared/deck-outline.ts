/**
 * Deck outline helpers shared by the renderer AI skill and the embed host's
 * agent-edit surface: turn RenderSlide trees into compact text the model reads
 * (per-page outline + one-page geometry dump). Pure data transforms — no
 * renderer or main-process dependencies.
 */
import type {
  GroupRenderNode,
  PictureRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
} from '@genoffice/pptx-render'

/** Element info shared by outline/read_slide/edit scripts (includes absolute geometry; locked = layout decoration, read-only). */
export interface DeckNodeInfo {
  id: string
  type: string
  /** Element text (tables joined row by row; empty string when no text) */
  text: string
  x: number
  y: number
  w: number
  h: number
  rotation: number
  /** Max font size of the text (pt; omitted when no text) */
  fontSizePt?: number
  /** Solid fill color (#RRGGBB; omitted for none/gradient/image fills) */
  fill?: string
  /** Dominant text color (#RRGGBB; omitted when no text) */
  textColor?: string
  /** Stroke/border color (#RRGGBB; omitted when no stroke) */
  strokeColor?: string
  /** Group children: coordinates are absolute; direct children of a top-level group are editable (ops carry groupId), deeper nesting is read-only */
  inGroup?: boolean
  /** Direct child of a top-level group: the parent group's id (write ops route through the in-group edit pipeline) */
  groupId?: string
  /** master/layout decoration: read-only, not modifiable */
  locked?: boolean
}

/** Find one node by id in the node tree (including groups). */
export function findNodeById(nodes: RenderNode[], id: string): RenderNode | undefined {
  for (const n of nodes) {
    if (n.sourceId === id || n.durableId === id) return n
    if (n.type === 'group') {
      const hit = findNodeById(n.children, id)
      if (hit) return hit
    }
  }
  return undefined
}

function nodeText(n: RenderNode): string {
  if (n.type === 'shape' || n.type === 'text') {
    return ((n as ShapeRenderNode).text?.lines ?? [])
      .map((line) => line.runs.map((r) => r.text).join(''))
      .join('\n')
  }
  if (n.type === 'table') {
    // Tables join cell text row by row (tab-separated) so the AI can read table content
    const byRow = new Map<number, string[]>()
    for (const c of n.cells) {
      const t = (c.text?.lines ?? []).map((l) => l.runs.map((r) => r.text).join('')).join(' ')
      const row = byRow.get(c.y) ?? []
      row.push(t)
      byRow.set(c.y, row)
    }
    return [...byRow.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([, r]) => r.join('\t'))
      .join('\n')
  }
  return ''
}

/** Max font size of the text (pt, converted back from px); returns undefined when there is no text. */
function nodeMaxFontPt(n: RenderNode): number | undefined {
  if (n.type !== 'shape' && n.type !== 'text') return undefined
  let maxPx = 0
  for (const line of (n as ShapeRenderNode).text?.lines ?? []) {
    for (const r of line.runs) if (r.fontSizePx > maxPx) maxPx = r.fontSizePx
  }
  return maxPx > 0 ? Math.round((maxPx * 72) / 96) : undefined
}

/** Normalize a render color to #RRGGBB (strips alpha); undefined when not a hex color. */
function hex6(color: string | undefined): string | undefined {
  if (!color) return undefined
  const m = /^#([0-9a-fA-F]{6})/.exec(color.trim())
  return m ? `#${m[1].toUpperCase()}` : undefined
}

/** Dominant text color = the run color covering the most characters (bullets excluded). */
function dominantTextColor(n: RenderNode): string | undefined {
  if (n.type !== 'shape' && n.type !== 'text') return undefined
  const weight = new Map<string, number>()
  for (const line of (n as ShapeRenderNode).text?.lines ?? []) {
    for (const r of line.runs) {
      if (r.isBullet) continue
      const c = hex6(r.color)
      if (c) weight.set(c, (weight.get(c) ?? 0) + r.text.length)
    }
  }
  let best: string | undefined
  let max = 0
  for (const [c, w] of weight) {
    if (w > max) {
      best = c
      max = w
    }
  }
  return best
}

/** Readable colors of a node (solid fill / dominant text color / stroke); pictures only expose stroke. */
function nodeColors(n: RenderNode): Pick<DeckNodeInfo, 'fill' | 'textColor' | 'strokeColor'> {
  const out: Pick<DeckNodeInfo, 'fill' | 'textColor' | 'strokeColor'> = {}
  if (n.type === 'shape' || n.type === 'text') {
    const s = n as ShapeRenderNode
    if (s.fill.kind === 'solid') {
      const c = hex6(s.fill.color)
      if (c) out.fill = c
    }
    const stroke = hex6(s.stroke?.color)
    if (stroke) out.strokeColor = stroke
    const text = dominantTextColor(n)
    if (text) out.textColor = text
  } else if (n.type === 'picture') {
    const stroke = hex6((n as PictureRenderNode).stroke?.color)
    if (stroke) out.strokeColor = stroke
  }
  return out
}

/**
 * Collect node info (including nested group children). A child's box is in group-local
 * coordinates (ext/chExt scaling already baked into geometry at build time); here we add the
 * group offset to convert to absolute coordinates and set the inGroup flag. Direct children of a
 * top-level group also carry groupId (editable via the in-group pipeline); deeper nesting stays
 * read-only (the main process patches one level only).
 */
export function collectNodeInfos(
  nodes: RenderNode[],
  ox = 0,
  oy = 0,
  parent?: { id: string; topLevel: boolean },
): DeckNodeInfo[] {
  const out: DeckNodeInfo[] = []
  for (const n of nodes) {
    const b = n.box
    const abs = {
      x: Math.round(ox + b.x),
      y: Math.round(oy + b.y),
      w: Math.round(b.w),
      h: Math.round(b.h),
    }
    const base: DeckNodeInfo = {
      // Durable id when the element's bytes carry one — survives regenerate/
      // ungroup/save, so the AI can keep addressing across turns
      id: n.durableId ?? n.sourceId,
      type: n.type,
      text: nodeText(n),
      ...abs,
      rotation: b.rotationDeg,
      ...(parent ? { inGroup: true } : {}),
      ...(parent?.topLevel ? { groupId: parent.id } : {}),
      ...(n.decoration ? { locked: true } : {}),
      ...nodeColors(n),
    }
    const fontPt = nodeMaxFontPt(n)
    if (fontPt !== undefined) base.fontSizePt = fontPt
    out.push(base)
    if (n.type === 'group') {
      const g = n as GroupRenderNode
      out.push(
        ...collectNodeInfos(g.children, abs.x, abs.y, {
          id: n.durableId ?? n.sourceId,
          topLevel: !parent,
        }),
      )
    }
  }
  return out
}

export function preview(text: string, max = 50): string {
  const flat = text.replace(/\n/g, ' / ')
  return flat.length > max ? `${flat.slice(0, max)}…` : flat
}

function pushPageLines(lines: string[], slide: RenderSlide, i: number): void {
  lines.push(`Page ${i + 1} (slideIndex=${i}):`)
  const infos = collectNodeInfos(slide.nodes)
  const fillCount = new Map<string, number>()
  for (const n of infos) {
    if (n.fill) fillCount.set(n.fill, (fillCount.get(n.fill) ?? 0) + 1)
  }
  const mainFills = [...fillCount.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([c, count]) => (count > 1 ? `${c}×${count}` : c))
  if (mainFills.length > 0) lines.push(`  main fills: ${mainFills.join(' ')}`)
  for (const n of infos) {
    lines.push(`  - ${n.id} | ${n.type}${n.text ? ` | "${preview(n.text)}"` : ''}`)
  }
}

export function buildDeckOutline(slides: RenderSlide[], current: number, selectedIds: string[]): string {
  const canvas = slides[0] ? `Canvas ${slides[0].widthPx}×${slides[0].heightPx}px.` : ''
  const lines: string[] = [
    `The presentation has ${slides.length} pages; page ${current + 1} is currently shown. ${canvas}`,
    `(Page order is the current actual order and may differ from generation time or earlier conversation; the user's "page N" refers to this outline)`,
  ]
  if (selectedIds.length > 0) {
    const currentSlide = slides[current]
    const selectedRefs = selectedIds.map((id) => {
      const node = currentSlide ? findNodeById(currentSlide.nodes, id) : undefined
      return node?.durableId ?? node?.sourceId ?? id
    })
    lines.push(`User selected elements: ${selectedRefs.join(', ')}`)
  }
  slides.forEach((slide, i) => pushPageLines(lines, slide, i))
  lines.push('(Use read_slide to see element positions/sizes/colors)')
  return lines.join('\n')
}

/** Host (embed agent) variant: no current page or user selection. */
export function buildAgentDeckOutline(slides: RenderSlide[]): string {
  const canvas = slides[0] ? `Canvas ${slides[0].widthPx}×${slides[0].heightPx}px.` : ''
  const lines: string[] = [
    `The presentation has ${slides.length} pages. ${canvas}`,
    `(Page order is the current actual order; the user's "page N" refers to this outline)`,
  ]
  slides.forEach((slide, i) => pushPageLines(lines, slide, i))
  return lines.join('\n')
}

/**
 * Element inventory of one slide as read_slide reports it (ids + geometry + colors + text).
 * Shared by the read_slide tool and the post-generation layout QC pass (slide-qc.ts), so the
 * QC model maps screenshot pixels back to the same ids/coordinates the edit tools accept.
 */
export function formatSlideDump(slide: RenderSlide): string {
  const infos = collectNodeInfos(slide.nodes)
  const parts = infos.map((n) => {
    const flags = [
      n.groupId
        ? `in group ${n.groupId} (directly editable)`
        : n.inGroup
          ? 'nested in a sub-group (read-only; ungroup_element the outer group to edit)'
          : '',
      n.locked ? 'layout decoration (read-only)' : '',
    ]
      .filter(Boolean)
      .join(' ')
    const rot = n.rotation ? ` rotation ${Math.round(n.rotation)}°` : ''
    const font = n.fontSizePt ? ` font ${n.fontSizePt}pt` : ''
    const colors = [
      n.fill ? `fill${n.fill}` : '',
      n.textColor ? `text${n.textColor}` : '',
      n.strokeColor ? `stroke${n.strokeColor}` : '',
    ]
      .filter(Boolean)
      .join(' ')
    const head = `${n.id} | ${n.type}${flags ? ` | ${flags}` : ''} | pos(${n.x},${n.y}) size ${n.w}×${n.h}${rot}${font}${colors ? ` | ${colors}` : ''}`
    return n.text ? `${head}\n${n.text}` : `${head} | (no text)`
  })
  const colorlessTypes = [
    ...new Set(
      infos
        .filter((n) => !n.fill && !n.textColor && !n.strokeColor)
        .filter((n) => n.type === 'picture' || n.type === 'chart')
        .map((n) => n.type),
    ),
  ]
  const colorNote = colorlessTypes.length
    ? `\n(${colorlessTypes.join('/')} colors not available)`
    : ''
  // Report the real px→EMU factor: render px carry the viewport scale, so ×9525 only
  // holds for decks whose baseline width is exactly the fit width (standard 16:9 at 1280).
  const pxToEmu = +(9525 / slide.scale).toFixed(2)
  return `Canvas ${slide.widthPx}×${slide.heightPx}px (1 px = ${pxToEmu} EMU)\n${parts.join('\n---\n') || '(no elements on this page)'}${colorNote}`
}
