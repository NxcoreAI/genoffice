import type { RenderNode, RenderSlide } from '@genoffice/pptx-render'

/**
 * 新插入元素的落位：在页面里找一个**尽量不压住已有内容**的矩形。
 *
 * 之前 `add_chart` 省略 x/y 时直接落在页面正中——实测一份「左文右图」的稿子，
 * 模型删掉右侧图表重建后命令它居中，新图正好盖住左栏正文，画面看起来像"图表被文字
 * 折叠"。这里改成：在网格上试算候选位置，优先选与已有元素零重叠的点，其次看重叠面积，
 * 同样干净时取离页面中心最近的（保持构图不跑偏）。
 */

export interface Box {
  x: number
  y: number
  w: number
  h: number
}

export interface PickFreeBoxOptions {
  pageW: number
  pageH: number
  w: number
  h: number
  occupied: readonly Box[]
  /** 距页边的最小留白。 */
  marginPx?: number
  /** 候选位置步长；越小越准，代价是候选数变多。 */
  stepPx?: number
}

interface Candidate {
  x: number
  y: number
  overlap: number
  centerDistance: number
}

function overlapArea(a: Box, b: Box): number {
  const x = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x))
  const y = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y))
  return x * y
}

export function pickFreeBox(options: PickFreeBoxOptions): Box {
  const { pageW, pageH, occupied } = options
  const margin = options.marginPx ?? Math.round(Math.min(pageW, pageH) * 0.04)
  const step = Math.max(1, options.stepPx ?? 8)
  const w = Math.min(options.w, Math.max(1, pageW - margin * 2))
  const h = Math.min(options.h, Math.max(1, pageH - margin * 2))

  const maxX = Math.max(margin, pageW - margin - w)
  const maxY = Math.max(margin, pageH - margin - h)
  const pageCenterX = pageW / 2
  const pageCenterY = pageH / 2

  let best: Candidate | null = null
  for (let y = margin; y <= maxY; y += step) {
    for (let x = margin; x <= maxX; x += step) {
      const candidate: Box = { x, y, w, h }
      let overlap = 0
      for (const box of occupied) {
        overlap += overlapArea(candidate, box)
        if (overlap > 0 && best && overlap > best.overlap) break
      }
      const centerDistance = Math.hypot(x + w / 2 - pageCenterX, y + h / 2 - pageCenterY)
      if (
        !best
        || overlap < best.overlap
        || (overlap === best.overlap && centerDistance < best.centerDistance)
      ) {
        best = { x, y, overlap, centerDistance }
        if (overlap === 0 && centerDistance === 0) return { x, y, w, h }
      }
      if (overlap === 0 && step >= 8) {
        // 零重叠已经达成，不必把同一行剩余候选都算完——但仍要继续扫，找到更靠中心的那一个。
        continue
      }
    }
  }
  const chosen = best ?? { x: margin, y: margin }
  return { x: Math.round(chosen.x), y: Math.round(chosen.y), w: Math.round(w), h: Math.round(h) }
}

/** 从渲染结果里取参与占位的顶层矩形（跳过母版/版式装饰）。 */
export function occupiedBoxes(slide: RenderSlide): Box[] {
  const out: Box[] = []
  const walk = (nodes: readonly RenderNode[]): void => {
    for (const node of nodes) {
      if (node.decoration) continue
      const { x, y, w, h } = node.box
      out.push({ x, y, w, h })
    }
  }
  walk(slide.nodes)
  return out
}
