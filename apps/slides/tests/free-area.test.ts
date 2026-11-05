import { describe, expect, it } from 'vitest'

import { pickFreeBox, type Box } from '../src/renderer/ai/free-area'

/** 实测的那页版式：左栏正文 0.67~6.77in，右栏空着。 */
function slideLikeBoxes(): Box[] {
  const IN = 96 // 96dpi 下 1 英寸
  return [
    { x: 0.67 * IN, y: 0.44 * IN, w: 12 * IN, h: 0.7 * IN },   // 标题
    { x: 0.67 * IN, y: 1.82 * IN, w: 6.1 * IN, h: 0.5 * IN },   // 提问句
    { x: 0.67 * IN, y: 2.36 * IN, w: 6.1 * IN, h: 2.3 * IN },   // 做法正文（被压的那块）
    { x: 0.67 * IN, y: 5.62 * IN, w: 3.83 * IN, h: 1.55 * IN }, // 卡片
  ]
}

function overlaps(a: Box, b: Box): number {
  const x = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x))
  const y = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y))
  return x * y
}

describe('新元素落位：优先放空白处', () => {
  it('左栏被正文占满时，图表落到右栏而不是压住正文', () => {
    const occupied = slideLikeBoxes()
    const picked = pickFreeBox({ pageW: 1280, pageH: 720, w: 540, h: 300, occupied })
    const totalOverlap = occupied.reduce((sum, box) => sum + overlaps(picked, box), 0)
    expect(totalOverlap).toBe(0)
    // 右栏：左边界应落在正文右边界（6.77in ≈ 650px）之后
    expect(picked.x).toBeGreaterThan(6.1 * 96)
  })

  it('页面全空时保持居中（不无谓地推到角落）', () => {
    const picked = pickFreeBox({ pageW: 1280, pageH: 720, w: 400, h: 200, occupied: [] })
    expect(picked.x).toBeGreaterThan(300)
    expect(picked.x).toBeLessThan(700)
    expect(picked.y).toBeGreaterThan(150)
    expect(picked.y).toBeLessThan(450)
  })

  it('放不下时退而求其次：挑重叠面积最小的位置', () => {
    const occupied: Box[] = [{ x: 100, y: 100, w: 900, h: 500 }]
    const picked = pickFreeBox({ pageW: 1280, pageH: 720, w: 500, h: 300, occupied })
    const overlap = overlaps(picked, occupied[0]!)
    const best = 900 * 500 // 完全被覆盖时的面积
    expect(overlap).toBeLessThan(best)
  })

  it('元素比页面还大时裁到页边留白内，不越界', () => {
    const picked = pickFreeBox({ pageW: 1280, pageH: 720, w: 5000, h: 5000, occupied: [] })
    expect(picked.x).toBeGreaterThanOrEqual(0)
    expect(picked.y).toBeGreaterThanOrEqual(0)
    expect(picked.x + picked.w).toBeLessThanOrEqual(1280)
    expect(picked.y + picked.h).toBeLessThanOrEqual(720)
  })
})
