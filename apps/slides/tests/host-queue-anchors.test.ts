/**
 * Host annotation queue → element geometry: a queued note carries a page hint
 * plus durable element ids, so anchors must survive page reorder, fold group
 * offsets into slide space, and drop notes whose element no longer exists.
 */
import { describe, expect, it } from 'vitest'
import type { PlacedBox, RenderSlide } from '@genoffice/pptx-render'
import {
  ASK_EXCERPT_MAX,
  buildElementExcerpt,
  findElements,
  flattenElements,
  resolveAskQueue,
  resolveAskQueueItem,
  unionBox,
} from '../src/renderer/ai/host-queue-anchors'
import type { SlidesAskQueueItem } from '../src/shared/ipc'

const box = (x: number, y: number, w: number, h: number): PlacedBox => ({
  x,
  y,
  w,
  h,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})

function slide(nodes: unknown[]): RenderSlide {
  return { widthPx: 1280, heightPx: 720, nodes } as unknown as RenderSlide
}

const title = {
  id: 't1',
  sourceId: 't1',
  durableId: 'dur-title',
  type: 'shape',
  box: box(80, 60, 600, 120),
  text: { lines: [{ runs: [{ text: 'Quarterly review', fontSizePx: 24 }] }] },
}
const picture = { id: 'p1', sourceId: 'p1', type: 'picture', box: box(700, 60, 400, 300) }
const grouped = {
  id: 'g1',
  sourceId: 'g1',
  durableId: 'dur-group',
  type: 'group',
  box: box(100, 400, 300, 200),
  children: [
    {
      id: 'c1',
      sourceId: 'c1',
      durableId: 'dur-child',
      type: 'shape',
      box: box(20, 30, 100, 50),
      text: { lines: [{ runs: [{ text: 'inside group', fontSizePx: 18 }] }] },
    },
  ],
}

const item = (over: Partial<SlidesAskQueueItem> = {}): SlidesAskQueueItem => ({
  id: 'd1',
  instruction: 'shrink it',
  slideIndex: 0,
  elementIds: ['dur-title'],
  ...over,
})

describe('flattenElements', () => {
  it('folds group offsets into slide-space coordinates', () => {
    const flat = flattenElements([grouped as never])
    const child = flat.find((placement) => placement.node.sourceId === 'c1')
    expect(child).toBeDefined()
    expect(child?.x).toBe(120)
    expect(child?.y).toBe(430)
  })
})

describe('findElements', () => {
  it('matches durable ids first and falls back to source ids', () => {
    expect(
      findElements(slide([title, picture]), ['dur-title'])?.map((p) => p.node.sourceId),
    ).toEqual(['t1'])
    expect(findElements(slide([title, picture]), ['p1'])?.map((p) => p.node.sourceId)).toEqual([
      'p1',
    ])
  })

  it('returns null when any anchored element is missing', () => {
    expect(findElements(slide([title]), ['dur-title', 'gone'])).toBeNull()
    expect(findElements(slide([]), [])).toBeNull()
  })
})

describe('resolveAskQueueItem', () => {
  it('resolves on the hinted page', () => {
    const resolved = resolveAskQueueItem([slide([title])], item())
    expect(resolved?.slideIndex).toBe(0)
    expect(resolved?.placements[0]?.node.sourceId).toBe('t1')
  })

  it('re-derives the page after pages are reordered', () => {
    const slides = [slide([picture]), slide([]), slide([title])]
    const resolved = resolveAskQueueItem(slides, item({ slideIndex: 0 }))
    expect(resolved?.slideIndex).toBe(2)
  })

  it('drops a note whose elements are gone', () => {
    expect(resolveAskQueueItem([slide([picture])], item())).toBeNull()
  })

  it('drops ambiguous multi-page matches instead of guessing', () => {
    // Durable ids repeat across slides for decks without creation ids; a match
    // on two non-hinted pages is ambiguous.
    const duplicate = { ...title, box: box(0, 0, 10, 10) }
    const slides = [slide([]), slide([title]), slide([duplicate])]
    expect(resolveAskQueueItem(slides, item({ slideIndex: 0 }))).toBeNull()
  })

  it('treats missing element ids as unrenderable (host flat payload)', () => {
    expect(resolveAskQueueItem([slide([title])], item({ elementIds: undefined }))).toBeNull()
  })

  it('resolves a group child through the embedded durable id', () => {
    const resolved = resolveAskQueueItem([slide([grouped])], item({ elementIds: ['dur-child'] }))
    expect(resolved?.placements[0]).toMatchObject({ x: 120, y: 430 })
  })
})

describe('resolveAskQueue', () => {
  it('keeps resolvable notes and drops the rest, preserving queue order', () => {
    const slides = [slide([title, picture])]
    const resolved = resolveAskQueue(slides, [
      item({ id: 'd1', elementIds: ['dur-title'] }),
      item({ id: 'd2', elementIds: ['gone'] }),
      item({ id: 'd3', elementIds: ['p1'] }),
    ])
    expect(resolved.map((anchor) => anchor.item.id)).toEqual(['d1', 'd3'])
  })
})

describe('unionBox', () => {
  it('covers every anchored element in slide space', () => {
    const placements = findElements(slide([title, picture]), ['dur-title', 'p1']) ?? []
    expect(unionBox(placements)).toEqual({ x: 80, y: 60, w: 1020, h: 300 })
  })

  it('returns null for an empty placement list', () => {
    expect(unionBox([])).toBeNull()
  })
})

describe('buildElementExcerpt', () => {
  it('joins element text and collapses whitespace', () => {
    expect(buildElementExcerpt([{ text: ' Quarterly  review ' }, { text: 'Q3 numbers' }])).toBe(
      'Quarterly review Q3 numbers',
    )
  })

  it('skips textless elements and caps the excerpt', () => {
    expect(buildElementExcerpt([{}, { text: 'a' }])).toBe('a')
    expect(buildElementExcerpt([{ text: 'word '.repeat(ASK_EXCERPT_MAX) }])).toHaveLength(
      ASK_EXCERPT_MAX,
    )
    expect(buildElementExcerpt([])).toBe('')
  })
})
