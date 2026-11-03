import { describe, expect, it } from 'vitest'
import type { Block, InkInfo } from '@genoffice/docx-engine'
import { pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import {
  annotationsFromParsed,
  buildInkImages,
  decodeInkPayload,
  encodeInkPayload,
  eraseStrokeAt,
  strokeBounds,
  strokeHitTest,
  type InkAnnotation,
  type InkRaster,
  type InkStroke,
} from '../src/renderer/editor/ink'

const STROKE: InkStroke = {
  tool: 'pen',
  color: 'C00000',
  width: 2,
  points: [
    { x: 10, y: 5 },
    { x: 60, y: 25 },
  ],
}

describe('ink payload round-trip', () => {
  it('encodes and decodes a stroke', () => {
    const decoded = decodeInkPayload(encodeInkPayload(STROKE))
    expect(decoded).toEqual(STROKE)
  })

  it('rejects malformed / foreign payloads', () => {
    expect(decodeInkPayload('not json')).toBeNull()
    expect(decodeInkPayload('{"v":99,"points":[[0,0]]}')).toBeNull()
    expect(decodeInkPayload('{"v":1,"points":[]}')).toBeNull()
  })
})

describe('annotationsFromParsed', () => {
  const base: Omit<InkInfo, 'payload' | 'dataUrl'> = {
    anchorIndex: 3,
    offsetXPx: 8,
    offsetYPx: -4,
    widthPx: 100,
    heightPx: 40,
  }

  it('restores editable strokes from the payload', () => {
    const [ann] = annotationsFromParsed([
      { ...base, dataUrl: 'data:image/png;base64,x', payload: encodeInkPayload(STROKE) },
    ])
    expect(ann.anchorIndex).toBe(3)
    expect(ann.stroke).toEqual(STROKE)
    expect(ann.image).toBeUndefined()
  })

  it('falls back to an image annotation when the payload is unreadable', () => {
    const [ann] = annotationsFromParsed([
      { ...base, dataUrl: 'data:image/png;base64,x', payload: null },
    ])
    expect(ann.stroke).toBeUndefined()
    expect(ann.image).toEqual({
      dataUrl: 'data:image/png;base64,x',
      offsetXPx: 8,
      offsetYPx: -4,
      widthPx: 100,
      heightPx: 40,
    })
  })
})

describe('stroke geometry', () => {
  it('bounds include the line width padding', () => {
    const b = strokeBounds(STROKE)
    expect(b.x).toBe(10 - 2)
    expect(b.y).toBe(5 - 2)
    expect(b.w).toBe(50 + 4)
    expect(b.h).toBe(20 + 4)
  })

  it('hit test matches near the segment and misses far away', () => {
    expect(strokeHitTest(STROKE, { x: 35, y: 15 }, 3)).toBe(true)
    expect(strokeHitTest(STROKE, { x: 35, y: 40 }, 3)).toBe(false)
  })
})

describe('partial erasing (eraseStrokeAt)', () => {
  const LINE: InkStroke = {
    tool: 'pen',
    color: '000000',
    width: 2,
    points: [
      { x: 0, y: 0 },
      { x: 10, y: 0 },
      { x: 20, y: 0 },
      { x: 30, y: 0 },
    ],
  }

  it('erasing mid-stroke splits it into two strokes instead of dropping everything', () => {
    const fragments = eraseStrokeAt(LINE, { x: 20, y: 0 }, 1)
    expect(fragments).not.toBeNull()
    expect(fragments).toHaveLength(2)
    expect(fragments![0]).toEqual({
      ...LINE,
      points: [{ x: 0, y: 0 }, { x: 10, y: 0 }, { x: 18, y: 0 }],
    })
    expect(fragments![1]).toEqual({
      ...LINE,
      points: [{ x: 22, y: 0 }, { x: 30, y: 0 }],
    })
    // style is preserved on both fragments
    expect(fragments![0]!.tool).toBe('pen')
    expect(fragments![1]!.color).toBe('000000')
  })

  it('erasing near the end trims only the covered part', () => {
    const fragments = eraseStrokeAt(LINE, { x: 0, y: 0 }, 1)
    expect(fragments).toHaveLength(1)
    expect(fragments![0]!.points).toEqual([
      { x: 2, y: 0 },
      { x: 10, y: 0 },
      { x: 20, y: 0 },
      { x: 30, y: 0 },
    ])
  })

  it('a click away from the stroke leaves it untouched (null)', () => {
    expect(eraseStrokeAt(LINE, { x: 20, y: 40 }, 1)).toBeNull()
  })

  it('covering the whole stroke returns no fragments', () => {
    const short: InkStroke = { ...LINE, points: [{ x: 0, y: 0 }, { x: 1, y: 0 }] }
    expect(eraseStrokeAt(short, { x: 0, y: 0 }, 4)).toEqual([])
  })

  it('wider strokes (highlighter) erase with the stroke width taken into account', () => {
    const highlighter: InkStroke = { ...LINE, tool: 'highlighter', width: 10 }
    const fragments = eraseStrokeAt(highlighter, { x: 20, y: 0 }, 1)
    expect(fragments![0]!.points.at(-1)).toEqual({ x: 14, y: 0 })
    expect(fragments![1]!.points[0]).toEqual({ x: 26, y: 0 })
  })
})

describe('buildInkImages', () => {
  const fakeRaster = (stroke: InkStroke): InkRaster => {
    const b = strokeBounds(stroke)
    return { base64: 'PNGDATA', widthPx: b.w, heightPx: b.h, offsetXPx: b.x, offsetYPx: b.y }
  }

  it('maps docx anchors to save-block indexes and encodes the payload', () => {
    const anns: InkAnnotation[] = [{ id: 'a', anchorIndex: 7, stroke: STROKE }]
    const out = buildInkImages(anns, new Map([[7, 2]]), fakeRaster)
    expect(out).toHaveLength(1)
    expect(out[0].blockIndex).toBe(2)
    expect(out[0].base64).toBe('PNGDATA')
    expect(decodeInkPayload(out[0].payload!)).toEqual(STROKE)
  })

  it('drops annotations whose anchor paragraph was deleted', () => {
    const anns: InkAnnotation[] = [{ id: 'a', anchorIndex: 7, stroke: STROKE }]
    expect(buildInkImages(anns, new Map(), fakeRaster)).toHaveLength(0)
  })

  it('carries restored image annotations through unchanged', () => {
    const anns: InkAnnotation[] = [
      {
        id: 'a',
        anchorIndex: 1,
        image: {
          dataUrl: 'data:image/png;base64,QUJD',
          offsetXPx: 5,
          offsetYPx: 6,
          widthPx: 70,
          heightPx: 30,
        },
      },
    ]
    const out = buildInkImages(anns, new Map([[1, 0]]), fakeRaster)
    expect(out).toEqual([
      {
        blockIndex: 0,
        base64: 'QUJD',
        widthPx: 70,
        heightPx: 30,
        offsetXPx: 5,
        offsetYPx: 6,
      },
    ])
  })
})

describe('pmDocToSavePlan saveBlockIndexByDocx', () => {
  const para = (docxIndex: number, text: string): Block => ({
    id: `b${docxIndex}`,
    docxIndex,
    originalXml: `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`,
    type: 'paragraph',
    runs: [{ text }],
  })
  const pmPara = (docxIndex: number | null, text: string): PmNode => ({
    type: 'docParagraph',
    attrs: { docxIndex },
    content: [{ type: 'text', text }],
  })

  it('maps unchanged, edited and shifted paragraphs to their save positions', () => {
    const originals = [para(0, 'a'), para(1, 'b')]
    const doc: PmNode = {
      type: 'doc',
      content: [
        pmPara(null, 'inserted'), // new paragraph at the top shifts everything
        pmPara(0, 'a'), // unchanged -> original
        pmPara(1, 'b CHANGED'), // edited -> generated, anchor consumed
      ],
    }
    const plan = pmDocToSavePlan(doc, originals)
    expect(plan.saveBlocks[1]).toEqual({ kind: 'original', docxIndex: 0 })
    expect(plan.saveBlocks[2].kind).toBe('generated')
    expect(plan.saveBlockIndexByDocx.get(0)).toBe(1)
    expect(plan.saveBlockIndexByDocx.get(1)).toBe(2)
  })

  it('deleted paragraphs get no mapping', () => {
    const originals = [para(0, 'a'), para(1, 'b')]
    const doc: PmNode = { type: 'doc', content: [pmPara(0, 'a')] }
    const plan = pmDocToSavePlan(doc, originals)
    expect(plan.saveBlockIndexByDocx.get(0)).toBe(0)
    expect(plan.saveBlockIndexByDocx.has(1)).toBe(false)
    expect(plan.deletedCount).toBe(1)
  })
})
