/**
 * Local page generation: LLM spec JSON → parse/validate → build a one-slide
 * pptx directly with pptx-engine primitives. The build tests reopen the bytes
 * with openPptx and assert on the parsed model (true roundtrip, no mocks).
 */
import { describe, it, expect } from 'vitest'
import { openPptx, type TextElement, type PictureElement } from '@genoffice/pptx-engine'
import { HeuristicMetrics } from '@genoffice/pptx-render'
import { parsePageSpec, buildPagePptx, type PageSpec } from '../src/main/page-spec'

// 1x1 red PNG
const PNG_1PX = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
)

const textSpec = (text: string, extra: Record<string, unknown> = {}) => ({
  type: 'text',
  x: 80,
  y: 60,
  w: 800,
  h: 90,
  paragraphs: [{ runs: [{ text, sizePt: 32, bold: true, color: '#112233' }] }],
  ...extra,
})

describe('parsePageSpec', () => {
  it('accepts fenced JSON with junk around it', () => {
    const raw =
      'Here is the design:\n```json\n{"background":"#0E1A2B","elements":[' +
      JSON.stringify(textSpec('Hello')) +
      ']}\n```\nDone.'
    const r = parsePageSpec(raw)
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.spec.background).toBe('#0E1A2B')
    expect(r.spec.elements).toHaveLength(1)
  })

  it('rejects output without a usable JSON object', () => {
    expect(parsePageSpec('sorry, I cannot').ok).toBe(false)
    expect(parsePageSpec('{"elements":[]}').ok).toBe(false)
    const bad = parsePageSpec('{"elements":[{"type":"text","x":0,"y":0,"w":100,"h":40}]}')
    expect(bad.ok).toBe(false) // text without any runs → all elements dropped
  })

  it('clamps out-of-canvas boxes and drops vanishing ones with warnings', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          { ...textSpec('kept'), x: 1200, w: 400 }, // clamped to 80px wide
          { ...textSpec('gone'), x: 5000, y: 5000 },
          { type: 'shape', shape: 'rect', x: 0, y: 0, w: 100, h: 100, fill: '#FFF' },
        ],
      }),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.spec.elements).toHaveLength(2)
    const kept = r.spec.elements[0]!
    expect(kept.x + kept.w).toBeLessThanOrEqual(1280)
    expect(r.warnings.some((w) => w.includes('outside'))).toBe(true)
    // #FFF expands to #FFFFFF
    expect((r.spec.elements[1] as { fill?: string }).fill).toBe('#FFFFFF')
  })

  it('falls back to rect for unknown shapes and drops invisible ones', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          { type: 'shape', shape: 'wavyMagicBlob', x: 0, y: 0, w: 10, h: 10, fill: '#123456' },
          { type: 'shape', shape: 'rect', x: 0, y: 0, w: 10, h: 10 }, // no fill/stroke → dropped
          { type: 'image', url: 'ftp://nope', x: 0, y: 0, w: 10, h: 10 }, // bad scheme → dropped
          textSpec('t'),
        ],
      }),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.spec.elements).toHaveLength(2)
    expect((r.spec.elements[0] as { shape: string }).shape).toBe('rect')
  })

  it('rejects emoji anywhere in text (pictographs, VS16, forced symbols) but keeps bare ★', () => {
    const withEmoji = parsePageSpec(
      JSON.stringify({ elements: [textSpec('季度目标 🎯 达成')] }),
    )
    expect(withEmoji.ok).toBe(false)
    if (!withEmoji.ok) expect(withEmoji.error).toContain('emoji')

    const forcedStar = parsePageSpec(
      JSON.stringify({ elements: [textSpec('重点☀️关注')] }),
    )
    expect(forcedStar.ok).toBe(false)

    const bareStar = parsePageSpec(JSON.stringify({ elements: [textSpec('重点 ★ 关注')] }))
    expect(bareStar.ok).toBe(true)
  })

  it('rejects text elements whose ink rectangles overlap, naming both indices', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          textSpec('标题占位', { y: 100, h: 60 }),
          textSpec('副标题占位', { y: 130, h: 60 }),
        ],
      }),
    )
    expect(r.ok).toBe(false)
    if (!r.ok) {
      expect(r.error).toContain('element 0')
      expect(r.error).toContain('element 1')
      expect(r.error).toContain('overlaps')
    }
  })

  it('rejects a labeled shape colliding with a text box, counting shape label ink', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          {
            type: 'shape',
            shape: 'rect',
            x: 80,
            y: 300,
            w: 400,
            h: 60,
            fill: '#EEF2F7',
            paragraphs: [{ runs: [{ text: '卡片标题', sizePt: 14 }] }],
          },
          textSpec('正文从同高度穿过', { x: 80, y: 310, w: 700, h: 40 }),
        ],
      }),
    )
    expect(r.ok).toBe(false)
  })

  it('accepts a text box sitting on a plain (text-less) card shape and non-colliding neighbors', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          { type: 'shape', shape: 'roundRect', x: 60, y: 80, w: 500, h: 300, fill: '#EEF2F7' },
          textSpec('卡片内标题', { x: 88, y: 108, w: 300, h: 60 }),
          textSpec('卡片内正文', { x: 88, y: 180, w: 440, h: 150 }),
        ],
      }),
    )
    expect(r.ok).toBe(true)
  })
})

describe('buildPagePptx', () => {
  const noImages = { fetchImage: async () => null }

  it('builds text/shape/background into a reopenable one-slide pptx', async () => {
    const spec: PageSpec = {
      background: '#0E1A2B',
      elements: [
        {
          type: 'shape',
          shape: 'roundRect',
          x: 80,
          y: 200,
          w: 400,
          h: 200,
          fill: '#FFFFFF14',
          stroke: { color: '#3B82F6', widthPt: 1 },
        },
        {
          type: 'text',
          x: 80,
          y: 60,
          w: 800,
          h: 90,
          paragraphs: [
            {
              align: 'left',
              lineSpacingPct: 110,
              runs: [{ text: 'Quarterly Wins', sizePt: 36, bold: true, color: '#FFFFFF' }],
            },
          ],
        },
      ],
    }
    const { bytes, imageFailures } = await buildPagePptx(spec, noImages)
    expect(imageFailures).toEqual([])
    const opened = await openPptx(bytes)
    expect(opened.deck.slides).toHaveLength(1)
    const slide = opened.deck.slides[0]!
    // The full-bleed background rect is promoted to the slide background at build time
    const texts = slide.elements.filter((e): e is TextElement => e.type === 'text')
    const shapes = slide.elements.filter((e) => e.type === 'shape')
    expect(shapes.length).toBeGreaterThanOrEqual(1)
    expect(texts).toHaveLength(1)
    const run = texts[0]!.text!.paragraphs[0]!.runs[0]!
    expect(run.text).toBe('Quarterly Wins')
    expect(run.bold).toBe(true)
    expect(run.fontSize).toBe(36)
    // 80px at the deck's 1280px-wide canvas = 80 * 9525 EMU
    expect(texts[0]!.transform.offset.x).toBe(80 * 9525)
  })

  it('adds images with cover-crop and reports failed downloads without failing the page', async () => {
    const spec: PageSpec = {
      elements: [
        { type: 'image', url: 'https://ok.example/a.png', x: 0, y: 0, w: 640, h: 720 },
        { type: 'image', url: 'https://dead.example/b.png', x: 640, y: 0, w: 640, h: 720 },
        { type: 'text', x: 100, y: 100, w: 400, h: 60, paragraphs: [{ runs: [{ text: 'cap' }] }] },
      ],
    }
    const { bytes, imageFailures } = await buildPagePptx(spec, {
      fetchImage: async (url) =>
        url.includes('ok.example') ? { bytes: PNG_1PX, ext: 'png' } : null,
      imageDims: () => ({ width: 200, height: 100 }),
    })
    expect(imageFailures).toEqual(['https://dead.example/b.png'])
    const opened = await openPptx(bytes)
    const pics = opened.deck.slides[0]!.elements.filter(
      (e): e is PictureElement => e.type === 'picture',
    )
    expect(pics).toHaveLength(1)
    // 200x100 source into a 640x720 portrait frame → horizontal crop applied
    expect(pics[0]!.srcRect?.l ?? 0).toBeGreaterThan(0)
  })
})

describe('buildPagePptx text-box height fix', () => {
  const deps = { fetchImage: async () => null, fontMetrics: new HeuristicMetrics() }
  const EMU_PER_PX = 9525
  const px = (emu: number) => emu / EMU_PER_PX
  const bigTitle = (extra: Record<string, unknown> = {}) => ({
    type: 'text' as const,
    x: 100,
    y: 100,
    w: 600,
    h: 30,
    paragraphs: [{ runs: [{ text: '成都来了就不想走', sizePt: 40, bold: true }] }],
    ...extra,
  })
  // 40pt = 53.33px glyphs on the default 1.2em line box → one line ≈ 64px
  const oneLinePx = 40 * (96 / 72) * 1.2

  it('grows an undersized top-anchored box to the measured content height', async () => {
    const { bytes } = await buildPagePptx({ elements: [bigTitle()] }, deps)
    const opened = await openPptx(bytes)
    const el = opened.deck.slides[0]!.elements.find((e): e is TextElement => e.type === 'text')!
    expect(px(el.transform.offset.cy)).toBeCloseTo(oneLinePx, 0)
    // Top anchor: the box only grows downward, glyphs don't move
    expect(el.transform.offset.y).toBe(100 * EMU_PER_PX)
  })

  it('shifts a middle-anchored box up so the rendered glyphs stay in place', async () => {
    const { bytes } = await buildPagePptx({ elements: [bigTitle({ valign: 'middle' })] }, deps)
    const opened = await openPptx(bytes)
    const el = opened.deck.slides[0]!.elements.find((e): e is TextElement => e.type === 'text')!
    expect(px(el.transform.offset.cy)).toBeCloseTo(oneLinePx, 0)
    expect(px(el.transform.offset.y)).toBeCloseTo(100 - (oneLinePx - 30) / 2, 0)
  })

  it('leaves tall-enough boxes and undersized shape labels untouched', async () => {
    const spec: PageSpec = {
      elements: [
        bigTitle({ h: 100 }),
        {
          type: 'shape',
          shape: 'roundRect',
          x: 100,
          y: 400,
          w: 600,
          h: 30,
          fill: '#FFFFFF',
          paragraphs: [{ runs: [{ text: '成都来了就不想走', sizePt: 40 }] }],
        },
      ],
    }
    const { bytes } = await buildPagePptx(spec, deps)
    const opened = await openPptx(bytes)
    const slide = opened.deck.slides[0]!
    const text = slide.elements.find((e): e is TextElement => e.type === 'text')!
    const shape = slide.elements.find((e) => e.type === 'shape')!
    // Content (≈64px) fits the 100px box → no change; shape height is design intent
    expect(text.transform.offset.cy).toBe(100 * EMU_PER_PX)
    expect(shape.transform.offset.cy).toBe(30 * EMU_PER_PX)
  })

  it('skips the fix entirely when no font metrics are injected', async () => {
    const { bytes } = await buildPagePptx(
      { elements: [bigTitle()] },
      { fetchImage: async () => null },
    )
    const opened = await openPptx(bytes)
    const el = opened.deck.slides[0]!.elements.find((e): e is TextElement => e.type === 'text')!
    expect(el.transform.offset.cy).toBe(30 * EMU_PER_PX)
  })
})
