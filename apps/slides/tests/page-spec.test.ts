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

describe('PageSpec gradients, shadows, charts and icons', () => {
  const noImages = { fetchImage: async () => null }

  it('parses gradient/shadow tiers, falling back with warnings when invalid', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          {
            type: 'shape', shape: 'rect', x: 0, y: 0, w: 400, h: 300,
            gradient: { stops: [{ pos: 0, color: '#0EA5E9' }, { pos: 1, color: '#1E3A8A' }], angle: 90 },
            shadow: 'soft',
          },
          textSpec('shadowed', { y: 400, shadow: 'strong' }),
          {
            type: 'shape', shape: 'rect', x: 0, y: 500, w: 100, h: 100, fill: '#111111',
            gradient: { stops: [{ pos: 0, color: 'nope' }] }, shadow: 'harsh',
          },
        ],
      }),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    const [a, b, c] = r.spec.elements
    expect(a).toMatchObject({ type: 'shape', shadow: 'soft' })
    expect((a as { gradient?: { angle?: number } }).gradient?.angle).toBe(90)
    expect(b).toMatchObject({ type: 'text', shadow: 'strong' })
    expect((c as { gradient?: unknown, shadow?: unknown }).gradient).toBeUndefined()
    expect((c as { shadow?: unknown }).shadow).toBeUndefined()
    expect(r.warnings.some((w) => w.includes('gradient needs'))).toBe(true)
    expect(r.warnings.some((w) => w.includes('shadow must be'))).toBe(true)
  })

  it('rejects a chart whose series data does not match the categories', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          {
            type: 'chart', chart: 'bar', x: 0, y: 0, w: 500, h: 300,
            categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'rev', values: [1, 2] }],
          },
          textSpec('t', { x: 0, y: 500 }),
        ],
      }),
    )
    expect(r.ok).toBe(false)
    if (r.ok) return
    expect(r.error).toContain('finite numbers')
  })

  it('keeps unknown chart kinds and unknown icons out with warnings', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          {
            type: 'chart', chart: 'doughnut', x: 0, y: 0, w: 400, h: 400,
            categories: ['a', 'b', 'c'], series: [{ name: 's', values: [1, 2, 3] }],
            colors: ['#111111', '#222222', '#333333'], legend: false, dataLabels: true,
          },
          {
            type: 'chart', chart: 'radar', x: 0, y: 450, w: 100, h: 100,
            categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }],
          },
          { type: 'icon', icon: 'check-circle', x: 500, y: 0, w: 48, h: 48, color: '#0EA5E9' },
          { type: 'icon', icon: 'nope', x: 500, y: 100, w: 48, h: 48 },
        ],
      }),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.spec.elements[0]).toMatchObject({
      type: 'chart', chart: 'doughnut', legend: false, dataLabels: true,
    })
    expect(r.spec.elements.filter((e) => e.type === 'icon')).toHaveLength(1)
    expect(r.warnings.some((w) => w.includes('unknown chart'))).toBe(true)
    expect(r.warnings.some((w) => w.includes('unknown icon'))).toBe(true)
  })

  it('builds gradient fills, shadow effects, a real chart part and custGeom icons', async () => {
    const spec: PageSpec = {
      elements: [
        {
          type: 'shape', shape: 'rect', x: 0, y: 0, w: 1280, h: 720,
          gradient: { stops: [{ pos: 0, color: '#F8FAFC' }, { pos: 1, color: '#CBD5E1' }], angle: 90 },
        },
        {
          type: 'shape', shape: 'roundRect', x: 60, y: 60, w: 300, h: 160,
          fill: '#FFFFFF', shadow: 'medium',
        },
        {
          type: 'text', x: 400, y: 40, w: 500, h: 60, shadow: 'soft',
          paragraphs: [{ runs: [{ text: 'Revenue', sizePt: 28, bold: true }] }],
        },
        { type: 'icon', icon: 'chart-bar', x: 340, y: 46, w: 32, h: 32, color: '#0F172A' },
        {
          type: 'chart', chart: 'bar', x: 400, y: 200, w: 800, h: 460, title: 'Quarterly revenue',
          categories: ['Q1', 'Q2', 'Q3', 'Q4'],
          series: [
            { name: '2026', values: [120, 180, 150, 210] },
            { name: '2025', values: [90, 110, 130, 140] },
          ],
          colors: ['#0EA5E9', '#94A3B8'], legend: true, gridlines: true,
        },
      ],
    }
    const { bytes, imageFailures } = await buildPagePptx(spec, noImages)
    expect(imageFailures).toEqual([])
    const opened = await openPptx(bytes)
    const slide = opened.deck.slides[0]!
    const xml = slide.elements.map((e) => e.anchor.originalXml).join('')
    expect(xml).toContain('<a:gradFill')
    expect(xml).toContain('<a:outerShdw')
    expect(xml).toContain('<a:custGeom')
    const chartEntry = [...opened.archive.entries.keys()].find((k) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(k),
    )
    expect(chartEntry).toBeTruthy()
    const chartXml = opened.archive.readText(chartEntry!)
    expect(chartXml).toContain('<c:barChart>')
    expect(chartXml).toContain('<a:srgbClr val="0EA5E9"/>')
    expect(chartXml).toContain('Quarterly revenue')
  })
})

describe('PageSpec tables, scatter and combo charts', () => {
  const noImages = { fetchImage: async () => null }

  it('parses tables with string shorthand, spans and per-cell styling', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          {
            type: 'table', x: 60, y: 200, w: 1160, h: 400, fontSize: 14,
            borderColor: '#CBD5E1', colWidths: [2, 1, 1, 1],
            rows: [
              [
                { text: '维度', bold: true, color: '#FFFFFF', fill: '#0F2A43', align: 'center' },
                { text: 'NexWing', bold: true, color: '#FFFFFF', fill: '#0F2A43', align: 'center' },
                { text: '竞品A', bold: true, color: '#FFFFFF', fill: '#0F2A43', align: 'center' },
                { text: '竞品B', bold: true, color: '#FFFFFF', fill: '#0F2A43', align: 'center' },
              ],
              ['价格', { text: '低', align: 'center' }, { text: '高', align: 'center' }, '高'],
              [
                { text: '综合', span: 3, bold: true, fill: '#F1F5F9' },
                { text: '优', align: 'center' },
              ],
            ],
          },
        ],
      }),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    const t = r.spec.elements[0] as unknown as {
      rows: Array<Array<Record<string, unknown>>>
      fontSize?: number
      borderColor?: string
    }
    expect(t.rows).toHaveLength(3)
    expect(t.rows[0]![0]).toMatchObject({ text: '维度', bold: true, fill: '#0F2A43' })
    expect(t.rows[1]![0]).toEqual({ text: '价格' })
    expect(t.rows[2]![0]).toMatchObject({ span: 3, fill: '#F1F5F9' })
    expect(t.fontSize).toBe(14)
    expect(t.borderColor).toBe('#CBD5E1')
  })

  it('rejects ragged tables and bad colWidths with a page-level error', () => {
    const ragged = parsePageSpec(
      JSON.stringify({
        elements: [
          { type: 'table', x: 0, y: 0, w: 800, h: 300, rows: [['a', 'b'], ['c']] },
          textSpec('t', { x: 0, y: 500 }),
        ],
      }),
    )
    expect(ragged.ok).toBe(false)
    if (ragged.ok) return
    expect(ragged.error).toContain('same number of columns')

    const badWidths = parsePageSpec(
      JSON.stringify({
        elements: [
          {
            type: 'table', x: 0, y: 0, w: 800, h: 300,
            colWidths: [1, 2], rows: [['a', 'b', 'c'], ['d', 'e', 'f']],
          },
          textSpec('t', { x: 0, y: 500 }),
        ],
      }),
    )
    expect(badWidths.ok).toBe(false)
    if (badWidths.ok) return
    expect(badWidths.error).toContain('colWidths needs 3 entries')
  })

  it('builds a real a:tbl graphicFrame that survives a reopen', async () => {
    const spec: PageSpec = {
      elements: [
        {
          type: 'table', x: 60, y: 150, w: 1160, h: 420, fontSize: 13,
          borderColor: '#D8DEE6', colWidths: [2, 1, 1],
          rows: [
            [
              { text: '构型', bold: true, color: '#FFFFFF', fill: '#1B2A4A', align: 'center' },
              { text: '多旋翼', bold: true, align: 'center' },
              { text: '倾转翼', bold: true, align: 'center' },
            ],
            ['续航', '40 min', '120 min'],
            [{ text: '结论', span: 2 }, '倾转翼胜出'],
          ],
        },
      ],
    }
    const { bytes } = await buildPagePptx(spec, noImages)
    const opened = await openPptx(bytes)
    const slide = opened.deck.slides[0]!
    const tableEl = slide.elements.find((e) => e.type === 'table')
    expect(tableEl).toBeTruthy()
    const xml = slide.elements.map((e) => e.anchor.originalXml).join('')
    expect(xml).toContain('<a:tbl>')
    expect(xml).toContain('gridSpan="2"')
    // OOXML: every row lists one tc per grid column; the span origin is
    // followed by a covered hMerge cell (LibreOffice/WPS mis-render without it)
    expect(xml).toContain('hMerge="1"')
    const trs = [...xml.matchAll(/<a:tr[^>]*>(.*?)<\/a:tr>/gs)]
    expect(trs[2]![1].match(/<a:tc[ >]/g)).toHaveLength(3)
    expect(xml).toContain('<a:srgbClr val="1B2A4A"/>')
    expect(xml).toContain('倾转翼胜出')
    // three grid columns, unequal per colWidths [2,1,1]
    const cols = [...xml.matchAll(/<a:gridCol w="(\d+)"\s*\/>/g)].map((m) => Number(m[1]))
    expect(cols).toHaveLength(3)
    expect(cols[1]).toBe(cols[2])
    expect(cols[0]).toBeGreaterThan(cols[1]!)
  })

  it('builds a rule-separated (insideH) table with zebra shading and default 13pt', async () => {
    const spec: PageSpec = {
      elements: [
        {
          type: 'table', x: 60, y: 150, w: 1160, h: 420,
          borderColor: '#CBD5E1', horizontalBordersOnly: true, borderWidthPt: 0.5,
          zebra: '#F1F5F9',
          rows: [
            [{ text: '指标', bold: true }, { text: '本季', bold: true }, { text: '上季', bold: true }],
            ['营收', '1,200', '900'],
            ['毛利', '480', '350'],
            ['客户数', '86', '61'],
          ],
        },
      ],
    }
    const { bytes } = await buildPagePptx(spec, noImages)
    const opened = await openPptx(bytes)
    const xml = opened.deck.slides[0]!.elements.map((e) => e.anchor.originalXml).join('')
    // insideH: no vertical rules anywhere, horizontal rules only between rows
    expect(xml).not.toContain('<a:lnL')
    expect(xml).not.toContain('<a:lnR')
    // header row cells carry no top rule, body rows do (3 body rows × 3 cells)
    expect(xml.match(/<a:lnT/g)?.length).toBe(9)
    // default font size 13 → sz=1300
    expect(xml).toContain('sz="1300"')
    // zebra lands on the 2nd body row only (ri=2)
    const zebraCells = [...xml.matchAll(/<a:tc>(?:(?!<\/a:tc>).)*?val="F1F5F9"(?:(?!<\/a:tc>).)*?<\/a:tc>/gs)]
    expect(zebraCells.length).toBe(3)
  })

  it('builds scatter and comboBarLine chart parts', async () => {
    const spec: PageSpec = {
      elements: [
        {
          type: 'chart', chart: 'scatter', x: 60, y: 80, w: 560, h: 400,
          title: '价格 vs 续航',
          categories: ['30', '55', '80', '120'],
          series: [{ name: '竞品', values: [3.2, 4.0, 4.6, 5.1] }],
          gridlines: true,
        },
        {
          type: 'chart', chart: 'comboBarLine', x: 660, y: 80, w: 560, h: 400,
          title: '收入与增长率',
          categories: ['2026', '2027', '2028', '2029'],
          series: [
            { name: '收入(万)', values: [800, 2400, 5200, 9800] },
            { name: '增长率(%)', values: [0, 200, 117, 88] },
          ],
          legend: true,
        },
      ],
    }
    const { bytes } = await buildPagePptx(spec, noImages)
    const opened = await openPptx(bytes)
    const charts = [...opened.archive.entries.keys()].filter((k) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(k),
    )
    expect(charts).toHaveLength(2)
    const all = charts.map((k) => opened.archive.readText(k)!).join('')
    expect(all).toContain('<c:scatterChart>')
    expect(all).toContain('<c:xVal>')
    expect(all).toContain('<c:barChart>')
    expect(all).toContain('<c:lineChart>')
  })

  it('rejects comboBarLine with a single series', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          {
            type: 'chart', chart: 'comboBarLine', x: 0, y: 0, w: 500, h: 300,
            categories: ['a', 'b'], series: [{ name: 'only', values: [1, 2] }],
          },
          textSpec('t', { x: 0, y: 500 }),
        ],
      }),
    )
    expect(r.ok).toBe(false)
    if (r.ok) return
    expect(r.error).toContain('comboBarLine needs 2-6 series')
  })
})

describe('buildPagePptx dense-page rejection and contain-fit images', () => {
  const EMU_PER_PX = 9525
  const px = (emu: number) => emu / EMU_PER_PX
  const deps = { fetchImage: async () => null, fontMetrics: new HeuristicMetrics() }

  it('rejects a table whose wrapped rows exceed the declared frame', async () => {
    // 8 rows of long CJK cells wrap to ~3 lines each at 13pt in a 200px column,
    // far past the 200px declared frame → the render pass grows the frame and
    // the validator must reject the page with actionable numbers.
    const longRow = ['渠道合作续约条款与返点比例说明', '季度返点按累计开票金额阶梯计算，超额部分单独结算']
    const spec: PageSpec = {
      elements: [
        {
          type: 'table', x: 60, y: 140, w: 400, h: 200, fontSize: 13,
          rows: [longRow, longRow, longRow, longRow, longRow, longRow, longRow, longRow],
        },
      ],
    }
    await expect(buildPagePptx(spec, deps)).rejects.toThrow(/page rejected — .+table content needs \d+px but the frame is \d+px/)
  })

  it('accepts a table whose rows fit the declared frame', async () => {
    const spec: PageSpec = {
      elements: [
        {
          type: 'table', x: 60, y: 140, w: 1160, h: 520, fontSize: 13,
          rows: [
            [{ text: '指标', bold: true }, { text: '本季', bold: true }],
            ['营收', '1,200'], ['毛利', '480'], ['客户数', '86'],
          ],
        },
      ],
    }
    const { bytes } = await buildPagePptx(spec, deps)
    const opened = await openPptx(bytes)
    expect(opened.deck.slides[0]!.elements.some((e) => e.type === 'table')).toBe(true)
  })

  it('rejects text that grows past the slide bottom', async () => {
    // 40pt CJK in a 400px-wide box wraps to ~3 lines (≈192px) but starts at y=600.
    const spec: PageSpec = {
      elements: [
        {
          type: 'text', x: 80, y: 600, w: 400, h: 50,
          paragraphs: [{ runs: [{ text: '这一段文字放在页面底部并且会折成好几行', sizePt: 40 }] }],
        },
      ],
    }
    await expect(buildPagePptx(spec, deps)).rejects.toThrow(/page rejected — .+text grows to \d+px, past the \d+px canvas/)
  })

  it('places fit:contain images at a fitted centered rect without cropping', async () => {
    const spec: PageSpec = {
      elements: [
        { type: 'image', url: 'https://ok.example/a.png', x: 100, y: 100, w: 400, h: 300, fit: 'contain' },
      ],
    }
    const { bytes } = await buildPagePptx(spec, {
      fetchImage: async () => ({ bytes: PNG_1PX, ext: 'png' }),
      imageDims: () => ({ width: 200, height: 100 }), // 2:1 source in a 4:3 frame
    })
    const opened = await openPptx(bytes)
    const pic = opened.deck.slides[0]!.elements.find(
      (e): e is PictureElement => e.type === 'picture',
    )!
    // Fitted to 400x200 and centered vertically in the frame (y 100 → 150)
    expect(px(pic.transform.offset.cx)).toBeCloseTo(400, 0)
    expect(px(pic.transform.offset.cy)).toBeCloseTo(200, 0)
    expect(px(pic.transform.offset.x)).toBeCloseTo(100, 0)
    expect(px(pic.transform.offset.y)).toBeCloseTo(150, 0)
    expect(pic.srcRect ?? undefined).toBeUndefined()
  })
})
