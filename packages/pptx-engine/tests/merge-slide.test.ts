/**
 * mergeSlideFromPptx integration tests (mirrors html→pptx per-slide conversion + deck-level merge):
 * each slide is converted to its own single-slide pptx and merged into an existing deck,
 * without re-converting earlier slides. Uses pptxgenjs to generate real single-slide pptx
 * (same library and structure as the html-pipeline convert-worker), running the real
 * openPptx → mergeSlideFromPptx → savePptx → openPptx chain with no mocks.
 */
import { describe, it, expect } from 'vitest'
import PptxGenJS from 'pptxgenjs'
import { openPptx, savePptx, mergeSlideFromPptx } from '../src/index'

// 1x1 red-dot PNG (base64)
const RED_DOT =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

async function onePagePptx(text: string, withImage = false): Promise<Uint8Array> {
  const p = new PptxGenJS()
  p.defineLayout({ name: 'W', width: 13.333, height: 7.5 })
  p.layout = 'W'
  const s = p.addSlide()
  s.addText(text, { x: 1, y: 1, w: 8, h: 1, fontSize: 32 })
  if (withImage) s.addImage({ data: 'image/png;base64,' + RED_DOT, x: 1, y: 3, w: 2, h: 2 })
  const buf = (await p.write({ outputType: 'nodebuffer' })) as Buffer
  return new Uint8Array(buf)
}

describe('mergeSlideFromPptx', () => {
  it('merges a single-slide pptx into an existing deck, slide count grows and content is kept', async () => {
    const base = await openPptx(await onePagePptx('PAGE_ONE'))
    expect(base.deck.slides.length).toBe(1)

    const s2 = await mergeSlideFromPptx(base, await onePagePptx('PAGE_TWO'))
    expect(s2).not.toBeNull()
    expect(base.deck.slides.length).toBe(2)

    const s3 = await mergeSlideFromPptx(base, await onePagePptx('PAGE_THREE'))
    expect(s3).not.toBeNull()
    expect(base.deck.slides.length).toBe(3)

    // Reopen after save: all three slides present, in the right order
    const reopened = await openPptx(await savePptx(base))
    expect(reopened.deck.slides.length).toBe(3)
    // pptxgenjs textboxes parse as text-bearing shapes (not pure text); accept both
    const texts = reopened.deck.slides.map((sl) =>
      sl.elements
        .filter((el) => el.type === 'text' || el.type === 'shape')
        .map(
          (el) =>
            (
              el as { text?: { paragraphs: Array<{ runs: Array<{ text: string }> }> } }
            ).text?.paragraphs
              ?.flatMap((pg) => pg.runs.map((r) => r.text))
              .join('') ?? '',
        )
        .join(' '),
    )
    expect(texts[0]).toContain('PAGE_ONE')
    expect(texts[1]).toContain('PAGE_TWO')
    expect(texts[2]).toContain('PAGE_THREE')
  })

  it('merging a slide with an image: media bytes moved in, rIds do not clash', async () => {
    const base = await openPptx(await onePagePptx('COVER', true))
    const before = [...base.archive.entries.keys()].filter((k) => /ppt\/media\//.test(k)).length
    expect(before).toBeGreaterThanOrEqual(1)

    // Merge another slide with an image — both media are originally named image-1-1.png, so remap must avoid overwriting
    await mergeSlideFromPptx(base, await onePagePptx('SECOND', true))
    expect(base.deck.slides.length).toBe(2)

    const reopened = await openPptx(await savePptx(base))
    expect(reopened.deck.slides.length).toBe(2)
    const media = [...reopened.archive.entries.keys()].filter((k) => /ppt\/media\//.test(k))
    // Both images exist independently (neither overwritten)
    expect(media.length).toBeGreaterThanOrEqual(2)
    // The second slide does contain a picture element
    const picCount = reopened.deck.slides[1]!.elements.filter((el) => el.type === 'picture').length
    expect(picCount).toBeGreaterThanOrEqual(1)
  })

  it('saved output opens with a standard zip and has the right number of slide parts', async () => {
    const base = await openPptx(await onePagePptx('A'))
    await mergeSlideFromPptx(base, await onePagePptx('B'))
    const bytes = await savePptx(base)
    const reopened = await openPptx(bytes)
    const slideParts = [...reopened.archive.entries.keys()].filter((k) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(k),
    )
    expect(slideParts.length).toBe(2)
    // presentation.xml's sldId list should also have 2 entries
    const pres = reopened.archive.readText('ppt/presentation.xml') ?? ''
    const sldIds = [...pres.matchAll(/<p:sldId\b/g)].length
    expect(sldIds).toBe(2)
  })

  it('merging a slide with a chart: part moved to a fresh name, rels and Content_Types stay valid', async () => {
    async function onePageWithChart(label: string, last: number): Promise<Uint8Array> {
      const p = new PptxGenJS()
      p.defineLayout({ name: 'W', width: 13.333, height: 7.5 })
      p.layout = 'W'
      const s = p.addSlide()
      s.addText(label, { x: 1, y: 1, w: 8, h: 1, fontSize: 32 })
      s.addChart(p.ChartType.bar, [{ name: 'S', labels: ['A', 'B', 'C'], values: [1, 2, last] }], {
        x: 1,
        y: 3,
        w: 6,
        h: 3,
      })
      const buf = (await p.write({ outputType: 'nodebuffer' })) as Buffer
      return new Uint8Array(buf)
    }

    // Base deck already owns ppt/charts/chart1.xml; the merged slide's chart
    // must land as chart2.xml (not overwrite), with its rel remapped.
    const base = await openPptx(await onePageWithChart('PAGE_ONE', 3))
    await mergeSlideFromPptx(base, await onePageWithChart('PAGE_TWO', 7))
    expect(base.deck.slides.length).toBe(2)

    const chartParts = [...base.archive.entries.keys()].filter((k) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(k),
    ).sort()
    expect(chartParts).toEqual(['ppt/charts/chart1.xml', 'ppt/charts/chart2.xml'])

    // The merged slide's rels point at chart2.xml and the Content_Types Override exists
    const rels = base.archive.readText('ppt/slides/_rels/slide2.xml.rels') ?? ''
    expect(rels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"')
    expect(rels).toContain('Target="../charts/chart2.xml"')
    const slide2 = base.archive.readText('ppt/slides/slide2.xml') ?? ''
    const rid = rels.match(/Id="(rId\d+)"[^>]*Target="\.\.\/charts\/chart2\.xml"/)?.[1]
    expect(rid).toBeTruthy()
    expect(slide2).toContain(`r:id="${rid}"`)
    const ct = base.archive.readText('[Content_Types].xml') ?? ''
    expect(ct).toContain('PartName="/ppt/charts/chart2.xml"')

    // Reopen after save: both slides keep a parsed chart element
    const reopenedChart = await openPptx(await savePptx(base))
    expect(reopenedChart.deck.slides.length).toBe(2)
    const chartCounts = reopenedChart.deck.slides.map(
      (sl) => sl.elements.filter((el) => el.type === 'chart').length,
    )
    expect(chartCounts[0]).toBe(1)
    expect(chartCounts[1]).toBe(1)
  })

  it('merging a slide with a jpeg: the new Default lands inside the <Types> root, not outside it', async () => {
    // Regression: ensureDefaultContentType once inserted the Default right after the
    // XML declaration (first '>' in the file), producing malformed XML that
    // LibreOffice/PowerPoint reject outright ("source file could not be loaded").
    async function onePageWithJpeg(label: string): Promise<Uint8Array> {
      const p = new PptxGenJS()
      p.defineLayout({ name: 'W', width: 13.333, height: 7.5 })
      p.layout = 'W'
      const s = p.addSlide()
      s.addText(label, { x: 1, y: 1, w: 8, h: 1, fontSize: 32 })
      // bytes are PNG but the data URI mime decides the part extension: .jpeg
      s.addImage({ data: 'image/jpeg;base64,' + RED_DOT, x: 1, y: 3, w: 2, h: 2 })
      const buf = (await p.write({ outputType: 'nodebuffer' })) as Buffer
      return new Uint8Array(buf)
    }

    // Target deck stripped to a bare Content_Types (like the engine's blank template:
    // no image Defaults) so the merge has to add the jpeg Default itself
    const base = await openPptx(await onePagePptx('COVER'))
    const bareCt = (base.archive.readText('[Content_Types].xml') ?? '').replace(
      /<Default Extension="(?:jpeg|jpg|png|gif|svg)"/g,
      '<Default Extension="x-$1"',
    )
    base.archive.entries.set('[Content_Types].xml', Buffer.from(bareCt, 'utf8'))
    expect(bareCt).not.toContain('Extension="jpeg"')

    await mergeSlideFromPptx(base, await onePageWithJpeg('WITH_JPEG'))
    const ct = base.archive.readText('[Content_Types].xml') ?? ''
    const typesOpen = ct.search(/<Types\b/)
    const jpegDefault = ct.indexOf('<Default Extension="jpeg"')
    expect(typesOpen).toBeGreaterThan(-1)
    // Inside the root element: after <Types …> opens, before </Types> closes
    expect(jpegDefault).toBeGreaterThan(typesOpen)
    expect(ct.lastIndexOf('</Types>')).toBeGreaterThan(jpegDefault)
    // The XML declaration must be followed directly by the root element
    expect(ct.replace(/^<\?xml[^?]*\?>/, '')).toMatch(/^\s*<Types\b/)

    const reopened = await openPptx(await savePptx(base))
    const picCount = reopened.deck.slides[1]!.elements.filter((el) => el.type === 'picture').length
    expect(picCount).toBeGreaterThanOrEqual(1)
  })
})
