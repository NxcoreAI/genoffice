import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import { AskQueueAnchors, HostNotePopover } from '../src/renderer/HostNotes'
import type { PdfAskQueueItem } from '../src/shared/ipc'

/** [x1,yMax,x2,yMax,x1,yMin,x2,yMin] — a 100×20 unit box on the page. */
const quad = [10, 720, 110, 720, 10, 700, 110, 700]

const item = (overrides: Partial<PdfAskQueueItem> = {}): PdfAskQueueItem => ({
  id: 'draft-1',
  instruction: 'tighten this',
  excerpt: 'the selected passage',
  pageIndex: 0,
  quads: [quad],
  ...overrides,
})

const noop = (): void => {}

function renderAnchors(
  entries: Array<{ item: PdfAskQueueItem; index: number }>,
  geom: { pw: number; ph: number; rot: number },
  scale = 1,
): string {
  return renderToStaticMarkup(
    createElement(AskQueueAnchors, {
      entries,
      geom,
      scale,
      editingId: null,
      onHover: noop,
      onLeave: noop,
      onOpen: noop,
    }),
  )
}

describe('AskQueueAnchors', () => {
  it('anchors every quad in display space and numbers the badge from the queue order', () => {
    const markup = renderAnchors(
      [{ item: item({ quads: [quad, [200, 500, 300, 500, 200, 480, 300, 480]] }), index: 2 }],
      { pw: 600, ph: 800, rot: 0 },
      2,
    )
    expect(markup.match(/data-pdf-ask-id="draft-1"/g)).toHaveLength(2)
    // pdf y-up → display y-down: y 700..720 on an 800pt page = 80..100 * scale 2
    expect(markup).toContain('left:20px;top:160px;width:200px;height:40px')
    // badge uses the queue index, not the per-page position
    expect(markup).toContain('>3</button>')
  })

  it('converts through page rotation (quads stay rotation-stable)', () => {
    const markup = renderAnchors([{ item: item(), index: 0 }], { pw: 600, ph: 800, rot: 90 })
    expect(markup).toContain('left:700px;top:10px;width:20px;height:100px')
  })

  it('renders nothing for a page without queued anchors', () => {
    const markup = renderAnchors([], { pw: 600, ph: 800, rot: 0 })
    expect(markup).not.toContain('data-pdf-ask-id')
  })
})

function renderPopover(mode: 'new' | 'edit', initialText?: string): string {
  return renderToStaticMarkup(
    createElement(
      LocaleProvider,
      { initial: 'zh' },
      createElement(HostNotePopover, {
        rect: { left: 100, top: 100, right: 300, bottom: 120 },
        mode,
        initialText,
        onQueue: noop,
        onUpdate: noop,
        onDelete: noop,
        onClose: noop,
      }),
    ),
  )
}

describe('HostNotePopover', () => {
  it('opens as a single-input queue popover for a new note', () => {
    const markup = renderPopover('new')
    expect(markup).toContain('写注释，回车加入队列…')
    expect(markup).not.toContain('host-note-pop-delete')
  })

  it('prefills and offers delete when editing a queued draft', () => {
    const markup = renderPopover('edit', 'tighten this')
    expect(markup).toContain('value="tighten this"')
    expect(markup).toContain('修改注释…')
    expect(markup).toContain('host-note-pop-delete')
  })
})
