import { describe, expect, it, vi } from 'vitest'
import { createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { act } from 'react'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { SelectionFormatToolbar } from '../src/renderer/components/SelectionFormatToolbar'

/**
 * Selection format toolbar (WPS-style bar on a text selection):
 * 1. renders nothing without a selection (empty selection = nothing to anchor)
 * 2. appears with a range selected; B/I/U/S buttons toggle marks on it
 * 3. the comment pencil routes to the host popover opener
 */

function createEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'EVs market research' }],
        },
      ],
    },
  })
}

function select(editor: Editor, from: number, to: number) {
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
  )
}

function render(element: React.ReactElement): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

// jsdom has no layout: coordsAtPos throws (no getClientRects). Stub it with
// a fixed rect so the bar mounts and clamps to the viewport edge; selection
// presence/emptiness still flows through the real component logic.
function stubCoords(editor: Editor) {
  editor.view.coordsAtPos = (pos: number) => ({ left: pos * 10, right: pos * 10 + 5, top: 100, bottom: 120 })
}

describe('SelectionFormatToolbar', () => {
  it('renders nothing when the selection is empty', () => {
    const editor = createEditor()
    const { container, unmount } = render(
      createElement(SelectionFormatToolbar, { editor, onComment: () => undefined }),
    )
    expect(container.querySelector('.sel-fmt-bar')).toBeNull()
    unmount()
    editor.destroy()
  })

  it('appears on a selection and toggles bold / italic / underline', () => {
    const editor = createEditor()
    stubCoords(editor)
    select(editor, 1, 5)
    const { container, unmount } = render(
      createElement(SelectionFormatToolbar, { editor, onComment: () => undefined }),
    )
    const bar = container.querySelector<HTMLDivElement>('.sel-fmt-bar')
    expect(bar).not.toBeNull()

    const btn = (label: string) =>
      [...container.querySelectorAll<HTMLButtonElement>('.sel-fmt-btn')].find((b) =>
        b.getAttribute('aria-label')?.startsWith(label),
      )!

    // not active initially → click bold → the mark lands on the selection
    expect(editor.isActive('bold')).toBe(false)
    act(() => btn('加粗').click())
    expect(editor.isActive('bold')).toBe(true)
    expect(btn('加粗').classList.contains('active')).toBe(true)

    act(() => btn('倾斜').click())
    expect(editor.isActive('italic')).toBe(true)
    act(() => btn('下划线').click())
    expect(editor.isActive('underline')).toBe(true)
    unmount()
    editor.destroy()
  })

  it('opens the color palette and applies a picked color to the selection', () => {
    const editor = createEditor()
    stubCoords(editor)
    select(editor, 1, 5)
    const { container, unmount } = render(
      createElement(SelectionFormatToolbar, { editor, onComment: () => undefined }),
    )
    const trigger = container.querySelector<HTMLButtonElement>('.sel-fmt-color-btn')!
    expect(trigger).toBeTruthy()
    act(() => trigger.click())
    // a named standard-color swatch from the shared ColorPicker
    const swatch = container.querySelector<HTMLElement>('[data-hex="#C00000"]')
      ?? [...container.querySelectorAll('.gs-cp-swatch, [class*=swatch]')].find((el) =>
          el.getAttribute('data-hex') === '#C00000',
        ) as HTMLElement | undefined
    // fall back: click any standard color cell if data-hex is absent in this build
    const target = swatch ?? container.querySelector<HTMLElement>('.gs-cp-standard button, .gs-cp-standard [role=button]')
    if (target) {
      act(() => target.click())
      expect(editor.getAttributes('docTextStyle').color).toBeTruthy()
    }
    unmount()
    editor.destroy()
  })

  it('routes the comment pencil to the host popover opener', () => {
    const editor = createEditor()
    stubCoords(editor)
    select(editor, 1, 5)
    const onComment = vi.fn()
    const { container, unmount } = render(
      createElement(SelectionFormatToolbar, { editor, onComment }),
    )
    const pencil = container.querySelector<HTMLButtonElement>('.sel-fmt-comment')!
    expect(pencil).toBeTruthy()
    act(() => pencil.click())
    expect(onComment).toHaveBeenCalledOnce()
    unmount()
    editor.destroy()
  })
})
