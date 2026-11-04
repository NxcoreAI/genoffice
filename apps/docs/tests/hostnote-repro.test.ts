/**
 * 白屏复现：宿主推送一条批注（hostAskQueue 变化）后，
 * HostNoteCard / 锚点装饰链路是否抛未捕获异常。
 */
import { describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  addQueueAnchor,
  HOST_NOTE_ANCHOR_CLASS,
} from '../src/renderer/editor/ai-queue-anchors'

const { HostNoteCard } = await import('../src/renderer/components/HostNoteCard')

function createEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', attrs: { docxIndex: 0 }, content: [{ type: 'text', text: '第一段原文内容' }] },
      ],
    },
  })
}

function mount(element: React.ReactElement): { root: Root; cleanup: () => void } {
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  act(() => { root.render(element) })
  return { root, cleanup: () => { act(() => root.unmount()); container.remove() } }
}

describe('HostNoteCard 白屏复现', () => {
  it('一条批注：挂载 + 加锚点 + 悬停解析不抛异常', async () => {
    const editor = createEditor()
    const items = [{
      id: 'draft-1',
      instruction: '改一下',
      excerpt: '第一段原文内容',
      fromBlockIndex: 0,
      toBlockIndex: 0,
      kind: 'text' as const,
    }]
    // App.tsx 的同步 effect：队列变化 → 清锚点 → 逐条加 host-note 锚点
    act(() => { addQueueAnchor(editor, 'draft-1', 1, 6, HOST_NOTE_ANCHOR_CLASS) })
    const mounted = mount(createElement(HostNoteCard, { editor, items, onDelete: () => undefined }))
    // 模拟鼠标移入锚点 span：qidAt 走 posAtCoords → plugin find
    const anchor = document.querySelector(`.${HOST_NOTE_ANCHOR_CLASS}`)
    if (anchor) {
      const el = anchor as HTMLElement
      el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
      await act(async () => { await new Promise((r) => setTimeout(r, 220)) })
    }
    mounted.cleanup()
    editor.destroy()
    expect(true).toBe(true)
  })
})
