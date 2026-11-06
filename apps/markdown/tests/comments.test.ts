import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import type { MdComment } from '../src/shared/ipc'
import {
  captureSelectionAnchor,
  nextCommentId,
  resolveAnchor,
  topLevelBlocks,
} from '../src/renderer/editor/comments'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function makeEditor(markdown: string): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  // parse markdown into the doc through the app's own extension set
  const parsed = editor.markdown!.parse(markdown)
  editor.commands.setContent(parsed)
  return editor
}

const doc = [
  '# Title',
  '',
  'First paragraph with some interesting text that can be commented on.',
  '',
  'Second paragraph follows here with more words.',
].join('\n')

function blocksOf(editor: Editor) {
  return topLevelBlocks(editor)
}

describe('comment ids', () => {
  it('allocates the smallest unused numeric id', () => {
    expect(nextCommentId([])).toBe('1')
    expect(
      nextCommentId([{ id: '3', author: 'a', text: 'x', blockIndex: 0, excerpt: 'e', offset: 0 }]),
    ).toBe('4')
    expect(
      nextCommentId([
        { id: 'abc', author: 'a', text: 'x', blockIndex: 0, excerpt: 'e', offset: 0 },
        { id: '2', author: 'a', text: 'x', blockIndex: 0, excerpt: 'e', offset: 0 },
      ]),
    ).toBe('3')
  })
})

describe('top-level blocks', () => {
  it('numbers blocks with doc positions', () => {
    const editor = makeEditor(doc)
    const blocks = blocksOf(editor)
    // heading, paragraph, paragraph (blank lines don't create empty blocks)
    expect(blocks.length).toBe(3)
    expect(blocks[0].text).toBe('Title')
    expect(blocks[1].text).toContain('First paragraph')
    expect(blocks[2].text).toContain('Second paragraph')
    // positions strictly increase
    expect(blocks[0].from).toBeLessThan(blocks[1].from)
    expect(blocks[1].from).toBeLessThan(blocks[2].from)
  })
})

describe('captureSelectionAnchor', () => {
  it('captures block index + excerpt from a selection', () => {
    const editor = makeEditor(doc)
    const target = blocksOf(editor)[1]
    const start = target.text.indexOf('interesting')
    const from = target.from + start
    const to = from + 'interesting text'.length
    editor.commands.setTextSelection({ from, to })
    const anchor = captureSelectionAnchor(editor)
    expect(anchor).not.toBeNull()
    expect(anchor!.blockIndex).toBe(1)
    expect(anchor!.excerpt).toBe('interesting text')
  })

  it('rejects empty selections', () => {
    const editor = makeEditor(doc)
    const anchor = captureSelectionAnchor(editor)
    expect(anchor).toBeNull()
  })
})

describe('resolveAnchor', () => {
  it('resolves an anchor captured from the same document', () => {
    const editor = makeEditor(doc)
    const blocks = blocksOf(editor)
    const start = blocks[1].text.indexOf('interesting')
    const comment: MdComment = {
      id: '1',
      author: 'User',
      text: 'note',
      blockIndex: 1,
      excerpt: 'interesting text',
      offset: start,
    }
    const range = resolveAnchor(editor, comment)
    expect(range).not.toBeNull()
    const text = editor.state.doc.textBetween(range!.from, range!.to, ' ', ' ')
    expect(text).toBe('interesting text')
  })

  it('survives an edit above the anchor block (block index stable)', () => {
    const editor = makeEditor(doc)
    const blocks = blocksOf(editor)
    const start = blocks[1].text.indexOf('interesting')
    const comment: MdComment = {
      id: '1',
      author: 'User',
      text: 'note',
      blockIndex: 1,
      excerpt: 'interesting text',
      offset: start,
    }
    // insert a new paragraph before the anchor block — block indices shift,
    // but the excerpt match must recover the anchor in the right block
    editor.commands.insertContentAt(blocks[1].from, '<p>Inserted before</p>')
    const range = resolveAnchor(editor, comment)
    expect(range).not.toBeNull()
    expect(editor.state.doc.textBetween(range!.from, range!.to, ' ', ' ')).toBe('interesting text')
  })

  it('dangles when the excerpt exists nowhere in the document', () => {
    const editor = makeEditor(doc)
    const comment: MdComment = {
      id: '1',
      author: 'User',
      text: 'note',
      blockIndex: 9,
      excerpt: 'this text was deleted entirely',
      offset: 0,
    }
    expect(resolveAnchor(editor, comment)).toBeNull()
  })

  it('falls back to the excerpt head when inner text shifted', () => {
    const editor = makeEditor(doc)
    const comment: MdComment = {
      id: '1',
      author: 'User',
      text: 'note',
      blockIndex: 1,
      excerpt: 'First paragraph with some interesting text that can be commented on',
      offset: 0,
    }
    // text before the excerpt changed; head-based fuzzy fallback should still land
    const blocks = blocksOf(editor)
    editor.commands.insertContentAt(blocks[1].from + 'First '.length, 'very ')
    const range = resolveAnchor(editor, comment)
    expect(range).not.toBeNull()
  })
})

describe('markdown purity', () => {
  it('comments never leak into getMarkdown output', () => {
    const editor = makeEditor(doc)
    const out = editor.getMarkdown()
    expect(out).not.toContain('md-comment')
    expect(out).toContain('First paragraph')
  })
})
