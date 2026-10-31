import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  agentImageInsertPos,
  clampImageWidth,
  scaledImageSize,
} from '../src/renderer/editor/agent-image'

/**
 * Pure behaviour of the host-driven image insert (`insert_image_base64`):
 * where the block lands for `afterIndex` and how `maxWidthPx` sizes it. The
 * image decode/insert itself is not covered here — jsdom has no `Image`
 * implementation — so these are exactly the two regressions from the agent
 * test run (dropped block anchor / fixed 620px width).
 */

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const heading = (t: string, level = 1): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level },
  content: [text(t)],
})
const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [text(t)],
})

/** Editors created during the current test; destroyed in afterEach so ProseMirror's
 * DOMObserver timers can't fire after the JSDOM environment is torn down. */
const liveEditors: Editor[] = []

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  liveEditors.push(editor)
  return editor
}

afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

/** 0 h1 | 1 p | 2 h2 | 3 p */
const fixture = () => [
  heading('Chapter 1 Overview', 1),
  para('GenSpark is an AI office suite.'),
  heading('Risk Notes', 2),
  para('This document is for reference only.'),
]

/** end position of top-level block `index`, computed independently of the helper */
function blockEnd(doc: Editor['state']['doc'], index: number): number {
  let pos = 0
  for (let i = 0; i <= index; i++) pos += doc.child(i).nodeSize
  return pos
}

describe('agentImageInsertPos', () => {
  it('places the image right after the block named by afterIndex', () => {
    const editor = createEditor(fixture())
    const doc = editor.state.doc
    expect(doc.childCount).toBe(4)
    for (let index = 0; index < doc.childCount; index++) {
      expect(agentImageInsertPos(doc, index)).toBe(blockEnd(doc, index))
    }
  })

  it('anchors after the requested block, not after the last one', () => {
    const editor = createEditor(fixture())
    const doc = editor.state.doc
    // block 1 is the first paragraph; the image must not follow block 3
    expect(agentImageInsertPos(doc, 1)).toBe(doc.child(0).nodeSize + doc.child(1).nodeSize)
    expect(agentImageInsertPos(doc, 1)).toBeLessThan(agentImageInsertPos(doc, 3))
  })

  it('treats afterIndex -1 as end of document', () => {
    const editor = createEditor(fixture())
    const doc = editor.state.doc
    expect(agentImageInsertPos(doc, -1)).toBe(doc.content.size)
  })

  it('treats out-of-range indexes as end of document', () => {
    const editor = createEditor(fixture())
    const doc = editor.state.doc
    expect(agentImageInsertPos(doc, doc.childCount)).toBe(doc.content.size)
    expect(agentImageInsertPos(doc, doc.childCount + 7)).toBe(doc.content.size)
    expect(agentImageInsertPos(doc, 99)).toBe(doc.content.size)
  })

  it('treats non-integer / NaN indexes as end of document', () => {
    const editor = createEditor(fixture())
    const doc = editor.state.doc
    expect(agentImageInsertPos(doc, 1.5)).toBe(doc.content.size)
    expect(agentImageInsertPos(doc, Number.NaN)).toBe(doc.content.size)
  })

  it('works for a single-block document (both ends coincide)', () => {
    const editor = createEditor([para('only paragraph')])
    const doc = editor.state.doc
    expect(agentImageInsertPos(doc, 0)).toBe(doc.content.size)
    expect(agentImageInsertPos(doc, -1)).toBe(doc.content.size)
  })
})

describe('scaledImageSize', () => {
  it('scales a wide image down to the cap, keeping the aspect ratio', () => {
    expect(scaledImageSize({ width: 1000, height: 500 }, 480)).toEqual({ width: 480, height: 240 })
    expect(scaledImageSize({ width: 1000, height: 333 }, 480)).toEqual({ width: 480, height: 160 })
  })

  it('never upscales an image narrower than the cap', () => {
    expect(scaledImageSize({ width: 320, height: 200 }, 620)).toEqual({ width: 320, height: 200 })
    expect(scaledImageSize({ width: 16, height: 16 }, 620)).toEqual({ width: 16, height: 16 })
  })

  it('keeps an image that exactly matches the cap untouched', () => {
    expect(scaledImageSize({ width: 620, height: 310 }, 620)).toEqual({ width: 620, height: 310 })
  })

  it('uses the renderer hard cap (620) when the tool allows more', () => {
    expect(scaledImageSize({ width: 1240, height: 620 }, 620)).toEqual({ width: 620, height: 310 })
  })
})

describe('clampImageWidth', () => {
  it('keeps a requested cap below the renderer limit', () => {
    expect(clampImageWidth(480, 620)).toBe(480)
    expect(clampImageWidth(16, 620)).toBe(16)
    expect(clampImageWidth(620, 620)).toBe(620)
  })

  it('touches the renderer hard cap when the request is larger', () => {
    expect(clampImageWidth(1200, 620)).toBe(620)
    expect(clampImageWidth(2000, 620)).toBe(620)
  })

  it('touches the schema maximum (1200) when the renderer allows more', () => {
    expect(clampImageWidth(5000, 5000)).toBe(1200)
  })

  it('touches the schema minimum (16) for tiny requests', () => {
    expect(clampImageWidth(4, 620)).toBe(16)
    expect(clampImageWidth(0, 620)).toBe(16)
    expect(clampImageWidth(-40, 620)).toBe(16)
  })

  it('falls back to the cap for non-finite input', () => {
    expect(clampImageWidth(Number.NaN, 620)).toBe(620)
    expect(clampImageWidth(Number.POSITIVE_INFINITY, 620)).toBe(620)
  })

  it('feeds scaledImageSize so the inserted width never exceeds schema or renderer limits', () => {
    const natural = { width: 2400, height: 1200 }
    const rendererLimited = scaledImageSize(natural, clampImageWidth(2000, 620))
    expect(rendererLimited).toEqual({ width: 620, height: 310 })
    const schemaLimited = scaledImageSize(natural, clampImageWidth(5000, 5000))
    expect(schemaLimited).toEqual({ width: 1200, height: 600 })
  })
})
