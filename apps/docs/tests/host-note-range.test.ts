import { describe, expect, it } from 'vitest'
import type { Editor } from '@tiptap/react'
import { Schema } from 'prosemirror-model'

import { blockRangePositions, hostNoteRange } from '../src/renderer/ai/protocol'

/**
 * 宿主批注锚点按摘要收窄（2026-11-10「只选了一句却高亮整段」回归）：
 * 宿主队列只带块索引 + 原文摘要，直接按块画锚点会把整段都标上。
 */
const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { content: 'inline*', group: 'block' },
    text: { group: 'inline' },
  },
  marks: {},
})

function editorOf(...texts: string[]): Editor {
  const doc = schema.node('doc', null, texts.map((text) => schema.node('paragraph', null, text ? [schema.text(text)] : [])))
  return { state: { doc } } as unknown as Editor
}

describe('hostNoteRange', () => {
  it('单块选区按摘要收窄到那句话（不是整段）', () => {
    const editor = editorOf('第一句。第二句。')
    const whole = blockRangePositions(editor, 0, 0)
    expect(whole).toEqual({ from: 0, to: 10 })
    // "第二句。" 在块内是第 5–8 个字（PM 位置 5..9，end 排他）
    expect(hostNoteRange(editor, 0, 0, '第二句。')).toEqual({ from: 5, to: 9 })
  })

  it('摘要匹配不唯一时回退整块，不猜位置', () => {
    const editor = editorOf('重复的一句。重复的一句。')
    expect(hostNoteRange(editor, 0, 0, '重复的一句。')).toEqual(blockRangePositions(editor, 0, 0))
  })

  it('摘要找不到（文档已改过）时回退整块', () => {
    const editor = editorOf('原来的内容')
    expect(hostNoteRange(editor, 0, 0, '已经被删掉的一句话')).toEqual(blockRangePositions(editor, 0, 0))
  })

  it('多块选区不收窄（摘要可能跨块，无法安全映射）', () => {
    const editor = editorOf('第一段。', '第二段。')
    expect(hostNoteRange(editor, 0, 1, '第一段。第二段。')).toEqual(blockRangePositions(editor, 0, 1))
  })

  it('缺摘要 / 摘要过短时回退整块', () => {
    const editor = editorOf('第一句。第二句。')
    const whole = blockRangePositions(editor, 0, 0)
    expect(hostNoteRange(editor, 0, 0, undefined)).toEqual(whole)
    expect(hostNoteRange(editor, 0, 0, '嗯')).toEqual(whole)
  })

  it('空白归一化后仍能命中（换行/多空格按单空格比对）', () => {
    const editor = editorOf('第一句。\n第二句。')
    // 块内文本是「第一句。\n第二句。」，摘要里的换行被归一化成空格；命中后
    // 收窄到整段文字（首尾各留一个块边界位置）。
    expect(hostNoteRange(editor, 0, 0, '第一句。  第二句。')).toEqual({ from: 1, to: 10 })
  })
})