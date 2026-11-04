/**
 * Normalization for the EverRoom embed annotation channel: payloads cross IPC
 * in both directions (renderer ask → host agent, host queue snapshot → editor)
 * and must arrive bounded, with unusable rows dropped instead of rendered.
 */
import { describe, expect, it } from 'vitest'
import {
  ASK_EXCERPT_MAX,
  ASK_INSTRUCTION_MAX,
  ASK_QUEUE_MAX,
  normalizeAgentAskOp,
  normalizeAskQueueAction,
  normalizeAskQueueItems,
} from '../src/main/ask-queue'

describe('normalizeAgentAskOp', () => {
  it('keeps instruction, page, targets, excerpt and queue mode', () => {
    const op = normalizeAgentAskOp({
      instruction: '  move this   down ',
      slideIndex: 3,
      targets: [{ id: 'dur-title', desc: { type: 'shape', text: 'Quarterly  review' } }],
      excerpt: '  Quarterly   review ',
      mode: 'queue',
    })
    expect(op).toMatchObject({
      instruction: 'move this down',
      slideIndex: 3,
      targets: [{ id: 'dur-title', desc: { type: 'shape', text: 'Quarterly review' } }],
      excerpt: 'Quarterly review',
      mode: 'queue',
    })
  })

  it('drops a now-mode flag and missing excerpt (defaults are the normal path)', () => {
    const op = normalizeAgentAskOp({
      instruction: 'make it bigger',
      slideIndex: 0,
      targets: [{ id: 'x', desc: { type: 'picture' } }],
      mode: 'now',
    })
    expect(op).not.toBeNull()
    expect(op?.mode).toBeUndefined()
    expect(op?.excerpt).toBeUndefined()
  })

  it('caps the excerpt and the instruction', () => {
    const op = normalizeAgentAskOp({
      instruction: 'a'.repeat(ASK_INSTRUCTION_MAX + 50),
      slideIndex: 1,
      targets: [{ id: 'x', desc: { type: 'text' } }],
      excerpt: ` ${'word '.repeat(ASK_EXCERPT_MAX)}`,
    })
    expect(op?.instruction).toHaveLength(ASK_INSTRUCTION_MAX)
    expect(op?.excerpt).toHaveLength(ASK_EXCERPT_MAX)
  })

  it('rejects payloads without instruction / targets / a real page index', () => {
    expect(normalizeAgentAskOp(null)).toBeNull()
    expect(
      normalizeAgentAskOp({ instruction: '   ', slideIndex: 0, targets: [{ id: 'x' }] }),
    ).toBeNull()
    expect(normalizeAgentAskOp({ instruction: 'do it', slideIndex: 0, targets: [] })).toBeNull()
    expect(normalizeAgentAskOp({ instruction: 'do it', targets: [{ id: 'x' }] })).toBeNull()
    expect(
      normalizeAgentAskOp({ instruction: 'do it', slideIndex: -1, targets: [{ id: 'x' }] }),
    ).toBeNull()
    // targets without an id are unusable
    expect(
      normalizeAgentAskOp({ instruction: 'do it', slideIndex: 0, targets: [{ desc: {} }] }),
    ).toBeNull()
  })

  it('bounds the target list and keeps table dimensions', () => {
    const targets = Array.from({ length: 30 }, (_, index) => ({
      id: `id-${index}`,
      desc: { type: 'table', rows: 2, cols: 4 },
    }))
    const op = normalizeAgentAskOp({ instruction: 'x', slideIndex: 0, targets })
    expect(op?.targets).toHaveLength(20)
    expect(op?.targets[0]?.desc).toMatchObject({ type: 'table', rows: 2, cols: 4 })
  })
})

describe('normalizeAskQueueItems', () => {
  it('accepts the host flat shape (optional page hint + element ids)', () => {
    const items = normalizeAskQueueItems([
      {
        id: 'd1',
        instruction: ' shrink it ',
        excerpt: ' Title ',
        slideIndex: 2,
        elementIds: ['a', 'b'],
      },
      { id: 'd2', instruction: 'recolor', elementIds: ['c'] },
    ])
    expect(items).toEqual([
      {
        id: 'd1',
        instruction: 'shrink it',
        excerpt: 'Title',
        slideIndex: 2,
        elementIds: ['a', 'b'],
      },
      { id: 'd2', instruction: 'recolor', slideIndex: 0, elementIds: ['c'] },
    ])
  })

  it('drops rows without an anchor, a note, or an id', () => {
    const items = normalizeAskQueueItems([
      { id: 'd1', instruction: 'note' }, // no element ids
      { id: 'd2', instruction: '   ', elementIds: ['a'] }, // empty note
      { instruction: 'note', elementIds: ['a'] }, // no id
      { id: 'd3', instruction: 'kept', elementIds: ['a'] },
    ])
    expect(items.map((item) => item.id)).toEqual(['d3'])
  })

  it('dedupes ids and anchors, and bounds the snapshot', () => {
    const items = normalizeAskQueueItems([
      { id: 'd1', instruction: 'one', elementIds: ['a', 'a', ' b '] },
      { id: 'd1', instruction: 'duplicate id', elementIds: ['b'] },
      { id: 'd2', instruction: 'two', elementIds: [] },
    ])
    expect(items).toHaveLength(1)
    expect(items[0]?.elementIds).toEqual(['a', 'b'])
  })

  it('caps the snapshot length and rejects non-arrays', () => {
    const rows = Array.from({ length: ASK_QUEUE_MAX + 10 }, (_, index) => ({
      id: `d${index}`,
      instruction: 'note',
      elementIds: ['a'],
    }))
    expect(normalizeAskQueueItems(rows)).toHaveLength(ASK_QUEUE_MAX)
    expect(normalizeAskQueueItems(null)).toEqual([])
  })
})

describe('normalizeAskQueueAction', () => {
  it('passes send-all through', () => {
    expect(normalizeAskQueueAction({ kind: 'send-all' })).toEqual({ kind: 'send-all' })
  })

  it('normalizes remove and update, trimming the draft id', () => {
    expect(normalizeAskQueueAction({ kind: 'remove', draftId: ' d1 ' })).toEqual({
      kind: 'remove',
      draftId: 'd1',
    })
    expect(
      normalizeAskQueueAction({ kind: 'update', draftId: ' d1 ', instruction: ' tighten   copy ' }),
    ).toEqual({ kind: 'update', draftId: 'd1', instruction: 'tighten copy' })
  })

  it('rejects unknown kinds and empty payloads', () => {
    expect(normalizeAskQueueAction({ kind: 'nope' })).toBeNull()
    expect(normalizeAskQueueAction({ kind: 'remove' })).toBeNull()
    expect(normalizeAskQueueAction({ kind: 'update', draftId: 'd1', instruction: '  ' })).toBeNull()
    expect(normalizeAskQueueAction(null)).toBeNull()
  })
})
