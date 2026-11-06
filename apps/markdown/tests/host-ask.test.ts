import { describe, expect, it } from 'vitest'
import {
  MAX_DRAFT_ID_CHARS,
  MAX_EXCERPT_CHARS,
  MAX_INSTRUCTION_CHARS,
  MAX_QUEUE_ITEMS,
  normalizeAgentAskOp,
  normalizeAskQueueAction,
  normalizeAskQueueItems,
} from '../src/main/host-ask'
import { MARKDOWN_CHANNELS } from '../src/shared/ipc'

const validOp = () => ({
  instruction: '  tighten this paragraph  ',
  blockIndex: 2,
  excerpt: '  the selected\n passage  ',
  offset: 14,
  mode: 'queue',
})

describe('host annotation IPC channels', () => {
  it('keeps the channel names the host wires against', () => {
    expect(MARKDOWN_CHANNELS.agentAsk).toBe('markdown:agent-ask')
    expect(MARKDOWN_CHANNELS.askQueue).toBe('markdown:ask-queue')
    expect(MARKDOWN_CHANNELS.askQueueAction).toBe('markdown:ask-queue-action')
  })
})

describe('normalizeAgentAskOp', () => {
  it('keeps a valid payload, trims text and collapses the excerpt', () => {
    expect(normalizeAgentAskOp(validOp())).toEqual({
      instruction: 'tighten this paragraph',
      blockIndex: 2,
      excerpt: 'the selected passage',
      offset: 14,
      mode: 'queue',
    })
  })

  it('omits the optional excerpt/offset/mode when absent or zero', () => {
    const op = normalizeAgentAskOp({ instruction: 'note', blockIndex: 0 })
    expect(op).toEqual({ instruction: 'note', blockIndex: 0 })
    expect(op && 'excerpt' in op).toBe(false)
    expect(op && 'offset' in op).toBe(false)
    expect(op && 'mode' in op).toBe(false)
  })

  it('only forwards the queue mode (now = default, no field)', () => {
    const op = normalizeAgentAskOp({ instruction: 'note', blockIndex: 0, mode: 'now' })
    expect(op && 'mode' in op).toBe(false)
  })

  it('drops payloads without an instruction or a usable block index', () => {
    expect(normalizeAgentAskOp(null)).toBeNull()
    expect(normalizeAgentAskOp({ blockIndex: 0 })).toBeNull()
    expect(normalizeAgentAskOp({ instruction: '  ', blockIndex: 0 })).toBeNull()
    expect(normalizeAgentAskOp({ instruction: 'note', blockIndex: -1 })).toBeNull()
    expect(normalizeAgentAskOp({ instruction: 'note', blockIndex: 1.5 })).toEqual({
      instruction: 'note',
      blockIndex: 1,
    })
    expect(normalizeAgentAskOp({ instruction: 'note', blockIndex: 0, offset: -3 })).toBeNull()
  })

  it('bounds the instruction and excerpt lengths', () => {
    const op = normalizeAgentAskOp({
      instruction: `x`.repeat(MAX_INSTRUCTION_CHARS + 50),
      blockIndex: 0,
      excerpt: `y`.repeat(MAX_EXCERPT_CHARS + 50),
    })
    expect(op?.instruction.length).toBe(MAX_INSTRUCTION_CHARS)
    expect(op?.excerpt?.length).toBe(MAX_EXCERPT_CHARS)
  })
})

describe('normalizeAskQueueItems', () => {
  it('keeps well-formed rows and drops malformed or duplicate ones', () => {
    const rows = normalizeAskQueueItems([
      { id: 'a', instruction: 'note a', excerpt: 'text a', blockIndex: 1, offset: 3 },
      { id: 'a', instruction: 'duplicate id', blockIndex: 1 },
      { id: '', instruction: 'no id', blockIndex: 1 },
      { id: 'b', instruction: 'no block', offset: 3 },
      { id: 'c', instruction: 'bad block', blockIndex: -1 },
      { id: 'd', blockIndex: 1 },
      'junk',
      null,
    ])
    expect(rows).toEqual([
      { id: 'a', instruction: 'note a', excerpt: 'text a', blockIndex: 1, offset: 3 },
    ])
  })

  it('caps the snapshot size', () => {
    const raw = Array.from({ length: MAX_QUEUE_ITEMS + 20 }, (_, i) => ({
      id: `id${i}`,
      instruction: `note ${i}`,
      blockIndex: i,
    }))
    expect(normalizeAskQueueItems(raw)).toHaveLength(MAX_QUEUE_ITEMS)
    expect(normalizeAskQueueItems('nope')).toEqual([])
  })
})

describe('normalizeAskQueueAction', () => {
  it('passes the three well-formed shapes through', () => {
    expect(normalizeAskQueueAction({ kind: 'send-all' })).toEqual({ kind: 'send-all' })
    expect(normalizeAskQueueAction({ kind: 'remove', draftId: ' d1 ' })).toEqual({
      kind: 'remove',
      draftId: 'd1',
    })
    expect(normalizeAskQueueAction({ kind: 'update', draftId: 'd1', instruction: ' x ' })).toEqual({
      kind: 'update',
      draftId: 'd1',
      instruction: 'x',
    })
  })

  it('rejects unknown kinds and missing fields', () => {
    expect(normalizeAskQueueAction(null)).toBeNull()
    expect(normalizeAskQueueAction({})).toBeNull()
    expect(normalizeAskQueueAction({ kind: 'remove' })).toBeNull()
    expect(normalizeAskQueueAction({ kind: 'update', draftId: 'd1' })).toBeNull()
    expect(normalizeAskQueueAction({ kind: 'explode', draftId: 'd1' })).toBeNull()
  })

  it('bounds the draft id and instruction', () => {
    expect(
      normalizeAskQueueAction({ kind: 'remove', draftId: 'd'.repeat(MAX_DRAFT_ID_CHARS + 5) })
        ?.draftId.length,
    ).toBe(MAX_DRAFT_ID_CHARS)
    expect(
      normalizeAskQueueAction({
        kind: 'update',
        draftId: 'd1',
        instruction: 'x'.repeat(MAX_INSTRUCTION_CHARS + 5),
      })?.instruction.length,
    ).toBe(MAX_INSTRUCTION_CHARS)
  })
})
