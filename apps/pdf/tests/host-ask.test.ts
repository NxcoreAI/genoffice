import { describe, expect, it } from 'vitest'
import {
  MAX_DRAFT_ID_CHARS,
  MAX_EXCERPT_CHARS,
  MAX_INSTRUCTION_CHARS,
  MAX_QUADS,
  MAX_QUEUE_ITEMS,
  normalizeAgentAskOp,
  normalizeAskQueueAction,
  normalizeAskQueueItems,
} from '../src/main/host-ask'
import { PDF_CHANNELS } from '../src/shared/ipc'

/** One well-formed PDF-space quad: [x1,yMax,x2,yMax,x1,yMin,x2,yMin]. */
const quad = (dx = 0): number[] => [10 + dx, 720, 110 + dx, 720, 10 + dx, 700, 110 + dx, 700]

const validOp = () => ({
  instruction: '  tighten this paragraph  ',
  pageIndex: 2,
  quads: [quad()],
  excerpt: '  the selected\n passage  ',
  mode: 'queue',
})

describe('host annotation IPC channels', () => {
  it('keeps the channel names the host wires against', () => {
    expect(PDF_CHANNELS.agentAsk).toBe('pdf:agent-ask')
    expect(PDF_CHANNELS.askQueue).toBe('pdf:ask-queue')
    expect(PDF_CHANNELS.askQueueAction).toBe('pdf:ask-queue-action')
  })
})

describe('normalizeAgentAskOp', () => {
  it('keeps a valid payload, trims text and collapses the excerpt', () => {
    expect(normalizeAgentAskOp(validOp())).toEqual({
      instruction: 'tighten this paragraph',
      pageIndex: 2,
      quads: [quad()],
      excerpt: 'the selected passage',
      mode: 'queue',
    })
  })

  it('omits the optional excerpt and mode when absent', () => {
    const op = normalizeAgentAskOp({ instruction: 'note', pageIndex: 0, quads: [quad()] })
    expect(op).toEqual({ instruction: 'note', pageIndex: 0, quads: [quad()] })
    expect(op && 'excerpt' in op).toBe(false)
    expect(op && 'mode' in op).toBe(false)
  })

  it('only forwards the queue mode (now = default, no field)', () => {
    const op = normalizeAgentAskOp({
      instruction: 'note',
      pageIndex: 0,
      quads: [quad()],
      mode: 'now',
    })
    expect(op && 'mode' in op).toBe(false)
  })

  it('rejects payloads without an instruction', () => {
    expect(normalizeAgentAskOp(null)).toBeNull()
    expect(normalizeAgentAskOp({})).toBeNull()
    expect(normalizeAgentAskOp({ ...validOp(), instruction: '   ' })).toBeNull()
    expect(normalizeAgentAskOp({ ...validOp(), instruction: 42 })).toBeNull()
  })

  it('rejects payloads without a usable page anchor', () => {
    expect(normalizeAgentAskOp({ ...validOp(), pageIndex: undefined })).toBeNull()
    expect(normalizeAgentAskOp({ ...validOp(), pageIndex: -1 })).toBeNull()
    expect(normalizeAgentAskOp({ ...validOp(), pageIndex: '2' })).toBeNull()
    expect(normalizeAgentAskOp({ ...validOp(), pageIndex: Number.NaN })).toBeNull()
    expect(normalizeAgentAskOp({ ...validOp(), quads: undefined })).toBeNull()
    expect(normalizeAgentAskOp({ ...validOp(), quads: [] })).toBeNull()
    expect(normalizeAgentAskOp({ ...validOp(), quads: '10,20' })).toBeNull()
  })

  it('drops malformed quads instead of forwarding them', () => {
    const op = normalizeAgentAskOp({
      ...validOp(),
      quads: [quad(), [1, 2, 3], [1, 2, 3, 4, 5, 6, 7, 'x'], null, quad(5)],
    })
    expect(op?.quads).toEqual([quad(), quad(5)])
    expect(normalizeAgentAskOp({ ...validOp(), quads: [[1, 2, 3]] })).toBeNull()
  })

  it('bounds instruction, excerpt and quad count', () => {
    const op = normalizeAgentAskOp({
      ...validOp(),
      instruction: 'x'.repeat(MAX_INSTRUCTION_CHARS + 50),
      excerpt: 'y'.repeat(MAX_EXCERPT_CHARS + 50),
      quads: Array.from({ length: MAX_QUADS + 10 }, (_, i) => quad(i)),
    })
    expect(op?.instruction).toHaveLength(MAX_INSTRUCTION_CHARS)
    expect(op?.excerpt).toHaveLength(MAX_EXCERPT_CHARS)
    expect(op?.quads).toHaveLength(MAX_QUADS)
  })

  it('truncates a fractional page index', () => {
    expect(normalizeAgentAskOp({ ...validOp(), pageIndex: 3.9 })?.pageIndex).toBe(3)
  })
})

describe('normalizeAskQueueItems', () => {
  const validItem = () => ({
    id: ' draft-1 ',
    instruction: ' rewrite the claim ',
    excerpt: '  claim  text ',
    pageIndex: 1,
    quads: [quad()],
  })

  it('normalizes a host snapshot and drops the optional excerpt when absent', () => {
    const items = normalizeAskQueueItems([
      validItem(),
      { ...validItem(), id: 'draft-2', excerpt: undefined },
    ])
    expect(items).toEqual([
      {
        id: 'draft-1',
        instruction: 'rewrite the claim',
        excerpt: 'claim text',
        pageIndex: 1,
        quads: [quad()],
      },
      { id: 'draft-2', instruction: 'rewrite the claim', pageIndex: 1, quads: [quad()] },
    ])
  })

  it('returns an empty list for non-array payloads', () => {
    expect(normalizeAskQueueItems(undefined)).toEqual([])
    expect(normalizeAskQueueItems(null)).toEqual([])
    expect(normalizeAskQueueItems({ id: 'x' })).toEqual([])
  })

  it('drops rows the page overlay could not anchor', () => {
    expect(
      normalizeAskQueueItems([
        { ...validItem(), id: '' },
        { ...validItem(), instruction: '' },
        { ...validItem(), pageIndex: -2 },
        { ...validItem(), pageIndex: undefined },
        { ...validItem(), quads: [] },
        { ...validItem(), quads: [[0, 0, 1, 1]] },
        null,
        'draft',
      ]),
    ).toEqual([])
  })

  it('deduplicates draft ids (anchor/card identity)', () => {
    const items = normalizeAskQueueItems([
      validItem(),
      { ...validItem(), instruction: 'second copy' },
      { ...validItem(), id: 'draft-2' },
    ])
    expect(items.map((i) => i.id)).toEqual(['draft-1', 'draft-2'])
    expect(items[0]?.instruction).toBe('rewrite the claim')
  })

  it('bounds the snapshot size, ids, instruction and excerpt', () => {
    const items = normalizeAskQueueItems(
      Array.from({ length: MAX_QUEUE_ITEMS + 5 }, (_, i) => ({ ...validItem(), id: `d${i}` })),
    )
    expect(items).toHaveLength(MAX_QUEUE_ITEMS)

    const [bounded] = normalizeAskQueueItems([
      {
        ...validItem(),
        id: 'i'.repeat(MAX_DRAFT_ID_CHARS + 10),
        instruction: 'x'.repeat(MAX_INSTRUCTION_CHARS + 10),
        excerpt: 'y'.repeat(MAX_EXCERPT_CHARS + 10),
      },
    ])
    expect(bounded?.id).toHaveLength(MAX_DRAFT_ID_CHARS)
    expect(bounded?.instruction).toHaveLength(MAX_INSTRUCTION_CHARS)
    expect(bounded?.excerpt).toHaveLength(MAX_EXCERPT_CHARS)
  })
})

describe('normalizeAskQueueAction', () => {
  it('accepts the three host actions', () => {
    expect(normalizeAskQueueAction({ kind: 'send-all' })).toEqual({ kind: 'send-all' })
    expect(normalizeAskQueueAction({ kind: 'remove', draftId: 'draft-1' })).toEqual({
      kind: 'remove',
      draftId: 'draft-1',
    })
    expect(
      normalizeAskQueueAction({ kind: 'update', draftId: 'draft-1', instruction: ' fix ' }),
    ).toEqual({
      kind: 'update',
      draftId: 'draft-1',
      instruction: 'fix',
    })
  })

  it('rejects unknown kinds and incomplete payloads', () => {
    expect(normalizeAskQueueAction(null)).toBeNull()
    expect(normalizeAskQueueAction({})).toBeNull()
    expect(normalizeAskQueueAction({ kind: 'send-all', draftId: '' })).toEqual({ kind: 'send-all' })
    expect(normalizeAskQueueAction({ kind: 'remove' })).toBeNull()
    expect(normalizeAskQueueAction({ kind: 'remove', draftId: '   ' })).toBeNull()
    expect(normalizeAskQueueAction({ kind: 'update', draftId: 'draft-1' })).toBeNull()
    expect(
      normalizeAskQueueAction({ kind: 'update', draftId: 'draft-1', instruction: ' ' }),
    ).toBeNull()
    expect(normalizeAskQueueAction({ kind: 'explode', draftId: 'draft-1' })).toBeNull()
  })

  it('bounds draft ids and update text', () => {
    const action = normalizeAskQueueAction({
      kind: 'update',
      draftId: 'd'.repeat(200),
      instruction: 'x'.repeat(MAX_INSTRUCTION_CHARS + 10),
    })
    expect(action).toEqual({
      kind: 'update',
      draftId: 'd'.repeat(MAX_DRAFT_ID_CHARS),
      instruction: 'x'.repeat(MAX_INSTRUCTION_CHARS),
    })
  })
})
