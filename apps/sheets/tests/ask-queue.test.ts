import { describe, expect, it } from 'vitest'
import {
  MAX_ASK_QUEUE_ITEMS,
  collapseExcerpt,
  excerptFromMatrix,
  normalizeSheetsAgentAskOp,
  normalizeSheetsAskQueueAction,
  normalizeSheetsAskQueueItems,
} from '../src/shared/ask-queue'

describe('sheets agent ask op normalization', () => {
  it('keeps the selection locator, excerpt and queue mode', () => {
    expect(
      normalizeSheetsAgentAskOp({
        instruction: '  Round the totals  ',
        sheetId: ' sheet-1 ',
        sheetName: ' Q1 ',
        range: ' B2:D5 ',
        excerpt: '  Total\n1  2  ',
        mode: 'queue',
      }),
    ).toEqual({
      instruction: 'Round the totals',
      sheetId: 'sheet-1',
      sheetName: 'Q1',
      range: 'B2:D5',
      excerpt: 'Total 1 2',
      mode: 'queue',
    })
  })

  it('preserves send-now mode and drops unknown modes', () => {
    expect(normalizeSheetsAgentAskOp({ instruction: 'go', mode: 'now' })?.mode).toBe('now')
    expect(normalizeSheetsAgentAskOp({ instruction: 'go', mode: 'later' })?.mode).toBeUndefined()
  })

  it('rejects empty instructions and non-object payloads', () => {
    expect(normalizeSheetsAgentAskOp(null)).toBeNull()
    expect(normalizeSheetsAgentAskOp('nope')).toBeNull()
    expect(normalizeSheetsAgentAskOp({ instruction: '   ' })).toBeNull()
  })

  it('bounds instruction and excerpt lengths', () => {
    const op = normalizeSheetsAgentAskOp({
      instruction: 'x'.repeat(3000),
      excerpt: 'y'.repeat(500),
    })
    expect(op?.instruction).toHaveLength(2000)
    expect(op?.excerpt).toHaveLength(200)
  })

  it('drops locator fields that are not strings', () => {
    const op = normalizeSheetsAgentAskOp({
      instruction: 'go',
      sheetId: 7,
      sheetName: { name: 'Q1' },
      range: ['A1:B2'],
    })
    expect(op).toEqual({ instruction: 'go' })
  })
})

describe('sheets ask queue snapshot normalization', () => {
  const valid = {
    id: 'draft-1',
    instruction: 'Bold the header',
    excerpt: '  Name\n Q1 ',
    sheetId: 'sheet-1',
    sheetName: ' Q1 ',
    range: ' A1:C1 ',
  }

  it('normalizes a flat host snapshot', () => {
    expect(normalizeSheetsAskQueueItems([valid])).toEqual([
      {
        id: 'draft-1',
        instruction: 'Bold the header',
        excerpt: 'Name Q1',
        sheetId: 'sheet-1',
        sheetName: 'Q1',
        range: 'A1:C1',
      },
    ])
  })

  it('drops malformed rows instead of failing the snapshot', () => {
    expect(
      normalizeSheetsAskQueueItems([
        null,
        'junk',
        { ...valid, id: '' },
        { ...valid, instruction: '  ' },
        valid,
      ]),
    ).toHaveLength(1)
  })

  it('keeps a note without a locator (the grid just cannot anchor it)', () => {
    expect(
      normalizeSheetsAskQueueItems([
        { ...valid, sheetId: undefined, range: 3, sheetName: undefined },
      ]),
    ).toEqual([{ id: 'draft-1', instruction: 'Bold the header', excerpt: 'Name Q1' }])
  })

  it('omits empty optional fields', () => {
    expect(
      normalizeSheetsAskQueueItems([
        { id: 'd', instruction: 'n', sheetId: 's', range: 'A1', excerpt: '   ' },
      ]),
    ).toEqual([{ id: 'd', instruction: 'n', sheetId: 's', range: 'A1' }])
  })

  it('caps the snapshot size', () => {
    const many = Array.from({ length: MAX_ASK_QUEUE_ITEMS + 20 }, (_value, index) => ({
      ...valid,
      id: `draft-${index}`,
    }))
    expect(normalizeSheetsAskQueueItems(many)).toHaveLength(MAX_ASK_QUEUE_ITEMS)
  })

  it('returns an empty list for non-array payloads', () => {
    expect(normalizeSheetsAskQueueItems(undefined)).toEqual([])
    expect(normalizeSheetsAskQueueItems({ items: [] })).toEqual([])
  })
})

describe('sheets ask queue action normalization', () => {
  it('accepts the docs-shaped actions', () => {
    expect(normalizeSheetsAskQueueAction({ kind: 'send-all' })).toEqual({ kind: 'send-all' })
    expect(normalizeSheetsAskQueueAction({ kind: 'remove', draftId: ' d1 ' })).toEqual({
      kind: 'remove',
      draftId: 'd1',
    })
    expect(
      normalizeSheetsAskQueueAction({ kind: 'update', draftId: 'd1', instruction: '  new\nnote ' }),
    ).toEqual({ kind: 'update', draftId: 'd1', instruction: 'new note' })
  })

  it('rejects actions without a target or with an empty update', () => {
    expect(normalizeSheetsAskQueueAction(null)).toBeNull()
    expect(normalizeSheetsAskQueueAction({ kind: 'remove' })).toBeNull()
    expect(
      normalizeSheetsAskQueueAction({ kind: 'update', draftId: 'd1', instruction: '  ' }),
    ).toBeNull()
    expect(normalizeSheetsAskQueueAction({ kind: 'unknown', draftId: 'd1' })).toBeNull()
  })
})

describe('selection excerpt helpers', () => {
  it('collapses whitespace and caps at 200 characters', () => {
    expect(collapseExcerpt('  a\n\tb  ')).toBe('a b')
    expect(collapseExcerpt('x'.repeat(400))).toHaveLength(200)
  })

  it('joins displayed cell values and skips blanks', () => {
    expect(
      excerptFromMatrix([
        ['Name', null, 'Q1'],
        [1, 2, 3],
      ]),
    ).toBe('Name Q1 1 2 3')
  })

  it('stops early once the cap is reached', () => {
    const matrix = [Array.from({ length: 20 }, () => 'y'.repeat(50))]
    expect(excerptFromMatrix(matrix)).toHaveLength(200)
  })
})
