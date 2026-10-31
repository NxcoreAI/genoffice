import { describe, expect, it } from 'vitest'

import {
  expandToPrimitiveOps,
  layoutOpLabel,
  MAX_DATA_TOOL_CELLS,
  structuralOpLabel,
  workbookCommandBatchSchema,
  workbookOperationSchema,
  type WorkbookOperation,
} from '../src/domain/workbook-dsl'
import { computeSortChanges } from '../src/domain/sort-range'
import { InMemoryWorkbookAdapter, WorkbookConflictError } from '../src/domain/in-memory-workbook'
import type { CellState, WorkbookSnapshot } from '../src/domain/workbook.types'

// ---------------------------------------------------------------------------
// Pure harness: the new batch-1 ops expand against a plain cell map instead of
// a live Univer/Electron grid, so the tests read the same way the adapter does.
// ---------------------------------------------------------------------------

type CellMap = Record<string, CellState>

function makeReader(bySheet: Record<string, CellMap>) {
  return (address: string, sheetId: string): CellState =>
    bySheet[sheetId]?.[address] ?? { value: null }
}

/** computeSortChanges reads single-sheet, so it takes a one-argument reader */
function makeGridReader(cells: CellMap) {
  return (address: string): CellState => cells[address] ?? { value: null }
}

function makeSheetNames(names: Record<string, string>) {
  return (sheetId: string): string => names[sheetId] ?? sheetId
}

function expandOne(
  operation: WorkbookOperation,
  cells: Record<string, CellMap> = {},
  names: Record<string, string> = {},
) {
  return expandToPrimitiveOps([operation], makeReader(cells), makeSheetNames(names))
}

/** address → written value, for the ops that only write cells */
function writtenValues(
  ops: readonly { op: string; address?: unknown; value?: unknown }[],
): Record<string, unknown> {
  return Object.fromEntries(
    ops
      .filter((op) => op.op === 'set_cell' && typeof op.address === 'string')
      .map((op) => [op.address as string, op.value]),
  )
}

function parseOne(input: unknown) {
  return workbookOperationSchema.safeParse(input)
}

describe('batch-1 op schemas', () => {
  it('parses sort_range in both the keys and the legacy single-key form', () => {
    expect(
      parseOne({
        op: 'sort_range',
        sheetId: 's1',
        range: 'A1:C5',
        keys: [
          { byColumn: 'B', order: 'desc' },
          { byColumn: 'C', order: 'asc' },
        ],
        hasHeader: true,
      }).success,
    ).toBe(true)
    expect(
      parseOne({ op: 'sort_range', sheetId: 's1', range: 'A1:C5', byColumn: 'B', order: 'asc' })
        .success,
    ).toBe(true)
  })

  it('caps sort keys at three and accepts at most one legacy order value', () => {
    const key = { byColumn: 'A', order: 'asc' as const }
    expect(
      parseOne({ op: 'sort_range', sheetId: 's1', range: 'A1:D5', keys: [key, key, key] }).success,
    ).toBe(true)
    const fourKeys = parseOne({
      op: 'sort_range',
      sheetId: 's1',
      range: 'A1:D5',
      keys: [key, key, key, key],
    })
    expect(fourKeys.success).toBe(false)
    expect(
      parseOne({ op: 'sort_range', sheetId: 's1', range: 'A1:C5', byColumn: 'B', order: 'up' })
        .success,
    ).toBe(false)
  })

  it('requires a sort key at expansion time (a schema refinement cannot live on a union member)', () => {
    const cells = { s1: { A1: { value: 1 }, A2: { value: 2 } } }
    expect(() => expandOne({ op: 'sort_range', sheetId: 's1', range: 'A1:A2' }, cells)).toThrow(
      /sort_range needs a sort key.*keys.*byColumn.*order/s,
    )
    // half of the legacy pair is not a key either
    expect(() =>
      expandOne({ op: 'sort_range', sheetId: 's1', range: 'A1:A2', byColumn: 'A' }, cells),
    ).toThrow(/needs a sort key/)
  })

  it('parses the four new data-tool ops', () => {
    expect(
      parseOne({ op: 'remove_duplicates', sheetId: 's1', range: 'A1:B9', hasHeader: true }).success,
    ).toBe(true)
    expect(
      parseOne({
        op: 'text_to_columns',
        sheetId: 's1',
        source: 'A2:A9',
        delimiter: ',',
        destination: 'C2',
        collapseConsecutive: true,
        clearSource: false,
      }).success,
    ).toBe(true)
    expect(
      parseOne({ op: 'text_to_columns', sheetId: 's1', source: 'A2:A9', delimiter: null }).success,
    ).toBe(true)
    expect(
      parseOne({
        op: 'consolidate',
        sheetId: 's1',
        target: 'D2',
        function: 'average',
        references: [{ range: 'A1:B3' }, { sheetId: 's2', range: 'A1:B3' }],
      }).success,
    ).toBe(true)
    expect(
      parseOne({ op: 'create_names_from_selection', sheetId: 's1', range: 'A1:B3', label: 'top' })
        .success,
    ).toBe(true)
    // a data tool with a bad function / delimiter / label direction is rejected up front
    expect(
      parseOne({
        op: 'consolidate',
        sheetId: 's1',
        target: 'D2',
        function: 'median',
        references: [{ range: 'A1:B3' }],
      }).success,
    ).toBe(false)
    expect(
      parseOne({ op: 'text_to_columns', sheetId: 's1', source: 'A1:A2', delimiter: '' }).success,
    ).toBe(false)
    expect(
      parseOne({
        op: 'create_names_from_selection',
        sheetId: 's1',
        range: 'A1:B3',
        label: 'middle',
      }).success,
    ).toBe(false)
  })

  it('parses the four outline ops', () => {
    for (const op of [
      'outline.group',
      'outline.ungroup',
      'outline.hideDetail',
      'outline.showDetail',
    ]) {
      expect(parseOne({ op, sheetId: 's1', axis: 'rows', range: '5:12' }).success).toBe(true)
      expect(parseOne({ op, sheetId: 's1', axis: 'cols', range: 'B:D' }).success).toBe(true)
      expect(parseOne({ op, sheetId: 's1', axis: 'sideways', range: '5:12' }).success).toBe(false)
      // "5:B" is neither a row span, a column span, nor an A1 range
      expect(parseOne({ op, sheetId: 's1', axis: 'rows', range: '5:B' }).success).toBe(false)
    }
  })

  it('rejects outline spans that do not match their axis', () => {
    expect(() =>
      expandOne({ op: 'outline.group', sheetId: 's1', axis: 'rows', range: 'B:D' }),
    ).toThrow(/is a column span but axis is "rows".*row span like "5:12"/s)
    expect(() =>
      expandOne({ op: 'outline.hideDetail', sheetId: 's1', axis: 'cols', range: '5:12' }),
    ).toThrow(/is a row span but axis is "cols".*column span like "B:D"/s)
  })

  it('requires exactly one of values / conditions on set_filter_criteria', () => {
    const both = () =>
      expandOne({
        op: 'set_filter_criteria',
        sheetId: 's1',
        column: 'B',
        values: ['East'],
        conditions: [{ operator: 'greaterThan', value: '100' }],
      })
    expect(both).toThrow(/exactly one of "values".*and "conditions"/s)
    const neither = () => expandOne({ op: 'set_filter_criteria', sheetId: 's1', column: 'B' })
    expect(neither).toThrow(/exactly one of "values"/)
    // values:null clears the column's criteria — that is still "one of" them
    expect(() =>
      expandOne({ op: 'set_filter_criteria', sheetId: 's1', column: 'B', values: null }),
    ).not.toThrow()
  })
})

describe('sort_range: multi-key expansion', () => {
  // keys: Team desc, then Score asc — blanks stay last in both directions.
  const cells: Record<string, CellMap> = {
    s1: {
      A1: { value: 'Name' },
      B1: { value: 'Team' },
      C1: { value: 'Score' },
      A2: { value: 'Ana' },
      B2: { value: 'Blue' },
      C2: { value: 10 },
      A3: { value: 'Ben' },
      B3: { value: 'Red' },
      C3: { value: 30 },
      A4: { value: 'Cy' },
      B4: { value: 'Blue' },
      C4: { value: 20 },
      A5: { value: 'Di' },
      B5: { value: 'Red' },
      C5: { value: null },
    },
  }

  it('orders rows by the second key within the first and reorders whole rows', () => {
    const ops = expandOne(
      {
        op: 'sort_range',
        sheetId: 's1',
        range: 'A1:C5',
        keys: [
          { byColumn: 'B', order: 'desc' },
          { byColumn: 'C', order: 'asc' },
        ],
        hasHeader: true,
      },
      cells,
    )
    expect(writtenValues(ops)).toEqual({
      A2: 'Ben',
      B2: 'Red',
      C2: 30,
      A3: 'Di',
      C3: null,
      A4: 'Ana',
      C4: 10,
      A5: 'Cy',
      B5: 'Blue',
      C5: 20,
    })
  })

  it('emits expectedValue = the pre-sort content of the target cell, skipping no-op cells', () => {
    const ops = expandOne(
      {
        op: 'sort_range',
        sheetId: 's1',
        range: 'A1:C5',
        keys: [
          { byColumn: 'B', order: 'desc' },
          { byColumn: 'C', order: 'asc' },
        ],
        hasHeader: true,
      },
      cells,
    )
    expect(ops).toEqual([
      { op: 'set_cell', sheetId: 's1', address: 'A2', value: 'Ben', expectedValue: 'Ana' },
      { op: 'set_cell', sheetId: 's1', address: 'B2', value: 'Red', expectedValue: 'Blue' },
      { op: 'set_cell', sheetId: 's1', address: 'C2', value: 30, expectedValue: 10 },
      { op: 'set_cell', sheetId: 's1', address: 'A3', value: 'Di', expectedValue: 'Ben' },
      // B3 stays "Red" for the second row too — no rewrite
      { op: 'set_cell', sheetId: 's1', address: 'C3', value: null, expectedValue: 30 },
      { op: 'set_cell', sheetId: 's1', address: 'A4', value: 'Ana', expectedValue: 'Cy' },
      { op: 'set_cell', sheetId: 's1', address: 'C4', value: 10, expectedValue: 20 },
      { op: 'set_cell', sheetId: 's1', address: 'A5', value: 'Cy', expectedValue: 'Di' },
      { op: 'set_cell', sheetId: 's1', address: 'B5', value: 'Blue', expectedValue: 'Red' },
      { op: 'set_cell', sheetId: 's1', address: 'C5', value: 20, expectedValue: null },
    ])
  })

  it('keeps blanks last under a descending secondary key', () => {
    const blankCells: CellMap = {
      A1: { value: 'x' },
      B1: { value: 1 },
      A2: { value: 'x' },
      B2: { value: 3 },
      A3: { value: 'x' },
      B3: { value: null },
      A4: { value: 'y' },
      B4: { value: 2 },
    }
    const changes = computeSortChanges(
      {
        range: 'A1:B4',
        hasHeader: false,
        keys: [
          { byColumn: 'A', ascending: true },
          { byColumn: 'B', ascending: false },
        ],
      },
      makeGridReader(blankCells),
    )
    // x-group sorted 3, 1, blank; the y row is untouched
    expect(changes).toEqual([
      { address: 'B1', before: 1, after: 3 },
      { address: 'B2', before: 3, after: 1 },
    ])
  })

  it('expands with the legacy byColumn/order pair exactly like a one-key sort', () => {
    const legacy = expandOne(
      {
        op: 'sort_range',
        sheetId: 's1',
        range: 'A1:C5',
        byColumn: 'C',
        order: 'desc',
        hasHeader: true,
      },
      cells,
    )
    expect(writtenValues(legacy)).toMatchObject({ C2: 30, C3: 20, C4: 10 })
  })

  it('rejects more than three keys at the sorting level and keys outside the range', () => {
    const fourKeys = (
      [
        { byColumn: 'A', ascending: true },
        { byColumn: 'B', ascending: true },
        { byColumn: 'C', ascending: true },
        { byColumn: 'D', ascending: true },
      ] as const
    ).map((key) => ({ ...key }))
    expect(() =>
      computeSortChanges(
        { range: 'A1:D5', hasHeader: false, keys: fourKeys },
        makeGridReader(cells.s1 ?? {}),
      ),
    ).toThrow(/at most 3 sort keys/)
    expect(() =>
      expandOne(
        {
          op: 'sort_range',
          sheetId: 's1',
          range: 'A1:C5',
          keys: [{ byColumn: 'D', order: 'asc' }],
        },
        cells,
      ),
    ).toThrow(/outside the range A1:C5/)
  })
})

describe('remove_duplicates', () => {
  const cells: Record<string, CellMap> = {
    s1: {
      A1: { value: 'Name' },
      B1: { value: 'Qty' },
      A2: { value: 'a' },
      B2: { value: 1 },
      A3: { value: 'b' },
      B3: { value: 2 },
      A4: { value: 'A' },
      B4: { value: 1 },
      A5: { value: 'c' },
      B5: { value: 3 },
      A6: { value: 'b' },
      B6: { value: 2 },
    },
  }

  it('deletes later duplicates as whole rows, bottom-up, keeping the first occurrence', () => {
    const ops = expandOne(
      { op: 'remove_duplicates', sheetId: 's1', range: 'A1:B6', hasHeader: true },
      cells,
    )
    expect(ops).toEqual([
      { op: 'delete_rows', sheetId: 's1', row: 6, count: 1 },
      { op: 'delete_rows', sheetId: 's1', row: 4, count: 1 },
    ])
    // and the labels name the rows the plan will delete
    expect(
      ops.map((op) => (op.op === 'delete_rows' ? structuralOpLabel(op) : 'not structural')),
    ).toEqual(['Delete row 6', 'Delete row 4'])
  })

  it('compares text case-insensitively but never matches the header row', () => {
    const headerRepeat: Record<string, CellMap> = {
      s1: {
        A1: { value: 'Name' },
        A2: { value: 'Name' },
        A3: { value: 'name' },
      },
    }
    expect(
      expandOne(
        { op: 'remove_duplicates', sheetId: 's1', range: 'A1:A3', hasHeader: true },
        headerRepeat,
      ),
    ).toEqual([{ op: 'delete_rows', sheetId: 's1', row: 3, count: 1 }])
    // without hasHeader every row is data, so rows 2 and 3 both duplicate row 1
    expect(
      expandOne({ op: 'remove_duplicates', sheetId: 's1', range: 'A1:A3' }, headerRepeat),
    ).toEqual([
      { op: 'delete_rows', sheetId: 's1', row: 3, count: 1 },
      { op: 'delete_rows', sheetId: 's1', row: 2, count: 1 },
    ])
  })

  it('is treated as structural: it may not share a batch with a content edit', () => {
    expect(
      workbookCommandBatchSchema.safeParse({
        dslVersion: 1,
        transactionId: 'tx-1',
        baseRevision: 0,
        summary: 'Dedupe A1:B9',
        operations: [{ op: 'remove_duplicates', sheetId: 's1', range: 'A1:B9', hasHeader: true }],
      }).success,
    ).toBe(true)

    const mixed = workbookCommandBatchSchema.safeParse({
      dslVersion: 1,
      transactionId: 'tx-2',
      baseRevision: 0,
      summary: 'Dedupe then write',
      operations: [
        { op: 'remove_duplicates', sheetId: 's1', range: 'A1:B9', hasHeader: true },
        { op: 'set_cell', sheetId: 's1', address: 'A1', value: 'x' },
      ],
    })
    expect(mixed.success).toBe(false)
    expect(JSON.stringify(mixed.error?.issues)).toMatch(/separate batches/)
  })

  it('enforces the data-tool cell budget', () => {
    expect(MAX_DATA_TOOL_CELLS).toBe(20_000)
    expect(() =>
      expandOne({ op: 'remove_duplicates', sheetId: 's1', range: 'A1:A20001' }, cells),
    ).toThrow(/remove_duplicates covers 20,001 cells — the limit is 20,000/)
  })

  it('needs the current grid to plan against', () => {
    expect(() =>
      expandToPrimitiveOps([
        { op: 'remove_duplicates', sheetId: 's1', range: 'A1:B6', hasHeader: true },
      ]),
    ).toThrow(/remove_duplicates needs the current cell contents/)
  })
})

describe('text_to_columns', () => {
  it('splits on a comma in place and clears the stale trailing cells', () => {
    const cells: Record<string, CellMap> = {
      s1: { A2: { value: 'a,b,c' }, A3: { value: 'd,e' } },
    }
    expect(
      expandOne({ op: 'text_to_columns', sheetId: 's1', source: 'A2:A3', delimiter: ',' }, cells),
    ).toEqual([
      { op: 'set_cell', sheetId: 's1', address: 'A2', value: 'a' },
      { op: 'set_cell', sheetId: 's1', address: 'B2', value: 'b' },
      { op: 'set_cell', sheetId: 's1', address: 'C2', value: 'c' },
      { op: 'set_cell', sheetId: 's1', address: 'A3', value: 'd' },
      { op: 'set_cell', sheetId: 's1', address: 'B3', value: 'e' },
      // row 3 only had two pieces, so C3's old content is wiped
      { op: 'clear_cell', sheetId: 's1', address: 'C3' },
    ])
  })

  it('collapses runs of the delimiter by default, or keeps the empty pieces when told to', () => {
    const cells: Record<string, CellMap> = { s1: { A1: { value: 'a,,b' } } }
    // default: consecutive delimiters fold into one — two pieces, so nothing
    // beyond B1 is touched (the widest split sets the trailing-clear width)
    expect(
      expandOne({ op: 'text_to_columns', sheetId: 's1', source: 'A1', delimiter: ',' }, cells),
    ).toEqual([
      { op: 'set_cell', sheetId: 's1', address: 'A1', value: 'a' },
      { op: 'set_cell', sheetId: 's1', address: 'B1', value: 'b' },
    ])
    // collapseConsecutive:false keeps the empty piece, which clears that cell
    expect(
      expandOne(
        {
          op: 'text_to_columns',
          sheetId: 's1',
          source: 'A1',
          delimiter: ',',
          collapseConsecutive: false,
        },
        cells,
      ),
    ).toEqual([
      { op: 'set_cell', sheetId: 's1', address: 'A1', value: 'a' },
      { op: 'set_cell', sheetId: 's1', address: 'B1', value: null },
      { op: 'set_cell', sheetId: 's1', address: 'C1', value: 'b' },
    ])
  })

  it('treats a null delimiter as any run of whitespace', () => {
    const cells: Record<string, CellMap> = { s1: { A1: { value: '  x   y  ' } } }
    expect(
      expandOne({ op: 'text_to_columns', sheetId: 's1', source: 'A1', delimiter: null }, cells),
    ).toEqual([
      { op: 'set_cell', sheetId: 's1', address: 'A1', value: 'x' },
      { op: 'set_cell', sheetId: 's1', address: 'B1', value: 'y' },
    ])
  })

  it('clears the whole source before writing to a destination elsewhere', () => {
    const cells: Record<string, CellMap> = {
      s1: { A1: { value: 'x,y' }, A2: { value: 'z' } },
    }
    expect(
      expandOne(
        {
          op: 'text_to_columns',
          sheetId: 's1',
          source: 'A1:A2',
          delimiter: ',',
          destination: 'C1',
        },
        cells,
      ),
    ).toEqual([
      // clearing first: an away destination may overlap the source
      { op: 'clear_cell', sheetId: 's1', address: 'A1' },
      { op: 'clear_cell', sheetId: 's1', address: 'A2' },
      { op: 'set_cell', sheetId: 's1', address: 'C1', value: 'x' },
      { op: 'set_cell', sheetId: 's1', address: 'D1', value: 'y' },
      { op: 'set_cell', sheetId: 's1', address: 'C2', value: 'z' },
    ])
  })

  it('keeps the source when clearSource is false', () => {
    const cells: Record<string, CellMap> = {
      s1: { A1: { value: 'x,y' }, A2: { value: 'z' } },
    }
    expect(
      expandOne(
        {
          op: 'text_to_columns',
          sheetId: 's1',
          source: 'A1:A2',
          delimiter: ',',
          destination: 'C1',
          clearSource: false,
        },
        cells,
      ),
    ).toEqual([
      { op: 'set_cell', sheetId: 's1', address: 'C1', value: 'x' },
      { op: 'set_cell', sheetId: 's1', address: 'D1', value: 'y' },
      { op: 'set_cell', sheetId: 's1', address: 'C2', value: 'z' },
    ])
  })

  it('splits a single row downward and rejects a two-dimensional source', () => {
    const cells: Record<string, CellMap> = {
      s1: { A1: { value: 'p|q' }, B1: { value: 'r' } },
    }
    // a one-row source spills DOWN within each source column; the column that
    // split into fewer pieces gets its stale trailing cell cleared
    expect(
      expandOne({ op: 'text_to_columns', sheetId: 's1', source: 'A1:B1', delimiter: '|' }, cells),
    ).toEqual([
      { op: 'set_cell', sheetId: 's1', address: 'A1', value: 'p' },
      { op: 'set_cell', sheetId: 's1', address: 'A2', value: 'q' },
      { op: 'set_cell', sheetId: 's1', address: 'B1', value: 'r' },
      { op: 'clear_cell', sheetId: 's1', address: 'B2' },
    ])
    expect(() =>
      expandOne({ op: 'text_to_columns', sheetId: 's1', source: 'A1:C3', delimiter: ',' }, cells),
    ).toThrow(/must be a single column.*or a single row/s)
  })

  it('rejects a formula cell whose computed value is unavailable', () => {
    const cells: Record<string, CellMap> = {
      s1: { A1: { value: null, formula: '=B1&","&C1' } },
    }
    expect(() =>
      expandOne({ op: 'text_to_columns', sheetId: 's1', source: 'A1', delimiter: ',' }, cells),
    ).toThrow(/found a formula at A1 whose computed value is unavailable.*convert_to_values/s)
  })

  it('enforces the data-tool cell budget on the source', () => {
    expect(() =>
      expandOne({ op: 'text_to_columns', sheetId: 's1', source: 'A1:A20001', delimiter: ',' }),
    ).toThrow(/text_to_columns source covers more than 20,000 cells/)
  })
})

describe('consolidate', () => {
  it('writes one bare-reference formula per target cell for same-sheet sources', () => {
    const ops = expandOne(
      {
        op: 'consolidate',
        sheetId: 's1',
        target: 'D2',
        function: 'sum',
        references: [{ range: 'A1:B2' }, { range: 'C1:D2' }],
      },
      {},
      { s1: 'Sheet1' },
    )
    expect(ops).toEqual([
      { op: 'set_formula', sheetId: 's1', address: 'D2', formula: '=SUM(A1,C1)' },
      { op: 'set_formula', sheetId: 's1', address: 'E2', formula: '=SUM(B1,D1)' },
      { op: 'set_formula', sheetId: 's1', address: 'D3', formula: '=SUM(A2,C2)' },
      { op: 'set_formula', sheetId: 's1', address: 'E3', formula: '=SUM(B2,D2)' },
    ])
  })

  it('qualifies cross-sheet references and quotes sheet names only when needed', () => {
    const flat = expandOne(
      {
        op: 'consolidate',
        sheetId: 's1',
        target: 'B2',
        function: 'sum',
        references: [{ range: 'A1:A1' }, { sheetId: 's2', range: 'A1:A1' }],
      },
      {},
      { s1: 'Sheet1', s2: 'Q2' },
    )
    expect(flat).toEqual([
      { op: 'set_formula', sheetId: 's1', address: 'B2', formula: '=SUM(A1,Q2!A1)' },
    ])

    const quoted = expandOne(
      {
        op: 'consolidate',
        sheetId: 's1',
        target: 'B2',
        function: 'max',
        references: [
          { sheetId: 's3', range: 'B2:B2' },
          { sheetId: 's4', range: 'C3:C3' },
        ],
      },
      {},
      { s1: 'Sheet1', s3: 'My Sheet', s4: "It's" },
    )
    expect(quoted).toEqual([
      {
        op: 'set_formula',
        sheetId: 's1',
        address: 'B2',
        formula: "=MAX('My Sheet'!B2,'It''s'!C3)",
      },
    ])
  })

  it('maps every supported function name', () => {
    const formulas = (['sum', 'count', 'average', 'max', 'min'] as const).map((fn) => {
      const [op] = expandOne(
        {
          op: 'consolidate',
          sheetId: 's1',
          target: 'A5',
          function: fn,
          references: [{ range: 'A1' }],
        },
        {},
        { s1: 'Sheet1' },
      )
      return op?.op === 'set_formula' ? op.formula : null
    })
    expect(formulas).toEqual(['=SUM(A1)', '=COUNT(A1)', '=AVERAGE(A1)', '=MAX(A1)', '=MIN(A1)'])
  })

  it('rejects references whose shapes differ', () => {
    expect(() =>
      expandOne(
        {
          op: 'consolidate',
          sheetId: 's1',
          target: 'D2',
          function: 'sum',
          references: [{ range: 'A1:B2' }, { range: 'A1:A2' }],
        },
        {},
        { s1: 'Sheet1' },
      ),
    ).toThrow(/reference #2 \(A1:A2\) is 2×1 but reference #1 \(A1:B2\) is 2×2/)
  })

  it('enforces the data-tool cell budget on the output block', () => {
    expect(() =>
      expandOne(
        {
          op: 'consolidate',
          sheetId: 's1',
          target: 'B1',
          function: 'sum',
          references: [{ range: 'A1:A20001' }],
        },
        {},
        { s1: 'Sheet1' },
      ),
    ).toThrow(/consolidate would write 20001 output cells — the limit is 20,000/)
  })

  it('needs a sheet-name resolver for cross-sheet references', () => {
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'consolidate',
          sheetId: 's1',
          target: 'B2',
          function: 'sum',
          references: [{ range: 'A1' }],
        },
      ]),
    ).toThrow(/consolidate needs sheet names/)
  })
})

describe('create_names_from_selection', () => {
  it('turns a top label row into defined names with absolute refs', () => {
    const cells: Record<string, CellMap> = {
      s1: {
        A1: { value: 'Revenue' },
        B1: { value: 'Cost' },
        A2: { value: 10 },
        B2: { value: 4 },
        A3: { value: 12 },
        B3: { value: 5 },
      },
    }
    expect(
      expandOne(
        { op: 'create_names_from_selection', sheetId: 's1', range: 'A1:B3', label: 'top' },
        cells,
        { s1: 'Sheet1' },
      ),
    ).toEqual([
      { op: 'add_defined_name', name: 'Revenue', ref: "'Sheet1'!$A$2:$A$3" },
      { op: 'add_defined_name', name: 'Cost', ref: "'Sheet1'!$B$2:$B$3" },
    ])
  })

  it('turns a left label column into one defined name per data row', () => {
    const cells: Record<string, CellMap> = {
      s1: {
        A1: { value: 'Alpha' },
        B1: { value: 1 },
        A2: { value: 'Beta' },
        B2: { value: 2 },
      },
    }
    expect(
      expandOne(
        { op: 'create_names_from_selection', sheetId: 's1', range: 'A1:B2', label: 'left' },
        cells,
        { s1: 'Data Set' },
      ),
    ).toEqual([
      { op: 'add_defined_name', name: 'Alpha', ref: "'Data Set'!$B$1:$B$1" },
      { op: 'add_defined_name', name: 'Beta', ref: "'Data Set'!$B$2:$B$2" },
    ])
  })

  it('sanitizes labels, numbers duplicates, and skips empty ones', () => {
    const cells: Record<string, CellMap> = {
      s1: {
        A1: { value: 'Total Sales' },
        B1: { value: 'Revenue' },
        C1: { value: '' },
        D1: { value: 'revenue' },
        E1: { value: '2nd Qtr' },
        F1: { value: 'A1' },
        G1: { value: '   ' },
        A2: { value: 1 },
        B2: { value: 2 },
        C2: { value: 3 },
        D2: { value: 4 },
        E2: { value: 5 },
        F2: { value: 6 },
        G2: { value: 7 },
      },
    }
    const ops = expandOne(
      { op: 'create_names_from_selection', sheetId: 's1', range: 'A1:G2', label: 'top' },
      cells,
      { s1: 'Sheet1' },
    )
    expect(ops.map((op) => (op.op === 'add_defined_name' ? op.name : op.op))).toEqual([
      'Total_Sales',
      'Revenue',
      // case-insensitive duplicate → numbered suffix
      'revenue_2',
      // leading digit gets an underscore prefix
      '_2nd_Qtr',
      // a cell-like label must not look like a reference
      '_A1',
    ])
    expect(ops).toHaveLength(5)
    expect(ops[0]).toEqual({
      op: 'add_defined_name',
      name: 'Total_Sales',
      ref: "'Sheet1'!$A$2:$A$2",
    })
  })

  it('rejects too-small ranges and all-empty labels', () => {
    expect(() =>
      expandOne(
        { op: 'create_names_from_selection', sheetId: 's1', range: 'A1:B1', label: 'top' },
        {},
        { s1: 'Sheet1' },
      ),
    ).toThrow(/needs at least two rows: one label row plus at least one data row/)

    const oneColumn: Record<string, CellMap> = { s1: { A1: { value: 'x' }, A2: { value: 1 } } }
    expect(() =>
      expandOne(
        { op: 'create_names_from_selection', sheetId: 's1', range: 'A1:A2', label: 'left' },
        oneColumn,
        { s1: 'Sheet1' },
      ),
    ).toThrow(/needs at least two columns: one label column plus at least one data column/)

    expect(() =>
      expandOne(
        { op: 'create_names_from_selection', sheetId: 's1', range: 'A1:B2', label: 'top' },
        { s1: { A2: { value: 1 }, B2: { value: 2 } } },
        { s1: 'Sheet1' },
      ),
    ).toThrow(/found no non-empty labels in the top row\/column of A1:B2/)
  })

  it('needs the grid and the sheet name', () => {
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'create_names_from_selection', sheetId: 's1', range: 'A1:B2', label: 'top' }],
        makeReader({}),
      ),
    ).toThrow(/needs the sheet name to write absolute defined-name references/)
    expect(() =>
      expandToPrimitiveOps([
        { op: 'create_names_from_selection', sheetId: 's1', range: 'A1:B2', label: 'top' },
      ]),
    ).toThrow(/needs the current cell contents to read its labels/)
  })
})

describe('outline ops and plan labels', () => {
  it('passes outline ops through unchanged and labels them', () => {
    const cases = [
      { op: 'outline.group', axis: 'rows', range: '5:12', label: 'Group rows 5:12' },
      { op: 'outline.ungroup', axis: 'cols', range: 'B:D', label: 'Ungroup columns B:D' },
      { op: 'outline.hideDetail', axis: 'rows', range: 'A5:C12', label: 'Hide detail rows A5:C12' },
      {
        op: 'outline.showDetail',
        axis: 'cols',
        range: 'B1:D1',
        label: 'Show detail columns B1:D1',
      },
    ] as const
    for (const { label, ...operation } of cases) {
      const expanded = expandOne({ ...operation, sheetId: 's1' })
      expect(expanded).toEqual([{ ...operation, sheetId: 's1' }])
      expect(layoutOpLabel(expanded[0] as never)).toBe(label)
    }
  })

  it('passes filter conditions through and labels the AND / OR join', () => {
    const and = expandOne({
      op: 'set_filter_criteria',
      sheetId: 's1',
      column: 'B',
      conditions: [
        { operator: 'greaterThanOrEqual', value: '10' },
        { operator: 'lessThan', value: '20' },
      ],
      and: true,
    })
    expect(and).toEqual([
      {
        op: 'set_filter_criteria',
        sheetId: 's1',
        column: 'B',
        conditions: [
          { operator: 'greaterThanOrEqual', value: '10' },
          { operator: 'lessThan', value: '20' },
        ],
        and: true,
      },
    ])
    expect(layoutOpLabel(and[0] as never)).toBe('Filter column B: >= 10 AND < 20')

    const or = expandOne({
      op: 'set_filter_criteria',
      sheetId: 's1',
      column: 'C',
      conditions: [{ operator: 'notEqual', value: 'x' }],
      and: false,
    })
    expect(layoutOpLabel(or[0] as never)).toBe('Filter column C: <> x')

    const list = expandOne({
      op: 'set_filter_criteria',
      sheetId: 's1',
      column: 'D',
      values: ['East', 'South'],
    })
    expect(layoutOpLabel(list[0] as never)).toBe('Filter column D to 2 values')

    const cleared = expandOne({
      op: 'set_filter_criteria',
      sheetId: 's1',
      column: 'D',
      values: null,
    })
    expect(layoutOpLabel(cleared[0] as never)).toBe('Clear filter criteria on column D')
  })

  it('labels the structural deletes remove_duplicates expands into', () => {
    expect(structuralOpLabel({ op: 'delete_rows', sheetId: 's1', row: 4, count: 1 })).toBe(
      'Delete row 4',
    )
    expect(structuralOpLabel({ op: 'delete_rows', sheetId: 's1', row: 4, count: 2 })).toBe(
      'Delete rows 4–5',
    )
    expect(structuralOpLabel({ op: 'insert_rows', sheetId: 's1', row: 3, count: 1 })).toBe(
      'Insert 1 row before row 3',
    )
  })
})

describe('in-memory adapter (demo workbook)', () => {
  function dedupeSnapshot(): WorkbookSnapshot {
    return {
      revision: 0,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: {
            A1: { value: 'Name' },
            B1: { value: 'Qty' },
            A2: { value: 'a' },
            B2: { value: 1 },
            A3: { value: 'b' },
            B3: { value: 2 },
            A4: { value: 'A' },
            B4: { value: 1 },
            A5: { value: 'c' },
            B5: { value: 3 },
            A6: { value: 'b' },
            B6: { value: 2 },
          },
        },
      ],
    }
  }

  it('applies remove_duplicates end to end: first occurrence kept, later rows deleted', () => {
    const adapter = new InMemoryWorkbookAdapter(dedupeSnapshot())
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'dedupe-1',
      baseRevision: 0,
      summary: 'Remove duplicate rows from A1:B6',
      operations: [
        { op: 'remove_duplicates', sheetId: 'sheet-1', range: 'A1:B6', hasHeader: true },
      ],
    })

    // dry run: nothing applied yet, two whole-row deletions planned
    expect(plan.cellChanges).toEqual([])
    expect(plan.structuralChanges.map((change) => change.label)).toEqual([
      'Delete row 6',
      'Delete row 4',
    ])
    expect(adapter.getSnapshot().sheets[0]?.cells.A4?.value).toBe('A')

    expect(adapter.apply(plan).revision).toBe(1)
    const cells = adapter.getSnapshot().sheets[0]?.cells ?? {}
    expect(cells.A1?.value).toBe('Name')
    // row 2 (a/1) is the kept first occurrence; the "A/1" copy at old row 4 is gone
    expect([cells.A2?.value, cells.B2?.value]).toEqual(['a', 1])
    expect([cells.A3?.value, cells.B3?.value]).toEqual(['b', 2])
    // old row 5 (c/3) shifted up by the two deletions
    expect([cells.A4?.value, cells.B4?.value]).toEqual(['c', 3])
    expect(cells.A5).toBeUndefined()
    expect(cells.A6).toBeUndefined()
  })

  it('undo restores the deleted rows', () => {
    const adapter = new InMemoryWorkbookAdapter(dedupeSnapshot())
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'dedupe-2',
      baseRevision: 0,
      summary: 'Remove duplicate rows from A1:B6',
      operations: [
        { op: 'remove_duplicates', sheetId: 'sheet-1', range: 'A1:B6', hasHeader: true },
      ],
    })
    adapter.apply(plan)
    adapter.undo()
    expect(adapter.getSnapshot().sheets[0]?.cells).toEqual(dedupeSnapshot().sheets[0]?.cells)
  })

  it('refuses outline ops: the demo workbook has no outline model', () => {
    const adapter = new InMemoryWorkbookAdapter(dedupeSnapshot())
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'outline-1',
        baseRevision: 0,
        summary: 'Group rows 2-3',
        operations: [{ op: 'outline.group', sheetId: 'sheet-1', axis: 'rows', range: '2:3' }],
      }),
    ).toThrow(WorkbookConflictError)
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'outline-2',
        baseRevision: 0,
        summary: 'Group rows 2-3',
        operations: [{ op: 'outline.group', sheetId: 'sheet-1', axis: 'rows', range: '2:3' }],
      }),
    ).toThrow(/outline\.group requires an imported xlsx file/)
  })
})
