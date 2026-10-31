import { describe, expect, it } from 'vitest'

import {
  expandToPrimitiveOps,
  isLayoutOp,
  isStructuralOp,
  layoutOpLabel,
  structuralOpLabel,
  workbookCommandBatchSchema,
  workbookOperationSchema,
  type EditChartOperation,
  type SubtotalOperation,
  type WorkbookOperation,
} from '../src/domain/workbook-dsl'
import { buildChartVisual } from '../src/domain/chart-visual'
import { InMemoryWorkbookAdapter, WorkbookConflictError } from '../src/domain/in-memory-workbook'
import type { CellState, WorkbookSnapshot } from '../src/domain/workbook.types'
import {
  applySubtotalToRange,
  MAX_SUBTOTAL_GROUPS,
  subtotalFailureMessage,
  type SubtotalFunction,
} from '../src/renderer/subtotal-ops'
import type { UniverRuntime } from '../src/renderer/univer-state'

// ---------------------------------------------------------------------------
// Pure harness: the subtotal op expands against a plain cell map and its
// executor runs against a recording fake worksheet (no Electron / Univer).
// ---------------------------------------------------------------------------

type CellMap = Record<string, CellState>

function makeReader(bySheet: Record<string, CellMap>) {
  return (address: string, sheetId: string): CellState =>
    bySheet[sheetId]?.[address] ?? { value: null }
}

function expandOne(operation: WorkbookOperation, cells: Record<string, CellMap> = {}) {
  return expandToPrimitiveOps([operation], makeReader(cells))
}

function parseOne(input: unknown) {
  return workbookOperationSchema.safeParse(input)
}

function batchOf(operations: unknown[]) {
  return workbookCommandBatchSchema.safeParse({
    dslVersion: 1,
    transactionId: 'tx-1',
    baseRevision: 0,
    summary: 'batch-3 test',
    operations,
  })
}

const SUBTOTAL_FUNCTIONS = [
  'sum',
  'average',
  'count',
  'counta',
  'max',
  'min',
  'product',
  'stdev',
  'stdevp',
  'var',
  'varp',
] as const

function subtotalOp(overrides: Partial<SubtotalOperation> = {}): SubtotalOperation {
  return {
    op: 'subtotal',
    sheetId: 's1',
    range: 'A1:D100',
    groupColumn: 'A',
    valueColumn: 'D',
    function: 'average',
    ...overrides,
  }
}

function editChart(overrides: Partial<EditChartOperation> = {}): EditChartOperation {
  return { op: 'edit_chart', chartPath: 'demo-chart-sheet-1-1', ...overrides }
}

/** deliberately invalid input for schema-rejection tests (the typed helper
 * above cannot express bad enums / colors) */
function rawEditChart(overrides: Record<string, unknown>) {
  return { op: 'edit_chart', chartPath: 'demo-chart-sheet-1-1', ...overrides }
}

// ---------------------------------------------------------------------------
// Recording fake worksheet: exactly the surface applySubtotalToRange touches.
// ---------------------------------------------------------------------------

interface FakeWrite {
  readonly row: number
  readonly column: number
  value: { v?: string; f?: string }
  bold: boolean
}

function fakeSheet(labels: readonly string[]) {
  const writes: FakeWrite[] = []
  const inserts: { row: number; count: number }[] = []
  const reads: { row: number; column: number; rows: number; columns: number }[] = []
  const pending = new Map<string, FakeWrite>()
  let current: FakeWrite | null = null
  const worksheet = {
    getRange(row: number, column: number, rows: number, columns: number) {
      reads.push({ row, column, rows, columns })
      if (rows === labels.length && columns === 1) {
        // the group-label column read
        return {
          getValues: () => labels.map((label) => [label]),
          setValue: () => undefined,
          setFontWeight: () => undefined,
        }
      }
      current = { row, column, value: {}, bold: false }
      pending.set(`${row}:${column}`, current)
      writes.push(current)
      return {
        getValues: () => [],
        setValue(value: { v?: string; f?: string }) {
          if (current) current.value = value
        },
        setFontWeight() {
          if (current) current.bold = true
        },
      }
    },
    insertRowsBefore(row: number, count: number) {
      inserts.push({ row, count })
      return true
    },
  }
  const runtime = {
    univerAPI: {
      getActiveWorkbook: () => ({
        getSheetBySheetId: (sheetId: string) => (sheetId === 'sheet-1' ? worksheet : undefined),
      }),
    },
  } as unknown as UniverRuntime
  return { runtime, writes, inserts, reads }
}

describe('subtotal: schema', () => {
  it('parses every Excel SUBTOTAL function', () => {
    expect(SUBTOTAL_FUNCTIONS).toHaveLength(11)
    for (const fn of SUBTOTAL_FUNCTIONS) {
      expect(parseOne(subtotalOp({ function: fn })).success, `function ${fn}`).toBe(true)
    }
  })

  it('rejects a missing or unknown function and bad geometry', () => {
    const { function: _fn, ...withoutFunction } = subtotalOp()
    expect(parseOne(withoutFunction).success).toBe(false)
    expect(parseOne(subtotalOp({ function: 'median' as SubtotalFunction })).success).toBe(false)
    expect(parseOne(subtotalOp({ range: 'A1:' })).success).toBe(false)
    expect(parseOne(subtotalOp({ groupColumn: '1' })).success).toBe(false)
    expect(parseOne(subtotalOp({ valueColumn: '' })).success).toBe(false)
  })
})

describe('subtotal: expansion validation', () => {
  it('needs a header row plus at least one data row', () => {
    expect(() => expandOne(subtotalOp({ range: 'A1:D1' }))).toThrow(
      /subtotal range A1:D1 needs at least two rows — a header row plus at least one data row\./,
    )
    // header + one data row is the minimum and passes through
    expect(expandOne(subtotalOp({ range: 'A1:D2' }))).toHaveLength(1)
  })

  it('rejects a group/value column outside the range', () => {
    expect(() => expandOne(subtotalOp({ range: 'B1:D100', groupColumn: 'A' }))).toThrow(
      /subtotal groupColumn A is outside the range B1:D100 — the grouping column must be one of the range's columns\./,
    )
    expect(() => expandOne(subtotalOp({ range: 'A1:D100', valueColumn: 'F' }))).toThrow(
      /subtotal valueColumn F is outside the range A1:D100 — the subtotaled column must be one of the range's columns\./,
    )
  })

  it('rejects grouping and subtotaling the same column', () => {
    expect(() => expandOne(subtotalOp({ groupColumn: 'D', valueColumn: 'D' }))).toThrow(
      /subtotal groupColumn and valueColumn are both D — group by one column and subtotal a different \(numeric\) column\./,
    )
  })

  it('passes a valid op through unchanged', () => {
    const operation = subtotalOp()
    expect(expandOne(operation)).toEqual([operation])
  })
})

describe('subtotal: structural classification and label', () => {
  it('is structural, not layout', () => {
    const operation = subtotalOp()
    expect(isStructuralOp(operation)).toBe(true)
    expect(isLayoutOp(operation)).toBe(false)
  })

  it('may share a batch with other structural ops but never with cell edits', () => {
    expect(batchOf([subtotalOp()]).success).toBe(true)
    expect(
      batchOf([subtotalOp(), { op: 'insert_rows', sheetId: 's1', row: 2, count: 1 }]).success,
    ).toBe(true)

    const mixedWithCell = batchOf([
      subtotalOp(),
      { op: 'set_cell', sheetId: 's1', address: 'A1', value: 'x' },
    ])
    expect(mixedWithCell.success).toBe(false)
    expect(JSON.stringify(mixedWithCell.error?.issues)).toMatch(/separate batches/)

    expect(
      batchOf([subtotalOp(), { op: 'clear_formats', sheetId: 's1', range: 'A1:B2' }]).success,
    ).toBe(false)
  })

  it('labels the summary rows it will insert', () => {
    expect(structuralOpLabel(subtotalOp())).toBe(
      'Insert SUBTOTAL average rows on A1:D100 grouped by A',
    )
    expect(
      structuralOpLabel(subtotalOp({ function: 'sum', range: 'B2:E50', groupColumn: 'B' })),
    ).toBe('Insert SUBTOTAL sum rows on B2:E50 grouped by B')
  })

  it('is rejected by the demo workbook', () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [{ id: 's1', name: 'Sheet1', cells: { A1: { value: 'Region' } } }],
    })
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'sub-1',
        baseRevision: 0,
        summary: 'Subtotal A1:D100',
        operations: [subtotalOp()],
      }),
    ).toThrow(WorkbookConflictError)
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'sub-2',
        baseRevision: 0,
        summary: 'Subtotal A1:D100',
        operations: [subtotalOp()],
      }),
    ).toThrow(
      /subtotal requires an imported xlsx file — the demo workbook cannot insert summary rows\./,
    )
  })
})

describe('subtotal-ops: function numbers and failure messages', () => {
  // SUBTOTAL_FUNCTION_NUMBERS is module-private, so the mapping is exercised
  // through the formulas the executor writes.
  const FUNCTION_NUMBERS: Record<SubtotalFunction, number> = {
    average: 1,
    count: 2,
    counta: 3,
    max: 4,
    min: 5,
    product: 6,
    stdev: 7,
    stdevp: 8,
    sum: 9,
    var: 10,
    varp: 11,
  }

  it('writes the Excel function_num for every function', () => {
    for (const fn of SUBTOTAL_FUNCTIONS) {
      const sheet = fakeSheet(['North', 'North'])
      const outcome = applySubtotalToRange(sheet.runtime, null, {
        sheetId: 'sheet-1',
        range: 'A1:D3',
        groupColumn: 'A',
        valueColumn: 'D',
        function: fn,
      })
      expect(outcome, `function ${fn}`).toEqual({ applied: true, groups: 1 })
      const valueWrites = sheet.writes.filter((write) => write.value.f !== undefined)
      expect(valueWrites[0]?.value.f, `function ${fn}`).toBe(
        `=SUBTOTAL(${FUNCTION_NUMBERS[fn]},$D$2:$D$3)`,
      )
    }
  })

  it('exposes the group ceiling', () => {
    expect(MAX_SUBTOTAL_GROUPS).toBe(200)
  })

  it('has English text for every failure reason', () => {
    expect(subtotalFailureMessage('sheet-missing')).toBe(
      'The subtotal sheet is no longer available.',
    )
    expect(subtotalFailureMessage('range-too-small')).toBe(
      'subtotal needs a range with a header row and at least one data row.',
    )
    expect(subtotalFailureMessage('no-data-rows')).toBe(
      'subtotal found no data rows below the header row.',
    )
    expect(subtotalFailureMessage('same-column')).toBe(
      'subtotal needs a groupColumn different from the valueColumn.',
    )
    expect(subtotalFailureMessage('column-outside-range')).toBe(
      'subtotal groupColumn and valueColumn must both lie inside the range.',
    )
    expect(subtotalFailureMessage('too-many-groups')).toBe(
      'subtotal would create more than 200 groups.',
    )
    expect(subtotalFailureMessage('too-many-groups', 251)).toBe(
      'subtotal would create 251 groups — the limit is 200. Sort the data by the group column first.',
    )
  })
})

describe('subtotal-ops: applying the runs', () => {
  it('inserts one bold SUBTOTAL row per run (bottom-up) plus a grand total', () => {
    // header + 4 data rows; runs: North×2, South, West
    const sheet = fakeSheet(['North', 'North', 'South', 'West'])
    const outcome = applySubtotalToRange(sheet.runtime, null, {
      sheetId: 'sheet-1',
      range: 'A1:D5',
      groupColumn: 'A',
      valueColumn: 'D',
      function: 'sum',
    })

    expect(outcome).toEqual({ applied: true, groups: 3 })
    // the group column is read as one column below the header row
    expect(sheet.reads[0]).toEqual({ row: 1, column: 0, rows: 4, columns: 1 })
    // inserts run bottom-up so earlier row indices stay valid
    expect(sheet.inserts).toEqual([
      { row: 5, count: 1 },
      { row: 4, count: 1 },
      { row: 3, count: 1 },
      { row: 8, count: 1 },
    ])
    const labelRows = sheet.writes
      .filter((write) => write.value.v !== undefined)
      .map((write) => [write.row, write.column, write.value.v])
    expect(labelRows).toEqual([
      [5, 0, 'West Total'],
      [4, 0, 'South Total'],
      [3, 0, 'North Total'],
      [8, 0, 'Grand Total'],
    ])
    const formulaRows = sheet.writes
      .filter((write) => write.value.f !== undefined)
      .map((write) => [write.row, write.column, write.value.f])
    expect(formulaRows).toEqual([
      [5, 3, '=SUBTOTAL(9,$D$5:$D$5)'],
      [4, 3, '=SUBTOTAL(9,$D$4:$D$4)'],
      [3, 3, '=SUBTOTAL(9,$D$2:$D$3)'],
      // the grand total spans every data row incl. the inserted summary rows
      [8, 3, '=SUBTOTAL(9,$D$2:$D$8)'],
    ])
    // summary labels and values are bold
    expect(sheet.writes.every((write) => write.bold)).toBe(true)
  })

  it('groups only contiguous runs of equal labels', () => {
    const sheet = fakeSheet(['North', 'South', 'North'])
    const outcome = applySubtotalToRange(sheet.runtime, null, {
      sheetId: 'sheet-1',
      range: 'A1:D4',
      groupColumn: 'A',
      valueColumn: 'D',
      function: 'count',
    })
    expect(outcome).toEqual({ applied: true, groups: 3 })
    // three groups, inserted bottom-up: North(row 2), South(row 3), North(row 4)
    expect(
      sheet.writes
        .filter((write) => write.value.v !== undefined)
        .map((write) => [write.row, write.value.v]),
    ).toEqual([
      [4, 'North Total'],
      [3, 'South Total'],
      [2, 'North Total'],
      [7, 'Grand Total'],
    ])
  })

  it('reports the validation failures as reason codes', () => {
    const sheet = fakeSheet(['North', 'North'])
    const target = {
      sheetId: 'sheet-1',
      range: 'A1:D3',
      groupColumn: 'A',
      valueColumn: 'D',
      function: 'sum' as const,
    }
    expect(applySubtotalToRange(sheet.runtime, null, { ...target, sheetId: 'nope' })).toEqual({
      applied: false,
      reason: 'sheet-missing',
    })
    expect(applySubtotalToRange(sheet.runtime, null, { ...target, range: 'A1:D1' })).toEqual({
      applied: false,
      reason: 'range-too-small',
    })
    expect(applySubtotalToRange(sheet.runtime, null, { ...target, groupColumn: 'F' })).toEqual({
      applied: false,
      reason: 'column-outside-range',
    })
    expect(applySubtotalToRange(sheet.runtime, null, { ...target, valueColumn: 'A' })).toEqual({
      applied: false,
      reason: 'same-column',
    })
    // a missing workbook is a missing sheet, not a crash
    const empty = { univerAPI: { getActiveWorkbook: () => null } } as unknown as UniverRuntime
    expect(applySubtotalToRange(empty, null, target)).toEqual({
      applied: false,
      reason: 'sheet-missing',
    })
  })

  it('refuses to create more groups than Excel allows', () => {
    const labels = Array.from({ length: 202 }, (_, index) => `g${index}`)
    const sheet = fakeSheet(labels)
    const outcome = applySubtotalToRange(sheet.runtime, null, {
      sheetId: 'sheet-1',
      range: `A1:D${labels.length + 1}`,
      groupColumn: 'A',
      valueColumn: 'D',
      function: 'sum',
    })
    expect(outcome).toEqual({ applied: false, reason: 'too-many-groups', groups: 202 })
    expect(sheet.inserts).toEqual([])
    expect(sheet.writes).toEqual([])
  })
})

describe('edit_chart: new fields parse', () => {
  it('accepts series → point → hex point colors and rejects bad keys or colors', () => {
    expect(
      parseOne(editChart({ pointColors: { '0': { '0': '#FF0000', '2': '#00FF00' } } })).success,
    ).toBe(true)
    expect(parseOne(editChart({ pointColors: {} })).success).toBe(true)
    // bad series index / point index / color
    expect(parseOne(editChart({ pointColors: { a: { '0': '#FF0000' } } })).success).toBe(false)
    expect(parseOne(editChart({ pointColors: { '0': { x: '#FF0000' } } })).success).toBe(false)
    expect(parseOne(editChart({ pointColors: { '0': { '0': '#GGGGGG' } } })).success).toBe(false)
    expect(parseOne(editChart({ pointColors: { '0': { '0': 'red' } } })).success).toBe(false)
    // the inner map must be an object of points, not a bare color
    expect(parseOne(rawEditChart({ pointColors: { '0': '#FF0000' } })).success).toBe(false)
  })

  it('accepts the three data-label positions only', () => {
    for (const position of ['center', 'inside-end', 'outside-end'] as const) {
      expect(parseOne(editChart({ dataLabelPosition: position })).success, position).toBe(true)
    }
    expect(parseOne(rawEditChart({ dataLabelPosition: 'outsideEnd' })).success).toBe(false)
    expect(parseOne(rawEditChart({ dataLabelPosition: 'outside' })).success).toBe(false)
  })

  it('bounds the data-label number format at 1..64 characters', () => {
    expect(parseOne(editChart({ dataLabelFormat: '#,##0' })).success).toBe(true)
    expect(parseOne(editChart({ dataLabelFormat: 'a'.repeat(64) })).success).toBe(true)
    expect(parseOne(editChart({ dataLabelFormat: 'a'.repeat(65) })).success).toBe(false)
    expect(parseOne(editChart({ dataLabelFormat: '' })).success).toBe(false)
  })

  it('accepts gridlines as a boolean', () => {
    expect(parseOne(editChart({ gridlines: true })).success).toBe(true)
    expect(parseOne(editChart({ gridlines: false })).success).toBe(true)
    expect(parseOne(rawEditChart({ gridlines: 'off' })).success).toBe(false)
  })

  it('requires valueAxis to name min or max, and allows null (back to auto)', () => {
    expect(parseOne(editChart({ valueAxis: { min: 0 } })).success).toBe(true)
    expect(parseOne(editChart({ valueAxis: { max: 1000 } })).success).toBe(true)
    expect(parseOne(editChart({ valueAxis: { min: 0, max: null } })).success).toBe(true)
    expect(parseOne(editChart({ valueAxis: { min: null, max: null } })).success).toBe(true)
    const empty = parseOne(editChart({ valueAxis: {} }))
    expect(empty.success).toBe(false)
    expect(JSON.stringify(empty.error?.issues)).toMatch(/valueAxis needs min or max/)
    expect(parseOne(editChart({ valueAxis: { min: Number.POSITIVE_INFINITY } })).success).toBe(
      false,
    )
    expect(parseOne(editChart({ valueAxis: { min: Number.NaN } })).success).toBe(false)
  })

  it('treats each new field as a complete edit on its own', () => {
    for (const operation of [
      editChart({ pointColors: { '0': { '1': '#FF0000' } } }),
      editChart({ dataLabelPosition: 'outside-end' }),
      editChart({ dataLabelFormat: '#,##0' }),
      editChart({ gridlines: false }),
      editChart({ valueAxis: { min: 0 } }),
    ]) {
      expect(expandOne(operation), JSON.stringify(operation)).toEqual([operation])
    }
  })

  it('still rejects an edit with no property, naming the new fields', () => {
    expect(() => expandOne(editChart())).toThrow(
      /edit_chart needs at least one of title \/ chartType \/ seriesColors \/ pointColors \/ legend \/ dataLabels \/ dataLabelPosition \/ dataLabelFormat \/ grouping \/ gridlines \/ valueAxis \/ axisTitles \/ seriesData\./,
    )
    // an empty pointColors map is not an edit either
    expect(() => expandOne(editChart({ pointColors: {} }))).toThrow(/needs at least one of/)
  })
})

describe('edit_chart: new field labels', () => {
  const label = (operation: EditChartOperation) => layoutOpLabel(operation)

  it('labels gridlines, value-axis bounds, label position/format, and point colors', () => {
    expect(label(editChart({ gridlines: true }))).toBe(
      'Edit chart demo-chart-sheet-1-1: gridlines on',
    )
    expect(label(editChart({ gridlines: false }))).toBe(
      'Edit chart demo-chart-sheet-1-1: gridlines off',
    )
    expect(label(editChart({ valueAxis: { min: 0, max: null } }))).toBe(
      'Edit chart demo-chart-sheet-1-1: value axis min 0 / max auto',
    )
    expect(label(editChart({ valueAxis: { min: null } }))).toBe(
      'Edit chart demo-chart-sheet-1-1: value axis min auto',
    )
    expect(label(editChart({ valueAxis: { max: 1000 } }))).toBe(
      'Edit chart demo-chart-sheet-1-1: value axis max 1000',
    )
    expect(label(editChart({ dataLabelPosition: 'outside-end' }))).toBe(
      'Edit chart demo-chart-sheet-1-1: label position outside-end',
    )
    expect(label(editChart({ dataLabelFormat: '#,##0' }))).toBe(
      'Edit chart demo-chart-sheet-1-1: label format "#,##0"',
    )
    // singular / plural point counts across series
    expect(label(editChart({ pointColors: { '0': { '1': '#FF0000' } } }))).toBe(
      'Edit chart demo-chart-sheet-1-1: 1 point color',
    )
    expect(
      label(
        editChart({
          pointColors: { '0': { '0': '#FF0000', '2': '#00FF00' }, '1': { '3': '#0000FF' } },
        }),
      ),
    ).toBe('Edit chart demo-chart-sheet-1-1: 3 point colors')
  })

  it('keeps a stable part order when several new fields are set together', () => {
    expect(
      label(
        editChart({
          gridlines: false,
          dataLabelPosition: 'outside-end',
          dataLabelFormat: '#,##0',
          valueAxis: { min: 0 },
          pointColors: { '0': { '1': '#FF0000' } },
        }),
      ),
    ).toBe(
      'Edit chart demo-chart-sheet-1-1: 1 point color, label position outside-end, ' +
        'label format "#,##0", gridlines off, value axis min 0',
    )
  })
})

describe('edit_chart: old fields are unchanged', () => {
  it('still parses and passes through the pre-batch fields', () => {
    const operation: EditChartOperation = {
      op: 'edit_chart',
      chartPath: 'xl/charts/chart1.xml',
      title: 'Revenue by region',
      chartType: 'bar',
      seriesColors: { '0': '#FF0000' },
      legend: 'bottom',
      dataLabels: 'value',
      grouping: 'stacked',
      axisTitles: { category: 'Region', value: null },
      seriesData: [{ index: 0, name: 'Revenue', valuesRange: 'B2:B5' }],
    }
    expect(parseOne(operation).success).toBe(true)
    expect(expandOne(operation)).toEqual([operation])
    expect(layoutOpLabel(operation)).toBe(
      'Edit chart xl/charts/chart1.xml: title "Revenue by region", type bar, ' +
        'series colors #0→#FF0000, legend bottom, category axis "Region", ' +
        'clear value-axis title, series #0 name "Revenue" values B2:B5',
    )
  })

  it('keeps the legacy single-field labels stable', () => {
    expect(layoutOpLabel(editChart({ title: 'Sales' }))).toBe(
      'Edit chart demo-chart-sheet-1-1: title "Sales"',
    )
    expect(layoutOpLabel(editChart({ legend: 'none' }))).toBe(
      'Edit chart demo-chart-sheet-1-1: hide legend',
    )
    expect(layoutOpLabel(editChart({ seriesColors: { '0': '#FF0000', '1': '#00FF00' } }))).toBe(
      'Edit chart demo-chart-sheet-1-1: series colors #0→#FF0000, #1→#00FF00',
    )
  })

  it('still rejects an entry that names neither a name nor a data range', () => {
    // the schema only requires an index; the "does something" rule lives at
    // expansion (a union member cannot carry a refinement)
    expect(
      parseOne(editChart({ seriesData: [{ index: 0, name: 'x' }, { index: 1 }] })).success,
    ).toBe(true)
    expect(() =>
      expandOne(editChart({ seriesData: [{ index: 0, name: 'x' }, { index: 1 }] })),
    ).toThrow(/seriesData\[index=1\] needs a name or a data range\./)
    expect(parseOne(editChart({ seriesData: [{ index: 0, valuesRange: 'B2:B5' }] })).success).toBe(
      true,
    )
    // the schema still bounds the series index
    expect(parseOne(editChart({ seriesData: [{ index: 256, name: 'x' }] })).success).toBe(false)
  })
})

describe('edit_chart: demo chart application', () => {
  function chartSnapshot(): WorkbookSnapshot {
    return {
      revision: 0,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: {
            A1: { value: 'Region' },
            B1: { value: 'Revenue' },
            A2: { value: 'North' },
            B2: { value: 10 },
            A3: { value: 'South' },
            B3: { value: 20 },
          },
          visuals: [
            buildChartVisual({
              id: 'demo-chart-sheet-1-1',
              sheetId: 'sheet-1',
              sheetName: 'Sheet1',
              chartType: 'column',
              dataRange: 'A1:B3',
              title: 'Revenue',
              values: [
                ['Region', 'Revenue'],
                ['North', 10],
                ['South', 20],
              ],
            }),
          ],
        },
      ],
    }
  }

  function applyEdit(overrides: Partial<EditChartOperation>) {
    const adapter = new InMemoryWorkbookAdapter(chartSnapshot())
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'chart-1',
      baseRevision: 0,
      summary: 'Edit demo chart',
      operations: [editChart(overrides)],
    })
    adapter.apply(plan)
    return { adapter, plan, chart: adapter.findVisual('demo-chart-sheet-1-1')?.chart }
  }

  it('keeps applying the pre-batch chart properties', () => {
    const { chart } = applyEdit({ title: 'Renamed', legend: 'bottom', gridlines: undefined })
    expect(chart?.title).toBe('Renamed')
    expect(chart?.legend).toBe('bottom')
  })

  it('demo adapter forwards the five new chart fields', () => {
    // The demo adapter maps the same ChartStateEdit fields as the file-chart
    // path (visual-actions.buildAiChartEdit).
    const { plan, chart } = applyEdit({
      gridlines: false,
      dataLabelPosition: 'outside-end',
      dataLabelFormat: '#,##0',
      valueAxis: { min: 0 },
      pointColors: { '0': { '1': '#FF0000' } },
    })
    expect(plan.structuralChanges.map((change) => change.label)).toEqual([
      'Edit chart demo-chart-sheet-1-1: 1 point color, label position outside-end, ' +
        'label format "#,##0", gridlines off, value axis min 0',
    ])
    expect(plan.warnings).toEqual([])
    expect(chart?.gridlines).toBe(false)
    expect(chart?.dataLabelPosition).toBe('outside-end')
    expect(chart?.dataLabelFormat).toBe('#,##0')
    expect(chart?.valueAxis).toEqual({ min: 0 })
    expect(chart?.series[0]?.pointColors).toEqual([{ index: 1, color: '#FF0000' }])
  })
})
