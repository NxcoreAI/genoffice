import { describe, expect, it } from 'vitest'

import {
  expandToPrimitiveOps,
  formatOpLabel,
  isLayoutOp,
  isStructuralOp,
  layoutOpLabel,
  MAX_DATA_TOOL_CELLS,
  workbookCommandBatchSchema,
  workbookOperationSchema,
  type CellFormatPatch,
  type FormatRangeOperation,
  type WorkbookOperation,
} from '../src/domain/workbook-dsl'
import { InMemoryWorkbookAdapter, WorkbookConflictError } from '../src/domain/in-memory-workbook'
import type { CellState, WorkbookSnapshot, WorksheetState } from '../src/domain/workbook.types'

// ---------------------------------------------------------------------------
// Pure harness: batch-2 ops expand against a plain cell map (no Electron /
// Univer), exactly like the in-memory adapter and the renderer plan paths do.
// ---------------------------------------------------------------------------

type CellMap = Record<string, CellState>

function makeReader(bySheet: Record<string, CellMap>) {
  return (address: string, sheetId: string): CellState =>
    bySheet[sheetId]?.[address] ?? { value: null }
}

function makeSheetNames(names: Record<string, string>) {
  return (sheetId: string): string => names[sheetId] ?? sheetId
}

function expandOne(operation: WorkbookOperation, cells: Record<string, CellMap> = {}) {
  return expandToPrimitiveOps([operation], makeReader(cells), makeSheetNames({}))
}

function parseOne(input: unknown) {
  return workbookOperationSchema.safeParse(input)
}

function batchOf(operations: unknown[]) {
  return workbookCommandBatchSchema.safeParse({
    dslVersion: 1,
    transactionId: 'tx-1',
    baseRevision: 0,
    summary: 'batch-2 test',
    operations,
  })
}

function formatRange(range: string, format: CellFormatPatch): FormatRangeOperation {
  return { op: 'format_range', sheetId: 's1', range, format }
}

/** ST_BorderStyle: the 13 OOXML line styles the save path also accepts. */
const BORDER_STYLES = [
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
  'hair',
  'dashDot',
  'dashDotDot',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
] as const

/** what clear_formats expands into — every format field in its cleared state */
const CLEARED_PATCH = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  fontFamily: null,
  fontSize: null,
  fontColor: null,
  fillColor: null,
  numberFormat: null,
  horizontalAlign: null,
  verticalAlign: null,
  wrapText: null,
  textRotation: null,
  indent: 0,
  border: { type: 'none' },
  locked: null,
  hidden: null,
}

describe('format_range: border styles', () => {
  it('accepts exactly the 13 OOXML line styles and rejects anything else', () => {
    expect(BORDER_STYLES).toHaveLength(13)
    for (const style of BORDER_STYLES) {
      const parsed = parseOne({
        op: 'format_range',
        sheetId: 's1',
        range: 'A1:B2',
        format: { border: { type: 'all', style } },
      })
      expect(parsed.success, `border style ${style} should parse`).toBe(true)
    }
    for (const bad of ['wavy', 'None', 'THIN', 'medium-dashed']) {
      expect(
        parseOne({
          op: 'format_range',
          sheetId: 's1',
          range: 'A1:B2',
          format: { border: { type: 'all', style: bad } },
        }).success,
        `border style ${bad} should be rejected`,
      ).toBe(false)
    }
  })

  it('parses border edge type + optional color and keeps border null invalid', () => {
    expect(
      parseOne({
        op: 'format_range',
        sheetId: 's1',
        range: 'A1',
        format: { border: { type: 'bottom', style: 'double', color: '#00FF00' } },
      }).success,
    ).toBe(true)
    // clearing is border {type:'none'}, not null
    expect(
      parseOne({ op: 'format_range', sheetId: 's1', range: 'A1', format: { border: null } })
        .success,
    ).toBe(false)
    expect(
      parseOne({
        op: 'format_range',
        sheetId: 's1',
        range: 'A1',
        format: { border: { type: 'diagonal' } },
      }).success,
    ).toBe(false)
  })

  it('labels a bordered format with its edge, line style, and color', () => {
    expect(
      formatOpLabel(
        formatRange('A1', { border: { type: 'all', style: 'dashed', color: '#000000' } }),
      ),
    ).toBe('A1: border all dashed #000000')
    expect(formatOpLabel(formatRange('A1', { border: { type: 'all', style: 'thick' } }))).toBe(
      'A1: border all thick',
    )
    expect(formatOpLabel(formatRange('A1', { border: { type: 'top' } }))).toBe('A1: border top')
    expect(formatOpLabel(formatRange('A1', { border: { type: 'none' } }))).toBe('A1: clear borders')
  })
})

describe('format_range: cell protection flags', () => {
  it('parses locked/hidden as booleans or null', () => {
    for (const format of [
      { locked: true },
      { locked: false },
      { locked: null },
      { hidden: true },
      { hidden: false },
      { hidden: null },
      { locked: false, hidden: true },
    ]) {
      expect(parseOne({ op: 'format_range', sheetId: 's1', range: 'A1', format }).success).toBe(
        true,
      )
    }
    expect(
      parseOne({ op: 'format_range', sheetId: 's1', range: 'A1', format: { locked: 'yes' } })
        .success,
    ).toBe(false)
  })

  it('labels unlocked / locked / reset and the hidden variants', () => {
    expect(formatOpLabel(formatRange('A1', { locked: false }))).toBe('A1: unlocked')
    expect(formatOpLabel(formatRange('A1', { locked: true }))).toBe('A1: locked')
    expect(formatOpLabel(formatRange('A1', { locked: null }))).toBe('A1: reset locked')
    expect(formatOpLabel(formatRange('A1', { hidden: true }))).toBe('A1: hidden')
    expect(formatOpLabel(formatRange('A1', { hidden: false }))).toBe('A1: not hidden')
    expect(formatOpLabel(formatRange('A1', { hidden: null }))).toBe('A1: reset hidden')
    expect(formatOpLabel(formatRange('A1', { bold: true, locked: false }))).toBe(
      'A1: bold, unlocked',
    )
  })
})

describe('format_range: cleared-patch collapse', () => {
  it('collapses the full clear patch into one "Clear formatting" line', () => {
    expect(formatOpLabel(formatRange('A1:B2', CLEARED_PATCH as CellFormatPatch))).toBe(
      'Clear formatting on A1:B2',
    )
  })

  it('does not collapse a near-miss patch (missing one field)', () => {
    const { hidden: _hidden, ...partial } = CLEARED_PATCH
    const label = formatOpLabel(formatRange('A1', partial as CellFormatPatch))
    expect(label.startsWith('A1: ')).toBe(true)
    expect(label).not.toContain('Clear formatting')
    // the individual clearing fragments still read sensibly
    expect(label).toContain('no bold')
    expect(label).toContain('clear font')
    expect(label).toContain('indent 0')
    expect(label).toContain('clear borders')
    expect(label).toContain('reset locked')
  })
})

describe('clear_formats', () => {
  it('expands to exactly one format_range carrying the full cleared set', () => {
    const ops = expandOne({ op: 'clear_formats', sheetId: 's1', range: 'A1:B2' })
    expect(ops).toHaveLength(1)
    expect(ops[0]).toEqual({
      op: 'format_range',
      sheetId: 's1',
      range: 'A1:B2',
      format: CLEARED_PATCH,
    })
    expect(ops.every((op) => op.op !== 'clear_cell' && op.op !== 'set_cell')).toBe(true)
  })

  it('is content-classed: batches with cell edits, never with structural ops', () => {
    const clear = { op: 'clear_formats', sheetId: 's1', range: 'A1:B2' }

    expect(isStructuralOp(clear as unknown as WorkbookOperation)).toBe(false)
    expect(isLayoutOp(clear as unknown as WorkbookOperation)).toBe(false)

    expect(
      batchOf([clear, { op: 'set_cell', sheetId: 's1', address: 'A1', value: 'x' }]).success,
    ).toBe(true)

    const mixed = batchOf([clear, { op: 'insert_rows', sheetId: 's1', row: 2, count: 1 }])
    expect(mixed.success).toBe(false)
    expect(JSON.stringify(mixed.error?.issues)).toMatch(/separate batches/)

    const deleteMixed = batchOf([
      { op: 'delete_cols', sheetId: 's1', column: 'A', count: 1 },
      clear,
    ])
    expect(deleteMixed.success).toBe(false)
  })

  it('enforces the data-tool cell budget', () => {
    expect(MAX_DATA_TOOL_CELLS).toBe(20_000)
    expect(() => expandOne({ op: 'clear_formats', sheetId: 's1', range: 'A1:A20001' })).toThrow(
      /clear_formats covers 20,001 cells — the limit is 20,000\. Clear it in several ranges\./,
    )
    // exactly at the limit is fine
    expect(expandOne({ op: 'clear_formats', sheetId: 's1', range: 'A1:A20000' })).toHaveLength(1)
  })

  it('rejects a bad range before anything else', () => {
    expect(parseOne({ op: 'clear_formats', sheetId: 's1', range: 'A1:' }).success).toBe(false)
  })
})

describe('set_theme', () => {
  it('accepts an empty op at the schema level but rejects it at expansion', () => {
    expect(parseOne({ op: 'set_theme' }).success).toBe(true)
    expect(() => expandOne({ op: 'set_theme' })).toThrow(
      /set_theme needs at least one of "theme".*"colors".*"fonts".*read_workbook_context/s,
    )
  })

  it('passes through unchanged for one, two, or three preset ids', () => {
    for (const operation of [
      { op: 'set_theme', theme: 'office' },
      { op: 'set_theme', theme: 'office', colors: 'warm' },
      { op: 'set_theme', theme: 'office', colors: 'warm', fonts: 'classic' },
      { op: 'set_theme', fonts: 'classic' },
    ] as const) {
      expect(expandOne(operation)).toEqual([operation])
    }
  })

  it('labels the applied presets', () => {
    expect(layoutOpLabel({ op: 'set_theme', theme: 'office' })).toBe('Apply theme office')
    expect(layoutOpLabel({ op: 'set_theme', colors: 'warm' })).toBe('Apply theme colors warm')
    expect(layoutOpLabel({ op: 'set_theme', fonts: 'classic' })).toBe('Apply theme fonts classic')
    expect(
      layoutOpLabel({ op: 'set_theme', theme: 'office', colors: 'warm', fonts: 'classic' }),
    ).toBe('Apply theme office, colors warm, fonts classic')
  })

  it('bounds preset ids at 1..64 characters', () => {
    expect(parseOne({ op: 'set_theme', theme: '' }).success).toBe(false)
    expect(parseOne({ op: 'set_theme', theme: 'a'.repeat(64) }).success).toBe(true)
    expect(parseOne({ op: 'set_theme', theme: 'a'.repeat(65) }).success).toBe(false)
  })
})

describe('protect_workbook', () => {
  it('parses, passes through, and labels both directions', () => {
    expect(parseOne({ op: 'protect_workbook', protected: true }).success).toBe(true)
    expect(parseOne({ op: 'protect_workbook', protected: false }).success).toBe(true)
    expect(parseOne({ op: 'protect_workbook' }).success).toBe(false)

    expect(expandOne({ op: 'protect_workbook', protected: true })).toEqual([
      { op: 'protect_workbook', protected: true },
    ])
    expect(layoutOpLabel({ op: 'protect_workbook', protected: true })).toBe(
      'Protect workbook structure',
    )
    expect(layoutOpLabel({ op: 'protect_workbook', protected: false })).toBe(
      'Unprotect workbook structure',
    )
  })
})

describe('allow_edit_ranges', () => {
  const ranges = (count: number) =>
    Array.from({ length: count }, (_, index) => ({ sqref: `A${index + 1}:B${index + 1}` }))

  it('accepts 0..100 ranges and clears with an empty list', () => {
    expect(parseOne({ op: 'allow_edit_ranges', sheetId: 's1', ranges: [] }).success).toBe(true)
    expect(parseOne({ op: 'allow_edit_ranges', sheetId: 's1', ranges: ranges(1) }).success).toBe(
      true,
    )
    expect(parseOne({ op: 'allow_edit_ranges', sheetId: 's1', ranges: ranges(100) }).success).toBe(
      true,
    )
    expect(parseOne({ op: 'allow_edit_ranges', sheetId: 's1', ranges: ranges(101) }).success).toBe(
      false,
    )
    // a missing ranges list is not "clear", it is invalid
    expect(parseOne({ op: 'allow_edit_ranges', sheetId: 's1' }).success).toBe(false)
    expect(
      parseOne({
        op: 'allow_edit_ranges',
        sheetId: 's1',
        ranges: [{ name: 'Input', sqref: 'not-a-range' }],
      }).success,
    ).toBe(false)
  })

  it('passes through unchanged', () => {
    const operation: WorkbookOperation = {
      op: 'allow_edit_ranges',
      sheetId: 's1',
      ranges: [{ name: 'Input', sqref: 'B2:D10' }, { sqref: 'F1:F9' }],
    } as const
    expect(expandOne(operation)).toEqual([operation])
  })

  it('labels empty, singular, plural, and long lists (first three + ellipsis)', () => {
    expect(layoutOpLabel({ op: 'allow_edit_ranges', sheetId: 's1', ranges: [] })).toBe(
      'Clear allowed edit ranges',
    )
    expect(
      layoutOpLabel({ op: 'allow_edit_ranges', sheetId: 's1', ranges: [{ sqref: 'B2:D10' }] }),
    ).toBe('Allow editing 1 range: B2:D10')
    expect(
      layoutOpLabel({
        op: 'allow_edit_ranges',
        sheetId: 's1',
        ranges: [{ name: 'Input', sqref: 'B2:D10' }, { sqref: 'F1:F9' }],
      }),
    ).toBe('Allow editing 2 ranges: Input, F1:F9')
    expect(
      layoutOpLabel({
        op: 'allow_edit_ranges',
        sheetId: 's1',
        ranges: [{ sqref: 'A1' }, { name: 'B2', sqref: 'B2' }, { sqref: 'C3' }, { sqref: 'D4' }],
      }),
    ).toBe('Allow editing 4 ranges: A1, B2, C3, …')
  })
})

describe('demo in-memory adapter: format + file-level settings', () => {
  function demoSnapshot(): WorkbookSnapshot {
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
            B2: { value: 128000 },
          },
          styles: {
            A1: { bold: true, fillColor: '#FF0000' },
            B2: { italic: true },
          },
        },
      ],
    }
  }

  /** styles are typed with the demo's CellFormatState; batch-2 adds flags the
   * type does not declare, so read them through a plain record */
  function styleAt(adapter: InMemoryWorkbookAdapter, address: string) {
    const sheet: WorksheetState | undefined = adapter.getSnapshot().sheets[0]
    const styles = sheet?.styles as unknown as Record<string, Record<string, unknown>> | undefined
    return styles?.[address]
  }

  it('applies locked:false plus a thick border, leaving values untouched', () => {
    const adapter = new InMemoryWorkbookAdapter(demoSnapshot())
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'fmt-1',
      baseRevision: 0,
      summary: 'Unlock A1 and add a thick border',
      operations: [
        {
          op: 'format_range',
          sheetId: 'sheet-1',
          range: 'A1',
          format: { locked: false, border: { type: 'all', style: 'thick' } },
        },
      ],
    })
    // NB: the op travels through the batch schema, so the label follows the
    // schema's field order (border before locked), not the literal's order
    expect(plan.formatChanges.map((change) => change.label)).toEqual([
      'A1: border all thick, unlocked',
    ])
    // dry run leaves the grid alone
    expect(styleAt(adapter, 'A1')?.locked).toBeUndefined()

    adapter.apply(plan)
    expect(styleAt(adapter, 'A1')).toEqual({
      bold: true,
      fillColor: '#FF0000',
      locked: false,
      border: { type: 'all', style: 'thick' },
    })
    expect(adapter.getSnapshot().sheets[0]?.cells.A1?.value).toBe('Region')
    expect(adapter.getSnapshot().sheets[0]?.cells.B2?.value).toBe(128000)
  })

  it('clear_formats clears styles through the plan/apply path and keeps cell values', () => {
    const adapter = new InMemoryWorkbookAdapter(demoSnapshot())
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'clear-1',
      baseRevision: 0,
      summary: 'Clear formatting on A1:B2',
      operations: [{ op: 'clear_formats', sheetId: 'sheet-1', range: 'A1:B2' }],
    })
    expect(plan.formatChanges).toHaveLength(1)
    expect(plan.formatChanges[0]?.label).toBe('Clear formatting on A1:B2')
    expect(plan.cellChanges).toEqual([])
    expect(styleAt(adapter, 'A1')).toEqual({ bold: true, fillColor: '#FF0000' })

    adapter.apply(plan)
    // every cleared value replaces the old styling: the falsy defaults and the
    // border 'none' stay in the demo style record, the nulls drop out of it
    expect(styleAt(adapter, 'A1')).toEqual({
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      indent: 0,
      border: { type: 'none' },
    })
    expect(styleAt(adapter, 'B2')).toEqual({
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      indent: 0,
      border: { type: 'none' },
    })
    // content survives
    expect(adapter.getSnapshot().sheets[0]?.cells.A1?.value).toBe('Region')
    expect(adapter.getSnapshot().sheets[0]?.cells.B2?.value).toBe(128000)
  })

  it('rejects the three file-level settings with the imported-xlsx error', () => {
    const cases = [
      { op: 'set_theme', theme: 'office' },
      { op: 'protect_workbook', protected: true },
      { op: 'allow_edit_ranges', sheetId: 'sheet-1', ranges: [{ sqref: 'A1:B2' }] },
    ]
    for (const operation of cases) {
      const adapter = new InMemoryWorkbookAdapter(demoSnapshot())
      expect(
        () =>
          adapter.plan({
            dslVersion: 1,
            transactionId: `tx-${operation.op}`,
            baseRevision: 0,
            summary: `Try ${operation.op}`,
            operations: [operation],
          }),
        `${operation.op} should fail on the demo workbook`,
      ).toThrow(WorkbookConflictError)
      expect(() =>
        adapter.plan({
          dslVersion: 1,
          transactionId: `tx-${operation.op}-2`,
          baseRevision: 0,
          summary: `Try ${operation.op}`,
          operations: [operation],
        }),
      ).toThrow(new RegExp(`${operation.op.replace('.', '\\.')} needs an imported xlsx file`))
    }
  })
})

describe('existing format_range behavior is unchanged', () => {
  it('passes a patch without the new fields through as-is and labels it like before', () => {
    const operation = formatRange('B2:C3', { bold: true, fillColor: '#FFCC00' })
    const expanded = expandOne(operation)
    expect(expanded).toEqual([operation])
    expect(expanded[0]).toEqual({
      op: 'format_range',
      sheetId: 's1',
      range: 'B2:C3',
      format: { bold: true, fillColor: '#FFCC00' },
    })
    expect(formatOpLabel(expanded[0] as FormatRangeOperation)).toBe('B2:C3: bold, fill #FFCC00')
  })

  it('keeps number formats, alignment, rotation, and indent labels stable', () => {
    expect(
      formatOpLabel(
        formatRange('A1', {
          numberFormat: '#,##0.00',
          horizontalAlign: 'center',
          indent: 2,
          wrapText: true,
        }),
      ),
    ).toBe('A1: number format #,##0.00, align center, indent 2, wrap')
    expect(formatOpLabel(formatRange('A1', { textRotation: 'vertical' }))).toBe(
      'A1: rotation vertical',
    )
    expect(formatOpLabel(formatRange('A1', { fontSize: null }))).toBe('A1: clear font size')
  })

  it('still rejects an empty patch and the 200k range cap', () => {
    expect(parseOne({ op: 'format_range', sheetId: 's1', range: 'A1', format: {} }).success).toBe(
      false,
    )
    expect(() => expandOne(formatRange('A1:A200001', { bold: true }))).toThrow(
      /format_range covers more than 200,000 cells/,
    )
  })
})
