/**
 * Data → Subtotal (and the AI `subtotal` op): insert one SUBTOTAL row after
 * each run of equal group values, plus a grand-total row at the bottom.
 * Shared by the ribbon dialog (current selection) and the plan executor
 * (explicit sheet + range). The row inserts are ordinary Univer mutations, so
 * the edit journal and the undo stack capture them exactly like a paste; this
 * module only decides where the rows go and what they contain.
 */
import { columnIndex, columnLabel, parseRange } from '../domain/cell-address'
import { isSheetRemoved, journalCellContentAt } from './edit-journal'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/** SUBTOTAL's `function_num` values (Excel's Data → Subtotal list). */
export type SubtotalFunction =
  | 'sum'
  | 'average'
  | 'count'
  | 'counta'
  | 'max'
  | 'min'
  | 'product'
  | 'stdev'
  | 'stdevp'
  | 'var'
  | 'varp'

/// Excel's SUBTOTAL ignores rows that are themselves SUBTOTALs, which is what
/// makes the grand total below safe over the per-group rows.
const SUBTOTAL_FUNCTION_NUMBERS: Record<SubtotalFunction, number> = {
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

/** Excel's own ceiling for this dialog; more groups than this is a mistake. */
export const MAX_SUBTOTAL_GROUPS = 200

/** One subtotal edit: which sheet/range, grouped and summed by which columns. */
export interface SubtotalTarget {
  readonly sheetId: string
  readonly range: string
  /** Column label the rows are grouped by (first row of the range is the header). */
  readonly groupColumn: string
  /** Column label that receives the SUBTOTAL formulas. */
  readonly valueColumn: string
  readonly function: SubtotalFunction
}

export type SubtotalFailure =
  | 'sheet-missing'
  | 'range-too-small'
  | 'no-data-rows'
  | 'same-column'
  | 'column-outside-range'
  | 'too-many-groups'

export interface SubtotalOutcome {
  readonly applied: boolean
  /** Number of group rows inserted (excludes the grand-total row). */
  readonly groups?: number
  readonly reason?: SubtotalFailure
}

/** English text for the executor's error path (the ribbon localizes instead). */
export function subtotalFailureMessage(reason: SubtotalFailure, groups?: number): string {
  switch (reason) {
    case 'sheet-missing':
      return 'The subtotal sheet is no longer available.'
    case 'range-too-small':
      return 'subtotal needs a range with a header row and at least one data row.'
    case 'no-data-rows':
      return 'subtotal found no data rows below the header row.'
    case 'same-column':
      return 'subtotal needs a groupColumn different from the valueColumn.'
    case 'column-outside-range':
      return 'subtotal groupColumn and valueColumn must both lie inside the range.'
    case 'too-many-groups':
      return groups === undefined
        ? `subtotal would create more than ${MAX_SUBTOTAL_GROUPS} groups.`
        : `subtotal would create ${groups.toLocaleString('en-US')} groups — the limit is ` +
            `${MAX_SUBTOTAL_GROUPS}. Sort the data by the group column first.`
  }
}

/**
 * Applies one subtotal to an explicit sheet/range. `lazyState` is null only for
 * the blank/demo workbook (no edit journal); the AI op always runs with one.
 * Validation failures come back as `reason` codes; a rejected Univer write
 * throws (the caller reports it like any other failed mutation).
 */
export function applySubtotalToRange(
  runtime: UniverRuntime,
  lazyState: LazyWorkbookState | null,
  target: SubtotalTarget,
): SubtotalOutcome {
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(target.sheetId)
  if (!worksheet || (lazyState && isSheetRemoved(lazyState.editJournal, target.sheetId))) {
    return { applied: false, reason: 'sheet-missing' }
  }
  const bounds = parseRange(target.range)
  const rowCount = bounds.endRow - bounds.startRow + 1
  if (rowCount < 2) return { applied: false, reason: 'range-too-small' }
  const groupColumn = columnIndex(target.groupColumn)
  const valueColumn = columnIndex(target.valueColumn)
  const insideRange = (column: number): boolean =>
    column >= bounds.startColumn && column <= bounds.endColumn
  if (!insideRange(groupColumn) || !insideRange(valueColumn)) {
    return { applied: false, reason: 'column-outside-range' }
  }
  if (groupColumn === valueColumn) return { applied: false, reason: 'same-column' }

  // Only the group column is read — the value column just receives formulas.
  // Journal-owned cells win over the grid: a lazy viewport may have been
  // evicted (or never loaded) since the edit landed in the journal.
  const bodyRows = rowCount - 1
  if (bodyRows === 0) return { applied: false, reason: 'no-data-rows' }
  const gridLabels = worksheet
    .getRange(bounds.startRow + 1, groupColumn, bodyRows, 1)
    .getValues() as (string | number | boolean | null)[][]
  const labels: string[] = []
  for (let offset = 0; offset < bodyRows; offset += 1) {
    const journalCell = lazyState
      ? journalCellContentAt(
          lazyState.editJournal,
          target.sheetId,
          bounds.startRow + 1 + offset,
          groupColumn,
        )
      : ({ found: false } as const)
    // A journaled formula's displayed value is unknown to the journal; the
    // grid still holds its computed value, so only literal edits overlay.
    const raw =
      journalCell.found && journalCell.formula === null
        ? journalCell.value
        : (gridLabels[offset]?.[0] ?? null)
    labels.push(String(raw ?? ''))
  }

  // Group by contiguous runs of equal stored values — Excel's "at each change
  // in" semantics, which is why sorting by the group column matters.
  const groups: { label: string; startAbs: number; endAbs: number }[] = []
  let currentLabel = labels[0] ?? ''
  let groupStart = bounds.startRow + 1
  for (let offset = 1; offset <= bodyRows; offset += 1) {
    const label = offset < bodyRows ? (labels[offset] ?? '') : null
    if (label !== currentLabel) {
      groups.push({ label: currentLabel, startAbs: groupStart, endAbs: bounds.startRow + offset })
      if (label !== null) {
        currentLabel = label
        groupStart = bounds.startRow + 1 + offset
      }
    }
  }
  if (groups.length === 0) return { applied: false, reason: 'no-data-rows' }
  if (groups.length > MAX_SUBTOTAL_GROUPS) {
    return { applied: false, reason: 'too-many-groups', groups: groups.length }
  }

  const fn = SUBTOTAL_FUNCTION_NUMBERS[target.function]
  const valueLetter = columnLabel(valueColumn)
  const writeTotalRow = (row: number, label: string, from: number, to: number): void => {
    const labelCell = worksheet.getRange(row, groupColumn, 1, 1)
    labelCell.setValue({ v: label })
    labelCell.setFontWeight('bold')
    const valueCell = worksheet.getRange(row, valueColumn, 1, 1)
    valueCell.setValue({
      f: `=SUBTOTAL(${fn},$${valueLetter}$${from + 1}:$${valueLetter}$${to + 1})`,
    })
    valueCell.setFontWeight('bold')
  }
  // Bottom-up so earlier group indices stay valid across the inserts.
  for (const group of [...groups].reverse()) {
    worksheet.insertRowsBefore(group.endAbs + 1, 1)
    writeTotalRow(group.endAbs + 1, `${group.label} Total`, group.startAbs, group.endAbs)
  }
  const lastRowAbs = bounds.startRow + bodyRows + groups.length
  worksheet.insertRowsBefore(lastRowAbs + 1, 1)
  // SUBTOTAL skips the nested per-group subtotal rows inside the span.
  writeTotalRow(lastRowAbs + 1, 'Grand Total', bounds.startRow + 1, lastRowAbs)
  return { applied: true, groups: groups.length }
}
