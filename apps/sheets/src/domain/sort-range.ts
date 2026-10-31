import { columnIndex, formatAddress, parseRange } from './cell-address'
import type { CellScalar, CellState } from './workbook.types'

/**
 * Computes the cell rewrites a sort produces, so sorting rides the existing
 * per-cell preview / CAS / apply machinery in both demo and lazy mode.
 * Regions containing formulas are rejected: moving formula text between rows
 * silently re-targets relative references, which is exactly the class of
 * damage the preview model exists to prevent.
 */

export interface SortComputedChange {
  readonly address: string
  readonly before: CellScalar
  readonly after: CellScalar
}

/** One sort key: the key column, and whether that key sorts ascending. */
export interface SortKey {
  readonly byColumn: string
  readonly ascending: boolean
}

export interface SortSpec {
  readonly range: string
  /** legacy single-key form; ignored when `keys` is present */
  readonly byColumn?: string
  readonly ascending?: boolean
  /** 1–3 keys applied in order — later keys break ties of earlier ones */
  readonly keys?: readonly SortKey[]
  readonly hasHeader: boolean
}

/// blanks always sort last; numbers before booleans before text (Excel order)
function compareScalars(a: CellScalar, b: CellScalar): number {
  if (a === null && b === null) return 0
  if (a === null) return 1
  if (b === null) return -1
  const rank = (value: Exclude<CellScalar, null>): number =>
    typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2
  const rankA = rank(a)
  const rankB = rank(b)
  if (rankA !== rankB) return rankA - rankB
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b)
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' })
}

/**
 * Normalizes `keys` vs the legacy byColumn/order pair into one key list.
 * Throws an actionable error when neither form is complete: sort_range is a
 * discriminated-union schema member, so this cross-field rule lives here
 * rather than as a zod refinement.
 */
function resolveSortKeys(spec: SortSpec): readonly SortKey[] {
  const keys =
    spec.keys ??
    (spec.byColumn !== undefined && spec.ascending !== undefined
      ? [{ byColumn: spec.byColumn, ascending: spec.ascending }]
      : null)
  if (keys === null || keys.length === 0) {
    throw new Error(
      'sort_range needs a sort key: pass "keys" (1–3 entries, e.g. ' +
        '[{"byColumn":"B","order":"desc"}]) or both "byColumn" and "order".',
    )
  }
  if (keys.length > 3) {
    throw new Error('sort_range accepts at most 3 sort keys — sort by fewer columns.')
  }
  return keys
}

export function computeSortChanges(
  spec: SortSpec,
  readCell: (address: string) => CellState,
): SortComputedChange[] {
  const bounds = parseRange(spec.range)
  const keys = resolveSortKeys(spec)
  // Validate every key column up front so the error names the offending key
  // rather than surfacing halfway through the read.
  const keyColumns = keys.map((key) => {
    const keyColumn = columnIndex(key.byColumn)
    if (keyColumn < bounds.startColumn || keyColumn > bounds.endColumn) {
      throw new Error(
        `Sort key column ${key.byColumn} is outside the range ${spec.range} — every sort key must be one of the range's columns.`,
      )
    }
    return keyColumn - bounds.startColumn
  })
  const firstDataRow = bounds.startRow + (spec.hasHeader ? 1 : 0)
  if (firstDataRow >= bounds.endRow) {
    throw new Error('The sort range needs at least two data rows.')
  }

  const rows: { keys: CellScalar[]; cells: CellScalar[] }[] = []
  for (let row = firstDataRow; row <= bounds.endRow; row += 1) {
    const cells: CellScalar[] = []
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      const state = readCell(formatAddress(row, column))
      if (state.formula) {
        throw new Error(
          `The sort range contains a formula at ${formatAddress(row, column)} — sorting would silently re-target its references. Sort values only, or convert formulas to values first.`,
        )
      }
      cells.push(state.value)
    }
    rows.push({ keys: keyColumns.map((offset) => cells[offset] ?? null), cells })
  }

  // Multi-key compare: the first key whose values differ decides the order.
  // For every key, blanks stay last in both directions (Excel behavior).
  const compareRows = (a: { keys: CellScalar[] }, b: { keys: CellScalar[] }): number => {
    for (let index = 0; index < keys.length; index += 1) {
      const cmp = compareScalars(a.keys[index] ?? null, b.keys[index] ?? null)
      if (cmp === 0) continue
      const key = keys[index]
      return a.keys[index] === null || b.keys[index] === null || key === undefined
        ? cmp
        : key.ascending
          ? cmp
          : -cmp
    }
    return 0
  }

  // Stable sort (the index tiebreak keeps equal rows in input order).
  const sorted = rows
    .map((row, index) => ({ row, index }))
    .sort((a, b) => {
      const cmp = compareRows(a.row, b.row)
      return cmp !== 0 ? cmp : a.index - b.index
    })
    .map((entry) => entry.row)

  const changes: SortComputedChange[] = []
  sorted.forEach((row, offset) => {
    const targetRow = firstDataRow + offset
    row.cells.forEach((value, columnOffset) => {
      const address = formatAddress(targetRow, bounds.startColumn + columnOffset)
      const before = rows[offset]?.cells[columnOffset] ?? null
      if (before !== value) changes.push({ address, before, after: value })
    })
  })
  return changes
}
