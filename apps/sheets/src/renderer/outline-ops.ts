/**
 * Row/column outline editing shared by the ribbon (Data → Group / Subtotal
 * outline buttons) and the AI executor (the `outline.*` layout operations).
 * Level edits journal directly — Univer has no outline model, so there is no
 * command to undo; Hide/Show Detail rides the normal hidden-rows pipeline
 * plus a collapsed flag on the summary line.
 */
import { columnIndex, columnLabel } from '../domain/cell-address'
import { isSheetRemoved, recordStructuralOp } from './edit-journal'
import { sheetOutline } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

export type OutlineAxis = 'rows' | 'cols'
export type OutlineAction = 'group' | 'ungroup' | 'hide-detail' | 'show-detail'

/** One outline edit: which sheet/axis/span, and what to do with it. */
export interface OutlineTarget {
  readonly sheetId: string
  readonly axis: OutlineAxis
  /** Row span "5:12" for rows, column span "B:D" for columns. */
  readonly range: string
  readonly action: OutlineAction
}

/** Inclusive 0-based index span. */
export interface OutlineRange {
  readonly start: number
  readonly end: number
}

/**
 * Why an outline edit did nothing, when it did nothing:
 * - max-level: group could not raise any run of the span (all at level 7)
 * - nothing-to-ungroup: every run of the span is already at level 0
 * - unavailable: the sheet is gone (removed this session or not loaded)
 */
export type OutlineSkipReason = 'max-level' | 'nothing-to-ungroup' | 'unavailable'

export interface OutlineOutcome {
  readonly applied: boolean
  readonly reason?: OutlineSkipReason
}

const MAX_ROW_INDEX = 1_048_575
const MAX_COL_INDEX = 16_383

/**
 * Parses an outline span in axis-local form: rows as 1-based row numbers
 * ("5:12", single row "5"), columns as letters ("B:D", single column "B").
 * A1-style tokens ("A5:A12", "B1:D1") are accepted too — the agent
 * occasionally sends them despite the guide — and a reversed span is
 * normalized. Numbers stay 1-based here so error messages read like Excel.
 */
export function parseOutlineRange(range: string, axis: OutlineAxis): OutlineRange {
  const example = axis === 'rows' ? '5:12' : 'B:D'
  const parts = range.trim().split(':')
  if (parts.length > 2 || !parts[0]?.trim()) {
    throw new Error(`Invalid outline range "${range}" — expected "${example}".`)
  }
  const parseOne = (token: string | undefined): number => {
    const trimmed = token?.trim() ?? ''
    if (trimmed === '') {
      throw new Error(`Invalid outline range "${range}" — expected "${example}".`)
    }
    if (axis === 'rows') {
      const match = /(\d+)/.exec(trimmed)
      if (!match?.[1]) {
        throw new Error(`Invalid row range "${range}" — expected row numbers like "${example}".`)
      }
      return Number(match[1])
    }
    const match = /^[^A-Za-z]*([A-Za-z]+)/.exec(trimmed)
    if (!match?.[1]) {
      throw new Error(
        `Invalid column range "${range}" — expected column letters like "${example}".`,
      )
    }
    return columnIndex(match[1].toUpperCase()) + 1
  }
  const first = parseOne(parts[0])
  const second = parts[1] === undefined ? first : parseOne(parts[1])
  const low = Math.min(first, second)
  const high = Math.max(first, second)
  const limit = axis === 'rows' ? MAX_ROW_INDEX + 1 : MAX_COL_INDEX + 1
  if (low < 1 || high > limit) {
    throw new Error(`Outline range "${range}" is outside the sheet (${axis} 1-${limit}).`)
  }
  return { start: low - 1, end: high - 1 }
}

/**
 * Applies one outline edit to an explicit sheet/span. Journaled through the
 * same `set-rows-outline` / `set-cols-outline` ops as the ribbon; the caller
 * owns setPendingEdits and user-facing messages.
 */
export function applyOutlineToRange(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  target: OutlineTarget,
): OutlineOutcome {
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(target.sheetId)
  if (!worksheet || isSheetRemoved(state.editJournal, target.sheetId)) {
    return { applied: false, reason: 'unavailable' }
  }
  const { start, end } = parseOutlineRange(target.range, target.axis)
  const outline = sheetOutline(state, target.sheetId)
  const entries = target.axis === 'rows' ? outline.rows : outline.cols
  const kind =
    target.axis === 'rows' ? ('set-rows-outline' as const) : ('set-cols-outline' as const)

  if (target.action === 'hide-detail' || target.action === 'show-detail') {
    // The target span is the group's detail span; the summary row/column is
    // the next one after it (Excel's summaryBelow/summaryRight default).
    const count = end - start + 1
    if (target.action === 'hide-detail') {
      if (target.axis === 'rows') worksheet.hideRows(start, count)
      else worksheet.hideColumns(start, count)
    } else if (target.axis === 'rows') {
      worksheet.showRows(start, count)
    } else {
      worksheet.showColumns(start, count)
    }
    const summary = end + 1
    const collapsed = target.action === 'hide-detail'
    const summaryLevel = entries.get(summary)?.level ?? 0
    entries.set(summary, { level: summaryLevel, collapsed })
    recordStructuralOp(state.editJournal, target.sheetId, {
      kind,
      start: summary,
      end: summary,
      level: summaryLevel,
      collapsed,
    })
    return { applied: true }
  }

  // Group/Ungroup shifts each contiguous run of equal levels by ±1
  // (levels clamp to 0-7). Runs already at the boundary are skipped.
  const delta = target.action === 'group' ? 1 : -1
  const ops: { start: number; end: number; level: number }[] = []
  let runStart = start
  let runLevel = entries.get(start)?.level ?? 0
  const closeRun = (runEnd: number): void => {
    const level = Math.min(7, Math.max(0, runLevel + delta))
    if (level !== runLevel) ops.push({ start: runStart, end: runEnd, level })
  }
  for (let index = start + 1; index <= end; index += 1) {
    const level = entries.get(index)?.level ?? 0
    if (level !== runLevel) {
      closeRun(index - 1)
      runStart = index
      runLevel = level
    }
  }
  closeRun(end)
  if (ops.length === 0) {
    return {
      applied: false,
      reason: target.action === 'group' ? 'max-level' : 'nothing-to-ungroup',
    }
  }
  for (const op of ops) {
    for (let index = op.start; index <= op.end; index += 1) {
      entries.set(index, {
        level: op.level,
        collapsed: entries.get(index)?.collapsed ?? false,
      })
    }
    recordStructuralOp(state.editJournal, target.sheetId, {
      kind,
      start: op.start,
      end: op.end,
      level: op.level,
    })
  }
  return { applied: true }
}

/** The selection-based ribbon command's span in the executor's range syntax. */
export function outlineRangeForSelection(axis: OutlineAxis, start: number, end: number): string {
  return axis === 'rows' ? `${start + 1}:${end + 1}` : `${columnLabel(start)}:${columnLabel(end)}`
}

/** `outline.hideDetail` → `hide-detail`, etc. Kept here so App and the ribbon agree. */
export function outlineActionFromOp(
  op: 'outline.group' | 'outline.ungroup' | 'outline.hideDetail' | 'outline.showDetail',
): OutlineAction {
  switch (op) {
    case 'outline.group':
      return 'group'
    case 'outline.ungroup':
      return 'ungroup'
    case 'outline.hideDetail':
      return 'hide-detail'
    case 'outline.showDetail':
      return 'show-detail'
  }
}
