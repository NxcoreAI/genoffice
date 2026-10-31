/**
 * Cell navigation (Name Box / Go To / formula / symbol) and data tools
 * (advanced filter, subtotal, consolidate, outline, format-as-table).
 * Extracted from App.tsx; the App component passes a DataToolsContext built
 * fresh per call so refs and state never go stale.
 */
import { ILayoutService } from '@univerjs/preset-sheets-core'

import { columnLabel } from '../domain/cell-address'
import { decodeCsvBuffer, isNumericCell, parseCsv } from '../gateway/csv-import'
import type { AdvancedFilterColumn, AdvancedFilterCriteria } from './AdvancedFilterDialog'
import {
  buildLabelMatrix,
  buildPositionMatrix,
  parseConsolidateReference,
  targetOverlapsSource,
  type ConsolidateArea,
  type ConsolidateConfig,
  type OutputCell,
} from './consolidate'
import { isSheetRemoved, journalSize } from './edit-journal'
import { resolveGoToRef, type GoToNameEntry } from './goto'
import { getLang, t } from './i18n/locale'
import { appendSymbol } from './SymbolDialog'
import {
  a1RangeRef,
  a1RowRangeRef,
  advancedFilterColumnOptions,
  applyFilterCriteria,
  columnLetter,
  loadVisibleRange,
  univerDefinedNames,
  revealCellBelowFreeze,
} from './univer-sync'
import { applyOutlineToRange, outlineRangeForSelection } from './outline-ops'
import { applySubtotalToRange, type SubtotalOutcome } from './subtotal-ops'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'
import { applyAiTableAdd } from './workbook-ops'

/** The App refs/state the data-tool actions need; built fresh per call. */
export interface DataToolsContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
  setAdvancedFilterColumns: (columns: readonly AdvancedFilterColumn[] | null) => void
}

/// Matches the paste ceiling in spirit: one setValues command, journaled and
/// undoable; larger files should open as their own workbook instead.
const CSV_IMPORT_MAX_CELLS = 50_000

/// Excel writes CSV in the system's legacy charset; the UI language is the
/// best tie-breaker we have (same map as the main-process open path).
const CSV_CHARSET_BY_LANG: Record<string, string> = {
  zh: 'gb18030',
  'zh-TW': 'big5',
  ja: 'shift_jis',
  ko: 'euc-kr',
}

/// Data → From Text/CSV: reads a delimited file (encoding and delimiter
/// sniffed by the shared CSV pipeline) into the active sheet at the selection.
export function handleImportCsv(ctx: DataToolsContext): void {
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = '.csv,.tsv,.txt,text/csv,text/tab-separated-values,text/plain'
  input.onchange = () => {
    const file = input.files?.[0]
    if (!file) return
    if (file.size > 32 * 1024 * 1024) {
      ctx.setMessage(t('appCsvTooLarge'))
      return
    }
    void file.arrayBuffer().then((buffer) => {
      importCsvText(ctx, decodeCsvBuffer(new Uint8Array(buffer), CSV_CHARSET_BY_LANG[getLang()]))
    })
  }
  input.click()
}

function importCsvText(ctx: DataToolsContext, text: string): void {
  const runtime = ctx.univerRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const rows = parseCsv(text)
  const columns = rows.reduce((width, row) => Math.max(width, row.length), 0)
  if (rows.length === 0 || columns === 0) {
    ctx.setMessage(t('appCsvEmpty'))
    return
  }
  if (rows.length * columns > CSV_IMPORT_MAX_CELLS) {
    ctx.setMessage(t('appCsvTooLarge'))
    return
  }
  const values = rows.map((row) =>
    Array.from({ length: columns }, (_, index) => {
      const cell = row[index] ?? ''
      return isNumericCell(cell) ? { v: Number(cell) } : { v: cell }
    }),
  )
  const row = range.getRow()
  const column = range.getColumn()
  try {
    worksheet.getRange(row, column, rows.length, columns).setValues(values)
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appCsvImportFailed'))
    return
  }
  ctx.setMessage(
    t('appCsvImported', {
      rows: rows.length,
      columns,
      cell: `${columnLetter(column)}${row + 1}`,
    }),
  )
}

export function activeCellLabel(ctx: DataToolsContext): string {
  const range = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()
  if (!range) return t('appActiveCellFallback')
  return `${columnLetter(range.getColumn())}${range.getRow() + 1}`
}

/// Name Box / Go To jump: resolves the typed reference (an A1 address or a
/// defined name) and selects it, switching sheets when the target lives on
/// another one. Returns null on success or a user-facing error message.
/// Univer's parser yields NaN rows instead of throwing on garbage, so
/// validity is decided by resolveGoToRef — the try/catch covers what
/// getRange itself rejects: unknown sheet names and out-of-bounds ranges.
export function goToReference(ctx: DataToolsContext, ref: string): string | null {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return t('appGoToNotReady')
  const trimmed = ref.trim()
  if (trimmed === '') return t('appGoToEmpty')
  const resolved = resolveGoToRef(trimmed, listDefinedNames(ctx))
  if (resolved === null) {
    return t('appGoToUnresolved', { ref: trimmed })
  }
  try {
    // A jump must not leave an editor open on the previous cell — later
    // keystrokes would land there. Commit it before moving.
    const editing = workbook as unknown as {
      isCellEditing?(): boolean
      endEditingAsync?(save: boolean): Promise<boolean>
    }
    if (editing.isCellEditing?.()) void editing.endEditingAsync?.(true)
    const range = worksheet.getRange(resolved)
    const target = workbook.getSheetBySheetId(range.getSheetId()) ?? worksheet
    workbook.setActiveRange(range)
    void revealCellBelowFreeze(target, range.getRow(), range.getColumn())
    // Hand keyboard focus back to the grid (Univer's hidden editor host, the
    // same handoff its own name box does); typing right after a jump then
    // lands in the target cell instead of being dropped on <body>.
    ctx.univerRef.current?.univer.__getInjector().get(ILayoutService).focus()
    // The single Scroll event a jump produces computes its viewport from a
    // stale getVisibleRange; load around the target explicitly.
    if (ctx.univerRef.current) {
      void loadVisibleRange(
        ctx.univerRef.current,
        ctx.lazyWorkbookRef as { current: LazyWorkbookState | null },
        target,
        ctx.setMessage,
        { row: range.getRow(), column: range.getColumn() },
      )
    }
  } catch (error: unknown) {
    const detail = error instanceof Error ? error.message : ''
    if (detail.includes('out of bounds')) return t('appGoToOutOfBounds', { ref: trimmed })
    if (detail.includes('Range not found')) return t('appGoToSheetNotFound', { ref: trimmed })
    return detail === ''
      ? t('appGoToInvalidAddress', { ref: trimmed })
      : t('appGoToFailed', { ref: trimmed, detail })
  }
  return null
}

/// All defined names of the active workbook, feeding both the Go To
/// dialog's list and Name Box resolution.
export function listDefinedNames(ctx: DataToolsContext): GoToNameEntry[] {
  return univerDefinedNames(ctx.univerRef.current).map((defined) => ({
    name: defined.getName(),
    ref: defined.getFormulaOrRefString(),
  }))
}

export function handleApplyFormula(ctx: DataToolsContext, formula: string): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) return t('appSelectCellFirst')
  const trimmed = formula.trim()
  if (!trimmed.startsWith('=')) return t('appFormulaStartsEquals')
  const opens = (trimmed.match(/\(/g) ?? []).length
  const closes = (trimmed.match(/\)/g) ?? []).length
  if (opens !== closes) return t('appUnbalancedParens')
  try {
    worksheet.getRange(range.getRow(), range.getColumn(), 1, 1).setValue({ f: trimmed })
  } catch (error: unknown) {
    return error instanceof Error ? error.message : t('appSetFormulaFailed')
  }
  ctx.setMessage(t('appFormulaSet', { cell: activeCellLabel(ctx) }))
  return null
}

/// A label turned into a legal defined name, Excel-style: illegal characters
/// become underscores, and anything that could read as a cell reference gets
/// an underscore prefix.
function definedNameFromLabel(label: string): string | null {
  const cleaned = label.trim().replace(/[^\p{L}\p{N}_.]/gu, '_')
  if (!cleaned || /^\.+$/.test(cleaned)) return null
  const named = /^[\p{L}_]/u.test(cleaned) ? cleaned : `_${cleaned}`
  const cellLike = /^[A-Za-z]{1,3}\d+$/.test(named) || /^[Rr]\d*([Cc]\d*)?$/.test(named)
  return cellLike ? `_${named}` : named
}

/// Formulas → Create from Selection: one defined name per column (labels in
/// the selection's top row) or per row (labels in its left column). The
/// defined-name mutation listener journals each insert.
export function handleCreateNamesFromSelection(ctx: DataToolsContext, mode: 'top' | 'left'): void {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const selection = workbook?.getActiveRange()?.getRange()
  if (!workbook || !worksheet || !selection) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const { startRow, endRow, startColumn, endColumn } = selection
  if (mode === 'top' ? endRow <= startRow : endColumn <= startColumn) {
    ctx.setMessage(t('appCreateNamesNeedsData'))
    return
  }
  const sheetName = worksheet.getSheetName()
  // Names span the full selection like Excel (empty data rows included),
  // but labels past the data extent are blank and would be skipped one by
  // one — stop the label walk there so a whole-column selection doesn't
  // read a million empty cells. File cells stream into Univer lazily, so
  // getLastRow/getLastColumn alone can undercount; the file's used range is
  // the floor.
  const fileSheet = ctx.lazyWorkbookRef.current?.file.sheets.find(
    (sheet) => sheet.id === worksheet.getSheetId(),
  )
  const labelEndRow = Math.min(
    endRow,
    Math.max(worksheet.getLastRow(), (fileSheet?.rowCount ?? 0) - 1),
  )
  const labelEndColumn = Math.min(
    endColumn,
    Math.max(worksheet.getLastColumn(), (fileSheet?.columnCount ?? 0) - 1),
  )
  // Labels come from the formatted display text — a date header must name
  // like "Jan_2023", not its raw serial.
  const entries: { label: string; ref: string }[] = []
  if (mode === 'top') {
    for (let column = startColumn; column <= labelEndColumn; column += 1) {
      entries.push({
        label: worksheet.getRange(startRow, column, 1, 1).getDisplayValue() ?? '',
        ref: a1RangeRef(sheetName, column, startRow + 1, endRow),
      })
    }
  } else {
    for (let row = startRow; row <= labelEndRow; row += 1) {
      entries.push({
        label: worksheet.getRange(row, startColumn, 1, 1).getDisplayValue() ?? '',
        ref: a1RowRangeRef(sheetName, row, startColumn + 1, endColumn),
      })
    }
  }
  // insertDefinedName keys entries by internal id and accepts duplicates
  // without error; a second entry with the same name fails the next save.
  // Dedupe here — against existing names and within the batch, both
  // case-insensitive like Excel.
  const taken = new Set(
    univerDefinedNames(ctx.univerRef.current).map((defined) => defined.getName().toLowerCase()),
  )
  let created = 0
  let skipped = 0
  for (const entry of entries) {
    const name = definedNameFromLabel(entry.label)
    if (!name || taken.has(name.toLowerCase())) {
      skipped += 1
      continue
    }
    try {
      workbook.insertDefinedName(name, entry.ref)
      taken.add(name.toLowerCase())
      created += 1
    } catch {
      // Reserved or otherwise rejected name — Excel prompts here; we skip.
      skipped += 1
    }
  }
  ctx.setMessage(
    skipped > 0
      ? t('appNamesCreatedSkipped', { count: created, skipped })
      : t('appNamesCreated', { count: created }),
  )
}

/// The Symbol dialog's click: appends the picked character to the active
/// range's top-left cell as text. setValue lands as a normal
/// set-range-values command, so it journals and undoes like typing.
export function handleInsertSymbol(ctx: DataToolsContext, char: string): void {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appSymbolNeedsCell'))
    return
  }
  const cell = worksheet.getRange(range.getRow(), range.getColumn(), 1, 1)
  try {
    cell.setValue(appendSymbol(cell.getValue(), char))
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appSymbolInsertFailed'))
    return
  }
  ctx.setMessage(t('appSymbolInserted', { char, cell: activeCellLabel(ctx) }))
}

/// Samples the active filter range's header row and opens the Advanced
/// Filter dialog; without an auto-filter it explains instead of opening.
export function openAdvancedFilterDialog(ctx: DataToolsContext): void {
  const worksheet = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  const filter = worksheet?.getFilter()
  if (!worksheet || !filter) {
    ctx.setMessage(t('appAdvFilterNeedsFilter'))
    return
  }
  ctx.setAdvancedFilterColumns(advancedFilterColumnOptions(worksheet, filter))
}

/// The Advanced Filter dialog's OK: lands through the same
/// applyFilterCriteria path as the AI op set_filter_criteria.
export function handleApplyAdvancedFilter(
  ctx: DataToolsContext,
  criteria: AdvancedFilterCriteria,
): string | null {
  const worksheet = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  const filter = worksheet?.getFilter()
  if (!worksheet || !filter) return t('appAdvFilterGone')
  const startColumn = filter.getRange().getRange().startColumn
  try {
    applyFilterCriteria(worksheet, columnLabel(startColumn + criteria.colId), {
      customs: { and: criteria.and, filters: criteria.filters },
    })
  } catch (error: unknown) {
    return error instanceof Error ? error.message : t('appAdvFilterFailed')
  }
  ctx.setMessage(t('appAdvFilterApplied'))
  return null
}

/// Data → Subtotal dialog: runs the shared subtotal engine (subtotal-ops.ts,
/// which the AI `subtotal` op also uses) over the current selection. The
/// dialog's "Count" has always meant COUNTA (SUBTOTAL 3), so it stays that.
export function handleCreateSubtotal(
  ctx: DataToolsContext,
  config: {
    groupCol: number
    valueCol: number
    agg: 'sum' | 'count' | 'average'
  },
): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) return t('appSelectSourceRangeFirst')
  const values = range.getValues()
  if (values.length < 2) return t('appSelectRangeWithHeader')
  const startRow = range.getRow()
  const startColumn = range.getColumn()
  if (config.groupCol === config.valueCol) return t('appSubtotalSameColumns')
  const state = ctx.lazyWorkbookRef.current
  const selection =
    `${columnLabel(startColumn)}${startRow + 1}:` +
    `${columnLabel(startColumn + range.getWidth() - 1)}${startRow + range.getHeight()}`
  let outcome: SubtotalOutcome
  try {
    outcome = applySubtotalToRange(runtime, state, {
      sheetId: worksheet.getSheetId(),
      range: selection,
      groupColumn: columnLabel(config.groupCol),
      valueColumn: columnLabel(config.valueCol),
      function: config.agg === 'sum' ? 'sum' : config.agg === 'average' ? 'average' : 'counta',
    })
  } catch (error: unknown) {
    return error instanceof Error ? error.message : t('appSubtotalInsertFailed')
  }
  if (!outcome.applied) {
    if (outcome.reason === 'too-many-groups') return t('appSubtotalTooManyGroups')
    if (outcome.reason === 'same-column') return t('appSubtotalSameColumns')
    if (outcome.reason === 'no-data-rows' || outcome.reason === 'range-too-small') {
      return t('appRangeNoDataRows')
    }
    return t('appSubtotalInsertFailed')
  }
  if (state) ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.setMessage(t('appSubtotalsInserted', { count: outcome.groups ?? 0 }))
  return null
}

export function consolidateDefaultReference(ctx: DataToolsContext): string {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range || (range.getHeight() === 1 && range.getWidth() === 1)) return ''
  const name = worksheet.getSheetName()
  const prefix = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replaceAll("'", "''")}'`
  return `${prefix}!${range.getA1Notation()}`
}

export function handleCreateConsolidate(
  ctx: DataToolsContext,
  config: ConsolidateConfig,
): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!workbook || !worksheet || !active) return t('appSelectTargetCellFirst')
  const activeSheetName = worksheet.getSheetName()
  const sheets = workbook.getSheets()
  const areas: ConsolidateArea[] = []
  const sources: UniverWorksheet[] = []
  for (const reference of config.references) {
    const parsed = parseConsolidateReference(reference)
    if (!parsed) return t('appConsolidateBadRef', { ref: reference })
    const name = parsed.sheetName ?? activeSheetName
    const source = sheets.find((sheet) => sheet.getSheetName().toLowerCase() === name.toLowerCase())
    if (!source) return t('appNoSheetNamed', { name })
    sources.push(source)
    areas.push({
      sheetName: source.getSheetName(),
      startRow: parsed.range.startRow,
      startColumn: parsed.range.startColumn,
      rows: parsed.range.endRow - parsed.range.startRow + 1,
      columns: parsed.range.endColumn - parsed.range.startColumn + 1,
    })
  }
  const target = { row: active.getRow(), column: active.getColumn() }
  let matrix: OutputCell[][]
  if (config.leftLabels) {
    const areaLabels = areas.map((area, index) =>
      (
        sources[index]?.getRange(area.startRow, area.startColumn, area.rows, 1).getValues() ?? []
      ).map((row) => String(row?.[0] ?? '')),
    )
    const built = buildLabelMatrix(config.fn, areas, areaLabels, target)
    if (typeof built === 'string') return built
    matrix = built
  } else {
    matrix = buildPositionMatrix(config.fn, areas)
  }
  const columns = matrix[0]?.length ?? 0
  if (matrix.length === 0 || columns === 0) return t('appConsolidateEmptySources')
  if (
    targetOverlapsSource(areas, activeSheetName, {
      ...target,
      rows: matrix.length,
      columns,
    })
  ) {
    return t('appConsolidateOverlap')
  }
  try {
    worksheet
      .getRange(target.row, target.column, matrix.length, columns)
      .setValues(
        matrix.map((line) =>
          line.map((cell) => (cell.f !== undefined ? { f: cell.f } : { v: cell.v ?? null })),
        ),
      )
  } catch (error: unknown) {
    return error instanceof Error ? error.message : t('appConsolidateWriteFailed')
  }
  ctx.setMessage(
    t('appConsolidateDone', {
      count: areas.length,
      cell: `${columnLetter(target.column)}${target.row + 1}`,
      rows: matrix.length,
      columns,
    }),
  )
  return null
}

/// Outline groups: level edits journal directly (Univer has no outline
/// model, so there is no command to undo); Hide/Show Detail rides the
/// normal hidden-rows pipeline plus a collapsed flag on the summary line.
/// The shared helper (outline-ops.ts) also backs the AI `outline.*` ops.
export function handleOutline(
  ctx: DataToolsContext,
  action: 'group' | 'ungroup' | 'hide-detail' | 'show-detail',
  axis: 'rows' | 'cols',
): void {
  const runtime = ctx.univerRef.current
  if (!runtime) return
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    ctx.setMessage(t('appOutlineNeedsFile'))
    return
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appOutlineSelectFirst'))
    return
  }
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  const start = axis === 'rows' ? range.getRow() : range.getColumn()
  const end = axis === 'rows' ? start + range.getHeight() - 1 : start + range.getWidth() - 1
  const outcome = applyOutlineToRange(runtime, state, {
    sheetId,
    axis,
    range: outlineRangeForSelection(axis, start, end),
    action,
  })
  if (!outcome.applied) {
    if (outcome.reason === 'max-level') ctx.setMessage(t('appOutlineMaxLevel'))
    else if (outcome.reason === 'nothing-to-ungroup') ctx.setMessage(t('appNothingToUngroup'))
    return
  }
  ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.setMessage(
    action === 'hide-detail'
      ? t('appDetailHidden')
      : action === 'show-detail'
        ? t('appDetailShown')
        : action === 'group'
          ? axis === 'rows'
            ? t('appRowsGrouped')
            : t('appColsGrouped')
          : axis === 'rows'
            ? t('appRowsUngrouped')
            : t('appColsUngrouped'),
  )
}

/// Home → Format as Table: the manual entry over the same engine as the
/// AI add_table op (Univer table rendering + journal + native table part).
export function handleFormatAsTable(ctx: DataToolsContext, style: string): void {
  const runtime = ctx.univerRef.current
  if (!runtime) return
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    ctx.setMessage(t('appTablesNeedFile'))
    return
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appTableSelectRange'))
    return
  }
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  const startRow = range.getRow()
  const startColumn = range.getColumn()
  const endRow = startRow + range.getHeight() - 1
  const endColumn = startColumn + range.getWidth() - 1
  try {
    applyAiTableAdd(runtime, state, {
      op: 'add_table',
      sheetId,
      range: `${columnLabel(startColumn)}${startRow + 1}:${columnLabel(endColumn)}${endRow + 1}`,
      style,
      bandedRows: true,
    })
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t('appTableCreated'))
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appTableCreateFailed'))
  }
}
