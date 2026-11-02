/**
 * 把「删行/删列的公式引用预检」接到 Univer 运行时上。
 *
 * 两个入口，覆盖两条真实路径，判定共用 `structural-shift-guard`：
 * 1. `structuralShiftBlock` —— AI/DSL 计划在 apply 前整批预检（错误回给 agent）；
 * 2. `installStructuralShiftVeto` —— 用户手势（右键/快捷键删行删列）走
 *    `beforeCommandExecuted` 否决，连内存都不会被改。
 *
 * 两条路都保证：被拒绝时**没有改内存、没有进编辑日志**，因此不会出现
 * 「日志写不出去 → 整份工作簿永远保存失败」的阻塞态。
 */
import { CustomCommandExecutionError, ICommandService } from '@univerjs/core'

import { columnLabel } from '../domain/cell-address'
import {
  findBlockingReference,
  structuralShiftBlockMessage,
  type BlockingReference,
  type DeletedAxis,
  type DeletedRange,
  type FormulaCell,
  type SheetFormulas,
} from './structural-shift-guard'
import { journalSuppression, type ActiveWorkbook, type UniverRuntime } from './univer-state'

/// 超过这个规模就不再整表读公式（读它比写盘还贵）；这类大表退回保存阶段的
/// 检查兜底（错误已本地化并提示撤销）。
const FORMULA_SCAN_CELL_LIMIT = 400_000

/// 用户在界面上「删除行 / 删除列」时命中的命令。带 -by-range 的变体参数形状
/// 不同，这里不猜参数，交给保存阶段的兜底。
const DELETE_COMMAND_AXIS: Readonly<Record<string, DeletedAxis>> = {
  'sheet.command.remove-row': 'row',
  'sheet.command.remove-col': 'column',
}

export function readSheetFormulas(workbook: ActiveWorkbook, sheetId: string): FormulaCell[] {
  const sheet = workbook.getSheetBySheetId(sheetId)
  if (!sheet) return []
  const lastRow = sheet.getLastRow()
  const lastColumn = sheet.getLastColumn()
  if (!(lastRow >= 0) || !(lastColumn >= 0)) return []
  const rows = lastRow + 1
  const columns = lastColumn + 1
  if (rows * columns > FORMULA_SCAN_CELL_LIMIT) return []
  const grid = sheet.getRange(0, 0, rows, columns).getFormulas() as unknown[][]
  const cells: FormulaCell[] = []
  for (let row = 0; row < rows; row += 1) {
    const line = grid[row]
    if (!Array.isArray(line)) continue
    for (let column = 0; column < columns; column += 1) {
      const formula = line[column]
      if (typeof formula !== 'string' || formula.trim() === '') continue
      cells.push({ cell: `${columnLabel(column)}${row + 1}`, formula })
    }
  }
  return cells
}

/** 整本工作簿的公式快照（跨表引用也要能查到）。 */
export function collectWorkbookFormulas(workbook: ActiveWorkbook): SheetFormulas[] {
  return workbook.getSheets().map((sheet) => ({
    sheetId: sheet.getSheetId(),
    name: sheet.getSheetName(),
    cells: readSheetFormulas(workbook, sheet.getSheetId()),
  }))
}

function blockingHit(
  workbook: ActiveWorkbook,
  target: { readonly sheetId: string; readonly sheetName: string },
  range: DeletedRange,
): BlockingReference | null {
  return findBlockingReference(collectWorkbookFormulas(workbook), target, range)
}

/**
 * 计划级预检：`sheets` 用 state.file.sheets（含未渲染的表），保证跨表引用也能查到。
 * 返回拒绝理由，null = 可以安全删除。
 */
export function structuralShiftBlock(
  runtime: UniverRuntime,
  sheets: readonly { readonly id: string; readonly name: string }[],
  target: { readonly sheetId: string },
  range: DeletedRange,
): string | null {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return null
  const sheetName =
    sheets.find((sheet) => sheet.id === target.sheetId)?.name ??
    workbook.getSheetBySheetId(target.sheetId)?.getSheetName()
  if (!sheetName) return null
  const hit = blockingHit(workbook, { sheetId: target.sheetId, sheetName }, range)
  return hit ? structuralShiftBlockMessage(hit, range) : null
}

/**
 * 用户手势否决：删行/删列前检查选区内的引用，命中就抛错取消命令并提示。
 * 预览（readonly）由 readonly-veto 统一拦，不需要重复安装。
 */
export function installStructuralShiftVeto(
  runtime: UniverRuntime,
  onBlocked: (message: string) => void,
): void {
  const commandService = runtime.univer.__getInjector().get(ICommandService)
  commandService.beforeCommandExecuted((command) => {
    const axis = DELETE_COMMAND_AXIS[command.id]
    if (!axis || journalSuppression.active) return
    const workbook = runtime.univerAPI.getActiveWorkbook()
    if (!workbook) return
    const sheet = workbook.getActiveSheet()
    const activeRange = workbook.getActiveRange()
    if (!sheet || !activeRange) return
    const range: DeletedRange =
      axis === 'row'
        ? { axis, index: activeRange.getRow(), count: activeRange.getHeight() }
        : { axis, index: activeRange.getColumn(), count: activeRange.getWidth() }
    if (range.count < 1) return
    const hit = blockingHit(
      workbook,
      { sheetId: sheet.getSheetId(), sheetName: sheet.getSheetName() },
      range,
    )
    if (!hit) return
    // Univer 只在成功路径清理命令执行栈；直接抛错会留下这次 dispatch 的条目，
    // 之后每个 mutation 都要 findLast 扫它（与 readonly-veto 同一处理）。
    const stack = (commandService as unknown as { _commandExecutionStack?: unknown[] })
      ._commandExecutionStack
    const index = stack?.indexOf(command) ?? -1
    if (index >= 0) stack?.splice(index, 1)
    onBlocked(structuralShiftBlockMessage(hit, range))
    throw new CustomCommandExecutionError('deleting this row/column would break a formula')
  })
}
