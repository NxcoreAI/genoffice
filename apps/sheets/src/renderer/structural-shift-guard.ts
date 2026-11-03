/**
 * 删行/删列前的公式引用预检。
 *
 * 为什么需要它：gateway 写盘时（`gateway/xlsx-structure.ts` 的
 * `shiftFormulaText`）如果发现有公式引用的单元格**整个落在被删区间里**，会抛
 * `StructuralShiftError` 中止这一批结构操作——这是有意的（照 Excel 的做法写出去
 * 就是 `#REF!`）。问题在于那一刻删除已经改进了内存、并且记进了编辑日志：日志
 * 再也写不出去，保存 / 自动保存 / 崩溃恢复三条路一起失败，用户看到的是
 * 「点保存一直报错」，只能靠撤销或丢弃改动逃生。
 *
 * 这里把**同一个判定**提到改内存之前：判定本身不复制一份，而是直接问
 * `xlsx-structure` 的 `referenceSurvivesOp`，保证预检与写盘结论永远一致。
 */
import {
  FORMULA_REFERENCE_PATTERN,
  type RowColumnOp,
  qualifierMatches,
  referenceSurvivesOp,
} from '../gateway/xlsx-structure'

export type DeletedAxis = 'row' | 'column'

/** 0-based 的删除区间（与 StructuralOp 的 index/count 同义）。 */
export interface DeletedRange {
  readonly axis: DeletedAxis
  readonly index: number
  readonly count: number
}

export interface FormulaCell {
  /** A1 坐标，用于报错定位。 */
  readonly cell: string
  readonly formula: string
}

export interface SheetFormulas {
  readonly sheetId: string
  readonly name: string
  readonly cells: readonly FormulaCell[]
}

export interface BlockingReference {
  readonly sheetId: string
  readonly sheetName: string
  /** 持有该公式的单元格（A1）。 */
  readonly cell: string
  /** 公式里指向被删区间的那个引用 token。 */
  readonly reference: string
  readonly formula: string
}

/** 删除区间 → gateway 的结构操作，供 `referenceSurvivesOp` 复用同一份平移逻辑。 */
export function deletionOp(range: DeletedRange): RowColumnOp {
  return range.axis === 'column'
    ? { kind: 'remove-cols', index: range.index, count: range.count }
    : { kind: 'remove-rows', index: range.index, count: range.count }
}

/**
 * 公式里第一处「指向被删区间」的引用 token；没有则返回 null。
 *
 * `acceptQualifier` 决定某个引用是否属于本次判定：无限定引用只在目标表里算，
 * 带 `目标表!` 前缀的跨表引用在任何表里都算——与 gateway 同表平移 + 跨表
 * qualifiedOnly 平移的语义一致。
 */
export function findBlockingToken(
  formula: string,
  op: RowColumnOp,
  acceptQualifier: (qualifier: string | undefined) => boolean,
): string | null {
  // 公式的字符串字面量用 "" 转义：按 `"` 切分后奇数段是字面量内容，跳过。
  const segments = formula.split('"')
  for (let index = 0; index < segments.length; index += 2) {
    const segment = segments[index]!
    FORMULA_REFERENCE_PATTERN.lastIndex = 0
    let match: RegExpExecArray | null
    while ((match = FORMULA_REFERENCE_PATTERN.exec(segment)) !== null) {
      const qualifier = match[2]
      const token = match[3]
      if (!token) continue
      if (!acceptQualifier(qualifier)) continue
      if (!referenceSurvivesOp(token, op)) return token
    }
  }
  return null
}

/**
 * 找出第一处「公式引用了将被删掉的行/列」的位置。
 *
 * `target` 是要删行/列的工作表。同表里无限定的引用、以及任何表里显式限定到
 * 目标表的引用都算——后者正是 gateway 跨表修订公式时会 abort 的那类。
 */
export function findBlockingReference(
  sheets: readonly SheetFormulas[],
  target: { readonly sheetId: string; readonly sheetName: string },
  range: DeletedRange,
): BlockingReference | null {
  const op = deletionOp(range)
  for (const sheet of sheets) {
    const isTargetSheet = sheet.sheetId === target.sheetId
    for (const cell of sheet.cells) {
      const token = findBlockingToken(cell.formula, op, (qualifier) => {
        if (qualifier === undefined) return isTargetSheet
        return qualifierMatches(qualifier, target.sheetName)
      })
      if (token) {
        return {
          sheetId: sheet.sheetId,
          sheetName: sheet.name,
          cell: cell.cell,
          reference: token,
          formula: cell.formula,
        }
      }
    }
  }
  return null
}

/**
 * 面向用户/agent 的拒绝理由：点名具体单元格与引用，并说明先改公式再删。
 * 与其它 gate 错误一样用英文（plan-operations 的 gate 文案同为英文）。
 */
export function structuralShiftBlockMessage(hit: BlockingReference, range: DeletedRange): string {
  const unit = range.axis === 'column' ? 'column' : 'row'
  const what = range.count > 1 ? `${range.count} ${unit}s` : `that ${unit}`
  return (
    `Cannot delete ${what}: the formula in ${hit.sheetName}!${hit.cell} ` +
    `(${hit.formula}) references ${hit.reference}, which sits inside the deleted ${unit}. ` +
    `Rewrite or clear that formula first, then delete — nothing was changed and nothing was saved.`
  )
}
