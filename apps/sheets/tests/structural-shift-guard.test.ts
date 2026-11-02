import { describe, expect, it } from 'vitest'

import {
  findBlockingReference,
  findBlockingToken,
  type DeletedRange,
  type SheetFormulas,
  structuralShiftBlockMessage,
  deletionOp,
} from '../src/renderer/structural-shift-guard'

const SHEET = { sheetId: 'sheet-1', sheetName: '薪酬明细' }
const OTHER = { sheetId: 'sheet-2', sheetName: '部门汇总' }

function sheets(
  mainCells: { cell: string; formula: string }[],
  otherCells: { cell: string; formula: string }[] = [],
): SheetFormulas[] {
  return [
    { sheetId: SHEET.sheetId, name: SHEET.sheetName, cells: mainCells },
    { sheetId: OTHER.sheetId, name: OTHER.sheetName, cells: otherCells },
  ]
}

/// 删除 G 列（0-based 6）。
const DELETE_G: DeletedRange = { axis: 'column', index: 6, count: 1 }

describe('删列预检：引用落在被删列里的公式', () => {
  it('应发合计 =E5+F5+G5 引用了 G5 → 阻止删除并点名 H5', () => {
    const hit = findBlockingReference(
      sheets([{ cell: 'H5', formula: '=E5+F5+G5' }]),
      SHEET,
      DELETE_G,
    )
    expect(hit).toMatchObject({
      sheetName: '薪酬明细',
      cell: 'H5',
      reference: 'G5',
    })
  })

  it('不引用被删列的公式放行', () => {
    const hit = findBlockingReference(
      sheets([
        { cell: 'H5', formula: '=E5+F5' },
        { cell: 'I5', formula: '=ROUND(E5*0.225,2)' },
      ]),
      SHEET,
      DELETE_G,
    )
    expect(hit).toBeNull()
  })

  it('字符串字面量里的 G5 不算引用', () => {
    const hit = findBlockingReference(
      sheets([{ cell: 'A1', formula: '=IF(B1="G5","G5",0)' }]),
      SHEET,
      DELETE_G,
    )
    expect(hit).toBeNull()
  })

  it('整列引用 SUM(G:G) 命中', () => {
    const hit = findBlockingReference(
      sheets([{ cell: 'Z1', formula: '=SUM(G:G)' }]),
      SHEET,
      DELETE_G,
    )
    expect(hit?.reference).toBe('G:G')
  })

  it('整行引用不受删列影响', () => {
    const hit = findBlockingReference(
      sheets([{ cell: 'Z1', formula: '=SUM(5:5)' }]),
      SHEET,
      DELETE_G,
    )
    expect(hit).toBeNull()
  })

  it('跨表显式限定到目标表 → 命中；指向自己那张表则放行', () => {
    const qualified = findBlockingReference(
      sheets([], [{ cell: 'A4', formula: '=SUMIF(薪酬明细!C5:C14,A4,薪酬明细!G5:G14)' }]),
      SHEET,
      DELETE_G,
    )
    expect(qualified).toMatchObject({ sheetName: '部门汇总', cell: 'A4' })

    const own = findBlockingReference(
      sheets([], [{ cell: 'B2', formula: '=SUM(G5:G14)' }]),
      SHEET,
      DELETE_G,
    )
    expect(own).toBeNull()
  })

  it('带引号的表名（含空格）也能匹配', () => {
    const hit = findBlockingReference(
      sheets([], [{ cell: 'A1', formula: "='薪酬 明细'!G5+1" }]),
      { sheetId: SHEET.sheetId, sheetName: '薪酬 明细' },
      DELETE_G,
    )
    expect(hit?.reference).toBe('G5')
  })
})

describe('删行预检', () => {
  it('把被引用的行删掉 → 命中', () => {
    const range: DeletedRange = { axis: 'row', index: 4, count: 1 }
    const hit = findBlockingReference(
      sheets([{ cell: 'H9', formula: '=E5+F5+G5' }]),
      SHEET,
      range,
    )
    expect(hit?.reference).toBe('E5')
  })

  it('合计行 SUM(E5:E14) 整体被删 → 命中；只删空行则放行', () => {
    const inside = findBlockingReference(
      sheets([{ cell: 'H15', formula: '=ROUND(SUM(E5:E14),2)' }]),
      SHEET,
      { axis: 'row', index: 4, count: 10 },
    )
    expect(inside?.reference).toBe('E5:E14')

    const empty = findBlockingReference(
      sheets([{ cell: 'H15', formula: '=ROUND(SUM(E5:E14),2)' }]),
      SHEET,
      { axis: 'row', index: 20, count: 1 },
    )
    expect(empty).toBeNull()
  })

  it('部分相交的区间不算「整个落在被删区间里」（保持与写盘判定一致）', () => {
    const hit = findBlockingReference(
      sheets([{ cell: 'H15', formula: '=SUM(E5:E14)' }]),
      SHEET,
      { axis: 'row', index: 8, count: 2 },
    )
    expect(hit).toBeNull()
  })
})

describe('细节', () => {
  it('findBlockingToken 尊重限定表名', () => {
    const op = deletionOp(DELETE_G)
    expect(findBlockingToken('=薪酬明细!G5', op, (q) => q === '薪酬明细')).toBe('G5')
    expect(findBlockingToken('=部门汇总!G5', op, (q) => q === '薪酬明细')).toBeNull()
    expect(findBlockingToken('=G5', op, (q) => q === undefined)).toBe('G5')
  })

  it('拒绝理由点名单元格、引用，并说明没有改动', () => {
    const message = structuralShiftBlockMessage(
      {
        sheetId: SHEET.sheetId,
        sheetName: '薪酬明细',
        cell: 'H5',
        reference: 'G5',
        formula: '=E5+F5+G5',
      },
      DELETE_G,
    )
    expect(message).toContain('薪酬明细!H5')
    expect(message).toContain('G5')
    expect(message).toContain('nothing was changed and nothing was saved')
  })
})
