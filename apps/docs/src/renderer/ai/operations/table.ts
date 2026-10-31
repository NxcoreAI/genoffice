import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { Command } from '@tiptap/pm/state'
import { TextSelection } from '@tiptap/pm/state'
import {
  CellSelection,
  TableMap,
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  mergeCells,
  splitCell,
} from '@tiptap/pm/tables'
import type {
  SectionSettings,
  TableBorders,
  TableCell,
  TableLook,
  TableModel,
} from '@genoffice/docx-engine'
import { TABLE_HEADER_FILL } from '@genoffice/docx-engine'
import {
  applyTablePreset,
  setTableAutoFit,
  setTableLookOption,
  type TablePreset,
} from '../../editor/table-properties'
import { tableModelToPmNode } from '../../editor/convert'
import { blockEndPos, clampInt, failure, hexOf, intOf, planned, success } from './shared'
import type { OperationContext, OperationDef, OperationOutcome } from './types'

/** 96dpi CSS px ⇄ twips (OOXML dxa). */
const TWIPS_PER_PX = 15
const MAX_INSERT_ROWS = 50
const MAX_INSERT_COLS = 20
const MAX_CELL_TEXT = 20_000
const MAX_HEX = /^[0-9A-Fa-f]{6}$/
const BORDER_STYLES = new Set(['single', 'double', 'dashed', 'dotted', 'none'])
type Obj = Record<string, unknown>
type BorderLine = { style: string; szEighths: number; color: string }

/** Text-column width in px (a table must fit inside it). */
export function contentWidthPxFromSection(section: SectionSettings | null): number {
  if (!section) return 600
  const contentTw = section.pageWidth - section.marginLeft - section.marginRight
  const cols = Math.max(1, section.columns ?? 1)
  const columnTw =
    cols > 1 ? (contentTw - (section.colSpace ?? 720) * (cols - 1)) / cols : contentTw
  return Math.max(120, Math.round(columnTw / TWIPS_PER_PX))
}

// ── table-specific validation helpers ──

function borderLine(value: unknown): BorderLine | null | string {
  if (value === null) return null
  if (!value || typeof value !== 'object') return 'border must be an object like { style, szEighths, color } or null'
  const raw = value as Obj
  const style = raw.style === undefined ? 'single' : String(raw.style)
  if (!BORDER_STYLES.has(style)) {
    return `border.style must be one of ${[...BORDER_STYLES].join('/')}`
  }
  const sz = raw.szEighths === undefined ? 4 : intOf(raw.szEighths)
  if (sz === null || sz < 0 || sz > 96) return 'border.szEighths must be an integer 0..96 (eighths of a point)'
  const color = raw.color === undefined ? 'auto' : String(raw.color).replace(/^#/, '').toUpperCase()
  if (color !== 'auto' && !MAX_HEX.test(color)) return 'border.color must be 6-digit hex (no #) or "auto"'
  return { style, szEighths: sz, color }
}

/** Merge one side of a borders object into the existing per-side model. */
function mergeBorders(
  current: unknown,
  patch: unknown,
  sides: readonly string[],
): Record<string, BorderLine | null> | null | string {
  if (patch === null) return null
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) {
    return `borders must be an object with keys ${sides.join('/')}, or null to clear`
  }
  const next: Record<string, BorderLine | null> = {
    ...((current && typeof current === 'object' ? current : {}) as Record<string, BorderLine | null>),
  }
  for (const key of Object.keys(patch as Obj)) {
    if (!sides.includes(key)) return `unknown border side "${key}"; use ${sides.join('/')}`
    const line = borderLine((patch as Obj)[key])
    if (typeof line === 'string') return line
    next[key] = line
  }
  return next
}

// ── table / cell addressing ──

interface AddressableTable {
  /** docTable node */
  node: PmNode
  /** doc position before the table node */
  pos: number
  /** doc position of the table's content start (node pos + 1) */
  start: number
  map: TableMap
  /** number of rows */
  rows: number
  /** number of grid columns (merged spans included) */
  cols: number
}

type TableLookup = { table: AddressableTable } | { error: string }

function resolveTable(editor: Editor, input: Obj): TableLookup {
  const doc = editor.state.doc
  const blockIndex = intOf(input.blockIndex)
  if (blockIndex === null || blockIndex < 0 || blockIndex >= doc.childCount) {
    return { error: `blockIndex must be an integer 0..${doc.childCount - 1} (a table block from the block list)` }
  }
  const node = doc.child(blockIndex)
  if (node.type.name !== 'docTable') {
    const tables: number[] = []
    doc.forEach((child, _offset, index) => {
      if (child.type.name === 'docTable') tables.push(index)
    })
    return {
      error:
        `block ${blockIndex} is a ${node.type.name}, not a table`
        + (tables.length ? `; table blocks: ${tables.join(', ')}` : '; the document has no table'),
    }
  }
  let pos = 0
  for (let i = 0; i < blockIndex; i++) pos += doc.child(i).nodeSize
  const map = TableMap.get(node)
  return { table: { node, pos, start: pos + 1, map, rows: node.childCount, cols: map.width } }
}

function resolveCell(
  table: AddressableTable,
  input: Obj,
): { row: number; col: number } | { error: string } {
  const row = intOf(input.row)
  const col = intOf(input.col)
  if (row === null || row < 0 || row >= table.rows) {
    return { error: `row must be an integer 0..${table.rows - 1}` }
  }
  if (col === null || col < 0 || col >= table.cols) {
    return { error: `col must be an integer 0..${table.cols - 1} (grid columns; a merged cell covers its span)` }
  }
  return { row, col }
}

/** Doc position of the cell node at (row, col). */
function cellDocPos(table: AddressableTable, row: number, col: number): number {
  return table.start + table.map.positionAt(row, col, table.node)
}

/** Put a synthetic text cursor inside the target cell (selectedRect()/isInTable rely on it). */
function selectCell(editor: Editor, table: AddressableTable, row: number, col: number): void {
  const cellPos = cellDocPos(table, row, col)
  const $inside = editor.state.doc.resolve(Math.min(cellPos + 1, editor.state.doc.content.size))
  const selection = TextSelection.near($inside, 1)
  editor.view.dispatch(editor.state.tr.setSelection(selection))
}

function runTableCommand(editor: Editor, command: Command): boolean {
  return command(editor.state, editor.view.dispatch)
}

/** Re-resolve the table after a dispatch (positions are stable, but never assume). */
function rereadTable(editor: Editor, blockIndex: number): AddressableTable | null {
  const lookup = resolveTable(editor, { blockIndex })
  return 'table' in lookup ? lookup.table : null
}

/**
 * Insert row/column position helper: `at` is the reference row/col, `position`
 * says whether the new band goes before or after it. Omitted reference = last band.
 */
function referencePosition(
  table: AddressableTable,
  input: Obj,
  axis: 'row' | 'col',
): { index: number; before: boolean } | { error: string } {
  const limit = axis === 'row' ? table.rows : table.cols
  const fallback = limit - 1
  const raw = axis === 'row' ? input.row : input.col
  const index = raw === undefined ? fallback : intOf(raw)
  if (index === null || index < 0 || index >= limit) {
    return { error: `${axis} must be an integer 0..${limit - 1}` }
  }
  const position = input.position === undefined ? 'after' : String(input.position)
  if (position !== 'before' && position !== 'after') {
    return { error: 'position must be "before" or "after"' }
  }
  return { index, before: position === 'before' }
}

function tableAddress(blockIndex: number): { blockIndex: number } {
  return { blockIndex }
}

// ── table.insert ──

function insertTable(input: Obj, ctx: OperationContext): OperationOutcome {
  const doc = ctx.editor.state.doc
  const afterRaw = input.afterBlockIndex === undefined ? -1 : intOf(input.afterBlockIndex)
  if (afterRaw === null || afterRaw < -1 || afterRaw >= doc.childCount) {
    return failure(`afterBlockIndex must be -1 (end of document) or 0..${doc.childCount - 1}`)
  }
  const rows = clampInt(input.rows, 2, 1, MAX_INSERT_ROWS)
  const cols = clampInt(input.cols, 2, 1, MAX_INSERT_COLS)
  let data: string[][] | null = null
  if (input.data !== undefined) {
    if (!Array.isArray(input.data)) return failure('data must be a 2-D array of strings')
    data = input.data.slice(0, rows).map((row) =>
      Array.isArray(row)
        ? row.slice(0, cols).map((cell) =>
            typeof cell === 'string' ? cell.slice(0, MAX_CELL_TEXT) : cell == null ? '' : String(cell).slice(0, MAX_CELL_TEXT),
          )
        : [],
    )
  }
  const headerRow = input.headerRow === undefined ? true : input.headerRow !== false
  const line = { style: 'single', szEighths: 4, color: 'auto' }
  const model: TableModel = {
    rows: Array.from({ length: rows }, (_r, r) =>
      Array.from({ length: cols }, (_c, c): TableCell => {
        const text = data?.[r]?.[c] ?? ''
        const header = headerRow && r === 0
        return {
          paras: [text],
          ...(header ? { bold: true, fill: TABLE_HEADER_FILL } : {}),
        }
      }),
    ),
    colWidthsPct: Array.from({ length: cols }, () => 100 / cols),
    borders: {
      top: line,
      right: line,
      bottom: line,
      left: line,
      insideH: line,
      insideV: line,
    } as TableBorders,
  }
  const atEnd = afterRaw < 0 || afterRaw >= doc.childCount - 1
  const pos = atEnd ? doc.content.size : blockEndPos(doc, afterRaw)
  const newBlockIndex = atEnd ? doc.childCount : afterRaw + 1
  if (ctx.dryRun) {
    return planned(
      `would insert a ${rows}×${cols} table ${atEnd ? 'at the end of the document' : `after block ${afterRaw}`}`,
      tableAddress(newBlockIndex),
    )
  }
  // tableModelToPmNode builds the editor's JSON node shape; materialize it
  // against the live schema so a raw transaction can insert it
  const node = ctx.editor.schema.nodeFromJSON(tableModelToPmNode(model))
  ctx.editor.view.dispatch(ctx.editor.state.tr.insert(pos, node))
  return success(
    `Inserted a ${rows}×${cols} table ${atEnd ? 'at the end of the document' : `after block ${afterRaw}`} (new block ${newBlockIndex}).`,
    'insert table',
    { blockIndex: newBlockIndex, rows, cols },
  )
}

// ── structural ops ──

function deleteTableOp(input: Obj, ctx: OperationContext): OperationOutcome {
  const lookup = resolveTable(ctx.editor, input)
  if ('error' in lookup) return failure(lookup.error)
  const table = lookup.table
  const blockIndex = intOf(input.blockIndex) as number
  if (ctx.dryRun) return planned(`would delete table at block ${blockIndex}`, tableAddress(blockIndex))
  ctx.editor.view.dispatch(ctx.editor.state.tr.delete(table.pos, table.pos + table.node.nodeSize))
  return success(`Deleted the table at block ${blockIndex}.`, 'delete table', tableAddress(blockIndex))
}

function insertRowOp(input: Obj, ctx: OperationContext): OperationOutcome {
  const lookup = resolveTable(ctx.editor, input)
  if ('error' in lookup) return failure(lookup.error)
  const table = lookup.table
  const ref = referencePosition(table, input, 'row')
  if ('error' in ref) return failure(ref.error)
  const blockIndex = intOf(input.blockIndex) as number
  const where = ref.before ? 'before' : 'after'
  if (ctx.dryRun) {
    return planned(`would insert a row ${where} row ${ref.index} (table block ${blockIndex})`)
  }
  selectCell(ctx.editor, table, ref.index, 0)
  const ok = runTableCommand(ctx.editor, ref.before ? addRowBefore : addRowAfter)
  if (!ok) return failure('row insert failed (the selection is not inside the table)')
  return success(
    `Inserted a row ${where} row ${ref.index}; the table now has ${table.rows + 1} rows.`,
    'insert table row',
    { blockIndex, row: ref.before ? ref.index : ref.index + 1 },
  )
}

function deleteRowOp(input: Obj, ctx: OperationContext): OperationOutcome {
  const lookup = resolveTable(ctx.editor, input)
  if ('error' in lookup) return failure(lookup.error)
  const table = lookup.table
  const row = intOf(input.row)
  if (row === null || row < 0 || row >= table.rows) {
    return failure(`row must be an integer 0..${table.rows - 1}`)
  }
  const blockIndex = intOf(input.blockIndex) as number
  if (ctx.dryRun) return planned(`would delete row ${row} from table block ${blockIndex}`)
  selectCell(ctx.editor, table, row, 0)
  const ok = runTableCommand(ctx.editor, deleteRow)
  if (!ok) {
    return failure(
      table.rows <= 1 ? 'cannot delete the only row; delete the table instead' : 'row delete failed',
    )
  }
  return success(
    `Deleted row ${row}; the table now has ${table.rows - 1} rows.`,
    'delete table row',
    { blockIndex, row },
  )
}

function insertColumnOp(input: Obj, ctx: OperationContext): OperationOutcome {
  const lookup = resolveTable(ctx.editor, input)
  if ('error' in lookup) return failure(lookup.error)
  const table = lookup.table
  const ref = referencePosition(table, input, 'col')
  if ('error' in ref) return failure(ref.error)
  const blockIndex = intOf(input.blockIndex) as number
  const where = ref.before ? 'before' : 'after'
  if (ctx.dryRun) {
    return planned(`would insert a column ${where} column ${ref.index} (table block ${blockIndex})`)
  }
  selectCell(ctx.editor, table, 0, ref.index)
  const ok = runTableCommand(ctx.editor, ref.before ? addColumnBefore : addColumnAfter)
  if (!ok) return failure('column insert failed (the selection is not inside the table)')
  return success(
    `Inserted a column ${where} column ${ref.index}; the table now has ${table.cols + 1} columns.`,
    'insert table column',
    { blockIndex, col: ref.before ? ref.index : ref.index + 1 },
  )
}

function deleteColumnOp(input: Obj, ctx: OperationContext): OperationOutcome {
  const lookup = resolveTable(ctx.editor, input)
  if ('error' in lookup) return failure(lookup.error)
  const table = lookup.table
  const col = intOf(input.col)
  if (col === null || col < 0 || col >= table.cols) {
    return failure(`col must be an integer 0..${table.cols - 1}`)
  }
  const blockIndex = intOf(input.blockIndex) as number
  if (ctx.dryRun) return planned(`would delete column ${col} from table block ${blockIndex}`)
  selectCell(ctx.editor, table, 0, col)
  const ok = runTableCommand(ctx.editor, deleteColumn)
  if (!ok) {
    return failure(
      table.cols <= 1
        ? 'cannot delete the only column; delete the table instead'
        : 'column delete failed',
    )
  }
  return success(
    `Deleted column ${col}; the table now has ${table.cols - 1} columns.`,
    'delete table column',
    { blockIndex, col },
  )
}

// ── merge / split ──

function mergeCellsOp(input: Obj, ctx: OperationContext): OperationOutcome {
  const lookup = resolveTable(ctx.editor, input)
  if ('error' in lookup) return failure(lookup.error)
  const table = lookup.table
  const cell = resolveCell(table, input)
  if ('error' in cell) return failure(cell.error)
  const rowSpan = input.rowspan === undefined ? 1 : intOf(input.rowspan)
  const colSpan = input.colspan === undefined ? 1 : intOf(input.colspan)
  if (rowSpan === null || rowSpan < 1 || rowSpan > table.rows) {
    return failure(`rowspan must be an integer 1..${table.rows}`)
  }
  if (colSpan === null || colSpan < 1 || colSpan > table.cols) {
    return failure(`colspan must be an integer 1..${table.cols}`)
  }
  const bottom = cell.row + rowSpan - 1
  const right = cell.col + colSpan - 1
  if (bottom >= table.rows || right >= table.cols) {
    return failure(
      `merge range out of bounds: rows ${cell.row}..${bottom} (table has ${table.rows}), `
      + `cols ${cell.col}..${right} (table has ${table.cols})`,
    )
  }
  if (rowSpan * colSpan < 2) return failure('merge needs at least two cells (rowspan*colspan > 1)')
  const blockIndex = intOf(input.blockIndex) as number
  if (ctx.dryRun) {
    return planned(
      `would merge rows ${cell.row}..${bottom} × cols ${cell.col}..${right} into one cell`,
      { blockIndex, row: cell.row, col: cell.col, rowspan: rowSpan, colspan: colSpan },
    )
  }
  const doc = ctx.editor.state.doc
  const from = doc.resolve(cellDocPos(table, cell.row, cell.col))
  const to = doc.resolve(cellDocPos(table, bottom, right))
  ctx.editor.view.dispatch(ctx.editor.state.tr.setSelection(new CellSelection(from, to)))
  const ok = runTableCommand(ctx.editor, mergeCells)
  if (!ok) return failure('merge failed: the selected range does not form a mergeable rectangle')
  return success(
    `Merged rows ${cell.row}..${bottom} × cols ${cell.col}..${right} into one cell.`,
    'merge table cells',
    { blockIndex, row: cell.row, col: cell.col, rowspan: rowSpan, colspan: colSpan },
  )
}

function splitCellOp(input: Obj, ctx: OperationContext): OperationOutcome {
  const lookup = resolveTable(ctx.editor, input)
  if ('error' in lookup) return failure(lookup.error)
  const table = lookup.table
  const cell = resolveCell(table, input)
  if ('error' in cell) return failure(cell.error)
  const blockIndex = intOf(input.blockIndex) as number
  const pos = cellDocPos(table, cell.row, cell.col)
  const node = ctx.editor.state.doc.nodeAt(pos)
  if (!node) return failure('cell not found')
  if (node.attrs.colspan === 1 && node.attrs.rowspan === 1) {
    return failure(`cell (${cell.row}, ${cell.col}) is not merged; nothing to split`)
  }
  if (ctx.dryRun) {
    return planned(
      `would split the merged cell at row ${cell.row}, col ${cell.col}`
      + ` (${node.attrs.rowspan}×${node.attrs.colspan} span)`,
      { blockIndex, row: cell.row, col: cell.col },
    )
  }
  selectCell(ctx.editor, table, cell.row, cell.col)
  const ok = runTableCommand(ctx.editor, splitCell)
  if (!ok) return failure('split failed: the cell is not merged')
  return success(
    `Split the merged cell at row ${cell.row}, col ${cell.col}.`,
    'split table cell',
    { blockIndex, row: cell.row, col: cell.col },
  )
}

// ── cell formatting ──

function setCellOp(input: Obj, ctx: OperationContext): OperationOutcome {
  const lookup = resolveTable(ctx.editor, input)
  if ('error' in lookup) return failure(lookup.error)
  const table = lookup.table
  const cell = resolveCell(table, input)
  if ('error' in cell) return failure(cell.error)
  const blockIndex = intOf(input.blockIndex) as number
  const pos = cellDocPos(table, cell.row, cell.col)
  const node = ctx.editor.state.doc.nodeAt(pos)
  if (!node) return failure('cell not found')

  const patch: Obj = {}
  const changed: string[] = []
  const hasText = input.text !== undefined
  const text = hasText ? (input.text === null ? '' : String(input.text)) : ''
  if (hasText) {
    if (text.length > MAX_CELL_TEXT) {
      return failure(`text must be at most ${MAX_CELL_TEXT} characters`)
    }
    changed.push('text')
  }
  if (input.shading !== undefined) {
    const fill = hexOf(input.shading)
    if (fill === undefined) return failure('shading must be 6-digit hex (no #) or null to clear')
    patch.fill = fill
    changed.push('shading')
  }
  if (input.vAlign !== undefined) {
    const vAlign = input.vAlign === null ? null : String(input.vAlign)
    if (vAlign !== null && vAlign !== 'top' && vAlign !== 'center' && vAlign !== 'bottom') {
      return failure('vAlign must be "top", "center", "bottom" or null')
    }
    patch.vAlign = vAlign
    changed.push('vAlign')
  }
  if (input.width !== undefined) {
    if (input.width === null) {
      patch.colwidth = null
    } else {
      const widthPx = Number(input.width)
      if (!Number.isFinite(widthPx) || widthPx < 24 || widthPx > 4000) {
        return failure('width must be a number of px 24..4000, or null to clear')
      }
      const span = Math.max(1, Number(node.attrs.colspan) || 1)
      patch.colwidth = Array.from({ length: span }, () =>
        Math.max(24, Math.round(widthPx / span)),
      )
    }
    changed.push('width')
  }
  if (input.borders !== undefined) {
    const borders = mergeBorders(node.attrs.borders, input.borders, [
      'top',
      'left',
      'bottom',
      'right',
      'insideH',
      'insideV',
    ])
    if (typeof borders === 'string') return failure(borders)
    patch.borders = borders
    changed.push('borders')
  }
  if (changed.length === 0) {
    return failure('nothing to change: pass text, shading, borders, vAlign and/or width')
  }
  if (ctx.dryRun) {
    return planned(
      `would set ${changed.join(', ')} on cell (${cell.row}, ${cell.col}) of table block ${blockIndex}`,
      { blockIndex, row: cell.row, col: cell.col },
    )
  }

  const editor = ctx.editor
  let tr = editor.state.tr
  if (hasText) {
    const schema = editor.schema
    const paragraph = schema.nodes.docParagraph.create(
      {},
      text ? schema.text(text) : undefined,
    )
    tr = tr.replaceWith(pos + 1, pos + node.nodeSize - 1, paragraph)
  }
  tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...patch })
  editor.view.dispatch(tr)
  return success(
    `Set ${changed.join(', ')} on cell (${cell.row}, ${cell.col}) of table block ${blockIndex}.`,
    'format table cell',
    { blockIndex, row: cell.row, col: cell.col, changed },
  )
}

// ── table properties ──

const STYLE_PRESETS: Record<string, TablePreset> = {
  plain: { headerFill: null, band1Fill: null, band2Fill: null, borderColor: '7F7F7F' },
  grid: { headerFill: 'D9E2F3', band1Fill: null, band2Fill: null, borderColor: '7F7F7F' },
  banded: { headerFill: 'D9E2F3', band1Fill: 'F2F2F2', band2Fill: null, borderColor: 'BFBFBF' },
}

function setPropertiesOp(input: Obj, ctx: OperationContext): OperationOutcome {
  const lookup = resolveTable(ctx.editor, input)
  if ('error' in lookup) return failure(lookup.error)
  const table = lookup.table
  const blockIndex = intOf(input.blockIndex) as number
  const editor = ctx.editor
  const doc = editor.state.doc
  const tablePatch: Obj = {}
  const rowPatch = new Map<number, Obj>()
  const changed: string[] = []

  if (input.alignment !== undefined) {
    const align = input.alignment === null ? null : String(input.alignment)
    if (align !== null && align !== 'left' && align !== 'center' && align !== 'right') {
      return failure('alignment must be "left", "center", "right" or null')
    }
    tablePatch.tblAlign = align
    changed.push('alignment')
  }
  if (input.indent !== undefined) {
    if (input.indent === null) {
      tablePatch.indentTwips = null
    } else {
      const indent = Number(input.indent)
      if (!Number.isFinite(indent) || indent < -7200 || indent > 31680) {
        return failure('indent must be twips between -7200 and 31680 (1 cm = 567), or null to clear')
      }
      tablePatch.indentTwips = Math.round(indent)
    }
    changed.push('indent')
  }
  if (input.rowHeight !== undefined) {
    if (input.rowHeight === null) {
      for (let row = 0; row < table.rows; row++) {
        rowPatch.set(row, { heightTwips: null, heightRule: null })
      }
    } else {
      if (typeof input.rowHeight !== 'object') {
        return failure('rowHeight must be { twips, rule?, row? } or null to clear')
      }
      const raw = input.rowHeight as Obj
      const twips = intOf(raw.twips)
      if (twips === null || twips < 0 || twips > 100_000) {
        return failure('rowHeight.twips must be an integer 0..100000')
      }
      const rule = raw.rule === undefined ? 'atLeast' : String(raw.rule)
      if (rule !== 'atLeast' && rule !== 'exact') {
        return failure('rowHeight.rule must be "atLeast" or "exact"')
      }
      const row = raw.row === undefined ? null : intOf(raw.row)
      if (row !== null && (row < 0 || row >= table.rows)) {
        return failure(`rowHeight.row must be an integer 0..${table.rows - 1}`)
      }
      const rows = row === null ? Array.from({ length: table.rows }, (_v, i) => i) : [row]
      for (const index of rows) rowPatch.set(index, { heightTwips: twips, heightRule: rule })
    }
    changed.push('rowHeight')
  }
  if (input.repeatHeaderRows !== undefined) {
    if (typeof input.repeatHeaderRows !== 'boolean' && intOf(input.repeatHeaderRows) === null) {
      return failure('repeatHeaderRows must be a boolean or the number of leading header rows')
    }
    const count =
      input.repeatHeaderRows === true
        ? 1
        : input.repeatHeaderRows === false
          ? 0
          : (intOf(input.repeatHeaderRows) as number)
    if (count < 0 || count > table.rows) {
      return failure(`repeatHeaderRows must be 0..${table.rows}`)
    }
    for (let row = 0; row < table.rows; row++) {
      rowPatch.set(row, {
        ...(rowPatch.get(row) ?? {}),
        repeatHeader: row < count,
        repeatHeaderEdited: true,
      })
    }
    changed.push('repeatHeaderRows')
  }
  if (input.colWidths !== undefined) {
    if (input.colWidths === null) {
      tablePatch.colWidthsPct = null
    } else {
      if (!Array.isArray(input.colWidths) || input.colWidths.length !== table.cols) {
        return failure(`colWidths must be an array of ${table.cols} relative widths (percent)`)
      }
      const widths = input.colWidths.map((value) => Number(value))
      if (widths.some((width) => !Number.isFinite(width) || width <= 0)) {
        return failure('colWidths entries must be positive numbers')
      }
      const total = widths.reduce((sum, width) => sum + width, 0)
      tablePatch.colWidthsPct = widths.map((width) => Math.round((width / total) * 10000) / 100)
      tablePatch.tblAutoFit = 'fixed'
      tablePatch.tblAutoFitEdited = true
    }
    changed.push('colWidths')
  }
  if (input.borders !== undefined) {
    const borders = mergeBorders(table.node.attrs.borders, input.borders, [
      'top',
      'left',
      'bottom',
      'right',
      'insideH',
      'insideV',
    ])
    if (typeof borders === 'string') return failure(borders)
    tablePatch.borders = borders
    changed.push('borders')
  }
  const LOOK_KEYS = [
    'firstRow',
    'lastRow',
    'firstColumn',
    'lastColumn',
    'bandedRows',
    'bandedColumns',
  ] as const
  let lookPatch: Array<[keyof TableLook, boolean]> | null = null
  if (input.look !== undefined) {
    if (!input.look || typeof input.look !== 'object' || Array.isArray(input.look)) {
      return failure('look must be an object: { firstRow?, lastRow?, firstColumn?, lastColumn?, bandedRows?, bandedColumns? }')
    }
    const entries: Array<[keyof TableLook, boolean]> = []
    for (const key of Object.keys(input.look as Obj)) {
      if (!(LOOK_KEYS as readonly string[]).includes(key)) {
        return failure(`unknown look option "${key}"; allowed: ${LOOK_KEYS.join('/')}`)
      }
      const value = (input.look as Obj)[key]
      if (typeof value !== 'boolean') return failure(`look.${key} must be a boolean`)
      entries.push([key as keyof TableLook, value])
    }
    if (entries.length === 0) return failure('look must contain at least one option')
    lookPatch = entries
    changed.push('look')
  }
  let preset: TablePreset | null = null
  if (input.stylePreset !== undefined) {
    if (input.stylePreset === null) {
      preset = null
    } else {
      preset = STYLE_PRESETS[String(input.stylePreset)] ?? null
      if (!preset) {
        return failure(`stylePreset must be one of ${Object.keys(STYLE_PRESETS).join('/')}`)
      }
    }
    changed.push('stylePreset')
  }
  if (changed.length === 0) {
    return failure(
      'nothing to change: pass alignment, indent, rowHeight, colWidths, borders, '
      + 'repeatHeaderRows, stylePreset and/or look',
    )
  }
  if (ctx.dryRun) {
    return planned(
      `would set ${changed.join(', ')} on table block ${blockIndex}`,
      { blockIndex },
    )
  }

  // one transaction for every attribute-level change
  if (Object.keys(tablePatch).length > 0 || rowPatch.size > 0) {
    let tr = editor.state.tr
    if (rowPatch.size > 0) {
      let rowPos = table.start
      for (let row = 0; row < table.rows; row++) {
        const patch = rowPatch.get(row)
        if (patch) tr = tr.setNodeMarkup(rowPos, undefined, { ...table.node.child(row).attrs, ...patch })
        rowPos += table.node.child(row).nodeSize
      }
    }
    if (Object.keys(tablePatch).length > 0) {
      tr = tr.setNodeMarkup(table.pos, undefined, { ...table.node.attrs, ...tablePatch })
    }
    editor.view.dispatch(tr)
  }
  // style preset / fixed grid are editor commands that need a cursor inside the table
  let current = rereadTable(editor, blockIndex)
  if (!current) return failure('table disappeared while applying properties')
  if (preset) {
    selectCell(editor, current, 0, 0)
    const ok = runTableCommand(editor, applyTablePreset(preset))
    if (!ok) return failure('style preset failed: table is not selectable')
    current = rereadTable(editor, blockIndex)
    if (!current) return failure('table disappeared while applying the style preset')
  }
  if (lookPatch) {
    selectCell(editor, current, 0, 0)
    for (const [key, value] of lookPatch) {
      if (!runTableCommand(editor, setTableLookOption(key, value))) {
        return failure(`table look option "${key}" could not be set (table not selectable)`)
      }
    }
    current = rereadTable(editor, blockIndex)
    if (!current) return failure('table disappeared while applying the look options')
  }
  if (tablePatch.colWidthsPct !== undefined && tablePatch.colWidthsPct !== null) {
    selectCell(editor, current, 0, 0)
    const ok = runTableCommand(editor, setTableAutoFit('fixed', ctx.contentWidthPx))
    if (!ok) return failure('colWidths applied but rebuilding the fixed grid failed')
  }
  return success(
    `Set ${changed.join(', ')} on table block ${blockIndex}.`,
    'format table properties',
    { blockIndex, changed },
  )
}

function autoFitOp(input: Obj, ctx: OperationContext): OperationOutcome {
  const lookup = resolveTable(ctx.editor, input)
  if ('error' in lookup) return failure(lookup.error)
  const table = lookup.table
  const blockIndex = intOf(input.blockIndex) as number
  const mode = String(input.mode ?? '')
  if (mode !== 'window' && mode !== 'contents' && mode !== 'fixed') {
    return failure('mode must be "window" (fill the text column), "contents" (shrink to content) or "fixed"')
  }
  if (ctx.dryRun) {
    return planned(`would set AutoFit "${mode}" on table block ${blockIndex}`, { blockIndex })
  }
  selectCell(ctx.editor, table, 0, 0)
  const ok = runTableCommand(ctx.editor, setTableAutoFit(mode, ctx.contentWidthPx))
  if (!ok) return failure('AutoFit failed: the table is not selectable')
  return success(
    mode === 'window'
      ? 'Table now fills the text column (AutoFit to window).'
      : mode === 'contents'
        ? 'Table shrunk to its content width (AutoFit to contents).'
        : 'Table switched to a fixed grid layout.',
    'table autofit',
    { blockIndex, mode },
  )
}

// ── registry ──

const blockIndexOf = {
  blockIndex: {
    type: 'integer',
    minimum: 0,
    description: 'index of the table block in the document block list (the "table" block)',
  },
} as const

const rowCol = {
  row: { type: 'integer', minimum: 0, description: '0-based grid row' },
  col: { type: 'integer', minimum: 0, description: '0-based grid column' },
} as const

const borderShape = {
  type: ['object', 'null'],
  description: 'per side; null clears that side; omit to keep it',
  properties: {
    style: { type: 'string', enum: ['single', 'double', 'dashed', 'dotted', 'none'] },
    szEighths: { type: 'integer', minimum: 0, maximum: 96, description: 'thickness in eighths of a point (4 = 0.5pt)' },
    color: { type: 'string', description: '6-digit hex without #, or "auto"' },
  },
} as const

const bordersShape = {
  type: ['object', 'null'],
  description:
    'keys top/left/bottom/right/insideH/insideV, each { style, szEighths, color } or null; null clears all borders',
  properties: {
    top: borderShape,
    left: borderShape,
    bottom: borderShape,
    right: borderShape,
    insideH: borderShape,
    insideV: borderShape,
  },
} as const

export const tableOperations: OperationDef[] = [
  {
    name: 'table.insert',
    group: 'table',
    description:
      'table.insert { afterBlockIndex?, rows?, cols?, data?, headerRow? } — insert a new table '
      + '(afterBlockIndex -1/omitted = end of document; data = row-major cell texts; headerRow defaults to true)',
    inputSchema: {
      type: 'object',
      properties: {
        afterBlockIndex: {
          type: 'integer',
          minimum: -1,
          description: 'insert after this block index; -1/omitted = end of document',
        },
        rows: { type: 'integer', minimum: 1, maximum: MAX_INSERT_ROWS },
        cols: { type: 'integer', minimum: 1, maximum: MAX_INSERT_COLS },
        data: {
          type: 'array',
          description: 'row-major cell texts (plain text; truncated to rows×cols)',
          items: { type: 'array', items: { type: 'string', maxLength: MAX_CELL_TEXT } },
        },
        headerRow: {
          type: 'boolean',
          description: 'treat the first row as a header row (bold, header cell type); default true',
        },
      },
    },
    handler: insertTable,
  },
  {
    name: 'table.delete',
    group: 'table',
    description: 'table.delete { blockIndex } — delete the whole table block',
    inputSchema: { type: 'object', properties: { ...blockIndexOf }, required: ['blockIndex'] },
    handler: deleteTableOp,
  },
  {
    name: 'table.insertRow',
    group: 'table',
    description:
      'table.insertRow { blockIndex, row?, position? } — insert an empty row before/after row '
      + '(default: after the last row)',
    inputSchema: {
      type: 'object',
      properties: {
        ...blockIndexOf,
        row: { type: 'integer', minimum: 0, description: 'reference row (default: last row)' },
        position: { type: 'string', enum: ['before', 'after'], description: 'default after' },
      },
      required: ['blockIndex'],
    },
    handler: insertRowOp,
  },
  {
    name: 'table.deleteRow',
    group: 'table',
    description: 'table.deleteRow { blockIndex, row } — delete one grid row',
    inputSchema: {
      type: 'object',
      properties: { ...blockIndexOf, ...rowCol },
      required: ['blockIndex', 'row'],
    },
    handler: deleteRowOp,
  },
  {
    name: 'table.insertColumn',
    group: 'table',
    description:
      'table.insertColumn { blockIndex, col?, position? } — insert an empty column before/after column '
      + '(default: after the last column)',
    inputSchema: {
      type: 'object',
      properties: {
        ...blockIndexOf,
        col: { type: 'integer', minimum: 0, description: 'reference column (default: last column)' },
        position: { type: 'string', enum: ['before', 'after'], description: 'default after' },
      },
      required: ['blockIndex'],
    },
    handler: insertColumnOp,
  },
  {
    name: 'table.deleteColumn',
    group: 'table',
    description: 'table.deleteColumn { blockIndex, col } — delete one grid column',
    inputSchema: {
      type: 'object',
      properties: { ...blockIndexOf, ...rowCol },
      required: ['blockIndex', 'col'],
    },
    handler: deleteColumnOp,
  },
  {
    name: 'table.mergeCells',
    group: 'table',
    description:
      'table.mergeCells { blockIndex, row, col, rowspan?, colspan? } — merge a rectangular range of cells '
      + 'into the top-left cell (defaults rowspan/colspan = 1; at least one must be > 1)',
    inputSchema: {
      type: 'object',
      properties: {
        ...blockIndexOf,
        ...rowCol,
        rowspan: { type: 'integer', minimum: 1, description: 'number of rows to merge' },
        colspan: { type: 'integer', minimum: 1, description: 'number of columns to merge' },
      },
      required: ['blockIndex', 'row', 'col'],
    },
    handler: mergeCellsOp,
  },
  {
    name: 'table.splitCell',
    group: 'table',
    description: 'table.splitCell { blockIndex, row, col } — split a previously merged cell back into grid cells',
    inputSchema: {
      type: 'object',
      properties: { ...blockIndexOf, ...rowCol },
      required: ['blockIndex', 'row', 'col'],
    },
    handler: splitCellOp,
  },
  {
    name: 'table.setCell',
    group: 'table',
    description:
      'table.setCell { blockIndex, row, col, text?, shading?, borders?, vAlign?, width? } — set cell text '
      + 'and/or formatting (text replaces the whole cell content; width in px; shading 6-digit hex or null)',
    inputSchema: {
      type: 'object',
      properties: {
        ...blockIndexOf,
        ...rowCol,
        text: { type: ['string', 'null'], maxLength: MAX_CELL_TEXT, description: 'plain cell text; null = clear' },
        shading: {
          type: ['string', 'null'],
          description: 'cell fill: 6-digit hex without # (e.g. "FFF2CC"), null = no fill',
        },
        borders: bordersShape,
        vAlign: {
          type: ['string', 'null'],
          enum: ['top', 'center', 'bottom', null],
          description: 'vertical cell alignment; null = default',
        },
        width: {
          type: ['number', 'null'],
          minimum: 24,
          maximum: 4000,
          description: 'preferred cell width in px; null = clear',
        },
      },
      required: ['blockIndex', 'row', 'col'],
    },
    handler: setCellOp,
  },
  {
    name: 'table.setProperties',
    group: 'table',
    description:
      'table.setProperties { blockIndex, alignment?, indent?, rowHeight?, colWidths?, borders?, '
      + 'repeatHeaderRows?, stylePreset?, look? } — table-level layout and look '
      + '(look = the Table Styles options: firstRow/lastRow/firstColumn/lastColumn/bandedRows/bandedColumns)',
    inputSchema: {
      type: 'object',
      properties: {
        ...blockIndexOf,
        alignment: {
          type: ['string', 'null'],
          enum: ['left', 'center', 'right', null],
          description: 'table horizontal alignment; null = default left',
        },
        indent: {
          type: ['number', 'null'],
          minimum: -7200,
          maximum: 31680,
          description: 'left indent in twips (1 cm = 567); null = clear',
        },
        rowHeight: {
          type: ['object', 'null'],
          description:
            'fixed/minimum row height; row omitted = all rows; null = clear all row heights',
          properties: {
            twips: { type: 'integer', minimum: 0, maximum: 100000 },
            rule: { type: 'string', enum: ['atLeast', 'exact'] },
            row: { type: 'integer', minimum: 0 },
          },
          required: ['twips'],
        },
        colWidths: {
          type: ['array', 'null'],
          description:
            'relative column widths (percent-like weights, one entry per grid column); null = clear',
          items: { type: 'number', minimum: 0.01 },
        },
        borders: bordersShape,
        repeatHeaderRows: {
          type: ['boolean', 'integer'],
          description: 'true = make row 0 a repeating header; false = clear; integer = count of leading header rows',
        },
        stylePreset: {
          type: ['string', 'null'],
          enum: ['plain', 'grid', 'banded', null],
          description: 'self-contained look preset (border color + header/band fills)',
        },
        look: {
          type: 'object',
          description: 'Table Styles options (w:tblLook); only the passed flags change',
          properties: {
            firstRow: { type: 'boolean' },
            lastRow: { type: 'boolean' },
            firstColumn: { type: 'boolean' },
            lastColumn: { type: 'boolean' },
            bandedRows: { type: 'boolean' },
            bandedColumns: { type: 'boolean' },
          },
        },
      },
      required: ['blockIndex'],
    },
    handler: setPropertiesOp,
  },
  {
    name: 'table.autoFit',
    group: 'table',
    description:
      'table.autoFit { blockIndex, mode } — mode "window" fills the text column, "contents" shrinks to '
      + 'content width, "fixed" keeps a fixed grid',
    inputSchema: {
      type: 'object',
      properties: {
        ...blockIndexOf,
        mode: { type: 'string', enum: ['window', 'contents', 'fixed'] },
      },
      required: ['blockIndex', 'mode'],
    },
    handler: autoFitOp,
  },
]

