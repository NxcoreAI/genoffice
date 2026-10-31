import { z } from 'zod'
import { ADDABLE_SHAPE_TYPES } from '../shared/shape-types'
import {
  columnIndex,
  columnLabel,
  formatAddress,
  parseAddress,
  parseRange,
  rangeCellCount,
} from './cell-address'
import { computeSortChanges, type SortSpec } from './sort-range'

const cellAddressSchema = z.string().regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}$/)
const cellRangeSchema = z.string().regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$/)
const columnLabelSchema = z.string().regex(/^[A-Z]{1,3}$/)
const sheetNameSchema = z
  .string()
  .trim()
  .min(1)
  .max(31)
  .refine((name) => !/[:\\/?*[\]]/.test(name))
const hexColorSchema = z.string().regex(/^#[0-9A-Fa-f]{6}$/)
const cellScalarSchema = z.union([z.string(), z.number().finite(), z.boolean(), z.null()])

const setCellSchema = z.object({
  op: z.literal('set_cell'),
  sheetId: z.string().min(1),
  address: cellAddressSchema,
  value: cellScalarSchema,
  expectedValue: cellScalarSchema.optional(),
})

const setFormulaSchema = z.object({
  op: z.literal('set_formula'),
  sheetId: z.string().min(1),
  address: cellAddressSchema,
  formula: z.string().startsWith('=').max(8192),
  expectedFormula: z.string().optional(),
})

const clearCellSchema = z.object({
  op: z.literal('clear_cell'),
  sheetId: z.string().min(1),
  address: cellAddressSchema,
})

// String values starting with "=" are treated as formulas, matching what
// typing the same text into the cell editor would do.
// `start` (top-left target cell) is canonical; `range` is also accepted
// because every other range-shaped op uses that field name and models keep
// reaching for it — when given, its size must match the values grid, which
// doubles as a misaligned-write check (validated in expandToPrimitiveOps).
const setRangeSchema = z.object({
  op: z.literal('set_range'),
  sheetId: z.string().min(1),
  start: cellAddressSchema.optional(),
  range: cellRangeSchema.optional(),
  values: z.array(z.array(cellScalarSchema).min(1).max(100)).min(1).max(500),
})

const clearRangeSchema = z.object({
  op: z.literal('clear_range'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
})

// Fill/copy (Excel's fill handle as one operation): the source block's values
// AND formulas tile across the target; relative references shift by each
// copy's offset, $-anchored axes stay pinned. This is the bulk path for
// "fill this formula down the whole column" — set_range would need the
// expanded values spelled out cell by cell, fill_range does not.
const fillRangeSchema = z.object({
  op: z.literal('fill_range'),
  sheetId: z.string().min(1),
  /** block to copy: single cell or rectangle (≤2000 cells), e.g. "A2" / "A2:C2" */
  source: cellRangeSchema,
  /** sheet the source lives on; defaults to the target sheet */
  sourceSheetId: z.string().min(1).optional(),
  /** range to fill; each dimension must be a whole multiple of the source's */
  target: cellRangeSchema,
})

// Copy one block to one destination (Excel copy → paste as one operation):
// values AND formulas copy, relative references shift by the block's offset,
// $-anchored axes stay pinned — exactly like pasting. Unlike fill_range
// (small source tiled across a big target), copy_range moves one block of up
// to 200,000 cells exactly once (duplicate a table, move a column's data).
const copyRangeSchema = z.object({
  op: z.literal('copy_range'),
  sheetId: z.string().min(1),
  /** block to copy, e.g. "A1:F5000" (up to 200,000 cells) */
  source: cellRangeSchema,
  /** sheet the source lives on; defaults to the target sheet */
  sourceSheetId: z.string().min(1).optional(),
  /** destination: its top-left cell, or a range exactly the source's size */
  target: cellRangeSchema,
})

// Freezes formulas into their current computed values (Excel's copy →
// paste-values onto the same cells). Non-formula cells and all formatting
// stay untouched.
const convertToValuesSchema = z.object({
  op: z.literal('convert_to_values'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
})

const insertRowsSchema = z.object({
  op: z.literal('insert_rows'),
  sheetId: z.string().min(1),
  /** 1-based; new rows are inserted before this row */
  row: z.number().int().min(1).max(9999999),
  count: z.number().int().min(1).max(500),
})

const deleteRowsSchema = z.object({
  op: z.literal('delete_rows'),
  sheetId: z.string().min(1),
  /** 1-based first row to delete */
  row: z.number().int().min(1).max(9999999),
  count: z.number().int().min(1).max(500),
})

const insertColsSchema = z.object({
  op: z.literal('insert_cols'),
  sheetId: z.string().min(1),
  /** new columns are inserted before this column */
  column: columnLabelSchema,
  count: z.number().int().min(1).max(100),
})

const deleteColsSchema = z.object({
  op: z.literal('delete_cols'),
  sheetId: z.string().min(1),
  column: columnLabelSchema,
  count: z.number().int().min(1).max(100),
})

const addSheetSchema = z.object({
  op: z.literal('add_sheet'),
  name: sheetNameSchema,
})

const deleteSheetSchema = z.object({
  op: z.literal('delete_sheet'),
  sheetId: z.string().min(1),
})

// Edits an EXISTING chart (charts are listed in get_workbook_context):
// file charts by chart part path, session-added and demo charts by their
// visual id. At least one property required — checked at expansion because
// discriminated-union members cannot carry refinements.
const editChartSchema = z.object({
  op: z.literal('edit_chart'),
  chartPath: z
    .string()
    .regex(/^(xl\/charts\/[A-Za-z0-9._-]+\.xml|(added|demo)-chart-[A-Za-z0-9._-]+)$/),
  title: z.string().max(255).optional(),
  chartType: z.enum(['column', 'bar', 'line', 'area', 'pie', 'doughnut']).optional(),
  /** series index ("0"-based, as string) → #RRGGBB */
  seriesColors: z.record(z.string().regex(/^[0-9]{1,3}$/), hexColorSchema).optional(),
  /** 'none' hides the legend; a side moves it (creating it if absent) */
  legend: z.enum(['none', 'right', 'bottom', 'top', 'left']).optional(),
  /** plot-level data labels; 'none' removes them */
  dataLabels: z.enum(['none', 'value', 'percent', 'category-percent']).optional(),
  /** pie/doughnut slice fills: series index → point index → #RRGGBB */
  pointColors: z
    .record(
      z.string().regex(/^[0-9]{1,3}$/),
      z.record(z.string().regex(/^[0-9]{1,3}$/), hexColorSchema),
    )
    .optional(),
  /** data-label placement (c:dLblPos) */
  dataLabelPosition: z.enum(['center', 'inside-end', 'outside-end']).optional(),
  /** number format for data labels, e.g. "#,##0.0" or "0%" */
  dataLabelFormat: z.string().min(1).max(64).optional(),
  /** bar/line/area stacking; 'clustered' restores side-by-side */
  grouping: z.enum(['clustered', 'stacked', 'percentStacked']).optional(),
  /** value-axis major gridlines on/off (axis-based charts only) */
  gridlines: z.boolean().optional(),
  /** explicit value-axis bounds; null resets that bound to auto */
  valueAxis: z
    .object({
      min: z.number().finite().nullable().optional(),
      max: z.number().finite().nullable().optional(),
    })
    .refine((axis) => axis.min !== undefined || axis.max !== undefined, {
      message: 'valueAxis needs min or max',
    })
    .optional(),
  /** axis-based charts only; null removes that axis title */
  axisTitles: z
    .object({
      category: z.string().max(255).nullable().optional(),
      value: z.string().max(255).nullable().optional(),
    })
    .optional(),
  /** rename a series and/or repoint its data at new single-row/column ranges */
  seriesData: z
    .array(
      z.object({
        index: z.number().int().min(0).max(255),
        name: z.string().max(255).optional(),
        valuesRange: cellRangeSchema.optional(),
        categoriesRange: cellRangeSchema.optional(),
        /** sheet the ranges live on; default: the chart's own sheet */
        sheetId: z.string().min(1).optional(),
      }),
    )
    .min(1)
    .max(24)
    .optional(),
})

// Creates a NEW chart from a data range (imported workbooks only). The range
// may include a header row and a leading category column — both are detected
// from the data, matching the ribbon's Insert Chart.
const addChartSchema = z.object({
  op: z.literal('add_chart'),
  sheetId: z.string().min(1),
  chartType: z.enum([
    'column',
    'bar',
    'line',
    'area',
    'pie',
    'doughnut',
    'scatter',
    'radar',
    'combo',
  ]),
  dataRange: cellRangeSchema,
  title: z.string().max(255).optional(),
  /** top-left cell of the chart frame; default: two columns right of the data */
  anchorCell: cellAddressSchema.optional(),
})

const addShapeSchema = z.object({
  op: z.literal('add_shape'),
  sheetId: z.string().min(1),
  shapeType: z.enum([...ADDABLE_SHAPE_TYPES, 'textbox']),
  anchorCell: cellAddressSchema,
  fillColor: hexColorSchema.optional(),
  text: z.string().max(1000).optional(),
})

// Edits a shape/text box ADDED THIS SESSION (ids listed by
// read_sheet_features); file-original drawings stay read-only. At least one
// property required — checked at expansion.
const editShapeSchema = z.object({
  op: z.literal('edit_shape'),
  visualId: z.string().min(1).max(128),
  text: z.string().max(1000).optional(),
  fillColor: hexColorSchema.optional(),
  /** moves the shape's top-left corner, keeping its size */
  anchorCell: cellAddressSchema.optional(),
})

// Inserts an image (PNG/JPEG/GIF): a local file (≤20MB) the user pointed at,
// or an https URL from image_search / generate_image.
const addImageSchema = z.object({
  op: z.literal('add_image'),
  sheetId: z.string().min(1),
  /** absolute path on this machine (~/ allowed) or an http(s) image URL */
  path: z.string().min(1).max(2048),
  anchorCell: cellAddressSchema,
})

// Creates a real Excel table (ListObject) over a data range (imported
// workbooks only). The range's first row must hold the column headers —
// non-empty and unique; they become the table's column names in the file.
const addTableSchema = z.object({
  op: z.literal('add_table'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  /** Excel table name (defined-name rules: letter/underscore start, no spaces); default Table1, Table2, … */
  name: z
    .string()
    .regex(/^[\p{L}_\\][\p{L}\p{N}_.\\]{0,254}$/u)
    .optional(),
  /** built-in table style, e.g. "TableStyleMedium2" (the default) */
  style: z
    .string()
    .regex(/^TableStyle(?:Light|Medium|Dark)[1-9][0-9]?$/)
    .optional(),
  /** banded (striped) rows — Excel's default is on */
  bandedRows: z.boolean().optional(),
})

// Inserts blank rows into a session-created table. The table must have been
// created with add_table in the current session (fail-closed: file tables not
// supported yet — save and reopen first). Rows are inserted into the data area
// (not the header row). row is 1-based within the data area; omit to append.
const addTableRowSchema = z.object({
  op: z.literal('add_table_row'),
  sheetId: z.string().min(1),
  /** Table name (must match an add_table op from this session) */
  tableName: z.string().min(1),
  /** 1-based position within the data area to insert before; omit to append */
  row: z.number().int().min(1).optional(),
  count: z.number().int().min(1).max(1000).default(1),
})

// Inserts blank columns into a session-created table. Column is 1-based
// within the table area; omit to append. columnName must be non-blank and
// unique within the table.
const addTableColumnSchema = z.object({
  op: z.literal('add_table_column'),
  sheetId: z.string().min(1),
  tableName: z.string().min(1),
  /** 1-based column position within the table to insert before; omit to append */
  column: z.number().int().min(1).optional(),
  columnName: z.string().min(1).max(255),
  count: z.number().int().min(1).max(100).default(1),
})

// Deletes rows from the data area of a session-created table.
// row is 1-based within the data area; table must retain ≥1 data row.
const deleteTableRowSchema = z.object({
  op: z.literal('delete_table_row'),
  sheetId: z.string().min(1),
  tableName: z.string().min(1),
  /** 1-based first row to delete within the data area */
  row: z.number().int().min(1),
  count: z.number().int().min(1).max(1000).default(1),
})

// Deletes columns from a session-created table; table must retain ≥1 column.
const deleteTableColumnSchema = z.object({
  op: z.literal('delete_table_column'),
  sheetId: z.string().min(1),
  tableName: z.string().min(1),
  /** 1-based column to delete within the table */
  column: z.number().int().min(1),
  count: z.number().int().min(1).max(100).default(1),
})

// Creates a real PivotTable (imported workbooks only): the apply bakes the
// aggregated grid into the target cells, and the save writes native pivot
// parts with refreshOnLoad so Excel turns it into a live pivot on open.
// Constraints: 1–8 row dimension levels; 0–8 column dimension levels; with
// column field(s) exactly one value entry; up to 4 page (report filter) fields.
const addPivotSchema = z.object({
  op: z.literal('add_pivot'),
  sheetId: z.string().min(1),
  /** source data range, first row = field headers (non-blank, unique) */
  sourceRange: cellRangeSchema,
  /** top-left cell of the pivot output; must not overlap the source */
  targetCell: cellAddressSchema,
  /** sheet for the output; default: the source sheet */
  targetSheetId: z.string().min(1).optional(),
  /**
   * Row dimension field(s) — 1 to 8 headers from the source, outer levels first.
   * For a single level you may pass a bare string or a one-element array.
   * With multiple fields the pivot groups hierarchically and subtotal rows
   * are automatically inserted for every non-leaf level.
   */
  rowFields: z.union([
    z.string().min(1).max(255),
    z.array(z.string().min(1).max(255)).min(1).max(8),
  ]),
  /**
   * Column dimension field(s) to spread across columns — 1 to 8 headers,
   * Outer levels first. Single level may be a bare string; multiple levels take
   * an array, expanding members as a cartesian product in column-field order,
   * with a subtotal column appended after each non-leaf member.
   */
  columnField: z
    .union([z.string().min(1).max(255), z.array(z.string().min(1).max(255)).min(1).max(8)])
    .optional(),
  /**
   * Report filter fields (pageFields) — up to 4 source headers that are
   * placed above the pivot as filter drop-downs in the saved Excel file.
   * The baked grid omits the filter row; Excel/LibreOffice shows them on open.
   */
  pageFields: z.array(z.string().min(1).max(255)).max(4).optional(),
  values: z
    .array(
      z.object({
        field: z.string().min(1).max(255),
        agg: z.enum(['sum', 'count', 'average', 'max', 'min']),
        /** Optional Excel number format string, e.g. "#,##0.00" or "0%" */
        numFmt: z.string().min(1).max(255).optional(),
        /**
         * "Show values as" mode: percentOfTotal = percent of grand total /
         * percentOfRow = percent of row total / percentOfCol = percent of column
         * total; defaults to the plain aggregate value.
         */
        showDataAs: z.enum(['percentOfTotal', 'percentOfRow', 'percentOfCol']).optional(),
        /**
         * Calculated field: when formula is present, field is the new data field's
         * name (must not clash with source headers); the formula does basic
         * arithmetic over other source field names (e.g. "Revenue-Cost"; wrap names
         * containing spaces in single quotes); each group sums the referenced fields
         * first, then evaluates the formula, and agg must be sum.
         */
        formula: z.string().min(1).max(1_024).optional(),
      }),
    )
    .min(1)
    .max(8),
  /**
   * Dimension-field grouping: date groups by year/quarter/month (quarter/month
   * merge across years); range groups into fixed-step numeric intervals (bucket
   * boundaries rangeStart + k*rangeStep, default start 0). field must be a
   * dimension field from rowFields/columnField.
   */
  groupings: z
    .array(
      z.discriminatedUnion('kind', [
        z.object({
          kind: z.literal('date'),
          field: z.string().min(1).max(255),
          dateUnit: z.enum(['year', 'quarter', 'month']),
        }),
        z.object({
          kind: z.literal('range'),
          field: z.string().min(1).max(255),
          rangeStep: z.number().positive().finite(),
          rangeStart: z.number().finite().optional(),
        }),
      ]),
    )
    .max(8)
    .optional(),
  /**
   * Value/label filters: label filters row/column field members by label
   * (equals/contains/begins-with, case-insensitive); value filters the field's
   * members by one values entry's aggregate (top: largest count / greaterThan:
   * > from / between: from ≤ x ≤ to). field must be a dimension field from
   * rowFields/columnField, at most one filter per field.
   */
  filters: z
    .array(
      z.discriminatedUnion('kind', [
        z.object({
          kind: z.literal('label'),
          field: z.string().min(1).max(255),
          op: z.enum(['equal', 'contains', 'beginsWith']),
          value: z.string().max(255),
        }),
        z.object({
          kind: z.literal('value'),
          field: z.string().min(1).max(255),
          /** The data field it measures (values array index, defaults to 0) */
          valueIndex: z.number().int().nonnegative().max(7).default(0),
          op: z.enum(['top', 'greaterThan', 'between']),
          count: z.number().int().min(1).max(10_000).optional(),
          from: z.number().finite().optional(),
          to: z.number().finite().optional(),
        }),
      ]),
    )
    .max(8)
    .optional(),
  /** pivot table name; default Pivot1, Pivot2, … */
  name: z
    .string()
    .regex(/^[\p{L}_\\][\p{L}\p{N}_.\\ ]{0,254}$/u)
    .optional(),
})

const setRowsHiddenSchema = z.object({
  op: z.literal('set_rows_hidden'),
  sheetId: z.string().min(1),
  /** 1-based first row */
  row: z.number().int().min(1).max(9999999),
  count: z.number().int().min(1).max(10000).default(1),
  hidden: z.boolean(),
})

const setColsHiddenSchema = z.object({
  op: z.literal('set_cols_hidden'),
  sheetId: z.string().min(1),
  column: columnLabelSchema,
  count: z.number().int().min(1).max(1000).default(1),
  hidden: z.boolean(),
})

// target: full URL (https://…) or internal reference like "Sheet1!A1";
// null removes the link.
const setHyperlinkSchema = z.object({
  op: z.literal('set_hyperlink'),
  sheetId: z.string().min(1),
  address: cellAddressSchema,
  target: z.string().min(1).max(2048).nullable(),
})

const protectSheetSchema = z.object({
  op: z.literal('protect_sheet'),
  sheetId: z.string().min(1),
  protected: z.boolean(),
})

// Workbook structure protection (Review → Protect Workbook): blocks sheet
// add/delete/move/rename/hide while protected. File-level setting — the
// executor writes it into workbook.xml on save.
const protectWorkbookSchema = z.object({
  op: z.literal('protect_workbook'),
  protected: z.boolean(),
})

// Page Layout → Themes / Theme Colors / Theme Fonts: preset ids are resolved
// against the renderer's theme catalogs by the executor (the catalogs live in
// the renderer, so id validation cannot happen here). At least one id is
// required — checked at expansion, a discriminated-union member cannot refine.
const themePresetIdSchema = z.string().min(1).max(64)
const setThemeSchema = z.object({
  op: z.literal('set_theme'),
  /** full theme preset id, e.g. "office" */
  theme: themePresetIdSchema.optional(),
  /** color-scheme preset id, e.g. "office" */
  colors: themePresetIdSchema.optional(),
  /** font-scheme preset id, e.g. "calibri" */
  fonts: themePresetIdSchema.optional(),
})

const allowEditRangeSchema = z.object({
  /** optional label shown in the Allow Edit Ranges dialog */
  name: z.string().min(1).max(255).optional(),
  /** the range that stays editable on the protected sheet, e.g. "B2:D10" */
  sqref: cellRangeSchema,
})

// Review → Allow Edit Ranges: which ranges remain editable while the sheet is
// protected. An empty `ranges` list clears the sheet's allow-edit ranges.
const allowEditRangesSchema = z.object({
  op: z.literal('allow_edit_ranges'),
  sheetId: z.string().min(1),
  ranges: z.array(allowEditRangeSchema).max(100),
})

// Creates (or replaces) the sheet's auto-filter over the range. Filter
// criteria are then chosen by the user through the column dropdowns.
const setFilterSchema = z.object({
  op: z.literal('set_filter'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
})

const clearFilterSchema = z.object({
  op: z.literal('clear_filter'),
  sheetId: z.string().min(1),
})

// Column criteria on an existing auto-filter, in one of two shapes:
// - checkbox style: rows whose displayed cell text is in `values` stay visible;
// - custom conditions: 1–2 value comparisons joined with AND (`and` default
//   true) or OR (`and:false`), e.g. [{operator:"greaterThan", value:"100"}].
// `values: null` clears the column's criteria. Exactly one of values /
// conditions must be present — checked at expansion because discriminated-
// union members cannot carry refinements.
const filterConditionSchema = z.object({
  operator: z.enum([
    'equal',
    'notEqual',
    'greaterThan',
    'greaterThanOrEqual',
    'lessThan',
    'lessThanOrEqual',
  ]),
  /** compared against the cell's displayed value (numbers/booleans as text) */
  value: z.string().max(255),
})

const setFilterCriteriaSchema = z.object({
  op: z.literal('set_filter_criteria'),
  sheetId: z.string().min(1),
  column: columnLabelSchema,
  values: z.array(z.string().max(255)).min(1).max(1000).nullable().optional(),
  conditions: z.array(filterConditionSchema).min(1).max(2).optional(),
  /** join for `conditions`: true = AND (default), false = OR */
  and: z.boolean().optional(),
})

const cfFormatSchema = z
  .object({
    fillColor: hexColorSchema.optional(),
    fontColor: hexColorSchema.optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
  })
  .refine((format) => Object.values(format).some((value) => value !== undefined), {
    message: 'The rule format must set at least one property.',
  })

export const cfRuleSchema = z.discriminatedUnion('kind', [
  z.object({
    kind: z.literal('number'),
    operator: z.enum([
      'greaterThan',
      'greaterThanOrEqual',
      'lessThan',
      'lessThanOrEqual',
      'equal',
      'notEqual',
      'between',
      'notBetween',
    ]),
    value: z.number().finite(),
    /** required for between / notBetween */
    value2: z.number().finite().optional(),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('text'),
    operator: z.enum(['contains', 'notContains', 'beginsWith', 'endsWith']),
    text: z.string().min(1).max(255),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('blank'),
    /** true highlights empty cells, false highlights non-empty ones */
    blank: z.boolean(),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('duplicate'),
    /** true highlights unique values instead of duplicates */
    unique: z.boolean().optional(),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('top10'),
    rank: z.number().int().min(1).max(1000),
    percent: z.boolean().optional(),
    bottom: z.boolean().optional(),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('formula'),
    formula: z.string().startsWith('=').max(8192),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('colorScale'),
    minColor: hexColorSchema,
    midColor: hexColorSchema.optional(),
    maxColor: hexColorSchema,
  }),
  z.object({
    kind: z.literal('dataBar'),
    color: hexColorSchema.optional(),
  }),
  z.object({
    kind: z.literal('iconSet'),
    /** OOXML icon set: 3Arrows / 3TrafficLights1 / 3Signs / 4Arrows / 4RedToBlack / 5Quarters / 5Rating … */
    iconSet: z.enum([
      '3Arrows',
      '3ArrowsGray',
      '3Flags',
      '3TrafficLights1',
      '3TrafficLights2',
      '3Signs',
      '3Symbols',
      '3Symbols2',
      '4Arrows',
      '4ArrowsGray',
      '4RedToBlack',
      '4Rating',
      '4TrafficLights',
      '5Arrows',
      '5ArrowsGray',
      '5Quarters',
      '5Rating',
    ]),
    /** show the cell values next to the icons (default true) */
    showValue: z.boolean().optional(),
    /** descending greaterThanOrEqual thresholds; length = icon count - 1; default percentile split */
    thresholds: z
      .array(
        z.object({ type: z.enum(['num', 'percent', 'percentile']), value: z.number().finite() }),
      )
      .max(4)
      .optional(),
  }),
])

const addConditionalFormatSchema = z.object({
  op: z.literal('add_conditional_format'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  rule: cfRuleSchema,
})

const clearConditionalFormatsSchema = z.object({
  op: z.literal('clear_conditional_formats'),
  sheetId: z.string().min(1),
})

export const dvRuleSchema = z.discriminatedUnion('kind', [
  z.object({
    kind: z.literal('list'),
    values: z.array(z.string().min(1).max(255)).min(1).max(100),
  }),
  z.object({
    kind: z.literal('listRef'),
    /** range whose cell values form the dropdown options */
    range: cellRangeSchema,
  }),
  z.object({
    kind: z.literal('numberBetween'),
    min: z.number().finite(),
    max: z.number().finite(),
  }),
  z.object({
    kind: z.literal('dateBetween'),
    /** ISO dates, e.g. "2026-01-31" */
    start: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
    end: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  }),
  z.object({ kind: z.literal('checkbox') }),
  z.object({
    kind: z.literal('formula'),
    formula: z.string().startsWith('=').max(8192),
  }),
])

// validation: null clears any rule on the range.
const setDataValidationSchema = z.object({
  op: z.literal('set_data_validation'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  validation: dvRuleSchema.nullable(),
})

// Print/page-layout settings; saved into the file, mirroring the Page Layout
// ribbon. At least one setting required (checked at expansion). scale and
// fitToWidth/fitToHeight are mutually exclusive.
const setPageSetupSchema = z.object({
  op: z.literal('set_page_setup'),
  sheetId: z.string().min(1),
  orientation: z.enum(['portrait', 'landscape']).optional(),
  /** OOXML paper-size code: 1 Letter, 8 A3, 9 A4, 11 A5 … */
  paperSize: z.number().int().min(1).max(118).optional(),
  /** print scale percent */
  scale: z.number().int().min(10).max(400).optional(),
  /** pages wide (0 = automatic) */
  fitToWidth: z.number().int().min(0).max(1000).optional(),
  /** pages tall (0 = automatic) */
  fitToHeight: z.number().int().min(0).max(1000).optional(),
  margins: z.enum(['normal', 'wide', 'narrow']).optional(),
  printGridlines: z.boolean().optional(),
  printHeadings: z.boolean().optional(),
  /** A1 range to print; null clears the print area */
  printArea: cellRangeSchema.nullable().optional(),
  /** Rows repeated at the top of every printed page, e.g. "1:2"; null clears */
  printTitles: z
    .string()
    .regex(/^\d+(:\d+)?$/)
    .nullable()
    .optional(),
  /** Printed header sections; text carries Excel codes (&P page, &N pages, &D date, &A sheet name); null clears */
  header: z
    .object({
      left: z.string().max(255).optional(),
      center: z.string().max(255).optional(),
      right: z.string().max(255).optional(),
    })
    .nullable()
    .optional(),
  /** Printed footer sections; same codes as header; null clears */
  footer: z
    .object({
      left: z.string().max(255).optional(),
      center: z.string().max(255).optional(),
      right: z.string().max(255).optional(),
    })
    .nullable()
    .optional(),
  /** Manual page breaks: 0-based index of the LAST ROW printed on the page */
  rowBreaks: z.array(z.number().int().min(0).max(1048574)).max(50).optional(),
  /** Manual page breaks: 0-based index of the last column printed on the page */
  colBreaks: z.array(z.number().int().min(0).max(16383)).max(50).optional(),
})

// Cell note (legacy comment): text null removes the note.
const setNoteSchema = z.object({
  op: z.literal('set_note'),
  sheetId: z.string().min(1),
  address: cellAddressSchema,
  text: z.string().min(1).max(32_000).nullable(),
})

// Recomputes every pivot table on the sheet from its current source data
// (layout unchanged; requires the fully-loaded mode).
const refreshPivotSchema = z.object({
  op: z.literal('refresh_pivot'),
  sheetId: z.string().min(1),
})

// Freeze panes: rows/columns counted from the top-left; 0/0 unfreezes.
const setFreezeSchema = z.object({
  op: z.literal('set_freeze'),
  sheetId: z.string().min(1),
  rows: z.number().int().min(0).max(100),
  columns: z.number().int().min(0).max(100),
})

const duplicateSheetSchema = z.object({
  op: z.literal('duplicate_sheet'),
  sheetId: z.string().min(1),
  name: sheetNameSchema.optional(),
})

const setSheetHiddenSchema = z.object({
  op: z.literal('set_sheet_hidden'),
  sheetId: z.string().min(1),
  hidden: z.boolean(),
})

const moveSheetSchema = z.object({
  op: z.literal('move_sheet'),
  sheetId: z.string().min(1),
  /** 1-based target tab position */
  position: z.number().int().min(1).max(1000),
})

// ref: an A1 reference like "Sheet1!$A$1:$B$10" or a constant/formula body
// (no leading =), as defined names are stored in the file.
const addDefinedNameSchema = z.object({
  op: z.literal('add_defined_name'),
  name: z.string().regex(/^[\p{L}_\\][\p{L}\p{N}_.\\]{0,254}$/u),
  ref: z.string().min(1).max(8192),
})

const deleteDefinedNameSchema = z.object({
  op: z.literal('delete_defined_name'),
  name: z.string().min(1).max(255),
})

/// OOXML ST_BorderStyle line styles — the same list as the shared
/// `editableBorderStyleSchema` (shared/desktop-api.ts) the save path validates
/// against, so an AI border style always round-trips into the file.
const borderLineStyleSchema = z.enum([
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
])

// Per-cell edge semantics: every cell in the range gets the named edge(s);
// 'none' removes all borders. Clearing is 'none', so null is not accepted.
const borderPatchSchema = z.object({
  type: z.enum(['all', 'top', 'bottom', 'left', 'right', 'none']),
  color: hexColorSchema.optional(),
  /** OOXML line style applied to the edge(s); ignored when type is 'none' */
  style: borderLineStyleSchema.optional(),
})

// Every field optional; null clears that property back to the default.
const formatPatchSchema = z
  .object({
    bold: z.boolean().nullable().optional(),
    italic: z.boolean().nullable().optional(),
    underline: z.boolean().nullable().optional(),
    strikethrough: z.boolean().nullable().optional(),
    fontFamily: z.string().min(1).max(128).nullable().optional(),
    fontSize: z.number().min(1).max(409).nullable().optional(),
    fontColor: hexColorSchema.nullable().optional(),
    fillColor: hexColorSchema.nullable().optional(),
    numberFormat: z.string().min(1).max(255).nullable().optional(),
    horizontalAlign: z.enum(['left', 'center', 'right']).nullable().optional(),
    verticalAlign: z.enum(['top', 'center', 'bottom']).nullable().optional(),
    wrapText: z.boolean().nullable().optional(),
    /** degrees, -90 (clockwise) to 90 (counterclockwise), or 'vertical' (stacked) */
    textRotation: z
      .union([z.number().int().min(-90).max(90), z.literal('vertical')])
      .nullable()
      .optional(),
    /** OOXML indent steps (0 or null clears); renders as left cell padding */
    indent: z.number().int().min(0).max(250).nullable().optional(),
    border: borderPatchSchema.optional(),
    /** cell protection: locked cells reject edits on a protected sheet */
    locked: z.boolean().nullable().optional(),
    /** cell protection: hidden cells hide their formula on a protected sheet */
    hidden: z.boolean().nullable().optional(),
  })
  .refine((patch) => Object.values(patch).some((value) => value !== undefined), {
    message: 'format must set at least one property',
  })

const formatRangeSchema = z.object({
  op: z.literal('format_range'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  format: formatPatchSchema,
})

// Clear Formats: resets the range's styling AND cell-protection flags by
// expanding to ONE format_range carrying every clearing value for the full
// patch shape (false / null / indent 0 / border 'none' / locked+hidden null).
// Content-classed, so it may share a batch with other cell edits.
const clearFormatsSchema = z.object({
  op: z.literal('clear_formats'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
})

/// Every field of formatPatchSchema in its cleared state — what clear_formats
/// expands into. null clears the property in the format engine, `indent: 0`
/// resets padding, border 'none' removes all edges, and locked/hidden null
/// restores the protection defaults (locked=true, hidden=false).
const CLEARED_FORMAT_PATCH: CellFormatPatch = {
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

// Sort rows by 1–3 key columns: later keys break ties of earlier ones, the
// sort is stable, and blanks always stay last. `keys` takes precedence over
// the legacy single-key pair `byColumn`+`order`; at least one of the two forms
// is required (checked at expansion — a zod refinement cannot live on a
// discriminated-union member).
const sortKeySchema = z.object({
  byColumn: columnLabelSchema,
  order: z.enum(['asc', 'desc']),
})

const sortRangeSchema = z.object({
  op: z.literal('sort_range'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  /** legacy single-key form; ignored when `keys` is present */
  byColumn: columnLabelSchema.optional(),
  /** legacy single-key form; ignored when `keys` is present */
  order: z.enum(['asc', 'desc']).optional(),
  /** 1–3 keys applied in order, e.g. [{byColumn:"B",order:"desc"}] */
  keys: z.array(sortKeySchema).min(1).max(3).optional(),
  hasHeader: z.boolean().optional(),
})

const mergeCellsSchema = z.object({
  op: z.literal('merge_cells'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
})

const unmergeCellsSchema = z.object({
  op: z.literal('unmerge_cells'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
})

const setRowHeightSchema = z.object({
  op: z.literal('set_row_height'),
  sheetId: z.string().min(1),
  /** 1-based first row */
  row: z.number().int().min(1).max(9999999),
  count: z.number().int().min(1).max(500).default(1),
  /** Excel points (2–409) */
  heightPoints: z.number().min(2).max(409),
})

const setColWidthSchema = z.object({
  op: z.literal('set_col_width'),
  sheetId: z.string().min(1),
  column: columnLabelSchema,
  count: z.number().int().min(1).max(100).default(1),
  /** pixels (10–2000) */
  widthPx: z.number().min(10).max(2000),
})

const renameSheetSchema = z.object({
  op: z.literal('rename_sheet'),
  sheetId: z.string().min(1),
  name: sheetNameSchema,
})

// Deletes a floating visual (chart/shape/image). Charts accept the chart
// part path or the visual id; shapes/images use the visual id (both listed
// in get_workbook_context / read_sheet_features).
const deleteVisualSchema = z.object({
  op: z.literal('delete_visual'),
  visualId: z.string().min(1).max(300),
})

// Removes a table created THIS session (converts it back to a plain range —
// values and formatting stay). File-native tables cannot be removed yet.
const deleteTableSchema = z.object({
  op: z.literal('delete_table'),
  sheetId: z.string().min(1),
  tableName: z.string().min(1).max(255),
})

// In-cell sparkline minicharts: one per dataRange row, hosted in the column
// right of the range (or starting at targetCell going down).
const addSparklineSchema = z.object({
  op: z.literal('add_sparkline'),
  sheetId: z.string().min(1),
  type: z.enum(['line', 'column', 'stacked']),
  dataRange: cellRangeSchema,
  targetCell: cellAddressSchema.optional(),
  color: hexColorSchema.optional(),
})

// Literal find & replace over TEXT values in a range (formulas and numbers
// are untouched). Expands to plain set_cell edits, so preview and undo
// behave like typing. Case-insensitive by default.
const findReplaceSchema = z.object({
  op: z.literal('find_replace'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  find: z.string().min(1).max(255),
  replace: z.string().max(255),
  matchCase: z.boolean().optional(),
  /** match the whole cell text instead of substrings */
  wholeCell: z.boolean().optional(),
})

// Data → Remove Duplicates: keeps the first occurrence of each identical row
// (text compared case-insensitively across the range's full width) and
// deletes the later duplicate rows as WHOLE rows, exactly like Excel. It is
// structural: deleting rows shifts every address below, so this op needs its
// own batch. Expands into delete_rows primitives (bottom-up) instead of
// carrying any new executor behavior.
const removeDuplicatesSchema = z.object({
  op: z.literal('remove_duplicates'),
  sheetId: z.string().min(1),
  /** the table, header row included when hasHeader is true */
  range: cellRangeSchema,
  hasHeader: z.boolean().optional(),
})

// Data → Text to Columns: splits every cell of a one-column (or one-row)
// source on a delimiter and writes the pieces into the adjacent cells.
// `delimiter: null` splits on runs of any whitespace. `destination` is the
// output's top-left cell (default: the source's top-left). `clearSource`
// removes the original text when the output lands elsewhere (default true).
// The exact in-place/away clear rules are documented at the expansion site.
const textToColumnsSchema = z.object({
  op: z.literal('text_to_columns'),
  sheetId: z.string().min(1),
  /** single column (spills right) or single row (spills down) */
  source: cellRangeSchema,
  /** null = split on any run of whitespace */
  delimiter: z.string().min(1).max(32).nullable(),
  /** top-left cell of the output block; defaults to the source's top-left */
  destination: cellAddressSchema.optional(),
  /** treat consecutive delimiters as one (default true) */
  collapseConsecutive: z.boolean().optional(),
  /** clear the source cells after splitting (default true; only matters when destination differs) */
  clearSource: z.boolean().optional(),
})

// Data → Consolidate: aggregates several same-shaped source ranges into a
// block of live formulas, one per target cell. Cross-sheet references are
// qualified (`Sheet2!B2`), same-sheet ones are bare, and sheet names are
// quoted only when they need to be. The target block is exactly the shape of
// the reference ranges.
const consolidateReferenceSchema = z.object({
  /** sheet the range lives on; defaults to the target sheet */
  sheetId: z.string().min(1).optional(),
  range: cellRangeSchema,
})

const consolidateSchema = z.object({
  op: z.literal('consolidate'),
  sheetId: z.string().min(1),
  /** top-left cell of the output block */
  target: cellAddressSchema,
  /** aggregation function applied to the aligned value of every reference */
  function: z.enum(['sum', 'count', 'average', 'max', 'min']),
  /** 1–20 source ranges, all of identical shape */
  references: z.array(consolidateReferenceSchema).min(1).max(20),
})

// Formulas → Create Names from Selection: turns a label row/column into
// defined names for the adjacent data row/column, like the ribbon command.
// Expands into add_defined_name primitives with absolute refs; labels are
// sanitized into legal names (invalid characters → _, leading digit or
// cell-like text gets an _ prefix) and empty/duplicate labels are skipped or
// numbered.
const createNamesSchema = z.object({
  op: z.literal('create_names_from_selection'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  /** where the labels live: the range's first/last row or column */
  label: z.enum(['top', 'left', 'bottom', 'right']),
})

// Data → Subtotal: inserts a bold SUBTOTAL summary row under each contiguous
// group of `groupColumn` values, plus a Grand Total row, aggregating
// `valueColumn` with the matching SUBTOTAL function. STRUCTURAL: summary rows
// shift every address below, so it needs its own batch. Passes through
// unchanged — the executor drives Univer's insertRowsBefore/formula mutations,
// which cannot be expressed as per-cell primitives here.
const subtotalSchema = z.object({
  op: z.literal('subtotal'),
  sheetId: z.string().min(1),
  /** the table; its FIRST row is the header and does not participate */
  range: cellRangeSchema,
  /** column whose runs of equal values define the groups, e.g. "A" */
  groupColumn: columnLabelSchema,
  /** column each group's SUBTOTAL is computed over, e.g. "D" */
  valueColumn: columnLabelSchema,
  /** SUBTOTAL function number: sum→9 average→1 count→2 counta→3 max→4 min→5 product→6 stdev→7 stdevp→8 var→10 varp→11 */
  function: z.enum([
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
  ]),
})

// Outline groups (Data → Group / Subtotal outline buttons). These are
// layout-class pass-through ops: the executor edits the outline journal and
// row/column visibility. `range` is axis-local like Excel's row/column
// ranges — "5:12" for rows, "B:D" for cols — A1-style ranges ("A5:C12",
// "B1:D1") are accepted too; the executor reads only the range's rows (axis
// rows) or columns (axis cols).
const outlineSpanSchema = z
  .string()
  .regex(
    /^(?:[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?|[1-9][0-9]{0,6}(?::[1-9][0-9]{0,6})?|[A-Z]{1,3}(?::[A-Z]{1,3})?)$/,
  )
const outlineAxisSchema = z.enum(['rows', 'cols'])
const outlineGroupSchema = z.object({
  op: z.literal('outline.group'),
  sheetId: z.string().min(1),
  axis: outlineAxisSchema,
  range: outlineSpanSchema,
})
const outlineUngroupSchema = z.object({
  op: z.literal('outline.ungroup'),
  sheetId: z.string().min(1),
  axis: outlineAxisSchema,
  range: outlineSpanSchema,
})
const outlineHideDetailSchema = z.object({
  op: z.literal('outline.hideDetail'),
  sheetId: z.string().min(1),
  axis: outlineAxisSchema,
  range: outlineSpanSchema,
})
const outlineShowDetailSchema = z.object({
  op: z.literal('outline.showDetail'),
  sheetId: z.string().min(1),
  axis: outlineAxisSchema,
  range: outlineSpanSchema,
})

export const workbookOperationSchema = z.discriminatedUnion('op', [
  setCellSchema,
  setFormulaSchema,
  clearCellSchema,
  setRangeSchema,
  clearRangeSchema,
  fillRangeSchema,
  copyRangeSchema,
  convertToValuesSchema,
  formatRangeSchema,
  clearFormatsSchema,
  sortRangeSchema,
  mergeCellsSchema,
  unmergeCellsSchema,
  setRowHeightSchema,
  setColWidthSchema,
  editChartSchema,
  addChartSchema,
  addShapeSchema,
  editShapeSchema,
  addImageSchema,
  addTableSchema,
  addTableRowSchema,
  addTableColumnSchema,
  deleteTableRowSchema,
  deleteTableColumnSchema,
  addPivotSchema,
  setRowsHiddenSchema,
  setColsHiddenSchema,
  setHyperlinkSchema,
  protectSheetSchema,
  protectWorkbookSchema,
  setThemeSchema,
  allowEditRangesSchema,
  setFilterSchema,
  clearFilterSchema,
  setFilterCriteriaSchema,
  addConditionalFormatSchema,
  clearConditionalFormatsSchema,
  setDataValidationSchema,
  addDefinedNameSchema,
  deleteDefinedNameSchema,
  setPageSetupSchema,
  setFreezeSchema,
  setNoteSchema,
  refreshPivotSchema,
  insertRowsSchema,
  deleteRowsSchema,
  insertColsSchema,
  deleteColsSchema,
  addSheetSchema,
  deleteSheetSchema,
  duplicateSheetSchema,
  setSheetHiddenSchema,
  moveSheetSchema,
  renameSheetSchema,
  deleteVisualSchema,
  deleteTableSchema,
  findReplaceSchema,
  addSparklineSchema,
  removeDuplicatesSchema,
  textToColumnsSchema,
  consolidateSchema,
  createNamesSchema,
  subtotalSchema,
  outlineGroupSchema,
  outlineUngroupSchema,
  outlineHideDetailSchema,
  outlineShowDetailSchema,
])

export type WorkbookOperation = z.infer<typeof workbookOperationSchema>
export type SetCellOperation = z.infer<typeof setCellSchema>
export type SetRangeOperation = z.infer<typeof setRangeSchema>
export type SetFormulaOperation = z.infer<typeof setFormulaSchema>
export type ClearCellOperation = z.infer<typeof clearCellSchema>
export type ClearRangeOperation = z.infer<typeof clearRangeSchema>
export type FillRangeOperation = z.infer<typeof fillRangeSchema>
export type FormatRangeOperation = z.infer<typeof formatRangeSchema>
export type CellFormatPatch = z.infer<typeof formatPatchSchema>
export type BorderPatch = z.infer<typeof borderPatchSchema>
export type StructuralOperation =
  | z.infer<typeof insertRowsSchema>
  | z.infer<typeof deleteRowsSchema>
  | z.infer<typeof insertColsSchema>
  | z.infer<typeof deleteColsSchema>
  | z.infer<typeof addSheetSchema>
  | z.infer<typeof deleteSheetSchema>
  | z.infer<typeof duplicateSheetSchema>
  | z.infer<typeof setSheetHiddenSchema>
  | z.infer<typeof moveSheetSchema>
export type EditChartOperation = z.infer<typeof editChartSchema>
export type AddChartOperation = z.infer<typeof addChartSchema>
export type AddShapeOperation = z.infer<typeof addShapeSchema>
export type EditShapeOperation = z.infer<typeof editShapeSchema>
export type AddImageOperation = z.infer<typeof addImageSchema>
export type AddTableOperation = z.infer<typeof addTableSchema>
export type AddTableRowOperation = z.infer<typeof addTableRowSchema>
export type AddTableColumnOperation = z.infer<typeof addTableColumnSchema>
export type DeleteTableRowOperation = z.infer<typeof deleteTableRowSchema>
export type DeleteTableColumnOperation = z.infer<typeof deleteTableColumnSchema>
export type AddPivotOperation = z.infer<typeof addPivotSchema>
export type SetHyperlinkOperation = z.infer<typeof setHyperlinkSchema>
export type AddConditionalFormatOperation = z.infer<typeof addConditionalFormatSchema>
export type SetDataValidationOperation = z.infer<typeof setDataValidationSchema>
export type SetPageSetupOperation = z.infer<typeof setPageSetupSchema>
export type CfRule = z.infer<typeof cfRuleSchema>
export type DvRule = z.infer<typeof dvRuleSchema>
/** sheet-layout edits: no address shifts, so they may mix with content ops */
export type LayoutOperation =
  | z.infer<typeof mergeCellsSchema>
  | z.infer<typeof unmergeCellsSchema>
  | z.infer<typeof setRowHeightSchema>
  | z.infer<typeof setColWidthSchema>
  | EditChartOperation
  | AddChartOperation
  | AddShapeOperation
  | EditShapeOperation
  | AddImageOperation
  | AddTableOperation
  | AddTableRowOperation
  | AddTableColumnOperation
  | DeleteTableRowOperation
  | DeleteTableColumnOperation
  | AddPivotOperation
  | z.infer<typeof setRowsHiddenSchema>
  | z.infer<typeof setColsHiddenSchema>
  | SetHyperlinkOperation
  | z.infer<typeof protectSheetSchema>
  | ProtectWorkbookOperation
  | SetThemeOperation
  | AllowEditRangesOperation
  | z.infer<typeof setFilterSchema>
  | z.infer<typeof clearFilterSchema>
  | z.infer<typeof setFilterCriteriaSchema>
  | AddConditionalFormatOperation
  | z.infer<typeof clearConditionalFormatsSchema>
  | SetDataValidationOperation
  | z.infer<typeof addDefinedNameSchema>
  | z.infer<typeof deleteDefinedNameSchema>
  | SetPageSetupOperation
  | z.infer<typeof setFreezeSchema>
  | z.infer<typeof setNoteSchema>
  | z.infer<typeof refreshPivotSchema>
  | DeleteVisualOperation
  | DeleteTableOperation
  | AddSparklineOperation
  | OutlineOperation
export type DeleteVisualOperation = z.infer<typeof deleteVisualSchema>
export type DeleteTableOperation = z.infer<typeof deleteTableSchema>
export type AddSparklineOperation = z.infer<typeof addSparklineSchema>
export type OutlineOperation =
  | z.infer<typeof outlineGroupSchema>
  | z.infer<typeof outlineUngroupSchema>
  | z.infer<typeof outlineHideDetailSchema>
  | z.infer<typeof outlineShowDetailSchema>
export type ProtectWorkbookOperation = z.infer<typeof protectWorkbookSchema>
export type SetThemeOperation = z.infer<typeof setThemeSchema>
export type AllowEditRangesOperation = z.infer<typeof allowEditRangesSchema>
export type ClearFormatsOperation = z.infer<typeof clearFormatsSchema>
export type RemoveDuplicatesOperation = z.infer<typeof removeDuplicatesSchema>
export type TextToColumnsOperation = z.infer<typeof textToColumnsSchema>
export type ConsolidateOperation = z.infer<typeof consolidateSchema>
export type CreateNamesOperation = z.infer<typeof createNamesSchema>
export type SubtotalOperation = z.infer<typeof subtotalSchema>
export type SortRangeOperation = z.infer<typeof sortRangeSchema>
export type FindReplaceOperation = z.infer<typeof findReplaceSchema>
export type CopyRangeOperation = z.infer<typeof copyRangeSchema>
export type ConvertToValuesOperation = z.infer<typeof convertToValuesSchema>
export type CellContentOperation = SetCellOperation | SetFormulaOperation | ClearCellOperation
/** what range ops expand into; the only shapes executors have to handle.
 * fill_range, copy_range, convert_to_values, and large clear_range /
 * find_replace (>MAX_EXPANDED_CELL_OPS cells) pass through as range-level
 * primitives — executors apply them with bulk grid reads/writes instead of
 * per-cell edits. */
export type PrimitiveOperation =
  | CellContentOperation
  | FormatRangeOperation
  | FillRangeOperation
  | CopyRangeOperation
  | ConvertToValuesOperation
  | ClearRangeOperation
  | FindReplaceOperation
  | LayoutOperation
  | StructuralOperation
  // Subtotal passes through as one structural op (it inserts summary rows the
  // executor performs via Univer commands); it is not part of the
  // StructuralOperation union, whose members the row/column shift helpers
  // switch over exhaustively.
  | SubtotalOperation
  | z.infer<typeof renameSheetSchema>

const STRUCTURAL_OPS = new Set([
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'add_sheet',
  'delete_sheet',
  'duplicate_sheet',
  'set_sheet_hidden',
  'move_sheet',
  // remove_duplicates expands into delete_rows: it shifts addresses exactly
  // like a row deletion, so it belongs to the structural lane (own batch).
  'remove_duplicates',
  // subtotal inserts summary rows (addresses below shift), so it must not
  // share a batch with cell edits. It passes through to the executor as-is.
  'subtotal',
])
const LAYOUT_OPS = new Set([
  'merge_cells',
  'unmerge_cells',
  'set_row_height',
  'set_col_width',
  'edit_chart',
  'add_chart',
  'add_shape',
  'edit_shape',
  'add_image',
  'add_table',
  'add_pivot',
  'add_table_row',
  'add_table_column',
  'delete_table_row',
  'delete_table_column',
  'set_rows_hidden',
  'set_cols_hidden',
  'set_hyperlink',
  'protect_sheet',
  'protect_workbook',
  'set_theme',
  'allow_edit_ranges',
  'set_filter',
  'clear_filter',
  'set_filter_criteria',
  'add_conditional_format',
  'clear_conditional_formats',
  'set_data_validation',
  'add_defined_name',
  'delete_defined_name',
  'set_page_setup',
  'set_freeze',
  'set_note',
  'refresh_pivot',
  'delete_visual',
  'delete_table',
  'add_sparkline',
  'outline.group',
  'outline.ungroup',
  'outline.hideDetail',
  'outline.showDetail',
])
// Content addresses cells, so it must not ride in the same batch as
// row/column shifts that would move those addresses. Layout ops keep their
// row/column coordinates too, so they count the same way.
const CELL_CONTENT_OPS = new Set([
  'set_cell',
  'set_formula',
  'clear_cell',
  'set_range',
  'clear_range',
  'fill_range',
  'copy_range',
  'convert_to_values',
  'format_range',
  'clear_formats',
  'sort_range',
  'find_replace',
  // Data tools that expand into per-cell writes / formulas / defined names.
  // They do not shift addresses, so they may share a batch with content edits.
  'text_to_columns',
  'consolidate',
  'create_names_from_selection',
  ...LAYOUT_OPS,
])

export function isStructuralOp(
  op: WorkbookOperation | PrimitiveOperation,
): op is StructuralOperation | SubtotalOperation {
  return STRUCTURAL_OPS.has(op.op)
}

export function isLayoutOp(op: WorkbookOperation | PrimitiveOperation): op is LayoutOperation {
  return LAYOUT_OPS.has(op.op)
}

export const workbookCommandBatchSchema = z
  .object({
    dslVersion: z.literal(1),
    transactionId: z.string().min(1).max(128),
    baseRevision: z.number().int().nonnegative(),
    summary: z.string().trim().min(1).max(500),
    operations: z.array(workbookOperationSchema).min(1).max(1000),
  })
  .refine(
    (batch) => {
      // Structural ops shift addresses, so mixing them with cell edits in one
      // batch would make previewed before/after states unverifiable. Propose
      // structure first, apply, then propose content.
      const hasStructural = batch.operations.some((op) => STRUCTURAL_OPS.has(op.op))
      const hasCellContent = batch.operations.some((op) => CELL_CONTENT_OPS.has(op.op))
      return !(hasStructural && hasCellContent)
    },
    {
      message:
        'Structural operations (rows/columns/sheets) and cell edits must be proposed in separate batches.',
    },
  )

export type WorkbookCommandBatch = z.infer<typeof workbookCommandBatchSchema>

export function parseWorkbookCommandBatch(input: unknown): WorkbookCommandBatch {
  return workbookCommandBatchSchema.parse(input)
}

export const MAX_EXPANDED_CELL_OPS = 2000
/** cap for range-level ops (fill_range / large clear_range / format_range),
 * which apply as one bulk grid write instead of per-cell edits */
export const MAX_RANGE_OP_CELLS = 200_000
/** a fill's source block is read cell by cell, so it stays small */
export const MAX_FILL_SOURCE_CELLS = 2000
/** data-tool ops (remove_duplicates / text_to_columns / consolidate) carry
 * their own cell budget: they expand into many plain cell edits, but a bulk
 * split/aggregation legitimately touches more cells than the per-cell preview
 * ceiling allows */
export const MAX_DATA_TOOL_CELLS = 20_000

/// fill_range geometry guards: the source must tile the target exactly, and
/// an overlapping source must sit at the target's top-left corner (the
/// classic fill-down/right shape) so no source cell is overwritten before
/// it is copied.
function validateFillRange(operation: FillRangeOperation): void {
  const source = parseRange(operation.source)
  const target = parseRange(operation.target)
  if (rangeCellCount(source) > MAX_FILL_SOURCE_CELLS) {
    throw new Error(`fill_range source covers more than ${MAX_FILL_SOURCE_CELLS} cells.`)
  }
  if (rangeCellCount(target) > MAX_RANGE_OP_CELLS) {
    throw new Error(
      `fill_range target covers more than ${MAX_RANGE_OP_CELLS.toLocaleString('en-US')} cells — fill it in several ranges.`,
    )
  }
  const sourceRows = source.endRow - source.startRow + 1
  const sourceColumns = source.endColumn - source.startColumn + 1
  const targetRows = target.endRow - target.startRow + 1
  const targetColumns = target.endColumn - target.startColumn + 1
  if (targetRows % sourceRows !== 0 || targetColumns % sourceColumns !== 0) {
    throw new Error(
      `fill_range target ${operation.target} (${targetRows}×${targetColumns}) is not a whole multiple of source ${operation.source} (${sourceRows}×${sourceColumns}) — adjust the target so the source tiles it exactly.`,
    )
  }
  const sameSheet =
    operation.sourceSheetId === undefined || operation.sourceSheetId === operation.sheetId
  const overlaps =
    sameSheet &&
    source.startRow <= target.endRow &&
    source.endRow >= target.startRow &&
    source.startColumn <= target.endColumn &&
    source.endColumn >= target.startColumn
  if (
    overlaps &&
    (source.startRow !== target.startRow || source.startColumn !== target.startColumn)
  ) {
    throw new Error(
      'fill_range source and target overlap — start the target at the source cell (fill-down/right includes the source as the first tile) or keep them disjoint.',
    )
  }
}

/**
 * The full destination rectangle of a copy_range: a single-cell target is
 * the paste anchor and extends to the source's size; a multi-cell target
 * must already match the source exactly.
 */
export function copyTargetBounds(operation: CopyRangeOperation): {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
} {
  const source = parseRange(operation.source)
  const target = parseRange(operation.target)
  if (rangeCellCount(target) === 1) {
    return {
      startRow: target.startRow,
      endRow: target.startRow + (source.endRow - source.startRow),
      startColumn: target.startColumn,
      endColumn: target.startColumn + (source.endColumn - source.startColumn),
    }
  }
  return target
}

/// copy_range geometry guards: one block, one destination. Overlaps are
/// rejected outright — the executor reads the source chunk by chunk while
/// writing the target, so an overlap would read back its own writes.
function validateCopyRange(operation: CopyRangeOperation): void {
  const source = parseRange(operation.source)
  if (rangeCellCount(source) > MAX_RANGE_OP_CELLS) {
    throw new Error(
      `copy_range source covers more than ${MAX_RANGE_OP_CELLS.toLocaleString('en-US')} cells — copy it in several blocks.`,
    )
  }
  const rawTarget = parseRange(operation.target)
  const target = copyTargetBounds(operation)
  if (
    rangeCellCount(rawTarget) !== 1 &&
    (target.endRow - target.startRow !== source.endRow - source.startRow ||
      target.endColumn - target.startColumn !== source.endColumn - source.startColumn)
  ) {
    throw new Error(
      `copy_range target ${operation.target} does not match the source's size — pass just the destination's top-left cell (e.g. "H1"), or a range exactly the size of ${operation.source}.`,
    )
  }
  const sameSheet =
    operation.sourceSheetId === undefined || operation.sourceSheetId === operation.sheetId
  if (
    sameSheet &&
    source.startRow <= target.endRow &&
    source.endRow >= target.startRow &&
    source.startColumn <= target.endColumn &&
    source.endColumn >= target.startColumn
  ) {
    throw new Error(
      'copy_range source and target overlap — choose a destination outside the source block (to shift data by whole rows/columns, use insert_rows/insert_cols instead).',
    )
  }
}

/**
 * Batch-order hazard guard: convert_to_values freezes what the grid holds
 * NOW, but same-batch formula writes land through a different plan lane
 * (per-cell changes apply after range-level bulk ops) and an async recalc —
 * so "write formulas, then freeze them" cannot work within one batch. The
 * demo path is worse: it expands the convert against the pre-batch grid.
 * Reject the mix and direct the writer to two batches.
 */
export function convertToValuesBatchError(operations: readonly WorkbookOperation[]): string | null {
  const converts = operations.filter((op) => op.op === 'convert_to_values')
  if (converts.length === 0) return null
  for (const convert of converts) {
    const bounds = parseRange(convert.range)
    for (const op of operations) {
      if (op === convert) continue
      let writes: { startRow: number; endRow: number; startColumn: number; endColumn: number }
      if (op.op === 'set_formula') {
        if (op.sheetId !== convert.sheetId) continue
        writes = parseRange(op.address)
      } else if (op.op === 'set_range') {
        if (op.sheetId !== convert.sheetId) continue
        // Plain-value writes commute with the convert (it only touches
        // formula cells); only "="-strings make the order observable.
        const writesFormulas = op.values.some((row) =>
          row.some((value) => typeof value === 'string' && value.startsWith('=')),
        )
        const anchor = op.range ?? op.start
        if (!writesFormulas || !anchor) continue
        const origin = parseRange(anchor)
        writes = {
          startRow: origin.startRow,
          endRow: origin.startRow + op.values.length - 1,
          startColumn: origin.startColumn,
          endColumn: origin.startColumn + (op.values[0]?.length ?? 1) - 1,
        }
      } else if (op.op === 'fill_range') {
        if (op.sheetId !== convert.sheetId) continue
        writes = parseRange(op.target)
      } else if (op.op === 'copy_range') {
        if (op.sheetId !== convert.sheetId) continue
        writes = copyTargetBounds(op)
      } else {
        continue
      }
      const overlaps =
        writes.startRow <= bounds.endRow &&
        writes.endRow >= bounds.startRow &&
        writes.startColumn <= bounds.endColumn &&
        writes.endColumn >= bounds.startColumn
      if (overlaps) {
        return (
          `convert_to_values on ${convert.range} cannot share a batch with a ${op.op} that writes ` +
          'formulas into that range — the writes would land after (or invisibly to) the convert. ' +
          'Propose the formula writes first, verify the results, then convert in a separate batch.'
        )
      }
    }
  }
  return null
}

export function fillOpLabel(op: FillRangeOperation): string {
  return `Fill ${op.source} → ${op.target}`
}

export function copyOpLabel(op: CopyRangeOperation): string {
  return `Copy ${op.source} → ${op.target}`
}

export function convertToValuesOpLabel(op: ConvertToValuesOperation): string {
  return `Convert ${op.range} to values`
}

export function clearRangeOpLabel(op: ClearRangeOperation): string {
  return `Clear ${op.range}`
}

export function findReplaceOpLabel(op: FindReplaceOperation): string {
  return `Replace "${op.find}" → "${op.replace}" in ${op.range}`
}

export function replaceOccurrences(
  text: string,
  find: string,
  replace: string,
  matchCase: boolean,
): string {
  if (matchCase) return text.split(find).join(replace)
  const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  // Callback form: a literal `$` in the replacement must stay literal.
  return text.replace(new RegExp(escaped, 'gi'), () => replace)
}

export type ExpandCellReader = (
  address: string,
  sheetId: string,
) => {
  value: string | number | boolean | null
  formula?: string | undefined
}

/// set_range guards, applied before any expansion. Jagged rows are rejected
/// because a shorter row silently leaves the old trailing cells in place —
/// the classic way a table rewrite shears its columns apart; requiring a
/// rectangle forces the writer to state the intended width (null clears a
/// cell). When the op targets a full `range`, its size must match the values
/// grid so misaligned writes fail before anything applies.
function setRangeOrigin(operation: SetRangeOperation): { startRow: number; startColumn: number } {
  const width = operation.values[0]?.length ?? 0
  const jaggedIndex = operation.values.findIndex((row) => row.length !== width)
  if (jaggedIndex !== -1) {
    throw new Error(
      `set_range values must be rectangular: row 1 has ${width} cell(s) but row ${jaggedIndex + 1} has ${operation.values[jaggedIndex]?.length}. ` +
        'Use null for cells that should be cleared, or split into separate set_range operations.',
    )
  }
  const { start, range } = operation
  if (!start && !range)
    throw new Error('set_range needs "start" — the top-left target cell, like "B2".')
  const bounds = parseRange(range ?? (start as string))
  if (start && range) {
    // Both fields together must agree; silently preferring one would let a
    // write land somewhere other than where the model believes it targeted.
    const startBounds = parseRange(start)
    if (
      startBounds.startRow !== bounds.startRow ||
      startBounds.startColumn !== bounds.startColumn
    ) {
      throw new Error(
        `set_range received both start ${start} and range ${range}, which disagree on the top-left cell — pass only one of them.`,
      )
    }
  }
  if (range) {
    const rows = bounds.endRow - bounds.startRow + 1
    const columns = bounds.endColumn - bounds.startColumn + 1
    if (rows !== operation.values.length || columns !== width) {
      throw new Error(
        `set_range range ${range} spans ${rows}×${columns} cells but values is ${operation.values.length} row(s) × ${width} cell(s) — ` +
          'make them match, or give "start" (the top-left cell) instead.',
      )
    }
  }
  return bounds
}

/// Excel formula prefix for a sheet: bare when the name is a simple
/// identifier, quoted (embedded ' doubled) otherwise. Mirrors the Consolidate
/// dialog's helper so AI-written and UI-written formulas look the same.
function sheetFormulaPrefix(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)
    ? `${sheetName}!`
    : `'${sheetName.replaceAll("'", "''")}'!`
}

/// Defined-name refs always quote the sheet ('Sheet1'!$B$2:$B$10): the
/// sheets UI writes refs that way and quoting is always legal Excel.
function absoluteQuotedRangeRef(
  sheetName: string,
  bounds: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): string {
  return (
    `'${sheetName.replaceAll("'", "''")}'!` +
    `$${columnLabel(bounds.startColumn)}$${bounds.startRow + 1}` +
    `:$${columnLabel(bounds.endColumn)}$${bounds.endRow + 1}`
  )
}

/// Excel defined-name rules applied to a label: invalid characters become
/// underscores and a name that could be read as a reference (leading digit,
/// "A1", "R1C1", the reserved "R"/"C") gets an underscore prefix. Returns
/// null for a label that is empty after trimming.
function sanitizeDefinedName(label: string): string | null {
  const trimmed = label.trim()
  if (trimmed === '') return null
  const cleaned = trimmed.replace(/[^\p{L}\p{N}_.\\]/gu, '_')
  if (cleaned === '' || /^\.+$/.test(cleaned)) return null
  // Defined names must start with a letter, underscore, or backslash.
  const named = /^[\p{L}_\\]/u.test(cleaned) ? cleaned : `_${cleaned}`
  const cellLike =
    /^[A-Za-z]{1,3}\d+$/.test(named) || /^[Rr]\d*([Cc]\d*)?$/.test(named) || /^[Cc]$/.test(named)
  // Excel name length limit is 255 characters.
  const bounded = (cellLike ? `_${named}` : named).slice(0, 255)
  return bounded
}

/// Makes a defined name unique within this expansion by appending _2, _3, …
/// (Excel names are case-insensitive, so the taken set is lowercased).
function uniqueDefinedName(base: string, taken: Set<string>): string {
  const stem = base.slice(0, 248)
  let candidate = stem
  let serial = 2
  while (taken.has(candidate.toLowerCase())) {
    candidate = `${stem}_${serial}`
    serial += 1
  }
  taken.add(candidate.toLowerCase())
  return candidate
}

/// Splits one cell's text into the pieces text-to-columns writes. A null
/// delimiter means "any run of whitespace" and trims the text; a literal
/// delimiter is escaped, and collapseConsecutive folds runs of it into one.
function splitTextColumnParts(
  text: string,
  delimiter: string | null,
  collapseConsecutive: boolean,
): string[] {
  if (delimiter === null) {
    return text
      .trim()
      .split(/\s+/)
      .filter((part) => part !== '')
  }
  if (!collapseConsecutive) return text.split(delimiter)
  const escaped = delimiter.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return text.split(new RegExp(`${escaped}+`)).filter((part) => part !== '')
}

/// set_filter_criteria carries its cross-field rule (exactly one of values /
/// conditions) here rather than as a zod refinement, because discriminated-
/// union members cannot carry refinements.
function validateFilterCriteria(operation: z.infer<typeof setFilterCriteriaSchema>): void {
  if ((operation.values !== undefined) === (operation.conditions !== undefined)) {
    throw new Error(
      'set_filter_criteria needs exactly one of "values" (checkbox list, or null to clear the column) ' +
        'and "conditions" (1–2 custom comparisons) — e.g. {"values":["East","South"]} or ' +
        '{"conditions":[{"operator":"greaterThan","value":"100"}],"and":true}.',
    )
  }
}

/// Outline spans are axis-local: "5:12" only makes sense for rows, "B:D" only
/// for columns. A1-style ranges are accepted for either axis (the executor
/// reads just that axis's extents). Rejecting the mismatch here gives the
/// model the fix instruction before a plan exists.
function validateOutlineOperation(operation: OutlineOperation): void {
  const rowSpan = /^[1-9][0-9]{0,6}(?::[1-9][0-9]{0,6})?$/.test(operation.range)
  const columnSpan = /^[A-Z]{1,3}(?::[A-Z]{1,3})?$/.test(operation.range)
  if (rowSpan && operation.axis === 'cols') {
    throw new Error(
      `${operation.op} range "${operation.range}" is a row span but axis is "cols" — pass a column span like "B:D" (or an A1 range like "B1:D1").`,
    )
  }
  if (columnSpan && operation.axis === 'rows') {
    throw new Error(
      `${operation.op} range "${operation.range}" is a column span but axis is "rows" — pass a row span like "5:12" (or an A1 range like "A5:A12").`,
    )
  }
}

/// Remove Duplicates, agent edition: read the range, keep the first occurrence
/// of each identical row (text compared case-insensitively across the full
/// range width, numbers/booleans by value — the same rule as the ribbon's
/// dedupe helper), and emit one whole-row delete_rows per later duplicate.
/// Deletes are emitted bottom-up (largest row first) so every row number stays
/// valid while the batch applies.
function expandRemoveDuplicates(
  operation: RemoveDuplicatesOperation,
  readCell: ExpandCellReader,
): PrimitiveOperation[] {
  const bounds = parseRange(operation.range)
  const cells = rangeCellCount(bounds)
  if (cells > MAX_DATA_TOOL_CELLS) {
    throw new Error(
      `remove_duplicates covers ${cells.toLocaleString('en-US')} cells — the limit is ` +
        `${MAX_DATA_TOOL_CELLS.toLocaleString('en-US')}. Remove duplicates on a smaller table or split the range.`,
    )
  }
  const firstDataRow = bounds.startRow + (operation.hasHeader === true ? 1 : 0)
  const seen = new Set<string>()
  const duplicateRows: number[] = []
  for (let row = firstDataRow; row <= bounds.endRow; row += 1) {
    const key: (string | number | boolean | null)[] = []
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      const value = readCell(formatAddress(row, column), operation.sheetId).value
      key.push(typeof value === 'string' ? value.toLowerCase() : value)
    }
    const encoded = JSON.stringify(key)
    if (seen.has(encoded)) duplicateRows.push(row)
    else seen.add(encoded)
  }
  return duplicateRows.reverse().map((row) => ({
    op: 'delete_rows',
    sheetId: operation.sheetId,
    row: row + 1,
    count: 1,
  }))
}

/// Text to Columns, agent edition. One-column sources spill parts rightwards
/// within each source row; one-row sources spill downward within each source
/// column (the range "orientation" decides the flow direction).
///
/// Clear rules, exactly:
/// - **destination == source top-left (in place)**: every source cell is
///   overwritten by its first part, and the trailing cells of each output
///   row/column that its source cell did not fill (up to the widest split) are
///   cleared with clear_cell — the stale values Excel's Text to Columns also
///   wipes when it widens a column.
/// - **destination elsewhere**: with clearSource (default true) the WHOLE
///   source is cleared first (so the data is moved, not duplicated), then the
///   parts are written; with clearSource:false the source text stays.
/// - Empty source cells are left untouched; numbers/booleans are split by
///   their text form. Formula cells are split from their computed value and
///   rejected when that value is unavailable (only the live grid knows it).
function expandTextToColumns(
  operation: TextToColumnsOperation,
  readCell: ExpandCellReader,
): PrimitiveOperation[] {
  const source = parseRange(operation.source)
  const sourceRows = source.endRow - source.startRow + 1
  const sourceColumns = source.endColumn - source.startColumn + 1
  if (sourceRows > 1 && sourceColumns > 1) {
    throw new Error(
      `text_to_columns source ${operation.source} must be a single column (spills right, e.g. "A2:A100") ` +
        'or a single row (spills down, e.g. "A2:Z2") — split one line of cells at a time.',
    )
  }
  if (rangeCellCount(source) > MAX_DATA_TOOL_CELLS) {
    throw new Error(
      `text_to_columns source covers more than ${MAX_DATA_TOOL_CELLS.toLocaleString('en-US')} cells — split it in several ranges.`,
    )
  }
  const destination =
    operation.destination === undefined
      ? { row: source.startRow, column: source.startColumn }
      : parseAddress(operation.destination)
  const columnFlow = sourceColumns === 1
  const collapseConsecutive = operation.collapseConsecutive ?? true

  interface SourceEntry {
    readonly row: number
    readonly column: number
    readonly parts: readonly string[]
  }
  const entries: SourceEntry[] = []
  let maxParts = 0
  for (let row = source.startRow; row <= source.endRow; row += 1) {
    for (let column = source.startColumn; column <= source.endColumn; column += 1) {
      const address = formatAddress(row, column)
      const state = readCell(address, operation.sheetId)
      if (state.formula !== undefined && state.value === null) {
        throw new Error(
          `text_to_columns found a formula at ${address} whose computed value is unavailable — ` +
            'convert that formula to a value first (convert_to_values), then split.',
        )
      }
      if (state.value === null) continue
      const text = typeof state.value === 'string' ? state.value : String(state.value)
      if (text.trim() === '') continue
      const parts = splitTextColumnParts(text, operation.delimiter, collapseConsecutive)
      if (parts.length === 0) continue
      entries.push({ row, column, parts })
      maxParts = Math.max(maxParts, parts.length)
    }
  }
  if (maxParts === 0) return []
  const outputRows = columnFlow ? sourceRows : maxParts
  const outputColumns = columnFlow ? maxParts : sourceColumns
  if (outputRows * outputColumns > MAX_DATA_TOOL_CELLS) {
    throw new Error(
      `text_to_columns would write ${outputRows * outputColumns} output cells (` +
        `${outputRows}×${outputColumns}) — the limit is ${MAX_DATA_TOOL_CELLS.toLocaleString('en-US')}. ` +
        'Split fewer rows at a time.',
    )
  }

  const inPlace = destination.row === source.startRow && destination.column === source.startColumn
  const expanded: PrimitiveOperation[] = []
  if (!inPlace && (operation.clearSource ?? true)) {
    // Clear before writing: an away destination may overlap the source, and
    // clearing afterwards would wipe the freshly split parts.
    for (let row = source.startRow; row <= source.endRow; row += 1) {
      for (let column = source.startColumn; column <= source.endColumn; column += 1) {
        expanded.push({
          op: 'clear_cell',
          sheetId: operation.sheetId,
          address: formatAddress(row, column),
        })
      }
    }
  }
  for (const entry of entries) {
    const rowOffset = columnFlow ? entry.row - source.startRow : 0
    const columnOffset = columnFlow ? 0 : entry.column - source.startColumn
    const positionOf = (index: number): { row: number; column: number } =>
      columnFlow
        ? { row: destination.row + rowOffset, column: destination.column + index }
        : { row: destination.row + index, column: destination.column + columnOffset }
    entry.parts.forEach((part, index) => {
      const position = positionOf(index)
      expanded.push({
        op: 'set_cell',
        sheetId: operation.sheetId,
        address: formatAddress(position.row, position.column),
        // An empty piece (leading/trailing delimiter) clears the cell, which
        // matches what typing an empty value into the grid does.
        value: part === '' ? null : part,
      })
    })
    if (!inPlace) continue
    // In-place: clear the trailing part of this output row/column that the
    // source cell did not fill (stale values from before the split).
    for (let index = entry.parts.length; index < maxParts; index += 1) {
      const position = positionOf(index)
      expanded.push({
        op: 'clear_cell',
        sheetId: operation.sheetId,
        address: formatAddress(position.row, position.column),
      })
    }
  }
  return expanded
}

const CONSOLIDATE_FUNCTION_NAMES: Record<ConsolidateOperation['function'], string> = {
  sum: 'SUM',
  count: 'COUNT',
  average: 'AVERAGE',
  max: 'MAX',
  min: 'MIN',
}

/// Consolidate, agent edition: one live formula per target cell, aggregating
/// the same offset from every reference range. All references must have the
/// same shape (that shape is the output shape). Cross-sheet refs are
/// qualified and quoted only when needed; same-sheet refs stay bare.
function expandConsolidate(
  operation: ConsolidateOperation,
  sheetNameById: (sheetId: string) => string,
): PrimitiveOperation[] {
  const first = operation.references[0]
  if (!first) throw new Error('consolidate needs at least one reference range.')
  const shape = parseRange(first.range)
  const rows = shape.endRow - shape.startRow + 1
  const columns = shape.endColumn - shape.startColumn + 1
  if (rows * columns > MAX_DATA_TOOL_CELLS) {
    throw new Error(
      `consolidate would write ${rows * columns} output cells — the limit is ` +
        `${MAX_DATA_TOOL_CELLS.toLocaleString('en-US')}. Aggregate a smaller block or split the target.`,
    )
  }
  const references = operation.references.map((reference, index) => {
    const bounds = parseRange(reference.range)
    if (
      bounds.endRow - bounds.startRow !== shape.endRow - shape.startRow ||
      bounds.endColumn - bounds.startColumn !== shape.endColumn - shape.startColumn
    ) {
      throw new Error(
        `consolidate reference #${index + 1} (${reference.range}) is ` +
          `${bounds.endRow - bounds.startRow + 1}×${bounds.endColumn - bounds.startColumn + 1} but ` +
          `reference #1 (${first.range}) is ${rows}×${columns} — every reference must have the same shape, ` +
          'because each target cell aggregates the same offset from each source.',
      )
    }
    if (reference.sheetId === undefined || reference.sheetId === operation.sheetId) {
      return { bounds, prefix: '' }
    }
    return { bounds, prefix: sheetFormulaPrefix(sheetNameById(reference.sheetId)) }
  })
  const target = parseAddress(operation.target)
  const fnName = CONSOLIDATE_FUNCTION_NAMES[operation.function]
  const expanded: PrimitiveOperation[] = []
  for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < columns; columnOffset += 1) {
      const args = references.map(
        (reference) =>
          `${reference.prefix}${formatAddress(
            reference.bounds.startRow + rowOffset,
            reference.bounds.startColumn + columnOffset,
          )}`,
      )
      expanded.push({
        op: 'set_formula',
        sheetId: operation.sheetId,
        address: formatAddress(target.row + rowOffset, target.column + columnOffset),
        formula: `=${fnName}(${args.join(',')})`,
      })
    }
  }
  return expanded
}

/// Create Names from Selection, agent edition: read the label row/column and
/// emit one add_defined_name per adjacent data column/row, with absolute A1
/// refs. Invalid characters become _, a leading digit / cell-like label gets
/// an _ prefix, and duplicate names get a numeric suffix (_2, _3, …).
/// Labels are read from the cell VALUE (the file's stored value); a label
/// formula whose computed value is unavailable counts as empty and is skipped.
function expandCreateNames(
  operation: CreateNamesOperation,
  readCell: ExpandCellReader,
  sheetNameById: (sheetId: string) => string,
): PrimitiveOperation[] {
  const bounds = parseRange(operation.range)
  const rows = bounds.endRow - bounds.startRow + 1
  const columns = bounds.endColumn - bounds.startColumn + 1
  const tableStyle = operation.label === 'top' || operation.label === 'bottom'
  if (tableStyle ? rows < 2 : columns < 2) {
    throw new Error(
      tableStyle
        ? `create_names_from_selection with label "${operation.label}" needs at least two rows: one label row plus at least one data row.`
        : `create_names_from_selection with label "${operation.label}" needs at least two columns: one label column plus at least one data column.`,
    )
  }
  const sheetName = sheetNameById(operation.sheetId)
  const labelText = (address: string): string => {
    const value = readCell(address, operation.sheetId).value
    if (value === null) return ''
    return typeof value === 'string' ? value : String(value)
  }
  const entries: { label: string; ref: string }[] = []
  if (tableStyle) {
    const labelRow = operation.label === 'top' ? bounds.startRow : bounds.endRow
    const dataStartRow = operation.label === 'top' ? bounds.startRow + 1 : bounds.startRow
    const dataEndRow = operation.label === 'top' ? bounds.endRow : bounds.endRow - 1
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      entries.push({
        label: labelText(formatAddress(labelRow, column)),
        ref: absoluteQuotedRangeRef(sheetName, {
          startRow: dataStartRow,
          endRow: dataEndRow,
          startColumn: column,
          endColumn: column,
        }),
      })
    }
  } else {
    const labelColumn = operation.label === 'left' ? bounds.startColumn : bounds.endColumn
    const dataStartColumn = operation.label === 'left' ? bounds.startColumn + 1 : bounds.startColumn
    const dataEndColumn = operation.label === 'left' ? bounds.endColumn : bounds.endColumn - 1
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      entries.push({
        label: labelText(formatAddress(row, labelColumn)),
        ref: absoluteQuotedRangeRef(sheetName, {
          startRow: row,
          endRow: row,
          startColumn: dataStartColumn,
          endColumn: dataEndColumn,
        }),
      })
    }
  }
  const taken = new Set<string>()
  const expanded: PrimitiveOperation[] = []
  for (const entry of entries) {
    const base = sanitizeDefinedName(entry.label)
    if (base === null) continue
    expanded.push({
      op: 'add_defined_name',
      name: uniqueDefinedName(base, taken),
      ref: entry.ref,
    })
  }
  if (expanded.length === 0) {
    throw new Error(
      `create_names_from_selection found no non-empty labels in the ${operation.label} ` +
        `row/column of ${operation.range} — check the range and the label direction.`,
    )
  }
  return expanded
}

/// Expands set_range/clear_range into per-cell primitives so the preview,
/// CAS checks, and both apply paths keep working on single cells. sort_range,
/// find_replace, remove_duplicates, text_to_columns, consolidate, and
/// create_names_from_selection additionally need `readCell` (the current grid
/// contents); consolidate/create_names_from_selection need `sheetNameById`
/// to qualify cross-sheet formula/defined-name references — callers pass a
/// real sheetId→sheetName resolver from the open workbook.
export function expandToPrimitiveOps(
  operations: readonly WorkbookOperation[],
  readCell?: ExpandCellReader,
  sheetNameById?: (sheetId: string) => string,
): PrimitiveOperation[] {
  const convertBatchError = convertToValuesBatchError(operations)
  if (convertBatchError) throw new Error(convertBatchError)
  const expanded: PrimitiveOperation[] = []
  let cellOps = 0
  const countCell = (): void => {
    cellOps += 1
    if (cellOps > MAX_EXPANDED_CELL_OPS) {
      throw new Error(`The batch expands to more than ${MAX_EXPANDED_CELL_OPS} cell edits.`)
    }
  }
  for (const operation of operations) {
    if (operation.op === 'set_range') {
      const origin = setRangeOrigin(operation)
      operation.values.forEach((rowValues, rowOffset) => {
        rowValues.forEach((value, columnOffset) => {
          countCell()
          const address = formatAddress(
            origin.startRow + rowOffset,
            origin.startColumn + columnOffset,
          )
          if (typeof value === 'string' && value.startsWith('=')) {
            expanded.push({
              op: 'set_formula',
              sheetId: operation.sheetId,
              address,
              formula: value,
            })
          } else {
            expanded.push({ op: 'set_cell', sheetId: operation.sheetId, address, value })
          }
        })
      })
    } else if (operation.op === 'clear_range') {
      const bounds = parseRange(operation.range)
      const cells = rangeCellCount(bounds)
      if (cells > MAX_RANGE_OP_CELLS) {
        throw new Error(
          `clear_range covers more than ${MAX_RANGE_OP_CELLS.toLocaleString('en-US')} cells — clear it in several ranges.`,
        )
      }
      if (cells > MAX_EXPANDED_CELL_OPS) {
        // Large clears stay range-level: per-cell expansion would blow the
        // preview/apply paths; executors clear the whole range in one call.
        expanded.push(operation)
        continue
      }
      for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
        for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
          countCell()
          expanded.push({
            op: 'clear_cell',
            sheetId: operation.sheetId,
            address: formatAddress(row, column),
          })
        }
      }
    } else if (operation.op === 'fill_range') {
      validateFillRange(operation)
      expanded.push(operation)
    } else if (operation.op === 'copy_range') {
      validateCopyRange(operation)
      expanded.push(operation)
    } else if (operation.op === 'convert_to_values') {
      // Range-level: the executor reads each cell's computed value from the
      // live grid (chunk-loading streamed regions first) — expansion here
      // could not see computed values, only stored ones.
      if (rangeCellCount(parseRange(operation.range)) > MAX_RANGE_OP_CELLS) {
        throw new Error(
          `convert_to_values covers more than ${MAX_RANGE_OP_CELLS.toLocaleString('en-US')} cells — convert it in several ranges.`,
        )
      }
      expanded.push(operation)
    } else if (operation.op === 'format_range') {
      // Range-level all the way through (never expanded per cell), so whole
      // columns of large files are fine up to the range-op cap.
      if (rangeCellCount(parseRange(operation.range)) > MAX_RANGE_OP_CELLS) {
        throw new Error(
          `format_range covers more than ${MAX_RANGE_OP_CELLS.toLocaleString('en-US')} cells.`,
        )
      }
      expanded.push(operation)
    } else if (operation.op === 'clear_formats') {
      // One range-level format_range carrying every clearing value: the same
      // shape the format engine already applies, so no new executor branch.
      const cells = rangeCellCount(parseRange(operation.range))
      if (cells > MAX_DATA_TOOL_CELLS) {
        throw new Error(
          `clear_formats covers ${cells.toLocaleString('en-US')} cells — the limit is ` +
            `${MAX_DATA_TOOL_CELLS.toLocaleString('en-US')}. Clear it in several ranges.`,
        )
      }
      expanded.push({
        op: 'format_range',
        sheetId: operation.sheetId,
        range: operation.range,
        // Fresh object per expansion: the patch is embedded in plans that a
        // future consumer could normalize in place, and sharing one module
        // constant would leak that mutation into every later clear_formats.
        format: { ...CLEARED_FORMAT_PATCH, border: { type: 'none' } },
      })
    } else if (operation.op === 'find_replace') {
      const bounds = parseRange(operation.range)
      const cells = rangeCellCount(bounds)
      if (cells > MAX_RANGE_OP_CELLS) {
        throw new Error(
          `find_replace covers more than ${MAX_RANGE_OP_CELLS.toLocaleString('en-US')} cells — replace in several ranges.`,
        )
      }
      if (cells > MAX_EXPANDED_CELL_OPS) {
        // Large replaces stay range-level: the executor scans loaded chunks
        // and rewrites only the matching cells (no per-cell preview).
        expanded.push(operation)
        continue
      }
      if (!readCell)
        throw new Error('find_replace needs the current cell contents to plan against.')
      const matchCase = operation.matchCase ?? false
      const needle = matchCase ? operation.find : operation.find.toLowerCase()
      for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
        for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
          const address = formatAddress(row, column)
          const current = readCell(address, operation.sheetId)
          if (current.formula !== undefined || typeof current.value !== 'string') continue
          const haystack = matchCase ? current.value : current.value.toLowerCase()
          let next: string | null = null
          if (operation.wholeCell) {
            if (haystack === needle) next = operation.replace
          } else if (haystack.includes(needle)) {
            next = replaceOccurrences(current.value, operation.find, operation.replace, matchCase)
          }
          if (next === null || next === current.value) continue
          countCell()
          expanded.push({ op: 'set_cell', sheetId: operation.sheetId, address, value: next })
        }
      }
    } else if (operation.op === 'sort_range') {
      if (!readCell) throw new Error('sort_range needs the current cell contents to plan against.')
      if (rangeCellCount(parseRange(operation.range)) > MAX_EXPANDED_CELL_OPS) {
        throw new Error(`sort_range covers more than ${MAX_EXPANDED_CELL_OPS} cells.`)
      }
      if (
        operation.keys === undefined &&
        (operation.byColumn === undefined || operation.order === undefined)
      ) {
        throw new Error(
          'sort_range needs a sort key: pass "keys" (1–3 entries, e.g. ' +
            '[{"byColumn":"B","order":"desc"}]) or both "byColumn" and "order".',
        )
      }
      const hasHeader = operation.hasHeader ?? false
      const spec: SortSpec =
        operation.keys !== undefined
          ? {
              range: operation.range,
              hasHeader,
              keys: operation.keys.map((key) => ({
                byColumn: key.byColumn,
                ascending: key.order === 'asc',
              })),
            }
          : {
              range: operation.range,
              hasHeader,
              byColumn: operation.byColumn as string,
              ascending: operation.order === 'asc',
            }
      const changes = computeSortChanges(spec, (address) => readCell(address, operation.sheetId))
      for (const change of changes) {
        countCell()
        expanded.push({
          op: 'set_cell',
          sheetId: operation.sheetId,
          address: change.address,
          value: change.after,
          expectedValue: change.before,
        })
      }
    } else if (operation.op === 'remove_duplicates') {
      if (!readCell)
        throw new Error('remove_duplicates needs the current cell contents to plan against.')
      expanded.push(...expandRemoveDuplicates(operation, readCell))
    } else if (operation.op === 'text_to_columns') {
      if (!readCell)
        throw new Error('text_to_columns needs the current cell contents to plan against.')
      // These data tools carry their own MAX_DATA_TOOL_CELLS budget instead of
      // the per-cell MAX_EXPANDED_CELL_OPS preview ceiling: a bulk split or
      // aggregation legitimately writes more than 2,000 plain cells.
      expanded.push(...expandTextToColumns(operation, readCell))
    } else if (operation.op === 'consolidate') {
      if (!sheetNameById)
        throw new Error(
          'consolidate needs sheet names to build cross-sheet references — open the workbook through a path that resolves sheet ids to names.',
        )
      expanded.push(...expandConsolidate(operation, sheetNameById))
    } else if (operation.op === 'create_names_from_selection') {
      if (!readCell)
        throw new Error(
          'create_names_from_selection needs the current cell contents to read its labels.',
        )
      if (!sheetNameById)
        throw new Error(
          'create_names_from_selection needs the sheet name to write absolute defined-name references.',
        )
      expanded.push(...expandCreateNames(operation, readCell, sheetNameById))
    } else if (operation.op === 'subtotal') {
      // Structural pass-through: the executor inserts the summary rows through
      // Univer. Validate the geometry here so the model gets a fix instruction
      // before a plan exists.
      const bounds = parseRange(operation.range)
      const rows = bounds.endRow - bounds.startRow + 1
      if (rows < 2) {
        throw new Error(
          `subtotal range ${operation.range} needs at least two rows — a header row plus at least one data row.`,
        )
      }
      const groupColumn = columnIndex(operation.groupColumn)
      const valueColumn = columnIndex(operation.valueColumn)
      if (groupColumn < bounds.startColumn || groupColumn > bounds.endColumn) {
        throw new Error(
          `subtotal groupColumn ${operation.groupColumn} is outside the range ${operation.range} — ` +
            "the grouping column must be one of the range's columns.",
        )
      }
      if (valueColumn < bounds.startColumn || valueColumn > bounds.endColumn) {
        throw new Error(
          `subtotal valueColumn ${operation.valueColumn} is outside the range ${operation.range} — ` +
            "the subtotaled column must be one of the range's columns.",
        )
      }
      if (groupColumn === valueColumn) {
        throw new Error(
          `subtotal groupColumn and valueColumn are both ${operation.groupColumn} — ` +
            'group by one column and subtotal a different (numeric) column.',
        )
      }
      expanded.push(operation)
    } else if (
      operation.op === 'outline.group' ||
      operation.op === 'outline.ungroup' ||
      operation.op === 'outline.hideDetail' ||
      operation.op === 'outline.showDetail'
    ) {
      // Layout-class pass-through: the executor edits the outline journal and
      // row/column visibility. Only axis/span consistency is validated here.
      validateOutlineOperation(operation)
      expanded.push(operation)
    } else if (operation.op === 'add_pivot') {
      const columnFieldsArray =
        operation.columnField === undefined
          ? []
          : Array.isArray(operation.columnField)
            ? operation.columnField
            : [operation.columnField]
      if (columnFieldsArray.length > 0 && operation.values.length !== 1) {
        throw new Error('With a columnField the pivot supports exactly one values entry.')
      }
      if (new Set(columnFieldsArray).size !== columnFieldsArray.length) {
        throw new Error('columnField must not repeat a field.')
      }
      const rowFieldsArray = Array.isArray(operation.rowFields)
        ? operation.rowFields
        : [operation.rowFields]
      if (new Set(rowFieldsArray).size !== rowFieldsArray.length) {
        throw new Error('rowFields must not repeat a field.')
      }
      if (columnFieldsArray.some((field) => rowFieldsArray.includes(field))) {
        throw new Error('A field cannot be both a row and a column dimension.')
      }
      const allDimensionFields = [
        ...rowFieldsArray,
        ...columnFieldsArray,
        ...(operation.pageFields ?? []),
      ]
      if (operation.values.some((value) => allDimensionFields.includes(value.field))) {
        throw new Error('A values field cannot also be a row, column, or page filter field.')
      }
      // Calculated fields always aggregate with SUM (the formula operates on each
      // referenced field's in-group sum).
      if (operation.values.some((value) => value.formula !== undefined && value.agg !== 'sum')) {
        throw new Error('A calculated field must use agg "sum".')
      }
      // Grouping may only target row/column dimension fields, with at most one
      // rule per field.
      const axisFields = [...rowFieldsArray, ...columnFieldsArray]
      const groupedFields = (operation.groupings ?? []).map((grouping) => grouping.field)
      if (groupedFields.some((field) => !axisFields.includes(field))) {
        throw new Error('A grouping field must be a row or column dimension field.')
      }
      if (new Set(groupedFields).size !== groupedFields.length) {
        throw new Error('A field can only have one grouping rule.')
      }
      // Filters may only target row/column dimension fields, at most one per
      // field; value-filter params must be complete for the op, and the
      // referenced values entry must exist.
      const filteredFields = (operation.filters ?? []).map((filter) => filter.field)
      if (filteredFields.some((field) => !axisFields.includes(field))) {
        throw new Error('A filter field must be a row or column dimension field.')
      }
      if (new Set(filteredFields).size !== filteredFields.length) {
        throw new Error('A field can only have one filter.')
      }
      for (const filter of operation.filters ?? []) {
        if (filter.kind !== 'value') continue
        if (filter.valueIndex >= operation.values.length) {
          throw new Error(`Filter valueIndex ${filter.valueIndex} is out of range.`)
        }
        if (filter.op === 'top' && filter.count === undefined) {
          throw new Error('A top-N value filter needs "count".')
        }
        if (filter.op === 'greaterThan' && filter.from === undefined) {
          throw new Error('A greaterThan value filter needs "from".')
        }
        if (filter.op === 'between' && (filter.from === undefined || filter.to === undefined)) {
          throw new Error('A between value filter needs "from" and "to".')
        }
      }
      // Normalize rowFields → always array for downstream consumers
      expanded.push({
        ...operation,
        rowFields: rowFieldsArray,
      })
    } else if (operation.op === 'edit_shape') {
      if (
        operation.text === undefined &&
        operation.fillColor === undefined &&
        operation.anchorCell === undefined
      ) {
        throw new Error('edit_shape needs at least one of text / fillColor / anchorCell.')
      }
      expanded.push(operation)
    } else if (operation.op === 'edit_chart') {
      if (
        operation.title === undefined &&
        operation.chartType === undefined &&
        (!operation.seriesColors || Object.keys(operation.seriesColors).length === 0) &&
        (!operation.pointColors || Object.keys(operation.pointColors).length === 0) &&
        operation.legend === undefined &&
        operation.axisTitles === undefined &&
        operation.dataLabels === undefined &&
        operation.dataLabelPosition === undefined &&
        operation.dataLabelFormat === undefined &&
        operation.grouping === undefined &&
        operation.gridlines === undefined &&
        operation.valueAxis === undefined &&
        (!operation.seriesData || operation.seriesData.length === 0)
      ) {
        throw new Error(
          'edit_chart needs at least one of title / chartType / seriesColors / pointColors / legend / dataLabels / dataLabelPosition / dataLabelFormat / grouping / gridlines / valueAxis / axisTitles / seriesData.',
        )
      }
      for (const entry of operation.seriesData ?? []) {
        if (
          entry.name === undefined &&
          entry.valuesRange === undefined &&
          entry.categoriesRange === undefined
        ) {
          throw new Error(`seriesData[index=${entry.index}] needs a name or a data range.`)
        }
      }
      expanded.push(operation)
    } else if (operation.op === 'add_conditional_format') {
      const rule = operation.rule
      if (
        rule.kind === 'number' &&
        (rule.operator === 'between' || rule.operator === 'notBetween') &&
        rule.value2 === undefined
      ) {
        throw new Error(`add_conditional_format with operator "${rule.operator}" needs value2.`)
      }
      expanded.push(operation)
    } else if (operation.op === 'set_data_validation') {
      const rule = operation.validation
      if (rule?.kind === 'numberBetween' && rule.min > rule.max) {
        throw new Error('set_data_validation numberBetween needs min ≤ max.')
      }
      if (rule?.kind === 'dateBetween' && rule.start > rule.end) {
        throw new Error('set_data_validation dateBetween needs start ≤ end.')
      }
      expanded.push(operation)
    } else if (operation.op === 'set_page_setup') {
      const { op: _op, sheetId: _sheetId, ...settings } = operation
      if (Object.values(settings).every((value) => value === undefined)) {
        throw new Error('set_page_setup needs at least one setting.')
      }
      if (
        operation.scale !== undefined &&
        (operation.fitToWidth !== undefined || operation.fitToHeight !== undefined)
      ) {
        throw new Error('set_page_setup: scale and fitToWidth/fitToHeight are mutually exclusive.')
      }
      expanded.push(operation)
    } else if (operation.op === 'set_theme') {
      if (
        operation.theme === undefined &&
        operation.colors === undefined &&
        operation.fonts === undefined
      ) {
        throw new Error(
          'set_theme needs at least one of "theme" (a full theme preset id), "colors" (a color-scheme preset id), ' +
            'or "fonts" (a font-scheme preset id) — read the available ids with read_workbook_context first.',
        )
      }
      expanded.push(operation)
    } else if (operation.op === 'set_filter_criteria') {
      validateFilterCriteria(operation)
      expanded.push(operation)
    } else {
      if (CELL_CONTENT_OPS.has(operation.op)) countCell()
      expanded.push(operation)
    }
  }
  return expanded
}

const FORMAT_FIELD_LABELS: Record<string, string> = {
  bold: 'bold',
  italic: 'italic',
  underline: 'underline',
  strikethrough: 'strikethrough',
  fontFamily: 'font',
  fontSize: 'font size',
  fontColor: 'font color',
  fillColor: 'fill',
  numberFormat: 'number format',
  horizontalAlign: 'align',
  verticalAlign: 'vertical align',
  wrapText: 'wrap',
  textRotation: 'rotation',
  indent: 'indent',
  locked: 'locked',
  hidden: 'hidden',
}

/// True when a patch is exactly what clear_formats expands to (every format
/// field present, in its cleared state) — labelled as one "Clear formatting"
/// line instead of a dozen "no bold / clear font …" fragments.
function isClearedFormatPatch(format: CellFormatPatch): boolean {
  const entries = Object.entries(CLEARED_FORMAT_PATCH)
  if (Object.keys(format).length !== entries.length) return false
  return entries.every(([key, cleared]) => {
    const value = (format as Record<string, unknown>)[key]
    if (key === 'border') return JSON.stringify(value) === JSON.stringify(cleared)
    return value === cleared
  })
}

export function formatOpLabel(op: FormatRangeOperation): string {
  if (isClearedFormatPatch(op.format)) return `Clear formatting on ${op.range}`
  const parts = Object.entries(op.format)
    .filter(([, value]) => value !== undefined)
    .map(([key, value]) => {
      const name = FORMAT_FIELD_LABELS[key] ?? key
      if (key === 'border') {
        const border = value as BorderPatch
        if (border.type === 'none') return 'clear borders'
        return `border ${border.type}${border.style ? ` ${border.style}` : ''}${
          border.color ? ` ${border.color}` : ''
        }`
      }
      if (key === 'locked') {
        if (value === null) return 'reset locked'
        return value === true ? 'locked' : 'unlocked'
      }
      if (key === 'hidden') {
        if (value === null) return 'reset hidden'
        return value === true ? 'hidden' : 'not hidden'
      }
      if (value === null) return `clear ${name}`
      if (value === true) return name
      if (value === false) return `no ${name}`
      return `${name} ${String(value)}`
    })
  return `${op.range}: ${parts.join(', ')}`
}

const FILTER_OPERATOR_SYMBOLS: Record<z.infer<typeof filterConditionSchema>['operator'], string> = {
  equal: '=',
  notEqual: '<>',
  greaterThan: '>',
  greaterThanOrEqual: '>=',
  lessThan: '<',
  lessThanOrEqual: '<=',
}

export function layoutOpLabel(op: LayoutOperation): string {
  switch (op.op) {
    case 'delete_visual':
      return `Delete visual ${op.visualId}`
    case 'delete_table':
      return `Delete table "${op.tableName}"`
    case 'add_sparkline':
      return `Insert ${op.type} sparklines for ${op.dataRange}`
    case 'merge_cells':
      return `Merge ${op.range}`
    case 'unmerge_cells':
      return `Unmerge ${op.range}`
    case 'set_row_height':
      return op.count === 1
        ? `Set row ${op.row} height to ${op.heightPoints}pt`
        : `Set rows ${op.row}–${op.row + op.count - 1} height to ${op.heightPoints}pt`
    case 'set_col_width':
      return op.count === 1
        ? `Set column ${op.column} width to ${op.widthPx}px`
        : `Set columns ${op.column}–${columnLabel(columnIndex(op.column) + op.count - 1)} width to ${op.widthPx}px`
    case 'edit_chart': {
      const parts: string[] = []
      if (op.title !== undefined) parts.push(`title "${op.title}"`)
      if (op.chartType !== undefined) parts.push(`type ${op.chartType}`)
      if (op.seriesColors && Object.keys(op.seriesColors).length > 0) {
        parts.push(
          `series colors ${Object.entries(op.seriesColors)
            .map(([i, c]) => `#${i}→${c}`)
            .join(', ')}`,
        )
      }
      if (op.pointColors && Object.keys(op.pointColors).length > 0) {
        const count = Object.values(op.pointColors).reduce(
          (total, points) => total + Object.keys(points).length,
          0,
        )
        parts.push(`${count} point color${count === 1 ? '' : 's'}`)
      }
      if (op.legend !== undefined) {
        parts.push(op.legend === 'none' ? 'hide legend' : `legend ${op.legend}`)
      }
      if (op.axisTitles?.category !== undefined) {
        parts.push(
          op.axisTitles.category === null
            ? 'clear category-axis title'
            : `category axis "${op.axisTitles.category}"`,
        )
      }
      if (op.axisTitles?.value !== undefined) {
        parts.push(
          op.axisTitles.value === null
            ? 'clear value-axis title'
            : `value axis "${op.axisTitles.value}"`,
        )
      }
      if (op.dataLabelPosition !== undefined) {
        parts.push(`label position ${op.dataLabelPosition}`)
      }
      if (op.dataLabelFormat !== undefined) {
        parts.push(`label format "${op.dataLabelFormat}"`)
      }
      if (op.gridlines !== undefined) {
        parts.push(op.gridlines ? 'gridlines on' : 'gridlines off')
      }
      if (op.valueAxis !== undefined) {
        const bounds: string[] = []
        if (op.valueAxis.min !== undefined) {
          bounds.push(op.valueAxis.min === null ? 'min auto' : `min ${op.valueAxis.min}`)
        }
        if (op.valueAxis.max !== undefined) {
          bounds.push(op.valueAxis.max === null ? 'max auto' : `max ${op.valueAxis.max}`)
        }
        parts.push(`value axis ${bounds.join(' / ')}`)
      }
      for (const entry of op.seriesData ?? []) {
        const detail = [
          entry.name === undefined ? '' : `name "${entry.name}"`,
          entry.valuesRange === undefined ? '' : `values ${entry.valuesRange}`,
          entry.categoriesRange === undefined ? '' : `categories ${entry.categoriesRange}`,
        ]
          .filter(Boolean)
          .join(' ')
        parts.push(`series #${entry.index} ${detail}`)
      }
      return `Edit chart ${op.chartPath}: ${parts.join(', ')}`
    }
    case 'add_chart':
      return `Insert ${op.chartType} chart from ${op.dataRange}${op.title ? ` "${op.title}"` : ''}`
    case 'add_shape':
      return `Insert ${op.shapeType === 'textbox' ? 'text box' : `${op.shapeType} shape`} at ${op.anchorCell}`
    case 'edit_shape': {
      const parts = [
        op.text === undefined
          ? ''
          : `text "${op.text.length > 40 ? `${op.text.slice(0, 40)}…` : op.text}"`,
        op.fillColor === undefined ? '' : `fill ${op.fillColor}`,
        op.anchorCell === undefined ? '' : `move to ${op.anchorCell}`,
      ].filter(Boolean)
      return `Edit shape ${op.visualId}: ${parts.join(', ')}`
    }
    case 'add_image':
      return `Insert image ${op.path.split('/').pop() ?? op.path} at ${op.anchorCell}`
    case 'add_table':
      return `Create table${op.name ? ` ${op.name}` : ''} over ${op.range}`
    case 'add_table_row':
      return `Insert ${op.count ?? 1} row(s) into table ${op.tableName}${op.row !== undefined ? ` at position ${op.row}` : ' (append)'}`
    case 'add_table_column':
      return `Insert column "${op.columnName}" into table ${op.tableName}`
    case 'delete_table_row':
      return `Delete ${op.count ?? 1} row(s) from table ${op.tableName} at position ${op.row}`
    case 'delete_table_column':
      return `Delete ${op.count ?? 1} column(s) from table ${op.tableName} at position ${op.column}`
    case 'add_pivot': {
      const rowFieldsArray = Array.isArray(op.rowFields) ? op.rowFields : [op.rowFields]
      const columnFieldsArray =
        op.columnField === undefined
          ? []
          : Array.isArray(op.columnField)
            ? op.columnField
            : [op.columnField]
      return (
        `Create pivot${op.name ? ` ${op.name}` : ''} from ${op.sourceRange} ` +
        `(rows: ${rowFieldsArray.join(' > ')}${columnFieldsArray.length > 0 ? `, columns: ${columnFieldsArray.join(' > ')}` : ''}` +
        `${op.pageFields && op.pageFields.length > 0 ? `, filters: ${op.pageFields.join(', ')}` : ''}, ` +
        `values: ${op.values.map((value) => `${value.agg} ${value.field}`).join(', ')}) at ${op.targetCell}`
      )
    }
    case 'set_rows_hidden':
      return `${op.hidden ? 'Hide' : 'Unhide'} row${op.count === 1 ? ` ${op.row}` : `s ${op.row}–${op.row + op.count - 1}`}`
    case 'set_cols_hidden':
      return `${op.hidden ? 'Hide' : 'Unhide'} column${
        op.count === 1
          ? ` ${op.column}`
          : `s ${op.column}–${columnLabel(columnIndex(op.column) + op.count - 1)}`
      }`
    case 'outline.group':
      return `Group ${op.axis === 'rows' ? 'rows' : 'columns'} ${op.range}`
    case 'outline.ungroup':
      return `Ungroup ${op.axis === 'rows' ? 'rows' : 'columns'} ${op.range}`
    case 'outline.hideDetail':
      return `Hide detail ${op.axis === 'rows' ? 'rows' : 'columns'} ${op.range}`
    case 'outline.showDetail':
      return `Show detail ${op.axis === 'rows' ? 'rows' : 'columns'} ${op.range}`
    case 'set_hyperlink':
      return op.target === null
        ? `Remove link at ${op.address}`
        : `Link ${op.address} → ${op.target}`
    case 'protect_sheet':
      return op.protected ? 'Protect sheet' : 'Unprotect sheet'
    case 'protect_workbook':
      return op.protected ? 'Protect workbook structure' : 'Unprotect workbook structure'
    case 'set_theme': {
      const parts: string[] = []
      if (op.theme !== undefined) parts.push(op.theme)
      if (op.colors !== undefined) parts.push(`colors ${op.colors}`)
      if (op.fonts !== undefined) parts.push(`fonts ${op.fonts}`)
      return `Apply theme ${parts.join(', ')}`
    }
    case 'allow_edit_ranges':
      if (op.ranges.length === 0) return 'Clear allowed edit ranges'
      return (
        `Allow editing ${op.ranges.length} range${op.ranges.length === 1 ? '' : 's'}: ` +
        `${op.ranges
          .slice(0, 3)
          .map((range) => range.name ?? range.sqref)
          .join(', ')}${op.ranges.length > 3 ? ', …' : ''}`
      )
    case 'set_filter':
      return `Add auto-filter on ${op.range}`
    case 'clear_filter':
      return 'Remove auto-filter'
    case 'set_filter_criteria':
      if (op.conditions !== undefined) {
        const joiner = (op.and ?? true) ? ' AND ' : ' OR '
        return `Filter column ${op.column}: ${op.conditions
          .map((condition) => `${FILTER_OPERATOR_SYMBOLS[condition.operator]} ${condition.value}`)
          .join(joiner)}`
      }
      return op.values === null || op.values === undefined
        ? `Clear filter criteria on column ${op.column}`
        : `Filter column ${op.column} to ${op.values.length} value${op.values.length === 1 ? '' : 's'}`
    case 'add_conditional_format':
      return `Conditional format ${op.range}: ${cfRuleLabel(op.rule)}`
    case 'clear_conditional_formats':
      return 'Clear conditional formats'
    case 'set_data_validation':
      return op.validation === null
        ? `Clear data validation on ${op.range}`
        : `Data validation ${op.range}: ${dvRuleLabel(op.validation)}`
    case 'add_defined_name':
      return `Define name ${op.name} = ${op.ref}`
    case 'delete_defined_name':
      return `Delete defined name ${op.name}`
    case 'set_page_setup': {
      const parts: string[] = []
      if (op.orientation !== undefined) parts.push(op.orientation)
      if (op.paperSize !== undefined) parts.push(`paper ${op.paperSize}`)
      if (op.scale !== undefined) parts.push(`scale ${op.scale}%`)
      if (op.fitToWidth !== undefined) parts.push(`fit width ${op.fitToWidth || 'auto'}`)
      if (op.fitToHeight !== undefined) parts.push(`fit height ${op.fitToHeight || 'auto'}`)
      if (op.margins !== undefined) parts.push(`${op.margins} margins`)
      if (op.printGridlines !== undefined)
        parts.push(`${op.printGridlines ? 'print' : 'no'} gridlines`)
      if (op.printHeadings !== undefined)
        parts.push(`${op.printHeadings ? 'print' : 'no'} headings`)
      if (op.printArea !== undefined)
        parts.push(op.printArea === null ? 'clear print area' : `print area ${op.printArea}`)
      return `Page setup: ${parts.join(', ')}`
    }
    case 'set_freeze':
      return op.rows === 0 && op.columns === 0
        ? 'Unfreeze panes'
        : `Freeze ${[
            op.rows > 0 ? `first ${op.rows} row${op.rows === 1 ? '' : 's'}` : '',
            op.columns > 0 ? `first ${op.columns} column${op.columns === 1 ? '' : 's'}` : '',
          ]
            .filter(Boolean)
            .join(' and ')}`
    case 'set_note':
      return op.text === null
        ? `Remove note at ${op.address}`
        : `Note at ${op.address}: "${op.text.length > 60 ? `${op.text.slice(0, 60)}…` : op.text}"`
    case 'refresh_pivot':
      return 'Refresh pivot tables'
  }
}

function cfRuleLabel(rule: CfRule): string {
  switch (rule.kind) {
    case 'number':
      return `${rule.operator} ${rule.value}${rule.value2 === undefined ? '' : ` and ${rule.value2}`}`
    case 'text':
      return `${rule.operator} "${rule.text}"`
    case 'blank':
      return rule.blank ? 'is empty' : 'is not empty'
    case 'duplicate':
      return rule.unique ? 'unique values' : 'duplicate values'
    case 'top10':
      return `${rule.bottom ? 'bottom' : 'top'} ${rule.rank}${rule.percent ? '%' : ''}`
    case 'formula':
      return rule.formula
    case 'colorScale':
      return `color scale ${rule.minColor} → ${rule.maxColor}`
    case 'iconSet':
      return rule.iconSet
    case 'dataBar':
      return `data bar${rule.color ? ` ${rule.color}` : ''}`
  }
}

function dvRuleLabel(rule: DvRule): string {
  switch (rule.kind) {
    case 'list':
      return `list (${rule.values.slice(0, 5).join(', ')}${rule.values.length > 5 ? ', …' : ''})`
    case 'listRef':
      return `list from ${rule.range}`
    case 'numberBetween':
      return `number between ${rule.min} and ${rule.max}`
    case 'dateBetween':
      return `date between ${rule.start} and ${rule.end}`
    case 'checkbox':
      return 'checkbox'
    case 'formula':
      return `custom ${rule.formula}`
  }
}

export function structuralOpLabel(op: StructuralOperation | SubtotalOperation): string {
  switch (op.op) {
    case 'subtotal':
      return `Insert SUBTOTAL ${op.function} rows on ${op.range} grouped by ${op.groupColumn}`
    case 'insert_rows':
      return op.count === 1
        ? `Insert 1 row before row ${op.row}`
        : `Insert ${op.count} rows before row ${op.row}`
    case 'delete_rows':
      return op.count === 1
        ? `Delete row ${op.row}`
        : `Delete rows ${op.row}–${op.row + op.count - 1}`
    case 'insert_cols':
      return op.count === 1
        ? `Insert 1 column before column ${op.column}`
        : `Insert ${op.count} columns before column ${op.column}`
    case 'delete_cols':
      return op.count === 1
        ? `Delete column ${op.column}`
        : `Delete columns ${op.column}–${columnLabel(columnIndex(op.column) + op.count - 1)}`
    case 'add_sheet':
      return `Add sheet "${op.name}"`
    case 'delete_sheet':
      return `Delete sheet ${op.sheetId}`
    case 'duplicate_sheet':
      return `Duplicate sheet ${op.sheetId}${op.name ? ` as "${op.name}"` : ''}`
    case 'set_sheet_hidden':
      return `${op.hidden ? 'Hide' : 'Unhide'} sheet ${op.sheetId}`
    case 'move_sheet':
      return `Move sheet ${op.sheetId} to position ${op.position}`
  }
}
