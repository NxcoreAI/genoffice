# Data tools guide (data)

All operations below are **layout-class**: they can share a batch with content/format operations, but not with structural operations (row/column or sheet insertion/deletion) — the exceptions are `remove_duplicates` (deletes rows) and `subtotal` (inserts rows), which are themselves structural (own batch). The feature-setting operations (hyperlinks, AutoFilter, conditional formatting, data validation, notes, defined names) work only on imported real xlsx files; once applied they are recorded in the edit journal and ⌘S saves them into the file.

**Before modifying existing settings, read_sheet_features first**: it reads back the current filter range and column criteria, all conditional formatting rules, data validation rules, defined names, and the shape/image inventory.

## Hyperlinks

- `{op:"set_hyperlink", sheetId, address:"B2", target:"https://example.com"}` — target can also be an in-workbook reference "Sheet1!A1" (# is added automatically); bare domains get https:// prepended automatically.
- `{op:"set_hyperlink", sheetId, address, target:null}` — removes the link (and clears the link styling).
- Setting a link gives the cell blue underlined styling; the cell content is unchanged.

## AutoFilter

- `{op:"set_filter", sheetId, range:"A1:D100"}` — creates an AutoFilter on the range (an existing filter is replaced); the first row of range is the filter header.
- `{op:"clear_filter", sheetId}` — removes the entire AutoFilter.
- `{op:"set_filter_criteria", sheetId, column:"B", values:["East","South"]}` — checkbox-style column criteria: rows whose **displayed cell text** is in values stay visible, other rows are hidden. Requires an existing AutoFilter on the sheet. values:null clears the column's criteria.
- `{op:"set_filter_criteria", sheetId, column:"B", conditions:[{operator:"equal",value:"East"},{operator:"greaterThanOrEqual",value:"100"}], and:true}` — custom conditions instead of checkboxes: 1–2 conditions combined with AND (`and:true`, the default) or OR (`and:false`). operator is one of equal/notEqual/greaterThan/greaterThanOrEqual/lessThan/lessThanOrEqual; value is a string compared against the cell's displayed text (numbers/booleans by their text form).
- Pass **exactly one** of `values` / `conditions`; `values:null` clears the column's criteria. Both forms require an existing AutoFilter (`set_filter` first).
- read_range the column first to see which values actually exist before writing values; text must match verbatim.

```json
{
  "op": "set_filter_criteria",
  "sheetId": "s1",
  "column": "B",
  "conditions": [
    { "operator": "equal", "value": "East" },
    { "operator": "greaterThanOrEqual", "value": "100" }
  ],
  "and": true
}
```

## Remove duplicates (structural)

`{op:"remove_duplicates", sheetId, range:"A1:D100", hasHeader?:true}`

- Keeps the **first occurrence** of each identical row and deletes the later duplicate rows (compared across the range's full width; text case-insensitively, numbers/booleans by value).
- Rows are deleted as whole rows, so the range must cover the **full table width** — anything in columns outside the range goes away with the row. Structural: it cannot share a batch with content/format ops, and after it applies you must re-read the layout before writing more.
- Max 20,000 cells (range area). hasHeader:true keeps the first row (headers) out of the comparison — **always pass it when the range starts with a header row**, otherwise row 1 counts as data and a duplicated header row would itself be deleted.
- **Irreversible after apply** (only the [Undo]/⌘Z rollback restores the rows): warn the user before proposing it and report how many rows were removed.

```json
{ "op": "remove_duplicates", "sheetId": "s1", "range": "A1:D100", "hasHeader": true }
```

## Text to columns

`{op:"text_to_columns", sheetId, source:"A2:A100", delimiter:",", destination?:"A2", collapseConsecutive?:true, clearSource?:true}`

- source must be a **single column** (pieces spill right, e.g. "A2:A100") or a **single row** (pieces spill down, e.g. "A2:Z2"); split one line of cells at a time. Part 1 goes at the source cell itself.
- `delimiter`: the separator string (1–32 chars), e.g. `","`, `";"` or `"\t"`; `null` splits on runs of any whitespace (the text is trimmed first).
- `collapseConsecutive` (default **true**) folds consecutive delimiters into one and drops empty pieces; false splits on every occurrence and writes the empty pieces as cleared cells.
- `destination`: top-left cell of the output block, default the source's top-left cell. In place, the first part overwrites the source cell and unfilled trailing cells up to the widest split are cleared; with a destination elsewhere, `clearSource` (default true) clears the whole source first (the data moves, not duplicates) — clearSource:false keeps the original text.
- Empty source cells are left untouched. Formula cells are split from their computed value — convert_to_values first if that value is unavailable.
- Limits: source ≤20,000 cells and the output block ≤20,000 cells (rows × widest split). Read the source first; make sure the target columns/rows are free.

```json
{
  "op": "text_to_columns",
  "sheetId": "s1",
  "source": "A2:A100",
  "delimiter": ",",
  "destination": "A2",
  "clearSource": true
}
```

## Consolidate

`{op:"consolidate", sheetId, target:"H2", function:"sum"|"count"|"average"|"max"|"min", references:[{sheetId?, range:"B2:D10"}...]}`

- Aggregates 1–20 reference areas **by position** into one output block anchored at `target`: each output cell is a live formula combining that offset from every area, e.g. `=SUM(B2,'Sheet2'!B2)` (sum/count/average/max/min map to SUM/COUNT/AVERAGE/MAX/MIN; same-sheet refs stay bare, cross-sheet refs are sheet-qualified and quoted only when the name needs it).
- All reference ranges must have the **same shape** (rows × columns) — reference #1's shape is the output shape. Max 20,000 output cells.
- `sheetId` on a reference defaults to the target's sheet; `range` uses A1 form (add the sheet via sheetId, not "Sheet1!B2").
- The target block must not overlap any source area (rejected). The output stays live (source edits flow through); freeze it with convert_to_values if the user wants constants.

```json
{
  "op": "consolidate",
  "sheetId": "s1",
  "target": "H2",
  "function": "sum",
  "references": [{ "range": "B2:D10" }, { "sheetId": "s2", "range": "B2:D10" }]
}
```

## Subtotal (structural)

`{op:"subtotal", sheetId, range:"A1:D100", groupColumn:"A", valueColumn:"D", function:"sum"}` — Data → Subtotal:

- The first row of range is the **header**; the rows below are grouped by the stored values of groupColumn. Grouping is by **consecutive runs** — sort by groupColumn first (sort_range) if the same group appears in several blocks.
- Inserts one bold "<group> Total" row after each group with a live `=SUBTOTAL(n,$D$x:$D$y)` in valueColumn, plus a bold "Grand Total" row at the end (its span also covers the nested per-group rows; SUBTOTAL skips other SUBTOTALs, so no double counting). Other columns of the inserted rows stay empty; existing values/formulas in every column are untouched.
- function: "sum" | "average" | "count" | "counta" | "max" | "min" | "product" | "stdev" | "stdevp" | "var" | "varp".
- groupColumn and valueColumn must both fall inside range and must differ. ≤200 groups; imported xlsx only; the range must be **loaded** first (read_range it) — grouping reads the grid's current values, and unloaded rows would be grouped as empty.
- **Structural**: it inserts rows, so it needs its own batch — re-read the layout (get_workbook_context / read_range) before writing anything else.

```json
{
  "op": "subtotal",
  "sheetId": "s1",
  "range": "A1:D100",
  "groupColumn": "A",
  "valueColumn": "D",
  "function": "sum"
}
```

## Conditional formatting

`{op:"add_conditional_format", sheetId, range, rule}` — rule discriminated by kind:

- `{kind:"number", operator:"greaterThan"|"greaterThanOrEqual"|"lessThan"|"lessThanOrEqual"|"equal"|"notEqual"|"between"|"notBetween", value, value2?, format}` — between/notBetween require value2
- `{kind:"text", operator:"contains"|"notContains"|"beginsWith"|"endsWith", text, format}`
- `{kind:"blank", blank:true|false, format}` — true highlights blank cells, false highlights non-blank
- `{kind:"duplicate", unique?, format}` — highlights duplicate values by default; unique:true highlights unique values instead
- `{kind:"top10", rank:10, percent?, bottom?, format}`
- `{kind:"formula", formula:"=B2>100", format}` — formula starts with =, written with relative references anchored at the range's top-left cell
- `{kind:"colorScale", minColor:"#63BE7B", midColor?, maxColor:"#F8696B"}` — color scale (no format field)
- `{kind:"dataBar", color?}` — data bar (no format field)
- `{kind:"iconSet", iconSet:"3TrafficLights1", showValue?, thresholds?}` — icon set (no format field); iconSet one of 3Arrows/3ArrowsGray/3Flags/3TrafficLights1/3TrafficLights2/3Signs/3Symbols/3Symbols2/4Arrows/4ArrowsGray/4RedToBlack/4Rating/4TrafficLights/5Arrows/5ArrowsGray/5Quarters/5Rating; thresholds default to an equal percentile split, pass `{type:"num"|"percent"|"percentile", value}[]` (length = icon count - 1, ascending) to customize

format: `{fillColor?, fontColor?, bold?, italic?}` with at least one property. Highlight colors follow the low-saturation principle (light background, dark text), e.g. red #FFC7CE/#9C0006, green #C6EFCE/#006100, yellow #FFEB9C/#9C6500.

`{op:"clear_conditional_formats", sheetId}` — clears **all** conditional formatting rules on the sheet (including ones that came with the file); confirm with the user before doing this.

## Data validation

`{op:"set_data_validation", sheetId, range, validation}` — validation discriminated by kind:

- `{kind:"list", values:["Yes","No"]}` — dropdown list (≤100 items)
- `{kind:"listRef", range:"H1:H20"}` — dropdown options taken from a cell range's values
- `{kind:"numberBetween", min, max}`
- `{kind:"dateBetween", start:"2026-01-01", end:"2026-12-31"}` — ISO dates
- `{kind:"checkbox"}`
- `{kind:"formula", formula:"=LEN(A1)<=10"}`
- `validation:null` — clears validation rules on the range

## Notes

- `{op:"set_note", sheetId, address:"B2", text:"This figure includes tax"}` — adds/edits a cell note (yellow sticky, shown on hover).
- `{op:"set_note", sheetId, address, text:null}` — deletes the note.
- Saved as Excel legacy comments; visible in Excel when hovering the cell. Read existing notes with read_sheet_features first.

## Pivot tables

- `{op:"add_pivot", sheetId, sourceRange, targetCell, rowFields, columnField?, values:[{field, agg, showDataAs?}]}` — **creates a native pivot table** (real OOXML, interactive and refreshable when opened in Excel); rowFields supports 1–8 levels of row grouping, showDataAs supports percent-of-total/row/column. Usage and limits in the `pivot` guide (load_guide).
- `{op:"refresh_pivot", sheetId}` — **recomputes all pivot tables on the sheet**: using the layout recorded in the file, recomputes the data area from current source data and writes it back (on save it also makes Excel rebuild the cache when the file is opened). Use it after the source data has changed to bring pivot results up to date.
- If refresh finds **new category members** in the source data, the layout grows automatically (new members appended, output region expanded — the growth area must be empty); moved field columns, renamed headers, or pivots using calculated fields/grouping/value filters still fail with an explicit error (suggest the user refresh in Excel in that case).
- read_sheet_features lists each pivot table's row/column/value fields, source range, and whether it is recomputable.
- Modifying an existing pivot table's field layout is not supported yet; the output region is read-only (direct AI writes are rejected).
- When the data is irregular, live formulas are needed for further computation, or add_pivot limits are exceeded, fall back to a **formula aggregation table** (SUMIFS/COUNTIFS, see the pivot guide's fallback plan).
- Summary tables follow the formatting guide's table style rules (headers, alignment, total rows).

## Defined names

- `{op:"add_defined_name", name:"SalesRegion", ref:"Sheet1!$A$1:$B$10"}` — name starts with a letter/underscore; ref is an A1 reference (absolute $ references recommended) or a constant.
- `{op:"delete_defined_name", name}` — deletes the name (formula text referencing it is untouched, but those formulas will turn into #NAME? — confirm no formulas use it before deleting).
- `{op:"create_names_from_selection", sheetId, range:"A1:D20", label:"top"|"left"|"bottom"|"right"}` — creates one defined name per column (`label:"top"`/`"bottom"`) or per row (`label:"left"`/`"right"`) from the label row/column, each pointing at the label's data span with absolute references (e.g. `Sheet1!$B$2:$B$20`). top/bottom need at least two rows, left/right at least two columns.
- Labels are the cells' stored values (a label formula whose computed value is unavailable counts as empty and is skipped); invalid characters become `_`, a leading digit or cell-like label gets a `_` prefix, empty labels are skipped, and duplicate labels within the batch get a numeric suffix (`_2`, `_3`, …).

```json
{ "op": "create_names_from_selection", "sheetId": "s1", "range": "A1:D20", "label": "top" }
```
