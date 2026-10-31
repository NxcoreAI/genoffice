# Sorting & layout guide (layout)

## Sorting

`{op:"sort_range", sheetId, range:"A1:C10", keys:[{byColumn:"B", order:"asc"}], hasHeader?:true}`

- `keys`: 1–3 sort keys applied in order — later keys break ties of earlier ones (multi-key sort). The legacy single-key form `{...byColumn:"B", order:"asc"}` is still accepted for one key (ignored when `keys` is present).
- Stable sort by **current values**: numbers < text < booleans; blank cells always sort last regardless of direction; string comparison is number-aware (item2 < item10).
- Each key's byColumn must fall within range; when hasHeader is true the first row does not participate in the sort.
- The sort expands into per-cell changes in the preview (with concurrency protection), so the user can see each cell's before/after values.
- **A range containing formulas is rejected outright** — moving formula text silently changes what relative references point at. Explain the reason to the user (you may suggest converting the formula column to values first); do not try to work around it.
- Sort the whole table together: range must cover all related columns — sorting a single column tears rows apart.

```json
{
  "op": "sort_range",
  "sheetId": "s1",
  "range": "A1:D100",
  "keys": [
    { "byColumn": "B", "order": "desc" },
    { "byColumn": "C", "order": "asc" }
  ],
  "hasHeader": true
}
```

## Merging

- `{op:"merge_cells", sheetId, range:"A1:C1"}` — rejected if it overlaps an existing merged region; needs at least two cells.
- `{op:"unmerge_cells", sheetId, range}` — clears all merges intersecting the range.
- Merging is mainly for title rows/category headers; avoid merging in data regions (it breaks sorting and formula filling).
- For title-like content prefer "merge + center + fill color" over scattered standalone cells.

## Row height & column width

- `{op:"set_row_height", sheetId, row:2, count?:1, heightPoints:24}` — height in points (2–409).
- `{op:"set_col_width", sheetId, column:"B", count?:1, widthPx:120}` — width in pixels (10–2000).
- Give report title rows extra height (e.g. 24–28pt); size data columns so content doesn't overflow, and keep similar columns the same width.

## Hiding rows/columns

- `{op:"set_rows_hidden", sheetId, row:5, count?:1, hidden:true|false}` — row is 1-based; hidden:false unhides.
- `{op:"set_cols_hidden", sheetId, column:"C", count?:1, hidden:true|false}`
- Hiding doesn't delete data and formula references are unaffected; good for hiding helper-calculation columns. To delete data use delete_rows/delete_cols (mind the batching discipline).

## Outline groups (imported files only)

`{op:"outline.group", sheetId, axis:"rows"|"cols", range:"5:12"}` — groups the range's rows (or columns): each contiguous run of the range's outline level shifts +1, clamped to Excel's 0–7 levels. Use axis-local spans: "5:12" for rows, "B:D" for cols (an A1 range like "A5:A12"/"B1:D1" is also accepted — only that axis's extent is used); a span/axis mismatch is rejected.

- `{op:"outline.ungroup", ...}` — shifts the range's level −1 (clamped at 0; runs already ungrouped are left alone).
- `{op:"outline.hideDetail", ...}` — hides the range (the group's detail rows/columns) and marks the summary row/column immediately after the range (end+1) as collapsed.
- `{op:"outline.showDetail", ...}` — shows the range again and clears that collapsed flag.
- Grouping an all-level-7 span does nothing (and ungrouping an all-level-0 span does nothing) — check the preview for a skip before telling the user it worked.
- Exactly the same behavior as the Data → Group / Subtotal outline buttons; levels are recorded in the edit journal and written on save.
- Put the summary row just **below** the detail rows (Excel's summary-below default) — hiding detail then leaves the summary visible. Requires an imported xlsx; unsupported in the demo workbook.

```json
{ "op": "outline.group", "sheetId": "s1", "axis": "rows", "range": "5:12" }
```

## Freeze panes

`{op:"set_freeze", sheetId, rows:1, columns:0}` — freezes the top rows rows and left columns columns; `rows:0, columns:0` unfreezes. Takes effect on the canvas immediately and is written into the file on save. Typical use: freeze the header row with `rows:1`.

## Page setup (printing)

`{op:"set_page_setup", sheetId, orientation?, paperSize?, scale?, fitToWidth?, fitToHeight?, margins?, printGridlines?, printHeadings?, printArea?, printTitles?, header?, footer?, rowBreaks?, colBreaks?}` — at least one property, only works on imported xlsx files, written on save:

- `orientation`: "portrait" | "landscape"
- `paperSize`: OOXML paper code (1=Letter, 8=A3, 9=A4, 11=A5)
- `scale`: print scaling 10–400(%); **mutually exclusive** with fitToWidth/fitToHeight (fit means "fit to N pages wide/tall", 0=automatic)
- `margins`: "normal" | "wide" | "narrow"
- `printGridlines` / `printHeadings`: print gridlines / row-column headings
- `printArea`: "A1:H40" sets the print area, null clears it
- `printTitles`: "1:2" repeats those rows at the top of every printed page (long-table headers), null clears
- `header` / `footer`: `{left?, center?, right?}` printed sections; text carries Excel codes — &P page number, &N total pages, &D date, &A sheet name, && literal &; null clears
- `rowBreaks` / `colBreaks`: arrays of 0-based indexes of the last row/column printed on a page, e.g. `[24, 49]` starts new pages after those; passing [] clears all manual breaks
- Settings are only written into the file (take effect when printing); the canvas does not show pagination.

## Protection (workbook level & editable ranges)

- `{op:"protect_workbook", protected:true|false}` — locks the workbook **structure** (no adding/deleting/moving/renaming/hiding sheets), no password; works only on imported xlsx files and is written on save. A workbook already protected with a password is refused in both directions — ask the user to unlock it in Excel first. This is workbook-level and independent of per-sheet protection (`protect_sheet`, structure guide): it does not stop cell edits.
- `{op:"allow_edit_ranges", sheetId, ranges:[{name?, sqref:"A1:B10"}]}` — **replaces** the sheet's allowed-edit ranges (ranges that stay editable while the sheet is protected); `ranges:[]` removes them all; up to 100 entries; imported xlsx only. Refused when an existing allowed-edit range carries a password. `name` is optional (auto-generated Range1, Range2, …), `sqref` is uppercase A1 form without `$` (a single cell or a range). Only meaningful on a protected sheet — pair it with `protect_sheet`, and lock the rest of the sheet with format_range `locked`.
- Both are journaled and written on save; neither moves addresses (layout-class).

```json
{ "op": "protect_workbook", "protected": true }
```

```json
{ "op": "allow_edit_ranges", "sheetId": "s1", "ranges": [{ "name": "Inputs", "sqref": "B2:C10" }] }
```
