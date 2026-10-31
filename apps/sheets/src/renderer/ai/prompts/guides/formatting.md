# Formatting guide (formatting)

## Operation definition

`{op:"format_range", sheetId, range:"A1:C1", format:{…}}` — format needs at least one field; a field value of **null clears that property**:

- `bold` / `italic` / `underline` / `strikethrough`: boolean or null
- `fontFamily`: font name (e.g. "Calibri", "微软雅黑") or null
- `fontSize`: point size (1–409) or null
- `fontColor` / `fillColor`: 6-digit hex ("#RRGGBB") or null
- `numberFormat`: Excel format string (e.g. "0.00%", "#,##0.00", "yyyy-mm-dd") or null (restores General)
- `horizontalAlign`: "left" | "center" | "right" or null
- `verticalAlign`: "top" | "center" | "bottom" or null
- `wrapText`: boolean or null (word wrap)
- `textRotation`: integer from -90 to 90 (positive = counterclockwise) or "vertical" (stacked vertical text) or null
- `indent`: indent level 0–250 (0 clears), rendered on screen as left padding and written into the file
- `border`: {type:"all"|"top"|"bottom"|"left"|"right"|"none", color?:"#RRGGBB", style?} — **per-cell edge** semantics (every cell in the range gets that edge); "none" clears all borders. `style` is the OOXML line style: "thin" (default) | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair" | "dashDot" | "dashDotDot" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" | "slantDashDot" (ignored when type is "none").
- `locked` / `hidden`: cell-protection flags, boolean or null — locked blocks editing and hidden removes the formula from the formula bar; **both only take effect while the sheet is protected** (`protect_sheet`, structure guide). null restores the defaults (locked:true, hidden:false). Lock the input cells and then protect the sheet for a "fill in the yellow cells" template.

One operation per range (range-level, up to 200,000 cells — formatting a whole data column of a large file is fine); when you need to "reuse the format from somewhere", read the current state with read_formats first.

## Clear formats

`{op:"clear_formats", sheetId, range:"A1:D20"}` — strips **every visual format** on the range (font, fill, borders, number format, alignment, wrap, rotation, indent) and restores default cell protection (locked, formula visible). Values and formulas are untouched. Up to 20,000 cells. This is the op for "reset the formatting / remove all styling" — do not spell it out as dozens of format_range nulls.

```json
{ "op": "clear_formats", "sheetId": "s1", "range": "A1:D20" }
```

## Document themes (imported files only)

`{op:"set_theme", theme?:"indigo", colors?:"forest", fonts?:"georgia"}` — at least one of the three; requires an imported xlsx whose file has a theme part (demo workbooks and theme-less files are refused):

- `theme`: applies a whole preset (palette + fonts). Ids: office (Office) | ember (Ember) | indigo (Indigo) | forest (Forest) | cream (Cream) | rose (Rose) | graphite (Graphite) | midnight (Midnight).
- `colors`: swaps only the palette. Same ids as `theme`: office | ember | indigo | forest | cream | rose | graphite | midnight.
- `fonts`: swaps only the font scheme. Ids: office (Calibri Light/Calibri) | arial (Arial) | georgia (Georgia) | candara (Candara) | segoe (Segoe UI) | trebuchet (Trebuchet MS, minor Calibri).
- Combine freely: `theme` sets both halves and an explicit `colors`/`fonts` overrides that half of the preset. Unknown ids are rejected with the valid id list.
- Re-theme is **live**: cells styled from theme colors/fonts recolor as soon as the op applies. Charts and conditional-formatting colors pick the new palette on reopen after ⌘S; hard-coded per-cell colors (`fillColor`, `fontColor`) never change.
- A theme whose file has no rewritable font scheme keeps its fonts — only the colors change.

```json
{ "op": "set_theme", "theme": "indigo" }
```

## Number formats

- Percentages: generally `0.0%`/`0.00%`; share columns (summing to 100%) at least `0.000%`
- Large numbers: `#,##0`; keep decimal places consistent within a column
- Dates: `yyyy-mm-dd`

### Currency must be locale-locked (critical, most common mistake)

**Never use bare `$#,##0.00`** — Chinese-locale Excel remaps a bare `$` to `¥`, making a USD table display CNY symbols. Always force the currency symbol with a bracketed locale prefix:

- USD: `[$$-409]#,##0.00` (or integer `[$$-409]#,##0`)
- CNY: `[$¥-804]#,##0.00`
- EUR: `[$€-x-euro2]#,##0.00`
- GBP: `[$£-809]#,##0.00`

At thousand/million scale, decimals are usually 0 (`[$$-409]#,##0`) and the unit is noted in the header (e.g. "Amount (thousands)").

### Negatives and zeros (mandatory for financial tables)

- Negatives in red parentheses: `[$$-409]#,##0;[Red]([$$-409]#,##0)`
- Zero displayed as a dash (distinguishes "0" from "blank = missing data"): `[$$-409]#,##0;[Red]([$$-409]#,##0);"-"`
- Three-section format string semantics: `positive;negative;zero`; plain-number version: `#,##0;[Red](#,##0);"-"`

> The full spec for financial statements/accounting tables (hierarchy, total rows, GAAP order, palette) is in the `financial-formatting` guide.

## Table style rules (when the user has no explicit styling requirements)

- **Headers are mandatory**: no data grid may be a bare matrix of numbers — column headers (dimension/period/metric names) are required, and cross-tabs also need row headers. Headers bold + light fill (e.g. #F2F2F2 or #D6E4F0).
- **Alignment**: numbers right-aligned, text left-aligned, dates centered or left-aligned; strictly consistent within a column and within a table.
- **Restrained font sizes**: one size for body text; section titles at most one step larger; never mix multiple font sizes within one table body.
- **Total/subtotal rows**: highlight with bold or a light fill; don't over-decorate.
- **Color accessibility**: keep high contrast between text and background; avoid harsh high-saturation palettes except to emphasize key values.
- **Financial model conventions** (when the user is building a financial model or asks for professional formatting): hard-coded inputs in blue #0000FF, formula calculations in black, cross-sheet references in green #008000.
- **Styling failures never block**: if a format-class operation is rejected, prioritize delivering correct data; do not retry styling repeatedly.
