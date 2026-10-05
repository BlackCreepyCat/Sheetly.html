# Sheetly

**A complete spreadsheet application in a single HTML file.**
No server, no build step, no dependencies, no network access. Open `Sheetly.html` in a modern browser and start working.

> Current version: **V2.3**

Sheetly offers a familiar Excel-like experience: formulas, multi-sheet workbooks, charts, pivot tables, conditional formatting, data validation, filters, import/export of `.xlsx` and `.csv`, printing, version history and 11 themes, all running locally in your browser.

---

<img width="1599" height="887" alt="image" src="https://github.com/user-attachments/assets/329a40ab-101f-4a00-83e8-6029fb8bd910" />


## Table of contents

- [Features](#features)
- [Getting started](#getting-started)
- [User interface](#user-interface)
- [Formulas and functions](#formulas-and-functions)
- [Editing and navigation](#editing-and-navigation)
- [Data tools](#data-tools)
- [Charts](#charts)
- [Pivot tables](#pivot-tables)
- [Files, import and export](#files-import-and-export)
- [Saving and version history](#saving-and-version-history)
- [Printing](#printing)
- [Themes](#themes)
- [Keyboard shortcuts](#keyboard-shortcuts)
- [Limits](#limits)
- [Browser compatibility](#browser-compatibility)
- [Privacy](#privacy)
- [Technical overview](#technical-overview)
- [Contributing](#contributing)
- [License](#license)

---

## Features

| Area | Highlights |
| --- | --- |
| **Formulas** | 130+ functions (math, statistics, logic, text, date/time, lookup, information, financial), cross-sheet references, whole-column/row references, function library with search |
| **Formatting** | Fonts, sizes, bold/italic/underline/strikethrough, text and fill colors, alignment, text wrapping, borders, merged cells, number formats, format painter |
| **Data tools** | Sorting, AutoFilter, data validation (dropdown lists and input checks), freeze panes, find & replace, hide rows/columns |
| **Conditional formatting** | Rules on values, text, duplicates/unique, top/bottom N, above/below average, custom formulas, color scales, data bars, icon sets |
| **Charts** | Columns, bars, lines, areas, pie, donut, scatter; stacked variants; PNG export |
| **Pivot tables** | Row/column fields, aggregated values (sum, count, average, min, max), refreshable |
| **Collaboration-style tools** | Cell comments, sheet protection with locked/unlocked cells |
| **Auditing** | Trace precedents/dependents, show formulas mode |
| **Files** | Native `.sheetly` format, Excel (`.xlsx`) import/export, CSV export/import |
| **Safety net** | 50-step undo/redo, automatic local saving, named version snapshots |
| **Printing** | Page setup (paper, orientation, margins, scaling), headers, grid lines, PDF through the browser printer |
| **Look & feel** | 11 light and dark themes plus an automatic system mode |

---

<img width="1602" height="888" alt="image" src="https://github.com/user-attachments/assets/e6b91b5b-95cb-4d8b-8759-79ed50373b58" />


## Getting started

1. Download `Sheetly.html`.
2. Open it in a modern browser (Chrome, Edge, Firefox, Safari).
3. That's it.

You can also host it anywhere that serves static files (GitHub Pages, an intranet, a USB stick...). Since it is one self-contained file, it works equally well opened straight from disk.

To explore the application quickly, use **File ▾ → Load the demo example**: it loads a three-sheet sample workbook (undoable with `Ctrl+Z`).

---

## User interface

- **Title bar**: workbook name (editable), **File** menu, **Theme** menu and **Help**.
- **Toolbar**: undo/redo, clipboard, format painter, font, size, text styling, colors, alignment, number format, decimals, borders, merge, insert/delete, sort, AutoSum, find, conditional formatting, data validation, filter, freeze panes, charts, pivot tables, comments, auditing, protection and printing.
- **Formula bar**: name box (type `A1` or `A1:B5` and press Enter to jump/select), `fx` function library button, and the formula/value input.
- **Grid**: the sheet itself, with fill handle, selection, multi-range selection, charts floating on top, and a cell editor with autocomplete.
- **Sheet tabs**: right-click for options, drag to reorder, rename, duplicate, delete (sheet names up to 31 characters).
- **Status bar**: current reference, selection statistics, zoom controls (`-` / `+` or `Ctrl` + mouse wheel), protection indicator and last-saved time.

---

## Formulas and functions

Formulas start with `=`. Sheetly aims at Excel compatibility while being forgiving:

- **Decimal point** is `.` (e.g. `=1.5*2`).
- **Argument separators**: both `,` and `;` are accepted.
- **Function names**: both **English and French** names work (e.g. `SUM` / `SOMME`, `IF` / `SI`, `VLOOKUP` / `RECHERCHEV`).
- **References**: relative (`A1`), absolute (`$A$1`), ranges (`A1:B5`), other sheets (`Sheet2!A1`, `'My data'!A:A`), whole columns (`A:A`) and whole rows (`2:2`).
- **Missing closing parentheses** are added automatically.
- **Errors** follow spreadsheet conventions: `#DIV/0!`, `#NAME?`, `#VALUE!`, `#N/A`, plus `#CIRC!` for circular references.
- **Click-to-reference**: while typing a formula, click a cell to insert its reference, drag for a range, click a header for a whole column/row.

### Function library

Press **`Shift+F3`** or click **`fx`** to open the function library: search by name or description, browse by category, see the syntax and an example, and insert into the cell. Typing in a cell also shows an autocomplete list.

| Category | Functions |
| --- | --- |
| **Math & Trig** | `SUM`, `PRODUCT`, `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `SUMSQ`, `ABS`, `SIGN`, `SQRT`, `POWER`, `EXP`, `LN`, `LOG`, `LOG10`, `MOD`, `QUOTIENT`, `INT`, `TRUNC`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `CEILING`, `FLOOR`, `MROUND`, `EVEN`, `ODD`, `FACT`, `COMBIN`, `GCD`, `LCM`, `PI`, `RAND`, `RANDBETWEEN`, `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `RADIANS`, `DEGREES` |
| **Statistical** | `AVERAGE`, `MEDIAN`, `MODE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`, `LARGE`, `SMALL`, `RANK`, `STDEV`, `STDEVP`, `VAR`, `VARP`, `PERCENTILE`, `QUARTILE`, `SUBTOTAL`, `CORREL`, `SLOPE`, `INTERCEPT`, `RSQ`, `FORECAST` |
| **Logical** | `IF`, `IFS`, `IFERROR`, `IFNA`, `AND`, `OR`, `NOT`, `XOR`, `TRUE`, `FALSE`, `SWITCH`, `CHOOSE` |
| **Text** | `CONCAT`, `CONCATENATE`, `TEXTJOIN`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`, `PROPER`, `TRIM`, `CLEAN`, `SUBSTITUTE`, `REPLACE`, `FIND`, `SEARCH`, `REPT`, `EXACT`, `VALUE`, `NUMBERVALUE`, `TEXT`, `CHAR`, `CODE` |
| **Date & Time** | `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`, `HOUR`, `MINUTE`, `SECOND`, `WEEKDAY`, `DAYS`, `EDATE`, `EOMONTH`, `DATEDIF`, `NETWORKDAYS`, `WORKDAY`, `WEEKNUM`, `ISOWEEKNUM` |
| **Lookup & Reference** | `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `LOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `INDIRECT`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS` |
| **Information** | `ISNUMBER`, `ISTEXT`, `ISNONTEXT`, `ISLOGICAL`, `ISBLANK`, `ISERROR`, `ISERR`, `ISNA`, `ISEVEN`, `ISODD`, `NA`, `N` |
| **Financial** | `NPV`, `IRR`, `NPER`, `RATE`, `PMT`, `FV`, `PV` |

> Note: as in Excel, `VLOOKUP(x, range, 2, FALSE)` — `FALSE` means exact match.

<img width="1597" height="888" alt="image" src="https://github.com/user-attachments/assets/9a302458-3be3-47c5-903b-5aca49215553" />

### Auditing tools

- **Show formulas** (`Ctrl+` `` ` ``) to display formulas instead of values.
- **Trace precedents / dependents** to draw arrows between related cells; **Remove arrows** to clear them.

---

## Editing and navigation

### Entering data

- Type to enter, `F2` or double-click to edit, `Alt+Enter` for a line break inside a cell.
- Dates (`24/12/2025`) and percentages (`15%`) are recognized automatically.
- **Smart fill** with the fill handle (▣): numeric series (1, 2, 3...), days, months, and formulas with adjusted references. `Ctrl+D` fills down, `Ctrl+R` fills right.
- **Dropdown lists** appear automatically on cells with a list validation.

### Selection

- Click and drag in any direction (with auto-scroll on all four edges).
- `Shift+click` extends the selection; `Ctrl+click` adds multiple ranges.
- `Ctrl+A` selects all, `Ctrl+Space` a column, `Shift+Space` a row.

### Clipboard

- Copy / cut / paste with Excel-compatible tab-separated text, so you can exchange data with Excel, LibreOffice or Google Sheets.
- **Format painter** to copy formatting from one range to another.

### Structure

- Insert/delete rows and columns, hide/show rows and columns, resize rows and columns, merge/unmerge cells.
- **Freeze panes**: first row, first column, or at the current selection.
- **Zoom** from the status bar or `Ctrl` + mouse wheel.

### Formatting

- Fonts (Georgia, Arial, Consolas, Courier New, Times New Roman, or default) and sizes 9 to 24.
- Number formats: Automatic, Number, Currency (€), Percentage, Date, with increase/decrease decimals.
- Borders (all, outside, and more), text and fill colors, alignment and text wrapping.

---

<img width="1601" height="884" alt="image" src="https://github.com/user-attachments/assets/8567dec2-f1e2-4d5e-bf64-048ae4dcd92a" />

## Data tools

### Sorting and filtering

- **Sort A to Z / Z to A** from the toolbar or context menu.
- **AutoFilter** with per-column value filters, reapply, and clear criteria.

### Find & Replace

`Ctrl+F` opens the find panel with *Next* and *Replace all*.

### Data validation

Restrict what can be typed into cells:

- **List** (dropdown, from a range or a typed list)
- **Whole number**, **Decimal number**, **Date**, **Text length**
- **Custom** formula

Error styles (stop / warning / information) are supported and exported to Excel.

### Conditional formatting

Create rules via the toolbar button (*New rule*, *Manage rules*, *Clear rules*):

- Cell value, text (contains / does not contain / begins with / ends with)
- Blank, non-blank and error cells
- Duplicate or unique values
- Top / bottom N (or N%)
- Above / below average
- Custom formula
- **Color scales**, **data bars** and **icon sets** (arrows, dots, squares, flags)

### Comments and protection

- **Cell comments** (`Shift+F2`).
- **Sheet protection**: protect a sheet, then lock or unlock individual cells. Protected sheets block edits to locked cells and structural changes. A status-bar indicator shows when protection is on.

---

## Charts

Insert a chart from the toolbar and choose its data range. Supported types:

- Columns, Bars, Lines, Areas (with stacked options for columns, bars and areas)
- Pie and Donut
- Scatter

Charts float over the grid, are linked to their source data, and can be exported as **PNG**. They are included in `.xlsx` exports and recognized when importing `.xlsx` files.

---

## Pivot tables

Build a summary of a data range:

- Choose one or more **row fields** and a **column field**.
- Add **value fields** with an aggregation: **sum, count, average, min, max**.
- Grand totals are computed automatically.
- Use **Refresh pivot tables** after changing source data, or **Edit this pivot table...** to reconfigure.

Access pivot tools from the toolbar or the right-click context menu.

---

## Files, import and export

Use the **File ▾** menu:

| Command | Description |
| --- | --- |
| **New** | Start an empty workbook |
| **Open...** (`Ctrl+O`) | Open `.sheetly`, `.xlsx`, `.xlsm`, `.csv`, `.tsv`, `.txt`, `.json` (and legacy `.registre`) files |
| **Save** (`Ctrl+S`) / **Save as...** (`Ctrl+Shift+S`) | Save the workbook as a `.sheetly` file |
| **Export to Excel (.xlsx)** | Includes values, formulas, styles, merged cells, comments, data validation, conditional formatting and charts |
| **Import a file (.xlsx, .csv...)** | Same as Open |
| **Export to CSV** | UTF-8 CSV (with BOM, for Excel compatibility) |
| **Page setup and printing...** (`Ctrl+P`) | Print dialog |
| **Load the demo example** | Replace the workbook with a 3-sheet demo |
| **Version history...** / **Save a version** | Manage snapshots |

### The `.sheetly` format

A `.sheetly` file is plain **JSON**:

```json
{
  "format": "sheetly",
  "version": 2,
  "name": "My workbook",
  "wb": { "v": 2, "active": 0, "sheets": [ ... ] }
}
```

It stores cells, formulas, styles, merges, column widths and row heights, conditional formatting, validation rules, filters, frozen panes, charts, pivot tables and protection state for every sheet. Files from older Sheetly/Registre versions (including single-sheet files) are still readable.

On browsers that support the File System Access API, **Save** writes back to the same file; otherwise a download is triggered.

### Excel compatibility notes

- `.xlsx` import and export are implemented natively (custom ZIP/OOXML reader and writer), without any external library.
- Importing `.xlsx` needs a browser with `DecompressionStream` support (all current mainstream browsers).
- Features without an equivalent in Sheetly may be simplified on import.

---

## Saving and version history

- **Automatic local save**: the workbook is saved to the browser's `localStorage` shortly after each change, and restored the next time you open the page. The status bar shows the last save time, or a warning if storage is full or blocked.
- **Always keep a real file copy** with **File → Save**: browser storage is convenient but not a backup.
- **Version history**: save named snapshots at any time, restore or delete them later. Up to 25 versions are kept (automatic snapshots are discarded first when the limit is reached).
- **Undo / redo**: up to 50 steps (`Ctrl+Z` / `Ctrl+Y`).

---

## Printing

Open **Page setup and printing...** (`Ctrl+P`):

- Scope: sheet (and more)
- Paper size (A4 by default), orientation, margins
- Scaling: fit to page or custom percentage
- Grid lines, row/column headers, title, date and page numbers

Printing uses the browser's print dialog, so you can also **save as PDF**. Your print preferences are remembered.

---

## Themes

Click **◐ Theme** to choose among 11 themes, or **Auto** to follow the system's light/dark mode:

| Light | Dark |
| --- | --- |
| Sheetly (paper, default) | Night |
| Ocean | Charcoal |
| Forest | Nord |
| Parchment | Twilight |
| Lavender | Solarized Dark |
| Solarized Light | |

The theme choice is remembered between sessions.

---

## Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| `Arrows`, `Tab`, `Enter` | Move (add `Shift` to extend the selection) |
| `Ctrl+Arrows`, `Home`, `End` | Jump to the edge of the data (`+Shift`: extend) |
| `Ctrl+PgDn` / `Ctrl+PgUp` | Next / previous sheet |
| `Ctrl+A` | Select all |
| `Ctrl+Space` / `Shift+Space` | Select column / row |
| `F2` / double-click | Edit cell |
| `Alt+Enter` | Line break in a cell |
| `Ctrl+C` / `Ctrl+X` / `Ctrl+V` | Copy / cut / paste |
| `Ctrl+Z` / `Ctrl+Y` | Undo / redo |
| `Ctrl+D` / `Ctrl+R` | Fill down / fill right |
| `Ctrl+B` / `Ctrl+I` / `Ctrl+U` | Bold / italic / underline |
| `Alt+=` | AutoSum |
| `Shift+F3` | Insert function |
| `Shift+F2` | Cell comment |
| `Ctrl+F` | Find & replace |
| `Ctrl+` `` ` `` | Show formulas |
| `Ctrl+O` / `Ctrl+S` / `Ctrl+Shift+S` | Open / save / save as |
| `Ctrl+P` | Print / page setup |
| `Ctrl` + wheel | Zoom |
| `Esc` | Close dialogs and panels |

A built-in **Help** panel (title bar) provides the same quick reference inside the app.

---

## Limits

| Item | Limit |
| --- | --- |
| Rows per sheet | 5,000 |
| Columns per sheet | 700 |
| Default sheet size | 200 rows x 52 columns (grows automatically as you type) |
| Whole-column / whole-row references | Interpreted against Excel's bounds (1,048,576 rows, 16,384 columns) |
| Sheet name length | 31 characters |
| Undo history | 50 steps |
| Version history | 25 snapshots |

---

## Browser compatibility

Sheetly runs in any current evergreen browser (Chrome, Edge, Firefox, Safari). Some features depend on browser capabilities:

- **File System Access API** (Chrome/Edge): in-place saving to a file. Other browsers fall back to downloads.
- **`DecompressionStream`**: required to import `.xlsx`.
- **`localStorage`**: required for automatic saving, themes, print preferences and version history. If it is blocked (e.g. some private modes), use **File → Save** regularly.

---

## Privacy

Sheetly is fully offline-capable and does **not** send any data anywhere. The file contains no external scripts, fonts, analytics or network calls. Your data stays in your browser (`localStorage`) and in the files you choose to save.

---

## Technical overview

- **Single file**: HTML + CSS + vanilla JavaScript (~6,000 lines, ~400 KB). No framework, no bundler, no third-party library.
- **Theming**: every color goes through CSS variables; the theme is applied before first render to avoid flashes.
- **Formula engine**: tokenizer/parser with a single function catalog that feeds the function library, autocomplete, English/French aliases and argument checking. Dependency tracking with circular-reference detection (`#CIRC!`).
- **Rendering**: HTML table grid with overlay layers (selection, active cell, marching-ants copy border, charts, filters, trace arrows) and SVG charts.
- **Excel I/O**: handwritten OOXML (`.xlsx`) reader/writer with its own ZIP implementation (CRC-32, `deflate-raw` decompression through the browser's native stream API).
- **Persistence**: `localStorage` for autosave, version history, theme and print settings; JSON for `.sheetly` files.

---

## Contributing

Issues and pull requests are welcome. Because the whole application lives in one file, please:

1. Keep it dependency-free and self-contained.
2. Route all colors through the theme CSS variables.
3. Register new spreadsheet functions in the central function catalog so they appear in the function library, autocomplete and help.
4. Test with several themes and with `.xlsx` round-trips (export, then re-import) when touching file I/O.

---

## License

MIT file in the repository with a mention for my work somewhere.
