# Known Limitations

This page lists known gaps in SheetKit 1.0 that affect results users can observe. Each entry gives the impact, a workaround where one exists, and the plan status. None of these items is in a current plan, so all are marked **Not scheduled**. Formula results that differ from Excel are listed in [Excel Compatibility](./excel-compatibility.md); only the larger formula engine gaps are repeated here.

## Formula engine

### Whole-column and whole-row references are not evaluated

`SUM(A:A)` and `SUM(1:1)` can be stored in a cell, and insert and remove operations shift them the same way Excel does. The calculator cannot parse them: `evaluate_formula("Sheet1", "SUM(A:A)")` returns a formula parse error, and `calculate_all` stores `#VALUE!` as the cached result.

- **Impact**: the cached value SheetKit stores for these formulas is wrong until another application, such as Excel, recalculates the workbook. The formula text itself is correct.
- **Workaround**: use a bounded range such as `SUM(A1:A1000)` when SheetKit must calculate the value.
- **Status**: Not scheduled.

### Array constants are not evaluated

Formulas containing array constants, such as `SUM({1,2;3,4})`, return a formula parse error from `evaluate_formula`, and `calculate_all` cannot compute them.

- **Impact**: as above. The formula text is stored and saved unchanged.
- **Workaround**: put the constants in cells and reference the range.
- **Status**: Not scheduled.

### Error arguments and several date and text edge cases

Text functions propagate an error given as their text argument, but an error in a numeric argument such as `LEFT("abc",1/0)` surfaces as an evaluation error rather than an error value. `COUNT` and `SUM` recognise numeric text only in SheetKit's plain number syntax, so `" 3 "`, `"50%"`, `"1,000"`, `"(5)"`, and `"12:00"` are not counted where Excel counts them. `EDATE(1/0,1)` and `YEAR(1/0)` return `#VALUE!` instead of `#DIV/0!`. `NETWORKDAYS`, `WEEKDAY`, and `WEEKNUM` still use calendar arithmetic, so serial 0 is rejected and weekdays before March 1900 differ from Excel. The full list, with Excel's results, is in [Excel Compatibility](./excel-compatibility.md).

- **Impact**: cached values differ from Excel for these inputs until the workbook is recalculated in Excel.
- **Workaround**: none in SheetKit. Avoid these inputs where SheetKit's cached value matters.
- **Status**: Not scheduled.

## Structural edits

### Pivot tables are not updated by row and column insertion or removal

`insert_rows`, `remove_row`, `insert_cols`, and `remove_col` do not change pivot table locations or pivot cache source ranges.

- **Impact**: after an edit above or left of a pivot table or its source data, the pivot table points at the old coordinates.
- **Workaround**: delete the pivot table with `delete_pivot_table` and add it again with `add_pivot_table` after the edit, or refresh it in Excel.
- **Status**: Not scheduled.

### Inserting columns inside a table does not add table columns

`insert_cols` inside a table widens the table range but does not add a `tableColumn` entry. For a two-column table on `E1:F3`, `insert_cols("Sheet1", "F", 1)` leaves the range `E1:G3` with two columns. Excel 16.114 reported "We found a problem with some content" when it opened the result and repaired the file.

- **Impact**: the saved file needs repair in Excel.
- **Workaround**: insert columns outside the table, or delete the table with `delete_table` and add it again with the new column list.
- **Status**: Not scheduled.

### Deleting a shared formula's master cell does not rewrite its dependants

In files written by Excel, a formula filled down or across is often stored once in a master cell and shared by the other cells. When `remove_row` or `remove_col` deletes the master cell, the other cells still refer to the shared formula, and no cell defines it any more.

- **Impact**: the dependent cells lose their formula text when read by SheetKit, and Excel may report the file as needing repair.
- **Workaround**: before the edit, set the dependent cells' formulas explicitly with `set_cell_formula`; SheetKit writes them as ordinary formulas.
- **Status**: Not scheduled.

### The streaming reader does not resolve shared formulas

`open_sheet_reader` (`openSheetReader` in Node.js) reads a cell that shares another cell's formula as its cached value only. Other reads, such as `get_cell_value`, `get_rows`, and `get_cols`, return the shifted formula.

- **Impact**: streamed rows show no formula for cells that share a formula filled down or across in Excel.
- **Workaround**: read those cells with `get_cell_value`, `get_rows`, or `get_cols`.
- **Status**: Not scheduled.

### Positions inside a deleted band move instead of becoming `#REF!`

Frozen panes, selections, drawing anchors (charts, images, shapes), sparkline locations, and internal hyperlink targets that point into deleted rows or columns move to the first row or column after the deleted band. Formulas, by contrast, become `#REF!`.

- **Impact**: an internal hyperlink or a sparkline can end up pointing at a different cell than before the edit.
- **Workaround**: check internal hyperlinks and sparklines that point into the deleted area, and reset them after the edit.
- **Status**: Not scheduled.

### Deleting a table leaves its totals row content

`delete_table` removes the table definition but leaves the totals row cells, including `SUBTOTAL(…,[Column])` formulas. Opened in Excel 16.114, such a leftover formula showed as `=SUBTOTAL(102,#REF!)`.

- **Impact**: orphaned formulas show `#REF!` in Excel.
- **Workaround**: clear the totals row cells after `delete_table`.
- **Status**: Not scheduled.

## Tables, slicers, and comments

### Calculated column and custom totals formulas cannot be set

`calculatedColumnFormula` and `totalsRowFormula` in existing tables are preserved when the workbook is saved, but `TableConfig` has no fields for them. Because of that, the `custom` totals function is rejected.

- **Impact**: new tables cannot have calculated columns or custom totals formulas.
- **Workaround**: write the formulas into the cells with `set_cell_formula`. Excel does not treat them as calculated columns.
- **Status**: Not scheduled.

### Slicers connected to pivot tables lose their pivot data on save

SheetKit models slicer caches for table slicers only. When a workbook with a slicer connected to a pivot table is opened and saved, the slicer cache is written without its pivot table list and item data.

- **Impact**: the pivot slicer loses its connection and items after a round trip through SheetKit.
- **Workaround**: none. Avoid saving such workbooks with SheetKit, or recreate the slicer in Excel.
- **Status**: Not scheduled.

### Note and thread operations on the same cell are mutually exclusive

Excel stores each threaded comment with a legacy note that holds a `[Threaded comment]` placeholder and keeps either a note or a thread on a cell, never both. SheetKit enforces the same rule: `add_comment` and `remove_comment` on a cell that holds a thread return `InvalidArgument`, and `add_threaded_comment` on a cell that holds an ordinary note returns `InvalidArgument`. Deleting a thread root also deletes its replies, as Excel does.

- **Impact**: code that used to overwrite a note with a thread, or edit the placeholder note directly, now receives an error instead of silently losing the thread in Excel.
- **Workaround**: remove the note before adding a thread; use `delete_threaded_comment` to clear a threaded cell. `get_comments` does not return placeholder notes, so cells with threads can be found with `get_threaded_comments`.
- **Status**: Resolved by design.

## VBA

### VBA source in East Asian codepages is decoded lossily

`get_vba_modules` decodes module source fully for codepages 1252 and 65001 (UTF-8). For codepages 932 (Japanese), 949 (Korean), and 936 (Simplified Chinese), ASCII and CP932 half-width katakana are kept, and other double-byte characters become U+FFFD. A per-module entry in `VbaProject::warnings` reports how many sequences were replaced. Other codepages fall back to lossy UTF-8 decoding with a warning.

- **Impact**: non-ASCII identifiers, strings, and comments in those modules are not readable. The VBA project bytes are preserved unchanged when the workbook is saved.
- **Workaround**: read the raw project with `get_vba_project` and decode it with a codepage-aware tool.
- **Status**: Not scheduled.

## Opening files

### Compression ratio checks start above a size floor

The per-entry compression ratio limit (default 100:1) applies only to ZIP entries that decompress to more than the floor (default 8 MiB). Smaller entries are not ratio checked, though they still count toward the entry and aggregate size limits. Above the floor, the ratio applies to the whole entry. A non-empty entry with zero compressed bytes is always rejected, and a floor of `0` checks every entry.

- **Impact**: a legitimate file with a large, highly repetitive part (more than 8 MiB and above 100:1) is rejected at the defaults.
- **Workaround**: for trusted files, raise the floor with `OpenOptions::compression_ratio_floor_bytes` (Node: `compressionRatioFloorBytes`), or raise the ratio with `OpenOptions::max_compression_ratio` (Rust only), to a larger finite value.
- **Status**: Behaves as designed. No change is scheduled.

### Resolved: prefixed SpreadsheetML namespace

Worksheets, workbook parts, shared strings, and styles whose root uses a prefixed main namespace (for example `<x:worksheet xmlns:x="…">`) open in eager, lazy, and stream read modes, and they read and save correctly. Parts that omit the relationship namespace declaration also open. This is covered by the namespace compatibility regression tests.

## For contributors

### `Cargo.lock` is not committed

`Cargo.lock` is listed in `.gitignore`, so each checkout resolves dependency versions when it first builds. Two builds of the same revision can therefore use different dependency versions.

- **Impact**: builds of the same revision are not guaranteed to be reproducible, and benchmark comparisons across revisions can mix dependency changes with code changes.
- **Workaround**: keep a `Cargo.lock` from a known-good build and copy it into other checkouts. The revision comparison script does this automatically; see `benchmarks/COMPARISON.md`.
- **Status**: Not scheduled.
