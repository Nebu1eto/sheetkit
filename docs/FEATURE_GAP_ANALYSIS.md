# SheetKit Feature Gap Analysis

Comparison of public API methods in code versus documentation coverage across
`docs/guide-en.md`, `docs/getting-started.md`, and `docs/api-reference.md`.

Generated: 2026-02-08

---

## 1. Complete Public API Inventory

### 1.1 Node.js Workbook Class (`#[napi]` methods)

| # | Method | Category |
|---|--------|----------|
| 1 | `new Workbook()` | Workbook I/O |
| 2 | `Workbook.open(path)` | Workbook I/O |
| 3 | `save(path)` | Workbook I/O |
| 4 | `sheetNames` (getter) | Workbook I/O |
| 5 | `getCellValue(sheet, cell)` | Cell Operations |
| 6 | `setCellValue(sheet, cell, value)` | Cell Operations |
| 7 | `newSheet(name)` | Sheet Management |
| 8 | `deleteSheet(name)` | Sheet Management |
| 9 | `setSheetName(old, new)` | Sheet Management |
| 10 | `copySheet(source, target)` | Sheet Management |
| 11 | `getSheetIndex(name)` | Sheet Management |
| 12 | `getActiveSheet()` | Sheet Management |
| 13 | `setActiveSheet(name)` | Sheet Management |
| 14 | `insertRows(sheet, row, count)` | Row Operations |
| 15 | `removeRow(sheet, row)` | Row Operations |
| 16 | `duplicateRow(sheet, row)` | Row Operations |
| 17 | `setRowHeight(sheet, row, height)` | Row Operations |
| 18 | `getRowHeight(sheet, row)` | Row Operations |
| 19 | `setRowVisible(sheet, row, visible)` | Row Operations |
| 20 | `getRowVisible(sheet, row)` | Row Operations |
| 21 | `setRowOutlineLevel(sheet, row, level)` | Row Operations |
| 22 | `getRowOutlineLevel(sheet, row)` | Row Operations |
| 23 | `setRowStyle(sheet, row, styleId)` | Row Operations |
| 24 | `getRowStyle(sheet, row)` | Row Operations |
| 25 | `setColWidth(sheet, col, width)` | Column Operations |
| 26 | `getColWidth(sheet, col)` | Column Operations |
| 27 | `setColVisible(sheet, col, visible)` | Column Operations |
| 28 | `getColVisible(sheet, col)` | Column Operations |
| 29 | `setColOutlineLevel(sheet, col, level)` | Column Operations |
| 30 | `getColOutlineLevel(sheet, col)` | Column Operations |
| 31 | `setColStyle(sheet, col, styleId)` | Column Operations |
| 32 | `getColStyle(sheet, col)` | Column Operations |
| 33 | `insertCols(sheet, col, count)` | Column Operations |
| 34 | `removeCol(sheet, col)` | Column Operations |
| 35 | `addStyle(style)` | Styles |
| 36 | `getCellStyle(sheet, cell)` | Styles |
| 37 | `setCellStyle(sheet, cell, styleId)` | Styles |
| 38 | `mergeCells(sheet, topLeft, bottomRight)` | Merge Cells |
| 39 | `unmergeCell(sheet, reference)` | Merge Cells |
| 40 | `getMergeCells(sheet)` | Merge Cells |
| 41 | `addChart(sheet, from, to, config)` | Charts |
| 42 | `addImage(sheet, config)` | Images |
| 43 | `addDataValidation(sheet, config)` | Data Validation |
| 44 | `getDataValidations(sheet)` | Data Validation |
| 45 | `removeDataValidation(sheet, sqref)` | Data Validation |
| 46 | `setConditionalFormat(sheet, sqref, rules)` | Conditional Formatting |
| 47 | `getConditionalFormats(sheet)` | Conditional Formatting |
| 48 | `deleteConditionalFormat(sheet, sqref)` | Conditional Formatting |
| 49 | `addComment(sheet, config)` | Comments |
| 50 | `getComments(sheet)` | Comments |
| 51 | `removeComment(sheet, cell)` | Comments |
| 52 | `setAutoFilter(sheet, range)` | Auto-Filter |
| 53 | `removeAutoFilter(sheet)` | Auto-Filter |
| 54 | `newStreamWriter(sheetName)` | StreamWriter |
| 55 | `applyStreamWriter(writer)` | StreamWriter |
| 56 | `setDocProps(props)` | Document Properties |
| 57 | `getDocProps()` | Document Properties |
| 58 | `setAppProps(props)` | Document Properties |
| 59 | `getAppProps()` | Document Properties |
| 60 | `setCustomProperty(name, value)` | Document Properties |
| 61 | `getCustomProperty(name)` | Document Properties |
| 62 | `deleteCustomProperty(name)` | Document Properties |
| 63 | `protectWorkbook(config)` | Workbook Protection |
| 64 | `unprotectWorkbook()` | Workbook Protection |
| 65 | `isWorkbookProtected()` | Workbook Protection |
| 66 | `setPanes(sheet, cell)` | Freeze/Split Panes |
| 67 | `unsetPanes(sheet)` | Freeze/Split Panes |
| 68 | `getPanes(sheet)` | Freeze/Split Panes |
| 69 | `setPageMargins(sheet, margins)` | Page Layout |
| 70 | `getPageMargins(sheet)` | Page Layout |
| 71 | `setPageSetup(sheet, setup)` | Page Layout |
| 72 | `getPageSetup(sheet)` | Page Layout |
| 73 | `setHeaderFooter(sheet, header, footer)` | Page Layout |
| 74 | `getHeaderFooter(sheet)` | Page Layout |
| 75 | `setPrintOptions(sheet, opts)` | Page Layout |
| 76 | `getPrintOptions(sheet)` | Page Layout |
| 77 | `insertPageBreak(sheet, row)` | Page Layout |
| 78 | `removePageBreak(sheet, row)` | Page Layout |
| 79 | `getPageBreaks(sheet)` | Page Layout |
| 80 | `setCellHyperlink(sheet, cell, opts)` | Hyperlinks |
| 81 | `getCellHyperlink(sheet, cell)` | Hyperlinks |
| 82 | `deleteCellHyperlink(sheet, cell)` | Hyperlinks |
| 83 | `getRows(sheet)` | Row/Column Iterators |
| 84 | `getCols(sheet)` | Row/Column Iterators |
| 85 | `evaluateFormula(sheet, formula)` | Formula Evaluation |
| 86 | `calculateAll()` | Formula Evaluation |
| 87 | `addPivotTable(config)` | Pivot Tables |
| 88 | `getPivotTables()` | Pivot Tables |
| 89 | `deletePivotTable(name)` | Pivot Tables |

### 1.2 Node.js JsStreamWriter Class (`#[napi]` methods)

| # | Method | Category |
|---|--------|----------|
| 1 | `sheetName` (getter) | StreamWriter |
| 2 | `setColWidth(col, width)` | StreamWriter |
| 3 | `setColWidthRange(min, max, width)` | StreamWriter |
| 4 | `writeRow(row, values)` | StreamWriter |
| 5 | `addMergeCell(reference)` | StreamWriter |

### 1.3 Rust-Only Public API (on Workbook, not in Node.js bindings)

| Method | Category | Notes |
|--------|----------|-------|
| `get_occupied_cells(sheet)` | Cell Operations | Returns Vec<(row, col)> of non-empty cells |
| `get_orientation(sheet)` | Page Layout | Separate getter (JS wraps into `getPageSetup`) |
| `get_paper_size(sheet)` | Page Layout | Separate getter (JS wraps into `getPageSetup`) |
| `get_page_setup_details(sheet)` | Page Layout | Separate getter (JS wraps into `getPageSetup`) |

### 1.4 Rust-Only StreamWriter Methods (not in Node.js bindings)

| Method | Category | Notes |
|--------|----------|-------|
| `set_freeze_panes(cell)` | StreamWriter | Set freeze panes before writing rows |
| `set_col_visible(col, visible)` | StreamWriter | Set column visibility |
| `set_col_outline_level(col, level)` | StreamWriter | Set column outline level (0-7) |
| `set_col_style(col, style_id)` | StreamWriter | Set column style |
| `write_row_with_options(row, values, options)` | StreamWriter | Write row with height/visibility/outline/style |

### 1.5 Rust-Only Low-Level Functions (not on Workbook)

| Function | Module | Notes |
|----------|--------|-------|
| `set_defined_name` | `sheetkit_core::defined_names` | Operates on WorkbookXml directly |
| `get_defined_name` | `sheetkit_core::defined_names` | Operates on WorkbookXml directly |
| `delete_defined_name` | `sheetkit_core::defined_names` | Operates on WorkbookXml directly |
| `protect_sheet` | `sheetkit_core::sheet` | Operates on WorksheetXml directly |
| `unprotect_sheet` | `sheetkit_core::sheet` | Operates on WorksheetXml directly |
| `cell_name_to_coordinates` | `sheetkit_core::utils::cell_ref` | Utility |
| `coordinates_to_cell_name` | `sheetkit_core::utils::cell_ref` | Utility |
| `column_name_to_number` | `sheetkit_core::utils::cell_ref` | Utility |
| `column_number_to_name` | `sheetkit_core::utils::cell_ref` | Utility |
| `date_to_serial` | `sheetkit_core::cell` | Date conversion |
| `datetime_to_serial` | `sheetkit_core::cell` | Date conversion |
| `serial_to_date` | `sheetkit_core::cell` | Date conversion |
| `serial_to_datetime` | `sheetkit_core::cell` | Date conversion |
| `is_date_format_code` | `sheetkit_core::cell` | Date detection |
| `is_date_num_fmt` | `sheetkit_core::cell` | Date detection |

---

## 2. Documentation Coverage Matrix

### Legend

- [x] = Documented with code examples for both Rust and TypeScript
- [~] = Partially documented (mentioned or only one language)
- [ ] = Not documented at all

### 2.1 Coverage in `guide-en.md`

| Feature | Status | Notes |
|---------|--------|-------|
| Workbook I/O (new/open/save/sheetNames) | [x] | |
| Cell Operations (get/set) | [x] | |
| Sheet Management (new/delete/rename/copy/index/active) | [x] | |
| Row Operations (insert/remove/duplicate/height/visible) | [x] | |
| Column Operations (width/visible/insert/remove) | [x] | |
| Row/Column Outline Levels | [ ] | Not mentioned in guide |
| Row/Column Styles | [ ] | Not mentioned in guide |
| Row/Column Iterators (getRows/getCols) | [ ] | Not mentioned in guide |
| Styles (add/get/set cell style) | [x] | |
| Merge Cells | [ ] | Not mentioned in guide |
| Hyperlinks | [ ] | Not mentioned in guide |
| Charts | [x] | Only 8 of 41 chart types listed |
| Images | [x] | |
| Data Validation | [x] | |
| Comments | [x] | |
| Auto-Filter | [x] | |
| Conditional Formatting | [ ] | Not mentioned in guide |
| Freeze/Split Panes | [ ] | Not mentioned in guide |
| Page Layout (margins/setup/print/header/breaks) | [ ] | Not mentioned in guide |
| Document Properties | [x] | |
| Workbook Protection | [x] | |
| Formula Evaluation | [ ] | Not mentioned in guide |
| Pivot Tables | [ ] | Not mentioned in guide |
| StreamWriter | [x] | |
| Utility Functions | [~] | Brief mention at end |

### 2.2 Coverage in `getting-started.md`

| Feature | Status | Notes |
|---------|--------|-------|
| Workbook I/O | [x] | |
| Cell Operations | [x] | Including DateValue |
| CellValue Types | [x] | Good table with all variants |
| Date Values and detection | [x] | |
| Cell References | [x] | |
| Style System (overview + example) | [x] | |
| Charts (overview + example) | [x] | |
| StreamWriter (overview + example) | [x] | |
| All other features | [ ] | Not covered (expected, it's a quick start) |

### 2.3 Coverage in `api-reference.md`

| Feature | Status | Notes |
|---------|--------|-------|
| Workbook I/O | [x] | Complete |
| Cell Operations | [x] | Complete with DateValue details |
| Sheet Management | [x] | Complete |
| Row Operations | [x] | Complete |
| Column Operations | [x] | Complete |
| Row/Column Iterators | [x] | Complete with TS interfaces |
| Styles | [x] | Complete with all sub-types |
| Merge Cells | [x] | Complete |
| Hyperlinks | [x] | Complete |
| Charts | [x] | All 41 types listed |
| Images | [x] | Complete |
| Data Validation | [x] | Complete |
| Comments | [x] | Complete |
| Auto-Filter | [x] | Complete |
| Conditional Formatting | [x] | Complete with 18 rule types |
| Freeze/Split Panes | [x] | Complete |
| Page Layout | [x] | Complete (margins, setup, print, header/footer, breaks) |
| Defined Names | [~] | Documented as Rust-only, low-level API |
| Document Properties | [x] | Complete |
| Workbook Protection | [x] | Complete |
| Sheet Protection | [~] | Documented as Rust-only, low-level API |
| Formula Evaluation | [x] | Complete with 110 function list |
| Pivot Tables | [x] | Complete |
| StreamWriter | [x] | Complete, including Rust-only methods |
| Utility Functions | [x] | Complete |
| `get_occupied_cells` | [ ] | Rust-only method, not documented anywhere |
| `is_date_format_code` / `is_date_num_fmt` | [ ] | Re-exported in facade, not documented |

---

## 3. Gap Summary

### 3.1 Features in Code but MISSING from `guide-en.md`

These features have full Node.js bindings and are documented in `api-reference.md`,
but are absent from the user guide. A user reading only the guide would not discover them.

1. **Merge Cells** (`mergeCells`, `unmergeCell`, `getMergeCells`)
   - Three methods fully implemented in Node.js bindings.
   - Fully documented in api-reference.md Section 8.
   - Entirely absent from guide-en.md.

2. **Hyperlinks** (`setCellHyperlink`, `getCellHyperlink`, `deleteCellHyperlink`)
   - Three methods with support for external, internal, and email link types.
   - Fully documented in api-reference.md Section 9.
   - Entirely absent from guide-en.md.

3. **Conditional Formatting** (`setConditionalFormat`, `getConditionalFormats`, `deleteConditionalFormat`)
   - 18 rule types (cellIs, colorScale, dataBar, duplicates, top10, etc.).
   - Fully documented in api-reference.md Section 15.
   - Entirely absent from guide-en.md.

4. **Freeze/Split Panes** (`setPanes`, `unsetPanes`, `getPanes`)
   - Three methods for freezing rows/columns.
   - Fully documented in api-reference.md Section 16.
   - Entirely absent from guide-en.md.

5. **Page Layout** (margins, setup, print options, header/footer, page breaks -- 11 methods total)
   - `setPageMargins` / `getPageMargins`
   - `setPageSetup` / `getPageSetup`
   - `setPrintOptions` / `getPrintOptions`
   - `setHeaderFooter` / `getHeaderFooter`
   - `insertPageBreak` / `removePageBreak` / `getPageBreaks`
   - Fully documented in api-reference.md Section 17.
   - Entirely absent from guide-en.md.

6. **Row/Column Iterators** (`getRows`, `getCols`)
   - Bulk data reading with typed cell results.
   - Fully documented in api-reference.md Section 6.
   - Entirely absent from guide-en.md.

7. **Row/Column Outline Levels** (`setRowOutlineLevel`, `getRowOutlineLevel`, `setColOutlineLevel`, `getColOutlineLevel`)
   - Grouping/outlining support (levels 0-7).
   - Documented in api-reference.md Sections 4 and 5.
   - Absent from guide-en.md (the guide covers row/column basics but not outline levels).

8. **Row/Column Styles** (`setRowStyle`, `getRowStyle`, `setColStyle`, `getColStyle`)
   - Apply styles to entire rows/columns.
   - Documented in api-reference.md Sections 4 and 5.
   - Absent from guide-en.md.

9. **Formula Evaluation** (`evaluateFormula`, `calculateAll`)
   - 110 supported Excel functions.
   - Fully documented in api-reference.md Section 22.
   - Entirely absent from guide-en.md.

10. **Pivot Tables** (`addPivotTable`, `getPivotTables`, `deletePivotTable`)
    - Full pivot table creation with row/column/data fields and 11 aggregate functions.
    - Fully documented in api-reference.md Section 23.
    - Entirely absent from guide-en.md.

### 3.2 Features in Code but NOT Documented Anywhere

1. **`get_occupied_cells(sheet)`** (Rust only)
   - Returns a list of (row, col) tuples for all non-empty cells.
   - Available as a public method on the Workbook struct.
   - Not documented in any doc file, not exposed in Node.js bindings.

2. **`is_date_format_code` / `is_date_num_fmt`** (Rust only)
   - Re-exported in the `sheetkit` facade crate.
   - Useful for determining if a number format represents a date.
   - Not mentioned in any documentation.

### 3.3 Chart Types: guide-en.md vs Code

The guide lists only 8 chart types; the code and api-reference support 41 chart types.
The following 33 chart types are implemented in code but missing from the guide:

- ColStacked, ColPercentStacked, Col3D, Col3DStacked, Col3DPercentStacked
- BarStacked, BarPercentStacked, Bar3D, Bar3DStacked, Bar3DPercentStacked
- LineStacked, LinePercentStacked, Line3D
- Pie3D, Doughnut
- Area, AreaStacked, AreaPercentStacked, Area3D, Area3DStacked, Area3DPercentStacked
- Scatter, ScatterSmooth, ScatterLine
- Radar, RadarFilled, RadarMarker
- StockHLC, StockOHLC, StockVHLC, StockVOHLC
- Bubble
- Surface, Surface3D, SurfaceWireframe, SurfaceWireframe3D
- ColLine, ColLineStacked, ColLinePercentStacked

### 3.4 Node.js Bindings Gaps (Features in Rust but not in Node.js)

The following Rust Workbook methods have no Node.js binding:

| Method | Difficulty to Add | Impact |
|--------|-------------------|--------|
| `get_occupied_cells` | Low | Useful for sheet inspection |

The following Rust StreamWriter methods have no Node.js binding:

| Method | Difficulty to Add | Impact |
|--------|-------------------|--------|
| `set_freeze_panes` | Low | Common use case for streamed sheets |
| `set_col_visible` | Low | Column hiding in streams |
| `set_col_outline_level` | Low | Column grouping in streams |
| `set_col_style` | Low | Column styling in streams |
| `write_row_with_options` | Medium | Row options (height, visibility, outline, style) |

The following Rust-only features operate on low-level XML types rather than the
Workbook facade. They cannot be exposed via Node.js without first adding Workbook-level wrappers:

| Feature | Rust Module | Impact |
|---------|-------------|--------|
| Defined Names (set/get/delete) | `sheetkit_core::defined_names` | Named ranges -- common Excel feature |
| Sheet Protection (protect/unprotect) | `sheetkit_core::sheet` | Individual sheet lock-down |

---

## 4. Recommendations

### 4.1 High Priority: Update `guide-en.md`

The guide is missing 10 major feature categories that all have working Node.js
bindings. These features are documented in api-reference.md but many users will
read the guide first and not discover them. Add sections for:

1. **Merge Cells** -- very commonly used, minimal documentation needed
2. **Hyperlinks** -- commonly used, needs examples for all 3 link types
3. **Conditional Formatting** -- complex feature, needs at least 3-4 examples (cellIs, colorScale, dataBar, text-based)
4. **Freeze/Split Panes** -- commonly used, simple API, one example suffices
5. **Page Layout** -- 11 methods; group into sub-sections (margins, setup, print, header/footer, page breaks)
6. **Row/Column Iterators** -- essential for reading data, needs TypeScript interface explanation
7. **Formula Evaluation** -- major feature with 110 functions, needs usage examples and function list
8. **Pivot Tables** -- advanced feature, needs a complete example
9. **Row/Column Outline Levels and Styles** -- add to existing Row/Column section
10. **Chart Types** -- update the chart types table from 8 to all 41

### 4.2 Medium Priority: Add Node.js Bindings for Missing StreamWriter Features

Five Rust StreamWriter methods lack Node.js bindings. These are useful for
streaming large files:

- `setFreezePanes(cell)` -- very common need
- `setColVisible(col, visible)`
- `setColOutlineLevel(col, level)`
- `setColStyle(col, styleId)`
- `writeRowWithOptions(row, values, options)`

### 4.3 Medium Priority: Elevate Low-Level APIs to Workbook Facade

Two features (Defined Names, Sheet Protection) exist only as low-level functions
operating on XML types. Adding Workbook-level wrappers would make them accessible
from both Rust and Node.js:

- `wb.set_defined_name(name, value, scope)` / `wb.setDefinedName(...)`
- `wb.protect_sheet(sheet, config)` / `wb.protectSheet(...)`

### 4.4 Low Priority: Document Undocumented Rust-Only Functions

- Add `get_occupied_cells` to the Utility Functions section of api-reference.md
- Add `is_date_format_code` / `is_date_num_fmt` to the Date section of api-reference.md

### 4.5 Low Priority: Sync `guide-ko.md` and `api-reference-ko.md`

The Korean-language docs likely have the same gaps as their English counterparts.
After updating the English docs, apply the same changes to the Korean versions.

---

## 5. Summary Statistics

| Metric | Count |
|--------|-------|
| Total Node.js Workbook methods | 89 |
| Total Node.js StreamWriter methods | 5 |
| Total Rust-only Workbook methods | 4 |
| Total Rust-only StreamWriter methods | 5 |
| Total Rust-only low-level functions | 15 |
| Features documented in api-reference.md | 25 sections |
| Features documented in guide-en.md | 12 sections |
| Features missing from guide-en.md | 10 categories (37 methods) |
| Features not documented anywhere | 4 methods/functions |
| Chart types in code | 41 |
| Chart types listed in guide | 8 |
