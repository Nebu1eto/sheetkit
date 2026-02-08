# SheetKit API Reference

SheetKit is a Rust library for reading and writing Excel (.xlsx) files, with Node.js bindings via napi-rs. This document covers every public API method available in both the Rust crate and the TypeScript/Node.js package.

---

## Table of Contents

- [1. Workbook I/O](#1-workbook-io)
- [2. Cell Operations](#2-cell-operations)
- [3. Sheet Management](#3-sheet-management)
- [4. Row Operations](#4-row-operations)
- [5. Column Operations](#5-column-operations)
- [6. Row/Column Iterators](#6-rowcolumn-iterators)
- [7. Styles](#7-styles)
- [8. Merge Cells](#8-merge-cells)
- [9. Hyperlinks](#9-hyperlinks)
- [10. Charts](#10-charts)
- [11. Images](#11-images)
- [12. Data Validation](#12-data-validation)
- [13. Comments](#13-comments)
- [14. Auto-Filter](#14-auto-filter)
- [15. Conditional Formatting](#15-conditional-formatting)
- [16. Freeze/Split Panes](#16-freezesplit-panes)
- [17. Page Layout](#17-page-layout)
- [18. Defined Names](#18-defined-names)
- [19. Document Properties](#19-document-properties)
- [20. Workbook Protection](#20-workbook-protection)
- [21. Sheet Protection](#21-sheet-protection)
- [22. Formula Evaluation](#22-formula-evaluation)
- [23. Pivot Tables](#23-pivot-tables)
- [24. StreamWriter](#24-streamwriter)
- [25. Utility Functions](#25-utility-functions)
- [26. Sparklines](#26-sparklines)
- [27. Theme Colors](#27-theme-colors)
- [28. Rich Text](#28-rich-text)

---

## 1. Workbook I/O

The `Workbook` is the central type. It represents an in-memory `.xlsx` file and provides all operations for reading and writing spreadsheet data.

### `Workbook::new()` / `new Workbook()`

Create a new empty workbook containing a single sheet named "Sheet1".

**Rust:**

```rust
use sheetkit::Workbook;

let wb = Workbook::new();
```

**TypeScript:**

```typescript
import { Workbook } from "sheetkit";

const wb = new Workbook();
```

### `Workbook::open(path)` / `Workbook.open(path)`

Open an existing `.xlsx` file from disk. Returns an error if the file cannot be read or is not a valid `.xlsx` archive.

**Rust:**

```rust
let wb = Workbook::open("report.xlsx")?;
```

**TypeScript:**

```typescript
const wb = Workbook.open("report.xlsx");
```

### `wb.save(path)` / `wb.save(path)`

Save the workbook to a `.xlsx` file on disk. Overwrites the file if it already exists.

**Rust:**

```rust
wb.save("output.xlsx")?;
```

**TypeScript:**

```typescript
wb.save("output.xlsx");
```

### `wb.sheet_names()` / `wb.sheetNames`

Return the names of all sheets in workbook order.

**Rust:**

```rust
let names: Vec<&str> = wb.sheet_names();
```

**TypeScript:**

```typescript
const names: string[] = wb.sheetNames;
```

> Note: In TypeScript, `sheetNames` is a getter property, not a method.

---

## 2. Cell Operations

### `get_cell_value` / `getCellValue`

Read the typed value of a single cell.

**Rust:**

```rust
use sheetkit::{Workbook, CellValue};

let wb = Workbook::open("data.xlsx")?;
let value: CellValue = wb.get_cell_value("Sheet1", "B3")?;

match value {
    CellValue::String(s) => println!("String: {s}"),
    CellValue::Number(n) => println!("Number: {n}"),
    CellValue::Bool(b) => println!("Bool: {b}"),
    CellValue::Date(serial) => println!("Date serial: {serial}"),
    CellValue::Formula { expr, result } => println!("Formula: ={expr}"),
    CellValue::Error(e) => println!("Error: {e}"),
    CellValue::RichString(runs) => println!("Rich text with {} runs", runs.len()),
    CellValue::Empty => println!("Empty"),
}
```

**TypeScript:**

```typescript
const value = wb.getCellValue("Sheet1", "B3");
// value is: string | number | boolean | DateValue | null
```

### `set_cell_value` / `setCellValue`

Write a value to a cell. Accepts strings, numbers, booleans, date values, or empty/null to clear.

**Rust:**

```rust
wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))?;
wb.set_cell_value("Sheet1", "A2", CellValue::Number(42.0))?;
wb.set_cell_value("Sheet1", "A3", CellValue::Bool(true))?;
wb.set_cell_value("Sheet1", "A4", CellValue::Empty)?;

// Date from chrono types
use chrono::NaiveDate;
let date = NaiveDate::from_ymd_opt(2025, 1, 15).unwrap();
wb.set_cell_value("Sheet1", "A5", CellValue::from(date))?;
```

**TypeScript:**

```typescript
wb.setCellValue("Sheet1", "A1", "Hello");
wb.setCellValue("Sheet1", "A2", 42);
wb.setCellValue("Sheet1", "A3", true);
wb.setCellValue("Sheet1", "A4", null); // clear cell

// Date value
wb.setCellValue("Sheet1", "A5", { type: "date", serial: 45672 });
```

### CellValue Type Mapping

| Rust Variant | TypeScript Type | Description |
|---|---|---|
| `CellValue::String(String)` | `string` | Text value |
| `CellValue::Number(f64)` | `number` | Numeric value (integers stored as f64) |
| `CellValue::Bool(bool)` | `boolean` | Boolean value |
| `CellValue::Date(f64)` | `DateValue` | Date as Excel serial number |
| `CellValue::Formula { expr, result }` | `string` | Formula expression (returned as string in TS) |
| `CellValue::Error(String)` | `string` | Error value (e.g., "#DIV/0!") |
| `CellValue::RichString(Vec<RichTextRun>)` | `string` | Rich text (returned as concatenated plain text in TS) |
| `CellValue::Empty` | `null` | Empty cell |

### DateValue (TypeScript)

The `DateValue` object is used for date cells in TypeScript:

```typescript
interface DateValue {
  type: "date";   // always "date"
  serial: number;  // Excel serial number (days since 1899-12-30)
  iso?: string;    // ISO 8601 string (e.g., "2025-01-15" or "2025-01-15T14:30:00")
}
```

When reading a date cell, `iso` is populated automatically. When writing, only `serial` is required.

### Date Detection

Cells with built-in date number formats (IDs 14-22 and 45-47) are automatically read as `CellValue::Date` in Rust, or `DateValue` in TypeScript. Custom number formats containing date/time tokens (y, m, d, h, s) are also detected.

### `get_occupied_cells(sheet)` (Rust only)

Return a list of `(col, row)` coordinate pairs for every non-empty cell in a sheet. Both values are 1-based. Useful for iterating only over cells that contain data without scanning the entire grid.

```rust
let cells = wb.get_occupied_cells("Sheet1")?;
for (col, row) in &cells {
    println!("Cell at col {}, row {}", col, row);
}
```

---

## 3. Sheet Management

### `new_sheet(name)` / `newSheet(name)`

Create a new empty sheet. Returns the 0-based sheet index.

**Rust:**

```rust
let index: usize = wb.new_sheet("Sales")?;
```

**TypeScript:**

```typescript
const index: number = wb.newSheet("Sales");
```

### `delete_sheet(name)` / `deleteSheet(name)`

Delete a sheet by name. Returns an error if the sheet does not exist or if it is the last remaining sheet (a workbook must always have at least one sheet).

**Rust:**

```rust
wb.delete_sheet("Sheet2")?;
```

**TypeScript:**

```typescript
wb.deleteSheet("Sheet2");
```

### `set_sheet_name(old, new)` / `setSheetName(old, new)`

Rename a sheet. Returns an error if the old name does not exist or the new name is invalid or already taken.

**Rust:**

```rust
wb.set_sheet_name("Sheet1", "Summary")?;
```

**TypeScript:**

```typescript
wb.setSheetName("Sheet1", "Summary");
```

### `copy_sheet(src, dst)` / `copySheet(src, dst)`

Copy a sheet. Creates a new sheet named `dst` with the same content as `src`. Returns the 0-based index of the new sheet.

**Rust:**

```rust
let index: usize = wb.copy_sheet("Sheet1", "Sheet1_Copy")?;
```

**TypeScript:**

```typescript
const index: number = wb.copySheet("Sheet1", "Sheet1_Copy");
```

### `get_sheet_index(name)` / `getSheetIndex(name)`

Get the 0-based index of a sheet by name, or `None`/`null` if not found.

**Rust:**

```rust
let idx: Option<usize> = wb.get_sheet_index("Sales");
```

**TypeScript:**

```typescript
const idx: number | null = wb.getSheetIndex("Sales");
```

### `get_active_sheet()` / `getActiveSheet()`

Get the name of the currently active sheet.

**Rust:**

```rust
let name: &str = wb.get_active_sheet();
```

**TypeScript:**

```typescript
const name: string = wb.getActiveSheet();
```

### `set_active_sheet(name)` / `setActiveSheet(name)`

Set the active sheet by name. Returns an error if the sheet does not exist.

**Rust:**

```rust
wb.set_active_sheet("Sales")?;
```

**TypeScript:**

```typescript
wb.setActiveSheet("Sales");
```

### Sheet Name Rules

Sheet names must:
- Be non-empty
- Be at most 31 characters
- Not contain `: \ / ? * [ ]`
- Not start or end with a single quote (`'`)

---

## 4. Row Operations

All row numbers are 1-based.

### `insert_rows(sheet, row, count)` / `insertRows(sheet, row, count)`

Insert `count` empty rows starting at `row`, shifting existing rows at and below that position downward. Cell references in shifted rows are updated automatically.

**Rust:**

```rust
wb.insert_rows("Sheet1", 3, 2)?; // insert 2 rows at row 3
```

**TypeScript:**

```typescript
wb.insertRows("Sheet1", 3, 2);
```

### `remove_row(sheet, row)` / `removeRow(sheet, row)`

Delete a single row and shift rows below it upward by one.

**Rust:**

```rust
wb.remove_row("Sheet1", 5)?;
```

**TypeScript:**

```typescript
wb.removeRow("Sheet1", 5);
```

### `duplicate_row(sheet, row)` / `duplicateRow(sheet, row)`

Copy a row and insert the duplicate directly below the source row. Existing rows below are shifted down.

**Rust:**

```rust
wb.duplicate_row("Sheet1", 2)?;
```

**TypeScript:**

```typescript
wb.duplicateRow("Sheet1", 2);
```

### `set_row_height` / `get_row_height`

Set or get the height of a row in points. Returns `None`/`null` when no explicit height is set.

**Rust:**

```rust
wb.set_row_height("Sheet1", 1, 30.0)?;
let height: Option<f64> = wb.get_row_height("Sheet1", 1)?;
```

**TypeScript:**

```typescript
wb.setRowHeight("Sheet1", 1, 30.0);
const height: number | null = wb.getRowHeight("Sheet1", 1);
```

### `set_row_visible` / `get_row_visible`

Set or check the visibility of a row. Rows are visible by default.

**Rust:**

```rust
wb.set_row_visible("Sheet1", 3, false)?; // hide row 3
let visible: bool = wb.get_row_visible("Sheet1", 3)?;
```

**TypeScript:**

```typescript
wb.setRowVisible("Sheet1", 3, false);
const visible: boolean = wb.getRowVisible("Sheet1", 3);
```

### `set_row_outline_level` / `get_row_outline_level`

Set or get the outline (grouping) level of a row. Valid range: 0-7. Returns 0 if not set.

**Rust:**

```rust
wb.set_row_outline_level("Sheet1", 5, 1)?;
let level: u8 = wb.get_row_outline_level("Sheet1", 5)?;
```

**TypeScript:**

```typescript
wb.setRowOutlineLevel("Sheet1", 5, 1);
const level: number = wb.getRowOutlineLevel("Sheet1", 5);
```

### `set_row_style` / `get_row_style`

Apply a style ID to an entire row, or retrieve the current row style ID (0 if not set).

**Rust:**

```rust
let style_id = wb.add_style(&style)?;
wb.set_row_style("Sheet1", 1, style_id)?;
let current: u32 = wb.get_row_style("Sheet1", 1)?;
```

**TypeScript:**

```typescript
const styleId = wb.addStyle(style);
wb.setRowStyle("Sheet1", 1, styleId);
const current: number = wb.getRowStyle("Sheet1", 1);
```

---

## 5. Column Operations

Columns are identified by letter names (e.g., "A", "B", "AA").

### `set_col_width` / `get_col_width`

Set or get the width of a column. Valid range: 0.0 to 255.0. Returns `None`/`null` when no explicit width is set.

**Rust:**

```rust
wb.set_col_width("Sheet1", "B", 20.0)?;
let width: Option<f64> = wb.get_col_width("Sheet1", "B")?;
```

**TypeScript:**

```typescript
wb.setColWidth("Sheet1", "B", 20.0);
const width: number | null = wb.getColWidth("Sheet1", "B");
```

### `set_col_visible` / `get_col_visible`

Set or check the visibility of a column. Columns are visible by default.

**Rust:**

```rust
wb.set_col_visible("Sheet1", "C", false)?;
let visible: bool = wb.get_col_visible("Sheet1", "C")?;
```

**TypeScript:**

```typescript
wb.setColVisible("Sheet1", "C", false);
const visible: boolean = wb.getColVisible("Sheet1", "C");
```

### `insert_cols(sheet, col, count)` / `insertCols(sheet, col, count)`

Insert `count` empty columns starting at the given column letter, shifting existing columns to the right. Cell references are updated automatically.

**Rust:**

```rust
wb.insert_cols("Sheet1", "B", 3)?; // insert 3 columns at B
```

**TypeScript:**

```typescript
wb.insertCols("Sheet1", "B", 3);
```

### `remove_col(sheet, col)` / `removeCol(sheet, col)`

Delete a single column and shift columns to its right leftward by one.

**Rust:**

```rust
wb.remove_col("Sheet1", "D")?;
```

**TypeScript:**

```typescript
wb.removeCol("Sheet1", "D");
```

### `set_col_outline_level` / `get_col_outline_level`

Set or get the outline (grouping) level of a column. Valid range: 0-7.

**Rust:**

```rust
wb.set_col_outline_level("Sheet1", "B", 2)?;
let level: u8 = wb.get_col_outline_level("Sheet1", "B")?;
```

**TypeScript:**

```typescript
wb.setColOutlineLevel("Sheet1", "B", 2);
const level: number = wb.getColOutlineLevel("Sheet1", "B");
```

### `set_col_style` / `get_col_style`

Apply a style ID to an entire column, or retrieve the current column style ID (0 if not set).

**Rust:**

```rust
wb.set_col_style("Sheet1", "A", style_id)?;
let current: u32 = wb.get_col_style("Sheet1", "A")?;
```

**TypeScript:**

```typescript
wb.setColStyle("Sheet1", "A", styleId);
const current: number = wb.getColStyle("Sheet1", "A");
```

---

## 6. Row/Column Iterators

### `get_rows(sheet)` / `getRows(sheet)`

Return all rows that contain at least one cell. Sparse: empty rows are omitted.

**Rust:**

```rust
let rows = wb.get_rows("Sheet1")?;
// Vec<(row_number: u32, cells: Vec<(column_name: String, CellValue)>)>
for (row_num, cells) in &rows {
    for (col_name, value) in cells {
        println!("Row {row_num}, Col {col_name}: {value}");
    }
}
```

**TypeScript:**

```typescript
const rows = wb.getRows("Sheet1");
// JsRowData[]
for (const row of rows) {
    console.log(`Row ${row.row}:`);
    for (const cell of row.cells) {
        console.log(`  ${cell.column}: ${cell.value} (${cell.valueType})`);
    }
}
```

### `get_cols(sheet)` / `getCols(sheet)`

Return all columns that contain data. The result is a column-oriented transpose of the row data.

**Rust:**

```rust
let cols = wb.get_cols("Sheet1")?;
// Vec<(column_name: String, cells: Vec<(row_number: u32, CellValue)>)>
```

**TypeScript:**

```typescript
const cols = wb.getCols("Sheet1");
// JsColData[]
```

### TypeScript Cell Types

**JsRowData:**

```typescript
interface JsRowData {
    row: number;         // 1-based row number
    cells: JsRowCell[];
}

interface JsRowCell {
    column: string;         // column name (e.g., "A", "B")
    valueType: string;      // "string" | "number" | "boolean" | "date" | "empty" | "error" | "formula"
    value?: string;         // string representation
    numberValue?: number;   // set when valueType is "number"
    boolValue?: boolean;    // set when valueType is "boolean"
}
```

**JsColData:**

```typescript
interface JsColData {
    column: string;         // column name
    cells: JsColCell[];
}

interface JsColCell {
    row: number;            // 1-based row number
    valueType: string;      // same types as JsRowCell
    value?: string;
    numberValue?: number;
    boolValue?: boolean;
}
```

---

## 7. Styles

Styles control the visual formatting of cells. A style is registered once with `add_style`, which returns a numeric style ID. That ID is then applied to cells, rows, or columns.

### `add_style(style)` / `addStyle(style)`

Register a style definition and return its style ID. Identical styles are deduplicated: registering the same style twice returns the same ID.

**Rust:**

```rust
use sheetkit::style::*;

let style = Style {
    font: Some(FontStyle {
        name: Some("Arial".to_string()),
        size: Some(12.0),
        bold: true,
        color: Some(StyleColor::Rgb("#FF0000".to_string())),
        ..Default::default()
    }),
    fill: Some(FillStyle {
        pattern: PatternType::Solid,
        fg_color: Some(StyleColor::Rgb("#FFFF00".to_string())),
        bg_color: None,
    }),
    num_fmt: Some(NumFmtStyle::Custom("#,##0.00".to_string())),
    ..Default::default()
};
let style_id: u32 = wb.add_style(&style)?;
```

**TypeScript:**

```typescript
const styleId = wb.addStyle({
    font: {
        name: "Arial",
        size: 12,
        bold: true,
        color: "#FF0000",
    },
    fill: {
        pattern: "solid",
        fgColor: "#FFFF00",
    },
    customNumFmt: "#,##0.00",
});
```

### `set_cell_style` / `get_cell_style`

Apply a style ID to a single cell, or get the style ID currently applied to a cell. Returns `None`/`null` for the default style.

**Rust:**

```rust
wb.set_cell_style("Sheet1", "A1", style_id)?;
let current: Option<u32> = wb.get_cell_style("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.setCellStyle("Sheet1", "A1", styleId);
const current: number | null = wb.getCellStyle("Sheet1", "A1");
```

### Style Components Reference

#### FontStyle

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `name` | `Option<String>` | `string?` | Font family (e.g., "Calibri") |
| `size` | `Option<f64>` | `number?` | Font size in points |
| `bold` | `bool` | `boolean?` | Bold text |
| `italic` | `bool` | `boolean?` | Italic text |
| `underline` | `bool` | `boolean?` | Underline text |
| `strikethrough` | `bool` | `boolean?` | Strikethrough text |
| `color` | `Option<StyleColor>` | `string?` | Font color |

**StyleColor (Rust):** `StyleColor::Rgb("#FF0000".into())`, `StyleColor::Theme(1)`, `StyleColor::Indexed(8)`

**Color strings (TypeScript):** `"#FF0000"` (RGB hex), `"theme:1"` (theme color), `"indexed:8"` (indexed color)

#### FillStyle

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `pattern` | `PatternType` | `string?` | Fill pattern type |
| `fg_color` | `Option<StyleColor>` | `string?` | Foreground color |
| `bg_color` | `Option<StyleColor>` | `string?` | Background color |

**PatternType values:**

| Rust | TypeScript | Description |
|---|---|---|
| `PatternType::None` | `"none"` | No fill |
| `PatternType::Solid` | `"solid"` | Solid fill |
| `PatternType::Gray125` | `"gray125"` | 12.5% gray |
| `PatternType::DarkGray` | `"darkGray"` | Dark gray |
| `PatternType::MediumGray` | `"mediumGray"` | Medium gray |
| `PatternType::LightGray` | `"lightGray"` | Light gray |

#### BorderStyle

Each side (`left`, `right`, `top`, `bottom`, `diagonal`) accepts a `BorderSideStyle`:

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `style` | `BorderLineStyle` | `string?` | Line style |
| `color` | `Option<StyleColor>` | `string?` | Border color |

**BorderLineStyle values:**

| Rust | TypeScript |
|---|---|
| `BorderLineStyle::Thin` | `"thin"` |
| `BorderLineStyle::Medium` | `"medium"` |
| `BorderLineStyle::Thick` | `"thick"` |
| `BorderLineStyle::Dashed` | `"dashed"` |
| `BorderLineStyle::Dotted` | `"dotted"` |
| `BorderLineStyle::Double` | `"double"` |
| `BorderLineStyle::Hair` | `"hair"` |
| `BorderLineStyle::MediumDashed` | `"mediumDashed"` |
| `BorderLineStyle::DashDot` | `"dashDot"` |
| `BorderLineStyle::MediumDashDot` | `"mediumDashDot"` |
| `BorderLineStyle::DashDotDot` | `"dashDotDot"` |
| `BorderLineStyle::MediumDashDotDot` | `"mediumDashDotDot"` |
| `BorderLineStyle::SlantDashDot` | `"slantDashDot"` |

**Rust example:**

```rust
use sheetkit::style::*;

let style = Style {
    border: Some(BorderStyle {
        top: Some(BorderSideStyle {
            style: BorderLineStyle::Thin,
            color: Some(StyleColor::Rgb("#000000".to_string())),
        }),
        bottom: Some(BorderSideStyle {
            style: BorderLineStyle::Double,
            color: Some(StyleColor::Rgb("#0000FF".to_string())),
        }),
        ..Default::default()
    }),
    ..Default::default()
};
```

**TypeScript example:**

```typescript
const styleId = wb.addStyle({
    border: {
        top: { style: "thin", color: "#000000" },
        bottom: { style: "double", color: "#0000FF" },
    },
});
```

#### AlignmentStyle

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `horizontal` | `Option<HorizontalAlign>` | `string?` | Horizontal alignment |
| `vertical` | `Option<VerticalAlign>` | `string?` | Vertical alignment |
| `wrap_text` | `bool` | `boolean?` | Enable text wrapping |
| `text_rotation` | `Option<u32>` | `number?` | Rotation angle in degrees |
| `indent` | `Option<u32>` | `number?` | Indentation level |
| `shrink_to_fit` | `bool` | `boolean?` | Shrink text to fit cell width |

**HorizontalAlign values:** `General`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterContinuous`, `Distributed`
In TypeScript: `"general"`, `"left"`, `"center"`, `"right"`, `"fill"`, `"justify"`, `"centerContinuous"`, `"distributed"`

**VerticalAlign values:** `Top`, `Center`, `Bottom`, `Justify`, `Distributed`
In TypeScript: `"top"`, `"center"`, `"bottom"`, `"justify"`, `"distributed"`

#### NumFmtStyle

Number formats control how values are displayed.

**Rust:**

```rust
use sheetkit::style::NumFmtStyle;

// Built-in format by ID
NumFmtStyle::Builtin(9)  // 0%

// Custom format string
NumFmtStyle::Custom("#,##0.00".to_string())
```

**TypeScript:**

Use `numFmtId` for built-in formats or `customNumFmt` for custom format strings on the style object:

```typescript
// Built-in format
wb.addStyle({ numFmtId: 9 }); // 0%

// Custom format
wb.addStyle({ customNumFmt: "#,##0.00" });
```

**Common built-in format IDs:**

| ID | Format | Description |
|---|---|---|
| 0 | General | General |
| 1 | 0 | Integer |
| 2 | 0.00 | 2 decimal places |
| 3 | #,##0 | Thousands separator |
| 4 | #,##0.00 | Thousands with 2 decimals |
| 9 | 0% | Percentage |
| 10 | 0.00% | Percentage with 2 decimals |
| 11 | 0.00E+00 | Scientific notation |
| 14 | m/d/yyyy | Date |
| 15 | d-mmm-yy | Date |
| 20 | h:mm | Time |
| 21 | h:mm:ss | Time |
| 22 | m/d/yyyy h:mm | Date and time |
| 49 | @ | Text |

#### ProtectionStyle

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `locked` | `bool` | `boolean?` | Lock the cell (default: true) |
| `hidden` | `bool` | `boolean?` | Hide formulas in protected sheet view |

---

## 8. Merge Cells

Merge cells combines a rectangular range of cells into a single visual cell.

### `merge_cells(sheet, top_left, bottom_right)` / `mergeCells(sheet, topLeft, bottomRight)`

Merge a range of cells defined by its top-left and bottom-right corners.

**Rust:**

```rust
wb.merge_cells("Sheet1", "A1", "C3")?;
```

**TypeScript:**

```typescript
wb.mergeCells("Sheet1", "A1", "C3");
```

### `unmerge_cell(sheet, reference)` / `unmergeCell(sheet, reference)`

Remove a merged cell range. The `reference` must match the exact range that was merged (e.g., "A1:C3").

**Rust:**

```rust
wb.unmerge_cell("Sheet1", "A1:C3")?;
```

**TypeScript:**

```typescript
wb.unmergeCell("Sheet1", "A1:C3");
```

### `get_merge_cells(sheet)` / `getMergeCells(sheet)`

Get all merged cell ranges on a sheet. Returns a list of range strings.

**Rust:**

```rust
let ranges: Vec<String> = wb.get_merge_cells("Sheet1")?;
// e.g., ["A1:C3", "E5:F6"]
```

**TypeScript:**

```typescript
const ranges: string[] = wb.getMergeCells("Sheet1");
```

---

## 9. Hyperlinks

Hyperlinks can link cells to external URLs, internal sheet references, or email addresses.

### `set_cell_hyperlink` / `setCellHyperlink`

Set a hyperlink on a cell.

**Rust:**

```rust
use sheetkit::hyperlink::HyperlinkType;

// External URL
wb.set_cell_hyperlink("Sheet1", "A1", HyperlinkType::External("https://example.com".into()), Some("Example"), Some("Click to visit"))?;

// Internal sheet reference
wb.set_cell_hyperlink("Sheet1", "B1", HyperlinkType::Internal("Sheet2!A1".into()), None, None)?;

// Email
wb.set_cell_hyperlink("Sheet1", "C1", HyperlinkType::Email("mailto:user@example.com".into()), None, None)?;
```

**TypeScript:**

```typescript
// External URL
wb.setCellHyperlink("Sheet1", "A1", {
    linkType: "external",
    target: "https://example.com",
    display: "Example",
    tooltip: "Click to visit",
});

// Internal sheet reference
wb.setCellHyperlink("Sheet1", "B1", {
    linkType: "internal",
    target: "Sheet2!A1",
});

// Email
wb.setCellHyperlink("Sheet1", "C1", {
    linkType: "email",
    target: "mailto:user@example.com",
});
```

### `get_cell_hyperlink` / `getCellHyperlink`

Get hyperlink information for a cell, or `None`/`null` if no hyperlink exists.

**Rust:**

```rust
if let Some(info) = wb.get_cell_hyperlink("Sheet1", "A1")? {
    // info.link_type, info.display, info.tooltip
}
```

**TypeScript:**

```typescript
const info = wb.getCellHyperlink("Sheet1", "A1");
if (info) {
    // info.linkType, info.target, info.display, info.tooltip
}
```

### `delete_cell_hyperlink` / `deleteCellHyperlink`

Remove a hyperlink from a cell.

**Rust:**

```rust
wb.delete_cell_hyperlink("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.deleteCellHyperlink("Sheet1", "A1");
```

### HyperlinkOptions (TypeScript)

```typescript
interface JsHyperlinkOptions {
    linkType: string;    // "external" | "internal" | "email"
    target: string;      // URL, sheet reference, or mailto address
    display?: string;    // optional display text
    tooltip?: string;    // optional tooltip text
}
```

> Note: External and email hyperlinks are stored in the worksheet `.rels` file with `TargetMode="External"`. Internal hyperlinks use only a `location` attribute.

---

## 10. Charts

Charts render data from cell ranges and are anchored between two cells (top-left and bottom-right corners of the chart area).

### `add_chart` / `addChart`

Add a chart to a sheet.

**Rust:**

```rust
use sheetkit::{ChartConfig, ChartSeries, ChartType};

let config = ChartConfig {
    chart_type: ChartType::Col,
    title: Some("Quarterly Sales".to_string()),
    series: vec![ChartSeries {
        name: "Revenue".to_string(),
        categories: "Sheet1!$A$2:$A$5".to_string(),
        values: "Sheet1!$B$2:$B$5".to_string(),
        x_values: None,
        bubble_sizes: None,
    }],
    show_legend: true,
    view_3d: None,
};
wb.add_chart("Sheet1", "D1", "K15", &config)?;
```

**TypeScript:**

```typescript
wb.addChart("Sheet1", "D1", "K15", {
    chartType: "col",
    title: "Quarterly Sales",
    series: [{
        name: "Revenue",
        categories: "Sheet1!$A$2:$A$5",
        values: "Sheet1!$B$2:$B$5",
    }],
    showLegend: true,
});
```

### ChartConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `chart_type` | `ChartType` | `string` | Chart type (see table below) |
| `title` | `Option<String>` | `string?` | Chart title |
| `series` | `Vec<ChartSeries>` | `JsChartSeries[]` | Data series |
| `show_legend` | `bool` | `boolean?` | Show legend (default: true) |
| `view_3d` | `Option<View3DConfig>` | `JsView3DConfig?` | 3D rotation settings |

### ChartSeries

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `name` | `String` | `string` | Series name |
| `categories` | `String` | `string` | Category axis range (e.g., "Sheet1!$A$2:$A$5") |
| `values` | `String` | `string` | Value axis range (e.g., "Sheet1!$B$2:$B$5") |
| `x_values` | `Option<String>` | `string?` | X-axis values (scatter/bubble charts only) |
| `bubble_sizes` | `Option<String>` | `string?` | Bubble sizes (bubble charts only) |

### Supported Chart Types (41 types)

**Column charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Col` | `"col"` | Clustered column |
| `ChartType::ColStacked` | `"colStacked"` | Stacked column |
| `ChartType::ColPercentStacked` | `"colPercentStacked"` | 100% stacked column |
| `ChartType::Col3D` | `"col3D"` | 3D clustered column |
| `ChartType::Col3DStacked` | `"col3DStacked"` | 3D stacked column |
| `ChartType::Col3DPercentStacked` | `"col3DPercentStacked"` | 3D 100% stacked column |

**Bar charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Bar` | `"bar"` | Clustered bar |
| `ChartType::BarStacked` | `"barStacked"` | Stacked bar |
| `ChartType::BarPercentStacked` | `"barPercentStacked"` | 100% stacked bar |
| `ChartType::Bar3D` | `"bar3D"` | 3D clustered bar |
| `ChartType::Bar3DStacked` | `"bar3DStacked"` | 3D stacked bar |
| `ChartType::Bar3DPercentStacked` | `"bar3DPercentStacked"` | 3D 100% stacked bar |

**Line charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Line` | `"line"` | Line |
| `ChartType::LineStacked` | `"lineStacked"` | Stacked line |
| `ChartType::LinePercentStacked` | `"linePercentStacked"` | 100% stacked line |
| `ChartType::Line3D` | `"line3D"` | 3D line |

**Pie charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Pie` | `"pie"` | Pie |
| `ChartType::Pie3D` | `"pie3D"` | 3D pie |

**Area charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Area` | `"area"` | Area |
| `ChartType::AreaStacked` | `"areaStacked"` | Stacked area |
| `ChartType::AreaPercentStacked` | `"areaPercentStacked"` | 100% stacked area |
| `ChartType::Area3D` | `"area3D"` | 3D area |
| `ChartType::Area3DStacked` | `"area3DStacked"` | 3D stacked area |
| `ChartType::Area3DPercentStacked` | `"area3DPercentStacked"` | 3D 100% stacked area |

**Scatter charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Scatter` | `"scatter"` | Scatter (markers only) |
| `ChartType::ScatterSmooth` | `"scatterSmooth"` | Scatter with smooth lines |
| `ChartType::ScatterLine` | `"scatterLine"` | Scatter with straight lines |

**Radar charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Radar` | `"radar"` | Radar |
| `ChartType::RadarFilled` | `"radarFilled"` | Filled radar |
| `ChartType::RadarMarker` | `"radarMarker"` | Radar with markers |

**Stock charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::StockHLC` | `"stockHLC"` | High-Low-Close |
| `ChartType::StockOHLC` | `"stockOHLC"` | Open-High-Low-Close |
| `ChartType::StockVHLC` | `"stockVHLC"` | Volume-High-Low-Close |
| `ChartType::StockVOHLC` | `"stockVOHLC"` | Volume-Open-High-Low-Close |

**Surface charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Surface` | `"surface"` | 3D surface |
| `ChartType::Surface3D` | `"surface3D"` | 3D surface (top view) |
| `ChartType::SurfaceWireframe` | `"surfaceWireframe"` | Wireframe surface |
| `ChartType::SurfaceWireframe3D` | `"surfaceWireframe3D"` | Wireframe surface (top view) |

**Other charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Doughnut` | `"doughnut"` | Doughnut |
| `ChartType::Bubble` | `"bubble"` | Bubble |

**Combo charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::ColLine` | `"colLine"` | Column + line combo |
| `ChartType::ColLineStacked` | `"colLineStacked"` | Stacked column + line |
| `ChartType::ColLinePercentStacked` | `"colLinePercentStacked"` | 100% stacked column + line |

### View3DConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `rot_x` | `Option<i32>` | `number?` | X-axis rotation angle |
| `rot_y` | `Option<i32>` | `number?` | Y-axis rotation angle |
| `depth_percent` | `Option<u32>` | `number?` | Depth as percentage of chart width |
| `right_angle_axes` | `Option<bool>` | `boolean?` | Use right-angle axes |
| `perspective` | `Option<u32>` | `number?` | Perspective field of view |

---

## 11. Images

Embed images into worksheets. Supported formats: PNG, JPEG, and GIF.

### `add_image` / `addImage`

Add an image to a sheet at the specified cell position.

**Rust:**

```rust
use sheetkit::image::{ImageConfig, ImageFormat};

let image_data = std::fs::read("logo.png")?;
let config = ImageConfig {
    data: image_data,
    format: ImageFormat::Png,
    from_cell: "B2".to_string(),
    width_px: 200,
    height_px: 100,
};
wb.add_image("Sheet1", &config)?;
```

**TypeScript:**

```typescript
import { readFileSync } from "fs";

const imageData = readFileSync("logo.png");
wb.addImage("Sheet1", {
    data: imageData,
    format: "png",   // "png" | "jpeg" | "gif"
    fromCell: "B2",
    widthPx: 200,
    heightPx: 100,
});
```

### ImageConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `data` | `Vec<u8>` | `Buffer` | Raw image bytes |
| `format` | `ImageFormat` | `string` | `Png`/`"png"`, `Jpeg`/`"jpeg"`, `Gif`/`"gif"` |
| `from_cell` | `String` | `string` | Anchor cell (top-left corner) |
| `width_px` | `u32` | `number` | Image width in pixels |
| `height_px` | `u32` | `number` | Image height in pixels |

---

## 12. Data Validation

Data validation restricts what values users can enter in cells.

### `add_data_validation` / `addDataValidation`

Add a validation rule to a cell range.

**Rust:**

```rust
use sheetkit::validation::*;

// Dropdown list
let config = DataValidationConfig {
    sqref: "B2:B100".to_string(),
    validation_type: ValidationType::List,
    operator: None,
    formula1: Some("\"Red,Green,Blue\"".to_string()),
    formula2: None,
    allow_blank: true,
    error_style: Some(ErrorStyle::Stop),
    error_title: Some("Invalid color".to_string()),
    error_message: Some("Please select from the list.".to_string()),
    prompt_title: Some("Color".to_string()),
    prompt_message: Some("Choose a color.".to_string()),
    show_input_message: true,
    show_error_message: true,
};
wb.add_data_validation("Sheet1", &config)?;
```

**TypeScript:**

```typescript
wb.addDataValidation("Sheet1", {
    sqref: "B2:B100",
    validationType: "list",
    formula1: '"Red,Green,Blue"',
    allowBlank: true,
    errorStyle: "stop",
    errorTitle: "Invalid color",
    errorMessage: "Please select from the list.",
    promptTitle: "Color",
    promptMessage: "Choose a color.",
    showInputMessage: true,
    showErrorMessage: true,
});
```

### `get_data_validations` / `getDataValidations`

Get all data validation rules on a sheet.

**Rust:**

```rust
let validations = wb.get_data_validations("Sheet1")?;
```

**TypeScript:**

```typescript
const validations = wb.getDataValidations("Sheet1");
```

### `remove_data_validation` / `removeDataValidation`

Remove a data validation rule by its cell range (sqref).

**Rust:**

```rust
wb.remove_data_validation("Sheet1", "B2:B100")?;
```

**TypeScript:**

```typescript
wb.removeDataValidation("Sheet1", "B2:B100");
```

### Validation Types

| Rust | TypeScript | Description |
|---|---|---|
| `ValidationType::Whole` | `"whole"` | Whole number |
| `ValidationType::Decimal` | `"decimal"` | Decimal number |
| `ValidationType::List` | `"list"` | Dropdown list |
| `ValidationType::Date` | `"date"` | Date value |
| `ValidationType::Time` | `"time"` | Time value |
| `ValidationType::TextLength` | `"textlength"` | Text length constraint |
| `ValidationType::Custom` | `"custom"` | Custom formula |

### Validation Operators

Used with `Whole`, `Decimal`, `Date`, `Time`, and `TextLength` types:

| Rust | TypeScript |
|---|---|
| `ValidationOperator::Between` | `"between"` |
| `ValidationOperator::NotBetween` | `"notbetween"` |
| `ValidationOperator::Equal` | `"equal"` |
| `ValidationOperator::NotEqual` | `"notequal"` |
| `ValidationOperator::LessThan` | `"lessthan"` |
| `ValidationOperator::LessThanOrEqual` | `"lessthanorequal"` |
| `ValidationOperator::GreaterThan` | `"greaterthan"` |
| `ValidationOperator::GreaterThanOrEqual` | `"greaterthanorequal"` |

### Error Styles

| Rust | TypeScript | Description |
|---|---|---|
| `ErrorStyle::Stop` | `"stop"` | Reject invalid input |
| `ErrorStyle::Warning` | `"warning"` | Warn but allow |
| `ErrorStyle::Information` | `"information"` | Inform only |

---

## 13. Comments

Comments (also known as notes) attach text annotations to individual cells.

### `add_comment` / `addComment`

Add a comment to a cell.

**Rust:**

```rust
use sheetkit::comment::CommentConfig;

wb.add_comment("Sheet1", &CommentConfig {
    cell: "A1".to_string(),
    author: "Alice".to_string(),
    text: "Please verify this value.".to_string(),
})?;
```

**TypeScript:**

```typescript
wb.addComment("Sheet1", {
    cell: "A1",
    author: "Alice",
    text: "Please verify this value.",
});
```

### `get_comments` / `getComments`

Get all comments on a sheet.

**Rust:**

```rust
let comments = wb.get_comments("Sheet1")?;
for c in &comments {
    println!("{}: {} (by {})", c.cell, c.text, c.author);
}
```

**TypeScript:**

```typescript
const comments = wb.getComments("Sheet1");
```

### `remove_comment` / `removeComment`

Remove a comment from a specific cell.

**Rust:**

```rust
wb.remove_comment("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.removeComment("Sheet1", "A1");
```

---

## 14. Auto-Filter

Auto-filter adds dropdown filter controls to a header row.

### `set_auto_filter` / `setAutoFilter`

Set an auto-filter on a cell range. The first row of the range becomes the filter header.

**Rust:**

```rust
wb.set_auto_filter("Sheet1", "A1:D100")?;
```

**TypeScript:**

```typescript
wb.setAutoFilter("Sheet1", "A1:D100");
```

### `remove_auto_filter` / `removeAutoFilter`

Remove the auto-filter from a sheet.

**Rust:**

```rust
wb.remove_auto_filter("Sheet1")?;
```

**TypeScript:**

```typescript
wb.removeAutoFilter("Sheet1");
```

---

## 15. Conditional Formatting

Conditional formatting changes the appearance of cells based on rules applied to their values.

### `set_conditional_format` / `setConditionalFormat`

Apply one or more conditional formatting rules to a cell range.

**Rust:**

```rust
use sheetkit::conditional::*;
use sheetkit::style::*;

// Highlight cells greater than 100 in red
let rules = vec![ConditionalFormatRule {
    rule_type: ConditionalFormatType::CellIs {
        operator: CfOperator::GreaterThan,
        formula: "100".to_string(),
        formula2: None,
    },
    format: Some(ConditionalStyle {
        font: Some(FontStyle {
            color: Some(StyleColor::Rgb("#FF0000".to_string())),
            ..Default::default()
        }),
        fill: Some(FillStyle {
            pattern: PatternType::Solid,
            fg_color: Some(StyleColor::Rgb("#FFCCCC".to_string())),
            bg_color: None,
        }),
        border: None,
        num_fmt: None,
    }),
    priority: Some(1),
    stop_if_true: false,
}];
wb.set_conditional_format("Sheet1", "A1:A100", &rules)?;
```

**TypeScript:**

```typescript
wb.setConditionalFormat("Sheet1", "A1:A100", [{
    ruleType: "cellIs",
    operator: "greaterThan",
    formula: "100",
    format: {
        font: { color: "#FF0000" },
        fill: { pattern: "solid", fgColor: "#FFCCCC" },
    },
    priority: 1,
}]);
```

### `get_conditional_formats` / `getConditionalFormats`

Get all conditional formatting rules for a sheet.

**Rust:**

```rust
let formats = wb.get_conditional_formats("Sheet1")?;
// Vec<(sqref: String, rules: Vec<ConditionalFormatRule>)>
```

**TypeScript:**

```typescript
const formats = wb.getConditionalFormats("Sheet1");
// JsConditionalFormatEntry[]
// { sqref: string, rules: JsConditionalFormatRule[] }[]
```

### `delete_conditional_format` / `deleteConditionalFormat`

Remove conditional formatting for a specific cell range.

**Rust:**

```rust
wb.delete_conditional_format("Sheet1", "A1:A100")?;
```

**TypeScript:**

```typescript
wb.deleteConditionalFormat("Sheet1", "A1:A100");
```

### Rule Types (18 types)

| Rule Type | Description | Key Fields |
|---|---|---|
| `cellIs` | Compare cell value against formula(s) | `operator`, `formula`, `formula2` |
| `expression` | Custom formula evaluates to true/false | `formula` |
| `colorScale` | Gradient color scale (2 or 3 colors) | `min/mid/max_type`, `min/mid/max_value`, `min/mid/max_color` |
| `dataBar` | Data bar proportional to value | `min/max_type`, `min/max_value`, `bar_color`, `show_value` |
| `duplicateValues` | Highlight duplicate values | (no extra fields) |
| `uniqueValues` | Highlight unique values | (no extra fields) |
| `top10` | Top N values | `rank`, `percent` |
| `bottom10` | Bottom N values | `rank`, `percent` |
| `aboveAverage` | Above/below average values | `above`, `equal_average` |
| `containsBlanks` | Cells that are blank | (no extra fields) |
| `notContainsBlanks` | Cells that are not blank | (no extra fields) |
| `containsErrors` | Cells containing errors | (no extra fields) |
| `notContainsErrors` | Cells without errors | (no extra fields) |
| `containsText` | Cells containing specific text | `text` |
| `notContainsText` | Cells not containing text | `text` |
| `beginsWith` | Cells beginning with text | `text` |
| `endsWith` | Cells ending with text | `text` |

### CfOperator Values

Used with `cellIs` rules: `lessThan`, `lessThanOrEqual`, `equal`, `notEqual`, `greaterThanOrEqual`, `greaterThan`, `between`, `notBetween`.

### CfValueType Values

Used with `colorScale` and `dataBar` for min/mid/max: `num`, `percent`, `min`, `max`, `percentile`, `formula`.

### Examples

**Color scale (green to red):**

```typescript
wb.setConditionalFormat("Sheet1", "B2:B50", [{
    ruleType: "colorScale",
    minType: "min",
    minColor: "#63BE7B",
    maxType: "max",
    maxColor: "#F8696B",
}]);
```

**Data bar:**

```typescript
wb.setConditionalFormat("Sheet1", "C2:C50", [{
    ruleType: "dataBar",
    barColor: "#638EC6",
    showValue: true,
}]);
```

**Contains text:**

```typescript
wb.setConditionalFormat("Sheet1", "A1:A100", [{
    ruleType: "containsText",
    text: "urgent",
    format: {
        font: { bold: true, color: "#FF0000" },
    },
}]);
```

---

## 16. Freeze/Split Panes

Freeze panes lock rows and/or columns so they remain visible while scrolling.

### `set_panes(sheet, cell)` / `setPanes(sheet, cell)`

Freeze rows and columns. The `cell` argument identifies the top-left cell of the scrollable area:
- `"A2"` freezes row 1
- `"B1"` freezes column A
- `"B2"` freezes row 1 and column A
- `"C3"` freezes rows 1-2 and columns A-B

**Rust:**

```rust
wb.set_panes("Sheet1", "B2")?; // freeze row 1 + column A
```

**TypeScript:**

```typescript
wb.setPanes("Sheet1", "B2");
```

### `unset_panes(sheet)` / `unsetPanes(sheet)`

Remove any freeze or split panes from a sheet.

**Rust:**

```rust
wb.unset_panes("Sheet1")?;
```

**TypeScript:**

```typescript
wb.unsetPanes("Sheet1");
```

### `get_panes(sheet)` / `getPanes(sheet)`

Get the current freeze pane cell reference, or `None`/`null` if no panes are set.

**Rust:**

```rust
let pane: Option<String> = wb.get_panes("Sheet1")?;
```

**TypeScript:**

```typescript
const pane: string | null = wb.getPanes("Sheet1");
```

---

## 17. Page Layout

Page layout settings control how a sheet appears when printed.

### Margins

#### `set_page_margins` / `setPageMargins`

Set page margins in inches.

**Rust:**

```rust
use sheetkit::page_layout::PageMarginsConfig;

wb.set_page_margins("Sheet1", &PageMarginsConfig {
    left: 0.7,
    right: 0.7,
    top: 0.75,
    bottom: 0.75,
    header: 0.3,
    footer: 0.3,
})?;
```

**TypeScript:**

```typescript
wb.setPageMargins("Sheet1", {
    left: 0.7,
    right: 0.7,
    top: 0.75,
    bottom: 0.75,
    header: 0.3,
    footer: 0.3,
});
```

#### `get_page_margins` / `getPageMargins`

Get page margins for a sheet. Returns default values if not explicitly set.

**Rust:**

```rust
let margins = wb.get_page_margins("Sheet1")?;
```

**TypeScript:**

```typescript
const margins = wb.getPageMargins("Sheet1");
```

### Page Setup

#### `set_page_setup` / `setPageSetup`

Set paper size, orientation, scale, and fit-to-page options.

**Rust:**

```rust
use sheetkit::page_layout::{Orientation, PaperSize};

wb.set_page_setup("Sheet1", Some(Orientation::Landscape), Some(PaperSize::A4), Some(100), None, None)?;
```

**TypeScript:**

```typescript
wb.setPageSetup("Sheet1", {
    paperSize: "a4",       // "letter" | "tabloid" | "legal" | "a3" | "a4" | "a5" | "b4" | "b5"
    orientation: "landscape",  // "portrait" | "landscape"
    scale: 100,            // 10-400
    fitToWidth: 1,         // number of pages wide
    fitToHeight: 1,        // number of pages tall
});
```

#### `get_page_setup` / `getPageSetup`

Get the current page setup for a sheet.

**TypeScript:**

```typescript
const setup = wb.getPageSetup("Sheet1");
// { paperSize?: string, orientation?: string, scale?: number, fitToWidth?: number, fitToHeight?: number }
```

### Print Options

#### `set_print_options` / `setPrintOptions`

Set print options: gridlines, headings, and centering.

**Rust:**

```rust
wb.set_print_options("Sheet1", Some(true), Some(false), Some(true), None)?;
```

**TypeScript:**

```typescript
wb.setPrintOptions("Sheet1", {
    gridLines: true,
    headings: false,
    horizontalCentered: true,
    verticalCentered: false,
});
```

#### `get_print_options` / `getPrintOptions`

Get print options for a sheet.

**TypeScript:**

```typescript
const opts = wb.getPrintOptions("Sheet1");
```

### Header and Footer

#### `set_header_footer` / `setHeaderFooter`

Set header and/or footer text for printing. Uses Excel formatting codes:
- `&L` left section
- `&C` center section
- `&R` right section
- `&P` page number
- `&N` total pages
- `&D` date
- `&T` time
- `&F` file name

**Rust:**

```rust
wb.set_header_footer("Sheet1", Some("&CMonthly Report"), Some("&LPage &P of &N"))?;
```

**TypeScript:**

```typescript
wb.setHeaderFooter("Sheet1", "&CMonthly Report", "&LPage &P of &N");
```

#### `get_header_footer` / `getHeaderFooter`

Get header and footer text for a sheet.

**Rust:**

```rust
let (header, footer) = wb.get_header_footer("Sheet1")?;
```

**TypeScript:**

```typescript
const result = wb.getHeaderFooter("Sheet1");
// { header?: string, footer?: string }
```

### Page Breaks

#### `insert_page_break` / `insertPageBreak`

Insert a horizontal page break before the given 1-based row number.

**Rust:**

```rust
wb.insert_page_break("Sheet1", 20)?;
```

**TypeScript:**

```typescript
wb.insertPageBreak("Sheet1", 20);
```

#### `remove_page_break` / `removePageBreak`

Remove a page break at the given row.

**Rust:**

```rust
wb.remove_page_break("Sheet1", 20)?;
```

**TypeScript:**

```typescript
wb.removePageBreak("Sheet1", 20);
```

#### `get_page_breaks` / `getPageBreaks`

Get all row page break positions (1-based).

**Rust:**

```rust
let breaks: Vec<u32> = wb.get_page_breaks("Sheet1")?;
```

**TypeScript:**

```typescript
const breaks: number[] = wb.getPageBreaks("Sheet1");
```

---

## 18. Defined Names

Defined names (named ranges) assign a symbolic name to a cell reference or formula. They are available as standalone functions in `sheetkit_core::defined_names` (Rust only). There is currently no Workbook-level wrapper or Node.js binding for defined names.

### `set_defined_name` (Rust only)

Add or update a defined name.

```rust
use sheetkit_core::defined_names::{set_defined_name, DefinedNameScope};

set_defined_name(
    &mut workbook_xml,
    "SalesTotal",
    "Sheet1!$B$10",
    DefinedNameScope::Workbook,
    None, // optional comment
)?;
```

### `get_defined_name` (Rust only)

Get the value of a defined name.

```rust
use sheetkit_core::defined_names::get_defined_name;

if let Some(info) = get_defined_name(&workbook_xml, "SalesTotal", DefinedNameScope::Workbook) {
    println!("Refers to: {}", info.value);
}
```

### `delete_defined_name` (Rust only)

Delete a defined name. Returns `true` if the name existed.

```rust
use sheetkit_core::defined_names::delete_defined_name;

let removed: bool = delete_defined_name(&mut workbook_xml, "SalesTotal", DefinedNameScope::Workbook);
```

> Note: These functions operate on the low-level `WorkbookXml` struct, not on the `Workbook` facade. They support both workbook-scoped and sheet-scoped names via `DefinedNameScope`.

---

## 19. Document Properties

Document properties store metadata about the workbook file.

### Core Properties

#### `set_doc_props` / `setDocProps`

Set core document properties (title, creator, etc.).

**Rust:**

```rust
use sheetkit::doc_props::DocProperties;

wb.set_doc_props(DocProperties {
    title: Some("Annual Report".to_string()),
    creator: Some("Finance Team".to_string()),
    subject: Some("Financial Summary".to_string()),
    ..Default::default()
});
```

**TypeScript:**

```typescript
wb.setDocProps({
    title: "Annual Report",
    creator: "Finance Team",
    subject: "Financial Summary",
});
```

#### `get_doc_props` / `getDocProps`

Get core document properties.

**Rust:**

```rust
let props = wb.get_doc_props();
```

**TypeScript:**

```typescript
const props = wb.getDocProps();
```

#### Core Properties Fields

| Field | Type | Description |
|---|---|---|
| `title` | `Option<String>` / `string?` | Document title |
| `subject` | `Option<String>` / `string?` | Subject |
| `creator` | `Option<String>` / `string?` | Author name |
| `keywords` | `Option<String>` / `string?` | Keywords |
| `description` | `Option<String>` / `string?` | Description/comments |
| `last_modified_by` | `Option<String>` / `string?` | Last editor |
| `revision` | `Option<String>` / `string?` | Revision number |
| `created` | `Option<String>` / `string?` | Creation timestamp (ISO 8601) |
| `modified` | `Option<String>` / `string?` | Last modified timestamp |
| `category` | `Option<String>` / `string?` | Category |
| `content_status` | `Option<String>` / `string?` | Content status |

### Application Properties

#### `set_app_props` / `setAppProps`

Set application properties.

**Rust:**

```rust
use sheetkit::doc_props::AppProperties;

wb.set_app_props(AppProperties {
    application: Some("SheetKit".to_string()),
    company: Some("Acme Corp".to_string()),
    ..Default::default()
});
```

**TypeScript:**

```typescript
wb.setAppProps({
    application: "SheetKit",
    company: "Acme Corp",
});
```

#### `get_app_props` / `getAppProps`

Get application properties.

#### Application Properties Fields

| Field | Type | Description |
|---|---|---|
| `application` | `Option<String>` / `string?` | Application name |
| `doc_security` | `Option<u32>` / `number?` | Document security level |
| `company` | `Option<String>` / `string?` | Company name |
| `app_version` | `Option<String>` / `string?` | Application version |
| `manager` | `Option<String>` / `string?` | Manager name |
| `template` | `Option<String>` / `string?` | Template name |

### Custom Properties

Custom properties store arbitrary key-value metadata.

#### `set_custom_property` / `setCustomProperty`

Set a custom property. Accepts String, Int, Float, Bool, or DateTime values.

**Rust:**

```rust
use sheetkit::doc_props::CustomPropertyValue;

wb.set_custom_property("Department", CustomPropertyValue::String("Engineering".to_string()));
wb.set_custom_property("Version", CustomPropertyValue::Int(3));
wb.set_custom_property("Approved", CustomPropertyValue::Bool(true));
```

**TypeScript:**

```typescript
wb.setCustomProperty("Department", "Engineering");
wb.setCustomProperty("Version", 3);
wb.setCustomProperty("Approved", true);
```

> Note: In TypeScript, numeric values are automatically distinguished as integer or float. Integer-like numbers (no fractional part) within the i32 range are stored as `Int`; others are stored as `Float`.

#### `get_custom_property` / `getCustomProperty`

Get a custom property value, or `None`/`null` if not found.

**Rust:**

```rust
if let Some(value) = wb.get_custom_property("Department") {
    // value is CustomPropertyValue
}
```

**TypeScript:**

```typescript
const value = wb.getCustomProperty("Department");
// string | number | boolean | null
```

#### `delete_custom_property` / `deleteCustomProperty`

Delete a custom property. Returns `true` if it existed.

**Rust:**

```rust
let existed: bool = wb.delete_custom_property("Department");
```

**TypeScript:**

```typescript
const existed: boolean = wb.deleteCustomProperty("Department");
```

---

## 20. Workbook Protection

Workbook protection prevents structural changes to the workbook (adding, removing, or renaming sheets).

### `protect_workbook` / `protectWorkbook`

Protect the workbook with optional password and lock settings.

**Rust:**

```rust
use sheetkit::protection::WorkbookProtectionConfig;

wb.protect_workbook(WorkbookProtectionConfig {
    password: Some("secret".to_string()),
    lock_structure: true,
    lock_windows: false,
    lock_revision: false,
});
```

**TypeScript:**

```typescript
wb.protectWorkbook({
    password: "secret",
    lockStructure: true,
    lockWindows: false,
    lockRevision: false,
});
```

### `unprotect_workbook` / `unprotectWorkbook`

Remove workbook protection.

**Rust:**

```rust
wb.unprotect_workbook();
```

**TypeScript:**

```typescript
wb.unprotectWorkbook();
```

### `is_workbook_protected` / `isWorkbookProtected`

Check whether the workbook is protected.

**Rust:**

```rust
let protected: bool = wb.is_workbook_protected();
```

**TypeScript:**

```typescript
const isProtected: boolean = wb.isWorkbookProtected();
```

### WorkbookProtectionConfig

| Field | Type | Description |
|---|---|---|
| `password` | `Option<String>` / `string?` | Password (hashed with legacy Excel algorithm) |
| `lock_structure` | `bool` / `boolean?` | Prevent adding/removing/renaming sheets |
| `lock_windows` | `bool` / `boolean?` | Prevent moving/resizing workbook windows |
| `lock_revision` | `bool` / `boolean?` | Lock revision tracking |

> Note: The password uses the legacy Excel hash algorithm, which is NOT cryptographically secure. It provides only basic deterrence.

---

## 21. Sheet Protection

Sheet protection prevents editing of cells within a single sheet. It is available as standalone functions in `sheetkit_core::sheet` (Rust only). There is currently no Workbook-level wrapper or Node.js binding for sheet protection.

### `protect_sheet` (Rust only)

Protect a sheet with optional password and granular permission settings.

```rust
use sheetkit_core::sheet::{protect_sheet, SheetProtectionConfig};

protect_sheet(&mut worksheet_xml, &SheetProtectionConfig {
    password: Some("mypass".to_string()),
    select_locked_cells: false,
    select_unlocked_cells: false,
    format_cells: false,
    format_columns: false,
    format_rows: false,
    insert_columns: false,
    insert_rows: false,
    insert_hyperlinks: false,
    delete_columns: false,
    delete_rows: false,
    sort: false,
    auto_filter: false,
    pivot_tables: false,
})?;
```

### `unprotect_sheet` (Rust only)

Remove protection from a sheet.

```rust
use sheetkit_core::sheet::unprotect_sheet;

unprotect_sheet(&mut worksheet_xml)?;
```

> Note: These functions operate on `WorksheetXml` directly. The permission booleans default to `false` (locked). Setting a permission to `true` allows that action even when the sheet is protected.

---

## 22. Formula Evaluation

SheetKit includes a formula evaluator that supports 110 Excel functions. Formulas are parsed using a nom-based parser and evaluated against the current workbook data.

### `evaluate_formula` / `evaluateFormula`

Evaluate a single formula string in the context of a specific sheet.

**Rust:**

```rust
let result: CellValue = wb.evaluate_formula("Sheet1", "SUM(A1:A10)")?;
```

**TypeScript:**

```typescript
const result = wb.evaluateFormula("Sheet1", "SUM(A1:A10)");
// returns: string | number | boolean | DateValue | null
```

### `calculate_all` / `calculateAll`

Recalculate all formula cells in the workbook. Uses a dependency graph with topological sort (Kahn's algorithm) to ensure formulas are calculated in the correct order.

**Rust:**

```rust
wb.calculate_all()?;
```

**TypeScript:**

```typescript
wb.calculateAll();
```

### Supported Functions (110)

#### Math (23 functions)

`SUM`, `ABS`, `INT`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `MOD`, `POWER`, `SQRT`, `CEILING`, `FLOOR`, `SIGN`, `RAND`, `RANDBETWEEN`, `PI`, `LOG`, `LOG10`, `LN`, `EXP`, `PRODUCT`, `QUOTIENT`, `FACT`, `SUMIF`, `SUMIFS`

#### Statistical (15 functions)

`AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `AVERAGEIF`, `AVERAGEIFS`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `MEDIAN`, `MODE`, `LARGE`, `SMALL`, `RANK`

#### Text (18 functions)

`LEN`, `LOWER`, `UPPER`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `CONCATENATE`, `CONCAT`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `REPT`, `EXACT`, `T`, `PROPER`, `VALUE`, `TEXT`

#### Logical (11 functions)

`IF`, `AND`, `OR`, `NOT`, `TRUE`, `FALSE`, `IFERROR`, `IFNA`, `IFS`, `SWITCH`, `XOR`

#### Information (13 functions)

`ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISERROR`, `ISERR`, `ISNA`, `ISLOGICAL`, `ISEVEN`, `ISODD`, `TYPE`, `N`, `NA`, `ERROR.TYPE`

#### Date/Time (17 functions)

`DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `HOUR`, `MINUTE`, `SECOND`, `DATEDIF`, `EDATE`, `EOMONTH`, `DATEVALUE`, `WEEKDAY`, `WEEKNUM`, `NETWORKDAYS`, `WORKDAY`

#### Lookup (11 functions)

`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `LOOKUP`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `CHOOSE`, `ADDRESS`

> Note: Function names are case-insensitive. Unsupported functions return an error. The evaluator supports cell references (A1, $B$2), range references (A1:C10), cross-sheet references (Sheet2!A1), and standard arithmetic operators (+, -, *, /, ^, &, comparison operators).

---

## 23. Pivot Tables

Pivot tables summarize data from a source range into a structured report.

### `add_pivot_table` / `addPivotTable`

Add a pivot table to the workbook.

**Rust:**

```rust
use sheetkit::pivot::{PivotTableConfig, PivotField, PivotDataField, AggregateFunction};

let config = PivotTableConfig {
    name: "SalesPivot".to_string(),
    source_sheet: "Data".to_string(),
    source_range: "A1:D100".to_string(),
    target_sheet: "PivotSheet".to_string(),
    target_cell: "A1".to_string(),
    rows: vec![PivotField { name: "Region".to_string() }],
    columns: vec![PivotField { name: "Quarter".to_string() }],
    data: vec![PivotDataField {
        name: "Revenue".to_string(),
        function: AggregateFunction::Sum,
        display_name: Some("Total Revenue".to_string()),
    }],
};
wb.add_pivot_table(&config)?;
```

**TypeScript:**

```typescript
wb.addPivotTable({
    name: "SalesPivot",
    sourceSheet: "Data",
    sourceRange: "A1:D100",
    targetSheet: "PivotSheet",
    targetCell: "A1",
    rows: [{ name: "Region" }],
    columns: [{ name: "Quarter" }],
    data: [{
        name: "Revenue",
        function: "sum",
        displayName: "Total Revenue",
    }],
});
```

### `get_pivot_tables` / `getPivotTables`

Get all pivot tables in the workbook.

**Rust:**

```rust
let tables: Vec<PivotTableInfo> = wb.get_pivot_tables();
for t in &tables {
    println!("{}: {} -> {}", t.name, t.source_range, t.location);
}
```

**TypeScript:**

```typescript
const tables = wb.getPivotTables();
```

### `delete_pivot_table` / `deletePivotTable`

Delete a pivot table by name.

**Rust:**

```rust
wb.delete_pivot_table("SalesPivot")?;
```

**TypeScript:**

```typescript
wb.deletePivotTable("SalesPivot");
```

### PivotTableConfig

| Field | Type | Description |
|---|---|---|
| `name` | `String` / `string` | Pivot table name |
| `source_sheet` | `String` / `string` | Source data sheet name |
| `source_range` | `String` / `string` | Source data range (e.g., "A1:D100") |
| `target_sheet` | `String` / `string` | Target sheet for the pivot table |
| `target_cell` | `String` / `string` | Top-left cell of the pivot table |
| `rows` | `Vec<PivotField>` / `PivotField[]` | Row fields |
| `columns` | `Vec<PivotField>` / `PivotField[]` | Column fields |
| `data` | `Vec<PivotDataField>` / `PivotDataField[]` | Data/value fields |

### PivotDataField

| Field | Type | Description |
|---|---|---|
| `name` | `String` / `string` | Column name from source data header |
| `function` | `AggregateFunction` / `string` | Aggregate function |
| `display_name` | `Option<String>` / `string?` | Custom display name |

### Aggregate Functions

| Rust | TypeScript | Description |
|---|---|---|
| `AggregateFunction::Sum` | `"sum"` | Sum of values |
| `AggregateFunction::Count` | `"count"` | Count of entries |
| `AggregateFunction::Average` | `"average"` | Average |
| `AggregateFunction::Max` | `"max"` | Maximum |
| `AggregateFunction::Min` | `"min"` | Minimum |
| `AggregateFunction::Product` | `"product"` | Product |
| `AggregateFunction::CountNums` | `"countNums"` | Count of numeric values |
| `AggregateFunction::StdDev` | `"stdDev"` | Standard deviation (sample) |
| `AggregateFunction::StdDevP` | `"stdDevP"` | Standard deviation (population) |
| `AggregateFunction::Var` | `"var"` | Variance (sample) |
| `AggregateFunction::VarP` | `"varP"` | Variance (population) |

---

## 24. StreamWriter

The `StreamWriter` provides a forward-only streaming API for writing large sheets without holding the entire worksheet in memory. Rows must be written in ascending order.

### Basic Workflow

1. Create a stream writer from the workbook
2. Set column widths and other column settings (must be done BEFORE writing any rows)
3. Write rows in ascending order
4. Apply the stream writer back to the workbook

**Rust:**

```rust
use sheetkit::cell::CellValue;

let mut sw = wb.new_stream_writer("LargeSheet")?;

// Set column widths BEFORE writing rows
sw.set_col_width(1, 20.0)?;
sw.set_col_width(2, 15.0)?;

// Write header
sw.write_row(1, &[
    CellValue::from("Name"),
    CellValue::from("Score"),
])?;

// Write data rows
for i in 2..=1000 {
    sw.write_row(i, &[
        CellValue::from(format!("Item {}", i - 1)),
        CellValue::from(i as f64 * 1.5),
    ])?;
}

// Apply to workbook
let sheet_index = wb.apply_stream_writer(sw)?;
wb.save("large_output.xlsx")?;
```

**TypeScript:**

```typescript
const sw = wb.newStreamWriter("LargeSheet");

// Set column widths BEFORE writing rows
sw.setColWidth(1, 20.0);
sw.setColWidth(2, 15.0);

// Write header
sw.writeRow(1, ["Name", "Score"]);

// Write data rows
for (let i = 2; i <= 1000; i++) {
    sw.writeRow(i, [`Item ${i - 1}`, i * 1.5]);
}

// Apply to workbook
const sheetIndex = wb.applyStreamWriter(sw);
wb.save("large_output.xlsx");
```

### StreamWriter API

#### `set_col_width(col, width)` / `setColWidth(col, width)`

Set the width of a single column. Column numbers are 1-based. Must be called before any `write_row`.

#### `set_col_width_range(min, max, width)` / `setColWidthRange(min, max, width)`

Set the width for a range of columns (inclusive). Must be called before any `write_row`.

#### `write_row(row, values)` / `writeRow(row, values)`

Write a row of values. Row numbers are 1-based and must be written in ascending order.

#### `add_merge_cell(reference)` / `addMergeCell(reference)`

Register a merge cell range (e.g., "A1:C1").

**Rust:**

```rust
sw.add_merge_cell("A1:C1")?;
```

**TypeScript:**

```typescript
sw.addMergeCell("A1:C1");
```

#### Rust-Only StreamWriter Methods

The following methods are available only in the Rust API:

- `set_freeze_panes(cell)` -- Set freeze panes for the streamed sheet (must be called before writing rows)
- `set_col_visible(col, visible)` -- Set column visibility
- `set_col_outline_level(col, level)` -- Set column outline level (0-7)
- `set_col_style(col, style_id)` -- Set column style
- `write_row_with_options(row, values, options)` -- Write a row with custom options (height, visibility, outline level, style)

```rust
use sheetkit::stream::StreamRowOptions;

sw.set_freeze_panes("A2")?; // freeze row 1
sw.set_col_visible(3, false)?; // hide column C
sw.set_col_style(1, style_id)?;

sw.write_row_with_options(1, &values, &StreamRowOptions {
    height: Some(25.0),
    visible: Some(true),
    outline_level: Some(1),
    style_id: Some(style_id),
})?;
```

> Important: Column widths, visibility, styles, outline levels, and freeze panes must ALL be set before the first `write_row` call. Setting them after writing any rows returns an error.

---

## 25. Utility Functions

These utility functions are available in the Rust API only (`sheetkit_core::utils::cell_ref`).

### `cell_name_to_coordinates`

Convert an A1-style cell reference to 1-based (column, row) coordinates. Supports absolute references (e.g., "$B$3").

```rust
use sheetkit_core::utils::cell_ref::cell_name_to_coordinates;

let (col, row) = cell_name_to_coordinates("B3")?;
assert_eq!((col, row), (2, 3));

let (col, row) = cell_name_to_coordinates("$AA$100")?;
assert_eq!((col, row), (27, 100));
```

### `coordinates_to_cell_name`

Convert 1-based (column, row) coordinates to an A1-style cell reference.

```rust
use sheetkit_core::utils::cell_ref::coordinates_to_cell_name;

let name = coordinates_to_cell_name(2, 3)?;
assert_eq!(name, "B3");
```

### `column_name_to_number`

Convert a column letter name to a 1-based column number.

```rust
use sheetkit_core::utils::cell_ref::column_name_to_number;

assert_eq!(column_name_to_number("A")?, 1);
assert_eq!(column_name_to_number("Z")?, 26);
assert_eq!(column_name_to_number("AA")?, 27);
assert_eq!(column_name_to_number("XFD")?, 16384);
```

### `column_number_to_name`

Convert a 1-based column number to its letter name.

```rust
use sheetkit_core::utils::cell_ref::column_number_to_name;

assert_eq!(column_number_to_name(1)?, "A");
assert_eq!(column_number_to_name(26)?, "Z");
assert_eq!(column_number_to_name(27)?, "AA");
assert_eq!(column_number_to_name(16384)?, "XFD");
```

### Date Conversion Functions

Available in `sheetkit_core::cell`:

- `date_to_serial(NaiveDate) -> f64` -- Convert a chrono date to an Excel serial number
- `datetime_to_serial(NaiveDateTime) -> f64` -- Convert a chrono datetime to an Excel serial number with time fraction
- `serial_to_date(f64) -> Option<NaiveDate>` -- Convert an Excel serial number to a date
- `serial_to_datetime(f64) -> Option<NaiveDateTime>` -- Convert an Excel serial number to a datetime

```rust
use chrono::NaiveDate;
use sheetkit_core::cell::{date_to_serial, serial_to_date};

let date = NaiveDate::from_ymd_opt(2025, 6, 15).unwrap();
let serial = date_to_serial(date);
let roundtrip = serial_to_date(serial).unwrap();
assert_eq!(date, roundtrip);
```

> Note: Excel uses the 1900 date system with a known bug where it incorrectly treats 1900 as a leap year. Serial number 60 (February 29, 1900) does not correspond to a real date. These conversion functions account for this bug.

### `is_date_num_fmt(num_fmt_id)` (Rust only)

Check whether a built-in number format ID represents a date or time format. Returns `true` for IDs 14-22 and 45-47.

```rust
use sheetkit::is_date_num_fmt;

assert!(is_date_num_fmt(14));   // m/d/yyyy
assert!(is_date_num_fmt(22));   // m/d/yyyy h:mm
assert!(!is_date_num_fmt(0));   // General
assert!(!is_date_num_fmt(49));  // @
```

### `is_date_format_code(code)` (Rust only)

Check whether a custom number format string represents a date or time format. Returns `true` if the format code contains date/time tokens (y, m, d, h, s) outside of quoted strings and escaped characters.

```rust
use sheetkit::is_date_format_code;

assert!(is_date_format_code("yyyy-mm-dd"));
assert!(is_date_format_code("h:mm:ss AM/PM"));
assert!(!is_date_format_code("#,##0.00"));
assert!(!is_date_format_code("0%"));
```

---

## 26. Sparklines

Sparklines are mini-charts embedded in worksheet cells. SheetKit supports three sparkline types: Line, Column, and Win/Loss. Excel defines 36 style presets (indices 0-35).

> **Note:** Sparkline types, configuration, and XML conversion are available. Workbook integration (`addSparkline` / `getSparklines`) will be added in a future release.

### Types

#### `SparklineType` (Rust) / `sparklineType` (TypeScript)

| Value | Rust | TypeScript | OOXML |
|-------|------|------------|-------|
| Line | `SparklineType::Line` | `"line"` | (default, omitted) |
| Column | `SparklineType::Column` | `"column"` | `"column"` |
| Win/Loss | `SparklineType::WinLoss` | `"winloss"` or `"stacked"` | `"stacked"` |

#### `SparklineConfig` (Rust)

```rust
use sheetkit::SparklineConfig;

let config = SparklineConfig::new("Sheet1!A1:A10", "B1");
```

Fields:

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `data_range` | `String` | (required) | Data source range (e.g., `"Sheet1!A1:A10"`) |
| `location` | `String` | (required) | Cell where sparkline is rendered (e.g., `"B1"`) |
| `sparkline_type` | `SparklineType` | `Line` | Sparkline chart type |
| `markers` | `bool` | `false` | Show data markers |
| `high_point` | `bool` | `false` | Highlight highest point |
| `low_point` | `bool` | `false` | Highlight lowest point |
| `first_point` | `bool` | `false` | Highlight first point |
| `last_point` | `bool` | `false` | Highlight last point |
| `negative_points` | `bool` | `false` | Highlight negative values |
| `show_axis` | `bool` | `false` | Show horizontal axis |
| `line_weight` | `Option<f64>` | `None` | Line weight in points |
| `style` | `Option<u32>` | `None` | Style preset index (0-35) |

#### `JsSparklineConfig` (TypeScript)

```typescript
const config = {
  dataRange: 'Sheet1!A1:A10',
  location: 'B1',
  sparklineType: 'line',    // "line" | "column" | "winloss" | "stacked"
  markers: true,
  highPoint: false,
  lowPoint: false,
  firstPoint: false,
  lastPoint: false,
  negativePoints: false,
  showAxis: false,
  lineWeight: 0.75,
  style: 1,
};
```

### Validation

The `validate_sparkline_config` function (Rust) checks that:
- `data_range` is not empty
- `location` is not empty
- `line_weight` (if set) is positive
- `style` (if set) is in range 0-35

```rust
use sheetkit_core::sparkline::{SparklineConfig, validate_sparkline_config};

let config = SparklineConfig::new("Sheet1!A1:A10", "B1");
validate_sparkline_config(&config).unwrap(); // Ok
```

## 27. Theme Colors

Resolve theme color slots (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink) with optional tint.

### Workbook.getThemeColor (Node.js) / Workbook::get_theme_color (Rust)

| Parameter | Type              | Description                                      |
| --------- | ----------------- | ------------------------------------------------ |
| index     | `u32` / `number`  | Theme color index (0-11)                         |
| tint      | `Option<f64>` / `number \| null` | Tint value: positive lightens, negative darkens |

**Returns:** ARGB hex string (e.g. `"FF4472C4"`) or `None`/`null` if out of range.

**Theme Color Indices:**

| Index | Slot Name | Default Color |
| ----- | --------- | ------------- |
| 0     | dk1       | FF000000      |
| 1     | lt1       | FFFFFFFF      |
| 2     | dk2       | FF44546A      |
| 3     | lt2       | FFE7E6E6      |
| 4     | accent1   | FF4472C4      |
| 5     | accent2   | FFED7D31      |
| 6     | accent3   | FFA5A5A5      |
| 7     | accent4   | FFFFC000      |
| 8     | accent5   | FF5B9BD5      |
| 9     | accent6   | FF70AD47      |
| 10    | hlink     | FF0563C1      |
| 11    | folHlink  | FF954F72      |

#### Node.js

```javascript
const wb = new Workbook();

// Get accent1 color (no tint)
const color = wb.getThemeColor(4, null); // "FF4472C4"

// Lighten black by 50%
const lightened = wb.getThemeColor(0, 0.5); // "FF7F7F7F"

// Darken white by 50%
const darkened = wb.getThemeColor(1, -0.5); // "FF7F7F7F"

// Out of range returns null
const invalid = wb.getThemeColor(99, null); // null
```

#### Rust

```rust
let wb = Workbook::new();

// Get accent1 color (no tint)
let color = wb.get_theme_color(4, None); // Some("FF4472C4")

// Apply tint
let tinted = wb.get_theme_color(0, Some(0.5)); // Some("FF7F7F7F")
```

### Gradient Fill

The `FillStyle` type supports gradient fills via the `gradient` field.

#### Types

```rust
pub struct GradientFillStyle {
    pub gradient_type: GradientType, // Linear or Path
    pub degree: Option<f64>,         // Rotation angle for linear gradients
    pub left: Option<f64>,           // Path gradient coordinates (0.0-1.0)
    pub right: Option<f64>,
    pub top: Option<f64>,
    pub bottom: Option<f64>,
    pub stops: Vec<GradientStop>,    // Color stops
}

pub struct GradientStop {
    pub position: f64,     // Position (0.0-1.0)
    pub color: StyleColor, // Color at this stop
}

pub enum GradientType {
    Linear,
    Path,
}
```

#### Rust Example

```rust
use sheetkit::*;

let mut wb = Workbook::new();
let style_id = wb.add_style(&Style {
    fill: Some(FillStyle {
        pattern: PatternType::None,
        fg_color: None,
        bg_color: None,
        gradient: Some(GradientFillStyle {
            gradient_type: GradientType::Linear,
            degree: Some(90.0),
            left: None,
            right: None,
            top: None,
            bottom: None,
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: StyleColor::Rgb("FFFFFFFF".to_string()),
                },
                GradientStop {
                    position: 1.0,
                    color: StyleColor::Rgb("FF4472C4".to_string()),
                },
            ],
        }),
    }),
    ..Style::default()
})?;
```

---

## 28. Rich Text

Rich text allows a single cell to contain multiple text segments (runs), each with independent formatting such as font, size, bold, italic, and color.

### `RichTextRun` Type

Each run in a rich text cell is described by a `RichTextRun`.

**Rust:**

```rust
pub struct RichTextRun {
    pub text: String,
    pub font: Option<String>,
    pub size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub color: Option<String>,
}
```

**TypeScript:**

```typescript
interface RichTextRun {
  text: string;
  font?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;  // RGB hex string, e.g. "#FF0000"
}
```

### `set_cell_rich_text` / `setCellRichText`

Set a cell value to rich text with multiple formatted runs.

**Rust:**

```rust
use sheetkit::{Workbook, RichTextRun};

let mut wb = Workbook::new();
let runs = vec![
    RichTextRun {
        text: "Bold text".to_string(),
        font: Some("Arial".to_string()),
        size: Some(14.0),
        bold: true,
        italic: false,
        color: Some("#FF0000".to_string()),
    },
    RichTextRun {
        text: " normal text".to_string(),
        font: None,
        size: None,
        bold: false,
        italic: false,
        color: None,
    },
];
wb.set_cell_rich_text("Sheet1", "A1", runs)?;
```

**TypeScript:**

```typescript
const wb = new Workbook();
wb.setCellRichText("Sheet1", "A1", [
  { text: "Bold text", font: "Arial", size: 14, bold: true, color: "#FF0000" },
  { text: " normal text" },
]);
```

### `get_cell_rich_text` / `getCellRichText`

Retrieve the rich text runs for a cell. Returns `None`/`null` for non-rich-text cells.

**Rust:**

```rust
let runs = wb.get_cell_rich_text("Sheet1", "A1")?;
if let Some(runs) = runs {
    for run in &runs {
        println!("Text: {:?}, Bold: {}", run.text, run.bold);
    }
}
```

**TypeScript:**

```typescript
const runs = wb.getCellRichText("Sheet1", "A1");
if (runs) {
  for (const run of runs) {
    console.log(`Text: ${run.text}, Bold: ${run.bold ?? false}`);
  }
}
```

### `CellValue::RichString` (Rust only)

Rich text cells use the `CellValue::RichString(Vec<RichTextRun>)` variant. When read through `get_cell_value`, the display value is the concatenation of all run texts.

```rust
match wb.get_cell_value("Sheet1", "A1")? {
    CellValue::RichString(runs) => {
        println!("Rich text with {} runs", runs.len());
    }
    _ => {}
}
```

### `rich_text_to_plain`

Utility function to extract the concatenated plain text from a slice of rich text runs.

**Rust:**

```rust
use sheetkit::rich_text_to_plain;

let plain = rich_text_to_plain(&runs);
```
