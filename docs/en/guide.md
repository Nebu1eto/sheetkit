# SheetKit User Guide

SheetKit is a Rust library for reading and writing Excel (.xlsx) files, with first-class Node.js bindings via napi-rs.

---

## Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [API Reference](#api-reference)
  - [Workbook I/O](#workbook-io)
  - [Cell Operations](#cell-operations)
  - [Sheet Management](#sheet-management)
  - [Row and Column Operations](#row-and-column-operations)
  - [Styles](#styles)
  - [Charts](#charts)
  - [Images](#images)
  - [Data Validation](#data-validation)
  - [Comments](#comments)
  - [Auto-Filter](#auto-filter)
  - [StreamWriter](#streamwriter)
  - [Document Properties](#document-properties)
  - [Workbook Protection](#workbook-protection)
- [Examples](#examples)

---

## Installation

### Rust

Add `sheetkit` to your `Cargo.toml`:

```toml
[dependencies]
sheetkit = "0.1"
```

### Node.js

```bash
npm install sheetkit
```

> The Node.js package is a native addon built with napi-rs. A Rust build toolchain (rustc, cargo) is required to compile the native module during installation.

---

## Quick Start

### Rust

```rust
use sheetkit::{CellValue, Workbook};

fn main() -> sheetkit::Result<()> {
    // Create a new workbook (contains "Sheet1" by default)
    let mut wb = Workbook::new();

    // Write cell values
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Name".into()))?;
    wb.set_cell_value("Sheet1", "B1", CellValue::String("Age".into()))?;
    wb.set_cell_value("Sheet1", "A2", CellValue::String("Alice".into()))?;
    wb.set_cell_value("Sheet1", "B2", CellValue::Number(30.0))?;

    // Read a cell value
    let val = wb.get_cell_value("Sheet1", "A1")?;
    println!("A1 = {:?}", val);

    // Save to file
    wb.save("output.xlsx")?;

    // Open an existing file
    let wb2 = Workbook::open("output.xlsx")?;
    println!("Sheets: {:?}", wb2.sheet_names());

    Ok(())
}
```

### TypeScript / Node.js

```typescript
import { Workbook } from 'sheetkit';

// Create a new workbook (contains "Sheet1" by default)
const wb = new Workbook();

// Write cell values
wb.setCellValue('Sheet1', 'A1', 'Name');
wb.setCellValue('Sheet1', 'B1', 'Age');
wb.setCellValue('Sheet1', 'A2', 'Alice');
wb.setCellValue('Sheet1', 'B2', 30);

// Read a cell value
const val = wb.getCellValue('Sheet1', 'A1');
console.log('A1 =', val);

// Save to file
wb.save('output.xlsx');

// Open an existing file
const wb2 = Workbook.open('output.xlsx');
console.log('Sheets:', wb2.sheetNames);
```

---

## API Reference

### Workbook I/O

Create, open, and save workbooks.

#### Rust

```rust
use sheetkit::Workbook;

// Create a new empty workbook with a single "Sheet1"
let mut wb = Workbook::new();

// Open an existing .xlsx file
let wb = Workbook::open("input.xlsx")?;

// Save the workbook to a .xlsx file
wb.save("output.xlsx")?;

// Get the names of all sheets
let names: Vec<&str> = wb.sheet_names();
```

#### TypeScript

```typescript
import { Workbook } from 'sheetkit';

// Create a new empty workbook with a single "Sheet1"
const wb = new Workbook();

// Open an existing .xlsx file
const wb2 = Workbook.open('input.xlsx');

// Save the workbook to a .xlsx file
wb.save('output.xlsx');

// Get the names of all sheets
const names: string[] = wb.sheetNames;
```

---

### Cell Operations

Read and write cell values. Cells are identified by sheet name and cell reference (e.g., `"A1"`, `"B2"`, `"AA100"`).

#### CellValue Types

| Rust Variant             | TypeScript Type | Description                                |
|--------------------------|-----------------|--------------------------------------------|
| `CellValue::String(s)`  | `string`        | Text value                                 |
| `CellValue::Number(n)`  | `number`        | Numeric value (stored as f64 internally)   |
| `CellValue::Bool(b)`    | `boolean`       | Boolean value                              |
| `CellValue::Empty`      | `null`          | Empty cell / clear value                   |
| `CellValue::Formula{..}`| --              | Formula (Rust only)                        |
| `CellValue::Error(e)`   | --              | Error value such as `#DIV/0!` (Rust only)  |

#### Rust

```rust
use sheetkit::{CellValue, Workbook};

let mut wb = Workbook::new();

// Set values of different types
wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))?;
wb.set_cell_value("Sheet1", "B1", CellValue::Number(42.0))?;
wb.set_cell_value("Sheet1", "C1", CellValue::Bool(true))?;
wb.set_cell_value("Sheet1", "D1", CellValue::Empty)?;

// Convenient From conversions
wb.set_cell_value("Sheet1", "A2", CellValue::from("Text"))?;
wb.set_cell_value("Sheet1", "B2", CellValue::from(100i32))?;
wb.set_cell_value("Sheet1", "C2", CellValue::from(3.14))?;

// Read a cell value
let val = wb.get_cell_value("Sheet1", "A1")?;
match val {
    CellValue::String(s) => println!("String: {}", s),
    CellValue::Number(n) => println!("Number: {}", n),
    CellValue::Bool(b) => println!("Bool: {}", b),
    CellValue::Empty => println!("(empty)"),
    _ => {}
}
```

#### TypeScript

```typescript
// Set values -- the type is inferred from the JavaScript value
wb.setCellValue('Sheet1', 'A1', 'Hello');       // string
wb.setCellValue('Sheet1', 'B1', 42);            // number
wb.setCellValue('Sheet1', 'C1', true);          // boolean
wb.setCellValue('Sheet1', 'D1', null);          // clear cell

// Read a cell value -- returns string | number | boolean | null
const val = wb.getCellValue('Sheet1', 'A1');
```

---

### Sheet Management

Create, delete, rename, copy, and navigate sheets.

#### Rust

```rust
let mut wb = Workbook::new();

// Create a new sheet (returns 0-based index)
let idx: usize = wb.new_sheet("Sales")?;

// Delete a sheet by name
wb.delete_sheet("Sales")?;

// Rename a sheet
wb.set_sheet_name("Sheet1", "Main")?;

// Copy a sheet (returns new sheet's 0-based index)
let idx: usize = wb.copy_sheet("Main", "Main_Copy")?;

// Get the index of a sheet (None if not found)
let idx: Option<usize> = wb.get_sheet_index("Main");

// Get/set the active sheet
let active: &str = wb.get_active_sheet();
wb.set_active_sheet("Main")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Create a new sheet (returns 0-based index)
const idx: number = wb.newSheet('Sales');

// Delete a sheet
wb.deleteSheet('Sales');

// Rename a sheet
wb.setSheetName('Sheet1', 'Main');

// Copy a sheet (returns new sheet's 0-based index)
const copyIdx: number = wb.copySheet('Main', 'Main_Copy');

// Get the index of a sheet (null if not found)
const sheetIdx: number | null = wb.getSheetIndex('Main');

// Get/set the active sheet
const active: string = wb.getActiveSheet();
wb.setActiveSheet('Main');
```

---

### Row and Column Operations

Insert, delete, and configure rows and columns.

#### Rust

```rust
let mut wb = Workbook::new();

// -- Rows (1-based row numbers) --

// Insert 3 empty rows starting at row 2
wb.insert_rows("Sheet1", 2, 3)?;

// Remove row 5
wb.remove_row("Sheet1", 5)?;

// Duplicate row 1 (inserts copy below)
wb.duplicate_row("Sheet1", 1)?;

// Set/get row height
wb.set_row_height("Sheet1", 1, 25.0)?;
let height: Option<f64> = wb.get_row_height("Sheet1", 1)?;

// Show/hide a row
wb.set_row_visible("Sheet1", 3, false)?;

// -- Columns (letter-based, e.g., "A", "B", "AA") --

// Set/get column width
wb.set_col_width("Sheet1", "A", 20.0)?;
let width: Option<f64> = wb.get_col_width("Sheet1", "A")?;

// Show/hide a column
wb.set_col_visible("Sheet1", "B", false)?;

// Insert 2 empty columns starting at column "C"
wb.insert_cols("Sheet1", "C", 2)?;

// Remove column "D"
wb.remove_col("Sheet1", "D")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// -- Rows (1-based row numbers) --
wb.insertRows('Sheet1', 2, 3);
wb.removeRow('Sheet1', 5);
wb.duplicateRow('Sheet1', 1);
wb.setRowHeight('Sheet1', 1, 25);
const height: number | null = wb.getRowHeight('Sheet1', 1);
wb.setRowVisible('Sheet1', 3, false);

// -- Columns (letter-based) --
wb.setColWidth('Sheet1', 'A', 20);
const width: number | null = wb.getColWidth('Sheet1', 'A');
wb.setColVisible('Sheet1', 'B', false);
wb.insertCols('Sheet1', 'C', 2);
wb.removeCol('Sheet1', 'D');
```

---

### Styles

Styles control the visual presentation of cells. Register a style definition to get a style ID, then apply that ID to cells. Identical style definitions are deduplicated automatically.

A `Style` can include any combination of: font, fill, border, alignment, number format, and protection.

#### Rust

```rust
use sheetkit::{
    AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle,
    FillStyle, FontStyle, HorizontalAlign, PatternType, Style,
    StyleColor, VerticalAlign, Workbook,
};

let mut wb = Workbook::new();

// Register a style
let style_id = wb.add_style(&Style {
    font: Some(FontStyle {
        name: Some("Arial".into()),
        size: Some(14.0),
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        color: Some(StyleColor::Rgb("#FFFFFF".into())),
    }),
    fill: Some(FillStyle {
        pattern: PatternType::Solid,
        fg_color: Some(StyleColor::Rgb("#4472C4".into())),
        bg_color: None,
    }),
    border: Some(BorderStyle {
        bottom: Some(BorderSideStyle {
            style: BorderLineStyle::Thin,
            color: Some(StyleColor::Rgb("#000000".into())),
        }),
        ..Default::default()
    }),
    alignment: Some(AlignmentStyle {
        horizontal: Some(HorizontalAlign::Center),
        vertical: Some(VerticalAlign::Center),
        wrap_text: true,
        ..Default::default()
    }),
    ..Default::default()
})?;

// Apply the style to a cell
wb.set_cell_style("Sheet1", "A1", style_id)?;

// Read the style ID of a cell (None if default)
let current_style: Option<u32> = wb.get_cell_style("Sheet1", "A1")?;
```

#### TypeScript

```typescript
// Register a style
const styleId = wb.addStyle({
    font: {
        name: 'Arial',
        size: 14,
        bold: true,
        color: '#FFFFFF',
    },
    fill: {
        pattern: 'solid',
        fgColor: '#4472C4',
    },
    border: {
        bottom: { style: 'thin', color: '#000000' },
    },
    alignment: {
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
    },
});

// Apply the style to a cell
wb.setCellStyle('Sheet1', 'A1', styleId);

// Read the style ID of a cell (null if default)
const currentStyle: number | null = wb.getCellStyle('Sheet1', 'A1');
```

#### Style Components Reference

**FontStyle**

| Field           | Rust Type          | TS Type    | Description                     |
|-----------------|--------------------|------------|---------------------------------|
| `name`          | `Option<String>`   | `string?`  | Font family (e.g., "Calibri")   |
| `size`          | `Option<f64>`      | `number?`  | Font size in points             |
| `bold`          | `bool`             | `boolean?` | Bold text                       |
| `italic`        | `bool`             | `boolean?` | Italic text                     |
| `underline`     | `bool`             | `boolean?` | Underline text                  |
| `strikethrough` | `bool`             | `boolean?` | Strikethrough text              |
| `color`         | `Option<StyleColor>` | `string?` | Font color (hex string in TS)  |

**FillStyle**

| Field      | Rust Type          | TS Type   | Description                             |
|------------|--------------------|-----------|-----------------------------------------|
| `pattern`  | `PatternType`      | `string?` | Pattern type (see values below)         |
| `fg_color` | `Option<StyleColor>` | `string?` | Foreground color                      |
| `bg_color` | `Option<StyleColor>` | `string?` | Background color                      |

PatternType values: `None`, `Solid`, `Gray125`, `DarkGray`, `MediumGray`, `LightGray`.
In TypeScript, use lowercase strings: `"none"`, `"solid"`, `"gray125"`, `"darkGray"`, `"mediumGray"`, `"lightGray"`.

**BorderStyle**

Each side (`left`, `right`, `top`, `bottom`, `diagonal`) accepts a `BorderSideStyle` with:
- `style`: one of `Thin`, `Medium`, `Thick`, `Dashed`, `Dotted`, `Double`, `Hair`, `MediumDashed`, `DashDot`, `MediumDashDot`, `DashDotDot`, `MediumDashDotDot`, `SlantDashDot`
- `color`: optional color

In TypeScript, use lowercase strings for border style: `"thin"`, `"medium"`, `"thick"`, etc.

**AlignmentStyle**

| Field           | Rust Type                | TS Type    | Description                 |
|-----------------|--------------------------|------------|-----------------------------|
| `horizontal`    | `Option<HorizontalAlign>`| `string?`  | Horizontal alignment        |
| `vertical`      | `Option<VerticalAlign>`  | `string?`  | Vertical alignment          |
| `wrap_text`     | `bool`                   | `boolean?` | Wrap text                   |
| `text_rotation` | `Option<u32>`            | `number?`  | Text rotation in degrees    |
| `indent`        | `Option<u32>`            | `number?`  | Indentation level           |
| `shrink_to_fit` | `bool`                   | `boolean?` | Shrink text to fit cell     |

HorizontalAlign values: `General`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterContinuous`, `Distributed`.
VerticalAlign values: `Top`, `Center`, `Bottom`, `Justify`, `Distributed`.

**NumFmtStyle** (Rust only)

```rust
use sheetkit::style::NumFmtStyle;

// Built-in format (e.g., percent, date, currency)
NumFmtStyle::Builtin(9)  // 0%

// Custom format string
NumFmtStyle::Custom("#,##0.00".to_string())
```

In TypeScript, use `numFmtId` (built-in format ID) or `customNumFmt` (custom format string) on the style object.

**ProtectionStyle**

| Field    | Rust Type | TS Type    | Description                     |
|----------|-----------|------------|---------------------------------|
| `locked` | `bool`    | `boolean?` | Lock the cell (default: true)   |
| `hidden` | `bool`    | `boolean?` | Hide formulas in protected view |

---

### Charts

Add charts to worksheets. Charts are anchored between two cells (top-left and bottom-right) and render data from specified cell ranges.

#### Supported Chart Types

| Rust Variant                | TS String            | Description                      |
|-----------------------------|----------------------|----------------------------------|
| `ChartType::Col`            | `"col"`              | Vertical bar chart (clustered)   |
| `ChartType::ColStacked`     | `"colStacked"`       | Vertical bar chart (stacked)     |
| `ChartType::ColPercentStacked` | `"colPercentStacked"` | Vertical bar chart (% stacked) |
| `ChartType::Bar`            | `"bar"`              | Horizontal bar chart (clustered) |
| `ChartType::BarStacked`     | `"barStacked"`       | Horizontal bar chart (stacked)   |
| `ChartType::BarPercentStacked` | `"barPercentStacked"` | Horizontal bar chart (% stacked) |
| `ChartType::Line`           | `"line"`             | Line chart                       |
| `ChartType::Pie`            | `"pie"`              | Pie chart                        |

#### Rust

```rust
use sheetkit::{ChartConfig, ChartSeries, ChartType, Workbook};

let mut wb = Workbook::new();

// Populate data first...
wb.set_cell_value("Sheet1", "A1", CellValue::String("Q1".into()))?;
wb.set_cell_value("Sheet1", "B1", CellValue::Number(1500.0))?;
// ... more data rows ...

// Add a chart anchored from D1 to K15
wb.add_chart(
    "Sheet1",
    "D1",   // top-left anchor cell
    "K15",  // bottom-right anchor cell
    &ChartConfig {
        chart_type: ChartType::Col,
        title: Some("Quarterly Revenue".into()),
        series: vec![ChartSeries {
            name: "Revenue".into(),
            categories: "Sheet1!$A$1:$A$4".into(),
            values: "Sheet1!$B$1:$B$4".into(),
        }],
        show_legend: true,
    },
)?;
```

#### TypeScript

```typescript
wb.addChart('Sheet1', 'D1', 'K15', {
    chartType: 'col',
    title: 'Quarterly Revenue',
    series: [
        {
            name: 'Revenue',
            categories: 'Sheet1!$A$1:$A$4',
            values: 'Sheet1!$B$1:$B$4',
        },
    ],
    showLegend: true,
});
```

---

### Images

Embed images (PNG, JPEG, GIF) into worksheets. Images are anchored to a cell and sized by pixel dimensions.

#### Rust

```rust
use sheetkit::{ImageConfig, ImageFormat, Workbook};

let mut wb = Workbook::new();

let image_bytes = std::fs::read("logo.png").unwrap();

wb.add_image(
    "Sheet1",
    &ImageConfig {
        data: image_bytes,
        format: ImageFormat::Png,
        from_cell: "B2".into(),
        width_px: 200,
        height_px: 100,
    },
)?;
```

#### TypeScript

```typescript
import { readFileSync } from 'fs';

const imageData = readFileSync('logo.png');

wb.addImage('Sheet1', {
    data: imageData,
    format: 'png',        // "png" | "jpeg" | "gif"
    fromCell: 'B2',
    widthPx: 200,
    heightPx: 100,
});
```

---

### Data Validation

Add data validation rules to cell ranges. These rules restrict what values users can enter in the specified cells.

#### Validation Types

| Rust Variant             | TS String       | Description                    |
|--------------------------|-----------------|--------------------------------|
| `ValidationType::Whole`  | `"whole"`       | Whole number constraint        |
| `ValidationType::Decimal`| `"decimal"`     | Decimal number constraint      |
| `ValidationType::List`   | `"list"`        | Dropdown list                  |
| `ValidationType::Date`   | `"date"`        | Date constraint                |
| `ValidationType::Time`   | `"time"`        | Time constraint                |
| `ValidationType::TextLength` | `"textLength"` | Text length constraint      |
| `ValidationType::Custom` | `"custom"`      | Custom formula constraint      |

#### Validation Operators

`Between`, `NotBetween`, `Equal`, `NotEqual`, `LessThan`, `LessThanOrEqual`, `GreaterThan`, `GreaterThanOrEqual`.

In TypeScript, use lowercase strings: `"between"`, `"notBetween"`, `"equal"`, etc.

#### Error Styles

`Stop`, `Warning`, `Information` -- controls the severity of the error dialog shown on invalid input.

#### Rust

```rust
use sheetkit::{DataValidationConfig, ErrorStyle, ValidationType, Workbook};

let mut wb = Workbook::new();

// Dropdown list validation
wb.add_data_validation(
    "Sheet1",
    &DataValidationConfig {
        sqref: "C2:C100".into(),
        validation_type: ValidationType::List,
        operator: None,
        formula1: Some("\"Option A,Option B,Option C\"".into()),
        formula2: None,
        allow_blank: true,
        show_input_message: true,
        prompt_title: Some("Select an option".into()),
        prompt_message: Some("Choose from the dropdown".into()),
        show_error_message: true,
        error_style: Some(ErrorStyle::Stop),
        error_title: Some("Invalid".into()),
        error_message: Some("Please select from the list".into()),
    },
)?;

// Read all validations on a sheet
let validations = wb.get_data_validations("Sheet1")?;

// Remove a validation by cell range reference
wb.remove_data_validation("Sheet1", "C2:C100")?;
```

#### TypeScript

```typescript
// Dropdown list validation
wb.addDataValidation('Sheet1', {
    sqref: 'C2:C100',
    validationType: 'list',
    formula1: '"Option A,Option B,Option C"',
    allowBlank: true,
    showInputMessage: true,
    promptTitle: 'Select an option',
    promptMessage: 'Choose from the dropdown',
    showErrorMessage: true,
    errorStyle: 'stop',
    errorTitle: 'Invalid',
    errorMessage: 'Please select from the list',
});

// Read all validations on a sheet
const validations = wb.getDataValidations('Sheet1');

// Remove a validation by cell range reference
wb.removeDataValidation('Sheet1', 'C2:C100');
```

---

### Comments

Add, read, and remove cell comments.

#### Rust

```rust
use sheetkit::{CommentConfig, Workbook};

let mut wb = Workbook::new();

// Add a comment
wb.add_comment(
    "Sheet1",
    &CommentConfig {
        cell: "A1".into(),
        author: "Admin".into(),
        text: "This cell contains the project name.".into(),
    },
)?;

// Get all comments on a sheet
let comments: Vec<CommentConfig> = wb.get_comments("Sheet1")?;

// Remove a comment from a specific cell
wb.remove_comment("Sheet1", "A1")?;
```

#### TypeScript

```typescript
// Add a comment
wb.addComment('Sheet1', {
    cell: 'A1',
    author: 'Admin',
    text: 'This cell contains the project name.',
});

// Get all comments on a sheet
const comments = wb.getComments('Sheet1');

// Remove a comment from a specific cell
wb.removeComment('Sheet1', 'A1');
```

---

### Auto-Filter

Apply or remove auto-filter dropdowns on a range of columns.

#### Rust

```rust
// Set auto-filter on a range
wb.set_auto_filter("Sheet1", "A1:D100")?;

// Remove auto-filter
wb.remove_auto_filter("Sheet1")?;
```

#### TypeScript

```typescript
// Set auto-filter on a range
wb.setAutoFilter('Sheet1', 'A1:D100');

// Remove auto-filter
wb.removeAutoFilter('Sheet1');
```

---

### StreamWriter

The StreamWriter provides a forward-only, streaming API for writing large sheets efficiently. It writes XML directly to an internal buffer, avoiding the need to build the entire worksheet in memory.

Rows must be written in ascending order. Column widths must be set before writing any rows.

#### Rust

```rust
use sheetkit::{CellValue, Workbook};

let mut wb = Workbook::new();

// Create a stream writer for a new sheet
let mut sw = wb.new_stream_writer("LargeSheet")?;

// Set column widths (must be done before writing rows)
sw.set_col_width(1, 20.0)?;     // column 1 (A)
sw.set_col_width(2, 15.0)?;     // column 2 (B)

// Write rows in ascending order (1-based)
sw.write_row(1, &[
    CellValue::from("Name"),
    CellValue::from("Score"),
])?;
for i in 2..=10_000 {
    sw.write_row(i, &[
        CellValue::from(format!("User_{}", i - 1)),
        CellValue::from(i as f64 * 1.5),
    ])?;
}

// Optionally add merge cells
sw.add_merge_cell("A1:B1")?;

// Apply the stream writer to the workbook
wb.apply_stream_writer(sw)?;

wb.save("large_file.xlsx")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Create a stream writer for a new sheet
const sw = wb.newStreamWriter('LargeSheet');

// Set column widths (must be done before writing rows)
sw.setColWidth(1, 20);     // column 1 (A)
sw.setColWidth(2, 15);     // column 2 (B)

// Write rows in ascending order (1-based)
sw.writeRow(1, ['Name', 'Score']);
for (let i = 2; i <= 10000; i++) {
    sw.writeRow(i, [`User_${i - 1}`, i * 1.5]);
}

// Optionally add merge cells
sw.addMergeCell('A1:B1');

// Apply the stream writer to the workbook
wb.applyStreamWriter(sw);

wb.save('large_file.xlsx');
```

#### StreamWriter API Summary

| Method                | Description                                     |
|-----------------------|-------------------------------------------------|
| `set_col_width`       | Set width for a single column (1-based number)  |
| `set_col_width_range` | Set width for a range of columns (Rust only)    |
| `write_row`           | Write a row of values at the given row number   |
| `add_merge_cell`      | Add a merge cell reference (e.g., `"A1:C3"`)    |

---

### Document Properties

Set and read document metadata: core properties (title, author, etc.), application properties, and custom properties.

#### Rust

```rust
use sheetkit::{AppProperties, CustomPropertyValue, DocProperties, Workbook};

let mut wb = Workbook::new();

// Core document properties
wb.set_doc_props(DocProperties {
    title: Some("Annual Report".into()),
    creator: Some("SheetKit".into()),
    description: Some("Financial data for 2025".into()),
    ..Default::default()
});
let props = wb.get_doc_props();

// Application properties
wb.set_app_props(AppProperties {
    application: Some("SheetKit".into()),
    company: Some("Acme Corp".into()),
    ..Default::default()
});
let app_props = wb.get_app_props();

// Custom properties (string, integer, float, boolean, or datetime)
wb.set_custom_property("Project", CustomPropertyValue::String("SheetKit".into()));
wb.set_custom_property("Version", CustomPropertyValue::Int(1));
wb.set_custom_property("Released", CustomPropertyValue::Bool(false));

let val = wb.get_custom_property("Project");
let deleted = wb.delete_custom_property("Version");
```

#### TypeScript

```typescript
// Core document properties
wb.setDocProps({
    title: 'Annual Report',
    creator: 'SheetKit',
    description: 'Financial data for 2025',
});
const props = wb.getDocProps();

// Application properties
wb.setAppProps({
    application: 'SheetKit',
    company: 'Acme Corp',
});
const appProps = wb.getAppProps();

// Custom properties (string, number, or boolean)
wb.setCustomProperty('Project', 'SheetKit');
wb.setCustomProperty('Version', 1);
wb.setCustomProperty('Released', false);

const val = wb.getCustomProperty('Project');       // string | number | boolean | null
const deleted: boolean = wb.deleteCustomProperty('Version');
```

#### DocProperties Fields

| Field              | Type             | Description                     |
|--------------------|------------------|---------------------------------|
| `title`            | `Option<String>` | Document title                  |
| `subject`          | `Option<String>` | Document subject                |
| `creator`          | `Option<String>` | Author name                     |
| `keywords`         | `Option<String>` | Keywords for search             |
| `description`      | `Option<String>` | Document description            |
| `last_modified_by` | `Option<String>` | Last editor                     |
| `revision`         | `Option<String>` | Revision number                 |
| `created`          | `Option<String>` | Creation timestamp              |
| `modified`         | `Option<String>` | Last modification timestamp     |
| `category`         | `Option<String>` | Category                        |
| `content_status`   | `Option<String>` | Content status                  |

#### AppProperties Fields

| Field          | Type             | Description                     |
|----------------|------------------|---------------------------------|
| `application`  | `Option<String>` | Application name                |
| `doc_security` | `Option<u32>`    | Document security level         |
| `company`      | `Option<String>` | Company name                    |
| `app_version`  | `Option<String>` | Application version             |
| `manager`      | `Option<String>` | Manager name                    |
| `template`     | `Option<String>` | Template name                   |

---

### Workbook Protection

Protect the workbook structure to prevent users from adding, deleting, or renaming sheets. An optional password can be set (legacy Excel hash -- not cryptographically secure).

#### Rust

```rust
use sheetkit::{Workbook, WorkbookProtectionConfig};

let mut wb = Workbook::new();

// Protect the workbook
wb.protect_workbook(WorkbookProtectionConfig {
    password: Some("secret".into()),
    lock_structure: true,    // prevent sheet add/delete/rename
    lock_windows: false,     // allow window resize
    lock_revision: false,    // allow revision tracking changes
});

// Check if protected
let is_protected: bool = wb.is_workbook_protected();

// Remove protection
wb.unprotect_workbook();
```

#### TypeScript

```typescript
// Protect the workbook
wb.protectWorkbook({
    password: 'secret',
    lockStructure: true,
    lockWindows: false,
    lockRevision: false,
});

// Check if protected
const isProtected: boolean = wb.isWorkbookProtected();

// Remove protection
wb.unprotectWorkbook();
```

---

## Examples

Complete example projects demonstrating all features are available in the repository:

- **Rust**: `examples/rust/` -- a standalone Cargo project (`cargo run` from within the directory)
- **Node.js**: `examples/node/` -- a TypeScript project (build the native module first, then run with `npx tsx index.ts`)

Each example walks through every feature: creating a workbook, setting cell values, managing sheets, applying styles, adding charts and images, data validation, comments, auto-filter, streaming large datasets, document properties, and workbook protection.

---

## Utility Functions

SheetKit also exposes helper functions for working with cell references:

```rust
use sheetkit::utils::cell_ref;

// Convert cell name to (column, row) coordinates
let (col, row) = cell_ref::cell_name_to_coordinates("B3")?;  // (2, 3)

// Convert coordinates to cell name
let name = cell_ref::coordinates_to_cell_name(2, 3)?;  // "B3"

// Convert column name to number
let num = cell_ref::column_name_to_number("AA")?;  // 27

// Convert column number to name
let name = cell_ref::column_number_to_name(27)?;  // "AA"
```
