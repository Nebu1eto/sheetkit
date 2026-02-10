# Getting Started with SheetKit

SheetKit is a high-performance SpreadsheetML library for Rust and TypeScript.
The Rust core handles all Excel (.xlsx) processing, and napi-rs bindings bring the same performance to TypeScript with minimal overhead.

## Why SheetKit?

- **Native Performance**: Up to 10x faster than JavaScript-only libraries for large datasets
- **Memory Efficient**: Buffer-based FFI transfer reduces memory usage by up to 96%
- **Type Safe**: Strongly typed APIs for both Rust and TypeScript
- **Complete**: 110+ formula functions, 43 chart types, streaming writer, and more

## Installation

### Rust Library

Add SheetKit using `cargo add` (recommended):

```bash
cargo add sheetkit
```

Or manually add to your `Cargo.toml`:

```toml
[dependencies]
sheetkit = { version = "0.3" }
```

For encryption support:

```bash
cargo add sheetkit --features encryption
```

[View on crates.io](https://crates.io/crates/sheetkit)

### Node.js Library

Install via your preferred package manager:

```bash
# npm
npm install @sheetkit/node

# yarn
yarn add @sheetkit/node

# pnpm
pnpm add @sheetkit/node
```

Note: A Rust toolchain is required for native compilation during installation.

[View on npm](https://www.npmjs.com/package/@sheetkit/node)

### CLI Tool

For command-line operations (sheet inspection, data conversion, etc.):

```bash
cargo install sheetkit --features cli
```

See the [CLI Guide](./guide/cli.md) for usage examples.

## Quick Start

### Creating a Workbook and Writing Cells

**Rust**

```rust
use sheetkit::{CellValue, Workbook};

fn main() -> sheetkit::Result<()> {
    let mut wb = Workbook::new();

    // Write different value types
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Name".into()))?;
    wb.set_cell_value("Sheet1", "B1", CellValue::Number(42.0))?;
    wb.set_cell_value("Sheet1", "C1", CellValue::Bool(true))?;
    wb.set_cell_value("Sheet1", "D1", CellValue::Empty)?;

    // Read a cell value
    let val = wb.get_cell_value("Sheet1", "A1")?;
    println!("A1 = {:?}", val); // CellValue::String("Name")

    // Save to file
    wb.save("output.xlsx")?;
    Ok(())
}
```

**TypeScript**

```typescript
import { Workbook } from "@sheetkit/node";

const wb = new Workbook();

// Write different value types
wb.setCellValue("Sheet1", "A1", "Name");
wb.setCellValue("Sheet1", "B1", 42);
wb.setCellValue("Sheet1", "C1", true);
wb.setCellValue("Sheet1", "D1", null); // clear/empty

// Read a cell value
const val = wb.getCellValue("Sheet1", "A1");
console.log("A1 =", val); // "Name"

// Save to file
await wb.save("output.xlsx");
```

### Opening an Existing File

**Rust**

```rust
use sheetkit::Workbook;

fn main() -> sheetkit::Result<()> {
    let wb = Workbook::open("input.xlsx")?;

    // List all sheet names
    let names = wb.sheet_names();
    println!("Sheets: {:?}", names);

    // Read a cell from the first sheet
    let val = wb.get_cell_value(&names[0], "A1")?;
    println!("A1 = {:?}", val);

    Ok(())
}
```

**TypeScript**

```typescript
import { Workbook } from "@sheetkit/node";

const wb = await Workbook.open("input.xlsx");

// List all sheet names
console.log("Sheets:", wb.sheetNames);

// Read a cell from the first sheet
const val = wb.getCellValue(wb.sheetNames[0], "A1");
console.log("A1 =", val);
```

## Core Concepts

### CellValue Types

SheetKit uses a typed cell value model. Every cell holds one of these variants:

| Type    | Rust                                            | TypeScript               |
| ------- | ----------------------------------------------- | ------------------------ |
| String  | `CellValue::String(String)`                     | `string`                 |
| Number  | `CellValue::Number(f64)`                        | `number`                 |
| Bool    | `CellValue::Bool(bool)`                         | `boolean`                |
| Empty   | `CellValue::Empty`                              | `null`                   |
| Date    | `CellValue::Date(f64)`                          | `{ type: 'date', serial: number, iso?: string }` |
| Formula | `CellValue::Formula { expr, result }`           | *(set via formula eval)* |
| Error   | `CellValue::Error(String)`                      | *(read-only)*            |

### Date Values

Dates are stored as Excel serial numbers (days since 1900-01-01). The integer part represents the date, and the fractional part represents the time of day (for example, 0.5 = noon).

To display a date correctly in Excel, you must apply a date number format style to the cell.

**Rust**

```rust
use sheetkit::{CellValue, Style, NumFmtStyle, Workbook};

let mut wb = Workbook::new();

// Excel serial 45292 = 2024-01-01
wb.set_cell_value("Sheet1", "A1", CellValue::Date(45292.0))?;

// Apply a date format so Excel renders it as a date
let style_id = wb.add_style(&Style {
    num_fmt: Some(NumFmtStyle { num_fmt_id: Some(14), custom_format: None }),
    ..Default::default()
})?;
wb.set_cell_style("Sheet1", "A1", style_id)?;
```

**TypeScript**

```typescript
const wb = new Workbook();

// Excel serial 45292 = 2024-01-01
wb.setCellValue("Sheet1", "A1", { type: "date", serial: 45292 });

// Apply a date format so Excel renders it as a date
const styleId = wb.addStyle({ numFmtId: 14 });
wb.setCellStyle("Sheet1", "A1", styleId);
```

When reading a date cell back, the `DateValue` object includes an optional `iso` field with the ISO 8601 string representation (e.g., `"2024-01-01"` or `"2024-01-01T12:00:00"`).

### Cell References

Cell references use A1-style notation: a column letter followed by a 1-based row number.

- `"A1"` -- column A, row 1
- `"B2"` -- column B, row 2
- `"AA100"` -- column AA (27th column), row 100

### Sheet Names

Sheet names are case-sensitive strings. A new workbook starts with a single sheet named `"Sheet1"`.

### 1-Based vs 0-Based Indexing

- **Rows**: 1-based (row 1 is the first row)
- **Column numbers**: 1-based (column 1 = "A")
- **Sheet indices**: 0-based (the first sheet has index 0)

### Style System

Styles follow a register-then-apply pattern:

1. Define a `Style` struct/object with font, fill, border, alignment, and number format options.
2. Register it with `add_style` / `addStyle` to get a numeric style ID.
3. Apply the style ID to cells, rows, or columns.

Style deduplication is automatic. Registering two identical styles returns the same ID.

## Working with Styles

**Rust**

```rust
use sheetkit::{
    CellValue, FillStyle, FontStyle, PatternType, Style, StyleColor, Workbook,
};

let mut wb = Workbook::new();
wb.set_cell_value("Sheet1", "A1", CellValue::String("Styled".into()))?;

let style_id = wb.add_style(&Style {
    font: Some(FontStyle {
        bold: true,
        size: Some(14.0),
        color: Some(StyleColor::Rgb("#FFFFFF".into())),
        ..Default::default()
    }),
    fill: Some(FillStyle {
        pattern: PatternType::Solid,
        fg_color: Some(StyleColor::Rgb("#4472C4".into())),
        bg_color: None,
    }),
    ..Default::default()
})?;

wb.set_cell_style("Sheet1", "A1", style_id)?;
wb.save("styled.xlsx")?;
```

**TypeScript**

```typescript
const wb = new Workbook();
wb.setCellValue("Sheet1", "A1", "Styled");

const styleId = wb.addStyle({
  font: { bold: true, size: 14, color: "#FFFFFF" },
  fill: { pattern: "solid", fgColor: "#4472C4" },
});

wb.setCellStyle("Sheet1", "A1", styleId);
await wb.save("styled.xlsx");
```

## Working with Charts

Add a chart by specifying the anchor range (top-left and bottom-right cells), chart type, and data series.

**Rust**

```rust
use sheetkit::{CellValue, ChartConfig, ChartSeries, ChartType, Workbook};

let mut wb = Workbook::new();

// Prepare data
wb.set_cell_value("Sheet1", "A1", CellValue::String("Q1".into()))?;
wb.set_cell_value("Sheet1", "A2", CellValue::String("Q2".into()))?;
wb.set_cell_value("Sheet1", "A3", CellValue::String("Q3".into()))?;
wb.set_cell_value("Sheet1", "B1", CellValue::Number(1500.0))?;
wb.set_cell_value("Sheet1", "B2", CellValue::Number(2300.0))?;
wb.set_cell_value("Sheet1", "B3", CellValue::Number(1800.0))?;

// Add a column chart
wb.add_chart(
    "Sheet1",
    "D1",   // top-left anchor
    "K15",  // bottom-right anchor
    &ChartConfig {
        chart_type: ChartType::Col,
        title: Some("Quarterly Revenue".into()),
        series: vec![ChartSeries {
            name: "Revenue".into(),
            categories: "Sheet1!$A$1:$A$3".into(),
            values: "Sheet1!$B$1:$B$3".into(),
            x_values: None,
            bubble_sizes: None,
        }],
        show_legend: true,
        view_3d: None,
    },
)?;

wb.save("chart.xlsx")?;
```

**TypeScript**

```typescript
const wb = new Workbook();

// Prepare data
wb.setCellValue("Sheet1", "A1", "Q1");
wb.setCellValue("Sheet1", "A2", "Q2");
wb.setCellValue("Sheet1", "A3", "Q3");
wb.setCellValue("Sheet1", "B1", 1500);
wb.setCellValue("Sheet1", "B2", 2300);
wb.setCellValue("Sheet1", "B3", 1800);

// Add a column chart
wb.addChart("Sheet1", "D1", "K15", {
  chartType: "col",
  title: "Quarterly Revenue",
  series: [
    {
      name: "Revenue",
      categories: "Sheet1!$A$1:$A$3",
      values: "Sheet1!$B$1:$B$3",
    },
  ],
  showLegend: true,
});

await wb.save("chart.xlsx");
```

## StreamWriter for Large Files

The `StreamWriter` writes rows sequentially to an internal buffer without building the entire worksheet in memory. Rows must be written in ascending order.

**Rust**

```rust
use sheetkit::{CellValue, Workbook};

let mut wb = Workbook::new();
let mut sw = wb.new_stream_writer("LargeSheet")?;

// Set column widths
sw.set_col_width(1, 20.0)?;
sw.set_col_width(2, 15.0)?;

// Write header
sw.write_row(1, &[
    CellValue::String("Item".into()),
    CellValue::String("Value".into()),
])?;

// Write 10,000 data rows
for i in 2..=10_001 {
    sw.write_row(i, &[
        CellValue::String(format!("Item_{}", i - 1)),
        CellValue::Number(i as f64 * 1.5),
    ])?;
}

// Apply the stream writer output to the workbook
wb.apply_stream_writer(sw)?;
wb.save("large.xlsx")?;
```

**TypeScript**

```typescript
const wb = new Workbook();
const sw = wb.newStreamWriter("LargeSheet");

// Set column widths
sw.setColWidth(1, 20);
sw.setColWidth(2, 15);

// Write header
sw.writeRow(1, ["Item", "Value"]);

// Write 10,000 data rows
for (let i = 2; i <= 10_001; i++) {
  sw.writeRow(i, [`Item_${i - 1}`, i * 1.5]);
}

// Apply the stream writer output to the workbook
wb.applyStreamWriter(sw);
await wb.save("large.xlsx");
```

## Working with Encrypted Files

SheetKit can read and write password-protected .xlsx files. Enable the `encryption` feature in Rust; Node.js bindings always include encryption support.

**Rust**

```rust
use sheetkit::Workbook;

// Save with password
let wb = Workbook::new();
wb.save_with_password("encrypted.xlsx", "secret")?;

// Open with password
let wb2 = Workbook::open_with_password("encrypted.xlsx", "secret")?;
```

**TypeScript**

```typescript
import { Workbook } from "@sheetkit/node";

// Save with password
const wb = new Workbook();
wb.saveWithPassword("encrypted.xlsx", "secret");

// Open with password
const wb2 = Workbook.openWithPasswordSync("encrypted.xlsx", "secret");
```

## Next Steps

- [API Reference](./api-reference/index.md) -- Full documentation for every method and type.
- [Architecture](./architecture.md) -- Internal design and crate structure.
- [Contributing](./contributing.md) -- Development setup and contribution guidelines.
