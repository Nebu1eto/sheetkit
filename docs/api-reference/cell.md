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

### `get_cell_formatted_value` / `getCellFormattedValue`

Return the display text for a cell by applying its number format. For numeric cells with a format style (date, percentage, thousands separator, etc.), the raw value is rendered through the format code. String cells return their text as-is. Empty cells return an empty string.

**Rust:**

```rust
use sheetkit::Workbook;

let wb = Workbook::open("data.xlsx")?;
let formatted: String = wb.get_cell_formatted_value("Sheet1", "A1")?;
// e.g., "1,234.50" for a number with #,##0.00 format
// e.g., "2024-12-31" for a date serial with yyyy-mm-dd format
```

**TypeScript:**

```typescript
const wb = await Workbook.open("data.xlsx");
const text = wb.getCellFormattedValue("Sheet1", "A1");
// "1,234.50", "2024-12-31", "85.00%", etc.
```

### `format_number` / `formatNumber`

Standalone utility that formats a numeric value using an Excel format code string. This does not require a workbook instance.

**Rust:**

```rust
use sheetkit::format_number;

assert_eq!(format_number(1234.5, "#,##0.00"), "1,234.50");
assert_eq!(format_number(0.85, "0.00%"), "85.00%");
assert_eq!(format_number(45657.0, "yyyy-mm-dd"), "2024-12-31");
assert_eq!(format_number(0.5, "h:mm AM/PM"), "12:00 PM");
```

**TypeScript:**

```typescript
import { formatNumber } from "sheetkit";

formatNumber(1234.5, "#,##0.00");   // "1,234.50"
formatNumber(0.85, "0.00%");        // "85.00%"
formatNumber(45657, "yyyy-mm-dd");  // "2024-12-31"
formatNumber(0.5, "h:mm AM/PM");    // "12:00 PM"
```

Supported format features:

| Feature | Examples | Description |
|---|---|---|
| General | `General` | Default display format |
| Integer / Decimal | `0`, `0.00`, `#,##0.00` | Digit placeholders with optional thousands separator |
| Percentage | `0%`, `0.00%` | Multiplies by 100 and appends % |
| Scientific | `0.00E+00` | Exponential notation |
| Date | `m/d/yyyy`, `yyyy-mm-dd`, `d-mmm-yy` | Date from Excel serial number |
| Time | `h:mm`, `h:mm:ss`, `h:mm AM/PM` | Time from fractional day |
| Date + Time | `m/d/yyyy h:mm` | Combined date and time |
| Fractions | `# ?/?`, `# ??/??` | Fractional representation |
| Multi-section | `pos;neg;zero;text` | Up to 4 sections separated by `;` |
| Color codes | `[Red]0.00` | Color tags are parsed and stripped |
| Literal text | `"text"`, `\x` | Quoted strings and escaped characters |

### `builtin_format_code` / `builtinFormatCode`

Look up the format code string for a built-in number format ID (0-49). Returns `None`/`null` for unrecognized IDs.

**Rust:**

```rust
use sheetkit::builtin_format_code;

assert_eq!(builtin_format_code(0), Some("General"));
assert_eq!(builtin_format_code(14), Some("m/d/yyyy"));
assert_eq!(builtin_format_code(100), None);
```

**TypeScript:**

```typescript
import { builtinFormatCode } from "sheetkit";

builtinFormatCode(0);   // "General"
builtinFormatCode(14);  // "m/d/yyyy"
builtinFormatCode(100); // null
```

### `get_occupied_cells(sheet)` (Rust only)

Return a list of `(col, row)` coordinate pairs for every non-empty cell in a sheet. Both values are 1-based. Useful for iterating only over cells that contain data without scanning the entire grid.

```rust
let cells = wb.get_occupied_cells("Sheet1")?;
for (col, row) in &cells {
    println!("Cell at col {}, row {}", col, row);
}
```

---
