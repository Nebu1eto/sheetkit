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
