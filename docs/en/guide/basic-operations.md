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