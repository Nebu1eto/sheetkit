## Installation

### Rust

Add `sheetkit` to your `Cargo.toml`:

```toml
[dependencies]
sheetkit = "0.5.0"
```

### Node.js

```bash
npm install @sheetkit/node
```

> The Node.js package is a native addon built with napi-rs. On supported OS/architecture pairs, prebuilt binaries are used and a Rust toolchain is not required. A Rust toolchain is needed only when building from source or when no prebuilt binary is available.

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
import { Workbook } from '@sheetkit/node';

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
await wb.save('output.xlsx');

// Open an existing file
const wb2 = await Workbook.open('output.xlsx');
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
import { Workbook } from '@sheetkit/node';

// Create a new empty workbook with a single "Sheet1"
const wb = new Workbook();

// Open an existing .xlsx file
const wb2 = await Workbook.open('input.xlsx');

// Save the workbook to a .xlsx file
await wb.save('output.xlsx');

// Get the names of all sheets
const names: string[] = wb.sheetNames;
```

#### Buffer I/O

Read from and write to in-memory buffers without touching the file system.

**Rust:**

```rust
// Save to buffer
let buf: Vec<u8> = wb.save_to_buffer()?;

// Open from buffer
let wb2 = Workbook::open_from_buffer(&buf)?;
```

**TypeScript:**

```typescript
// Save to buffer
const buf: Buffer = wb.writeBufferSync();

// Open from buffer
const wb2 = Workbook.openBufferSync(buf);

// Async versions
const buf2: Buffer = await wb.writeBuffer();
const wb3 = await Workbook.openBuffer(buf2);
```

---

### Workbook Format and VBA Preservation

SheetKit supports multiple Excel file formats beyond the standard `.xlsx`. When opening a file, the format is automatically detected from the package content types. When saving, the format is inferred from the file extension.

#### Supported Formats

| Extension | Description |
|-----------|-------------|
| `.xlsx` | Standard spreadsheet (default) |
| `.xlsm` | Macro-enabled spreadsheet |
| `.xltx` | Template |
| `.xltm` | Macro-enabled template |
| `.xlam` | Macro-enabled add-in |

#### Rust

```rust
use sheetkit::{Workbook, WorkbookFormat};

// Format is auto-detected on open
let wb = Workbook::open("macros.xlsm")?;
assert_eq!(wb.format(), WorkbookFormat::Xlsm);

// Format is inferred from extension on save
let mut wb2 = Workbook::new();
wb2.save("template.xltx")?;  // saved as template format

// Explicit format control
let mut wb3 = Workbook::new();
wb3.set_format(WorkbookFormat::Xlsm);
wb3.save_to_buffer()?;  // buffer uses xlsm content type
```

#### TypeScript

```typescript
// Format is auto-detected on open
const wb = await Workbook.open("macros.xlsm");
console.log(wb.format);  // "xlsm"

// Format is inferred from extension on save
const wb2 = new Workbook();
await wb2.save("template.xltx");  // saved as template format

// Explicit format control
const wb3 = new Workbook();
wb3.format = "xlsm";
const buf = wb3.writeBufferSync();  // buffer uses xlsm content type
```

#### VBA Preservation

Macro-enabled files (`.xlsm`, `.xltm`) preserve their VBA project through open/save round-trips. No user code is needed -- the VBA binary blob is kept in memory and written back on save automatically.

```typescript
const wb = await Workbook.open("with_macros.xlsm");
wb.setCellValue("Sheet1", "A1", "Updated");
await wb.save("with_macros.xlsm");  // macros are preserved
```

For full API details, see the [API Reference](../api-reference/workbook.md).

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