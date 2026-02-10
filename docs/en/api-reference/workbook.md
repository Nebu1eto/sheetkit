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
import { Workbook } from "@sheetkit/node";

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
const wb = await Workbook.open("report.xlsx");
```

> Note (Node.js): `Workbook.open(path)` is async and returns `Promise<Workbook>`. Use `Workbook.openSync(path)` for synchronous behavior.

### `wb.save(path)` / `wb.save(path)`

Save the workbook to a `.xlsx` file on disk. Overwrites the file if it already exists.

**Rust:**

```rust
wb.save("output.xlsx")?;
```

**TypeScript:**

```typescript
await wb.save("output.xlsx");
```

> Note (Node.js): `wb.save(path)` is async and returns `Promise<void>`. Use `wb.saveSync(path)` for synchronous behavior.

### `Workbook::open_from_buffer(data)` / `Workbook.openBufferSync(data)`

Open a workbook from an in-memory byte buffer instead of a file path. Useful for processing uploaded files or data received over the network.

**Rust:**

```rust
let data: Vec<u8> = std::fs::read("report.xlsx")?;
let wb = Workbook::open_from_buffer(&data)?;
```

**TypeScript:**

```typescript
// Sync
const data: Buffer = fs.readFileSync("report.xlsx");
const wb = Workbook.openBufferSync(data);

// Async
const wb2 = await Workbook.openBuffer(data);
```

> Note (Node.js): `Workbook.openBuffer(data)` is async and returns `Promise<Workbook>`. Use `Workbook.openBufferSync(data)` for synchronous behavior.

### `wb.save_to_buffer()` / `wb.writeBufferSync()`

Serialize the workbook to an in-memory byte buffer without writing to disk. Useful for sending files as HTTP responses or passing data between services.

**Rust:**

```rust
let buf: Vec<u8> = wb.save_to_buffer()?;
// buf contains valid .xlsx data
```

**TypeScript:**

```typescript
// Sync
const buf: Buffer = wb.writeBufferSync();

// Async
const buf2: Buffer = await wb.writeBuffer();
```

> Note (Node.js): `wb.writeBuffer()` is async and returns `Promise<Buffer>`. Use `wb.writeBufferSync()` for synchronous behavior.

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

### `wb.format()` / `wb.format`

Get the detected or assigned workbook format. The format determines which OOXML content type is used for `xl/workbook.xml` on save. When a workbook is opened from a file, the format is auto-detected from the content types in the package.

**Rust:**

```rust
use sheetkit::{Workbook, WorkbookFormat};

let wb = Workbook::open("report.xlsm")?;
let fmt: WorkbookFormat = wb.format();
assert_eq!(fmt, WorkbookFormat::Xlsm);
```

**TypeScript:**

```typescript
const wb = await Workbook.open("report.xlsm");
const fmt: string = wb.format; // "xlsm"
```

### `wb.set_format(format)` / `wb.format = ...`

Set the workbook format explicitly. This overrides the auto-detected format and controls the content type emitted on save.

**Rust:**

```rust
use sheetkit::{Workbook, WorkbookFormat};

let mut wb = Workbook::new();
wb.set_format(WorkbookFormat::Xlsm);
wb.save("macros.xlsm")?;
```

**TypeScript:**

```typescript
const wb = new Workbook();
wb.format = "xlsm";
await wb.save("macros.xlsm");
```

### WorkbookFormat

| Rust | TypeScript | Extension | Description |
|---|---|---|---|
| `WorkbookFormat::Xlsx` | `"xlsx"` | `.xlsx` | Standard spreadsheet (default) |
| `WorkbookFormat::Xlsm` | `"xlsm"` | `.xlsm` | Macro-enabled spreadsheet |
| `WorkbookFormat::Xltx` | `"xltx"` | `.xltx` | Template |
| `WorkbookFormat::Xltm` | `"xltm"` | `.xltm` | Macro-enabled template |
| `WorkbookFormat::Xlam` | `"xlam"` | `.xlam` | Macro-enabled add-in |

### Extension-Based Save

When saving, the target format is automatically inferred from the file extension. If the file path ends with a recognized extension (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xlam`), the workbook format is updated before writing. Unrecognized extensions return an error.

**Rust:**

```rust
let mut wb = Workbook::new();
// Format is inferred from the ".xlsm" extension
wb.save("output.xlsm")?;
assert_eq!(wb.format(), WorkbookFormat::Xlsm);
```

**TypeScript:**

```typescript
const wb = new Workbook();
await wb.save("output.xlsm"); // format set to xlsm automatically
```

When using `save_to_buffer()` / `writeBufferSync()`, the stored format is used as-is since there is no file extension to infer from. Set the format explicitly with `set_format()` before calling buffer save.

### VBA Preservation

Macro-enabled workbooks (`.xlsm`, `.xltm`) contain a VBA project blob (`xl/vbaProject.bin`). When such a file is opened and re-saved, the VBA project is preserved transparently. No additional API calls are needed -- the binary blob is stored in memory and written back on save.

```rust
// Open a macro-enabled file, modify data, save -- VBA macros are preserved
let mut wb = Workbook::open("with_macros.xlsm")?;
wb.set_cell_value("Sheet1", "A1", CellValue::String("Updated".into()))?;
wb.save("with_macros.xlsm")?;
```

```typescript
const wb = await Workbook.open("with_macros.xlsm");
wb.setCellValue("Sheet1", "A1", "Updated");
await wb.save("with_macros.xlsm"); // VBA preserved
```

### `OpenOptions` / `JsOpenOptions`

Options for controlling how a workbook is opened and parsed. All fields are optional and default to no limit / parse everything.

| Field | Rust type | TypeScript type | Description |
|---|---|---|---|
| `sheet_rows` / `sheetRows` | `Option<u32>` | `number?` | Maximum number of rows to read per sheet. Rows beyond this limit are discarded. |
| `sheets` | `Option<Vec<String>>` | `string[]?` | Only parse sheets whose names are in this list. Unselected sheets exist in the workbook but contain no data. |
| `max_unzip_size` / `maxUnzipSize` | `Option<u64>` | `number?` | Maximum total decompressed size of the ZIP archive in bytes. Prevents zip bombs. |
| `max_zip_entries` / `maxZipEntries` | `Option<usize>` | `number?` | Maximum number of entries in the ZIP archive. Prevents zip bombs. |

### `Workbook::open_with_options(path, options)` / `Workbook.open(path, options?)`

Open a `.xlsx` file with custom parsing options. When no options are provided, behaves identically to `Workbook::open`.

**Rust:**

```rust
use sheetkit::{Workbook, OpenOptions};

// Read only the first 100 rows of the "Sales" sheet
let opts = OpenOptions::new()
    .sheet_rows(100)
    .sheets(vec!["Sales".to_string()]);
let wb = Workbook::open_with_options("report.xlsx", &opts)?;
```

**TypeScript:**

```typescript
// Read only the first 100 rows of the "Sales" sheet
const wb = Workbook.openSync("report.xlsx", {
  sheetRows: 100,
  sheets: ["Sales"],
});

// With ZIP safety limits
const wb2 = await Workbook.open("untrusted.xlsx", {
  maxUnzipSize: 500_000_000,  // 500 MB
  maxZipEntries: 5000,
});
```

### `Workbook::open_from_buffer_with_options(data, options)` / `Workbook.openBufferSync(data, options?)`

Open a workbook from an in-memory buffer with custom parsing options.

**Rust:**

```rust
let data = std::fs::read("report.xlsx")?;
let opts = OpenOptions::new().sheet_rows(50);
let wb = Workbook::open_from_buffer_with_options(&data, &opts)?;
```

**TypeScript:**

```typescript
const data = fs.readFileSync("report.xlsx");
const wb = Workbook.openBufferSync(data, { sheetRows: 50 });
```

> Note (Node.js): The options parameter is optional in all open methods. Omitting it preserves backward compatibility.

---
