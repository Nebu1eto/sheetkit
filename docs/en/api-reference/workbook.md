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

---
