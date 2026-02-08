# SheetKit

A Rust library for reading and writing Excel (.xlsx) files, with Node.js bindings via napi-rs.

한국어 버전은 [README.ko.md](README.ko.md)를 참조하세요.

> **Warning**: SheetKit is experimental. APIs may change without notice. This project is under active development.

## Features

- Read/write .xlsx files
- Rust core + Node.js bindings (napi-rs)
- Cell operations (string, number, boolean, date, formula)
- Sheet management (create, delete, rename, copy, active sheet)
- Row/column operations (insert, delete, resize, hide, outline)
- Style system (font, fill, border, alignment, number format, protection)
- 41 chart types with 3D support
- Images (PNG, JPEG, GIF)
- Conditional formatting (18 rule types)
- Data validation, comments, auto-filter
- Formula evaluation (110+ functions)
- Streaming writer for large datasets
- Merge cells, hyperlinks, freeze/split panes
- Page layout and print settings
- Document properties, workbook/sheet protection
- Pivot tables
- Defined names (named ranges)

## Quick Start

**Rust:**

```rust
use sheetkit::{CellValue, Workbook};

fn main() -> sheetkit::Result<()> {
    let mut wb = Workbook::new();
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))?;
    wb.set_cell_value("Sheet1", "B1", CellValue::Number(42.0))?;
    wb.save("output.xlsx")?;
    Ok(())
}
```

**TypeScript:**

```typescript
import { Workbook } from "sheetkit";

const wb = new Workbook();
wb.setCellValue("Sheet1", "A1", "Hello");
wb.setCellValue("Sheet1", "B1", 42);
wb.save("output.xlsx");
```

## Installation

**Rust** -- add to your `Cargo.toml`:

```toml
[dependencies]
sheetkit = "0.1"
```

**Node.js:**

```bash
npm install sheetkit
```

## Documentation

See the [docs/en/](docs/en/) folder for detailed documentation.

## Acknowledgements

SheetKit is a Rust rewrite inspired by [Excelize](https://github.com/qax-os/excelize), the Go library for Excel files.

## License

MIT OR Apache-2.0
