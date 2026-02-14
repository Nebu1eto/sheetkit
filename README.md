<h1>
  <img src="./logo.svg" alt="SheetKit logo" width="40" height="40" align="absmiddle" />
  SheetKit
</h1>

[![crates.io](https://img.shields.io/crates/v/sheetkit?logo=rust)](https://crates.io/crates/sheetkit)
[![npm](https://img.shields.io/npm/v/%40sheetkit%2Fnode?logo=npm)](https://www.npmjs.com/package/@sheetkit/node)
[![Docs](https://img.shields.io/badge/docs-vitepress-059669)](https://nebu1eto.github.io/sheetkit/)
[![CI](https://github.com/Nebu1eto/sheetkit/actions/workflows/ci.yml/badge.svg)](https://github.com/Nebu1eto/sheetkit/actions/workflows/ci.yml)
[![License: MIT OR Apache-2.0](https://img.shields.io/badge/license-MIT%20OR%20Apache--2.0-blue.svg)](#license)

A Rust library for reading and writing Excel (.xlsx) files, with Node.js bindings via napi-rs.

한국어 버전은 [README.ko.md](README.ko.md)를 참조하세요.

> **Warning**: SheetKit is experimental. APIs may change without notice. This project is under active development.

Quick links:
- [Docs](https://nebu1eto.github.io/sheetkit/)
- [Rust crate (crates.io)](https://crates.io/crates/sheetkit)
- [npm package (@sheetkit/node)](https://www.npmjs.com/package/@sheetkit/node)
- [Repository (GitHub)](https://github.com/Nebu1eto/sheetkit)

## Features

- Read/write .xlsx files
- Rust core + Node.js bindings via napi-rs (also works with Deno and Bun)
- Cell operations (string, number, boolean, date, formula)
- Sheet management (create, delete, rename, copy, active sheet)
- Row/column operations (insert, delete, resize, hide, outline)
- Style system (font, fill, border, alignment, number format, protection)
- 43 chart types with 3D support
- Images (11 formats: PNG, JPEG, GIF, BMP, ICO, TIFF, SVG, EMF, EMZ, WMF, WMZ)
- Conditional formatting (17 rule types)
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
import { Workbook } from "@sheetkit/node";

const wb = new Workbook();
wb.setCellValue("Sheet1", "A1", "Hello");
wb.setCellValue("Sheet1", "B1", 42);
wb.save("output.xlsx");
```

## Installation

**Rust** -- use `cargo add` (recommended):

```bash
cargo add sheetkit
```

Or add to your `Cargo.toml`:

```toml
[dependencies]
sheetkit = "0.5.0"
```

[View on crates.io](https://crates.io/crates/sheetkit)

**Node.js:**

```bash
# npm
npm install @sheetkit/node

# yarn
yarn add @sheetkit/node

# pnpm
pnpm add @sheetkit/node
```

[View on npm](https://www.npmjs.com/package/@sheetkit/node)

### Deno / Bun

SheetKit's Node.js bindings use [napi-rs](https://napi.rs/), which is compatible with other JavaScript runtimes that support Node-API:

- **Deno**: Supports napi-rs native addons via the [`--allow-ffi`](https://docs.deno.com/runtime/fundamentals/security/#ffi-(foreign-function-interface)) permission flag.
- **Bun**: Supports Node-API natively. Most napi-rs modules [work out of the box](https://bun.com/docs/runtime/node-api).

## Documentation

**[Documentation Site](https://nebu1eto.github.io/sheetkit/)** | [Korean (한국어)](https://nebu1eto.github.io/sheetkit/ko/)

See the [docs/](docs/) folder for documentation source files.

## Acknowledgements

SheetKit is heavily inspired by the implementation of [Excelize](https://github.com/qax-os/excelize), the Go library for Excel files. This project was built for the Rust and TypeScript projects, and we have great respect and admiration for the work done by the Excelize team and its contributors.

## License

MIT OR Apache-2.0
