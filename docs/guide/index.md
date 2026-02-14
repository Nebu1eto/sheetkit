# SheetKit User Guide

SheetKit is a high-performance SpreadsheetML library for Rust and TypeScript. The Rust core handles all Excel (.xlsx) processing, and napi-rs bindings bring the same performance to TypeScript with minimal overhead.

---

## Table of Contents

- [Basic Operations](./basic-operations.md)
  - Installation
  - Quick Start
  - Workbook I/O
  - Cell Operations
  - Workbook Format and VBA Preservation
- [Styling](./styling.md)
  - Styles
- [Data Features](./data-features.md)
  - Sheet Management
  - Row and Column Operations
  - Row/Column Iterators
  - Row/Column Outline Levels and Styles
  - Charts
  - Images
  - Merge Cells
  - Hyperlinks
  - Conditional Formatting
  - Tables
  - Data Conversion Utilities
- [Advanced](./advanced.md)
  - Freeze/Split Panes
  - Page Layout
  - Data Validation
  - Comments
  - Auto-Filter
  - Formula Evaluation
  - Pivot Tables
  - StreamWriter
  - Document Properties
  - Workbook Protection
  - Sparklines
  - Defined Names
  - Sheet Protection
  - Sheet View Options
  - Sheet Visibility
  - Examples
  - Utility Functions
  - Theme Colors
  - Rich Text
  - File Encryption
- [Migration: Async-First Lazy-Open](./migration-async-first.md)
  - ReadMode mapping
  - OpenOptions changes
  - Streaming reader migration
  - Raw buffer V2
  - Copy-on-write save

---

## Getting Started

Start with [Basic Operations](./basic-operations.md) to learn how to create and manipulate workbooks, then continue with [Styling](./styling.md) and [Data Features](./data-features.md) for advanced capabilities.

For a full list of API methods, see the [API Reference](../api-reference/index.md).

---

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

> The Node.js package is a native addon built with napi-rs. A Rust build toolchain (rustc, cargo) is required to compile the native module during installation.

---

## Quick Example

### Rust

```rust
use sheetkit::{CellValue, Workbook};

fn main() -> sheetkit::Result<()> {
    let mut wb = Workbook::new();
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))?;
    wb.save("output.xlsx")?;
    Ok(())
}
```

### TypeScript / Node.js

```typescript
import { Workbook } from '@sheetkit/node';

const wb = new Workbook();
wb.setCellValue('Sheet1', 'A1', 'Hello');
await wb.save('output.xlsx');
```

---

## Next Steps

- Learn [Basic Operations](./basic-operations.md) with detailed examples
- Apply [Styling](./styling.md) to make your spreadsheets look professional
- Add [Data Features](./data-features.md) like charts, validation, and comments
- See [Advanced](./advanced.md) features for complex workbooks
- Check out complete example projects in `examples/rust/` and `examples/node/` in the repository
