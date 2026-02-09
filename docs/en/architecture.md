# SheetKit Architecture

## 1. Overview

SheetKit is a Rust rewrite of the Go Excelize library for reading and writing Excel (.xlsx) files. The .xlsx format is OOXML (Office Open XML), which is a ZIP archive containing XML parts. SheetKit reads the ZIP, deserializes each XML part into typed Rust structs, exposes a high-level API for manipulation, and serializes everything back into a valid .xlsx file on save.

## 2. Crate Structure

```
crates/
  sheetkit-xml/     # XML schema types (serde-based)
  sheetkit-core/    # Business logic
  sheetkit/         # Public facade (re-exports)
packages/
  sheetkit/         # Node.js bindings (napi-rs)
```

### sheetkit-xml

Low-level XML data structures mapping to OOXML schemas. Each file corresponds to a major OOXML part:

| File | OOXML Part |
|---|---|
| `worksheet.rs` | Worksheet (sheet data, merge cells, conditional formatting, validations) |
| `shared_strings.rs` | SharedStrings (SST) |
| `styles.rs` | Stylesheet (fonts, fills, borders, number formats, XF records, DXF records) |
| `workbook.rs` | Workbook (sheets, defined names, calc properties, pivot caches) |
| `content_types.rs` | `[Content_Types].xml` |
| `relationships.rs` | `.rels` relationship files |
| `chart.rs` | Chart definitions (DrawingML charts) |
| `drawing.rs` | DrawingML (anchors, shapes, image references) |
| `comments.rs` | Comment data and authors |
| `doc_props.rs` | Core, App, and Custom document properties |
| `pivot_table.rs` | Pivot table definitions |
| `pivot_cache.rs` | Pivot cache definitions and records |
| `namespaces.rs` | OOXML namespace constants |

All types use `serde::Deserialize` and `serde::Serialize` derive macros with `quick-xml` attributes for XML element/attribute mapping.

### sheetkit-core

All business logic lives here. The central type is `Workbook` in `workbook.rs`, which owns the deserialized XML state and provides the public API.

**Core modules:**

| Module | Responsibility |
|---|---|
| `workbook.rs` | Opens ZIP, deserializes XML parts, manages mutable state, serializes and saves |
| `cell.rs` | `CellValue` enum (String, Number, Bool, Empty, Date, Formula, Error), date serial number conversion via chrono |
| `sst.rs` | Shared Strings Table runtime with HashMap for O(1) string deduplication |
| `sheet.rs` | Sheet management: create, delete, rename, copy, set active, freeze/split panes, sheet properties, sheet protection |
| `row.rs` | Row operations: insert, delete, duplicate, set height, visibility, outline level, row style, iterators |
| `col.rs` | Column operations: set width, visibility, insert, delete, outline level, column style |
| `style.rs` | Style system: font, fill, border, alignment, number format, cell protection. StyleBuilder API with automatic XF deduplication |
| `conditional.rs` | Conditional formatting: 17 rule types (cell value, color scale, data bar, top/bottom, above/below average, duplicates, blanks, errors, text matching, etc.) using DXF records |
| `chart.rs` | Chart creation for 43 chart types (bar, line, pie, area, scatter, radar, stock, surface, doughnut, combo, 3D variants). Manages DrawingML anchors and relationships |
| `image.rs` | Image embedding (11 formats: PNG, JPEG, GIF, BMP, ICO, TIFF, SVG, EMF, EMZ, WMF, WMZ) with two-cell anchors in DrawingML |
| `validation.rs` | Data validation rules (dropdown, whole number, decimal, text length, date, time, custom formula) |
| `comment.rs` | Cell comments using VML drawing format |
| `table.rs` | Table and auto-filter support |
| `hyperlink.rs` | Hyperlinks: external and email use worksheet .rels, internal (sheet-to-sheet) use location attribute only |
| `merge.rs` | Merge and unmerge cell ranges |
| `doc_props.rs` | Core (dc:, dcterms:, cp:), App, and Custom document properties. DC namespace requires manual quick-xml Writer/Reader |
| `protection.rs` | Workbook-level protection with legacy password hash |
| `page_layout.rs` | Page margins, page setup, print options, header/footer |
| `defined_names.rs` | Named ranges (workbook-scoped and sheet-scoped) |
| `pivot.rs` | Pivot tables: cache definition, cache records, table definition, workbook pivot cache collection |
| `stream.rs` | StreamWriter: forward-only XML writer for generating large files without holding the full sheet in memory. Supports SST merging, freeze panes, row options, column style/visibility/outline |
| `error.rs` | Error types using thiserror |

**Formula subsystem** (`formula/`):

| File | Responsibility |
|---|---|
| `parser.rs` | nom-based formula parser producing an AST. Handles operator precedence, cell references, range references, function calls, string/number/boolean literals |
| `ast.rs` | AST node types (BinaryOp, UnaryOp, FunctionCall, CellRef, RangeRef, Literal, etc.) |
| `eval.rs` | Formula evaluator. Uses `CellDataProvider` trait for workbook data access and `CellSnapshot` (HashMap) to avoid borrow checker issues. `calculate_all()` builds a dependency graph and uses Kahn's algorithm for topological sort |
| `functions/mod.rs` | Function dispatch table mapping function names to implementations |
| `functions/math.rs` | Math functions (SUM, AVERAGE, ABS, ROUND, etc.) |
| `functions/statistical.rs` | Statistical functions (COUNT, COUNTA, MAX, MIN, STDEV, etc.) |
| `functions/text.rs` | Text functions (CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, etc.) |
| `functions/logical.rs` | Logical functions (IF, AND, OR, NOT, IFERROR, etc.) |
| `functions/information.rs` | Information functions (ISBLANK, ISERROR, ISNUMBER, TYPE, etc.) |
| `functions/date_time.rs` | Date/time functions (DATE, TODAY, NOW, YEAR, MONTH, DAY, etc.) |
| `functions/lookup.rs` | Lookup functions (VLOOKUP, HLOOKUP, INDEX, MATCH, etc.) |

**Utilities** (`utils/`):

| File | Responsibility |
|---|---|
| `cell_ref.rs` | Cell reference parsing: "A1" to (row, col), column letter conversion, range parsing |
| `constants.rs` | Shared constants |

### sheetkit (facade)

Thin re-export crate. Its `lib.rs` contains `pub use sheetkit_core::*;` so that end users depend on `sheetkit` and get the full public API.

### sheetkit-node (packages/sheetkit)

Node.js bindings via napi-rs (v3, no compat-mode).

- `src/lib.rs` -- Single file containing all bindings. The `#[napi]` `Workbook` class wraps `sheetkit_core::workbook::Workbook` in an `inner` field. Methods delegate to `inner` and convert between Rust types and napi-compatible types.
- `#[napi(object)]` structs define JS-friendly data transfer types (e.g., `JsStyle`, `JsChartConfig`, `JsPivotTableOption`).
- `Either` enums from napi v3 handle polymorphic values (e.g., cell values that can be string, number, or boolean).
- `index.js` -- ESM module generated by `napi build --esm`. Exports the native addon bindings directly.
- `index.d.ts` -- TypeScript type definitions generated by napi-derive.

## 3. Key Design Decisions

### XML Processing

Most XML parts are handled by `serde` derive macros with `quick-xml`. However, some OOXML parts use namespace prefixes that serde cannot handle correctly:

- **DC namespace** (`dc:`, `dcterms:`, `cp:` in core properties) -- serialized and deserialized using manual quick-xml Writer/Reader.
- **vt: namespace** (variant types in custom properties) -- also handled manually.

The XML declaration `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` is prepended manually to each serialized XML part before writing to the ZIP.

### ZIP Archive Handling

The `zip` crate handles reading and writing .xlsx archives. All ZIP entries use `SimpleFileOptions::default().compression_method(CompressionMethod::Deflated)`. On `open()`, every XML part is read into memory and deserialized. On `save()`, every part is re-serialized and written to a new ZIP archive atomically.

### SharedStrings (SST)

The shared strings table is optional in .xlsx files. If `sharedStrings.xml` is not present in the archive, `Sst::default()` is used. At runtime, the SST maintains a `HashMap<String, usize>` for O(1) deduplication: when a string cell value is set, the SST returns an existing index if the string is already present, or inserts it and returns the new index.

### Style Deduplication

When `add_style()` is called, the style components (font, fill, border, alignment, number format, protection) are each checked against existing records. If all components match an existing XF (cell format) record, that record's index is returned without creating a duplicate. This keeps the styles.xml compact.

### CellValue and Date Detection

`CellValue` is an enum with variants: `String`, `Number`, `Bool`, `Empty`, `Date`, `Formula`, `Error`.

On read, numeric cells are checked against the applied number format. If the format ID falls within known date format ranges (built-in IDs 14-22, 27-36, 45-47) or the custom format string contains date/time patterns, the cell is returned as `CellValue::Date(serial_number)` instead of `CellValue::Number`. The chrono crate handles conversion between Excel serial numbers and `NaiveDate`/`NaiveDateTime`.

### Formula System

The formula system has two independent parts:

1. **Parser**: Uses the `nom` crate to parse formula strings into an AST. The parser handles Excel formula syntax including operator precedence, nested function calls, absolute/relative cell references, range references, and all literal types.

2. **Evaluator**: Walks the AST and computes results. The `CellDataProvider` trait abstracts workbook data access so the evaluator does not directly borrow the workbook. Before evaluation, cell values are snapshotted into a `CellSnapshot` (`HashMap<(String, u32, u32), CellValue>`) to avoid mutable/immutable borrow conflicts when one cell's formula references another cell.

   `calculate_all()` evaluates every formula cell in the workbook in dependency order. It builds a directed graph of cell dependencies, performs a topological sort using Kahn's algorithm, and evaluates cells from leaves to roots. Circular references are detected and reported as errors.

### Conditional Formatting

Conditional formatting rules reference DXF (Differential Formatting) records in styles.xml rather than full XF records. DXF records contain only the formatting properties that differ from the cell's base style. SheetKit supports 17 rule types including cell value comparisons, color scales, data bars, top/bottom N, above/below average, duplicates, unique values, blanks, errors, text matching, and formula-based rules.

### Pivot Tables

A pivot table in OOXML consists of 4 interconnected parts:

1. **pivotTable definition** (`pivotTable{n}.xml`) -- field layout, row/column/data/page fields
2. **pivotCacheDefinition** (`pivotCacheDefinition{n}.xml`) -- source range, field definitions
3. **pivotCacheRecords** (`pivotCacheRecords{n}.xml`) -- cached source data
4. **Workbook pivot caches** -- collection in workbook.xml linking cache IDs to definitions

Each pivot table gets its own dedicated cache. The cache definition specifies the source data range, and the cache records store a snapshot of that data.

### StreamWriter

The StreamWriter provides a forward-only API for generating large .xlsx files. Instead of building the full worksheet XML in memory, it writes XML elements directly to the output as rows are added. Constraints:

- Rows must be written in ascending order (no random access).
- The sheet XML is finalized when `flush()` is called.
- SST entries from the stream are merged into the workbook SST on flush.
- Supports freeze panes, row options (height, visibility, outline), and column-level style/visibility/outline.

### napi Bindings Design

The Node.js bindings follow the `inner` field pattern: the napi `Workbook` class contains a `sheetkit_core::workbook::Workbook` as its `inner` field. Each napi method unwraps arguments from JS types, calls the corresponding method on `inner`, and converts the result back to a JS-compatible type.

napi v3 `Either` types are used for polymorphic values instead of `JsUnknown`, providing type safety on both the Rust and TypeScript sides.

## 4. Data Flow

### Reading an .xlsx file

```
.xlsx file (ZIP archive)
  |
  v
zip::ZipArchive::new(reader)
  |
  v
For each known XML part path (e.g., "xl/workbook.xml", "xl/worksheets/sheet1.xml"):
  zip_archive.by_name(path) -> raw XML bytes
  |
  v
quick_xml::de::from_str() or manual Reader -> sheetkit-xml typed struct
  |
  v
All deserialized parts assembled into Workbook struct
  (sheets, styles, SST, relationships, drawings, charts, etc.)
```

### Manipulating data

```
User code calls Workbook methods:
  wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))
  |
  v
Workbook locates the target worksheet (by name -> sheet index)
  |
  v
SST.get_or_insert("Hello") -> string index (O(1) via HashMap)
  |
  v
Worksheet row/cell data updated with SST index reference
```

### Writing an .xlsx file

```
Workbook.save_as("output.xlsx")
  |
  v
For each XML part:
  quick_xml::se::to_string() or manual Writer -> XML string
  Prepend XML declaration
  |
  v
zip::ZipWriter::start_file(path, options)
  writer.write_all(xml_bytes)
  |
  v
zip::ZipWriter::finish() -> .xlsx file
```

## 5. Testing Strategy

- **Unit tests**: Co-located with their modules using `#[cfg(test)]` inline test blocks. Each module tests its own functionality in isolation.
- **Node.js tests**: Located at `packages/sheetkit/__test__/index.spec.ts`. Uses vitest to test the napi bindings end-to-end.
- **Test coverage**: The project maintains over 1,300 Rust tests and 200 Node.js tests across all modules.
- **Test output files**: Any `.xlsx` files generated during tests are gitignored to keep the repository clean.
