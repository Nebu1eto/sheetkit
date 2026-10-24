# Migrating from 0.5.1 to 1.0

SheetKit 1.0 puts finite limits on every operation, rejects malformed input that
0.5.1 accepted, and fixes many results that differed from Excel. Most
applications upgrade without code changes, but some defaults, return shapes,
errors, and saved output change on purpose.

This guide covers the Rust crates, the `@sheetkit/node` package, and the CLI.
After upgrading, rebuild the native addon and recompile TypeScript against the
new declarations. The Node.js package requires Node.js 24 or later.

## What to check first

- **Resource limits.** Open, formula, traversal, stream, save, encryption, and
  VBA operations use finite defaults. You can raise a limit for trusted input,
  but you cannot remove it.
- **Lazy open.** Node.js opens workbooks lazily and defers auxiliary parts.
  Numeric cells with a date number format are read as dates.
- **Sparse rows.** Row and column APIs return occupied cells only. Dense output
  is an explicit, bounded option.
- **Structured values.** Node.js returns formula and date cells as objects
  instead of plain scalars.
- **Atomic failures.** Formula, import, structural-edit, and save operations
  either complete or leave the workbook unchanged.
- **CSV.** Output neutralizes formula-like text by default.
- **Async saves.** `save`, `writeBuffer`, and `saveWithPassword` run in the
  background. The workbook cannot be modified until they settle.
- **Macros.** Saving to `.xlsx` or `.xltx` removes VBA content.
- **Strict validation.** Malformed references, buffers, tables, colors,
  package parts, and Node.js arguments now fail instead of loading or applying
  partially.

## Resource limits and errors

Each operation applies its own limits. Limits are not global, are not stored
on the workbook, and cannot exceed a fixed hard ceiling.

| Operation | Rust | Node.js | Main defaults |
| --- | --- | --- | --- |
| Open | `OpenOptions`, `OpenLimits`, `CryptLimits`, `VbaLimits` | `OpenOptions` | 128 MiB input, 4,096 ZIP entries, 512 MiB total expanded size, 256 MiB per XML part |
| Formula | `FormulaLimits` | `FormulaLimits` | 32 KiB per formula, 4 MiB total input, 16 MiB total output |
| Traversal | `TraversalOptions`, `TraversalLimits` | `TraversalOptions` | 1,048,576 rows, 16,384 columns, 10 million cells, 512 MiB output |
| Save | `SaveLimits` | `SaveLimits` | 512 MiB archive, 128 MiB per XML part, 10 million rows and cells |
| Stream write | `StreamLimits` | `StreamLimits` | 16 MiB per row, 512 MiB total, 1,048,576 rows, 10 million cells |
| JSON import | none | `limits` option | 1,048,576 rows, 16,384 columns, 10 million cells, 512 MiB retained |
| CSV | CSV options | `ToCsvOptions` | 64 MiB output |

An operation counts several dimensions at once, so an input can fail even when
each individual value is below its default.

### Handling limit errors

In Rust, a limit failure is `Error::ResourceLimitExceeded`. Match it separately
from invalid data and I/O errors:

```rust
use sheetkit::{Error, OpenOptions, Workbook};

match Workbook::open_with_options("report.xlsx", &OpenOptions::default()) {
    Ok(workbook) => println!("{} sheets", workbook.sheet_names().len()),
    Err(Error::ResourceLimitExceeded(details)) => eprintln!(
        "{} exceeded {}: {} > {}",
        details.operation_code(),
        details.dimension_code(),
        details.observed(),
        details.limit(),
    ),
    Err(error) => return Err(error),
}
```

In Node.js, use `isResourceLimitError` instead of matching error messages.
`limit` and `observed` are `bigint`:

```typescript
import { isResourceLimitError, Workbook } from '@sheetkit/node';

try {
  const workbook = await Workbook.open('report.xlsx');
} catch (error) {
  if (!isResourceLimitError(error)) throw error;
  console.error(error.operation, error.dimension, error.limit, error.observed);
}
```

Argument errors, lifecycle errors, and wrong passwords remain ordinary errors.

### Raising a limit

Raise only the dimension that a known, trusted file needs:

```typescript
const workbook = await Workbook.open('trusted-large.xlsx', {
  maxOuterInputBytes: 256 * 1024 * 1024,
  maxUnzipSize: 768 * 1024 * 1024,
  maxZipEntries: 8_000,
});
```

`maxUnzipSize` and `maxZipEntries` are still accepted, but no value disables
the archive checks. Zero, unsafe JavaScript numbers, values above the hard
ceiling, and unknown modes are rejected.

`maxXmlPartBytes` can lower the 256 MiB per-part default. Raising it has no
effect, because each ZIP entry is also limited to 256 MiB and Node.js does not
expose that limit.

The 100:1 compression-ratio check applies only to entries larger than 8 MiB
after decompression, so small repetitive parts such as shared strings open
normally. Change the threshold with `compressionRatioFloorBytes` (Node.js) or
`OpenOptions::compression_ratio_floor_bytes` (Rust); `0` checks every entry.

## Opening workbooks

### Defaults

Node.js open methods default to:

```typescript
{ readMode: 'lazy', auxParts: 'deferred', dateInterpretation: 'numFmt' }
```

The Rust equivalents are `ReadMode::Lazy`, `AuxParts::Deferred`, and
`DateInterpretation::NumFmt`.

A lazy workbook parses a worksheet or auxiliary part on first access, under the
limits chosen at open. A part that exceeds `maxXmlPartBytes` can therefore fail
on first access even though `open` succeeded. To validate the whole file up
front, open eagerly:

```typescript
const workbook = await Workbook.open('input.xlsx', {
  readMode: 'eager',
  auxParts: 'eager',
});
```

A lazy workbook keeps reading the file as it was at open time. Replacing the
file afterward does not change what the workbook reads.

### Filtering and partial reads

Sheets excluded with `sheets` are no longer decompressed or validated, and are
copied unchanged on save. `sheetRows` stops parsing after the requested row.
Neither option validates the rest of the file.

### Legacy option names

`parseMode: 'full'` maps to eager open and `parseMode: 'readfast'` to lazy
open. Use `readMode` in new code. `readMode: 'stream'` currently behaves like
`'lazy'`; use `openSheetReader` for forward streaming.

### Files that now fail to open

Open now fails, in every read mode, when:

- a part fails its CRC or decompression check (0.5.1 dropped such parts);
- a ZIP entry name is duplicated, or two part names differ only by case;
- a worksheet relationship points to another part, outside the package, or to
  an external target;
- an entry's declared size does not match its content (a lazy workbook fails
  when it reads or copies that part).

A part that decompresses correctly but contains malformed XML is still kept.

## Cell values and dates

Node.js cell values now use this type:

```typescript
type CellValue = string | number | boolean | DateValue | FormulaValue | null;
```

With the default `dateInterpretation: 'numFmt'`, a numeric cell with a date
number format becomes a `DateValue`. Choose `'cellType'` to follow the stored
cell type literally. Rust returns `CellValue::Date`.

A formula cell keeps its formula and cached result:

```typescript
const value = workbook.getCellValue('Sheet1', 'D2');
if (typeof value === 'object' && value?.type === 'formula') {
  console.log(value.formula, value.result);
}
```

Dates and formulas keep this shape in direct reads, row traversal, raw
buffers, stream reads, and JSON import. Update type guards and serializers that
expected a `number` or a flattened formula result.

In Rust, `CompactCellRef::new` and `CompactCellRef::from_coordinates` now
return `Result` and enforce Excel's row and column bounds.

## Formulas

### Limits and atomic updates

Parsing, evaluation, `calculateAll`, `setCellFormula`, `fillFormula`,
`FormulaValue` assignment, stream writes, JSON import, and raw buffer import
share one `FormulaLimits` budget per operation. The budget is checked before
any cell changes, so a failed bulk operation leaves no partial results.

```typescript
workbook.fillFormula('Sheet1', 'D2:D1000', 'B2*C2', {
  operationInputBytes: 4 * 1024 * 1024,
  operationEvaluationSteps: 1_048_576,
});
```

Rust provides `set_cell_formula_with_limits`, `fill_formula_with_limits`,
`evaluate_formula_with_limits`, and `calculate_all_with_limits`. The methods
without limits use `FormulaLimits::secure_default()`.

An error inside one cell, such as `#DIV/0!`, is stored as that cell's cached
result. Only cycles and engine failures abort `calculateAll`.

### Changed results

Update expected values in formula tests. Results now match Excel for:

- ranges, blank cells, and exponentiation;
- financial, holiday, `TEXT`, Unicode `LEN`, and Bessel functions;
- `SUM`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`, `MEDIAN`, `LARGE`, `SMALL`,
  and `RANK`, which ignore booleans, text, and empty cells in references.
  `AVERAGE` over a reference with no numbers returns `#DIV/0!`. Values typed
  directly into the formula are still coerced, so `SUM("5", TRUE)` is 6;
- `DATE`, `EDATE`, `EOMONTH`, and `WORKDAY`, which return `#NUM!` outside
  the Excel date range;
- `LEFT`, `RIGHT`, `MID`, `REPT`, and `REPLACE`, which return `#VALUE!`
  for negative counts and clamp very large counts;
- `ROUND`, `ROUNDUP`, and `ROUNDDOWN`, which never return NaN or infinity;
- formulas reached through another sheet, which evaluate unqualified
  references, `ROW()`, and `COLUMN()` in their own sheet and cell.

Some inputs that 0.5.1 accepted are now rejected or bounded: wildcard criteria
in `COUNTIF(S)`, `SUMIF(S)`, `AVERAGEIF(S)`, and `SEARCH`; large periods in
depreciation and cumulative-interest functions; `SUBSTITUTE` results that
would grow too large; and `places` outside 1 to 10 in base-conversion
functions.

## Structural edits

Inserting, deleting, and copying rows and columns now shifts references token
by token instead of by text replacement. Formulas, defined names, data
validations, conditional formats, charts, tables, and other dependent parts
update together, or the edit fails without changes. Remove any workaround that
rewrote formula text yourself.

Behavior to expect:

- `removeRow` and `removeCol` replace references to deleted cells with
  `#REF!` and shrink ranges that lose rows or columns. Merges, validations,
  conditional formats, hyperlinks, auto-filters, comments, and row page breaks
  shrink or are removed.
- Removing a column inside a table removes that table column. An edit that
  would remove a table's header row, totals row, all data rows, or all columns
  is rejected.
- `insertRows` and `insertCols` fail if a comment would move off the sheet.
- `duplicateRow` writes shared formulas on the copy as ordinary formulas.
- Whole-row and whole-column references such as `1:1` and `A:A` shift.
  Whole-sheet references (`1:1048576`, `A:XFD`) and 3-D references do not.

```typescript
workbook.setCellFormula('Sheet1', 'B1', 'A5+SUM(A1:A5)');
workbook.removeRow('Sheet1', 5);
// B1 now holds '#REF!+SUM(A1:A4)'.
```

Sheet copy and deletion now handle every part the sheet owns: drawings,
charts, images, tables, comments, defined names, and controls. `copySheet`
gives copied tables names that do not collide with existing table or defined
names, ignoring case.

### Generated IDs

New relationship IDs, sheet IDs, custom property IDs, pivot cache IDs, table
IDs, and custom number format IDs take the lowest unused value. After deleting
`rId2`, the next relationship is `rId2`, not `rId4`. Conditional format
priorities still use the current maximum plus one, so evaluation order does not
change.

## Rows, columns, raw buffers, JSON, CSV, and HTML

### Sparse traversal

In 0.5.1, one distant cell could make row APIs expand the whole rectangle up to
it. In 1.0, `getRows`, `getCols`, raw buffers, `toJSON`, `toCSV`, and
`toHTML` return occupied data only, within traversal limits:

```typescript
const rows = workbook.getRows('Sheet1', {
  mode: 'sparse',
  limits: { maxCells: 2_000_000, maxOutputBytes: 128 * 1024 * 1024 },
});
```

Each row carries its absolute row and column numbers. Do not count array
positions to infer blank rows. Sparse CSV and HTML also drop leading and
internal empty columns.

If a consumer needs the full rectangle, request dense v1 output with explicit
limits and decode it with `SheetData`:

```typescript
import { SheetData } from '@sheetkit/node';

const buffer = workbook.getRowsBuffer('Sheet1', {
  mode: 'dense',
  limits: { maxCells: 250_000, maxOutputBytes: 64 * 1024 * 1024 },
});
const dense = new SheetData(buffer).toArray();
```

Dense mode exists only for the v1 buffer format. `getCols` rejects it.

### Raw buffers

The v1 format now defaults to sparse output. The v2 format is always sparse,
with absolute coordinates and inline strings. Both decoders validate the whole
buffer before returning data, and `setSheetDataBuffer` validates everything
before writing any cell. Treat buffers from outside the process as untrusted;
a rejected buffer leaves the workbook unchanged and usable.

### JSON

`toJSON` returns plain objects with own data properties. A header named
`__proto__` becomes an ordinary property and no longer replaces the row's
prototype.

`fromJSON` reads array indexes and own data properties once, and ignores
custom iterators. Proxies, getters, duplicate headers, invalid coordinates, and
malformed date or formula objects are rejected without changing the sheet.

```typescript
workbook.fromJSON('Sheet1', records, {
  header: true,
  startCell: 'A1',
  limits: { maxRows: 10_000, maxColumns: 100 },
});
```

### CSV and HTML

CSV output adds an apostrophe before text that starts like a formula, so a
spreadsheet does not execute it. Use `raw-exchange` only when the receiver is not a spreadsheet and
the data is trusted:

```typescript
const safeCsv = workbook.toCSV('Sheet1');
const rawCsv = workbook.toCSV('Sheet1', { formulaPolicy: 'raw-exchange' });
```

CSV output is limited to 64 MiB. HTML escapes text and attributes for their
context and counts the final HTML size against `limits.maxOutputBytes`.

## Streaming

### Reading

Use `openSheetReader` for large forward reads:

```typescript
const reader = await workbook.openSheetReader('Sheet1', { batchSize: 500 });
for await (const batch of reader) {
  for (const row of batch) processRow(row);
}
```

The reader reads from the package or a bounded temporary file instead of
cloning the worksheet. Batch size must be a positive integer up to 100,000. A
truncated worksheet raises an error instead of returning partial rows. Rust
uses `open_sheet_reader` or `open_sheet_reader_owned`.

### Writing

`newStreamWriter` (Node.js) and `StreamWriter::new` (Rust) use finite
defaults. Pass limits when creating the writer:

```typescript
const writer = workbook.newStreamWriter('Export', {
  rowCount: 500_000,
  outputBytes: 256 * 1024 * 1024,
});
```

Rust uses `StreamWriter::new_with_limits`. A row or batch is checked before any
of it is written, so a rejected row is not appended. After an I/O failure the
writer refuses further calls. `addMergeCell` rejects overlapping ranges.

## Saving

### Limits and atomic replacement

All save methods use `SaveLimits`. In Rust, choose limits with
`save_with_limits` or `save_to_buffer_with_limits`; in Node.js, pass them as
the second argument:

```typescript
await workbook.save('output.xlsx', {
  outputBytes: 768 * 1024 * 1024,
  xmlPartBytes: 256 * 1024 * 1024,
});
```

`xmlPartBytes` applies to every worksheet written, including unchanged sheets
that a lazy save copies as is. A save that passed under a low limit in 0.5.1
can now fail.

A path save writes a temporary file in the same directory and replaces the
destination only when the write is complete. If the save fails, the existing
file is unchanged. Saving through a symlink replaces the target file and keeps
the link. On Unix, an existing file keeps its permission bits, and a new file
is created with mode `0o666` minus the process umask.

In Node.js, `saveWithPassword` and `saveWithPasswordSync` take `CryptLimits`
instead of `SaveLimits`. In Rust, `save_with_password_with_output_limits`
accepts both.

<a id="opt-in-save-durability"></a>

### Durable saves

By default, a save is atomic but not flushed to disk. A power loss or OS crash
right after a save can leave the old file, the new file, or, on some file
systems, an empty or incomplete file. An application crash still leaves the
previous file intact, though a forced kill can leave a `.sheetkit-*`
temporary file behind.

Request durability when the file must survive a power loss:

```rust
use sheetkit::{SaveLimits, SaveOptions};

workbook.save_with_options("output.xlsx", SaveLimits::default(), SaveOptions { durable: true })?;
```

```typescript
await workbook.save('output.xlsx', undefined, { durable: true });
```

A durable save flushes the temporary file before the rename and, on Unix,
syncs the directory afterward. On macOS this uses `F_FULLFSYNC` and adds a few
milliseconds per save. If the directory sync fails, the error says that the new
file is already in place. Encrypted saves accept the same option through
`save_with_password_with_options` in Rust and the last `options` argument in
Node.js. Network file systems and drives that ignore flush requests weaken
these guarantees.

### Async saves

`save`, `writeBuffer`, and `saveWithPassword` now run in a background task.
Modifying the workbook while one is pending throws
`workbook has an async save operation in progress`:

```typescript
await workbook.save('output.xlsx');
workbook.setCellValue('Sheet1', 'A1', 'next');
```

Once the task settles, successfully or not, the workbook can be modified and
saved again. Separate workbooks can save at the same time.

### Saved content

A lazy save copies unchanged parts byte for byte and rewrites only edited
parts. Shared strings keep significant whitespace and rich-text runs, and a
rich-text string no longer shares an entry with a plain string of the same
text. Custom document properties keep empty strings. Compare saved files by
reopening them rather than by ZIP bytes, which can change even when the
content does not.

### Macros

The file extension decides whether VBA is kept:

| Target | Result |
| --- | --- |
| `.xlsx`, `.xltx` | VBA parts, relationships, and content types are removed |
| `.xlsm`, `.xltm`, `.xlam` | VBA is kept |

If you saved a macro workbook as `.xlsx` and relied on the VBA staying inside,
save it as `.xlsm` instead.

## Encryption and VBA

Password open applies `CryptLimits` before key derivation and keeps the same
open limits for the decrypted package. Standard and Agile encryption, integrity
keys, and package HMAC now follow Office behavior. Agile files whose password
key and package key differ in size now open; a `saltSize` outside 1 to 65536 is
rejected. Re-check encrypted test files made with 0.5.1 in Excel or another
Office-compatible tool. To raise limits for a large encrypted file, set both
the open and the crypt limits.

VBA inspection uses one `VbaLimits` budget per operation. Module names come
from the Unicode stream-name record when present. Source in CP932, CP949, or
CP936 is decoded lossily: ASCII (and half-width katakana for CP932) is kept,
other double-byte characters become U+FFFD, and each affected module adds a
warning to the project's `warnings` list.

## Charts, tables, pivots, and other parts

Adding or changing these features now writes every related part and
relationship, or fails without changes:

- pivot tables write their cache, records, and relationships with unique IDs;
- slicers write their cache and view extensions and validate placement;
- form controls keep their VML and property relationships and escape labels;
- charts, drawings, images, shapes, and sparklines use unique IDs; stock charts
  require the documented series count and order;
- tables check their width before changing anything and always write headers;
  names follow Excel rules (1 to 255 characters, not cell-reference-like, and
  unique among tables and defined names, ignoring case);
- conditional formats write consistent priorities, formulas, and styles;
- protection hashes, data validation, page breaks, and comments follow Office
  semantics.

A table part that no worksheet clearly owns is reported as invalid instead of
being attached to the first sheet. Lazy workbooks load sparklines, pictures,
comments, controls, names, and merges before any edit that needs them, so you
can remove eager-open workarounds added to keep them.

Regenerate golden files for these features and check them by saving and
reopening, or in Excel.

## Styles and rendering

Theme tints, inherited styles, and number formats now render as in Excel, so
formatted text and SVG snapshots change. SVG rendering escapes font names
itself; pass plain names such as `A&B Sans`, not escaped ones. Colors must be
6-digit RGB or 8-digit ARGB hex, optionally prefixed with `#`; anything else is
rejected.

## Node.js argument checks

- Integer arguments and numeric options reject fractions, NaN, infinity,
  unsafe integers, and out-of-range values instead of truncating them.
- Unknown style and conditional format enum strings throw an error that lists
  the accepted values.

## CLI

Text output from `info`, `sheets`, `read`, and `get` escapes terminal control
characters and is size-limited. Scripts should use JSON:

```bash
sheetkit sheets input.xlsx --format json
```

`read --format csv` and `convert` neutralize formula-like text by default.
Keep it raw only for trusted exchange:

```bash
sheetkit read input.xlsx --sheet Sheet1 --format csv \
  --csv-formula-policy raw-exchange
```

If output exceeds its limit, nothing is written. A failed `convert` leaves the
existing destination file unchanged.

## Performance

On a 100,000-row workbook, compared with 0.5.1 on an Apple M4 Pro with Node.js
24, lazy open is about 29% faster and saving an unchanged lazy workbook about
45% faster. Reading rows, streaming, and saving after edits are within 4% of
0.5.1. Producing a v2 buffer is about 7% (1.8 ms) slower because of the added
validation. If your workload differs, measure it with your own files.

## Upgrade checklist

1. Upgrade Rust and Node.js packages together, rebuild, and recompile
   TypeScript. Use Node.js 24 or later.
2. Replace error-message matching with `Error::ResourceLimitExceeded` or
   `isResourceLimitError`.
3. Raise limits only for trusted files that need it.
4. Decide between lazy and eager open, and check date handling.
5. Update type guards and serializers for `DateValue` and `FormulaValue`.
6. Update row consumers for sparse output with absolute coordinates.
7. Update expected values for formulas, structural edits, number formats, and
   SVG output.
8. Choose a CSV formula policy explicitly and use JSON output from the CLI.
9. Await async saves before modifying the workbook.
10. Pass `durable: true` to saves that must survive a power loss.
11. Save macro workbooks as `.xlsm`, `.xltm`, or `.xlam`.
12. Save and reopen representative files with tables, pivots, charts, slicers,
    controls, comments, and drawings.
