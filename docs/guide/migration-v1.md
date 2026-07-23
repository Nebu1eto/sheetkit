# Migrating from v0.5.1 to 1.0

SheetKit 1.0 changes several defaults and rejects inputs that v0.5.1 could
accept without a finite resource policy. It also corrects formula, package,
rendering, and Office interoperability behavior accumulated since v0.5.1.

This guide is for applications upgrading directly from v0.5.1. It covers the
Rust crate, the `@sheetkit/node` package, and the CLI.

## Migration at a glance

Plan for these changes before updating:

- Workbook open, traversal, formula, stream, save, encryption, VBA, and output
  operations now use finite defaults. Callers can select larger finite limits,
  but cannot disable the hard ceilings.
- Node open remains lazy by default. Auxiliary parts are deferred, and
  date-formatted numeric cells use number-format interpretation by default.
- Row and column traversal is sparse by default. Dense traversal is an
  explicit, bounded v1 transfer-format compatibility mode.
- Node formula and date cells are structured objects rather than flattened
  scalar values.
- Formula parsing, calculation, assignment, cached results, and reference
  shifts share finite operation budgets. Failed bulk operations do not commit
  partial results.
- CSV output is spreadsheet-safe by default. Raw formula-capable text requires
  an explicit `raw-exchange` policy.
- Async Node save methods perform work off the JavaScript thread. The workbook
  cannot be mutated while an async save or buffer write is in progress.
- Saving a macro-enabled workbook to a macro-free format removes its VBA
  content. Choose an XLSM, XLTM, or XLAM target to retain macros.
- Strict cell references, transfer buffers, table ownership, style colors, and
  advanced OOXML package graphs may reject malformed data that previously
  loaded or mutated partially.

Recompile Rust and TypeScript consumers after upgrading. The generated Node
declarations compile with Node 24 types, and the addon uses Node-API 8.

## 1. Handle finite resource limits and structured errors

### What changed

The v0.5.1 compatibility methods remain available, but they now select secure
finite defaults. Option-taking methods can raise a specific limit for a trusted
workload. Every caller-selected value must still fit the non-overridable hard
ceiling.

Limits apply to one operation. They do not alter process-global state or store
an unlimited policy on a workbook.

| Operation | Rust policy | Node policy | Selected default ceilings |
|---|---|---|---|
| Open | `OpenOptions`, `OpenLimits`, `CryptLimits`, `VbaLimits` | `OpenOptions` | 128 MiB outer input, 4,096 ZIP entries, 512 MiB aggregate expansion, 256 MiB XML part |
| Formula | `FormulaLimits` | `FormulaLimits` | 32 KiB per formula input, 4 MiB cumulative input, 16 MiB cumulative output |
| Traversal | `TraversalOptions`, `TraversalLimits` | `TraversalOptions` | 1,048,576 rows, 16,384 columns, 10 million cells, 512 MiB output |
| Save | `SaveLimits` | `SaveLimits` | 512 MiB archive, 128 MiB XML part, 10 million rows and cells |
| Stream write | `StreamLimits` | `StreamLimits` | 16 MiB row XML, 512 MiB row output, 1,048,576 rows, 10 million cells |
| CSV | bounded CSV options | `ToCsvOptions` | 64 MiB returned UTF-8 |

These are defaults, not a promise that every input below one individual value
will succeed. Operations charge multiple dimensions and cumulative work.

### Rust error handling

Resource exhaustion is `Error::ResourceLimitExceeded`. Match it separately from
invalid data and I/O failures:

```rust
use sheetkit::{Error, OpenOptions, Workbook};

fn inspect_report() -> sheetkit::Result<()> {
    match Workbook::open_with_options("report.xlsx", &OpenOptions::default()) {
        Ok(workbook) => {
            println!("opened {} sheets", workbook.sheet_names().len());
        }
        Err(Error::ResourceLimitExceeded(details)) => {
            eprintln!(
                "{} exceeded {}: observed {}, limit {}",
                details.operation_code(),
                details.dimension_code(),
                details.observed(),
                details.limit(),
            );
        }
        Err(error) => return Err(error),
    }

    Ok(())
}
```

Propagate `Result` from constructors and mutators. In particular,
`CompactCellRef::new` and `CompactCellRef::from_coordinates` are fallible in
1.0 and enforce Excel's row and column bounds.

### Node error handling

Do not parse error messages. Use `isResourceLimitError` and its stable fields:

```typescript
import { isResourceLimitError, Workbook } from '@sheetkit/node';

try {
  const workbook = await Workbook.open('report.xlsx');
  await workbook.save('copy.xlsx');
} catch (error) {
  if (isResourceLimitError(error)) {
    console.error({
      operation: error.operation,
      category: error.category,
      dimension: error.dimension,
      limit: error.limit,
      observed: error.observed,
      stage: error.stage,
    });
  } else {
    throw error;
  }
}
```

`limit` and `observed` are `bigint`. Lifecycle and ordinary argument failures
remain normal JavaScript errors.

### Raising a limit for trusted input

Override only the dimension required by a known workbook:

```typescript
const workbook = await Workbook.open('trusted-large.xlsx', {
  maxOuterInputBytes: 256 * 1024 * 1024,
  maxUnzipSize: 768 * 1024 * 1024,
  maxZipEntries: 8_000,
});
```

`maxUnzipSize` and `maxZipEntries` remain accepted compatibility names, but no
value turns the archive policy off. Unknown modes, unsafe JavaScript numbers,
zero limits, and values above hard ceilings reject instead of silently
falling back.

`maxXmlPartBytes` can lower the 256 MiB default XML ceiling. Raising that field
alone does not admit an XML ZIP entry above the separate 256 MiB per-entry
declared and actual defaults; Node 1.0 does not expose those two per-entry
overrides.

## 2. Open, lazy hydration, and date interpretation

### Defaults in 1.0

Node open methods use:

```typescript
{
  readMode: 'lazy',
  auxParts: 'deferred',
  dateInterpretation: 'numFmt',
}
```

The equivalent Rust values are `ReadMode::Lazy`, `AuxParts::Deferred`, and
`DateInterpretation::NumFmt`.

Lazy hydration retains the resolved open policy. A worksheet or auxiliary XML
part that exceeds `maxXmlPartBytes` can therefore fail when first accessed,
even if the initial `open` call succeeded. If the whole package must be
validated immediately, open eagerly:

```typescript
const workbook = await Workbook.open('input.xlsx', {
  readMode: 'eager',
  auxParts: 'eager',
});
```

Filtering with `sheets` no longer decompresses or validates excluded
worksheets. Do not treat a filtered open as validation of the complete file.
Unchanged excluded sheets remain package-backed for save.

### Legacy parse modes

The compatibility adapter maps `parseMode: 'full'` to eager behavior and
`parseMode: 'readfast'` to lazy behavior. New code should use `readMode`:

```typescript
// v0.5.1 compatibility spelling
const before = await Workbook.open('input.xlsx', {
  parseMode: 'readfast',
});

// 1.0 spelling
const after = await Workbook.open('input.xlsx', {
  readMode: 'lazy',
  auxParts: 'deferred',
});
```

`readMode: 'stream'` is currently a reserved selector with the same open-path
behavior as `lazy`. `openSheetReader` is the forward batch API and does not
require `readMode: 'stream'`. Random cell access still requires worksheet
hydration.

### Date values

The default `numFmt` policy recognizes how Excel commonly stores dates: a
numeric cell with a date number format becomes a structured `DateValue`.
Choose `cellType` only when the XML cell type must be followed literally:

```typescript
const workbook = await Workbook.open('input.xlsx', {
  dateInterpretation: 'cellType',
});
```

Audit TypeScript code that assumed every numeric XML cell returned a `number`.
In Node, direct reads, row traversal, raw transfer, and stream reads preserve
`DateValue` consistently. Rust exposes the corresponding value as
`CellValue::Date`.

## 3. Formula values, mutation, and recalculation

### Structured values

Formula cells and cached results survive Node round trips:

```typescript
const value = workbook.getCellValue('Sheet1', 'D2');

if (typeof value === 'object' && value?.type === 'formula') {
  console.log(value.formula);
  console.log(value.result);
}
```

Update serializers and type guards that expected a cached formula result to
replace the formula object.

### Finite, atomic operations

The same `FormulaLimits` model covers parsing, evaluation, `calculateAll`,
direct assignment, range fill, generic `FormulaValue` mutations, stream
writes, JSON import, and raw v1/v2 import.

Before 1.0, generated formulas could consume unbounded nesting, range, matching,
or output work. Some bulk paths could discover an error after earlier cells had
changed. In 1.0, the operation preflights its cumulative budget and stages bulk
results before commit.

```typescript
workbook.setCellFormula('Sheet1', 'D2', 'B2*C2', {
  perFormulaInputBytes: 32 * 1024,
  operationOutputBytes: 16 * 1024 * 1024,
});

workbook.fillFormula('Sheet1', 'D2:D1000', 'B2*C2', {
  operationInputBytes: 4 * 1024 * 1024,
  operationEvaluationSteps: 1_048_576,
});
```

Rust provides `set_cell_formula_with_limits`,
`fill_formula_with_limits`, `evaluate_formula_with_limits`, and
`calculate_all_with_limits`. Their no-options counterparts use secure defaults.

### Result changes

Rebaseline formula tests where v0.5.1 differed from Excel. The corrections
include:

- range and blank handling;
- exponentiation and several financial functions;
- holiday, `TEXT`, Unicode `LEN`, and Bessel behavior;
- checked `places` values for base conversions;
- bounded periods for depreciation and cumulative-interest functions;
- bounded wildcard criteria for `COUNTIF(S)`, `SUMIF(S)`,
  `AVERAGEIF(S)`, and `SEARCH`;
- preflighted `SUBSTITUTE` amplification.

Cell-domain errors are stored as cached errors for the affected formula cell.
Cycle and engine failures still abort calculation. Do not assume every
spreadsheet-domain error rejects the entire `calculateAll` call.

### Structural edits and references

Row and column insertion, deletion, copy, and formula fill use token-aware A1
reference shifts. Qualified references, quoted sheet names, string literals,
absolute rows, defined names, validations, charts, tables, and related package
parts are updated as one operation.

Remove workarounds that performed textual formula replacement. Rebaseline
golden files that encoded the old partial-reference behavior.

## 4. Sparse traversal, raw transfer, JSON, CSV, and HTML

### Sparse is the default

In v0.5.1, a distant occupied cell could make row-oriented APIs expand the
whole bounding rectangle. In 1.0, `getRows`, `getCols`, v2 raw buffers, JSON,
CSV, and HTML operate on occupied data under finite traversal limits.

```typescript
const rows = workbook.getRows('Sheet1', {
  mode: 'sparse',
  limits: {
    maxRows: 100_000,
    maxColumns: 1_000,
    maxCells: 2_000_000,
    maxOutputBytes: 128 * 1024 * 1024,
  },
});
```

Sparse rows preserve their absolute row and column coordinates. Do not infer
blank row counts from array positions.

If a legacy consumer requires the complete bounding rectangle, request bounded
dense v1 transfer. `getRowsBuffer` returns an opaque buffer; decode it through
`SheetData` or the exported buffer helpers:

```typescript
import { SheetData } from '@sheetkit/node';

const denseV1Buffer = workbook.getRowsBuffer('Sheet1', {
  mode: 'dense',
  limits: { maxCells: 250_000, maxOutputBytes: 64 * 1024 * 1024 },
});
const dense = new SheetData(denseV1Buffer).toArray();
```

Dense mode is not available for v2. `getRows`, `getRowsRaw`, and
`getRowsIterator` select the bounded v1 transfer internally when dense mode is
requested. `getCols` accepts finite sparse limits and rejects dense mode.

### Raw v1 and v2 buffers

V1 remains supported and now defaults to sparse output. V2 is the canonical
sparse format with absolute coordinates and inline strings.

Both decoders validate headers, dimensions, offsets, tags, UTF-8, counts,
ordering, and trailing bytes before exposing data. `setSheetDataBuffer`
preflights shifted coordinates, retained values, strings, formulas, and stale
SST references before one commit.

Treat raw buffers as untrusted input:

```typescript
try {
  workbook.setSheetDataBuffer('Sheet1', receivedBuffer, 'B2');
} catch (error) {
  // The workbook is reusable and unchanged when validation fails.
  throw error;
}
```

Custom decoders must use the published absolute coordinates. Do not depend on
gaps being represented by allocated blank rows.

### JSON export

`toJSON` creates row records with enumerable own data properties. A worksheet
header such as `__proto__` remains ordinary data and does not replace the row
object's prototype:

```typescript
const rows = workbook.toJSON('Sheet1', { header: true });
const first = rows[0];
if (first && Object.hasOwn(first, '__proto__')) {
  console.log(first['__proto__']);
}
```

Use own-property access for data-derived headers. Do not rely on prototype
mutation from v0.5.1.

### JSON import

`fromJSON` accepts ordinary arrays of records and optional finite import
limits:

```typescript
workbook.fromJSON(
  'Sheet1',
  [
    { Name: 'Alice', Score: 42 },
    { Name: 'Bob', Score: 37 },
  ],
  {
    header: true,
    startCell: 'A1',
    limits: {
      maxRows: 10_000,
      maxColumns: 100,
      maxCells: 1_000_000,
      maxRetainedBytes: 64 * 1024 * 1024,
    },
  },
);
```

Proxies, accessors, duplicate custom headers, non-canonical coordinates, and
malformed date/formula objects can now reject. Overridden array iterators are
ignored; imports use indexed own data properties and snapshot accepted
descriptors once. Do not rely on repeated getter evaluation during import.

### CSV policy

Text beginning with a spreadsheet formula prefix is neutralized by default:

```typescript
const safeCsv = workbook.toCSV('Sheet1');
```

Use raw exchange only when the receiver treats the data as trusted and the
formula-capable text is intentional:

```typescript
const rawCsv = workbook.toCSV('Sheet1', {
  formulaPolicy: 'raw-exchange',
  limits: { maxOutputBytes: 32 * 1024 * 1024 },
});
```

CSV syntax values are validated, and returned UTF-8 cannot exceed 64 MiB.
HTML and SVG use sink-specific escaping rather than CSV or XML pre-escaping.
Sparse CSV and HTML compact each occupied row, so leading columns and internal
column gaps are not represented as empty fields or cells. Select a finite dense
v1 traversal when those positions or other blank regions are part of the output
contract.

## 5. Streaming reads and writes

### Stream reads

Prefer `openSheetReader` over the synchronous `getRowsIterator` for large
forward reads:

```typescript
const workbook = await Workbook.open('large.xlsx');
const reader = await workbook.openSheetReader('Sheet1', {
  batchSize: 500,
});

for await (const batch of reader) {
  for (const row of batch) {
    processRow(row);
  }
}
```

Batch size is a positive finite integer with a fixed maximum. Package-backed
bytes or a bounded disk spool replace full worksheet clones. Formula, rich
text, date, and absolute coordinate values are retained across the stream.
Rust uses `open_sheet_reader` or `open_sheet_reader_owned`; `ReadMode::Stream`
does not select a separate parser in 1.0.

`sheetRows` stops worksheet parsing after the requested row count. It is a
read filter, not full-file validation: malformed skipped rows may remain
unread.

### Stream writes

Rust `StreamWriter::new` and Node `newStreamWriter` use finite defaults. Select
limits at construction for a trusted workload:

```typescript
const writer = workbook.newStreamWriter('Export', {
  rowCount: 500_000,
  cellCount: 5_000_000,
  rowXmlBytes: 8 * 1024 * 1024,
  outputBytes: 256 * 1024 * 1024,
});
```

One row is measured before publication, and a batch is admitted before its
first row is written. Formula input and output consume the same cumulative
formula budget as ordinary mutations. A rejected row is not appended. An I/O
failure poisons the writer so later calls cannot report a misleading success.

Rust callers that construct explicit limits use
`StreamWriter::new_with_limits`. Keep row order and Excel bounds canonical.

## 6. Save, async work, encryption, and macros

### Bounded, atomic save

Rust `save`, `save_to_buffer`, and password variants use finite defaults.
Use `save_with_limits` or `save_to_buffer_with_limits` to select a trusted
profile:

```rust
use sheetkit::{SaveLimits, Workbook};

let workbook = Workbook::open("input.xlsx")?;
let limits = SaveLimits {
    output_bytes: 768 * 1024 * 1024,
    xml_part_bytes: 256 * 1024 * 1024,
    ..SaveLimits::default()
};
workbook.save_with_limits("output.xlsx", limits)?;
```

Node `saveSync`, `save`, `writeBufferSync`, and `writeBuffer` accept
`SaveLimits`:

```typescript
await workbook.save('output.xlsx', {
  outputBytes: 768 * 1024 * 1024,
  xmlPartBytes: 256 * 1024 * 1024,
});
```

Node `saveWithPasswordSync` and `saveWithPassword` accept `CryptLimits`, not
`SaveLimits`; their intermediate archive uses finite save defaults. Rust can
select both policies with `save_with_password_with_output_limits`.

Path saves use a same-directory temporary file and replace the destination only
after a complete write. If a limit or I/O error occurs, an existing destination
remains intact. Buffer save is all-or-error under the same save policy.

Lazy saves copy unchanged package-backed parts without parsing and
re-serializing them. Edited sheets and auxiliary parts hydrate on demand;
unrelated preserved OOXML can therefore remain byte-identical while changed
parts are canonicalized.

Shared-string saves preserve significant whitespace and rich-text runs while
emitting canonical SST counts and relationships. Remove workarounds that
flattened rich text or restored whitespace after a save. Archive-byte snapshots
may change even when the workbook meaning does not.

### Await async save before mutation

`save`, `writeBuffer`, and `saveWithPassword` now perform their work in a
background task. The task holds a shared reference to the workbook. Mutation
while it is pending rejects with
`workbook has an async save operation in progress`.

```typescript
// Incorrect in 1.0: mutation races the pending save task.
const pending = workbook.save('output.xlsx');
workbook.setCellValue('Sheet1', 'A1', 'next');
await pending;

// Correct: await the task before reusing the workbook for mutation.
await workbook.save('output.xlsx');
workbook.setCellValue('Sheet1', 'A1', 'next');
```

After a successful or failed task settles, the workbook can be mutated and
saved again. Independent workbook instances can save concurrently.

### Office encryption

Password open applies `CryptLimits` before expensive KDF work and carries the
same open budget into the decrypted ZIP. Standard and Agile parsing, integrity
keys, and package HMAC handling now follow Office-compatible behavior.

Audit private encrypted fixtures made for earlier SheetKit behavior. Recreate
nonstandard files and verify representative files with Microsoft Office or an
independent Office-compatible implementation. When raising limits for a trusted
legacy file, configure both the open and crypt dimensions; neither policy
resets the other.

VBA project and module inspection use one finite `VbaLimits` budget for
container streams, compressed input, decoded modules, retained data, and
returned output. VBA token decoding and comment-text interpretation also use
corrected Office semantics, so rebaseline fixtures that asserted the
v0.5.1-decoded text.

### Macro target format

The destination extension controls macro retention:

| Target | 1.0 behavior |
|---|---|
| XLSX, XLTX | Remove VBA bytes, relationships, and macro content types |
| XLSM, XLTM, XLAM | Preserve canonical VBA content |

This is an intentional content change. Code that saved an XLSM source as XLSX
and expected the hidden VBA project to remain must choose a macro-enabled
target.

## 7. OOXML packages, charts, tables, and rendering

### Package graph corrections

Sheet copy and deletion now process the complete sheet-owned relationship
closure. Drawings, charts, images, tables, comments, names, controls, and
package paths are copied or removed atomically without ID collisions.

Generated or modified advanced features now emit complete package graphs:

- pivot tables include cache, table, record, workbook, and worksheet
  relationships with collision-free IDs;
- slicers include cache and view extension graphs and validate placement;
- form controls preserve VML and property relationships, escape labels, and
  avoid corrupting unrelated comments or VML;
- charts, stock charts, drawings, images, shapes, and sparklines use unique IDs
  and validated references;
- tables preflight width before mutation and emit complete headers;
- conditional formatting emits coherent priorities, formulas, and differential
  styles;
- workbook and sheet protection, data validation, page breaks, and comments
  use corrected Office field, hash, range, and text semantics.

Stock charts require the documented series count and order. Invalid placement,
range, width, validation list, page-break, and package ownership data now
reject rather than leaving a partial graph.

Rebaseline generated OOXML and save/reopen tests. Compare workbook semantics
and required package relationships instead of requiring complete ZIP byte
identity.

### Strict ownership and deferred state

An orphan or ambiguous table part is no longer assigned to the first
worksheet. Content type, package path, relationship, and worksheet
`tableParts` must agree.

Lazy sparkline, picture, comment, control, name, merge, and auxiliary state is
materialized before an edit that needs authoritative data. Remove eager-open
workarounds used only to prevent these parts from disappearing.

### Styles and rendering

Theme tint, inherited styles, Excel number formats, and formatted values have
corrected output in 1.0. SVG rendering validates RGB or ARGB colors, encodes
font names separately for CSS and XML attributes, and bounds dense layout and
output.

Pass logical, unescaped font names and canonical ASCII colors:

```typescript
// Supply data, not pre-escaped CSS/XML fragments.
const styleId = workbook.addStyle({
  font: { name: 'A&B Sans', color: 'FF336699' },
});
workbook.setCellStyle('Sheet1', 'A1', styleId);
```

Rebaseline formatted text and SVG snapshots. Invalid non-ASCII or malformed
color strings now reject instead of being copied into the document.

## 8. CLI output and automation

Human-readable CLI output escapes terminal control characters and has finite
byte limits. JSON output preserves the original structured values.

Use JSON for automation:

```bash
# v0.5.1 scripts sometimes parsed the human table.
sheetkit sheets input.xlsx

# 1.0 machine-readable form.
sheetkit sheets input.xlsx --format json
```

The same rule applies to `info`, `read`, and `get`. Rebaseline terminal output
if workbook metadata, sheet names, or cell values contain control characters.

CSV reads and conversions use spreadsheet-safe output by default:

```bash
sheetkit read input.xlsx --sheet Sheet1 --format csv

# Trusted data exchange that intentionally preserves formula-capable text.
sheetkit read input.xlsx \
  --sheet Sheet1 \
  --format csv \
  --csv-formula-policy raw-exchange
```

Destination conversion is staged and replaced atomically. Existing output is
preserved if serialization fails.

<a id="cumulative-change-map"></a>

## 9. Cumulative change map since v0.5.1

This map accounts for the 23 public or contributor changes included between
v0.5.1 and the 1.0 security work.

| Change area | What to verify in 1.0 |
|---|---|
| AI contribution policy | Runtime behavior is unchanged. Contributors follow `AI_POLICY.md` and the repository's commit trailer rules. |
| Unsafe workbook inputs | Formatting, formula, raw-transfer, stream, and sheet-edit paths return errors instead of panicking or leaving partial mutations. |
| Reference-aware structural edits | Row and column changes update formulas and other reference-bearing parts atomically using token-aware shifts. |
| Sheet lifecycle graphs | Sheet copy and deletion process the complete sheet-owned relationship closure. |
| Save pipeline fidelity | Saves emit canonical SST metadata and preserve untouched lazy bytes, rich text, and significant whitespace. |
| Advanced OOXML round trips | Unknown OOXML and reviewed style/property fields survive unrelated edits; graph mutations preflight before commit. |
| Formula recalculation errors | Cell-domain errors become cached cell errors; engine and cycle errors still abort the operation. |
| Excel formula results | Rebaseline range, blank, financial, date/text, Unicode, and Bessel results. |
| Buffers and streaming | V1/v2 transfer validates structure and preserves absolute coordinates, formulas, and rich text. |
| Node value and option parity | `DateValue`, `FormulaValue`, and open defaults have the same meaning as their Rust counterparts. |
| Office encryption | Standard and Agile parsing and package HMAC use Office-compatible behavior. |
| Office metadata semantics | VBA, protection, validation, page-break, and comment interpretation and validation are corrected. |
| Tables and conditional formats | Width preflight, headers, formulas, priorities, and differential styles are complete. |
| Formatting and theme rendering | Theme tint, inherited styles, and number-format display results are corrected. |
| Lazy workbook features | Pictures, sparklines, comments, controls, names, and merges survive lazy workflows. |
| Pivot packages | Cache, table, record, workbook, and worksheet graphs are complete. |
| Slicer package graphs | Cache/view extensions, placement, and stale-part cleanup are validated. |
| Form-control package graphs | VML/property relationships, label escaping, and placement are validated. |
| Charts and drawings | IDs are unique; stock-chart series/axes and escaped ranges are valid. |
| Agile integrity keys | SheetKit emits hash-sized keys and accepts the published salt-sized variant. |
| Chart/control documentation | Stock-series order and control metadata requirements are explicit. |
| Lazy sparkline state | Copy and shift hydrate sparkline state without dirtying unrelated tables. |
| Audit-review parity | Lazy graphs, Node values/defaults, raw transfer, package paths, and Agile variants agree across Rust and Node. |

The 77 reportable security findings map to the guide by domain:

- Formula: `SK-DS-001` through `SK-DS-013`, `SK-DS-035` through
  `SK-DS-041`, `SK-DS-044`, `SK-DS-051`, and `SK-DS-053` through
  `SK-DS-061`.
- Open, encryption, and VBA: `SK-DS-014` through `SK-DS-017`,
  `SK-DS-045`, `SK-DS-046`, `SK-DS-048`, `SK-DS-049`, `SK-DS-064`,
  `SK-DS-065`, `SK-DS-068`, `SK-DS-069`, and `SK-DS-077`.
- Traversal and transfer: `SK-DS-023` through `SK-DS-034`, `SK-DS-043`,
  `SK-DS-062`, `SK-DS-063`, `SK-DS-066`, `SK-DS-067`, and `SK-DS-070`.
- SVG, output, and CLI: `SK-DS-018` through `SK-DS-022`, `SK-DS-042`,
  `SK-DS-047`, `SK-DS-050`, `SK-DS-052`, and `SK-DS-071` through
  `SK-DS-074`.
- Package and macro handling: `SK-DS-076` and `SK-DS-078`.

`SK-DS-075` was not reportable and is not part of the 77-finding count.

## 10. Performance status

The final clean-environment comparison used the repository's existing focused
benchmark, identical fixtures, one warm-up, five measured samples, and median
latency. It compared the 1.0 candidate with v0.5.1 and excluded installation,
build, linking, and fixture-generation time.

On the 100,000-row fixture, lazy path open improved by 28.6%, ordinary
`getRows` changed by +0.2%, `getRowsRaw-only` by +1.4%, streaming by +3.3%,
untouched copy-on-write save by -44.6%, and edited save by -1.0%. V2-to-V2
buffer production changed by +7.1%, or 1.772 ms. That path now performs
canonical coordinate, shared-formula, date-style, and finite-output checks; it
remains below the 10% release-stop threshold.

These measurements describe the recorded Apple M4 Pro and Node.js 24.18.0
environment. They are release evidence, not a performance guarantee for every
workbook or machine.

## 11. Upgrade checklist

1. Update Rust and Node dependencies together, then rebuild the native addon
   and recompile TypeScript against the generated declarations.
2. Run application tests on Node 24 when using `@sheetkit/node`.
3. Replace message matching with `ResourceLimitExceeded` or
   `isResourceLimitError`.
4. Inventory trusted large inputs and set explicit finite open, crypt, VBA,
   formula, traversal, stream, and save limits only where the defaults are too
   small.
5. Test lazy first access, eager validation, filtered sheets, deferred
   auxiliary parts, and both date-interpretation policies.
6. Update Node type guards and serializers for `DateValue`, `FormulaValue`, and
   cached formula errors.
7. Rebaseline formulas, structural reference edits, number formats, themes,
   and SVG output.
8. Update row and column consumers for sparse absolute coordinates. Use
   bounded dense v1 output only where blank positions are part of the
   application contract.
9. Validate every external raw buffer before use and confirm failed
   `setSheetDataBuffer` or `fromJSON` imports leave the workbook reusable.
10. Select `spreadsheet-safe` or `raw-exchange` CSV policy explicitly and use
    CLI JSON output for automation.
11. Await each async plain save, encrypted save, or buffer write before
    mutating that workbook.
12. Verify encrypted files with Office-compatible tooling and choose a
    macro-enabled target when VBA must survive.
13. Open, modify, save, and reopen representative files containing tables,
    pivots, charts, slicers, controls, comments, drawings, sparklines, and
    unknown OOXML extensions.
14. Confirm failed path saves preserve the existing destination and that the
    workbook remains usable after an error.
15. Re-run the same application-specific fixtures if your workload differs
    materially from the release benchmark corpus.

Contributor-only changes do not affect runtime migration. AI-assisted
contributions must follow `AI_POLICY.md` and the repository's commit trailer
convention.
