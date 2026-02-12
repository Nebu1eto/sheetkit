# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-12T11:16:53.575Z

## Environment

| Item | Value |
|------|-------|
| CPU | Apple M4 Pro |
| RAM | 24 GB |
| OS | darwin arm64 |
| Node.js | v25.3.0 |
| Rust | rustc 1.93.0 (254b59607 2026-01-19) (SheetKit native backend) |

## Methodology

- **All libraries**: 1 warmup run(s) + 5 measured runs per scenario. Median time reported.
- **RSS (Resident Set Size)**: Total process memory delta measured before/after each run via `process.memoryUsage().rss`. Includes V8 heap, native (Rust/napi) allocations, and OS overhead. This is a post-operation residual measurement, not peak usage during the operation.
- **Heap Used**: V8 heap delta measured before/after each run via `process.memoryUsage().heapUsed`. Isolates JavaScript-side memory growth, excluding native allocations. Useful for comparing JS-only libraries against napi-rs libraries where native memory dominates.
- **GC**: When `--expose-gc` is enabled, `global.gc()` is called before each measurement to reduce noise from deferred garbage collection.
- **Limitations**: Both RSS and heapUsed measure post-operation residual, not peak. Actual peak memory during an operation may be higher due to intermediate allocations freed before measurement. For napi-rs libraries, most memory lives in native heap (visible in RSS but not heapUsed).

## Libraries

| Library | Description |
|---------|-------------|
| **SheetKit** (`@sheetkit/node`) | Rust-based Excel library with Node.js bindings via napi-rs |
| **ExcelJS** (`exceljs`) | Pure JavaScript Excel library with streaming support |
| **SheetJS** (`xlsx`) | Pure JavaScript spreadsheet library (community edition) |

## Test Fixtures

| Fixture | Description |
|---------|-------------|
| `large-data.xlsx` | 50,000 rows x 20 columns, mixed types (numbers, strings, floats, booleans) |
| `heavy-styles.xlsx` | 5,000 rows x 10 columns with rich formatting |
| `multi-sheet.xlsx` | 10 sheets, each with 5,000 rows x 10 columns |
| `formulas.xlsx` | 10,000 rows with 5 formula columns |
| `strings.xlsx` | 20,000 rows x 10 columns of text data (SST stress test) |
| `data-validation.xlsx` | 5,000 rows with 8 validation rules (list, whole, decimal, textLength, custom) |
| `comments.xlsx` | 2,000 rows with cell comments (2,667 total comments) |
| `merged-cells.xlsx` | 500 merged regions (section headers and sub-headers) |
| `mixed-workload.xlsx` | Multi-sheet ERP document with styles, formulas, validation, comments |
| `scale-{1k,10k,100k}.xlsx` | Scaling benchmarks at 1K, 10K, and 100K rows |

## Results

### Read

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Large Data (50k rows x 20 cols) | 546ms | 2.91s | 1.65s | SheetKit |
| Read Large Data (50k rows x 20 cols) (async) | 530ms | N/A | N/A | SheetKit |
| Read Large Data (50k rows x 20 cols) (lazy) | 530ms | N/A | N/A | SheetKit |
| Read Large Data (50k rows x 20 cols) (getRowsRaw) | 535ms | N/A | N/A | SheetKit |
| Read Large Data (50k rows x 20 cols) (lazy+raw) | 529ms | N/A | N/A | SheetKit |
| Read Large Data (50k rows x 20 cols) (bufferV2) | 506ms | N/A | N/A | SheetKit |
| Read Large Data (50k rows x 20 cols) (stream) | 739ms | N/A | N/A | SheetKit |
| Read Heavy Styles (5k rows, formatted) | 29ms | 183ms | 90ms | SheetKit |
| Read Heavy Styles (5k rows, formatted) (async) | 28ms | N/A | N/A | SheetKit |
| Read Heavy Styles (5k rows, formatted) (lazy) | 28ms | N/A | N/A | SheetKit |
| Read Heavy Styles (5k rows, formatted) (getRowsRaw) | 28ms | N/A | N/A | SheetKit |
| Read Heavy Styles (5k rows, formatted) (lazy+raw) | 28ms | N/A | N/A | SheetKit |
| Read Heavy Styles (5k rows, formatted) (bufferV2) | 26ms | N/A | N/A | SheetKit |
| Read Heavy Styles (5k rows, formatted) (stream) | 38ms | N/A | N/A | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) | 625ms | 1.64s | 698ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) (async) | 636ms | N/A | N/A | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) (lazy) | 607ms | N/A | N/A | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) (getRowsRaw) | 609ms | N/A | N/A | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) (lazy+raw) | 606ms | N/A | N/A | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) (bufferV2) | 280ms | N/A | N/A | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) (stream) | 408ms | N/A | N/A | SheetKit |
| Read Formulas (10k rows) | 42ms | 200ms | 85ms | SheetKit |
| Read Formulas (10k rows) (async) | 39ms | N/A | N/A | SheetKit |
| Read Formulas (10k rows) (lazy) | 40ms | N/A | N/A | SheetKit |
| Read Formulas (10k rows) (getRowsRaw) | 40ms | N/A | N/A | SheetKit |
| Read Formulas (10k rows) (lazy+raw) | 40ms | N/A | N/A | SheetKit |
| Read Formulas (10k rows) (bufferV2) | 32ms | N/A | N/A | SheetKit |
| Read Formulas (10k rows) (stream) | 47ms | N/A | N/A | SheetKit |
| Read Strings (20k rows text-heavy) | 112ms | 648ms | 295ms | SheetKit |
| Read Strings (20k rows text-heavy) (async) | 110ms | N/A | N/A | SheetKit |
| Read Strings (20k rows text-heavy) (lazy) | 114ms | N/A | N/A | SheetKit |
| Read Strings (20k rows text-heavy) (getRowsRaw) | 110ms | N/A | N/A | SheetKit |
| Read Strings (20k rows text-heavy) (lazy+raw) | 111ms | N/A | N/A | SheetKit |
| Read Strings (20k rows text-heavy) (bufferV2) | 105ms | N/A | N/A | SheetKit |
| Read Strings (20k rows text-heavy) (stream) | 154ms | N/A | N/A | SheetKit |
| Read Data Validation (5k rows, 8 rules) | 22ms | 147ms | 64ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) (async) | 22ms | N/A | N/A | SheetKit |
| Read Data Validation (5k rows, 8 rules) (lazy) | 23ms | N/A | N/A | SheetKit |
| Read Data Validation (5k rows, 8 rules) (getRowsRaw) | 22ms | N/A | N/A | SheetKit |
| Read Data Validation (5k rows, 8 rules) (lazy+raw) | 22ms | N/A | N/A | SheetKit |
| Read Data Validation (5k rows, 8 rules) (bufferV2) | 21ms | N/A | N/A | SheetKit |
| Read Data Validation (5k rows, 8 rules) (stream) | 30ms | N/A | N/A | SheetKit |
| Read Comments (2k rows with comments) | 9ms | 117ms | 30ms | SheetKit |
| Read Comments (2k rows with comments) (async) | 9ms | N/A | N/A | SheetKit |
| Read Comments (2k rows with comments) (lazy) | 8ms | N/A | N/A | SheetKit |
| Read Comments (2k rows with comments) (getRowsRaw) | 9ms | N/A | N/A | SheetKit |
| Read Comments (2k rows with comments) (lazy+raw) | 8ms | N/A | N/A | SheetKit |
| Read Comments (2k rows with comments) (bufferV2) | 8ms | N/A | N/A | SheetKit |
| Read Comments (2k rows with comments) (stream) | 9ms | N/A | N/A | SheetKit |
| Read Merged Cells (500 regions) | 2ms | 24ms | 6ms | SheetKit |
| Read Merged Cells (500 regions) (async) | 2ms | N/A | N/A | SheetKit |
| Read Merged Cells (500 regions) (lazy) | 2ms | N/A | N/A | SheetKit |
| Read Merged Cells (500 regions) (getRowsRaw) | 2ms | N/A | N/A | SheetKit |
| Read Merged Cells (500 regions) (lazy+raw) | 2ms | N/A | N/A | SheetKit |
| Read Merged Cells (500 regions) (bufferV2) | 2ms | N/A | N/A | SheetKit |
| Read Merged Cells (500 regions) (stream) | 3ms | N/A | N/A | SheetKit |
| Read Mixed Workload (ERP document) | 32ms | 198ms | 77ms | SheetKit |
| Read Mixed Workload (ERP document) (async) | 32ms | N/A | N/A | SheetKit |
| Read Mixed Workload (ERP document) (lazy) | 31ms | N/A | N/A | SheetKit |
| Read Mixed Workload (ERP document) (getRowsRaw) | 33ms | N/A | N/A | SheetKit |
| Read Mixed Workload (ERP document) (lazy+raw) | 31ms | N/A | N/A | SheetKit |
| Read Mixed Workload (ERP document) (bufferV2) | 26ms | N/A | N/A | SheetKit |
| Read Mixed Workload (ERP document) (stream) | 38ms | N/A | N/A | SheetKit |

### Read (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Scale 1k rows | 6ms | 42ms | 20ms | SheetKit |
| Read Scale 1k rows (async) | 6ms | N/A | N/A | SheetKit |
| Read Scale 1k rows (lazy) | 6ms | N/A | N/A | SheetKit |
| Read Scale 1k rows (getRowsRaw) | 6ms | N/A | N/A | SheetKit |
| Read Scale 1k rows (lazy+raw) | 6ms | N/A | N/A | SheetKit |
| Read Scale 1k rows (bufferV2) | 5ms | N/A | N/A | SheetKit |
| Read Scale 1k rows (stream) | 8ms | N/A | N/A | SheetKit |
| Read Scale 10k rows | 55ms | 308ms | 147ms | SheetKit |
| Read Scale 10k rows (async) | 55ms | N/A | N/A | SheetKit |
| Read Scale 10k rows (lazy) | 53ms | N/A | N/A | SheetKit |
| Read Scale 10k rows (getRowsRaw) | 53ms | N/A | N/A | SheetKit |
| Read Scale 10k rows (lazy+raw) | 54ms | N/A | N/A | SheetKit |
| Read Scale 10k rows (bufferV2) | 49ms | N/A | N/A | SheetKit |
| Read Scale 10k rows (stream) | 73ms | N/A | N/A | SheetKit |
| Read Scale 100k rows | 565ms | 2.96s | 1.63s | SheetKit |
| Read Scale 100k rows (async) | 569ms | N/A | N/A | SheetKit |
| Read Scale 100k rows (lazy) | 573ms | N/A | N/A | SheetKit |
| Read Scale 100k rows (getRowsRaw) | 565ms | N/A | N/A | SheetKit |
| Read Scale 100k rows (lazy+raw) | 574ms | N/A | N/A | SheetKit |
| Read Scale 100k rows (bufferV2) | 526ms | N/A | N/A | SheetKit |
| Read Scale 100k rows (stream) | 773ms | N/A | N/A | SheetKit |

### Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 50000 rows x 20 cols | 489ms | 2.73s | 1.37s | SheetKit |
| Write 5000 styled rows | 36ms | 168ms | 46ms | SheetKit |
| Write 10 sheets x 5000 rows | 260ms | 1.31s | 462ms | SheetKit |
| Write 10000 rows with formulas | 30ms | 167ms | 63ms | SheetKit |
| Write 20000 text-heavy rows | 90ms | 511ms | 232ms | SheetKit |

### Write (DV)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 10ms | 94ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 2000 rows with comments | 8ms | 69ms | 71ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 500 merged regions | 2ms | 28ms | 3ms | SheetKit |

### Write (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 1k rows x 10 cols | 5ms | 38ms | 9ms | SheetKit |
| Write 10k rows x 10 cols | 48ms | 264ms | 97ms | SheetKit |
| Write 50k rows x 10 cols | 244ms | 1.33s | 532ms | SheetKit |
| Write 100k rows x 10 cols | 508ms | 3.00s | 1.27s | SheetKit |

### Round-Trip

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 131ms | 498ms | 198ms | SheetKit |
| Buffer round-trip (10000 rows) (lazy) | 125ms | N/A | N/A | SheetKit |

### Streaming

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Streaming write (50000 rows) | 332ms | 550ms | N/A | SheetKit |

### Random Access

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Random-access (open+1000 lookups) | 489ms | 3.04s | 1.37s | SheetKit |
| Random-access (open+1000 lookups) (async) | 473ms | N/A | N/A | SheetKit |
| Random-access (open+1000 lookups) (lazy) | 465ms | N/A | N/A | SheetKit |
| Random-access (lookup-only, 1000 cells) | 473ms | 2.98s | 1.35s | SheetKit |
| Random-access (lookup-only, 1000 cells) (async open) | 464ms | N/A | N/A | SheetKit |

### Mixed Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Mixed workload write (ERP-style) | 20ms | 113ms | N/A | SheetKit |

### COW Save

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Copy-on-write save (untouched) (lazy) | 217ms | N/A | N/A | SheetKit |
| Copy-on-write save (untouched) (eager) | 682ms | N/A | N/A | SheetKit |
| Copy-on-write save (single-cell edit) (lazy) | 692ms | N/A | N/A | SheetKit |
| Copy-on-write save (single-cell edit) (eager) | 685ms | N/A | N/A | SheetKit |

### Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | RSS (median) | Heap (median) |
|----------|---------|--------|-----|-----|-----|--------------|---------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 546ms | 534ms | 557ms | 557ms | 192.4MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) | ExcelJS | 2.91s | 2.88s | 3.00s | 3.00s | 0.0MB | 0.1MB |
| Read Large Data (50k rows x 20 cols) | SheetJS | 1.65s | 1.63s | 1.66s | 1.66s | 0.0MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) (async) | SheetKit | 530ms | 525ms | 535ms | 535ms | 4.3MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) (lazy) | SheetKit | 530ms | 526ms | 544ms | 544ms | 0.0MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) (getRowsRaw) | SheetKit | 535ms | 525ms | 553ms | 553ms | 4.2MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) (lazy+raw) | SheetKit | 529ms | 520ms | 549ms | 549ms | 0.0MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) (bufferV2) | SheetKit | 506ms | 489ms | 554ms | 554ms | 3.6MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) (stream) | SheetKit | 739ms | 735ms | 779ms | 779ms | 0.0MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 29ms | 28ms | 29ms | 29ms | 5.4MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | ExcelJS | 183ms | 179ms | 184ms | 184ms | 0.2MB | 0.1MB |
| Read Heavy Styles (5k rows, formatted) | SheetJS | 90ms | 84ms | 91ms | 91ms | 0.0MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (async) | SheetKit | 28ms | 28ms | 30ms | 30ms | 0.0MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (lazy) | SheetKit | 28ms | 28ms | 28ms | 28ms | 0.0MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (getRowsRaw) | SheetKit | 28ms | 28ms | 29ms | 29ms | 0.0MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (lazy+raw) | SheetKit | 28ms | 28ms | 29ms | 29ms | 0.0MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (bufferV2) | SheetKit | 26ms | 26ms | 26ms | 26ms | 0.0MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (stream) | SheetKit | 38ms | 37ms | 44ms | 44ms | 0.0MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 625ms | 619ms | 655ms | 655ms | 121.7MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | ExcelJS | 1.64s | 1.62s | 1.67s | 1.67s | 0.0MB | 0.1MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetJS | 698ms | 674ms | 722ms | 722ms | 0.0MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (async) | SheetKit | 636ms | 630ms | 652ms | 652ms | 0.4MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (lazy) | SheetKit | 607ms | 598ms | 628ms | 628ms | 10.3MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (getRowsRaw) | SheetKit | 609ms | 599ms | 612ms | 612ms | 0.4MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (lazy+raw) | SheetKit | 606ms | 592ms | 635ms | 635ms | 0.4MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (bufferV2) | SheetKit | 280ms | 278ms | 290ms | 290ms | 0.4MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (stream) | SheetKit | 408ms | 404ms | 419ms | 419ms | 0.4MB | 0.0MB |
| Read Formulas (10k rows) | SheetKit | 42ms | 41ms | 44ms | 44ms | 9.3MB | 0.0MB |
| Read Formulas (10k rows) | ExcelJS | 200ms | 196ms | 205ms | 205ms | 0.0MB | 0.0MB |
| Read Formulas (10k rows) | SheetJS | 85ms | 83ms | 86ms | 86ms | 0.0MB | 0.0MB |
| Read Formulas (10k rows) (async) | SheetKit | 39ms | 39ms | 40ms | 40ms | 0.0MB | 0.0MB |
| Read Formulas (10k rows) (lazy) | SheetKit | 40ms | 39ms | 41ms | 41ms | 0.0MB | 0.0MB |
| Read Formulas (10k rows) (getRowsRaw) | SheetKit | 40ms | 39ms | 43ms | 43ms | 0.0MB | 0.0MB |
| Read Formulas (10k rows) (lazy+raw) | SheetKit | 40ms | 39ms | 40ms | 40ms | 0.0MB | 0.0MB |
| Read Formulas (10k rows) (bufferV2) | SheetKit | 32ms | 32ms | 32ms | 32ms | 0.0MB | 0.0MB |
| Read Formulas (10k rows) (stream) | SheetKit | 47ms | 46ms | 48ms | 48ms | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) | SheetKit | 112ms | 111ms | 117ms | 117ms | 10.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) | ExcelJS | 648ms | 645ms | 661ms | 661ms | 0.0MB | 0.1MB |
| Read Strings (20k rows text-heavy) | SheetJS | 295ms | 292ms | 326ms | 326ms | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) (async) | SheetKit | 110ms | 109ms | 112ms | 112ms | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) (lazy) | SheetKit | 114ms | 112ms | 115ms | 115ms | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) (getRowsRaw) | SheetKit | 110ms | 110ms | 111ms | 111ms | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) (lazy+raw) | SheetKit | 111ms | 111ms | 115ms | 115ms | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) (bufferV2) | SheetKit | 105ms | 104ms | 106ms | 106ms | 2.7MB | 0.0MB |
| Read Strings (20k rows text-heavy) (stream) | SheetKit | 154ms | 153ms | 165ms | 165ms | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | SheetKit | 22ms | 22ms | 23ms | 23ms | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | ExcelJS | 147ms | 139ms | 150ms | 150ms | 3.1MB | 16.4MB |
| Read Data Validation (5k rows, 8 rules) | SheetJS | 64ms | 61ms | 68ms | 68ms | 0.1MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (async) | SheetKit | 22ms | 22ms | 22ms | 22ms | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (lazy) | SheetKit | 23ms | 22ms | 23ms | 23ms | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (getRowsRaw) | SheetKit | 22ms | 22ms | 23ms | 23ms | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (lazy+raw) | SheetKit | 22ms | 22ms | 22ms | 22ms | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (bufferV2) | SheetKit | 21ms | 21ms | 21ms | 21ms | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (stream) | SheetKit | 30ms | 30ms | 31ms | 31ms | 0.0MB | 0.0MB |
| Read Comments (2k rows with comments) | SheetKit | 9ms | 9ms | 9ms | 9ms | 0.6MB | 0.0MB |
| Read Comments (2k rows with comments) | ExcelJS | 117ms | 114ms | 121ms | 121ms | 0.0MB | 0.1MB |
| Read Comments (2k rows with comments) | SheetJS | 30ms | 26ms | 32ms | 32ms | 0.0MB | 0.0MB |
| Read Comments (2k rows with comments) (async) | SheetKit | 9ms | 9ms | 9ms | 9ms | 0.0MB | 0.0MB |
| Read Comments (2k rows with comments) (lazy) | SheetKit | 8ms | 7ms | 8ms | 8ms | 0.0MB | 0.0MB |
| Read Comments (2k rows with comments) (getRowsRaw) | SheetKit | 9ms | 9ms | 9ms | 9ms | 0.0MB | 0.0MB |
| Read Comments (2k rows with comments) (lazy+raw) | SheetKit | 8ms | 8ms | 8ms | 8ms | 0.0MB | 0.0MB |
| Read Comments (2k rows with comments) (bufferV2) | SheetKit | 8ms | 8ms | 8ms | 8ms | 0.0MB | 0.0MB |
| Read Comments (2k rows with comments) (stream) | SheetKit | 9ms | 9ms | 9ms | 9ms | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) | ExcelJS | 24ms | 22ms | 25ms | 25ms | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) | SheetJS | 6ms | 6ms | 6ms | 6ms | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) (async) | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) (lazy) | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) (getRowsRaw) | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) (lazy+raw) | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) (bufferV2) | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) (stream) | SheetKit | 3ms | 2ms | 3ms | 3ms | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) | SheetKit | 32ms | 31ms | 34ms | 34ms | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) | ExcelJS | 198ms | 195ms | 227ms | 227ms | 0.1MB | 0.1MB |
| Read Mixed Workload (ERP document) | SheetJS | 77ms | 75ms | 82ms | 82ms | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) (async) | SheetKit | 32ms | 31ms | 32ms | 32ms | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) (lazy) | SheetKit | 31ms | 30ms | 31ms | 31ms | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) (getRowsRaw) | SheetKit | 33ms | 32ms | 35ms | 35ms | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) (lazy+raw) | SheetKit | 31ms | 31ms | 32ms | 32ms | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) (bufferV2) | SheetKit | 26ms | 26ms | 27ms | 27ms | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) (stream) | SheetKit | 38ms | 38ms | 39ms | 39ms | 0.0MB | 0.0MB |
| Read Scale 1k rows | SheetKit | 6ms | 6ms | 6ms | 6ms | 0.0MB | 0.0MB |
| Read Scale 1k rows | ExcelJS | 42ms | 40ms | 43ms | 43ms | 0.0MB | 0.0MB |
| Read Scale 1k rows | SheetJS | 20ms | 19ms | 22ms | 22ms | 0.0MB | 0.0MB |
| Read Scale 1k rows (async) | SheetKit | 6ms | 6ms | 6ms | 6ms | 0.0MB | 0.0MB |
| Read Scale 1k rows (lazy) | SheetKit | 6ms | 6ms | 6ms | 6ms | 0.0MB | 0.0MB |
| Read Scale 1k rows (getRowsRaw) | SheetKit | 6ms | 6ms | 6ms | 6ms | 0.0MB | 0.0MB |
| Read Scale 1k rows (lazy+raw) | SheetKit | 6ms | 5ms | 6ms | 6ms | 0.0MB | 0.0MB |
| Read Scale 1k rows (bufferV2) | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0MB | 0.0MB |
| Read Scale 1k rows (stream) | SheetKit | 8ms | 7ms | 8ms | 8ms | 0.0MB | 0.0MB |
| Read Scale 10k rows | SheetKit | 55ms | 54ms | 56ms | 56ms | 0.4MB | 0.0MB |
| Read Scale 10k rows | ExcelJS | 308ms | 298ms | 323ms | 323ms | 0.0MB | 0.1MB |
| Read Scale 10k rows | SheetJS | 147ms | 144ms | 152ms | 152ms | 0.0MB | 0.0MB |
| Read Scale 10k rows (async) | SheetKit | 55ms | 54ms | 55ms | 55ms | 0.0MB | 0.0MB |
| Read Scale 10k rows (lazy) | SheetKit | 53ms | 53ms | 54ms | 54ms | 0.0MB | 0.0MB |
| Read Scale 10k rows (getRowsRaw) | SheetKit | 53ms | 53ms | 55ms | 55ms | 0.0MB | 0.0MB |
| Read Scale 10k rows (lazy+raw) | SheetKit | 54ms | 53ms | 55ms | 55ms | 0.0MB | 0.0MB |
| Read Scale 10k rows (bufferV2) | SheetKit | 49ms | 49ms | 50ms | 50ms | 0.0MB | 0.0MB |
| Read Scale 10k rows (stream) | SheetKit | 73ms | 72ms | 75ms | 75ms | 0.0MB | 0.0MB |
| Read Scale 100k rows | SheetKit | 565ms | 562ms | 571ms | 571ms | 184.5MB | 0.0MB |
| Read Scale 100k rows | ExcelJS | 2.96s | 2.95s | 3.06s | 3.06s | 0.7MB | 0.0MB |
| Read Scale 100k rows | SheetJS | 1.63s | 1.58s | 1.67s | 1.67s | 0.1MB | 0.0MB |
| Read Scale 100k rows (async) | SheetKit | 569ms | 560ms | 575ms | 575ms | 16.1MB | 0.0MB |
| Read Scale 100k rows (lazy) | SheetKit | 573ms | 570ms | 581ms | 581ms | 8.4MB | 0.0MB |
| Read Scale 100k rows (getRowsRaw) | SheetKit | 565ms | 553ms | 582ms | 582ms | 8.4MB | 0.0MB |
| Read Scale 100k rows (lazy+raw) | SheetKit | 574ms | 562ms | 579ms | 579ms | 8.4MB | 0.0MB |
| Read Scale 100k rows (bufferV2) | SheetKit | 526ms | 508ms | 557ms | 557ms | 8.4MB | 0.0MB |
| Read Scale 100k rows (stream) | SheetKit | 773ms | 740ms | 791ms | 791ms | 0.0MB | 0.0MB |
| Write 50000 rows x 20 cols | SheetKit | 489ms | 483ms | 497ms | 497ms | 157.0MB | 0.0MB |
| Write 50000 rows x 20 cols | ExcelJS | 2.73s | 2.62s | 3.12s | 3.12s | 5.0MB | 0.0MB |
| Write 50000 rows x 20 cols | SheetJS | 1.37s | 1.29s | 1.64s | 1.64s | 0.0MB | 0.0MB |
| Write 5000 styled rows | SheetKit | 36ms | 35ms | 36ms | 36ms | 8.1MB | 0.0MB |
| Write 5000 styled rows | ExcelJS | 168ms | 164ms | 185ms | 185ms | 0.3MB | 0.0MB |
| Write 5000 styled rows | SheetJS | 46ms | 44ms | 49ms | 49ms | 0.0MB | 0.0MB |
| Write 10 sheets x 5000 rows | SheetKit | 260ms | 256ms | 277ms | 277ms | 97.7MB | 0.0MB |
| Write 10 sheets x 5000 rows | ExcelJS | 1.31s | 1.28s | 1.39s | 1.39s | 0.4MB | 0.0MB |
| Write 10 sheets x 5000 rows | SheetJS | 462ms | 434ms | 501ms | 501ms | 0.0MB | 0.0MB |
| Write 10000 rows with formulas | SheetKit | 30ms | 29ms | 30ms | 30ms | 7.5MB | 0.0MB |
| Write 10000 rows with formulas | ExcelJS | 167ms | 166ms | 171ms | 171ms | 0.2MB | 0.0MB |
| Write 10000 rows with formulas | SheetJS | 63ms | 62ms | 68ms | 68ms | 0.0MB | 0.1MB |
| Write 20000 text-heavy rows | SheetKit | 90ms | 90ms | 94ms | 94ms | 7.0MB | 0.0MB |
| Write 20000 text-heavy rows | ExcelJS | 511ms | 497ms | 730ms | 730ms | 0.0MB | 0.0MB |
| Write 20000 text-heavy rows | SheetJS | 232ms | 225ms | 239ms | 239ms | 0.2MB | 0.0MB |
| Write 5000 rows + 8 validation rules | SheetKit | 10ms | 9ms | 10ms | 10ms | 1.8MB | 0.0MB |
| Write 5000 rows + 8 validation rules | ExcelJS | 94ms | 91ms | 97ms | 97ms | 0.3MB | 0.1MB |
| Write 2000 rows with comments | SheetKit | 8ms | 8ms | 8ms | 8ms | 0.2MB | 0.0MB |
| Write 2000 rows with comments | ExcelJS | 69ms | 68ms | 71ms | 71ms | 0.1MB | 0.1MB |
| Write 2000 rows with comments | SheetJS | 71ms | 71ms | 73ms | 73ms | 0.0MB | 0.0MB |
| Write 500 merged regions | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB | 0.0MB |
| Write 500 merged regions | ExcelJS | 28ms | 19ms | 29ms | 29ms | 0.0MB | 0.1MB |
| Write 500 merged regions | SheetJS | 3ms | 3ms | 3ms | 3ms | 0.0MB | 0.0MB |
| Write 1k rows x 10 cols | SheetKit | 5ms | 5ms | 6ms | 6ms | 0.0MB | 0.0MB |
| Write 1k rows x 10 cols | ExcelJS | 38ms | 35ms | 39ms | 39ms | 0.1MB | 0.0MB |
| Write 1k rows x 10 cols | SheetJS | 9ms | 9ms | 10ms | 10ms | 0.0MB | 0.0MB |
| Write 10k rows x 10 cols | SheetKit | 48ms | 46ms | 49ms | 49ms | 3.3MB | 0.0MB |
| Write 10k rows x 10 cols | ExcelJS | 264ms | 263ms | 265ms | 265ms | 0.3MB | 0.0MB |
| Write 10k rows x 10 cols | SheetJS | 97ms | 92ms | 100ms | 100ms | 0.0MB | 0.0MB |
| Write 50k rows x 10 cols | SheetKit | 244ms | 243ms | 245ms | 245ms | 40.2MB | 0.0MB |
| Write 50k rows x 10 cols | ExcelJS | 1.33s | 1.29s | 1.36s | 1.36s | 10.3MB | 0.0MB |
| Write 50k rows x 10 cols | SheetJS | 532ms | 522ms | 538ms | 538ms | 0.0MB | 0.0MB |
| Write 100k rows x 10 cols | SheetKit | 508ms | 493ms | 518ms | 518ms | 117.5MB | 0.0MB |
| Write 100k rows x 10 cols | ExcelJS | 3.00s | 2.81s | 3.27s | 3.27s | 0.0MB | 0.0MB |
| Write 100k rows x 10 cols | SheetJS | 1.27s | 1.24s | 1.40s | 1.40s | 0.0MB | 0.0MB |
| Buffer round-trip (10000 rows) | SheetKit | 131ms | 129ms | 132ms | 132ms | 13.7MB | 0.0MB |
| Buffer round-trip (10000 rows) | ExcelJS | 498ms | 493ms | 504ms | 504ms | 0.3MB | 0.0MB |
| Buffer round-trip (10000 rows) | SheetJS | 198ms | 189ms | 222ms | 222ms | 0.0MB | 0.0MB |
| Buffer round-trip (10000 rows) (lazy) | SheetKit | 125ms | 125ms | 126ms | 126ms | 0.7MB | 0.0MB |
| Streaming write (50000 rows) | SheetKit | 332ms | 331ms | 357ms | 357ms | 0.0MB | 0.0MB |
| Streaming write (50000 rows) | ExcelJS | 550ms | 546ms | 553ms | 553ms | 2.0MB | 0.0MB |
| Random-access (open+1000 lookups) | SheetKit | 489ms | 488ms | 495ms | 495ms | 160.8MB | 0.0MB |
| Random-access (open+1000 lookups) | ExcelJS | 3.04s | 2.94s | 3.07s | 3.07s | 0.0MB | 0.0MB |
| Random-access (open+1000 lookups) | SheetJS | 1.37s | 1.35s | 1.38s | 1.38s | 0.0MB | 0.0MB |
| Random-access (open+1000 lookups) (async) | SheetKit | 473ms | 456ms | 481ms | 481ms | 17.7MB | 0.0MB |
| Random-access (open+1000 lookups) (lazy) | SheetKit | 465ms | 463ms | 466ms | 466ms | 21.4MB | 0.0MB |
| Random-access (lookup-only, 1000 cells) | SheetKit | 473ms | 471ms | 482ms | 482ms | 55.0MB | 0.0MB |
| Random-access (lookup-only, 1000 cells) | ExcelJS | 2.98s | 2.93s | 3.00s | 3.00s | 2.3MB | 0.0MB |
| Random-access (lookup-only, 1000 cells) | SheetJS | 1.35s | 1.34s | 1.37s | 1.37s | 0.0MB | 0.0MB |
| Random-access (lookup-only, 1000 cells) (async open) | SheetKit | 464ms | 462ms | 469ms | 469ms | 14.6MB | 0.0MB |
| Mixed workload write (ERP-style) | SheetKit | 20ms | 20ms | 21ms | 21ms | 3.3MB | 0.0MB |
| Mixed workload write (ERP-style) | ExcelJS | 113ms | 109ms | 148ms | 148ms | 0.3MB | 0.0MB |
| Copy-on-write save (untouched) (lazy) | SheetKit | 217ms | 214ms | 221ms | 221ms | 24.2MB | 0.0MB |
| Copy-on-write save (untouched) (eager) | SheetKit | 682ms | 682ms | 693ms | 693ms | 36.1MB | 0.0MB |
| Copy-on-write save (single-cell edit) (lazy) | SheetKit | 692ms | 687ms | 701ms | 701ms | 36.1MB | 0.0MB |
| Copy-on-write save (single-cell edit) (eager) | SheetKit | 685ms | 679ms | 691ms | 691ms | 46.3MB | 0.0MB |

### Memory Usage (RSS / Heap Used)

RSS = Resident Set Size delta (total process memory). Heap = V8 heapUsed delta (JS-only memory).

| Scenario | SheetKit (RSS/Heap) | ExcelJS (RSS/Heap) | SheetJS (RSS/Heap) |
|----------|---------------------|--------------------|--------------------|
| Read Large Data (50k rows x 20 cols) | 192.4MB / 0.0MB | 0.0MB / 0.1MB | 0.0MB / 0.0MB |
| Read Large Data (50k rows x 20 cols) (async) | 4.3MB / 0.0MB | N/A | N/A |
| Read Large Data (50k rows x 20 cols) (lazy) | 0.0MB / 0.0MB | N/A | N/A |
| Read Large Data (50k rows x 20 cols) (getRowsRaw) | 4.2MB / 0.0MB | N/A | N/A |
| Read Large Data (50k rows x 20 cols) (lazy+raw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Large Data (50k rows x 20 cols) (bufferV2) | 3.6MB / 0.0MB | N/A | N/A |
| Read Large Data (50k rows x 20 cols) (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) | 5.4MB / 0.0MB | 0.2MB / 0.1MB | 0.0MB / 0.0MB |
| Read Heavy Styles (5k rows, formatted) (async) | 0.0MB / 0.0MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) (lazy) | 0.0MB / 0.0MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) (lazy+raw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) | 121.7MB / 0.0MB | 0.0MB / 0.1MB | 0.0MB / 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (async) | 0.4MB / 0.0MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) (lazy) | 10.3MB / 0.0MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) (getRowsRaw) | 0.4MB / 0.0MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) (lazy+raw) | 0.4MB / 0.0MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) (bufferV2) | 0.4MB / 0.0MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) (stream) | 0.4MB / 0.0MB | N/A | N/A |
| Read Formulas (10k rows) | 9.3MB / 0.0MB | 0.0MB / 0.0MB | 0.0MB / 0.0MB |
| Read Formulas (10k rows) (async) | 0.0MB / 0.0MB | N/A | N/A |
| Read Formulas (10k rows) (lazy) | 0.0MB / 0.0MB | N/A | N/A |
| Read Formulas (10k rows) (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Formulas (10k rows) (lazy+raw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Formulas (10k rows) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Formulas (10k rows) (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Strings (20k rows text-heavy) | 10.0MB / 0.0MB | 0.0MB / 0.1MB | 0.0MB / 0.0MB |
| Read Strings (20k rows text-heavy) (async) | 0.0MB / 0.0MB | N/A | N/A |
| Read Strings (20k rows text-heavy) (lazy) | 0.0MB / 0.0MB | N/A | N/A |
| Read Strings (20k rows text-heavy) (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Strings (20k rows text-heavy) (lazy+raw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Strings (20k rows text-heavy) (bufferV2) | 2.7MB / 0.0MB | N/A | N/A |
| Read Strings (20k rows text-heavy) (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) | 0.0MB / 0.0MB | 3.1MB / 16.4MB | 0.1MB / 0.0MB |
| Read Data Validation (5k rows, 8 rules) (async) | 0.0MB / 0.0MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) (lazy) | 0.0MB / 0.0MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) (lazy+raw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Comments (2k rows with comments) | 0.6MB / 0.0MB | 0.0MB / 0.1MB | 0.0MB / 0.0MB |
| Read Comments (2k rows with comments) (async) | 0.0MB / 0.0MB | N/A | N/A |
| Read Comments (2k rows with comments) (lazy) | 0.0MB / 0.0MB | N/A | N/A |
| Read Comments (2k rows with comments) (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Comments (2k rows with comments) (lazy+raw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Comments (2k rows with comments) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Comments (2k rows with comments) (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Merged Cells (500 regions) | 0.0MB / 0.0MB | 0.0MB / 0.0MB | 0.0MB / 0.0MB |
| Read Merged Cells (500 regions) (async) | 0.0MB / 0.0MB | N/A | N/A |
| Read Merged Cells (500 regions) (lazy) | 0.0MB / 0.0MB | N/A | N/A |
| Read Merged Cells (500 regions) (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Merged Cells (500 regions) (lazy+raw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Merged Cells (500 regions) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Merged Cells (500 regions) (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Mixed Workload (ERP document) | 0.0MB / 0.0MB | 0.1MB / 0.1MB | 0.0MB / 0.0MB |
| Read Mixed Workload (ERP document) (async) | 0.0MB / 0.0MB | N/A | N/A |
| Read Mixed Workload (ERP document) (lazy) | 0.0MB / 0.0MB | N/A | N/A |
| Read Mixed Workload (ERP document) (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Mixed Workload (ERP document) (lazy+raw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Mixed Workload (ERP document) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Mixed Workload (ERP document) (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 1k rows | 0.0MB / 0.0MB | 0.0MB / 0.0MB | 0.0MB / 0.0MB |
| Read Scale 1k rows (async) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 1k rows (lazy) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 1k rows (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 1k rows (lazy+raw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 1k rows (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 1k rows (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 10k rows | 0.4MB / 0.0MB | 0.0MB / 0.1MB | 0.0MB / 0.0MB |
| Read Scale 10k rows (async) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 10k rows (lazy) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 10k rows (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 10k rows (lazy+raw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 10k rows (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 10k rows (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 100k rows | 184.5MB / 0.0MB | 0.7MB / 0.0MB | 0.1MB / 0.0MB |
| Read Scale 100k rows (async) | 16.1MB / 0.0MB | N/A | N/A |
| Read Scale 100k rows (lazy) | 8.4MB / 0.0MB | N/A | N/A |
| Read Scale 100k rows (getRowsRaw) | 8.4MB / 0.0MB | N/A | N/A |
| Read Scale 100k rows (lazy+raw) | 8.4MB / 0.0MB | N/A | N/A |
| Read Scale 100k rows (bufferV2) | 8.4MB / 0.0MB | N/A | N/A |
| Read Scale 100k rows (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Write 50000 rows x 20 cols | 157.0MB / 0.0MB | 5.0MB / 0.0MB | 0.0MB / 0.0MB |
| Write 5000 styled rows | 8.1MB / 0.0MB | 0.3MB / 0.0MB | 0.0MB / 0.0MB |
| Write 10 sheets x 5000 rows | 97.7MB / 0.0MB | 0.4MB / 0.0MB | 0.0MB / 0.0MB |
| Write 10000 rows with formulas | 7.5MB / 0.0MB | 0.2MB / 0.0MB | 0.0MB / 0.1MB |
| Write 20000 text-heavy rows | 7.0MB / 0.0MB | 0.0MB / 0.0MB | 0.2MB / 0.0MB |
| Write 5000 rows + 8 validation rules | 1.8MB / 0.0MB | 0.3MB / 0.1MB | N/A |
| Write 2000 rows with comments | 0.2MB / 0.0MB | 0.1MB / 0.1MB | 0.0MB / 0.0MB |
| Write 500 merged regions | 0.0MB / 0.0MB | 0.0MB / 0.1MB | 0.0MB / 0.0MB |
| Write 1k rows x 10 cols | 0.0MB / 0.0MB | 0.1MB / 0.0MB | 0.0MB / 0.0MB |
| Write 10k rows x 10 cols | 3.3MB / 0.0MB | 0.3MB / 0.0MB | 0.0MB / 0.0MB |
| Write 50k rows x 10 cols | 40.2MB / 0.0MB | 10.3MB / 0.0MB | 0.0MB / 0.0MB |
| Write 100k rows x 10 cols | 117.5MB / 0.0MB | 0.0MB / 0.0MB | 0.0MB / 0.0MB |
| Buffer round-trip (10000 rows) | 13.7MB / 0.0MB | 0.3MB / 0.0MB | 0.0MB / 0.0MB |
| Buffer round-trip (10000 rows) (lazy) | 0.7MB / 0.0MB | N/A | N/A |
| Streaming write (50000 rows) | 0.0MB / 0.0MB | 2.0MB / 0.0MB | N/A |
| Random-access (open+1000 lookups) | 160.8MB / 0.0MB | 0.0MB / 0.0MB | 0.0MB / 0.0MB |
| Random-access (open+1000 lookups) (async) | 17.7MB / 0.0MB | N/A | N/A |
| Random-access (open+1000 lookups) (lazy) | 21.4MB / 0.0MB | N/A | N/A |
| Random-access (lookup-only, 1000 cells) | 55.0MB / 0.0MB | 2.3MB / 0.0MB | 0.0MB / 0.0MB |
| Random-access (lookup-only, 1000 cells) (async open) | 14.6MB / 0.0MB | N/A | N/A |
| Mixed workload write (ERP-style) | 3.3MB / 0.0MB | 0.3MB / 0.0MB | N/A |
| Copy-on-write save (untouched) (lazy) | 24.2MB / 0.0MB | N/A | N/A |
| Copy-on-write save (untouched) (eager) | 36.1MB / 0.0MB | N/A | N/A |
| Copy-on-write save (single-cell edit) (lazy) | 36.1MB / 0.0MB | N/A | N/A |
| Copy-on-write save (single-cell edit) (eager) | 46.3MB / 0.0MB | N/A | N/A |

## Summary

Total scenarios: 109

| Library | Wins |
|---------|------|
| SheetKit | 109/109 |
| ExcelJS | 0/109 |
| SheetJS | 0/109 |
