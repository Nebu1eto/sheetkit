# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-12T12:10:16.496Z

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
| Read Large Data (50k rows x 20 cols) | 541ms | 1.24s | 1.56s | SheetKit |
| Read Large Data (50k rows x 20 cols) (sync) | 525ms | N/A | N/A | - |
| Read Large Data (50k rows x 20 cols) (getRowsRaw) | 530ms | N/A | N/A | - |
| Read Large Data (50k rows x 20 cols) (bufferV2) | 491ms | N/A | N/A | - |
| Read Large Data (50k rows x 20 cols) (stream) | 696ms | N/A | N/A | - |
| Read Heavy Styles (5k rows, formatted) | 27ms | 67ms | 70ms | SheetKit |
| Read Heavy Styles (5k rows, formatted) (sync) | 27ms | N/A | N/A | - |
| Read Heavy Styles (5k rows, formatted) (getRowsRaw) | 27ms | N/A | N/A | - |
| Read Heavy Styles (5k rows, formatted) (bufferV2) | 24ms | N/A | N/A | - |
| Read Heavy Styles (5k rows, formatted) (stream) | 35ms | N/A | N/A | - |
| Read Multi-Sheet (10 sheets x 5k rows) | 290ms | 613ms | 602ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) (sync) | 294ms | N/A | N/A | - |
| Read Multi-Sheet (10 sheets x 5k rows) (getRowsRaw) | 292ms | N/A | N/A | - |
| Read Multi-Sheet (10 sheets x 5k rows) (bufferV2) | 272ms | N/A | N/A | - |
| Read Multi-Sheet (10 sheets x 5k rows) (stream) | 397ms | N/A | N/A | - |
| Read Formulas (10k rows) | 34ms | 79ms | 68ms | SheetKit |
| Read Formulas (10k rows) (sync) | 35ms | N/A | N/A | - |
| Read Formulas (10k rows) (getRowsRaw) | 33ms | N/A | N/A | - |
| Read Formulas (10k rows) (bufferV2) | 29ms | N/A | N/A | - |
| Read Formulas (10k rows) (stream) | 45ms | N/A | N/A | - |
| Read Strings (20k rows text-heavy) | 117ms | 232ms | 267ms | SheetKit |
| Read Strings (20k rows text-heavy) (sync) | 117ms | N/A | N/A | - |
| Read Strings (20k rows text-heavy) (getRowsRaw) | 116ms | N/A | N/A | - |
| Read Strings (20k rows text-heavy) (bufferV2) | 99ms | N/A | N/A | - |
| Read Strings (20k rows text-heavy) (stream) | 150ms | N/A | N/A | - |
| Read Data Validation (5k rows, 8 rules) | 21ms | 53ms | 45ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) (sync) | 21ms | N/A | N/A | - |
| Read Data Validation (5k rows, 8 rules) (getRowsRaw) | 21ms | N/A | N/A | - |
| Read Data Validation (5k rows, 8 rules) (bufferV2) | 19ms | N/A | N/A | - |
| Read Data Validation (5k rows, 8 rules) (stream) | 28ms | N/A | N/A | - |
| Read Comments (2k rows with comments) | 8ms | 32ms | 20ms | SheetKit |
| Read Comments (2k rows with comments) (sync) | 8ms | N/A | N/A | - |
| Read Comments (2k rows with comments) (getRowsRaw) | 8ms | N/A | N/A | - |
| Read Comments (2k rows with comments) (bufferV2) | 7ms | N/A | N/A | - |
| Read Comments (2k rows with comments) (stream) | 8ms | N/A | N/A | - |
| Read Merged Cells (500 regions) | 1ms | 9ms | 3ms | SheetKit |
| Read Merged Cells (500 regions) (sync) | 1ms | N/A | N/A | - |
| Read Merged Cells (500 regions) (getRowsRaw) | 1ms | N/A | N/A | - |
| Read Merged Cells (500 regions) (bufferV2) | 1ms | N/A | N/A | - |
| Read Merged Cells (500 regions) (stream) | 2ms | N/A | N/A | - |
| Read Mixed Workload (ERP document) | 29ms | 69ms | 60ms | SheetKit |
| Read Mixed Workload (ERP document) (sync) | 28ms | N/A | N/A | - |
| Read Mixed Workload (ERP document) (getRowsRaw) | 28ms | N/A | N/A | - |
| Read Mixed Workload (ERP document) (bufferV2) | 27ms | N/A | N/A | - |
| Read Mixed Workload (ERP document) (stream) | 39ms | N/A | N/A | - |

### Read (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Scale 1k rows | 5ms | 13ms | 11ms | SheetKit |
| Read Scale 1k rows (sync) | 5ms | N/A | N/A | - |
| Read Scale 1k rows (getRowsRaw) | 5ms | N/A | N/A | - |
| Read Scale 1k rows (bufferV2) | 5ms | N/A | N/A | - |
| Read Scale 1k rows (stream) | 7ms | N/A | N/A | - |
| Read Scale 10k rows | 50ms | 121ms | 118ms | SheetKit |
| Read Scale 10k rows (sync) | 49ms | N/A | N/A | - |
| Read Scale 10k rows (getRowsRaw) | 50ms | N/A | N/A | - |
| Read Scale 10k rows (bufferV2) | 47ms | N/A | N/A | - |
| Read Scale 10k rows (stream) | 70ms | N/A | N/A | - |
| Read Scale 100k rows | 535ms | 1.30s | 1.60s | SheetKit |
| Read Scale 100k rows (sync) | 539ms | N/A | N/A | - |
| Read Scale 100k rows (getRowsRaw) | 513ms | N/A | N/A | - |
| Read Scale 100k rows (bufferV2) | 487ms | N/A | N/A | - |
| Read Scale 100k rows (stream) | 713ms | N/A | N/A | - |

### Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 50000 rows x 20 cols | 469ms | 2.62s | 1.09s | SheetKit |
| Write 5000 styled rows | 32ms | 147ms | 34ms | SheetKit |
| Write 10 sheets x 5000 rows | 237ms | 1.27s | 331ms | SheetKit |
| Write 10000 rows with formulas | 27ms | 133ms | 54ms | SheetKit |
| Write 20000 text-heavy rows | 86ms | 466ms | 178ms | SheetKit |

### Write (DV)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 9ms | 69ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 2000 rows with comments | 7ms | 47ms | 81ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 500 merged regions | 2ms | 14ms | 1ms | SheetJS |

### Write (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 1k rows x 10 cols | 4ms | 25ms | 6ms | SheetKit |
| Write 10k rows x 10 cols | 46ms | 239ms | 67ms | SheetKit |
| Write 50k rows x 10 cols | 231ms | 1.26s | 452ms | SheetKit |
| Write 100k rows x 10 cols | 485ms | 2.85s | 1.12s | SheetKit |

### Round-Trip

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 123ms | 319ms | 163ms | SheetKit |

### Streaming

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Streaming write (50000 rows) | 313ms | 450ms | N/A | SheetKit |
| Streaming write (50000 rows) (writeRows) | 306ms | N/A | N/A | - |

### Random Access

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Random-access (open+1000 lookups) | 453ms | 1.27s | 1.34s | SheetKit |
| Random-access (open+1000 lookups) (sync) | 451ms | N/A | N/A | - |
| Random-access (lookup-only, 1000 cells) | 457ms | 1.27s | 1.30s | SheetKit |
| Random-access (lookup-only, 1000 cells) (sync) | 455ms | N/A | N/A | - |

### Mixed Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Mixed workload write (ERP-style) | 18ms | 83ms | N/A | SheetKit |

### COW Save

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Copy-on-write save (untouched) (lazy) | 217ms | N/A | N/A | - |
| Copy-on-write save (untouched) (eager) | 661ms | N/A | N/A | - |
| Copy-on-write save (single-cell edit) (lazy) | 670ms | N/A | N/A | - |
| Copy-on-write save (single-cell edit) (eager) | 688ms | N/A | N/A | - |

### Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | RSS (median) | Heap (median) |
|----------|---------|--------|-----|-----|-----|--------------|---------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 541ms | 527ms | 547ms | 547ms | 17.4MB | 58.7MB |
| Read Large Data (50k rows x 20 cols) | ExcelJS | 1.24s | 1.17s | 1.26s | 1.26s | 11.3MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) | SheetJS | 1.56s | 1.49s | 1.58s | 1.58s | 0.0MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) (sync) | SheetKit | 525ms | 522ms | 534ms | 534ms | 190.4MB | 58.8MB |
| Read Large Data (50k rows x 20 cols) (getRowsRaw) | SheetKit | 530ms | 524ms | 545ms | 545ms | 0.0MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) (bufferV2) | SheetKit | 491ms | 481ms | 507ms | 507ms | 14.3MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) (stream) | SheetKit | 696ms | 694ms | 697ms | 697ms | 17.2MB | 10.1MB |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 27ms | 26ms | 27ms | 27ms | 5.9MB | 6.9MB |
| Read Heavy Styles (5k rows, formatted) | ExcelJS | 67ms | 65ms | 71ms | 71ms | 2.3MB | 27.9MB |
| Read Heavy Styles (5k rows, formatted) | SheetJS | 70ms | 69ms | 72ms | 72ms | 5.3MB | 1.0MB |
| Read Heavy Styles (5k rows, formatted) (sync) | SheetKit | 27ms | 26ms | 28ms | 28ms | 8.6MB | 6.9MB |
| Read Heavy Styles (5k rows, formatted) (getRowsRaw) | SheetKit | 27ms | 26ms | 27ms | 27ms | 0.7MB | 3.3MB |
| Read Heavy Styles (5k rows, formatted) (bufferV2) | SheetKit | 24ms | 24ms | 24ms | 24ms | 0.0MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (stream) | SheetKit | 35ms | 35ms | 35ms | 35ms | 0.0MB | 7.6MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 290ms | 289ms | 291ms | 291ms | 3.3MB | 13.4MB |
| Read Multi-Sheet (10 sheets x 5k rows) | ExcelJS | 613ms | 604ms | 632ms | 632ms | 0.0MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetJS | 602ms | 590ms | 611ms | 611ms | 0.0MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (sync) | SheetKit | 294ms | 291ms | 296ms | 296ms | 62.5MB | 13.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (getRowsRaw) | SheetKit | 292ms | 291ms | 300ms | 300ms | 0.0MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (bufferV2) | SheetKit | 272ms | 269ms | 290ms | 290ms | 10.2MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (stream) | SheetKit | 397ms | 392ms | 398ms | 398ms | 0.0MB | 4.4MB |
| Read Formulas (10k rows) | SheetKit | 34ms | 33ms | 34ms | 34ms | 8.8MB | 14.9MB |
| Read Formulas (10k rows) | ExcelJS | 79ms | 75ms | 101ms | 101ms | 2.7MB | 32.6MB |
| Read Formulas (10k rows) | SheetJS | 68ms | 67ms | 69ms | 69ms | 1.4MB | 43.4MB |
| Read Formulas (10k rows) (sync) | SheetKit | 35ms | 34ms | 35ms | 35ms | 9.3MB | 14.9MB |
| Read Formulas (10k rows) (getRowsRaw) | SheetKit | 33ms | 33ms | 34ms | 34ms | 0.5MB | 9.5MB |
| Read Formulas (10k rows) (bufferV2) | SheetKit | 29ms | 29ms | 30ms | 30ms | 0.0MB | 0.0MB |
| Read Formulas (10k rows) (stream) | SheetKit | 45ms | 45ms | 49ms | 49ms | 2.3MB | 10.9MB |
| Read Strings (20k rows text-heavy) | SheetKit | 117ms | 116ms | 121ms | 121ms | 6.9MB | 0.0MB |
| Read Strings (20k rows text-heavy) | ExcelJS | 232ms | 230ms | 241ms | 241ms | 10.5MB | 66.9MB |
| Read Strings (20k rows text-heavy) | SheetJS | 267ms | 242ms | 284ms | 284ms | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) (sync) | SheetKit | 117ms | 117ms | 119ms | 119ms | 13.3MB | 22.7MB |
| Read Strings (20k rows text-heavy) (getRowsRaw) | SheetKit | 116ms | 115ms | 117ms | 117ms | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) (bufferV2) | SheetKit | 99ms | 99ms | 99ms | 99ms | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) (stream) | SheetKit | 150ms | 149ms | 150ms | 150ms | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | SheetKit | 21ms | 21ms | 22ms | 22ms | 0.0MB | 6.7MB |
| Read Data Validation (5k rows, 8 rules) | ExcelJS | 53ms | 51ms | 57ms | 57ms | 3.1MB | 4.1MB |
| Read Data Validation (5k rows, 8 rules) | SheetJS | 45ms | 44ms | 46ms | 46ms | 2.7MB | 10.1MB |
| Read Data Validation (5k rows, 8 rules) (sync) | SheetKit | 21ms | 21ms | 22ms | 22ms | 0.0MB | 6.7MB |
| Read Data Validation (5k rows, 8 rules) (getRowsRaw) | SheetKit | 21ms | 21ms | 22ms | 22ms | 0.3MB | 3.7MB |
| Read Data Validation (5k rows, 8 rules) (bufferV2) | SheetKit | 19ms | 19ms | 19ms | 19ms | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (stream) | SheetKit | 28ms | 28ms | 29ms | 29ms | 0.0MB | 5.9MB |
| Read Comments (2k rows with comments) | SheetKit | 8ms | 8ms | 9ms | 9ms | 0.6MB | 1.9MB |
| Read Comments (2k rows with comments) | ExcelJS | 32ms | 32ms | 35ms | 35ms | 1.9MB | 10.2MB |
| Read Comments (2k rows with comments) | SheetJS | 20ms | 20ms | 22ms | 22ms | 1.4MB | 30.3MB |
| Read Comments (2k rows with comments) (sync) | SheetKit | 8ms | 8ms | 8ms | 8ms | 0.6MB | 1.9MB |
| Read Comments (2k rows with comments) (getRowsRaw) | SheetKit | 8ms | 8ms | 8ms | 8ms | 0.7MB | 1.0MB |
| Read Comments (2k rows with comments) (bufferV2) | SheetKit | 7ms | 7ms | 8ms | 8ms | 0.6MB | 0.0MB |
| Read Comments (2k rows with comments) (stream) | SheetKit | 8ms | 8ms | 9ms | 9ms | 0.0MB | 1.5MB |
| Read Merged Cells (500 regions) | SheetKit | 1ms | 1ms | 2ms | 2ms | 0.0MB | 0.4MB |
| Read Merged Cells (500 regions) | ExcelJS | 9ms | 9ms | 12ms | 12ms | 0.2MB | 10.2MB |
| Read Merged Cells (500 regions) | SheetJS | 3ms | 3ms | 3ms | 3ms | 0.2MB | 4.6MB |
| Read Merged Cells (500 regions) (sync) | SheetKit | 1ms | 1ms | 1ms | 1ms | 0.0MB | 0.4MB |
| Read Merged Cells (500 regions) (getRowsRaw) | SheetKit | 1ms | 1ms | 1ms | 1ms | 0.1MB | 0.1MB |
| Read Merged Cells (500 regions) (bufferV2) | SheetKit | 1ms | 1ms | 1ms | 1ms | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) (stream) | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB | 0.4MB |
| Read Mixed Workload (ERP document) | SheetKit | 29ms | 28ms | 30ms | 30ms | 0.0MB | 9.5MB |
| Read Mixed Workload (ERP document) | ExcelJS | 69ms | 66ms | 78ms | 78ms | 3.0MB | 31.3MB |
| Read Mixed Workload (ERP document) | SheetJS | 60ms | 60ms | 68ms | 68ms | 4.4MB | 0.0MB |
| Read Mixed Workload (ERP document) (sync) | SheetKit | 28ms | 28ms | 29ms | 29ms | 0.0MB | 9.5MB |
| Read Mixed Workload (ERP document) (getRowsRaw) | SheetKit | 28ms | 28ms | 29ms | 29ms | 0.5MB | 6.1MB |
| Read Mixed Workload (ERP document) (bufferV2) | SheetKit | 27ms | 26ms | 29ms | 29ms | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) (stream) | SheetKit | 39ms | 37ms | 45ms | 45ms | 0.0MB | 7.5MB |
| Read Scale 1k rows | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0MB | 1.2MB |
| Read Scale 1k rows | ExcelJS | 13ms | 12ms | 18ms | 18ms | 0.0MB | 0.0MB |
| Read Scale 1k rows | SheetJS | 11ms | 11ms | 13ms | 13ms | 0.9MB | 18.2MB |
| Read Scale 1k rows (sync) | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0MB | 1.2MB |
| Read Scale 1k rows (getRowsRaw) | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0MB | 0.5MB |
| Read Scale 1k rows (bufferV2) | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0MB | 0.0MB |
| Read Scale 1k rows (stream) | SheetKit | 7ms | 7ms | 7ms | 7ms | 0.0MB | 1.5MB |
| Read Scale 10k rows | SheetKit | 50ms | 49ms | 50ms | 50ms | 0.0MB | 12.3MB |
| Read Scale 10k rows | ExcelJS | 121ms | 117ms | 141ms | 141ms | 4.0MB | 11.4MB |
| Read Scale 10k rows | SheetJS | 118ms | 117ms | 122ms | 122ms | 11.6MB | 7.4MB |
| Read Scale 10k rows (sync) | SheetKit | 49ms | 49ms | 50ms | 50ms | 0.0MB | 12.3MB |
| Read Scale 10k rows (getRowsRaw) | SheetKit | 50ms | 49ms | 51ms | 51ms | 0.8MB | 5.0MB |
| Read Scale 10k rows (bufferV2) | SheetKit | 47ms | 46ms | 47ms | 47ms | 0.0MB | 0.0MB |
| Read Scale 10k rows (stream) | SheetKit | 70ms | 69ms | 77ms | 77ms | 0.0MB | 14.5MB |
| Read Scale 100k rows | SheetKit | 535ms | 534ms | 540ms | 540ms | 24.2MB | 85.2MB |
| Read Scale 100k rows | ExcelJS | 1.30s | 1.23s | 1.36s | 1.36s | 0.0MB | 25.2MB |
| Read Scale 100k rows | SheetJS | 1.60s | 1.56s | 1.61s | 1.61s | 0.0MB | 5.4MB |
| Read Scale 100k rows (sync) | SheetKit | 539ms | 534ms | 544ms | 544ms | 169.7MB | 84.1MB |
| Read Scale 100k rows (getRowsRaw) | SheetKit | 513ms | 511ms | 518ms | 518ms | 44.8MB | 0.0MB |
| Read Scale 100k rows (bufferV2) | SheetKit | 487ms | 485ms | 487ms | 487ms | 52.5MB | 0.0MB |
| Read Scale 100k rows (stream) | SheetKit | 713ms | 712ms | 716ms | 716ms | 0.0MB | 0.0MB |
| Write 50000 rows x 20 cols | SheetKit | 469ms | 465ms | 479ms | 479ms | 44.2MB | 37.6MB |
| Write 50000 rows x 20 cols | ExcelJS | 2.62s | 2.57s | 2.70s | 2.70s | 21.2MB | 0.0MB |
| Write 50000 rows x 20 cols | SheetJS | 1.09s | 1.08s | 1.44s | 1.44s | 0.0MB | 0.0MB |
| Write 5000 styled rows | SheetKit | 32ms | 32ms | 66ms | 66ms | 2.8MB | 3.4MB |
| Write 5000 styled rows | ExcelJS | 147ms | 145ms | 156ms | 156ms | 7.6MB | 69.4MB |
| Write 5000 styled rows | SheetJS | 34ms | 34ms | 36ms | 36ms | 3.0MB | 11.3MB |
| Write 10 sheets x 5000 rows | SheetKit | 237ms | 235ms | 244ms | 244ms | 40.8MB | 0.0MB |
| Write 10 sheets x 5000 rows | ExcelJS | 1.27s | 1.26s | 1.28s | 1.28s | 0.0MB | 0.0MB |
| Write 10 sheets x 5000 rows | SheetJS | 331ms | 327ms | 352ms | 352ms | 30.2MB | 45.2MB |
| Write 10000 rows with formulas | SheetKit | 27ms | 26ms | 38ms | 38ms | 2.5MB | 4.7MB |
| Write 10000 rows with formulas | ExcelJS | 133ms | 130ms | 136ms | 136ms | 19.7MB | 90.7MB |
| Write 10000 rows with formulas | SheetJS | 54ms | 52ms | 55ms | 55ms | 5.1MB | 0.0MB |
| Write 20000 text-heavy rows | SheetKit | 86ms | 86ms | 91ms | 91ms | 13.4MB | 21.6MB |
| Write 20000 text-heavy rows | ExcelJS | 466ms | 450ms | 481ms | 481ms | 75.2MB | 278.3MB |
| Write 20000 text-heavy rows | SheetJS | 178ms | 176ms | 210ms | 210ms | 22.5MB | 79.0MB |
| Write 5000 rows + 8 validation rules | SheetKit | 9ms | 9ms | 9ms | 9ms | 1.2MB | 0.8MB |
| Write 5000 rows + 8 validation rules | ExcelJS | 69ms | 68ms | 72ms | 72ms | 2.8MB | 26.3MB |
| Write 2000 rows with comments | SheetKit | 7ms | 7ms | 7ms | 7ms | 0.0MB | 0.7MB |
| Write 2000 rows with comments | ExcelJS | 47ms | 46ms | 48ms | 48ms | 2.5MB | 13.6MB |
| Write 2000 rows with comments | SheetJS | 81ms | 81ms | 82ms | 82ms | 0.8MB | 4.7MB |
| Write 500 merged regions | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB | 0.3MB |
| Write 500 merged regions | ExcelJS | 14ms | 14ms | 20ms | 20ms | 1.0MB | 15.4MB |
| Write 500 merged regions | SheetJS | 1ms | 1ms | 2ms | 2ms | 0.0MB | 4.0MB |
| Write 1k rows x 10 cols | SheetKit | 4ms | 4ms | 5ms | 5ms | 0.0MB | 0.5MB |
| Write 1k rows x 10 cols | ExcelJS | 25ms | 21ms | 27ms | 27ms | 0.0MB | 0.0MB |
| Write 1k rows x 10 cols | SheetJS | 6ms | 6ms | 9ms | 9ms | 1.2MB | 15.2MB |
| Write 10k rows x 10 cols | SheetKit | 46ms | 45ms | 47ms | 47ms | 0.0MB | 4.6MB |
| Write 10k rows x 10 cols | ExcelJS | 239ms | 234ms | 245ms | 245ms | 29.9MB | 122.5MB |
| Write 10k rows x 10 cols | SheetJS | 67ms | 66ms | 68ms | 68ms | 9.3MB | 32.2MB |
| Write 50k rows x 10 cols | SheetKit | 231ms | 228ms | 245ms | 245ms | 27.4MB | 23.6MB |
| Write 50k rows x 10 cols | ExcelJS | 1.26s | 1.25s | 1.34s | 1.34s | 170.9MB | 673.1MB |
| Write 50k rows x 10 cols | SheetJS | 452ms | 437ms | 550ms | 550ms | 41.8MB | 104.3MB |
| Write 100k rows x 10 cols | SheetKit | 485ms | 476ms | 490ms | 490ms | 9.0MB | 1.8MB |
| Write 100k rows x 10 cols | ExcelJS | 2.85s | 2.80s | 3.02s | 3.02s | 1.5MB | 0.0MB |
| Write 100k rows x 10 cols | SheetJS | 1.12s | 1.12s | 1.16s | 1.16s | 0.0MB | 0.0MB |
| Buffer round-trip (10000 rows) | SheetKit | 123ms | 122ms | 123ms | 123ms | 2.4MB | 11.1MB |
| Buffer round-trip (10000 rows) | ExcelJS | 319ms | 317ms | 367ms | 367ms | 29.5MB | 156.1MB |
| Buffer round-trip (10000 rows) | SheetJS | 163ms | 161ms | 193ms | 193ms | 28.3MB | 33.2MB |
| Streaming write (50000 rows) | SheetKit | 313ms | 309ms | 316ms | 316ms | 0.0MB | 0.0MB |
| Streaming write (50000 rows) | ExcelJS | 450ms | 449ms | 461ms | 461ms | 0.0MB | 0.0MB |
| Streaming write (50000 rows) (writeRows) | SheetKit | 306ms | 303ms | 326ms | 326ms | 0.0MB | 0.0MB |
| Random-access (open+1000 lookups) | SheetKit | 453ms | 450ms | 456ms | 456ms | 61.0MB | 0.1MB |
| Random-access (open+1000 lookups) | ExcelJS | 1.27s | 1.18s | 1.29s | 1.29s | 6.0MB | 70.6MB |
| Random-access (open+1000 lookups) | SheetJS | 1.34s | 1.28s | 1.37s | 1.37s | 23.0MB | 0.0MB |
| Random-access (open+1000 lookups) (sync) | SheetKit | 451ms | 449ms | 468ms | 468ms | 150.5MB | 0.1MB |
| Random-access (lookup-only, 1000 cells) | SheetKit | 457ms | 450ms | 477ms | 477ms | 4.2MB | 0.1MB |
| Random-access (lookup-only, 1000 cells) | ExcelJS | 1.27s | 1.18s | 1.28s | 1.28s | 0.0MB | 0.0MB |
| Random-access (lookup-only, 1000 cells) | SheetJS | 1.30s | 1.29s | 1.38s | 1.38s | 0.0MB | 0.0MB |
| Random-access (lookup-only, 1000 cells) (sync) | SheetKit | 455ms | 449ms | 457ms | 457ms | 53.3MB | 0.1MB |
| Mixed workload write (ERP-style) | SheetKit | 18ms | 18ms | 19ms | 19ms | 0.4MB | 1.6MB |
| Mixed workload write (ERP-style) | ExcelJS | 83ms | 78ms | 88ms | 88ms | 5.5MB | 26.1MB |
| Copy-on-write save (untouched) (lazy) | SheetKit | 217ms | 214ms | 219ms | 219ms | 34.8MB | 0.0MB |
| Copy-on-write save (untouched) (eager) | SheetKit | 661ms | 659ms | 677ms | 677ms | 12.3MB | 0.0MB |
| Copy-on-write save (single-cell edit) (lazy) | SheetKit | 670ms | 663ms | 681ms | 681ms | 69.7MB | 0.0MB |
| Copy-on-write save (single-cell edit) (eager) | SheetKit | 688ms | 671ms | 698ms | 698ms | 146.3MB | 0.0MB |

### Memory Usage (RSS / Heap Used)

RSS = Resident Set Size delta (total process memory). Heap = V8 heapUsed delta (JS-only memory).

| Scenario | SheetKit (RSS/Heap) | ExcelJS (RSS/Heap) | SheetJS (RSS/Heap) |
|----------|---------------------|--------------------|--------------------|
| Read Large Data (50k rows x 20 cols) | 17.4MB / 58.7MB | 11.3MB / 0.0MB | 0.0MB / 0.0MB |
| Read Large Data (50k rows x 20 cols) (sync) | 190.4MB / 58.8MB | N/A | N/A |
| Read Large Data (50k rows x 20 cols) (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Large Data (50k rows x 20 cols) (bufferV2) | 14.3MB / 0.0MB | N/A | N/A |
| Read Large Data (50k rows x 20 cols) (stream) | 17.2MB / 10.1MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) | 5.9MB / 6.9MB | 2.3MB / 27.9MB | 5.3MB / 1.0MB |
| Read Heavy Styles (5k rows, formatted) (sync) | 8.6MB / 6.9MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) (getRowsRaw) | 0.7MB / 3.3MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) (stream) | 0.0MB / 7.6MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) | 3.3MB / 13.4MB | 0.0MB / 0.0MB | 0.0MB / 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (sync) | 62.5MB / 13.0MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) (bufferV2) | 10.2MB / 0.0MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) (stream) | 0.0MB / 4.4MB | N/A | N/A |
| Read Formulas (10k rows) | 8.8MB / 14.9MB | 2.7MB / 32.6MB | 1.4MB / 43.4MB |
| Read Formulas (10k rows) (sync) | 9.3MB / 14.9MB | N/A | N/A |
| Read Formulas (10k rows) (getRowsRaw) | 0.5MB / 9.5MB | N/A | N/A |
| Read Formulas (10k rows) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Formulas (10k rows) (stream) | 2.3MB / 10.9MB | N/A | N/A |
| Read Strings (20k rows text-heavy) | 6.9MB / 0.0MB | 10.5MB / 66.9MB | 0.0MB / 0.0MB |
| Read Strings (20k rows text-heavy) (sync) | 13.3MB / 22.7MB | N/A | N/A |
| Read Strings (20k rows text-heavy) (getRowsRaw) | 0.0MB / 0.0MB | N/A | N/A |
| Read Strings (20k rows text-heavy) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Strings (20k rows text-heavy) (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) | 0.0MB / 6.7MB | 3.1MB / 4.1MB | 2.7MB / 10.1MB |
| Read Data Validation (5k rows, 8 rules) (sync) | 0.0MB / 6.7MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) (getRowsRaw) | 0.3MB / 3.7MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) (stream) | 0.0MB / 5.9MB | N/A | N/A |
| Read Comments (2k rows with comments) | 0.6MB / 1.9MB | 1.9MB / 10.2MB | 1.4MB / 30.3MB |
| Read Comments (2k rows with comments) (sync) | 0.6MB / 1.9MB | N/A | N/A |
| Read Comments (2k rows with comments) (getRowsRaw) | 0.7MB / 1.0MB | N/A | N/A |
| Read Comments (2k rows with comments) (bufferV2) | 0.6MB / 0.0MB | N/A | N/A |
| Read Comments (2k rows with comments) (stream) | 0.0MB / 1.5MB | N/A | N/A |
| Read Merged Cells (500 regions) | 0.0MB / 0.4MB | 0.2MB / 10.2MB | 0.2MB / 4.6MB |
| Read Merged Cells (500 regions) (sync) | 0.0MB / 0.4MB | N/A | N/A |
| Read Merged Cells (500 regions) (getRowsRaw) | 0.1MB / 0.1MB | N/A | N/A |
| Read Merged Cells (500 regions) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Merged Cells (500 regions) (stream) | 0.0MB / 0.4MB | N/A | N/A |
| Read Mixed Workload (ERP document) | 0.0MB / 9.5MB | 3.0MB / 31.3MB | 4.4MB / 0.0MB |
| Read Mixed Workload (ERP document) (sync) | 0.0MB / 9.5MB | N/A | N/A |
| Read Mixed Workload (ERP document) (getRowsRaw) | 0.5MB / 6.1MB | N/A | N/A |
| Read Mixed Workload (ERP document) (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Mixed Workload (ERP document) (stream) | 0.0MB / 7.5MB | N/A | N/A |
| Read Scale 1k rows | 0.0MB / 1.2MB | 0.0MB / 0.0MB | 0.9MB / 18.2MB |
| Read Scale 1k rows (sync) | 0.0MB / 1.2MB | N/A | N/A |
| Read Scale 1k rows (getRowsRaw) | 0.0MB / 0.5MB | N/A | N/A |
| Read Scale 1k rows (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 1k rows (stream) | 0.0MB / 1.5MB | N/A | N/A |
| Read Scale 10k rows | 0.0MB / 12.3MB | 4.0MB / 11.4MB | 11.6MB / 7.4MB |
| Read Scale 10k rows (sync) | 0.0MB / 12.3MB | N/A | N/A |
| Read Scale 10k rows (getRowsRaw) | 0.8MB / 5.0MB | N/A | N/A |
| Read Scale 10k rows (bufferV2) | 0.0MB / 0.0MB | N/A | N/A |
| Read Scale 10k rows (stream) | 0.0MB / 14.5MB | N/A | N/A |
| Read Scale 100k rows | 24.2MB / 85.2MB | 0.0MB / 25.2MB | 0.0MB / 5.4MB |
| Read Scale 100k rows (sync) | 169.7MB / 84.1MB | N/A | N/A |
| Read Scale 100k rows (getRowsRaw) | 44.8MB / 0.0MB | N/A | N/A |
| Read Scale 100k rows (bufferV2) | 52.5MB / 0.0MB | N/A | N/A |
| Read Scale 100k rows (stream) | 0.0MB / 0.0MB | N/A | N/A |
| Write 50000 rows x 20 cols | 44.2MB / 37.6MB | 21.2MB / 0.0MB | 0.0MB / 0.0MB |
| Write 5000 styled rows | 2.8MB / 3.4MB | 7.6MB / 69.4MB | 3.0MB / 11.3MB |
| Write 10 sheets x 5000 rows | 40.8MB / 0.0MB | 0.0MB / 0.0MB | 30.2MB / 45.2MB |
| Write 10000 rows with formulas | 2.5MB / 4.7MB | 19.7MB / 90.7MB | 5.1MB / 0.0MB |
| Write 20000 text-heavy rows | 13.4MB / 21.6MB | 75.2MB / 278.3MB | 22.5MB / 79.0MB |
| Write 5000 rows + 8 validation rules | 1.2MB / 0.8MB | 2.8MB / 26.3MB | N/A |
| Write 2000 rows with comments | 0.0MB / 0.7MB | 2.5MB / 13.6MB | 0.8MB / 4.7MB |
| Write 500 merged regions | 0.0MB / 0.3MB | 1.0MB / 15.4MB | 0.0MB / 4.0MB |
| Write 1k rows x 10 cols | 0.0MB / 0.5MB | 0.0MB / 0.0MB | 1.2MB / 15.2MB |
| Write 10k rows x 10 cols | 0.0MB / 4.6MB | 29.9MB / 122.5MB | 9.3MB / 32.2MB |
| Write 50k rows x 10 cols | 27.4MB / 23.6MB | 170.9MB / 673.1MB | 41.8MB / 104.3MB |
| Write 100k rows x 10 cols | 9.0MB / 1.8MB | 1.5MB / 0.0MB | 0.0MB / 0.0MB |
| Buffer round-trip (10000 rows) | 2.4MB / 11.1MB | 29.5MB / 156.1MB | 28.3MB / 33.2MB |
| Streaming write (50000 rows) | 0.0MB / 0.0MB | 0.0MB / 0.0MB | N/A |
| Streaming write (50000 rows) (writeRows) | 0.0MB / 0.0MB | N/A | N/A |
| Random-access (open+1000 lookups) | 61.0MB / 0.1MB | 6.0MB / 70.6MB | 23.0MB / 0.0MB |
| Random-access (open+1000 lookups) (sync) | 150.5MB / 0.1MB | N/A | N/A |
| Random-access (lookup-only, 1000 cells) | 4.2MB / 0.1MB | 0.0MB / 0.0MB | 0.0MB / 0.0MB |
| Random-access (lookup-only, 1000 cells) (sync) | 53.3MB / 0.1MB | N/A | N/A |
| Mixed workload write (ERP-style) | 0.4MB / 1.6MB | 5.5MB / 26.1MB | N/A |
| Copy-on-write save (untouched) (lazy) | 34.8MB / 0.0MB | N/A | N/A |
| Copy-on-write save (untouched) (eager) | 12.3MB / 0.0MB | N/A | N/A |
| Copy-on-write save (single-cell edit) (lazy) | 69.7MB / 0.0MB | N/A | N/A |
| Copy-on-write save (single-cell edit) (eager) | 146.3MB / 0.0MB | N/A | N/A |

## Summary

Contested scenarios (>= 2 libraries): 29

| Library | Wins |
|---------|------|
| SheetKit | 28/29 |
| SheetJS | 1/29 |
| ExcelJS | 0/29 |
