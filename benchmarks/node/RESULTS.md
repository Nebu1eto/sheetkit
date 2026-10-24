# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-10-24T06:03:05.996Z

## Environment

| Item    | Value                                                         |
| ------- | ------------------------------------------------------------- |
| CPU     | Apple M4 Pro                                                  |
| RAM     | 24 GB                                                         |
| OS      | darwin arm64                                                  |
| Node.js | v26.12.0                                                      |
| Rust    | rustc 1.99.1 (58cf8f9c3 2026-10-13) (SheetKit native backend) |

## Methodology

- **All libraries**: 1 warmup run(s) + 5 measured runs per scenario. Median time reported.
- **RSS (Resident Set Size)**: Total process memory delta measured before/after each run via `process.memoryUsage().rss`. Includes V8 heap, native (Rust/napi) allocations, and OS overhead. This is a post-operation residual measurement, not peak usage during the operation.
- **Heap Used**: V8 heap delta measured before/after each run via `process.memoryUsage().heapUsed`. Isolates JavaScript-side memory growth, excluding native allocations. Useful for comparing JS-only libraries against napi-rs libraries where native memory dominates.
- **GC**: When `--expose-gc` is enabled, `global.gc()` is called before each measurement to reduce noise from deferred garbage collection.
- **Limitations**: Both RSS and heapUsed measure post-operation residual, not peak. Actual peak memory during an operation may be higher due to intermediate allocations freed before measurement. For napi-rs libraries, most memory lives in native heap (visible in RSS but not heapUsed).

## Libraries

| Library                         | Description                                                |
| ------------------------------- | ---------------------------------------------------------- |
| **SheetKit** (`@sheetkit/node`) | Rust-based Excel library with Node.js bindings via napi-rs |
| **ExcelJS** (`exceljs`)         | Pure JavaScript Excel library with streaming support       |
| **SheetJS** (`xlsx`)            | Pure JavaScript spreadsheet library (community edition)    |

## Test Fixtures

| Fixture                    | Description                                                                   |
| -------------------------- | ----------------------------------------------------------------------------- |
| `large-data.xlsx`          | 50,000 rows x 20 columns, mixed types (numbers, strings, floats, booleans)    |
| `heavy-styles.xlsx`        | 5,000 rows x 10 columns with rich formatting                                  |
| `multi-sheet.xlsx`         | 10 sheets, each with 5,000 rows x 10 columns                                  |
| `formulas.xlsx`            | 10,000 rows with 5 formula columns                                            |
| `strings.xlsx`             | 20,000 rows x 10 columns of text data (SST stress test)                       |
| `data-validation.xlsx`     | 5,000 rows with 8 validation rules (list, whole, decimal, textLength, custom) |
| `comments.xlsx`            | 2,000 rows with cell comments (2,667 total comments)                          |
| `merged-cells.xlsx`        | 500 merged regions (section headers and sub-headers)                          |
| `mixed-workload.xlsx`      | Multi-sheet ERP document with styles, formulas, validation, comments          |
| `scale-{1k,10k,100k}.xlsx` | Scaling benchmarks at 1K, 10K, and 100K rows                                  |

## Results

### Read

| Scenario                                             | SheetKit | ExcelJS | SheetJS | Winner   |
| ---------------------------------------------------- | -------- | ------- | ------- | -------- |
| Read Large Data (50k rows x 20 cols)                 | 455ms    | 2.80s   | 1.82s   | SheetKit |
| Read Large Data (50k rows x 20 cols) (sync)          | 458ms    | N/A     | N/A     | -        |
| Read Large Data (50k rows x 20 cols) (getRowsRaw)    | 459ms    | N/A     | N/A     | -        |
| Read Large Data (50k rows x 20 cols) (bufferV2)      | 422ms    | N/A     | N/A     | -        |
| Read Large Data (50k rows x 20 cols) (stream)        | 394ms    | N/A     | N/A     | -        |
| Read Heavy Styles (5k rows, formatted)               | 24ms     | 170ms   | 99ms    | SheetKit |
| Read Heavy Styles (5k rows, formatted) (sync)        | 24ms     | N/A     | N/A     | -        |
| Read Heavy Styles (5k rows, formatted) (getRowsRaw)  | 26ms     | N/A     | N/A     | -        |
| Read Heavy Styles (5k rows, formatted) (bufferV2)    | 23ms     | N/A     | N/A     | -        |
| Read Heavy Styles (5k rows, formatted) (stream)      | 24ms     | N/A     | N/A     | -        |
| Read Multi-Sheet (10 sheets x 5k rows)               | 263ms    | 1.54s   | 766ms   | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) (sync)        | 270ms    | N/A     | N/A     | -        |
| Read Multi-Sheet (10 sheets x 5k rows) (getRowsRaw)  | 275ms    | N/A     | N/A     | -        |
| Read Multi-Sheet (10 sheets x 5k rows) (bufferV2)    | 248ms    | N/A     | N/A     | -        |
| Read Multi-Sheet (10 sheets x 5k rows) (stream)      | 244ms    | N/A     | N/A     | -        |
| Read Formulas (10k rows)                             | 29ms     | 192ms   | 103ms   | SheetKit |
| Read Formulas (10k rows) (sync)                      | 29ms     | N/A     | N/A     | -        |
| Read Formulas (10k rows) (getRowsRaw)                | 34ms     | N/A     | N/A     | -        |
| Read Formulas (10k rows) (bufferV2)                  | 26ms     | N/A     | N/A     | -        |
| Read Formulas (10k rows) (stream)                    | 32ms     | N/A     | N/A     | -        |
| Read Strings (20k rows text-heavy)                   | 102ms    | 606ms   | 340ms   | SheetKit |
| Read Strings (20k rows text-heavy) (sync)            | 102ms    | N/A     | N/A     | -        |
| Read Strings (20k rows text-heavy) (getRowsRaw)      | 109ms    | N/A     | N/A     | -        |
| Read Strings (20k rows text-heavy) (bufferV2)        | 88ms     | N/A     | N/A     | -        |
| Read Strings (20k rows text-heavy) (stream)          | 95ms     | N/A     | N/A     | -        |
| Read Data Validation (5k rows, 8 rules)              | 19ms     | 132ms   | 69ms    | SheetKit |
| Read Data Validation (5k rows, 8 rules) (sync)       | 19ms     | N/A     | N/A     | -        |
| Read Data Validation (5k rows, 8 rules) (getRowsRaw) | 21ms     | N/A     | N/A     | -        |
| Read Data Validation (5k rows, 8 rules) (bufferV2)   | 18ms     | N/A     | N/A     | -        |
| Read Data Validation (5k rows, 8 rules) (stream)     | 20ms     | N/A     | N/A     | -        |
| Read Comments (2k rows with comments)                | 7ms      | 110ms   | 45ms    | SheetKit |
| Read Comments (2k rows with comments) (sync)         | 7ms      | N/A     | N/A     | -        |
| Read Comments (2k rows with comments) (getRowsRaw)   | 7ms      | N/A     | N/A     | -        |
| Read Comments (2k rows with comments) (bufferV2)     | 6ms      | N/A     | N/A     | -        |
| Read Comments (2k rows with comments) (stream)       | 7ms      | N/A     | N/A     | -        |
| Read Merged Cells (500 regions)                      | 2ms      | 21ms    | 6ms     | SheetKit |
| Read Merged Cells (500 regions) (sync)               | 1ms      | N/A     | N/A     | -        |
| Read Merged Cells (500 regions) (getRowsRaw)         | 2ms      | N/A     | N/A     | -        |
| Read Merged Cells (500 regions) (bufferV2)           | 1ms      | N/A     | N/A     | -        |
| Read Merged Cells (500 regions) (stream)             | 2ms      | N/A     | N/A     | -        |
| Read Mixed Workload (ERP document)                   | 24ms     | 188ms   | 96ms    | SheetKit |
| Read Mixed Workload (ERP document) (sync)            | 24ms     | N/A     | N/A     | -        |
| Read Mixed Workload (ERP document) (getRowsRaw)      | 27ms     | N/A     | N/A     | -        |
| Read Mixed Workload (ERP document) (bufferV2)        | 22ms     | N/A     | N/A     | -        |
| Read Mixed Workload (ERP document) (stream)          | 26ms     | N/A     | N/A     | -        |

### Read (Scale)

| Scenario                          | SheetKit | ExcelJS | SheetJS | Winner   |
| --------------------------------- | -------- | ------- | ------- | -------- |
| Read Scale 1k rows                | 5ms      | 40ms    | 21ms    | SheetKit |
| Read Scale 1k rows (sync)         | 5ms      | N/A     | N/A     | -        |
| Read Scale 1k rows (getRowsRaw)   | 5ms      | N/A     | N/A     | -        |
| Read Scale 1k rows (bufferV2)     | 5ms      | N/A     | N/A     | -        |
| Read Scale 1k rows (stream)       | 6ms      | N/A     | N/A     | -        |
| Read Scale 10k rows               | 45ms     | 290ms   | 160ms   | SheetKit |
| Read Scale 10k rows (sync)        | 46ms     | N/A     | N/A     | -        |
| Read Scale 10k rows (getRowsRaw)  | 49ms     | N/A     | N/A     | -        |
| Read Scale 10k rows (bufferV2)    | 43ms     | N/A     | N/A     | -        |
| Read Scale 10k rows (stream)      | 42ms     | N/A     | N/A     | -        |
| Read Scale 100k rows              | 470ms    | 2.91s   | 1.84s   | SheetKit |
| Read Scale 100k rows (sync)       | 489ms    | N/A     | N/A     | -        |
| Read Scale 100k rows (getRowsRaw) | 475ms    | N/A     | N/A     | -        |
| Read Scale 100k rows (bufferV2)   | 446ms    | N/A     | N/A     | -        |
| Read Scale 100k rows (stream)     | 400ms    | N/A     | N/A     | -        |

### Write

| Scenario                       | SheetKit | ExcelJS | SheetJS | Winner   |
| ------------------------------ | -------- | ------- | ------- | -------- |
| Write 50000 rows x 20 cols     | 474ms    | 2.68s   | 1.15s   | SheetKit |
| Write 5000 styled rows         | 38ms     | 163ms   | 42ms    | SheetKit |
| Write 10 sheets x 5000 rows    | 249ms    | 1.36s   | 409ms   | SheetKit |
| Write 10000 rows with formulas | 34ms     | 157ms   | 59ms    | SheetKit |
| Write 20000 text-heavy rows    | 87ms     | 482ms   | 215ms   | SheetKit |

### Write (DV)

| Scenario                             | SheetKit | ExcelJS | SheetJS | Winner   |
| ------------------------------------ | -------- | ------- | ------- | -------- |
| Write 5000 rows + 8 validation rules | 10ms     | 92ms    | N/A     | SheetKit |

### Write (Comments)

| Scenario                      | SheetKit | ExcelJS | SheetJS | Winner   |
| ----------------------------- | -------- | ------- | ------- | -------- |
| Write 2000 rows with comments | 8ms      | 65ms    | 10ms    | SheetKit |

### Write (Merge)

| Scenario                 | SheetKit | ExcelJS | SheetJS | Winner   |
| ------------------------ | -------- | ------- | ------- | -------- |
| Write 500 merged regions | 2ms      | 27ms    | 3ms     | SheetKit |

### Write (Scale)

| Scenario                  | SheetKit | ExcelJS | SheetJS | Winner   |
| ------------------------- | -------- | ------- | ------- | -------- |
| Write 1k rows x 10 cols   | 5ms      | 39ms    | 8ms     | SheetKit |
| Write 10k rows x 10 cols  | 45ms     | 260ms   | 87ms    | SheetKit |
| Write 50k rows x 10 cols  | 225ms    | 1.39s   | 544ms   | SheetKit |
| Write 100k rows x 10 cols | 502ms    | 2.79s   | 1.26s   | SheetKit |

### Round-Trip

| Scenario                       | SheetKit | ExcelJS | SheetJS | Winner   |
| ------------------------------ | -------- | ------- | ------- | -------- |
| Buffer round-trip (10000 rows) | 122ms    | 511ms   | 196ms   | SheetKit |

### Streaming

| Scenario                                 | SheetKit | ExcelJS | SheetJS | Winner   |
| ---------------------------------------- | -------- | ------- | ------- | -------- |
| Streaming write (50000 rows)             | 280ms    | 553ms   | N/A     | SheetKit |
| Streaming write (50000 rows) (writeRows) | 303ms    | N/A     | N/A     | -        |

### Random Access

| Scenario                                       | SheetKit | ExcelJS | SheetJS | Winner   |
| ---------------------------------------------- | -------- | ------- | ------- | -------- |
| Random-access (open+1000 lookups)              | 415ms    | 2.87s   | 1.62s   | SheetKit |
| Random-access (open+1000 lookups) (sync)       | 434ms    | N/A     | N/A     | -        |
| Random-access (lookup-only, 1000 cells)        | 416ms    | 2.83s   | 1.57s   | SheetKit |
| Random-access (lookup-only, 1000 cells) (sync) | 425ms    | N/A     | N/A     | -        |

### Mixed Write

| Scenario                         | SheetKit | ExcelJS | SheetJS | Winner   |
| -------------------------------- | -------- | ------- | ------- | -------- |
| Mixed workload write (ERP-style) | 22ms     | 103ms   | N/A     | SheetKit |

### COW Save

| Scenario                                      | SheetKit | ExcelJS | SheetJS | Winner |
| --------------------------------------------- | -------- | ------- | ------- | ------ |
| Copy-on-write save (untouched) (lazy)         | 203ms    | N/A     | N/A     | -      |
| Copy-on-write save (untouched) (eager)        | 536ms    | N/A     | N/A     | -      |
| Copy-on-write save (single-cell edit) (lazy)  | 614ms    | N/A     | N/A     | -      |
| Copy-on-write save (single-cell edit) (eager) | 620ms    | N/A     | N/A     | -      |

### Detailed Statistics

| Scenario                                             | Library  | Median | Min   | Max   | P95   | RSS (median) | Heap (median) |
| ---------------------------------------------------- | -------- | ------ | ----- | ----- | ----- | ------------ | ------------- |
| Read Large Data (50k rows x 20 cols)                 | SheetKit | 455ms  | 440ms | 475ms | 475ms | 24.3MB       | 0.0MB         |
| Read Large Data (50k rows x 20 cols)                 | ExcelJS  | 2.80s  | 2.73s | 2.85s | 2.85s | 23.4MB       | 0.1MB         |
| Read Large Data (50k rows x 20 cols)                 | SheetJS  | 1.82s  | 1.79s | 1.87s | 1.87s | 1.1MB        | 0.0MB         |
| Read Large Data (50k rows x 20 cols) (sync)          | SheetKit | 458ms  | 452ms | 475ms | 475ms | 183.1MB      | 0.0MB         |
| Read Large Data (50k rows x 20 cols) (getRowsRaw)    | SheetKit | 459ms  | 454ms | 464ms | 464ms | 25.2MB       | 0.0MB         |
| Read Large Data (50k rows x 20 cols) (bufferV2)      | SheetKit | 422ms  | 419ms | 430ms | 430ms | 24.2MB       | 0.0MB         |
| Read Large Data (50k rows x 20 cols) (stream)        | SheetKit | 394ms  | 392ms | 396ms | 396ms | 25.5MB       | 0.0MB         |
| Read Heavy Styles (5k rows, formatted)               | SheetKit | 24ms   | 24ms  | 25ms  | 25ms  | 0.0MB        | 0.0MB         |
| Read Heavy Styles (5k rows, formatted)               | ExcelJS  | 170ms  | 167ms | 174ms | 174ms | 0.2MB        | 0.1MB         |
| Read Heavy Styles (5k rows, formatted)               | SheetJS  | 99ms   | 95ms  | 106ms | 106ms | 0.0MB        | 0.0MB         |
| Read Heavy Styles (5k rows, formatted) (sync)        | SheetKit | 24ms   | 24ms  | 24ms  | 24ms  | 5.5MB        | 0.0MB         |
| Read Heavy Styles (5k rows, formatted) (getRowsRaw)  | SheetKit | 26ms   | 26ms  | 27ms  | 27ms  | 0.0MB        | 0.0MB         |
| Read Heavy Styles (5k rows, formatted) (bufferV2)    | SheetKit | 23ms   | 23ms  | 23ms  | 23ms  | 0.0MB        | 0.0MB         |
| Read Heavy Styles (5k rows, formatted) (stream)      | SheetKit | 24ms   | 24ms  | 25ms  | 25ms  | 0.0MB        | 0.0MB         |
| Read Multi-Sheet (10 sheets x 5k rows)               | SheetKit | 263ms  | 260ms | 266ms | 266ms | 23.8MB       | 0.0MB         |
| Read Multi-Sheet (10 sheets x 5k rows)               | ExcelJS  | 1.54s  | 1.52s | 1.60s | 1.60s | 0.0MB        | 0.1MB         |
| Read Multi-Sheet (10 sheets x 5k rows)               | SheetJS  | 766ms  | 753ms | 780ms | 780ms | 0.0MB        | 0.1MB         |
| Read Multi-Sheet (10 sheets x 5k rows) (sync)        | SheetKit | 270ms  | 266ms | 270ms | 270ms | 88.4MB       | 0.0MB         |
| Read Multi-Sheet (10 sheets x 5k rows) (getRowsRaw)  | SheetKit | 275ms  | 269ms | 285ms | 285ms | 2.6MB        | 0.0MB         |
| Read Multi-Sheet (10 sheets x 5k rows) (bufferV2)    | SheetKit | 248ms  | 245ms | 249ms | 249ms | 0.0MB        | 0.0MB         |
| Read Multi-Sheet (10 sheets x 5k rows) (stream)      | SheetKit | 244ms  | 240ms | 248ms | 248ms | 0.0MB        | 0.0MB         |
| Read Formulas (10k rows)                             | SheetKit | 29ms   | 29ms  | 30ms  | 30ms  | 0.0MB        | 0.0MB         |
| Read Formulas (10k rows)                             | ExcelJS  | 192ms  | 189ms | 193ms | 193ms | 0.3MB        | 0.0MB         |
| Read Formulas (10k rows)                             | SheetJS  | 103ms  | 96ms  | 105ms | 105ms | 0.0MB        | 0.0MB         |
| Read Formulas (10k rows) (sync)                      | SheetKit | 29ms   | 29ms  | 29ms  | 29ms  | 9.2MB        | 0.0MB         |
| Read Formulas (10k rows) (getRowsRaw)                | SheetKit | 34ms   | 31ms  | 35ms  | 35ms  | 0.0MB        | 0.0MB         |
| Read Formulas (10k rows) (bufferV2)                  | SheetKit | 26ms   | 26ms  | 27ms  | 27ms  | 0.0MB        | 0.0MB         |
| Read Formulas (10k rows) (stream)                    | SheetKit | 32ms   | 31ms  | 33ms  | 33ms  | 0.0MB        | 0.0MB         |
| Read Strings (20k rows text-heavy)                   | SheetKit | 102ms  | 100ms | 104ms | 104ms | 3.5MB        | 0.0MB         |
| Read Strings (20k rows text-heavy)                   | ExcelJS  | 606ms  | 595ms | 621ms | 621ms | 0.3MB        | 0.0MB         |
| Read Strings (20k rows text-heavy)                   | SheetJS  | 340ms  | 337ms | 355ms | 355ms | 0.0MB        | 0.0MB         |
| Read Strings (20k rows text-heavy) (sync)            | SheetKit | 102ms  | 100ms | 103ms | 103ms | 12.5MB       | 0.0MB         |
| Read Strings (20k rows text-heavy) (getRowsRaw)      | SheetKit | 109ms  | 106ms | 111ms | 111ms | 3.5MB        | 0.0MB         |
| Read Strings (20k rows text-heavy) (bufferV2)        | SheetKit | 88ms   | 88ms  | 89ms  | 89ms  | 3.5MB        | 0.0MB         |
| Read Strings (20k rows text-heavy) (stream)          | SheetKit | 95ms   | 94ms  | 96ms  | 96ms  | 3.4MB        | 0.0MB         |
| Read Data Validation (5k rows, 8 rules)              | SheetKit | 19ms   | 19ms  | 20ms  | 20ms  | 0.0MB        | 0.0MB         |
| Read Data Validation (5k rows, 8 rules)              | ExcelJS  | 132ms  | 131ms | 133ms | 133ms | 3.6MB        | 16.1MB        |
| Read Data Validation (5k rows, 8 rules)              | SheetJS  | 69ms   | 61ms  | 72ms  | 72ms  | 0.0MB        | 0.0MB         |
| Read Data Validation (5k rows, 8 rules) (sync)       | SheetKit | 19ms   | 19ms  | 19ms  | 19ms  | 0.0MB        | 0.0MB         |
| Read Data Validation (5k rows, 8 rules) (getRowsRaw) | SheetKit | 21ms   | 20ms  | 22ms  | 22ms  | 0.0MB        | 0.0MB         |
| Read Data Validation (5k rows, 8 rules) (bufferV2)   | SheetKit | 18ms   | 18ms  | 18ms  | 18ms  | 0.0MB        | 0.0MB         |
| Read Data Validation (5k rows, 8 rules) (stream)     | SheetKit | 20ms   | 17ms  | 20ms  | 20ms  | 0.0MB        | 0.0MB         |
| Read Comments (2k rows with comments)                | SheetKit | 7ms    | 6ms   | 7ms   | 7ms   | 0.0MB        | 0.0MB         |
| Read Comments (2k rows with comments)                | ExcelJS  | 110ms  | 107ms | 115ms | 115ms | 0.0MB        | 0.1MB         |
| Read Comments (2k rows with comments)                | SheetJS  | 45ms   | 44ms  | 52ms  | 52ms  | 0.1MB        | 0.0MB         |
| Read Comments (2k rows with comments) (sync)         | SheetKit | 7ms    | 6ms   | 7ms   | 7ms   | 0.0MB        | 0.0MB         |
| Read Comments (2k rows with comments) (getRowsRaw)   | SheetKit | 7ms    | 7ms   | 7ms   | 7ms   | 0.0MB        | 0.0MB         |
| Read Comments (2k rows with comments) (bufferV2)     | SheetKit | 6ms    | 6ms   | 6ms   | 6ms   | 0.0MB        | 0.0MB         |
| Read Comments (2k rows with comments) (stream)       | SheetKit | 7ms    | 7ms   | 7ms   | 7ms   | 0.0MB        | 0.0MB         |
| Read Merged Cells (500 regions)                      | SheetKit | 2ms    | 2ms   | 2ms   | 2ms   | 0.0MB        | 0.0MB         |
| Read Merged Cells (500 regions)                      | ExcelJS  | 21ms   | 20ms  | 23ms  | 23ms  | 0.0MB        | 0.0MB         |
| Read Merged Cells (500 regions)                      | SheetJS  | 6ms    | 4ms   | 6ms   | 6ms   | 0.0MB        | 0.0MB         |
| Read Merged Cells (500 regions) (sync)               | SheetKit | 1ms    | 1ms   | 2ms   | 2ms   | 0.0MB        | 0.0MB         |
| Read Merged Cells (500 regions) (getRowsRaw)         | SheetKit | 2ms    | 2ms   | 2ms   | 2ms   | 0.0MB        | 0.0MB         |
| Read Merged Cells (500 regions) (bufferV2)           | SheetKit | 1ms    | 1ms   | 2ms   | 2ms   | 0.0MB        | 0.0MB         |
| Read Merged Cells (500 regions) (stream)             | SheetKit | 2ms    | 2ms   | 2ms   | 2ms   | 0.0MB        | 0.0MB         |
| Read Mixed Workload (ERP document)                   | SheetKit | 24ms   | 24ms  | 25ms  | 25ms  | 0.0MB        | 0.0MB         |
| Read Mixed Workload (ERP document)                   | ExcelJS  | 188ms  | 183ms | 198ms | 198ms | 0.5MB        | 0.1MB         |
| Read Mixed Workload (ERP document)                   | SheetJS  | 96ms   | 87ms  | 97ms  | 97ms  | 0.0MB        | 0.1MB         |
| Read Mixed Workload (ERP document) (sync)            | SheetKit | 24ms   | 24ms  | 24ms  | 24ms  | 0.0MB        | 0.0MB         |
| Read Mixed Workload (ERP document) (getRowsRaw)      | SheetKit | 27ms   | 26ms  | 28ms  | 28ms  | 0.0MB        | 0.0MB         |
| Read Mixed Workload (ERP document) (bufferV2)        | SheetKit | 22ms   | 22ms  | 22ms  | 22ms  | 0.0MB        | 0.0MB         |
| Read Mixed Workload (ERP document) (stream)          | SheetKit | 26ms   | 26ms  | 26ms  | 26ms  | 0.0MB        | 0.0MB         |
| Read Scale 1k rows                                   | SheetKit | 5ms    | 5ms   | 5ms   | 5ms   | 0.0MB        | 0.0MB         |
| Read Scale 1k rows                                   | ExcelJS  | 40ms   | 39ms  | 40ms  | 40ms  | 0.0MB        | 0.0MB         |
| Read Scale 1k rows                                   | SheetJS  | 21ms   | 19ms  | 24ms  | 24ms  | 0.0MB        | 0.0MB         |
| Read Scale 1k rows (sync)                            | SheetKit | 5ms    | 5ms   | 5ms   | 5ms   | 0.0MB        | 0.0MB         |
| Read Scale 1k rows (getRowsRaw)                      | SheetKit | 5ms    | 5ms   | 6ms   | 6ms   | 0.0MB        | 0.0MB         |
| Read Scale 1k rows (bufferV2)                        | SheetKit | 5ms    | 5ms   | 5ms   | 5ms   | 0.0MB        | 0.0MB         |
| Read Scale 1k rows (stream)                          | SheetKit | 6ms    | 5ms   | 6ms   | 6ms   | 0.0MB        | 0.0MB         |
| Read Scale 10k rows                                  | SheetKit | 45ms   | 44ms  | 45ms  | 45ms  | 0.0MB        | 0.0MB         |
| Read Scale 10k rows                                  | ExcelJS  | 290ms  | 286ms | 295ms | 295ms | 0.3MB        | 0.0MB         |
| Read Scale 10k rows                                  | SheetJS  | 160ms  | 158ms | 169ms | 169ms | 0.0MB        | 0.0MB         |
| Read Scale 10k rows (sync)                           | SheetKit | 46ms   | 46ms  | 47ms  | 47ms  | 2.1MB        | 0.0MB         |
| Read Scale 10k rows (getRowsRaw)                     | SheetKit | 49ms   | 48ms  | 50ms  | 50ms  | 0.0MB        | 0.0MB         |
| Read Scale 10k rows (bufferV2)                       | SheetKit | 43ms   | 43ms  | 43ms  | 43ms  | 0.0MB        | 0.0MB         |
| Read Scale 10k rows (stream)                         | SheetKit | 42ms   | 41ms  | 43ms  | 43ms  | 0.0MB        | 0.0MB         |
| Read Scale 100k rows                                 | SheetKit | 470ms  | 461ms | 471ms | 471ms | 35.9MB       | 0.0MB         |
| Read Scale 100k rows                                 | ExcelJS  | 2.91s  | 2.87s | 2.93s | 2.93s | 3.1MB        | 0.0MB         |
| Read Scale 100k rows                                 | SheetJS  | 1.84s  | 1.82s | 1.89s | 1.89s | 0.0MB        | 0.0MB         |
| Read Scale 100k rows (sync)                          | SheetKit | 489ms  | 473ms | 503ms | 503ms | 174.4MB      | 0.0MB         |
| Read Scale 100k rows (getRowsRaw)                    | SheetKit | 475ms  | 470ms | 480ms | 480ms | 30.8MB       | 0.0MB         |
| Read Scale 100k rows (bufferV2)                      | SheetKit | 446ms  | 443ms | 451ms | 451ms | 0.0MB        | 0.0MB         |
| Read Scale 100k rows (stream)                        | SheetKit | 400ms  | 398ms | 405ms | 405ms | 0.0MB        | 0.0MB         |
| Write 50000 rows x 20 cols                           | SheetKit | 474ms  | 464ms | 488ms | 488ms | 201.3MB      | 0.0MB         |
| Write 50000 rows x 20 cols                           | ExcelJS  | 2.68s  | 2.63s | 2.80s | 2.80s | 51.8MB       | 0.0MB         |
| Write 50000 rows x 20 cols                           | SheetJS  | 1.15s  | 1.14s | 1.19s | 1.19s | 0.0MB        | 0.0MB         |
| Write 5000 styled rows                               | SheetKit | 38ms   | 37ms  | 39ms  | 39ms  | 4.9MB        | 0.0MB         |
| Write 5000 styled rows                               | ExcelJS  | 163ms  | 159ms | 179ms | 179ms | 0.2MB        | 0.0MB         |
| Write 5000 styled rows                               | SheetJS  | 42ms   | 40ms  | 44ms  | 44ms  | 0.0MB        | 0.0MB         |
| Write 10 sheets x 5000 rows                          | SheetKit | 249ms  | 247ms | 256ms | 256ms | 75.2MB       | 0.0MB         |
| Write 10 sheets x 5000 rows                          | ExcelJS  | 1.36s  | 1.35s | 1.39s | 1.39s | 0.2MB        | 0.0MB         |
| Write 10 sheets x 5000 rows                          | SheetJS  | 409ms  | 406ms | 425ms | 425ms | 0.0MB        | 0.1MB         |
| Write 10000 rows with formulas                       | SheetKit | 34ms   | 34ms  | 35ms  | 35ms  | 7.5MB        | 0.0MB         |
| Write 10000 rows with formulas                       | ExcelJS  | 157ms  | 150ms | 163ms | 163ms | 0.0MB        | 0.0MB         |
| Write 10000 rows with formulas                       | SheetJS  | 59ms   | 53ms  | 62ms  | 62ms  | 0.0MB        | 0.0MB         |
| Write 20000 text-heavy rows                          | SheetKit | 87ms   | 86ms  | 87ms  | 87ms  | 13.8MB       | 0.0MB         |
| Write 20000 text-heavy rows                          | ExcelJS  | 482ms  | 481ms | 502ms | 502ms | 0.1MB        | 0.0MB         |
| Write 20000 text-heavy rows                          | SheetJS  | 215ms  | 205ms | 222ms | 222ms | 0.0MB        | 0.0MB         |
| Write 5000 rows + 8 validation rules                 | SheetKit | 10ms   | 10ms  | 11ms  | 11ms  | 1.2MB        | 0.0MB         |
| Write 5000 rows + 8 validation rules                 | ExcelJS  | 92ms   | 91ms  | 93ms  | 93ms  | 0.4MB        | 0.1MB         |
| Write 2000 rows with comments                        | SheetKit | 8ms    | 8ms   | 8ms   | 8ms   | 0.1MB        | 0.0MB         |
| Write 2000 rows with comments                        | ExcelJS  | 65ms   | 63ms  | 67ms  | 67ms  | 0.2MB        | 0.1MB         |
| Write 2000 rows with comments                        | SheetJS  | 10ms   | 9ms   | 10ms  | 10ms  | 0.0MB        | 0.0MB         |
| Write 500 merged regions                             | SheetKit | 2ms    | 2ms   | 3ms   | 3ms   | 0.0MB        | 0.0MB         |
| Write 500 merged regions                             | ExcelJS  | 27ms   | 19ms  | 29ms  | 29ms  | 0.3MB        | 0.1MB         |
| Write 500 merged regions                             | SheetJS  | 3ms    | 3ms   | 3ms   | 3ms   | 0.0MB        | 0.0MB         |
| Write 1k rows x 10 cols                              | SheetKit | 5ms    | 5ms   | 5ms   | 5ms   | 0.0MB        | 0.0MB         |
| Write 1k rows x 10 cols                              | ExcelJS  | 39ms   | 38ms  | 40ms  | 40ms  | 0.4MB        | 0.0MB         |
| Write 1k rows x 10 cols                              | SheetJS  | 8ms    | 7ms   | 9ms   | 9ms   | 0.3MB        | 0.0MB         |
| Write 10k rows x 10 cols                             | SheetKit | 45ms   | 44ms  | 47ms  | 47ms  | 1.8MB        | 0.0MB         |
| Write 10k rows x 10 cols                             | ExcelJS  | 260ms  | 258ms | 262ms | 262ms | 0.4MB        | 0.0MB         |
| Write 10k rows x 10 cols                             | SheetJS  | 87ms   | 83ms  | 90ms  | 90ms  | 0.0MB        | 0.0MB         |
| Write 50k rows x 10 cols                             | SheetKit | 225ms  | 222ms | 229ms | 229ms | 32.9MB       | 0.0MB         |
| Write 50k rows x 10 cols                             | ExcelJS  | 1.39s  | 1.36s | 1.39s | 1.39s | 10.2MB       | 0.0MB         |
| Write 50k rows x 10 cols                             | SheetJS  | 544ms  | 535ms | 547ms | 547ms | 0.0MB        | 0.0MB         |
| Write 100k rows x 10 cols                            | SheetKit | 502ms  | 494ms | 508ms | 508ms | 173.8MB      | 0.0MB         |
| Write 100k rows x 10 cols                            | ExcelJS  | 2.79s  | 2.78s | 2.81s | 2.81s | 6.9MB        | 0.0MB         |
| Write 100k rows x 10 cols                            | SheetJS  | 1.26s  | 1.23s | 1.41s | 1.41s | 12.1MB       | 0.0MB         |
| Buffer round-trip (10000 rows)                       | SheetKit | 122ms  | 119ms | 131ms | 131ms | 0.7MB        | 0.0MB         |
| Buffer round-trip (10000 rows)                       | ExcelJS  | 511ms  | 493ms | 564ms | 564ms | 2.4MB        | 0.0MB         |
| Buffer round-trip (10000 rows)                       | SheetJS  | 196ms  | 195ms | 203ms | 203ms | 0.0MB        | 0.1MB         |
| Streaming write (50000 rows)                         | SheetKit | 280ms  | 278ms | 287ms | 287ms | 0.0MB        | 0.0MB         |
| Streaming write (50000 rows)                         | ExcelJS  | 553ms  | 551ms | 559ms | 559ms | 1.1MB        | 0.0MB         |
| Streaming write (50000 rows) (writeRows)             | SheetKit | 303ms  | 301ms | 338ms | 338ms | 0.0MB        | 0.0MB         |
| Random-access (open+1000 lookups)                    | SheetKit | 415ms  | 411ms | 422ms | 422ms | 31.8MB       | 0.0MB         |
| Random-access (open+1000 lookups)                    | ExcelJS  | 2.87s  | 2.84s | 2.92s | 2.92s | 0.1MB        | 0.0MB         |
| Random-access (open+1000 lookups)                    | SheetJS  | 1.62s  | 1.59s | 1.62s | 1.62s | 0.9MB        | 0.0MB         |
| Random-access (open+1000 lookups) (sync)             | SheetKit | 434ms  | 429ms | 443ms | 443ms | 167.4MB      | 0.0MB         |
| Random-access (lookup-only, 1000 cells)              | SheetKit | 416ms  | 405ms | 435ms | 435ms | 24.2MB       | 0.0MB         |
| Random-access (lookup-only, 1000 cells)              | ExcelJS  | 2.83s  | 2.79s | 2.89s | 2.89s | 0.0MB        | 0.0MB         |
| Random-access (lookup-only, 1000 cells)              | SheetJS  | 1.57s  | 1.55s | 1.58s | 1.58s | 0.0MB        | 0.0MB         |
| Random-access (lookup-only, 1000 cells) (sync)       | SheetKit | 425ms  | 421ms | 444ms | 444ms | 148.5MB      | 0.0MB         |
| Mixed workload write (ERP-style)                     | SheetKit | 22ms   | 22ms  | 23ms  | 23ms  | 3.5MB        | 0.0MB         |
| Mixed workload write (ERP-style)                     | ExcelJS  | 103ms  | 102ms | 106ms | 106ms | 4.0MB        | 0.0MB         |
| Copy-on-write save (untouched) (lazy)                | SheetKit | 203ms  | 202ms | 232ms | 232ms | 31.2MB       | 0.0MB         |
| Copy-on-write save (untouched) (eager)               | SheetKit | 536ms  | 534ms | 544ms | 544ms | 31.3MB       | 0.0MB         |
| Copy-on-write save (single-cell edit) (lazy)         | SheetKit | 614ms  | 614ms | 620ms | 620ms | 77.9MB       | 0.0MB         |
| Copy-on-write save (single-cell edit) (eager)        | SheetKit | 620ms  | 610ms | 626ms | 626ms | 78.0MB       | 0.0MB         |

### Memory Usage (RSS / Heap Used)

RSS = Resident Set Size delta (total process memory). Heap = V8 heapUsed delta (JS-only memory).

| Scenario                                             | SheetKit (RSS/Heap) | ExcelJS (RSS/Heap) | SheetJS (RSS/Heap) |
| ---------------------------------------------------- | ------------------- | ------------------ | ------------------ |
| Read Large Data (50k rows x 20 cols)                 | 24.3MB / 0.0MB      | 23.4MB / 0.1MB     | 1.1MB / 0.0MB      |
| Read Large Data (50k rows x 20 cols) (sync)          | 183.1MB / 0.0MB     | N/A                | N/A                |
| Read Large Data (50k rows x 20 cols) (getRowsRaw)    | 25.2MB / 0.0MB      | N/A                | N/A                |
| Read Large Data (50k rows x 20 cols) (bufferV2)      | 24.2MB / 0.0MB      | N/A                | N/A                |
| Read Large Data (50k rows x 20 cols) (stream)        | 25.5MB / 0.0MB      | N/A                | N/A                |
| Read Heavy Styles (5k rows, formatted)               | 0.0MB / 0.0MB       | 0.2MB / 0.1MB      | 0.0MB / 0.0MB      |
| Read Heavy Styles (5k rows, formatted) (sync)        | 5.5MB / 0.0MB       | N/A                | N/A                |
| Read Heavy Styles (5k rows, formatted) (getRowsRaw)  | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Heavy Styles (5k rows, formatted) (bufferV2)    | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Heavy Styles (5k rows, formatted) (stream)      | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Multi-Sheet (10 sheets x 5k rows)               | 23.8MB / 0.0MB      | 0.0MB / 0.1MB      | 0.0MB / 0.1MB      |
| Read Multi-Sheet (10 sheets x 5k rows) (sync)        | 88.4MB / 0.0MB      | N/A                | N/A                |
| Read Multi-Sheet (10 sheets x 5k rows) (getRowsRaw)  | 2.6MB / 0.0MB       | N/A                | N/A                |
| Read Multi-Sheet (10 sheets x 5k rows) (bufferV2)    | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Multi-Sheet (10 sheets x 5k rows) (stream)      | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Formulas (10k rows)                             | 0.0MB / 0.0MB       | 0.3MB / 0.0MB      | 0.0MB / 0.0MB      |
| Read Formulas (10k rows) (sync)                      | 9.2MB / 0.0MB       | N/A                | N/A                |
| Read Formulas (10k rows) (getRowsRaw)                | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Formulas (10k rows) (bufferV2)                  | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Formulas (10k rows) (stream)                    | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Strings (20k rows text-heavy)                   | 3.5MB / 0.0MB       | 0.3MB / 0.0MB      | 0.0MB / 0.0MB      |
| Read Strings (20k rows text-heavy) (sync)            | 12.5MB / 0.0MB      | N/A                | N/A                |
| Read Strings (20k rows text-heavy) (getRowsRaw)      | 3.5MB / 0.0MB       | N/A                | N/A                |
| Read Strings (20k rows text-heavy) (bufferV2)        | 3.5MB / 0.0MB       | N/A                | N/A                |
| Read Strings (20k rows text-heavy) (stream)          | 3.4MB / 0.0MB       | N/A                | N/A                |
| Read Data Validation (5k rows, 8 rules)              | 0.0MB / 0.0MB       | 3.6MB / 16.1MB     | 0.0MB / 0.0MB      |
| Read Data Validation (5k rows, 8 rules) (sync)       | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Data Validation (5k rows, 8 rules) (getRowsRaw) | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Data Validation (5k rows, 8 rules) (bufferV2)   | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Data Validation (5k rows, 8 rules) (stream)     | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Comments (2k rows with comments)                | 0.0MB / 0.0MB       | 0.0MB / 0.1MB      | 0.1MB / 0.0MB      |
| Read Comments (2k rows with comments) (sync)         | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Comments (2k rows with comments) (getRowsRaw)   | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Comments (2k rows with comments) (bufferV2)     | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Comments (2k rows with comments) (stream)       | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Merged Cells (500 regions)                      | 0.0MB / 0.0MB       | 0.0MB / 0.0MB      | 0.0MB / 0.0MB      |
| Read Merged Cells (500 regions) (sync)               | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Merged Cells (500 regions) (getRowsRaw)         | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Merged Cells (500 regions) (bufferV2)           | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Merged Cells (500 regions) (stream)             | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Mixed Workload (ERP document)                   | 0.0MB / 0.0MB       | 0.5MB / 0.1MB      | 0.0MB / 0.1MB      |
| Read Mixed Workload (ERP document) (sync)            | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Mixed Workload (ERP document) (getRowsRaw)      | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Mixed Workload (ERP document) (bufferV2)        | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Mixed Workload (ERP document) (stream)          | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Scale 1k rows                                   | 0.0MB / 0.0MB       | 0.0MB / 0.0MB      | 0.0MB / 0.0MB      |
| Read Scale 1k rows (sync)                            | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Scale 1k rows (getRowsRaw)                      | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Scale 1k rows (bufferV2)                        | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Scale 1k rows (stream)                          | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Scale 10k rows                                  | 0.0MB / 0.0MB       | 0.3MB / 0.0MB      | 0.0MB / 0.0MB      |
| Read Scale 10k rows (sync)                           | 2.1MB / 0.0MB       | N/A                | N/A                |
| Read Scale 10k rows (getRowsRaw)                     | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Scale 10k rows (bufferV2)                       | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Scale 10k rows (stream)                         | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Scale 100k rows                                 | 35.9MB / 0.0MB      | 3.1MB / 0.0MB      | 0.0MB / 0.0MB      |
| Read Scale 100k rows (sync)                          | 174.4MB / 0.0MB     | N/A                | N/A                |
| Read Scale 100k rows (getRowsRaw)                    | 30.8MB / 0.0MB      | N/A                | N/A                |
| Read Scale 100k rows (bufferV2)                      | 0.0MB / 0.0MB       | N/A                | N/A                |
| Read Scale 100k rows (stream)                        | 0.0MB / 0.0MB       | N/A                | N/A                |
| Write 50000 rows x 20 cols                           | 201.3MB / 0.0MB     | 51.8MB / 0.0MB     | 0.0MB / 0.0MB      |
| Write 5000 styled rows                               | 4.9MB / 0.0MB       | 0.2MB / 0.0MB      | 0.0MB / 0.0MB      |
| Write 10 sheets x 5000 rows                          | 75.2MB / 0.0MB      | 0.2MB / 0.0MB      | 0.0MB / 0.1MB      |
| Write 10000 rows with formulas                       | 7.5MB / 0.0MB       | 0.0MB / 0.0MB      | 0.0MB / 0.0MB      |
| Write 20000 text-heavy rows                          | 13.8MB / 0.0MB      | 0.1MB / 0.0MB      | 0.0MB / 0.0MB      |
| Write 5000 rows + 8 validation rules                 | 1.2MB / 0.0MB       | 0.4MB / 0.1MB      | N/A                |
| Write 2000 rows with comments                        | 0.1MB / 0.0MB       | 0.2MB / 0.1MB      | 0.0MB / 0.0MB      |
| Write 500 merged regions                             | 0.0MB / 0.0MB       | 0.3MB / 0.1MB      | 0.0MB / 0.0MB      |
| Write 1k rows x 10 cols                              | 0.0MB / 0.0MB       | 0.4MB / 0.0MB      | 0.3MB / 0.0MB      |
| Write 10k rows x 10 cols                             | 1.8MB / 0.0MB       | 0.4MB / 0.0MB      | 0.0MB / 0.0MB      |
| Write 50k rows x 10 cols                             | 32.9MB / 0.0MB      | 10.2MB / 0.0MB     | 0.0MB / 0.0MB      |
| Write 100k rows x 10 cols                            | 173.8MB / 0.0MB     | 6.9MB / 0.0MB      | 12.1MB / 0.0MB     |
| Buffer round-trip (10000 rows)                       | 0.7MB / 0.0MB       | 2.4MB / 0.0MB      | 0.0MB / 0.1MB      |
| Streaming write (50000 rows)                         | 0.0MB / 0.0MB       | 1.1MB / 0.0MB      | N/A                |
| Streaming write (50000 rows) (writeRows)             | 0.0MB / 0.0MB       | N/A                | N/A                |
| Random-access (open+1000 lookups)                    | 31.8MB / 0.0MB      | 0.1MB / 0.0MB      | 0.9MB / 0.0MB      |
| Random-access (open+1000 lookups) (sync)             | 167.4MB / 0.0MB     | N/A                | N/A                |
| Random-access (lookup-only, 1000 cells)              | 24.2MB / 0.0MB      | 0.0MB / 0.0MB      | 0.0MB / 0.0MB      |
| Random-access (lookup-only, 1000 cells) (sync)       | 148.5MB / 0.0MB     | N/A                | N/A                |
| Mixed workload write (ERP-style)                     | 3.5MB / 0.0MB       | 4.0MB / 0.0MB      | N/A                |
| Copy-on-write save (untouched) (lazy)                | 31.2MB / 0.0MB      | N/A                | N/A                |
| Copy-on-write save (untouched) (eager)               | 31.3MB / 0.0MB      | N/A                | N/A                |
| Copy-on-write save (single-cell edit) (lazy)         | 77.9MB / 0.0MB      | N/A                | N/A                |
| Copy-on-write save (single-cell edit) (eager)        | 78.0MB / 0.0MB      | N/A                | N/A                |

## Summary

Contested scenarios (>= 2 libraries): 29

| Library  | Wins  |
| -------- | ----- |
| SheetKit | 29/29 |
| ExcelJS  | 0/29  |
| SheetJS  | 0/29  |
