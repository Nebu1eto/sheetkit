# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-09T13:13:56.418Z
Platform: darwin arm64
Node.js: v25.3.0

## Methodology

- **SheetKit**: 1 warmup run(s) + 5 measured runs per scenario. Median time reported.
- **ExcelJS / SheetJS**: Single run per scenario (comparison context only).

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

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Read Large Data (50k rows x 20 cols) | 1.24s | 1.99s | 2.09s | SheetKit |
| Read Heavy Styles (5k rows, formatted) | 59ms | 331ms | 146ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) | 597ms | 2.48s | 999ms | SheetKit |
| Read Formulas (10k rows) | 76ms | 310ms | 111ms | SheetKit |
| Read Strings (20k rows text-heavy) | 235ms | 967ms | 406ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) | 45ms | 211ms | 86ms | SheetKit |
| Read Comments (2k rows with comments) | 15ms | 169ms | 38ms | SheetKit |
| Read Merged Cells (500 regions) | 3ms | 29ms | 8ms | SheetKit |
| Read Mixed Workload (ERP document) | 58ms | 278ms | 114ms | SheetKit |

### Read (Scale)

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Read Scale 1k rows | 11ms | 57ms | 27ms | SheetKit |
| Read Scale 10k rows | 109ms | 456ms | 213ms | SheetKit |
| Read Scale 100k rows | 1.21s | 4.48s | 2.08s | SheetKit |

### Write

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Write 50000 rows x 20 cols | 678ms | 4.48s | 1.60s | SheetKit |
| Write 5000 styled rows | 50ms | 270ms | 84ms | SheetKit |
| Write 10 sheets x 5000 rows | 393ms | 2.08s | 716ms | SheetKit |
| Write 10000 rows with formulas | 40ms | 273ms | 76ms | SheetKit |
| Write 20000 text-heavy rows | 126ms | 757ms | 307ms | SheetKit |

### Write (DV)

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 15ms | 146ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Write 2000 rows with comments | 11ms | 101ms | 87ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Write 500 merged regions | 15ms | 39ms | 5ms | SheetJS |

### Write (Scale)

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Write 1k rows x 10 cols | 7ms | 55ms | 14ms | SheetKit |
| Write 10k rows x 10 cols | 67ms | 389ms | 123ms | SheetKit |
| Write 50k rows x 10 cols | 337ms | 2.02s | 721ms | SheetKit |
| Write 100k rows x 10 cols | 730ms | 4.39s | 1.70s | SheetKit |

### Round-Trip

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 219ms | 821ms | 272ms | SheetKit |

### Streaming

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Streaming write (50000 rows) | 1.18s | 1.17s | N/A | ExcelJS |

### Random Access

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Random-access read (1000 cells from 50k-row file) | 545ms | 4.56s | 1.79s | SheetKit |

### Mixed Write

| Scenario | SheetKit (median) | ExcelJS | SheetJS | Winner |
|----------|-------------------|---------|---------|--------|
| Mixed workload write (ERP-style) | 29ms | 177ms | N/A | SheetKit |

### SheetKit Detailed Statistics

| Scenario | Median | Min | Max | P95 | Memory (median) |
|----------|--------|-----|-----|-----|-----------------|
| Read Large Data (50k rows x 20 cols) | 1.24s | 1.22s | 1.27s | 1.27s | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | 59ms | 58ms | 59ms | 59ms | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | 597ms | 591ms | 600ms | 600ms | 0.0MB |
| Read Formulas (10k rows) | 76ms | 73ms | 79ms | 79ms | 0.0MB |
| Read Strings (20k rows text-heavy) | 235ms | 231ms | 240ms | 240ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | 45ms | 44ms | 46ms | 46ms | 0.0MB |
| Read Comments (2k rows with comments) | 15ms | 15ms | 15ms | 15ms | 0.0MB |
| Read Merged Cells (500 regions) | 3ms | 3ms | 3ms | 3ms | 0.0MB |
| Read Mixed Workload (ERP document) | 58ms | 57ms | 58ms | 58ms | 0.0MB |
| Read Scale 1k rows | 11ms | 11ms | 11ms | 11ms | 0.0MB |
| Read Scale 10k rows | 109ms | 107ms | 109ms | 109ms | 0.0MB |
| Read Scale 100k rows | 1.21s | 1.18s | 1.22s | 1.22s | 0.0MB |
| Write 50000 rows x 20 cols | 678ms | 675ms | 686ms | 686ms | 0.0MB |
| Write 5000 styled rows | 50ms | 49ms | 50ms | 50ms | 0.0MB |
| Write 10 sheets x 5000 rows | 393ms | 379ms | 416ms | 416ms | 0.0MB |
| Write 10000 rows with formulas | 40ms | 39ms | 41ms | 41ms | 0.0MB |
| Write 20000 text-heavy rows | 126ms | 124ms | 127ms | 127ms | 0.0MB |
| Write 5000 rows + 8 validation rules | 15ms | 14ms | 21ms | 21ms | 0.0MB |
| Write 2000 rows with comments | 11ms | 10ms | 12ms | 12ms | 0.0MB |
| Write 500 merged regions | 15ms | 15ms | 15ms | 15ms | 0.0MB |
| Write 1k rows x 10 cols | 7ms | 7ms | 7ms | 7ms | 0.0MB |
| Write 10k rows x 10 cols | 67ms | 65ms | 68ms | 68ms | 0.0MB |
| Write 50k rows x 10 cols | 337ms | 333ms | 352ms | 352ms | 0.0MB |
| Write 100k rows x 10 cols | 730ms | 675ms | 774ms | 774ms | 0.0MB |
| Buffer round-trip (10000 rows) | 219ms | 217ms | 221ms | 221ms | 0.0MB |
| Streaming write (50000 rows) | 1.18s | 1.14s | 1.20s | 1.20s | 0.0MB |
| Random-access read (1000 cells from 50k-row file) | 545ms | 545ms | 554ms | 554ms | 0.0MB |
| Mixed workload write (ERP-style) | 29ms | 28ms | 29ms | 29ms | 0.0MB |

## Summary

Total scenarios: 28

| Library | Wins |
|---------|------|
| SheetKit | 26/28 |
| ExcelJS | 1/28 |
| SheetJS | 1/28 |
