# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-09T12:19:15.384Z
Platform: darwin arm64
Node.js: v25.3.0

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
| Read Large Data (50k rows x 20 cols) | 1.28s | 1.81s | 2.12s | SheetKit |
| Read Heavy Styles (5k rows, formatted) | 60ms | 136ms | 124ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) | 566ms | 845ms | 899ms | SheetKit |
| Read Formulas (10k rows) | 78ms | 108ms | 161ms | SheetKit |
| Read Strings (20k rows text-heavy) | 236ms | 334ms | 331ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) | 43ms | 70ms | 67ms | SheetKit |
| Read Comments (2k rows with comments) | 17ms | 48ms | 30ms | SheetKit |
| Read Merged Cells (500 regions) | 4ms | 16ms | 5ms | SheetKit |
| Read Mixed Workload (ERP document) | 56ms | 100ms | 93ms | SheetKit |

### Read (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Scale 1k rows | 12ms | 20ms | 16ms | SheetKit |
| Read Scale 10k rows | 105ms | 160ms | 174ms | SheetKit |
| Read Scale 100k rows | 1.18s | 1.81s | 2.09s | SheetKit |

### Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 50000 rows x 20 cols | 695ms | 3.66s | 1.51s | SheetKit |
| Write 5000 styled rows | 99ms | 231ms | 89ms | SheetJS |
| Write 10 sheets x 5000 rows | 363ms | 1.81s | 526ms | SheetKit |
| Write 10000 rows with formulas | 43ms | 200ms | 91ms | SheetKit |
| Write 20000 text-heavy rows | 131ms | 657ms | 339ms | SheetKit |

### Write (DV)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 14ms | 123ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 2000 rows with comments | 12ms | 76ms | 101ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 500 merged regions | 15ms | 24ms | 4ms | SheetJS |

### Write (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 1k rows x 10 cols | 7ms | 40ms | 11ms | SheetKit |
| Write 10k rows x 10 cols | 66ms | 347ms | 130ms | SheetKit |
| Write 50k rows x 10 cols | 343ms | 1.92s | 671ms | SheetKit |
| Write 100k rows x 10 cols | 680ms | 3.86s | 1.52s | SheetKit |

### Round-Trip

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 268ms | 506ms | 245ms | SheetJS |

### Streaming

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Streaming write (50000 rows) | 1.11s | 796ms | N/A | ExcelJS |

### Random Access

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Random-access read (1000 cells from 50k-row file) | 512ms | 1.76s | 1.72s | SheetKit |

### Mixed Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Mixed workload write (ERP-style) | 31ms | 161ms | N/A | SheetKit |

## Summary

Total scenarios: 28

| Library | Wins |
|---------|------|
| SheetKit | 24/28 |
| SheetJS | 3/28 |
| ExcelJS | 1/28 |
