# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-09T10:32:03.077Z
Platform: linux x64
Node.js: v22.22.0

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
| Read Large Data (50k rows x 20 cols) | 4.03s | 3.13s | 4.20s | ExcelJS |
| Read Heavy Styles (5k rows, formatted) | 96ms | 404ms | 219ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) | 1.04s | 1.40s | 1.75s | SheetKit |
| Read Formulas (10k rows) | 121ms | 210ms | 216ms | SheetKit |
| Read Strings (20k rows text-heavy) | 456ms | 754ms | 720ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) | 75ms | 133ms | 120ms | SheetKit |
| Read Comments (2k rows with comments) | 23ms | 163ms | 63ms | SheetKit |
| Read Merged Cells (500 regions) | 6ms | 31ms | 9ms | SheetKit |
| Read Mixed Workload (ERP document) | 102ms | 181ms | 180ms | SheetKit |

### Read (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Scale 1k rows | 20ms | 30ms | 36ms | SheetKit |
| Read Scale 10k rows | 193ms | 274ms | 386ms | SheetKit |
| Read Scale 100k rows | 2.68s | 3.16s | 4.52s | SheetKit |

### Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 50000 rows x 20 cols | 49.74s | 7.29s | 2.73s | SheetJS |
| Write 5000 styled rows | 478ms | 431ms | 139ms | SheetJS |
| Write 10 sheets x 5000 rows | 2.00s | 3.45s | 1.07s | SheetJS |
| Write 10000 rows with formulas | 226ms | 371ms | 196ms | SheetJS |
| Write 20000 text-heavy rows | 1.39s | 1.28s | 657ms | SheetJS |

### Write (DV)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 64ms | 232ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 2000 rows with comments | 28ms | 151ms | 352ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 500 merged regions | 19ms | 63ms | 8ms | SheetJS |

### Write (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 1k rows x 10 cols | 29ms | 83ms | 30ms | SheetKit |
| Write 10k rows x 10 cols | 475ms | 706ms | 304ms | SheetJS |
| Write 50k rows x 10 cols | 29.10s | 4.23s | 1.67s | SheetJS |
| Write 100k rows x 10 cols | 143.60s | 8.16s | 2.83s | SheetJS |

### Round-Trip

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 789ms | 862ms | 518ms | SheetJS |

### Streaming

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Streaming write (50000 rows) | 3.28s | 1.54s | N/A | ExcelJS |

### Random Access

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Random-access read (1000 cells from 50k-row file) | 1.45s | 3.25s | 3.53s | SheetKit |

### Mixed Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Mixed workload write (ERP-style) | 121ms | 254ms | N/A | SheetKit |

## Summary

Total scenarios: 28

| Library | Wins |
|---------|------|
| SheetKit | 16/28 |
| SheetJS | 10/28 |
| ExcelJS | 2/28 |
