# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-10T00:47:06.135Z

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
- **Memory**: Measured as RSS (Resident Set Size) delta before/after each run. RSS includes both V8 heap and native (Rust) heap allocations, providing accurate measurements for napi-rs based libraries.

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
| Read Large Data (50k rows x 20 cols) | 710ms | 3.72s | 2.04s | SheetKit |
| Read Heavy Styles (5k rows, formatted) | 37ms | 239ms | 114ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) | 820ms | 2.11s | 866ms | SheetKit |
| Read Formulas (10k rows) | 50ms | 263ms | 112ms | SheetKit |
| Read Strings (20k rows text-heavy) | 146ms | 845ms | 374ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) | 29ms | 178ms | 81ms | SheetKit |
| Read Comments (2k rows with comments) | 12ms | 151ms | 37ms | SheetKit |
| Read Merged Cells (500 regions) | 2ms | 30ms | 8ms | SheetKit |
| Read Mixed Workload (ERP document) | 39ms | 255ms | 101ms | SheetKit |

### Read (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Scale 1k rows | 8ms | 53ms | 25ms | SheetKit |
| Read Scale 10k rows | 69ms | 397ms | 195ms | SheetKit |
| Read Scale 100k rows | 738ms | 3.92s | 2.08s | SheetKit |

### Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 50000 rows x 20 cols | 696ms | 3.54s | 1.60s | SheetKit |
| Write 5000 styled rows | 51ms | 226ms | 59ms | SheetKit |
| Write 10 sheets x 5000 rows | 349ms | 1.77s | 546ms | SheetKit |
| Write 10000 rows with formulas | 40ms | 213ms | 79ms | SheetKit |
| Write 20000 text-heavy rows | 127ms | 634ms | 275ms | SheetKit |

### Write (DV)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 13ms | 118ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 2000 rows with comments | 11ms | 89ms | 94ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 500 merged regions | 14ms | 36ms | 4ms | SheetJS |

### Write (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 1k rows x 10 cols | 7ms | 51ms | 13ms | SheetKit |
| Write 10k rows x 10 cols | 66ms | 354ms | 118ms | SheetKit |
| Write 50k rows x 10 cols | 339ms | 1.76s | 664ms | SheetKit |
| Write 100k rows x 10 cols | 690ms | 3.68s | 1.59s | SheetKit |

### Round-Trip

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 173ms | 660ms | 214ms | SheetKit |

### Streaming

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Streaming write (50000 rows) | 709ms | 707ms | N/A | ExcelJS |

### Random Access

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Random-access read (1000 cells from 50k-row file) | 560ms | 3.82s | 1.75s | SheetKit |

### Mixed Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Mixed workload write (ERP-style) | 28ms | 147ms | N/A | SheetKit |

### Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Memory (median) |
|----------|---------|--------|-----|-----|-----|-----------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 710ms | 692ms | 720ms | 720ms | 349.3MB |
| Read Large Data (50k rows x 20 cols) | ExcelJS | 3.72s | 3.69s | 3.75s | 3.75s | 0.9MB |
| Read Large Data (50k rows x 20 cols) | SheetJS | 2.04s | 2.01s | 2.05s | 2.05s | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 37ms | 35ms | 40ms | 40ms | 15.3MB |
| Read Heavy Styles (5k rows, formatted) | ExcelJS | 239ms | 233ms | 243ms | 243ms | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | SheetJS | 114ms | 113ms | 121ms | 121ms | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 820ms | 818ms | 825ms | 825ms | 27.3MB |
| Read Multi-Sheet (10 sheets x 5k rows) | ExcelJS | 2.11s | 2.10s | 2.14s | 2.14s | 0.1MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetJS | 866ms | 856ms | 876ms | 876ms | 0.0MB |
| Read Formulas (10k rows) | SheetKit | 50ms | 49ms | 53ms | 53ms | 13.3MB |
| Read Formulas (10k rows) | ExcelJS | 263ms | 259ms | 265ms | 265ms | 0.2MB |
| Read Formulas (10k rows) | SheetJS | 112ms | 110ms | 112ms | 112ms | 0.0MB |
| Read Strings (20k rows text-heavy) | SheetKit | 146ms | 142ms | 147ms | 147ms | 7.5MB |
| Read Strings (20k rows text-heavy) | ExcelJS | 845ms | 832ms | 852ms | 852ms | 0.0MB |
| Read Strings (20k rows text-heavy) | SheetJS | 374ms | 369ms | 375ms | 375ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | SheetKit | 29ms | 27ms | 30ms | 30ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | ExcelJS | 178ms | 173ms | 185ms | 185ms | 3.2MB |
| Read Data Validation (5k rows, 8 rules) | SheetJS | 81ms | 79ms | 82ms | 82ms | 0.0MB |
| Read Comments (2k rows with comments) | SheetKit | 12ms | 11ms | 12ms | 12ms | 0.6MB |
| Read Comments (2k rows with comments) | ExcelJS | 151ms | 149ms | 158ms | 158ms | 0.0MB |
| Read Comments (2k rows with comments) | SheetJS | 37ms | 36ms | 43ms | 43ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.1MB |
| Read Merged Cells (500 regions) | ExcelJS | 30ms | 28ms | 31ms | 31ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetJS | 8ms | 5ms | 8ms | 8ms | 0.0MB |
| Read Mixed Workload (ERP document) | SheetKit | 39ms | 38ms | 41ms | 41ms | 0.0MB |
| Read Mixed Workload (ERP document) | ExcelJS | 255ms | 253ms | 258ms | 258ms | 0.1MB |
| Read Mixed Workload (ERP document) | SheetJS | 101ms | 100ms | 105ms | 105ms | 0.0MB |
| Read Scale 1k rows | SheetKit | 8ms | 7ms | 8ms | 8ms | 0.0MB |
| Read Scale 1k rows | ExcelJS | 53ms | 50ms | 54ms | 54ms | 0.0MB |
| Read Scale 1k rows | SheetJS | 25ms | 20ms | 27ms | 27ms | 0.0MB |
| Read Scale 10k rows | SheetKit | 69ms | 66ms | 72ms | 72ms | 0.0MB |
| Read Scale 10k rows | ExcelJS | 397ms | 392ms | 405ms | 405ms | 0.0MB |
| Read Scale 10k rows | SheetJS | 195ms | 191ms | 196ms | 196ms | 0.0MB |
| Read Scale 100k rows | SheetKit | 738ms | 713ms | 749ms | 749ms | 90.7MB |
| Read Scale 100k rows | ExcelJS | 3.92s | 3.86s | 3.95s | 3.95s | 1.4MB |
| Read Scale 100k rows | SheetJS | 2.08s | 2.06s | 2.08s | 2.08s | 0.0MB |
| Write 50000 rows x 20 cols | SheetKit | 696ms | 688ms | 706ms | 706ms | 110.0MB |
| Write 50000 rows x 20 cols | ExcelJS | 3.54s | 3.43s | 3.69s | 3.69s | 52.2MB |
| Write 50000 rows x 20 cols | SheetJS | 1.60s | 1.59s | 2.06s | 2.06s | 0.0MB |
| Write 5000 styled rows | SheetKit | 51ms | 50ms | 52ms | 52ms | 13.7MB |
| Write 5000 styled rows | ExcelJS | 226ms | 219ms | 236ms | 236ms | 0.3MB |
| Write 5000 styled rows | SheetJS | 59ms | 57ms | 61ms | 61ms | 0.0MB |
| Write 10 sheets x 5000 rows | SheetKit | 349ms | 345ms | 354ms | 354ms | 166.8MB |
| Write 10 sheets x 5000 rows | ExcelJS | 1.77s | 1.75s | 1.80s | 1.80s | 0.3MB |
| Write 10 sheets x 5000 rows | SheetJS | 546ms | 532ms | 556ms | 556ms | 0.0MB |
| Write 10000 rows with formulas | SheetKit | 40ms | 40ms | 41ms | 41ms | 13.1MB |
| Write 10000 rows with formulas | ExcelJS | 213ms | 208ms | 225ms | 225ms | 0.0MB |
| Write 10000 rows with formulas | SheetJS | 79ms | 76ms | 83ms | 83ms | 0.0MB |
| Write 20000 text-heavy rows | SheetKit | 127ms | 127ms | 128ms | 128ms | 7.9MB |
| Write 20000 text-heavy rows | ExcelJS | 634ms | 626ms | 640ms | 640ms | 0.3MB |
| Write 20000 text-heavy rows | SheetJS | 275ms | 272ms | 277ms | 277ms | 0.0MB |
| Write 5000 rows + 8 validation rules | SheetKit | 13ms | 13ms | 14ms | 14ms | 3.1MB |
| Write 5000 rows + 8 validation rules | ExcelJS | 118ms | 115ms | 128ms | 128ms | 0.2MB |
| Write 2000 rows with comments | SheetKit | 11ms | 10ms | 11ms | 11ms | 0.1MB |
| Write 2000 rows with comments | ExcelJS | 89ms | 86ms | 90ms | 90ms | 0.1MB |
| Write 2000 rows with comments | SheetJS | 94ms | 93ms | 95ms | 95ms | 0.0MB |
| Write 500 merged regions | SheetKit | 14ms | 14ms | 15ms | 15ms | 0.0MB |
| Write 500 merged regions | ExcelJS | 36ms | 25ms | 39ms | 39ms | 0.0MB |
| Write 500 merged regions | SheetJS | 4ms | 4ms | 4ms | 4ms | 0.0MB |
| Write 1k rows x 10 cols | SheetKit | 7ms | 7ms | 8ms | 8ms | 0.0MB |
| Write 1k rows x 10 cols | ExcelJS | 51ms | 39ms | 54ms | 54ms | 0.0MB |
| Write 1k rows x 10 cols | SheetJS | 13ms | 12ms | 14ms | 14ms | 0.0MB |
| Write 10k rows x 10 cols | SheetKit | 66ms | 64ms | 68ms | 68ms | 1.6MB |
| Write 10k rows x 10 cols | ExcelJS | 354ms | 349ms | 366ms | 366ms | 0.2MB |
| Write 10k rows x 10 cols | SheetJS | 118ms | 115ms | 119ms | 119ms | 0.0MB |
| Write 50k rows x 10 cols | SheetKit | 339ms | 327ms | 343ms | 343ms | 23.4MB |
| Write 50k rows x 10 cols | ExcelJS | 1.76s | 1.75s | 1.78s | 1.78s | 0.0MB |
| Write 50k rows x 10 cols | SheetJS | 664ms | 658ms | 675ms | 675ms | 0.0MB |
| Write 100k rows x 10 cols | SheetKit | 690ms | 668ms | 719ms | 719ms | 79.6MB |
| Write 100k rows x 10 cols | ExcelJS | 3.68s | 3.61s | 3.84s | 3.84s | 0.0MB |
| Write 100k rows x 10 cols | SheetJS | 1.59s | 1.57s | 1.61s | 1.61s | 0.0MB |
| Buffer round-trip (10000 rows) | SheetKit | 173ms | 170ms | 175ms | 175ms | 3.4MB |
| Buffer round-trip (10000 rows) | ExcelJS | 660ms | 654ms | 666ms | 666ms | 0.5MB |
| Buffer round-trip (10000 rows) | SheetJS | 214ms | 209ms | 222ms | 222ms | 0.0MB |
| Streaming write (50000 rows) | SheetKit | 709ms | 701ms | 750ms | 750ms | 73.4MB |
| Streaming write (50000 rows) | ExcelJS | 707ms | 702ms | 729ms | 729ms | 0.2MB |
| Random-access read (1000 cells from 50k-row file) | SheetKit | 560ms | 560ms | 564ms | 564ms | 28.1MB |
| Random-access read (1000 cells from 50k-row file) | ExcelJS | 3.82s | 3.75s | 3.89s | 3.89s | 0.0MB |
| Random-access read (1000 cells from 50k-row file) | SheetJS | 1.75s | 1.75s | 1.76s | 1.76s | 0.0MB |
| Mixed workload write (ERP-style) | SheetKit | 28ms | 28ms | 29ms | 29ms | 6.5MB |
| Mixed workload write (ERP-style) | ExcelJS | 147ms | 144ms | 150ms | 150ms | 0.1MB |

### Memory Usage

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| Read Large Data (50k rows x 20 cols) | 349.3MB | 0.9MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | 15.3MB | 0.0MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | 27.3MB | 0.1MB | 0.0MB |
| Read Formulas (10k rows) | 13.3MB | 0.2MB | 0.0MB |
| Read Strings (20k rows text-heavy) | 7.5MB | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | 0.0MB | 3.2MB | 0.0MB |
| Read Comments (2k rows with comments) | 0.6MB | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) | 0.1MB | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) | 0.0MB | 0.1MB | 0.0MB |
| Read Scale 1k rows | 0.0MB | 0.0MB | 0.0MB |
| Read Scale 10k rows | 0.0MB | 0.0MB | 0.0MB |
| Read Scale 100k rows | 90.7MB | 1.4MB | 0.0MB |
| Write 50000 rows x 20 cols | 110.0MB | 52.2MB | 0.0MB |
| Write 5000 styled rows | 13.7MB | 0.3MB | 0.0MB |
| Write 10 sheets x 5000 rows | 166.8MB | 0.3MB | 0.0MB |
| Write 10000 rows with formulas | 13.1MB | 0.0MB | 0.0MB |
| Write 20000 text-heavy rows | 7.9MB | 0.3MB | 0.0MB |
| Write 5000 rows + 8 validation rules | 3.1MB | 0.2MB | N/A |
| Write 2000 rows with comments | 0.1MB | 0.1MB | 0.0MB |
| Write 500 merged regions | 0.0MB | 0.0MB | 0.0MB |
| Write 1k rows x 10 cols | 0.0MB | 0.0MB | 0.0MB |
| Write 10k rows x 10 cols | 1.6MB | 0.2MB | 0.0MB |
| Write 50k rows x 10 cols | 23.4MB | 0.0MB | 0.0MB |
| Write 100k rows x 10 cols | 79.6MB | 0.0MB | 0.0MB |
| Buffer round-trip (10000 rows) | 3.4MB | 0.5MB | 0.0MB |
| Streaming write (50000 rows) | 73.4MB | 0.2MB | N/A |
| Random-access read (1000 cells from 50k-row file) | 28.1MB | 0.0MB | 0.0MB |
| Mixed workload write (ERP-style) | 6.5MB | 0.1MB | N/A |

## Summary

Total scenarios: 28

| Library | Wins |
|---------|------|
| SheetKit | 26/28 |
| ExcelJS | 1/28 |
| SheetJS | 1/28 |
