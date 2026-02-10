# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-10T14:20:27.878Z

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
| Read Large Data (50k rows x 20 cols) | 950ms | 5.20s | 2.56s | SheetKit |
| Read Heavy Styles (5k rows, formatted) | 60ms | 363ms | 228ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) | 833ms | 2.40s | 948ms | SheetKit |
| Read Formulas (10k rows) | 53ms | 299ms | 120ms | SheetKit |
| Read Strings (20k rows text-heavy) | 143ms | 918ms | 402ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) | 33ms | 218ms | 94ms | SheetKit |
| Read Comments (2k rows with comments) | 13ms | 209ms | 60ms | SheetKit |
| Read Merged Cells (500 regions) | 7ms | 54ms | 12ms | SheetKit |
| Read Mixed Workload (ERP document) | 55ms | 655ms | 115ms | SheetKit |

### Read (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Scale 1k rows | 8ms | 63ms | 31ms | SheetKit |
| Read Scale 10k rows | 70ms | 482ms | 224ms | SheetKit |
| Read Scale 100k rows | 2.59s | 4.54s | 2.32s | SheetJS |

### Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 50000 rows x 20 cols | 1.97s | 5.02s | 2.34s | SheetKit |
| Write 5000 styled rows | 86ms | 1.52s | 483ms | SheetKit |
| Write 10 sheets x 5000 rows | 2.93s | 9.05s | 928ms | SheetJS |
| Write 10000 rows with formulas | 70ms | 274ms | 129ms | SheetKit |
| Write 20000 text-heavy rows | 148ms | 788ms | 354ms | SheetKit |

### Write (DV)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 22ms | 157ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 2000 rows with comments | 17ms | 120ms | 143ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 500 merged regions | 20ms | 46ms | 7ms | SheetJS |

### Write (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 1k rows x 10 cols | 15ms | 87ms | 21ms | SheetKit |
| Write 10k rows x 10 cols | 79ms | 442ms | 152ms | SheetKit |
| Write 50k rows x 10 cols | 351ms | 2.19s | 864ms | SheetKit |
| Write 100k rows x 10 cols | 689ms | 4.69s | 1.90s | SheetKit |

### Round-Trip

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 233ms | 772ms | 368ms | SheetKit |

### Streaming

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Streaming write (50000 rows) | 997ms | 846ms | N/A | ExcelJS |

### Random Access

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Random-access read (1000 cells from 50k-row file) | 577ms | 4.17s | 1.81s | SheetKit |

### Mixed Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Mixed workload write (ERP-style) | 27ms | 144ms | N/A | SheetKit |

### Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Memory (median) |
|----------|---------|--------|-----|-----|-----|-----------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 950ms | 919ms | 1.39s | 1.39s | 195.4MB |
| Read Large Data (50k rows x 20 cols) | ExcelJS | 5.20s | 4.58s | 11.24s | 11.24s | 0.3MB |
| Read Large Data (50k rows x 20 cols) | SheetJS | 2.56s | 2.23s | 3.57s | 3.57s | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 60ms | 51ms | 66ms | 66ms | 5.3MB |
| Read Heavy Styles (5k rows, formatted) | ExcelJS | 363ms | 328ms | 1.84s | 1.84s | 0.2MB |
| Read Heavy Styles (5k rows, formatted) | SheetJS | 228ms | 199ms | 1.31s | 1.31s | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 833ms | 807ms | 844ms | 844ms | 114.3MB |
| Read Multi-Sheet (10 sheets x 5k rows) | ExcelJS | 2.40s | 2.30s | 9.42s | 9.42s | 0.1MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetJS | 948ms | 938ms | 961ms | 961ms | 1.6MB |
| Read Formulas (10k rows) | SheetKit | 53ms | 49ms | 55ms | 55ms | 9.3MB |
| Read Formulas (10k rows) | ExcelJS | 299ms | 280ms | 305ms | 305ms | 0.0MB |
| Read Formulas (10k rows) | SheetJS | 120ms | 118ms | 136ms | 136ms | 0.0MB |
| Read Strings (20k rows text-heavy) | SheetKit | 143ms | 142ms | 145ms | 145ms | 2.8MB |
| Read Strings (20k rows text-heavy) | ExcelJS | 918ms | 913ms | 968ms | 968ms | 0.0MB |
| Read Strings (20k rows text-heavy) | SheetJS | 402ms | 397ms | 405ms | 405ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | SheetKit | 33ms | 30ms | 35ms | 35ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | ExcelJS | 218ms | 199ms | 364ms | 364ms | 3.1MB |
| Read Data Validation (5k rows, 8 rules) | SheetJS | 94ms | 87ms | 98ms | 98ms | 0.0MB |
| Read Comments (2k rows with comments) | SheetKit | 13ms | 12ms | 14ms | 14ms | 0.6MB |
| Read Comments (2k rows with comments) | ExcelJS | 209ms | 169ms | 254ms | 254ms | 0.0MB |
| Read Comments (2k rows with comments) | SheetJS | 60ms | 56ms | 76ms | 76ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetKit | 7ms | 5ms | 9ms | 9ms | 0.0MB |
| Read Merged Cells (500 regions) | ExcelJS | 54ms | 49ms | 55ms | 55ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetJS | 12ms | 8ms | 12ms | 12ms | 0.0MB |
| Read Mixed Workload (ERP document) | SheetKit | 55ms | 48ms | 63ms | 63ms | 0.0MB |
| Read Mixed Workload (ERP document) | ExcelJS | 655ms | 342ms | 3.33s | 3.33s | 2.1MB |
| Read Mixed Workload (ERP document) | SheetJS | 115ms | 108ms | 137ms | 137ms | 0.0MB |
| Read Scale 1k rows | SheetKit | 8ms | 7ms | 8ms | 8ms | 0.0MB |
| Read Scale 1k rows | ExcelJS | 63ms | 62ms | 70ms | 70ms | 0.0MB |
| Read Scale 1k rows | SheetJS | 31ms | 24ms | 36ms | 36ms | 0.0MB |
| Read Scale 10k rows | SheetKit | 70ms | 66ms | 72ms | 72ms | 0.0MB |
| Read Scale 10k rows | ExcelJS | 482ms | 435ms | 515ms | 515ms | 0.0MB |
| Read Scale 10k rows | SheetJS | 224ms | 216ms | 248ms | 248ms | 0.0MB |
| Read Scale 100k rows | SheetKit | 2.59s | 934ms | 4.96s | 4.96s | 175.1MB |
| Read Scale 100k rows | ExcelJS | 4.54s | 4.53s | 4.67s | 4.67s | 0.0MB |
| Read Scale 100k rows | SheetJS | 2.32s | 2.23s | 2.44s | 2.44s | 0.0MB |
| Write 50000 rows x 20 cols | SheetKit | 1.97s | 706ms | 4.63s | 4.63s | 98.1MB |
| Write 50000 rows x 20 cols | ExcelJS | 5.02s | 4.10s | 18.85s | 18.85s | 13.9MB |
| Write 50000 rows x 20 cols | SheetJS | 2.34s | 1.78s | 2.44s | 2.44s | 0.0MB |
| Write 5000 styled rows | SheetKit | 86ms | 84ms | 356ms | 356ms | 0.0MB |
| Write 5000 styled rows | ExcelJS | 1.52s | 399ms | 2.35s | 2.35s | 0.2MB |
| Write 5000 styled rows | SheetJS | 483ms | 436ms | 584ms | 584ms | 0.0MB |
| Write 10 sheets x 5000 rows | SheetKit | 2.93s | 2.01s | 3.18s | 3.18s | 22.3MB |
| Write 10 sheets x 5000 rows | ExcelJS | 9.05s | 5.22s | 15.07s | 15.07s | 0.0MB |
| Write 10 sheets x 5000 rows | SheetJS | 928ms | 900ms | 1.46s | 1.46s | 0.0MB |
| Write 10000 rows with formulas | SheetKit | 70ms | 65ms | 74ms | 74ms | 2.9MB |
| Write 10000 rows with formulas | ExcelJS | 274ms | 262ms | 318ms | 318ms | 0.2MB |
| Write 10000 rows with formulas | SheetJS | 129ms | 99ms | 256ms | 256ms | 0.0MB |
| Write 20000 text-heavy rows | SheetKit | 148ms | 148ms | 165ms | 165ms | 25.4MB |
| Write 20000 text-heavy rows | ExcelJS | 788ms | 750ms | 1.50s | 1.50s | 0.2MB |
| Write 20000 text-heavy rows | SheetJS | 354ms | 345ms | 396ms | 396ms | 0.0MB |
| Write 5000 rows + 8 validation rules | SheetKit | 22ms | 18ms | 23ms | 23ms | 1.2MB |
| Write 5000 rows + 8 validation rules | ExcelJS | 157ms | 149ms | 168ms | 168ms | 0.1MB |
| Write 2000 rows with comments | SheetKit | 17ms | 13ms | 18ms | 18ms | 0.0MB |
| Write 2000 rows with comments | ExcelJS | 120ms | 105ms | 136ms | 136ms | 0.0MB |
| Write 2000 rows with comments | SheetJS | 143ms | 130ms | 162ms | 162ms | 0.0MB |
| Write 500 merged regions | SheetKit | 20ms | 14ms | 26ms | 26ms | 0.0MB |
| Write 500 merged regions | ExcelJS | 46ms | 37ms | 59ms | 59ms | 0.0MB |
| Write 500 merged regions | SheetJS | 7ms | 6ms | 11ms | 11ms | 0.0MB |
| Write 1k rows x 10 cols | SheetKit | 15ms | 13ms | 16ms | 16ms | 0.0MB |
| Write 1k rows x 10 cols | ExcelJS | 87ms | 71ms | 94ms | 94ms | 0.0MB |
| Write 1k rows x 10 cols | SheetJS | 21ms | 18ms | 23ms | 23ms | 0.0MB |
| Write 10k rows x 10 cols | SheetKit | 79ms | 72ms | 89ms | 89ms | 0.2MB |
| Write 10k rows x 10 cols | ExcelJS | 442ms | 411ms | 492ms | 492ms | 0.3MB |
| Write 10k rows x 10 cols | SheetJS | 152ms | 149ms | 164ms | 164ms | 0.0MB |
| Write 50k rows x 10 cols | SheetKit | 351ms | 350ms | 354ms | 354ms | 42.4MB |
| Write 50k rows x 10 cols | ExcelJS | 2.19s | 2.12s | 2.88s | 2.88s | 0.2MB |
| Write 50k rows x 10 cols | SheetJS | 864ms | 790ms | 1.18s | 1.18s | 0.0MB |
| Write 100k rows x 10 cols | SheetKit | 689ms | 685ms | 695ms | 695ms | 100.1MB |
| Write 100k rows x 10 cols | ExcelJS | 4.69s | 3.80s | 4.92s | 4.92s | 0.8MB |
| Write 100k rows x 10 cols | SheetJS | 1.90s | 1.73s | 1.96s | 1.96s | 0.0MB |
| Buffer round-trip (10000 rows) | SheetKit | 233ms | 220ms | 255ms | 255ms | 1.5MB |
| Buffer round-trip (10000 rows) | ExcelJS | 772ms | 730ms | 881ms | 881ms | 0.2MB |
| Buffer round-trip (10000 rows) | SheetJS | 368ms | 325ms | 425ms | 425ms | 0.0MB |
| Streaming write (50000 rows) | SheetKit | 997ms | 768ms | 1.36s | 1.36s | 93.4MB |
| Streaming write (50000 rows) | ExcelJS | 846ms | 798ms | 934ms | 934ms | 0.2MB |
| Random-access read (1000 cells from 50k-row file) | SheetKit | 577ms | 572ms | 581ms | 581ms | 12.4MB |
| Random-access read (1000 cells from 50k-row file) | ExcelJS | 4.17s | 4.08s | 4.37s | 4.37s | 0.0MB |
| Random-access read (1000 cells from 50k-row file) | SheetJS | 1.81s | 1.78s | 1.87s | 1.87s | 0.0MB |
| Mixed workload write (ERP-style) | SheetKit | 27ms | 27ms | 28ms | 28ms | 0.0MB |
| Mixed workload write (ERP-style) | ExcelJS | 144ms | 141ms | 151ms | 151ms | 0.2MB |

### Memory Usage

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| Read Large Data (50k rows x 20 cols) | 195.4MB | 0.3MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | 5.3MB | 0.2MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | 114.3MB | 0.1MB | 1.6MB |
| Read Formulas (10k rows) | 9.3MB | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) | 2.8MB | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | 0.0MB | 3.1MB | 0.0MB |
| Read Comments (2k rows with comments) | 0.6MB | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) | 0.0MB | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) | 0.0MB | 2.1MB | 0.0MB |
| Read Scale 1k rows | 0.0MB | 0.0MB | 0.0MB |
| Read Scale 10k rows | 0.0MB | 0.0MB | 0.0MB |
| Read Scale 100k rows | 175.1MB | 0.0MB | 0.0MB |
| Write 50000 rows x 20 cols | 98.1MB | 13.9MB | 0.0MB |
| Write 5000 styled rows | 0.0MB | 0.2MB | 0.0MB |
| Write 10 sheets x 5000 rows | 22.3MB | 0.0MB | 0.0MB |
| Write 10000 rows with formulas | 2.9MB | 0.2MB | 0.0MB |
| Write 20000 text-heavy rows | 25.4MB | 0.2MB | 0.0MB |
| Write 5000 rows + 8 validation rules | 1.2MB | 0.1MB | N/A |
| Write 2000 rows with comments | 0.0MB | 0.0MB | 0.0MB |
| Write 500 merged regions | 0.0MB | 0.0MB | 0.0MB |
| Write 1k rows x 10 cols | 0.0MB | 0.0MB | 0.0MB |
| Write 10k rows x 10 cols | 0.2MB | 0.3MB | 0.0MB |
| Write 50k rows x 10 cols | 42.4MB | 0.2MB | 0.0MB |
| Write 100k rows x 10 cols | 100.1MB | 0.8MB | 0.0MB |
| Buffer round-trip (10000 rows) | 1.5MB | 0.2MB | 0.0MB |
| Streaming write (50000 rows) | 93.4MB | 0.2MB | N/A |
| Random-access read (1000 cells from 50k-row file) | 12.4MB | 0.0MB | 0.0MB |
| Mixed workload write (ERP-style) | 0.0MB | 0.2MB | N/A |

## Summary

Total scenarios: 28

| Library | Wins |
|---------|------|
| SheetKit | 24/28 |
| SheetJS | 3/28 |
| ExcelJS | 1/28 |
