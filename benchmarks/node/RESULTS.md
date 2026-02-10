# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-10T14:32:55.706Z

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
| Read Large Data (50k rows x 20 cols) | 680ms | 3.88s | 2.06s | SheetKit |
| Read Large Data (50k rows x 20 cols) (async) | 655ms | N/A | N/A | SheetKit |
| Read Heavy Styles (5k rows, formatted) | 37ms | 239ms | 113ms | SheetKit |
| Read Heavy Styles (5k rows, formatted) (async) | 36ms | N/A | N/A | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) | 781ms | 2.15s | 896ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) (async) | 777ms | N/A | N/A | SheetKit |
| Read Formulas (10k rows) | 49ms | 279ms | 119ms | SheetKit |
| Read Formulas (10k rows) (async) | 49ms | N/A | N/A | SheetKit |
| Read Strings (20k rows text-heavy) | 140ms | 886ms | 387ms | SheetKit |
| Read Strings (20k rows text-heavy) (async) | 143ms | N/A | N/A | SheetKit |
| Read Data Validation (5k rows, 8 rules) | 28ms | 190ms | 80ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) (async) | 28ms | N/A | N/A | SheetKit |
| Read Comments (2k rows with comments) | 11ms | 161ms | 41ms | SheetKit |
| Read Comments (2k rows with comments) (async) | 11ms | N/A | N/A | SheetKit |
| Read Merged Cells (500 regions) | 2ms | 30ms | 7ms | SheetKit |
| Read Merged Cells (500 regions) (async) | 2ms | N/A | N/A | SheetKit |
| Read Mixed Workload (ERP document) | 39ms | 278ms | 103ms | SheetKit |
| Read Mixed Workload (ERP document) (async) | 41ms | N/A | N/A | SheetKit |

### Read (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Scale 1k rows | 7ms | 55ms | 25ms | SheetKit |
| Read Scale 1k rows (async) | 7ms | N/A | N/A | SheetKit |
| Read Scale 10k rows | 68ms | 432ms | 189ms | SheetKit |
| Read Scale 10k rows (async) | 68ms | N/A | N/A | SheetKit |
| Read Scale 100k rows | 714ms | 4.03s | 2.11s | SheetKit |
| Read Scale 100k rows (async) | 683ms | N/A | N/A | SheetKit |

### Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 50000 rows x 20 cols | 657ms | 3.49s | 1.59s | SheetKit |
| Write 5000 styled rows | 48ms | 227ms | 56ms | SheetKit |
| Write 10 sheets x 5000 rows | 344ms | 1.83s | 611ms | SheetKit |
| Write 10000 rows with formulas | 39ms | 221ms | 84ms | SheetKit |
| Write 20000 text-heavy rows | 123ms | 626ms | 271ms | SheetKit |

### Write (DV)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 13ms | 118ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 2000 rows with comments | 11ms | 89ms | 93ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 500 merged regions | 14ms | 34ms | 4ms | SheetJS |

### Write (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 1k rows x 10 cols | 7ms | 52ms | 13ms | SheetKit |
| Write 10k rows x 10 cols | 66ms | 353ms | 113ms | SheetKit |
| Write 50k rows x 10 cols | 332ms | 1.74s | 653ms | SheetKit |
| Write 100k rows x 10 cols | 665ms | 3.68s | 1.56s | SheetKit |

### Round-Trip

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 167ms | 674ms | 211ms | SheetKit |

### Streaming

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Streaming write (50000 rows) | 669ms | 702ms | N/A | SheetKit |

### Random Access

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Random-access read (1000 cells from 50k-row file) | 550ms | 3.97s | 1.74s | SheetKit |
| Random-access read (1000 cells from 50k-row file) (async) | 549ms | N/A | N/A | SheetKit |

### Mixed Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Mixed workload write (ERP-style) | 28ms | 146ms | N/A | SheetKit |

### Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Memory (median) |
|----------|---------|--------|-----|-----|-----|-----------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 680ms | 673ms | 684ms | 684ms | 195.4MB |
| Read Large Data (50k rows x 20 cols) | ExcelJS | 3.88s | 3.88s | 4.15s | 4.15s | 0.0MB |
| Read Large Data (50k rows x 20 cols) | SheetJS | 2.06s | 2.03s | 2.07s | 2.07s | 0.0MB |
| Read Large Data (50k rows x 20 cols) (async) | SheetKit | 655ms | 649ms | 663ms | 663ms | 17.2MB |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 37ms | 34ms | 37ms | 37ms | 6.6MB |
| Read Heavy Styles (5k rows, formatted) | ExcelJS | 239ms | 236ms | 247ms | 247ms | 0.1MB |
| Read Heavy Styles (5k rows, formatted) | SheetJS | 113ms | 110ms | 114ms | 114ms | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (async) | SheetKit | 36ms | 34ms | 37ms | 37ms | 0.1MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 781ms | 777ms | 786ms | 786ms | 132.1MB |
| Read Multi-Sheet (10 sheets x 5k rows) | ExcelJS | 2.15s | 2.13s | 2.20s | 2.20s | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetJS | 896ms | 846ms | 927ms | 927ms | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (async) | SheetKit | 777ms | 773ms | 790ms | 790ms | 17.6MB |
| Read Formulas (10k rows) | SheetKit | 49ms | 49ms | 54ms | 54ms | 9.0MB |
| Read Formulas (10k rows) | ExcelJS | 279ms | 271ms | 283ms | 283ms | 0.2MB |
| Read Formulas (10k rows) | SheetJS | 119ms | 117ms | 125ms | 125ms | 0.0MB |
| Read Formulas (10k rows) (async) | SheetKit | 49ms | 48ms | 53ms | 53ms | 0.0MB |
| Read Strings (20k rows text-heavy) | SheetKit | 140ms | 138ms | 142ms | 142ms | 2.5MB |
| Read Strings (20k rows text-heavy) | ExcelJS | 886ms | 871ms | 895ms | 895ms | 0.1MB |
| Read Strings (20k rows text-heavy) | SheetJS | 387ms | 374ms | 393ms | 393ms | 0.0MB |
| Read Strings (20k rows text-heavy) (async) | SheetKit | 143ms | 139ms | 144ms | 144ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | SheetKit | 28ms | 26ms | 29ms | 29ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | ExcelJS | 190ms | 179ms | 191ms | 191ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | SheetJS | 80ms | 77ms | 83ms | 83ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (async) | SheetKit | 28ms | 26ms | 29ms | 29ms | 0.0MB |
| Read Comments (2k rows with comments) | SheetKit | 11ms | 11ms | 12ms | 12ms | 0.6MB |
| Read Comments (2k rows with comments) | ExcelJS | 161ms | 153ms | 163ms | 163ms | 0.0MB |
| Read Comments (2k rows with comments) | SheetJS | 41ms | 36ms | 42ms | 42ms | 0.0MB |
| Read Comments (2k rows with comments) (async) | SheetKit | 11ms | 11ms | 12ms | 12ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB |
| Read Merged Cells (500 regions) | ExcelJS | 30ms | 29ms | 31ms | 31ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetJS | 7ms | 5ms | 8ms | 8ms | 0.0MB |
| Read Merged Cells (500 regions) (async) | SheetKit | 2ms | 2ms | 2ms | 2ms | 0.0MB |
| Read Mixed Workload (ERP document) | SheetKit | 39ms | 37ms | 41ms | 41ms | 0.0MB |
| Read Mixed Workload (ERP document) | ExcelJS | 278ms | 264ms | 282ms | 282ms | 0.3MB |
| Read Mixed Workload (ERP document) | SheetJS | 103ms | 101ms | 110ms | 110ms | 0.0MB |
| Read Mixed Workload (ERP document) (async) | SheetKit | 41ms | 38ms | 42ms | 42ms | 0.0MB |
| Read Scale 1k rows | SheetKit | 7ms | 7ms | 8ms | 8ms | 0.0MB |
| Read Scale 1k rows | ExcelJS | 55ms | 54ms | 58ms | 58ms | 0.0MB |
| Read Scale 1k rows | SheetJS | 25ms | 21ms | 29ms | 29ms | 0.0MB |
| Read Scale 1k rows (async) | SheetKit | 7ms | 7ms | 7ms | 7ms | 0.0MB |
| Read Scale 10k rows | SheetKit | 68ms | 65ms | 71ms | 71ms | 2.1MB |
| Read Scale 10k rows | ExcelJS | 432ms | 415ms | 438ms | 438ms | 0.0MB |
| Read Scale 10k rows | SheetJS | 189ms | 186ms | 192ms | 192ms | 0.0MB |
| Read Scale 10k rows (async) | SheetKit | 68ms | 67ms | 72ms | 72ms | 0.0MB |
| Read Scale 100k rows | SheetKit | 714ms | 705ms | 779ms | 779ms | 161.1MB |
| Read Scale 100k rows | ExcelJS | 4.03s | 3.97s | 4.14s | 4.14s | 0.4MB |
| Read Scale 100k rows | SheetJS | 2.11s | 2.06s | 2.13s | 2.13s | 0.0MB |
| Read Scale 100k rows (async) | SheetKit | 683ms | 678ms | 684ms | 684ms | 0.0MB |
| Write 50000 rows x 20 cols | SheetKit | 657ms | 639ms | 668ms | 668ms | 89.4MB |
| Write 50000 rows x 20 cols | ExcelJS | 3.49s | 3.43s | 3.74s | 3.74s | 6.1MB |
| Write 50000 rows x 20 cols | SheetJS | 1.59s | 1.57s | 2.09s | 2.09s | 0.0MB |
| Write 5000 styled rows | SheetKit | 48ms | 48ms | 49ms | 49ms | 0.0MB |
| Write 5000 styled rows | ExcelJS | 227ms | 220ms | 238ms | 238ms | 0.1MB |
| Write 5000 styled rows | SheetJS | 56ms | 53ms | 59ms | 59ms | 0.0MB |
| Write 10 sheets x 5000 rows | SheetKit | 344ms | 343ms | 348ms | 348ms | 38.2MB |
| Write 10 sheets x 5000 rows | ExcelJS | 1.83s | 1.74s | 1.85s | 1.85s | 0.1MB |
| Write 10 sheets x 5000 rows | SheetJS | 611ms | 581ms | 624ms | 624ms | 0.0MB |
| Write 10000 rows with formulas | SheetKit | 39ms | 39ms | 41ms | 41ms | 2.3MB |
| Write 10000 rows with formulas | ExcelJS | 221ms | 213ms | 232ms | 232ms | 0.0MB |
| Write 10000 rows with formulas | SheetJS | 84ms | 81ms | 85ms | 85ms | 0.0MB |
| Write 20000 text-heavy rows | SheetKit | 123ms | 122ms | 125ms | 125ms | 15.7MB |
| Write 20000 text-heavy rows | ExcelJS | 626ms | 618ms | 635ms | 635ms | 0.1MB |
| Write 20000 text-heavy rows | SheetJS | 271ms | 269ms | 273ms | 273ms | 0.0MB |
| Write 5000 rows + 8 validation rules | SheetKit | 13ms | 13ms | 14ms | 14ms | 1.2MB |
| Write 5000 rows + 8 validation rules | ExcelJS | 118ms | 116ms | 123ms | 123ms | 0.2MB |
| Write 2000 rows with comments | SheetKit | 11ms | 10ms | 11ms | 11ms | 0.0MB |
| Write 2000 rows with comments | ExcelJS | 89ms | 85ms | 91ms | 91ms | 0.1MB |
| Write 2000 rows with comments | SheetJS | 93ms | 92ms | 95ms | 95ms | 0.0MB |
| Write 500 merged regions | SheetKit | 14ms | 14ms | 15ms | 15ms | 0.0MB |
| Write 500 merged regions | ExcelJS | 34ms | 26ms | 39ms | 39ms | 0.0MB |
| Write 500 merged regions | SheetJS | 4ms | 4ms | 4ms | 4ms | 0.0MB |
| Write 1k rows x 10 cols | SheetKit | 7ms | 7ms | 7ms | 7ms | 0.0MB |
| Write 1k rows x 10 cols | ExcelJS | 52ms | 50ms | 55ms | 55ms | 0.0MB |
| Write 1k rows x 10 cols | SheetJS | 13ms | 12ms | 13ms | 13ms | 0.0MB |
| Write 10k rows x 10 cols | SheetKit | 66ms | 65ms | 68ms | 68ms | 0.0MB |
| Write 10k rows x 10 cols | ExcelJS | 353ms | 353ms | 359ms | 359ms | 0.2MB |
| Write 10k rows x 10 cols | SheetJS | 113ms | 111ms | 119ms | 119ms | 0.0MB |
| Write 50k rows x 10 cols | SheetKit | 332ms | 328ms | 335ms | 335ms | 39.1MB |
| Write 50k rows x 10 cols | ExcelJS | 1.74s | 1.73s | 1.79s | 1.79s | 0.3MB |
| Write 50k rows x 10 cols | SheetJS | 653ms | 650ms | 660ms | 660ms | 0.0MB |
| Write 100k rows x 10 cols | SheetKit | 665ms | 662ms | 667ms | 667ms | 70.4MB |
| Write 100k rows x 10 cols | ExcelJS | 3.68s | 3.57s | 3.77s | 3.77s | 0.1MB |
| Write 100k rows x 10 cols | SheetJS | 1.56s | 1.53s | 1.59s | 1.59s | 0.0MB |
| Buffer round-trip (10000 rows) | SheetKit | 167ms | 165ms | 168ms | 168ms | 0.0MB |
| Buffer round-trip (10000 rows) | ExcelJS | 674ms | 652ms | 681ms | 681ms | 0.2MB |
| Buffer round-trip (10000 rows) | SheetJS | 211ms | 204ms | 218ms | 218ms | 0.0MB |
| Streaming write (50000 rows) | SheetKit | 669ms | 651ms | 694ms | 694ms | 80.0MB |
| Streaming write (50000 rows) | ExcelJS | 702ms | 700ms | 712ms | 712ms | 0.0MB |
| Random-access read (1000 cells from 50k-row file) | SheetKit | 550ms | 545ms | 556ms | 556ms | 18.2MB |
| Random-access read (1000 cells from 50k-row file) | ExcelJS | 3.97s | 3.88s | 3.99s | 3.99s | 0.0MB |
| Random-access read (1000 cells from 50k-row file) | SheetJS | 1.74s | 1.73s | 1.74s | 1.74s | 0.0MB |
| Random-access read (1000 cells from 50k-row file) (async) | SheetKit | 549ms | 546ms | 551ms | 551ms | 4.2MB |
| Mixed workload write (ERP-style) | SheetKit | 28ms | 27ms | 29ms | 29ms | 0.0MB |
| Mixed workload write (ERP-style) | ExcelJS | 146ms | 144ms | 150ms | 150ms | 0.2MB |

### Memory Usage

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| Read Large Data (50k rows x 20 cols) | 195.4MB | 0.0MB | 0.0MB |
| Read Large Data (50k rows x 20 cols) (async) | 17.2MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) | 6.6MB | 0.1MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (async) | 0.1MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) | 132.1MB | 0.0MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (async) | 17.6MB | N/A | N/A |
| Read Formulas (10k rows) | 9.0MB | 0.2MB | 0.0MB |
| Read Formulas (10k rows) (async) | 0.0MB | N/A | N/A |
| Read Strings (20k rows text-heavy) | 2.5MB | 0.1MB | 0.0MB |
| Read Strings (20k rows text-heavy) (async) | 0.0MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) | 0.0MB | 0.0MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (async) | 0.0MB | N/A | N/A |
| Read Comments (2k rows with comments) | 0.6MB | 0.0MB | 0.0MB |
| Read Comments (2k rows with comments) (async) | 0.0MB | N/A | N/A |
| Read Merged Cells (500 regions) | 0.0MB | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) (async) | 0.0MB | N/A | N/A |
| Read Mixed Workload (ERP document) | 0.0MB | 0.3MB | 0.0MB |
| Read Mixed Workload (ERP document) (async) | 0.0MB | N/A | N/A |
| Read Scale 1k rows | 0.0MB | 0.0MB | 0.0MB |
| Read Scale 1k rows (async) | 0.0MB | N/A | N/A |
| Read Scale 10k rows | 2.1MB | 0.0MB | 0.0MB |
| Read Scale 10k rows (async) | 0.0MB | N/A | N/A |
| Read Scale 100k rows | 161.1MB | 0.4MB | 0.0MB |
| Read Scale 100k rows (async) | 0.0MB | N/A | N/A |
| Write 50000 rows x 20 cols | 89.4MB | 6.1MB | 0.0MB |
| Write 5000 styled rows | 0.0MB | 0.1MB | 0.0MB |
| Write 10 sheets x 5000 rows | 38.2MB | 0.1MB | 0.0MB |
| Write 10000 rows with formulas | 2.3MB | 0.0MB | 0.0MB |
| Write 20000 text-heavy rows | 15.7MB | 0.1MB | 0.0MB |
| Write 5000 rows + 8 validation rules | 1.2MB | 0.2MB | N/A |
| Write 2000 rows with comments | 0.0MB | 0.1MB | 0.0MB |
| Write 500 merged regions | 0.0MB | 0.0MB | 0.0MB |
| Write 1k rows x 10 cols | 0.0MB | 0.0MB | 0.0MB |
| Write 10k rows x 10 cols | 0.0MB | 0.2MB | 0.0MB |
| Write 50k rows x 10 cols | 39.1MB | 0.3MB | 0.0MB |
| Write 100k rows x 10 cols | 70.4MB | 0.1MB | 0.0MB |
| Buffer round-trip (10000 rows) | 0.0MB | 0.2MB | 0.0MB |
| Streaming write (50000 rows) | 80.0MB | 0.0MB | N/A |
| Random-access read (1000 cells from 50k-row file) | 18.2MB | 0.0MB | 0.0MB |
| Random-access read (1000 cells from 50k-row file) (async) | 4.2MB | N/A | N/A |
| Mixed workload write (ERP-style) | 0.0MB | 0.2MB | N/A |

## Summary

Total scenarios: 41

| Library | Wins |
|---------|------|
| SheetKit | 40/41 |
| SheetJS | 1/41 |
| ExcelJS | 0/41 |
