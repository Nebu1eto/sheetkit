# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-09T13:56:11.832Z

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
| Read Large Data (50k rows x 20 cols) | 1.26s | 3.82s | 2.11s | SheetKit |
| Read Heavy Styles (5k rows, formatted) | 59ms | 243ms | 113ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) | 617ms | 2.10s | 860ms | SheetKit |
| Read Formulas (10k rows) | 78ms | 264ms | 120ms | SheetKit |
| Read Strings (20k rows text-heavy) | 240ms | 845ms | 400ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) | 46ms | 179ms | 86ms | SheetKit |
| Read Comments (2k rows with comments) | 15ms | 154ms | 44ms | SheetKit |
| Read Merged Cells (500 regions) | 3ms | 30ms | 8ms | SheetKit |
| Read Mixed Workload (ERP document) | 63ms | 268ms | 110ms | SheetKit |

### Read (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Scale 1k rows | 12ms | 56ms | 31ms | SheetKit |
| Read Scale 10k rows | 117ms | 402ms | 210ms | SheetKit |
| Read Scale 100k rows | 1.23s | 4.06s | 2.13s | SheetKit |

### Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 50000 rows x 20 cols | 750ms | 3.58s | 1.66s | SheetKit |
| Write 5000 styled rows | 59ms | 228ms | 59ms | SheetKit |
| Write 10 sheets x 5000 rows | 365ms | 1.82s | 559ms | SheetKit |
| Write 10000 rows with formulas | 42ms | 220ms | 78ms | SheetKit |
| Write 20000 text-heavy rows | 135ms | 658ms | 273ms | SheetKit |

### Write (DV)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 14ms | 124ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 2000 rows with comments | 11ms | 92ms | 95ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 500 merged regions | 15ms | 37ms | 4ms | SheetJS |

### Write (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 1k rows x 10 cols | 7ms | 50ms | 13ms | SheetKit |
| Write 10k rows x 10 cols | 68ms | 359ms | 116ms | SheetKit |
| Write 50k rows x 10 cols | 335ms | 1.79s | 705ms | SheetKit |
| Write 100k rows x 10 cols | 755ms | 3.82s | 1.58s | SheetKit |

### Round-Trip

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 218ms | 678ms | 237ms | SheetKit |

### Streaming

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Streaming write (50000 rows) | 1.20s | 716ms | N/A | ExcelJS |

### Random Access

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Random-access read (1000 cells from 50k-row file) | 549ms | 3.93s | 1.75s | SheetKit |

### Mixed Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Mixed workload write (ERP-style) | 29ms | 156ms | N/A | SheetKit |

### Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Memory (median) |
|----------|---------|--------|-----|-----|-----|-----------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 1.26s | 1.21s | 1.28s | 1.28s | 405.6MB |
| Read Large Data (50k rows x 20 cols) | ExcelJS | 3.82s | 3.79s | 3.98s | 3.98s | 0.6MB |
| Read Large Data (50k rows x 20 cols) | SheetJS | 2.11s | 2.10s | 2.13s | 2.13s | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 59ms | 59ms | 60ms | 60ms | 20.0MB |
| Read Heavy Styles (5k rows, formatted) | ExcelJS | 243ms | 235ms | 251ms | 251ms | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | SheetJS | 113ms | 108ms | 115ms | 115ms | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 617ms | 611ms | 618ms | 618ms | 207.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | ExcelJS | 2.10s | 2.08s | 2.11s | 2.11s | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetJS | 860ms | 854ms | 912ms | 912ms | 0.8MB |
| Read Formulas (10k rows) | SheetKit | 78ms | 76ms | 83ms | 83ms | 16.3MB |
| Read Formulas (10k rows) | ExcelJS | 264ms | 259ms | 269ms | 269ms | 0.3MB |
| Read Formulas (10k rows) | SheetJS | 120ms | 119ms | 123ms | 123ms | 0.0MB |
| Read Strings (20k rows text-heavy) | SheetKit | 240ms | 239ms | 242ms | 242ms | 12.4MB |
| Read Strings (20k rows text-heavy) | ExcelJS | 845ms | 843ms | 852ms | 852ms | 0.2MB |
| Read Strings (20k rows text-heavy) | SheetJS | 400ms | 398ms | 407ms | 407ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | SheetKit | 46ms | 46ms | 47ms | 47ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | ExcelJS | 179ms | 175ms | 184ms | 184ms | 3.3MB |
| Read Data Validation (5k rows, 8 rules) | SheetJS | 86ms | 84ms | 88ms | 88ms | 0.0MB |
| Read Comments (2k rows with comments) | SheetKit | 15ms | 15ms | 16ms | 16ms | 0.6MB |
| Read Comments (2k rows with comments) | ExcelJS | 154ms | 149ms | 160ms | 160ms | 0.1MB |
| Read Comments (2k rows with comments) | SheetJS | 44ms | 36ms | 45ms | 45ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetKit | 3ms | 3ms | 4ms | 4ms | 0.2MB |
| Read Merged Cells (500 regions) | ExcelJS | 30ms | 30ms | 32ms | 32ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetJS | 8ms | 6ms | 9ms | 9ms | 0.0MB |
| Read Mixed Workload (ERP document) | SheetKit | 63ms | 62ms | 68ms | 68ms | 0.0MB |
| Read Mixed Workload (ERP document) | ExcelJS | 268ms | 253ms | 273ms | 273ms | 0.3MB |
| Read Mixed Workload (ERP document) | SheetJS | 110ms | 101ms | 113ms | 113ms | 0.0MB |
| Read Scale 1k rows | SheetKit | 12ms | 11ms | 14ms | 14ms | 0.0MB |
| Read Scale 1k rows | ExcelJS | 56ms | 55ms | 57ms | 57ms | 0.0MB |
| Read Scale 1k rows | SheetJS | 31ms | 21ms | 84ms | 84ms | 0.0MB |
| Read Scale 10k rows | SheetKit | 117ms | 114ms | 121ms | 121ms | 23.8MB |
| Read Scale 10k rows | ExcelJS | 402ms | 390ms | 428ms | 428ms | 0.1MB |
| Read Scale 10k rows | SheetJS | 210ms | 204ms | 210ms | 210ms | 0.0MB |
| Read Scale 100k rows | SheetKit | 1.23s | 1.21s | 1.26s | 1.26s | 361.1MB |
| Read Scale 100k rows | ExcelJS | 4.06s | 4.03s | 4.10s | 4.10s | 0.0MB |
| Read Scale 100k rows | SheetJS | 2.13s | 2.12s | 2.19s | 2.19s | 0.0MB |
| Write 50000 rows x 20 cols | SheetKit | 750ms | 724ms | 824ms | 824ms | 255.0MB |
| Write 50000 rows x 20 cols | ExcelJS | 3.58s | 3.52s | 3.76s | 3.76s | 28.2MB |
| Write 50000 rows x 20 cols | SheetJS | 1.66s | 1.58s | 2.17s | 2.17s | 0.0MB |
| Write 5000 styled rows | SheetKit | 59ms | 50ms | 84ms | 84ms | 20.2MB |
| Write 5000 styled rows | ExcelJS | 228ms | 218ms | 238ms | 238ms | 0.3MB |
| Write 5000 styled rows | SheetJS | 59ms | 58ms | 63ms | 63ms | 0.0MB |
| Write 10 sheets x 5000 rows | SheetKit | 365ms | 360ms | 388ms | 388ms | 204.3MB |
| Write 10 sheets x 5000 rows | ExcelJS | 1.82s | 1.81s | 1.84s | 1.84s | 0.3MB |
| Write 10 sheets x 5000 rows | SheetJS | 559ms | 489ms | 603ms | 603ms | 0.0MB |
| Write 10000 rows with formulas | SheetKit | 42ms | 41ms | 42ms | 42ms | 16.0MB |
| Write 10000 rows with formulas | ExcelJS | 220ms | 210ms | 223ms | 223ms | 0.0MB |
| Write 10000 rows with formulas | SheetJS | 78ms | 74ms | 80ms | 80ms | 0.0MB |
| Write 20000 text-heavy rows | SheetKit | 135ms | 133ms | 141ms | 141ms | 17.4MB |
| Write 20000 text-heavy rows | ExcelJS | 658ms | 628ms | 681ms | 681ms | 0.1MB |
| Write 20000 text-heavy rows | SheetJS | 273ms | 270ms | 294ms | 294ms | 0.0MB |
| Write 5000 rows + 8 validation rules | SheetKit | 14ms | 13ms | 15ms | 15ms | 3.7MB |
| Write 5000 rows + 8 validation rules | ExcelJS | 124ms | 117ms | 128ms | 128ms | 0.2MB |
| Write 2000 rows with comments | SheetKit | 11ms | 11ms | 11ms | 11ms | 0.1MB |
| Write 2000 rows with comments | ExcelJS | 92ms | 89ms | 95ms | 95ms | 0.1MB |
| Write 2000 rows with comments | SheetJS | 95ms | 93ms | 100ms | 100ms | 0.0MB |
| Write 500 merged regions | SheetKit | 15ms | 14ms | 15ms | 15ms | 0.0MB |
| Write 500 merged regions | ExcelJS | 37ms | 26ms | 41ms | 41ms | 0.0MB |
| Write 500 merged regions | SheetJS | 4ms | 4ms | 4ms | 4ms | 0.0MB |
| Write 1k rows x 10 cols | SheetKit | 7ms | 7ms | 7ms | 7ms | 0.0MB |
| Write 1k rows x 10 cols | ExcelJS | 50ms | 40ms | 53ms | 53ms | 0.0MB |
| Write 1k rows x 10 cols | SheetJS | 13ms | 12ms | 13ms | 13ms | 0.0MB |
| Write 10k rows x 10 cols | SheetKit | 68ms | 66ms | 69ms | 69ms | 0.0MB |
| Write 10k rows x 10 cols | ExcelJS | 359ms | 350ms | 365ms | 365ms | 0.3MB |
| Write 10k rows x 10 cols | SheetJS | 116ms | 111ms | 128ms | 128ms | 0.0MB |
| Write 50k rows x 10 cols | SheetKit | 335ms | 325ms | 356ms | 356ms | 56.6MB |
| Write 50k rows x 10 cols | ExcelJS | 1.79s | 1.76s | 1.85s | 1.85s | 0.4MB |
| Write 50k rows x 10 cols | SheetJS | 705ms | 699ms | 1.04s | 1.04s | 0.0MB |
| Write 100k rows x 10 cols | SheetKit | 755ms | 699ms | 782ms | 782ms | 268.7MB |
| Write 100k rows x 10 cols | ExcelJS | 3.82s | 3.74s | 5.17s | 5.17s | 13.5MB |
| Write 100k rows x 10 cols | SheetJS | 1.58s | 1.55s | 1.69s | 1.69s | 0.0MB |
| Buffer round-trip (10000 rows) | SheetKit | 218ms | 215ms | 220ms | 220ms | 4.1MB |
| Buffer round-trip (10000 rows) | ExcelJS | 678ms | 670ms | 694ms | 694ms | 0.4MB |
| Buffer round-trip (10000 rows) | SheetJS | 237ms | 233ms | 240ms | 240ms | 0.0MB |
| Streaming write (50000 rows) | SheetKit | 1.20s | 1.18s | 1.21s | 1.21s | 323.4MB |
| Streaming write (50000 rows) | ExcelJS | 716ms | 712ms | 720ms | 720ms | 1.6MB |
| Random-access read (1000 cells from 50k-row file) | SheetKit | 549ms | 536ms | 551ms | 551ms | 29.8MB |
| Random-access read (1000 cells from 50k-row file) | ExcelJS | 3.93s | 3.89s | 4.24s | 4.24s | 0.0MB |
| Random-access read (1000 cells from 50k-row file) | SheetJS | 1.75s | 1.74s | 1.86s | 1.86s | 0.0MB |
| Mixed workload write (ERP-style) | SheetKit | 29ms | 28ms | 31ms | 31ms | 8.5MB |
| Mixed workload write (ERP-style) | ExcelJS | 156ms | 151ms | 165ms | 165ms | 0.4MB |

### Memory Usage

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| Read Large Data (50k rows x 20 cols) | 405.6MB | 0.6MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | 20.0MB | 0.0MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | 207.0MB | 0.0MB | 0.8MB |
| Read Formulas (10k rows) | 16.3MB | 0.3MB | 0.0MB |
| Read Strings (20k rows text-heavy) | 12.4MB | 0.2MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | 0.0MB | 3.3MB | 0.0MB |
| Read Comments (2k rows with comments) | 0.6MB | 0.1MB | 0.0MB |
| Read Merged Cells (500 regions) | 0.2MB | 0.0MB | 0.0MB |
| Read Mixed Workload (ERP document) | 0.0MB | 0.3MB | 0.0MB |
| Read Scale 1k rows | 0.0MB | 0.0MB | 0.0MB |
| Read Scale 10k rows | 23.8MB | 0.1MB | 0.0MB |
| Read Scale 100k rows | 361.1MB | 0.0MB | 0.0MB |
| Write 50000 rows x 20 cols | 255.0MB | 28.2MB | 0.0MB |
| Write 5000 styled rows | 20.2MB | 0.3MB | 0.0MB |
| Write 10 sheets x 5000 rows | 204.3MB | 0.3MB | 0.0MB |
| Write 10000 rows with formulas | 16.0MB | 0.0MB | 0.0MB |
| Write 20000 text-heavy rows | 17.4MB | 0.1MB | 0.0MB |
| Write 5000 rows + 8 validation rules | 3.7MB | 0.2MB | N/A |
| Write 2000 rows with comments | 0.1MB | 0.1MB | 0.0MB |
| Write 500 merged regions | 0.0MB | 0.0MB | 0.0MB |
| Write 1k rows x 10 cols | 0.0MB | 0.0MB | 0.0MB |
| Write 10k rows x 10 cols | 0.0MB | 0.3MB | 0.0MB |
| Write 50k rows x 10 cols | 56.6MB | 0.4MB | 0.0MB |
| Write 100k rows x 10 cols | 268.7MB | 13.5MB | 0.0MB |
| Buffer round-trip (10000 rows) | 4.1MB | 0.4MB | 0.0MB |
| Streaming write (50000 rows) | 323.4MB | 1.6MB | N/A |
| Random-access read (1000 cells from 50k-row file) | 29.8MB | 0.0MB | 0.0MB |
| Mixed workload write (ERP-style) | 8.5MB | 0.4MB | N/A |

## Summary

Total scenarios: 28

| Library | Wins |
|---------|------|
| SheetKit | 26/28 |
| ExcelJS | 1/28 |
| SheetJS | 1/28 |
