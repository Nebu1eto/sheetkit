# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-11T04:21:31.870Z

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
| Read Large Data (50k rows x 20 cols) | 454ms | 2.69s | 1.59s | SheetKit |
| Read Large Data (50k rows x 20 cols) (async) | 435ms | N/A | N/A | SheetKit |
| Read Heavy Styles (5k rows, formatted) | 24ms | 175ms | 79ms | SheetKit |
| Read Heavy Styles (5k rows, formatted) (async) | 23ms | N/A | N/A | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) | 530ms | 1.55s | 590ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) (async) | 525ms | N/A | N/A | SheetKit |
| Read Formulas (10k rows) | 33ms | 196ms | 78ms | SheetKit |
| Read Formulas (10k rows) (async) | 32ms | N/A | N/A | SheetKit |
| Read Strings (20k rows text-heavy) | 95ms | 610ms | 260ms | SheetKit |
| Read Strings (20k rows text-heavy) (async) | 95ms | N/A | N/A | SheetKit |
| Read Data Validation (5k rows, 8 rules) | 19ms | 131ms | 56ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) (async) | 19ms | N/A | N/A | SheetKit |
| Read Comments (2k rows with comments) | 8ms | 113ms | 27ms | SheetKit |
| Read Comments (2k rows with comments) (async) | 8ms | N/A | N/A | SheetKit |
| Read Merged Cells (500 regions) | 2ms | 21ms | 5ms | SheetKit |
| Read Merged Cells (500 regions) (async) | 1ms | N/A | N/A | SheetKit |
| Read Mixed Workload (ERP document) | 27ms | 191ms | 71ms | SheetKit |
| Read Mixed Workload (ERP document) (async) | 27ms | N/A | N/A | SheetKit |

### Read (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Scale 1k rows | 5ms | 39ms | 18ms | SheetKit |
| Read Scale 1k rows (async) | 5ms | N/A | N/A | SheetKit |
| Read Scale 10k rows | 45ms | 292ms | 132ms | SheetKit |
| Read Scale 10k rows (async) | 44ms | N/A | N/A | SheetKit |
| Read Scale 100k rows | 474ms | 2.84s | 1.51s | SheetKit |
| Read Scale 100k rows (async) | 474ms | N/A | N/A | SheetKit |

### Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 50000 rows x 20 cols | 461ms | 2.55s | 1.25s | SheetKit |
| Write 5000 styled rows | 35ms | 171ms | 41ms | SheetKit |
| Write 10 sheets x 5000 rows | 251ms | 1.25s | 415ms | SheetKit |
| Write 10000 rows with formulas | 28ms | 152ms | 55ms | SheetKit |
| Write 20000 text-heavy rows | 87ms | 458ms | 199ms | SheetKit |

### Write (DV)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 9ms | 84ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 2000 rows with comments | 8ms | 62ms | 61ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 500 merged regions | 10ms | 26ms | 3ms | SheetJS |

### Write (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 1k rows x 10 cols | 5ms | 37ms | 9ms | SheetKit |
| Write 10k rows x 10 cols | 47ms | 250ms | 79ms | SheetKit |
| Write 50k rows x 10 cols | 235ms | 1.24s | 480ms | SheetKit |
| Write 100k rows x 10 cols | 476ms | 2.64s | 1.22s | SheetKit |

### Round-Trip

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 118ms | 479ms | 147ms | SheetKit |

### Streaming

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Streaming write (50000 rows) | 309ms | 499ms | N/A | SheetKit |

### Random Access

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Random-access read (1000 cells from 50k-row file) | 387ms | 2.81s | 1.29s | SheetKit |
| Random-access read (1000 cells from 50k-row file) (async) | 382ms | N/A | N/A | SheetKit |

### Mixed Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Mixed workload write (ERP-style) | 19ms | 104ms | N/A | SheetKit |

### Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Memory (median) |
|----------|---------|--------|-----|-----|-----|-----------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 454ms | 443ms | 457ms | 457ms | 195.3MB |
| Read Large Data (50k rows x 20 cols) | ExcelJS | 2.69s | 2.67s | 2.72s | 2.72s | 0.2MB |
| Read Large Data (50k rows x 20 cols) | SheetJS | 1.59s | 1.51s | 1.62s | 1.62s | 0.1MB |
| Read Large Data (50k rows x 20 cols) (async) | SheetKit | 435ms | 432ms | 452ms | 452ms | 17.2MB |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 24ms | 23ms | 25ms | 25ms | 6.0MB |
| Read Heavy Styles (5k rows, formatted) | ExcelJS | 175ms | 169ms | 176ms | 176ms | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | SheetJS | 79ms | 77ms | 82ms | 82ms | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (async) | SheetKit | 23ms | 23ms | 24ms | 24ms | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 530ms | 526ms | 532ms | 532ms | 112.7MB |
| Read Multi-Sheet (10 sheets x 5k rows) | ExcelJS | 1.55s | 1.54s | 1.56s | 1.56s | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetJS | 590ms | 582ms | 596ms | 596ms | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (async) | SheetKit | 525ms | 523ms | 530ms | 530ms | 0.4MB |
| Read Formulas (10k rows) | SheetKit | 33ms | 33ms | 35ms | 35ms | 9.3MB |
| Read Formulas (10k rows) | ExcelJS | 196ms | 190ms | 197ms | 197ms | 0.1MB |
| Read Formulas (10k rows) | SheetJS | 78ms | 77ms | 78ms | 78ms | 0.0MB |
| Read Formulas (10k rows) (async) | SheetKit | 32ms | 32ms | 35ms | 35ms | 0.0MB |
| Read Strings (20k rows text-heavy) | SheetKit | 95ms | 93ms | 95ms | 95ms | 2.5MB |
| Read Strings (20k rows text-heavy) | ExcelJS | 610ms | 602ms | 623ms | 623ms | 0.0MB |
| Read Strings (20k rows text-heavy) | SheetJS | 260ms | 255ms | 271ms | 271ms | 0.0MB |
| Read Strings (20k rows text-heavy) (async) | SheetKit | 95ms | 94ms | 96ms | 96ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | SheetKit | 19ms | 18ms | 20ms | 20ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | ExcelJS | 131ms | 127ms | 138ms | 138ms | 3.2MB |
| Read Data Validation (5k rows, 8 rules) | SheetJS | 56ms | 55ms | 56ms | 56ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (async) | SheetKit | 19ms | 18ms | 19ms | 19ms | 0.0MB |
| Read Comments (2k rows with comments) | SheetKit | 8ms | 8ms | 8ms | 8ms | 0.5MB |
| Read Comments (2k rows with comments) | ExcelJS | 113ms | 110ms | 117ms | 117ms | 0.0MB |
| Read Comments (2k rows with comments) | SheetJS | 27ms | 24ms | 29ms | 29ms | 0.0MB |
| Read Comments (2k rows with comments) (async) | SheetKit | 8ms | 7ms | 8ms | 8ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetKit | 2ms | 1ms | 2ms | 2ms | 0.0MB |
| Read Merged Cells (500 regions) | ExcelJS | 21ms | 20ms | 23ms | 23ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetJS | 5ms | 4ms | 5ms | 5ms | 0.0MB |
| Read Merged Cells (500 regions) (async) | SheetKit | 1ms | 1ms | 1ms | 1ms | 0.0MB |
| Read Mixed Workload (ERP document) | SheetKit | 27ms | 27ms | 28ms | 28ms | 0.0MB |
| Read Mixed Workload (ERP document) | ExcelJS | 191ms | 183ms | 193ms | 193ms | 0.1MB |
| Read Mixed Workload (ERP document) | SheetJS | 71ms | 69ms | 72ms | 72ms | 0.0MB |
| Read Mixed Workload (ERP document) (async) | SheetKit | 27ms | 25ms | 27ms | 27ms | 0.0MB |
| Read Scale 1k rows | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0MB |
| Read Scale 1k rows | ExcelJS | 39ms | 37ms | 39ms | 39ms | 0.0MB |
| Read Scale 1k rows | SheetJS | 18ms | 17ms | 19ms | 19ms | 0.0MB |
| Read Scale 1k rows (async) | SheetKit | 5ms | 4ms | 5ms | 5ms | 0.0MB |
| Read Scale 10k rows | SheetKit | 45ms | 44ms | 47ms | 47ms | 2.1MB |
| Read Scale 10k rows | ExcelJS | 292ms | 286ms | 310ms | 310ms | 0.1MB |
| Read Scale 10k rows | SheetJS | 132ms | 129ms | 136ms | 136ms | 0.0MB |
| Read Scale 10k rows (async) | SheetKit | 44ms | 43ms | 47ms | 47ms | 0.0MB |
| Read Scale 100k rows | SheetKit | 474ms | 473ms | 484ms | 484ms | 175.2MB |
| Read Scale 100k rows | ExcelJS | 2.84s | 2.79s | 2.86s | 2.86s | 0.0MB |
| Read Scale 100k rows | SheetJS | 1.51s | 1.47s | 1.53s | 1.53s | 0.0MB |
| Read Scale 100k rows (async) | SheetKit | 474ms | 469ms | 500ms | 500ms | 15.9MB |
| Write 50000 rows x 20 cols | SheetKit | 461ms | 459ms | 507ms | 507ms | 67.3MB |
| Write 50000 rows x 20 cols | ExcelJS | 2.55s | 2.47s | 2.68s | 2.68s | 54.4MB |
| Write 50000 rows x 20 cols | SheetJS | 1.25s | 1.18s | 1.54s | 1.54s | 0.0MB |
| Write 5000 styled rows | SheetKit | 35ms | 34ms | 35ms | 35ms | 0.0MB |
| Write 5000 styled rows | ExcelJS | 171ms | 154ms | 175ms | 175ms | 0.1MB |
| Write 5000 styled rows | SheetJS | 41ms | 39ms | 42ms | 42ms | 0.0MB |
| Write 10 sheets x 5000 rows | SheetKit | 251ms | 248ms | 253ms | 253ms | 38.2MB |
| Write 10 sheets x 5000 rows | ExcelJS | 1.25s | 1.23s | 1.35s | 1.35s | 0.3MB |
| Write 10 sheets x 5000 rows | SheetJS | 415ms | 388ms | 432ms | 432ms | 0.0MB |
| Write 10000 rows with formulas | SheetKit | 28ms | 28ms | 29ms | 29ms | 0.0MB |
| Write 10000 rows with formulas | ExcelJS | 152ms | 148ms | 156ms | 156ms | 0.2MB |
| Write 10000 rows with formulas | SheetJS | 55ms | 53ms | 60ms | 60ms | 0.0MB |
| Write 20000 text-heavy rows | SheetKit | 87ms | 86ms | 87ms | 87ms | 12.2MB |
| Write 20000 text-heavy rows | ExcelJS | 458ms | 447ms | 460ms | 460ms | 0.3MB |
| Write 20000 text-heavy rows | SheetJS | 199ms | 198ms | 212ms | 212ms | 0.0MB |
| Write 5000 rows + 8 validation rules | SheetKit | 9ms | 9ms | 9ms | 9ms | 1.2MB |
| Write 5000 rows + 8 validation rules | ExcelJS | 84ms | 81ms | 89ms | 89ms | 0.3MB |
| Write 2000 rows with comments | SheetKit | 8ms | 7ms | 8ms | 8ms | 0.0MB |
| Write 2000 rows with comments | ExcelJS | 62ms | 62ms | 64ms | 64ms | 0.1MB |
| Write 2000 rows with comments | SheetJS | 61ms | 61ms | 62ms | 62ms | 0.0MB |
| Write 500 merged regions | SheetKit | 10ms | 10ms | 11ms | 11ms | 0.0MB |
| Write 500 merged regions | ExcelJS | 26ms | 18ms | 28ms | 28ms | 0.0MB |
| Write 500 merged regions | SheetJS | 3ms | 3ms | 3ms | 3ms | 0.0MB |
| Write 1k rows x 10 cols | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0MB |
| Write 1k rows x 10 cols | ExcelJS | 37ms | 35ms | 39ms | 39ms | 0.2MB |
| Write 1k rows x 10 cols | SheetJS | 9ms | 9ms | 9ms | 9ms | 0.0MB |
| Write 10k rows x 10 cols | SheetKit | 47ms | 46ms | 48ms | 48ms | 0.1MB |
| Write 10k rows x 10 cols | ExcelJS | 250ms | 250ms | 252ms | 252ms | 0.2MB |
| Write 10k rows x 10 cols | SheetJS | 79ms | 77ms | 82ms | 82ms | 0.0MB |
| Write 50k rows x 10 cols | SheetKit | 235ms | 229ms | 235ms | 235ms | 35.0MB |
| Write 50k rows x 10 cols | ExcelJS | 1.24s | 1.22s | 1.26s | 1.26s | 0.2MB |
| Write 50k rows x 10 cols | SheetJS | 480ms | 469ms | 523ms | 523ms | 0.0MB |
| Write 100k rows x 10 cols | SheetKit | 476ms | 474ms | 478ms | 478ms | 65.1MB |
| Write 100k rows x 10 cols | ExcelJS | 2.64s | 2.60s | 2.79s | 2.79s | 0.0MB |
| Write 100k rows x 10 cols | SheetJS | 1.22s | 1.21s | 1.25s | 1.25s | 0.0MB |
| Buffer round-trip (10000 rows) | SheetKit | 118ms | 115ms | 120ms | 120ms | 1.4MB |
| Buffer round-trip (10000 rows) | ExcelJS | 479ms | 470ms | 482ms | 482ms | 0.1MB |
| Buffer round-trip (10000 rows) | SheetJS | 147ms | 140ms | 167ms | 167ms | 0.0MB |
| Streaming write (50000 rows) | SheetKit | 309ms | 308ms | 309ms | 309ms | 0.0MB |
| Streaming write (50000 rows) | ExcelJS | 499ms | 497ms | 504ms | 504ms | 0.0MB |
| Random-access read (1000 cells from 50k-row file) | SheetKit | 387ms | 385ms | 390ms | 390ms | 61.6MB |
| Random-access read (1000 cells from 50k-row file) | ExcelJS | 2.81s | 2.72s | 2.90s | 2.90s | 0.0MB |
| Random-access read (1000 cells from 50k-row file) | SheetJS | 1.29s | 1.25s | 1.32s | 1.32s | 0.0MB |
| Random-access read (1000 cells from 50k-row file) (async) | SheetKit | 382ms | 381ms | 385ms | 385ms | 0.0MB |
| Mixed workload write (ERP-style) | SheetKit | 19ms | 19ms | 20ms | 20ms | 0.0MB |
| Mixed workload write (ERP-style) | ExcelJS | 104ms | 102ms | 106ms | 106ms | 0.3MB |

### Memory Usage

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| Read Large Data (50k rows x 20 cols) | 195.3MB | 0.2MB | 0.1MB |
| Read Large Data (50k rows x 20 cols) (async) | 17.2MB | N/A | N/A |
| Read Heavy Styles (5k rows, formatted) | 6.0MB | 0.0MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) (async) | 0.0MB | N/A | N/A |
| Read Multi-Sheet (10 sheets x 5k rows) | 112.7MB | 0.0MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) (async) | 0.4MB | N/A | N/A |
| Read Formulas (10k rows) | 9.3MB | 0.1MB | 0.0MB |
| Read Formulas (10k rows) (async) | 0.0MB | N/A | N/A |
| Read Strings (20k rows text-heavy) | 2.5MB | 0.0MB | 0.0MB |
| Read Strings (20k rows text-heavy) (async) | 0.0MB | N/A | N/A |
| Read Data Validation (5k rows, 8 rules) | 0.0MB | 3.2MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) (async) | 0.0MB | N/A | N/A |
| Read Comments (2k rows with comments) | 0.5MB | 0.0MB | 0.0MB |
| Read Comments (2k rows with comments) (async) | 0.0MB | N/A | N/A |
| Read Merged Cells (500 regions) | 0.0MB | 0.0MB | 0.0MB |
| Read Merged Cells (500 regions) (async) | 0.0MB | N/A | N/A |
| Read Mixed Workload (ERP document) | 0.0MB | 0.1MB | 0.0MB |
| Read Mixed Workload (ERP document) (async) | 0.0MB | N/A | N/A |
| Read Scale 1k rows | 0.0MB | 0.0MB | 0.0MB |
| Read Scale 1k rows (async) | 0.0MB | N/A | N/A |
| Read Scale 10k rows | 2.1MB | 0.1MB | 0.0MB |
| Read Scale 10k rows (async) | 0.0MB | N/A | N/A |
| Read Scale 100k rows | 175.2MB | 0.0MB | 0.0MB |
| Read Scale 100k rows (async) | 15.9MB | N/A | N/A |
| Write 50000 rows x 20 cols | 67.3MB | 54.4MB | 0.0MB |
| Write 5000 styled rows | 0.0MB | 0.1MB | 0.0MB |
| Write 10 sheets x 5000 rows | 38.2MB | 0.3MB | 0.0MB |
| Write 10000 rows with formulas | 0.0MB | 0.2MB | 0.0MB |
| Write 20000 text-heavy rows | 12.2MB | 0.3MB | 0.0MB |
| Write 5000 rows + 8 validation rules | 1.2MB | 0.3MB | N/A |
| Write 2000 rows with comments | 0.0MB | 0.1MB | 0.0MB |
| Write 500 merged regions | 0.0MB | 0.0MB | 0.0MB |
| Write 1k rows x 10 cols | 0.0MB | 0.2MB | 0.0MB |
| Write 10k rows x 10 cols | 0.1MB | 0.2MB | 0.0MB |
| Write 50k rows x 10 cols | 35.0MB | 0.2MB | 0.0MB |
| Write 100k rows x 10 cols | 65.1MB | 0.0MB | 0.0MB |
| Buffer round-trip (10000 rows) | 1.4MB | 0.1MB | 0.0MB |
| Streaming write (50000 rows) | 0.0MB | 0.0MB | N/A |
| Random-access read (1000 cells from 50k-row file) | 61.6MB | 0.0MB | 0.0MB |
| Random-access read (1000 cells from 50k-row file) (async) | 0.0MB | N/A | N/A |
| Mixed workload write (ERP-style) | 0.0MB | 0.3MB | N/A |

## Summary

Total scenarios: 41

| Library | Wins |
|---------|------|
| SheetKit | 40/41 |
| SheetJS | 1/41 |
| ExcelJS | 0/41 |
