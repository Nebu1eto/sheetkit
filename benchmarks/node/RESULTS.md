# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-09T13:35:13.725Z

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
| Read Large Data (50k rows x 20 cols) | 1.26s | 4.14s | 2.11s | SheetKit |
| Read Heavy Styles (5k rows, formatted) | 62ms | 258ms | 122ms | SheetKit |
| Read Multi-Sheet (10 sheets x 5k rows) | 602ms | 2.23s | 883ms | SheetKit |
| Read Formulas (10k rows) | 76ms | 289ms | 119ms | SheetKit |
| Read Strings (20k rows text-heavy) | 234ms | 905ms | 378ms | SheetKit |
| Read Data Validation (5k rows, 8 rules) | 45ms | 191ms | 82ms | SheetKit |
| Read Comments (2k rows with comments) | 15ms | 162ms | 42ms | SheetKit |
| Read Merged Cells (500 regions) | 3ms | 30ms | 7ms | SheetKit |
| Read Mixed Workload (ERP document) | 58ms | 299ms | 110ms | SheetKit |

### Read (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Read Scale 1k rows | 11ms | 56ms | 26ms | SheetKit |
| Read Scale 10k rows | 113ms | 437ms | 208ms | SheetKit |
| Read Scale 100k rows | 1.23s | 4.27s | 2.12s | SheetKit |

### Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 50000 rows x 20 cols | 727ms | 3.58s | 1.69s | SheetKit |
| Write 5000 styled rows | 51ms | 226ms | 59ms | SheetKit |
| Write 10 sheets x 5000 rows | 369ms | 1.84s | 596ms | SheetKit |
| Write 10000 rows with formulas | 42ms | 212ms | 82ms | SheetKit |
| Write 20000 text-heavy rows | 130ms | 679ms | 297ms | SheetKit |

### Write (DV)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 5000 rows + 8 validation rules | 14ms | 124ms | N/A | SheetKit |

### Write (Comments)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 2000 rows with comments | 11ms | 91ms | 94ms | SheetKit |

### Write (Merge)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 500 merged regions | 14ms | 36ms | 4ms | SheetJS |

### Write (Scale)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Write 1k rows x 10 cols | 7ms | 53ms | 13ms | SheetKit |
| Write 10k rows x 10 cols | 68ms | 377ms | 120ms | SheetKit |
| Write 50k rows x 10 cols | 363ms | 1.84s | 702ms | SheetKit |
| Write 100k rows x 10 cols | 699ms | 3.90s | 1.64s | SheetKit |

### Round-Trip

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Buffer round-trip (10000 rows) | 221ms | 703ms | 223ms | SheetKit |

### Streaming

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Streaming write (50000 rows) | 1.22s | 757ms | N/A | ExcelJS |

### Random Access

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Random-access read (1000 cells from 50k-row file) | 579ms | 4.15s | 1.82s | SheetKit |

### Mixed Write

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Mixed workload write (ERP-style) | 29ms | 159ms | N/A | SheetKit |

### Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Memory (median) |
|----------|---------|--------|-----|-----|-----|-----------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 1.26s | 1.20s | 1.59s | 1.59s | 0.0MB |
| Read Large Data (50k rows x 20 cols) | ExcelJS | 4.14s | 3.97s | 5.28s | 5.28s | 0.2MB |
| Read Large Data (50k rows x 20 cols) | SheetJS | 2.11s | 2.10s | 2.65s | 2.65s | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 62ms | 60ms | 63ms | 63ms | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | ExcelJS | 258ms | 256ms | 262ms | 262ms | 0.2MB |
| Read Heavy Styles (5k rows, formatted) | SheetJS | 122ms | 119ms | 140ms | 140ms | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 602ms | 595ms | 613ms | 613ms | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | ExcelJS | 2.23s | 2.22s | 2.26s | 2.26s | 0.2MB |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetJS | 883ms | 868ms | 935ms | 935ms | 0.0MB |
| Read Formulas (10k rows) | SheetKit | 76ms | 75ms | 78ms | 78ms | 0.0MB |
| Read Formulas (10k rows) | ExcelJS | 289ms | 286ms | 307ms | 307ms | 0.1MB |
| Read Formulas (10k rows) | SheetJS | 119ms | 116ms | 125ms | 125ms | 0.0MB |
| Read Strings (20k rows text-heavy) | SheetKit | 234ms | 230ms | 241ms | 241ms | 0.0MB |
| Read Strings (20k rows text-heavy) | ExcelJS | 905ms | 882ms | 1.07s | 1.07s | 0.2MB |
| Read Strings (20k rows text-heavy) | SheetJS | 378ms | 373ms | 395ms | 395ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | SheetKit | 45ms | 45ms | 47ms | 47ms | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | ExcelJS | 191ms | 180ms | 196ms | 196ms | 16.5MB |
| Read Data Validation (5k rows, 8 rules) | SheetJS | 82ms | 78ms | 83ms | 83ms | 0.0MB |
| Read Comments (2k rows with comments) | SheetKit | 15ms | 15ms | 15ms | 15ms | 0.0MB |
| Read Comments (2k rows with comments) | ExcelJS | 162ms | 158ms | 168ms | 168ms | 0.2MB |
| Read Comments (2k rows with comments) | SheetJS | 42ms | 35ms | 43ms | 43ms | 0.0MB |
| Read Merged Cells (500 regions) | SheetKit | 3ms | 3ms | 3ms | 3ms | 0.0MB |
| Read Merged Cells (500 regions) | ExcelJS | 30ms | 29ms | 34ms | 34ms | 0.1MB |
| Read Merged Cells (500 regions) | SheetJS | 7ms | 5ms | 8ms | 8ms | 0.0MB |
| Read Mixed Workload (ERP document) | SheetKit | 58ms | 58ms | 80ms | 80ms | 0.0MB |
| Read Mixed Workload (ERP document) | ExcelJS | 299ms | 277ms | 364ms | 364ms | 0.3MB |
| Read Mixed Workload (ERP document) | SheetJS | 110ms | 107ms | 113ms | 113ms | 0.0MB |
| Read Scale 1k rows | SheetKit | 11ms | 11ms | 11ms | 11ms | 0.0MB |
| Read Scale 1k rows | ExcelJS | 56ms | 53ms | 58ms | 58ms | 0.1MB |
| Read Scale 1k rows | SheetJS | 26ms | 20ms | 29ms | 29ms | 0.0MB |
| Read Scale 10k rows | SheetKit | 113ms | 111ms | 114ms | 114ms | 0.0MB |
| Read Scale 10k rows | ExcelJS | 437ms | 410ms | 443ms | 443ms | 0.2MB |
| Read Scale 10k rows | SheetJS | 208ms | 198ms | 215ms | 215ms | 0.0MB |
| Read Scale 100k rows | SheetKit | 1.23s | 1.21s | 1.25s | 1.25s | 0.0MB |
| Read Scale 100k rows | ExcelJS | 4.27s | 4.21s | 4.46s | 4.46s | 0.2MB |
| Read Scale 100k rows | SheetJS | 2.12s | 2.11s | 2.18s | 2.18s | 0.0MB |
| Write 50000 rows x 20 cols | SheetKit | 727ms | 706ms | 743ms | 743ms | 0.0MB |
| Write 50000 rows x 20 cols | ExcelJS | 3.58s | 3.48s | 3.81s | 3.81s | 0.2MB |
| Write 50000 rows x 20 cols | SheetJS | 1.69s | 1.63s | 2.64s | 2.64s | 0.0MB |
| Write 5000 styled rows | SheetKit | 51ms | 49ms | 52ms | 52ms | 0.0MB |
| Write 5000 styled rows | ExcelJS | 226ms | 225ms | 239ms | 239ms | 0.2MB |
| Write 5000 styled rows | SheetJS | 59ms | 56ms | 61ms | 61ms | 0.0MB |
| Write 10 sheets x 5000 rows | SheetKit | 369ms | 364ms | 393ms | 393ms | 0.0MB |
| Write 10 sheets x 5000 rows | ExcelJS | 1.84s | 1.78s | 1.90s | 1.90s | 0.3MB |
| Write 10 sheets x 5000 rows | SheetJS | 596ms | 570ms | 622ms | 622ms | 0.0MB |
| Write 10000 rows with formulas | SheetKit | 42ms | 41ms | 94ms | 94ms | 0.0MB |
| Write 10000 rows with formulas | ExcelJS | 212ms | 209ms | 228ms | 228ms | 0.2MB |
| Write 10000 rows with formulas | SheetJS | 82ms | 79ms | 86ms | 86ms | 0.0MB |
| Write 20000 text-heavy rows | SheetKit | 130ms | 129ms | 133ms | 133ms | 0.0MB |
| Write 20000 text-heavy rows | ExcelJS | 679ms | 658ms | 695ms | 695ms | 0.2MB |
| Write 20000 text-heavy rows | SheetJS | 297ms | 276ms | 304ms | 304ms | 0.0MB |
| Write 5000 rows + 8 validation rules | SheetKit | 14ms | 13ms | 15ms | 15ms | 0.0MB |
| Write 5000 rows + 8 validation rules | ExcelJS | 124ms | 116ms | 130ms | 130ms | 0.1MB |
| Write 2000 rows with comments | SheetKit | 11ms | 11ms | 11ms | 11ms | 0.0MB |
| Write 2000 rows with comments | ExcelJS | 91ms | 85ms | 93ms | 93ms | 0.1MB |
| Write 2000 rows with comments | SheetJS | 94ms | 91ms | 99ms | 99ms | 0.0MB |
| Write 500 merged regions | SheetKit | 14ms | 14ms | 15ms | 15ms | 0.0MB |
| Write 500 merged regions | ExcelJS | 36ms | 26ms | 39ms | 39ms | 0.1MB |
| Write 500 merged regions | SheetJS | 4ms | 4ms | 4ms | 4ms | 0.0MB |
| Write 1k rows x 10 cols | SheetKit | 7ms | 7ms | 7ms | 7ms | 0.0MB |
| Write 1k rows x 10 cols | ExcelJS | 53ms | 50ms | 56ms | 56ms | 0.1MB |
| Write 1k rows x 10 cols | SheetJS | 13ms | 12ms | 14ms | 14ms | 0.0MB |
| Write 10k rows x 10 cols | SheetKit | 68ms | 66ms | 73ms | 73ms | 0.0MB |
| Write 10k rows x 10 cols | ExcelJS | 377ms | 354ms | 383ms | 383ms | 0.2MB |
| Write 10k rows x 10 cols | SheetJS | 120ms | 118ms | 123ms | 123ms | 0.0MB |
| Write 50k rows x 10 cols | SheetKit | 363ms | 359ms | 457ms | 457ms | 0.0MB |
| Write 50k rows x 10 cols | ExcelJS | 1.84s | 1.79s | 1.86s | 1.86s | 0.2MB |
| Write 50k rows x 10 cols | SheetJS | 702ms | 688ms | 739ms | 739ms | 0.0MB |
| Write 100k rows x 10 cols | SheetKit | 699ms | 677ms | 727ms | 727ms | 0.0MB |
| Write 100k rows x 10 cols | ExcelJS | 3.90s | 3.76s | 3.93s | 3.93s | 0.2MB |
| Write 100k rows x 10 cols | SheetJS | 1.64s | 1.60s | 1.70s | 1.70s | 0.0MB |
| Buffer round-trip (10000 rows) | SheetKit | 221ms | 217ms | 228ms | 228ms | 0.0MB |
| Buffer round-trip (10000 rows) | ExcelJS | 703ms | 685ms | 723ms | 723ms | 0.3MB |
| Buffer round-trip (10000 rows) | SheetJS | 223ms | 218ms | 242ms | 242ms | 0.0MB |
| Streaming write (50000 rows) | SheetKit | 1.22s | 1.19s | 1.31s | 1.31s | 0.0MB |
| Streaming write (50000 rows) | ExcelJS | 757ms | 740ms | 805ms | 805ms | 0.2MB |
| Random-access read (1000 cells from 50k-row file) | SheetKit | 579ms | 564ms | 589ms | 589ms | 0.0MB |
| Random-access read (1000 cells from 50k-row file) | ExcelJS | 4.15s | 4.12s | 4.34s | 4.34s | 0.2MB |
| Random-access read (1000 cells from 50k-row file) | SheetJS | 1.82s | 1.78s | 1.93s | 1.93s | 0.0MB |
| Mixed workload write (ERP-style) | SheetKit | 29ms | 28ms | 29ms | 29ms | 0.0MB |
| Mixed workload write (ERP-style) | ExcelJS | 159ms | 148ms | 163ms | 163ms | 0.2MB |

### Memory Usage

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| Read Large Data (50k rows x 20 cols) | 0.0MB | 0.2MB | 0.0MB |
| Read Heavy Styles (5k rows, formatted) | 0.0MB | 0.2MB | 0.0MB |
| Read Multi-Sheet (10 sheets x 5k rows) | 0.0MB | 0.2MB | 0.0MB |
| Read Formulas (10k rows) | 0.0MB | 0.1MB | 0.0MB |
| Read Strings (20k rows text-heavy) | 0.0MB | 0.2MB | 0.0MB |
| Read Data Validation (5k rows, 8 rules) | 0.0MB | 16.5MB | 0.0MB |
| Read Comments (2k rows with comments) | 0.0MB | 0.2MB | 0.0MB |
| Read Merged Cells (500 regions) | 0.0MB | 0.1MB | 0.0MB |
| Read Mixed Workload (ERP document) | 0.0MB | 0.3MB | 0.0MB |
| Read Scale 1k rows | 0.0MB | 0.1MB | 0.0MB |
| Read Scale 10k rows | 0.0MB | 0.2MB | 0.0MB |
| Read Scale 100k rows | 0.0MB | 0.2MB | 0.0MB |
| Write 50000 rows x 20 cols | 0.0MB | 0.2MB | 0.0MB |
| Write 5000 styled rows | 0.0MB | 0.2MB | 0.0MB |
| Write 10 sheets x 5000 rows | 0.0MB | 0.3MB | 0.0MB |
| Write 10000 rows with formulas | 0.0MB | 0.2MB | 0.0MB |
| Write 20000 text-heavy rows | 0.0MB | 0.2MB | 0.0MB |
| Write 5000 rows + 8 validation rules | 0.0MB | 0.1MB | N/A |
| Write 2000 rows with comments | 0.0MB | 0.1MB | 0.0MB |
| Write 500 merged regions | 0.0MB | 0.1MB | 0.0MB |
| Write 1k rows x 10 cols | 0.0MB | 0.1MB | 0.0MB |
| Write 10k rows x 10 cols | 0.0MB | 0.2MB | 0.0MB |
| Write 50k rows x 10 cols | 0.0MB | 0.2MB | 0.0MB |
| Write 100k rows x 10 cols | 0.0MB | 0.2MB | 0.0MB |
| Buffer round-trip (10000 rows) | 0.0MB | 0.3MB | 0.0MB |
| Streaming write (50000 rows) | 0.0MB | 0.2MB | N/A |
| Random-access read (1000 cells from 50k-row file) | 0.0MB | 0.2MB | 0.0MB |
| Mixed workload write (ERP-style) | 0.0MB | 0.2MB | N/A |

## Summary

Total scenarios: 28

| Library | Wins |
|---------|------|
| SheetKit | 26/28 |
| ExcelJS | 1/28 |
| SheetJS | 1/28 |
