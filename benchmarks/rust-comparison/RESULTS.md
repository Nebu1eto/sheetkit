# Rust Excel Library Comparison Benchmark

Benchmark run: 2026-02-12T11:10:02Z

## Libraries

| Library | Description | Capability |
|---------|-------------|------------|
| **SheetKit** | Rust Excel library (this project) | Read + Write + Modify |
| **calamine** | Fast Excel/ODS reader | Read-only |
| **rust_xlsxwriter** | Excel writer (port of libxlsxwriter) | Write-only |
| **edit-xlsx** | Excel read/modify/write library | Read + Write + Modify |

## Environment

| Item | Value |
|------|-------|
| CPU | Apple M4 Pro |
| RAM | 24 GB |
| OS | macos aarch64 |
| Rust | rustc 1.93.0 (254b59607 2026-01-19) |
| Profile | release (opt-level=3, LTO=fat) |
| Iterations | 5 runs per scenario, 1 warmup |

## Read

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Read Large Data (50k rows x 20 cols) | 525ms | 500ms | 331ms | 39ms* | N/A | N/A | calamine |
| Read Heavy Styles (5k rows, formatted) | 27ms | 27ms | 18ms | 2ms* | N/A | N/A | calamine |
| Read Multi-Sheet (10 sheets x 5k rows) | 304ms | 305ms | 185ms | 39ms* | N/A | N/A | calamine |
| Read Formulas (10k rows) | 32ms | 31ms | 15ms | 0ms* | N/A | N/A | calamine |
| Read Strings (20k rows text-heavy) | 105ms | 104ms | 71ms | 9ms* | N/A | N/A | calamine |

## Read (Scale)

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Read Scale 1k rows | 5ms | 5ms | 3ms | 1ms* | N/A | N/A | calamine |
| Read Scale 10k rows | 50ms | 50ms | 34ms | 5ms* | N/A | N/A | calamine |
| Read Scale 100k rows | 537ms | 527ms | 355ms | 45ms* | N/A | N/A | calamine |

## Write

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Write 50000 rows x 20 cols | 527ms | N/A | N/A | 971ms | 940ms | N/A | SheetKit |
| Write 5000 styled rows | 27ms | N/A | N/A | 53ms | 40ms | N/A | SheetKit |
| Write 10 sheets x 5000 rows | 256ms | N/A | N/A | 430ms | 353ms | N/A | SheetKit |
| Write 10000 rows with formulas | 23ms | N/A | N/A | 62ms | 37ms | N/A | SheetKit |
| Write 20000 text-heavy rows | 58ms | N/A | N/A | 73ms | 69ms | N/A | SheetKit |
| Write 500 merged regions | 1ms | N/A | N/A | 5ms | 2ms | N/A | SheetKit |

## Write (Scale)

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Write 1k rows x 10 cols | 4ms | N/A | N/A | 12ms | 6ms | N/A | SheetKit |
| Write 10k rows x 10 cols | 47ms | N/A | N/A | 89ms | 77ms | N/A | SheetKit |
| Write 50k rows x 10 cols | 250ms | N/A | N/A | 437ms | 406ms | N/A | SheetKit |
| Write 100k rows x 10 cols | 507ms | N/A | N/A | 882ms | 830ms | N/A | SheetKit |

## Round-Trip

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Buffer round-trip (10000 rows) | 123ms | N/A | N/A | N/A | N/A | 90ms | xlsxwriter+calamine |

## Streaming

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Streaming write (50000 rows) | 202ms | N/A | N/A | N/A | 931ms | N/A | SheetKit |

## Random Access

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Random-access read (1000 cells) | 491ms | 481ms | 340ms | N/A | N/A | N/A | calamine |

## Modify

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Modify 1000 cells in 50k-row file | 715ms | 688ms | N/A | N/A | N/A | N/A | SheetKit (lazy) |

* `*` indicates `cells_read = 0`; excluded from Winner selection as non-comparable.

## Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Peak Mem (MB) | Cells Read |
|----------|---------|--------|-----|-----|-----|---------------|------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 525ms | 510ms | 546ms | 546ms | 60.2 | 1000020 |
| Read Large Data (50k rows x 20 cols) | SheetKit (lazy) | 500ms | 495ms | 510ms | 510ms | 17.2 | 1000020 |
| Read Large Data (50k rows x 20 cols) | calamine | 331ms | 330ms | 340ms | 340ms | 38.2 | 1000020 |
| Read Large Data (50k rows x 20 cols) | edit-xlsx | 39ms | 39ms | 41ms | 41ms | 3.4 | 0 |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 27ms | 27ms | 28ms | 28ms | 1.9 | 50010 |
| Read Heavy Styles (5k rows, formatted) | SheetKit (lazy) | 27ms | 27ms | 28ms | 28ms | 0.7 | 50010 |
| Read Heavy Styles (5k rows, formatted) | calamine | 18ms | 18ms | 18ms | 18ms | 0.0 | 50010 |
| Read Heavy Styles (5k rows, formatted) | edit-xlsx | 2ms | 2ms | 2ms | 2ms | 0.0 | 0 |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 304ms | 301ms | 308ms | 308ms | 1.2 | 500100 |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit (lazy) | 305ms | 294ms | 306ms | 306ms | 8.0 | 500100 |
| Read Multi-Sheet (10 sheets x 5k rows) | calamine | 185ms | 181ms | 187ms | 187ms | 0.0 | 500100 |
| Read Multi-Sheet (10 sheets x 5k rows) | edit-xlsx | 39ms | 38ms | 39ms | 39ms | 0.0 | 0 |
| Read Formulas (10k rows) | SheetKit | 32ms | 32ms | 33ms | 33ms | 0.0 | 70007 |
| Read Formulas (10k rows) | SheetKit (lazy) | 31ms | 31ms | 32ms | 32ms | 0.0 | 70007 |
| Read Formulas (10k rows) | calamine | 15ms | 15ms | 16ms | 16ms | 0.0 | 20007 |
| Read Formulas (10k rows) | edit-xlsx | 0ms | 0ms | 0ms | 0ms | 0.0 | 0 |
| Read Strings (20k rows text-heavy) | SheetKit | 105ms | 104ms | 106ms | 106ms | 0.0 | 200010 |
| Read Strings (20k rows text-heavy) | SheetKit (lazy) | 104ms | 104ms | 105ms | 105ms | 0.0 | 200010 |
| Read Strings (20k rows text-heavy) | calamine | 71ms | 70ms | 72ms | 72ms | 7.7 | 200010 |
| Read Strings (20k rows text-heavy) | edit-xlsx | 9ms | 9ms | 9ms | 9ms | 0.0 | 0 |
| Read Scale 1k rows | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0 | 10010 |
| Read Scale 1k rows | SheetKit (lazy) | 5ms | 5ms | 5ms | 5ms | 0.0 | 10010 |
| Read Scale 1k rows | calamine | 3ms | 3ms | 3ms | 3ms | 0.0 | 10010 |
| Read Scale 1k rows | edit-xlsx | 1ms | 1ms | 1ms | 1ms | 0.0 | 0 |
| Read Scale 10k rows | SheetKit | 50ms | 50ms | 51ms | 51ms | 0.0 | 100010 |
| Read Scale 10k rows | SheetKit (lazy) | 50ms | 50ms | 51ms | 51ms | 0.0 | 100010 |
| Read Scale 10k rows | calamine | 34ms | 33ms | 35ms | 35ms | 0.0 | 100010 |
| Read Scale 10k rows | edit-xlsx | 5ms | 4ms | 5ms | 5ms | 0.0 | 0 |
| Read Scale 100k rows | SheetKit | 537ms | 531ms | 543ms | 543ms | 41.4 | 1000010 |
| Read Scale 100k rows | SheetKit (lazy) | 527ms | 514ms | 554ms | 554ms | 29.0 | 1000010 |
| Read Scale 100k rows | calamine | 355ms | 344ms | 365ms | 365ms | 0.0 | 1000010 |
| Read Scale 100k rows | edit-xlsx | 45ms | 44ms | 46ms | 46ms | 0.6 | 0 |
| Write 50000 rows x 20 cols | SheetKit | 527ms | 517ms | 542ms | 542ms | 40.0 | N/A |
| Write 50000 rows x 20 cols | rust_xlsxwriter | 940ms | 932ms | 949ms | 949ms | 0.0 | N/A |
| Write 50000 rows x 20 cols | edit-xlsx | 971ms | 967ms | 1.00s | 1.00s | 112.0 | N/A |
| Write 5000 styled rows | SheetKit | 27ms | 27ms | 28ms | 28ms | 0.0 | N/A |
| Write 5000 styled rows | rust_xlsxwriter | 40ms | 39ms | 40ms | 40ms | 0.0 | N/A |
| Write 5000 styled rows | edit-xlsx | 53ms | 53ms | 54ms | 54ms | 0.0 | N/A |
| Write 10 sheets x 5000 rows | SheetKit | 256ms | 255ms | 263ms | 263ms | 12.0 | N/A |
| Write 10 sheets x 5000 rows | rust_xlsxwriter | 353ms | 352ms | 359ms | 359ms | 9.9 | N/A |
| Write 10 sheets x 5000 rows | edit-xlsx | 430ms | 427ms | 437ms | 437ms | 0.0 | N/A |
| Write 10000 rows with formulas | SheetKit | 23ms | 23ms | 23ms | 23ms | 0.0 | N/A |
| Write 10000 rows with formulas | rust_xlsxwriter | 37ms | 37ms | 37ms | 37ms | 0.0 | N/A |
| Write 10000 rows with formulas | edit-xlsx | 62ms | 61ms | 66ms | 66ms | 0.0 | N/A |
| Write 20000 text-heavy rows | SheetKit | 58ms | 58ms | 61ms | 61ms | 0.0 | N/A |
| Write 20000 text-heavy rows | rust_xlsxwriter | 69ms | 68ms | 69ms | 69ms | 0.0 | N/A |
| Write 20000 text-heavy rows | edit-xlsx | 73ms | 72ms | 74ms | 74ms | 0.0 | N/A |
| Write 500 merged regions | SheetKit | 1ms | 1ms | 1ms | 1ms | 0.0 | N/A |
| Write 500 merged regions | rust_xlsxwriter | 2ms | 2ms | 2ms | 2ms | 0.0 | N/A |
| Write 500 merged regions | edit-xlsx | 5ms | 5ms | 5ms | 5ms | 0.0 | N/A |
| Write 1k rows x 10 cols | SheetKit | 4ms | 4ms | 5ms | 5ms | 0.0 | N/A |
| Write 1k rows x 10 cols | rust_xlsxwriter | 6ms | 6ms | 6ms | 6ms | 0.0 | N/A |
| Write 1k rows x 10 cols | edit-xlsx | 12ms | 11ms | 12ms | 12ms | 0.0 | N/A |
| Write 10k rows x 10 cols | SheetKit | 47ms | 47ms | 47ms | 47ms | 0.0 | N/A |
| Write 10k rows x 10 cols | rust_xlsxwriter | 77ms | 76ms | 78ms | 78ms | 0.0 | N/A |
| Write 10k rows x 10 cols | edit-xlsx | 89ms | 88ms | 90ms | 90ms | 0.0 | N/A |
| Write 50k rows x 10 cols | SheetKit | 250ms | 246ms | 251ms | 251ms | 0.0 | N/A |
| Write 50k rows x 10 cols | rust_xlsxwriter | 406ms | 403ms | 433ms | 433ms | 0.0 | N/A |
| Write 50k rows x 10 cols | edit-xlsx | 437ms | 435ms | 449ms | 449ms | 0.0 | N/A |
| Write 100k rows x 10 cols | SheetKit | 507ms | 500ms | 510ms | 510ms | 0.0 | N/A |
| Write 100k rows x 10 cols | rust_xlsxwriter | 830ms | 829ms | 850ms | 850ms | 0.0 | N/A |
| Write 100k rows x 10 cols | edit-xlsx | 882ms | 861ms | 914ms | 914ms | 0.1 | N/A |
| Buffer round-trip (10000 rows) | SheetKit | 123ms | 122ms | 128ms | 128ms | 0.1 | N/A |
| Buffer round-trip (10000 rows) | xlsxwriter+calamine | 90ms | 86ms | 91ms | 91ms | 0.0 | N/A |
| Streaming write (50000 rows) | SheetKit | 202ms | 196ms | 204ms | 204ms | 0.0 | N/A |
| Streaming write (50000 rows) | rust_xlsxwriter | 931ms | 920ms | 933ms | 933ms | 41.0 | N/A |
| Random-access read (1000 cells) | SheetKit | 491ms | 484ms | 494ms | 494ms | 21.4 | 1000 |
| Random-access read (1000 cells) | SheetKit (lazy) | 481ms | 473ms | 491ms | 491ms | 17.2 | 1000 |
| Random-access read (1000 cells) | calamine | 340ms | 338ms | 395ms | 395ms | 38.2 | 1000 |
| Modify 1000 cells in 50k-row file | SheetKit | 715ms | 704ms | 722ms | 722ms | 26.8 | N/A |
| Modify 1000 cells in 50k-row file | SheetKit (lazy) | 688ms | 684ms | 696ms | 696ms | 2.0 | N/A |

## Win Summary

| Library | Wins |
|---------|------|
| SheetKit | 11/22 |
| calamine | 9/22 |
| SheetKit (lazy) | 1/22 |
| xlsxwriter+calamine | 1/22 |
