# Rust Excel Library Comparison Benchmark

Benchmark run: 2026-02-12T13:20:18Z

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
| Read Large Data (50k rows x 20 cols) | 486ms | 489ms | 323ms | 39ms* | N/A | N/A | calamine |
| Read Heavy Styles (5k rows, formatted) | 26ms | 26ms | 17ms | 2ms* | N/A | N/A | calamine |
| Read Multi-Sheet (10 sheets x 5k rows) | 293ms | 294ms | 181ms | 38ms* | N/A | N/A | calamine |
| Read Formulas (10k rows) | 32ms | 32ms | 15ms | 0ms* | N/A | N/A | calamine |
| Read Strings (20k rows text-heavy) | 105ms | 105ms | 70ms | 10ms* | N/A | N/A | calamine |

## Read (Scale)

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Read Scale 1k rows | 5ms | 5ms | 3ms | 1ms* | N/A | N/A | calamine |
| Read Scale 10k rows | 50ms | 50ms | 33ms | 4ms* | N/A | N/A | calamine |
| Read Scale 100k rows | 507ms | 510ms | 327ms | 43ms* | N/A | N/A | calamine |

## Write

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Write 50000 rows x 20 cols | 497ms | N/A | N/A | 946ms | 917ms | N/A | SheetKit |
| Write 5000 styled rows | 27ms | N/A | N/A | 54ms | 40ms | N/A | SheetKit |
| Write 10 sheets x 5000 rows | 256ms | N/A | N/A | 424ms | 350ms | N/A | SheetKit |
| Write 10000 rows with formulas | 23ms | N/A | N/A | 60ms | 37ms | N/A | SheetKit |
| Write 20000 text-heavy rows | 58ms | N/A | N/A | 72ms | 69ms | N/A | SheetKit |
| Write 500 merged regions | 1ms | N/A | N/A | 5ms | 2ms | N/A | SheetKit |

## Write (Scale)

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Write 1k rows x 10 cols | 5ms | N/A | N/A | 12ms | 6ms | N/A | SheetKit |
| Write 10k rows x 10 cols | 46ms | N/A | N/A | 89ms | 77ms | N/A | SheetKit |
| Write 50k rows x 10 cols | 244ms | N/A | N/A | 432ms | 401ms | N/A | SheetKit |
| Write 100k rows x 10 cols | 492ms | N/A | N/A | 883ms | 809ms | N/A | SheetKit |

## Round-Trip

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Buffer round-trip (10000 rows) | 120ms | N/A | N/A | N/A | N/A | 85ms | xlsxwriter+calamine |

## Streaming

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Streaming write (50000 rows) | 200ms | N/A | N/A | N/A | 922ms | N/A | SheetKit |

## Random Access

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Random-access read (1000 cells) | 476ms | 469ms | 326ms | N/A | N/A | N/A | calamine |

## Modify

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Modify 1000 cells in 50k-row file | 701ms | 688ms | N/A | N/A | N/A | N/A | SheetKit (lazy) |

* `*` indicates `cells_read = 0`; excluded from Winner selection as non-comparable.

## Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Peak Mem (MB) | Cells Read |
|----------|---------|--------|-----|-----|-----|---------------|------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 486ms | 483ms | 491ms | 491ms | 17.4 | 1000020 |
| Read Large Data (50k rows x 20 cols) | SheetKit (lazy) | 489ms | 487ms | 492ms | 492ms | 0.0 | 1000020 |
| Read Large Data (50k rows x 20 cols) | calamine | 323ms | 321ms | 326ms | 326ms | 38.2 | 1000020 |
| Read Large Data (50k rows x 20 cols) | edit-xlsx | 39ms | 39ms | 40ms | 40ms | 5.3 | 0 |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 26ms | 26ms | 26ms | 26ms | 0.0 | 50010 |
| Read Heavy Styles (5k rows, formatted) | SheetKit (lazy) | 26ms | 26ms | 26ms | 26ms | 0.0 | 50010 |
| Read Heavy Styles (5k rows, formatted) | calamine | 17ms | 17ms | 17ms | 17ms | 0.0 | 50010 |
| Read Heavy Styles (5k rows, formatted) | edit-xlsx | 2ms | 2ms | 2ms | 2ms | 0.0 | 0 |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 293ms | 291ms | 296ms | 296ms | 0.9 | 500100 |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit (lazy) | 294ms | 292ms | 296ms | 296ms | 0.0 | 500100 |
| Read Multi-Sheet (10 sheets x 5k rows) | calamine | 181ms | 181ms | 183ms | 183ms | 0.0 | 500100 |
| Read Multi-Sheet (10 sheets x 5k rows) | edit-xlsx | 38ms | 38ms | 39ms | 39ms | 4.0 | 0 |
| Read Formulas (10k rows) | SheetKit | 32ms | 31ms | 32ms | 32ms | 0.0 | 70007 |
| Read Formulas (10k rows) | SheetKit (lazy) | 32ms | 32ms | 32ms | 32ms | 0.0 | 70007 |
| Read Formulas (10k rows) | calamine | 15ms | 15ms | 15ms | 15ms | 0.0 | 20007 |
| Read Formulas (10k rows) | edit-xlsx | 0ms | 0ms | 0ms | 0ms | 0.0 | 0 |
| Read Strings (20k rows text-heavy) | SheetKit | 105ms | 103ms | 105ms | 105ms | 3.5 | 200010 |
| Read Strings (20k rows text-heavy) | SheetKit (lazy) | 105ms | 105ms | 106ms | 106ms | 3.5 | 200010 |
| Read Strings (20k rows text-heavy) | calamine | 70ms | 69ms | 71ms | 71ms | 4.2 | 200010 |
| Read Strings (20k rows text-heavy) | edit-xlsx | 10ms | 10ms | 10ms | 10ms | 0.0 | 0 |
| Read Scale 1k rows | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0 | 10010 |
| Read Scale 1k rows | SheetKit (lazy) | 5ms | 5ms | 5ms | 5ms | 0.0 | 10010 |
| Read Scale 1k rows | calamine | 3ms | 3ms | 4ms | 4ms | 0.0 | 10010 |
| Read Scale 1k rows | edit-xlsx | 1ms | 1ms | 1ms | 1ms | 0.0 | 0 |
| Read Scale 10k rows | SheetKit | 50ms | 50ms | 51ms | 51ms | 0.0 | 100010 |
| Read Scale 10k rows | SheetKit (lazy) | 50ms | 50ms | 50ms | 50ms | 0.0 | 100010 |
| Read Scale 10k rows | calamine | 33ms | 32ms | 33ms | 33ms | 0.4 | 100010 |
| Read Scale 10k rows | edit-xlsx | 4ms | 4ms | 4ms | 4ms | 0.0 | 0 |
| Read Scale 100k rows | SheetKit | 507ms | 507ms | 515ms | 515ms | 10.1 | 1000010 |
| Read Scale 100k rows | SheetKit (lazy) | 510ms | 507ms | 530ms | 530ms | 8.4 | 1000010 |
| Read Scale 100k rows | calamine | 327ms | 325ms | 328ms | 328ms | 38.7 | 1000010 |
| Read Scale 100k rows | edit-xlsx | 43ms | 43ms | 44ms | 44ms | 15.6 | 0 |
| Write 50000 rows x 20 cols | SheetKit | 497ms | 493ms | 500ms | 500ms | 17.5 | N/A |
| Write 50000 rows x 20 cols | rust_xlsxwriter | 917ms | 911ms | 925ms | 925ms | 41.1 | N/A |
| Write 50000 rows x 20 cols | edit-xlsx | 946ms | 944ms | 954ms | 954ms | 117.5 | N/A |
| Write 5000 styled rows | SheetKit | 27ms | 27ms | 27ms | 27ms | 0.0 | N/A |
| Write 5000 styled rows | rust_xlsxwriter | 40ms | 39ms | 41ms | 41ms | 0.0 | N/A |
| Write 5000 styled rows | edit-xlsx | 54ms | 53ms | 54ms | 54ms | 2.8 | N/A |
| Write 10 sheets x 5000 rows | SheetKit | 256ms | 254ms | 258ms | 258ms | 15.4 | N/A |
| Write 10 sheets x 5000 rows | rust_xlsxwriter | 350ms | 349ms | 352ms | 352ms | 3.6 | N/A |
| Write 10 sheets x 5000 rows | edit-xlsx | 424ms | 424ms | 425ms | 425ms | 0.0 | N/A |
| Write 10000 rows with formulas | SheetKit | 23ms | 23ms | 23ms | 23ms | 0.0 | N/A |
| Write 10000 rows with formulas | rust_xlsxwriter | 37ms | 37ms | 37ms | 37ms | 0.0 | N/A |
| Write 10000 rows with formulas | edit-xlsx | 60ms | 60ms | 61ms | 61ms | 6.1 | N/A |
| Write 20000 text-heavy rows | SheetKit | 58ms | 57ms | 58ms | 58ms | 1.7 | N/A |
| Write 20000 text-heavy rows | rust_xlsxwriter | 69ms | 68ms | 69ms | 69ms | 0.0 | N/A |
| Write 20000 text-heavy rows | edit-xlsx | 72ms | 71ms | 72ms | 72ms | 1.0 | N/A |
| Write 500 merged regions | SheetKit | 1ms | 1ms | 1ms | 1ms | 0.0 | N/A |
| Write 500 merged regions | rust_xlsxwriter | 2ms | 2ms | 2ms | 2ms | 0.0 | N/A |
| Write 500 merged regions | edit-xlsx | 5ms | 5ms | 5ms | 5ms | 0.0 | N/A |
| Write 1k rows x 10 cols | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0 | N/A |
| Write 1k rows x 10 cols | rust_xlsxwriter | 6ms | 6ms | 6ms | 6ms | 0.0 | N/A |
| Write 1k rows x 10 cols | edit-xlsx | 12ms | 11ms | 12ms | 12ms | 0.0 | N/A |
| Write 10k rows x 10 cols | SheetKit | 46ms | 45ms | 46ms | 46ms | 0.0 | N/A |
| Write 10k rows x 10 cols | rust_xlsxwriter | 77ms | 76ms | 82ms | 82ms | 0.0 | N/A |
| Write 10k rows x 10 cols | edit-xlsx | 89ms | 88ms | 90ms | 90ms | 0.1 | N/A |
| Write 50k rows x 10 cols | SheetKit | 244ms | 243ms | 254ms | 254ms | 6.9 | N/A |
| Write 50k rows x 10 cols | rust_xlsxwriter | 401ms | 398ms | 401ms | 401ms | 4.6 | N/A |
| Write 50k rows x 10 cols | edit-xlsx | 432ms | 429ms | 433ms | 433ms | 2.5 | N/A |
| Write 100k rows x 10 cols | SheetKit | 492ms | 487ms | 494ms | 494ms | 8.0 | N/A |
| Write 100k rows x 10 cols | rust_xlsxwriter | 809ms | 806ms | 824ms | 824ms | 9.6 | N/A |
| Write 100k rows x 10 cols | edit-xlsx | 883ms | 879ms | 896ms | 896ms | 9.9 | N/A |
| Buffer round-trip (10000 rows) | SheetKit | 120ms | 119ms | 121ms | 121ms | 0.0 | N/A |
| Buffer round-trip (10000 rows) | xlsxwriter+calamine | 85ms | 85ms | 86ms | 86ms | 0.0 | N/A |
| Streaming write (50000 rows) | SheetKit | 200ms | 198ms | 218ms | 218ms | 0.0 | N/A |
| Streaming write (50000 rows) | rust_xlsxwriter | 922ms | 910ms | 929ms | 929ms | 20.2 | N/A |
| Random-access read (1000 cells) | SheetKit | 476ms | 467ms | 477ms | 477ms | 14.1 | 1000 |
| Random-access read (1000 cells) | SheetKit (lazy) | 469ms | 464ms | 478ms | 478ms | 17.2 | 1000 |
| Random-access read (1000 cells) | calamine | 326ms | 324ms | 329ms | 329ms | 10.7 | 1000 |
| Modify 1000 cells in 50k-row file | SheetKit | 701ms | 689ms | 715ms | 715ms | 39.4 | N/A |
| Modify 1000 cells in 50k-row file | SheetKit (lazy) | 688ms | 682ms | 689ms | 689ms | 7.7 | N/A |

## Win Summary

| Library | Wins |
|---------|------|
| SheetKit | 11/22 |
| calamine | 9/22 |
| SheetKit (lazy) | 1/22 |
| xlsxwriter+calamine | 1/22 |
