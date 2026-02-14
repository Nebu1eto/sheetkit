# Rust Excel Library Comparison Benchmark

Benchmark run: 2026-02-14T06:21:07Z

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
| Read Large Data (50k rows x 20 cols) | 593ms | 607ms | 358ms | 40ms* | N/A | N/A | calamine |
| Read Heavy Styles (5k rows, formatted) | 32ms | 33ms | 19ms | 2ms* | N/A | N/A | calamine |
| Read Multi-Sheet (10 sheets x 5k rows) | 344ms | 353ms | 197ms | 38ms* | N/A | N/A | calamine |
| Read Formulas (10k rows) | 42ms* | 36ms* | 16ms* | 0ms* | N/A | N/A | N/A |
| Read Strings (20k rows text-heavy) | 132ms | 132ms | 76ms | 10ms* | N/A | N/A | calamine |

## Read (Scale)

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Read Scale 1k rows | 6ms | 6ms | 4ms | 1ms* | N/A | N/A | calamine |
| Read Scale 10k rows | 62ms | 62ms | 37ms | 5ms* | N/A | N/A | calamine |
| Read Scale 100k rows | 625ms | 635ms | 369ms | 44ms* | N/A | N/A | calamine |

## Write

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Write 50000 rows x 20 cols | 594ms | N/A | N/A | 1.01s | 962ms | N/A | SheetKit |
| Write 5000 styled rows | 32ms | N/A | N/A | 58ms | 40ms | N/A | SheetKit |
| Write 10 sheets x 5000 rows | 286ms | N/A | N/A | 446ms | 349ms | N/A | SheetKit |
| Write 10000 rows with formulas | 25ms | N/A | N/A | 65ms | 37ms | N/A | SheetKit |
| Write 20000 text-heavy rows | 66ms | N/A | N/A | 82ms | 69ms | N/A | SheetKit |
| Write 500 merged regions | 1ms | N/A | N/A | 5ms | 2ms | N/A | SheetKit |

## Write (Scale)

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Write 1k rows x 10 cols | 5ms | N/A | N/A | 12ms | 7ms | N/A | SheetKit |
| Write 10k rows x 10 cols | 53ms | N/A | N/A | 91ms | 76ms | N/A | SheetKit |
| Write 50k rows x 10 cols | 265ms | N/A | N/A | 445ms | 397ms | N/A | SheetKit |
| Write 100k rows x 10 cols | 552ms | N/A | N/A | 916ms | 802ms | N/A | SheetKit |

## Round-Trip

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Buffer round-trip (10000 rows) | 130ms | N/A | N/A | N/A | N/A | 89ms | xlsxwriter+calamine |

## Streaming

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Streaming write (50000 rows) | 224ms | N/A | N/A | N/A | 961ms | N/A | SheetKit |

## Random Access

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Random-access read (1000 cells) | 538ms | 537ms | 365ms | N/A | N/A | N/A | calamine |

## Modify

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Modify 1000 cells in 50k-row file | 771ms | 777ms | N/A | N/A | N/A | N/A | SheetKit |

* `*` indicates workload-count mismatch, value-probe mismatch, or zero read counts; excluded from Winner selection as non-comparable.

## Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Peak Mem (MB) | Rows Read | Rows Expected | Cells Read | Cells Expected | Comparable |
|----------|---------|--------|-----|-----|-----|---------------|----------|---------------|------------|----------------|------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 593ms | 586ms | 597ms | 597ms | 0.0 | 50001 | 50001 | 1000020 | 1000020 | Yes |
| Read Large Data (50k rows x 20 cols) | SheetKit (lazy) | 607ms | 605ms | 609ms | 609ms | 0.0 | 50001 | 50001 | 1000020 | 1000020 | Yes |
| Read Large Data (50k rows x 20 cols) | calamine | 358ms | 356ms | 361ms | 361ms | 0.0 | 50001 | 50001 | 1000020 | 1000020 | Yes |
| Read Large Data (50k rows x 20 cols) | edit-xlsx | 40ms | 39ms | 41ms | 41ms | 8.5 | 0 | 50001 | 0 | 1000020 | No |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 32ms | 32ms | 32ms | 32ms | 0.0 | 5001 | 5001 | 50010 | 50010 | Yes |
| Read Heavy Styles (5k rows, formatted) | SheetKit (lazy) | 33ms | 33ms | 33ms | 33ms | 0.0 | 5001 | 5001 | 50010 | 50010 | Yes |
| Read Heavy Styles (5k rows, formatted) | calamine | 19ms | 19ms | 19ms | 19ms | 0.0 | 5001 | 5001 | 50010 | 50010 | Yes |
| Read Heavy Styles (5k rows, formatted) | edit-xlsx | 2ms | 2ms | 2ms | 2ms | 0.0 | 0 | 5001 | 0 | 50010 | No |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 344ms | 344ms | 345ms | 345ms | 1.3 | 50010 | 50010 | 500100 | 500100 | Yes |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit (lazy) | 353ms | 350ms | 356ms | 356ms | 17.7 | 50010 | 50010 | 500100 | 500100 | Yes |
| Read Multi-Sheet (10 sheets x 5k rows) | calamine | 197ms | 196ms | 198ms | 198ms | 0.0 | 50010 | 50010 | 500100 | 500100 | Yes |
| Read Multi-Sheet (10 sheets x 5k rows) | edit-xlsx | 38ms | 38ms | 39ms | 39ms | 4.0 | 0 | 50010 | 0 | 500100 | No |
| Read Formulas (10k rows) | SheetKit | 42ms | 42ms | 42ms | 42ms | 0.0 | 10001 | 10001 | 70007 | N/A | No |
| Read Formulas (10k rows) | SheetKit (lazy) | 36ms | 36ms | 36ms | 36ms | 0.0 | 10001 | 10001 | 70007 | N/A | No |
| Read Formulas (10k rows) | calamine | 16ms | 16ms | 16ms | 16ms | 0.0 | 10001 | 10001 | 20007 | N/A | No |
| Read Formulas (10k rows) | edit-xlsx | 0ms | 0ms | 0ms | 0ms | 0.0 | 0 | 10001 | 0 | N/A | No |
| Read Strings (20k rows text-heavy) | SheetKit | 132ms | 131ms | 134ms | 134ms | 3.5 | 20001 | 20001 | 200010 | 200010 | Yes |
| Read Strings (20k rows text-heavy) | SheetKit (lazy) | 132ms | 131ms | 134ms | 134ms | 3.5 | 20001 | 20001 | 200010 | 200010 | Yes |
| Read Strings (20k rows text-heavy) | calamine | 76ms | 76ms | 80ms | 80ms | 4.2 | 20001 | 20001 | 200010 | 200010 | Yes |
| Read Strings (20k rows text-heavy) | edit-xlsx | 10ms | 10ms | 10ms | 10ms | 0.0 | 0 | 20001 | 0 | 200010 | No |
| Read Scale 1k rows | SheetKit | 6ms | 6ms | 6ms | 6ms | 0.0 | 1001 | 1001 | 10010 | 10010 | Yes |
| Read Scale 1k rows | SheetKit (lazy) | 6ms | 6ms | 6ms | 6ms | 0.0 | 1001 | 1001 | 10010 | 10010 | Yes |
| Read Scale 1k rows | calamine | 4ms | 3ms | 4ms | 4ms | 0.0 | 1001 | 1001 | 10010 | 10010 | Yes |
| Read Scale 1k rows | edit-xlsx | 1ms | 1ms | 1ms | 1ms | 0.0 | 0 | 1001 | 0 | 10010 | No |
| Read Scale 10k rows | SheetKit | 62ms | 61ms | 63ms | 63ms | 0.0 | 10001 | 10001 | 100010 | 100010 | Yes |
| Read Scale 10k rows | SheetKit (lazy) | 62ms | 61ms | 63ms | 63ms | 0.0 | 10001 | 10001 | 100010 | 100010 | Yes |
| Read Scale 10k rows | calamine | 37ms | 36ms | 41ms | 41ms | 3.8 | 10001 | 10001 | 100010 | 100010 | Yes |
| Read Scale 10k rows | edit-xlsx | 5ms | 4ms | 5ms | 5ms | 0.0 | 0 | 10001 | 0 | 100010 | No |
| Read Scale 100k rows | SheetKit | 625ms | 623ms | 635ms | 635ms | 8.9 | 100001 | 100001 | 1000010 | 1000010 | Yes |
| Read Scale 100k rows | SheetKit (lazy) | 635ms | 627ms | 640ms | 640ms | 8.4 | 100001 | 100001 | 1000010 | 1000010 | Yes |
| Read Scale 100k rows | calamine | 369ms | 367ms | 374ms | 374ms | 38.7 | 100001 | 100001 | 1000010 | 1000010 | Yes |
| Read Scale 100k rows | edit-xlsx | 44ms | 43ms | 44ms | 44ms | 15.1 | 0 | 100001 | 0 | 1000010 | No |
| Write 50000 rows x 20 cols | SheetKit | 594ms | 591ms | 605ms | 605ms | 57.5 | N/A | N/A | N/A | N/A | Yes |
| Write 50000 rows x 20 cols | rust_xlsxwriter | 962ms | 954ms | 973ms | 973ms | 3.7 | N/A | N/A | N/A | N/A | Yes |
| Write 50000 rows x 20 cols | edit-xlsx | 1.01s | 992ms | 1.01s | 1.01s | 2.5 | N/A | N/A | N/A | N/A | Yes |
| Write 5000 styled rows | SheetKit | 32ms | 31ms | 33ms | 33ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 5000 styled rows | rust_xlsxwriter | 40ms | 39ms | 41ms | 41ms | 0.2 | N/A | N/A | N/A | N/A | Yes |
| Write 5000 styled rows | edit-xlsx | 58ms | 56ms | 59ms | 59ms | 0.3 | N/A | N/A | N/A | N/A | Yes |
| Write 10 sheets x 5000 rows | SheetKit | 286ms | 285ms | 292ms | 292ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10 sheets x 5000 rows | rust_xlsxwriter | 349ms | 347ms | 354ms | 354ms | 8.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10 sheets x 5000 rows | edit-xlsx | 446ms | 439ms | 473ms | 473ms | 5.7 | N/A | N/A | N/A | N/A | Yes |
| Write 10000 rows with formulas | SheetKit | 25ms | 25ms | 26ms | 26ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10000 rows with formulas | rust_xlsxwriter | 37ms | 36ms | 37ms | 37ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10000 rows with formulas | edit-xlsx | 65ms | 65ms | 65ms | 65ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 20000 text-heavy rows | SheetKit | 66ms | 65ms | 66ms | 66ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 20000 text-heavy rows | rust_xlsxwriter | 69ms | 69ms | 70ms | 70ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 20000 text-heavy rows | edit-xlsx | 82ms | 77ms | 83ms | 83ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 500 merged regions | SheetKit | 1ms | 1ms | 1ms | 1ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 500 merged regions | rust_xlsxwriter | 2ms | 2ms | 2ms | 2ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 500 merged regions | edit-xlsx | 5ms | 5ms | 6ms | 6ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 1k rows x 10 cols | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 1k rows x 10 cols | rust_xlsxwriter | 7ms | 6ms | 7ms | 7ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 1k rows x 10 cols | edit-xlsx | 12ms | 12ms | 12ms | 12ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10k rows x 10 cols | SheetKit | 53ms | 53ms | 56ms | 56ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10k rows x 10 cols | rust_xlsxwriter | 76ms | 75ms | 77ms | 77ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10k rows x 10 cols | edit-xlsx | 91ms | 90ms | 117ms | 117ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 50k rows x 10 cols | SheetKit | 265ms | 264ms | 284ms | 284ms | 32.6 | N/A | N/A | N/A | N/A | Yes |
| Write 50k rows x 10 cols | rust_xlsxwriter | 397ms | 396ms | 398ms | 398ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 50k rows x 10 cols | edit-xlsx | 445ms | 443ms | 447ms | 447ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 100k rows x 10 cols | SheetKit | 552ms | 547ms | 557ms | 557ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 100k rows x 10 cols | rust_xlsxwriter | 802ms | 800ms | 804ms | 804ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 100k rows x 10 cols | edit-xlsx | 916ms | 907ms | 976ms | 976ms | 9.1 | N/A | N/A | N/A | N/A | Yes |
| Buffer round-trip (10000 rows) | SheetKit | 130ms | 130ms | 183ms | 183ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Buffer round-trip (10000 rows) | xlsxwriter+calamine | 89ms | 88ms | 90ms | 90ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Streaming write (50000 rows) | SheetKit | 224ms | 221ms | 235ms | 235ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Streaming write (50000 rows) | rust_xlsxwriter | 961ms | 956ms | 969ms | 969ms | 32.9 | N/A | N/A | N/A | N/A | Yes |
| Random-access read (1000 cells) | SheetKit | 538ms | 532ms | 542ms | 542ms | 23.5 | N/A | N/A | 1000 | N/A | Yes |
| Random-access read (1000 cells) | SheetKit (lazy) | 537ms | 535ms | 538ms | 538ms | 21.4 | N/A | N/A | 1000 | N/A | Yes |
| Random-access read (1000 cells) | calamine | 365ms | 356ms | 372ms | 372ms | 38.2 | N/A | N/A | 1000 | N/A | Yes |
| Modify 1000 cells in 50k-row file | SheetKit | 771ms | 769ms | 774ms | 774ms | 0.7 | N/A | N/A | N/A | N/A | Yes |
| Modify 1000 cells in 50k-row file | SheetKit (lazy) | 777ms | 771ms | 785ms | 785ms | 0.0 | N/A | N/A | N/A | N/A | Yes |

## Win Summary

| Library | Wins |
|---------|------|
| SheetKit | 12/21 |
| calamine | 8/21 |
| xlsxwriter+calamine | 1/21 |
