# Rust Excel Library Comparison Benchmark

Benchmark run: 2026-02-14T07:24:31Z

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

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Read Large Data (50k rows x 20 cols) | 494ms | 324ms | 372ms* | N/A | N/A | calamine |
| Read Heavy Styles (5k rows, formatted) | 26ms | 17ms | 21ms* | N/A | N/A | calamine |
| Read Multi-Sheet (10 sheets x 5k rows) | 289ms | 179ms | 199ms* | N/A | N/A | calamine |
| Read Formulas (10k rows) | 32ms* | 15ms* | 23ms* | N/A | N/A | N/A |
| Read Strings (20k rows text-heavy) | 105ms | 70ms | 81ms* | N/A | N/A | calamine |

## Read (Scale)

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Read Scale 1k rows | 5ms | 3ms | 4ms* | N/A | N/A | calamine |
| Read Scale 10k rows | 51ms | 33ms | 36ms* | N/A | N/A | calamine |
| Read Scale 100k rows | 511ms | 325ms | 373ms* | N/A | N/A | calamine |

## Write

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Write 50000 rows x 20 cols | 475ms | N/A | 939ms | 886ms | N/A | SheetKit |
| Write 5000 styled rows | 27ms | N/A | 52ms | 38ms | N/A | SheetKit |
| Write 10 sheets x 5000 rows | 249ms | N/A | 414ms | 338ms | N/A | SheetKit |
| Write 10000 rows with formulas | 23ms | N/A | 59ms | 36ms | N/A | SheetKit |
| Write 20000 text-heavy rows | 57ms | N/A | 71ms | 66ms | N/A | SheetKit |
| Write 500 merged regions | 1ms | N/A | 5ms | 2ms | N/A | SheetKit |

## Write (Scale)

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Write 1k rows x 10 cols | 5ms | N/A | 11ms | 6ms | N/A | SheetKit |
| Write 10k rows x 10 cols | 46ms | N/A | 86ms | 74ms | N/A | SheetKit |
| Write 50k rows x 10 cols | 238ms | N/A | 423ms | 388ms | N/A | SheetKit |
| Write 100k rows x 10 cols | 493ms | N/A | 837ms | 780ms | N/A | SheetKit |

## Round-Trip

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Buffer round-trip (10000 rows) | 118ms | N/A | N/A | N/A | 82ms | xlsxwriter+calamine |

## Streaming

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Streaming write (50000 rows) | 191ms | N/A | N/A | 885ms | N/A | SheetKit |

## Random Access

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Random-access read (1000 cells) | 465ms | 321ms | N/A | N/A | N/A | calamine |

## Modify

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Modify 1000 cells in 50k-row file | 668ms | N/A | 537ms | N/A | N/A | edit-xlsx |

* `*` indicates workload-count mismatch, value-probe mismatch, or zero read counts; excluded from Winner selection as non-comparable.

## Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Peak Mem (MB) | Rows Read | Rows Expected | Cells Read | Cells Expected | Comparable |
|----------|---------|--------|-----|-----|-----|---------------|----------|---------------|------------|----------------|------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 494ms | 490ms | 501ms | 501ms | 4.3 | 50001 | 50001 | 1000020 | 1000020 | Yes |
| Read Large Data (50k rows x 20 cols) | calamine | 324ms | 321ms | 327ms | 327ms | 38.2 | 50001 | 50001 | 1000020 | 1000020 | Yes |
| Read Large Data (50k rows x 20 cols) | edit-xlsx | 372ms | 372ms | 373ms | 373ms | 0.0 | 0 | 50001 | 0 | 1000020 | No |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 26ms | 26ms | 26ms | 26ms | 0.0 | 5001 | 5001 | 50010 | 50010 | Yes |
| Read Heavy Styles (5k rows, formatted) | calamine | 17ms | 17ms | 17ms | 17ms | 0.0 | 5001 | 5001 | 50010 | 50010 | Yes |
| Read Heavy Styles (5k rows, formatted) | edit-xlsx | 21ms | 21ms | 21ms | 21ms | 0.0 | 0 | 5001 | 0 | 50010 | No |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 289ms | 287ms | 300ms | 300ms | 0.8 | 50010 | 50010 | 500100 | 500100 | Yes |
| Read Multi-Sheet (10 sheets x 5k rows) | calamine | 179ms | 178ms | 179ms | 179ms | 0.0 | 50010 | 50010 | 500100 | 500100 | Yes |
| Read Multi-Sheet (10 sheets x 5k rows) | edit-xlsx | 199ms | 199ms | 200ms | 200ms | 4.0 | 0 | 50010 | 0 | 500100 | No |
| Read Formulas (10k rows) | SheetKit | 32ms | 31ms | 32ms | 32ms | 0.0 | 10001 | 10001 | 70007 | N/A | No |
| Read Formulas (10k rows) | calamine | 15ms | 15ms | 15ms | 15ms | 0.0 | 10001 | 10001 | 20007 | N/A | No |
| Read Formulas (10k rows) | edit-xlsx | 23ms | 23ms | 24ms | 24ms | 0.0 | 0 | 10001 | 0 | N/A | No |
| Read Strings (20k rows text-heavy) | SheetKit | 105ms | 105ms | 105ms | 105ms | 0.0 | 20001 | 20001 | 200010 | 200010 | Yes |
| Read Strings (20k rows text-heavy) | calamine | 70ms | 69ms | 71ms | 71ms | 0.0 | 20001 | 20001 | 200010 | 200010 | Yes |
| Read Strings (20k rows text-heavy) | edit-xlsx | 81ms | 80ms | 81ms | 81ms | 7.4 | 0 | 20001 | 0 | 200010 | No |
| Read Scale 1k rows | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0 | 1001 | 1001 | 10010 | 10010 | Yes |
| Read Scale 1k rows | calamine | 3ms | 3ms | 3ms | 3ms | 0.0 | 1001 | 1001 | 10010 | 10010 | Yes |
| Read Scale 1k rows | edit-xlsx | 4ms | 4ms | 4ms | 4ms | 0.0 | 0 | 1001 | 0 | 10010 | No |
| Read Scale 10k rows | SheetKit | 51ms | 51ms | 51ms | 51ms | 0.0 | 10001 | 10001 | 100010 | 100010 | Yes |
| Read Scale 10k rows | calamine | 33ms | 32ms | 33ms | 33ms | 0.0 | 10001 | 10001 | 100010 | 100010 | Yes |
| Read Scale 10k rows | edit-xlsx | 36ms | 36ms | 37ms | 37ms | 0.0 | 0 | 10001 | 0 | 100010 | No |
| Read Scale 100k rows | SheetKit | 511ms | 509ms | 513ms | 513ms | 8.4 | 100001 | 100001 | 1000010 | 1000010 | Yes |
| Read Scale 100k rows | calamine | 325ms | 324ms | 328ms | 328ms | 38.2 | 100001 | 100001 | 1000010 | 1000010 | Yes |
| Read Scale 100k rows | edit-xlsx | 373ms | 372ms | 374ms | 374ms | 0.0 | 0 | 100001 | 0 | 1000010 | No |
| Write 50000 rows x 20 cols | SheetKit | 475ms | 471ms | 476ms | 476ms | 100.7 | N/A | N/A | N/A | N/A | Yes |
| Write 50000 rows x 20 cols | rust_xlsxwriter | 886ms | 883ms | 896ms | 896ms | 41.1 | N/A | N/A | N/A | N/A | Yes |
| Write 50000 rows x 20 cols | edit-xlsx | 939ms | 937ms | 948ms | 948ms | 2.6 | N/A | N/A | N/A | N/A | Yes |
| Write 5000 styled rows | SheetKit | 27ms | 27ms | 28ms | 28ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 5000 styled rows | rust_xlsxwriter | 38ms | 38ms | 39ms | 39ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 5000 styled rows | edit-xlsx | 52ms | 51ms | 52ms | 52ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10 sheets x 5000 rows | SheetKit | 249ms | 248ms | 254ms | 254ms | 1.3 | N/A | N/A | N/A | N/A | Yes |
| Write 10 sheets x 5000 rows | rust_xlsxwriter | 338ms | 336ms | 338ms | 338ms | 9.8 | N/A | N/A | N/A | N/A | Yes |
| Write 10 sheets x 5000 rows | edit-xlsx | 414ms | 412ms | 416ms | 416ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10000 rows with formulas | SheetKit | 23ms | 22ms | 23ms | 23ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10000 rows with formulas | rust_xlsxwriter | 36ms | 35ms | 36ms | 36ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10000 rows with formulas | edit-xlsx | 59ms | 59ms | 60ms | 60ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 20000 text-heavy rows | SheetKit | 57ms | 56ms | 57ms | 57ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 20000 text-heavy rows | rust_xlsxwriter | 66ms | 65ms | 66ms | 66ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 20000 text-heavy rows | edit-xlsx | 71ms | 70ms | 71ms | 71ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 500 merged regions | SheetKit | 1ms | 1ms | 1ms | 1ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 500 merged regions | rust_xlsxwriter | 2ms | 2ms | 2ms | 2ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 500 merged regions | edit-xlsx | 5ms | 5ms | 5ms | 5ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 1k rows x 10 cols | SheetKit | 5ms | 4ms | 5ms | 5ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 1k rows x 10 cols | rust_xlsxwriter | 6ms | 6ms | 6ms | 6ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 1k rows x 10 cols | edit-xlsx | 11ms | 11ms | 12ms | 12ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10k rows x 10 cols | SheetKit | 46ms | 46ms | 47ms | 47ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10k rows x 10 cols | rust_xlsxwriter | 74ms | 73ms | 74ms | 74ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 10k rows x 10 cols | edit-xlsx | 86ms | 86ms | 86ms | 86ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 50k rows x 10 cols | SheetKit | 238ms | 236ms | 243ms | 243ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 50k rows x 10 cols | rust_xlsxwriter | 388ms | 385ms | 402ms | 402ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Write 50k rows x 10 cols | edit-xlsx | 423ms | 420ms | 427ms | 427ms | 3.1 | N/A | N/A | N/A | N/A | Yes |
| Write 100k rows x 10 cols | SheetKit | 493ms | 483ms | 499ms | 499ms | 18.2 | N/A | N/A | N/A | N/A | Yes |
| Write 100k rows x 10 cols | rust_xlsxwriter | 780ms | 780ms | 783ms | 783ms | 1.2 | N/A | N/A | N/A | N/A | Yes |
| Write 100k rows x 10 cols | edit-xlsx | 837ms | 832ms | 860ms | 860ms | 37.9 | N/A | N/A | N/A | N/A | Yes |
| Buffer round-trip (10000 rows) | SheetKit | 118ms | 116ms | 119ms | 119ms | 2.0 | N/A | N/A | N/A | N/A | Yes |
| Buffer round-trip (10000 rows) | xlsxwriter+calamine | 82ms | 82ms | 84ms | 84ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Streaming write (50000 rows) | SheetKit | 191ms | 190ms | 197ms | 197ms | 0.0 | N/A | N/A | N/A | N/A | Yes |
| Streaming write (50000 rows) | rust_xlsxwriter | 885ms | 881ms | 893ms | 893ms | 41.0 | N/A | N/A | N/A | N/A | Yes |
| Random-access read (1000 cells) | SheetKit | 465ms | 458ms | 468ms | 468ms | 13.4 | N/A | N/A | 1000 | N/A | Yes |
| Random-access read (1000 cells) | calamine | 321ms | 316ms | 322ms | 322ms | 38.2 | N/A | N/A | 1000 | N/A | Yes |
| Modify 1000 cells in 50k-row file | SheetKit | 668ms | 666ms | 678ms | 678ms | 53.3 | N/A | N/A | N/A | N/A | Yes |
| Modify 1000 cells in 50k-row file | edit-xlsx | 537ms | 536ms | 547ms | 547ms | 28.1 | N/A | N/A | N/A | N/A | Yes |

## Win Summary

| Library | Wins |
|---------|------|
| SheetKit | 11/21 |
| calamine | 8/21 |
| edit-xlsx | 1/21 |
| xlsxwriter+calamine | 1/21 |
