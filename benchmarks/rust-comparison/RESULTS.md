# Rust Excel Library Comparison Benchmark

Benchmark run: 2026-02-11T04:17:20Z

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
| Read Large Data (50k rows x 20 cols) | 390ms | 299ms | 35ms | N/A | N/A | edit-xlsx |
| Read Heavy Styles (5k rows, formatted) | 20ms | 16ms | 2ms | N/A | N/A | edit-xlsx |
| Read Multi-Sheet (10 sheets x 5k rows) | 228ms | 170ms | 35ms | N/A | N/A | edit-xlsx |
| Read Formulas (10k rows) | 24ms | 14ms | 0ms | N/A | N/A | edit-xlsx |
| Read Strings (20k rows text-heavy) | 83ms | 66ms | 9ms | N/A | N/A | edit-xlsx |

## Read (Scale)

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Read Scale 1k rows | 4ms | 3ms | 1ms | N/A | N/A | edit-xlsx |
| Read Scale 10k rows | 39ms | 31ms | 4ms | N/A | N/A | edit-xlsx |
| Read Scale 100k rows | 406ms | 314ms | 41ms | N/A | N/A | edit-xlsx |

## Write

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Write 50000 rows x 20 cols | 459ms | N/A | 886ms | 847ms | N/A | SheetKit |
| Write 5000 styled rows | 25ms | N/A | 49ms | 37ms | N/A | SheetKit |
| Write 10 sheets x 5000 rows | 237ms | N/A | 393ms | 326ms | N/A | SheetKit |
| Write 10000 rows with formulas | 21ms | N/A | 56ms | 34ms | N/A | SheetKit |
| Write 20000 text-heavy rows | 53ms | N/A | 66ms | 64ms | N/A | SheetKit |
| Write 500 merged regions | 8ms | N/A | 5ms | 2ms | N/A | rust_xlsxwriter |

## Write (Scale)

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Write 1k rows x 10 cols | 4ms | N/A | 10ms | 6ms | N/A | SheetKit |
| Write 10k rows x 10 cols | 42ms | N/A | 81ms | 71ms | N/A | SheetKit |
| Write 50k rows x 10 cols | 230ms | N/A | 402ms | 371ms | N/A | SheetKit |
| Write 100k rows x 10 cols | 466ms | N/A | 813ms | 748ms | N/A | SheetKit |

## Round-Trip

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Buffer round-trip (10000 rows) | 105ms | N/A | N/A | N/A | 79ms | xlsxwriter+calamine |

## Streaming

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Streaming write (50000 rows) | 184ms | N/A | N/A | 858ms | N/A | SheetKit |

## Random Access

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Random-access read (1000 cells) | 382ms | 308ms | N/A | N/A | N/A | calamine |

## Modify

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Modify 1000 cells in 50k-row file | 588ms | N/A | N/A | N/A | N/A | SheetKit |

## Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|---------|--------|-----|-----|-----|---------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 390ms | 388ms | 401ms | 401ms | 68.7 |
| Read Large Data (50k rows x 20 cols) | calamine | 299ms | 298ms | 301ms | 301ms | 43.9 |
| Read Large Data (50k rows x 20 cols) | edit-xlsx | 35ms | 35ms | 36ms | 36ms | 13.3 |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 20ms | 20ms | 21ms | 21ms | 0.0 |
| Read Heavy Styles (5k rows, formatted) | calamine | 16ms | 15ms | 16ms | 16ms | 0.0 |
| Read Heavy Styles (5k rows, formatted) | edit-xlsx | 2ms | 2ms | 2ms | 2ms | 0.0 |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 228ms | 227ms | 231ms | 231ms | 0.0 |
| Read Multi-Sheet (10 sheets x 5k rows) | calamine | 170ms | 169ms | 170ms | 170ms | 0.0 |
| Read Multi-Sheet (10 sheets x 5k rows) | edit-xlsx | 35ms | 35ms | 35ms | 35ms | 4.0 |
| Read Formulas (10k rows) | SheetKit | 24ms | 24ms | 25ms | 25ms | 2.7 |
| Read Formulas (10k rows) | calamine | 14ms | 14ms | 14ms | 14ms | 0.0 |
| Read Formulas (10k rows) | edit-xlsx | 0ms | 0ms | 0ms | 0ms | 0.0 |
| Read Strings (20k rows text-heavy) | SheetKit | 83ms | 83ms | 84ms | 84ms | 3.5 |
| Read Strings (20k rows text-heavy) | calamine | 66ms | 66ms | 66ms | 66ms | 0.0 |
| Read Strings (20k rows text-heavy) | edit-xlsx | 9ms | 9ms | 9ms | 9ms | 0.0 |
| Read Scale 1k rows | SheetKit | 4ms | 4ms | 4ms | 4ms | 0.0 |
| Read Scale 1k rows | calamine | 3ms | 3ms | 3ms | 3ms | 0.0 |
| Read Scale 1k rows | edit-xlsx | 1ms | 1ms | 1ms | 1ms | 0.0 |
| Read Scale 10k rows | SheetKit | 39ms | 39ms | 39ms | 39ms | 0.0 |
| Read Scale 10k rows | calamine | 31ms | 30ms | 31ms | 31ms | 0.0 |
| Read Scale 10k rows | edit-xlsx | 4ms | 4ms | 4ms | 4ms | 0.0 |
| Read Scale 100k rows | SheetKit | 406ms | 402ms | 408ms | 408ms | 41.4 |
| Read Scale 100k rows | calamine | 314ms | 313ms | 315ms | 315ms | 0.0 |
| Read Scale 100k rows | edit-xlsx | 41ms | 40ms | 41ms | 41ms | 17.6 |
| Write 50000 rows x 20 cols | SheetKit | 459ms | 455ms | 461ms | 461ms | 15.8 |
| Write 50000 rows x 20 cols | rust_xlsxwriter | 847ms | 844ms | 849ms | 849ms | 41.1 |
| Write 50000 rows x 20 cols | edit-xlsx | 886ms | 882ms | 889ms | 889ms | 117.6 |
| Write 5000 styled rows | SheetKit | 25ms | 25ms | 26ms | 26ms | 0.0 |
| Write 5000 styled rows | rust_xlsxwriter | 37ms | 37ms | 37ms | 37ms | 0.0 |
| Write 5000 styled rows | edit-xlsx | 49ms | 49ms | 49ms | 49ms | 2.1 |
| Write 10 sheets x 5000 rows | SheetKit | 237ms | 235ms | 237ms | 237ms | 15.4 |
| Write 10 sheets x 5000 rows | rust_xlsxwriter | 326ms | 324ms | 384ms | 384ms | 3.4 |
| Write 10 sheets x 5000 rows | edit-xlsx | 393ms | 393ms | 399ms | 399ms | 0.0 |
| Write 10000 rows with formulas | SheetKit | 21ms | 21ms | 21ms | 21ms | 0.0 |
| Write 10000 rows with formulas | rust_xlsxwriter | 34ms | 34ms | 34ms | 34ms | 0.0 |
| Write 10000 rows with formulas | edit-xlsx | 56ms | 56ms | 56ms | 56ms | 6.1 |
| Write 20000 text-heavy rows | SheetKit | 53ms | 52ms | 53ms | 53ms | 1.7 |
| Write 20000 text-heavy rows | rust_xlsxwriter | 64ms | 64ms | 64ms | 64ms | 0.0 |
| Write 20000 text-heavy rows | edit-xlsx | 66ms | 65ms | 66ms | 66ms | 1.3 |
| Write 500 merged regions | SheetKit | 8ms | 8ms | 9ms | 9ms | 0.0 |
| Write 500 merged regions | rust_xlsxwriter | 2ms | 2ms | 2ms | 2ms | 0.0 |
| Write 500 merged regions | edit-xlsx | 5ms | 4ms | 5ms | 5ms | 0.0 |
| Write 1k rows x 10 cols | SheetKit | 4ms | 4ms | 4ms | 4ms | 0.0 |
| Write 1k rows x 10 cols | rust_xlsxwriter | 6ms | 6ms | 6ms | 6ms | 0.0 |
| Write 1k rows x 10 cols | edit-xlsx | 10ms | 10ms | 11ms | 11ms | 0.0 |
| Write 10k rows x 10 cols | SheetKit | 42ms | 42ms | 43ms | 43ms | 0.0 |
| Write 10k rows x 10 cols | rust_xlsxwriter | 71ms | 71ms | 71ms | 71ms | 0.0 |
| Write 10k rows x 10 cols | edit-xlsx | 81ms | 80ms | 81ms | 81ms | 0.1 |
| Write 50k rows x 10 cols | SheetKit | 230ms | 228ms | 231ms | 231ms | 6.9 |
| Write 50k rows x 10 cols | rust_xlsxwriter | 371ms | 370ms | 373ms | 373ms | 4.6 |
| Write 50k rows x 10 cols | edit-xlsx | 402ms | 400ms | 403ms | 403ms | 19.0 |
| Write 100k rows x 10 cols | SheetKit | 466ms | 460ms | 467ms | 467ms | 18.5 |
| Write 100k rows x 10 cols | rust_xlsxwriter | 748ms | 745ms | 750ms | 750ms | 41.1 |
| Write 100k rows x 10 cols | edit-xlsx | 813ms | 810ms | 814ms | 814ms | 43.3 |
| Buffer round-trip (10000 rows) | SheetKit | 105ms | 103ms | 105ms | 105ms | 0.0 |
| Buffer round-trip (10000 rows) | xlsxwriter+calamine | 79ms | 79ms | 79ms | 79ms | 0.0 |
| Streaming write (50000 rows) | SheetKit | 184ms | 183ms | 184ms | 184ms | 0.0 |
| Streaming write (50000 rows) | rust_xlsxwriter | 858ms | 856ms | 919ms | 919ms | 29.0 |
| Random-access read (1000 cells) | SheetKit | 382ms | 380ms | 383ms | 383ms | 12.2 |
| Random-access read (1000 cells) | calamine | 308ms | 308ms | 310ms | 310ms | 0.0 |
| Modify 1000 cells in 50k-row file | SheetKit | 588ms | 582ms | 599ms | 599ms | 65.7 |

## Win Summary

| Library | Wins |
|---------|------|
| SheetKit | 11/22 |
| edit-xlsx | 8/22 |
| calamine | 1/22 |
| xlsxwriter+calamine | 1/22 |
| rust_xlsxwriter | 1/22 |
