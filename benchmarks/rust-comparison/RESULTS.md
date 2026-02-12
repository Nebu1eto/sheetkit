# Rust Excel Library Comparison Benchmark

Benchmark run: 2026-02-12T12:05:55Z

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
| Read Large Data (50k rows x 20 cols) | 497ms | 480ms | 322ms | 39ms* | N/A | N/A | calamine |
| Read Heavy Styles (5k rows, formatted) | 26ms | 26ms | 17ms | 2ms* | N/A | N/A | calamine |
| Read Multi-Sheet (10 sheets x 5k rows) | 290ms | 287ms | 181ms | 39ms* | N/A | N/A | calamine |
| Read Formulas (10k rows) | 32ms | 31ms | 15ms | 0ms* | N/A | N/A | calamine |
| Read Strings (20k rows text-heavy) | 104ms | 104ms | 70ms | 10ms* | N/A | N/A | calamine |

## Read (Scale)

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Read Scale 1k rows | 5ms | 5ms | 3ms | 1ms* | N/A | N/A | calamine |
| Read Scale 10k rows | 51ms | 49ms | 32ms | 4ms* | N/A | N/A | calamine |
| Read Scale 100k rows | 511ms | 498ms | 331ms | 43ms* | N/A | N/A | calamine |

## Write

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Write 50000 rows x 20 cols | 488ms | N/A | N/A | 937ms | 911ms | N/A | SheetKit |
| Write 5000 styled rows | 27ms | N/A | N/A | 54ms | 39ms | N/A | SheetKit |
| Write 10 sheets x 5000 rows | 256ms | N/A | N/A | 425ms | 350ms | N/A | SheetKit |
| Write 10000 rows with formulas | 23ms | N/A | N/A | 60ms | 37ms | N/A | SheetKit |
| Write 20000 text-heavy rows | 57ms | N/A | N/A | 72ms | 68ms | N/A | SheetKit |
| Write 500 merged regions | 1ms | N/A | N/A | 6ms | 2ms | N/A | SheetKit |

## Write (Scale)

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Write 1k rows x 10 cols | 5ms | N/A | N/A | 12ms | 7ms | N/A | SheetKit |
| Write 10k rows x 10 cols | 47ms | N/A | N/A | 89ms | 77ms | N/A | SheetKit |
| Write 50k rows x 10 cols | 240ms | N/A | N/A | 427ms | 396ms | N/A | SheetKit |
| Write 100k rows x 10 cols | 482ms | N/A | N/A | 861ms | 799ms | N/A | SheetKit |

## Round-Trip

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Buffer round-trip (10000 rows) | 120ms | N/A | N/A | N/A | N/A | 84ms | xlsxwriter+calamine |

## Streaming

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Streaming write (50000 rows) | 195ms | N/A | N/A | N/A | 910ms | N/A | SheetKit |

## Random Access

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Random-access read (1000 cells) | 505ms | 460ms | 321ms | N/A | N/A | N/A | calamine |

## Modify

| Scenario | SheetKit | SheetKit (lazy) | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|--------|
| Modify 1000 cells in 50k-row file | 693ms | 671ms | N/A | N/A | N/A | N/A | SheetKit (lazy) |

* `*` indicates `cells_read = 0`; excluded from Winner selection as non-comparable.

## Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Peak Mem (MB) | Cells Read |
|----------|---------|--------|-----|-----|-----|---------------|------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 497ms | 489ms | 504ms | 504ms | 60.3 | 1000020 |
| Read Large Data (50k rows x 20 cols) | SheetKit (lazy) | 480ms | 478ms | 483ms | 483ms | 17.3 | 1000020 |
| Read Large Data (50k rows x 20 cols) | calamine | 322ms | 320ms | 323ms | 323ms | 38.2 | 1000020 |
| Read Large Data (50k rows x 20 cols) | edit-xlsx | 39ms | 38ms | 39ms | 39ms | 3.3 | 0 |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 26ms | 26ms | 27ms | 27ms | 1.5 | 50010 |
| Read Heavy Styles (5k rows, formatted) | SheetKit (lazy) | 26ms | 26ms | 26ms | 26ms | 0.0 | 50010 |
| Read Heavy Styles (5k rows, formatted) | calamine | 17ms | 17ms | 17ms | 17ms | 0.0 | 50010 |
| Read Heavy Styles (5k rows, formatted) | edit-xlsx | 2ms | 2ms | 2ms | 2ms | 0.0 | 0 |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 290ms | 289ms | 291ms | 291ms | 0.9 | 500100 |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit (lazy) | 287ms | 286ms | 290ms | 290ms | 0.2 | 500100 |
| Read Multi-Sheet (10 sheets x 5k rows) | calamine | 181ms | 178ms | 191ms | 191ms | 0.0 | 500100 |
| Read Multi-Sheet (10 sheets x 5k rows) | edit-xlsx | 39ms | 38ms | 39ms | 39ms | 0.0 | 0 |
| Read Formulas (10k rows) | SheetKit | 32ms | 31ms | 32ms | 32ms | 0.0 | 70007 |
| Read Formulas (10k rows) | SheetKit (lazy) | 31ms | 30ms | 31ms | 31ms | 0.0 | 70007 |
| Read Formulas (10k rows) | calamine | 15ms | 15ms | 15ms | 15ms | 0.0 | 20007 |
| Read Formulas (10k rows) | edit-xlsx | 0ms | 0ms | 0ms | 0ms | 0.0 | 0 |
| Read Strings (20k rows text-heavy) | SheetKit | 104ms | 104ms | 105ms | 105ms | 0.0 | 200010 |
| Read Strings (20k rows text-heavy) | SheetKit (lazy) | 104ms | 103ms | 105ms | 105ms | 0.0 | 200010 |
| Read Strings (20k rows text-heavy) | calamine | 70ms | 69ms | 70ms | 70ms | 7.7 | 200010 |
| Read Strings (20k rows text-heavy) | edit-xlsx | 10ms | 9ms | 10ms | 10ms | 0.0 | 0 |
| Read Scale 1k rows | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0 | 10010 |
| Read Scale 1k rows | SheetKit (lazy) | 5ms | 5ms | 5ms | 5ms | 0.0 | 10010 |
| Read Scale 1k rows | calamine | 3ms | 3ms | 3ms | 3ms | 0.0 | 10010 |
| Read Scale 1k rows | edit-xlsx | 1ms | 1ms | 1ms | 1ms | 0.0 | 0 |
| Read Scale 10k rows | SheetKit | 51ms | 51ms | 51ms | 51ms | 0.0 | 100010 |
| Read Scale 10k rows | SheetKit (lazy) | 49ms | 49ms | 50ms | 50ms | 0.0 | 100010 |
| Read Scale 10k rows | calamine | 32ms | 32ms | 33ms | 33ms | 0.0 | 100010 |
| Read Scale 10k rows | edit-xlsx | 4ms | 4ms | 5ms | 5ms | 0.0 | 0 |
| Read Scale 100k rows | SheetKit | 511ms | 510ms | 513ms | 513ms | 41.4 | 1000010 |
| Read Scale 100k rows | SheetKit (lazy) | 498ms | 497ms | 502ms | 502ms | 29.4 | 1000010 |
| Read Scale 100k rows | calamine | 331ms | 330ms | 333ms | 333ms | 0.0 | 1000010 |
| Read Scale 100k rows | edit-xlsx | 43ms | 43ms | 44ms | 44ms | 0.6 | 0 |
| Write 50000 rows x 20 cols | SheetKit | 488ms | 486ms | 491ms | 491ms | 40.0 | N/A |
| Write 50000 rows x 20 cols | rust_xlsxwriter | 911ms | 911ms | 914ms | 914ms | 0.0 | N/A |
| Write 50000 rows x 20 cols | edit-xlsx | 937ms | 934ms | 939ms | 939ms | 112.1 | N/A |
| Write 5000 styled rows | SheetKit | 27ms | 26ms | 27ms | 27ms | 0.0 | N/A |
| Write 5000 styled rows | rust_xlsxwriter | 39ms | 39ms | 40ms | 40ms | 0.0 | N/A |
| Write 5000 styled rows | edit-xlsx | 54ms | 53ms | 54ms | 54ms | 0.0 | N/A |
| Write 10 sheets x 5000 rows | SheetKit | 256ms | 255ms | 257ms | 257ms | 12.0 | N/A |
| Write 10 sheets x 5000 rows | rust_xlsxwriter | 350ms | 346ms | 368ms | 368ms | 9.8 | N/A |
| Write 10 sheets x 5000 rows | edit-xlsx | 425ms | 421ms | 427ms | 427ms | 0.0 | N/A |
| Write 10000 rows with formulas | SheetKit | 23ms | 23ms | 23ms | 23ms | 0.0 | N/A |
| Write 10000 rows with formulas | rust_xlsxwriter | 37ms | 36ms | 37ms | 37ms | 0.0 | N/A |
| Write 10000 rows with formulas | edit-xlsx | 60ms | 60ms | 61ms | 61ms | 0.0 | N/A |
| Write 20000 text-heavy rows | SheetKit | 57ms | 56ms | 58ms | 58ms | 0.0 | N/A |
| Write 20000 text-heavy rows | rust_xlsxwriter | 68ms | 67ms | 68ms | 68ms | 0.0 | N/A |
| Write 20000 text-heavy rows | edit-xlsx | 72ms | 72ms | 77ms | 77ms | 0.0 | N/A |
| Write 500 merged regions | SheetKit | 1ms | 1ms | 1ms | 1ms | 0.0 | N/A |
| Write 500 merged regions | rust_xlsxwriter | 2ms | 2ms | 2ms | 2ms | 0.0 | N/A |
| Write 500 merged regions | edit-xlsx | 6ms | 5ms | 7ms | 7ms | 0.0 | N/A |
| Write 1k rows x 10 cols | SheetKit | 5ms | 5ms | 5ms | 5ms | 0.0 | N/A |
| Write 1k rows x 10 cols | rust_xlsxwriter | 7ms | 6ms | 7ms | 7ms | 0.0 | N/A |
| Write 1k rows x 10 cols | edit-xlsx | 12ms | 11ms | 12ms | 12ms | 0.0 | N/A |
| Write 10k rows x 10 cols | SheetKit | 47ms | 46ms | 47ms | 47ms | 0.0 | N/A |
| Write 10k rows x 10 cols | rust_xlsxwriter | 77ms | 76ms | 77ms | 77ms | 0.0 | N/A |
| Write 10k rows x 10 cols | edit-xlsx | 89ms | 88ms | 89ms | 89ms | 0.0 | N/A |
| Write 50k rows x 10 cols | SheetKit | 240ms | 239ms | 241ms | 241ms | 0.0 | N/A |
| Write 50k rows x 10 cols | rust_xlsxwriter | 396ms | 394ms | 404ms | 404ms | 0.0 | N/A |
| Write 50k rows x 10 cols | edit-xlsx | 427ms | 424ms | 429ms | 429ms | 0.0 | N/A |
| Write 100k rows x 10 cols | SheetKit | 482ms | 477ms | 487ms | 487ms | 66.0 | N/A |
| Write 100k rows x 10 cols | rust_xlsxwriter | 799ms | 798ms | 802ms | 802ms | 0.0 | N/A |
| Write 100k rows x 10 cols | edit-xlsx | 861ms | 858ms | 865ms | 865ms | 0.1 | N/A |
| Buffer round-trip (10000 rows) | SheetKit | 120ms | 119ms | 122ms | 122ms | 0.0 | N/A |
| Buffer round-trip (10000 rows) | xlsxwriter+calamine | 84ms | 83ms | 85ms | 85ms | 0.0 | N/A |
| Streaming write (50000 rows) | SheetKit | 195ms | 194ms | 196ms | 196ms | 0.0 | N/A |
| Streaming write (50000 rows) | rust_xlsxwriter | 910ms | 908ms | 911ms | 911ms | 0.0 | N/A |
| Random-access read (1000 cells) | SheetKit | 505ms | 500ms | 516ms | 516ms | 13.0 | 1000 |
| Random-access read (1000 cells) | SheetKit (lazy) | 460ms | 455ms | 462ms | 462ms | 17.2 | 1000 |
| Random-access read (1000 cells) | calamine | 321ms | 320ms | 326ms | 326ms | 38.9 | 1000 |
| Modify 1000 cells in 50k-row file | SheetKit | 693ms | 691ms | 696ms | 696ms | 3.1 | N/A |
| Modify 1000 cells in 50k-row file | SheetKit (lazy) | 671ms | 671ms | 674ms | 674ms | 26.2 | N/A |

## Win Summary

| Library | Wins |
|---------|------|
| SheetKit | 11/22 |
| calamine | 9/22 |
| SheetKit (lazy) | 1/22 |
| xlsxwriter+calamine | 1/22 |
