# Rust Excel Library Comparison Benchmark

Benchmark run: 2026-02-11T03:39:36Z

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
| CPU | unknown |
| RAM | 21 GB |
| OS | linux x86_64 |
| Rust | rustc 1.93.0 (254b59607 2026-01-19) |
| Profile | release (opt-level=3, LTO=fat) |
| Iterations | 5 runs per scenario, 1 warmup |

## Read

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Read Large Data (50k rows x 20 cols) | 1.05s | 620ms | 88ms | N/A | N/A | edit-xlsx |
| Read Heavy Styles (5k rows, formatted) | 43ms | 31ms | 5ms | N/A | N/A | edit-xlsx |
| Read Multi-Sheet (10 sheets x 5k rows) | 588ms | 327ms | 92ms | N/A | N/A | edit-xlsx |
| Read Formulas (10k rows) | 55ms | 29ms | 2ms | N/A | N/A | edit-xlsx |
| Read Strings (20k rows text-heavy) | 197ms | 131ms | 23ms | N/A | N/A | edit-xlsx |

## Read (Scale)

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Read Scale 1k rows | 9ms | 7ms | 3ms | N/A | N/A | edit-xlsx |
| Read Scale 10k rows | 92ms | 64ms | 12ms | N/A | N/A | edit-xlsx |
| Read Scale 100k rows | 1.20s | 694ms | 116ms | N/A | N/A | edit-xlsx |

## Write

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Write 50000 rows x 20 cols | 1.15s | N/A | 2.28s | 2.19s | N/A | SheetKit |
| Write 5000 styled rows | 48ms | N/A | 121ms | 85ms | N/A | SheetKit |
| Write 10 sheets x 5000 rows | 528ms | N/A | 964ms | 941ms | N/A | SheetKit |
| Write 10000 rows with formulas | 39ms | N/A | 143ms | 86ms | N/A | SheetKit |
| Write 20000 text-heavy rows | 96ms | N/A | 153ms | 160ms | N/A | SheetKit |
| Write 500 merged regions | 14ms | N/A | 25ms | 6ms | N/A | rust_xlsxwriter |

## Write (Scale)

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Write 1k rows x 10 cols | 10ms | N/A | 41ms | 17ms | N/A | SheetKit |
| Write 10k rows x 10 cols | 90ms | N/A | 208ms | 188ms | N/A | SheetKit |
| Write 50k rows x 10 cols | 527ms | N/A | 983ms | 962ms | N/A | SheetKit |
| Write 100k rows x 10 cols | 1.22s | N/A | 2.10s | 1.97s | N/A | SheetKit |

## Round-Trip

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Buffer round-trip (10000 rows) | 239ms | N/A | N/A | N/A | 178ms | xlsxwriter+calamine |

## Streaming

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Streaming write (50000 rows) | 1.32s | N/A | N/A | 2.12s | N/A | SheetKit |

## Random Access

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Random-access read (1000 cells) | 966ms | 602ms | N/A | N/A | N/A | calamine |

## Modify

| Scenario | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner |
|----------|--------|--------|--------|--------|--------|--------|
| Modify 1000 cells in 50k-row file | 1.46s | N/A | N/A | N/A | N/A | SheetKit |

## Detailed Statistics

| Scenario | Library | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|---------|--------|-----|-----|-----|---------------|
| Read Large Data (50k rows x 20 cols) | SheetKit | 1.05s | 1.02s | 1.36s | 1.36s | 168.3 |
| Read Large Data (50k rows x 20 cols) | calamine | 620ms | 618ms | 652ms | 652ms | 0.0 |
| Read Large Data (50k rows x 20 cols) | edit-xlsx | 88ms | 84ms | 92ms | 92ms | 0.0 |
| Read Heavy Styles (5k rows, formatted) | SheetKit | 43ms | 42ms | 45ms | 45ms | 0.0 |
| Read Heavy Styles (5k rows, formatted) | calamine | 31ms | 30ms | 34ms | 34ms | 0.0 |
| Read Heavy Styles (5k rows, formatted) | edit-xlsx | 5ms | 5ms | 6ms | 6ms | 0.0 |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit | 588ms | 541ms | 594ms | 594ms | 0.0 |
| Read Multi-Sheet (10 sheets x 5k rows) | calamine | 327ms | 324ms | 333ms | 333ms | 0.0 |
| Read Multi-Sheet (10 sheets x 5k rows) | edit-xlsx | 92ms | 86ms | 94ms | 94ms | 0.0 |
| Read Formulas (10k rows) | SheetKit | 55ms | 54ms | 67ms | 67ms | 0.0 |
| Read Formulas (10k rows) | calamine | 29ms | 28ms | 34ms | 34ms | 0.0 |
| Read Formulas (10k rows) | edit-xlsx | 2ms | 2ms | 2ms | 2ms | 0.0 |
| Read Strings (20k rows text-heavy) | SheetKit | 197ms | 188ms | 239ms | 239ms | 0.0 |
| Read Strings (20k rows text-heavy) | calamine | 131ms | 128ms | 141ms | 141ms | 0.0 |
| Read Strings (20k rows text-heavy) | edit-xlsx | 23ms | 21ms | 24ms | 24ms | 0.0 |
| Read Scale 1k rows | SheetKit | 9ms | 9ms | 10ms | 10ms | 0.0 |
| Read Scale 1k rows | calamine | 7ms | 6ms | 7ms | 7ms | 0.0 |
| Read Scale 1k rows | edit-xlsx | 3ms | 3ms | 3ms | 3ms | 0.0 |
| Read Scale 10k rows | SheetKit | 92ms | 84ms | 94ms | 94ms | 0.0 |
| Read Scale 10k rows | calamine | 64ms | 61ms | 70ms | 70ms | 0.0 |
| Read Scale 10k rows | edit-xlsx | 12ms | 12ms | 13ms | 13ms | 0.0 |
| Read Scale 100k rows | SheetKit | 1.20s | 1.00s | 1.27s | 1.27s | 68.7 |
| Read Scale 100k rows | calamine | 694ms | 682ms | 715ms | 715ms | 0.0 |
| Read Scale 100k rows | edit-xlsx | 116ms | 111ms | 121ms | 121ms | 0.0 |
| Write 50000 rows x 20 cols | SheetKit | 1.15s | 1.13s | 1.28s | 1.28s | 56.6 |
| Write 50000 rows x 20 cols | rust_xlsxwriter | 2.19s | 2.16s | 2.22s | 2.22s | 0.0 |
| Write 50000 rows x 20 cols | edit-xlsx | 2.28s | 2.26s | 2.34s | 2.34s | 0.0 |
| Write 5000 styled rows | SheetKit | 48ms | 46ms | 56ms | 56ms | 0.0 |
| Write 5000 styled rows | rust_xlsxwriter | 85ms | 84ms | 86ms | 86ms | 0.0 |
| Write 5000 styled rows | edit-xlsx | 121ms | 118ms | 127ms | 127ms | 0.0 |
| Write 10 sheets x 5000 rows | SheetKit | 528ms | 513ms | 586ms | 586ms | 0.0 |
| Write 10 sheets x 5000 rows | rust_xlsxwriter | 941ms | 918ms | 951ms | 951ms | 0.0 |
| Write 10 sheets x 5000 rows | edit-xlsx | 964ms | 940ms | 1.01s | 1.01s | 0.0 |
| Write 10000 rows with formulas | SheetKit | 39ms | 39ms | 40ms | 40ms | 0.0 |
| Write 10000 rows with formulas | rust_xlsxwriter | 86ms | 84ms | 87ms | 87ms | 0.0 |
| Write 10000 rows with formulas | edit-xlsx | 143ms | 140ms | 145ms | 145ms | 0.0 |
| Write 20000 text-heavy rows | SheetKit | 96ms | 93ms | 101ms | 101ms | 0.0 |
| Write 20000 text-heavy rows | rust_xlsxwriter | 160ms | 158ms | 160ms | 160ms | 0.0 |
| Write 20000 text-heavy rows | edit-xlsx | 153ms | 146ms | 160ms | 160ms | 0.0 |
| Write 500 merged regions | SheetKit | 14ms | 14ms | 14ms | 14ms | 0.0 |
| Write 500 merged regions | rust_xlsxwriter | 6ms | 6ms | 6ms | 6ms | 0.0 |
| Write 500 merged regions | edit-xlsx | 25ms | 23ms | 30ms | 30ms | 0.0 |
| Write 1k rows x 10 cols | SheetKit | 10ms | 9ms | 10ms | 10ms | 0.0 |
| Write 1k rows x 10 cols | rust_xlsxwriter | 17ms | 16ms | 19ms | 19ms | 0.0 |
| Write 1k rows x 10 cols | edit-xlsx | 41ms | 38ms | 47ms | 47ms | 0.0 |
| Write 10k rows x 10 cols | SheetKit | 90ms | 86ms | 95ms | 95ms | 0.0 |
| Write 10k rows x 10 cols | rust_xlsxwriter | 188ms | 186ms | 195ms | 195ms | 0.0 |
| Write 10k rows x 10 cols | edit-xlsx | 208ms | 203ms | 211ms | 211ms | 0.0 |
| Write 50k rows x 10 cols | SheetKit | 527ms | 510ms | 551ms | 551ms | 0.0 |
| Write 50k rows x 10 cols | rust_xlsxwriter | 962ms | 938ms | 971ms | 971ms | 0.0 |
| Write 50k rows x 10 cols | edit-xlsx | 983ms | 961ms | 1.00s | 1.00s | 0.0 |
| Write 100k rows x 10 cols | SheetKit | 1.22s | 1.18s | 1.24s | 1.24s | 0.0 |
| Write 100k rows x 10 cols | rust_xlsxwriter | 1.97s | 1.93s | 1.99s | 1.99s | 0.0 |
| Write 100k rows x 10 cols | edit-xlsx | 2.10s | 2.06s | 2.12s | 2.12s | 0.0 |
| Buffer round-trip (10000 rows) | SheetKit | 239ms | 226ms | 244ms | 244ms | 0.0 |
| Buffer round-trip (10000 rows) | xlsxwriter+calamine | 178ms | 170ms | 189ms | 189ms | 0.0 |
| Streaming write (50000 rows) | SheetKit | 1.32s | 1.28s | 1.33s | 1.33s | 0.0 |
| Streaming write (50000 rows) | rust_xlsxwriter | 2.12s | 2.10s | 2.23s | 2.23s | 0.0 |
| Random-access read (1000 cells) | SheetKit | 966ms | 921ms | 993ms | 993ms | 0.0 |
| Random-access read (1000 cells) | calamine | 602ms | 598ms | 625ms | 625ms | 0.0 |
| Modify 1000 cells in 50k-row file | SheetKit | 1.46s | 1.40s | 1.54s | 1.54s | 0.0 |

## Win Summary

| Library | Wins |
|---------|------|
| SheetKit | 11/22 |
| edit-xlsx | 8/22 |
| rust_xlsxwriter | 1/22 |
| xlsxwriter+calamine | 1/22 |
| calamine | 1/22 |
