# SheetKit Rust Native Benchmark

Benchmark run: 2026-02-12T12:02:02Z

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

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Read Large Data (50k rows x 20 cols) | 499ms | 497ms | 512ms | 512ms | 60.2 |
| Read Large Data (50k rows x 20 cols) (lazy) | 491ms | 488ms | 495ms | 495ms | 17.3 |
| Read Heavy Styles (5k rows, formatted) | 26ms | 26ms | 27ms | 27ms | 0.0 |
| Read Heavy Styles (5k rows, formatted) (lazy) | 26ms | 26ms | 26ms | 26ms | 0.0 |
| Read Multi-Sheet (10 sheets x 5k rows) | 299ms | 297ms | 300ms | 300ms | 19.0 |
| Read Multi-Sheet (10 sheets x 5k rows) (lazy) | 292ms | 291ms | 293ms | 293ms | 17.2 |
| Read Formulas (10k rows) | 33ms | 32ms | 34ms | 34ms | 2.7 |
| Read Formulas (10k rows) (lazy) | 31ms | 31ms | 31ms | 31ms | 0.0 |
| Read Strings (20k rows text-heavy) | 107ms | 106ms | 107ms | 107ms | 0.0 |
| Read Strings (20k rows text-heavy) (lazy) | 106ms | 105ms | 107ms | 107ms | 0.0 |
| Read Data Validation (5k rows, 8 rules) | 21ms | 20ms | 21ms | 21ms | 0.0 |
| Read Data Validation (5k rows, 8 rules) (lazy) | 20ms | 20ms | 20ms | 20ms | 0.0 |
| Read Comments (2k rows with comments) | 8ms | 8ms | 9ms | 9ms | 0.0 |
| Read Comments (2k rows with comments) (lazy) | 7ms | 7ms | 7ms | 7ms | 0.0 |
| Read Merged Cells (500 regions) | 1ms | 1ms | 1ms | 1ms | 0.0 |
| Read Merged Cells (500 regions) (lazy) | 1ms | 1ms | 1ms | 1ms | 0.0 |
| Read Mixed Workload (ERP document) | 26ms | 26ms | 27ms | 27ms | 0.0 |
| Read Mixed Workload (ERP document) (lazy) | 26ms | 25ms | 26ms | 26ms | 0.0 |

## Read (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Read Scale 1k rows | 5ms | 5ms | 5ms | 5ms | 0.0 |
| Read Scale 1k rows (lazy) | 5ms | 5ms | 5ms | 5ms | 0.0 |
| Read Scale 10k rows | 52ms | 51ms | 54ms | 54ms | 0.0 |
| Read Scale 10k rows (lazy) | 52ms | 51ms | 53ms | 53ms | 2.4 |
| Read Scale 100k rows | 525ms | 524ms | 527ms | 527ms | 33.0 |
| Read Scale 100k rows (lazy) | 511ms | 511ms | 513ms | 513ms | 20.6 |

## Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 50000 rows x 20 cols | 510ms | 506ms | 511ms | 511ms | 31.5 |
| Write 5000 styled rows | 28ms | 27ms | 28ms | 28ms | 0.0 |
| Write 10 sheets x 5000 rows | 269ms | 266ms | 270ms | 270ms | 0.8 |
| Write 10000 rows with formulas | 23ms | 23ms | 23ms | 23ms | 0.0 |
| Write 20000 text-heavy rows | 105ms | 104ms | 105ms | 105ms | 0.0 |

## Write (DV)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 5000 rows + 8 validation rules | 9ms | 9ms | 9ms | 9ms | 0.0 |

## Write (Comments)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 2000 rows with comments | 6ms | 6ms | 7ms | 7ms | 0.0 |

## Write (Merge)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 500 merged regions | 1ms | 1ms | 1ms | 1ms | 0.0 |

## Write (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 1k rows x 10 cols | 5ms | 5ms | 5ms | 5ms | 0.0 |
| Write 10k rows x 10 cols | 48ms | 48ms | 48ms | 48ms | 0.0 |
| Write 50k rows x 10 cols | 251ms | 248ms | 251ms | 251ms | 0.0 |
| Write 100k rows x 10 cols | 515ms | 512ms | 517ms | 517ms | 0.0 |

## Round-Trip

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Buffer round-trip (10000 rows) | 124ms | 123ms | 127ms | 127ms | 0.0 |

## Streaming

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Streaming write (50000 rows) | 204ms | 202ms | 206ms | 206ms | 0.0 |

## Random Access

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Random-access (open+1000 lookups) | 485ms | 483ms | 486ms | 486ms | 0.0 |
| Random-access (open+1000 lookups, lazy) | 479ms | 476ms | 481ms | 481ms | 21.4 |
| Random-access (lookup-only, 1000 cells) | 13ms | 13ms | 13ms | 13ms | 21.4 |

## Mixed Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Mixed workload write (ERP-style) | 16ms | 16ms | 16ms | 16ms | 0.0 |
