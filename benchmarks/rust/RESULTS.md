# SheetKit Rust Native Benchmark

Benchmark run: 2026-02-10T00:41:38Z

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
| Read Large Data (50k rows x 20 cols) | 606ms | 604ms | 628ms | 628ms | 37.6 |
| Read Heavy Styles (5k rows, formatted) | 32ms | 32ms | 33ms | 33ms | 0.0 |
| Read Multi-Sheet (10 sheets x 5k rows) | 350ms | 349ms | 351ms | 351ms | 20.1 |
| Read Formulas (10k rows) | 38ms | 37ms | 38ms | 38ms | 0.0 |
| Read Strings (20k rows text-heavy) | 132ms | 132ms | 134ms | 134ms | 3.5 |
| Read Data Validation (5k rows, 8 rules) | 25ms | 24ms | 25ms | 25ms | 0.0 |
| Read Comments (2k rows with comments) | 10ms | 10ms | 10ms | 10ms | 0.0 |
| Read Merged Cells (500 regions) | 2ms | 2ms | 2ms | 2ms | 0.0 |
| Read Mixed Workload (ERP document) | 33ms | 33ms | 33ms | 33ms | 0.0 |

## Read (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Read Scale 1k rows | 6ms | 6ms | 6ms | 6ms | 0.0 |
| Read Scale 10k rows | 62ms | 62ms | 63ms | 63ms | 0.0 |
| Read Scale 100k rows | 625ms | 621ms | 639ms | 639ms | 41.5 |

## Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 50000 rows x 20 cols | 694ms | 689ms | 700ms | 700ms | 40.4 |
| Write 5000 styled rows | 38ms | 38ms | 38ms | 38ms | 0.0 |
| Write 10 sheets x 5000 rows | 356ms | 355ms | 357ms | 357ms | 1.8 |
| Write 10000 rows with formulas | 32ms | 31ms | 32ms | 32ms | 0.0 |
| Write 20000 text-heavy rows | 143ms | 142ms | 143ms | 143ms | 1.6 |

## Write (DV)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 5000 rows + 8 validation rules | 12ms | 12ms | 12ms | 12ms | 0.0 |

## Write (Comments)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 2000 rows with comments | 9ms | 9ms | 9ms | 9ms | 0.0 |

## Write (Merge)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 500 merged regions | 13ms | 13ms | 13ms | 13ms | 0.0 |

## Write (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 1k rows x 10 cols | 6ms | 6ms | 6ms | 6ms | 0.0 |
| Write 10k rows x 10 cols | 67ms | 67ms | 69ms | 69ms | 0.2 |
| Write 50k rows x 10 cols | 344ms | 342ms | 346ms | 346ms | 0.0 |
| Write 100k rows x 10 cols | 707ms | 705ms | 722ms | 722ms | 0.0 |

## Round-Trip

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Buffer round-trip (10000 rows) | 162ms | 162ms | 164ms | 164ms | 0.0 |

## Streaming

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Streaming write (50000 rows) | 535ms | 531ms | 537ms | 537ms | 0.0 |

## Random Access

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Random-access read (1000 cells from 50k-row file) | 583ms | 581ms | 586ms | 586ms | 0.0 |

## Mixed Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Mixed workload write (ERP-style) | 21ms | 21ms | 22ms | 22ms | 0.0 |
