# SheetKit Rust Native Benchmark

Benchmark run: 2026-02-09T13:29:12Z

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
| Read Large Data (50k rows x 20 cols) | 625ms | 619ms | 668ms | 668ms | 52.6 |
| Read Heavy Styles (5k rows, formatted) | 33ms | 33ms | 34ms | 34ms | 0.1 |
| Read Multi-Sheet (10 sheets x 5k rows) | 369ms | 356ms | 382ms | 382ms | 25.3 |
| Read Formulas (10k rows) | 43ms | 39ms | 50ms | 50ms | 0.0 |
| Read Strings (20k rows text-heavy) | 137ms | 134ms | 141ms | 141ms | 3.5 |
| Read Data Validation (5k rows, 8 rules) | 26ms | 26ms | 26ms | 26ms | 0.7 |
| Read Comments (2k rows with comments) | 10ms | 10ms | 10ms | 10ms | 0.0 |
| Read Merged Cells (500 regions) | 2ms | 2ms | 2ms | 2ms | 0.1 |
| Read Mixed Workload (ERP document) | 34ms | 34ms | 34ms | 34ms | 0.0 |

## Read (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Read Scale 1k rows | 6ms | 6ms | 7ms | 7ms | 0.0 |
| Read Scale 10k rows | 65ms | 64ms | 67ms | 67ms | 0.0 |
| Read Scale 100k rows | 650ms | 644ms | 651ms | 651ms | 33.1 |

## Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 50000 rows x 20 cols | 742ms | 722ms | 759ms | 759ms | 126.7 |
| Write 5000 styled rows | 41ms | 41ms | 42ms | 42ms | 0.0 |
| Write 10 sheets x 5000 rows | 381ms | 373ms | 392ms | 392ms | 0.1 |
| Write 10000 rows with formulas | 34ms | 33ms | 34ms | 34ms | 0.0 |
| Write 20000 text-heavy rows | 148ms | 147ms | 158ms | 158ms | 0.0 |

## Write (DV)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 5000 rows + 8 validation rules | 12ms | 12ms | 12ms | 12ms | 0.0 |

## Write (Comments)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 2000 rows with comments | 9ms | 9ms | 10ms | 10ms | 0.0 |

## Write (Merge)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 500 merged regions | 13ms | 13ms | 14ms | 14ms | 0.0 |

## Write (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 1k rows x 10 cols | 6ms | 6ms | 7ms | 7ms | 0.0 |
| Write 10k rows x 10 cols | 69ms | 67ms | 69ms | 69ms | 0.0 |
| Write 50k rows x 10 cols | 347ms | 344ms | 353ms | 353ms | 0.0 |
| Write 100k rows x 10 cols | 705ms | 695ms | 730ms | 730ms | 66.0 |

## Round-Trip

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Buffer round-trip (10000 rows) | 169ms | 164ms | 175ms | 175ms | 0.0 |

## Streaming

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Streaming write (50000 rows) | 1.02s | 994ms | 1.05s | 1.05s | 23.8 |

## Random Access

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Random-access read (1000 cells from 50k-row file) | 585ms | 576ms | 598ms | 598ms | 0.0 |

## Mixed Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Mixed workload write (ERP-style) | 23ms | 23ms | 23ms | 23ms | 0.0 |
