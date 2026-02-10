# SheetKit Rust Native Benchmark

Benchmark run: 2026-02-10T14:04:14Z

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
| Read Large Data (50k rows x 20 cols) | 616ms | 607ms | 632ms | 632ms | 68.7 |
| Read Heavy Styles (5k rows, formatted) | 33ms | 33ms | 33ms | 33ms | 0.7 |
| Read Multi-Sheet (10 sheets x 5k rows) | 360ms | 356ms | 363ms | 363ms | 19.1 |
| Read Formulas (10k rows) | 40ms | 39ms | 40ms | 40ms | 2.7 |
| Read Strings (20k rows text-heavy) | 140ms | 131ms | 191ms | 191ms | 3.5 |
| Read Data Validation (5k rows, 8 rules) | 25ms | 25ms | 25ms | 25ms | 0.0 |
| Read Comments (2k rows with comments) | 10ms | 10ms | 11ms | 11ms | 0.1 |
| Read Merged Cells (500 regions) | 2ms | 2ms | 2ms | 2ms | 0.0 |
| Read Mixed Workload (ERP document) | 34ms | 33ms | 34ms | 34ms | 0.0 |

## Read (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Read Scale 1k rows | 6ms | 6ms | 7ms | 7ms | 0.0 |
| Read Scale 10k rows | 62ms | 61ms | 62ms | 62ms | 0.9 |
| Read Scale 100k rows | 659ms | 630ms | 968ms | 968ms | 41.5 |

## Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 50000 rows x 20 cols | 1.03s | 727ms | 1.97s | 1.97s | 38.7 |
| Write 5000 styled rows | 39ms | 39ms | 39ms | 39ms | 0.0 |
| Write 10 sheets x 5000 rows | 377ms | 372ms | 405ms | 405ms | 0.1 |
| Write 10000 rows with formulas | 35ms | 33ms | 36ms | 36ms | 0.0 |
| Write 20000 text-heavy rows | 145ms | 145ms | 147ms | 147ms | 0.0 |

## Write (DV)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 5000 rows + 8 validation rules | 16ms | 14ms | 19ms | 19ms | 0.0 |

## Write (Comments)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 2000 rows with comments | 14ms | 14ms | 15ms | 15ms | 0.0 |

## Write (Merge)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 500 merged regions | 16ms | 14ms | 19ms | 19ms | 0.0 |

## Write (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 1k rows x 10 cols | 7ms | 6ms | 8ms | 8ms | 0.0 |
| Write 10k rows x 10 cols | 68ms | 67ms | 68ms | 68ms | 0.0 |
| Write 50k rows x 10 cols | 456ms | 415ms | 508ms | 508ms | 0.1 |
| Write 100k rows x 10 cols | 735ms | 721ms | 832ms | 832ms | 8.1 |

## Round-Trip

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Buffer round-trip (10000 rows) | 165ms | 165ms | 171ms | 171ms | 0.0 |

## Streaming

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Streaming write (50000 rows) | 555ms | 548ms | 570ms | 570ms | 24.8 |

## Random Access

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Random-access read (1000 cells from 50k-row file) | 592ms | 586ms | 601ms | 601ms | 0.0 |

## Mixed Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Mixed workload write (ERP-style) | 22ms | 21ms | 22ms | 22ms | 0.0 |
