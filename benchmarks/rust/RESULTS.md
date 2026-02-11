# SheetKit Rust Native Benchmark

Benchmark run: 2026-02-11T04:15:49Z

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
| Read Large Data (50k rows x 20 cols) | 387ms | 384ms | 403ms | 403ms | 68.6 |
| Read Heavy Styles (5k rows, formatted) | 20ms | 20ms | 20ms | 20ms | 0.1 |
| Read Multi-Sheet (10 sheets x 5k rows) | 223ms | 222ms | 224ms | 224ms | 17.3 |
| Read Formulas (10k rows) | 24ms | 24ms | 24ms | 24ms | 2.7 |
| Read Strings (20k rows text-heavy) | 83ms | 82ms | 83ms | 83ms | 3.5 |
| Read Data Validation (5k rows, 8 rules) | 16ms | 15ms | 16ms | 16ms | 0.0 |
| Read Comments (2k rows with comments) | 6ms | 6ms | 6ms | 6ms | 0.0 |
| Read Merged Cells (500 regions) | 1ms | 1ms | 1ms | 1ms | 0.0 |
| Read Mixed Workload (ERP document) | 21ms | 21ms | 21ms | 21ms | 0.0 |

## Read (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Read Scale 1k rows | 4ms | 4ms | 4ms | 4ms | 0.0 |
| Read Scale 10k rows | 39ms | 39ms | 39ms | 39ms | 0.0 |
| Read Scale 100k rows | 410ms | 401ms | 418ms | 418ms | 41.5 |

## Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 50000 rows x 20 cols | 478ms | 473ms | 478ms | 478ms | 38.7 |
| Write 5000 styled rows | 25ms | 25ms | 25ms | 25ms | 0.1 |
| Write 10 sheets x 5000 rows | 246ms | 245ms | 246ms | 246ms | 0.1 |
| Write 10000 rows with formulas | 21ms | 21ms | 21ms | 21ms | 0.0 |
| Write 20000 text-heavy rows | 92ms | 91ms | 94ms | 94ms | 0.0 |

## Write (DV)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 5000 rows + 8 validation rules | 8ms | 8ms | 9ms | 9ms | 0.0 |

## Write (Comments)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 2000 rows with comments | 6ms | 6ms | 7ms | 7ms | 0.0 |

## Write (Merge)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 500 merged regions | 9ms | 9ms | 9ms | 9ms | 0.0 |

## Write (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 1k rows x 10 cols | 4ms | 4ms | 5ms | 5ms | 0.0 |
| Write 10k rows x 10 cols | 48ms | 43ms | 52ms | 52ms | 1.1 |
| Write 50k rows x 10 cols | 226ms | 223ms | 228ms | 228ms | 0.0 |
| Write 100k rows x 10 cols | 454ms | 452ms | 457ms | 457ms | 0.0 |

## Round-Trip

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Buffer round-trip (10000 rows) | 106ms | 103ms | 109ms | 109ms | 0.0 |

## Streaming

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Streaming write (50000 rows) | 186ms | 182ms | 190ms | 190ms | 0.0 |

## Random Access

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Random-access read (1000 cells from 50k-row file) | 412ms | 411ms | 427ms | 427ms | 0.0 |

## Mixed Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Mixed workload write (ERP-style) | 14ms | 14ms | 16ms | 16ms | 0.0 |
