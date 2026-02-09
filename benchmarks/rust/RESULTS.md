# SheetKit Rust Native Benchmark

Benchmark run: 2026-02-09T13:11:58Z
Platform: macos aarch64
Profile: release
Iterations: 5 runs per scenario, 1 warmup

## Read

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Read Large Data (50k rows x 20 cols) | 636ms | 630ms | 680ms | 680ms | 52.7 |
| Read Heavy Styles (5k rows, formatted) | 34ms | 34ms | 37ms | 37ms | 3.7 |
| Read Multi-Sheet (10 sheets x 5k rows) | 379ms | 368ms | 382ms | 382ms | 24.8 |
| Read Formulas (10k rows) | 40ms | 40ms | 40ms | 40ms | 0.0 |
| Read Strings (20k rows text-heavy) | 138ms | 136ms | 146ms | 146ms | 3.5 |
| Read Data Validation (5k rows, 8 rules) | 25ms | 25ms | 25ms | 25ms | 0.0 |
| Read Comments (2k rows with comments) | 10ms | 10ms | 10ms | 10ms | 0.0 |
| Read Merged Cells (500 regions) | 2ms | 2ms | 2ms | 2ms | 0.0 |
| Read Mixed Workload (ERP document) | 34ms | 34ms | 34ms | 34ms | 0.0 |

## Read (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Read Scale 1k rows | 6ms | 6ms | 6ms | 6ms | 0.0 |
| Read Scale 10k rows | 63ms | 63ms | 63ms | 63ms | 0.0 |
| Read Scale 100k rows | 642ms | 629ms | 659ms | 659ms | 33.1 |

## Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 50000 rows x 20 cols | 702ms | 690ms | 960ms | 960ms | 127.5 |
| Write 5000 styled rows | 40ms | 40ms | 49ms | 49ms | 0.0 |
| Write 10 sheets x 5000 rows | 380ms | 376ms | 382ms | 382ms | 0.7 |
| Write 10000 rows with formulas | 33ms | 33ms | 33ms | 33ms | 0.0 |
| Write 20000 text-heavy rows | 152ms | 151ms | 153ms | 153ms | 0.0 |

## Write (DV)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 5000 rows + 8 validation rules | 13ms | 13ms | 13ms | 13ms | 0.0 |

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
| Write 1k rows x 10 cols | 7ms | 7ms | 7ms | 7ms | 0.0 |
| Write 10k rows x 10 cols | 68ms | 68ms | 71ms | 71ms | 0.0 |
| Write 50k rows x 10 cols | 356ms | 352ms | 363ms | 363ms | 0.0 |
| Write 100k rows x 10 cols | 709ms | 685ms | 717ms | 717ms | 66.0 |

## Round-Trip

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Buffer round-trip (10000 rows) | 163ms | 163ms | 165ms | 165ms | 0.0 |

## Streaming

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Streaming write (50000 rows) | 1.00s | 962ms | 1.01s | 1.01s | 26.3 |

## Random Access

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Random-access read (1000 cells from 50k-row file) | 593ms | 569ms | 609ms | 609ms | 0.0 |

## Mixed Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Mixed workload write (ERP-style) | 23ms | 23ms | 25ms | 25ms | 0.0 |
