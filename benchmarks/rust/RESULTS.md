# SheetKit Rust Native Benchmark

Benchmark run: 2026-02-12T11:08:17Z

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
| Read Large Data (50k rows x 20 cols) | 518ms | 509ms | 526ms | 526ms | 60.2 |
| Read Large Data (50k rows x 20 cols) (lazy) | 506ms | 501ms | 514ms | 514ms | 17.2 |
| Read Heavy Styles (5k rows, formatted) | 27ms | 27ms | 27ms | 27ms | 3.2 |
| Read Heavy Styles (5k rows, formatted) (lazy) | 27ms | 26ms | 27ms | 27ms | 0.8 |
| Read Multi-Sheet (10 sheets x 5k rows) | 301ms | 298ms | 306ms | 306ms | 19.0 |
| Read Multi-Sheet (10 sheets x 5k rows) (lazy) | 295ms | 291ms | 300ms | 300ms | 24.7 |
| Read Formulas (10k rows) | 33ms | 33ms | 34ms | 34ms | 2.7 |
| Read Formulas (10k rows) (lazy) | 32ms | 31ms | 32ms | 32ms | 0.0 |
| Read Strings (20k rows text-heavy) | 108ms | 107ms | 109ms | 109ms | 0.0 |
| Read Strings (20k rows text-heavy) (lazy) | 107ms | 107ms | 108ms | 108ms | 0.0 |
| Read Data Validation (5k rows, 8 rules) | 21ms | 21ms | 21ms | 21ms | 0.0 |
| Read Data Validation (5k rows, 8 rules) (lazy) | 20ms | 20ms | 20ms | 20ms | 0.0 |
| Read Comments (2k rows with comments) | 8ms | 8ms | 8ms | 8ms | 0.0 |
| Read Comments (2k rows with comments) (lazy) | 7ms | 6ms | 7ms | 7ms | 0.0 |
| Read Merged Cells (500 regions) | 1ms | 1ms | 1ms | 1ms | 0.0 |
| Read Merged Cells (500 regions) (lazy) | 1ms | 1ms | 1ms | 1ms | 0.0 |
| Read Mixed Workload (ERP document) | 27ms | 27ms | 27ms | 27ms | 0.0 |
| Read Mixed Workload (ERP document) (lazy) | 26ms | 26ms | 26ms | 26ms | 0.0 |

## Read (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Read Scale 1k rows | 5ms | 5ms | 5ms | 5ms | 0.0 |
| Read Scale 1k rows (lazy) | 5ms | 5ms | 5ms | 5ms | 0.0 |
| Read Scale 10k rows | 51ms | 50ms | 51ms | 51ms | 0.0 |
| Read Scale 10k rows (lazy) | 50ms | 50ms | 52ms | 52ms | 0.0 |
| Read Scale 100k rows | 530ms | 522ms | 537ms | 537ms | 33.0 |
| Read Scale 100k rows (lazy) | 524ms | 508ms | 531ms | 531ms | 20.6 |

## Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 50000 rows x 20 cols | 503ms | 502ms | 507ms | 507ms | 31.5 |
| Write 5000 styled rows | 28ms | 27ms | 28ms | 28ms | 0.0 |
| Write 10 sheets x 5000 rows | 258ms | 257ms | 259ms | 259ms | 0.1 |
| Write 10000 rows with formulas | 24ms | 23ms | 24ms | 24ms | 0.0 |
| Write 20000 text-heavy rows | 107ms | 104ms | 112ms | 112ms | 0.0 |

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
| Write 1k rows x 10 cols | 5ms | 4ms | 5ms | 5ms | 0.0 |
| Write 10k rows x 10 cols | 47ms | 47ms | 48ms | 48ms | 0.0 |
| Write 50k rows x 10 cols | 247ms | 246ms | 253ms | 253ms | 0.0 |
| Write 100k rows x 10 cols | 518ms | 514ms | 519ms | 519ms | 0.0 |

## Round-Trip

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Buffer round-trip (10000 rows) | 122ms | 122ms | 123ms | 123ms | 0.0 |

## Streaming

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Streaming write (50000 rows) | 198ms | 195ms | 199ms | 199ms | 0.0 |

## Random Access

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Random-access (open+1000 lookups) | 484ms | 474ms | 491ms | 491ms | 2.6 |
| Random-access (open+1000 lookups, lazy) | 482ms | 477ms | 487ms | 487ms | 21.4 |
| Random-access (lookup-only, 1000 cells) | 12ms | 12ms | 13ms | 13ms | 21.4 |

## Mixed Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Mixed workload write (ERP-style) | 15ms | 15ms | 15ms | 15ms | 0.0 |
