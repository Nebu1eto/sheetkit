# SheetKit Rust Native Benchmark

Benchmark run: 2026-02-12T13:16:13Z

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
| Read Large Data (50k rows x 20 cols) | 509ms | 508ms | 519ms | 519ms | 17.6 |
| Read Large Data (50k rows x 20 cols) (lazy) | 518ms | 510ms | 520ms | 520ms | 17.2 |
| Read Heavy Styles (5k rows, formatted) | 27ms | 26ms | 27ms | 27ms | 1.7 |
| Read Heavy Styles (5k rows, formatted) (lazy) | 27ms | 26ms | 27ms | 27ms | 0.0 |
| Read Multi-Sheet (10 sheets x 5k rows) | 298ms | 296ms | 300ms | 300ms | 22.8 |
| Read Multi-Sheet (10 sheets x 5k rows) (lazy) | 301ms | 297ms | 303ms | 303ms | 17.6 |
| Read Formulas (10k rows) | 33ms | 33ms | 33ms | 33ms | 0.0 |
| Read Formulas (10k rows) (lazy) | 33ms | 33ms | 33ms | 33ms | 0.0 |
| Read Strings (20k rows text-heavy) | 109ms | 109ms | 110ms | 110ms | 3.5 |
| Read Strings (20k rows text-heavy) (lazy) | 109ms | 109ms | 110ms | 110ms | 0.0 |
| Read Data Validation (5k rows, 8 rules) | 21ms | 21ms | 21ms | 21ms | 0.0 |
| Read Data Validation (5k rows, 8 rules) (lazy) | 21ms | 20ms | 21ms | 21ms | 0.0 |
| Read Comments (2k rows with comments) | 7ms | 7ms | 7ms | 7ms | 0.0 |
| Read Comments (2k rows with comments) (lazy) | 7ms | 7ms | 7ms | 7ms | 0.0 |
| Read Merged Cells (500 regions) | 1ms | 1ms | 1ms | 1ms | 0.0 |
| Read Merged Cells (500 regions) (lazy) | 1ms | 1ms | 1ms | 1ms | 0.0 |
| Read Mixed Workload (ERP document) | 26ms | 26ms | 26ms | 26ms | 0.0 |
| Read Mixed Workload (ERP document) (lazy) | 26ms | 26ms | 27ms | 27ms | 0.0 |

## Read (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Read Scale 1k rows | 5ms | 5ms | 5ms | 5ms | 0.0 |
| Read Scale 1k rows (lazy) | 5ms | 5ms | 5ms | 5ms | 0.0 |
| Read Scale 10k rows | 51ms | 51ms | 52ms | 52ms | 0.0 |
| Read Scale 10k rows (lazy) | 51ms | 51ms | 52ms | 52ms | 0.0 |
| Read Scale 100k rows | 526ms | 521ms | 531ms | 531ms | 0.1 |
| Read Scale 100k rows (lazy) | 539ms | 522ms | 543ms | 543ms | 0.0 |

## Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 50000 rows x 20 cols | 544ms | 537ms | 555ms | 555ms | 12.8 |
| Write 5000 styled rows | 28ms | 28ms | 29ms | 29ms | 0.1 |
| Write 10 sheets x 5000 rows | 274ms | 271ms | 276ms | 276ms | 1.6 |
| Write 10000 rows with formulas | 24ms | 24ms | 25ms | 25ms | 0.0 |
| Write 20000 text-heavy rows | 108ms | 106ms | 108ms | 108ms | 1.6 |

## Write (DV)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 5000 rows + 8 validation rules | 9ms | 9ms | 9ms | 9ms | 0.0 |

## Write (Comments)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 2000 rows with comments | 7ms | 7ms | 7ms | 7ms | 0.0 |

## Write (Merge)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 500 merged regions | 1ms | 1ms | 2ms | 2ms | 0.0 |

## Write (Scale)

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Write 1k rows x 10 cols | 5ms | 5ms | 5ms | 5ms | 0.0 |
| Write 10k rows x 10 cols | 48ms | 48ms | 49ms | 49ms | 0.0 |
| Write 50k rows x 10 cols | 258ms | 256ms | 260ms | 260ms | 0.0 |
| Write 100k rows x 10 cols | 531ms | 518ms | 543ms | 543ms | 33.3 |

## Round-Trip

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Buffer round-trip (10000 rows) | 128ms | 128ms | 128ms | 128ms | 0.0 |

## Streaming

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Streaming write (50000 rows) | 207ms | 205ms | 210ms | 210ms | 0.0 |

## Random Access

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Random-access (open+1000 lookups) | 489ms | 488ms | 491ms | 491ms | 0.0 |
| Random-access (open+1000 lookups, lazy) | 488ms | 485ms | 491ms | 491ms | 6.5 |
| Random-access (lookup-only, 1000 cells) | 366ms | 363ms | 374ms | 374ms | 4.2 |

## Mixed Write

| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |
|----------|--------|-----|-----|-----|---------------|
| Mixed workload write (ERP-style) | 16ms | 16ms | 16ms | 16ms | 0.0 |
