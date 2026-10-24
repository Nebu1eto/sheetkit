# SheetKit Rust Native Benchmark

Benchmark run: 2026-10-24T06:06:40Z

## Environment

| Item       | Value                               |
| ---------- | ----------------------------------- |
| CPU        | Apple M4 Pro                        |
| RAM        | 24 GB                               |
| OS         | macos aarch64                       |
| Rust       | rustc 1.99.1 (58cf8f9c3 2026-10-13) |
| Profile    | release (opt-level=3, LTO=fat)      |
| Iterations | 5 runs per scenario, 1 warmup       |

## Read

| Scenario                                       | Median | Min   | Max   | P95   | Peak Mem (MB) |
| ---------------------------------------------- | ------ | ----- | ----- | ----- | ------------- |
| Read Large Data (50k rows x 20 cols)           | 450ms  | 447ms | 458ms | 458ms | 25.9          |
| Read Large Data (50k rows x 20 cols) (lazy)    | 431ms  | 429ms | 433ms | 433ms | 24.2          |
| Read Heavy Styles (5k rows, formatted)         | 24ms   | 24ms  | 24ms  | 24ms  | 0.0           |
| Read Heavy Styles (5k rows, formatted) (lazy)  | 23ms   | 22ms  | 23ms  | 23ms  | 0.0           |
| Read Multi-Sheet (10 sheets x 5k rows)         | 262ms  | 260ms | 265ms | 265ms | 26.7          |
| Read Multi-Sheet (10 sheets x 5k rows) (lazy)  | 253ms  | 251ms | 254ms | 254ms | 26.7          |
| Read Formulas (10k rows)                       | 27ms   | 27ms  | 28ms  | 28ms  | 0.0           |
| Read Formulas (10k rows) (lazy)                | 26ms   | 26ms  | 27ms  | 27ms  | 0.0           |
| Read Strings (20k rows text-heavy)             | 97ms   | 96ms  | 97ms  | 97ms  | 3.5           |
| Read Strings (20k rows text-heavy) (lazy)      | 94ms   | 93ms  | 96ms  | 96ms  | 3.5           |
| Read Data Validation (5k rows, 8 rules)        | 19ms   | 19ms  | 19ms  | 19ms  | 0.0           |
| Read Data Validation (5k rows, 8 rules) (lazy) | 18ms   | 18ms  | 18ms  | 18ms  | 0.0           |
| Read Comments (2k rows with comments)          | 6ms    | 6ms   | 6ms   | 6ms   | 0.0           |
| Read Comments (2k rows with comments) (lazy)   | 6ms    | 6ms   | 6ms   | 6ms   | 0.0           |
| Read Merged Cells (500 regions)                | 1ms    | 1ms   | 1ms   | 1ms   | 0.0           |
| Read Merged Cells (500 regions) (lazy)         | 1ms    | 1ms   | 1ms   | 1ms   | 0.0           |
| Read Mixed Workload (ERP document)             | 24ms   | 23ms  | 24ms  | 24ms  | 0.0           |
| Read Mixed Workload (ERP document) (lazy)      | 23ms   | 23ms  | 23ms  | 23ms  | 0.0           |

## Read (Scale)

| Scenario                    | Median | Min   | Max   | P95   | Peak Mem (MB) |
| --------------------------- | ------ | ----- | ----- | ----- | ------------- |
| Read Scale 1k rows          | 4ms    | 4ms   | 4ms   | 4ms   | 0.0           |
| Read Scale 1k rows (lazy)   | 4ms    | 4ms   | 4ms   | 4ms   | 0.0           |
| Read Scale 10k rows         | 45ms   | 45ms  | 46ms  | 46ms  | 0.0           |
| Read Scale 10k rows (lazy)  | 44ms   | 43ms  | 44ms  | 44ms  | 0.0           |
| Read Scale 100k rows        | 457ms  | 455ms | 478ms | 478ms | 39.0          |
| Read Scale 100k rows (lazy) | 444ms  | 442ms | 446ms | 446ms | 35.9          |

## Write

| Scenario                       | Median | Min   | Max   | P95   | Peak Mem (MB) |
| ------------------------------ | ------ | ----- | ----- | ----- | ------------- |
| Write 50000 rows x 20 cols     | 461ms  | 459ms | 471ms | 471ms | 78.9          |
| Write 5000 styled rows         | 24ms   | 23ms  | 24ms  | 24ms  | 0.0           |
| Write 10 sheets x 5000 rows    | 248ms  | 243ms | 248ms | 248ms | 16.2          |
| Write 10000 rows with formulas | 21ms   | 21ms  | 21ms  | 21ms  | 0.0           |
| Write 20000 text-heavy rows    | 92ms   | 91ms  | 93ms  | 93ms  | 0.4           |

## Write (DV)

| Scenario                             | Median | Min | Max | P95 | Peak Mem (MB) |
| ------------------------------------ | ------ | --- | --- | --- | ------------- |
| Write 5000 rows + 8 validation rules | 8ms    | 8ms | 8ms | 8ms | 0.0           |

## Write (Comments)

| Scenario                      | Median | Min | Max | P95 | Peak Mem (MB) |
| ----------------------------- | ------ | --- | --- | --- | ------------- |
| Write 2000 rows with comments | 6ms    | 6ms | 6ms | 6ms | 0.0           |

## Write (Merge)

| Scenario                 | Median | Min | Max | P95 | Peak Mem (MB) |
| ------------------------ | ------ | --- | --- | --- | ------------- |
| Write 500 merged regions | 1ms    | 1ms | 1ms | 1ms | 0.0           |

## Write (Scale)

| Scenario                  | Median | Min   | Max   | P95   | Peak Mem (MB) |
| ------------------------- | ------ | ----- | ----- | ----- | ------------- |
| Write 1k rows x 10 cols   | 4ms    | 4ms   | 4ms   | 4ms   | 0.0           |
| Write 10k rows x 10 cols  | 43ms   | 42ms  | 43ms  | 43ms  | 1.0           |
| Write 50k rows x 10 cols  | 224ms  | 222ms | 225ms | 225ms | 13.2          |
| Write 100k rows x 10 cols | 461ms  | 453ms | 477ms | 477ms | 35.4          |

## Round-Trip

| Scenario                       | Median | Min   | Max   | P95   | Peak Mem (MB) |
| ------------------------------ | ------ | ----- | ----- | ----- | ------------- |
| Buffer round-trip (10000 rows) | 108ms  | 107ms | 108ms | 108ms | 0.0           |

## Streaming

| Scenario                     | Median | Min   | Max   | P95   | Peak Mem (MB) |
| ---------------------------- | ------ | ----- | ----- | ----- | ------------- |
| Streaming write (50000 rows) | 145ms  | 144ms | 148ms | 148ms | 0.0           |

## Random Access

| Scenario                                | Median | Min   | Max   | P95   | Peak Mem (MB) |
| --------------------------------------- | ------ | ----- | ----- | ----- | ------------- |
| Random-access (open+1000 lookups)       | 416ms  | 413ms | 420ms | 420ms | 10.7          |
| Random-access (open+1000 lookups, lazy) | 418ms  | 411ms | 424ms | 424ms | 7.9           |
| Random-access (lookup-only, 1000 cells) | 344ms  | 343ms | 352ms | 352ms | 24.2          |

## Mixed Write

| Scenario                         | Median | Min  | Max  | P95  | Peak Mem (MB) |
| -------------------------------- | ------ | ---- | ---- | ---- | ------------- |
| Mixed workload write (ERP-style) | 14ms   | 13ms | 14ms | 14ms | 0.0           |
