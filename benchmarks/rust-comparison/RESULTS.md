# Rust Excel Library Comparison Benchmark

Benchmark run: 2026-10-24T06:11:34Z

## Libraries

| Library             | Description                          | Capability            |
| ------------------- | ------------------------------------ | --------------------- |
| **SheetKit**        | Rust Excel library (this project)    | Read + Write + Modify |
| **calamine**        | Fast Excel/ODS reader                | Read-only             |
| **rust_xlsxwriter** | Excel writer (port of libxlsxwriter) | Write-only            |
| **edit-xlsx**       | Excel read/modify/write library      | Read + Write + Modify |

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

| Scenario                               | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner   |
| -------------------------------------- | -------- | -------- | --------- | --------------- | ------------------- | -------- |
| Read Large Data (50k rows x 20 cols)   | 416ms    | 236ms    | 333ms*    | N/A             | N/A                 | calamine |
| Read Heavy Styles (5k rows, formatted) | 22ms     | 12ms     | 19ms*     | N/A             | N/A                 | calamine |
| Read Multi-Sheet (10 sheets x 5k rows) | 246ms    | 135ms    | 182ms*    | N/A             | N/A                 | calamine |
| Read Formulas (10k rows)               | 26ms*    | 13ms*    | 21ms*     | N/A             | N/A                 | N/A      |
| Read Strings (20k rows text-heavy)     | 89ms     | 49ms     | 72ms*     | N/A             | N/A                 | calamine |

## Read (Scale)

| Scenario             | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner   |
| -------------------- | -------- | -------- | --------- | --------------- | ------------------- | -------- |
| Read Scale 1k rows   | 4ms      | 2ms      | 3ms*      | N/A             | N/A                 | calamine |
| Read Scale 10k rows  | 41ms     | 23ms     | 33ms*     | N/A             | N/A                 | calamine |
| Read Scale 100k rows | 444ms    | 239ms    | 338ms*    | N/A             | N/A                 | calamine |

## Write

| Scenario                       | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner          |
| ------------------------------ | -------- | -------- | --------- | --------------- | ------------------- | --------------- |
| Write 50000 rows x 20 cols     | 472ms    | N/A      | 577ms     | 491ms           | N/A                 | SheetKit        |
| Write 5000 styled rows         | 25ms     | N/A      | 40ms      | 25ms            | N/A                 | SheetKit        |
| Write 10 sheets x 5000 rows    | 249ms    | N/A      | 282ms     | 213ms           | N/A                 | rust_xlsxwriter |
| Write 10000 rows with formulas | 23ms     | N/A      | 44ms      | 24ms            | N/A                 | SheetKit        |
| Write 20000 text-heavy rows    | 53ms     | N/A      | 61ms      | 48ms            | N/A                 | rust_xlsxwriter |
| Write 500 merged regions       | 2ms      | N/A      | 5ms       | 2ms             | N/A                 | SheetKit        |

## Write (Scale)

| Scenario                  | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner          |
| ------------------------- | -------- | -------- | --------- | --------------- | ------------------- | --------------- |
| Write 1k rows x 10 cols   | 4ms      | N/A      | 9ms       | 4ms             | N/A                 | rust_xlsxwriter |
| Write 10k rows x 10 cols  | 44ms     | N/A      | 58ms      | 44ms            | N/A                 | SheetKit        |
| Write 50k rows x 10 cols  | 230ms    | N/A      | 280ms     | 236ms           | N/A                 | SheetKit        |
| Write 100k rows x 10 cols | 470ms    | N/A      | 566ms     | 481ms           | N/A                 | SheetKit        |

## Round-Trip

| Scenario                       | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner              |
| ------------------------------ | -------- | -------- | --------- | --------------- | ------------------- | ------------------- |
| Buffer round-trip (10000 rows) | 81ms     | N/A      | N/A       | N/A             | 51ms                | xlsxwriter+calamine |

## Streaming

| Scenario                     | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner   |
| ---------------------------- | -------- | -------- | --------- | --------------- | ------------------- | -------- |
| Streaming write (50000 rows) | 161ms    | N/A      | N/A       | 495ms           | N/A                 | SheetKit |

## Random Access

| Scenario                        | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner   |
| ------------------------------- | -------- | -------- | --------- | --------------- | ------------------- | -------- |
| Random-access read (1000 cells) | 412ms    | 241ms    | N/A       | N/A             | N/A                 | calamine |

## Modify

| Scenario                          | SheetKit | calamine | edit-xlsx | rust_xlsxwriter | xlsxwriter+calamine | Winner    |
| --------------------------------- | -------- | -------- | --------- | --------------- | ------------------- | --------- |
| Modify 1000 cells in 50k-row file | 637ms    | N/A      | 439ms     | N/A             | N/A                 | edit-xlsx |
- `*` indicates workload-count mismatch, value-probe mismatch, or zero read counts; excluded from Winner selection as non-comparable.

## Detailed Statistics

| Scenario                               | Library             | Median | Min   | Max   | P95   | Peak Mem (MB) | Rows Read | Rows Expected | Cells Read | Cells Expected | Comparable |
| -------------------------------------- | ------------------- | ------ | ----- | ----- | ----- | ------------- | --------- | ------------- | ---------- | -------------- | ---------- |
| Read Large Data (50k rows x 20 cols)   | SheetKit            | 416ms  | 414ms | 418ms | 418ms | 24.3          | 50001     | 50001         | 1000020    | 1000020        | Yes        |
| Read Large Data (50k rows x 20 cols)   | calamine            | 236ms  | 235ms | 239ms | 239ms | 38.2          | 50001     | 50001         | 1000020    | 1000020        | Yes        |
| Read Large Data (50k rows x 20 cols)   | edit-xlsx           | 333ms  | 331ms | 336ms | 336ms | 51.1          | 0         | 50001         | 0          | 1000020        | No         |
| Read Heavy Styles (5k rows, formatted) | SheetKit            | 22ms   | 22ms  | 22ms  | 22ms  | 0.0           | 5001      | 5001          | 50010      | 50010          | Yes        |
| Read Heavy Styles (5k rows, formatted) | calamine            | 12ms   | 12ms  | 12ms  | 12ms  | 0.0           | 5001      | 5001          | 50010      | 50010          | Yes        |
| Read Heavy Styles (5k rows, formatted) | edit-xlsx           | 19ms   | 19ms  | 19ms  | 19ms  | 0.0           | 0         | 5001          | 0          | 50010          | No         |
| Read Multi-Sheet (10 sheets x 5k rows) | SheetKit            | 246ms  | 245ms | 247ms | 247ms | 24.4          | 50010     | 50010         | 500100     | 500100         | Yes        |
| Read Multi-Sheet (10 sheets x 5k rows) | calamine            | 135ms  | 134ms | 137ms | 137ms | 0.0           | 50010     | 50010         | 500100     | 500100         | Yes        |
| Read Multi-Sheet (10 sheets x 5k rows) | edit-xlsx           | 182ms  | 181ms | 183ms | 183ms | 18.0          | 0         | 50010         | 0          | 500100         | No         |
| Read Formulas (10k rows)               | SheetKit            | 26ms   | 26ms  | 26ms  | 26ms  | 0.0           | 10001     | 10001         | 70007      | N/A            | No         |
| Read Formulas (10k rows)               | calamine            | 13ms   | 13ms  | 13ms  | 13ms  | 0.0           | 10001     | 10001         | 20007      | N/A            | No         |
| Read Formulas (10k rows)               | edit-xlsx           | 21ms   | 21ms  | 22ms  | 22ms  | 0.0           | 0         | 10001         | 0          | N/A            | No         |
| Read Strings (20k rows text-heavy)     | SheetKit            | 89ms   | 88ms  | 89ms  | 89ms  | 3.5           | 20001     | 20001         | 200010     | 200010         | Yes        |
| Read Strings (20k rows text-heavy)     | calamine            | 49ms   | 48ms  | 49ms  | 49ms  | 4.2           | 20001     | 20001         | 200010     | 200010         | Yes        |
| Read Strings (20k rows text-heavy)     | edit-xlsx           | 72ms   | 72ms  | 73ms  | 73ms  | 7.0           | 0         | 20001         | 0          | 200010         | No         |
| Read Scale 1k rows                     | SheetKit            | 4ms    | 4ms   | 4ms   | 4ms   | 0.0           | 1001      | 1001          | 10010      | 10010          | Yes        |
| Read Scale 1k rows                     | calamine            | 2ms    | 2ms   | 2ms   | 2ms   | 0.0           | 1001      | 1001          | 10010      | 10010          | Yes        |
| Read Scale 1k rows                     | edit-xlsx           | 3ms    | 3ms   | 3ms   | 3ms   | 0.0           | 0         | 1001          | 0          | 10010          | No         |
| Read Scale 10k rows                    | SheetKit            | 41ms   | 41ms  | 41ms  | 41ms  | 0.0           | 10001     | 10001         | 100010     | 100010         | Yes        |
| Read Scale 10k rows                    | calamine            | 23ms   | 23ms  | 24ms  | 24ms  | 3.8           | 10001     | 10001         | 100010     | 100010         | Yes        |
| Read Scale 10k rows                    | edit-xlsx           | 33ms   | 32ms  | 34ms  | 34ms  | 0.0           | 0         | 10001         | 0          | 100010         | No         |
| Read Scale 100k rows                   | SheetKit            | 444ms  | 431ms | 456ms | 456ms | 11.4          | 100001    | 100001        | 1000010    | 1000010        | Yes        |
| Read Scale 100k rows                   | calamine            | 239ms  | 234ms | 240ms | 240ms | 31.7          | 100001    | 100001        | 1000010    | 1000010        | Yes        |
| Read Scale 100k rows                   | edit-xlsx           | 338ms  | 332ms | 340ms | 340ms | 35.7          | 0         | 100001        | 0          | 1000010        | No         |
| Write 50000 rows x 20 cols             | SheetKit            | 472ms  | 470ms | 477ms | 477ms | 103.0         | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 50000 rows x 20 cols             | rust_xlsxwriter     | 491ms  | 486ms | 492ms | 492ms | 44.2          | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 50000 rows x 20 cols             | edit-xlsx           | 577ms  | 574ms | 581ms | 581ms | 47.2          | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 5000 styled rows                 | SheetKit            | 25ms   | 25ms  | 25ms  | 25ms  | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 5000 styled rows                 | rust_xlsxwriter     | 25ms   | 25ms  | 25ms  | 25ms  | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 5000 styled rows                 | edit-xlsx           | 40ms   | 39ms  | 40ms  | 40ms  | 2.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 10 sheets x 5000 rows            | SheetKit            | 249ms  | 247ms | 250ms | 250ms | 26.0          | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 10 sheets x 5000 rows            | rust_xlsxwriter     | 213ms  | 213ms | 214ms | 214ms | 3.5           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 10 sheets x 5000 rows            | edit-xlsx           | 282ms  | 280ms | 283ms | 283ms | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 10000 rows with formulas         | SheetKit            | 23ms   | 23ms  | 23ms  | 23ms  | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 10000 rows with formulas         | rust_xlsxwriter     | 24ms   | 23ms  | 24ms  | 24ms  | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 10000 rows with formulas         | edit-xlsx           | 44ms   | 44ms  | 46ms  | 46ms  | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 20000 text-heavy rows            | SheetKit            | 53ms   | 52ms  | 54ms  | 54ms  | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 20000 text-heavy rows            | rust_xlsxwriter     | 48ms   | 48ms  | 49ms  | 49ms  | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 20000 text-heavy rows            | edit-xlsx           | 61ms   | 60ms  | 63ms  | 63ms  | 6.8           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 500 merged regions               | SheetKit            | 2ms    | 1ms   | 2ms   | 2ms   | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 500 merged regions               | rust_xlsxwriter     | 2ms    | 2ms   | 2ms   | 2ms   | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 500 merged regions               | edit-xlsx           | 5ms    | 5ms   | 6ms   | 6ms   | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 1k rows x 10 cols                | SheetKit            | 4ms    | 4ms   | 5ms   | 5ms   | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 1k rows x 10 cols                | rust_xlsxwriter     | 4ms    | 4ms   | 5ms   | 5ms   | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 1k rows x 10 cols                | edit-xlsx           | 9ms    | 8ms   | 9ms   | 9ms   | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 10k rows x 10 cols               | SheetKit            | 44ms   | 44ms  | 44ms  | 44ms  | 1.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 10k rows x 10 cols               | rust_xlsxwriter     | 44ms   | 44ms  | 45ms  | 45ms  | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 10k rows x 10 cols               | edit-xlsx           | 58ms   | 58ms  | 58ms  | 58ms  | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 50k rows x 10 cols               | SheetKit            | 230ms  | 229ms | 237ms | 237ms | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 50k rows x 10 cols               | rust_xlsxwriter     | 236ms  | 234ms | 239ms | 239ms | 12.7          | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 50k rows x 10 cols               | edit-xlsx           | 280ms  | 279ms | 281ms | 281ms | 14.4          | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 100k rows x 10 cols              | SheetKit            | 470ms  | 467ms | 477ms | 477ms | 62.7          | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 100k rows x 10 cols              | rust_xlsxwriter     | 481ms  | 475ms | 484ms | 484ms | 33.8          | N/A       | N/A           | N/A        | N/A            | Yes        |
| Write 100k rows x 10 cols              | edit-xlsx           | 566ms  | 564ms | 568ms | 568ms | 31.5          | N/A       | N/A           | N/A        | N/A            | Yes        |
| Buffer round-trip (10000 rows)         | SheetKit            | 81ms   | 81ms  | 82ms  | 82ms  | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Buffer round-trip (10000 rows)         | xlsxwriter+calamine | 51ms   | 50ms  | 51ms  | 51ms  | 0.8           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Streaming write (50000 rows)           | SheetKit            | 161ms  | 160ms | 169ms | 169ms | 0.0           | N/A       | N/A           | N/A        | N/A            | Yes        |
| Streaming write (50000 rows)           | rust_xlsxwriter     | 495ms  | 492ms | 498ms | 498ms | 33.3          | N/A       | N/A           | N/A        | N/A            | Yes        |
| Random-access read (1000 cells)        | SheetKit            | 412ms  | 407ms | 415ms | 415ms | 11.2          | N/A       | N/A           | 1000       | N/A            | Yes        |
| Random-access read (1000 cells)        | calamine            | 241ms  | 238ms | 244ms | 244ms | 34.0          | N/A       | N/A           | 1000       | N/A            | Yes        |
| Modify 1000 cells in 50k-row file      | SheetKit            | 637ms  | 635ms | 648ms | 648ms | 33.7          | N/A       | N/A           | N/A        | N/A            | Yes        |
| Modify 1000 cells in 50k-row file      | edit-xlsx           | 439ms  | 438ms | 441ms | 441ms | 29.1          | N/A       | N/A           | N/A        | N/A            | Yes        |

## Win Summary

| Library             | Wins |
| ------------------- | ---- |
| calamine            | 8/21 |
| SheetKit            | 8/21 |
| rust_xlsxwriter     | 3/21 |
| xlsxwriter+calamine | 1/21 |
| edit-xlsx           | 1/21 |
