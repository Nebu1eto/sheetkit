# SheetKit: Rust Native vs Node.js (napi-rs) Benchmark Comparison

Benchmark run: 2026-02-11

## Environment

| Item | Value |
|------|-------|
| CPU | Apple M4 Pro |
| RAM | 24 GB |
| OS | macOS arm64 (Apple Silicon) |
| Node.js | v25.3.0 |
| Rust | rustc 1.93.0 (254b59607 2026-01-19) |
| Methodology | 1 warmup + 5 measured runs per scenario, median reported |

## napi-rs Overhead Analysis

This comparison measures the overhead introduced by the napi-rs FFI layer when calling
SheetKit from Node.js vs using it directly from Rust. The Rust numbers are from the
Rust native benchmark (run separately); the Node.js numbers are from the latest
Node.js benchmark run.

### Read Benchmarks (sync)

| Scenario | Rust (median) | Node.js sync (median) | Overhead | Ratio |
|----------|--------------|----------------------|----------|-------|
| Large Data (50k rows x 20 cols) | 387ms | 454ms | +67ms | 1.17x |
| Heavy Styles (5k rows, formatted) | 20ms | 24ms | +4ms | 1.20x |
| Multi-Sheet (10 sheets x 5k rows) | 223ms | 530ms | +307ms | 2.38x |
| Formulas (10k rows) | 24ms | 33ms | +9ms | 1.38x |
| Strings (20k rows text-heavy) | 83ms | 95ms | +12ms | 1.14x |
| Data Validation (5k rows, 8 rules) | 16ms | 19ms | +3ms | 1.19x |
| Comments (2k rows with comments) | 6ms | 8ms | +2ms | 1.33x |
| Merged Cells (500 regions) | 1ms | 2ms | +1ms | 2.00x |
| Mixed Workload (ERP document) | 21ms | 27ms | +6ms | 1.29x |

### Read Benchmarks (async)

| Scenario | Rust (median) | Node.js async (median) | Overhead | Ratio |
|----------|--------------|------------------------|----------|-------|
| Large Data (50k rows x 20 cols) | 387ms | 435ms | +48ms | 1.12x |
| Heavy Styles (5k rows, formatted) | 20ms | 23ms | +3ms | 1.15x |
| Multi-Sheet (10 sheets x 5k rows) | 223ms | 525ms | +302ms | 2.35x |
| Formulas (10k rows) | 24ms | 32ms | +8ms | 1.33x |
| Strings (20k rows text-heavy) | 83ms | 95ms | +12ms | 1.14x |
| Data Validation (5k rows, 8 rules) | 16ms | 19ms | +3ms | 1.19x |
| Comments (2k rows with comments) | 6ms | 8ms | +2ms | 1.33x |
| Merged Cells (500 regions) | 1ms | 1ms | +0ms | 1.00x |
| Mixed Workload (ERP document) | 21ms | 27ms | +6ms | 1.29x |

### Read Scaling

| Scenario | Rust (median) | Node.js sync (median) | Node.js async (median) | Sync Ratio | Async Ratio |
|----------|--------------|----------------------|------------------------|------------|-------------|
| Scale 1k rows | 4ms | 5ms | 5ms | 1.25x | 1.25x |
| Scale 10k rows | 39ms | 45ms | 44ms | 1.15x | 1.13x |
| Scale 100k rows | 410ms | 474ms | 474ms | 1.16x | 1.16x |

### Write Benchmarks

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| 50k rows x 20 cols | 478ms | 461ms | -17ms | 0.96x |
| 5k styled rows | 25ms | 35ms | +10ms | 1.40x |
| 10 sheets x 5k rows | 246ms | 251ms | +5ms | 1.02x |
| 10k rows with formulas | 21ms | 28ms | +7ms | 1.33x |
| 20k text-heavy rows | 92ms | 87ms | -5ms | 0.95x |
| 5k rows + 8 DV rules | 8ms | 9ms | +1ms | 1.13x |
| 2k rows with comments | 6ms | 8ms | +2ms | 1.33x |
| 500 merged regions | 9ms | 10ms | +1ms | 1.11x |

### Write Scaling

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| 1k rows x 10 cols | 4ms | 5ms | +1ms | 1.25x |
| 10k rows x 10 cols | 48ms | 47ms | -1ms | 0.98x |
| 50k rows x 10 cols | 226ms | 235ms | +9ms | 1.04x |
| 100k rows x 10 cols | 454ms | 476ms | +22ms | 1.05x |

### Other

| Scenario | Rust (median) | Node.js sync (median) | Node.js async (median) | Sync Ratio | Async Ratio |
|----------|--------------|----------------------|------------------------|------------|-------------|
| Buffer round-trip (10k rows) | 106ms | 118ms | - | 1.11x | - |
| Streaming write (50k rows) | 186ms | 309ms | - | 1.66x | - |
| Random-access read (1k cells/50k file) | 412ms | 387ms | 382ms | 0.94x | 0.93x |
| Mixed workload write (ERP-style) | 14ms | 19ms | - | 1.36x | - |

## Node.js Memory Usage

RSS (Resident Set Size) measured for SheetKit in the Node.js benchmark.

| Scenario | Sync Memory | Async Memory |
|----------|-------------|--------------|
| Read Large Data (50k x 20) | 195.3MB | 17.2MB |
| Read Multi-Sheet (10 x 5k) | 112.7MB | 0.4MB |
| Read Scale 100k rows | 175.2MB | 15.9MB |
| Read Heavy Styles (5k) | 6.0MB | 0.0MB |
| Read Formulas (10k) | 9.3MB | 0.0MB |
| Read Strings (20k) | 2.5MB | 0.0MB |
| Write 50k x 20 | 67.3MB | - |
| Write 100k x 10 | 65.1MB | - |
| Streaming 50k rows | 0.0MB | - |
| Random-access read (1k cells) | 61.6MB | 0.0MB |

## Rust Ecosystem Comparison

SheetKit was benchmarked against other popular Rust Excel libraries. Each library targets different use cases: calamine is read-only, rust_xlsxwriter is write-only, and edit-xlsx supports read/modify/write.

### Read (Rust libraries)

| Scenario | SheetKit | calamine | edit-xlsx | Winner |
|----------|----------|----------|-----------|--------|
| Large Data (50k rows x 20 cols) | 494ms | 324ms | 372ms* | calamine |
| Heavy Styles (5k rows, formatted) | 26ms | 17ms | 21ms* | calamine |
| Multi-Sheet (10 sheets x 5k rows) | 289ms | 179ms | 199ms* | calamine |
| Formulas (10k rows) | 32ms* | 15ms* | 23ms* | N/A |
| Strings (20k rows text-heavy) | 105ms | 70ms | 81ms* | calamine |

### Write (Rust libraries)

| Scenario | SheetKit | rust_xlsxwriter | edit-xlsx | Winner |
|----------|----------|-----------------|-----------|--------|
| 50k rows x 20 cols | 475ms | 886ms | 939ms | SheetKit |
| 5k styled rows | 27ms | 38ms | 52ms | SheetKit |
| 10 sheets x 5k rows | 249ms | 338ms | 414ms | SheetKit |
| 10k rows with formulas | 23ms | 36ms | 59ms | SheetKit |
| 20k text-heavy rows | 57ms | 66ms | 71ms | SheetKit |
| 500 merged regions | 1ms | 2ms | 5ms | SheetKit |

### Other (Rust libraries)

| Scenario | SheetKit | Best alternative | Winner |
|----------|----------|-----------------|--------|
| Buffer round-trip (10k rows) | 118ms | 82ms (xlsxwriter+calamine) | xlsxwriter+calamine |
| Streaming write (50k rows) | 191ms | 885ms (rust_xlsxwriter) | SheetKit |
| Random-access read (1k cells) | 465ms | 321ms (calamine) | calamine |
| Modify 1k cells in 50k file | 668ms | 537ms (edit-xlsx) | edit-xlsx |

Win summary: SheetKit 11/21, calamine 8/21, edit-xlsx 1/21, xlsxwriter+calamine 1/21.

Note: `edit-xlsx` read results marked with `*` are excluded from winner selection when workload counts or sampled value probes do not match. In particular, for some files `edit-xlsx` can fail to deserialize minimal `workbook.xml` structures and fall back to defaults, which can report very low times with `rows=0`/`cells=0`. calamine, as a dedicated read-only library, remains the fastest comparable reader in this run.

## Key Findings

### 1. Read operations: ~1.2x napi overhead (typical)

Read operations show modest overhead from the napi-rs FFI layer, typically around 1.15-1.20x.
The multi-sheet scenario remains the outlier at ~2.4x due to per-sheet FFI costs. Async reads
are slightly faster than sync, averaging ~1.15x overhead.

### 2. Write operations: near parity (~1.0x)

Write operations through napi-rs show near parity with native Rust. Large data-heavy writes
(50k+ rows) are at 0.95-1.05x, while smaller writes with more overhead per row show
1.1-1.4x. The batch `setSheetData()` API keeps large writes efficient.

### 3. Async read shows dramatically lower memory

The async `Workbook.open()` API shows near-zero RSS delta for read operations compared
to sync. Read Large Data: 195.3MB (sync) vs 17.2MB (async). This is because the async
path processes data on a worker thread, reducing V8 heap pressure.

### 4. Rust ecosystem: fastest writer, competitive reader

In this run, SheetKit is the fastest writer across all write scenarios. For reads, calamine
(read-only) is fastest on comparable workloads. `edit-xlsx` can win specific modify workloads,
but read results with workload/value mismatches are excluded from winner selection.

## Summary

| Category | Sync Overhead | Async Overhead |
|----------|--------------|----------------|
| Read (typical) | ~1.2x | ~1.15x |
| Write (batch) | ~1.0x | - |
| Streaming write | ~1.66x | - |
| Round-trip | ~1.11x | - |
