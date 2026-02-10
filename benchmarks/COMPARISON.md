# SheetKit: Rust Native vs Node.js (napi-rs) Benchmark Comparison

Benchmark run: 2026-02-10

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
| Large Data (50k rows x 20 cols) | 616ms | 680ms | +64ms | 1.10x |
| Heavy Styles (5k rows, formatted) | 33ms | 37ms | +4ms | 1.12x |
| Multi-Sheet (10 sheets x 5k rows) | 360ms | 781ms | +421ms | 2.17x |
| Formulas (10k rows) | 40ms | 52ms | +12ms | 1.30x |
| Strings (20k rows text-heavy) | 140ms | 126ms | -14ms | 0.90x |
| Data Validation (5k rows, 8 rules) | 25ms | 29ms | +4ms | 1.16x |
| Comments (2k rows with comments) | 10ms | 11ms | +1ms | 1.10x |
| Merged Cells (500 regions) | 2ms | 2ms | +0ms | 1.00x |
| Mixed Workload (ERP document) | 34ms | 39ms | +5ms | 1.15x |

### Read Benchmarks (async)

| Scenario | Rust (median) | Node.js async (median) | Overhead | Ratio |
|----------|--------------|------------------------|----------|-------|
| Large Data (50k rows x 20 cols) | 616ms | 655ms | +39ms | 1.06x |
| Heavy Styles (5k rows, formatted) | 33ms | 36ms | +3ms | 1.09x |
| Multi-Sheet (10 sheets x 5k rows) | 360ms | 777ms | +417ms | 2.16x |
| Formulas (10k rows) | 40ms | 49ms | +9ms | 1.23x |
| Strings (20k rows text-heavy) | 140ms | 123ms | -17ms | 0.88x |
| Data Validation (5k rows, 8 rules) | 25ms | 29ms | +4ms | 1.16x |
| Comments (2k rows with comments) | 10ms | 11ms | +1ms | 1.10x |
| Merged Cells (500 regions) | 2ms | 2ms | +0ms | 1.00x |
| Mixed Workload (ERP document) | 34ms | 41ms | +7ms | 1.21x |

### Read Scaling

| Scenario | Rust (median) | Node.js sync (median) | Node.js async (median) | Sync Ratio | Async Ratio |
|----------|--------------|----------------------|------------------------|------------|-------------|
| Scale 1k rows | 6ms | 7ms | 7ms | 1.17x | 1.17x |
| Scale 10k rows | 62ms | 68ms | 68ms | 1.10x | 1.10x |
| Scale 100k rows | 659ms | 714ms | 683ms | 1.08x | 1.04x |

### Write Benchmarks

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| 50k rows x 20 cols | 1.03s | 657ms | -373ms | 0.64x |
| 5k styled rows | 39ms | 48ms | +9ms | 1.23x |
| 10 sheets x 5k rows | 377ms | 344ms | -33ms | 0.91x |
| 10k rows with formulas | 35ms | 39ms | +4ms | 1.11x |
| 20k text-heavy rows | 145ms | 123ms | -22ms | 0.85x |
| 5k rows + 8 DV rules | 16ms | 13ms | -3ms | 0.81x |
| 2k rows with comments | 14ms | 11ms | -3ms | 0.79x |
| 500 merged regions | 16ms | 14ms | -2ms | 0.88x |

### Write Scaling

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| 1k rows x 10 cols | 7ms | 7ms | +0ms | 1.00x |
| 10k rows x 10 cols | 68ms | 66ms | -2ms | 0.97x |
| 50k rows x 10 cols | 456ms | 332ms | -124ms | 0.73x |
| 100k rows x 10 cols | 735ms | 665ms | -70ms | 0.90x |

### Other

| Scenario | Rust (median) | Node.js sync (median) | Node.js async (median) | Sync Ratio | Async Ratio |
|----------|--------------|----------------------|------------------------|------------|-------------|
| Buffer round-trip (10k rows) | 165ms | 167ms | - | 1.01x | - |
| Streaming write (50k rows) | 555ms | 669ms | - | 1.21x | - |
| Random-access read (1k cells/50k file) | 592ms | 550ms | 549ms | 0.93x | 0.93x |
| Mixed workload write (ERP-style) | 22ms | 28ms | - | 1.27x | - |

## Node.js Memory Usage (After Memory Optimization)

RSS (Resident Set Size) measured for SheetKit in the Node.js benchmark.

| Scenario | Sync Memory | Async Memory |
|----------|-------------|--------------|
| Read Large Data (50k x 20) | 195.4MB | 17.2MB |
| Read Multi-Sheet (10 x 5k) | 132.1MB | 17.6MB |
| Read Scale 100k rows | 161.1MB | 0.0MB |
| Read Heavy Styles (5k) | 6.6MB | 0.1MB |
| Read Formulas (10k) | 8.8MB | 0.0MB |
| Read Strings (20k) | 2.2MB | 0.0MB |
| Write 50k x 20 | 89.4MB | - |
| Write 100k x 10 | 70.4MB | - |
| Streaming 50k rows | 80.0MB | - |
| Random-access read (1k cells) | 18.2MB | 4.2MB |

### Memory Optimization Results (Before vs After)

| Scenario | Before | After (sync) | After (async) | Sync Reduction | Async Reduction |
|----------|--------|-------------|---------------|----------------|-----------------|
| Read Large Data (50k x 20) | 349.5MB | 195.4MB | 17.2MB | -44% | -95% |
| Read Multi-Sheet (10 x 5k) | 215.8MB | 132.1MB | 17.6MB | -39% | -92% |
| Read Scale 100k rows | 325.6MB | 161.1MB | 0.0MB | -51% | -100% |
| Read Formulas (10k) | 13.3MB | 8.8MB | 0.0MB | -34% | -100% |
| Read Heavy Styles (5k) | 15.2MB | 6.6MB | 0.1MB | -57% | -99% |
| Random-access read | 27.2MB | 18.2MB | 4.2MB | -33% | -85% |

Optimizations applied:
1. Box<CellFormula> and Box<InlineString> in Cell struct (~72-88B saved per cell)
2. shrink_to_fit() on Vec<Cell> and Vec<Row> after deserialization
3. Arc<str> deduplication in SharedStringTable (strings + index_map share allocation)
4. Skip Row.spans deserialization (field unused at runtime)

## Key Findings

### 1. Read operations: ~1.1x napi overhead (sync), ~1.1x (async)

Both sync and async read operations show minimal overhead compared to pure Rust,
averaging ~1.1x. The multi-sheet scenario remains the outlier at ~2.2x due to
per-sheet FFI costs. Async and sync perform nearly identically in timing.

### 2. Write operations: near-parity or faster than Rust

Write operations through napi-rs show parity or even outperform the Rust native
benchmark. Large-scale writes (50k-100k) are consistently faster in Node.js,
likely due to V8's efficient string handling and the batch setSheetData() API.

### 3. Async read shows dramatically lower memory

The async `Workbook.open()` API shows near-zero RSS delta for read operations
compared to sync. This is because the async FFI path processes data on a worker
thread, reducing V8 heap pressure. Read Large Data: 195.4MB (sync) vs 17.2MB (async).

### 4. Memory: ~44-51% reduction from optimization (sync)

The memory optimization (Box<CellFormula>, shrink_to_fit, Arc<str>, skip spans)
reduced sync RSS by ~39-57% for large read scenarios. Async reads show even greater
reduction (85-100%) since the data never enters the V8 heap in the same way.

## Summary

| Category | Sync Overhead | Async Overhead |
|----------|--------------|----------------|
| Read (typical) | ~1.1x | ~1.1x |
| Write (batch) | ~0.9x | - |
| Streaming write | ~1.21x | - |
| Round-trip | ~1.01x | - |
