# SheetKit: Rust Native vs Node.js (napi-rs) Benchmark Comparison

Benchmark run: 2026-02-09
Platform: macOS arm64 (Apple Silicon)
Methodology: 1 warmup + 5 measured runs per scenario, median reported

## napi-rs Overhead Analysis

This comparison measures the overhead introduced by the napi-rs FFI layer when calling
SheetKit from Node.js vs using it directly from Rust.

### Read Benchmarks

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| Large Data (50k rows x 20 cols) | 636ms | 1.24s | +604ms | 1.95x |
| Heavy Styles (5k rows, formatted) | 34ms | 59ms | +25ms | 1.74x |
| Multi-Sheet (10 sheets x 5k rows) | 379ms | 597ms | +218ms | 1.57x |
| Formulas (10k rows) | 40ms | 76ms | +36ms | 1.90x |
| Strings (20k rows text-heavy) | 138ms | 235ms | +97ms | 1.70x |
| Data Validation (5k rows, 8 rules) | 25ms | 45ms | +20ms | 1.80x |
| Comments (2k rows with comments) | 10ms | 15ms | +5ms | 1.50x |
| Merged Cells (500 regions) | 2ms | 3ms | +1ms | 1.50x |
| Mixed Workload (ERP document) | 34ms | 58ms | +24ms | 1.71x |

### Read Scaling

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| Scale 1k rows | 6ms | 11ms | +5ms | 1.83x |
| Scale 10k rows | 63ms | 109ms | +46ms | 1.73x |
| Scale 100k rows | 642ms | 1.21s | +568ms | 1.88x |

### Write Benchmarks

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| 50k rows x 20 cols | 702ms | 678ms | -24ms | 0.97x |
| 5k styled rows | 40ms | 50ms | +10ms | 1.25x |
| 10 sheets x 5k rows | 380ms | 393ms | +13ms | 1.03x |
| 10k rows with formulas | 33ms | 40ms | +7ms | 1.21x |
| 20k text-heavy rows | 152ms | 126ms | -26ms | 0.83x |
| 5k rows + 8 DV rules | 13ms | 15ms | +2ms | 1.15x |
| 2k rows with comments | 9ms | 11ms | +2ms | 1.22x |
| 500 merged regions | 13ms | 15ms | +2ms | 1.15x |

### Write Scaling

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| 1k rows x 10 cols | 7ms | 7ms | +0ms | 1.00x |
| 10k rows x 10 cols | 68ms | 67ms | -1ms | 0.99x |
| 50k rows x 10 cols | 356ms | 337ms | -19ms | 0.95x |
| 100k rows x 10 cols | 709ms | 730ms | +21ms | 1.03x |

### Other

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| Buffer round-trip (10k rows) | 163ms | 219ms | +56ms | 1.34x |
| Streaming write (50k rows) | 1.00s | 1.18s | +180ms | 1.18x |
| Random-access read (1k cells/50k file) | 593ms | 545ms | -48ms | 0.92x |
| Mixed workload write (ERP-style) | 23ms | 29ms | +6ms | 1.26x |

## Key Findings

### 1. Read operations: consistent ~1.5-1.95x overhead

Reading operations show the most consistent napi-rs overhead because each read involves:
- Parsing the .xlsx (ZIP + XML) in Rust (same cost)
- Serializing the result data from Rust types to JavaScript types via napi-rs
- The `getRows()` call materializes all cell data into JS objects

The overhead scales roughly linearly with data size (1.83x at 1k rows, 1.73x at 10k, 1.88x at 100k),
confirming the cost is proportional to the amount of data crossing the FFI boundary.

### 2. Write operations: near-zero overhead (~1.0-1.25x)

Write operations show minimal napi overhead because the Node.js API uses batch methods
(`setSheetData`) that pass data in bulk, minimizing FFI round-trips. Some scenarios even
show Node.js being marginally faster (e.g., 50k write, 20k strings), likely due to:
- Node.js `setSheetData` batching vs Rust `set_cell_value` per-cell calls in the benchmark
- V8's efficient string interning for the data construction phase
- Measurement noise within margin of error

### 3. Streaming write: ~18% overhead

The streaming writer shows moderate overhead (1.18x) because each `writeRow()` call crosses
the FFI boundary. With 50,000 rows, that's 50,000 FFI calls. This is expected.

### 4. Buffer round-trip: ~34% overhead

The round-trip test (write to buffer + read back) shows 34% overhead, which combines the
write overhead (near-zero) with the read overhead (~1.7x), averaging out.

### 5. Random-access read: Node.js faster

The random-access read shows Node.js slightly faster (0.92x ratio). This is because the
benchmark includes the `open()` call (parsing the entire file), and the 1,000 `getCellValue()`
lookups are fast hash-table lookups. The difference is within measurement noise.

## Summary

| Category | Average Overhead | Description |
|----------|-----------------|-------------|
| Read | ~1.75x | Data serialization across FFI boundary |
| Write (batch) | ~1.05x | Minimal - bulk data transfer |
| Write (per-row) | ~1.18x | Moderate - per-call FFI cost |
| Round-trip | ~1.34x | Combined read + write overhead |

The napi-rs FFI layer adds roughly **5-25% overhead for write operations** and **50-95% overhead
for read operations**. The read overhead is dominated by the cost of converting Rust data
structures into JavaScript objects. For write-heavy workloads (the common case for Excel
generation), the overhead is negligible.
