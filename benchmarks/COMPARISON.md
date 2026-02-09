# SheetKit: Rust Native vs Node.js (napi-rs) Benchmark Comparison

Benchmark run: 2026-02-09

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
SheetKit from Node.js vs using it directly from Rust.

### Read Benchmarks

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| Large Data (50k rows x 20 cols) | 625ms | 1.26s | +635ms | 2.02x |
| Heavy Styles (5k rows, formatted) | 33ms | 62ms | +29ms | 1.88x |
| Multi-Sheet (10 sheets x 5k rows) | 369ms | 602ms | +233ms | 1.63x |
| Formulas (10k rows) | 43ms | 76ms | +33ms | 1.77x |
| Strings (20k rows text-heavy) | 137ms | 234ms | +97ms | 1.71x |
| Data Validation (5k rows, 8 rules) | 26ms | 45ms | +19ms | 1.73x |
| Comments (2k rows with comments) | 10ms | 15ms | +5ms | 1.50x |
| Merged Cells (500 regions) | 2ms | 3ms | +1ms | 1.50x |
| Mixed Workload (ERP document) | 34ms | 58ms | +24ms | 1.71x |

### Read Scaling

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| Scale 1k rows | 6ms | 11ms | +5ms | 1.83x |
| Scale 10k rows | 65ms | 113ms | +48ms | 1.74x |
| Scale 100k rows | 650ms | 1.23s | +580ms | 1.89x |

### Write Benchmarks

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| 50k rows x 20 cols | 742ms | 727ms | -15ms | 0.98x |
| 5k styled rows | 41ms | 51ms | +10ms | 1.24x |
| 10 sheets x 5k rows | 381ms | 369ms | -12ms | 0.97x |
| 10k rows with formulas | 34ms | 42ms | +8ms | 1.24x |
| 20k text-heavy rows | 148ms | 130ms | -18ms | 0.88x |
| 5k rows + 8 DV rules | 12ms | 14ms | +2ms | 1.17x |
| 2k rows with comments | 9ms | 11ms | +2ms | 1.22x |
| 500 merged regions | 13ms | 14ms | +1ms | 1.08x |

### Write Scaling

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| 1k rows x 10 cols | 6ms | 7ms | +1ms | 1.17x |
| 10k rows x 10 cols | 69ms | 68ms | -1ms | 0.99x |
| 50k rows x 10 cols | 347ms | 363ms | +16ms | 1.05x |
| 100k rows x 10 cols | 705ms | 699ms | -6ms | 0.99x |

### Other

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| Buffer round-trip (10k rows) | 169ms | 221ms | +52ms | 1.31x |
| Streaming write (50k rows) | 1.02s | 1.22s | +200ms | 1.20x |
| Random-access read (1k cells/50k file) | 585ms | 579ms | -6ms | 0.99x |
| Mixed workload write (ERP-style) | 23ms | 29ms | +6ms | 1.26x |

## Key Findings

### 1. Read operations: consistent ~1.5-2.0x overhead

Reading operations show the most consistent napi-rs overhead because each read involves:
- Parsing the .xlsx (ZIP + XML) in Rust (same cost)
- Serializing the result data from Rust types to JavaScript types via napi-rs
- The `getRows()` call materializes all cell data into JS objects

The overhead scales roughly linearly with data size (1.83x at 1k rows, 1.74x at 10k, 1.89x at 100k),
confirming the cost is proportional to the amount of data crossing the FFI boundary.

### 2. Write operations: near-zero overhead (~0.9-1.25x)

Write operations show minimal napi overhead because the Node.js API uses batch methods
(`setSheetData`) that pass data in bulk, minimizing FFI round-trips. Some scenarios even
show Node.js being marginally faster (e.g., 20k strings at 0.88x), likely due to:
- Node.js `setSheetData` batching vs Rust `set_cell_value` per-cell calls in the benchmark
- V8's efficient string interning for the data construction phase
- Measurement noise within margin of error

### 3. Streaming write: ~20% overhead

The streaming writer shows moderate overhead (1.20x) because each `writeRow()` call crosses
the FFI boundary. With 50,000 rows, that's 50,000 FFI calls. This is expected.

### 4. Buffer round-trip: ~31% overhead

The round-trip test (write to buffer + read back) shows 31% overhead, which combines the
write overhead (near-zero) with the read overhead (~1.7x), averaging out.

### 5. Random-access read: near parity

The random-access read shows near parity (0.99x ratio). The benchmark includes the `open()`
call (parsing the entire file), and the 1,000 `getCellValue()` lookups are fast hash-table
lookups. The difference is within measurement noise.

## Summary

| Category | Average Overhead | Description |
|----------|-----------------|-------------|
| Read | ~1.75x | Data serialization across FFI boundary |
| Write (batch) | ~1.05x | Minimal - bulk data transfer |
| Write (per-row) | ~1.20x | Moderate - per-call FFI cost |
| Round-trip | ~1.31x | Combined read + write overhead |

The napi-rs FFI layer adds roughly **0-25% overhead for write operations** and **50-100% overhead
for read operations**. The read overhead is dominated by the cost of converting Rust data
structures into JavaScript objects. For write-heavy workloads (the common case for Excel
generation), the overhead is negligible.
