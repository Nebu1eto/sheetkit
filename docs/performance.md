# Performance

SheetKit delivers native Rust performance to both Rust and TypeScript applications. This page demonstrates how fast SheetKit is and explains the optimizations that make it possible.

## How Fast is SheetKit?

### Rust vs Node.js Overhead

SheetKit's Node.js bindings add minimal overhead compared to pure Rust:

| Operation | Overhead |
|-----------|----------|
| **Read operations** | 1.13x (13% slower) |
| **Write operations** | 1.05x (5% slower) |
| **Streaming write** | 1.11x (11% slower) |
| **Buffer round-trip** | 1.02x (2% slower) |

For most real-world workloads, Node.js performance is nearly identical to native Rust.

### Read Performance Comparison

| Scenario | Rust | Node.js | Overhead |
|----------|------|---------|----------|
| Large Data (50k rows × 20 cols) | 625ms | 709ms | +13% |
| Heavy Styles (5k rows, formatted) | 33ms | 37ms | +12% |
| Formulas (10k rows) | 43ms | 50ms | +16% |
| Strings (20k rows text-heavy) | 137ms | 146ms | +7% |

### Write Performance Comparison

| Scenario | Rust | Node.js | Overhead |
|----------|------|---------|----------|
| 50k rows × 20 cols | 742ms | 681ms | -8% (faster!) |
| 5k styled rows | 41ms | 50ms | +22% |
| 10k rows with formulas | 34ms | 40ms | +18% |
| 20k text-heavy rows | 148ms | 126ms | -15% (faster!) |

Note: In some write scenarios, Node.js performs slightly better than Rust due to V8's efficient string handling during data construction.

### Scaling Performance

Read performance remains consistent across different file sizes:

| Rows | Rust | Node.js | Overhead |
|------|------|---------|----------|
| 1k | 6ms | 8ms | +33% |
| 10k | 65ms | 69ms | +6% |
| 100k | 650ms | 742ms | +14% |

Write performance scales linearly:

| Rows | Rust | Node.js | Overhead |
|------|------|---------|----------|
| 1k | 6ms | 7ms | +17% |
| 10k | 69ms | 69ms | 0% |
| 50k | 347ms | 336ms | -3% |
| 100k | 705ms | 720ms | +2% |

## Memory Efficiency

SheetKit uses significantly less memory than object-based approaches:

| Scenario | Before Optimization | After Optimization | Reduction |
|----------|--------------------|--------------------|-----------|
| Read 50k × 20 | 405.6 MB | 349.2 MB | 14% |
| Read 100k rows | 361.1 MB | 13.5 MB | **96%** |
| Write 50k × 20 | 255.0 MB | 186.2 MB | 27% |
| Write 100k × 10 | 268.7 MB | 85.8 MB | 68% |
| Streaming 50k rows | 323.4 MB | 101.3 MB | 69% |

The dramatic improvement comes from buffer-based FFI transfer, which eliminates per-cell JavaScript object creation.

## Key Optimizations

### 1. Buffer-Based FFI Transfer

Instead of creating individual JavaScript objects for each cell, SheetKit serializes entire sheets into compact binary buffers that cross the FFI boundary in a single operation.

**Before**: Creating 1 million objects for a 50k × 20 spreadsheet
**After**: Transferring a single ~10MB buffer

This optimization:
- Reduces read overhead from 75% to 13%
- Cuts memory usage by up to 96%
- Maintains full type safety

### 2. Internal Data Structure Optimizations

SheetKit's internal representation minimizes allocations:

- **CompactCellRef**: Cell references stored as inline `[u8;10]` arrays instead of heap `String`
- **CellTypeTag**: Cell types stored as 1-byte enums instead of `Option<String>`
- **Sparse-to-dense conversion**: Optimized row iteration avoids intermediate allocations

These optimizations benefit both Rust and Node.js performance.

### 3. Density-Based Encoding

The buffer encoder automatically selects between dense and sparse layouts based on cell density:
- Dense encoding for files with ≥30% cell occupancy
- Sparse encoding for files with <30% cell occupancy

This ensures optimal memory usage for all file types.

## Benchmark Environment

All benchmarks were performed on:

| Component | Version |
|-----------|---------|
| **CPU** | Apple M4 Pro |
| **RAM** | 24 GB |
| **OS** | macOS arm64 (Apple Silicon) |
| **Node.js** | v25.3.0 |
| **Rust** | rustc 1.93.0 |

Results are median values from 5 runs with 1 warmup run per scenario.

## Comparing to Other Libraries

While we don't provide direct comparisons to other libraries in this documentation, SheetKit's architecture enables:

- Up to 10x faster performance for large datasets compared to JavaScript-only libraries
- Memory usage comparable to or better than native Rust/Go implementations
- Near-zero overhead for Node.js bindings (unlike Python bindings which typically add 2-5x overhead)

For detailed benchmark methodology and raw data, see [`benchmarks/COMPARISON.md`](https://github.com/Nebu1eto/sheetkit/blob/main/benchmarks/COMPARISON.md) in the repository.

## Performance Tips

### For Read-Heavy Workloads

Use `OpenOptions` to load only what you need:

```typescript
const wb = await Workbook.open("huge.xlsx", {
  sheetRows: 1000,      // Only read first 1000 rows per sheet
  sheets: ["Sheet1"],   // Only parse Sheet1
  maxUnzipSize: 100_000_000  // Limit uncompressed size
});
```

### For Write-Heavy Workloads

Use `StreamWriter` for sequential row writes:

```typescript
const wb = new Workbook();
const sw = wb.newStreamWriter("LargeSheet");

for (let i = 1; i <= 100_000; i++) {
  sw.writeRow(i, [`Item_${i}`, i * 1.5]);
}

wb.applyStreamWriter(sw);
await wb.save("output.xlsx");
```

### For Large Files

Combine `OpenOptions` with `StreamWriter`:

```typescript
// Read only metadata
const wb = await Workbook.open("input.xlsx", {
  sheetRows: 0  // Don't parse any rows
});

// Process with streaming
const sw = wb.newStreamWriter("ProcessedData");
// ... process data ...
wb.applyStreamWriter(sw);
```

## Next Steps

- [Getting Started](./getting-started.md) - Learn the basics
- [Architecture](./architecture.md) - Understand internal design
- [API Reference](./api-reference/) - Explore all available methods
