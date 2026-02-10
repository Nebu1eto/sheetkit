# Performance

SheetKit delivers native Rust performance to both Rust and TypeScript applications. This page demonstrates how fast SheetKit is and explains the optimizations that make it possible.

## How Fast is SheetKit?

### Rust vs Node.js Overhead

SheetKit's Node.js bindings stay close to native Rust performance, and in several write-heavy paths they are faster:

| Operation | Overhead |
|-----------|----------|
| **Read operations (sync)** | ~1.10x (~10% slower, typical) |
| **Read operations (async)** | ~1.10x (~10% slower, typical) |
| **Write operations (batch)** | ~0.90x (~10% faster, typical) |
| **Streaming write** | 1.21x (21% slower) |
| **Buffer round-trip** | 1.01x (near parity) |

For most real-world workloads, Node.js performance remains close to native Rust.

### Read Performance Comparison

| Scenario | Rust | Node.js | Overhead |
|----------|------|---------|----------|
| Large Data (50k rows × 20 cols) | 616ms | 680ms | +10% |
| Heavy Styles (5k rows, formatted) | 33ms | 37ms | +12% |
| Multi-Sheet (10 sheets × 5k rows) | 360ms | 781ms | +117% |
| Formulas (10k rows) | 40ms | 52ms | +30% |
| Strings (20k rows text-heavy) | 140ms | 126ms | -10% (faster) |

### Write Performance Comparison

| Scenario | Rust | Node.js | Overhead |
|----------|------|---------|----------|
| 50k rows × 20 cols | 1.03s | 657ms | -36% (faster) |
| 5k styled rows | 39ms | 48ms | +23% |
| 10k rows with formulas | 35ms | 39ms | +11% |
| 20k text-heavy rows | 145ms | 123ms | -15% (faster) |

Note: In some write scenarios, Node.js performs slightly better than Rust due to V8's efficient string handling during data construction.

### Scaling Performance

Read performance remains consistent across different file sizes:

| Rows | Rust | Node.js | Overhead |
|------|------|---------|----------|
| 1k | 6ms | 7ms | +17% |
| 10k | 62ms | 68ms | +10% |
| 100k | 659ms | 714ms | +8% |

Write performance scales linearly:

| Rows | Rust | Node.js | Overhead |
|------|------|---------|----------|
| 1k | 7ms | 7ms | 0% |
| 10k | 68ms | 66ms | -3% (faster) |
| 50k | 456ms | 332ms | -27% (faster) |
| 100k | 735ms | 665ms | -10% (faster) |

## Raw Buffer Transfer and Memory Behavior

SheetKit reduces Node.js-Rust boundary cost by transferring sheet data as raw buffers instead of per-cell JavaScript objects. This transfer model keeps the FFI boundary coarse-grained, reduces object marshalling overhead, and lowers GC pressure in read-heavy paths.

## Key Optimizations

### 1. Buffer-Based FFI Transfer

Instead of creating individual JavaScript objects for each cell, SheetKit serializes entire sheets into compact binary buffers that cross the FFI boundary in a single operation.

**Before**: Per-cell object transfer across the FFI boundary
**After**: Single raw-buffer transfer for a sheet payload

This optimization:
- Reduces read-side FFI overhead
- Reduces allocation and GC pressure from per-cell object creation
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

## Benchmark Scope and Data

The numbers on this page are from SheetKit's own Rust and Node.js benchmark suites in this repository. Results vary based on data shape, feature usage, and runtime environment.

For benchmark methodology and raw data, see [`benchmarks/COMPARISON.md`](https://github.com/Nebu1eto/sheetkit/blob/main/benchmarks/COMPARISON.md) in the repository.

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
