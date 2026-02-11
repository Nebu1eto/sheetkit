# Performance

SheetKit delivers native Rust performance to both Rust and TypeScript applications. This page demonstrates how fast SheetKit is and explains the optimizations that make it possible.

## How Fast is SheetKit?

### Compared with ExcelJS and SheetJS (Node.js)

In the existing Node.js benchmark suite (`benchmarks/node/RESULTS.md`), SheetKit is consistently faster than both ExcelJS and SheetJS in representative read/write workloads:

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| Read Large Data (50k rows x 20 cols) | 454ms | 2.69s | 1.59s |
| Write 50k rows x 20 cols | 461ms | 2.55s | 1.25s |
| Buffer round-trip (10k rows) | 118ms | 479ms | 147ms |
| Random-access read (1k cells from 50k-row file) | 387ms | 2.81s | 1.29s |

### Compared with Rust Excel Libraries

Among pure Rust libraries, SheetKit is the fastest writer. For reads, calamine (read-only) and edit-xlsx (lazy parsing) are faster, but SheetKit is the only library supporting full read+modify+write in a single crate.

| Scenario | SheetKit | calamine | rust_xlsxwriter | edit-xlsx |
|----------|----------|----------|-----------------|-----------|
| Read Large Data (50k rows) | 390ms | 299ms | N/A | 35ms |
| Write 50k rows x 20 cols | 459ms | N/A | 847ms | 886ms |
| Streaming write (50k rows) | 184ms | N/A | 858ms | N/A |
| Modify 1k cells in 50k file | 588ms | N/A | N/A | N/A |

### Rust vs Node.js Overhead

SheetKit's Node.js bindings stay close to native Rust performance:

| Operation | Overhead |
|-----------|----------|
| **Read operations (sync)** | ~1.20x (~20% slower, typical) |
| **Read operations (async)** | ~1.15x (~15% slower, typical) |
| **Write operations (batch)** | ~1.0x (near parity) |
| **Streaming write** | 1.66x (66% slower) |
| **Buffer round-trip** | 1.11x (11% slower) |

For most real-world workloads, Node.js performance remains close to native Rust.

### Read Performance Comparison

| Scenario | Rust | Node.js | Overhead |
|----------|------|---------|----------|
| Large Data (50k rows x 20 cols) | 387ms | 454ms | +17% |
| Heavy Styles (5k rows, formatted) | 20ms | 24ms | +20% |
| Multi-Sheet (10 sheets x 5k rows) | 223ms | 530ms | +138% |
| Formulas (10k rows) | 24ms | 33ms | +38% |
| Strings (20k rows text-heavy) | 83ms | 95ms | +14% |

### Write Performance Comparison

| Scenario | Rust | Node.js | Overhead |
|----------|------|---------|----------|
| 50k rows x 20 cols | 478ms | 461ms | -4% (faster) |
| 5k styled rows | 25ms | 35ms | +40% |
| 10k rows with formulas | 21ms | 28ms | +33% |
| 20k text-heavy rows | 92ms | 87ms | -5% (faster) |

Note: In some write scenarios, Node.js performs slightly better than Rust due to V8's efficient string handling during data construction and the batch `setSheetData()` API.

### Scaling Performance

Read performance remains consistent across different file sizes:

| Rows | Rust | Node.js | Overhead |
|------|------|---------|----------|
| 1k | 4ms | 5ms | +25% |
| 10k | 39ms | 45ms | +15% |
| 100k | 410ms | 474ms | +16% |

Write performance scales linearly:

| Rows | Rust | Node.js | Overhead |
|------|------|---------|----------|
| 1k | 4ms | 5ms | +25% |
| 10k | 48ms | 47ms | -2% (faster) |
| 50k | 226ms | 235ms | +4% |
| 100k | 454ms | 476ms | +5% |

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

Use `StreamWriter` for sequential row writes. Each `write_row()` writes directly to a temp file on disk, so memory usage stays constant regardless of the number of rows:

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

> **Note:** Cell values in streamed sheets cannot be read directly after `applyStreamWriter`. Save the workbook and reopen it to read the data.

## Next Steps

- [Getting Started](./getting-started.md) - Learn the basics
- [Architecture](./architecture.md) - Understand internal design
- [API Reference](./api-reference/) - Explore all available methods
