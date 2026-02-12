# Performance

SheetKit delivers native Rust performance to both Rust and TypeScript applications. This page demonstrates how fast SheetKit is and explains the optimizations that make it possible.

## How Fast is SheetKit?

### Compared with ExcelJS and SheetJS (Node.js)

In the existing Node.js benchmark suite (`benchmarks/node/RESULTS.md`), SheetKit is consistently faster than both ExcelJS and SheetJS in representative read/write workloads:

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| Read Large Data (50k rows x 20 cols) | 546ms | 2.91s | 1.65s |
| Write 50k rows x 20 cols | 489ms | 2.73s | 1.37s |
| Buffer round-trip (10k rows) | 131ms | 498ms | 198ms |
| Random-access read (1k cells from 50k-row file) | 489ms | 3.04s | 1.37s |

### Compared with Rust Excel Libraries

Among pure Rust libraries, SheetKit is the fastest writer. For reads, calamine (read-only) and edit-xlsx (lazy parsing) are faster, but SheetKit is the only library supporting full read+modify+write in a single crate.

| Scenario | SheetKit | calamine | rust_xlsxwriter | edit-xlsx |
|----------|----------|----------|-----------------|-----------|
| Read Large Data (50k rows) | 525ms | 331ms | N/A | 39ms* |
| Write 50k rows x 20 cols | 527ms | N/A | 940ms | 971ms |
| Streaming write (50k rows) | 202ms | N/A | 931ms | N/A |
| Modify 1k cells in 50k file (lazy) | 688ms | N/A | N/A | N/A |

\* edit-xlsx reads 0 cells (lazy open only); not directly comparable.

### Rust vs Node.js Overhead

SheetKit's Node.js bindings stay close to native Rust performance:

| Operation | Overhead |
|-----------|----------|
| **Read operations (sync)** | ~1.05x (~5% slower, typical) |
| **Read operations (async)** | ~1.02x (~2% slower, typical) |
| **Write operations (batch)** | ~1.0x (near parity) |
| **Streaming write** | 1.68x (68% slower) |
| **Buffer round-trip** | 1.07x (7% slower) |

For most real-world workloads, Node.js performance remains close to native Rust.

### Read Performance Comparison

| Scenario | Rust | Node.js | Overhead |
|----------|------|---------|----------|
| Large Data (50k rows x 20 cols) | 518ms | 546ms | +5% |
| Heavy Styles (5k rows, formatted) | 27ms | 29ms | +7% |
| Multi-Sheet (10 sheets x 5k rows) | 301ms | 625ms | +108% |
| Formulas (10k rows) | 33ms | 42ms | +27% |
| Strings (20k rows text-heavy) | 108ms | 112ms | +4% |

### Write Performance Comparison

| Scenario | Rust | Node.js | Overhead |
|----------|------|---------|----------|
| 50k rows x 20 cols | 503ms | 489ms | -3% (faster) |
| 5k styled rows | 28ms | 36ms | +29% |
| 10k rows with formulas | 24ms | 30ms | +25% |
| 20k text-heavy rows | 107ms | 90ms | -16% (faster) |

Note: In some write scenarios, Node.js performs slightly better than Rust due to V8's efficient string handling during data construction and the batch `setSheetData()` API.

### Scaling Performance

Read performance remains consistent across different file sizes:

| Rows | Rust | Node.js | Overhead |
|------|------|---------|----------|
| 1k | 5ms | 6ms | +20% |
| 10k | 51ms | 55ms | +8% |
| 100k | 530ms | 565ms | +7% |

Write performance scales linearly:

| Rows | Rust | Node.js | Overhead |
|------|------|---------|----------|
| 1k | 5ms | 5ms | 0% |
| 10k | 47ms | 48ms | +2% |
| 50k | 247ms | 244ms | -1% (faster) |
| 100k | 518ms | 508ms | -2% (faster) |

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

### Choosing a Read Mode

SheetKit supports three read modes that control how much parsing is done during `open()`:

| Read Mode | Open Cost | Memory on Open | Best For |
|-----------|-----------|----------------|----------|
| `lazy` (default) | Low -- ZIP index + metadata only | Minimal | Most workloads. Sheets are parsed on first access. |
| `eager` | High -- all sheets parsed | Full workbook in memory | When you need all sheets immediately after open. |
| `stream` | Minimal | Near-zero | Forward-only iteration over very large sheets. |

```typescript
// Lazy open (default): fastest open, parses sheets on demand
const wb = await Workbook.open("huge.xlsx");
const rows = wb.getRows("Sheet1"); // Sheet1 parsed here

// Eager open: all sheets parsed during open
const wb2 = await Workbook.open("huge.xlsx", { readMode: "eager" });

// Stream mode: bounded-memory forward-only reading
const wb3 = await Workbook.open("huge.xlsx", { readMode: "stream" });
const reader = await wb3.openSheetReader("Sheet1", { batchSize: 500 });
for await (const batch of reader) {
  for (const row of batch) {
    process(row);
  }
}
```

### Deferred Auxiliary Parts

By default, auxiliary parts (comments, charts, images, pivot tables) are not parsed during open. They load on-demand when you first call a method that needs them. Set `auxParts: 'eager'` if you need all parts available immediately:

```typescript
const wb = await Workbook.open("report.xlsx", { auxParts: "eager" });
```

### For Read-Heavy Workloads

Use `OpenOptions` to load only what you need:

```typescript
const wb = await Workbook.open("huge.xlsx", {
  readMode: "lazy",
  sheetRows: 1000,      // Only read first 1000 rows per sheet
  sheets: ["Sheet1"],   // Only parse Sheet1
  maxUnzipSize: 100_000_000  // Limit uncompressed size
});
```

### Streaming Reads with SheetStreamReader

For very large files where you do not need random cell access, use `openSheetReader()` for forward-only bounded-memory iteration:

```typescript
const wb = await Workbook.open("huge.xlsx", { readMode: "stream" });
const reader = await wb.openSheetReader("Sheet1", { batchSize: 1000 });

for await (const batch of reader) {
  for (const row of batch) {
    // Process each row -- only one batch in memory at a time
  }
}
```

### Raw Buffer V2 Transfer

`getRowsBufferV2()` produces a v2 binary buffer with inline strings, enabling incremental row-by-row decoding without eagerly materializing a global string table:

```typescript
const bufV2 = wb.getRowsBufferV2("Sheet1");
```

### Copy-on-Write Save

When saving a workbook opened in `lazy` mode, unchanged sheets are written directly from the original ZIP entry without parse-serialize round-trips. This significantly reduces save latency for workloads that modify only a few sheets in a large workbook.

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

Combine lazy open with `StreamWriter`:

```typescript
// Lazy open -- only metadata is parsed
const wb = await Workbook.open("input.xlsx");

// Process with streaming
const sw = wb.newStreamWriter("ProcessedData");
// ... process data ...
wb.applyStreamWriter(sw);
```

> **Note:** Cell values in streamed sheets cannot be read directly after `applyStreamWriter`. Save the workbook and reopen it to read the data.

## Performance KPIs and Regression Criteria

This section defines measurable KPI thresholds for the three read modes (`lazy`, `stream`, `eager`) introduced in the async-first lazy-open refactor. All targets are expressed relative to the current `Full` mode baseline. A regression is any measured value that exceeds the tolerance defined here.

### Reference Fixtures

The primary fixtures used for KPI measurement are:

| Fixture | Rows | Columns | Size | Purpose |
|---------|------|---------|------|---------|
| `large-data.xlsx` | 50,001 | 20 | 7.2 MB | Peak RSS and throughput (primary) |
| `scale-100k.xlsx` | 100,001 | 10 | 7.2 MB | Scaling upper bound |
| `scale-10k.xlsx` | 10,001 | 10 | 727 KB | Mid-range baseline |
| `scale-1k.xlsx` | 1,001 | 10 | 75 KB | Overhead measurement |
| `multi-sheet.xlsx` | 50,010 | 10 | 3.7 MB | Multi-sheet lazy hydration |

### 1. Open Latency

Measured as time from `open()` / `Workbook::open()` call to returned handle, excluding any subsequent data access.

| Read Mode | Target (relative to Full baseline) | Tolerance | Notes |
|-----------|-------------------------------------|-----------|-------|
| `lazy` | < 30% of Full | +5% | Metadata + ZIP index only; no sheet XML parse |
| `stream` | < 20% of Full | +5% | Minimal parse; no materialization |
| `eager` | No regression | +5% | Same behavior as current Full mode |

**Measurement**: Compare median open latency across all reference fixtures. Lazy and stream targets apply to every fixture individually (not just the aggregate).

### 2. Peak RSS (Memory)

Measured as peak resident set size during the operation. Use `--expose-gc` in Node.js benchmarks and external RSS sampling for Rust benchmarks.

| Read Mode | Scenario | Target (relative to Full baseline) | Tolerance |
|-----------|----------|------------------------------------|-----------|
| `lazy` | Open only (no data access) | < 20% of Full peak RSS | +5% |
| `lazy` | Open + single sheet read | < 40% of Full peak RSS | +5% |
| `lazy` | Open + all sheets read | No regression | +10% |
| `stream` | Any file size | < 50 MB absolute | +10 MB |
| `eager` | Full workbook | No regression | +5% |

**Primary fixture for RSS**: `large-data.xlsx` (50k x 20, 1M cells). The `stream` mode 50 MB bound must hold for `scale-100k.xlsx` as well.

### 3. getRows Throughput

Measured as total time for `getRows("Sheet1")` on a pre-opened workbook, or equivalent Rust `get_rows()`.

| Read Mode | Scenario | Target (relative to Full baseline) | Tolerance |
|-----------|----------|------------------------------------|-----------|
| `lazy` | Lazy open then getRows | No regression | +10% |
| `stream` | Batch iteration (all rows) | >= 80% of eager throughput | -20% floor |
| `eager` | getRows on pre-opened | No regression | +5% |

**Measurement**: Use `getRows-only` (pre-opened) category from `bench-baseline.ts` and `get_rows` group from `open_modes.rs`. The lazy-then-read scenario includes the on-demand hydration cost.

### 4. Save Latency

Measured as time from `save()` call to file written, using a temporary file target.

| Read Mode | Scenario | Target (relative to Full baseline) | Tolerance |
|-----------|----------|------------------------------------|-----------|
| `lazy` | Untouched workbook save | < 50% of Full save | +10% |
| `lazy` | Single-cell edit then save | No regression | +10% |
| `eager` | Full save | No regression | +5% |

**Rationale**: Untouched lazy workbooks should benefit from passthrough (no parse-serialize round-trip for unchanged parts). The single-cell edit scenario verifies that selective materialization does not introduce unexpected overhead.

### 5. Node.js Async Overhead

Measured as the ratio of async API latency to sync API latency for equivalent operations.

| Operation | Max Async/Sync Ratio | Notes |
|-----------|---------------------|-------|
| `open` (path) | 1.3x | Worker thread dispatch overhead |
| `openBuffer` | 1.3x | Buffer transfer + dispatch |
| `getRows` | 1.2x | Data marshalling overhead |
| `save` | 1.3x | File I/O + dispatch |

These ratios apply to the `eager` mode baseline. In `lazy` mode, async open should be faster in absolute terms since it does less work.

### Regression Fail Criteria

A benchmark result is a **regression** if any of the following hold:

1. **Open latency** for any read mode exceeds the target + tolerance on any reference fixture.
2. **Peak RSS** for any read mode exceeds the target + tolerance on `large-data.xlsx` or `scale-100k.xlsx`.
3. **getRows throughput** for any read mode is slower than the target - tolerance on any reference fixture.
4. **Save latency** for any scenario exceeds the target + tolerance on any reference fixture.
5. **Stream mode RSS** exceeds 60 MB (50 MB target + 10 MB tolerance) on any fixture regardless of file size.

**Blocking policy**: Any regression blocks PR merge. The PR author must either fix the regression or update the baseline with justification reviewed by at least one maintainer.

### Baseline Capture Process

#### Capturing a baseline

1. Ensure fixtures exist: `cd benchmarks/node && pnpm generate`
2. Run Rust benchmarks: `cargo bench --bench open_modes`
3. Run Node.js benchmarks: `node --expose-gc --import tsx benchmarks/node/bench-baseline.ts`
4. Criterion results are stored in `target/criterion/` (Rust). Node.js results are written to `benchmarks/node/baseline-results.json`.

#### Baseline file format

The Node.js baseline is a JSON file with this structure:

```json
{
  "timestamp": "ISO-8601",
  "platform": "darwin arm64",
  "nodeVersion": "v25.x.x",
  "cpu": "Apple M4 Pro",
  "ramGb": 24,
  "config": { "warmupRuns": 1, "benchRuns": 5 },
  "results": [
    {
      "category": "openSync-path",
      "fixture": "large-data",
      "median": 387.0,
      "p95": 395.0,
      "rssMedianMb": 120.0
    }
  ]
}
```

The Rust baseline uses Criterion's built-in JSON output under `target/criterion/<group>/<benchmark>/new/estimates.json`.

#### Updating baselines

1. Run benchmarks on consistent hardware (same machine, minimal background load).
2. Compare new results against stored baseline using percentage delta.
3. If all KPIs pass, overwrite the baseline JSON with the new results.
4. Commit the updated baseline with a message explaining the reason (optimization landed, fixture changed, etc.).

#### Comparison logic

For each KPI entry, compute:

```
delta_pct = ((new_value - baseline_value) / baseline_value) * 100
```

- Latency KPIs: positive delta = regression (slower).
- Throughput KPIs: negative delta = regression (slower).
- RSS KPIs: positive delta = regression (more memory).

A KPI **passes** if `delta_pct` is within the tolerance column. A KPI **fails** if it exceeds tolerance.

### CI Integration

#### Trigger policy

Benchmarks do not run on every PR due to execution cost and environment variability. They run under these conditions:

- **Manual dispatch**: Maintainer triggers the benchmark workflow via `workflow_dispatch`.
- **Label trigger**: Adding the `bench` label to a PR triggers the benchmark job.
- **Release branches**: Benchmarks run automatically on PRs targeting `main` that modify files in `crates/sheetkit-core/src/` or `packages/sheetkit/src/`.

#### CI workflow steps

1. Generate fixtures if not cached.
2. Run Rust benchmarks (`cargo bench --bench open_modes -- --output-format bencher`).
3. Run Node.js benchmarks (`node --expose-gc --import tsx benchmarks/node/bench-baseline.ts`).
4. Compare results against stored baseline.
5. Post a summary comment on the PR with pass/fail per KPI.

#### Summary output format

The comparison script produces a markdown table:

```
| KPI | Fixture | Baseline | Current | Delta | Status |
|-----|---------|----------|---------|-------|--------|
| open_latency/lazy | large-data | - | 45ms | - | PASS (new) |
| open_latency/eager | large-data | 387ms | 392ms | +1.3% | PASS |
| peak_rss/lazy_open | large-data | - | 18MB | - | PASS (new) |
| getRows/lazy | large-data | 387ms | 410ms | +5.9% | PASS |
| save/untouched | large-data | 520ms | 245ms | -52.9% | PASS |
```

Status values: `PASS`, `FAIL`, `PASS (new)` (no baseline to compare against).

### Running Benchmarks Locally

**Rust**:

```bash
# Full benchmark suite
cargo bench --bench open_modes

# Single benchmark group
cargo bench --bench open_modes -- open_latency

# With verbose output
cargo bench --bench open_modes -- --verbose
```

**Node.js**:

```bash
# With GC exposure for accurate RSS measurement
node --expose-gc --import tsx benchmarks/node/bench-baseline.ts

# Quick spot-check (fewer runs, less accurate)
BENCH_RUNS=3 node --expose-gc --import tsx benchmarks/node/bench-baseline.ts
```

**Prerequisites**: Fixtures must be generated first. See [Benchmark Fixtures](../benchmarks/FIXTURES.md) for details.

## Next Steps

- [Getting Started](./getting-started.md) - Learn the basics
- [Architecture](./architecture.md) - Understand internal design
- [API Reference](./api-reference/) - Explore all available methods
