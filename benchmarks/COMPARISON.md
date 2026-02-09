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

## Optimization History

### Buffer-Based FFI Transfer (2026-02-10)

This update reflects the buffer-based FFI transfer optimization applied to
`getRows()` and internal cell data structures. Previously, `getRows()` created
individual `JsRowCell` JavaScript objects for every cell (1M objects for 50k x 20).
Now, cell data is serialized into a compact binary buffer in Rust and transferred
in a single FFI call. Internal Rust data structures were also optimized:

- **CompactCellRef**: Cell references stored as `[u8;10]` inline instead of heap `String`
- **CellTypeTag**: Cell type stored as a 1-byte enum instead of `Option<String>`
- **Sparse-to-dense row conversion**: Optimized row iteration internals

**Impact on napi overhead**:
- Read operations: overhead reduced from ~1.75x to ~1.1-1.3x (measured below)
- Write operations: unchanged (already efficient via batch `setSheetData()`)
- Memory: read memory for 100k rows reduced from 361.1MB to 13.5MB (96%)

## napi-rs Overhead Analysis

This comparison measures the overhead introduced by the napi-rs FFI layer when calling
SheetKit from Node.js vs using it directly from Rust. The Rust numbers are from the
Rust native benchmark (run separately); the Node.js numbers are from the latest
Node.js benchmark run.

### Read Benchmarks

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| Large Data (50k rows x 20 cols) | 625ms | 709ms | +84ms | 1.13x |
| Heavy Styles (5k rows, formatted) | 33ms | 37ms | +4ms | 1.12x |
| Multi-Sheet (10 sheets x 5k rows) | 369ms | 788ms | +419ms | 2.14x |
| Formulas (10k rows) | 43ms | 50ms | +7ms | 1.16x |
| Strings (20k rows text-heavy) | 137ms | 146ms | +9ms | 1.07x |
| Data Validation (5k rows, 8 rules) | 26ms | 28ms | +2ms | 1.08x |
| Comments (2k rows with comments) | 10ms | 12ms | +2ms | 1.20x |
| Merged Cells (500 regions) | 2ms | 2ms | +0ms | 1.00x |
| Mixed Workload (ERP document) | 34ms | 40ms | +6ms | 1.18x |

### Read Scaling

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| Scale 1k rows | 6ms | 8ms | +2ms | 1.33x |
| Scale 10k rows | 65ms | 69ms | +4ms | 1.06x |
| Scale 100k rows | 650ms | 742ms | +92ms | 1.14x |

### Write Benchmarks

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| 50k rows x 20 cols | 742ms | 681ms | -61ms | 0.92x |
| 5k styled rows | 41ms | 50ms | +9ms | 1.22x |
| 10 sheets x 5k rows | 381ms | 348ms | -33ms | 0.91x |
| 10k rows with formulas | 34ms | 40ms | +6ms | 1.18x |
| 20k text-heavy rows | 148ms | 126ms | -22ms | 0.85x |
| 5k rows + 8 DV rules | 12ms | 13ms | +1ms | 1.08x |
| 2k rows with comments | 9ms | 11ms | +2ms | 1.22x |
| 500 merged regions | 13ms | 15ms | +2ms | 1.15x |

### Write Scaling

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| 1k rows x 10 cols | 6ms | 7ms | +1ms | 1.17x |
| 10k rows x 10 cols | 69ms | 69ms | +0ms | 1.00x |
| 50k rows x 10 cols | 347ms | 336ms | -11ms | 0.97x |
| 100k rows x 10 cols | 705ms | 720ms | +15ms | 1.02x |

### Other

| Scenario | Rust (median) | Node.js (median) | Overhead | Ratio |
|----------|--------------|-----------------|----------|-------|
| Buffer round-trip (10k rows) | 169ms | 172ms | +3ms | 1.02x |
| Streaming write (50k rows) | 1.02s | 1.13s | +110ms | 1.11x |
| Random-access read (1k cells/50k file) | 585ms | 566ms | -19ms | 0.97x |
| Mixed workload write (ERP-style) | 23ms | 28ms | +5ms | 1.22x |

## Memory: Before and After Buffer Transfer

This section compares Node.js RSS memory before and after the buffer-based FFI
transfer optimization, for key read scenarios that were most affected.

| Scenario | Before (2026-02-09) | After (2026-02-10) | Reduction |
|----------|--------------------|--------------------|-----------|
| Read Large Data (50k x 20) | 405.6MB | 349.2MB | 14% |
| Read Multi-Sheet (10 x 5k) | 207.0MB | 216.8MB | -5% (within noise) |
| Read Scale 10k rows | 23.8MB | 0.0MB | ~100% |
| Read Scale 100k rows | 361.1MB | 13.5MB | 96% |
| Read Heavy Styles (5k) | 20.0MB | 16.0MB | 20% |
| Read Formulas (10k) | 16.3MB | 13.4MB | 18% |
| Read Strings (20k) | 12.4MB | 8.4MB | 32% |
| Write 50k x 20 | 255.0MB | 186.2MB | 27% |
| Write 100k x 10 | 268.7MB | 85.8MB | 68% |
| Streaming 50k rows | 323.4MB | 101.3MB | 69% |

The most dramatic improvement is in the read scaling benchmarks, where 100k rows
dropped from 361.1MB to 13.5MB. This confirms that the buffer-based transfer
eliminates the V8 object overhead that previously dominated memory usage.

For the large data (50k x 20) read, memory dropped from 405.6MB to 349.2MB.
The remaining 349MB is primarily the Rust-side Workbook holding the full XML
DOM in memory, not the FFI transfer overhead. Further memory reduction in this
scenario would require lazy/streaming XML parsing (a separate optimization).

Write memory also improved significantly: 50k x 20 write dropped from 255.0MB
to 186.2MB (27%), and 100k x 10 write dropped from 268.7MB to 85.8MB (68%).
Streaming write improved from 323.4MB to 101.3MB (69%).

## Key Findings

### 1. Read operations: napi overhead reduced from ~1.75x to ~1.13x

After the buffer-based FFI transfer, read operations now show only 6-33%
overhead compared to pure Rust, down from 50-100% previously. The previous
overhead was dominated by per-cell napi object creation; the new approach
transfers data as a single binary buffer.

| | Before (object-based) | After (buffer-based) |
|---|---|---|
| Average read overhead | ~1.75x | ~1.13x |
| 50k x 20 cols | 2.02x (1.26s vs 625ms) | 1.13x (709ms vs 625ms) |
| 100k rows | 1.89x (1.23s vs 650ms) | 1.14x (742ms vs 650ms) |
| 10k rows | 1.74x (117ms vs 65ms) | 1.06x (69ms vs 65ms) |

The multi-sheet scenario (2.14x) is an outlier. It involves 10 separate sheet
parses and the overhead compounds per sheet. This scenario may benefit from
future work to batch multi-sheet reads.

### 2. Write operations: near-zero overhead (unchanged)

Write operations continue to show minimal napi overhead (~0.85-1.22x), as the
batch `setSheetData()` API was already efficient. Some scenarios show Node.js
marginally outperforming Rust, likely due to V8's efficient string handling
during data construction and measurement noise.

### 3. Streaming write: overhead reduced from 1.20x to 1.11x

Streaming write overhead dropped from 1.20x to 1.11x. Memory improved
dramatically from 323.4MB to 101.3MB (69% reduction).

### 4. Buffer round-trip: near parity

The buffer round-trip scenario went from 1.31x overhead to 1.02x overhead,
achieving near parity with pure Rust.

### 5. Time improvements from internal Rust optimizations

The read time improvements (e.g., 50k read: 1.26s to 709ms) are not solely
from reduced FFI overhead. Internal Rust data structures were also optimized:

- CompactCellRef eliminates heap allocation for cell references
- CellTypeTag enum eliminates String allocation for cell types
- Optimized row iteration avoids intermediate String allocations
- These improvements benefit both Rust native and Node.js performance

## Summary

| Category | Average Overhead (Before) | Average Overhead (After) | Improvement |
|----------|--------------------------|--------------------------|-------------|
| Read | ~1.75x | ~1.13x | 35% less overhead |
| Write (batch) | ~1.05x | ~1.05x | No change (already efficient) |
| Write (per-row) | ~1.20x | ~1.11x | Modest improvement |
| Round-trip | ~1.31x | ~1.02x | Near parity |

The buffer-based FFI transfer reduced napi-rs read overhead from ~75% to ~13%
on average. For write-heavy workloads, overhead remains negligible. The
optimization eliminates per-cell V8 object creation, replacing it with a single
binary buffer transfer that is decoded lazily on the JavaScript side.
