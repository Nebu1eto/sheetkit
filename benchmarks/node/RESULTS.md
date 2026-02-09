# Excel Library Benchmark: SheetKit vs ExcelJS vs SheetJS

Benchmark run: 2026-02-09
Platform: Linux x64
Node.js: v22.22.0

## Libraries Tested

| Library | Version | Description |
|---------|---------|-------------|
| **SheetKit** (`@sheetkit/node`) | 0.2.0 | Rust-based Excel library with Node.js bindings via napi-rs |
| **ExcelJS** (`exceljs`) | 4.4.0 | Pure JavaScript Excel library with streaming support |
| **SheetJS** (`xlsx`) | 0.18.5 | Pure JavaScript spreadsheet library (community edition) |

## Test Fixtures

All fixtures were generated programmatically using SheetKit.

| Fixture | Rows | Cols | Size | Description |
|---------|------|------|------|-------------|
| `large-data.xlsx` | 50,000 | 20 | 5.1 MB | Mixed types: numbers, strings, floats, booleans |
| `heavy-styles.xlsx` | 5,000 | 10 | 255 KB | Rich formatting: fonts, fills, borders, number formats, alignment |
| `multi-sheet.xlsx` | 50,000 total | 10 | 2.8 MB | 10 sheets, each with 5,000 rows |
| `formulas.xlsx` | 10,000 | 7 | 462 KB | 5 formula columns (SUM, PRODUCT, AVERAGE, MAX, IF) |
| `strings.xlsx` | 20,000 | 10 | 898 KB | Text-heavy data (shared string table stress test) |

## Results

### Read Performance

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| Large Data (50k rows x 20 cols) | 3.86s | 3.12s | 4.00s | ExcelJS |
| Heavy Styles (5k rows, formatted) | **100ms** | 374ms | 247ms | SheetKit (3.7x) |
| Multi-Sheet (10 sheets x 5k rows) | **1.07s** | 3.37s | 1.68s | SheetKit (1.6x) |
| Formulas (10k rows) | **110ms** | 403ms | 191ms | SheetKit (1.7x) |
| Strings (20k text-heavy rows) | **409ms** | 1.29s | 632ms | SheetKit (1.5x) |

SheetKit wins **4 out of 5** read scenarios. For the large-data case, ExcelJS is slightly faster -- likely because SheetKit's `getRows()` materializes all data through the napi boundary in one call, and for very large row counts, the overhead of converting 1M+ cells from Rust to JS objects becomes visible.

### Write Performance (cell-by-cell API)

| Scenario | SheetKit | ExcelJS | SheetJS | Winner |
|----------|----------|---------|---------|--------|
| 50k rows x 20 cols | 46.12s | 7.38s | **2.86s** | SheetJS |
| 5k styled rows | 236ms | 450ms | **112ms** | SheetJS |
| 10 sheets x 5k rows | 1.92s | 3.57s | **1.08s** | SheetJS |
| 10k rows with formulas | 225ms | 382ms | **160ms** | SheetJS |
| 20k text-heavy rows | 1.45s | 1.34s | **529ms** | SheetJS |

SheetJS wins all write scenarios when using per-cell APIs. This is expected: SheetJS builds an in-memory JS object directly, while SheetKit's `setCellValue()` crosses the napi-rs FFI boundary on every call. For 50k x 20 = 1 million cells, that's 1M FFI round-trips.

**Key insight:** SheetKit's cell-by-cell write API is not suited for bulk data generation. Use the streaming writer instead (see below).

### Streaming Write

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| 50k rows x 20 cols | 3.19s | **1.57s** | N/A |

SheetKit's streaming writer (`JsStreamWriter.writeRow()`) reduces 50k-row write time from **46s to 3.2s** (14x improvement) by batching data per-row instead of per-cell. ExcelJS's streaming writer is still faster because it writes directly to a file stream without building a full in-memory workbook.

SheetJS has no streaming API.

### Buffer Round-Trip (write + read back from memory)

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| 10k rows x 10 cols | 602ms | 1.15s | **431ms** |

### Memory Usage (heap delta during operation)

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| Read Large Data | **0.1 MB** | 1.5 MB | 30.4 MB |
| Read Multi-Sheet | **0.0 MB** | 0.2 MB | 31.2 MB |
| Read Strings | **0.0 MB** | 0.1 MB | 9.0 MB |
| Write 50k rows | **0.0 MB** | 299.1 MB | 0.2 MB |
| Write 10 sheets | **0.0 MB** | 136.1 MB | 0.1 MB |
| Write 20k strings | **0.0 MB** | 45.7 MB | 0.0 MB |

SheetKit keeps JS heap usage near zero because all data lives in Rust memory. ExcelJS shows the highest memory usage as it stores everything in JavaScript objects. SheetJS is moderate for reads but efficient for writes.

### Output File Size

| Scenario | SheetKit | ExcelJS | SheetJS |
|----------|----------|---------|---------|
| Write 50k rows | **5.8 MB** | 6.2 MB | 33.8 MB |
| Write styled rows | **0.2 MB** | 0.3 MB | 1.6 MB |
| Write 10 sheets | **2.8 MB** | 2.9 MB | 17.4 MB |
| Write formulas | **0.3 MB** | 0.3 MB | 2.1 MB |
| Write strings | 0.8 MB | **0.8 MB** | 10.3 MB |
| Streaming 50k rows | **5.8 MB** | 6.0 MB | N/A |

SheetKit produces the smallest files. SheetJS (community edition) produces significantly larger files because it does not apply deflate compression by default and does not use shared strings.

## Summary

| Category | Winner | Notes |
|----------|--------|-------|
| **Read speed** | SheetKit (4/5) | 1.5x-3.7x faster than alternatives on most workloads |
| **Write speed (cell API)** | SheetJS (5/5) | SheetKit's per-cell napi calls are costly for bulk writes |
| **Write speed (streaming)** | ExcelJS | True streaming-to-disk architecture |
| **Memory efficiency** | SheetKit | Near-zero JS heap; data stays in Rust |
| **File size** | SheetKit | Best compression; proper shared string table |
| **Style support** | SheetKit / ExcelJS | SheetJS community edition lacks style write support |

### Wins by Library

| Library | Wins (speed) | Strengths |
|---------|-------------|-----------|
| **SheetKit** | 4/12 | Read performance, memory efficiency, file size, full OOXML feature set |
| **SheetJS** | 6/12 | Fast bulk writes via in-memory JS objects, minimal API surface |
| **ExcelJS** | 2/12 | Streaming write, large-data reads, mature ecosystem |

### Recommendations

- **Read-heavy workloads:** SheetKit is the best choice -- fastest reads with minimal memory.
- **Bulk data generation:** Use SheetKit's `JsStreamWriter` for large writes instead of `setCellValue()`.
- **Quick data export (no formatting needed):** SheetJS is fast if you don't need styles.
- **Streaming to disk:** ExcelJS excels when you need to stream very large files without holding them in memory.

## Reproducing

```bash
cd benchmarks/node
pnpm install
pnpm generate    # Generate test fixtures
pnpm bench       # Run benchmarks
```
