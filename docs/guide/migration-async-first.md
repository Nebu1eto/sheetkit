# Migration Guide: Async-First Lazy-Open

This guide covers the breaking changes and new APIs introduced by the async-first lazy-open refactor. All changes apply to the Node.js (`@sheetkit/node`) package. The Rust crate API is not affected.

## Summary of Changes

### Breaking Changes

1. **`ParseMode` replaced by `ReadMode`**: The `full`/`readfast` mode names are replaced by `eager`/`lazy`/`stream`.
2. **`AuxParts` policy**: New option controlling when auxiliary parts (comments, charts, images) are parsed.
3. **`OpenOptions` restructured**: New typed options interface with `readMode`, `auxParts`, `sheets`, `sheetRows`.
4. **`getRowsIterator` replaced by `openSheetReader`**: Stream-based reader with async iterator protocol.
5. **Default read mode changed**: `open()` now defaults to `readMode: 'lazy'` instead of eagerly parsing all sheets.

### New APIs

1. `Workbook.openSheetReader(sheet, opts?)` -- returns a `SheetStreamReader`
2. `SheetStreamReader` class with `for await...of` support
3. `getRowsBufferV2(sheet)` -- v2 raw buffer with inline strings
4. `ReadMode` type: `'eager' | 'lazy' | 'stream'`
5. `AuxParts` type: `'eager' | 'deferred'`

## ReadMode Mapping

| Old (`parseMode`) | New (`readMode`) | Behavior |
|-------------------|-----------------|----------|
| `'full'` | `'eager'` | Parse all sheets and auxiliary parts during open. Same behavior as before. |
| `'readfast'` | `'lazy'` | Parse ZIP index and metadata only. Sheet XML is parsed on first access (e.g., `getRows()`). |
| (no equivalent) | `'stream'` | Minimal parse. Use with `openSheetReader()` for forward-only bounded-memory iteration. |

### Before (v0.4)

```typescript
const wb = await Workbook.open('large.xlsx', {
  parseMode: 'readfast',
  sheetRows: 1000,
});
```

### After (v0.5)

```typescript
const wb = await Workbook.open('large.xlsx', {
  readMode: 'lazy',
  sheetRows: 1000,
});
```

If you were using `parseMode: 'full'` or no `parseMode`, the equivalent is `readMode: 'eager'`. However, consider switching to `readMode: 'lazy'` for improved open latency and lower memory usage. Lazy mode transparently hydrates sheets on first access, so most existing code works without further changes.

## OpenOptions Changes

The new `OpenOptions` interface:

```typescript
interface OpenOptions {
  readMode?: 'lazy' | 'stream' | 'eager';  // default: 'lazy'
  auxParts?: 'deferred' | 'eager';          // default: 'deferred'
  sheets?: string[];
  sheetRows?: number;
  maxUnzipSize?: number;
  maxZipEntries?: number;
}
```

Key differences from the previous `JsOpenOptions`:
- `parseMode` is removed. Use `readMode` instead.
- `auxParts` is new. It controls whether comments, charts, images, and other auxiliary parts are parsed during open (`'eager'`) or deferred until first access (`'deferred'`).
- The old `JsOpenOptions` type is still accepted for backward compatibility, but `parseMode` values are not recognized by the new code path.

### AuxParts Policy

When `auxParts` is `'deferred'` (the default), auxiliary parts like comments, charts, images, and pivot tables are not parsed during `open()`. They are loaded on-demand the first time you call a method that needs them (e.g., `getComments()`, `getPivotTables()`).

This reduces open latency and memory usage for workloads that only need cell data.

```typescript
// Deferred (default): comments are not parsed during open
const wb = await Workbook.open('report.xlsx');
// Comments are loaded on first access:
const comments = wb.getComments('Sheet1');

// Eager: all parts parsed during open (old behavior)
const wb2 = await Workbook.open('report.xlsx', { auxParts: 'eager' });
```

## Streaming Reader Migration

### Before: `getRowsIterator()`

The old `getRowsIterator()` was a synchronous generator that materialized the entire sheet buffer first, then yielded rows one at a time. It reduced peak JS object count but not peak Rust memory.

```typescript
const wb = Workbook.openSync('large.xlsx');
for (const row of wb.getRowsIterator('Sheet1')) {
  process(row);
}
```

### After: `openSheetReader()`

The new `openSheetReader()` returns a `SheetStreamReader` that reads rows in batches directly from the ZIP entry without materializing the entire sheet. Memory usage is bounded by the batch size.

```typescript
const wb = await Workbook.open('large.xlsx', { readMode: 'stream' });
const reader = await wb.openSheetReader('Sheet1', { batchSize: 500 });

for await (const batch of reader) {
  for (const row of batch) {
    process(row);
  }
}
```

Key differences:
- `openSheetReader()` is async and returns a `SheetStreamReader`.
- The reader implements the async iterator protocol (`for await...of`).
- Each iteration yields an array of rows (a batch), not a single row.
- Memory usage is bounded: only one batch of rows is in memory at a time.
- Call `reader.close()` to release resources early, or let the `for await` loop handle it automatically.

### Manual Batch Control

```typescript
const reader = await wb.openSheetReader('Sheet1');

while (true) {
  const batch = await reader.next(1000); // custom batch size per call
  if (batch === null) break;
  for (const row of batch) {
    process(row);
  }
}

await reader.close();
```

## Raw Buffer V2

`getRowsBufferV2()` produces a v2 binary buffer that inlines string data instead of using a separate global string table. This enables incremental row-by-row decoding without eagerly materializing all strings.

The v1 format (`getRowsBuffer()`) is still available and fully supported.

```typescript
// V1 (unchanged)
const bufV1 = wb.getRowsBuffer('Sheet1');

// V2 (new): inline strings, better for streaming decoders
const bufV2 = wb.getRowsBufferV2('Sheet1');
```

Use v2 when building custom streaming decoders or when you need to decode rows incrementally. For most use cases, the higher-level APIs (`getRows()`, `getRowsRaw()`, `SheetData`) handle format selection automatically.

## Copy-on-Write Save

Workbooks opened in `lazy` mode use a copy-on-write save pipeline. When you save a lazy-opened workbook without modifying any sheets, unchanged sheet XML is written directly from the original ZIP entry without being parsed and re-serialized. This reduces save latency for workflows that modify only a subset of sheets.

```typescript
// Open large file, modify one sheet, save
const wb = await Workbook.open('100-sheets.xlsx');
wb.setCellValue('Sheet1', 'A1', 'Updated');
await wb.save('output.xlsx');
// Only Sheet1 is re-serialized. The other 99 sheets are copied as-is.
```

This is transparent and requires no code changes. Modified sheets are tracked automatically via per-sheet dirty flags.

## Full Migration Checklist

1. Replace `parseMode: 'full'` with `readMode: 'eager'` (or remove it to use the default `'lazy'`).
2. Replace `parseMode: 'readfast'` with `readMode: 'lazy'`.
3. Replace `getRowsIterator()` with `openSheetReader()` and update to async iteration.
4. If you rely on all auxiliary parts being available immediately after open, add `auxParts: 'eager'` to your open options.
5. Test that cell-access methods still work -- lazy mode hydrates sheets transparently.
6. Consider using `readMode: 'stream'` with `openSheetReader()` for very large files where you do not need random cell access.

## Compatibility Notes

- The old `JsOpenOptions` type with `sheetRows`, `sheets`, `maxUnzipSize`, and `maxZipEntries` fields is still accepted by all open methods. Only `parseMode` is not recognized.
- `getRowsIterator()` still exists and works. It is not removed, but `openSheetReader()` is preferred for new code because it provides true bounded-memory streaming.
- `getRowsBuffer()` (v1) is not deprecated. V1 and V2 formats coexist.
- All sync methods (`openSync`, `saveSync`, `openBufferSync`, `writeBufferSync`) continue to work unchanged.
