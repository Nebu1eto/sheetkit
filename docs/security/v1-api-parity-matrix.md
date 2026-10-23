# SheetKit 1.0 API security parity matrix

Date: 2026-07-23

This matrix records the security-relevant 1.0 behavior shared by the Rust
facade, the napi binding, and the TypeScript wrapper. A current-source release
ESM addon was generated and replayed on Node 24.18.0 on 2026-07-23. Performance
measurements from the current busy machine are reference-only; release
performance must be measured again in the user-requested idle, clean
environment.

| Surface | Rust facade/core | Node.js | 1.0 behavior and evidence |
| --- | --- | --- | --- |
| Workbook open | Path, buffer, password, and option-taking variants | Sync/async path, buffer, and password variants | Secure finite `OpenOptions`; lazy reads and deferred auxiliary parts by default; invalid enum strings and unsafe numeric limits reject instead of falling back. `maxXmlPartBytes` also governs later lazy hydration. Covered by `open-options-defaults.spec.ts`, `security-node-error-parity.spec.ts`, and the passing Node 24 replay. |
| Compression-ratio floor | `OpenLimits::compression_ratio_floor_bytes`, `OpenOptions::compression_ratio_floor_bytes` | `OpenOptions.compressionRatioFloorBytes` | Same 8 MiB default and 1 GiB hard ceiling. Entries at or below the floor skip the 100:1 ratio check; larger entries are checked as a whole; non-empty entries with zero compressed bytes always reject. `0` enforces every entry. Values above the ceiling reject with dimension `compression_ratio_floor_bytes`; fractional, negative, and unsafe Node numbers reject as invalid arguments. Covered by `compat_compression_ratio_regression.rs` and `open-compression-ratio-floor.spec.ts`. |
| Traversal | `get_rows_with_limits`, `get_cols_with_limits`, raw v1/v2 options | `getRows`, `getCols`, buffer/iterator traversal options | Sparse and finite by default. Row, distinct-column, cell, and retained/output byte limits use exact-limit acceptance and plus-one rejection. Dense traversal is explicit v1 compatibility behavior; low-level sparse-only methods reject dense mode. Regression coverage is in core row/column tests, `security-traversal.spec.ts`, and `security-node-error-parity.spec.ts`. |
| Formula evaluation and mutation | `FormulaLimits`, limit-taking evaluation, calculation, and mutation variants | `FormulaLimits`/`JsFormulaLimits` on evaluation, calculation, and direct formula mutation | Secure finite defaults and caller-selected finite ceilings. Resource failures expose the same structured metadata; failed bulk calculation and mutation preserve prior values. Raw v1/v2 imports use one non-reset operation budget and cover per-formula 32 KiB and cumulative 4 MiB exact/plus-one boundaries. Cached formula results cover the cumulative 16 MiB output exact/plus-one boundary. |
| CSV output | Bounded CSV options and formula policy | `ToCsvOptions` | Spreadsheet-safe formula neutralization and 64 MiB hard output ceiling by default; raw exchange is explicit. Covered by CSV core/Node security tests and LibreOffice receipt. |
| Other output | Bounded render and transfer options | JSON/HTML/SVG/raw wrapper paths | Traversal/output ceilings are finite and typed limit failures remain distinguishable from invalid arguments. Covered by render, transfer, JSON, output, and streaming security suites. |
| Error contract | `ResourceLimitError` metadata | `ResourceLimitError` plus `isResourceLimitError` | Closed string codes and exact integer metadata are preserved; JavaScript exposes count/byte fields as `bigint`. Synchronous core-result boundaries and background open/save/write tasks preserve the same typed metadata, including lazy hydration. The Node 24 focused replay passed 79/79. |
| Declarations/runtime | `napi8` feature | Node 24.18.0, `@types/node` 24 | The generated `binding.d.ts` and public declarations compile together using `__test__/declarations/consumer.ts`. The current-source release addon passed the complete Node suite, 643/643. |
| Integer arguments (audit v2) | Typed integer parameters (`u32`, `u8`, `i32`), so fractional or non-finite values cannot be passed | Row, column, count, style, level, page-break, StreamWriter, view, page-setup, and numeric open options | Fractional, NaN, `Infinity`, `-Infinity`, unsafe (beyond `Number.MAX_SAFE_INTEGER`), and out-of-range numbers throw an invalid-argument error ("must be an integer from ... to ...") before mutation instead of being truncated, wrapped, or coerced to zero. Exact boundary values still pass, and in-range values keep the same domain errors as Rust (for example `invalid row number: 0`). Covered by `numeric-boundary.spec.ts` through both the wrapper and the raw binding. |
| HTML output limit (audit v2) | Not applicable; HTML export is Node-only | `toHTML` `limits.maxOutputBytes` (default 512 MiB) | The limit bounds the native transfer and the final UTF-8 HTML, escaping included. Exceeding it throws `ResourceLimitError` with operation `export`, category `output`, dimension `output_bytes`, and stage `serialize`. Exact-limit output passes and one byte less rejects. |
| Style and conditional-format enum strings (audit v2) | Typed enums | Fill pattern, border style, horizontal and vertical alignment, and conditional-format value type strings (case-insensitive), plus conditional-format `cellIs` operators | An unknown string throws an invalid-argument error that lists the accepted values (`unknown fill pattern 'polkadot'; expected one of: ...`) instead of falling back to a default. |
| StreamWriter merge cells (audit v2) | `StreamWriter::add_merge_cell` | `StreamWriter.addMergeCell` | Ranges are stored in canonical form (absolute markers dropped, uppercase, top-left to bottom-right). Re-adding an equivalent range is a no-op, and an overlapping range returns `MergeCellOverlap`. Covered by `stream.rs` unit tests and `core-contract.spec.ts`. |
| Stream reader truncated XML (audit v2) | `SheetStreamReader` | `openSheetReader` batches | A worksheet that ends inside an open row, cell, or value element is an XML parse error, not a shorter row list. Covered by `stream_reader.rs` unit tests and `core-contract.spec.ts`. |
| Table names (audit v2) | `add_table`, `copy_sheet` | `addTable`, `copySheet` | Names and display names must be 1 to 255 characters, start with a letter, underscore, or backslash, continue with letters, digits, periods, or underscores, and must not look like an A1 or R1C1 reference. An empty display name falls back to the table name. Names are unique case-insensitively across table names, display names, and defined names. A clash with a table is `TableAlreadyExists`; a clash with a defined name or an invalid name is `InvalidArgument`. `copy_sheet` generates collision-free table names and rejects an invalid generated name without mutation. Covered by `audit_table_regression.rs`, `audit_sheet_structural.rs`, and `core-contract.spec.ts`. |

The TypeScript declaration fixture is a release gate:

```sh
pnpm --filter @sheetkit/node typecheck:declarations
```

The runtime parity result is also recorded in
`docs/plans/artifacts/2026-07-15-001-fix-security-hardening-v1/u12-interop/results.md`
for the same 2026-07-23 replay.
