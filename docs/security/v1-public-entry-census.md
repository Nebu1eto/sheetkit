# SheetKit 1.0 Public Entry-Point Census

This document is U1 evidence for the security-hardening plan. It records the
public boundaries present at scan revision `150e579` and is not a progress
tracker. Paths below were verified against that revision's workspace layout.

## Untrusted package ingress

- Core path and buffer opens are `Workbook::open`, `open_with_options`,
  `open_from_buffer`, and `open_from_buffer_with_options` in
  `crates/sheetkit-core/src/workbook/io.rs`. Their policy type and compatibility
  defaults are in `crates/sheetkit-core/src/workbook/open_options.rs`.
- Encrypted path ingress is `Workbook::open_with_password` in
  `crates/sheetkit-core/src/workbook/io.rs`; CFB parsing, password verification,
  decryption, and VBA extraction continue through
  `crates/sheetkit-core/src/crypt/` and `crates/sheetkit-core/src/vba.rs`.
- Deferred/lazy worksheet ingress and later hydration are owned by
  `crates/sheetkit-core/src/workbook/source.rs`,
  `crates/sheetkit-core/src/workbook/aux_parts.rs`, and
  `crates/sheetkit-core/src/workbook/mod.rs`.
- Streaming worksheet XML ingress is exposed by `open_sheet_reader` and
  `open_sheet_reader_owned` in `crates/sheetkit-core/src/workbook/mod.rs`, backed
  by `crates/sheetkit-core/src/stream_reader.rs`.
- Binary raw-row ingress is `raw_buffer_to_cells` in
  `crates/sheetkit-core/src/raw_transfer_write.rs`; the v1 and v2 producers and
  readers are in `crates/sheetkit-core/src/raw_transfer.rs` and
  `crates/sheetkit-core/src/raw_transfer_v2.rs`.

## Formula and model ingress

- Formula text enters through `set_cell_formula` and `fill_formula` in
  `crates/sheetkit-core/src/workbook/cell_ops.rs`, and through
  `evaluate_formula`/`calculate_all` in
  `crates/sheetkit-core/src/workbook/data.rs`.
- Parsing, dependency-graph construction, range expansion, and evaluation are
  public or transitively reachable in `crates/sheetkit-core/src/formula/`.
- Untrusted coordinates, ranges, hyperlinks, images, charts, validations,
  defined names, tables, pivots, slicers, comments, and controls enter through
  the public `Workbook` methods distributed across
  `crates/sheetkit-core/src/workbook/cell_ops.rs`, `sheet_ops.rs`, `data.rs`,
  `drawing.rs`, and `features.rs`.

## Traversal and serialization egress

- Sparse and dense row traversal starts at `Workbook::get_rows` in
  `crates/sheetkit-core/src/workbook/sheet_ops.rs`; streaming traversal is in
  `crates/sheetkit-core/src/stream_reader.rs`, and streaming construction is in
  `crates/sheetkit-core/src/stream.rs`.
- File, buffer, and encrypted outputs are `save`, `save_to_buffer`, and
  `save_with_password` in `crates/sheetkit-core/src/workbook/io.rs`.
- SVG output is `Workbook::render_to_svg` in
  `crates/sheetkit-core/src/workbook/features.rs`, implemented by
  `crates/sheetkit-core/src/render.rs`.
- Raw row egress is implemented by `crates/sheetkit-core/src/raw_transfer.rs`,
  `raw_transfer_v2.rs`, and `raw_transfer_write.rs`.

## Rust facade and CLI

- `crates/sheetkit/src/lib.rs` re-exports the core `Workbook`, `OpenOptions`,
  `StreamWriter`, rendering options, VBA types, and the value/configuration
  types that make the core boundaries public to Rust consumers.
- `crates/sheetkit/src/bin/main.rs` exposes `info`, `sheets`, `read`, `get`,
  `set`, and `convert`. All file-taking commands use the default path open;
  `read` emits table or CSV text, `convert` writes CSV, and the other read
  commands write human-readable terminal text.

## napi Rust boundary

- `packages/sheetkit/src/lib.rs` exposes sync and async path/buffer/password
  opens; sync and async path/buffer/encrypted saves; row/raw-buffer APIs;
  formula APIs; streaming readers and writers; VBA inspection; and SVG output.
- `packages/sheetkit/src/types.rs` defines the napi option and payload objects,
  including `JsOpenOptions`, while `packages/sheetkit/src/conversions.rs` maps
  them to core types and retains the legacy `parse_mode` adapter.
- Native errors are currently converted at the individual call sites in
  `packages/sheetkit/src/lib.rs`, `stream.rs`, and `conversions.rs`.

## TypeScript boundary and compatibility adapters

- `packages/sheetkit/index.ts` wraps every native workbook boundary: path,
  buffer, and password opens; path, buffer, and password outputs; formula
  mutation/evaluation; row, raw-buffer, and async-stream traversal; JSON, CSV,
  HTML, and SVG output; JSON import; and VBA inspection.
- `toNativeOpenOptions` in `packages/sheetkit/index.ts` accepts both the public
  `OpenOptions` and legacy generated `JsOpenOptions`; the latter can carry
  `parseMode`, which is normalized by `packages/sheetkit/src/conversions.rs`.
- `packages/sheetkit/buffer-codec.ts` directly decodes native v1/v2 row buffers
  for array, raw, and iterator consumers. `packages/sheetkit/sheet-data.ts`
  exposes random access and dense-array materialization over those buffers.

## Seven parity and security gaps at census time

1. **Default package limits are not finite.** Core `OpenOptions` leaves ZIP
   size and entry limits optional, the napi/TypeScript options preserve omitted
   values, and CLI callers cannot select a policy. Therefore equivalent default
   path and buffer opens do not yet meet the 1.0 finite-default contract.
2. **Encrypted opens bypass caller policy.** Rust, napi, and TypeScript password
   opens accept no `OpenOptions`; core decrypts into a complete in-memory buffer
   and then uses `OpenOptions::default()`. The encrypted surface cannot express
   the same admission controls as plain path/buffer opens.
3. **Lazy and streaming work can escape the open operation.** Deferred package
   hydration and `open_sheet_reader` perform parsing after initial open, while
   the current option object does not provide an immutable policy snapshot plus
   fresh operation-scoped budget for those later public calls.
4. **Raw and traversal adapters expose independent allocation boundaries.** Core
   raw decoding, TypeScript buffer decoding/iteration, `SheetData.toArray`,
   `get_rows`, and stream batch sizing are separate entry points; they do not
   yet share finite row/cell/string/gap dimensions or a sparse-first default.
5. **Formula controls are not cross-surface policy.** Parse, direct evaluation,
   dependency construction, `calculate_all`, and formula assignment are exposed
   through both Rust and Node.js without a common operation-scoped formula
   budget or stable limit errors.
6. **Output safety depends on the sink.** Core/napi SVG, TypeScript JSON/CSV/HTML,
   and CLI table/CSV/human output currently use separate ad hoc formatting.
   They lack frozen sink-specific policies for XML/CSS, prototype-sensitive
   keys, spreadsheet formulas, HTML, and terminal control characters.
7. **Errors and compatibility behavior lose parity at adapters.** napi generally
   reduces core failures to message-only `Error::from_reason` values, CLI prints
   display text, and adapters silently fall back for unknown values. There is
   also a confirmed mode-parity bug: TypeScript permits `readMode: "eager"`, but
   `js_open_options_to_core` in `packages/sheetkit/src/conversions.rs` matches
   only `"lazy"` and `"stream"`, so `"eager"` falls back to the core default
   `Lazy`; the legacy `parseMode: "full"` path does select `Eager`. Stable
   resource categories, dimensions, modes, and bounded metadata therefore do
   not round-trip consistently across Rust, Node.js, and CLI.

U12 must re-run this census after implementation and demonstrate that every
listed adapter inherits the same secure default or an explicit finite profile;
no alternate entry point may bypass a shared control.
