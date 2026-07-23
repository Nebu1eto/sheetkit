# SheetKit 1.0 Public Entry-Point Census

This document is the U1 baseline and U12 release-candidate census for the
security-hardening plan. The baseline revision is
`150e57939ac816a9bfa86daa856b432a135e3ca9`; the release-candidate column
describes the current `prepare-v1` working tree. It is evidence, not a progress
tracker or a substitute for the U13 migration classification.

Performance measurements gathered on the busy development machine remain
reference-only. The release performance gate must be repeated in the
user-requested idle, clean environment.

## Shared root controls

| Control | Finite policy or invariant | Public boundary |
|---|---|---|
| RC-U2 | `OpenLimits` and one operation-owned `OpenBudget` | ZIP path and buffer open |
| RC-U3 | `CryptLimits`, `VbaLimits`, retained decrypted-package accounting, and target-format macro rules | Password open/save and VBA inspection |
| RC-U4/U5/U9 | `FormulaLimits` and a shared `FormulaBudget` | Parse, evaluate, calculate, matching, and formula output |
| RC-U6 | Sparse-first `TraversalOptions` and finite `TraversalLimits` | Rows, columns, raw buffers, pivots, and render rectangles |
| RC-U7/U11 | Sink-specific CSV, JSON, SVG, HTML, and terminal policies | Returned output, CLI output, and conversion |
| RC-U8 | Pre-allocation package checks, bounded package-backed hydration, and finite stream batches | Lazy/filtered open and streaming readers |
| RC-U10 | Fallible cell references, strict transfer decoding, and exact package-part ownership | Coordinates, raw transfer, tables, and hydration |
| RC-U12 | `SaveLimits`, `StreamLimits`, atomic background tasks, and the structured napi error bridge | File/buffer save, streaming construction, and every Node resource-producing boundary |

No caller-selected profile can exceed a hard ceiling. A profile applies only to
the operation that receives it; there is no environment, process-global, or
mutable-workbook override.

## Untrusted package ingress

| Surface | Release-candidate entry points | Root control |
|---|---|---|
| Rust path open | `Workbook::open`, `open_with_options` | The no-options adapter calls finite `OpenOptions::default()`. The option-taking path resolves `OpenLimits`, creates one `OpenBudget`, and checks outer input before ZIP entry parsing. |
| Rust buffer open | `open_from_buffer`, `open_from_buffer_with_options` | The no-options adapter uses the same finite defaults; the option-taking path checks and charges the retained outer buffer before archive parsing. |
| Rust encrypted path/buffer open | `open_with_password`, `open_with_password_with_options`, `open_from_buffer_with_password`, `open_from_buffer_with_password_with_options` | The compatibility adapters use finite defaults. Explicit variants combine `OpenLimits`, `CryptLimits`, and `VbaLimits`; the same open budget continues into the decrypted ZIP without reset. |
| Deferred and filtered hydration | `PackageSource`, lazy worksheet access, deferred auxiliary access | The package source stores the resolved immutable `OpenLimits`; later XML reads use the stored per-part ceiling. Filtered worksheets keep package-backed raw data instead of unbounded typed hydration. |
| Rust streaming read | `open_sheet_reader`, `open_sheet_reader_owned`, reader `next_batch` | The reader inherits the workbook's worksheet spool ceiling; each batch validates a finite maximum before reserve or read. |
| Node path/buffer/password open | napi sync/async methods and TypeScript `openSync`, `open`, `openBufferSync`, `openBuffer`, `openWithPasswordSync`, `openWithPassword`, `openBufferWithPasswordSync`, `openBufferWithPassword` | `JsOpenOptions` maps to the same finite core options. Sync and async tasks retain structured resource metadata. The TypeScript wrapper supplies `lazy`/`deferred` when omitted. |
| CLI file ingress | `info`, `sheets`, `read`, `get`, `set`, `convert` | Every command opens through finite `Workbook::open`; there is no CLI unlimited-open switch. |

`maxUnzipSize` and `maxZipEntries` remain compatibility names, but they only
override the corresponding finite `OpenLimits` dimensions. `readMode`,
`auxParts`, and `dateInterpretation` reject unknown strings at the napi
conversion boundary.

## Formula and model ingress

| Surface | Release-candidate entry points | Root control |
|---|---|---|
| Direct parser | `sheetkit_core::formula::parse_formula` | Creates a finite parse budget before tokenization and AST construction. |
| Rust evaluation | `evaluate_formula`, `evaluate_formula_with_limits`, `calculate_all`, `calculate_all_with_limits` | Default adapters use `FormulaLimits::secure_default()`; explicit variants validate hard ceilings and share cumulative counters across parsing, graph construction, evaluation, matching, and output. Failed `calculate_all` stages results before commit. |
| Node evaluation | napi and TypeScript `evaluateFormula`, `calculateAll` | `JsFormulaLimits`/`FormulaLimits` use finite defaults, reject unsafe numeric values, and map to the same core operations. |
| Formula assignment | Rust/napi/TypeScript `set_cell_formula`/`setCellFormula` and `fill_formula`/`fillFormula` | Direct and bulk mutations use finite input, output, work, range, and cached-result limits. Fill operations preflight all shifted formulas before commit. Raw v1/v2 imports consume one non-reset formula budget through an opaque validated-transfer token. |
| Checked model coordinates | Cell getters/setters, ranges, validations, tables, pivots, drawings, hyperlinks, comments, controls, and defined names | Canonical A1 parsing and fallible coordinate construction enforce Excel row/column bounds. Rectangle-producing pivots and SVG render add finite traversal admission. |

Formula evaluation and mutation passed the focused core replay and the final
Node 24 integration replay.

## Traversal, transfer, and serialization

| Surface | Release-candidate entry points | Root control |
|---|---|---|
| Rust row traversal | `get_rows`, `get_rows_with_limits` | Sparse occupied rows are the default; cloned rows, cells, columns, and retained payload bytes are checked before result allocation. |
| Rust column traversal | `get_cols`, `get_cols_with_limits` | Sparse occupied columns use caller-selected finite limits and retained-value accounting. Exact-limit acceptance and plus-one rejection match row traversal. |
| Rust raw egress | `sheet_to_raw_buffer*`, `sheet_to_raw_buffer_v2*` | v1 defaults to sparse and permits only explicit finite dense traversal; v2 is sparse-only. Counts and encoded bytes use checked arithmetic before allocation. |
| Node row/column egress | native `getRows`, `getRowsBuffer`, `getRowsBufferV2`, `getCols`; wrapper `getRows`, `getRowsRaw`, `getRowsIterator`, `getRowsBuffer`, `getRowsBufferV2`, `getCols` | All adapters accept `TraversalOptions`; omitted options are finite and sparse. v1 is the only dense compatibility format. `getRows` and `getCols` apply the caller limits once and expose typed exact/plus-one failures. |
| TypeScript transfer consumers | `decodeRowsBuffer`, `decodeRowsRawBuffer`, `decodeRowsIterator`, `SheetData` and `SheetData.toArray` | v1/v2 headers, offsets, tags, UTF-8, dimensions, and trailing bytes are validated before exposure. Sparse consumers preserve absolute row/column coordinates; dense materialization remains explicit and finite. |
| Raw ingress | `raw_buffer_to_cells`, `cells_to_raw_buffer`, napi/TypeScript `setSheetDataBuffer` | Strict v1/v2 structural decoding, checked collection/string accounting, shifted-coordinate preflight, stale-SST validation, and an opaque validated-transfer token prevent partial mutation and formula-budget reset. |
| Streaming construction | Rust/napi `StreamWriter` row append and finish methods | `StreamLimits` bounds rows, cells, one row's XML, configuration, and total output. Checked admission occurs before publication; rejected rows remain unpublished, partial I/O poisons the writer, and a failed workbook guard does not consume it. |
| JSON/CSV/HTML | TypeScript `toJSON`, `toCSV`, `toHTML` | Each obtains data through finite transfer options. JSON uses own data properties for attacker-controlled headers; CSV has a separate returned-byte budget and formula policy; HTML performs context-specific escaping. |
| Pivot and SVG | Rust/napi/TypeScript `addPivotTable`, `renderToSvg` | Dense source/render rectangles are admitted against explicit finite traversal limits before proportional work; SVG uses a separate output ceiling. |

The default transfer ceiling is not an unlimited compatibility promise.
`TraversalLimits` also has non-overridable row, column, cell, and output-byte
hard ceilings. Coordinate distance cannot select sparse allocation size.

## Save, encryption, macro, and output sinks

| Surface | Release-candidate entry points | Root control |
|---|---|---|
| Rust file save | `save`, `save_with_limits`, `save_with_password`, `save_with_password_with_limits` | `SaveLimits` bounds archive output, XML parts, rows, and cells across hydrated, raw, lazy package-backed, and streamed sheets. File destinations use same-directory staging, flush/sync, and atomic replacement. |
| Rust buffer save | `save_to_buffer`, `save_to_buffer_with_limits` | The returned archive is constructed fallibly under the same finite save model and output limits while preserving the stored target format. |
| Node save | sync/async path, buffer, and password methods | `SaveLimits` and crypt limits map to the same core rules. Async save/write operations run as background tasks, retain structured limit metadata, preserve destination atomicity, and release the workbook for later mutation. |
| Macro conversion | normal and encrypted save to XLSX/XLSM/XLTX/XLTM/XLAM | The destination format selects the rule: macro-free formats remove VBA bytes, relationships, and content types; macro formats preserve canonical VBA content. |
| VBA inspection | `get_vba_project`, `get_vba_modules`, `get_vba_modules_with_limits`, direct `extract_vba_modules*`/`decompress_vba_stream`, and Node equivalents | Raw-project retention is bounded during open; default direct helpers create finite `VbaLimits`, and module extraction uses one budget for container, compressed, decoded, retained, and output totals. |
| Core/Node CSV | `serialize_csv*`, TypeScript `toCSV` | Spreadsheet-safe text is the default, raw exchange is a named opt-in, syntax cannot introduce formula prefixes, and returned UTF-8 bytes have a hard 64 MiB ceiling. |
| SVG | `render_to_svg`, `render_to_svg_with_limits`, Node `renderToSvg` | XML text, XML attributes, CSS strings, and colors use separate validation/encoding. Output is staged under a finite byte ceiling. |
| CLI human/JSON output | `info`, `sheets`, `read`, `get` | Human output visibly escapes C0/C1/DEL and is byte-bounded; JSON preserves structured values through a bounded writer. |
| CLI CSV and destination conversion | `read --format csv`, `convert --format csv` | Spreadsheet-safe is the default, `raw-exchange` is explicit, serialization is byte-bounded, and file conversion stages and syncs before replacement. |

## Rust facade, napi, and compatibility adapters

- `crates/sheetkit/src/lib.rs` re-exports `OpenLimits`, `OpenOptions`,
  `CryptLimits` when encryption is enabled, `VbaLimits`, `FormulaLimits`,
  `TraversalLimits`, `TraversalMode`, `TraversalOptions`, and the structured
  resource-error vocabulary. The facade does not add a separate unlimited
  profile.
- `toNativeOpenOptions` accepts the public TypeScript `OpenOptions` and the
  generated `JsOpenOptions`. Legacy `parseMode: "full"` maps to eager and
  `parseMode: "readfast"` maps to lazy behavior; invalid legacy modes reject.
  The adapter cannot remove finite archive/container limits.
- `maxUnzipSize` and `maxZipEntries` are finite compatibility overrides.
  Omission selects secure defaults rather than `u64::MAX`/`usize::MAX`.
- v1 raw buffers and dense v1 traversal are bounded compatibility formats.
  v2 and default traversal remain sparse.
- `raw-exchange` is an explicit data-exchange policy, not an unlimited resource
  profile: traversal and output ceilings still apply.
- TypeScript `fromJSON` validates canonical Excel bounds and finite row,
  column, cell, property-work, and retained-byte limits before materialization.
  It uses indexed array traversal, snapshots formula/date objects once, and
  passes the same admitted plain values to accounting and native mutation.

## Structured error bridge

Rust resource failures expose `Error::ResourceLimitExceeded` and stable
`ResourceLimitError` fields: code, operation, category, dimension, limit,
observed, and stage. TypeScript exposes the closed `ResourceLimitError` shape
and `isResourceLimitError`.

The dedicated napi bridge preserves these fields at all synchronous
resource-producing core boundaries. Background open, save, buffer-save, and
encrypted-save tasks retain metadata during worker execution and enrich the
rejected JavaScript error on the main thread. Plain lifecycle and argument
errors remain ordinary errors by design. Static review found no residual
resource-producing message-only boundary, and the final Node 24 replay covered
lazy hydration, formula, traversal, render, pivot, save, streaming, VBA, and
encryption paths.

## Freeze evidence and open rows

| Audit row | Current evidence | Freeze status |
|---|---|---|
| Finding shape | `v1-finding-ledger.md` parses to 77 unique reportable IDs: 1 High, 36 Medium, 40 Low; only ignored `SK-DS-075` is absent from 001–078 | Complete |
| Package/default parity | `open_budget_regression.rs`, `encrypted_open_budget_regression.rs`, `workbook_u3_security_regression.rs`, `open-options-defaults.spec.ts`, `resource-limit-error.spec.ts` | Implemented; final repository gate still required |
| Formula | Core formula replay passed 67/67; raw v1/v2 boundary replay passed 2/2; Node formula and aggregate security suites passed | Complete |
| Rows and columns | Core row/column unit tests, `security-traversal.spec.ts`, and `security-node-error-parity.spec.ts` cover finite sparse traversal | Complete |
| Transfer and SST | Core raw-transfer unit tests plus `security_fuzz_corpus_replay.rs` cover strict decoding, stale SST, cumulative formula limits, shifted coordinates, and adversarial structure | Complete |
| Stream writer | Core stream replay passed 6/6 plus poison 1/1; Node replay covers rejected-row nonpublication | Complete |
| Path and buffer save | Core save/encryption replay passed 5/5; Node sync/async replay covers typed limits, atomic destination preservation, and reuse | Complete |
| TypeScript `fromJSON` | `security-json-import.spec.ts` passed 22/22, including iterator, property-work, snapshot, canonical-bound, typed formula, and atomicity cases | Complete |
| Node error shape | Current-source release ESM addon focused replay passed 79/79 on Node 24.18.0; the complete Node suite passed 643/643 | Complete |
| Interoperability | The dated U12 receipt records Rust, CLI, fuzz, LibreOffice, msoffcrypto, macro, Node 24, and declaration evidence observed on 2026-07-23 | Complete for the local macOS arm64 release artifact |
| Fuzz corpus | `security_fuzz_corpus_replay.rs` records 3/3 tests, 33 seeds, and 2,048 deterministic mutations on 2026-07-23 | Complete; sanitizer/nightly fuzz remains supplementary |
| Performance | Existing U1/U2 reports | Reference-only; idle clean-machine rerun required |

The U12 security entry-point census is closed. U13 may use it as the source
inventory for old behavior, 1.0 behavior, breaking status, required action, and
bilingual guide destinations. The separate performance gate remains
reference-only until it is repeated on the requested idle, clean machine.
