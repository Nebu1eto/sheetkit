# SheetKit 1.0 API security parity matrix

Date: 2026-07-23

This matrix records the security-relevant 1.0 behavior shared by the Rust
facade, the napi binding, and the TypeScript wrapper. A current-source release
ESM addon was generated and replayed on Node 24.18.0 on 2026-07-23. Performance
measurements from the current busy machine are reference-only; release
performance must be measured again in the user-requested idle, clean
environment.

| Surface | Rust facade/core | Node.js | 1.0 behavior and evidence |
|---|---|---|---|
| Workbook open | Path, buffer, password, and option-taking variants | Sync/async path, buffer, and password variants | Secure finite `OpenOptions`; lazy reads and deferred auxiliary parts by default; invalid enum strings and unsafe numeric limits reject instead of falling back. `maxXmlPartBytes` also governs later lazy hydration. Covered by `open-options-defaults.spec.ts`, `security-node-error-parity.spec.ts`, and the passing Node 24 replay. |
| Traversal | `get_rows_with_limits`, `get_cols_with_limits`, raw v1/v2 options | `getRows`, `getCols`, buffer/iterator traversal options | Sparse and finite by default. Row, distinct-column, cell, and retained/output byte limits use exact-limit acceptance and plus-one rejection. Dense traversal is explicit v1 compatibility behavior; low-level sparse-only methods reject dense mode. Regression coverage is in core row/column tests, `security-traversal.spec.ts`, and `security-node-error-parity.spec.ts`. |
| Formula evaluation and mutation | `FormulaLimits`, limit-taking evaluation, calculation, and mutation variants | `FormulaLimits`/`JsFormulaLimits` on evaluation, calculation, and direct formula mutation | Secure finite defaults and caller-selected finite ceilings. Resource failures expose the same structured metadata; failed bulk calculation and mutation preserve prior values. Raw v1/v2 imports use one non-reset operation budget and cover per-formula 32 KiB and cumulative 4 MiB exact/plus-one boundaries. Cached formula results cover the cumulative 16 MiB output exact/plus-one boundary. |
| CSV output | Bounded CSV options and formula policy | `ToCsvOptions` | Spreadsheet-safe formula neutralization and 64 MiB hard output ceiling by default; raw exchange is explicit. Covered by CSV core/Node security tests and LibreOffice receipt. |
| Other output | Bounded render and transfer options | JSON/HTML/SVG/raw wrapper paths | Traversal/output ceilings are finite and typed limit failures remain distinguishable from invalid arguments. Covered by render, transfer, JSON, output, and streaming security suites. |
| Error contract | `ResourceLimitError` metadata | `ResourceLimitError` plus `isResourceLimitError` | Closed string codes and exact integer metadata are preserved; JavaScript exposes count/byte fields as `bigint`. Synchronous core-result boundaries and background open/save/write tasks preserve the same typed metadata, including lazy hydration. The Node 24 focused replay passed 79/79. |
| Declarations/runtime | `napi8` feature | Node 24.18.0, `@types/node` 24 | The generated `binding.d.ts` and public declarations compile together using `__test__/declarations/consumer.ts`. The current-source release addon passed the complete Node suite, 643/643. |

The TypeScript declaration fixture is a release gate:

```sh
pnpm --filter @sheetkit/node typecheck:declarations
```

The runtime parity result is also recorded in
`docs/plans/artifacts/2026-07-15-001-fix-security-hardening-v1/u12-interop/results.md`
for the same 2026-07-23 replay.
