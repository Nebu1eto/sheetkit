# SheetKit Agent Guide

SheetKit reads and writes Excel `.xlsx` files. Rust owns all logic; Node.js bindings wrap it with napi-rs.

## Layout

- `crates/sheetkit-xml`: serde XML types for OOXML.
- `crates/sheetkit-core`: all business logic (workbook, cells, styles, formulas, charts, streaming, encryption behind the `encryption` feature).
- `crates/sheetkit`: public facade that re-exports `sheetkit-core`.
- `packages/sheetkit`: napi cdylib plus the TypeScript wrapper (`index.ts`, `buffer-codec.ts`, `sheet-data.ts`, `open-options.ts`). `binding.js`/`binding.d.ts` and the `.js`/`.d.ts` next to the sources are generated.
- `examples/`, `benchmarks/`, `docs/` (VitePress, English and Korean), `tools/oxlint/anti-slop` (vendored lint rules; see its `UPSTREAM.md`).

Toolchain: `mise` provides Node.js (`.node-version`) and pnpm (`mise.toml`); Rust comes from rustup. Node.js 24 is the minimum supported runtime, so keep `@types/node` on 24.x and avoid newer APIs.

## Working rules

- Do the requested change and the validation it needs without asking first. Ask only when missing input would change the outcome. Confirm before destructive actions, publishing, releases, or dependency and toolchain changes the request did not cover.
- Keep the diff to the request. Do not fix, refactor, or extend nearby code, and add tests only for behavior that the change introduces or protects.
- Rust and Node.js APIs must behave identically: defaults, return values, and errors. A change on one side updates the other.
- Zero tolerance for diagnostics: no warning or error from any linter, formatter, type checker, compiler, or test may remain. Fix the cause. Never suppress with disable comments, `@ts-ignore`/`@ts-expect-error`, `any`, assertion chains, rule downgrades, or new ignore patterns.
- Prefer targeted edits to whole-file rewrites.

## Code style

- English for all identifiers, strings, comments, and example data, including demos and tests.
- Doc comments (`///`, `/** */`) state inputs, behavior, and outputs briefly. Inline comments only for non-obvious logic. No decorative comments or emoji anywhere.
- TypeScript: ESM only, no hand-written `.js`/`.mjs`/`.cjs`, no non-null `!` (use `assert(value != null)` from `node:assert`). Scripts run with `tsx`; library code builds with tsdown.
- Korean docs (`docs/ko/`, `README.ko.md`): polite declarative endings only (`~합니다/~됩니다/~입니다/~세요`), never `~이다/~한다/~된다`. Keep common technical terms untranslated. Code blocks may carry Korean comments; identifiers and data stay English.

## Commands and gates

Run everything relevant to the change before reporting completion; CI runs the same set.

```bash
cargo build --workspace
cargo clippy --workspace -- -D warnings
cargo fmt --check
cargo test --workspace --features sheetkit-core/encryption
RUSTDOCFLAGS="-D warnings" cargo doc --no-deps --workspace
pnpm check            # oxlint + oxfmt --check + vendored rule typecheck
pnpm -r build         # napi build -> tsc --noEmit -> tsdown
pnpm --filter @sheetkit/node typecheck:declarations
pnpm test             # vitest for the Node.js binding
pnpm docs:build       # rustdoc (warnings are errors) + typedoc + VitePress
```

`pnpm check:write` applies lint autofixes and formatting. Update `docs/` when a public API changes.

## Commits

Use the `commit` skill. Subject `[<crate-or-package>] <type>(#<issue>): <summary>` (`[*]` for repo-wide, omit the issue segment when unknown), body of `- <what and why>` bullets, all lines at most 72 characters, `Assisted-by: <agent>:<model>` and `Signed-off-by` trailers, signed when signing is configured. Stage files explicitly; never `git add -A` or `git add .`. See `AI_POLICY.md`.

## Facts that are easy to get wrong

- The XML declaration is prepended by hand: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`.
- `dc:`, `dcterms:`, `cp:`, and `vt:` prefixes need manual quick-xml readers and writers; serde does not handle the namespaces.
- `sharedStrings.xml` is optional; use `Sst::default()` when absent.
- ZIP entries use `SimpleFileOptions::default().compression_method(CompressionMethod::Deflated)`.
- Encrypted files are OLE/CFB containers: ZIP starts with `PK\x03\x04`, CFB with `D0CF11E0A1B11AE1`. Node bindings always enable `encryption`.
- `cargo fmt` may touch other crates; include those files in the commit.
- Test output `.xlsx` files are gitignored except under `fixtures/`.
