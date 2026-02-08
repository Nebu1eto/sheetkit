# Contributing to SheetKit

## 1. Prerequisites

- **Rust toolchain** (rustc, cargo) -- latest stable release
- **Node.js** >= 18
- **pnpm** >= 9

## 2. Repository Setup

```bash
git clone <repo-url>
cd sheetkit
pnpm install
cargo build --workspace
```

After cloning, verify that the full workspace builds and all tests pass:

```bash
cargo test --workspace
cargo clippy --workspace
cargo fmt --check
```

## 3. Project Structure

```
sheetkit/
  crates/
    sheetkit-xml/      # XML schema types (serde-based OOXML mappings)
    sheetkit-core/     # Core business logic (all features implemented here)
    sheetkit/          # Public facade crate (re-exports from sheetkit-core)
  packages/
    sheetkit/          # Node.js bindings via napi-rs
  examples/
    rust/              # Rust usage examples
    node/              # Node.js usage examples
  docs/                # Documentation
```

## 4. Build Commands

### Rust workspace

```bash
cargo build --workspace        # Build all crates
cargo test --workspace         # Run all Rust tests
cargo clippy --workspace       # Lint (must produce zero warnings)
cargo fmt --check              # Verify formatting
```

### Node.js bindings

The napi build process has three steps that must be followed in order:

```bash
cd packages/sheetkit

# Step 1: Build the native addon. This compiles the Rust cdylib and generates
# index.js (CJS binding loader) and index.d.ts (TypeScript declarations).
npx napi build --platform

# Step 2: Rename the generated CJS file. The napi build overwrites index.js
# with a CJS module, but we need to keep our ESM wrapper as index.js.
cp index.js binding.cjs

# Step 3: Restore the ESM wrapper. The real index.js is an ESM module that
# loads binding.cjs via createRequire. Restore it from git.
git checkout -- index.js

# Step 4: Run the Node.js test suite.
npx vitest run
```

The ESM wrapper (`index.js`) is checked into the repository and must never be overwritten permanently by the napi build output. It loads the CJS binding using Node.js `createRequire`.

## 5. Development Workflow

SheetKit follows a TDD (Test-Driven Development) approach:

1. **Write tests first** that describe the expected behavior.
2. **Implement the feature** until all tests pass.
3. **Run the full verification checklist** before considering the work complete.

### Verification Checklist

Every change must pass all of the following before submission:

- [ ] `cargo build --workspace` -- compiles without errors
- [ ] `cargo test --workspace` -- all tests pass
- [ ] `cargo clippy --workspace` -- no warnings
- [ ] `cargo fmt --check` -- formatting is correct
- [ ] `cd packages/sheetkit && npx vitest run` -- Node.js tests pass (if bindings were changed)

## 6. Code Style

### Rust

- Follow standard Rust conventions. Use `cargo fmt` for automated formatting.
- Note that `cargo fmt` may reformat files in crates you did not directly modify. Include those reformatted files in your commits.

### TypeScript / JavaScript

- Use Biome for formatting and linting.
- ESM only for all JavaScript and TypeScript code. The napi CJS output (`binding.cjs`) is the sole exception and is loaded via `createRequire` in the ESM wrapper.

### General Rules

- **English everywhere in code**: All variable names, string literals, comments, and example data values must be in English. Even demo or test data should use English strings (e.g., "Name", "Sales", "Employee List").
- **Doc comments**: Use `///` for Rust and `/** */` for TypeScript. Describe inputs, behavior, and outputs concisely.
- **Inline comments**: Only for logic that is not self-evident from the code itself.
- **No section markers or decorative comments**: Do not add comment banners, separators, or ornamental markers.

## 7. Adding a New Feature

Follow these steps when implementing a new feature:

### Step 1: XML types (if needed)

If the feature requires new OOXML XML structures, add serde-based types to `crates/sheetkit-xml/src/`. Create a new file or extend an existing one depending on the OOXML part involved.

### Step 2: Core business logic

Implement the feature in `crates/sheetkit-core/src/`. Place the logic in its own module file (e.g., `feature_name.rs`) to minimize merge conflicts when multiple contributors work in parallel.

Write tests in `#[cfg(test)]` inline test modules within the same file. Tests should verify essential behaviors, not trivial properties.

Register the new module in `crates/sheetkit-core/src/lib.rs`.

### Step 3: Facade re-exports (if needed)

If the feature introduces new public types that end users should access, ensure they are re-exported through `crates/sheetkit/src/lib.rs`.

### Step 4: Node.js bindings

Add napi bindings in `packages/sheetkit/src/lib.rs`:

- Add `#[napi]` methods to the `Workbook` class that delegate to the core implementation.
- Define `#[napi(object)]` structs for any new configuration or result types.
- Use `Either` types for polymorphic parameters or return values.

### Step 5: Node.js tests

Add test cases in `packages/sheetkit/__test__/index.spec.ts` covering the new bindings.

### Step 6: Rebuild and verify

Rebuild the napi bindings and run the full verification checklist (see Section 5).

## 8. Workspace Layout

### Cargo workspace

The Cargo workspace includes:

- `crates/sheetkit-xml`
- `crates/sheetkit-core`
- `crates/sheetkit`
- `packages/sheetkit` (the napi-rs crate, named `sheetkit-node` in Cargo.toml)
- `examples/rust`

### pnpm workspace

The pnpm workspace includes:

- `packages/*`
- `examples/*`

## 9. Key Dependencies

| Crate | Purpose |
|---|---|
| `quick-xml` | XML parsing and serialization (with `serialize` and `overlapped-lists` features) |
| `serde` | Derive-based (de)serialization for XML types |
| `zip` | ZIP archive handling for .xlsx files (with `deflate` feature) |
| `thiserror` | Ergonomic error type definitions |
| `nom` | Formula string parsing (produces an AST) |
| `napi` / `napi-derive` | Node.js native addon bindings (v3, no compat-mode) |
| `chrono` | Date/time handling and Excel serial number conversion |
| `tempfile` | Temporary file creation in tests |
| `pretty_assertions` | Improved assertion diff output in tests |

## 10. Common Gotchas

### cargo fmt side effects

`cargo fmt` may reformat files in crates you did not directly change. Always check `git diff` after formatting and include any reformatted files in your commit.

### ZIP compression options

When writing ZIP entries, always use:

```rust
SimpleFileOptions::default().compression_method(CompressionMethod::Deflated)
```

### napi build output

The napi build generates a CJS `index.js` that overwrites the ESM wrapper. Always rename it to `binding.cjs` and restore the ESM wrapper from git after building.

### File organization

Place new features in their own source files rather than adding to existing large files (especially `workbook.rs`). This reduces merge conflicts when multiple people work on different features simultaneously.

### Staging files for commits

Stage only the specific files you changed. Do not use `git add -A` or `git add .`, as these can accidentally include generated files, test output, or other unintended changes.

### Test output files

Any `.xlsx` files generated by tests are gitignored and should not be committed to the repository.

### Namespace prefix handling

Some OOXML namespace prefixes (`dc:`, `dcterms:`, `cp:`, `vt:`) cannot be handled by serde and require manual quick-xml Writer/Reader code. If your feature involves document properties or similar namespace-prefixed elements, you will need to write manual serialization/deserialization logic.

### Formula parser

The formula parser uses the `nom` crate, not a hand-written parser. If you need to extend formula support, familiarize yourself with nom combinators in `crates/sheetkit-core/src/formula/parser.rs`.
