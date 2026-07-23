# U14 1.0.0 release-candidate receipt

Observed: 2026-07-29

This receipt covers local release-candidate correctness and package contents.
It does not authorize or record a tag, push, publication, or remote release.
The final clean-environment performance comparison is recorded in the
[U14 performance receipt](../u14-performance/results.md).

## Version and generated-artifact consistency

- Workspace Rust version: `1.0.0`
- napi Rust package version: `1.0.0`
- Main npm package and all eight optional platform packages: `1.0.0`
- All eight platform package manifests name an existing matching `.node` file.
- macOS arm64 addon SHA-256:
  `0c4d141d61f9193efadc3e7ad5e002d7c3af8532233c1ab0a417e7a839c40ad1`
- Generated `binding.d.ts` SHA-256:
  `357296b431ef66b5b123b55cf0129a223b72f0d937672545f8cdfe87279f80dd`
- Generated `index.js` SHA-256:
  `99f07e674c6f0f517cac74ea22ec633a2725dec100df43a4325e491f32359bb0`
- `pnpm install --frozen-lockfile` state is represented by the refreshed
  `pnpm-lock.yaml`; `cargo metadata --locked` resolved all four SheetKit Rust
  packages at `1.0.0`.

## Package contents

`npm pack --dry-run --json` for `@sheetkit/node@1.0.0` reported ten entries:
the package README, `package.json`, four generated JavaScript modules, and four
matching declaration files. Scan reports, exploit fixtures, fuzz corpora,
source files, tests, and internal plan artifacts were absent. The package README
links the public documentation and v0.5.1-to-1.0 migration guide.

## Local final gates

| Gate | Result |
|---|---|
| `cargo build --workspace` | Pass |
| `cargo test --workspace` | Pass; core 2,279/2,279 plus workspace integration and doc targets |
| `cargo test --workspace --features sheetkit-core/encryption` | Pass |
| `cargo clippy --workspace --all-targets --all-features -- -D warnings` | Pass |
| `cargo fmt --all -- --check` | Pass |
| `pnpm check` | Pass; 38 files |
| `pnpm -r build` | Pass on Node.js 24 |
| `pnpm test` | Pass; 646/646 across 18 files on Node.js 24 |
| napi source typecheck | Pass |
| generated declaration consumer typecheck | Pass |
| CLI build and test | Pass; 27/27 |
| VitePress documentation build | Pass; 171 Markdown files |
| `actionlint` | Pass |

The U12 interoperability receipt separately records LibreOffice,
msoffcrypto-tool, VBA package inspection, atomicity, exploit corpus replay, and
fuzz replay results against the same generated 1.0 addon.

## Performance gate

The idle clean-environment comparison used the repository's existing focused
Node benchmark against v0.5.1. Build time was excluded. Open, traversal,
streaming, V2 transfer, and save rows all remained below the 10% release-stop
threshold after the affected raw-transfer and COW save paths were optimized.
R7 passes.

## Clean-checkout reproduction

A detached worktree at the signed candidate commit completed a frozen pnpm
install, `cargo build --workspace`, and `pnpm -r build`. The generated
declaration and JavaScript hashes matched the primary checkout, and
`npm pack --dry-run --json` reproduced the same ten entries, sizes, package
SHA-1, and integrity value. The native addon hash differed between checkout
paths, so native binaries are not claimed to be byte-reproducible; both builds
completed from the same source and passed the runtime suite in the primary
checkout.

## Final review follow-up

The full correctness, standards, testing, maintainability, security,
performance, API-contract, and reliability review applied the concrete
correctness and security findings: shared raw-transfer budgeting and target
preflight, transactional stream formula charging, explicit post-rename commit
semantics, exact Rust/Node/CLI CSV formula parity, release CI gates, and
canonical API documentation. Focused regressions and the repository-wide gates
above pass after those fixes.

The independent Claude adversarial route ran against the full diff but returned
no usable schema-shaped output, so it is not counted as corroboration.

The retained performance fixes remove duplicate validation work without
weakening limits: raw decode performs one structural pass and one materializing
pass, transfer scans aggregate checks per row, and immutable worksheet save
metrics are cached while each caller's limits are rechecked. The proposed
financial-module split was rejected as a non-behavioral organization preference
without a repository rule or runtime failure.
