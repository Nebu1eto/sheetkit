# U12 structured fuzz and corpus replay evidence

Revalidated on 2026-07-23 with stable `rustc 1.93.0` and `cargo 1.93.0`
on `aarch64-apple-darwin`.

## Stable release gate

Command:

```text
cargo +stable test -p sheetkit-core --test security_fuzz_corpus_replay
```

Result: 3 passed, 0 failed. The gate freezes and replays the exact 33-file
canonical inventory twice, executes 2,048 deterministic structured mutations,
and checks exact worksheet coordinate transfer round trips plus
invalid-coordinate rejection.

The first red replay exposed unchecked raw-transfer coordinates. The finalized
raw-transfer controls validate coordinates and structural counts before
allocation, use checked encoder arithmetic, and account decoded collections and
string copies in one cumulative output budget. The same corpus now passes
including tiny inputs declaring 65,535 cells against a one-column span,
`u32::MAX` rows, and `u32::MAX` v2 string bytes.

## Fuzz target build and smoke

Command:

```text
CARGO_TARGET_DIR=/tmp/sheetkit-u12-fuzz-final-20260723-1 \
  cargo +stable build --manifest-path fuzz/Cargo.toml --bins
```

Result: all four targets built on stable from a clean target directory. Direct
`libfuzzer-sys` smoke runs used `-timeout=10`, `-rss_limit_mb=2048`, and each
target's production input ceiling. They completed without crash or panic:

```text
/tmp/sheetkit-u12-fuzz-final-20260723-1/debug/workbook_open \
  security/fuzz-corpus/workbook_open -runs=50 -timeout=10 \
  -rss_limit_mb=2048 -max_len=8388608
/tmp/sheetkit-u12-fuzz-final-20260723-1/debug/worksheet_xml \
  security/fuzz-corpus/worksheet_xml -runs=100 -timeout=10 \
  -rss_limit_mb=2048 -max_len=1048576
/tmp/sheetkit-u12-fuzz-final-20260723-1/debug/formula \
  security/fuzz-corpus/formula -runs=100 -timeout=10 \
  -rss_limit_mb=2048 -max_len=65537
/tmp/sheetkit-u12-fuzz-final-20260723-1/debug/transfer \
  security/fuzz-corpus/transfer -runs=100 -timeout=10 \
  -rss_limit_mb=2048 -max_len=1048576
```

| Target | Corpus files | Runs | Result |
| --- | ---: | ---: | --- |
| `workbook_open` | 6 | 50 | pass |
| `worksheet_xml` | 9 | 100 | pass |
| `formula` | 6 | 100 | pass |
| `transfer` | 12 | 100 | pass |

`cargo-fuzz` was not installed on this machine. The direct stable smoke lacked
coverage instrumentation, so it proves target linkage and corpus execution but
does not substitute for a nightly instrumented campaign. Nightly cargo-fuzz and
sanitizer commands are documented as supplementary release-runner checks in
`docs/security/v1-fuzz-verification.md`; neither is the sole release gate.

## Static validation

Commands:

```text
actionlint .github/workflows/ci.yml
ruby -e 'require "yaml"; YAML.load_file(".github/workflows/ci.yml", aliases: true)'
cargo +stable clippy -p sheetkit-core \
  --test security_fuzz_corpus_replay -- -D warnings
cargo +stable clippy --manifest-path fuzz/Cargo.toml --bins -- -D warnings
rustfmt --edition 2021 --check security/fuzz_support.rs \
  crates/sheetkit-core/tests/security_fuzz_corpus_replay.rs \
  fuzz/fuzz_targets/*.rs
```

Result: the workflow syntax and paths are valid, fuzz targets are warning-free,
and the replay/support sources are formatted. CI runs the deterministic stable
replay explicitly. Nightly sanitizer and cargo-fuzz work remains supplementary.

## Corpus identity

The committed text recipes and hex seeds are deterministic. Their aggregate
identity can be reproduced with:

```text
find security/fuzz-corpus -type f -print0 | LC_ALL=C sort -z \
  | xargs -0 shasum -a 256
find security/fuzz-corpus -type f -print0 | LC_ALL=C sort -z \
  | xargs -0 shasum -a 256 | shasum -a 256
```

The sorted per-file manifest contains 33 entries. Hashing that manifest produces
`deaab90f835d7a25895e18027e8a34449e671d71e9a51d18d34e3ceb609ea462`.

## Performance status

These checks establish bounded correctness, not an R7 performance verdict.
Benchmark values collected while other workloads are active on this machine are
reference-only. Final latency and peak-memory acceptance remains deferred until
the same workloads and baselines can be measured again in an otherwise idle,
clean environment.
