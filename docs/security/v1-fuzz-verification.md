# SheetKit 1.0 Security Fuzz Verification

The release gate uses deterministic corpus replay on the stable Rust toolchain.
The corpus covers workbook ZIP metadata and open modes, worksheet XML and A1
references, formula parsing and budgeted evaluation, and v1/v2 raw-transfer
validation. Seeds include accepted boundaries, one-over limits, misleading ZIP
metadata, oversized coordinate text, truncated transfer sections, and
attacker-declared counts that are disproportionate to the input.

Run the stable release gate:

```bash
cargo +stable test -p sheetkit-core --test security_fuzz_corpus_replay
```

The replay test materializes text recipes, runs every seed twice, and requires
identical outcomes without panic. It also executes 2,048 deterministic
truncate, bit-flip, integer-overwrite, and append mutations. The target harness
caps only the input already owned by the fuzzer; production APIs retain their
own finite policies and are not replaced by harness limits.

The four targets can also be built and replayed without nightly instrumentation
as a bounded linkage and corpus smoke check:

```bash
cargo +stable build --manifest-path fuzz/Cargo.toml --bins
fuzz/target/debug/workbook_open security/fuzz-corpus/workbook_open \
  -runs=50 -timeout=10 -rss_limit_mb=2048 -max_len=8388608
fuzz/target/debug/worksheet_xml security/fuzz-corpus/worksheet_xml \
  -runs=100 -timeout=10 -rss_limit_mb=2048 -max_len=1048576
fuzz/target/debug/formula security/fuzz-corpus/formula \
  -runs=100 -timeout=10 -rss_limit_mb=2048 -max_len=65537
fuzz/target/debug/transfer security/fuzz-corpus/transfer \
  -runs=100 -timeout=10 -rss_limit_mb=2048 -max_len=1048576
```

This smoke check is not coverage-instrumented and does not replace either the
stable deterministic release gate or a nightly fuzzing campaign.

Four cargo-fuzz targets reuse the exact stable harness. They are supplementary
because cargo-fuzz requires nightly compiler instrumentation and is not the sole
release gate:

```bash
cargo fuzz run workbook_open security/fuzz-corpus/workbook_open -- -max_total_time=60
cargo fuzz run worksheet_xml security/fuzz-corpus/worksheet_xml -- -max_total_time=60
cargo fuzz run formula security/fuzz-corpus/formula -- -max_total_time=60
cargo fuzz run transfer security/fuzz-corpus/transfer -- -max_total_time=60
```

Recommended supplementary checks for a dedicated runner are:

```bash
RUSTFLAGS="-Zsanitizer=address" \
  cargo +nightly test -Zbuild-std --target aarch64-apple-darwin \
  -p sheetkit-core --test security_fuzz_corpus_replay

RUSTFLAGS="-Zsanitizer=address" \
  cargo +nightly test -Zbuild-std --target x86_64-unknown-linux-gnu \
  -p sheetkit-core --test security_fuzz_corpus_replay
```

Record the compiler, target, corpus revision, duration, and any minimized crash
artifact. A crash, panic, non-deterministic replay, or allocation proportional
to an attacker declaration is a release blocker. Sanitizer and randomized fuzz
runs supplement, but never replace, the stable deterministic replay gate.
