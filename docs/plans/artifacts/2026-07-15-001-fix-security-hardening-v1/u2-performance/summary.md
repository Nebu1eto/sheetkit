# U2 provisional performance evidence

## Status

These measurements are reference-only. The machine was under material load
during the run, so R7 is deferred until the same scenarios are repeated in a
clean environment. No release verdict is derived from this artifact.

- Working-tree base: `150e57939ac816a9bfa86daa856b432a135e3ca9`.
- Cumulative baseline: v0.5.1 tag
  `f9eac5366eb83b6dd79d5dd6d03a8988b5c319a8`.
- Captured on Apple M4 Pro, 24 GiB RAM, macOS 26.5.2.
- Toolchain: Rust 1.93.0, Node.js 25.3.0, pnpm 10.29.1.
- Node runner policy: one warmup plus five samples in one process.
- Rust runner policy: Criterion defaults with 95% confidence intervals.

The release rerun must use the same fixture checksums recorded by U1, warmed
repeated distributions, and fresh-process memory measurements. In particular,
the provisional Node path scale-10k result must be repeated with at least 20
samples and alternating tag/working-tree order before it can pass or fail R7.

## Implemented hot-path changes

- Shared strings are opened and inflated once; serde parsing and the exact-text
  fallback share one bounded byte buffer.
- Ordinary SSTs skip the second raw XML scan. The fallback remains enabled for
  `xml:space` and CDATA and is covered by exact-text parity tests.
- Optional package entries use structured `ZipError::FileNotFound` handling,
  preserving all other ZIP and resource-limit errors.
- Clean control metadata is not rescanned during save.
- Rust and Node benchmark labels now distinguish default-lazy open from
  explicit eager open. Save-only cases pre-open their workbook, and Rust output
  files carry an `.xlsx` suffix.

## Provisional observations

Positive percentages mean the hardened working tree was slower than the named
baseline. These values are not final R7 evidence.

| Runner / scenario | v0.5.1 | pre-hardening | working tree | vs tag | vs pre |
|---|---:|---:|---:|---:|---:|
| Rust default lazy, scale-1k | 1.463 ms | 2.065 ms | 1.419 ms | -3.1% | -31.3% |
| Rust default lazy, scale-10k | 14.383 ms | 21.340 ms | 14.926 ms | +3.8% | -30.1% |
| Rust default lazy, large-data | 144.222 ms | 220.286 ms | 152.877 ms | +6.0% | -30.6% |
| Node path sync, scale-10k | 13.939 ms | 18.777 ms | 16.341 ms | +17.2% | -13.0% |
| Node path async, scale-10k | 13.373 ms | 18.386 ms | 15.081 ms | +12.8% | -18.0% |
| Node buffer sync, scale-10k | 16.556 ms | 19.315 ms | 15.135 ms | -8.6% | -21.6% |
| Node buffer async, scale-10k | 15.381 ms | 19.059 ms | 14.991 ms | -2.5% | -21.3% |

The provisional Node scale-10k path values cross the nominal 10% threshold,
but the same run's buffer values improve and the Rust result is +3.8%. Because
the user identified current-machine load, this discrepancy is retained as a
mandatory clean-environment rerun, not interpreted as noise or as a release
failure.

The newly added Rust scale-100k scenario completed at a mean of 166.662 ms
(95% CI 162.327-171.773 ms). U1 has no equivalent Rust scale-100k baseline, so
no ratio is synthesized. The current strings mean is 31.632 ms
(95% CI 30.789-32.584 ms), also without a U1 Rust baseline.

## Deferred diagnostics

The planned fresh-process path rerun and separate microbenchmarks for
`File::metadata` and the N-API error bridge were stopped before execution when
the run was declared load-contaminated. Static inspection shows that
`File::metadata` is on successful path opens, while structured JS error-object
construction runs only on error. Both are still listed for the clean runner so
the path-only discrepancy is resolved with measurements rather than inference.

## Artifact map

- `comparison.json`: machine-readable provisional Rust and Node comparisons.
- `node-provisional-results.json`: all raw Node samples and runner metadata.
- U1 `fixtures.sha256`: authoritative checksums for the shared fixture corpus.
