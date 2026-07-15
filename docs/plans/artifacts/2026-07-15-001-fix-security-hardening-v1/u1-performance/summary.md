# U1 performance baseline summary

## Baseline identity

- Remediation baseline: `150e57939ac816a9bfa86daa856b432a135e3ca9`.
- Cumulative upgrade baseline: v0.5.1 tag commit
  `f9eac5366eb83b6dd79d5dd6d03a8988b5c319a8`.
- Machine: Apple M4 Pro, 24 GiB RAM, macOS 26.5.2, arm64.
- Toolchain: Rust 1.93.0, Cargo 1.93.0, Node.js 25.3.0, pnpm 10.29.1.
- The two benchmark source files, Cargo lockfile, pnpm lockfile, and all fixture
  checksums were identical across the two source archives.

The current revision is the comparison point for future security remediation.
The v0.5.1 measurement is an independent cumulative-upgrade comparison. The
tag-to-current differences below predate the security implementation and must
not be attributed to later hardening work.

## Captured evidence

- Rust Criterion: 9 successful scenarios per revision across full open, lazy
  open, and `get_rows`; raw samples, 95% estimates, and stdout are retained.
- Node focused runner: 60 successful scenarios per revision; all five measured
  samples, median, p95, and runner RSS deltas are retained as JSON.
- Aggregate fresh-process Node high-water RSS: v0.5.1 used 3,811,000,320 bytes;
  current used 3,453,747,200 bytes. This is a whole-suite auxiliary signal, not
  a per-scenario memory gate.

## Representative tag-to-current deltas

Positive percentages mean the current revision was slower than v0.5.1.

| Runner / scenario | v0.5.1 | current | delta |
|---|---:|---:|---:|
| Rust full open, scale-1k | 1.463 ms | 2.065 ms | +41.1% |
| Rust full open, scale-10k | 14.383 ms | 21.340 ms | +48.4% |
| Rust full open, large-data | 144.222 ms | 220.286 ms | +52.7% |
| Rust lazy open, scale-10k | 16.885 ms | 24.221 ms | +43.4% |
| Rust lazy open, large-data | 149.248 ms | 184.367 ms | +23.5% |
| Rust get rows, scale-10k | 2.339 ms | 2.368 ms | +1.2% |
| Node async path open, scale-100k | 150.928 ms | 195.666 ms | +29.6% |
| Node lazy open, scale-100k | 148.660 ms | 209.508 ms | +40.9% |
| Node lazy open plus rows, scale-100k | 609.138 ms | 652.361 ms | +7.1% |
| Node stream reader, scale-100k | 776.752 ms | 851.250 ms | +9.6% |
| Node untouched COW save, scale-100k | 267.349 ms | 331.430 ms | +24.0% |
| Node edited COW save, scale-100k | 791.629 ms | 925.016 ms | +16.9% |

Across all Node rows, tag-to-current median deltas ranged from -23.2% to
+49.7%; 34 of 60 were above 10%. Across the 9 successful Rust rows, five were
above 10%. `comparison.json` retains every ratio without averaging away an
individual workload.

## Noise and application floor

The existing Node runner uses one warmup plus five samples. Its median
within-run relative MAD was 1.60% for v0.5.1 and 1.47% for current, but small
sub-2 ms operations and a few save cases were materially noisier. Criterion
retains its per-scenario 95% confidence intervals. For subsequent comparisons on
this same machine, use both the retained distribution/CI and these small
absolute floors before applying the plan's percentage bands:

- latency: 0.5 ms for Node and 0.05 ms for Rust;
- aggregate process high-water memory: 16 MiB.

A percentage-only change below the corresponding absolute floor is noise, not a
gate failure. Any primary workload above 10% and above the floor remains a hard
stop under R7; a 5-10% result requires profiling and justification.

## Exact limitations and blocker

- The existing focused Node runner mixes warmup and warm samples in one fresh
  process. It does not publish a separately timed cold-open scenario.
- Its per-scenario `rssMedianMb` is a before/after RSS delta, not process peak
  RSS. Only the additional whole-run `/usr/bin/time -l` result is a true
  fresh-process high-water signal.
- The focused suite does not cover encrypted open, scalar/formula-heavy work,
  sparse traversal, or CSV/JSON/HTML/SVG output. Later units must add or use
  representative runners for those R7 workloads before release readiness.
- Rust `save_latency` is not a valid baseline: both revisions panic at the same
  source line with `UnsupportedFileExtension("")` because the benchmark creates
  an extensionless temporary output. The raw logs record the identical failure;
  no value was synthesized.
- Five Node samples are suitable for a frozen starting snapshot, not for a
  noisy shared-CI release decision. Stable-runner repeats remain required for
  the final gate.

## Artifact map

- `manifest.json`: revisions, environment, resolutions, and artifact hashes.
- `fixtures.sha256`: the exact shared workbook corpus.
- `comparison.json`: all machine-readable tag-to-current ratios.
- `raw/*/node-bench-baseline.json`: Node samples and statistics.
- `raw/*/rust-criterion-samples.json`: Criterion iteration/time samples.
- `raw/*/rust-criterion-estimates.json`: Criterion confidence estimates.
- `raw/*/*.stdout.txt`: captured benchmark console output and wrapper status.
- `raw/*/node-bench-baseline.process-rss.txt`: fresh-process memory evidence.
