# U14 clean performance receipt

## Verdict

R7 passes for the 1.0 release candidate against v0.5.1. No measured primary
workload remains above the 10% release-stop threshold.

Build, dependency installation, fixture generation, and linking time are
excluded. Every reported latency measures a SheetKit operation on an existing
Excel fixture.

## Comparison identity

- Baseline: v0.5.1 commit
  `f9eac5366eb83b6dd79d5dd6d03a8988b5c319a8`.
- Candidate source: `prepare-v1`, based on signed commit
  `35bf560609b3d99838d96b8eda38e417b8619582`, plus the performance fixes
  recorded with this receipt.
- Machine: Apple M4 Pro, 24 GiB RAM, macOS 26.5.2.
- Toolchain: Rust 1.93.0, Cargo 1.93.0, Node.js 24.18.0, pnpm 10.29.1.
- Primary runner: `benchmarks/node/bench-baseline.ts`, SHA-256
  `6e8de47eaa8736350cd720fdd38e16089a157d3e59c442c7518007fcee4d6e54`.
- Cargo lock SHA-256:
  `3b28ba98279f18056553f1b6fd569db7532ca9c76a9a5796007789a0759466a1`.
- pnpm lock SHA-256:
  `6543744f193cc3d71fda60caa6fb4eed4550358bb512fbfac0e5677a533b736a`.

The same runner source and fixture bytes were used for both revisions. The
v0.5.1 checkout used the runner's explicit eager compatibility options where
the 1.0 public spelling differs.

## Method

The primary comparison follows the repository's existing focused baseline:

- one untimed warm-up;
- five measured samples per scenario;
- forced GC before each sample;
- median latency as the comparison statistic;
- identical fixtures and operation boundaries;
- pre-opened workbooks for `*-only` traversal and save rows.

The final `getRowsRaw-only` fix crossed the decision threshold, so it received
two additional five-sample blocks with the same method. Their ten samples were
combined before taking the final median. V2 is compared with V2 because both
revisions produce the same canonical occupied-row result.

Earlier three-revision release-gate runs are diagnostic only. They are not the
primary acceptance result.

## Primary results

The table uses the 100,000-row fixture. Positive percentages are regressions.

| Operation | v0.5.1 median | 1.0 median | Change | Verdict |
|---|---:|---:|---:|---|
| Lazy `openSync` from path | 144.368 ms | 103.1 ms | -28.6% | Pass |
| `getRows` including open | 637.850 ms | 639.1 ms | +0.2% | Pass |
| `getRowsRaw-only` | 64.093 ms | 64.975 ms | +1.4% | Pass |
| `getRowsBufferV2-only` | 25.028 ms | 26.8 ms | +7.1% | Profiled |
| `openSheetReader` | 761.651 ms | 786.5 ms | +3.3% | Pass |
| Untouched COW `saveSync` | 111.606 ms | 61.8 ms | -44.6% | Pass |
| Edited COW `saveSync` | 251.750 ms | 249.2 ms | -1.0% | Pass |

The two final raw-traversal blocks had medians of 63.590 ms and 66.012 ms.
Their combined median was 64.975 ms.

## Profiled results

V2 buffer production is 1.772 ms, or 7.1%, slower than v0.5.1. The 1.0 path
also validates canonical coordinates, resolves shared formulas and date
styles, and enforces exact finite output accounting. Row-level aggregation
removed the avoidable per-cell limit-check overhead. The remaining absolute
cost is bounded and stays below the 10% stop threshold.

The candidate-only Criterion traversal suite also passed its asymptotic
checks. Sparse v1 near and far cases were 181.58 ns and 185.48 ns, while sparse
v2 far was 168.34 ns. Coordinate distance therefore does not amplify sparse
work. The 50,000-value pivot diagnostic was 20.149 ms and 6.1% above its local
prior estimate; it remains linear, below the stop threshold, and is unrelated
to the release comparison adapter.

## Performance fixes retained

- V1 and V2 transfer dimension scans aggregate cell-limit checks per row
  without weakening the cumulative limit.
- Raw decoding reuses counts collected during structural validation and
  performs fatal UTF-8 decoding once, before returning the completed result.
- The V2 raw path keeps all offset, range, tag, coordinate-order, boolean, and
  trailing-byte checks in one specialized validation pass.
- Immutable package-backed worksheet metrics are cached after their first
  bounded scan. Every save still rechecks the caller's current XML, row, and
  cell limits.

The Node security suite passed all 646 tests after these changes. The Rust raw
transfer tests and save-output regression suite also pass.

## Fixture checksums

| Fixture | SHA-256 |
|---|---|
| `scale-1k.xlsx` | `08d7a657973ac6c757b7af9dac36692378dddc76426df32fa25612973b5c0a29` |
| `scale-10k.xlsx` | `4c57df3e1850fec424c39e2dd6657cf804c76fb694a07a8718c7dbd261dfdfbe` |
| `scale-100k.xlsx` | `13aa19cfbf08ccffe24be7302007accb377b8eb1e4247c4fd2f18f50f4cd97bd` |
| `large-data.xlsx` | `fdd1ae3dc0272a66fde15cc898f8c8f82160ffbba5333beaa0db943c50591d72` |
| `strings.xlsx` | `30eaa89c478f53c1bb971c31e2bc7c95aef376fabd44d7a08c19acab46077db0` |
