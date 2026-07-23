# U14 clean performance gate runbook

This runbook defines the deferred R7 release measurement. Results collected
while the machine is carrying unrelated work are reference-only and must not
produce a release verdict.

## Final acceptance method

The final user-directed release comparison uses the repository's existing
`benchmarks/node/bench-baseline.ts` method against the last release, v0.5.1:

- copy the same benchmark source into both isolated checkouts;
- use identical fixture bytes and Node.js 24;
- build both revisions before measurement;
- run one warm-up and five measured samples per scenario;
- report operation medians only;
- exclude dependency installation, compilation, linking, fixture generation,
  and output cleanup;
- repeat only a threshold-crossing affected workload.

Use V2-to-V2 for occupied-row transfer because it materializes the same public
result in both revisions. The candidate-only Rust `traversal` Criterion suite
checks bounded and distance-independent invariants; it is not used to invent a
cross-revision ratio where v0.5.1 has no equivalent benchmark.

The broader three-revision release gate below remains a reproducible diagnostic
method. Its earlier results are reference-only and do not replace the final
two-revision acceptance receipt in [results.md](results.md).

## Diagnostic comparison set

- Release candidate: the signed, clean `prepare-v1` HEAD containing this
  runbook and `benchmarks/node/release-gate.ts`. Record its full
  `git rev-parse HEAD` value before extraction; do not run from a dirty tree.
- Remediation baseline: commit
  `150e57939ac816a9bfa86daa856b432a135e3ca9`.
- Cumulative baseline: v0.5.1 commit
  `f9eac5366eb83b6dd79d5dd6d03a8988b5c319a8`.
- Fixture manifest:
  `../u1-performance/fixtures.sha256`.
- Fixture-manifest SHA-256:
  `c4410eca9914fa3ca2a20199affe2d216b079620cb9952005dd278de67261d84`.
- Node release-gate runner:
  `benchmarks/node/release-gate.ts`.
- Runner SHA-256:
  `30b21def45aef8d110de15b4f4b13deed0c8fd393ce673907ee59fb9b5ecc31d`.

Record the exact OS, CPU, RAM, Rust, Cargo, Node.js, pnpm, Cargo.lock, and
pnpm-lock.yaml identities with the results. All three revisions must use the
same machine, release profile, toolchain family, fixture bytes, and dependency
resolution. If an older API needs an adapter, keep the measured operation
equivalent and retain the adapter source with the raw results.

## Admission check

Run only after unrelated builds, tests, simulators, virtual machines, indexers,
and package installs have stopped. Before starting:

1. Reboot or allow the machine to return to an idle thermal state.
2. Connect power and disable sleep for the measurement window.
3. Record `uptime`, `pmset -g therm`, `rustc -Vv`, `cargo -V`,
   `node --version`, `pnpm --version`, and `sw_vers`.
4. Confirm the fixture manifest with
   `shasum -a 256 -c fixtures.sha256` from the fixture directory.
5. Abort and restart the complete comparison if thermal pressure appears or an
   unrelated sustained CPU, memory, or disk workload begins.

Do not mix accepted and load-contaminated samples.

## Execution order

Extract the three revisions into separate directories and give each one an
independent `CARGO_TARGET_DIR`. Copy the candidate's release-gate runner
unchanged into both baselines and verify its SHA-256 before building. The runner
uses the common v0.5.1/`150e579`/1.0 API subset; extra 1.0 traversal options are
ignored by the older JavaScript wrappers.

Build all three before timing. Alternate the revision order for each successive
workload group:

1. v0.5.1, remediation baseline, release candidate.
2. release candidate, v0.5.1, remediation baseline.
3. remediation baseline, release candidate, v0.5.1.

For each workload, run three short blocks of seven warmed samples per revision.
Use the three revision orders above so every revision appears once in each
position. Each invocation writes to a caller-selected, create-only JSON path
that includes the block number. Concatenate the three raw arrays only after all
nine blocks complete, yielding 21 interleaved samples per revision. Do not
average block medians. Retain Criterion's raw samples and 95% confidence
intervals. A cold measurement uses a fresh process and is reported separately;
it is never included in the warmed median.

Run the existing focused Rust suites on the release candidate as invariant and
diagnostic evidence:

```sh
cargo bench --locked -p sheetkit-core --bench open_modes -- --noplot
cargo bench --locked -p sheetkit-core --bench traversal -- --noplot
cargo bench --locked -p sheetkit-core --bench formula_security -- --noplot
```

`traversal` and `formula_security` do not exist in both baselines, and
`open_modes` changed benchmark identities and default-open semantics. Do not
synthesize cross-revision Rust ratios from those non-equivalent sources. The
three-revision release gate uses the unchanged Node runner copied from the
candidate; the Rust suites prove the candidate's asymptotic invariants and
provide profiles when a Node row needs investigation.

`cargo bench` uses Cargo's optimized bench profile. Dependency compilation and
linking occur before Criterion starts and are not benchmark samples or reported
as library latency.

Run a warmed Node operation from its extracted revision as follows. Use a unique
output path for every revision, operation, and fixture:

```sh
cd /absolute/path/to/<revision>/benchmarks/node
SHEETKIT_GATE_OPERATION=ordinary-open \
SHEETKIT_GATE_FIXTURE=/absolute/path/to/scale-10k.xlsx \
SHEETKIT_GATE_OUTPUT=/absolute/path/to/results/<revision>/ordinary-open-scale-10k-block-<1-3>.json \
SHEETKIT_GATE_SAMPLES=7 \
SHEETKIT_GATE_WARMUPS=3 \
node --expose-gc --import tsx release-gate.ts
```

The runner refuses to overwrite an existing result. Use these exact operation
and fixture pairs on all three revisions:

| Gate | Runner operation and fixture |
|---|---|
| Ordinary open | `ordinary-open`: `scale-10k.xlsx` and `scale-100k.xlsx` |
| Encrypted open | `encrypted-open`: fixed encrypted `scale-10k.xlsx` |
| Scalar calculation | `scalar-calculation`: built-in fixed 1,000-formula batch |
| Dense traversal | `dense-traversal`: V2 transfer of fully occupied `scale-10k.xlsx` |
| Streaming | `streaming`: `scale-100k.xlsx` |
| Output | `output-csv`, `output-json`, `output-html`, and `output-svg`: `scale-1k.xlsx` |

Generate the encrypted fixture once from the release-candidate build:

```sh
cd /absolute/path/to/release-candidate/benchmarks/node
SHEETKIT_GATE_FIXTURE=/absolute/path/to/scale-10k.xlsx \
SHEETKIT_GATE_ENCRYPTED_FIXTURE=/absolute/path/to/encrypted-scale-10k.xlsx \
SHEETKIT_GATE_PASSWORD=<fixed-password> \
node --import tsx release-gate.ts --generate-encrypted
shasum -a 256 /absolute/path/to/encrypted-scale-10k.xlsx
```

Record its SHA-256 and give the identical bytes and password to every revision.
Do not include fixture generation, workbook construction, pre-open work, or
output cleanup in the timed region. Set `SHEETKIT_GATE_ENCRYPTED_FIXTURE` and
`SHEETKIT_GATE_PASSWORD` for `encrypted-open`; the ordinary fixture argument is
still required as run metadata.

`scale-10k.xlsx` has no coordinate gaps, so its V2 occupied-row transfer
materializes the same dense row array on all three revisions. Explicit dense
requests containing coordinate gaps remain covered by deterministic bounded
invariants rather than a ratio against the pre-1.0 API.

## Peak memory

Measure each primary operation in its own fresh process. On macOS, run the
following five times per revision, operation, and fixture, changing `<run>` from
1 through 5:

```sh
cd /absolute/path/to/<revision>/benchmarks/node
SHEETKIT_GATE_OPERATION=ordinary-open \
SHEETKIT_GATE_FIXTURE=/absolute/path/to/scale-10k.xlsx \
SHEETKIT_GATE_OUTPUT=/absolute/path/to/results/<revision>/cold-ordinary-open-scale-10k-<run>.json \
SHEETKIT_GATE_SAMPLES=1 \
SHEETKIT_GATE_WARMUPS=0 \
/usr/bin/time -l \
node --expose-gc --import tsx release-gate.ts \
2> /absolute/path/to/results/<revision>/rss-ordinary-open-scale-10k-<run>.txt
```

The JSON is the cold latency sample. The complete `/usr/bin/time -l` output is
the isolated peak-memory evidence. The direct Node command deliberately
excludes the pnpm process from RSS. Use the median of the five maximum resident
set sizes. Per-operation before/after RSS deltas from a shared process are
diagnostic only and cannot pass the memory gate.

## Decision rules

Compare the release candidate independently with both baselines. Positive
percentages mean regression. For each comparison, bootstrap the difference of
medians from the two raw sample arrays with 10,000 resamples and a recorded
random seed. Report the raw median ratio and the 95% percentile confidence
interval for the regression ratio.

- Ignore a percentage-only latency change below 0.5 ms for Node or 0.05 ms for
  Rust.
- Ignore a percentage-only peak-memory change below 16 MiB.
- If the absolute change is below its floor, classify it as noise.
- Otherwise, a confidence-interval upper bound at or below 5% passes.
- A confidence interval wholly above 5% and at or below 10% requires profiling
  and written justification.
- A confidence-interval lower bound above 10% is a release stop unless the user
  explicitly accepts it.
- An interval crossing 5% or 10% is indeterminate and must be repeated; it
  cannot be declared a pass or stop.
- Never average a stopped workload away with an improvement elsewhere.

Also record the non-timing invariants: sparse traversal cost is independent of
coordinate distance, rejected package work is independent of declared
expansion, formula work remains bounded by its budgets, streaming incremental
memory does not scale with the full worksheet, and output cost is linear in
input plus emitted bytes.

## Required receipt

Retain:

- environment and source metadata;
- dependency and fixture checksums;
- exact adapter and runner source;
- raw sample arrays and Criterion estimates;
- fresh-process `/usr/bin/time -l` output;
- per-workload medians, confidence/noise information, absolute deltas, and
  ratios against both baselines;
- profiles and justification for every 5-10% result;
- a separate PASS or STOP verdict for every primary workload.

Only after all rows pass, or an exception is explicitly accepted, may the U14
receipt state a final R7 and release-readiness verdict.
