# U1 dual-baseline commands

All commands ran on 2026-07-15 from
`/Users/Nebuleto/Workspace/sheetkit`. The temporary root was
`/tmp/sheetkit-u1-20260715`. No Git worktree was created.

## Source isolation

```sh
git archive 150e57939ac816a9bfa86daa856b432a135e3ca9 | \
  tar -x -C /tmp/sheetkit-u1-20260715/current
git archive f9eac5366eb83b6dd79d5dd6d03a8988b5c319a8 | \
  tar -x -C /tmp/sheetkit-u1-20260715/tag
```

The ignored `benchmarks/node/fixtures` directory was copied once to
`/tmp/sheetkit-u1-20260715/fixtures` and symlinked into both archives.
`fixtures.sha256` records the inputs.

## Dependency resolution

```sh
cd /tmp/sheetkit-u1-20260715/current
cargo generate-lockfile
cp Cargo.lock ../tag/Cargo.lock
pnpm install --frozen-lockfile

cd /tmp/sheetkit-u1-20260715/tag
pnpm install --frozen-lockfile
```

Both Cargo lockfiles had SHA-256
`40016c54b483615aaca6875c4274a0a749fd3237b228fe1c840de7be96e98809`.
All three pnpm lockfiles had SHA-256
`608513dffc382d0f98f2d623d9ee8618fec66011a07efcabbd6fb362c140ba95`.

## Rust focused benchmark

Each revision was built first with its own `CARGO_TARGET_DIR`:

```sh
CARGO_TARGET_DIR=/tmp/sheetkit-u1-20260715/targets/<revision> \
  cargo bench --locked -p sheetkit-core --bench open_modes --no-run
```

The following command then ran for tag and current, alternating revisions for
each group:

```sh
CARGO_TARGET_DIR=/tmp/sheetkit-u1-20260715/targets/<revision> \
  cargo bench --locked -p sheetkit-core --bench open_modes -- \
  '<group>/' --noplot
```

`<group>` was `^open_latency`, `^open_readfast_latency`, `^get_rows`, then
`^save_latency`. The first three groups passed on both revisions. The save group
failed identically on both revisions with exit 101 because the existing
benchmark passes an extensionless `NamedTempFile` to `Workbook::save`.

## Node focused benchmark

Each N-API package was built with a separate Cargo target:

```sh
CARGO_TARGET_DIR=/tmp/sheetkit-u1-20260715/targets/<revision>-node \
  pnpm --filter @sheetkit/node build
```

The focused benchmark then ran once per revision:

```sh
pnpm --dir benchmarks/node bench:baseline
```

For an aggregate fresh-process high-water RSS signal, the same runner was
executed once more per revision:

```sh
cd benchmarks/node
/usr/bin/time -l node --expose-gc --import tsx bench-baseline.ts >/dev/null
```

The JSON from the first timed-content run is retained. The second invocation
exists only for the process-level RSS evidence in
`node-bench-baseline.process-rss.txt`.
