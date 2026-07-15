# SheetKit 1.0 Resource-Limit Calibration

Status: tentative U1 calibration for U2 implementation. These values are not
the frozen SheetKit 1.0 API contract. U12 must recheck the accepted corpus,
near-limit regressions, interoperability corpus, and performance evidence before
freezing public defaults.

## Scope and evidence

This calibration covers one workbook-open operation over 14 `.xlsx` fixtures:
13 ZIP packages and one encrypted CFB container. Twelve generated benchmark
fixtures are intentionally ignored by Git; two encryption interoperability
fixtures are tracked. The measurement artifact records a SHA-256 hash for every
file, the container type, outer size, ZIP entry count, declared sizes,
compression ratios, XML sizes, and CRC result:

`docs/plans/artifacts/2026-07-15-001-fix-security-hardening-v1/u1-calibration/fixture-metrics.json`

The artifact was measured at revision
`150e57939ac816a9bfa86daa856b432a135e3ca9`. Its SHA-256 is
`5a77d053355d25809461ffc75782d5c95e21b065dd6547c4ff3a7b9ae0c06219`.
All 14 current fixture hashes match the artifact. Python independently reopened
all 13 ZIP packages, validated their CRCs, and reproduced every recorded
central-directory measurement and summary maximum. The CFB fixture was checked
only for outer size, signature, and hash; this corpus does not calibrate its
decrypted package or KDF cost.

The ratio in this report is uncompressed bytes divided by compressed bytes.
Entries with zero compressed bytes do not contribute a finite observed ratio.
U2 must treat a non-empty result from zero compressed bytes as an unbounded
ratio and reject it.

## Supported single-operation envelope

The evidence supports normal ZIP workbook opens up to the observed envelope
below. It includes scalar, formula, string, style, validation, comment, merged
cell, and multi-sheet benchmark workbooks up to 100,000 rows. It does not prove
that inputs outside this envelope are unsafe or unsupported; it only establishes
the valid corpus against which finite defaults receive documented headroom.

| Dimension | Observed maximum | Fixture-derived meaning |
| --- | ---: | --- |
| Outer input | 7,582,404 bytes (7.23 MiB) | Largest `.xlsx` file |
| Archive entries | 17 | Largest central-directory entry count |
| Per-entry declared bytes | 34,597,993 bytes (33.00 MiB) | Largest ZIP member |
| Aggregate declared bytes | 41,770,451 bytes (39.84 MiB) | Largest sum of ZIP members |
| Compression ratio | 26.260155:1 | Largest finite per-entry ratio |
| Single XML part | 34,597,993 bytes (33.00 MiB) | Largest `.xml` member |
| Aggregate XML | 41,769,299 bytes (39.83 MiB) | Largest sum of `.xml` members |
| Archive-material proxy | 49,352,855 bytes (47.07 MiB) | Maximum outer plus aggregate declared bytes |

The archive-material proxy is not a heap or retained-memory measurement. It is
a conservative content-byte proxy used only to select initial retained-byte
headroom. Parsed Rust structures, allocator overhead, string decoding, shared
ownership, and temporary copies can make process memory larger.

## Tentative U2 limits

U2 should use the secure defaults below for every equivalent default open path.
Path, buffer, encrypted-wrapper, and compatibility adapters must normalize to
the same finite profile. Declared limits are preflight signals; actual limits
must be charged from bytes read or produced and cannot trust ZIP metadata.

| Dimension | Secure default | Corpus headroom | Non-overridable hard ceiling |
| --- | ---: | ---: | ---: |
| Outer input bytes | 128 MiB | 17.70x | 2 GiB |
| Archive entry count | 4,096 | 240.94x | 65,536 |
| Per-entry declared bytes | 256 MiB | 7.76x | 1 GiB |
| Per-entry actual bytes | 256 MiB | 7.76x declared proxy | 1 GiB |
| Aggregate declared bytes | 512 MiB | 12.85x | 4 GiB |
| Aggregate actual bytes | 512 MiB | 12.85x declared proxy | 4 GiB |
| Compression ratio | 100:1 | 3.81x | 1,000:1 |
| Single XML part bytes | 256 MiB | 7.76x | 1 GiB |
| Retained package bytes | 512 MiB | 10.88x archive-material proxy | 2 GiB |

All byte constants use binary units. The exact integer defaults are therefore
`134217728`, `268435456`, and `536870912`; hard ceilings use `1073741824`,
`2147483648`, and `4294967296` as applicable.

The hard ceilings apply after profile and per-operation override resolution.
A caller can select a finite value above the default but cannot exceed the hard
ceiling, disable a dimension, or raise a different dimension implicitly. Each
value must also fit the platform address space and pass checked conversion and
fallible reservation; a hard ceiling is not a promise that an allocation will
succeed. Exact-limit input is accepted, and the first unit over the limit is
rejected with the stable resource dimension.

For XML, the table bounds each extracted XML part. The aggregate actual budget
continues to bound the package-wide sum. Retained package bytes count owned raw
package buffers that remain live after extraction or transfer; releasing a
temporary buffer may release its retained charge, while total actual bytes read
remain cumulative for the operation.

## Unmeasured dimensions and follow-up

This calibration does not measure wall-clock runtime, CPU work, fresh-process
peak RSS, allocator high-water marks, repeated reads, or concurrent opens. The
separate performance baseline owns latency and memory regression evidence.
Formula, traversal, rendering, export, streaming, macro, and encryption work
budgets require their unit-specific calibration rather than reuse of these byte
limits.

The following gaps must remain explicit through U12:

- The corpus contains no large encrypted CFB input, `.xlsm`/VBA package, ZIP64
  package, data-descriptor edge case, or adversarial archive.
- Declared sizes are not actual decompression measurements. U2 regression
  fixtures must prove actual per-entry and aggregate accounting with misleading
  headers, truncation, CRC failure, and high-ratio data.
- Retained package bytes use a content proxy, not runtime instrumentation. U2
  must validate the charging lifecycle, and the release performance run must
  record process-isolated peak memory.
- Repeated public calls receive new operation-local budgets. They must not reuse
  a depleted mutable workbook budget or inherit a preceding trusted override.
- A per-operation library budget does not bound aggregate work across concurrent
  calls. Hosts accepting untrusted jobs remain responsible for an aggregate
  concurrency limit, queue, memory limit, and cancellation policy. SheetKit must
  not introduce ambient process-global policy to simulate that host control.

## Reproduction

Environment: macOS 26.5.2 (Darwin 25.5.0), arm64, Apple M4 Pro; Python 3.14.6;
Git 2.55.0; Apple `jq` 1.7.1; `shasum` 6.02; Python standard-library
`zipfile`.

From the repository root, verify artifact shape and every fixture hash:

```sh
jq -e '.schema_version == 1 and
  .revision == "150e57939ac816a9bfa86daa856b432a135e3ca9" and
  (.fixtures | length) == 14 and
  ([.fixtures[].sha256] | unique | length) == 14' \
  docs/plans/artifacts/2026-07-15-001-fix-security-hardening-v1/u1-calibration/fixture-metrics.json

jq -r '.fixtures[] | "\(.sha256)  \(.path)"' \
  docs/plans/artifacts/2026-07-15-001-fix-security-hardening-v1/u1-calibration/fixture-metrics.json \
  | shasum -a 256 -c -
```

Recompute the summary maxima after independently opening the ZIP files:

```sh
python3 - <<'PY'
import json
import pathlib
import zipfile

artifact = pathlib.Path(
    "docs/plans/artifacts/2026-07-15-001-fix-security-hardening-v1/"
    "u1-calibration/fixture-metrics.json"
)
data = json.loads(artifact.read_text())
rows = []
for fixture in data["fixtures"]:
    if fixture["container"] != "zip":
        continue
    with zipfile.ZipFile(fixture["path"]) as archive:
        assert archive.testzip() is None
        entries = archive.infolist()
        rows.append((
            pathlib.Path(fixture["path"]).stat().st_size,
            len(entries),
            sum(entry.file_size for entry in entries),
            max(entry.file_size for entry in entries),
            max(
                entry.file_size / entry.compress_size
                for entry in entries
                if entry.compress_size
            ),
        ))

actual = {
    "fixture_count": len(data["fixtures"]),
    "zip_fixture_count": len(rows),
    "cfb_fixture_count": len(data["fixtures"]) - len(rows),
    "max_outer_bytes": max(
        pathlib.Path(fixture["path"]).stat().st_size
        for fixture in data["fixtures"]
    ),
    "max_archive_entry_count": max(row[1] for row in rows),
    "max_aggregate_declared_uncompressed_bytes": max(row[2] for row in rows),
    "max_single_entry_declared_uncompressed_bytes": max(row[3] for row in rows),
    "max_observed_compression_ratio": round(max(row[4] for row in rows), 6),
}
assert actual == data["summary"], (actual, data["summary"])
print(actual)
PY
```
