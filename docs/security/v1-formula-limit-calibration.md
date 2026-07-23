# SheetKit 1.0 Formula-Limit Calibration

Status: tentative U4 calibration for U4, U5, and U9 implementation. These
values are not frozen SheetKit 1.0 API commitments. U12 must recheck boundary
fixtures, Rust/Node parity, interoperability, and R7 performance evidence
before freezing them.

## Scope and evidence

This calibration measures formula-bearing core tests, Node.js tests, examples,
benchmark source, and every formula cell in the benchmark `.xlsx` fixtures
present in this workspace. The machine-readable artifact is:

`docs/plans/artifacts/2026-07-15-001-fix-security-hardening-v1/u4-calibration/formula-corpus-metrics.json`

The artifact was measured from revision
`150e57939ac816a9bfa86daa856b432a135e3ca9`. It records the source and fixture
hashes, corpus digests, metric definitions, workbook graph measurements, and
exact tentative integer limits. The input-path SHA-256 manifest is
`237de740f1abcd4274a6dfea839fc5b84c925995c65d3cf4dcfcdc584abb697f`.

The scan found 58,372 formula occurrences: 365 in formula-bearing source and
58,007 in benchmark workbook cells. The current parser accepted 58,360
occurrences representing 58,276 distinct formula strings. The valid source
subset contains 289 core-test, 44 Node-test, 2 example, and 25 benchmark-source
occurrences; 58,000 valid occurrences come from workbook fixtures.

Only formulas accepted by the current `parse_formula` contract contribute to
the maxima. That contract excludes a leading `=`. Formula-like validation
strings and deliberately invalid parser cases remain in the invalid digest but
do not inflate valid limits. Generated benchmark templates are measured from
their expanded workbook cells where available.

The parser exposes an AST but no token stream. Token count therefore uses the
artifact's deterministic lexical definition: a maximal number,
identifier/reference fragment, string literal, or error literal is one token;
operators and punctuation are individual tokens; whitespace is ignored. AST
nodes and depth come from the current `Expr` tree. Root depth is 1. Every
function argument, binary child, unary operand, or parenthesized expression
adds one level. Binary-chain depth separately counts consecutive
left-associated `BinaryOp` nodes because a flat chain can still recurse deeply
during evaluation.

No wall-clock time, peak RSS, allocation count, or stack high-water mark was
measured. All figures below are deterministic input, structure, or work-count
measurements; none is a runtime claim.

## Valid formula envelope

| Dimension | Observed maximum | Example |
| --- | ---: | --- |
| Input bytes | 65 | Inline-list string formula in `generate-fixtures.ts` |
| Lexical tokens | 29 | Mixed sheet references, string, and `LOG10` |
| AST nodes | 10 | `FV(0.06/12,10,-200,-500,1)` |
| AST depth | 5 | Five-term reference expression |
| Binary-chain depth | 4 | Five-term addition chain |
| Expanded range cells per formula | 20 | Two ten-cell ranges |
| Dependency reference edges per formula | 20 | Two ten-cell ranges |
| Wildcard-pattern bytes | 3 | `SEARCH("A*C","ABCABC")` |

The source corpus is shallow and small. It is useful for proving that finite
defaults retain substantial headroom, but it does not independently justify
the hard ceilings. Microsoft documents 8,192 characters for formula contents,
64 nested function levels, and 32,767 characters in one cell in its
[Excel specifications and limits](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3).
Those interoperability limits inform the tentative structural bounds; they do
not replace byte and operation accounting.

## Ordinary workbook graph envelope

| Fixture | Valid formula cells | Input bytes | Tokens | AST nodes | Reference edges | Internal formula edges | Ready-queue ops |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `formulas.xlsx` | 50,000 | 798,980 | 280,000 | 180,000 | 100,000 | 0 | 100,000 |
| `mixed-workload.xlsx` | 8,000 | 84,489 | 32,000 | 30,000 | 13,000 | 3,000 | 16,000 |

Ready-queue work is defined as one push plus one pop per acyclic formula node.
`formulas.xlsx` is the largest observed calculation graph and establishes the
operation-wide headroom baseline. Its 50,000 formulas all depend on non-formula
cells. In `mixed-workload.xlsx`, 3,000 formula cells depend on another formula
cell, so 5,000 cells are initially ready.

Seven `Summary` formulas in `mixed-workload.xlsx` incorrectly retain a leading
`=` and are invalid under the current parser contract. Removing exactly that
one byte only for sensitivity analysis yields ranges of 2,000 or 3,000 cells,
18,000 expanded cells in total. These values are not included in the valid
maxima, but the tentative range budget preserves more than seven times that
intended ordinary workload.

Neither workbook stores cached formula results, so it provides no output-byte
measurement. The existing text-function tests accept the six-byte result of
`REPT("AB",3)` and reject a requested 32,768-character result. This establishes
an interoperability boundary, not a representative output distribution.

## Tentative per-formula structural limits

| Dimension | Secure default | Corpus headroom | Non-overridable hard ceiling |
| --- | ---: | ---: | ---: |
| Input bytes | 32,768 | 504.12x | 65,536 |
| Nesting depth | 64 | 12.80x | 128 |
| Binary-chain depth | 512 | 128.00x | 4,096 |
| Lexical tokens | 16,384 | 564.97x | 65,536 |
| AST nodes | 16,384 | 1,638.40x | 32,768 |
| String output bytes | 131,068 | No representative output corpus | 131,068 |
| String output characters | 32,767 | Existing semantic boundary | 32,767 |

Input bytes, tokens, AST nodes, nesting, and binary-chain depth are checked per
formula before proportional reserve or recursion. The 32 KiB input default can
represent 8,192 four-byte Unicode scalar values; a separate 8,192-character
interoperability check remains an implementation decision rather than being
silently conflated with the byte budget.

The string-output ceiling is four UTF-8 bytes times 32,767 Unicode scalar
values. Both the character and checked byte product must pass before allocation.
This fixed structural ceiling is distinct from the calculation-wide cumulative
output budget below.

## Tentative operation-wide budgets

| Dimension | Secure default | Corpus headroom | Non-overridable hard ceiling |
| --- | ---: | ---: | ---: |
| Formula input bytes | 4 MiB | 5.25x `formulas.xlsx` | 32 MiB |
| Lexical tokens | 1,048,576 | 3.75x `formulas.xlsx` | 8,388,608 |
| AST nodes | 1,048,576 | 5.83x `formulas.xlsx` | 8,388,608 |
| Expanded range cells | 131,072 | 7.28x normalized sensitivity total | 1,048,576 |
| Dependency edges | 524,288 | 5.24x `formulas.xlsx` | 4,194,304 |
| Ready-queue operations | 524,288 | 5.24x `formulas.xlsx` | 4,194,304 |
| Evaluation steps | 1,048,576 | 5.83x AST-node lower-bound proxy | 8,388,608 |
| Matcher work | 1,048,576 | Not calibrated by normal corpus | 16,777,216 |
| Output bytes | 16 MiB | Not calibrated by normal corpus | 64 MiB |

All operation figures are cumulative. Releasing an AST or intermediate string
does not refund work already performed. A single `calculate_all` owns one
`FormulaBudget`; parsing, graph construction, scheduling, evaluation, and every
U5/U9 helper borrow it. Direct `evaluate_formula` creates one fresh operation
budget. `set_cell_formula` and `fill_formula` apply the per-formula structural
checks but must not spend calculation work until evaluation.

`evaluation_steps` charges AST visits plus analytically precharged loop work.
Numeric functions should charge the checked iteration count before entering a
loop. `matcher_work` charges wildcard state transitions, comparisons, and
backtracks. Criteria families additionally charge range expansion through the
same operation budget. Output-producing helpers charge the checked requested
byte count before reserve and actual bytes as they append. A conservative
precharge may exceed actual work; it must never undercharge an attacker-selected
loop or reset in a nested helper.

The hard ceilings apply after default, finite-profile, and per-operation option
resolution. A caller may raise one finite dimension to its ceiling but cannot
disable it, exceed it, or implicitly raise another dimension. Every conversion
and sum uses checked arithmetic. Exact-limit work is accepted; the first unit
over returns the stable formula resource error. A hard ceiling is not an
allocation promise: fallible reservation remains required.

## Why these values remain tentative

The per-formula structural defaults have large corpus headroom and align with
Excel's documented formula and nesting envelope. The cumulative parse, graph,
and queue defaults use powers of two with at least 3.74x headroom over the
largest valid benchmark workload. The range default also accepts the intended,
but currently malformed, mixed-workload summary with 7.28x total headroom.

Evaluation, matching, and output are less certain. The corpus has no valid deep
nesting, wide binary chain, large range, large accepted output, adversarial
wildcard, or high-iteration numeric fixture. U4, U5, and U9 must add the plan's
boundary and cumulative-bypass fixtures before treating these defaults as
supported. In particular:

- U4 must measure exact-minus/exact/plus-one structural and cumulative
  boundaries, repeated parser calls, large dependency graphs, queue churn, and
  unchanged workbook state after failure.
- U5 must instrument periods/iterations, criteria range scans, wildcard search,
  and engineering conversion widths using the work-unit definitions above.
- U9 must cover factorial, depreciation, Bessel, search, substitute, and every
  remaining width/output helper at its analytical boundary.
- U12 must compare ordinary scalar and `calculate_all` workloads against
  revision `150e579` and v0.5.1 under the R7 cold/warm and process-isolated
  policy. Static counts do not substitute for those latency and memory runs.

If measured normal workloads approach a default, increase only that named
default while retaining a finite ceiling. If stack or allocation evidence makes
a proposed ceiling unsafe, lower the ceiling and record the compatibility
effect; corpus headroom alone does not override safety evidence.

## Reproduction and validation

The measurement used a temporary, read-only analyzer because this unit owns
only this report and JSON artifacts. The exact command and SHA-256 hashes of the
scanner, Rust metric source, Cargo manifest, and Cargo lockfile are recorded in
the JSON artifact. At measurement time the sources were under
`/tmp/sheetkit-formula-calibration/`.

From the repository root, validate the artifact shape and immutable corpus
identity:

```sh
jq -e '
  .schema_version == 1 and
  .revision == "150e57939ac816a9bfa86daa856b432a135e3ca9" and
  .runtime_measurements_recorded == false and
  .scope.all_occurrences == 58372 and
  .scope.parser_valid_occurrences == 58360 and
  .corpus_digests.valid_occurrences_sha256 ==
    "596c76f38605177b268634f1f13c4b791a8888903de2377fd00dc10ce9084545"
' docs/plans/artifacts/2026-07-15-001-fix-security-hardening-v1/u4-calibration/formula-corpus-metrics.json
```

Verify the key live inputs used by the measurement:

```sh
jq -r '.key_input_hashes | to_entries[] | "\(.value)  \(.key)"' \
  docs/plans/artifacts/2026-07-15-001-fix-security-hardening-v1/u4-calibration/formula-corpus-metrics.json \
  | shasum -a 256 -c -
```

When the temporary analyzer sources are present and match their recorded
hashes, rebuild and reproduce the artifact values with:

```sh
shasum -a 256 \
  /tmp/sheetkit-formula-calibration/scan.py \
  /tmp/sheetkit-formula-calibration/src/main.rs \
  /tmp/sheetkit-formula-calibration/Cargo.toml \
  /tmp/sheetkit-formula-calibration/Cargo.lock

cargo build \
  --manifest-path /tmp/sheetkit-formula-calibration/Cargo.toml

python3 /tmp/sheetkit-formula-calibration/scan.py \
  /Users/Nebuleto/Workspace/sheetkit \
  /tmp/sheetkit-formula-calibration/target/debug/sheetkit-formula-calibration
```

The emitted JSON's counts, corpus digests, maxima, and fixture graphs must match
the artifact. A changed input hash intentionally invalidates the comparison and
requires a new calibration artifact rather than editing observed values by hand.
