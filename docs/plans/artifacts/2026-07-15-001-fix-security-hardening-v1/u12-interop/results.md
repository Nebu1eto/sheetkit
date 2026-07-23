# U12 interoperability, atomicity, and exploit replay

Observed: 2026-07-23

This receipt records correctness and interoperability observations only. It
contains no wall-clock or release-performance verdict. Performance remains
provisional and must be repeated on the user-requested idle, clean machine.

## Environment

- macOS arm64
- Rust `1.93.0`; Cargo `1.93.0`
- Node.js `24.18.0`; current-source release ESM addon
- macOS arm64 addon SHA-256:
  `0c4d141d61f9193efadc3e7ad5e002d7c3af8532233c1ab0a417e7a839c40ad1`
- Generated `binding.d.ts` SHA-256:
  `357296b431ef66b5b123b55cf0129a223b72f0d937672545f8cdfe87279f80dd`
- Generated `index.js` SHA-256:
  `99f07e674c6f0f517cac74ea22ec633a2725dec100df43a4325e491f32359bb0`
- pnpm `10.29.1`
- LibreOfficeDev `26.8.0.0.alpha0`
- msoffcrypto-tool `6.0.0`
- Microsoft Excel `16.111.1` is installed but was not automated in this run.
  Excel is not counted as a passed check; LibreOffice supplied the observed
  desktop spreadsheet round trips in this receipt.

## Validation rubric

- [x] Replay attacker-controlled packages and data through the real Rust, Node,
  and CLI entry points.
- [x] Prove LibreOffice, msoffcrypto-tool, and checked-in Office fixture
  interoperability semantically rather than by synthetic byte equality.
- [x] Inspect macro-free output and prove VBA bytes, relationships, and content
  types are stripped or preserved according to the selected target format.
- [x] Prove failed staged writes preserve the destination, remove temporary
  output, and leave the workbook reusable.
- [x] Prove CLI diagnostics and structured output remain deterministic,
  terminal-safe, and atomic.

## Results

| Check | Result | Observed evidence |
|---|---|---|
| Selected Rust exploit corpus | Pass | 79/79 across archive, encrypted-container, Office fixture, VBA, CSV, SVG, streaming, and password path/buffer parity targets |
| Workbook I/O integration | Pass | 140/140, including macro stripping/preservation, encrypted target-format handling, destination preservation, and post-failure reuse |
| Atomic temporary cleanup | Pass | Final current-source exact core and CLI replays passed: the injected stage leaves only the existing destination, the workbook remains reusable, and failed CLI conversion leaves one destination entry and empty stdout |
| CLI output/diagnostics | Pass | 27/27, including spreadsheet-safe CSV, explicit raw exchange, output-limit atomicity, destination preservation, and terminal-control sanitation |
| Canonical fuzz/exploit replay | Pass | 3/3; all 33 checked-in seeds replay deterministically, 2,048 structured mutations are panic-free, and transfer exact/extreme cases pass |
| Node.js security parity on Node 24 | Pass | Current-source release ESM addon focused replay passed 79/79; an expanded receipt replay including the ResourceLimitError contract passed 83/83; the complete Node suite passed 644/644 across 18 files |
| Generated/public declarations | Pass | `binding.d.ts` and public declarations passed the Node 24 consumer fixture and typecheck; generated signatures include `JsSaveLimits`, async password save, and formula mutation limits |
| SheetKit workbook generation | Pass | Rust example generated a four-sheet workbook, reopened its encrypted form, and recovered `Sheet1!A1 = Employee List` |
| LibreOffice workbook round trip | Pass | Headless XLSX-to-XLSX conversion retained four sheets, active sheet, title, creator, and `Sheet1!A1` |
| SheetKit CSV to LibreOffice | Pass | SheetKit emitted 73 bytes with four CRLF records and no lone LF; LibreOffice converted it to XLSX and SheetKit reopened `output!A1 = Employee List`. A neutralized `'=1+1` field remained a string and was not evaluated |
| SheetKit encryption to msoffcrypto | Pass | msoffcrypto identified Agile encryption, decrypted successfully, and the resulting valid ZIP reopened with four sheets and the expected A1 value |
| msoffcrypto encryption to SheetKit | Pass | The rebuilt Node 24 addon reopened the checked-in msoffcrypto-tool 6 Agile fixture through sync/async path and buffer APIs; it also opened the freshly generated external Agile file and recovered four sheets plus `Sheet1!A1 = Employee List` |
| Macro-free package inspection | Pass | Generated XLSX has 24 entries, no VBA/macro path, no macro-enabled content type, and no VBA relationship |
| VBA preservation | Pass | Final current-source exact macro-free and encrypted-target tests passed; they preserve `xl/vbaProject.bin` bytes, relationship, and content type for XLSM and remove all three for XLSX |

The msoffcrypto-decrypted XLSX is not byte-identical to the source because the
tool rewrites ZIP compression. `unzip -t`, SheetKit package parsing, workbook
metadata, sheet count, and the sentinel cell all passed, so the comparison is
semantic rather than a synthetic byte equality.

No `.sheetkit-*` temporary file remained under the interoperability output
directory after successful operations. The rejected core and CLI regression
tests each assert that the existing destination is the only remaining directory
entry; the core test then saves the same workbook successfully and verifies the
preserved VBA bytes.

## Remaining uncertainty

- Microsoft Excel automation was not observed in this run. LibreOffice supplied
  the desktop spreadsheet semantic checks.
- Local Node/N-API execution covers the macOS arm64 artifact. Cross-platform
  package builds remain a repository release-gate responsibility.
- Deterministic fuzz corpus replay is recorded here; instrumented fuzz and
  sanitizer scope is recorded separately in the U12 fuzz receipt.
- Performance remains intentionally unjudged until the user-requested idle,
  clean-machine rerun.

## Commands

Selected Rust exploit and interoperability targets:

```sh
cargo test -p sheetkit-core --features encryption \
  --test open_budget_regression \
  --test encrypted_open_budget_regression \
  --test office_encryption_fixture_regression \
  --test workbook_u3_security_regression \
  --test vba_security_regression \
  --test csv_security_regression \
  --test render_security \
  --test workbook_stream_budget_regression

cargo test -p sheetkit-core --features encryption workbook::io::tests::

cargo test -p sheetkit-core --features encryption \
  workbook::io::tests::failed_atomic_xlsx_stage_preserves_destination_and_workbook \
  -- --exact

cargo test -p sheetkit-core --features encryption \
  --test security_fuzz_corpus_replay

cargo build -p sheetkit --features cli
cargo test -p sheetkit --features cli --test cli_tests
```

Node 24 replay commands:

```sh
pnpm --filter @sheetkit/node build
mise exec node@24 -- pnpm --filter @sheetkit/node typecheck
mise exec node@24 -- pnpm --filter @sheetkit/node typecheck:declarations
# The eight security-* files below are the 79-test core subset. Including
# resource-limit-error.spec.ts produces the expanded 83-test receipt replay.
mise exec node@24 -- pnpm --filter @sheetkit/node exec vitest run \
  __test__/resource-limit-error.spec.ts \
  __test__/security-formula-mutation.spec.ts \
  __test__/security-encrypted-vba.spec.ts \
  __test__/security-json-import.spec.ts \
  __test__/security-json.spec.ts \
  __test__/security-node-error-parity.spec.ts \
  __test__/security-output.spec.ts \
  __test__/security-streaming.spec.ts \
  __test__/security-traversal.spec.ts
mise exec node@24 -- pnpm test
```

Workbook generation and msoffcrypto interoperability:

```sh
mkdir -p \
  /tmp/sheetkit-u12-interop-20260723/example \
  /tmp/sheetkit-u12-interop-20260723/lo-xlsx \
  /tmp/sheetkit-u12-interop-20260723/lo-csv \
  /tmp/sheetkit-u12-interop-20260723/lo-safe-csv
cd /tmp/sheetkit-u12-interop-20260723/example

cargo run \
  --manifest-path /Users/Nebuleto/Workspace/sheetkit/examples/rust/Cargo.toml \
  --quiet

uvx --from msoffcrypto-tool msoffcrypto-tool \
  -p secret123 output_encrypted.xlsx decrypted-by-msoffcrypto.xlsx

uvx --from msoffcrypto-tool msoffcrypto-tool \
  -p external123 -e output.xlsx encrypted-by-msoffcrypto.xlsx

uvx --from msoffcrypto-tool msoffcrypto-tool \
  -v -t output_encrypted.xlsx

npx --yes node@24 --input-type=module - <<'JS'
import { Workbook } from '/Users/Nebuleto/Workspace/sheetkit/packages/sheetkit/index.js'
const workbook = Workbook.openWithPasswordSync(
  'encrypted-by-msoffcrypto.xlsx',
  'external123',
)
if (workbook.getCellValue('Sheet1', 'A1') !== 'Employee List') process.exit(1)
JS

python3 - <<'PY'
from zipfile import ZipFile

with ZipFile('output.xlsx') as archive:
    names = archive.namelist()
    content_types = archive.read('[Content_Types].xml').decode()
    relationships = archive.read('xl/_rels/workbook.xml.rels').decode()
    assert not [name for name in names if 'vba' in name.lower() or 'macro' in name.lower()]
    assert 'vba' not in content_types.lower()
    assert 'macroenabled' not in content_types.lower()
    assert 'vba' not in relationships.lower()
PY
```

LibreOffice and CSV interoperability:

```sh
soffice -env:UserInstallation=file:///tmp/sheetkit-u12-interop-20260723/lo-profile \
  --headless --convert-to xlsx \
  --outdir /tmp/sheetkit-u12-interop-20260723/lo-xlsx output.xlsx

/Users/Nebuleto/Workspace/sheetkit/target/debug/sheetkit info \
  /tmp/sheetkit-u12-interop-20260723/lo-xlsx/output.xlsx
/Users/Nebuleto/Workspace/sheetkit/target/debug/sheetkit get \
  /tmp/sheetkit-u12-interop-20260723/lo-xlsx/output.xlsx Sheet1 A1

/Users/Nebuleto/Workspace/sheetkit/target/debug/sheetkit convert --format csv \
  --output /tmp/sheetkit-u12-interop-20260723/output.csv output.xlsx

soffice -env:UserInstallation=file:///tmp/sheetkit-u12-interop-20260723/lo-profile \
  --headless --convert-to xlsx \
  --outdir /tmp/sheetkit-u12-interop-20260723/lo-csv \
  /tmp/sheetkit-u12-interop-20260723/output.csv

/Users/Nebuleto/Workspace/sheetkit/target/debug/sheetkit get \
  /tmp/sheetkit-u12-interop-20260723/lo-csv/output.xlsx output A1

/Users/Nebuleto/Workspace/sheetkit/target/debug/sheetkit set \
  /tmp/sheetkit-u12-interop-20260723/example/output.xlsx Sheet1 A1 '=1+1' \
  --output /tmp/sheetkit-u12-interop-20260723/formula-text.xlsx
/Users/Nebuleto/Workspace/sheetkit/target/debug/sheetkit convert \
  /tmp/sheetkit-u12-interop-20260723/formula-text.xlsx \
  --format csv \
  --output /tmp/sheetkit-u12-interop-20260723/safe-formula.csv
soffice -env:UserInstallation=file:///tmp/sheetkit-u12-interop-20260723/lo-profile \
  --headless --convert-to xlsx \
  --outdir /tmp/sheetkit-u12-interop-20260723/lo-safe-csv \
  /tmp/sheetkit-u12-interop-20260723/safe-formula.csv
```

## Ephemeral output checksums

These files live under `/tmp/sheetkit-u12-interop-20260723` and are not release
artifacts. The hashes make this receipt independently auditable for the current
run.

```text
e2b9cd42ba7ac2a42b1728b69c45b5a5df34509c7b093987262ea4133f30100b  example/output.xlsx
293d03b482ff8f871eeb775ba4c41d035a24ed9b6c1819d15db7f3bc4fb846b3  example/output_encrypted.xlsx
e57b51510b2b8dbabe904e60da686b6d892adb4c1cb1a4f2ad55198d1fde67df  example/decrypted-by-msoffcrypto.xlsx
9b8e9f164a5dde1bc49d40661e2a7986537817be4265d43ffd6f11bb5c848964  example/encrypted-by-msoffcrypto.xlsx
396234a9da4a94e660efb4a31f8b8da7bda47be18440102dfa56a774a3b50ece  lo-xlsx/output.xlsx
39526d835c36c9667b21f5c2b13a09efae04c63425d21417a16ed63dc6f18077  output.csv
7dd6307ba913a0e66c3445296194bd544779fe174e9e1c427962e88fb990bb79  lo-csv/output.xlsx
4f4cf17ad70be3b326c34efee7ce84c21fcc4f8ed567bcc06fa875b438cd8000  formula-text.xlsx
e1ea4b22c3b86fd2b321556fdd83d9bf7fd51c08ff4494d24a11cd0381c3e0b1  safe-formula.csv
1c90e93cc3790d5b2a4a0c77814d3540c4cc88258f990eab3bf0705a900d6ae8  lo-safe-csv/safe-formula.xlsx
```
