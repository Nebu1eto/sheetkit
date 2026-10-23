# 1.0 마이그레이션 가이드

이 문서는 SheetKit `0.5.1` 사용자가 `1.0`으로 이전할 때 확인해야 하는 누적 변경 사항을 설명합니다. Rust, Node.js, CLI의 공개 동작을 함께 다루며, 보안 점검 산출물이나 commit 기록을 따로 읽지 않아도 이전할 수 있도록 구성했습니다.

1.0에서는 잘못되거나 과도한 입력을 허용하던 동작보다 유한한 자원 사용과 원자적 실패를 우선합니다. 따라서 일부 기본값, 반환 형태, 오류, 저장 결과가 의도적으로 달라집니다.

Node.js package의 생성 declaration은 Node.js 24와 Node-API 8을 기준으로 검증되었습니다. Current-source release ESM addon은 Node.js 24.18.0에서 전체 test 646건을 통과했으며 declaration compile도 통과했습니다. Upgrade 후 native addon을 다시 build하고 TypeScript application을 새 declaration으로 typecheck하세요.

## 먼저 확인할 변경 사항

| 영역 | 0.5.1 동작 | 1.0 동작 | 필요한 조치 |
|---|---|---|---|
| 워크북 열기 | ZIP, 암호화 container, XML의 일부 자원 사용이 사실상 무제한이었습니다. | `OpenLimits`, `CryptLimits`, `VbaLimits`가 유한한 기본값과 해제할 수 없는 상한을 적용합니다. | `ResourceLimitError`를 처리하고, 신뢰할 수 있는 큰 파일에만 필요한 항목을 유한하게 높이세요. |
| 읽기 기본값 | eager 처리와 dense 결과를 전제로 한 코드가 있었습니다. | lazy/deferred open과 sparse traversal이 기본값입니다. | 빈 행과 열을 암묵적으로 세던 코드를 수정하고, 필요한 경우에만 bounded dense v1 형식을 선택하세요. |
| 날짜와 수식 | Node.js에서 날짜와 수식 cached result가 평탄화되거나 손실될 수 있었습니다. | `DateValue`와 `FormulaValue`가 직접 호출, raw transfer, stream에서 보존됩니다. | TypeScript의 value narrowing과 직렬화 코드를 갱신하세요. |
| 수식 처리 | 깊거나 넓은 수식, 큰 range, wildcard, 증폭 결과가 과도한 작업을 만들 수 있었습니다. | parse, evaluate, calculate, assign, fill이 한 operation의 `FormulaLimits`를 공유하며 실패 전에 검증합니다. | 극단적으로 생성된 수식을 나누고 typed 자원 오류를 처리하세요. |
| JSON, CSV, HTML | sparse sheet가 큰 dense rectangle로 확장될 수 있었고 CSV가 수식 형태의 text를 그대로 내보냈습니다. | sparse traversal과 출력 byte 제한을 적용하며 CSV는 `spreadsheet-safe`가 기본값입니다. | 자동화 결과를 다시 확인하고, 신뢰할 수 있는 교환에만 `raw-exchange`를 지정하세요. |
| 저장과 stream write | 출력 크기와 stream row의 일부 경계가 없었으며 실패 중간 상태가 노출될 수 있었습니다. | `SaveLimits`와 `StreamLimits`를 적용하며 path 저장과 batch mutation이 원자적으로 실패합니다. | 제한 오류를 처리하고 async 저장 중에는 같은 workbook을 변경하지 마세요. |
| macro와 package | macro workbook을 `.xlsx`로 저장해도 VBA가 남을 수 있었습니다. | 대상 확장자가 macro-free 형식이면 VBA byte, relationship, content type을 제거합니다. | macro를 보존하려면 `.xlsm`, `.xltm`, `.xlam`으로 저장하세요. |
| CLI | 사람이 보는 출력에 terminal control 문자가 포함될 수 있었고 CSV가 raw였습니다. | human output은 control 문자를 보이게 escape하며 CSV는 안전하고 bounded입니다. | 자동화에서는 `--format json`을 사용하고 CSV policy를 명시하세요. |

## 유한한 limit과 typed 오류

제한을 생략해도 무제한으로 처리되지 않습니다. 기본 profile은 operation마다 새로 적용되며 process 전체나 workbook 전체를 변경하는 전역 override는 없습니다. 사용자가 지정한 값도 hard ceiling을 넘을 수 없습니다.

주요 기본 제한은 다음과 같습니다.

| Profile | 주요 기본값 |
|---|---|
| Open | outer input 128 MiB, archive entry 4,096개, aggregate expanded data 512 MiB, XML part 256 MiB, compression ratio 100배입니다. |
| Formula | formula당 input 32 KiB, operation 누적 input 4 MiB, cached result를 포함한 누적 output 16 MiB입니다. |
| Traversal | row 1,048,576개, column 16,384개, cell 10,000,000개, output 512 MiB입니다. |
| Save | archive output 512 MiB, XML part 128 MiB, row 10,000,000개, cell 10,000,000개입니다. |
| Stream | row XML 16 MiB, total row XML 512 MiB, row 1,048,576개, cell 10,000,000개입니다. |
| JSON import | row 1,048,576개, column 16,384개, cell 10,000,000개, retained data 512 MiB입니다. |
| CSV return | UTF-8 output 64 MiB입니다. |

Node.js에서는 자원 제한 오류를 다음과 같이 구분할 수 있습니다. `limit`과 `observed`는 정밀도를 보존하기 위해 `bigint`입니다.

```typescript
import { isResourceLimitError, Workbook } from '@sheetkit/node';

try {
  await Workbook.open('large.xlsx');
} catch (error) {
  if (isResourceLimitError(error)) {
    console.error(error.operation, error.dimension, error.limit, error.observed);
  } else {
    throw error;
  }
}
```

Rust에서는 `Error::ResourceLimitExceeded`를 match하면 `operation`, `category`, `dimension`, `limit`, `observed`, `stage`를 확인할 수 있습니다. 잘못된 argument, lifecycle 오류, password 오류는 자원 제한 오류와 구분됩니다.

## 워크북 열기와 lazy hydration

`Workbook::open`과 Node.js의 path, buffer, password variant는 모두 유한한 기본 제한을 사용합니다. ZIP32/ZIP64 entry count, 개별 entry, aggregate expansion, 압축 비율, outer buffer, 암호화 CFB stream을 allocation이나 큰 read 전에 검사합니다. 암호를 해제한 뒤에도 같은 open budget이 이어지므로 decrypted ZIP에서 제한이 초기화되지 않습니다.

Node.js의 기본값은 `readMode: 'lazy'`, `auxParts: 'deferred'`, `dateInterpretation: 'numFmt'`입니다. `parseMode: 'full'`과 `parseMode: 'readfast'`는 각각 eager와 lazy로 매핑되는 compatibility 이름이며, 알 수 없는 mode는 더 이상 조용히 기본값으로 바뀌지 않습니다. `maxUnzipSize`와 `maxZipEntries`도 유한한 항목만 덮어쓰는 compatibility 이름입니다.

```typescript
const workbook = await Workbook.open('report.xlsx', {
  readMode: 'lazy',
  auxParts: 'deferred',
  dateInterpretation: 'numFmt',
  maxOuterInputBytes: 256 * 1024 * 1024,
  maxXmlPartBytes: 128 * 1024 * 1024,
  maxZipEntries: 8192,
});
```

Rust에서는 builder를 사용하여 같은 정책을 지정할 수 있습니다.

```rust
use sheetkit::{AuxParts, OpenOptions, ReadMode, Workbook};

let options = OpenOptions::new()
    .read_mode(ReadMode::Lazy)
    .aux_parts(AuxParts::Deferred)
    .max_outer_input_bytes(256 * 1024 * 1024)
    .max_xml_part_bytes(128 * 1024 * 1024);
let workbook = Workbook::open_with_options("report.xlsx", &options)?;
```

lazy workbook은 worksheet와 comments, charts, images, pivot, slicer, form control 같은 auxiliary part를 첫 접근까지 package-backed 상태로 보존합니다. `maxXmlPartBytes`는 open 시점뿐 아니라 나중의 hydration에도 적용됩니다. 잘못된 A1 reference나 table relationship은 해당 part를 hydrate할 때 오류가 날 수 있습니다.

`sheets` filter로 제외한 sheet는 압축을 풀거나 hydrate하지 않으며, 저장 시에는 원래 compressed data를 이용하여 복사합니다. 따라서 제외한 sheet의 잘못된 내부 구조까지 검사한다고 가정하면 안 됩니다. 전체 파일을 검증해야 한다면 eager open을 사용하세요. `sheetRows`도 지정한 row에서 parsing을 멈추는 기능이며 파일 전체 검증 기능은 아닙니다.

lazy copy-on-write는 pictures, sparklines, comments, controls, names, merges, tables를 보존합니다. sheet copy/delete와 row/column insert/delete는 먼저 관련 package graph와 수식 reference를 검증한 뒤 원자적으로 반영합니다. lazy 처리 때문에 eager open을 강제하던 workaround는 제거할 수 있습니다.

## 셀 값, 날짜, 수식, reference

Node.js의 cell value는 다음 union으로 처리해야 합니다.

```typescript
type CellValue =
  | string
  | number
  | boolean
  | { type: 'date'; serial: number; iso?: string }
  | {
      type: 'formula';
      formula: string;
      result?: {
        valueType: 'empty' | 'string' | 'number' | 'boolean' | 'date' | 'error' | 'formula';
        value?: string;
        numberValue?: number;
        boolValue?: boolean;
        date?: { type: 'date'; serial: number; iso?: string };
      };
    }
  | null;
```

기본 `dateInterpretation: 'numFmt'`는 Excel이 일반적으로 기록하는 date number format을 보고 number cell을 `DateValue`로 올립니다. OOXML의 cell type만 그대로 따르려면 `'cellType'`을 지정하세요. date와 formula cached result는 direct call, stream, v1/v2 raw transfer, JSON import에서 같은 형태로 보존됩니다.

수식 parse, evaluation, `calculateAll`, `setCellFormula`, `fillFormula`, generic `FormulaValue` mutation은 유한한 `FormulaLimits`를 사용합니다. 하나의 bulk operation은 입력 byte, token, AST node, range cell, dependency edge, ready queue work, evaluation step, wildcard match, output byte를 누적하여 계산합니다. 한 항목이 실패하면 앞선 항목만 반영되는 partial mutation이 남지 않습니다.

```typescript
workbook.setCellFormula('Sheet1', 'C2', 'SUM(A2:B2)', {
  perFormulaInputBytes: 32 * 1024,
  operationInputBytes: 4 * 1024 * 1024,
  operationOutputBytes: 16 * 1024 * 1024,
});

workbook.calculateAll({
  operationRangeCells: 1_000_000,
  operationEvaluationSteps: 5_000_000,
});
```

Rust에서는 `set_cell_formula_with_limits`, `fill_formula_with_limits`, `evaluate_formula_with_limits`, `calculate_all_with_limits`를 사용합니다. 기본 method도 `FormulaLimits::secure_default()`를 사용하므로 무제한 compatibility path가 아닙니다.

1.0의 계산 결과도 일부 달라집니다. range, blank, power, financial, holiday, `TEXT`, Unicode `LEN`, Bessel 함수가 Excel에 맞게 수정되었습니다. cell 단위 domain 오류는 해당 formula의 cached error가 되며, cycle이나 engine 오류만 전체 계산을 중단합니다. `COUNTIF(S)`, `SUMIF(S)`, `AVERAGEIF(S)`, `SEARCH`의 wildcard 처리는 bounded matcher를 사용하며 극단적인 pattern은 거부될 수 있습니다. `FACT`, `DB`, `DDB`, `CUMIPMT`, `CUMPRINC`, `BESSELK`, `BESSELY`, `SUBSTITUTE`도 작업량이나 출력 증폭을 미리 검사합니다. `BIN2HEX`, `BIN2OCT`, `DEC2BIN`, `DEC2HEX`, `DEC2OCT`, `HEX2BIN`, `HEX2OCT`, `OCT2BIN`, `OCT2HEX`의 `places`는 1부터 10까지만 허용됩니다.

Rust의 `CompactCellRef::new`와 `CompactCellRef::from_coordinates`는 panic 대신 `Result`를 반환하며 Excel의 row와 column 범위를 엄격하게 적용합니다. row/column 구조 변경은 문자열 치환이 아니라 token-aware reference shift를 사용하므로 literal, qualified reference, validation, table, chart, name 같은 종속 항목이 함께 갱신됩니다. 이전의 잘못된 textual shift 결과를 snapshot으로 고정했다면 새 결과로 바꾸세요.

## Sparse traversal, raw transfer, JSON, CSV, HTML

`getRows`, `getCols`, `getRowsRaw`, `getRowsIterator`, `toJSON`, `toCSV`, `toHTML`은 occupied row와 cell을 기준으로 처리합니다. 큰 coordinate gap은 더 이상 빈 객체나 row를 자동으로 생성하지 않습니다.

```typescript
const rows = workbook.getRows('Sheet1', {
  mode: 'sparse',
  limits: {
    maxRows: 100_000,
    maxColumns: 1000,
    maxCells: 1_000_000,
    maxOutputBytes: 128 * 1024 * 1024,
  },
});

const columns = workbook.getCols('Sheet1', {
  limits: { maxCells: 1_000_000 },
});
```

v2 raw buffer는 absolute row와 column coordinate를 담는 canonical sparse 형식입니다. v1도 기본값은 sparse이며, 빈 row가 필요한 legacy consumer만 `mode: 'dense'`를 명시할 수 있습니다. dense mode는 v1에서만 허용되며 동일한 유한 limit을 적용합니다. decoder는 header, dimension, offset, tag, UTF-8, trailing byte, stale SST를 먼저 검증합니다. `setSheetDataBuffer`는 이동된 coordinate와 전체 formula budget을 검증한 뒤 한 번에 반영합니다.

`fromJSON`은 실제 array index와 own data property만 읽습니다. `__proto__` 같은 이름도 prototype을 바꾸지 않는 일반 own property로 기록합니다. custom header, duplicate key, accessor, iterator override, FormulaValue와 DateValue getter를 검증하고 snapshot한 뒤 native mutation에 넘깁니다. row, column, cell, retained byte, Excel coordinate limit을 넘으면 sheet를 변경하지 않습니다.

```typescript
workbook.fromJSON(
  'Sheet1',
  [{ Name: 'Ada', Score: 98 }],
  {
    header: ['Name', 'Score'],
    startCell: 'B2',
    limits: {
      maxRows: 10_000,
      maxColumns: 100,
      maxCells: 1_000_000,
      maxRetainedBytes: 128 * 1024 * 1024,
    },
  },
);
```

CSV는 formula-capable text 앞에 apostrophe를 붙이는 `spreadsheet-safe`가 기본값입니다. 외부 system과 원문을 교환해야 하고 입력을 신뢰할 수 있을 때만 `raw-exchange`를 지정하세요. 두 policy 모두 traversal과 64 MiB 반환 제한을 적용합니다.

```typescript
const safeCsv = workbook.toCSV('Sheet1');
const rawCsv = workbook.toCSV('Sheet1', {
  formulaPolicy: 'raw-exchange',
  limits: { maxOutputBytes: 8 * 1024 * 1024 },
});
```

HTML은 text와 attribute를 context에 맞게 escape하며 sparse traversal을 사용합니다. `addPivotTable`처럼 본질적으로 dense range가 필요한 API는 작업을 시작하기 전에 source rectangle을 유한한 traversal limit과 비교합니다.

## Streaming read와 write

`openSheetReader`는 package-backed data나 bounded disk spool에서 batch를 읽습니다. borrowed Rust reader와 owned/Node reader 모두 batch size를 reserve 전에 검증하며, 지나치게 큰 batch는 reader state를 소비하지 않고 거부합니다.

```typescript
const reader = await workbook.openSheetReader('Sheet1', { batchSize: 500 });
for await (const batch of reader) {
  for (const row of batch) {
    processRow(row);
  }
}
```

`StreamWriter`는 row count, cell count, 한 row의 XML byte, 전체 row XML byte, column/merge/pane configuration count와 byte를 제한합니다. batch는 모든 row를 검증한 뒤 처음 byte를 게시하므로 실패한 batch의 앞부분도 남지 않습니다. I/O가 일부 진행된 뒤 실패한 writer는 poisoned 상태가 되며 다시 성공한 것처럼 사용할 수 없습니다. streamed formula도 일반 formula mutation과 같은 누적 budget을 사용합니다.

async `save()`와 `writeBuffer()`는 background task에서 실행됩니다. 작업 중에는 native workbook snapshot을 보호하기 위해 같은 `Workbook`의 mutation이 `"workbook has an async save operation in progress"` 오류로 거부됩니다. 저장 Promise를 기다린 뒤 변경하세요.

```typescript
await workbook.save('snapshot.xlsx', {
  outputBytes: 512 * 1024 * 1024,
  xmlPartBytes: 128 * 1024 * 1024,
  rowCount: 10_000_000,
  cellCount: 10_000_000,
});

workbook.setCellValue('Sheet1', 'A1', 'saved');
```

저장이 자원 제한이나 I/O 오류로 실패해도 기존 destination byte는 유지되고 temporary file은 정리됩니다. task가 끝나면 workbook을 다시 변경하거나 저장할 수 있습니다.

## 저장, 암호화, macro, package fidelity

Rust의 `save`, `save_to_buffer`와 Node.js의 sync/async path 및 buffer 저장은 유한한 `SaveLimits`를 사용합니다. Rust에서 다른 profile이 필요하면 `save_with_limits`와 `save_to_buffer_with_limits`를 사용하세요.

```rust
use sheetkit::{SaveLimits, Workbook};

let workbook = Workbook::new();
let limits = SaveLimits {
    output_bytes: 512 * 1024 * 1024,
    xml_part_bytes: 128 * 1024 * 1024,
    row_count: 10_000_000,
    cell_count: 10_000_000,
};
workbook.save_with_limits("report.xlsx", limits)?;
```

1.0.0 후보는 다시 serialize하지 않고 복사한 clean worksheet에 `xml_part_bytes`를 적용하지 않았으며, 원자적 교체 과정에서 symlink 대상을 일반 파일로 바꾸고 기존 permission bit를 유지하지 않았습니다. 두 문제 모두 수정되었습니다. 자세한 내용은 [Audit v2 동작 변경](#audit-v2-behavior-changes)을 참고하세요.

Standard와 Agile encryption parsing, password derivation, package HMAC는 Office-compatible 동작으로 정리되었습니다. `CryptLimits`는 outer CFB, named stream, encrypted package, KDF iteration과 work를 제한하며 `VbaLimits`는 compressed, decoded, retained, output byte와 module 수를 한 operation에서 누적합니다. 기존에 만든 비표준 암호화 fixture는 Office나 msoffcrypto로 다시 검증하세요.

저장 대상 확장자가 최종 형식을 결정합니다. `.xlsx`와 `.xltx`는 VBA를 제거하며 `.xlsm`, `.xltm`, `.xlam`은 canonical VBA package를 보존합니다.

```typescript
await workbook.save('macro-free.xlsx');
await workbook.save('macro-preserved.xlsm');
```

save cycle은 untouched OOXML과 lazy auxiliary byte를 가능한 한 그대로 보존하지만 SST metadata, relationship, content type은 canonical 형태로 쓸 수 있습니다. ZIP archive의 byte-for-byte 동일성을 테스트하지 말고 Excel/LibreOffice에서의 의미와 package relationship을 검증하세요.

## OOXML feature와 package graph

1.0은 sheet copy/delete, table, pivot, slicer, form control, chart, drawing, image, shape, sparkline의 package graph를 mutation 전에 검사합니다. 실패하면 일부 relationship이나 part만 남기지 않습니다.

- Sheet copy/delete는 drawing, chart, image, table, comment, defined name과 package path의 전체 closure를 함께 처리합니다.
- Table은 width와 header를 먼저 검증하고 conditional formula, priority, DXF를 완전하게 기록합니다.
- Pivot은 cache, table, record, workbook, worksheet relationship과 ID를 충돌 없이 생성합니다.
- Slicer는 cache와 view extension graph, placement, stale part 제거를 함께 처리합니다.
- Form control은 VML과 property relationship을 보존하고 label과 placement를 검증합니다.
- Chart와 drawing은 ID를 중복 없이 생성하며 stock chart의 series 순서와 수, axis를 검증합니다.
- Sparkline은 lazy state를 먼저 materialize한 뒤 copy와 structural shift를 처리합니다.
- Validation, page break, comment, workbook protection hash와 VBA token은 Office semantics에 맞게 수정되었습니다.
- Orphan table part나 모호한 `tableParts` ownership은 첫 sheet에 임의로 붙지 않고 invalid package 오류가 됩니다.

이 변경으로 생성된 OOXML과 기존 golden file이 달라질 수 있습니다. 이전 결과를 복사하는 대신 1.0에서 다시 생성하고 save/reopen 및 Excel/LibreOffice 확인을 수행하세요.

## Style, rendering, SVG

theme와 inherited style, HSL tint, Excel number format의 rendering이 수정되었습니다. 표시 text와 SVG snapshot은 다시 기준을 잡아야 합니다.

SVG는 XML text, XML attribute, CSS string을 별도로 escape합니다. font name을 미리 escape하지 말고 원래 논리 이름을 넘기세요. color는 선택적인 `#` 뒤에 6자리 RGB 또는 8자리 ARGB ASCII hex만 허용합니다. 잘못된 color는 panic이나 raw attribute가 아니라 sanitized 오류가 됩니다. 큰 render rectangle은 layout allocation과 output을 시작하기 전에 traversal limit으로 검사합니다.

## CLI 자동화

`info`, `sheets`, `read`, `get`의 human output은 C0, C1, DEL control 문자를 보이는 형태로 escape합니다. 원래 값을 기계적으로 처리하려면 JSON을 사용하세요.

```bash
sheetkit info report.xlsx --format json
sheetkit sheets report.xlsx --format json
sheetkit read report.xlsx --sheet Sheet1 --format json
sheetkit get report.xlsx Sheet1 A1 --format json
```

CSV read와 convert는 `spreadsheet-safe`가 기본값입니다. 신뢰할 수 있는 data exchange에만 explicit policy를 사용하세요.

```bash
sheetkit read report.xlsx --format csv \
  --csv-formula-policy raw-exchange

sheetkit convert report.xlsx --format csv --output report.csv \
  --csv-formula-policy raw-exchange
```

CLI의 table, JSON, CSV 출력은 byte 제한을 넘으면 성공한 prefix를 stdout이나 destination에 남기지 않습니다. `convert`가 실패하면 기존 destination을 보존합니다.

## 0.5.1 이후 누적 변경 매핑

다음 표는 0.5.1 이후 1.0에 포함된 23개 공개 또는 contributor 변경을 빠짐없이 연결합니다.

| 변경 영역 | 1.0에서 확인할 내용 |
|---|---|
| AI contribution policy | Runtime 변경은 없으며 contributor는 `AI_POLICY.md`와 commit trailer 규칙을 따릅니다. |
| Unsafe workbook input 정리 | formatting, formula, raw transfer, stream, sheet edit가 panic이나 partial mutation 대신 오류를 반환합니다. |
| Reference-aware structural edit | row/column 변경이 formula와 모든 reference-bearing part를 token 단위로 원자적으로 갱신합니다. |
| Sheet lifecycle graph | copy/delete가 sheet-owned relationship closure 전체를 처리합니다. |
| Save pipeline fidelity | canonical SST metadata를 쓰며 untouched lazy byte와 rich text/whitespace를 보존합니다. |
| Advanced OOXML round trip | unknown OOXML과 검토된 style/property field를 보존하며 graph mutation을 preflight합니다. |
| Formula recalculation error | cell domain 오류는 cached error가 되고 engine/cycle 오류만 operation을 중단합니다. |
| Excel formula result | range, blank, financial, date/text, Unicode, Bessel 결과를 새 기준으로 확인해야 합니다. |
| Buffer와 streaming | v1/v2 transfer를 엄격하게 검증하고 absolute coordinate, formula, rich text를 보존합니다. |
| Node value/option parity | DateValue, FormulaValue, open default가 Rust와 같은 의미를 가집니다. |
| Office encryption | Standard/Agile parsing과 HMAC가 Office-compatible 동작을 사용합니다. |
| Office metadata semantics | VBA, protection, validation, page break, comment의 해석과 검증이 수정되었습니다. |
| Table과 conditional format | width preflight, header, formula, priority, DXF를 완전하게 기록합니다. |
| Formatting과 theme rendering | theme tint, inherited style, number format 표시 결과가 수정되었습니다. |
| Lazy workbook feature | picture, sparkline, comment, control, name, merge를 lazy mode에서도 보존합니다. |
| Pivot package | cache, table, record와 workbook/worksheet graph를 완전하게 생성합니다. |
| Slicer package graph | cache/view extension, placement, stale part 정리를 검증합니다. |
| Form control package graph | VML/property relationship, label escape, placement를 검증합니다. |
| Chart와 drawing | unique ID, stock chart series/axis, escaped range를 생성합니다. |
| Agile integrity key | hash-sized key를 생성하며 공개된 salt-sized variant도 읽습니다. |
| Chart/control documentation | stock series 순서와 control metadata 요구 사항이 명확해졌습니다. |
| Lazy sparkline state | copy/shift 전에 sparkline을 hydrate하며 무관한 table을 dirty로 만들지 않습니다. |
| Audit review parity | lazy graph, Node value/default, raw transfer, package path, Agile variant가 Rust/Node에서 일치합니다. |

보안 hardening 77건은 이 문서의 영역에 모두 반영되었습니다. Formula 관련 항목은 `SK-DS-001`부터 `013`, `035`부터 `041`, `044`, `051`, `053`부터 `061`까지이며, open/encryption/VBA 항목은 `014`부터 `017`, `045`, `046`, `048`, `049`, `064`, `065`, `068`, `069`, `077`입니다. Traversal/transfer 항목은 `023`부터 `034`, `043`, `062`, `063`, `066`, `067`, `070`이며, SVG/output/CLI 항목은 `018`부터 `022`, `042`, `047`, `050`, `052`, `071`부터 `074`입니다. Package와 macro 항목은 `076`, `078`입니다. Reportable finding이 아니었던 `SK-DS-075`는 77건에 포함되지 않습니다.

## 성능 측정 상태

최종 clean 환경 비교는 repository의 기존 focused benchmark, 같은 fixture, warm-up 1회, 측정 5회와 median latency를 사용했습니다. 1.0 후보를 v0.5.1과 비교했으며 install, build, linking, fixture 생성 시간은 모두 제외했습니다.

100,000-row fixture에서 lazy path open은 28.6% 개선되었고 일반 `getRows`는 +0.2%, `getRowsRaw-only`는 +1.4%, streaming은 +3.3%, 변경하지 않은 copy-on-write save는 -44.6%, 변경한 save는 -1.0%였습니다. V2-to-V2 buffer 생성은 +7.1%, 절대값 1.772 ms였습니다. 이 경로는 canonical coordinate, shared formula, date style과 유한한 output을 추가로 검증하며 10% release-stop 기준 아래에 있습니다.

이 값은 Apple M4 Pro와 Node.js 24.18.0 환경에서 수집한 release evidence입니다. 모든 workbook과 기기에서 같은 성능을 보장하지는 않습니다.

## 이전 확인 checklist

- Node.js 24와 생성된 TypeScript declaration으로 application을 typecheck하세요.
- open call에서 lazy/deferred 기본값과 date interpretation이 의도와 맞는지 확인하세요.
- `ResourceLimitError`를 다른 argument, lifecycle, password 오류와 구분하여 처리하세요.
- sparse row/column과 absolute raw coordinate를 기준으로 consumer를 수정하세요.
- DateValue와 FormulaValue의 object shape를 보존하도록 serializer와 type narrowing을 수정하세요.
- formula, JSON import, traversal, save, stream에 필요한 유한 limit만 명시하세요.
- CSV automation은 `spreadsheet-safe`와 `raw-exchange` 중 하나를 의도적으로 선택하세요.
- async save와 writeBuffer가 끝난 뒤 workbook을 변경하세요.
- macro를 보존할 파일은 macro-capable 확장자로 저장하세요.
- sheet copy/delete와 table, pivot, slicer, control, chart golden file을 새로 생성하세요.
- byte-for-byte ZIP 비교 대신 save/reopen과 Excel/LibreOffice 의미 검증을 수행하세요.
- CLI automation은 human text 대신 JSON을 사용하세요.
- 실제 workload가 release fixture와 크게 다르면 같은 방식으로 application fixture를 다시 측정하세요.

<a id="audit-v2-behavior-changes"></a>

## Audit v2 동작 변경

1.0.0 후보에 대한 두 번째 audit에서 아래 수정이 적용되었습니다. 각 수정은 application이 의존할 수 있는 동작을 바꿉니다. Finding ID는 [보안 finding ledger](../../security/v1-finding-ledger.md#audit-v2-post-1-0-0-candidate)에 정리되어 있습니다.

### 열기와 저장

- Path 저장은 symlink를 따라가 대상 파일을 교체하므로 link 자체는 그대로 남습니다. Unix에서는 기존 대상 파일의 permission bit를 유지하며, 새 파일은 `0o666`에서 process umask를 적용한 mode로 생성됩니다.
- 어떤 part든 CRC 또는 압축 해제 검사에 실패하면 모든 read mode에서 open이 실패합니다. 이전 후보는 이런 part를 조용히 제외했습니다. 내용은 손상되지 않았지만 XML 형식이 잘못된 optional part는 계속 보존됩니다.
- Worksheet relationship이 다른 part를 가리키는 alias이거나, package 밖을 가리키거나, external이면 open이 실패합니다.
- ZIP entry 이름이 중복되거나 대소문자만 다른 part 이름이 있으면 open이 실패합니다.
- 선언된 크기와 실제 내용이 다른 entry는 거부됩니다. Eager open은 즉시 실패하며, lazy workbook은 해당 part를 hydrate하거나 저장 중 복사할 때 실패합니다.
- 저장 시 `xml_part_bytes`(`xmlPartBytes`)는 모든 worksheet에 대해 실제로 쓰는 byte 수를 기준으로 적용됩니다. Lazy 저장과 sheet filtering이 다시 serialize하지 않고 복사하는 clean worksheet도 포함됩니다. 낮은 limit으로 통과하던 저장이 이제 실패할 수 있습니다.
- Lazy workbook은 open 시점의 package를 계속 읽습니다. 이후 원본 path를 교체하거나 다시 써도 읽기 결과는 바뀌지 않습니다.

### 구조 편집

- `removeRow`와 `removeCol`은 삭제된 cell을 가리키는 reference를 `#REF!`로 바꾸고, 행이나 열을 잃은 range를 축소합니다. Defined name도 같은 규칙을 따릅니다. Merge, data validation, conditional format, hyperlink, auto-filter, comment, threaded comment, 행 page break도 축소되거나 제거됩니다.
- Table 안의 열을 삭제하면 해당 table column도 제거됩니다. Table의 header 행, totals 행, 모든 data 행, 모든 열을 제거하는 편집은 workbook을 변경하지 않고 거부됩니다.
- Comment가 grid 밖으로 밀려나는 경우 `insertRows`와 `insertCols`가 실패합니다.
- `duplicateRow`는 복사본의 shared formula를 일반 formula로 펼칩니다.
- `A:A`, `1:1` 같은 전체 열/행 reference도 이동합니다. 전체 sheet span(`1:1048576`, `A:XFD`)과 3-D reference는 이동하지 않습니다.
- `copySheet`는 기존 table 이름 및 defined name과 대소문자 구분 없이 겹치지 않는 table 이름을 생성합니다.

```typescript
workbook.setCellFormula('Sheet1', 'B1', 'A5+SUM(A1:A5)');
workbook.removeRow('Sheet1', 5);
// B1의 값은 '#REF!+SUM(A1:A4)'가 됩니다.
```

### 수식 결과

- `SUM`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`, `MEDIAN`, `LARGE`, `SMALL`, `RANK` 같은 numeric aggregate는 reference 안의 boolean, text, 빈 cell을 무시합니다. 숫자가 없는 reference에 대한 `AVERAGE`는 `#DIV/0!`를 반환합니다. 수식에 직접 입력한 scalar 인수는 기존 coercion을 유지하므로 `SUM("5", TRUE)`는 여전히 6입니다.
- 연도가 0..=9999 범위를 벗어난 `DATE`는 `#NUM!`을 반환합니다. `DATE`, `EDATE`, `EOMONTH`, `WORKDAY`는 결과가 Excel 날짜 범위를 벗어나면 `#NUM!`을 반환하며, `WORKDAY`는 serial 1 이전의 결과에도 `#NUM!`을 반환합니다.
- `LEFT`, `RIGHT`, `MID`, `REPT`, `REPLACE`는 음수 count에 `#VALUE!`를 반환합니다. 매우 큰 count는 overflow 대신 clamp됩니다.
- `ROUND`, `ROUNDUP`, `ROUNDDOWN`은 NaN이나 infinity를 반환하지 않습니다. `f64` 정밀도를 넘는 digit 수는 값을 그대로 두거나 0으로 만들며, non-finite 결과가 될 경우 `#NUM!`을 반환합니다.
- Cross-sheet reference로 도달한 formula는 unqualified reference, `ROW()`, `COLUMN()`을 해당 formula의 sheet와 cell 기준으로 계산합니다.

### ID와 저장 값

- 새 sheet relationship ID, sheet ID, hyperlink relationship ID, custom property `pid`, pivot `cacheId`, table ID, custom number format ID는 사용되지 않은 가장 작은 값을 사용합니다. `rId2`를 삭제한 뒤 추가하는 relationship은 `rId4`가 아니라 `rId2`가 됩니다. Conditional format priority는 평가 순서가 바뀌지 않도록 계속 현재 최댓값에 1을 더한 값을 사용합니다.
- 같은 text를 가진 rich text 문자열과 일반 문자열은 더 이상 shared string entry를 공유하지 않습니다.
- Custom document property는 빈 문자열을 유지하고 XML entity를 decode합니다. Slicer cache attribute도 entity를 decode합니다.

### Node.js 인수와 출력

- 정수 인수와 numeric option은 소수, NaN, infinity, safe integer를 넘는 값, 범위를 벗어난 값을 잘라내거나 wrap하지 않고 거부합니다.
- 알 수 없는 style 및 conditional format enum 문자열은 허용되는 값 목록과 함께 오류를 발생시킵니다.
- `toHTML`은 최종 UTF-8 HTML 크기를 `limits.maxOutputBytes`에 반영합니다.
- `StreamWriter.addMergeCell`은 range를 canonical 형태로 저장하며 겹치는 range에는 오류를 발생시킵니다.
- Stream reader는 잘린 worksheet에서 일부 행을 반환하지 않고 오류를 발생시킵니다.
- Table 이름은 Excel 규칙을 따릅니다. 길이는 1~255자이며 cell reference처럼 보이는 이름은 사용할 수 없고, table과 defined name 전체에서 대소문자 구분 없이 고유해야 합니다. 자세한 내용은 [API parity matrix](../../security/v1-api-parity-matrix.md)를 참고하세요.

### 암호화와 VBA

- Agile decryption은 password key와 package key의 크기가 다른 파일도 읽습니다. 잘못된 numeric attribute와 1..=65536 범위를 벗어난 `saltSize`는 거부합니다.
- VBA module 이름은 Unicode stream-name record를 우선 사용합니다. CP932, CP949, CP936 source code는 lossy하게 decode됩니다. ASCII와 CP932의 반각 가타카나는 유지되며 그 밖의 double-byte 문자는 U+FFFD로 바뀝니다. 영향을 받은 module마다 project의 `warnings` 목록에 warning이 추가됩니다.
