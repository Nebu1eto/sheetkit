# 0.5.1에서 1.0으로 마이그레이션

SheetKit 1.0은 모든 작업에 유한한 제한을 적용하고, 0.5.1에서 허용하던 잘못된 입력을 거부하며, Excel과 다르던 계산 결과를 바로잡았습니다. 대부분의 애플리케이션은 코드를 고치지 않고 업그레이드할 수 있지만, 일부 기본값, 반환 형태, 오류, 저장 결과는 의도적으로 바뀌었습니다.

이 문서는 Rust crate, `@sheetkit/node` 패키지, CLI를 다룹니다. 업그레이드한 뒤에는 native addon을 다시 빌드하고 새 declaration으로 TypeScript를 다시 컴파일하세요. Node.js 패키지는 Node.js 24 이상이 필요합니다.

## 먼저 확인할 사항

- **자원 제한.** 열기, 수식, 순회, 스트림, 저장, 암호화, VBA 작업에 유한한 기본 제한이 적용됩니다. 신뢰할 수 있는 입력에 한해 제한을 높일 수 있지만 제한을 없앨 수는 없습니다.
- **Lazy 열기.** Node.js는 워크북을 lazy 방식으로 열고 보조 part를 나중에 읽습니다. 날짜 서식이 지정된 숫자 셀은 날짜로 읽습니다.
- **Sparse 행.** 행과 열 API는 값이 있는 셀만 반환합니다. Dense 출력은 제한을 명시해서 선택하는 옵션입니다.
- **구조화된 값.** Node.js는 수식 셀과 날짜 셀을 단순 값이 아닌 객체로 반환합니다.
- **원자적 실패.** 수식, 가져오기, 구조 편집, 저장 작업은 끝까지 완료되거나 워크북을 바꾸지 않습니다.
- **CSV.** 수식처럼 보이는 텍스트는 기본적으로 무력화해서 출력합니다.
- **비동기 저장.** `save`, `writeBuffer`, `saveWithPassword`는 백그라운드에서 실행되며, 완료될 때까지 워크북을 수정할 수 없습니다.
- **매크로.** `.xlsx`나 `.xltx`로 저장하면 VBA가 제거됩니다.
- **엄격한 검증.** 잘못된 참조, 버퍼, 표, 색상, package part, Node.js 인수는 일부만 반영되는 대신 오류가 됩니다.

## 자원 제한과 오류

제한은 작업마다 따로 적용됩니다. 전역 설정이 아니며 워크북에 저장되지 않고, 고정된 상한을 넘을 수 없습니다.

| 작업 | Rust | Node.js | 주요 기본값 |
| --- | --- | --- | --- |
| 열기 | `OpenOptions`, `OpenLimits`, `CryptLimits`, `VbaLimits` | `OpenOptions` | 입력 128 MiB, ZIP entry 4,096개, 압축 해제 합계 512 MiB, XML part당 256 MiB |
| 수식 | `FormulaLimits` | `FormulaLimits` | 수식당 32 KiB, 입력 합계 4 MiB, 출력 합계 16 MiB |
| 순회 | `TraversalOptions`, `TraversalLimits` | `TraversalOptions` | 행 1,048,576개, 열 16,384개, 셀 1천만 개, 출력 512 MiB |
| 저장 | `SaveLimits` | `SaveLimits` | archive 512 MiB, XML part당 128 MiB, 행과 셀 각 1천만 개 |
| 스트림 쓰기 | `StreamLimits` | `StreamLimits` | 행당 16 MiB, 합계 512 MiB, 행 1,048,576개, 셀 1천만 개 |
| JSON 가져오기 | 없음 | `limits` 옵션 | 행 1,048,576개, 열 16,384개, 셀 1천만 개, 보관 데이터 512 MiB |
| CSV | CSV 옵션 | `ToCsvOptions` | 출력 64 MiB |

한 작업은 여러 항목을 함께 계산하므로, 각 값이 기본값보다 작아도 실패할 수 있습니다.

### 제한 오류 처리

Rust에서는 제한 초과가 `Error::ResourceLimitExceeded`로 반환됩니다. 잘못된 데이터나 I/O 오류와 구분해서 처리하세요.

```rust
use sheetkit::{Error, OpenOptions, Workbook};

match Workbook::open_with_options("report.xlsx", &OpenOptions::default()) {
    Ok(workbook) => println!("{} sheets", workbook.sheet_names().len()),
    Err(Error::ResourceLimitExceeded(details)) => eprintln!(
        "{} exceeded {}: {} > {}",
        details.operation_code(),
        details.dimension_code(),
        details.observed(),
        details.limit(),
    ),
    Err(error) => return Err(error),
}
```

Node.js에서는 오류 메시지를 비교하지 말고 `isResourceLimitError`를 사용하세요. `limit`과 `observed`는 `bigint`입니다.

```typescript
import { isResourceLimitError, Workbook } from '@sheetkit/node';

try {
  const workbook = await Workbook.open('report.xlsx');
} catch (error) {
  if (!isResourceLimitError(error)) throw error;
  console.error(error.operation, error.dimension, error.limit, error.observed);
}
```

인수 오류, 수명 주기 오류, 잘못된 비밀번호는 기존처럼 일반 오류입니다.

### 제한 높이기

신뢰할 수 있는 파일에 필요한 항목만 높이세요.

```typescript
const workbook = await Workbook.open('trusted-large.xlsx', {
  maxOuterInputBytes: 256 * 1024 * 1024,
  maxUnzipSize: 768 * 1024 * 1024,
  maxZipEntries: 8_000,
});
```

`maxUnzipSize`와 `maxZipEntries`도 계속 사용할 수 있지만, 어떤 값으로도 archive 검사를 끌 수 없습니다. 0, 안전하지 않은 JavaScript 숫자, 상한을 넘는 값, 알 수 없는 모드는 거부됩니다.

`maxXmlPartBytes`로 part당 기본값 256 MiB를 낮출 수 있습니다. 각 ZIP entry도 256 MiB로 제한되고 Node.js에서는 이 제한을 바꿀 수 없으므로, 256 MiB보다 높게 지정해도 효과가 없습니다.

100:1 압축률 검사는 압축을 풀었을 때 8 MiB를 넘는 entry에만 적용되므로, shared strings처럼 작고 반복이 많은 part는 그대로 열립니다. 기준은 Node.js의 `compressionRatioFloorBytes`나 Rust의 `OpenOptions::compression_ratio_floor_bytes`로 바꿀 수 있으며, `0`을 지정하면 모든 entry를 검사합니다.

## 워크북 열기

### 기본값

Node.js의 열기 메서드는 다음 기본값을 사용합니다.

```typescript
{ readMode: 'lazy', auxParts: 'deferred', dateInterpretation: 'numFmt' }
```

Rust에서는 `ReadMode::Lazy`, `AuxParts::Deferred`, `DateInterpretation::NumFmt`에 해당합니다.

Lazy 워크북은 워크시트나 보조 part를 처음 접근할 때 열기 시점의 제한으로 파싱합니다. 그래서 `open`이 성공했더라도 `maxXmlPartBytes`를 넘는 part는 처음 접근할 때 실패할 수 있습니다. 파일 전체를 미리 검증하려면 eager로 여세요.

```typescript
const workbook = await Workbook.open('input.xlsx', {
  readMode: 'eager',
  auxParts: 'eager',
});
```

Lazy 워크북은 열었을 때의 파일 내용을 계속 읽습니다. 연 뒤에 파일을 교체해도 워크북이 읽는 내용은 바뀌지 않습니다.

### 필터링과 부분 읽기

`sheets`로 제외한 시트는 압축을 풀거나 검증하지 않고, 저장할 때 그대로 복사합니다. `sheetRows`는 지정한 행까지만 파싱합니다. 두 옵션 모두 나머지 내용을 검증하지 않습니다.

### 이전 옵션 이름

`parseMode: 'full'`은 eager 열기로, `parseMode: 'readfast'`는 lazy 열기로 처리됩니다. 새 코드에서는 `readMode`를 사용하세요. `readMode: 'stream'`은 현재 `'lazy'`와 같게 동작하며, 순방향 스트리밍에는 `openSheetReader`를 사용하세요.

### 이제 열리지 않는 파일

다음 경우에는 읽기 모드와 관계없이 열기가 실패합니다.

- part가 CRC 검사나 압축 해제에 실패한 경우(0.5.1은 이런 part를 버렸습니다)
- ZIP entry 이름이 중복되거나, 대소문자만 다른 part 이름이 있는 경우
- 워크시트 relationship이 다른 part, package 바깥, 외부 대상을 가리키는 경우
- entry에 선언된 크기가 실제 내용과 다른 경우(lazy 워크북은 해당 part를 읽거나 복사할 때 실패합니다)

압축은 정상적으로 풀리지만 XML 형식이 잘못된 part는 계속 보존됩니다.

## 셀 값과 날짜

Node.js의 셀 값 타입은 다음과 같습니다.

```typescript
type CellValue = string | number | boolean | DateValue | FormulaValue | null;
```

기본값인 `dateInterpretation: 'numFmt'`에서는 날짜 서식이 지정된 숫자 셀이 `DateValue`가 됩니다. 저장된 셀 타입을 그대로 따르려면 `'cellType'`을 지정하세요. Rust에서는 `CellValue::Date`를 반환합니다.

수식 셀은 수식과 캐시된 결과를 함께 유지합니다.

```typescript
const value = workbook.getCellValue('Sheet1', 'D2');
if (typeof value === 'object' && value?.type === 'formula') {
  console.log(value.formula, value.result);
}
```

날짜는 직접 읽기, 행 순회, raw 버퍼, 스트림 읽기, JSON 가져오기에서 모두 같은 형태를 유지합니다. 수식은 `getCellValue`에서 `FormulaValue`로 반환되고, `getRows`, `getRowsIterator`, `getCols`, `openSheetReader`의 셀에서는 `formula` 필드로 제공되며, 쓰기와 JSON 가져오기도 같은 형태를 받습니다. raw 버퍼(`getRowsBuffer`, `getRowsBufferV2`, `getRowsRaw`), `toJSON`, `toCSV`에는 수식 식만 담깁니다. `number`나 평탄화된 수식 결과를 기대하던 타입 가드와 직렬화 코드를 수정하세요.

Rust의 `CompactCellRef::new`와 `CompactCellRef::from_coordinates`는 이제 `Result`를 반환하며 Excel의 행과 열 범위를 검사합니다.

## 수식

### 제한과 원자적 갱신

파싱, 계산, `calculateAll`, `setCellFormula`, `fillFormula`, `FormulaValue` 대입, 스트림 쓰기, JSON 가져오기, raw 버퍼 가져오기는 작업마다 하나의 `FormulaLimits` 예산을 공유합니다. 셀을 바꾸기 전에 예산을 검사하므로, 실패한 일괄 작업은 일부 결과도 남기지 않습니다.

```typescript
workbook.fillFormula('Sheet1', 'D2:D1000', 'B2*C2', {
  operationInputBytes: 4 * 1024 * 1024,
  operationEvaluationSteps: 1_048_576,
});
```

Rust에서는 `set_cell_formula_with_limits`, `fill_formula_with_limits`, `evaluate_formula_with_limits`, `calculate_all_with_limits`를 사용합니다. 제한을 받지 않는 메서드는 `FormulaLimits::secure_default()`를 사용합니다.

`#DIV/0!`처럼 한 셀 안에서 생긴 오류는 그 셀의 캐시된 결과로 저장됩니다. 순환 참조와 엔진 오류만 `calculateAll`을 중단합니다.

### 바뀐 계산 결과

수식 테스트의 기대값을 갱신하세요. 다음 결과가 Excel과 같아졌습니다.

- 범위, 빈 셀, 거듭제곱
- 재무, 휴일, `TEXT`, Unicode `LEN`, Bessel 함수
- `SUM`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`, `MEDIAN`, `LARGE`, `SMALL`, `RANK`는 참조 안의 boolean, 텍스트, 빈 셀을 무시합니다. 숫자가 없는 참조에 대한 `AVERAGE`는 `#DIV/0!`를 반환합니다. 수식에 직접 입력한 값은 계속 변환되므로 `SUM("5", TRUE)`는 6입니다.
- `DATE`, `EDATE`, `EOMONTH`, `WORKDAY`는 Excel 날짜 범위를 벗어나면 `#NUM!`을 반환합니다.
- `LEFT`, `RIGHT`, `MID`, `REPT`, `REPLACE`는 음수 개수에 `#VALUE!`를 반환하고, 매우 큰 개수는 범위 안으로 제한합니다.
- `ROUND`, `ROUNDUP`, `ROUNDDOWN`은 NaN이나 무한대를 반환하지 않습니다.
- 다른 시트를 거쳐 계산되는 수식은 시트 이름이 없는 참조, `ROW()`, `COLUMN()`을 자기 시트와 셀 기준으로 계산합니다.

0.5.1에서 허용하던 일부 입력은 이제 거부되거나 제한됩니다. `COUNTIF(S)`, `SUMIF(S)`, `AVERAGEIF(S)`, `SEARCH`의 와일드카드 조건, 감가상각과 누적 이자 함수의 큰 기간 값, 결과가 지나치게 커지는 `SUBSTITUTE`, 진법 변환 함수에서 1~10을 벗어난 `places`가 해당합니다.

## 구조 편집

행과 열의 삽입, 삭제, 복사는 이제 텍스트 치환이 아니라 수식 token 단위로 참조를 이동합니다. 수식, 정의된 이름, 데이터 유효성 검사, 조건부 서식, 차트, 표 등 관련된 부분이 함께 갱신되며, 그렇지 못하면 아무것도 바꾸지 않고 실패합니다. 수식 텍스트를 직접 고치던 우회 코드는 제거하세요.

예상되는 동작은 다음과 같습니다.

- `removeRow`와 `removeCol`은 삭제된 셀을 가리키는 참조를 `#REF!`로 바꾸고, 행이나 열을 잃은 범위를 줄입니다. 병합, 유효성 검사, 조건부 서식, 하이퍼링크, 자동 필터, 메모, 행 페이지 나누기도 줄어들거나 제거됩니다.
- 표 안의 열을 삭제하면 해당 표 열도 제거됩니다. 표의 머리글 행, 요약 행, 모든 데이터 행, 모든 열을 없애는 편집은 거부됩니다.
- 메모가 시트 밖으로 밀려나면 `insertRows`와 `insertCols`가 실패합니다.
- `duplicateRow`는 복사본의 공유 수식을 일반 수식으로 기록합니다.
- `1:1`, `A:A` 같은 전체 행과 전체 열 참조도 이동합니다. 시트 전체 참조(`1:1048576`, `A:XFD`)와 3차원 참조는 이동하지 않습니다.

```typescript
workbook.setCellFormula('Sheet1', 'B1', 'A5+SUM(A1:A5)');
workbook.removeRow('Sheet1', 5);
// B1은 '#REF!+SUM(A1:A4)'가 됩니다.
```

시트 복사와 삭제는 이제 시트에 속한 그림, 차트, 이미지, 표, 메모, 정의된 이름, 컨트롤을 모두 함께 처리합니다. `copySheet`는 복사한 표에 기존 표 이름이나 정의된 이름과 대소문자 구분 없이 겹치지 않는 이름을 붙입니다.

### 생성되는 ID

새 relationship ID, 시트 ID, 사용자 지정 속성 ID, 피벗 캐시 ID, 표 ID, 사용자 지정 숫자 서식 ID는 사용하지 않는 가장 작은 값을 받습니다. `rId2`를 삭제한 뒤 추가하는 relationship은 `rId4`가 아니라 `rId2`가 됩니다. 조건부 서식 우선순위는 계산 순서가 바뀌지 않도록 계속 현재 최댓값에 1을 더한 값을 사용합니다.

## 행, 열, raw 버퍼, JSON, CSV, HTML

### Sparse 순회

0.5.1에서는 멀리 떨어진 셀 하나 때문에 행 API가 그 셀까지의 사각형 전체를 펼칠 수 있었습니다. 1.0의 `getRows`, `getCols`, raw 버퍼, `toJSON`, `toCSV`, `toHTML`은 순회 제한 안에서 값이 있는 데이터만 반환합니다.

```typescript
const rows = workbook.getRows('Sheet1', {
  mode: 'sparse',
  limits: { maxCells: 2_000_000, maxOutputBytes: 128 * 1024 * 1024 },
});
```

각 행은 절대 행 번호와 열 번호를 가집니다. 배열 위치로 빈 행을 추정하지 마세요. Sparse CSV와 HTML에서는 앞쪽과 중간의 빈 열도 빠집니다.

사각형 전체가 필요한 코드라면 제한을 명시해서 dense v1 출력을 요청하고 `SheetData`로 해석하세요.

```typescript
import { SheetData } from '@sheetkit/node';

const buffer = workbook.getRowsBuffer('Sheet1', {
  mode: 'dense',
  limits: { maxCells: 250_000, maxOutputBytes: 64 * 1024 * 1024 },
});
const dense = new SheetData(buffer).toArray();
```

Dense 모드는 v1 버퍼 형식에만 있으며, `getCols`는 dense 모드를 거부합니다.

### Raw 버퍼

v1 형식은 이제 기본으로 sparse 출력을 만듭니다. v2 형식은 항상 sparse이며 절대 좌표와 인라인 문자열을 사용합니다. 두 디코더 모두 데이터를 반환하기 전에 버퍼 전체를 검증하고, `setSheetDataBuffer`는 셀을 쓰기 전에 모든 내용을 검증합니다. 프로세스 밖에서 받은 버퍼는 신뢰할 수 없는 입력으로 다루세요. 거부된 버퍼는 워크북을 바꾸지 않으며, 워크북은 계속 사용할 수 있습니다.

### JSON

`toJSON`은 자체 데이터 속성만 가진 일반 객체를 반환합니다. 머리글 이름이 `__proto__`여도 일반 속성이 되며 행 객체의 prototype을 바꾸지 않습니다.

`fromJSON`은 배열 인덱스와 자체 데이터 속성을 한 번만 읽고, 사용자 지정 iterator는 무시합니다. Proxy, getter, 중복된 머리글, 잘못된 좌표, 형식이 잘못된 날짜나 수식 객체는 시트를 바꾸지 않고 거부됩니다.

```typescript
workbook.fromJSON('Sheet1', records, {
  header: true,
  startCell: 'A1',
  limits: { maxRows: 10_000, maxColumns: 100 },
});
```

### CSV와 HTML

CSV 출력은 수식처럼 시작하는 텍스트 앞에 작은따옴표(`'`)를 붙여 스프레드시트가 실행하지 않게 합니다. 받는 쪽이 스프레드시트가 아니고 데이터를 신뢰할 수 있을 때만 `raw-exchange`를 사용하세요.

```typescript
const safeCsv = workbook.toCSV('Sheet1');
const rawCsv = workbook.toCSV('Sheet1', { formulaPolicy: 'raw-exchange' });
```

CSV 출력은 64 MiB로 제한됩니다. HTML은 텍스트와 속성을 위치에 맞게 escape하고, 최종 HTML 크기를 `limits.maxOutputBytes`에 포함합니다.

## 스트리밍

### 읽기

큰 시트를 앞에서부터 읽을 때는 `openSheetReader`를 사용하세요.

```typescript
const reader = await workbook.openSheetReader('Sheet1', { batchSize: 500 });
for await (const batch of reader) {
  for (const row of batch) processRow(row);
}
```

Reader는 워크시트를 복제하지 않고 package나 크기가 제한된 임시 파일에서 읽습니다. 배치 크기는 100,000 이하의 양의 정수여야 합니다. 잘린 워크시트는 일부 행을 반환하지 않고 오류를 발생시킵니다. Rust에서는 `open_sheet_reader`나 `open_sheet_reader_owned`를 사용합니다.

### 쓰기

Node.js의 `newStreamWriter`와 Rust의 `StreamWriter::new`는 유한한 기본값을 사용합니다. 제한은 writer를 만들 때 지정하세요.

```typescript
const writer = workbook.newStreamWriter('Export', {
  rowCount: 500_000,
  outputBytes: 256 * 1024 * 1024,
});
```

Rust에서는 `StreamWriter::new_with_limits`를 사용합니다. 행이나 배치는 쓰기 전에 검사하므로 거부된 행은 추가되지 않습니다. I/O 오류가 난 writer는 이후 호출을 거부합니다. `addMergeCell`은 겹치는 범위를 거부합니다.

## 저장

### 제한과 원자적 교체

모든 저장 메서드는 `SaveLimits`를 사용합니다. Rust에서는 `save_with_limits`나 `save_to_buffer_with_limits`로, Node.js에서는 두 번째 인수로 제한을 지정합니다.

```typescript
await workbook.save('output.xlsx', {
  outputBytes: 768 * 1024 * 1024,
  xmlPartBytes: 256 * 1024 * 1024,
});
```

`xmlPartBytes`는 lazy 저장이 그대로 복사하는 변경 없는 시트를 포함해 기록하는 모든 워크시트에 적용됩니다. 0.5.1에서 낮은 제한으로 통과하던 저장이 이제 실패할 수 있습니다.

경로 저장은 같은 디렉터리에 임시 파일을 쓰고, 쓰기가 끝난 뒤에만 대상 파일을 교체합니다. 저장이 실패하면 기존 파일은 그대로 남습니다. Symlink 경로로 저장하면 link는 그대로 두고 link가 가리키는 파일을 교체합니다. Unix에서는 기존 파일의 권한 비트를 유지하고, 새 파일은 `0o666`에서 프로세스 umask를 뺀 권한으로 만듭니다.

Node.js의 `saveWithPassword`와 `saveWithPasswordSync`는 `SaveLimits` 대신 `CryptLimits`를 받습니다. Rust의 `save_with_password_with_output_limits`는 두 제한을 모두 받습니다.

<a id="opt-in-save-durability"></a>

### Durable 저장

저장은 기본적으로 원자적이지만 디스크에 flush하지는 않습니다. 저장 직후 전원이 끊기거나 OS가 비정상 종료되면 이전 파일이나 새 파일이 남고, 파일 시스템에 따라서는 비어 있거나 불완전한 파일이 남을 수 있습니다. 애플리케이션이 비정상 종료되어도 이전 파일은 유지되지만, 강제 종료된 경우 `.sheetkit-*` 임시 파일이 남을 수 있습니다.

전원이 끊겨도 파일이 남아야 한다면 durability를 요청하세요.

```rust
use sheetkit::{SaveLimits, SaveOptions};

workbook.save_with_options("output.xlsx", SaveLimits::default(), SaveOptions { durable: true })?;
```

```typescript
await workbook.save('output.xlsx', undefined, { durable: true });
```

Durable 저장은 rename 전에 임시 파일을 flush하고, Unix에서는 rename 뒤에 디렉터리를 sync합니다. macOS에서는 `F_FULLFSYNC`를 사용하므로 저장마다 수 밀리초가 더 걸립니다. 디렉터리 sync가 실패하면 새 파일이 이미 제자리에 있다는 오류를 반환합니다. 암호화 저장은 Rust의 `save_with_password_with_options`와 Node.js의 마지막 `options` 인수로 같은 옵션을 지원합니다. 네트워크 파일 시스템이나 flush 요청을 무시하는 드라이브에서는 이 보장이 약해집니다.

### 비동기 저장

`save`, `writeBuffer`, `saveWithPassword`는 이제 백그라운드 작업으로 실행됩니다. 작업이 끝나기 전에 워크북을 수정하면 `workbook has an async save operation in progress` 오류가 발생합니다.

```typescript
await workbook.save('output.xlsx');
workbook.setCellValue('Sheet1', 'A1', 'next');
```

작업이 성공하든 실패하든 끝나면 워크북을 다시 수정하고 저장할 수 있습니다. 서로 다른 워크북은 동시에 저장할 수 있습니다.

### 저장 내용

Lazy 저장은 변경되지 않은 part를 바이트 그대로 복사하고 편집한 part만 다시 씁니다. Shared strings는 의미 있는 공백과 서식 있는 텍스트를 보존하며, 서식 있는 텍스트와 같은 내용의 일반 문자열은 더 이상 항목을 공유하지 않습니다. 사용자 지정 문서 속성은 빈 문자열을 유지합니다. 저장 결과는 ZIP 바이트가 아니라 다시 열어서 비교하세요. 내용이 같아도 ZIP 바이트는 달라질 수 있습니다.

### 매크로

파일 확장자가 VBA 보존 여부를 결정합니다.

| 대상 | 결과 |
| --- | --- |
| `.xlsx`, `.xltx` | VBA part, relationship, content type을 제거합니다 |
| `.xlsm`, `.xltm`, `.xlam` | VBA를 유지합니다 |

매크로 워크북을 `.xlsx`로 저장하면서 VBA가 남아 있기를 기대했다면 `.xlsm`으로 저장하세요.

## 암호화와 VBA

비밀번호로 열 때는 키 유도 전에 `CryptLimits`를 적용하고, 복호화한 package에도 같은 열기 제한을 이어서 적용합니다. Standard와 Agile 암호화, integrity key, package HMAC는 Office와 같은 방식으로 동작합니다. 비밀번호 키와 package 키의 크기가 다른 Agile 파일도 이제 열리며, 1~65536을 벗어난 `saltSize`는 거부됩니다. 0.5.1로 만든 암호화 테스트 파일은 Excel이나 Office 호환 도구로 다시 확인하세요. 큰 암호화 파일의 제한을 높이려면 열기 제한과 암호화 제한을 모두 지정해야 합니다.

VBA 검사는 작업마다 하나의 `VbaLimits` 예산을 사용합니다. 모듈 이름은 Unicode stream-name record가 있으면 그 값을 사용합니다. CP932, CP949, CP936 소스는 손실이 있는 방식으로 디코딩합니다. ASCII와 CP932의 반각 가타카나는 유지되고 나머지 2바이트 문자는 U+FFFD가 되며, 해당 모듈마다 프로젝트의 `warnings` 목록에 경고가 추가됩니다.

## 차트, 표, 피벗과 기타 part

다음 기능을 추가하거나 바꾸면 관련된 모든 part와 relationship을 함께 기록하며, 그렇지 못하면 아무것도 바꾸지 않고 실패합니다.

- 피벗 테이블은 캐시, 레코드, relationship을 겹치지 않는 ID로 기록합니다.
- 슬라이서는 캐시와 view extension을 기록하고 배치를 검증합니다.
- 양식 컨트롤은 VML과 속성 relationship을 유지하고 레이블을 escape합니다.
- 차트, 그림, 이미지, 도형, 스파크라인은 고유한 ID를 사용합니다. 주식형 차트는 문서에 정해진 계열 수와 순서가 필요합니다.
- 표는 무엇이든 바꾸기 전에 너비를 확인하고 항상 머리글을 기록합니다. 표 이름은 Excel 규칙을 따릅니다. 1~255자여야 하고, 셀 참조처럼 보이면 안 되며, 표와 정의된 이름 사이에서 대소문자 구분 없이 고유해야 합니다.
- 조건부 서식은 일관된 우선순위, 수식, 스타일을 기록합니다.
- 보호 해시, 데이터 유효성 검사, 페이지 나누기, 메모는 Office 동작을 따릅니다.

어느 워크시트에 속하는지 분명하지 않은 표 part는 첫 번째 시트에 붙이지 않고 잘못된 package로 보고합니다. Lazy 워크북은 스파크라인, 그림, 메모, 컨트롤, 이름, 병합이 필요한 편집 전에 이를 먼저 읽으므로, 이 데이터를 지키려고 넣은 eager 열기 우회 코드는 제거할 수 있습니다.

이 기능들의 기준 파일은 새로 만들고, 저장 후 다시 열거나 Excel에서 확인하세요.

## 스타일과 렌더링

테마 tint, 상속된 스타일, 숫자 서식이 Excel과 같게 렌더링되므로 서식 적용 텍스트와 SVG 스냅샷이 바뀝니다. SVG 렌더링은 글꼴 이름을 직접 escape하므로 `A&B Sans`처럼 escape하지 않은 이름을 전달하세요. 색상은 6자리 RGB나 8자리 ARGB 16진수여야 하며 앞에 `#`를 붙일 수 있습니다. 그 밖의 값은 거부됩니다.

## Node.js 인수 검사

- 정수 인수와 숫자 옵션은 소수, NaN, 무한대, 안전하지 않은 정수, 범위를 벗어난 값을 잘라내지 않고 거부합니다.
- 알 수 없는 스타일 및 조건부 서식 enum 문자열은 허용되는 값 목록과 함께 오류를 발생시킵니다.

## CLI

`info`, `sheets`, `read`, `get`의 텍스트 출력은 터미널 제어 문자를 escape하고 크기가 제한됩니다. 스크립트에서는 JSON을 사용하세요.

```bash
sheetkit sheets input.xlsx --format json
```

`read --format csv`와 `convert`는 기본적으로 수식처럼 보이는 텍스트를 무력화합니다. 신뢰할 수 있는 데이터를 교환할 때만 원문을 유지하세요.

```bash
sheetkit read input.xlsx --sheet Sheet1 --format csv \
  --csv-formula-policy raw-exchange
```

출력이 제한을 넘으면 아무것도 기록하지 않습니다. `convert`가 실패하면 기존 대상 파일은 그대로 남습니다.

## 성능

Apple M4 Pro와 Node.js 24 환경에서 100,000행 워크북으로 0.5.1과 비교했을 때, lazy 열기는 약 29%, 변경하지 않은 lazy 워크북 저장은 약 45% 빨라졌습니다. 행 읽기, 스트리밍, 편집 후 저장은 0.5.1과 4% 이내로 차이가 납니다. v2 버퍼 생성은 추가된 검증 때문에 약 7%(1.8 ms) 느려졌습니다. 작업 환경이 다르다면 실제 파일로 측정하세요.

## 업그레이드 체크리스트

1. Rust와 Node.js 패키지를 함께 업그레이드하고, 다시 빌드하고, TypeScript를 다시 컴파일하세요. Node.js 24 이상을 사용하세요.
2. 오류 메시지 비교를 `Error::ResourceLimitExceeded`나 `isResourceLimitError`로 바꾸세요.
3. 필요한 신뢰할 수 있는 파일에 한해서만 제한을 높이세요.
4. Lazy와 eager 열기 중 무엇을 쓸지 정하고 날짜 처리를 확인하세요.
5. `DateValue`와 `FormulaValue`에 맞게 타입 가드와 직렬화 코드를 수정하세요.
6. 행을 사용하는 코드를 절대 좌표 기반 sparse 출력에 맞게 수정하세요.
7. 수식, 구조 편집, 숫자 서식, SVG 출력의 기대값을 갱신하세요.
8. CSV 수식 정책을 명시하고, CLI에서는 JSON 출력을 사용하세요.
9. 비동기 저장이 끝난 뒤에 워크북을 수정하세요.
10. 전원이 끊겨도 남아야 하는 저장에는 `durable: true`를 전달하세요.
11. 매크로 워크북은 `.xlsm`, `.xltm`, `.xlam`으로 저장하세요.
12. 표, 피벗, 차트, 슬라이서, 컨트롤, 메모, 그림이 있는 대표 파일을 저장한 뒤 다시 열어 확인하세요.
