# SheetKit 시작 가이드

SheetKit은 Rust와 TypeScript를 위한 고성능 SpreadsheetML 라이브러리입니다.
Rust 코어가 모든 Excel (.xlsx) 처리를 담당하며, napi-rs 바인딩을 통해 TypeScript에서도 최소한의 overhead로 동일한 성능을 제공합니다.

## SheetKit을 선택해야 하는 이유

- **네이티브 성능**: Rust 코어와 저오버헤드 Node.js 바인딩으로 대용량 스프레드시트를 빠르게 처리합니다
- **FFI 오버헤드 최소화**: Raw Buffer 기반 전송으로 Node.js와 Rust 간 FFI 경계 오버헤드를 줄입니다
- **타입 안전성**: Rust와 TypeScript 모두에서 강력한 타입 안전 API를 제공합니다
- **완전한 기능**: 110개 이상의 수식 함수, 43가지 차트 타입, 스트리밍 쓰기 등을 지원합니다

## 설치

### Rust 라이브러리

`cargo add` 명령어를 사용하여 설치합니다 (권장):

```bash
cargo add sheetkit
```

또는 `Cargo.toml`에 직접 추가합니다:

```toml
[dependencies]
sheetkit = { version = "0.4" }
```

암호화 기능이 필요한 경우:

```bash
cargo add sheetkit --features encryption
```

[crates.io에서 보기](https://crates.io/crates/sheetkit)

### Node.js 라이브러리

선호하는 패키지 매니저를 사용하여 설치합니다:

```bash
# npm
npm install @sheetkit/node

# yarn
yarn add @sheetkit/node

# pnpm
pnpm add @sheetkit/node
```

참고: 주요 플랫폼에는 사전 빌드 바이너리가 제공됩니다. 사전 빌드 바이너리가 없는 환경이거나 소스에서 직접 빌드하는 경우에만 Rust 툴체인이 필요합니다.

[npm에서 보기](https://www.npmjs.com/package/@sheetkit/node)

### CLI 도구

커맨드 라인에서 시트 검사, 데이터 변환 등의 작업을 수행하려면:

```bash
cargo install sheetkit --features cli
```

사용 방법은 [CLI 가이드](./guide/cli.md)를 참조하세요.

## 빠른 시작

### 워크북 생성 및 셀 쓰기

**Rust**

```rust
use sheetkit::{CellValue, Workbook};

fn main() -> sheetkit::Result<()> {
    let mut wb = Workbook::new();

    // Write different value types
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Name".into()))?;
    wb.set_cell_value("Sheet1", "B1", CellValue::Number(42.0))?;
    wb.set_cell_value("Sheet1", "C1", CellValue::Bool(true))?;
    wb.set_cell_value("Sheet1", "D1", CellValue::Empty)?;

    // Read a cell value
    let val = wb.get_cell_value("Sheet1", "A1")?;
    println!("A1 = {:?}", val); // CellValue::String("Name")

    // Save to file
    wb.save("output.xlsx")?;
    Ok(())
}
```

**TypeScript**

```typescript
import { Workbook } from "@sheetkit/node";

const wb = new Workbook();

// Write different value types
wb.setCellValue("Sheet1", "A1", "Name");
wb.setCellValue("Sheet1", "B1", 42);
wb.setCellValue("Sheet1", "C1", true);
wb.setCellValue("Sheet1", "D1", null); // clear/empty

// Read a cell value
const val = wb.getCellValue("Sheet1", "A1");
console.log("A1 =", val); // "Name"

// Save to file
await wb.save("output.xlsx");
```

### 기존 파일 열기

**Rust**

```rust
use sheetkit::Workbook;

fn main() -> sheetkit::Result<()> {
    let wb = Workbook::open("input.xlsx")?;

    // List all sheet names
    let names = wb.sheet_names();
    println!("Sheets: {:?}", names);

    // Read a cell from the first sheet
    let val = wb.get_cell_value(&names[0], "A1")?;
    println!("A1 = {:?}", val);

    Ok(())
}
```

**TypeScript**

```typescript
import { Workbook } from "@sheetkit/node";

const wb = await Workbook.open("input.xlsx");

// List all sheet names
console.log("Sheets:", wb.sheetNames);

// Read a cell from the first sheet
const val = wb.getCellValue(wb.sheetNames[0], "A1");
console.log("A1 =", val);
```

## 핵심 개념

### CellValue 타입

SheetKit은 타입이 지정된 셀 값 모델을 사용합니다. 모든 셀은 다음 중 하나의 값을 가집니다:

| 타입    | Rust                                            | TypeScript               |
| ------- | ----------------------------------------------- | ------------------------ |
| String  | `CellValue::String(String)`                     | `string`                 |
| Number  | `CellValue::Number(f64)`                        | `number`                 |
| Bool    | `CellValue::Bool(bool)`                         | `boolean`                |
| Empty   | `CellValue::Empty`                              | `null`                   |
| Date    | `CellValue::Date(f64)`                          | `{ type: 'date', serial: number, iso?: string }` |
| Formula | `CellValue::Formula { expr, result }`           | *(수식 평가를 통해 설정)* |
| Error   | `CellValue::Error(String)`                      | *(읽기 전용)*            |

### 날짜 값

날짜는 Excel 시리얼 넘버(1900-01-01 기준 경과 일수)로 저장됩니다. 정수 부분은 날짜, 소수 부분은 시간을 나타냅니다 (예: 0.5 = 정오).

날짜를 Excel에서 올바르게 표시하려면 셀에 날짜 숫자 형식 스타일을 적용해야 합니다.

**Rust**

```rust
use sheetkit::{CellValue, Style, NumFmtStyle, Workbook};

let mut wb = Workbook::new();

// Excel serial 45292 = 2024-01-01
wb.set_cell_value("Sheet1", "A1", CellValue::Date(45292.0))?;

// Apply a date format so Excel renders it as a date
let style_id = wb.add_style(&Style {
    num_fmt: Some(NumFmtStyle { num_fmt_id: Some(14), custom_format: None }),
    ..Default::default()
})?;
wb.set_cell_style("Sheet1", "A1", style_id)?;
```

**TypeScript**

```typescript
const wb = new Workbook();

// Excel serial 45292 = 2024-01-01
wb.setCellValue("Sheet1", "A1", { type: "date", serial: 45292 });

// Apply a date format so Excel renders it as a date
const styleId = wb.addStyle({ numFmtId: 14 });
wb.setCellStyle("Sheet1", "A1", styleId);
```

날짜 셀을 다시 읽으면 `DateValue` 객체에 ISO 8601 문자열 표현이 포함된 선택적 `iso` 필드가 포함됩니다 (예: `"2024-01-01"` 또는 `"2024-01-01T12:00:00"`).

### 셀 참조

셀 참조는 A1 스타일 표기법을 사용합니다: 열 문자 뒤에 1 기반 행 번호가 옵니다.

- `"A1"` -- A열, 1행
- `"B2"` -- B열, 2행
- `"AA100"` -- AA열 (27번째 열), 100행

### 시트 이름

시트 이름은 대소문자를 구분하는 문자열입니다. 새 워크북은 `"Sheet1"`이라는 단일 시트로 시작합니다.

### 1 기반 vs 0 기반 인덱싱

- **행**: 1 기반 (행 1이 첫 번째 행)
- **열 번호**: 1 기반 (열 1 = "A")
- **시트 인덱스**: 0 기반 (첫 번째 시트의 인덱스는 0)

### 스타일 시스템

스타일은 등록 후 적용 패턴을 따릅니다:

1. 글꼴, 채우기, 테두리, 정렬, 숫자 형식 옵션으로 `Style` 구조체/객체를 정의합니다.
2. `add_style` / `addStyle`로 등록하여 숫자 스타일 ID를 받습니다.
3. 스타일 ID를 셀, 행 또는 열에 적용합니다.

스타일 중복 제거는 자동으로 처리됩니다. 동일한 스타일을 두 번 등록하면 같은 ID가 반환됩니다.

## 스타일 사용하기

아래 예시는 재사용 가능한 스타일 정의를 한 번 등록하고, 반환된 스타일 ID를 대상 셀에 적용하는 방식입니다. 이 패턴을 사용하면 `styles.xml`의 중복 스타일 레코드를 줄여 워크북 구조를 더 간결하게 유지할 수 있습니다.

**Rust**

```rust
use sheetkit::{
    CellValue, FillStyle, FontStyle, PatternType, Style, StyleColor, Workbook,
};

let mut wb = Workbook::new();
wb.set_cell_value("Sheet1", "A1", CellValue::String("Styled".into()))?;

let style_id = wb.add_style(&Style {
    font: Some(FontStyle {
        bold: true,
        size: Some(14.0),
        color: Some(StyleColor::Rgb("#FFFFFF".into())),
        ..Default::default()
    }),
    fill: Some(FillStyle {
        pattern: PatternType::Solid,
        fg_color: Some(StyleColor::Rgb("#4472C4".into())),
        bg_color: None,
    }),
    ..Default::default()
})?;

wb.set_cell_style("Sheet1", "A1", style_id)?;
wb.save("styled.xlsx")?;
```

**TypeScript**

```typescript
const wb = new Workbook();
wb.setCellValue("Sheet1", "A1", "Styled");

const styleId = wb.addStyle({
  font: { bold: true, size: 14, color: "#FFFFFF" },
  fill: { pattern: "solid", fgColor: "#4472C4" },
});

wb.setCellStyle("Sheet1", "A1", styleId);
await wb.save("styled.xlsx");
```

## 차트 사용하기

앵커 범위(왼쪽 상단, 오른쪽 하단 셀), 차트 유형, 데이터 시리즈를 지정하여 차트를 추가합니다.
앵커 셀은 차트의 배치 위치를 제어하고, `categories`와 `values` 범위는 실제 플롯 데이터를 제어합니다. 범주 범위와 값 범위의 길이를 맞춰야 의도한 차트 결과를 얻을 수 있습니다.

**Rust**

```rust
use sheetkit::{CellValue, ChartConfig, ChartSeries, ChartType, Workbook};

let mut wb = Workbook::new();

// Prepare data
wb.set_cell_value("Sheet1", "A1", CellValue::String("Q1".into()))?;
wb.set_cell_value("Sheet1", "A2", CellValue::String("Q2".into()))?;
wb.set_cell_value("Sheet1", "A3", CellValue::String("Q3".into()))?;
wb.set_cell_value("Sheet1", "B1", CellValue::Number(1500.0))?;
wb.set_cell_value("Sheet1", "B2", CellValue::Number(2300.0))?;
wb.set_cell_value("Sheet1", "B3", CellValue::Number(1800.0))?;

// Add a column chart
wb.add_chart(
    "Sheet1",
    "D1",   // top-left anchor
    "K15",  // bottom-right anchor
    &ChartConfig {
        chart_type: ChartType::Col,
        title: Some("Quarterly Revenue".into()),
        series: vec![ChartSeries {
            name: "Revenue".into(),
            categories: "Sheet1!$A$1:$A$3".into(),
            values: "Sheet1!$B$1:$B$3".into(),
            x_values: None,
            bubble_sizes: None,
        }],
        show_legend: true,
        view_3d: None,
    },
)?;

wb.save("chart.xlsx")?;
```

**TypeScript**

```typescript
const wb = new Workbook();

// Prepare data
wb.setCellValue("Sheet1", "A1", "Q1");
wb.setCellValue("Sheet1", "A2", "Q2");
wb.setCellValue("Sheet1", "A3", "Q3");
wb.setCellValue("Sheet1", "B1", 1500);
wb.setCellValue("Sheet1", "B2", 2300);
wb.setCellValue("Sheet1", "B3", 1800);

// Add a column chart
wb.addChart("Sheet1", "D1", "K15", {
  chartType: "col",
  title: "Quarterly Revenue",
  series: [
    {
      name: "Revenue",
      categories: "Sheet1!$A$1:$A$3",
      values: "Sheet1!$B$1:$B$3",
    },
  ],
  showLegend: true,
});

await wb.save("chart.xlsx");
```

## 대용량 파일을 위한 StreamWriter

`StreamWriter`는 행 데이터를 디스크의 임시 파일에 직접 기록하여, 행 수에 관계없이 일정한 메모리 사용량을 유지합니다. 각 `write_row()` 호출은 행을 XML로 serialize하여 임시 파일에 추가합니다. 저장 시 행 데이터가 임시 파일에서 ZIP 아카이브로 직접 스트리밍됩니다.

행은 오름차순으로 작성해야 하며, 열 너비는 행을 쓰기 전에 설정해야 합니다. `apply_stream_writer` 이후 스트리밍된 시트의 셀 값은 직접 읽을 수 없습니다. 데이터를 읽으려면 워크북을 저장한 후 다시 열어야 합니다.

이 방식은 배치 내보내기, ETL 작업, 스트림 입력처럼 행 데이터가 점진적으로 생성되는 워크로드에 적합합니다.

**Rust**

```rust
use sheetkit::{CellValue, Workbook};

let mut wb = Workbook::new();
let mut sw = wb.new_stream_writer("LargeSheet")?;

// Set column widths
sw.set_col_width(1, 20.0)?;
sw.set_col_width(2, 15.0)?;

// Write header
sw.write_row(1, &[
    CellValue::String("Item".into()),
    CellValue::String("Value".into()),
])?;

// Write 10,000 data rows
for i in 2..=10_001 {
    sw.write_row(i, &[
        CellValue::String(format!("Item_{}", i - 1)),
        CellValue::Number(i as f64 * 1.5),
    ])?;
}

// Apply the stream writer output to the workbook
wb.apply_stream_writer(sw)?;
wb.save("large.xlsx")?;
```

**TypeScript**

```typescript
const wb = new Workbook();
const sw = wb.newStreamWriter("LargeSheet");

// Set column widths
sw.setColWidth(1, 20);
sw.setColWidth(2, 15);

// Write header
sw.writeRow(1, ["Item", "Value"]);

// Write 10,000 data rows
for (let i = 2; i <= 10_001; i++) {
  sw.writeRow(i, [`Item_${i - 1}`, i * 1.5]);
}

// Apply the stream writer output to the workbook
wb.applyStreamWriter(sw);
await wb.save("large.xlsx");
```

## 암호화된 파일 다루기

SheetKit은 비밀번호로 보호된 .xlsx 파일의 읽기/쓰기를 지원합니다. Rust에서는 `encryption` feature를 활성화해야 하며, Node.js 바인딩에는 항상 암호화 지원이 포함됩니다.
이 기능은 OOXML 패키지 수준의 파일 암호화를 적용하므로 비밀번호 없이는 파일을 열 수 없습니다. 운영 환경에서는 암호화된 파일 경로를 도입하기 전에 비밀번호 관리 및 복구 절차를 함께 설계하는 것이 좋습니다.

**Rust**

```rust
use sheetkit::Workbook;

// 비밀번호로 저장
let wb = Workbook::new();
wb.save_with_password("encrypted.xlsx", "secret")?;

// 비밀번호로 열기
let wb2 = Workbook::open_with_password("encrypted.xlsx", "secret")?;
```

**TypeScript**

```typescript
import { Workbook } from "@sheetkit/node";

// 비밀번호로 저장
const wb = new Workbook();
wb.saveWithPassword("encrypted.xlsx", "secret");

// 비밀번호로 열기
const wb2 = Workbook.openWithPasswordSync("encrypted.xlsx", "secret");
```

## 다음 단계

- [API 레퍼런스](./api-reference/index.md) -- 모든 메서드와 타입에 대한 전체 문서.
- [아키텍처](./architecture.md) -- 내부 설계 및 크레이트 구조.
- [기여 가이드](./contributing.md) -- 개발 환경 설정 및 기여 가이드라인.
