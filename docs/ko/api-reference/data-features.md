## 8. 셀 병합

여러 셀을 하나로 병합하거나 해제하는 기능을 다룬다.

### `merge_cells(sheet, top_left, bottom_right)` / `mergeCells(sheet, topLeft, bottomRight)`

셀 범위를 병합한다.

**Rust:**

```rust
wb.merge_cells("Sheet1", "A1", "D1")?;
```

**TypeScript:**

```typescript
wb.mergeCells("Sheet1", "A1", "D1");
```

### `unmerge_cell(sheet, reference)` / `unmergeCell(sheet, reference)`

병합을 해제한다. 참조는 "A1:D1" 형식의 전체 범위 문자열이다.

**Rust:**

```rust
wb.unmerge_cell("Sheet1", "A1:D1")?;
```

**TypeScript:**

```typescript
wb.unmergeCell("Sheet1", "A1:D1");
```

### `get_merge_cells(sheet)` / `getMergeCells(sheet)`

시트의 모든 병합 범위를 반환한다.

**Rust:**

```rust
let merged: Vec<String> = wb.get_merge_cells("Sheet1")?;
// ["A1:D1", "B3:C5"]
```

**TypeScript:**

```typescript
const merged: string[] = wb.getMergeCells("Sheet1");
```

---

## 9. 하이퍼링크

셀에 하이퍼링크를 설정, 조회, 삭제하는 기능을 다룬다. 외부 URL, 내부 시트 참조, 이메일의 세 가지 유형을 지원한다.

### `set_cell_hyperlink` / `setCellHyperlink`

셀에 하이퍼링크를 설정한다.

**Rust:**

```rust
use sheetkit::hyperlink::HyperlinkType;

// External URL
wb.set_cell_hyperlink(
    "Sheet1", "A1",
    HyperlinkType::External("https://example.com".into()),
    Some("Example Site"),  // display text
    Some("Click here"),    // tooltip
)?;

// Internal sheet reference
wb.set_cell_hyperlink(
    "Sheet1", "A2",
    HyperlinkType::Internal("Sheet2!A1".into()),
    None, None,
)?;

// Email
wb.set_cell_hyperlink(
    "Sheet1", "A3",
    HyperlinkType::Email("mailto:user@example.com".into()),
    Some("Send email"), None,
)?;
```

**TypeScript:**

```typescript
// External URL
wb.setCellHyperlink("Sheet1", "A1", {
    linkType: "external",
    target: "https://example.com",
    display: "Example Site",
    tooltip: "Click here",
});

// Internal sheet reference
wb.setCellHyperlink("Sheet1", "A2", {
    linkType: "internal",
    target: "Sheet2!A1",
});

// Email
wb.setCellHyperlink("Sheet1", "A3", {
    linkType: "email",
    target: "mailto:user@example.com",
    display: "Send email",
});
```

### `get_cell_hyperlink` / `getCellHyperlink`

셀의 하이퍼링크 정보를 조회한다. 없으면 None / null을 반환한다.

**Rust:**

```rust
if let Some(info) = wb.get_cell_hyperlink("Sheet1", "A1")? {
    // info.link_type, info.display, info.tooltip
}
```

**TypeScript:**

```typescript
const info = wb.getCellHyperlink("Sheet1", "A1");
if (info) {
    console.log(info.linkType, info.target, info.display, info.tooltip);
}
```

**JsHyperlinkInfo 구조:**

```typescript
interface JsHyperlinkInfo {
  linkType: string;     // "external" | "internal" | "email"
  target: string;       // URL, sheet reference, or email address
  display?: string;     // display text
  tooltip?: string;     // tooltip text
}
```

### `delete_cell_hyperlink` / `deleteCellHyperlink`

셀의 하이퍼링크를 삭제한다.

**Rust:**

```rust
wb.delete_cell_hyperlink("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.deleteCellHyperlink("Sheet1", "A1");
```

> 외부/이메일 하이퍼링크는 워크시트 .rels 파일에 저장되고, 내부 하이퍼링크는 location 속성만 사용한다.

---

## 10. 차트

41가지 차트 유형을 지원한다. `add_chart`로 시트에 차트를 추가하며, 셀 범위로 위치와 크기를 지정한다.

### `add_chart(sheet, from_cell, to_cell, config)` / `addChart(sheet, fromCell, toCell, config)`

시트에 차트를 추가한다. `from_cell`은 차트의 왼쪽 상단, `to_cell`은 오른쪽 하단 위치를 나타낸다.

**Rust:**

```rust
use sheetkit::chart::*;

wb.add_chart("Sheet1", "E1", "L15", &ChartConfig {
    chart_type: ChartType::Col,
    title: Some("Sales Report".into()),
    series: vec![
        ChartSeries {
            name: "Revenue".into(),
            categories: "Sheet1!$A$2:$A$6".into(),
            values: "Sheet1!$B$2:$B$6".into(),
            x_values: None,
            bubble_sizes: None,
        },
    ],
    show_legend: true,
    view_3d: None,
})?;
```

**TypeScript:**

```typescript
wb.addChart("Sheet1", "E1", "L15", {
    chartType: "col",
    title: "Sales Report",
    series: [
        {
            name: "Revenue",
            categories: "Sheet1!$A$2:$A$6",
            values: "Sheet1!$B$2:$B$6",
        },
    ],
    showLegend: true,
});
```

### ChartSeries 구조

| 속성 | 타입 | 필수 | 설명 |
|------|------|------|------|
| `name` | `string` | O | 시리즈 이름 또는 셀 참조 |
| `categories` | `string` | O | 카테고리 데이터 범위 |
| `values` | `string` | O | 값 데이터 범위 |
| `x_values` / `xValues` | `string?` | X | Scatter/Bubble용 X 축 범위 |
| `bubble_sizes` / `bubbleSizes` | `string?` | X | Bubble 차트용 크기 범위 |

### View3DConfig 구조

3D 차트의 시점을 설정한다. 3D 차트 유형에서는 지정하지 않으면 자동으로 기본값이 적용된다.

| 속성 | 타입 | 설명 |
|------|------|------|
| `rot_x` / `rotX` | `i32?` / `number?` | X축 회전 각도 |
| `rot_y` / `rotY` | `i32?` / `number?` | Y축 회전 각도 |
| `depth_percent` / `depthPercent` | `u32?` / `number?` | 깊이 비율 (100 = 기본) |
| `right_angle_axes` / `rightAngleAxes` | `bool?` / `boolean?` | 직각 축 사용 여부 |
| `perspective` | `u32?` / `number?` | 원근 각도 |

### 차트 유형 전체 목록 (41종)

#### 세로 막대 (Column) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `col` | `ChartType::Col` | 세로 막대 |
| `colStacked` | `ChartType::ColStacked` | 누적 세로 막대 |
| `colPercentStacked` | `ChartType::ColPercentStacked` | 100% 누적 세로 막대 |
| `col3D` | `ChartType::Col3D` | 3D 세로 막대 |
| `col3DStacked` | `ChartType::Col3DStacked` | 3D 누적 세로 막대 |
| `col3DPercentStacked` | `ChartType::Col3DPercentStacked` | 3D 100% 누적 세로 막대 |

#### 가로 막대 (Bar) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `bar` | `ChartType::Bar` | 가로 막대 |
| `barStacked` | `ChartType::BarStacked` | 누적 가로 막대 |
| `barPercentStacked` | `ChartType::BarPercentStacked` | 100% 누적 가로 막대 |
| `bar3D` | `ChartType::Bar3D` | 3D 가로 막대 |
| `bar3DStacked` | `ChartType::Bar3DStacked` | 3D 누적 가로 막대 |
| `bar3DPercentStacked` | `ChartType::Bar3DPercentStacked` | 3D 100% 누적 가로 막대 |

#### 꺾은선 (Line) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `line` | `ChartType::Line` | 꺾은선 |
| `lineStacked` | `ChartType::LineStacked` | 누적 꺾은선 |
| `linePercentStacked` | `ChartType::LinePercentStacked` | 100% 누적 꺾은선 |
| `line3D` | `ChartType::Line3D` | 3D 꺾은선 |

#### 원형 (Pie) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `pie` | `ChartType::Pie` | 원형 |
| `pie3D` | `ChartType::Pie3D` | 3D 원형 |
| `doughnut` | `ChartType::Doughnut` | 도넛형 |

#### 영역 (Area) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `area` | `ChartType::Area` | 영역 |
| `areaStacked` | `ChartType::AreaStacked` | 누적 영역 |
| `areaPercentStacked` | `ChartType::AreaPercentStacked` | 100% 누적 영역 |
| `area3D` | `ChartType::Area3D` | 3D 영역 |
| `area3DStacked` | `ChartType::Area3DStacked` | 3D 누적 영역 |
| `area3DPercentStacked` | `ChartType::Area3DPercentStacked` | 3D 100% 누적 영역 |

#### 분산형 (Scatter) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `scatter` | `ChartType::Scatter` | 분산형 (표식만) |
| `scatterSmooth` | `ChartType::ScatterSmooth` | 부드러운 선 |
| `scatterStraight` | `ChartType::ScatterLine` | 직선 |

#### 방사형 (Radar) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `radar` | `ChartType::Radar` | 방사형 |
| `radarFilled` | `ChartType::RadarFilled` | 채워진 방사형 |
| `radarMarker` | `ChartType::RadarMarker` | 표식이 있는 방사형 |

#### 주식 (Stock) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `stockHLC` | `ChartType::StockHLC` | 고가-저가-종가 |
| `stockOHLC` | `ChartType::StockOHLC` | 시가-고가-저가-종가 |
| `stockVHLC` | `ChartType::StockVHLC` | 거래량-고가-저가-종가 |
| `stockVOHLC` | `ChartType::StockVOHLC` | 거래량-시가-고가-저가-종가 |

#### 기타 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `bubble` | `ChartType::Bubble` | 거품형 |
| `surface` | `ChartType::Surface` | 표면형 |
| `surfaceTop` | `ChartType::Surface3D` | 3D 표면형 |
| `surfaceWireframe` | `ChartType::SurfaceWireframe` | 와이어프레임 표면형 |
| `surfaceTopWireframe` | `ChartType::SurfaceWireframe3D` | 3D 와이어프레임 표면형 |

#### 콤보 (Combo) 차트

| 타입 문자열 | Rust Enum | 설명 |
|------------|-----------|------|
| `colLine` | `ChartType::ColLine` | 세로 막대 + 꺾은선 |
| `colLineStacked` | `ChartType::ColLineStacked` | 누적 세로 막대 + 꺾은선 |
| `colLinePercentStacked` | `ChartType::ColLinePercentStacked` | 100% 누적 세로 막대 + 꺾은선 |

---

## 11. 이미지

시트에 이미지를 삽입하는 기능을 다룬다. PNG, JPEG, GIF 형식을 지원한다.

### `add_image(sheet, config)` / `addImage(sheet, config)`

시트에 이미지를 추가한다.

**Rust:**

```rust
use sheetkit::image::{ImageConfig, ImageFormat};

let data = std::fs::read("logo.png")?;
wb.add_image("Sheet1", &ImageConfig {
    data,
    format: ImageFormat::Png,
    from_cell: "A1".into(),
    width_px: 200,
    height_px: 100,
})?;
```

**TypeScript:**

```typescript
import { readFileSync } from "fs";

const data = readFileSync("logo.png");
wb.addImage("Sheet1", {
    data: data,
    format: "png",
    fromCell: "A1",
    widthPx: 200,
    heightPx: 100,
});
```

### ImageConfig 구조

| 속성 | 타입 | 설명 |
|------|------|------|
| `data` | `Vec<u8>` / `Buffer` | 이미지 바이너리 데이터 |
| `format` | `ImageFormat` / `string` | `"png"`, `"jpeg"` (`"jpg"`), `"gif"` |
| `from_cell` / `fromCell` | `string` | 이미지 시작 위치 셀 |
| `width_px` / `widthPx` | `u32` / `number` | 너비 (픽셀) |
| `height_px` / `heightPx` | `u32` / `number` | 높이 (픽셀) |

---

## 12. 데이터 유효성 검사

셀에 입력 제한 규칙을 설정하여 사용자 입력을 검증하는 기능을 다룬다.

### `add_data_validation` / `addDataValidation`

데이터 유효성 검사 규칙을 추가한다.

**Rust:**

```rust
use sheetkit::validation::*;

// Dropdown list
wb.add_data_validation("Sheet1", &DataValidationConfig {
    sqref: "A1:A100".into(),
    validation_type: ValidationType::List,
    operator: None,
    formula1: Some("\"Option1,Option2,Option3\"".into()),
    formula2: None,
    allow_blank: true,
    error_style: Some(ErrorStyle::Stop),
    error_title: Some("Invalid".into()),
    error_message: Some("Select from the list".into()),
    prompt_title: Some("Selection".into()),
    prompt_message: Some("Choose an option".into()),
    show_input_message: true,
    show_error_message: true,
})?;

// Number range
wb.add_data_validation("Sheet1", &DataValidationConfig {
    sqref: "B1:B100".into(),
    validation_type: ValidationType::Whole,
    operator: Some(ValidationOperator::Between),
    formula1: Some("1".into()),
    formula2: Some("100".into()),
    allow_blank: true,
    error_style: Some(ErrorStyle::Stop),
    error_title: None,
    error_message: Some("Enter a number between 1 and 100".into()),
    prompt_title: None,
    prompt_message: None,
    show_input_message: false,
    show_error_message: true,
})?;
```

**TypeScript:**

```typescript
// Dropdown list
wb.addDataValidation("Sheet1", {
    sqref: "A1:A100",
    validationType: "list",
    formula1: '"Option1,Option2,Option3"',
    allowBlank: true,
    errorStyle: "stop",
    errorTitle: "Invalid",
    errorMessage: "Select from the list",
    promptTitle: "Selection",
    promptMessage: "Choose an option",
    showInputMessage: true,
    showErrorMessage: true,
});

// Number range
wb.addDataValidation("Sheet1", {
    sqref: "B1:B100",
    validationType: "whole",
    operator: "between",
    formula1: "1",
    formula2: "100",
    errorMessage: "Enter a number between 1 and 100",
    showErrorMessage: true,
});
```

> Node.js에서 `validationType`은 지원되는 값(`list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom`)만 허용되며, 지원되지 않는 값은 오류를 반환한다.

### `get_data_validations` / `getDataValidations`

시트의 모든 유효성 검사 규칙을 반환한다.

**Rust:**

```rust
let validations = wb.get_data_validations("Sheet1")?;
```

**TypeScript:**

```typescript
const validations = wb.getDataValidations("Sheet1");
```

### `remove_data_validation` / `removeDataValidation`

sqref로 유효성 검사를 제거한다.

**Rust:**

```rust
wb.remove_data_validation("Sheet1", "A1:A100")?;
```

**TypeScript:**

```typescript
wb.removeDataValidation("Sheet1", "A1:A100");
```

### 유효성 검사 유형 (7종)

| 값 | 설명 |
|----|------|
| `whole` | 정수 |
| `decimal` | 소수 |
| `list` | 드롭다운 목록 |
| `date` | 날짜 |
| `time` | 시각 |
| `textLength` | 텍스트 길이 |
| `custom` | 사용자 정의 수식 |

### 연산자 (8종)

| 값 | 설명 |
|----|------|
| `between` | 사이 (범위) |
| `notBetween` | 사이 아님 |
| `equal` | 같음 |
| `notEqual` | 같지 않음 |
| `lessThan` | 미만 |
| `lessThanOrEqual` | 이하 |
| `greaterThan` | 초과 |
| `greaterThanOrEqual` | 이상 |

### 오류 스타일 (3종)

| 값 | 설명 |
|----|------|
| `stop` | 입력 차단 (기본) |
| `warning` | 경고 (사용자가 무시 가능) |
| `information` | 정보 (사용자가 무시 가능) |

---

## 13. 코멘트

셀에 메모(코멘트)를 추가, 조회, 삭제하는 기능을 다룬다.

### `add_comment` / `addComment`

셀에 코멘트를 추가한다.

**Rust:**

```rust
use sheetkit::comment::CommentConfig;

wb.add_comment("Sheet1", &CommentConfig {
    cell: "A1".into(),
    author: "Admin".into(),
    text: "Review this value".into(),
})?;
```

**TypeScript:**

```typescript
wb.addComment("Sheet1", {
    cell: "A1",
    author: "Admin",
    text: "Review this value",
});
```

### `get_comments` / `getComments`

시트의 모든 코멘트를 반환한다.

**Rust:**

```rust
let comments = wb.get_comments("Sheet1")?;
for c in &comments {
    println!("{}: {} - {}", c.cell, c.author, c.text);
}
```

**TypeScript:**

```typescript
const comments = wb.getComments("Sheet1");
for (const c of comments) {
    console.log(`${c.cell}: ${c.author} - ${c.text}`);
}
```

### `remove_comment` / `removeComment`

셀의 코멘트를 삭제한다.

**Rust:**

```rust
wb.remove_comment("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.removeComment("Sheet1", "A1");
```

---

## 14. 자동 필터

데이터 범위에 자동 필터를 설정하거나 제거하는 기능을 다룬다.

### `set_auto_filter(sheet, range)` / `setAutoFilter(sheet, range)`

셀 범위에 자동 필터를 설정한다.

**Rust:**

```rust
wb.set_auto_filter("Sheet1", "A1:D10")?;
```

**TypeScript:**

```typescript
wb.setAutoFilter("Sheet1", "A1:D10");
```

### `remove_auto_filter(sheet)` / `removeAutoFilter(sheet)`

시트의 자동 필터를 제거한다.

**Rust:**

```rust
wb.remove_auto_filter("Sheet1")?;
```

**TypeScript:**

```typescript
wb.removeAutoFilter("Sheet1");
```

---

## 15. 조건부 서식

셀 값이나 수식에 따라 자동으로 서식을 적용하는 18가지 규칙 유형을 지원한다.

### `set_conditional_format` / `setConditionalFormat`

셀 범위에 조건부 서식 규칙을 설정한다.

**Rust:**

```rust
use sheetkit::conditional::*;
use sheetkit::style::*;

// cellIs rule
wb.set_conditional_format("Sheet1", "A1:A100", &[
    ConditionalFormatRule {
        rule_type: ConditionalFormatType::CellIs {
            operator: CfOperator::GreaterThan,
            formula: "90".into(),
            formula2: None,
        },
        format: Some(ConditionalStyle {
            font: Some(FontStyle {
                bold: true,
                color: Some(StyleColor::Rgb("#006100".into())),
                ..Default::default()
            }),
            fill: Some(FillStyle {
                pattern: PatternType::Solid,
                fg_color: Some(StyleColor::Rgb("#C6EFCE".into())),
                bg_color: None,
            }),
            border: None,
            num_fmt: None,
        }),
        priority: Some(1),
        stop_if_true: false,
    },
])?;
```

**TypeScript:**

```typescript
// cellIs rule
wb.setConditionalFormat("Sheet1", "A1:A100", [
    {
        ruleType: "cellIs",
        operator: "greaterThan",
        formula: "90",
        format: {
            font: { bold: true, color: "#006100" },
            fill: { pattern: "solid", fgColor: "#C6EFCE" },
        },
        priority: 1,
    },
]);
```

### 조건부 서식 예제

#### colorScale (색상 스케일)

```typescript
wb.setConditionalFormat("Sheet1", "B1:B50", [
    {
        ruleType: "colorScale",
        minType: "min",
        minColor: "FFF8696B",  // red
        midType: "percentile",
        midValue: "50",
        midColor: "FFFFEB84",  // yellow
        maxType: "max",
        maxColor: "FF63BE7B",  // green
    },
]);
```

#### dataBar (데이터 막대)

```typescript
wb.setConditionalFormat("Sheet1", "C1:C50", [
    {
        ruleType: "dataBar",
        barColor: "FF638EC6",
        showValue: true,
    },
]);
```

#### containsText (텍스트 포함)

```typescript
wb.setConditionalFormat("Sheet1", "D1:D100", [
    {
        ruleType: "containsText",
        text: "Error",
        format: {
            font: { color: "#FF0000", bold: true },
        },
    },
]);
```

### `get_conditional_formats` / `getConditionalFormats`

시트의 모든 조건부 서식 규칙을 반환한다.

**Rust:**

```rust
let formats = wb.get_conditional_formats("Sheet1")?;
// Vec<(String, Vec<ConditionalFormatRule>)>
// (sqref, rules)
```

**TypeScript:**

```typescript
const formats = wb.getConditionalFormats("Sheet1");
// JsConditionalFormatEntry[]
for (const entry of formats) {
    console.log(`Range: ${entry.sqref}, Rules: ${entry.rules.length}`);
}
```

### `delete_conditional_format` / `deleteConditionalFormat`

특정 셀 범위의 조건부 서식을 삭제한다.

**Rust:**

```rust
wb.delete_conditional_format("Sheet1", "A1:A100")?;
```

**TypeScript:**

```typescript
wb.deleteConditionalFormat("Sheet1", "A1:A100");
```

### 규칙 유형 (18종)

| 규칙 유형 | 설명 | 필수 속성 |
|-----------|------|-----------|
| `cellIs` | 셀 값 비교 | `operator`, `formula`, `formula2`(between용) |
| `expression` | 수식 결과 기반 | `formula` |
| `colorScale` | 색상 스케일 | `minType/minColor`, `maxType/maxColor`, `midType/midColor`(선택) |
| `dataBar` | 데이터 막대 | `barColor`, `showValue` |
| `duplicateValues` | 중복 값 | -- |
| `uniqueValues` | 고유 값 | -- |
| `top10` | 상위 N개 | `rank`, `percent` |
| `bottom10` | 하위 N개 | `rank`, `percent` |
| `aboveAverage` | 평균 이상/이하 | `above`, `equalAverage` |
| `containsBlanks` | 빈 셀 포함 | -- |
| `notContainsBlanks` | 빈 셀 미포함 | -- |
| `containsErrors` | 오류 포함 | -- |
| `notContainsErrors` | 오류 미포함 | -- |
| `containsText` | 텍스트 포함 | `text` |
| `notContainsText` | 텍스트 미포함 | `text` |
| `beginsWith` | 텍스트로 시작 | `text` |
| `endsWith` | 텍스트로 끝남 | `text` |
| `expression` | 사용자 정의 수식 | `formula` |

### CfValueType (색상 스케일/데이터 막대용)

| 값 | 설명 |
|----|------|
| `num` | 숫자 값 |
| `percent` | 백분율 |
| `min` | 최소값 |
| `max` | 최대값 |
| `percentile` | 백분위수 |
| `formula` | 수식 |

---
