## 8. 셀 병합

여러 셀을 하나로 병합하거나 해제하는 기능을 다룹니다.

### `merge_cells(sheet, top_left, bottom_right)` / `mergeCells(sheet, topLeft, bottomRight)`

셀 범위를 병합합니다.

**Rust:**

```rust
wb.merge_cells("Sheet1", "A1", "D1")?;
```

**TypeScript:**

```typescript
wb.mergeCells("Sheet1", "A1", "D1");
```

### `unmerge_cell(sheet, reference)` / `unmergeCell(sheet, reference)`

병합을 해제합니다. 참조는 "A1:D1" 형식의 전체 범위 문자열입니다.

**Rust:**

```rust
wb.unmerge_cell("Sheet1", "A1:D1")?;
```

**TypeScript:**

```typescript
wb.unmergeCell("Sheet1", "A1:D1");
```

### `get_merge_cells(sheet)` / `getMergeCells(sheet)`

시트의 모든 병합 범위를 반환합니다.

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

셀에 하이퍼링크를 설정, 조회, 삭제하는 기능을 다룹니다. 외부 URL, 내부 시트 참조, 이메일의 세 가지 유형을 지원합니다.

### `set_cell_hyperlink` / `setCellHyperlink`

셀에 하이퍼링크를 설정합니다.

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

셀의 하이퍼링크 정보를 조회합니다. 없으면 None / null을 반환합니다.

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

셀의 하이퍼링크를 삭제합니다.

**Rust:**

```rust
wb.delete_cell_hyperlink("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.deleteCellHyperlink("Sheet1", "A1");
```

> 외부/이메일 하이퍼링크는 워크시트 .rels 파일에 저장되고, 내부 하이퍼링크는 location 속성만 사용합니다.

---

## 10. 차트

43가지 차트 유형을 지원합니다. `add_chart`로 시트에 차트를 추가하며, 셀 범위로 위치와 크기를 지정합니다.

### `add_chart(sheet, from_cell, to_cell, config)` / `addChart(sheet, fromCell, toCell, config)`

시트에 차트를 추가합니다. `from_cell`은 차트의 왼쪽 상단, `to_cell`은 오른쪽 하단 위치를 나타냅니다.

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

3D 차트의 시점을 설정합니다. 3D 차트 유형에서는 지정하지 않으면 자동으로 기본값이 적용됩니다.

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

시트에 이미지를 삽입하는 기능을 다룹니다. 11가지 형식을 지원합니다: PNG, JPEG, GIF, BMP, ICO, TIFF, SVG, EMF, EMZ, WMF, WMZ.

### `add_image(sheet, config)` / `addImage(sheet, config)`

시트에 이미지를 추가합니다.

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
| `format` | `ImageFormat` / `string` | 11가지 형식 지원 ([이미지](./image.md#지원-형식) 참조) |
| `from_cell` / `fromCell` | `string` | 이미지 시작 위치 셀 |
| `width_px` / `widthPx` | `u32` / `number` | 너비 (픽셀) |
| `height_px` / `heightPx` | `u32` / `number` | 높이 (픽셀) |

---

## 12. 데이터 유효성 검사

셀에 입력 제한 규칙을 설정하여 사용자 입력을 검증하는 기능을 다룹니다.

### `add_data_validation` / `addDataValidation`

데이터 유효성 검사 규칙을 추가합니다.

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

> Node.js에서 `validationType`은 지원되는 값(`none`, `list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom`)만 허용되며, 지원되지 않는 값은 오류를 반환합니다. `sqref`는 유효한 셀 범위여야 합니다(예: `"A1:B10"`). `none` 이외의 타입에는 `formula1`이 필수이며, `between`/`notBetween` 연산자에는 `formula2`도 필수입니다.

### `get_data_validations` / `getDataValidations`

시트의 모든 유효성 검사 규칙을 반환합니다.

**Rust:**

```rust
let validations = wb.get_data_validations("Sheet1")?;
```

**TypeScript:**

```typescript
const validations = wb.getDataValidations("Sheet1");
```

### `remove_data_validation` / `removeDataValidation`

sqref로 유효성 검사를 제거합니다.

**Rust:**

```rust
wb.remove_data_validation("Sheet1", "A1:A100")?;
```

**TypeScript:**

```typescript
wb.removeDataValidation("Sheet1", "A1:A100");
```

### 유효성 검사 유형 (8종)

| 값 | 설명 |
|----|------|
| `none` | 제한 없음 (프롬프트/메시지만 표시) |
| `whole` | 정수 |
| `decimal` | 소수 |
| `list` | 드롭다운 목록 |
| `date` | 날짜 |
| `time` | 시각 |
| `textLength` | 텍스트 길이 |
| `custom` | 사용자 정의 수식 |

### 연산자 (8종)

TypeScript 입력은 대소문자를 구분하지 않으며, 출력은 OOXML 규격에 맞는 camelCase를 사용합니다:

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

셀에 메모(코멘트)를 추가, 조회, 삭제하는 기능을 다룹니다.

코멘트를 추가하면 SheetKit이 자동으로 VML(Vector Markup Language) 드로잉 파트(`xl/drawings/vmlDrawingN.vml`)와 워크시트 XML의 `<legacyDrawing>` 참조를 생성합니다. 이를 통해 Excel UI에서 코멘트 팝업 상자가 올바르게 렌더링됩니다.

기존 VML 코멘트 드로잉이 포함된 워크북을 열면 SheetKit이 저장/열기 사이클을 통해 VML 파트를 보존합니다. 시트의 모든 코멘트가 제거되면 관련 VML 파트와 관계가 자동으로 정리됩니다.

### `add_comment` / `addComment`

셀에 코멘트를 추가합니다.

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

시트의 모든 코멘트를 반환합니다.

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

셀의 코멘트를 삭제합니다.

**Rust:**

```rust
wb.remove_comment("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.removeComment("Sheet1", "A1");
```

### VML 호환성

Excel은 코멘트 노트 상자를 렌더링하기 위해 VML(Vector Markup Language)을 사용합니다. SheetKit은 다음을 자동으로 처리합니다:

- 새 코멘트 생성 시 최소한의 VML 드로잉 파트 생성
- 디스크에서 열린 워크북의 기존 VML 파트 보존
- 워크시트 XML에 `<legacyDrawing>` 관계 참조 연결
- VML 파트에 대한 적절한 콘텐츠 타입 항목 추가
- 시트의 모든 코멘트가 제거될 때 VML 파트 및 관계 정리

추가적인 API 호출이 필요하지 않습니다. VML 처리는 사용자에게 투명하게 수행됩니다.

---

## 14. 자동 필터

데이터 범위에 자동 필터를 설정하거나 제거하는 기능을 다룹니다.

### `set_auto_filter(sheet, range)` / `setAutoFilter(sheet, range)`

셀 범위에 자동 필터를 설정합니다.

**Rust:**

```rust
wb.set_auto_filter("Sheet1", "A1:D10")?;
```

**TypeScript:**

```typescript
wb.setAutoFilter("Sheet1", "A1:D10");
```

### `remove_auto_filter(sheet)` / `removeAutoFilter(sheet)`

시트의 자동 필터를 제거합니다.

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

셀 값이나 수식에 따라 자동으로 서식을 적용하는 17가지 규칙 유형을 지원합니다.

### `set_conditional_format` / `setConditionalFormat`

셀 범위에 조건부 서식 규칙을 설정합니다.

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

시트의 모든 조건부 서식 규칙을 반환합니다.

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

특정 셀 범위의 조건부 서식을 삭제합니다.

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

## 16. 테이블

테이블은 헤더, 스타일, 선택적 자동 필터가 포함된 구조화된 데이터 범위입니다. 테이블은 별도의 OOXML 파트(`xl/tables/tableN.xml`)로 저장되며 관계 및 콘텐츠 타입이 자동으로 연결됩니다.

### `add_table(sheet, config)` / `addTable(sheet, config)`

시트에 테이블을 생성합니다. 테이블 이름은 워크북 전체에서 고유해야 합니다.

**Rust:**

```rust
use sheetkit::table::{TableConfig, TableColumn};

let config = TableConfig {
    name: "Sales".to_string(),
    display_name: "Sales".to_string(),
    range: "A1:C10".to_string(),
    columns: vec![
        TableColumn { name: "Product".to_string(), totals_row_function: None, totals_row_label: None },
        TableColumn { name: "Quantity".to_string(), totals_row_function: None, totals_row_label: None },
        TableColumn { name: "Price".to_string(), totals_row_function: None, totals_row_label: None },
    ],
    show_header_row: true,
    style_name: Some("TableStyleMedium2".to_string()),
    auto_filter: true,
    ..TableConfig::default()
};
wb.add_table("Sheet1", &config)?;
```

**TypeScript:**

```typescript
wb.addTable("Sheet1", {
    name: "Sales",
    displayName: "Sales",
    range: "A1:C10",
    columns: [
        { name: "Product" },
        { name: "Quantity" },
        { name: "Price" },
    ],
    showHeaderRow: true,
    styleName: "TableStyleMedium2",
    autoFilter: true,
});
```

### `get_tables(sheet)` / `getTables(sheet)`

시트의 모든 테이블을 조회합니다.

**Rust:**

```rust
let tables = wb.get_tables("Sheet1")?;
for t in &tables {
    println!("{}: {} ({})", t.name, t.range, t.columns.join(", "));
}
```

**TypeScript:**

```typescript
const tables = wb.getTables("Sheet1");
for (const t of tables) {
    console.log(`${t.name}: ${t.range}`);
}
```

### `delete_table(sheet, name)` / `deleteTable(sheet, name)`

시트에서 이름으로 테이블을 삭제합니다.

**Rust:**

```rust
wb.delete_table("Sheet1", "Sales")?;
```

**TypeScript:**

```typescript
wb.deleteTable("Sheet1", "Sales");
```

### TableConfig

| 필드 | Rust 타입 | TS 타입 | 설명 |
|------|-----------|---------|------|
| `name` | `String` | `string` | 내부 테이블 이름 (워크북 내 고유) |
| `display_name` | `String` | `string` | UI에 표시되는 이름 |
| `range` | `String` | `string` | 셀 범위 (예: "A1:D10") |
| `columns` | `Vec<TableColumn>` | `TableColumn[]` | 열 정의 |
| `show_header_row` | `bool` | `boolean?` | 헤더 행 표시 (기본값: true) |
| `style_name` | `Option<String>` | `string?` | 테이블 스타일 (예: "TableStyleMedium2") |
| `auto_filter` | `bool` | `boolean?` | 자동 필터 사용 (기본값: true) |
| `show_first_column` | `bool` | `boolean?` | 첫 번째 열 강조 (기본값: false) |
| `show_last_column` | `bool` | `boolean?` | 마지막 열 강조 (기본값: false) |
| `show_row_stripes` | `bool` | `boolean?` | 행 줄무늬 표시 (기본값: true) |
| `show_column_stripes` | `bool` | `boolean?` | 열 줄무늬 표시 (기본값: false) |

### TableColumn

| 필드 | Rust 타입 | TS 타입 | 설명 |
|------|-----------|---------|------|
| `name` | `String` | `string` | 열 헤더 이름 |
| `totals_row_function` | `Option<String>` | `string?` | 합계 행 함수 (예: "sum", "count", "average") |
| `totals_row_label` | `Option<String>` | `string?` | 합계 행 레이블 (첫 번째 열용) |

### TableInfo (`get_tables` 반환 타입)

| 필드 | Rust 타입 | TS 타입 | 설명 |
|------|-----------|---------|------|
| `name` | `String` | `string` | 테이블 이름 |
| `display_name` | `String` | `string` | 표시 이름 |
| `range` | `String` | `string` | 셀 범위 |
| `show_header_row` | `bool` | `boolean` | 헤더 행 표시 여부 |
| `auto_filter` | `bool` | `boolean` | 자동 필터 사용 여부 |
| `columns` | `Vec<String>` | `string[]` | 열 헤더 이름 목록 |
| `style_name` | `Option<String>` | `string \| null` | 테이블 스타일 이름 |

> 테이블 이름은 단일 시트가 아닌 워크북 전체에서 고유해야 합니다. 시트가 삭제되면 해당 시트의 모든 테이블이 자동으로 제거됩니다.

---

## 17. 데이터 변환 유틸리티 (Node.js 전용)

시트 데이터와 일반적인 형식(JSON, CSV, HTML) 간의 변환을 위한 편의 메서드입니다. TypeScript/Node.js 바인딩에서만 사용할 수 있습니다.

### `toJSON(sheet, options?)`

시트를 객체 배열로 변환합니다. 각 객체는 첫 번째 행의 열 헤더를 키로 사용합니다.

```typescript
const wb = await Workbook.open("data.xlsx");
const records = wb.toJSON("Sheet1");
// [{ Name: "Alice", Age: 30, City: "Seoul" }, ...]

// 옵션 사용
const records2 = wb.toJSON("Sheet1", { headerRow: 2, range: "A2:C100" });
```

### `toCSV(sheet, options?)`

시트를 CSV 문자열로 변환합니다. 필요한 경우 값이 인용되며 쉼표로 구분됩니다.

```typescript
const csv = wb.toCSV("Sheet1");
// "Name,Age,City\nAlice,30,Seoul\n..."

// 사용자 정의 구분자 사용
const tsv = wb.toCSV("Sheet1", { separator: "\t" });
```

### `toHTML(sheet, options?)`

시트를 HTML `<table>` 문자열로 변환합니다. 모든 텍스트 내용은 XSS 안전합니다(HTML 이스케이프 처리됩니다).

```typescript
const html = wb.toHTML("Sheet1");
// "<table><thead><tr><th>Name</th>..."

// CSS 클래스 적용
const html2 = wb.toHTML("Sheet1", { tableClass: "data-table" });
```

### `fromJSON(sheet, data, options?)`

객체 배열을 시트에 씁니다. 키가 헤더 행이 되고 값이 데이터 행을 채웁니다.

```typescript
const wb = new Workbook();
wb.fromJSON("Sheet1", [
    { Name: "Alice", Age: 30, City: "Seoul" },
    { Name: "Bob", Age: 25, City: "Busan" },
]);
await wb.save("output.xlsx");

// 옵션 사용
wb.fromJSON("Sheet1", data, { startCell: "B2", writeHeaders: true });
```

### 변환 옵션

**ToJSONOptions:**

| 필드 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `headerRow` | `number` | `1` | 열 헤더로 사용할 행 번호 (1 기반) |
| `range` | `string?` | `undefined` | 특정 셀 범위로 제한 |

**ToCSVOptions:**

| 필드 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `separator` | `string` | `","` | 필드 구분자 |
| `lineEnding` | `string` | `"\n"` | 줄 바꿈 문자 |
| `escapeFormulas` | `boolean` | `false` | `=`, `+`, `-`, `@`로 시작하는 셀에 탭 문자를 접두사로 추가하여 수식 삽입 공격을 방지합니다 |

**ToHTMLOptions:**

| 필드 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `tableClass` | `string?` | `undefined` | `<table>` 요소의 CSS 클래스 |
| `includeHeaders` | `boolean` | `true` | `<thead>` 섹션 포함 여부 |

**FromJSONOptions:**

| 필드 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `startCell` | `string` | `"A1"` | 쓰기 시작할 왼쪽 상단 셀 |
| `writeHeaders` | `boolean` | `true` | 객체 키를 헤더 행으로 작성 |

---
