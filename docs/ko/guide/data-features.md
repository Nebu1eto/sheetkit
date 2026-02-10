### 시트 관리

시트를 생성, 삭제, 이름 변경, 복사하고 활성 시트를 설정합니다.

#### Rust

```rust
let mut wb = Workbook::new();

// Create new sheet (returns 0-based index)
let idx: usize = wb.new_sheet("Sales")?;

// Delete sheet
wb.delete_sheet("Sales")?;

// Rename sheet
wb.set_sheet_name("Sheet1", "Main")?;

// Copy sheet (returns index of new sheet)
let idx: usize = wb.copy_sheet("Main", "Main_Copy")?;

// Get sheet index (None if not found)
let idx: Option<usize> = wb.get_sheet_index("Main");

// Get/set active sheet
let active: &str = wb.get_active_sheet();
wb.set_active_sheet("Main")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Create new sheet (returns 0-based index)
const idx: number = wb.newSheet('Sales');

// Delete sheet
wb.deleteSheet('Sales');

// Rename sheet
wb.setSheetName('Sheet1', 'Main');

// Copy sheet (returns index of new sheet)
const copyIdx: number = wb.copySheet('Main', 'Main_Copy');

// Get sheet index (null if not found)
const sheetIdx: number | null = wb.getSheetIndex('Main');

// Get/set active sheet
const active: string = wb.getActiveSheet();
wb.setActiveSheet('Main');
```

---

### 행/열 조작

행과 열을 삽입, 삭제하고 크기 및 표시 여부를 설정합니다.

#### Rust

```rust
let mut wb = Workbook::new();

// -- Rows (1-based row numbers) --

// Insert 3 empty rows starting at row 2
wb.insert_rows("Sheet1", 2, 3)?;

// Delete row 5
wb.remove_row("Sheet1", 5)?;

// Duplicate row 1 (inserts copy below)
wb.duplicate_row("Sheet1", 1)?;

// Set/get row height
wb.set_row_height("Sheet1", 1, 25.0)?;
let height: Option<f64> = wb.get_row_height("Sheet1", 1)?;

// Show/hide row
wb.set_row_visible("Sheet1", 3, false)?;

// -- Columns (letter-based, e.g., "A", "B", "AA") --

// Set/get column width
wb.set_col_width("Sheet1", "A", 20.0)?;
let width: Option<f64> = wb.get_col_width("Sheet1", "A")?;

// Show/hide column
wb.set_col_visible("Sheet1", "B", false)?;

// Insert 2 empty columns starting at column "C"
wb.insert_cols("Sheet1", "C", 2)?;

// Delete column "D"
wb.remove_col("Sheet1", "D")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// -- Rows (1-based row numbers) --
wb.insertRows('Sheet1', 2, 3);
wb.removeRow('Sheet1', 5);
wb.duplicateRow('Sheet1', 1);
wb.setRowHeight('Sheet1', 1, 25);
const height: number | null = wb.getRowHeight('Sheet1', 1);
wb.setRowVisible('Sheet1', 3, false);

// -- Columns (letter-based) --
wb.setColWidth('Sheet1', 'A', 20);
const width: number | null = wb.getColWidth('Sheet1', 'A');
wb.setColVisible('Sheet1', 'B', false);
wb.insertCols('Sheet1', 'C', 2);
wb.removeCol('Sheet1', 'D');
```

---

### 스타일

스타일은 셀의 시각적 표현을 제어합니다. 스타일 정의를 등록하면 스타일 ID를 받고, 이 ID를 셀에 적용합니다. 동일한 스타일 정의는 자동으로 중복 제거됩니다.

`Style`은 글꼴, 채우기, 테두리, 정렬, 숫자 형식, 보호 속성을 자유롭게 조합할 수 있습니다.

#### Rust

```rust
use sheetkit::{
    AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle,
    FillStyle, FontStyle, HorizontalAlign, PatternType, Style,
    StyleColor, VerticalAlign, Workbook,
};

let mut wb = Workbook::new();

// Register style
let style_id = wb.add_style(&Style {
    font: Some(FontStyle {
        name: Some("Arial".into()),
        size: Some(14.0),
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        color: Some(StyleColor::Rgb("#FFFFFF".into())),
    }),
    fill: Some(FillStyle {
        pattern: PatternType::Solid,
        fg_color: Some(StyleColor::Rgb("#4472C4".into())),
        bg_color: None,
    }),
    border: Some(BorderStyle {
        bottom: Some(BorderSideStyle {
            style: BorderLineStyle::Thin,
            color: Some(StyleColor::Rgb("#000000".into())),
        }),
        ..Default::default()
    }),
    alignment: Some(AlignmentStyle {
        horizontal: Some(HorizontalAlign::Center),
        vertical: Some(VerticalAlign::Center),
        wrap_text: true,
        ..Default::default()
    }),
    ..Default::default()
})?;

// Apply style to cell
wb.set_cell_style("Sheet1", "A1", style_id)?;

// Get cell style ID (None if default style)
let current_style: Option<u32> = wb.get_cell_style("Sheet1", "A1")?;
```

#### TypeScript

```typescript
// Register style
const styleId = wb.addStyle({
    font: {
        name: 'Arial',
        size: 14,
        bold: true,
        color: '#FFFFFF',
    },
    fill: {
        pattern: 'solid',
        fgColor: '#4472C4',
    },
    border: {
        bottom: { style: 'thin', color: '#000000' },
    },
    alignment: {
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
    },
});

// Apply style to cell
wb.setCellStyle('Sheet1', 'A1', styleId);

// Get cell style ID (null if default style)
const currentStyle: number | null = wb.getCellStyle('Sheet1', 'A1');
```

#### 스타일 구성 요소 상세

**FontStyle (글꼴)**

| 필드             | Rust 타입           | TS 타입    | 설명                        |
|-----------------|---------------------|------------|----------------------------|
| `name`          | `Option<String>`    | `string?`  | 글꼴 이름 (예: "Calibri")   |
| `size`          | `Option<f64>`       | `number?`  | 글꼴 크기 (포인트)           |
| `bold`          | `bool`              | `boolean?` | 굵게                        |
| `italic`        | `bool`              | `boolean?` | 기울임꼴                     |
| `underline`     | `bool`              | `boolean?` | 밑줄                        |
| `strikethrough` | `bool`              | `boolean?` | 취소선                       |
| `color`         | `Option<StyleColor>` | `string?` | 글꼴 색상 (TS에서는 hex 문자열) |

**FillStyle (채우기)**

| 필드       | Rust 타입           | TS 타입   | 설명                        |
|-----------|---------------------|-----------|----------------------------|
| `pattern` | `PatternType`       | `string?` | 패턴 유형 (아래 값 참조)      |
| `fg_color`| `Option<StyleColor>` | `string?` | 전경 색상                   |
| `bg_color`| `Option<StyleColor>` | `string?` | 배경 색상                   |

PatternType 값: `None`, `Solid`, `Gray125`, `DarkGray`, `MediumGray`, `LightGray`.
TypeScript에서는 소문자 문자열 사용: `"none"`, `"solid"`, `"gray125"`, `"darkGray"`, `"mediumGray"`, `"lightGray"`.

**BorderStyle (테두리)**

각 면(`left`, `right`, `top`, `bottom`, `diagonal`)은 `BorderSideStyle`을 받으며 다음을 포함합니다:
- `style`: `Thin`, `Medium`, `Thick`, `Dashed`, `Dotted`, `Double`, `Hair`, `MediumDashed`, `DashDot`, `MediumDashDot`, `DashDotDot`, `MediumDashDotDot`, `SlantDashDot` 중 하나
- `color`: 선택적 색상

TypeScript에서는 소문자 문자열 사용: `"thin"`, `"medium"`, `"thick"` 등.

**AlignmentStyle (정렬)**

| 필드             | Rust 타입                 | TS 타입    | 설명                  |
|-----------------|--------------------------|------------|----------------------|
| `horizontal`    | `Option<HorizontalAlign>` | `string?` | 가로 정렬             |
| `vertical`      | `Option<VerticalAlign>`   | `string?` | 세로 정렬             |
| `wrap_text`     | `bool`                    | `boolean?`| 텍스트 줄 바꿈        |
| `text_rotation` | `Option<u32>`             | `number?` | 텍스트 회전 각도(도)   |
| `indent`        | `Option<u32>`             | `number?` | 들여쓰기 수준          |
| `shrink_to_fit` | `bool`                    | `boolean?`| 셀에 맞게 텍스트 축소  |

HorizontalAlign 값: `General`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterContinuous`, `Distributed`.
VerticalAlign 값: `Top`, `Center`, `Bottom`, `Justify`, `Distributed`.

**NumFmtStyle (숫자 형식, Rust 전용)**

```rust
use sheetkit::style::NumFmtStyle;

// Built-in format (e.g., percent, date, currency)
NumFmtStyle::Builtin(9)  // 0%

// Custom format string
NumFmtStyle::Custom("#,##0.00".to_string())
```

TypeScript에서는 스타일 객체의 `numFmtId`(기본 제공 형식 ID) 또는 `customNumFmt`(사용자 정의 형식 문자열)을 사용합니다.

**ProtectionStyle (보호)**

| 필드     | Rust 타입 | TS 타입    | 설명                              |
|---------|-----------|------------|----------------------------------|
| `locked`| `bool`    | `boolean?` | 셀 잠금 (기본값: true)             |
| `hidden`| `bool`    | `boolean?` | 보호 모드에서 수식 숨김             |

---

### 차트

워크시트에 차트를 추가합니다. 차트는 두 셀(좌상단, 우하단) 사이에 앵커링되어 지정된 셀 범위의 데이터를 시각화합니다.

#### 지원 차트 유형 (43종)

**세로 막대 (Column)**

| Rust 변형 | TS 문자열 | 설명 |
|-----------|----------|------|
| `ChartType::Col` | `"col"` | 세로 막대 |
| `ChartType::ColStacked` | `"colStacked"` | 누적 세로 막대 |
| `ChartType::ColPercentStacked` | `"colPercentStacked"` | 100% 누적 세로 막대 |
| `ChartType::Col3D` | `"col3D"` | 3D 세로 막대 |
| `ChartType::Col3DStacked` | `"col3DStacked"` | 3D 누적 세로 막대 |
| `ChartType::Col3DPercentStacked` | `"col3DPercentStacked"` | 3D 100% 누적 세로 막대 |

**가로 막대 (Bar)**

| Rust 변형 | TS 문자열 | 설명 |
|-----------|----------|------|
| `ChartType::Bar` | `"bar"` | 가로 막대 |
| `ChartType::BarStacked` | `"barStacked"` | 누적 가로 막대 |
| `ChartType::BarPercentStacked` | `"barPercentStacked"` | 100% 누적 가로 막대 |
| `ChartType::Bar3D` | `"bar3D"` | 3D 가로 막대 |
| `ChartType::Bar3DStacked` | `"bar3DStacked"` | 3D 누적 가로 막대 |
| `ChartType::Bar3DPercentStacked` | `"bar3DPercentStacked"` | 3D 100% 누적 가로 막대 |

**꺾은선 (Line)**

| Rust 변형 | TS 문자열 | 설명 |
|-----------|----------|------|
| `ChartType::Line` | `"line"` | 꺾은선 |
| `ChartType::LineStacked` | `"lineStacked"` | 누적 꺾은선 |
| `ChartType::LinePercentStacked` | `"linePercentStacked"` | 100% 누적 꺾은선 |
| `ChartType::Line3D` | `"line3D"` | 3D 꺾은선 |

**원형 (Pie)**

| Rust 변형 | TS 문자열 | 설명 |
|-----------|----------|------|
| `ChartType::Pie` | `"pie"` | 원형 |
| `ChartType::Pie3D` | `"pie3D"` | 3D 원형 |
| `ChartType::Doughnut` | `"doughnut"` | 도넛형 |

**영역 (Area)**

| Rust 변형 | TS 문자열 | 설명 |
|-----------|----------|------|
| `ChartType::Area` | `"area"` | 영역 |
| `ChartType::AreaStacked` | `"areaStacked"` | 누적 영역 |
| `ChartType::AreaPercentStacked` | `"areaPercentStacked"` | 100% 누적 영역 |
| `ChartType::Area3D` | `"area3D"` | 3D 영역 |
| `ChartType::Area3DStacked` | `"area3DStacked"` | 3D 누적 영역 |
| `ChartType::Area3DPercentStacked` | `"area3DPercentStacked"` | 3D 100% 누적 영역 |

**분산형 (Scatter)**

| Rust 변형 | TS 문자열 | 설명 |
|-----------|----------|------|
| `ChartType::Scatter` | `"scatter"` | 분산형 (표식만) |
| `ChartType::ScatterSmooth` | `"scatterSmooth"` | 부드러운 선 |
| `ChartType::ScatterLine` | `"scatterStraight"` | 직선 |

**방사형 (Radar)**

| Rust 변형 | TS 문자열 | 설명 |
|-----------|----------|------|
| `ChartType::Radar` | `"radar"` | 방사형 |
| `ChartType::RadarFilled` | `"radarFilled"` | 채워진 방사형 |
| `ChartType::RadarMarker` | `"radarMarker"` | 표식이 있는 방사형 |

**주식 (Stock)**

| Rust 변형 | TS 문자열 | 설명 |
|-----------|----------|------|
| `ChartType::StockHLC` | `"stockHLC"` | 고가-저가-종가 |
| `ChartType::StockOHLC` | `"stockOHLC"` | 시가-고가-저가-종가 |
| `ChartType::StockVHLC` | `"stockVHLC"` | 거래량-고가-저가-종가 |
| `ChartType::StockVOHLC` | `"stockVOHLC"` | 거래량-시가-고가-저가-종가 |

**기타**

| Rust 변형 | TS 문자열 | 설명 |
|-----------|----------|------|
| `ChartType::Bubble` | `"bubble"` | 거품형 |
| `ChartType::Surface` | `"surface"` | 표면형 |
| `ChartType::Surface3D` | `"surfaceTop"` | 3D 표면형 |
| `ChartType::SurfaceWireframe` | `"surfaceWireframe"` | 와이어프레임 표면형 |
| `ChartType::SurfaceWireframe3D` | `"surfaceTopWireframe"` | 3D 와이어프레임 표면형 |

**콤보 (Combo)**

| Rust 변형 | TS 문자열 | 설명 |
|-----------|----------|------|
| `ChartType::ColLine` | `"colLine"` | 세로 막대 + 꺾은선 |
| `ChartType::ColLineStacked` | `"colLineStacked"` | 누적 세로 막대 + 꺾은선 |
| `ChartType::ColLinePercentStacked` | `"colLinePercentStacked"` | 100% 누적 세로 막대 + 꺾은선 |

#### Rust

```rust
use sheetkit::{ChartConfig, ChartSeries, ChartType, Workbook};

let mut wb = Workbook::new();

// First, populate data...
wb.set_cell_value("Sheet1", "A1", CellValue::String("Q1".into()))?;
wb.set_cell_value("Sheet1", "B1", CellValue::Number(1500.0))?;
// ... more data rows ...

// Add chart in D1:K15 area
wb.add_chart(
    "Sheet1",
    "D1",   // Top-left anchor cell
    "K15",  // Bottom-right anchor cell
    &ChartConfig {
        chart_type: ChartType::Col,
        title: Some("Quarterly Revenue".into()),
        series: vec![ChartSeries {
            name: "Revenue".into(),
            categories: "Sheet1!$A$1:$A$4".into(),
            values: "Sheet1!$B$1:$B$4".into(),
        }],
        show_legend: true,
    },
)?;
```

#### TypeScript

```typescript
wb.addChart('Sheet1', 'D1', 'K15', {
    chartType: 'col',
    title: 'Quarterly Revenue',
    series: [
        {
            name: 'Revenue',
            categories: 'Sheet1!$A$1:$A$4',
            values: 'Sheet1!$B$1:$B$4',
        },
    ],
    showLegend: true,
});
```

---

### 이미지

워크시트에 이미지를 삽입합니다. 11가지 형식을 지원합니다: PNG, JPEG, GIF, BMP, ICO, TIFF, SVG, EMF, EMZ, WMF, WMZ. 이미지는 셀에 앵커링되며 픽셀 크기로 지정합니다.

#### Rust

```rust
use sheetkit::{ImageConfig, ImageFormat, Workbook};

let mut wb = Workbook::new();

let image_bytes = std::fs::read("logo.png").unwrap();

wb.add_image(
    "Sheet1",
    &ImageConfig {
        data: image_bytes,
        format: ImageFormat::Png,
        from_cell: "B2".into(),
        width_px: 200,
        height_px: 100,
    },
)?;
```

#### TypeScript

```typescript
import { readFileSync } from 'fs';

const imageData = readFileSync('logo.png');

wb.addImage('Sheet1', {
    data: imageData,
    format: 'png',        // "png" | "jpeg" | "jpg" | "gif" | "bmp" | "ico" | "tiff" | "tif" | "svg" | "emf" | "emz" | "wmf" | "wmz"
    fromCell: 'B2',
    widthPx: 200,
    heightPx: 100,
});
```

---

### 데이터 유효성 검사

셀 범위에 데이터 유효성 검사 규칙을 추가합니다. 이 규칙은 사용자가 해당 셀에 입력할 수 있는 값을 제한합니다.

#### 유효성 검사 유형

| Rust 변형                   | TS 문자열       | 설명                         |
|----------------------------|----------------|------------------------------|
| `ValidationType::None`     | `"none"`       | 제한 없음 (프롬프트/메시지만) |
| `ValidationType::Whole`    | `"whole"`      | 정수 제약                     |
| `ValidationType::Decimal`  | `"decimal"`    | 소수 제약                     |
| `ValidationType::List`     | `"list"`       | 드롭다운 목록                 |
| `ValidationType::Date`     | `"date"`       | 날짜 제약                     |
| `ValidationType::Time`     | `"time"`       | 시간 제약                     |
| `ValidationType::TextLength`| `"textLength"` | 텍스트 길이 제약              |
| `ValidationType::Custom`   | `"custom"`     | 사용자 정의 수식 제약          |

#### 비교 연산자

`Between`, `NotBetween`, `Equal`, `NotEqual`, `LessThan`, `LessThanOrEqual`, `GreaterThan`, `GreaterThanOrEqual`.

TypeScript 입력은 대소문자를 구분하지 않으며, 출력은 OOXML 규격에 맞는 camelCase를 사용: `"between"`, `"notBetween"`, `"lessThan"` 등.

`sqref`는 유효한 셀 범위 참조여야 합니다. `none` 이외의 타입에는 `formula1`이 필수이며, `between`/`notBetween` 연산자에는 `formula2`도 필수입니다.

#### 오류 스타일

`Stop`, `Warning`, `Information` -- 잘못된 입력 시 표시되는 오류 대화 상자의 심각도를 제어합니다.

#### Rust

```rust
use sheetkit::{DataValidationConfig, ErrorStyle, ValidationType, Workbook};

let mut wb = Workbook::new();

// Dropdown list validation
wb.add_data_validation(
    "Sheet1",
    &DataValidationConfig {
        sqref: "C2:C100".into(),
        validation_type: ValidationType::List,
        operator: None,
        formula1: Some("\"Achieved,Not Achieved,In Progress\"".into()),
        formula2: None,
        allow_blank: true,
        show_input_message: true,
        prompt_title: Some("Select Status".into()),
        prompt_message: Some("Choose from the dropdown".into()),
        show_error_message: true,
        error_style: Some(ErrorStyle::Stop),
        error_title: Some("Invalid".into()),
        error_message: Some("Please select from the list".into()),
    },
)?;

// Get all validations for the sheet
let validations = wb.get_data_validations("Sheet1")?;

// Remove validation by cell range reference
wb.remove_data_validation("Sheet1", "C2:C100")?;
```

#### TypeScript

```typescript
// Dropdown list validation
wb.addDataValidation('Sheet1', {
    sqref: 'C2:C100',
    validationType: 'list',
    formula1: '"Achieved,Not Achieved,In Progress"',
    allowBlank: true,
    showInputMessage: true,
    promptTitle: 'Select Status',
    promptMessage: 'Choose from the dropdown',
    showErrorMessage: true,
    errorStyle: 'stop',
    errorTitle: 'Invalid',
    errorMessage: 'Please select from the list',
});

// Get all validations for the sheet
const validations = wb.getDataValidations('Sheet1');

// Remove validation by cell range reference
wb.removeDataValidation('Sheet1', 'C2:C100');
```

---

### 코멘트

셀 코멘트를 추가, 조회, 삭제합니다.

#### Rust

```rust
use sheetkit::{CommentConfig, Workbook};

let mut wb = Workbook::new();

// Add comment
wb.add_comment(
    "Sheet1",
    &CommentConfig {
        cell: "A1".into(),
        author: "Admin".into(),
        text: "This cell contains the project name.".into(),
    },
)?;

// Get all comments for the sheet
let comments: Vec<CommentConfig> = wb.get_comments("Sheet1")?;

// Delete comment for a specific cell
wb.remove_comment("Sheet1", "A1")?;
```

#### TypeScript

```typescript
// Add comment
wb.addComment('Sheet1', {
    cell: 'A1',
    author: 'Admin',
    text: 'This cell contains the project name.',
});

// Get all comments for the sheet
const comments = wb.getComments('Sheet1');

// Delete comment for a specific cell
wb.removeComment('Sheet1', 'A1');
```

---

### 자동 필터

열 범위에 자동 필터 드롭다운을 적용하거나 제거합니다.

#### Rust

```rust
// Set auto filter on range
wb.set_auto_filter("Sheet1", "A1:D100")?;

// Remove auto filter
wb.remove_auto_filter("Sheet1")?;
```

#### TypeScript

```typescript
// Set auto filter on range
wb.setAutoFilter('Sheet1', 'A1:D100');

// Remove auto filter
wb.removeAutoFilter('Sheet1');
```

---

### 테이블

테이블은 열 헤더, 스타일, 선택적 자동 필터가 포함된 구조화된 데이터 범위입니다. 테이블은 별도의 OOXML 파트로 저장되며 저장/열기 라운드트립을 완벽하게 지원합니다.

#### Rust

```rust
use sheetkit::table::{TableConfig, TableColumn};

let mut wb = Workbook::new();

// Create table
let config = TableConfig {
    name: "EmployeeList".to_string(),
    display_name: "EmployeeList".to_string(),
    range: "A1:C5".to_string(),
    columns: vec![
        TableColumn { name: "Name".to_string(), totals_row_function: None, totals_row_label: None },
        TableColumn { name: "Department".to_string(), totals_row_function: None, totals_row_label: None },
        TableColumn { name: "Salary".to_string(), totals_row_function: None, totals_row_label: None },
    ],
    style_name: Some("TableStyleMedium2".to_string()),
    ..TableConfig::default()
};
wb.add_table("Sheet1", &config)?;

// List tables
let tables = wb.get_tables("Sheet1")?;

// Delete table
wb.delete_table("Sheet1", "EmployeeList")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Create table
wb.addTable("Sheet1", {
    name: "EmployeeList",
    displayName: "EmployeeList",
    range: "A1:C5",
    columns: [
        { name: "Name" },
        { name: "Department" },
        { name: "Salary" },
    ],
    styleName: "TableStyleMedium2",
});

// List tables
const tables = wb.getTables("Sheet1");

// Delete table
wb.deleteTable("Sheet1", "EmployeeList");
```

자세한 API 설명은 [API 레퍼런스](../api-reference/data-features.md#16-테이블)를 참조하세요.

---

### 데이터 변환 유틸리티 (Node.js 전용)

SheetKit은 시트 데이터와 JSON, CSV, HTML 등 일반적인 형식 간의 변환을 위한 편의 메서드를 제공합니다. TypeScript/Node.js 바인딩에서만 사용할 수 있습니다.

#### JSON 변환

```typescript
import { Workbook } from '@sheetkit/node';

// Read sheet as array of objects
const wb = await Workbook.open("data.xlsx");
const records = wb.toJSON("Sheet1");
// [{ Name: "Alice", Age: 30 }, { Name: "Bob", Age: 25 }, ...]

// Write array of objects to sheet
const wb2 = new Workbook();
wb2.fromJSON("Sheet1", [
    { Name: "Alice", Age: 30, City: "Seoul" },
    { Name: "Bob", Age: 25, City: "Busan" },
]);
await wb2.save("output.xlsx");
```

#### CSV 내보내기

```typescript
const csv = wb.toCSV("Sheet1");
// "Name,Age,City\nAlice,30,Seoul\n..."

// Tab-separated values
const tsv = wb.toCSV("Sheet1", { separator: "\t" });
```

#### HTML 내보내기

```typescript
const html = wb.toHTML("Sheet1");
// "<table><thead><tr><th>Name</th>..."

const html2 = wb.toHTML("Sheet1", { tableClass: "data-table" });
```

자세한 API 설명은 [API 레퍼런스](../api-reference/data-features.md#17-데이터-변환-유틸리티-nodejs-전용)를 참조하세요.

---
