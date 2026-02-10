### StreamWriter

StreamWriter는 대용량 시트를 효율적으로 작성하기 위한 순방향 전용 스트리밍 API입니다. 워크시트 데이터 구조를 메모리에 직접 빌드하며, `apply_stream_writer()`를 통해 워크북에 적용할 때 XML 직렬화/역직렬화 없이 데이터를 직접 전달합니다.

행은 반드시 오름차순으로 작성해야 합니다. 열 너비는 행을 쓰기 전에 설정해야 합니다.

#### Rust

```rust
use sheetkit::{CellValue, Workbook};

let mut wb = Workbook::new();

// Create a stream writer for a new sheet
let mut sw = wb.new_stream_writer("LargeSheet")?;

// Set column widths (must be done before writing rows)
sw.set_col_width(1, 20.0)?;     // Column 1 (A)
sw.set_col_width(2, 15.0)?;     // Column 2 (B)

// Write rows in ascending order (starting from 1)
sw.write_row(1, &[
    CellValue::from("Name"),
    CellValue::from("Score"),
])?;
for i in 2..=10_000 {
    sw.write_row(i, &[
        CellValue::from(format!("User_{}", i - 1)),
        CellValue::from(i as f64 * 1.5),
    ])?;
}

// Add merge cell (optional)
sw.add_merge_cell("A1:B1")?;

// Apply stream writer to the workbook
wb.apply_stream_writer(sw)?;

wb.save("large_file.xlsx")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Create a stream writer for a new sheet
const sw = wb.newStreamWriter('LargeSheet');

// Set column widths (must be done before writing rows)
sw.setColWidth(1, 20);     // Column 1 (A)
sw.setColWidth(2, 15);     // Column 2 (B)

// Write rows in ascending order (starting from 1)
sw.writeRow(1, ['Name', 'Score']);
for (let i = 2; i <= 10000; i++) {
    sw.writeRow(i, [`User_${i - 1}`, i * 1.5]);
}

// Add merge cell (optional)
sw.addMergeCell('A1:B1');

// Apply stream writer to the workbook
wb.applyStreamWriter(sw);

await wb.save('large_file.xlsx');
```

#### StreamWriter API 요약

| 메서드                 | 설명                                        |
|-----------------------|---------------------------------------------|
| `set_col_width`       | 단일 열 너비 설정 (1부터 시작하는 열 번호)     |
| `set_col_width_range` | 열 범위의 너비 설정 (Rust 전용)               |
| `write_row`           | 지정한 행 번호에 값 배열 작성                  |
| `add_merge_cell`      | 셀 병합 참조 추가 (예: `"A1:C3"`)            |

#### 성능 참고

StreamWriter는 대규모 쓰기에 최적화되어 있습니다. 내부적으로 힙 할당 없는 셀 참조(`CompactCellRef`)를 사용하여 `Row`와 `Cell` 구조체를 직접 빌드하며, 문자열 기반 XML 생성을 피합니다. `apply_stream_writer()`를 통해 적용할 때 XML 직렬화 후 다시 파싱하는 과정 없이 데이터가 워크북으로 직접 전달되어, 이전에 스트리밍 쓰기의 주요 병목이었던 부분을 제거합니다.

50,000행 x 20열 기준으로, 이 최적화는 스트리밍 쓰기 시간을 약 2배 단축하고 최대 메모리 사용량을 크게 줄입니다.

---

### 문서 속성

문서 메타데이터를 설정하고 읽습니다: 핵심 속성(제목, 작성자 등), 애플리케이션 속성, 사용자 정의 속성.

#### Rust

```rust
use sheetkit::{AppProperties, CustomPropertyValue, DocProperties, Workbook};

let mut wb = Workbook::new();

// Core document properties
wb.set_doc_props(DocProperties {
    title: Some("Annual Report".into()),
    creator: Some("SheetKit".into()),
    description: Some("Financial data for 2025".into()),
    ..Default::default()
});
let props = wb.get_doc_props();

// Application properties
wb.set_app_props(AppProperties {
    application: Some("SheetKit".into()),
    company: Some("Acme Corp".into()),
    ..Default::default()
});
let app_props = wb.get_app_props();

// Custom properties (string, int, float, boolean, datetime)
wb.set_custom_property("Project", CustomPropertyValue::String("SheetKit".into()));
wb.set_custom_property("Version", CustomPropertyValue::Int(1));
wb.set_custom_property("Released", CustomPropertyValue::Bool(false));

let val = wb.get_custom_property("Project");
let deleted = wb.delete_custom_property("Version");
```

#### TypeScript

```typescript
// Core document properties
wb.setDocProps({
    title: 'Annual Report',
    creator: 'SheetKit',
    description: 'Financial data for 2025',
});
const props = wb.getDocProps();

// Application properties
wb.setAppProps({
    application: 'SheetKit',
    company: 'Acme Corp',
});
const appProps = wb.getAppProps();

// Custom properties (string, number, boolean)
wb.setCustomProperty('Project', 'SheetKit');
wb.setCustomProperty('Version', 1);
wb.setCustomProperty('Released', false);

const val = wb.getCustomProperty('Project');       // string | number | boolean | null
const deleted: boolean = wb.deleteCustomProperty('Version');
```

#### DocProperties 필드

| 필드                | 타입              | 설명              |
|--------------------|-------------------|-------------------|
| `title`            | `Option<String>`  | 문서 제목          |
| `subject`          | `Option<String>`  | 문서 주제          |
| `creator`          | `Option<String>`  | 작성자 이름        |
| `keywords`         | `Option<String>`  | 검색용 키워드       |
| `description`      | `Option<String>`  | 문서 설명          |
| `last_modified_by` | `Option<String>`  | 마지막 편집자       |
| `revision`         | `Option<String>`  | 수정 번호          |
| `created`          | `Option<String>`  | 생성 일시          |
| `modified`         | `Option<String>`  | 최종 수정 일시      |
| `category`         | `Option<String>`  | 분류              |
| `content_status`   | `Option<String>`  | 콘텐츠 상태        |

#### AppProperties 필드

| 필드            | 타입              | 설명              |
|----------------|-------------------|-------------------|
| `application`  | `Option<String>`  | 애플리케이션 이름   |
| `doc_security` | `Option<u32>`     | 문서 보안 수준      |
| `company`      | `Option<String>`  | 회사 이름          |
| `app_version`  | `Option<String>`  | 앱 버전            |
| `manager`      | `Option<String>`  | 관리자 이름        |
| `template`     | `Option<String>`  | 템플릿 이름        |

---

### 워크북 보호

워크북 구조를 보호하여 사용자가 시트를 추가, 삭제, 이름 변경하는 것을 방지합니다. 선택적으로 비밀번호를 설정할 수 있습니다 (레거시 Excel 해시 -- 암호학적으로 안전하지 않음).

#### Rust

```rust
use sheetkit::{Workbook, WorkbookProtectionConfig};

let mut wb = Workbook::new();

// Protect workbook
wb.protect_workbook(WorkbookProtectionConfig {
    password: Some("secret".into()),
    lock_structure: true,    // Prevent adding/deleting/renaming sheets
    lock_windows: false,     // Allow window resizing
    lock_revision: false,    // Allow revision tracking changes
});

// Check protection status
let is_protected: bool = wb.is_workbook_protected();

// Remove protection
wb.unprotect_workbook();
```

#### TypeScript

```typescript
// Protect workbook
wb.protectWorkbook({
    password: 'secret',
    lockStructure: true,
    lockWindows: false,
    lockRevision: false,
});

// Check protection status
const isProtected: boolean = wb.isWorkbookProtected();

// Remove protection
wb.unprotectWorkbook();
```

---

### 셀 병합

여러 셀을 하나로 병합하거나 해제합니다. 병합된 셀의 값은 좌상단 셀에 저장됩니다.

#### Rust

```rust
let mut wb = Workbook::new();

wb.set_cell_value("Sheet1", "A1", CellValue::String("Header".into()))?;
wb.merge_cells("Sheet1", "A1", "C1")?;

let merged: Vec<String> = wb.get_merge_cells("Sheet1")?;
wb.unmerge_cell("Sheet1", "A1:C1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

wb.setCellValue('Sheet1', 'A1', 'Header');
wb.mergeCells('Sheet1', 'A1', 'C1');

const merged: string[] = wb.getMergeCells('Sheet1');
wb.unmergeCell('Sheet1', 'A1:C1');
```

---

### 하이퍼링크

셀에 하이퍼링크를 설정합니다. 외부 URL, 이메일, 내부 시트 참조의 세 가지 유형을 지원합니다.

#### Rust

```rust
use sheetkit::hyperlink::HyperlinkType;

let mut wb = Workbook::new();

// External URL
wb.set_cell_hyperlink(
    "Sheet1", "A1",
    HyperlinkType::External("https://example.com".into()),
    Some("Example Site"), Some("Click here"),
)?;

// Email
wb.set_cell_hyperlink(
    "Sheet1", "A2",
    HyperlinkType::Email("mailto:user@example.com".into()),
    Some("Send email"), None,
)?;

// Internal sheet reference
wb.set_cell_hyperlink(
    "Sheet1", "A3",
    HyperlinkType::Internal("Sheet2!A1".into()),
    None, None,
)?;

// Get and delete
let info = wb.get_cell_hyperlink("Sheet1", "A1")?;
wb.delete_cell_hyperlink("Sheet1", "A1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// External URL
wb.setCellHyperlink('Sheet1', 'A1', {
    linkType: 'external',
    target: 'https://example.com',
    display: 'Example Site',
    tooltip: 'Click here',
});

// Email
wb.setCellHyperlink('Sheet1', 'A2', {
    linkType: 'email',
    target: 'mailto:user@example.com',
    display: 'Send email',
});

// Internal sheet reference
wb.setCellHyperlink('Sheet1', 'A3', {
    linkType: 'internal',
    target: 'Sheet2!A1',
});

// Get and delete
const info = wb.getCellHyperlink('Sheet1', 'A1');
wb.deleteCellHyperlink('Sheet1', 'A1');
```

---

### 조건부 서식

셀 값이나 수식에 따라 자동으로 서식을 적용합니다. 17가지 규칙 유형을 지원합니다.

#### cellIs (셀 값 비교)

##### Rust

```rust
use sheetkit::conditional::*;
use sheetkit::style::*;

let mut wb = Workbook::new();

wb.set_conditional_format("Sheet1", "A1:A100", &[
    ConditionalFormatRule {
        rule_type: ConditionalFormatType::CellIs {
            operator: CfOperator::GreaterThan,
            formula: "90".into(),
            formula2: None,
        },
        format: Some(ConditionalStyle {
            font: Some(FontStyle { bold: true, color: Some(StyleColor::Rgb("#006100".into())), ..Default::default() }),
            fill: Some(FillStyle { pattern: PatternType::Solid, fg_color: Some(StyleColor::Rgb("#C6EFCE".into())), bg_color: None }),
            border: None,
            num_fmt: None,
        }),
        priority: Some(1),
        stop_if_true: false,
    },
])?;
```

##### TypeScript

```typescript
wb.setConditionalFormat('Sheet1', 'A1:A100', [
    {
        ruleType: 'cellIs',
        operator: 'greaterThan',
        formula: '90',
        format: {
            font: { bold: true, color: '#006100' },
            fill: { pattern: 'solid', fgColor: '#C6EFCE' },
        },
        priority: 1,
    },
]);
```

#### colorScale (색상 스케일)

```typescript
wb.setConditionalFormat('Sheet1', 'B1:B50', [
    {
        ruleType: 'colorScale',
        minType: 'min',
        minColor: 'FFF8696B',
        midType: 'percentile',
        midValue: '50',
        midColor: 'FFFFEB84',
        maxType: 'max',
        maxColor: 'FF63BE7B',
    },
]);
```

#### dataBar (데이터 막대)

```typescript
wb.setConditionalFormat('Sheet1', 'C1:C50', [
    { ruleType: 'dataBar', barColor: 'FF638EC6', showValue: true },
]);
```

#### 조회 및 삭제

```typescript
const formats = wb.getConditionalFormats('Sheet1');
wb.deleteConditionalFormat('Sheet1', 'A1:A100');
```

---

### 틀 고정/분할

특정 행이나 열을 고정하여 스크롤 시에도 항상 보이게 합니다. 셀 참조는 스크롤 가능 영역의 좌상단 셀입니다.

#### Rust

```rust
let mut wb = Workbook::new();

wb.set_panes("Sheet1", "A2")?;    // Freeze first row
wb.set_panes("Sheet1", "B1")?;    // Freeze first column
wb.set_panes("Sheet1", "B2")?;    // Freeze first row + first column

let pane = wb.get_panes("Sheet1")?;
wb.unset_panes("Sheet1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

wb.setPanes('Sheet1', 'A2');       // Freeze first row
wb.setPanes('Sheet1', 'B2');       // Freeze first row + first column

const pane = wb.getPanes('Sheet1');
wb.unsetPanes('Sheet1');
```

---

### 페이지 레이아웃

인쇄 관련 설정을 다룹니다. 여백, 용지 크기, 방향, 인쇄 옵션, 머리글/바닥글, 페이지 나누기를 포함합니다.

#### Rust

```rust
use sheetkit::page_layout::*;

let mut wb = Workbook::new();

// Margins (in inches)
wb.set_page_margins("Sheet1", &PageMarginsConfig {
    left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3,
})?;

// Page setup
wb.set_page_setup("Sheet1", Some(Orientation::Landscape), Some(PaperSize::A4), Some(100), None, None)?;

// Header/footer
wb.set_header_footer("Sheet1", Some("&CMonthly Report"), Some("&LPage &P of &N"))?;

// Page break
wb.insert_page_break("Sheet1", 20)?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Margins (in inches)
wb.setPageMargins('Sheet1', {
    left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3,
});

// Page setup
wb.setPageSetup('Sheet1', {
    paperSize: 'a4', orientation: 'landscape', scale: 100,
});

// Print options
wb.setPrintOptions('Sheet1', { gridLines: true, horizontalCentered: true });

// Header/footer
wb.setHeaderFooter('Sheet1', '&CMonthly Report', '&LPage &P of &N');

// Page break
wb.insertPageBreak('Sheet1', 20);
```

---

### 행/열 이터레이터

시트의 모든 행 또는 열 데이터를 한 번에 조회합니다. 데이터가 있는 행/열만 포함됩니다.

#### Rust

```rust
let wb = Workbook::open("data.xlsx")?;

// Get all rows
let rows = wb.get_rows("Sheet1")?;
for (row_num, cells) in &rows {
    for (col, val) in cells {
        println!("Row {}, Col {}: {:?}", row_num, col, val);
    }
}

// Get all columns
let cols = wb.get_cols("Sheet1")?;
```

#### TypeScript

```typescript
const wb = await Workbook.open('data.xlsx');

const rows = wb.getRows('Sheet1');
for (const row of rows) {
    for (const cell of row.cells) {
        console.log(`Row ${row.row}, ${cell.column}: ${cell.value}`);
    }
}

const cols = wb.getCols('Sheet1');
```

---

### 행/열 아웃라인 및 스타일

행과 열의 아웃라인(그룹) 수준(0-7)을 설정하고, 행/열 전체에 스타일을 적용합니다.

#### Rust

```rust
let mut wb = Workbook::new();

// Outline level
wb.set_row_outline_level("Sheet1", 2, 1)?;
let level: u8 = wb.get_row_outline_level("Sheet1", 2)?;

wb.set_col_outline_level("Sheet1", "B", 2)?;
let col_level: u8 = wb.get_col_outline_level("Sheet1", "B")?;

// Row/column style
let style_id = wb.add_style(&style)?;
wb.set_row_style("Sheet1", 1, style_id)?;
wb.set_col_style("Sheet1", "A", style_id)?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Outline level
wb.setRowOutlineLevel('Sheet1', 2, 1);
const level: number = wb.getRowOutlineLevel('Sheet1', 2);

wb.setColOutlineLevel('Sheet1', 'B', 2);
const colLevel: number = wb.getColOutlineLevel('Sheet1', 'B');

// Row/column style
const styleId = wb.addStyle({ font: { bold: true } });
wb.setRowStyle('Sheet1', 1, styleId);
wb.setColStyle('Sheet1', 'A', styleId);
```

---

### 수식 계산

셀 수식을 평가합니다. `evaluate_formula`는 단일 수식을 계산하고, `calculate_all`은 워크북의 모든 수식 셀을 의존성 순서대로 재계산합니다. 110개 이상의 함수를 지원합니다 (SUM, VLOOKUP, IF, DATE 등).

#### Rust

```rust
let mut wb = Workbook::new();

wb.set_cell_value("Sheet1", "A1", CellValue::Number(10.0))?;
wb.set_cell_value("Sheet1", "A2", CellValue::Number(20.0))?;
wb.set_cell_value("Sheet1", "A3", CellValue::Formula {
    expr: "SUM(A1:A2)".into(),
    result: None,
})?;

let result = wb.evaluate_formula("Sheet1", "SUM(A1:A2)")?;
wb.calculate_all()?;
```

#### TypeScript

```typescript
const wb = new Workbook();

wb.setCellValue('Sheet1', 'A1', 10);
wb.setCellValue('Sheet1', 'A2', 20);

const result = wb.evaluateFormula('Sheet1', 'SUM(A1:A2)');
wb.calculateAll();
```

---

### 피벗 테이블

소스 데이터 범위로부터 행/열/데이터 필드를 지정하여 피벗 테이블을 생성합니다.

#### Rust

```rust
use sheetkit::pivot::*;

let mut wb = Workbook::new();
wb.new_sheet("PivotSheet")?;

wb.add_pivot_table(&PivotTableConfig {
    name: "SalesPivot".into(),
    source_sheet: "Sheet1".into(),
    source_range: "A1:D100".into(),
    target_sheet: "PivotSheet".into(),
    target_cell: "A3".into(),
    rows: vec![PivotField { name: "Region".into() }],
    columns: vec![PivotField { name: "Quarter".into() }],
    data: vec![PivotDataField {
        name: "Revenue".into(),
        function: AggregateFunction::Sum,
        display_name: Some("Total Revenue".into()),
    }],
})?;

let tables = wb.get_pivot_tables();
wb.delete_pivot_table("SalesPivot")?;
```

#### TypeScript

```typescript
const wb = new Workbook();
wb.newSheet('PivotSheet');

wb.addPivotTable({
    name: 'SalesPivot',
    sourceSheet: 'Sheet1',
    sourceRange: 'A1:D100',
    targetSheet: 'PivotSheet',
    targetCell: 'A3',
    rows: [{ name: 'Region' }],
    columns: [{ name: 'Quarter' }],
    data: [{ name: 'Revenue', function: 'sum', displayName: 'Total Revenue' }],
});

const tables = wb.getPivotTables();
wb.deletePivotTable('SalesPivot');
```

#### 집계 함수

`sum`, `count`, `average`, `max`, `min`, `product`, `countNums`

---

### 파일 암호화

SheetKit은 ECMA-376 표준에 따른 .xlsx 파일의 파일 수준 암호화를 지원합니다. 암호화된 파일은 일반 ZIP 아카이브가 아닌 OLE/CFB 복합 컨테이너에 저장됩니다.

- **읽기**: Standard Encryption (Office 2007, AES-128-ECB)과 Agile Encryption (Office 2010+, AES-256-CBC) 모두 지원
- **쓰기**: Agile Encryption (AES-256-CBC + SHA-512, 100,000회 반복) 사용

> 파일 암호화는 워크북/시트 보호와 다릅니다. 암호화는 올바른 비밀번호 없이는 파일 자체를 열 수 없게 하지만, 보호는 편집 작업만 제한합니다.

#### Rust

```rust
use sheetkit::Workbook;

let mut wb = Workbook::new();
wb.set_cell_value("Sheet1", "A1", CellValue::from("Confidential"))?;

// Save with password (Agile Encryption)
wb.save_with_password("encrypted.xlsx", "mypassword")?;

// Open encrypted file
let wb2 = Workbook::open_with_password("encrypted.xlsx", "mypassword")?;
let val = wb2.get_cell_value("Sheet1", "A1")?;

// Opening without password returns FileEncrypted error
match Workbook::open("encrypted.xlsx") {
    Err(sheetkit::Error::FileEncrypted) => {
        println!("Password required");
    }
    _ => {}
}
```

#### TypeScript

```typescript
import { Workbook } from '@sheetkit/node';

const wb = new Workbook();
wb.setCellValue('Sheet1', 'A1', 'Confidential');

// Save with password (Agile Encryption)
wb.saveWithPassword('encrypted.xlsx', 'mypassword');

// Open encrypted file (sync)
const wb2 = Workbook.openWithPasswordSync('encrypted.xlsx', 'mypassword');
const val = wb2.getCellValue('Sheet1', 'A1');

// Async version
const wb3 = await Workbook.openWithPassword('encrypted.xlsx', 'mypassword');
await wb3.saveWithPassword('encrypted_copy.xlsx', 'newpassword');
```

> Rust에서는 `encryption` feature를 활성화해야 합니다 (`sheetkit = { features = ["encryption"] }`). Node.js 바인딩에는 항상 암호화 지원이 포함됩니다.

---

### 스파크라인

스파크라인은 개별 셀 안에 렌더링되는 미니 차트입니다. 라인, 컬럼, 승패 세 가지 유형을 지원합니다. 스타일 프리셋(0-35)은 Excel 기본 제공 스파크라인 스타일에 대응합니다.

#### Rust

```rust
use sheetkit::{SparklineConfig, SparklineType, Workbook};

let mut wb = Workbook::new();

// Enter data
for i in 1..=10 {
    wb.set_cell_value("Sheet1", &format!("A{i}"), CellValue::from(i as f64 * 1.5))?;
}

// Add column sparkline to cell B1
let mut config = SparklineConfig::new("Sheet1!A1:A10", "B1");
config.sparkline_type = SparklineType::Column;
config.high_point = true;
config.low_point = true;
config.style = Some(5);

wb.add_sparkline("Sheet1", &config)?;

// Get sparklines
let sparklines = wb.get_sparklines("Sheet1")?;

// Remove sparkline by location
wb.remove_sparkline("Sheet1", "B1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Enter data
for (let i = 1; i <= 10; i++) {
    wb.setCellValue('Sheet1', `A${i}`, i * 1.5);
}

// Add column sparkline to cell B1
wb.addSparkline('Sheet1', {
    dataRange: 'Sheet1!A1:A10',
    location: 'B1',
    sparklineType: 'column',
    highPoint: true,
    lowPoint: true,
    style: 5,
});

// Get sparklines
const sparklines = wb.getSparklines('Sheet1');

// Remove sparkline by location
wb.removeSparkline('Sheet1', 'B1');
```

#### SparklineConfig 필드

| 필드 | Rust 타입 | TS 타입 | 설명 |
|------|-----------|---------|------|
| `data_range` / `dataRange` | `String` | `string` | 데이터 소스 범위 (예: `"Sheet1!A1:A10"`) |
| `location` | `String` | `string` | 스파크라인이 렌더링되는 셀 |
| `sparkline_type` / `sparklineType` | `SparklineType` | `string?` | `"line"`, `"column"`, `"stacked"` (승패) |
| `markers` | `bool` | `boolean?` | 데이터 마커 표시 |
| `high_point` / `highPoint` | `bool` | `boolean?` | 최고점 강조 |
| `low_point` / `lowPoint` | `bool` | `boolean?` | 최저점 강조 |
| `first_point` / `firstPoint` | `bool` | `boolean?` | 첫 번째 포인트 강조 |
| `last_point` / `lastPoint` | `bool` | `boolean?` | 마지막 포인트 강조 |
| `negative_points` / `negativePoints` | `bool` | `boolean?` | 음수 값 강조 |
| `show_axis` / `showAxis` | `bool` | `boolean?` | 가로축 표시 |
| `line_weight` / `lineWeight` | `Option<f64>` | `number?` | 선 두께 (포인트) |
| `style` | `Option<u32>` | `number?` | 스타일 프리셋 인덱스 (0-35) |

---

### 정의된 이름

정의된 이름은 셀 범위나 수식에 사람이 읽기 쉬운 이름을 부여합니다. 워크북 범위(모든 곳에서 사용 가능)와 시트 범위(해당 시트에서만 사용 가능)를 지원합니다.

#### Rust

```rust
use sheetkit::Workbook;

let mut wb = Workbook::new();

// Workbook-scoped name
wb.set_defined_name("SalesTotal", "Sheet1!$B$10", None, None)?;

// Sheet-scoped name (with comment)
wb.set_defined_name(
    "LocalRange", "Sheet1!$A$1:$D$10",
    Some("Sheet1"), Some("Local data range"),
)?;

// Get defined name
if let Some(info) = wb.get_defined_name("SalesTotal", None)? {
    println!("Value: {}", info.value);
}

// List all defined names
let names = wb.get_all_defined_names();

// Delete defined name
wb.delete_defined_name("SalesTotal", None)?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Workbook-scoped name
wb.setDefinedName({
    name: 'SalesTotal',
    value: 'Sheet1!$B$10',
});

// Sheet-scoped name (with comment)
wb.setDefinedName({
    name: 'LocalRange',
    value: 'Sheet1!$A$1:$D$10',
    scope: 'Sheet1',
    comment: 'Local data range',
});

// Get defined name (null = workbook scope)
const info = wb.getDefinedName('SalesTotal', null);

// List all defined names
const names = wb.getDefinedNames();

// Delete defined name
wb.deleteDefinedName('SalesTotal', null);
```

---

### 시트 보호

시트 보호는 개별 워크시트의 편집 작업을 제한합니다. 워크북 보호(구조적 변경 방지)와 달리, 시트 보호는 서식 지정, 삽입, 삭제, 정렬, 필터링 등 셀 수준 작업을 제어합니다. 선택적으로 비밀번호를 설정할 수 있습니다.

#### Rust

```rust
use sheetkit::Workbook;
use sheetkit::sheet::SheetProtectionConfig;

let mut wb = Workbook::new();

// Protect sheet with password (allow sorting)
wb.protect_sheet("Sheet1", SheetProtectionConfig {
    password: Some("secret".into()),
    sort: true,
    auto_filter: true,
    ..Default::default()
})?;

// Check if sheet is protected
let is_protected: bool = wb.is_sheet_protected("Sheet1")?;

// Remove protection
wb.unprotect_sheet("Sheet1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Protect sheet with password (allow sorting)
wb.protectSheet('Sheet1', {
    password: 'secret',
    sort: true,
    autoFilter: true,
});

// Check if sheet is protected
const isProtected: boolean = wb.isSheetProtected('Sheet1');

// Remove protection
wb.unprotectSheet('Sheet1');
```

#### SheetProtectionConfig 필드

| 필드 | Rust 타입 | TS 타입 | 설명 |
|------|-----------|---------|------|
| `password` | `Option<String>` | `string?` | 선택적 비밀번호 (레거시 Excel 해시) |
| `select_locked_cells` / `selectLockedCells` | `bool` | `boolean?` | 잠긴 셀 선택 허용 |
| `select_unlocked_cells` / `selectUnlockedCells` | `bool` | `boolean?` | 잠기지 않은 셀 선택 허용 |
| `format_cells` / `formatCells` | `bool` | `boolean?` | 셀 서식 지정 허용 |
| `format_columns` / `formatColumns` | `bool` | `boolean?` | 열 서식 지정 허용 |
| `format_rows` / `formatRows` | `bool` | `boolean?` | 행 서식 지정 허용 |
| `insert_columns` / `insertColumns` | `bool` | `boolean?` | 열 삽입 허용 |
| `insert_rows` / `insertRows` | `bool` | `boolean?` | 행 삽입 허용 |
| `insert_hyperlinks` / `insertHyperlinks` | `bool` | `boolean?` | 하이퍼링크 삽입 허용 |
| `delete_columns` / `deleteColumns` | `bool` | `boolean?` | 열 삭제 허용 |
| `delete_rows` / `deleteRows` | `bool` | `boolean?` | 행 삭제 허용 |
| `sort` | `bool` | `boolean?` | 정렬 허용 |
| `auto_filter` / `autoFilter` | `bool` | `boolean?` | 자동 필터 사용 허용 |
| `pivot_tables` / `pivotTables` | `bool` | `boolean?` | 피벗 테이블 사용 허용 |

---

### 시트 보기 옵션

시트 보기 옵션은 워크시트의 시각적 표시를 제어합니다. 눈금선, 수식 표시, 확대/축소 수준, 보기 모드 등을 포함합니다. 보기 옵션을 설정해도 틀 고정 설정에는 영향을 주지 않습니다.

#### Rust

```rust
use sheetkit::sheet::{SheetViewOptions, ViewMode};

let mut wb = Workbook::new();

// Hide gridlines and set zoom to 150%
wb.set_sheet_view_options("Sheet1", &SheetViewOptions {
    show_gridlines: Some(false),
    zoom_scale: Some(150),
    ..Default::default()
})?;

// Switch to page break preview
wb.set_sheet_view_options("Sheet1", &SheetViewOptions {
    view_mode: Some(ViewMode::PageBreak),
    ..Default::default()
})?;

// Read current settings
let opts = wb.get_sheet_view_options("Sheet1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Hide gridlines and set zoom to 150%
wb.setSheetViewOptions("Sheet1", {
    showGridlines: false,
    zoomScale: 150,
});

// Switch to page break preview
wb.setSheetViewOptions("Sheet1", {
    viewMode: "pageBreak",
});

// Read current settings
const opts = wb.getSheetViewOptions("Sheet1");
```

보기 모드: `"normal"` (기본값), `"pageBreak"`, `"pageLayout"`. 확대/축소 범위: 10-400.

자세한 API 설명은 [API 레퍼런스](../api-reference/advanced.md#31-시트-보기-옵션)를 참조하세요.

---

### 시트 표시 여부

Excel UI에서 시트 탭의 표시 여부를 제어합니다. 세 가지 표시 상태를 사용할 수 있습니다: 표시(기본값), 숨김(사용자가 UI를 통해 숨김 해제 가능), 매우 숨김(코드를 통해서만 숨김 해제 가능). 최소 하나의 시트는 항상 표시 상태여야 합니다.

#### Rust

```rust
use sheetkit::sheet::SheetVisibility;

let mut wb = Workbook::new();
wb.new_sheet("Config")?;
wb.new_sheet("Internal")?;

// Hide Config sheet (user can unhide via Excel UI)
wb.set_sheet_visibility("Config", SheetVisibility::Hidden)?;

// Set Internal sheet to very hidden (only unhidable via code)
wb.set_sheet_visibility("Internal", SheetVisibility::VeryHidden)?;

// Check visibility status
let vis = wb.get_sheet_visibility("Config")?;
assert_eq!(vis, SheetVisibility::Hidden);

// Make visible again
wb.set_sheet_visibility("Config", SheetVisibility::Visible)?;
```

#### TypeScript

```typescript
const wb = new Workbook();
wb.newSheet("Config");
wb.newSheet("Internal");

// Hide sheets
wb.setSheetVisibility("Config", "hidden");
wb.setSheetVisibility("Internal", "veryHidden");

// Check visibility status
const vis = wb.getSheetVisibility("Config"); // "hidden"

// Make visible again
wb.setSheetVisibility("Config", "visible");
```

자세한 API 설명은 [API 레퍼런스](../api-reference/advanced.md#32-시트-표시-여부)를 참조하세요.

---

## 예제 프로젝트

모든 기능을 보여주는 완전한 예제 프로젝트가 저장소에 포함되어 있습니다:

- **Rust**: `examples/rust/` -- 독립된 Cargo 프로젝트 (해당 디렉토리에서 `cargo run` 실행)
- **Node.js**: `examples/node/` -- TypeScript 프로젝트 (네이티브 모듈을 먼저 빌드한 후 `npx tsx index.ts`로 실행)

각 예제는 워크북 생성, 셀 값 설정, 시트 관리, 스타일 적용, 차트와 이미지 추가, 데이터 유효성 검사, 코멘트, 자동 필터, 대용량 데이터 스트리밍, 문서 속성, 워크북 보호, 셀 병합, 하이퍼링크, 조건부 서식, 틀 고정, 페이지 레이아웃, 수식 계산, 피벗 테이블, 파일 암호화, 스파크라인, 정의된 이름, 시트 보호 등 모든 기능을 순서대로 시연합니다.

---

## 유틸리티 함수

SheetKit은 셀 참조 변환을 위한 도우미 함수도 제공합니다:

```rust
use sheetkit::utils::cell_ref;

// Convert cell name to (column, row) coordinates
let (col, row) = cell_ref::cell_name_to_coordinates("B3")?;  // (2, 3)

// Convert coordinates to cell name
let name = cell_ref::coordinates_to_cell_name(2, 3)?;  // "B3"

// Convert column name to number
let num = cell_ref::column_name_to_number("AA")?;  // 27

// Convert column number to name
let name = cell_ref::column_number_to_name(27)?;  // "AA"
```

---

### 테마 색상

테마 색상 슬롯(dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink)을 선택적 틴트 값과 함께 조회합니다.

#### Rust

```rust
use sheetkit::Workbook;

let wb = Workbook::new();

// Get accent1 color (no tint)
let color = wb.get_theme_color(4, None); // Some("FF4472C4")

// Lighten black (index 0) by 50%
let lightened = wb.get_theme_color(0, Some(0.5)); // Some("FF7F7F7F")

// Out of range returns None
let invalid = wb.get_theme_color(99, None); // None
```

#### TypeScript

```typescript
const wb = new Workbook();

// Get accent1 color (no tint)
const color = wb.getThemeColor(4, null); // "FF4472C4"

// Lighten black by 50%
const lightened = wb.getThemeColor(0, 0.5); // "FF7F7F7F"

// Darken white by 50%
const darkened = wb.getThemeColor(1, -0.5); // "FF7F7F7F"

// Out of range returns null
const invalid = wb.getThemeColor(99, null); // null
```

테마 색상 인덱스: 0 (dk1), 1 (lt1), 2 (dk2), 3 (lt2), 4-9 (accent1-6), 10 (hlink), 11 (folHlink).

그래디언트 채우기를 포함한 자세한 내용은 [API 레퍼런스](../api-reference/advanced.md#27-theme-colors)를 참조하세요.

---

### 서식 있는 텍스트

서식 있는 텍스트(Rich Text)를 사용하면 하나의 셀에 각각 독립적인 서식을 가진 여러 텍스트 조각(run)을 넣을 수 있습니다.

#### Rust

```rust
use sheetkit::{Workbook, RichTextRun};

let mut wb = Workbook::new();

// Set rich text with multiple formatted runs
wb.set_cell_rich_text("Sheet1", "A1", vec![
    RichTextRun {
        text: "Bold red".to_string(),
        font: Some("Arial".to_string()),
        size: Some(14.0),
        bold: true,
        italic: false,
        color: Some("#FF0000".to_string()),
    },
    RichTextRun {
        text: " normal text".to_string(),
        font: None,
        size: None,
        bold: false,
        italic: false,
        color: None,
    },
])?;

// Read rich text
if let Some(runs) = wb.get_cell_rich_text("Sheet1", "A1")? {
    for run in &runs {
        println!("Text: {:?}, Bold: {}", run.text, run.bold);
    }
}
```

#### TypeScript

```typescript
const wb = new Workbook();

// Set rich text with multiple formatted runs
wb.setCellRichText('Sheet1', 'A1', [
  { text: 'Bold red', font: 'Arial', size: 14, bold: true, color: '#FF0000' },
  { text: ' normal text' },
]);

// Read rich text
const runs = wb.getCellRichText('Sheet1', 'A1');
if (runs) {
  for (const run of runs) {
    console.log(`Text: ${run.text}, Bold: ${run.bold ?? false}`);
  }
}
```

`RichTextRun` 필드 및 `rich_text_to_plain`을 포함한 자세한 내용은 [API 레퍼런스](../api-reference/advanced.md#28-서식-있는-텍스트)를 참조하세요.

---

### 파일 암호화

전체 .xlsx 파일을 비밀번호로 보호합니다. 암호화된 파일은 일반 ZIP 대신 OLE/CFB 컨테이너를 사용합니다.

> Rust에서는 `encryption` feature가 필요합니다: `sheetkit = { features = ["encryption"] }`. Node.js 바인딩에는 항상 암호화 지원이 포함됩니다.

#### Rust

```rust
use sheetkit::Workbook;

// Save with password (Agile Encryption, AES-256-CBC)
let mut wb = Workbook::new();
wb.save_with_password("encrypted.xlsx", "secret")?;

// Open with password
let wb2 = Workbook::open_with_password("encrypted.xlsx", "secret")?;

// Detect encrypted file
match Workbook::open("file.xlsx") {
    Ok(wb) => { /* Unencrypted file */ }
    Err(sheetkit::Error::FileEncrypted) => {
        let wb = Workbook::open_with_password("file.xlsx", "password")?;
    }
    Err(e) => return Err(e),
}
```

#### TypeScript

```typescript
const wb = new Workbook();

// Save with password
wb.saveWithPassword('encrypted.xlsx', 'secret');

// Open with password (sync)
const wb2 = Workbook.openWithPasswordSync('encrypted.xlsx', 'secret');

// Open with password (async)
const wb3 = await Workbook.openWithPassword('encrypted.xlsx', 'secret');
```

에러 타입 및 암호화 사양을 포함한 자세한 내용은 [API 레퍼런스](../api-reference/advanced.md#29-파일-암호화)를 참조하세요.
---

### 성능 최적화

대용량 시트를 읽을 때 SheetKit은 Buffer 기반 FFI 전송을 사용하여 메모리 사용량을 대폭 줄입니다. Node.js 바인딩은 셀 데이터를 읽기 위한 세 가지 API를 제공하며, 사용 사례에 따라 적합한 API를 선택할 수 있습니다.

#### getRows() -- 기존 코드 호환

가장 단순한 방법으로 기존 코드를 변경할 필요가 없습니다. 내부적으로 Buffer 전송을 사용하므로 이전 버전보다 메모리 효율이 높습니다.

```typescript
const wb = Workbook.openSync('large.xlsx');
const rows = wb.getRows('Sheet1');
for (const row of rows) {
  for (const cell of row.cells) {
    console.log(`${cell.column}: ${cell.value ?? cell.numberValue}`);
  }
}
```

#### getRowsBuffer() + SheetData -- 대용량 시트에 최적

`SheetData` 클래스를 사용하면 전체 시트를 디코딩하지 않고 필요한 셀만 O(1)로 접근할 수 있습니다. 대용량 시트에서 특정 영역만 읽을 때 가장 효율적입니다.

```typescript
import { SheetData } from '@sheetkit/node/sheet-data';

const wb = Workbook.openSync('large.xlsx');
const buf = wb.getRowsBuffer('Sheet1');
const sheet = new SheetData(buf);

// Access specific cells (1-based row/column)
const header = sheet.getRow(1);
const value = sheet.getCell(100, 3);  // 100행, C열

// Iterate over all rows
for (const { row, values } of sheet.rows()) {
  // row: row number, values: value array
}

// Convert to 2D array
const data = sheet.toArray();
```

#### getRowsBuffer() -- 커스텀 처리

raw Buffer를 직접 사용하여 커스텀 디코더를 구현하거나 네트워크로 전송할 수 있습니다. Buffer 형식은 [아키텍처](../architecture.md#6-buffer-기반-ffi-전송) 문서를 참조하세요.

```typescript
const buf = wb.getRowsBuffer('Sheet1');
// Pass to custom decoder, send over network, cache, etc.
```

#### 메모리 비교

50,000행 x 20열 시트 읽기 시 예상 메모리 사용량:

| API | Node.js 메모리 | 비고 |
|-----|---------------|------|
| `getRows()` (이전 버전) | ~400MB | 100만 napi 객체 생성 |
| `getRows()` (Buffer 기반) | ~80MB | Buffer 디코딩 후 JS 객체 생성 |
| `getRowsBuffer()` + `SheetData` | ~50MB | Buffer만 유지, 필요 시 디코딩 |
| `getRowsBuffer()` (raw) | ~30MB | Buffer만 유지, 디코딩 없음 |

자세한 API 설명은 [API 레퍼런스](../api-reference/advanced.md#30-대량-데이터-전송)를 참조하세요.
