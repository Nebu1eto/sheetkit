# SheetKit 사용 가이드

SheetKit은 Excel(.xlsx) 파일을 읽고 쓰기 위한 Rust 라이브러리이며, napi-rs를 통해 Node.js 바인딩을 제공합니다.

---

## 목차

- [설치](#설치)
- [빠른 시작](#빠른-시작)
- [API 레퍼런스](#api-레퍼런스)
  - [워크북 I/O](#워크북-io)
  - [셀 조작](#셀-조작)
  - [시트 관리](#시트-관리)
  - [행/열 조작](#행열-조작)
  - [스타일](#스타일)
  - [차트](#차트)
  - [이미지](#이미지)
  - [데이터 유효성 검사](#데이터-유효성-검사)
  - [코멘트](#코멘트)
  - [자동 필터](#자동-필터)
  - [StreamWriter](#streamwriter)
  - [문서 속성](#문서-속성)
  - [워크북 보호](#워크북-보호)
  - [셀 병합](#셀-병합)
  - [하이퍼링크](#하이퍼링크)
  - [조건부 서식](#조건부-서식)
  - [틀 고정/분할](#틀-고정분할)
  - [페이지 레이아웃](#페이지-레이아웃)
  - [행/열 이터레이터](#행열-이터레이터)
  - [행/열 아웃라인 및 스타일](#행열-아웃라인-및-스타일)
  - [수식 계산](#수식-계산)
  - [피벗 테이블](#피벗-테이블)
- [예제 프로젝트](#예제-프로젝트)

---

## 설치

### Rust

`Cargo.toml`에 `sheetkit`을 추가합니다:

```toml
[dependencies]
sheetkit = "0.1"
```

### Node.js

```bash
npm install sheetkit
```

> Node.js 패키지는 napi-rs로 빌드된 네이티브 애드온입니다. 설치 시 네이티브 모듈을 컴파일하기 위해 Rust 빌드 도구(rustc, cargo)가 필요합니다.

---

## 빠른 시작

### Rust

```rust
use sheetkit::{CellValue, Workbook};

fn main() -> sheetkit::Result<()> {
    // 새 워크북 생성 (기본적으로 "Sheet1" 포함)
    let mut wb = Workbook::new();

    // 셀 값 쓰기
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Name".into()))?;
    wb.set_cell_value("Sheet1", "B1", CellValue::String("Age".into()))?;
    wb.set_cell_value("Sheet1", "A2", CellValue::String("John Doe".into()))?;
    wb.set_cell_value("Sheet1", "B2", CellValue::Number(30.0))?;

    // 셀 값 읽기
    let val = wb.get_cell_value("Sheet1", "A1")?;
    println!("A1 = {:?}", val);

    // 파일로 저장
    wb.save("output.xlsx")?;

    // 기존 파일 열기
    let wb2 = Workbook::open("output.xlsx")?;
    println!("Sheets: {:?}", wb2.sheet_names());

    Ok(())
}
```

### TypeScript / Node.js

```typescript
import { Workbook } from 'sheetkit';

// 새 워크북 생성 (기본적으로 "Sheet1" 포함)
const wb = new Workbook();

// 셀 값 쓰기
wb.setCellValue('Sheet1', 'A1', 'Name');
wb.setCellValue('Sheet1', 'B1', 'Age');
wb.setCellValue('Sheet1', 'A2', 'John Doe');
wb.setCellValue('Sheet1', 'B2', 30);

// 셀 값 읽기
const val = wb.getCellValue('Sheet1', 'A1');
console.log('A1 =', val);

// 파일로 저장
wb.save('output.xlsx');

// 기존 파일 열기
const wb2 = Workbook.open('output.xlsx');
console.log('Sheets:', wb2.sheetNames);
```

---

## API 레퍼런스

### 워크북 I/O

워크북을 생성, 열기, 저장하는 기본 기능입니다.

#### Rust

```rust
use sheetkit::Workbook;

// "Sheet1"이 포함된 빈 워크북 생성
let mut wb = Workbook::new();

// 기존 .xlsx 파일 열기
let wb = Workbook::open("input.xlsx")?;

// .xlsx 파일로 저장
wb.save("output.xlsx")?;

// 모든 시트 이름 조회
let names: Vec<&str> = wb.sheet_names();
```

#### TypeScript

```typescript
import { Workbook } from 'sheetkit';

// "Sheet1"이 포함된 빈 워크북 생성
const wb = new Workbook();

// 기존 .xlsx 파일 열기
const wb2 = Workbook.open('input.xlsx');

// .xlsx 파일로 저장
wb.save('output.xlsx');

// 모든 시트 이름 조회
const names: string[] = wb.sheetNames;
```

---

### 셀 조작

셀 값을 읽고 씁니다. 셀은 시트 이름과 셀 참조(예: `"A1"`, `"B2"`, `"AA100"`)로 식별합니다.

#### CellValue 타입

| Rust 변형               | TypeScript 타입 | 설명                                |
|--------------------------|-----------------|-------------------------------------|
| `CellValue::String(s)`  | `string`        | 텍스트 값                            |
| `CellValue::Number(n)`  | `number`        | 숫자 값 (내부적으로 f64로 저장)       |
| `CellValue::Bool(b)`    | `boolean`       | 불리언 값                            |
| `CellValue::Empty`      | `null`          | 빈 셀 / 값 지우기                    |
| `CellValue::Formula{..}`| --              | 수식 (Rust 전용)                     |
| `CellValue::Error(e)`   | --              | `#DIV/0!` 같은 에러 값 (Rust 전용)   |

#### Rust

```rust
use sheetkit::{CellValue, Workbook};

let mut wb = Workbook::new();

// 다양한 타입의 값 설정
wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))?;
wb.set_cell_value("Sheet1", "B1", CellValue::Number(42.0))?;
wb.set_cell_value("Sheet1", "C1", CellValue::Bool(true))?;
wb.set_cell_value("Sheet1", "D1", CellValue::Empty)?;

// From 트레이트를 활용한 편리한 변환
wb.set_cell_value("Sheet1", "A2", CellValue::from("Text"))?;
wb.set_cell_value("Sheet1", "B2", CellValue::from(100i32))?;
wb.set_cell_value("Sheet1", "C2", CellValue::from(3.14))?;

// 셀 값 읽기
let val = wb.get_cell_value("Sheet1", "A1")?;
match val {
    CellValue::String(s) => println!("String: {}", s),
    CellValue::Number(n) => println!("Number: {}", n),
    CellValue::Bool(b) => println!("Bool: {}", b),
    CellValue::Empty => println!("(empty)"),
    _ => {}
}
```

#### TypeScript

```typescript
// 값 설정 -- JavaScript 값의 타입에 따라 자동으로 결정됨
wb.setCellValue('Sheet1', 'A1', 'Hello');       // string
wb.setCellValue('Sheet1', 'B1', 42);            // number
wb.setCellValue('Sheet1', 'C1', true);          // boolean
wb.setCellValue('Sheet1', 'D1', null);          // 셀 비우기

// 셀 값 읽기 -- string | number | boolean | null 반환
const val = wb.getCellValue('Sheet1', 'A1');
```

---

### 시트 관리

시트를 생성, 삭제, 이름 변경, 복사하고 활성 시트를 설정합니다.

#### Rust

```rust
let mut wb = Workbook::new();

// 새 시트 생성 (0부터 시작하는 인덱스 반환)
let idx: usize = wb.new_sheet("Sales")?;

// 시트 삭제
wb.delete_sheet("Sales")?;

// 시트 이름 변경
wb.set_sheet_name("Sheet1", "Main")?;

// 시트 복사 (새 시트의 인덱스 반환)
let idx: usize = wb.copy_sheet("Main", "Main_Copy")?;

// 시트 인덱스 조회 (없으면 None)
let idx: Option<usize> = wb.get_sheet_index("Main");

// 활성 시트 조회/설정
let active: &str = wb.get_active_sheet();
wb.set_active_sheet("Main")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// 새 시트 생성 (0부터 시작하는 인덱스 반환)
const idx: number = wb.newSheet('Sales');

// 시트 삭제
wb.deleteSheet('Sales');

// 시트 이름 변경
wb.setSheetName('Sheet1', 'Main');

// 시트 복사 (새 시트의 인덱스 반환)
const copyIdx: number = wb.copySheet('Main', 'Main_Copy');

// 시트 인덱스 조회 (없으면 null)
const sheetIdx: number | null = wb.getSheetIndex('Main');

// 활성 시트 조회/설정
const active: string = wb.getActiveSheet();
wb.setActiveSheet('Main');
```

---

### 행/열 조작

행과 열을 삽입, 삭제하고 크기 및 표시 여부를 설정합니다.

#### Rust

```rust
let mut wb = Workbook::new();

// -- 행 (1부터 시작하는 행 번호) --

// 2번 행부터 3개의 빈 행 삽입
wb.insert_rows("Sheet1", 2, 3)?;

// 5번 행 삭제
wb.remove_row("Sheet1", 5)?;

// 1번 행 복제 (아래에 복사본 삽입)
wb.duplicate_row("Sheet1", 1)?;

// 행 높이 설정/조회
wb.set_row_height("Sheet1", 1, 25.0)?;
let height: Option<f64> = wb.get_row_height("Sheet1", 1)?;

// 행 표시/숨김
wb.set_row_visible("Sheet1", 3, false)?;

// -- 열 (알파벳 기반, 예: "A", "B", "AA") --

// 열 너비 설정/조회
wb.set_col_width("Sheet1", "A", 20.0)?;
let width: Option<f64> = wb.get_col_width("Sheet1", "A")?;

// 열 표시/숨김
wb.set_col_visible("Sheet1", "B", false)?;

// "C" 열부터 2개의 빈 열 삽입
wb.insert_cols("Sheet1", "C", 2)?;

// "D" 열 삭제
wb.remove_col("Sheet1", "D")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// -- 행 (1부터 시작하는 행 번호) --
wb.insertRows('Sheet1', 2, 3);
wb.removeRow('Sheet1', 5);
wb.duplicateRow('Sheet1', 1);
wb.setRowHeight('Sheet1', 1, 25);
const height: number | null = wb.getRowHeight('Sheet1', 1);
wb.setRowVisible('Sheet1', 3, false);

// -- 열 (알파벳 기반) --
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

// 스타일 등록
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

// 셀에 스타일 적용
wb.set_cell_style("Sheet1", "A1", style_id)?;

// 셀의 스타일 ID 조회 (기본 스타일이면 None)
let current_style: Option<u32> = wb.get_cell_style("Sheet1", "A1")?;
```

#### TypeScript

```typescript
// 스타일 등록
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

// 셀에 스타일 적용
wb.setCellStyle('Sheet1', 'A1', styleId);

// 셀의 스타일 ID 조회 (기본 스타일이면 null)
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

// 기본 제공 형식 (예: 퍼센트, 날짜, 통화)
NumFmtStyle::Builtin(9)  // 0%

// 사용자 정의 형식 문자열
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

#### 지원 차트 유형 (41종)

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

// 먼저 데이터를 채웁니다...
wb.set_cell_value("Sheet1", "A1", CellValue::String("Q1".into()))?;
wb.set_cell_value("Sheet1", "B1", CellValue::Number(1500.0))?;
// ... 더 많은 데이터 행 ...

// D1~K15 영역에 차트 추가
wb.add_chart(
    "Sheet1",
    "D1",   // 좌상단 앵커 셀
    "K15",  // 우하단 앵커 셀
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

워크시트에 이미지(PNG, JPEG, GIF)를 삽입합니다. 이미지는 셀에 앵커링되며 픽셀 크기로 지정합니다.

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
    format: 'png',        // "png" | "jpeg" | "gif"
    fromCell: 'B2',
    widthPx: 200,
    heightPx: 100,
});
```

---

### 데이터 유효성 검사

셀 범위에 데이터 유효성 검사 규칙을 추가합니다. 이 규칙은 사용자가 해당 셀에 입력할 수 있는 값을 제한합니다.

#### 유효성 검사 유형

| Rust 변형                   | TS 문자열       | 설명                  |
|----------------------------|----------------|-----------------------|
| `ValidationType::Whole`    | `"whole"`      | 정수 제약              |
| `ValidationType::Decimal`  | `"decimal"`    | 소수 제약              |
| `ValidationType::List`     | `"list"`       | 드롭다운 목록          |
| `ValidationType::Date`     | `"date"`       | 날짜 제약              |
| `ValidationType::Time`     | `"time"`       | 시간 제약              |
| `ValidationType::TextLength`| `"textLength"` | 텍스트 길이 제약       |
| `ValidationType::Custom`   | `"custom"`     | 사용자 정의 수식 제약   |

#### 비교 연산자

`Between`, `NotBetween`, `Equal`, `NotEqual`, `LessThan`, `LessThanOrEqual`, `GreaterThan`, `GreaterThanOrEqual`.

TypeScript에서는 소문자 문자열 사용: `"between"`, `"notBetween"`, `"equal"` 등.

#### 오류 스타일

`Stop`, `Warning`, `Information` -- 잘못된 입력 시 표시되는 오류 대화 상자의 심각도를 제어합니다.

#### Rust

```rust
use sheetkit::{DataValidationConfig, ErrorStyle, ValidationType, Workbook};

let mut wb = Workbook::new();

// 드롭다운 목록 유효성 검사
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

// 시트의 모든 유효성 검사 조회
let validations = wb.get_data_validations("Sheet1")?;

// 셀 범위 참조로 유효성 검사 제거
wb.remove_data_validation("Sheet1", "C2:C100")?;
```

#### TypeScript

```typescript
// 드롭다운 목록 유효성 검사
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

// 시트의 모든 유효성 검사 조회
const validations = wb.getDataValidations('Sheet1');

// 셀 범위 참조로 유효성 검사 제거
wb.removeDataValidation('Sheet1', 'C2:C100');
```

---

### 코멘트

셀 코멘트를 추가, 조회, 삭제합니다.

#### Rust

```rust
use sheetkit::{CommentConfig, Workbook};

let mut wb = Workbook::new();

// 코멘트 추가
wb.add_comment(
    "Sheet1",
    &CommentConfig {
        cell: "A1".into(),
        author: "Admin".into(),
        text: "This cell contains the project name.".into(),
    },
)?;

// 시트의 모든 코멘트 조회
let comments: Vec<CommentConfig> = wb.get_comments("Sheet1")?;

// 특정 셀의 코멘트 삭제
wb.remove_comment("Sheet1", "A1")?;
```

#### TypeScript

```typescript
// 코멘트 추가
wb.addComment('Sheet1', {
    cell: 'A1',
    author: 'Admin',
    text: 'This cell contains the project name.',
});

// 시트의 모든 코멘트 조회
const comments = wb.getComments('Sheet1');

// 특정 셀의 코멘트 삭제
wb.removeComment('Sheet1', 'A1');
```

---

### 자동 필터

열 범위에 자동 필터 드롭다운을 적용하거나 제거합니다.

#### Rust

```rust
// 범위에 자동 필터 설정
wb.set_auto_filter("Sheet1", "A1:D100")?;

// 자동 필터 제거
wb.remove_auto_filter("Sheet1")?;
```

#### TypeScript

```typescript
// 범위에 자동 필터 설정
wb.setAutoFilter('Sheet1', 'A1:D100');

// 자동 필터 제거
wb.removeAutoFilter('Sheet1');
```

---

### StreamWriter

StreamWriter는 대용량 시트를 효율적으로 작성하기 위한 순방향 전용 스트리밍 API입니다. 전체 워크시트를 메모리에 올리지 않고 내부 버퍼에 XML을 직접 씁니다.

행은 반드시 오름차순으로 작성해야 합니다. 열 너비는 행을 쓰기 전에 설정해야 합니다.

#### Rust

```rust
use sheetkit::{CellValue, Workbook};

let mut wb = Workbook::new();

// 새 시트를 위한 스트림 라이터 생성
let mut sw = wb.new_stream_writer("LargeSheet")?;

// 열 너비 설정 (행 작성 전에 해야 함)
sw.set_col_width(1, 20.0)?;     // 1번 열 (A)
sw.set_col_width(2, 15.0)?;     // 2번 열 (B)

// 오름차순으로 행 작성 (1부터 시작)
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

// 셀 병합 추가 (선택 사항)
sw.add_merge_cell("A1:B1")?;

// 스트림 라이터를 워크북에 적용
wb.apply_stream_writer(sw)?;

wb.save("large_file.xlsx")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// 새 시트를 위한 스트림 라이터 생성
const sw = wb.newStreamWriter('LargeSheet');

// 열 너비 설정 (행 작성 전에 해야 함)
sw.setColWidth(1, 20);     // 1번 열 (A)
sw.setColWidth(2, 15);     // 2번 열 (B)

// 오름차순으로 행 작성 (1부터 시작)
sw.writeRow(1, ['Name', 'Score']);
for (let i = 2; i <= 10000; i++) {
    sw.writeRow(i, [`User_${i - 1}`, i * 1.5]);
}

// 셀 병합 추가 (선택 사항)
sw.addMergeCell('A1:B1');

// 스트림 라이터를 워크북에 적용
wb.applyStreamWriter(sw);

wb.save('large_file.xlsx');
```

#### StreamWriter API 요약

| 메서드                 | 설명                                        |
|-----------------------|---------------------------------------------|
| `set_col_width`       | 단일 열 너비 설정 (1부터 시작하는 열 번호)     |
| `set_col_width_range` | 열 범위의 너비 설정 (Rust 전용)               |
| `write_row`           | 지정한 행 번호에 값 배열 작성                  |
| `add_merge_cell`      | 셀 병합 참조 추가 (예: `"A1:C3"`)            |

---

### 문서 속성

문서 메타데이터를 설정하고 읽습니다: 핵심 속성(제목, 작성자 등), 애플리케이션 속성, 사용자 정의 속성.

#### Rust

```rust
use sheetkit::{AppProperties, CustomPropertyValue, DocProperties, Workbook};

let mut wb = Workbook::new();

// 핵심 문서 속성
wb.set_doc_props(DocProperties {
    title: Some("Annual Report".into()),
    creator: Some("SheetKit".into()),
    description: Some("Financial data for 2025".into()),
    ..Default::default()
});
let props = wb.get_doc_props();

// 애플리케이션 속성
wb.set_app_props(AppProperties {
    application: Some("SheetKit".into()),
    company: Some("Acme Corp".into()),
    ..Default::default()
});
let app_props = wb.get_app_props();

// 사용자 정의 속성 (문자열, 정수, 실수, 불리언, 날짜시간)
wb.set_custom_property("Project", CustomPropertyValue::String("SheetKit".into()));
wb.set_custom_property("Version", CustomPropertyValue::Int(1));
wb.set_custom_property("Released", CustomPropertyValue::Bool(false));

let val = wb.get_custom_property("Project");
let deleted = wb.delete_custom_property("Version");
```

#### TypeScript

```typescript
// 핵심 문서 속성
wb.setDocProps({
    title: 'Annual Report',
    creator: 'SheetKit',
    description: 'Financial data for 2025',
});
const props = wb.getDocProps();

// 애플리케이션 속성
wb.setAppProps({
    application: 'SheetKit',
    company: 'Acme Corp',
});
const appProps = wb.getAppProps();

// 사용자 정의 속성 (문자열, 숫자, 불리언)
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

// 워크북 보호
wb.protect_workbook(WorkbookProtectionConfig {
    password: Some("secret".into()),
    lock_structure: true,    // 시트 추가/삭제/이름 변경 방지
    lock_windows: false,     // 창 크기 조정 허용
    lock_revision: false,    // 수정 내용 추적 변경 허용
});

// 보호 상태 확인
let is_protected: bool = wb.is_workbook_protected();

// 보호 해제
wb.unprotect_workbook();
```

#### TypeScript

```typescript
// 워크북 보호
wb.protectWorkbook({
    password: 'secret',
    lockStructure: true,
    lockWindows: false,
    lockRevision: false,
});

// 보호 상태 확인
const isProtected: boolean = wb.isWorkbookProtected();

// 보호 해제
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

// 외부 URL
wb.set_cell_hyperlink(
    "Sheet1", "A1",
    HyperlinkType::External("https://example.com".into()),
    Some("Example Site"), Some("Click here"),
)?;

// 이메일
wb.set_cell_hyperlink(
    "Sheet1", "A2",
    HyperlinkType::Email("mailto:user@example.com".into()),
    Some("Send email"), None,
)?;

// 내부 시트 참조
wb.set_cell_hyperlink(
    "Sheet1", "A3",
    HyperlinkType::Internal("Sheet2!A1".into()),
    None, None,
)?;

// 조회 및 삭제
let info = wb.get_cell_hyperlink("Sheet1", "A1")?;
wb.delete_cell_hyperlink("Sheet1", "A1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// 외부 URL
wb.setCellHyperlink('Sheet1', 'A1', {
    linkType: 'external',
    target: 'https://example.com',
    display: 'Example Site',
    tooltip: 'Click here',
});

// 이메일
wb.setCellHyperlink('Sheet1', 'A2', {
    linkType: 'email',
    target: 'mailto:user@example.com',
    display: 'Send email',
});

// 내부 시트 참조
wb.setCellHyperlink('Sheet1', 'A3', {
    linkType: 'internal',
    target: 'Sheet2!A1',
});

// 조회 및 삭제
const info = wb.getCellHyperlink('Sheet1', 'A1');
wb.deleteCellHyperlink('Sheet1', 'A1');
```

---

### 조건부 서식

셀 값이나 수식에 따라 자동으로 서식을 적용합니다. 18가지 규칙 유형을 지원합니다.

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

wb.set_panes("Sheet1", "A2")?;    // 첫 행 고정
wb.set_panes("Sheet1", "B1")?;    // 첫 열 고정
wb.set_panes("Sheet1", "B2")?;    // 첫 행 + 첫 열 고정

let pane = wb.get_panes("Sheet1")?;
wb.unset_panes("Sheet1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

wb.setPanes('Sheet1', 'A2');       // 첫 행 고정
wb.setPanes('Sheet1', 'B2');       // 첫 행 + 첫 열 고정

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

// 여백 (인치 단위)
wb.set_page_margins("Sheet1", &PageMarginsConfig {
    left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3,
})?;

// 페이지 설정
wb.set_page_setup("Sheet1", Some(Orientation::Landscape), Some(PaperSize::A4), Some(100), None, None)?;

// 머리글/바닥글
wb.set_header_footer("Sheet1", Some("&CMonthly Report"), Some("&LPage &P of &N"))?;

// 페이지 나누기
wb.insert_page_break("Sheet1", 20)?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// 여백 (인치 단위)
wb.setPageMargins('Sheet1', {
    left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3,
});

// 페이지 설정
wb.setPageSetup('Sheet1', {
    paperSize: 'a4', orientation: 'landscape', scale: 100,
});

// 인쇄 옵션
wb.setPrintOptions('Sheet1', { gridLines: true, horizontalCentered: true });

// 머리글/바닥글
wb.setHeaderFooter('Sheet1', '&CMonthly Report', '&LPage &P of &N');

// 페이지 나누기
wb.insertPageBreak('Sheet1', 20);
```

---

### 행/열 이터레이터

시트의 모든 행 또는 열 데이터를 한 번에 조회합니다. 데이터가 있는 행/열만 포함됩니다.

#### Rust

```rust
let wb = Workbook::open("data.xlsx")?;

// 모든 행 조회
let rows = wb.get_rows("Sheet1")?;
for (row_num, cells) in &rows {
    for (col, val) in cells {
        println!("Row {}, Col {}: {:?}", row_num, col, val);
    }
}

// 모든 열 조회
let cols = wb.get_cols("Sheet1")?;
```

#### TypeScript

```typescript
const wb = Workbook.open('data.xlsx');

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

// 아웃라인 수준
wb.set_row_outline_level("Sheet1", 2, 1)?;
let level: u8 = wb.get_row_outline_level("Sheet1", 2)?;

wb.set_col_outline_level("Sheet1", "B", 2)?;
let col_level: u8 = wb.get_col_outline_level("Sheet1", "B")?;

// 행/열 스타일
let style_id = wb.add_style(&style)?;
wb.set_row_style("Sheet1", 1, style_id)?;
wb.set_col_style("Sheet1", "A", style_id)?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// 아웃라인 수준
wb.setRowOutlineLevel('Sheet1', 2, 1);
const level: number = wb.getRowOutlineLevel('Sheet1', 2);

wb.setColOutlineLevel('Sheet1', 'B', 2);
const colLevel: number = wb.getColOutlineLevel('Sheet1', 'B');

// 행/열 스타일
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

## 예제 프로젝트

모든 기능을 보여주는 완전한 예제 프로젝트가 저장소에 포함되어 있습니다:

- **Rust**: `examples/rust/` -- 독립된 Cargo 프로젝트 (해당 디렉토리에서 `cargo run` 실행)
- **Node.js**: `examples/node/` -- TypeScript 프로젝트 (네이티브 모듈을 먼저 빌드한 후 `npx tsx index.ts`로 실행)

각 예제는 워크북 생성, 셀 값 설정, 시트 관리, 스타일 적용, 차트와 이미지 추가, 데이터 유효성 검사, 코멘트, 자동 필터, 대용량 데이터 스트리밍, 문서 속성, 워크북 보호, 셀 병합, 하이퍼링크, 조건부 서식, 틀 고정, 페이지 레이아웃, 수식 계산, 피벗 테이블 등 모든 기능을 순서대로 시연합니다.

---

## 유틸리티 함수

SheetKit은 셀 참조 변환을 위한 도우미 함수도 제공합니다:

```rust
use sheetkit::utils::cell_ref;

// 셀 이름을 (열, 행) 좌표로 변환
let (col, row) = cell_ref::cell_name_to_coordinates("B3")?;  // (2, 3)

// 좌표를 셀 이름으로 변환
let name = cell_ref::coordinates_to_cell_name(2, 3)?;  // "B3"

// 열 이름을 번호로 변환
let num = cell_ref::column_name_to_number("AA")?;  // 27

// 열 번호를 이름으로 변환
let name = cell_ref::column_number_to_name(27)?;  // "AA"
```
