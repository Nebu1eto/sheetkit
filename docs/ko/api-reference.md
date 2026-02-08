# SheetKit API Reference (Korean)

SheetKit은 Excel (.xlsx) 파일을 읽고 쓰기 위한 Rust 라이브러리이며, napi-rs를 통해 Node.js 바인딩을 제공한다. 이 문서는 Rust와 TypeScript 양쪽 API를 포괄적으로 설명한다.

---

## 목차

1. [워크북 입출력 (Workbook I/O)](#1-워크북-입출력)
2. [셀 조작 (Cell Operations)](#2-셀-조작)
3. [시트 관리 (Sheet Management)](#3-시트-관리)
4. [행 조작 (Row Operations)](#4-행-조작)
5. [열 조작 (Column Operations)](#5-열-조작)
6. [행/열 반복자 (Row/Column Iterators)](#6-행열-반복자)
7. [스타일 (Styles)](#7-스타일)
8. [셀 병합 (Merge Cells)](#8-셀-병합)
9. [하이퍼링크 (Hyperlinks)](#9-하이퍼링크)
10. [차트 (Charts)](#10-차트)
11. [이미지 (Images)](#11-이미지)
12. [데이터 유효성 검사 (Data Validation)](#12-데이터-유효성-검사)
13. [코멘트 (Comments)](#13-코멘트)
14. [자동 필터 (Auto-Filter)](#14-자동-필터)
15. [조건부 서식 (Conditional Formatting)](#15-조건부-서식)
16. [틀 고정 (Freeze/Split Panes)](#16-틀-고정)
17. [페이지 레이아웃 (Page Layout)](#17-페이지-레이아웃)
18. [정의된 이름 (Defined Names)](#18-정의된-이름)
19. [문서 속성 (Document Properties)](#19-문서-속성)
20. [워크북 보호 (Workbook Protection)](#20-워크북-보호)
21. [시트 보호 (Sheet Protection)](#21-시트-보호)
22. [수식 평가 (Formula Evaluation)](#22-수식-평가)
23. [피벗 테이블 (Pivot Tables)](#23-피벗-테이블)
24. [스트림 라이터 (StreamWriter)](#24-스트림-라이터)
25. [유틸리티 함수 (Utility Functions)](#25-유틸리티-함수)

---

## 1. 워크북 입출력

워크북의 생성, 열기, 저장 및 시트 이름 조회를 다루는 기본 API이다.

### `Workbook::new()` / `new Workbook()`

빈 워크북을 생성한다. 기본적으로 "Sheet1"이라는 시트 하나가 포함된다.

**Rust:**

```rust
use sheetkit::Workbook;

let wb = Workbook::new();
```

**TypeScript:**

```typescript
import { Workbook } from "sheetkit";

const wb = new Workbook();
```

### `Workbook::open(path)` / `Workbook.open(path)`

기존 .xlsx 파일을 열어 메모리에 로드한다.

**Rust:**

```rust
let wb = Workbook::open("report.xlsx")?;
```

**TypeScript:**

```typescript
const wb = Workbook.open("report.xlsx");
```

> 파일이 존재하지 않거나 유효한 .xlsx 형식이 아니면 오류가 발생한다.

### `wb.save(path)`

워크북을 .xlsx 파일로 저장한다.

**Rust:**

```rust
wb.save("output.xlsx")?;
```

**TypeScript:**

```typescript
wb.save("output.xlsx");
```

> ZIP 압축은 Deflate 방식을 사용한다.

### `wb.sheet_names()` / `wb.sheetNames`

워크북에 포함된 모든 시트의 이름을 순서대로 반환한다.

**Rust:**

```rust
let names: Vec<&str> = wb.sheet_names();
// ["Sheet1", "Sheet2"]
```

**TypeScript:**

```typescript
const names: string[] = wb.sheetNames;
// ["Sheet1", "Sheet2"]
```

> TypeScript에서는 getter 프로퍼티로 접근한다.

---

## 2. 셀 조작

셀 값의 읽기와 쓰기를 다룬다. 셀 값은 다양한 타입으로 표현된다.

### CellValue 타입

| 타입 | Rust | TypeScript | 설명 |
|------|------|-----------|------|
| 빈 셀 | `CellValue::Empty` | `null` | 값이 없는 셀 |
| 문자열 | `CellValue::String(String)` | `string` | 텍스트 값 |
| 숫자 | `CellValue::Number(f64)` | `number` | 정수와 실수 모두 f64로 저장 |
| 불리언 | `CellValue::Bool(bool)` | `boolean` | true / false |
| 날짜 | `CellValue::Date(f64)` | `DateValue` | Excel 시리얼 번호로 저장 |
| 수식 | `CellValue::Formula { expr, result }` | `string` (수식 문자열) | 수식 및 캐시된 결과 |
| 오류 | `CellValue::Error(String)` | `string` (오류 문자열) | #DIV/0!, #N/A 등 |

### DateValue (TypeScript)

날짜 셀을 표현하는 객체이다.

```typescript
interface DateValue {
  type: "date";     // always "date"
  serial: number;   // Excel serial number (1 = 1900-01-01)
  iso?: string;     // ISO format string ("2024-01-15" or "2024-01-15T14:30:00")
}
```

> Excel은 내부적으로 날짜를 시리얼 번호(정수부: 날짜, 소수부: 시각)로 저장한다. 1900년 윤년 버그가 포함되어 있으며 SheetKit은 이를 올바르게 처리한다.

### `get_cell_value(sheet, cell)` / `getCellValue(sheet, cell)`

지정한 셀의 값을 읽는다.

**Rust:**

```rust
let value = wb.get_cell_value("Sheet1", "A1")?;
match value {
    CellValue::String(s) => println!("text: {}", s),
    CellValue::Number(n) => println!("number: {}", n),
    CellValue::Bool(b) => println!("bool: {}", b),
    CellValue::Date(serial) => println!("date serial: {}", serial),
    CellValue::Empty => println!("empty"),
    CellValue::Formula { expr, result } => println!("formula: {}", expr),
    CellValue::Error(e) => println!("error: {}", e),
}
```

**TypeScript:**

```typescript
const value = wb.getCellValue("Sheet1", "A1");
// value: null | boolean | number | string | DateValue
```

### `set_cell_value(sheet, cell, value)` / `setCellValue(sheet, cell, value)`

셀에 값을 설정한다. null을 전달하면 셀이 비워진다.

**Rust:**

```rust
use sheetkit::CellValue;
use chrono::NaiveDate;

wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))?;
wb.set_cell_value("Sheet1", "B1", CellValue::Number(42.0))?;
wb.set_cell_value("Sheet1", "C1", CellValue::Bool(true))?;
wb.set_cell_value("Sheet1", "D1", CellValue::Empty)?;

// Date
let date = NaiveDate::from_ymd_opt(2024, 6, 15).unwrap();
wb.set_cell_value("Sheet1", "E1", CellValue::from(date))?;

// Formula
wb.set_cell_value("Sheet1", "F1", CellValue::Formula {
    expr: "SUM(A1:B1)".into(),
    result: None,
})?;
```

**TypeScript:**

```typescript
wb.setCellValue("Sheet1", "A1", "Hello");
wb.setCellValue("Sheet1", "B1", 42);
wb.setCellValue("Sheet1", "C1", true);
wb.setCellValue("Sheet1", "D1", null);      // clear

// Date
wb.setCellValue("Sheet1", "E1", { type: "date", serial: 45458 });
```

> 시트 이름이 존재하지 않거나 셀 참조가 유효하지 않으면 오류가 발생한다.

---

## 3. 시트 관리

시트의 생성, 삭제, 이름 변경, 복사 등 시트 단위 조작을 다룬다.

### `new_sheet(name)` / `newSheet(name)`

빈 시트를 추가한다. 0부터 시작하는 시트 인덱스를 반환한다.

**Rust:**

```rust
let index = wb.new_sheet("Data")?;
```

**TypeScript:**

```typescript
const index: number = wb.newSheet("Data");
```

### `delete_sheet(name)` / `deleteSheet(name)`

시트를 삭제한다. 마지막 시트는 삭제할 수 없다.

**Rust:**

```rust
wb.delete_sheet("Data")?;
```

**TypeScript:**

```typescript
wb.deleteSheet("Data");
```

### `set_sheet_name(old, new)` / `setSheetName(old, new)`

시트 이름을 변경한다.

**Rust:**

```rust
wb.set_sheet_name("Sheet1", "Summary")?;
```

**TypeScript:**

```typescript
wb.setSheetName("Sheet1", "Summary");
```

### `copy_sheet(source, target)` / `copySheet(source, target)`

기존 시트를 복사하여 새 시트를 생성한다. 새 시트의 0부터 시작하는 인덱스를 반환한다.

**Rust:**

```rust
let index = wb.copy_sheet("Sheet1", "Sheet1_Copy")?;
```

**TypeScript:**

```typescript
const index: number = wb.copySheet("Sheet1", "Sheet1_Copy");
```

### `get_sheet_index(name)` / `getSheetIndex(name)`

시트의 0부터 시작하는 인덱스를 반환한다. 존재하지 않으면 None / null을 반환한다.

**Rust:**

```rust
let index: Option<usize> = wb.get_sheet_index("Sheet1");
```

**TypeScript:**

```typescript
const index: number | null = wb.getSheetIndex("Sheet1");
```

### `get_active_sheet()` / `getActiveSheet()`

현재 활성 시트의 이름을 반환한다.

**Rust:**

```rust
let name: &str = wb.get_active_sheet();
```

**TypeScript:**

```typescript
const name: string = wb.getActiveSheet();
```

### `set_active_sheet(name)` / `setActiveSheet(name)`

활성 시트를 변경한다.

**Rust:**

```rust
wb.set_active_sheet("Data")?;
```

**TypeScript:**

```typescript
wb.setActiveSheet("Data");
```

---

## 4. 행 조작

행의 삽입, 삭제, 복제 및 높이/가시성/아웃라인/스타일 설정을 다룬다. 모든 행 번호는 1부터 시작한다.

### `insert_rows(sheet, start_row, count)` / `insertRows(sheet, startRow, count)`

지정한 행 번호부터 빈 행을 삽입한다.

**Rust:**

```rust
wb.insert_rows("Sheet1", 3, 2)?; // 3rd row onwards, insert 2 rows
```

**TypeScript:**

```typescript
wb.insertRows("Sheet1", 3, 2);
```

### `remove_row(sheet, row)` / `removeRow(sheet, row)`

행을 삭제한다. 아래쪽 행이 위로 이동한다.

**Rust:**

```rust
wb.remove_row("Sheet1", 5)?;
```

**TypeScript:**

```typescript
wb.removeRow("Sheet1", 5);
```

### `duplicate_row(sheet, row)` / `duplicateRow(sheet, row)`

지정한 행을 바로 아래에 복제한다.

**Rust:**

```rust
wb.duplicate_row("Sheet1", 2)?;
```

**TypeScript:**

```typescript
wb.duplicateRow("Sheet1", 2);
```

### `set_row_height` / `get_row_height`

행 높이를 포인트 단위로 설정하거나 조회한다.

**Rust:**

```rust
wb.set_row_height("Sheet1", 1, 30.0)?;
let height: Option<f64> = wb.get_row_height("Sheet1", 1)?;
```

**TypeScript:**

```typescript
wb.setRowHeight("Sheet1", 1, 30);
const height: number | null = wb.getRowHeight("Sheet1", 1);
```

### `set_row_visible` / `get_row_visible`

행의 표시/숨김을 설정하거나 조회한다.

**Rust:**

```rust
wb.set_row_visible("Sheet1", 3, false)?; // hide
let visible: bool = wb.get_row_visible("Sheet1", 3)?;
```

**TypeScript:**

```typescript
wb.setRowVisible("Sheet1", 3, false);
const visible: boolean = wb.getRowVisible("Sheet1", 3);
```

### `set_row_outline_level` / `get_row_outline_level`

행의 아웃라인(그룹) 수준을 설정하거나 조회한다. 범위는 0-7이다.

**Rust:**

```rust
wb.set_row_outline_level("Sheet1", 2, 1)?;
let level: u8 = wb.get_row_outline_level("Sheet1", 2)?;
```

**TypeScript:**

```typescript
wb.setRowOutlineLevel("Sheet1", 2, 1);
const level: number = wb.getRowOutlineLevel("Sheet1", 2);
```

### `set_row_style` / `get_row_style`

행 전체에 스타일 ID를 적용하거나 조회한다.

**Rust:**

```rust
let style_id = wb.add_style(&style)?;
wb.set_row_style("Sheet1", 1, style_id)?;
let sid: u32 = wb.get_row_style("Sheet1", 1)?;
```

**TypeScript:**

```typescript
const styleId = wb.addStyle({ font: { bold: true } });
wb.setRowStyle("Sheet1", 1, styleId);
const sid: number = wb.getRowStyle("Sheet1", 1);
```

---

## 5. 열 조작

열의 너비, 가시성, 삽입, 삭제 및 아웃라인/스타일 설정을 다룬다. 열은 문자열("A", "B", "AA" 등)로 지정한다.

### `set_col_width` / `get_col_width`

열 너비를 문자 단위로 설정하거나 조회한다.

**Rust:**

```rust
wb.set_col_width("Sheet1", "A", 20.0)?;
let width: Option<f64> = wb.get_col_width("Sheet1", "A")?;
```

**TypeScript:**

```typescript
wb.setColWidth("Sheet1", "A", 20);
const width: number | null = wb.getColWidth("Sheet1", "A");
```

### `set_col_visible` / `get_col_visible`

열의 표시/숨김을 설정하거나 조회한다.

**Rust:**

```rust
wb.set_col_visible("Sheet1", "B", false)?;
let visible: bool = wb.get_col_visible("Sheet1", "B")?;
```

**TypeScript:**

```typescript
wb.setColVisible("Sheet1", "B", false);
const visible: boolean = wb.getColVisible("Sheet1", "B");
```

### `insert_cols(sheet, col, count)` / `insertCols(sheet, col, count)`

지정한 열부터 빈 열을 삽입한다.

**Rust:**

```rust
wb.insert_cols("Sheet1", "C", 3)?;
```

**TypeScript:**

```typescript
wb.insertCols("Sheet1", "C", 3);
```

### `remove_col(sheet, col)` / `removeCol(sheet, col)`

열을 삭제한다. 오른쪽 열이 왼쪽으로 이동한다.

**Rust:**

```rust
wb.remove_col("Sheet1", "B")?;
```

**TypeScript:**

```typescript
wb.removeCol("Sheet1", "B");
```

### `set_col_outline_level` / `get_col_outline_level`

열의 아웃라인(그룹) 수준을 설정하거나 조회한다. 범위는 0-7이다.

**Rust:**

```rust
wb.set_col_outline_level("Sheet1", "A", 2)?;
let level: u8 = wb.get_col_outline_level("Sheet1", "A")?;
```

**TypeScript:**

```typescript
wb.setColOutlineLevel("Sheet1", "A", 2);
const level: number = wb.getColOutlineLevel("Sheet1", "A");
```

### `set_col_style` / `get_col_style`

열 전체에 스타일 ID를 적용하거나 조회한다.

**Rust:**

```rust
wb.set_col_style("Sheet1", "A", style_id)?;
let sid: u32 = wb.get_col_style("Sheet1", "A")?;
```

**TypeScript:**

```typescript
wb.setColStyle("Sheet1", "A", styleId);
const sid: number = wb.getColStyle("Sheet1", "A");
```

---

## 6. 행/열 반복자

시트의 모든 행 또는 열 데이터를 한 번에 조회한다. 데이터가 있는 행/열만 포함된다.

### `get_rows(sheet)` / `getRows(sheet)`

시트의 모든 행과 셀 데이터를 반환한다.

**Rust:**

```rust
let rows = wb.get_rows("Sheet1")?;
// Vec<(u32, Vec<(String, CellValue)>)>
// (row_number, [(column_name, value), ...])
for (row_num, cells) in &rows {
    for (col, val) in cells {
        println!("Row {}, Col {}: {}", row_num, col, val);
    }
}
```

**TypeScript:**

```typescript
const rows: JsRowData[] = wb.getRows("Sheet1");
for (const row of rows) {
    console.log(`Row ${row.row}:`);
    for (const cell of row.cells) {
        console.log(`  ${cell.column}: ${cell.valueType} = ${cell.value}`);
    }
}
```

**JsRowData 구조:**

```typescript
interface JsRowData {
  row: number;           // 1-based row number
  cells: JsRowCell[];
}

interface JsRowCell {
  column: string;        // "A", "B", "AA" ...
  valueType: string;     // "string" | "number" | "boolean" | "date" | "empty" | "error" | "formula"
  value?: string;        // string representation
  numberValue?: number;  // set when valueType is "number"
  boolValue?: boolean;   // set when valueType is "boolean"
}
```

### `get_cols(sheet)` / `getCols(sheet)`

시트의 모든 열과 셀 데이터를 반환한다.

**Rust:**

```rust
let cols = wb.get_cols("Sheet1")?;
// Vec<(String, Vec<(u32, CellValue)>)>
// (column_name, [(row_number, value), ...])
```

**TypeScript:**

```typescript
const cols: JsColData[] = wb.getCols("Sheet1");
for (const col of cols) {
    console.log(`Column ${col.column}:`);
    for (const cell of col.cells) {
        console.log(`  Row ${cell.row}: ${cell.valueType} = ${cell.value}`);
    }
}
```

**JsColData 구조:**

```typescript
interface JsColData {
  column: string;        // "A", "B", "AA" ...
  cells: JsColCell[];
}

interface JsColCell {
  row: number;           // 1-based row number
  valueType: string;     // "string" | "number" | "boolean" | "date" | "empty" | "error" | "formula"
  value?: string;
  numberValue?: number;
  boolValue?: boolean;
}
```

---

## 7. 스타일

셀의 폰트, 채우기, 테두리, 정렬, 숫자 서식 및 보호 설정을 다룬다. 스타일은 먼저 `add_style`로 등록한 후 셀에 적용하는 2단계 방식이다. 동일한 스타일은 자동으로 중복 제거된다.

### `add_style(style)` / `addStyle(style)`

스타일을 등록하고 스타일 ID를 반환한다.

**Rust:**

```rust
use sheetkit::style::*;

let style = Style {
    font: Some(FontStyle {
        name: Some("Arial".into()),
        size: Some(14.0),
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        color: Some(StyleColor::Rgb("#FF0000".into())),
    }),
    fill: Some(FillStyle {
        pattern: PatternType::Solid,
        fg_color: Some(StyleColor::Rgb("#FFFF00".into())),
        bg_color: None,
    }),
    border: Some(BorderStyle {
        left: Some(BorderSideStyle {
            style: BorderLineStyle::Thin,
            color: Some(StyleColor::Rgb("#000000".into())),
        }),
        right: None,
        top: None,
        bottom: None,
        diagonal: None,
    }),
    alignment: Some(AlignmentStyle {
        horizontal: Some(HorizontalAlign::Center),
        vertical: Some(VerticalAlign::Center),
        wrap_text: true,
        text_rotation: None,
        indent: None,
        shrink_to_fit: false,
    }),
    num_fmt: Some(NumFmtStyle::Custom("#,##0.00".into())),
    protection: Some(ProtectionStyle {
        locked: true,
        hidden: false,
    }),
};
let style_id = wb.add_style(&style)?;
```

**TypeScript:**

```typescript
const styleId = wb.addStyle({
    font: {
        name: "Arial",
        size: 14,
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        color: "#FF0000",
    },
    fill: {
        pattern: "solid",
        fgColor: "#FFFF00",
    },
    border: {
        left: { style: "thin", color: "#000000" },
    },
    alignment: {
        horizontal: "center",
        vertical: "center",
        wrapText: true,
    },
    customNumFmt: "#,##0.00",
    protection: {
        locked: true,
        hidden: false,
    },
});
```

### `set_cell_style` / `get_cell_style`

셀에 스타일 ID를 적용하거나 조회한다.

**Rust:**

```rust
wb.set_cell_style("Sheet1", "A1", style_id)?;
let sid: Option<u32> = wb.get_cell_style("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.setCellStyle("Sheet1", "A1", styleId);
const sid: number | null = wb.getCellStyle("Sheet1", "A1");
```

### 스타일 구성 요소 테이블

#### Font (폰트)

| 속성 | 타입 | 설명 |
|------|------|------|
| `name` | `string?` | 폰트 이름 (예: "Calibri", "Arial") |
| `size` | `f64?` / `number?` | 폰트 크기 (포인트) |
| `bold` | `bool` / `boolean?` | 굵게 |
| `italic` | `bool` / `boolean?` | 기울임 |
| `underline` | `bool` / `boolean?` | 밑줄 |
| `strikethrough` | `bool` / `boolean?` | 취소선 |
| `color` | `StyleColor?` / `string?` | 폰트 색상 |

> TypeScript에서 색상은 문자열로 지정한다: `"#RRGGBB"` (RGB), `"theme:N"` (테마), `"indexed:N"` (인덱스).

#### Fill (채우기)

| 속성 | 타입 | 설명 |
|------|------|------|
| `pattern` | `PatternType` / `string?` | 패턴 종류 |
| `fg_color` / `fgColor` | `StyleColor?` / `string?` | 전경색 |
| `bg_color` / `bgColor` | `StyleColor?` / `string?` | 배경색 |

**PatternType 값:**

| 값 | 설명 |
|----|------|
| `none` | 없음 |
| `solid` | 단색 채우기 |
| `gray125` | 12.5% 회색 |
| `darkGray` | 진한 회색 |
| `mediumGray` | 중간 회색 |
| `lightGray` | 연한 회색 |

#### Border (테두리)

| 속성 | 타입 | 설명 |
|------|------|------|
| `left` | `BorderSideStyle?` | 왼쪽 테두리 |
| `right` | `BorderSideStyle?` | 오른쪽 테두리 |
| `top` | `BorderSideStyle?` | 위쪽 테두리 |
| `bottom` | `BorderSideStyle?` | 아래쪽 테두리 |
| `diagonal` | `BorderSideStyle?` | 대각선 테두리 |

각 `BorderSideStyle`은 `style`과 `color`를 포함한다.

**BorderLineStyle 값:**

| 값 | 설명 |
|----|------|
| `thin` | 가는 선 |
| `medium` | 중간 선 |
| `thick` | 굵은 선 |
| `dashed` | 파선 |
| `dotted` | 점선 |
| `double` | 이중선 |
| `hair` | 머리카락 선 |
| `mediumDashed` | 중간 파선 |
| `dashDot` | 일점쇄선 |
| `mediumDashDot` | 중간 일점쇄선 |
| `dashDotDot` | 이점쇄선 |
| `mediumDashDotDot` | 중간 이점쇄선 |
| `slantDashDot` | 기울어진 일점쇄선 |

#### Alignment (정렬)

| 속성 | 타입 | 설명 |
|------|------|------|
| `horizontal` | `HorizontalAlign?` / `string?` | 가로 정렬 |
| `vertical` | `VerticalAlign?` / `string?` | 세로 정렬 |
| `wrap_text` / `wrapText` | `bool` / `boolean?` | 텍스트 줄바꿈 |
| `text_rotation` / `textRotation` | `u32?` / `number?` | 텍스트 회전 각도 |
| `indent` | `u32?` / `number?` | 들여쓰기 수준 |
| `shrink_to_fit` / `shrinkToFit` | `bool` / `boolean?` | 셀에 맞춰 축소 |

**HorizontalAlign 값:** `general`, `left`, `center`, `right`, `fill`, `justify`, `centerContinuous`, `distributed`

**VerticalAlign 값:** `top`, `center`, `bottom`, `justify`, `distributed`

#### NumFmt (숫자 서식)

Rust에서는 `NumFmtStyle` 열거형을 사용한다:
- `NumFmtStyle::Builtin(id)` -- 내장 서식 ID 사용
- `NumFmtStyle::Custom(code)` -- 사용자 정의 서식 코드

TypeScript에서는 두 가지 방식으로 지정한다:
- `numFmtId: number` -- 내장 서식 ID
- `customNumFmt: string` -- 사용자 정의 서식 코드 (우선 적용)

**주요 내장 숫자 서식 ID:**

| ID | 서식 | 설명 |
|----|------|------|
| 0 | General | 일반 |
| 1 | 0 | 정수 |
| 2 | 0.00 | 소수 2자리 |
| 3 | #,##0 | 천 단위 구분 |
| 4 | #,##0.00 | 천 단위 구분 + 소수 |
| 9 | 0% | 백분율 |
| 10 | 0.00% | 소수 백분율 |
| 11 | 0.00E+00 | 과학적 표기 |
| 14 | m/d/yyyy | 날짜 |
| 15 | d-mmm-yy | 날짜 |
| 20 | h:mm | 시각 |
| 21 | h:mm:ss | 시각(초) |
| 22 | m/d/yyyy h:mm | 날짜+시각 |
| 49 | @ | 텍스트 |

#### Protection (보호)

| 속성 | 타입 | 설명 |
|------|------|------|
| `locked` | `bool` / `boolean?` | 셀 잠금 (기본값: true) |
| `hidden` | `bool` / `boolean?` | 수식 숨기기 |

> 셀 보호는 시트 보호가 활성화된 경우에만 효과가 있다.

---

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

## 16. 틀 고정

특정 행/열을 고정하여 스크롤 시에도 보이도록 하는 기능을 다룬다.

### `set_panes(sheet, cell)` / `setPanes(sheet, cell)`

틀 고정을 설정한다. 셀 참조는 스크롤 가능한 영역의 왼쪽 상단 셀을 나타낸다.

**Rust:**

```rust
// Freeze first row
wb.set_panes("Sheet1", "A2")?;

// Freeze first column
wb.set_panes("Sheet1", "B1")?;

// Freeze first row and first column
wb.set_panes("Sheet1", "B2")?;
```

**TypeScript:**

```typescript
wb.setPanes("Sheet1", "A2");  // freeze row 1
wb.setPanes("Sheet1", "B1");  // freeze column A
wb.setPanes("Sheet1", "B2");  // freeze row 1 + column A
```

### `unset_panes(sheet)` / `unsetPanes(sheet)`

틀 고정을 제거한다.

**Rust:**

```rust
wb.unset_panes("Sheet1")?;
```

**TypeScript:**

```typescript
wb.unsetPanes("Sheet1");
```

### `get_panes(sheet)` / `getPanes(sheet)`

현재 틀 고정 설정을 조회한다. 설정이 없으면 None / null을 반환한다.

**Rust:**

```rust
if let Some(cell) = wb.get_panes("Sheet1")? {
    println!("Frozen at: {}", cell);  // e.g., "B2"
}
```

**TypeScript:**

```typescript
const pane = wb.getPanes("Sheet1");
if (pane) {
    console.log(`Frozen at: ${pane}`);
}
```

---

## 17. 페이지 레이아웃

인쇄 관련 설정을 다룬다. 여백, 용지 크기, 방향, 배율, 머리글/바닥글, 인쇄 옵션, 페이지 나누기를 포함한다.

### 여백 (Margins)

`set_page_margins` / `get_page_margins`로 페이지 여백을 인치 단위로 설정하거나 조회한다.

**Rust:**

```rust
use sheetkit::page_layout::PageMarginsConfig;

wb.set_page_margins("Sheet1", &PageMarginsConfig {
    left: 0.7,
    right: 0.7,
    top: 0.75,
    bottom: 0.75,
    header: 0.3,
    footer: 0.3,
})?;

let margins = wb.get_page_margins("Sheet1")?;
```

**TypeScript:**

```typescript
wb.setPageMargins("Sheet1", {
    left: 0.7,
    right: 0.7,
    top: 0.75,
    bottom: 0.75,
    header: 0.3,
    footer: 0.3,
});

const margins = wb.getPageMargins("Sheet1");
```

### 페이지 설정 (Page Setup)

용지 크기, 방향, 배율, 페이지 맞춤 설정을 다룬다.

**TypeScript:**

```typescript
wb.setPageSetup("Sheet1", {
    paperSize: "a4",
    orientation: "landscape",
    scale: 80,
    fitToWidth: 1,
    fitToHeight: 0,
});

const setup = wb.getPageSetup("Sheet1");
```

**용지 크기 값:** `letter`, `tabloid`, `legal`, `a3`, `a4`, `a5`, `b4`, `b5`

**방향 값:** `portrait` (세로), `landscape` (가로)

### 인쇄 옵션 (Print Options)

**TypeScript:**

```typescript
wb.setPrintOptions("Sheet1", {
    gridLines: true,
    headings: true,
    horizontalCentered: true,
    verticalCentered: false,
});

const opts = wb.getPrintOptions("Sheet1");
```

| 속성 | 타입 | 설명 |
|------|------|------|
| `grid_lines` / `gridLines` | `bool?` / `boolean?` | 눈금선 인쇄 |
| `headings` | `bool?` / `boolean?` | 행/열 머리글 인쇄 |
| `horizontal_centered` / `horizontalCentered` | `bool?` / `boolean?` | 가로 가운데 정렬 |
| `vertical_centered` / `verticalCentered` | `bool?` / `boolean?` | 세로 가운데 정렬 |

### 머리글/바닥글 (Header/Footer)

```typescript
wb.setHeaderFooter("Sheet1", "&LLeft Text&CCenter Text&RRight Text", "&CPage &P of &N");

const hf = wb.getHeaderFooter("Sheet1");
// hf.header, hf.footer
```

> Excel 서식 코드: `&L` (왼쪽), `&C` (가운데), `&R` (오른쪽), `&P` (현재 페이지), `&N` (총 페이지 수)

### 페이지 나누기 (Page Breaks)

```typescript
wb.insertPageBreak("Sheet1", 20);  // insert break before row 20
wb.insertPageBreak("Sheet1", 40);

const breaks: number[] = wb.getPageBreaks("Sheet1");
// [20, 40]

wb.removePageBreak("Sheet1", 20);
```

**Rust:**

```rust
wb.insert_page_break("Sheet1", 20)?;
let breaks: Vec<u32> = wb.get_page_breaks("Sheet1")?;
wb.remove_page_break("Sheet1", 20)?;
```

---

## 18. 정의된 이름

워크북 내에서 셀 범위에 이름을 부여하는 기능을 다룬다. 워크북 범위 또는 시트 범위로 정의할 수 있다.

> 현재 이 기능은 Rust 코어 모듈(`sheetkit_core::defined_names`)에서 함수로 제공되며, Workbook 구조체 메서드 및 Node.js 바인딩에는 아직 노출되지 않았다.

### Rust (모듈 함수 직접 사용)

```rust
use sheetkit_core::defined_names::*;

// set_defined_name: add or update a defined name
set_defined_name(
    &mut workbook_xml,
    "SalesData",                   // name
    "Sheet1!$A$1:$D$100",         // reference
    DefinedNameScope::Workbook,    // scope
    Some("Monthly sales data"),    // comment
)?;

// get_defined_name: retrieve a defined name
let info = get_defined_name(&workbook_xml, "SalesData", DefinedNameScope::Workbook)?;
// info.name, info.value, info.scope, info.comment

// delete_defined_name: remove a defined name
delete_defined_name(&mut workbook_xml, "SalesData", DefinedNameScope::Workbook)?;
```

### DefinedNameScope

| 값 | 설명 |
|----|------|
| `Workbook` | 워크북 전체에서 사용 가능 |
| `Sheet(index)` | 특정 시트(0부터 시작하는 인덱스)에서만 사용 가능 |

> 정의된 이름에는 `\ / ? * [ ]` 문자를 사용할 수 없으며, 앞뒤 공백도 허용되지 않는다.

---

## 19. 문서 속성

워크북의 메타데이터를 설정하고 조회하는 기능을 다룬다. 핵심 속성, 앱 속성, 사용자 정의 속성의 세 가지 유형이 있다.

### 핵심 속성 (Core Properties)

제목, 작성자 등 표준 문서 메타데이터를 다룬다.

**Rust:**

```rust
use sheetkit::doc_props::DocProperties;

wb.set_doc_props(DocProperties {
    title: Some("Annual Report".into()),
    subject: Some("Financial Data".into()),
    creator: Some("Finance Team".into()),
    keywords: Some("finance, annual, 2024".into()),
    description: Some("Annual financial report".into()),
    last_modified_by: Some("Admin".into()),
    revision: Some("3".into()),
    created: Some("2024-01-01T00:00:00Z".into()),
    modified: Some("2024-06-15T10:30:00Z".into()),
    category: Some("Reports".into()),
    content_status: Some("Final".into()),
});

let props = wb.get_doc_props();
```

**TypeScript:**

```typescript
wb.setDocProps({
    title: "Annual Report",
    subject: "Financial Data",
    creator: "Finance Team",
    keywords: "finance, annual, 2024",
    description: "Annual financial report",
    lastModifiedBy: "Admin",
    revision: "3",
    created: "2024-01-01T00:00:00Z",
    modified: "2024-06-15T10:30:00Z",
    category: "Reports",
    contentStatus: "Final",
});

const props = wb.getDocProps();
```

**DocProperties 속성:**

| 속성 | 타입 | 설명 |
|------|------|------|
| `title` | `string?` | 제목 |
| `subject` | `string?` | 주제 |
| `creator` | `string?` | 작성자 |
| `keywords` | `string?` | 키워드 |
| `description` | `string?` | 설명 |
| `last_modified_by` / `lastModifiedBy` | `string?` | 마지막 수정자 |
| `revision` | `string?` | 수정 번호 |
| `created` | `string?` | 생성 날짜 (ISO 8601) |
| `modified` | `string?` | 수정 날짜 (ISO 8601) |
| `category` | `string?` | 분류 |
| `content_status` / `contentStatus` | `string?` | 콘텐츠 상태 |

### 앱 속성 (App Properties)

애플리케이션 관련 메타데이터를 다룬다.

**Rust:**

```rust
use sheetkit::doc_props::AppProperties;

wb.set_app_props(AppProperties {
    application: Some("SheetKit".into()),
    doc_security: Some(0),
    company: Some("ACME Corp".into()),
    app_version: Some("1.0".into()),
    manager: Some("Department Lead".into()),
    template: None,
});

let app_props = wb.get_app_props();
```

**TypeScript:**

```typescript
wb.setAppProps({
    application: "SheetKit",
    docSecurity: 0,
    company: "ACME Corp",
    appVersion: "1.0",
    manager: "Department Lead",
});

const appProps = wb.getAppProps();
```

**AppProperties 속성:**

| 속성 | 타입 | 설명 |
|------|------|------|
| `application` | `string?` | 애플리케이션 이름 |
| `doc_security` / `docSecurity` | `u32?` / `number?` | 보안 수준 |
| `company` | `string?` | 회사 이름 |
| `app_version` / `appVersion` | `string?` | 앱 버전 |
| `manager` | `string?` | 관리자 |
| `template` | `string?` | 템플릿 |

### 사용자 정의 속성 (Custom Properties)

키-값 쌍으로 사용자 정의 메타데이터를 저장한다. 값은 문자열, 숫자, 불리언 타입을 지원한다.

**Rust:**

```rust
use sheetkit::doc_props::CustomPropertyValue;

wb.set_custom_property("Department", CustomPropertyValue::String("Engineering".into()));
wb.set_custom_property("Version", CustomPropertyValue::Int(3));
wb.set_custom_property("Approved", CustomPropertyValue::Bool(true));
wb.set_custom_property("Rating", CustomPropertyValue::Float(4.5));

let val = wb.get_custom_property("Department");
// Some(CustomPropertyValue::String("Engineering"))

let deleted = wb.delete_custom_property("Deprecated");
// true if existed
```

**TypeScript:**

```typescript
wb.setCustomProperty("Department", "Engineering");
wb.setCustomProperty("Version", 3);
wb.setCustomProperty("Approved", true);
wb.setCustomProperty("Rating", 4.5);

const val = wb.getCustomProperty("Department");
// "Engineering"

const deleted: boolean = wb.deleteCustomProperty("Deprecated");
```

> TypeScript에서 정수 값은 자동으로 Int로 변환되고, 소수점이 있는 숫자는 Float으로 저장된다.

---

## 20. 워크북 보호

워크북 구조(시트 추가/삭제/이름 변경)와 창 위치를 보호하는 기능을 다룬다. 선택적으로 비밀번호를 설정할 수 있다.

### `protect_workbook(config)` / `protectWorkbook(config)`

워크북 보호를 설정한다.

**Rust:**

```rust
use sheetkit::protection::WorkbookProtectionConfig;

wb.protect_workbook(WorkbookProtectionConfig {
    password: Some("secret123".into()),
    lock_structure: true,
    lock_windows: false,
    lock_revision: false,
});
```

**TypeScript:**

```typescript
wb.protectWorkbook({
    password: "secret123",
    lockStructure: true,
    lockWindows: false,
    lockRevision: false,
});
```

### `unprotect_workbook()` / `unprotectWorkbook()`

워크북 보호를 해제한다.

**Rust:**

```rust
wb.unprotect_workbook();
```

**TypeScript:**

```typescript
wb.unprotectWorkbook();
```

### `is_workbook_protected()` / `isWorkbookProtected()`

워크북이 보호되어 있는지 확인한다.

**Rust:**

```rust
let protected: bool = wb.is_workbook_protected();
```

**TypeScript:**

```typescript
const isProtected: boolean = wb.isWorkbookProtected();
```

### WorkbookProtectionConfig 속성

| 속성 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `password` | `string?` | `None` | 보호 비밀번호 (레거시 해시로 저장) |
| `lock_structure` / `lockStructure` | `bool` / `boolean?` | `false` | 시트 추가/삭제/이름 변경 차단 |
| `lock_windows` / `lockWindows` | `bool` / `boolean?` | `false` | 창 위치/크기 변경 차단 |
| `lock_revision` / `lockRevision` | `bool` / `boolean?` | `false` | 수정 추적 잠금 |

> 비밀번호는 Excel의 레거시 16비트 해시 알고리즘으로 저장된다. 이 해시는 암호학적으로 안전하지 않다.

---

## 21. 시트 보호

개별 시트의 편집을 제한하는 기능을 다룬다. 선택적으로 특정 작업을 허용할 수 있다.

> 현재 이 기능은 Rust 코어 모듈(`sheetkit_core::sheet`)에서 함수로 제공되며, Workbook 구조체 메서드 및 Node.js 바인딩에는 아직 노출되지 않았다.

### Rust (모듈 함수 직접 사용)

```rust
use sheetkit_core::sheet::{protect_sheet, unprotect_sheet, is_sheet_protected, SheetProtectionConfig};

protect_sheet(&mut worksheet_xml, &SheetProtectionConfig {
    password: Some("pass".into()),
    select_locked_cells: true,
    select_unlocked_cells: true,
    format_cells: false,
    format_columns: false,
    format_rows: false,
    insert_columns: false,
    insert_rows: false,
    insert_hyperlinks: false,
    delete_columns: false,
    delete_rows: false,
    sort: false,
    auto_filter: false,
    pivot_tables: false,
})?;

let protected: bool = is_sheet_protected(&worksheet_xml);

unprotect_sheet(&mut worksheet_xml)?;
```

### SheetProtectionConfig 속성

| 속성 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `password` | `Option<String>` | `None` | 보호 비밀번호 |
| `select_locked_cells` | `bool` | `false` | 잠긴 셀 선택 허용 |
| `select_unlocked_cells` | `bool` | `false` | 잠기지 않은 셀 선택 허용 |
| `format_cells` | `bool` | `false` | 셀 서식 변경 허용 |
| `format_columns` | `bool` | `false` | 열 서식 변경 허용 |
| `format_rows` | `bool` | `false` | 행 서식 변경 허용 |
| `insert_columns` | `bool` | `false` | 열 삽입 허용 |
| `insert_rows` | `bool` | `false` | 행 삽입 허용 |
| `insert_hyperlinks` | `bool` | `false` | 하이퍼링크 삽입 허용 |
| `delete_columns` | `bool` | `false` | 열 삭제 허용 |
| `delete_rows` | `bool` | `false` | 행 삭제 허용 |
| `sort` | `bool` | `false` | 정렬 허용 |
| `auto_filter` | `bool` | `false` | 자동 필터 사용 허용 |
| `pivot_tables` | `bool` | `false` | 피벗 테이블 사용 허용 |

---

## 22. 수식 평가

셀 수식을 파싱하고 실행하는 기능을 다룬다. nom 파서로 수식을 AST로 변환한 후 평가 엔진이 결과를 계산한다.

### `evaluate_formula(sheet, formula)` / `evaluateFormula(sheet, formula)`

주어진 시트 컨텍스트에서 수식 문자열을 평가하여 결과를 반환한다. 워크북의 현재 셀 데이터를 참조할 수 있다.

**Rust:**

```rust
wb.set_cell_value("Sheet1", "A1", CellValue::Number(10.0))?;
wb.set_cell_value("Sheet1", "A2", CellValue::Number(20.0))?;

let result = wb.evaluate_formula("Sheet1", "SUM(A1:A2)")?;
// CellValue::Number(30.0)
```

**TypeScript:**

```typescript
wb.setCellValue("Sheet1", "A1", 10);
wb.setCellValue("Sheet1", "A2", 20);

const result = wb.evaluateFormula("Sheet1", "SUM(A1:A2)");
// 30
```

### `calculate_all()` / `calculateAll()`

워크북의 모든 수식 셀을 재계산한다. 의존성 그래프를 구축하고 위상 정렬을 수행하여 올바른 순서로 평가한다.

**Rust:**

```rust
wb.calculate_all()?;
```

**TypeScript:**

```typescript
wb.calculateAll();
```

> 순환 참조가 발견되면 오류가 발생한다. 최대 재귀 깊이는 256이다.

### 지원 함수 목록 (110개, 8개 카테고리)

#### 수학 함수 (Math) -- 20개

| 함수 | 설명 |
|------|------|
| `SUM` | 합계 |
| `PRODUCT` | 곱 |
| `ABS` | 절대값 |
| `INT` | 정수 변환 (내림) |
| `MOD` | 나머지 |
| `POWER` | 거듭제곱 |
| `SQRT` | 제곱근 |
| `ROUND` | 반올림 |
| `ROUNDUP` | 올림 |
| `ROUNDDOWN` | 내림 |
| `CEILING` | 올림 (배수) |
| `FLOOR` | 내림 (배수) |
| `SIGN` | 부호 |
| `RAND` | 난수 (0-1) |
| `RANDBETWEEN` | 정수 난수 (범위) |
| `PI` | 원주율 |
| `LOG` | 로그 |
| `LOG10` | 상용 로그 |
| `LN` | 자연 로그 |
| `EXP` | 지수 함수 |
| `QUOTIENT` | 정수 몫 |
| `FACT` | 팩토리얼 |
| `SUMIF` | 조건부 합계 |
| `SUMIFS` | 다중 조건부 합계 |

#### 통계 함수 (Statistical) -- 16개

| 함수 | 설명 |
|------|------|
| `AVERAGE` | 평균 |
| `COUNT` | 숫자 셀 개수 |
| `COUNTA` | 비어 있지 않은 셀 개수 |
| `COUNTBLANK` | 빈 셀 개수 |
| `COUNTIF` | 조건부 개수 |
| `COUNTIFS` | 다중 조건부 개수 |
| `MIN` | 최소값 |
| `MAX` | 최대값 |
| `MEDIAN` | 중앙값 |
| `MODE` | 최빈값 |
| `LARGE` | N번째 큰 값 |
| `SMALL` | N번째 작은 값 |
| `RANK` | 순위 |
| `AVERAGEIF` | 조건부 평균 |
| `AVERAGEIFS` | 다중 조건부 평균 |

#### 논리 함수 (Logical) -- 10개

| 함수 | 설명 |
|------|------|
| `IF` | 조건 분기 |
| `AND` | 논리곱 |
| `OR` | 논리합 |
| `NOT` | 부정 |
| `XOR` | 배타적 논리합 |
| `TRUE` | TRUE 상수 |
| `FALSE` | FALSE 상수 |
| `IFERROR` | 오류 시 대체값 |
| `IFNA` | #N/A 시 대체값 |
| `IFS` | 다중 조건 분기 |
| `SWITCH` | 값 기반 분기 |

#### 텍스트 함수 (Text) -- 15개

| 함수 | 설명 |
|------|------|
| `LEN` | 문자열 길이 |
| `LOWER` | 소문자 변환 |
| `UPPER` | 대문자 변환 |
| `TRIM` | 공백 제거 |
| `LEFT` | 왼쪽 문자 추출 |
| `RIGHT` | 오른쪽 문자 추출 |
| `MID` | 중간 문자 추출 |
| `CONCATENATE` | 문자열 연결 |
| `CONCAT` | 문자열 연결 (최신) |
| `FIND` | 문자열 찾기 (대소문자 구분) |
| `SEARCH` | 문자열 찾기 (대소문자 무시) |
| `SUBSTITUTE` | 문자열 치환 |
| `REPLACE` | 위치 기반 문자열 교체 |
| `REPT` | 문자열 반복 |
| `EXACT` | 완전 일치 비교 |
| `T` | 텍스트 변환 |
| `PROPER` | 단어 첫 글자 대문자 |

#### 정보 함수 (Information) -- 11개

| 함수 | 설명 |
|------|------|
| `ISNUMBER` | 숫자 여부 |
| `ISTEXT` | 텍스트 여부 |
| `ISBLANK` | 빈 셀 여부 |
| `ISERROR` | 오류 여부 (#N/A 포함) |
| `ISERR` | 오류 여부 (#N/A 제외) |
| `ISNA` | #N/A 여부 |
| `ISLOGICAL` | 논리값 여부 |
| `ISEVEN` | 짝수 여부 |
| `ISODD` | 홀수 여부 |
| `TYPE` | 값 유형 번호 |
| `N` | 숫자 변환 |
| `NA` | #N/A 생성 |
| `ERROR.TYPE` | 오류 유형 번호 |

#### 변환 함수 (Conversion) -- 2개

| 함수 | 설명 |
|------|------|
| `VALUE` | 텍스트를 숫자로 변환 |
| `TEXT` | 값을 서식 문자열로 변환 |

#### 날짜/시각 함수 (Date/Time) -- 17개

| 함수 | 설명 |
|------|------|
| `DATE` | 연/월/일로 날짜 생성 |
| `TODAY` | 오늘 날짜 |
| `NOW` | 현재 날짜 및 시각 |
| `YEAR` | 연도 추출 |
| `MONTH` | 월 추출 |
| `DAY` | 일 추출 |
| `HOUR` | 시 추출 |
| `MINUTE` | 분 추출 |
| `SECOND` | 초 추출 |
| `DATEDIF` | 날짜 차이 계산 |
| `EDATE` | N개월 후 날짜 |
| `EOMONTH` | N개월 후 월말 |
| `DATEVALUE` | 텍스트를 날짜로 변환 |
| `WEEKDAY` | 요일 번호 |
| `WEEKNUM` | 주차 번호 |
| `NETWORKDAYS` | 근무일수 계산 |
| `WORKDAY` | N 근무일 후 날짜 |

#### 찾기/참조 함수 (Lookup) -- 11개

| 함수 | 설명 |
|------|------|
| `VLOOKUP` | 세로 방향 조회 |
| `HLOOKUP` | 가로 방향 조회 |
| `INDEX` | 범위에서 값 추출 |
| `MATCH` | 위치 찾기 |
| `LOOKUP` | 벡터 조회 |
| `ROW` | 행 번호 |
| `COLUMN` | 열 번호 |
| `ROWS` | 범위의 행 수 |
| `COLUMNS` | 범위의 열 수 |
| `CHOOSE` | 인덱스로 값 선택 |
| `ADDRESS` | 셀 주소 문자열 생성 |

---

## 23. 피벗 테이블

피벗 테이블을 생성, 조회, 삭제하는 기능을 다룬다. 소스 데이터 범위로부터 행/열/데이터 필드를 지정하여 피벗 테이블을 구성한다.

### `add_pivot_table(config)` / `addPivotTable(config)`

피벗 테이블을 추가한다.

**Rust:**

```rust
use sheetkit::pivot::*;

wb.add_pivot_table(&PivotTableConfig {
    name: "SalesPivot".into(),
    source_sheet: "RawData".into(),
    source_range: "A1:D100".into(),
    target_sheet: "PivotSheet".into(),
    target_cell: "A3".into(),
    rows: vec![
        PivotField { name: "Region".into() },
        PivotField { name: "Product".into() },
    ],
    columns: vec![
        PivotField { name: "Quarter".into() },
    ],
    data: vec![
        PivotDataField {
            name: "Revenue".into(),
            function: AggregateFunction::Sum,
            display_name: Some("Total Revenue".into()),
        },
        PivotDataField {
            name: "Quantity".into(),
            function: AggregateFunction::Count,
            display_name: None,
        },
    ],
})?;
```

**TypeScript:**

```typescript
wb.addPivotTable({
    name: "SalesPivot",
    sourceSheet: "RawData",
    sourceRange: "A1:D100",
    targetSheet: "PivotSheet",
    targetCell: "A3",
    rows: [
        { name: "Region" },
        { name: "Product" },
    ],
    columns: [
        { name: "Quarter" },
    ],
    data: [
        { name: "Revenue", function: "sum", displayName: "Total Revenue" },
        { name: "Quantity", function: "count" },
    ],
});
```

### `get_pivot_tables()` / `getPivotTables()`

워크북의 모든 피벗 테이블 정보를 반환한다.

**Rust:**

```rust
let tables = wb.get_pivot_tables();
for t in &tables {
    println!("{}: {}!{} -> {}!{}", t.name, t.source_sheet, t.source_range, t.target_sheet, t.location);
}
```

**TypeScript:**

```typescript
const tables = wb.getPivotTables();
for (const t of tables) {
    console.log(`${t.name}: ${t.sourceSheet}!${t.sourceRange} -> ${t.targetSheet}!${t.location}`);
}
```

**PivotTableInfo 구조:**

| 속성 | 타입 | 설명 |
|------|------|------|
| `name` | `string` | 피벗 테이블 이름 |
| `source_sheet` / `sourceSheet` | `string` | 소스 데이터 시트 이름 |
| `source_range` / `sourceRange` | `string` | 소스 데이터 범위 |
| `target_sheet` / `targetSheet` | `string` | 대상 시트 이름 |
| `location` | `string` | 피벗 테이블 위치 |

### `delete_pivot_table(name)` / `deletePivotTable(name)`

이름으로 피벗 테이블을 삭제한다.

**Rust:**

```rust
wb.delete_pivot_table("SalesPivot")?;
```

**TypeScript:**

```typescript
wb.deletePivotTable("SalesPivot");
```

### PivotTableConfig 구조

| 속성 | 타입 | 설명 |
|------|------|------|
| `name` | `string` | 피벗 테이블 이름 |
| `source_sheet` / `sourceSheet` | `string` | 소스 데이터가 있는 시트 |
| `source_range` / `sourceRange` | `string` | 소스 데이터 범위 (예: "A1:D100") |
| `target_sheet` / `targetSheet` | `string` | 피벗 테이블을 배치할 시트 |
| `target_cell` / `targetCell` | `string` | 피벗 테이블 시작 셀 |
| `rows` | `PivotField[]` | 행 필드 |
| `columns` | `PivotField[]` | 열 필드 |
| `data` | `PivotDataField[]` | 데이터(값) 필드 |

### AggregateFunction (집계 함수)

| 값 | 설명 |
|----|------|
| `sum` | 합계 |
| `count` | 개수 |
| `average` | 평균 |
| `max` | 최대값 |
| `min` | 최소값 |
| `product` | 곱 |
| `countNums` | 숫자 개수 |
| `stdDev` | 표준편차 |
| `stdDevP` | 모표준편차 |
| `var` | 분산 |
| `varP` | 모분산 |

---

## 24. 스트림 라이터

대용량 데이터를 메모리 효율적으로 쓰기 위한 스트리밍 API이다. 행은 오름차순으로만 쓸 수 있으며, 전체 워크시트 XML을 메모리에 구축하지 않고 직접 버퍼에 기록한다.

### 사용 흐름

1. `new_stream_writer`로 스트림 라이터 생성
2. 열 너비, 셀 병합 등 설정
3. `write_row`로 행 데이터를 순서대로 기록
4. `apply_stream_writer`로 워크북에 적용

**Rust:**

```rust
use sheetkit::cell::CellValue;

let mut sw = wb.new_stream_writer("LargeData")?;

// Set column widths
sw.set_col_width(1, 15.0)?;   // Column A
sw.set_col_width(2, 20.0)?;   // Column B

// Add merge cells
sw.add_merge_cell("A1:B1")?;

// Write header row
sw.write_row(1, &[
    CellValue::from("Name"),
    CellValue::from("Value"),
])?;

// Write data rows (must be in ascending order)
for i in 2..=10000 {
    sw.write_row(i, &[
        CellValue::from(format!("Item {}", i - 1)),
        CellValue::Number(i as f64 * 1.5),
    ])?;
}

// Apply to workbook
let sheet_index = wb.apply_stream_writer(sw)?;
wb.save("large_data.xlsx")?;
```

**TypeScript:**

```typescript
const sw = wb.newStreamWriter("LargeData");

// Set column widths
sw.setColWidth(1, 15);   // Column A
sw.setColWidth(2, 20);   // Column B

// Set width for a range of columns
sw.setColWidthRange(3, 10, 12);  // Columns C-J

// Add merge cells
sw.addMergeCell("A1:B1");

// Write header row
sw.writeRow(1, ["Name", "Value"]);

// Write data rows
for (let i = 2; i <= 10000; i++) {
    sw.writeRow(i, [`Item ${i - 1}`, i * 1.5]);
}

// Apply to workbook
const sheetIndex: number = wb.applyStreamWriter(sw);
wb.save("large_data.xlsx");
```

### StreamWriter API

#### `new_stream_writer(sheet_name)` / `newStreamWriter(sheetName)`

새 시트를 위한 스트림 라이터를 생성한다.

#### `write_row(row, values)` / `writeRow(row, values)`

행 데이터를 기록한다. 행 번호는 1부터 시작하며 반드시 오름차순이어야 한다.

#### `set_col_width(col, width)` / `setColWidth(col, width)`

열 너비를 설정한다. `col`은 1부터 시작하는 열 번호이다.

#### `set_col_width_range(min_col, max_col, width)` / `setColWidthRange(minCol, maxCol, width)`

열 범위의 너비를 한 번에 설정한다.

#### `add_merge_cell(reference)` / `addMergeCell(reference)`

셀 병합을 추가한다 (예: "A1:C3").

#### `apply_stream_writer(writer)` / `applyStreamWriter(writer)`

스트림 라이터의 결과를 워크북에 적용한다. 시트 인덱스를 반환한다. 적용 후 스트림 라이터는 소비(consumed)되어 더 이상 사용할 수 없다.

### StreamRowOptions (Rust 전용)

Rust에서는 `write_row_with_options`를 사용하여 행별 옵션을 지정할 수 있다.

| 속성 | 타입 | 설명 |
|------|------|------|
| `height` | `Option<f64>` | 행 높이 (포인트) |
| `visible` | `Option<bool>` | 행 표시 여부 |
| `outline_level` | `Option<u8>` | 아웃라인 수준 (0-7) |
| `style_id` | `Option<u32>` | 행 스타일 ID |

---

## 25. 유틸리티 함수

셀 참조 변환에 사용되는 유틸리티 함수들이다. Rust에서는 `sheetkit_core::utils::cell_ref` 모듈에서 제공한다.

### `cell_name_to_coordinates`

A1 형식의 셀 참조를 (열, 행) 좌표로 변환한다. 열과 행 모두 1부터 시작한다.

**Rust:**

```rust
use sheetkit_core::utils::cell_ref::cell_name_to_coordinates;

let (col, row) = cell_name_to_coordinates("B3")?;
// col = 2, row = 3

let (col, row) = cell_name_to_coordinates("$AB$100")?;
// col = 28, row = 100 (absolute references are supported)
```

### `coordinates_to_cell_name`

(열, 행) 좌표를 A1 형식의 셀 참조로 변환한다.

**Rust:**

```rust
use sheetkit_core::utils::cell_ref::coordinates_to_cell_name;

let name = coordinates_to_cell_name(2, 3)?;
// "B3"

let name = coordinates_to_cell_name(28, 100)?;
// "AB100"
```

### `column_name_to_number`

열 이름을 1부터 시작하는 열 번호로 변환한다.

**Rust:**

```rust
use sheetkit_core::utils::cell_ref::column_name_to_number;

assert_eq!(column_name_to_number("A")?, 1);
assert_eq!(column_name_to_number("Z")?, 26);
assert_eq!(column_name_to_number("AA")?, 27);
assert_eq!(column_name_to_number("XFD")?, 16384);  // maximum column
```

### `column_number_to_name`

1부터 시작하는 열 번호를 열 이름으로 변환한다.

**Rust:**

```rust
use sheetkit_core::utils::cell_ref::column_number_to_name;

assert_eq!(column_number_to_name(1)?, "A");
assert_eq!(column_number_to_name(26)?, "Z");
assert_eq!(column_number_to_name(27)?, "AA");
assert_eq!(column_number_to_name(16384)?, "XFD");
```

> 유틸리티 함수는 현재 Rust 전용으로 제공된다. TypeScript에서는 문자열 기반 셀 참조("A1", "B2" 등)를 직접 사용한다.

---

## 부록: 제한 사항

| 항목 | 제한 |
|------|------|
| 최대 열 수 | 16,384 (XFD) |
| 최대 행 수 | 1,048,576 |
| 최대 셀 문자 수 | 32,767 |
| 최대 행 높이 | 409 포인트 |
| 최대 아웃라인 수준 | 7 |
| 최대 스타일 XF 수 | 65,430 |
| 수식 최대 재귀 깊이 | 256 |
| 지원 수식 함수 수 | 110 / 456 |
