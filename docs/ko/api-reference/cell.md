## 셀 조작

셀 값의 읽기와 쓰기를 다룹니다. 셀 값은 다양한 타입으로 표현됩니다.

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
| 서식 있는 텍스트 | `CellValue::RichString(Vec<RichTextRun>)` | `string` (연결된 텍스트) | 여러 서식 run으로 구성된 텍스트 |

### DateValue (TypeScript)

날짜 셀을 표현하는 객체입니다.

```typescript
interface DateValue {
  type: "date";     // always "date"
  serial: number;   // Excel serial number (1 = 1900-01-01)
  iso?: string;     // ISO format string ("2024-01-15" or "2024-01-15T14:30:00")
}
```

> Excel은 내부적으로 날짜를 시리얼 번호(정수부: 날짜, 소수부: 시각)로 저장합니다. 1900년 윤년 버그가 포함되어 있으며 SheetKit은 이를 올바르게 처리합니다.

### `get_cell_value(sheet, cell)` / `getCellValue(sheet, cell)`

지정한 셀의 값을 읽습니다.

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

셀에 값을 설정합니다. null을 전달하면 셀이 비워집니다.

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

> 시트 이름이 존재하지 않거나 셀 참조가 유효하지 않으면 오류가 발생합니다.

### `get_cell_formatted_value(sheet, cell)` / `getCellFormattedValue(sheet, cell)`

셀의 표시 텍스트를 반환합니다. 숫자 셀에 서식 스타일(날짜, 백분율, 천 단위 구분 등)이 적용되어 있으면 해당 서식 코드를 통해 렌더링됩니다. 문자열 셀은 텍스트를 그대로 반환합니다. 빈 셀은 빈 문자열을 반환합니다.

**Rust:**

```rust
let formatted = wb.get_cell_formatted_value("Sheet1", "A1")?;
// "1,234.50" -- #,##0.00 서식이 적용된 숫자
// "2024-12-31" -- yyyy-mm-dd 서식이 적용된 날짜 시리얼
```

**TypeScript:**

```typescript
const text = wb.getCellFormattedValue("Sheet1", "A1");
// "1,234.50", "2024-12-31", "85.00%" 등
```

### `format_number(value, format_code)` / `formatNumber(value, formatCode)`

숫자 값을 Excel 서식 코드 문자열로 포맷하는 독립 유틸리티 함수입니다. Workbook 인스턴스가 필요하지 않습니다.

**Rust:**

```rust
use sheetkit::format_number;

assert_eq!(format_number(1234.5, "#,##0.00"), "1,234.50");
assert_eq!(format_number(0.85, "0.00%"), "85.00%");
assert_eq!(format_number(45657.0, "yyyy-mm-dd"), "2024-12-31");
assert_eq!(format_number(0.5, "h:mm AM/PM"), "12:00 PM");
```

**TypeScript:**

```typescript
import { formatNumber } from "sheetkit";

formatNumber(1234.5, "#,##0.00");   // "1,234.50"
formatNumber(0.85, "0.00%");        // "85.00%"
formatNumber(45657, "yyyy-mm-dd");  // "2024-12-31"
formatNumber(0.5, "h:mm AM/PM");    // "12:00 PM"
```

지원하는 서식 기능은 다음과 같습니다.

| 기능 | 예시 | 설명 |
|---|---|---|
| General | `General` | 기본 표시 서식입니다 |
| 정수 / 소수 | `0`, `0.00`, `#,##0.00` | 천 단위 구분 기호를 포함한 숫자 서식입니다 |
| 백분율 | `0%`, `0.00%` | 100을 곱하고 % 기호를 추가합니다 |
| 지수 | `0.00E+00` | 지수 표기법입니다 |
| 날짜 | `m/d/yyyy`, `yyyy-mm-dd`, `d-mmm-yy` | Excel 시리얼 번호에서 날짜를 렌더링합니다 |
| 시간 | `h:mm`, `h:mm:ss`, `h:mm AM/PM` | 소수부에서 시각을 렌더링합니다 |
| 날짜 + 시간 | `m/d/yyyy h:mm` | 날짜와 시간을 함께 표시합니다 |
| 분수 | `# ?/?`, `# ??/??` | 분수 표현입니다 |
| 다중 섹션 | `pos;neg;zero;text` | `;`로 구분된 최대 4개 섹션입니다 |
| 색상 코드 | `[Red]0.00` | 색상 태그를 파싱하고 제거합니다 |
| 리터럴 텍스트 | `"text"`, `\x` | 따옴표로 둘러싼 문자열과 이스케이프 문자입니다 |

### `builtin_format_code(id)` / `builtinFormatCode(id)`

기본 제공 서식 ID(0-49)에 대한 서식 코드 문자열을 조회합니다. 인식되지 않는 ID에 대해서는 `None`/`null`을 반환합니다.

**Rust:**

```rust
use sheetkit::builtin_format_code;

assert_eq!(builtin_format_code(0), Some("General"));
assert_eq!(builtin_format_code(14), Some("m/d/yyyy"));
assert_eq!(builtin_format_code(100), None);
```

**TypeScript:**

```typescript
import { builtinFormatCode } from "sheetkit";

builtinFormatCode(0);   // "General"
builtinFormatCode(14);  // "m/d/yyyy"
builtinFormatCode(100); // null
```

### `get_occupied_cells(sheet)` (Rust 전용)

시트에서 값이 있는 모든 셀의 `(col, row)` 좌표 쌍을 반환합니다. 두 값 모두 1부터 시작합니다. 전체 그리드를 탐색하지 않고 데이터가 존재하는 셀만 순회할 때 유용합니다.

```rust
let cells = wb.get_occupied_cells("Sheet1")?;
for (col, row) in &cells {
    println!("Cell at col {}, row {}", col, row);
}
```

---
