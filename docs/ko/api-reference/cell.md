## 2. 셀 조작

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

### `get_occupied_cells(sheet)` (Rust 전용)

시트에서 값이 있는 모든 셀의 `(col, row)` 좌표 쌍을 반환합니다. 두 값 모두 1부터 시작합니다. 전체 그리드를 탐색하지 않고 데이터가 존재하는 셀만 순회할 때 유용합니다.

```rust
let cells = wb.get_occupied_cells("Sheet1")?;
for (col, row) in &cells {
    println!("Cell at col {}, row {}", col, row);
}
```

---
