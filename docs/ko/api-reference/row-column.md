## 행 조작

행의 삽입, 삭제, 복제 및 높이/가시성/아웃라인/스타일 설정을 다룹니다. 모든 행 번호는 1부터 시작합니다.

### `insert_rows(sheet, start_row, count)` / `insertRows(sheet, startRow, count)`

지정한 행 번호부터 빈 행을 삽입합니다.

**Rust:**

```rust
wb.insert_rows("Sheet1", 3, 2)?; // 3rd row onwards, insert 2 rows
```

**TypeScript:**

```typescript
wb.insertRows("Sheet1", 3, 2);
```

### `remove_row(sheet, row)` / `removeRow(sheet, row)`

행을 삭제합니다. 아래쪽 행이 위로 이동합니다.

**Rust:**

```rust
wb.remove_row("Sheet1", 5)?;
```

**TypeScript:**

```typescript
wb.removeRow("Sheet1", 5);
```

### `duplicate_row(sheet, row)` / `duplicateRow(sheet, row)`

지정한 행을 바로 아래에 복제합니다.

**Rust:**

```rust
wb.duplicate_row("Sheet1", 2)?;
```

**TypeScript:**

```typescript
wb.duplicateRow("Sheet1", 2);
```

### `set_row_height` / `get_row_height`

행 높이를 포인트 단위로 설정하거나 조회합니다.

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

행의 표시/숨김을 설정하거나 조회합니다.

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

행의 아웃라인(그룹) 수준을 설정하거나 조회합니다. 범위는 0-7입니다.

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

행 전체에 스타일 ID를 적용하거나 조회합니다.

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

## 열 조작

열의 너비, 가시성, 삽입, 삭제 및 아웃라인/스타일 설정을 다룹니다. 열은 문자열("A", "B", "AA" 등)로 지정합니다.

### `set_col_width` / `get_col_width`

열 너비를 문자 단위로 설정하거나 조회합니다.

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

열의 표시/숨김을 설정하거나 조회합니다.

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

지정한 열부터 빈 열을 삽입합니다.

**Rust:**

```rust
wb.insert_cols("Sheet1", "C", 3)?;
```

**TypeScript:**

```typescript
wb.insertCols("Sheet1", "C", 3);
```

### `remove_col(sheet, col)` / `removeCol(sheet, col)`

열을 삭제합니다. 오른쪽 열이 왼쪽으로 이동합니다.

**Rust:**

```rust
wb.remove_col("Sheet1", "B")?;
```

**TypeScript:**

```typescript
wb.removeCol("Sheet1", "B");
```

### `set_col_outline_level` / `get_col_outline_level`

열의 아웃라인(그룹) 수준을 설정하거나 조회합니다. 범위는 0-7입니다.

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

열 전체에 스타일 ID를 적용하거나 조회합니다.

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

## 행/열 반복자

시트의 모든 행 또는 열 데이터를 한 번에 조회합니다. 데이터가 있는 행/열만 포함됩니다.

### `get_rows(sheet)` / `getRows(sheet)`

시트의 모든 행과 셀 데이터를 반환합니다.

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

시트의 모든 열과 셀 데이터를 반환합니다.

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
