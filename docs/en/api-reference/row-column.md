## 4. Row Operations

All row numbers are 1-based.

### `insert_rows(sheet, row, count)` / `insertRows(sheet, row, count)`

Insert `count` empty rows starting at `row`, shifting existing rows at and below that position downward. Cell references in shifted rows are updated automatically.

**Rust:**

```rust
wb.insert_rows("Sheet1", 3, 2)?; // insert 2 rows at row 3
```

**TypeScript:**

```typescript
wb.insertRows("Sheet1", 3, 2);
```

### `remove_row(sheet, row)` / `removeRow(sheet, row)`

Delete a single row and shift rows below it upward by one.

**Rust:**

```rust
wb.remove_row("Sheet1", 5)?;
```

**TypeScript:**

```typescript
wb.removeRow("Sheet1", 5);
```

### `duplicate_row(sheet, row)` / `duplicateRow(sheet, row)`

Copy a row and insert the duplicate directly below the source row. Existing rows below are shifted down.

**Rust:**

```rust
wb.duplicate_row("Sheet1", 2)?;
```

**TypeScript:**

```typescript
wb.duplicateRow("Sheet1", 2);
```

### `set_row_height` / `get_row_height`

Set or get the height of a row in points. Returns `None`/`null` when no explicit height is set.

**Rust:**

```rust
wb.set_row_height("Sheet1", 1, 30.0)?;
let height: Option<f64> = wb.get_row_height("Sheet1", 1)?;
```

**TypeScript:**

```typescript
wb.setRowHeight("Sheet1", 1, 30.0);
const height: number | null = wb.getRowHeight("Sheet1", 1);
```

### `set_row_visible` / `get_row_visible`

Set or check the visibility of a row. Rows are visible by default.

**Rust:**

```rust
wb.set_row_visible("Sheet1", 3, false)?; // hide row 3
let visible: bool = wb.get_row_visible("Sheet1", 3)?;
```

**TypeScript:**

```typescript
wb.setRowVisible("Sheet1", 3, false);
const visible: boolean = wb.getRowVisible("Sheet1", 3);
```

### `set_row_outline_level` / `get_row_outline_level`

Set or get the outline (grouping) level of a row. Valid range: 0-7. Returns 0 if not set.

**Rust:**

```rust
wb.set_row_outline_level("Sheet1", 5, 1)?;
let level: u8 = wb.get_row_outline_level("Sheet1", 5)?;
```

**TypeScript:**

```typescript
wb.setRowOutlineLevel("Sheet1", 5, 1);
const level: number = wb.getRowOutlineLevel("Sheet1", 5);
```

### `set_row_style` / `get_row_style`

Apply a style ID to an entire row, or retrieve the current row style ID (0 if not set).

**Rust:**

```rust
let style_id = wb.add_style(&style)?;
wb.set_row_style("Sheet1", 1, style_id)?;
let current: u32 = wb.get_row_style("Sheet1", 1)?;
```

**TypeScript:**

```typescript
const styleId = wb.addStyle(style);
wb.setRowStyle("Sheet1", 1, styleId);
const current: number = wb.getRowStyle("Sheet1", 1);
```

---

## 5. Column Operations

Columns are identified by letter names (e.g., "A", "B", "AA").

### `set_col_width` / `get_col_width`

Set or get the width of a column. Valid range: 0.0 to 255.0. Returns `None`/`null` when no explicit width is set.

**Rust:**

```rust
wb.set_col_width("Sheet1", "B", 20.0)?;
let width: Option<f64> = wb.get_col_width("Sheet1", "B")?;
```

**TypeScript:**

```typescript
wb.setColWidth("Sheet1", "B", 20.0);
const width: number | null = wb.getColWidth("Sheet1", "B");
```

### `set_col_visible` / `get_col_visible`

Set or check the visibility of a column. Columns are visible by default.

**Rust:**

```rust
wb.set_col_visible("Sheet1", "C", false)?;
let visible: bool = wb.get_col_visible("Sheet1", "C")?;
```

**TypeScript:**

```typescript
wb.setColVisible("Sheet1", "C", false);
const visible: boolean = wb.getColVisible("Sheet1", "C");
```

### `insert_cols(sheet, col, count)` / `insertCols(sheet, col, count)`

Insert `count` empty columns starting at the given column letter, shifting existing columns to the right. Cell references are updated automatically.

**Rust:**

```rust
wb.insert_cols("Sheet1", "B", 3)?; // insert 3 columns at B
```

**TypeScript:**

```typescript
wb.insertCols("Sheet1", "B", 3);
```

### `remove_col(sheet, col)` / `removeCol(sheet, col)`

Delete a single column and shift columns to its right leftward by one.

**Rust:**

```rust
wb.remove_col("Sheet1", "D")?;
```

**TypeScript:**

```typescript
wb.removeCol("Sheet1", "D");
```

### `set_col_outline_level` / `get_col_outline_level`

Set or get the outline (grouping) level of a column. Valid range: 0-7.

**Rust:**

```rust
wb.set_col_outline_level("Sheet1", "B", 2)?;
let level: u8 = wb.get_col_outline_level("Sheet1", "B")?;
```

**TypeScript:**

```typescript
wb.setColOutlineLevel("Sheet1", "B", 2);
const level: number = wb.getColOutlineLevel("Sheet1", "B");
```

### `set_col_style` / `get_col_style`

Apply a style ID to an entire column, or retrieve the current column style ID (0 if not set).

**Rust:**

```rust
wb.set_col_style("Sheet1", "A", style_id)?;
let current: u32 = wb.get_col_style("Sheet1", "A")?;
```

**TypeScript:**

```typescript
wb.setColStyle("Sheet1", "A", styleId);
const current: number = wb.getColStyle("Sheet1", "A");
```

---

## 6. Row/Column Iterators

### `get_rows(sheet)` / `getRows(sheet)`

Return all rows that contain at least one cell. Sparse: empty rows are omitted.

**Rust:**

```rust
let rows = wb.get_rows("Sheet1")?;
// Vec<(row_number: u32, cells: Vec<(column_name: String, CellValue)>)>
for (row_num, cells) in &rows {
    for (col_name, value) in cells {
        println!("Row {row_num}, Col {col_name}: {value}");
    }
}
```

**TypeScript:**

```typescript
const rows = wb.getRows("Sheet1");
// JsRowData[]
for (const row of rows) {
    console.log(`Row ${row.row}:`);
    for (const cell of row.cells) {
        console.log(`  ${cell.column}: ${cell.value} (${cell.valueType})`);
    }
}
```

### `get_cols(sheet)` / `getCols(sheet)`

Return all columns that contain data. The result is a column-oriented transpose of the row data.

**Rust:**

```rust
let cols = wb.get_cols("Sheet1")?;
// Vec<(column_name: String, cells: Vec<(row_number: u32, CellValue)>)>
```

**TypeScript:**

```typescript
const cols = wb.getCols("Sheet1");
// JsColData[]
```

### TypeScript Cell Types

**JsRowData:**

```typescript
interface JsRowData {
    row: number;         // 1-based row number
    cells: JsRowCell[];
}

interface JsRowCell {
    column: string;         // column name (e.g., "A", "B")
    valueType: string;      // "string" | "number" | "boolean" | "date" | "empty" | "error" | "formula"
    value?: string;         // string representation
    numberValue?: number;   // set when valueType is "number"
    boolValue?: boolean;    // set when valueType is "boolean"
}
```

**JsColData:**

```typescript
interface JsColData {
    column: string;         // column name
    cells: JsColCell[];
}

interface JsColCell {
    row: number;            // 1-based row number
    valueType: string;      // same types as JsRowCell
    value?: string;
    numberValue?: number;
    boolValue?: boolean;
}
```

---
