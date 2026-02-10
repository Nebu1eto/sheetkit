## Slicers

Slicers are visual filter controls that let users interactively filter tables in Excel. SheetKit supports adding, querying, and deleting table-based slicers (introduced in Excel 2010).

### `add_slicer(sheet, config)` / `addSlicer(sheet, config)`

Add a slicer to a worksheet targeting a table column.

**Rust:**

```rust
use sheetkit::SlicerConfig;

let config = SlicerConfig {
    name: "StatusFilter".to_string(),
    cell: "F1".to_string(),
    table_name: "Table1".to_string(),
    column_name: "Status".to_string(),
    caption: Some("Status".to_string()),
    style: Some("SlicerStyleLight1".to_string()),
    width: Some(200),
    height: Some(200),
    show_caption: Some(true),
    column_count: None,
};
wb.add_slicer("Sheet1", &config)?;
```

**TypeScript:**

```typescript
wb.addSlicer("Sheet1", {
    name: "StatusFilter",
    cell: "F1",
    tableName: "Table1",
    columnName: "Status",
    caption: "Status",
    style: "SlicerStyleLight1",
    width: 200,
    height: 200,
    showCaption: true,
});
```

### SlicerConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `name` | `String` | `string` | Unique slicer name |
| `cell` | `String` | `string` | Anchor cell (top-left corner, e.g. "F1") |
| `table_name` | `String` | `string` | Source table name |
| `column_name` | `String` | `string` | Column from the table to filter |
| `caption` | `Option<String>` | `string?` | Caption header text. Defaults to column name |
| `style` | `Option<String>` | `string?` | Visual style (e.g. "SlicerStyleLight1") |
| `width` | `Option<u32>` | `number?` | Width in pixels (default 200) |
| `height` | `Option<u32>` | `number?` | Height in pixels (default 200) |
| `show_caption` | `Option<bool>` | `boolean?` | Show the caption header |
| `column_count` | `Option<u32>` | `number?` | Number of columns in the slicer display |

### `get_slicers(sheet)` / `getSlicers(sheet)`

Get information about all slicers on a worksheet.

**Rust:**

```rust
let slicers = wb.get_slicers("Sheet1")?;
for s in &slicers {
    println!("{}: filtering column '{}'", s.name, s.column_name);
}
```

**TypeScript:**

```typescript
const slicers = wb.getSlicers("Sheet1");
for (const s of slicers) {
    console.log(`${s.name}: filtering column '${s.columnName}'`);
}
```

### SlicerInfo

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `name` | `String` | `string` | Slicer name |
| `caption` | `String` | `string` | Display caption |
| `table_name` | `String` | `string` | Source table name |
| `column_name` | `String` | `string` | Column being filtered |
| `style` | `Option<String>` | `string \| null` | Visual style, if set |

### `delete_slicer(sheet, name)` / `deleteSlicer(sheet, name)`

Delete a slicer by name from a worksheet. Removes the slicer definition, cache, content types, and relationships.

**Rust:**

```rust
wb.delete_slicer("Sheet1", "StatusFilter")?;
```

**TypeScript:**

```typescript
wb.deleteSlicer("Sheet1", "StatusFilter");
```

### Slicer Styles

Excel provides built-in slicer styles. Pass one of these as the `style` field:

- `SlicerStyleLight1` through `SlicerStyleLight6`
- `SlicerStyleDark1` through `SlicerStyleDark6`
- `SlicerStyleOther1` through `SlicerStyleOther2`

### Notes

- Slicers are an Excel 2010+ feature. Files with slicers may not display correctly in older spreadsheet applications.
- The slicer is stored as OOXML parts (`xl/slicers/`, `xl/slicerCaches/`) with proper content types and relationships.
- Multiple slicers can be added to the same worksheet, each filtering a different table column.
- Slicers survive save/open round-trips.

---
