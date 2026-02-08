## 16. Freeze/Split Panes

Freeze panes lock rows and/or columns so they remain visible while scrolling.

### `set_panes(sheet, cell)` / `setPanes(sheet, cell)`

Freeze rows and columns. The `cell` argument identifies the top-left cell of the scrollable area:
- `"A2"` freezes row 1
- `"B1"` freezes column A
- `"B2"` freezes row 1 and column A
- `"C3"` freezes rows 1-2 and columns A-B

**Rust:**

```rust
wb.set_panes("Sheet1", "B2")?; // freeze row 1 + column A
```

**TypeScript:**

```typescript
wb.setPanes("Sheet1", "B2");
```

### `unset_panes(sheet)` / `unsetPanes(sheet)`

Remove any freeze or split panes from a sheet.

**Rust:**

```rust
wb.unset_panes("Sheet1")?;
```

**TypeScript:**

```typescript
wb.unsetPanes("Sheet1");
```

### `get_panes(sheet)` / `getPanes(sheet)`

Get the current freeze pane cell reference, or `None`/`null` if no panes are set.

**Rust:**

```rust
let pane: Option<String> = wb.get_panes("Sheet1")?;
```

**TypeScript:**

```typescript
const pane: string | null = wb.getPanes("Sheet1");
```

---

## 17. Page Layout

Page layout settings control how a sheet appears when printed.

### Margins

#### `set_page_margins` / `setPageMargins`

Set page margins in inches.

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
```

#### `get_page_margins` / `getPageMargins`

Get page margins for a sheet. Returns default values if not explicitly set.

**Rust:**

```rust
let margins = wb.get_page_margins("Sheet1")?;
```

**TypeScript:**

```typescript
const margins = wb.getPageMargins("Sheet1");
```

### Page Setup

#### `set_page_setup` / `setPageSetup`

Set paper size, orientation, scale, and fit-to-page options.

**Rust:**

```rust
use sheetkit::page_layout::{Orientation, PaperSize};

wb.set_page_setup("Sheet1", Some(Orientation::Landscape), Some(PaperSize::A4), Some(100), None, None)?;
```

**TypeScript:**

```typescript
wb.setPageSetup("Sheet1", {
    paperSize: "a4",       // "letter" | "tabloid" | "legal" | "a3" | "a4" | "a5" | "b4" | "b5"
    orientation: "landscape",  // "portrait" | "landscape"
    scale: 100,            // 10-400
    fitToWidth: 1,         // number of pages wide
    fitToHeight: 1,        // number of pages tall
});
```

#### `get_page_setup` / `getPageSetup`

Get the current page setup for a sheet.

**TypeScript:**

```typescript
const setup = wb.getPageSetup("Sheet1");
// { paperSize?: string, orientation?: string, scale?: number, fitToWidth?: number, fitToHeight?: number }
```

### Print Options

#### `set_print_options` / `setPrintOptions`

Set print options: gridlines, headings, and centering.

**Rust:**

```rust
wb.set_print_options("Sheet1", Some(true), Some(false), Some(true), None)?;
```

**TypeScript:**

```typescript
wb.setPrintOptions("Sheet1", {
    gridLines: true,
    headings: false,
    horizontalCentered: true,
    verticalCentered: false,
});
```

#### `get_print_options` / `getPrintOptions`

Get print options for a sheet.

**TypeScript:**

```typescript
const opts = wb.getPrintOptions("Sheet1");
```

### Header and Footer

#### `set_header_footer` / `setHeaderFooter`

Set header and/or footer text for printing. Uses Excel formatting codes:
- `&L` left section
- `&C` center section
- `&R` right section
- `&P` page number
- `&N` total pages
- `&D` date
- `&T` time
- `&F` file name

**Rust:**

```rust
wb.set_header_footer("Sheet1", Some("&CMonthly Report"), Some("&LPage &P of &N"))?;
```

**TypeScript:**

```typescript
wb.setHeaderFooter("Sheet1", "&CMonthly Report", "&LPage &P of &N");
```

#### `get_header_footer` / `getHeaderFooter`

Get header and footer text for a sheet.

**Rust:**

```rust
let (header, footer) = wb.get_header_footer("Sheet1")?;
```

**TypeScript:**

```typescript
const result = wb.getHeaderFooter("Sheet1");
// { header?: string, footer?: string }
```

### Page Breaks

#### `insert_page_break` / `insertPageBreak`

Insert a horizontal page break before the given 1-based row number.

**Rust:**

```rust
wb.insert_page_break("Sheet1", 20)?;
```

**TypeScript:**

```typescript
wb.insertPageBreak("Sheet1", 20);
```

#### `remove_page_break` / `removePageBreak`

Remove a page break at the given row.

**Rust:**

```rust
wb.remove_page_break("Sheet1", 20)?;
```

**TypeScript:**

```typescript
wb.removePageBreak("Sheet1", 20);
```

#### `get_page_breaks` / `getPageBreaks`

Get all row page break positions (1-based).

**Rust:**

```rust
let breaks: Vec<u32> = wb.get_page_breaks("Sheet1")?;
```

**TypeScript:**

```typescript
const breaks: number[] = wb.getPageBreaks("Sheet1");
```

---

## 18. Defined Names

Defined names (named ranges) assign a symbolic name to a cell reference or formula. Names can be workbook-scoped (visible from all sheets) or sheet-scoped (visible only within a specific sheet).

### `set_defined_name` / `setDefinedName`

Add or update a defined name. If a name with the same name and scope already exists, its value and comment are updated (no duplication).

**Rust:**

```rust
// Workbook-scoped name
wb.set_defined_name("SalesTotal", "Sheet1!$B$10", None, None)?;

// Sheet-scoped name with comment
wb.set_defined_name("LocalRange", "Sheet1!$A$1:$D$10", Some("Sheet1"), Some("Local data range"))?;
```

**TypeScript:**

```typescript
// Workbook-scoped name
wb.setDefinedName({ name: "SalesTotal", value: "Sheet1!$B$10" });

// Sheet-scoped name with comment
wb.setDefinedName({
    name: "LocalRange",
    value: "Sheet1!$A$1:$D$10",
    scope: "Sheet1",
    comment: "Local data range",
});
```

### `get_defined_name` / `getDefinedName`

Get a defined name by name and optional scope. Returns `None`/`null` if not found.

**Rust:**

```rust
// Get workbook-scoped name
if let Some(info) = wb.get_defined_name("SalesTotal", None)? {
    println!("Refers to: {}", info.value);
}

// Get sheet-scoped name
if let Some(info) = wb.get_defined_name("LocalRange", Some("Sheet1"))? {
    println!("Sheet-scoped: {}", info.value);
}
```

**TypeScript:**

```typescript
// Get workbook-scoped name
const info = wb.getDefinedName("SalesTotal");
if (info) {
    console.log(`Refers to: ${info.value}`);
}

// Get sheet-scoped name
const local = wb.getDefinedName("LocalRange", "Sheet1");
```

### `get_all_defined_names` / `getDefinedNames`

List all defined names in the workbook.

**Rust:**

```rust
let names = wb.get_all_defined_names();
for dn in &names {
    println!("{}: {} (scope: {:?})", dn.name, dn.value, dn.scope);
}
```

**TypeScript:**

```typescript
const names = wb.getDefinedNames();
for (const dn of names) {
    console.log(`${dn.name}: ${dn.value} (scope: ${dn.scope ?? "workbook"})`);
}
```

### `delete_defined_name` / `deleteDefinedName`

Delete a defined name by name and optional scope. Returns an error if the name does not exist.

**Rust:**

```rust
wb.delete_defined_name("SalesTotal", None)?;
wb.delete_defined_name("LocalRange", Some("Sheet1"))?;
```

**TypeScript:**

```typescript
wb.deleteDefinedName("SalesTotal");
wb.deleteDefinedName("LocalRange", "Sheet1");
```

### DefinedNameInfo

| Field | Rust Type | TypeScript Type | Description |
|---|---|---|---|
| `name` | `String` | `string` | The defined name |
| `value` | `String` | `string` | The reference or formula |
| `scope` | `DefinedNameScope` | `string?` | Sheet name if sheet-scoped, or `None`/`undefined` if workbook-scoped |
| `comment` | `Option<String>` | `string?` | Optional comment |

> Defined names cannot contain `\ / ? * [ ]` characters and cannot start or end with whitespace.

---

## 19. Document Properties

Document properties store metadata about the workbook file.

### Core Properties

#### `set_doc_props` / `setDocProps`

Set core document properties (title, creator, etc.).

**Rust:**

```rust
use sheetkit::doc_props::DocProperties;

wb.set_doc_props(DocProperties {
    title: Some("Annual Report".to_string()),
    creator: Some("Finance Team".to_string()),
    subject: Some("Financial Summary".to_string()),
    ..Default::default()
});
```

**TypeScript:**

```typescript
wb.setDocProps({
    title: "Annual Report",
    creator: "Finance Team",
    subject: "Financial Summary",
});
```

#### `get_doc_props` / `getDocProps`

Get core document properties.

**Rust:**

```rust
let props = wb.get_doc_props();
```

**TypeScript:**

```typescript
const props = wb.getDocProps();
```

#### Core Properties Fields

| Field | Type | Description |
|---|---|---|
| `title` | `Option<String>` / `string?` | Document title |
| `subject` | `Option<String>` / `string?` | Subject |
| `creator` | `Option<String>` / `string?` | Author name |
| `keywords` | `Option<String>` / `string?` | Keywords |
| `description` | `Option<String>` / `string?` | Description/comments |
| `last_modified_by` | `Option<String>` / `string?` | Last editor |
| `revision` | `Option<String>` / `string?` | Revision number |
| `created` | `Option<String>` / `string?` | Creation timestamp (ISO 8601) |
| `modified` | `Option<String>` / `string?` | Last modified timestamp |
| `category` | `Option<String>` / `string?` | Category |
| `content_status` | `Option<String>` / `string?` | Content status |

### Application Properties

#### `set_app_props` / `setAppProps`

Set application properties.

**Rust:**

```rust
use sheetkit::doc_props::AppProperties;

wb.set_app_props(AppProperties {
    application: Some("SheetKit".to_string()),
    company: Some("Acme Corp".to_string()),
    ..Default::default()
});
```

**TypeScript:**

```typescript
wb.setAppProps({
    application: "SheetKit",
    company: "Acme Corp",
});
```

#### `get_app_props` / `getAppProps`

Get application properties.

#### Application Properties Fields

| Field | Type | Description |
|---|---|---|
| `application` | `Option<String>` / `string?` | Application name |
| `doc_security` | `Option<u32>` / `number?` | Document security level |
| `company` | `Option<String>` / `string?` | Company name |
| `app_version` | `Option<String>` / `string?` | Application version |
| `manager` | `Option<String>` / `string?` | Manager name |
| `template` | `Option<String>` / `string?` | Template name |

### Custom Properties

Custom properties store arbitrary key-value metadata.

#### `set_custom_property` / `setCustomProperty`

Set a custom property. Accepts String, Int, Float, Bool, or DateTime values.

**Rust:**

```rust
use sheetkit::doc_props::CustomPropertyValue;

wb.set_custom_property("Department", CustomPropertyValue::String("Engineering".to_string()));
wb.set_custom_property("Version", CustomPropertyValue::Int(3));
wb.set_custom_property("Approved", CustomPropertyValue::Bool(true));
```

**TypeScript:**

```typescript
wb.setCustomProperty("Department", "Engineering");
wb.setCustomProperty("Version", 3);
wb.setCustomProperty("Approved", true);
```

> Note: In TypeScript, numeric values are automatically distinguished as integer or float. Integer-like numbers (no fractional part) within the i32 range are stored as `Int`; others are stored as `Float`.

#### `get_custom_property` / `getCustomProperty`

Get a custom property value, or `None`/`null` if not found.

**Rust:**

```rust
if let Some(value) = wb.get_custom_property("Department") {
    // value is CustomPropertyValue
}
```

**TypeScript:**

```typescript
const value = wb.getCustomProperty("Department");
// string | number | boolean | null
```

#### `delete_custom_property` / `deleteCustomProperty`

Delete a custom property. Returns `true` if it existed.

**Rust:**

```rust
let existed: bool = wb.delete_custom_property("Department");
```

**TypeScript:**

```typescript
const existed: boolean = wb.deleteCustomProperty("Department");
```

---

## 20. Workbook Protection

Workbook protection prevents structural changes to the workbook (adding, removing, or renaming sheets).

### `protect_workbook` / `protectWorkbook`

Protect the workbook with optional password and lock settings.

**Rust:**

```rust
use sheetkit::protection::WorkbookProtectionConfig;

wb.protect_workbook(WorkbookProtectionConfig {
    password: Some("secret".to_string()),
    lock_structure: true,
    lock_windows: false,
    lock_revision: false,
});
```

**TypeScript:**

```typescript
wb.protectWorkbook({
    password: "secret",
    lockStructure: true,
    lockWindows: false,
    lockRevision: false,
});
```

### `unprotect_workbook` / `unprotectWorkbook`

Remove workbook protection.

**Rust:**

```rust
wb.unprotect_workbook();
```

**TypeScript:**

```typescript
wb.unprotectWorkbook();
```

### `is_workbook_protected` / `isWorkbookProtected`

Check whether the workbook is protected.

**Rust:**

```rust
let protected: bool = wb.is_workbook_protected();
```

**TypeScript:**

```typescript
const isProtected: boolean = wb.isWorkbookProtected();
```

### WorkbookProtectionConfig

| Field | Type | Description |
|---|---|---|
| `password` | `Option<String>` / `string?` | Password (hashed with legacy Excel algorithm) |
| `lock_structure` | `bool` / `boolean?` | Prevent adding/removing/renaming sheets |
| `lock_windows` | `bool` / `boolean?` | Prevent moving/resizing workbook windows |
| `lock_revision` | `bool` / `boolean?` | Lock revision tracking |

> Note: The password uses the legacy Excel hash algorithm, which is NOT cryptographically secure. It provides only basic deterrence.

---

## 21. Sheet Protection

Sheet protection prevents editing of cells within a single sheet. You can optionally specify a password and grant specific permissions.

### `protect_sheet` / `protectSheet`

Protect a sheet with optional password and granular permission settings. All permission booleans default to `false` (forbidden). Set a permission to `true` to allow that action even when the sheet is protected.

**Rust:**

```rust
use sheetkit::sheet::SheetProtectionConfig;

wb.protect_sheet("Sheet1", &SheetProtectionConfig {
    password: Some("mypass".to_string()),
    format_cells: true,
    insert_rows: true,
    sort: true,
    ..SheetProtectionConfig::default()
})?;
```

**TypeScript:**

```typescript
wb.protectSheet("Sheet1", {
    password: "mypass",
    formatCells: true,
    insertRows: true,
    sort: true,
});

// Protect with defaults (all actions forbidden, no password)
wb.protectSheet("Sheet1");
```

### `unprotect_sheet` / `unprotectSheet`

Remove protection from a sheet.

**Rust:**

```rust
wb.unprotect_sheet("Sheet1")?;
```

**TypeScript:**

```typescript
wb.unprotectSheet("Sheet1");
```

### `is_sheet_protected` / `isSheetProtected`

Check if a sheet is protected.

**Rust:**

```rust
let protected: bool = wb.is_sheet_protected("Sheet1")?;
```

**TypeScript:**

```typescript
const isProtected: boolean = wb.isSheetProtected("Sheet1");
```

### SheetProtectionConfig

| Field | Rust Type | TypeScript Type | Default | Description |
|---|---|---|---|---|
| `password` | `Option<String>` | `string?` | `None` | Password (hashed with legacy Excel algorithm) |
| `select_locked_cells` / `selectLockedCells` | `bool` | `boolean?` | `false` | Allow selecting locked cells |
| `select_unlocked_cells` / `selectUnlockedCells` | `bool` | `boolean?` | `false` | Allow selecting unlocked cells |
| `format_cells` / `formatCells` | `bool` | `boolean?` | `false` | Allow formatting cells |
| `format_columns` / `formatColumns` | `bool` | `boolean?` | `false` | Allow formatting columns |
| `format_rows` / `formatRows` | `bool` | `boolean?` | `false` | Allow formatting rows |
| `insert_columns` / `insertColumns` | `bool` | `boolean?` | `false` | Allow inserting columns |
| `insert_rows` / `insertRows` | `bool` | `boolean?` | `false` | Allow inserting rows |
| `insert_hyperlinks` / `insertHyperlinks` | `bool` | `boolean?` | `false` | Allow inserting hyperlinks |
| `delete_columns` / `deleteColumns` | `bool` | `boolean?` | `false` | Allow deleting columns |
| `delete_rows` / `deleteRows` | `bool` | `boolean?` | `false` | Allow deleting rows |
| `sort` | `bool` | `boolean?` | `false` | Allow sorting |
| `auto_filter` / `autoFilter` | `bool` | `boolean?` | `false` | Allow using auto-filter |
| `pivot_tables` / `pivotTables` | `bool` | `boolean?` | `false` | Allow using pivot tables |

> Note: The password uses the legacy Excel hash algorithm, which is NOT cryptographically secure. It provides only basic deterrence.

---

## 22. Formula Evaluation

SheetKit includes a formula evaluator that supports 110 Excel functions. Formulas are parsed using a nom-based parser and evaluated against the current workbook data.

### `evaluate_formula` / `evaluateFormula`

Evaluate a single formula string in the context of a specific sheet.

**Rust:**

```rust
let result: CellValue = wb.evaluate_formula("Sheet1", "SUM(A1:A10)")?;
```

**TypeScript:**

```typescript
const result = wb.evaluateFormula("Sheet1", "SUM(A1:A10)");
// returns: string | number | boolean | DateValue | null
```

### `calculate_all` / `calculateAll`

Recalculate all formula cells in the workbook. Uses a dependency graph with topological sort (Kahn's algorithm) to ensure formulas are calculated in the correct order.

**Rust:**

```rust
wb.calculate_all()?;
```

**TypeScript:**

```typescript
wb.calculateAll();
```

### Supported Functions (110)

#### Math (23 functions)

`SUM`, `ABS`, `INT`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `MOD`, `POWER`, `SQRT`, `CEILING`, `FLOOR`, `SIGN`, `RAND`, `RANDBETWEEN`, `PI`, `LOG`, `LOG10`, `LN`, `EXP`, `PRODUCT`, `QUOTIENT`, `FACT`, `SUMIF`, `SUMIFS`

#### Statistical (15 functions)

`AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `AVERAGEIF`, `AVERAGEIFS`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `MEDIAN`, `MODE`, `LARGE`, `SMALL`, `RANK`

#### Text (18 functions)

`LEN`, `LOWER`, `UPPER`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `CONCATENATE`, `CONCAT`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `REPT`, `EXACT`, `T`, `PROPER`, `VALUE`, `TEXT`

#### Logical (11 functions)

`IF`, `AND`, `OR`, `NOT`, `TRUE`, `FALSE`, `IFERROR`, `IFNA`, `IFS`, `SWITCH`, `XOR`

#### Information (13 functions)

`ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISERROR`, `ISERR`, `ISNA`, `ISLOGICAL`, `ISEVEN`, `ISODD`, `TYPE`, `N`, `NA`, `ERROR.TYPE`

#### Date/Time (17 functions)

`DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `HOUR`, `MINUTE`, `SECOND`, `DATEDIF`, `EDATE`, `EOMONTH`, `DATEVALUE`, `WEEKDAY`, `WEEKNUM`, `NETWORKDAYS`, `WORKDAY`

#### Lookup (11 functions)

`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `LOOKUP`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `CHOOSE`, `ADDRESS`

> Note: Function names are case-insensitive. Unsupported functions return an error. The evaluator supports cell references (A1, $B$2), range references (A1:C10), cross-sheet references (Sheet2!A1), and standard arithmetic operators (+, -, *, /, ^, &, comparison operators).

---

## 23. Pivot Tables

Pivot tables summarize data from a source range into a structured report.

### `add_pivot_table` / `addPivotTable`

Add a pivot table to the workbook.

**Rust:**

```rust
use sheetkit::pivot::{PivotTableConfig, PivotField, PivotDataField, AggregateFunction};

let config = PivotTableConfig {
    name: "SalesPivot".to_string(),
    source_sheet: "Data".to_string(),
    source_range: "A1:D100".to_string(),
    target_sheet: "PivotSheet".to_string(),
    target_cell: "A1".to_string(),
    rows: vec![PivotField { name: "Region".to_string() }],
    columns: vec![PivotField { name: "Quarter".to_string() }],
    data: vec![PivotDataField {
        name: "Revenue".to_string(),
        function: AggregateFunction::Sum,
        display_name: Some("Total Revenue".to_string()),
    }],
};
wb.add_pivot_table(&config)?;
```

**TypeScript:**

```typescript
wb.addPivotTable({
    name: "SalesPivot",
    sourceSheet: "Data",
    sourceRange: "A1:D100",
    targetSheet: "PivotSheet",
    targetCell: "A1",
    rows: [{ name: "Region" }],
    columns: [{ name: "Quarter" }],
    data: [{
        name: "Revenue",
        function: "sum",
        displayName: "Total Revenue",
    }],
});
```

> Note (Node.js): each `data[].function` must be a supported aggregate (`sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevP`, `var`, `varP`). Unknown values return an error.

### `get_pivot_tables` / `getPivotTables`

Get all pivot tables in the workbook.

**Rust:**

```rust
let tables: Vec<PivotTableInfo> = wb.get_pivot_tables();
for t in &tables {
    println!("{}: {} -> {}", t.name, t.source_range, t.location);
}
```

**TypeScript:**

```typescript
const tables = wb.getPivotTables();
```

### `delete_pivot_table` / `deletePivotTable`

Delete a pivot table by name.

**Rust:**

```rust
wb.delete_pivot_table("SalesPivot")?;
```

**TypeScript:**

```typescript
wb.deletePivotTable("SalesPivot");
```

### PivotTableConfig

| Field | Type | Description |
|---|---|---|
| `name` | `String` / `string` | Pivot table name |
| `source_sheet` | `String` / `string` | Source data sheet name |
| `source_range` | `String` / `string` | Source data range (e.g., "A1:D100") |
| `target_sheet` | `String` / `string` | Target sheet for the pivot table |
| `target_cell` | `String` / `string` | Top-left cell of the pivot table |
| `rows` | `Vec<PivotField>` / `PivotField[]` | Row fields |
| `columns` | `Vec<PivotField>` / `PivotField[]` | Column fields |
| `data` | `Vec<PivotDataField>` / `PivotDataField[]` | Data/value fields |

### PivotDataField

| Field | Type | Description |
|---|---|---|
| `name` | `String` / `string` | Column name from source data header |
| `function` | `AggregateFunction` / `string` | Aggregate function |
| `display_name` | `Option<String>` / `string?` | Custom display name |

### Aggregate Functions

| Rust | TypeScript | Description |
|---|---|---|
| `AggregateFunction::Sum` | `"sum"` | Sum of values |
| `AggregateFunction::Count` | `"count"` | Count of entries |
| `AggregateFunction::Average` | `"average"` | Average |
| `AggregateFunction::Max` | `"max"` | Maximum |
| `AggregateFunction::Min` | `"min"` | Minimum |
| `AggregateFunction::Product` | `"product"` | Product |
| `AggregateFunction::CountNums` | `"countNums"` | Count of numeric values |
| `AggregateFunction::StdDev` | `"stdDev"` | Standard deviation (sample) |
| `AggregateFunction::StdDevP` | `"stdDevP"` | Standard deviation (population) |
| `AggregateFunction::Var` | `"var"` | Variance (sample) |
| `AggregateFunction::VarP` | `"varP"` | Variance (population) |

---

## 24. StreamWriter

The `StreamWriter` provides a forward-only streaming API for writing large sheets without holding the entire worksheet in memory. Rows must be written in ascending order.

### Basic Workflow

1. Create a stream writer from the workbook
2. Set column widths and other column settings (must be done BEFORE writing any rows)
3. Write rows in ascending order
4. Apply the stream writer back to the workbook

**Rust:**

```rust
use sheetkit::cell::CellValue;

let mut sw = wb.new_stream_writer("LargeSheet")?;

// Set column widths BEFORE writing rows
sw.set_col_width(1, 20.0)?;
sw.set_col_width(2, 15.0)?;

// Write header
sw.write_row(1, &[
    CellValue::from("Name"),
    CellValue::from("Score"),
])?;

// Write data rows
for i in 2..=1000 {
    sw.write_row(i, &[
        CellValue::from(format!("Item {}", i - 1)),
        CellValue::from(i as f64 * 1.5),
    ])?;
}

// Apply to workbook
let sheet_index = wb.apply_stream_writer(sw)?;
wb.save("large_output.xlsx")?;
```

**TypeScript:**

```typescript
const sw = wb.newStreamWriter("LargeSheet");

// Set column widths BEFORE writing rows
sw.setColWidth(1, 20.0);
sw.setColWidth(2, 15.0);

// Write header
sw.writeRow(1, ["Name", "Score"]);

// Write data rows
for (let i = 2; i <= 1000; i++) {
    sw.writeRow(i, [`Item ${i - 1}`, i * 1.5]);
}

// Apply to workbook
const sheetIndex = wb.applyStreamWriter(sw);
await wb.save("large_output.xlsx");
```

### StreamWriter API

#### `set_col_width(col, width)` / `setColWidth(col, width)`

Set the width of a single column. Column numbers are 1-based. Must be called before any `write_row`.

#### `set_col_width_range(min, max, width)` / `setColWidthRange(min, max, width)`

Set the width for a range of columns (inclusive). Must be called before any `write_row`.

#### `write_row(row, values)` / `writeRow(row, values)`

Write a row of values. Row numbers are 1-based and must be written in ascending order.

#### `add_merge_cell(reference)` / `addMergeCell(reference)`

Register a merge cell range (e.g., "A1:C1").

**Rust:**

```rust
sw.add_merge_cell("A1:C1")?;
```

**TypeScript:**

```typescript
sw.addMergeCell("A1:C1");
```

#### Rust-Only StreamWriter Methods

The following methods are available only in the Rust API:

- `set_freeze_panes(cell)` -- Set freeze panes for the streamed sheet (must be called before writing rows)
- `set_col_visible(col, visible)` -- Set column visibility
- `set_col_outline_level(col, level)` -- Set column outline level (0-7)
- `set_col_style(col, style_id)` -- Set column style
- `write_row_with_options(row, values, options)` -- Write a row with custom options (height, visibility, outline level, style)

```rust
use sheetkit::stream::StreamRowOptions;

sw.set_freeze_panes("A2")?; // freeze row 1
sw.set_col_visible(3, false)?; // hide column C
sw.set_col_style(1, style_id)?;

sw.write_row_with_options(1, &values, &StreamRowOptions {
    height: Some(25.0),
    visible: Some(true),
    outline_level: Some(1),
    style_id: Some(style_id),
})?;
```

> Important: Column widths, visibility, styles, outline levels, and freeze panes must ALL be set before the first `write_row` call. Setting them after writing any rows returns an error.

---

## 25. Utility Functions

These utility functions are available in the Rust API only (`sheetkit_core::utils::cell_ref`).

### `cell_name_to_coordinates`

Convert an A1-style cell reference to 1-based (column, row) coordinates. Supports absolute references (e.g., "$B$3").

```rust
use sheetkit_core::utils::cell_ref::cell_name_to_coordinates;

let (col, row) = cell_name_to_coordinates("B3")?;
assert_eq!((col, row), (2, 3));

let (col, row) = cell_name_to_coordinates("$AA$100")?;
assert_eq!((col, row), (27, 100));
```

### `coordinates_to_cell_name`

Convert 1-based (column, row) coordinates to an A1-style cell reference.

```rust
use sheetkit_core::utils::cell_ref::coordinates_to_cell_name;

let name = coordinates_to_cell_name(2, 3)?;
assert_eq!(name, "B3");
```

### `column_name_to_number`

Convert a column letter name to a 1-based column number.

```rust
use sheetkit_core::utils::cell_ref::column_name_to_number;

assert_eq!(column_name_to_number("A")?, 1);
assert_eq!(column_name_to_number("Z")?, 26);
assert_eq!(column_name_to_number("AA")?, 27);
assert_eq!(column_name_to_number("XFD")?, 16384);
```

### `column_number_to_name`

Convert a 1-based column number to its letter name.

```rust
use sheetkit_core::utils::cell_ref::column_number_to_name;

assert_eq!(column_number_to_name(1)?, "A");
assert_eq!(column_number_to_name(26)?, "Z");
assert_eq!(column_number_to_name(27)?, "AA");
assert_eq!(column_number_to_name(16384)?, "XFD");
```

### Date Conversion Functions

Available in `sheetkit_core::cell`:

- `date_to_serial(NaiveDate) -> f64` -- Convert a chrono date to an Excel serial number
- `datetime_to_serial(NaiveDateTime) -> f64` -- Convert a chrono datetime to an Excel serial number with time fraction
- `serial_to_date(f64) -> Option<NaiveDate>` -- Convert an Excel serial number to a date
- `serial_to_datetime(f64) -> Option<NaiveDateTime>` -- Convert an Excel serial number to a datetime

```rust
use chrono::NaiveDate;
use sheetkit_core::cell::{date_to_serial, serial_to_date};

let date = NaiveDate::from_ymd_opt(2025, 6, 15).unwrap();
let serial = date_to_serial(date);
let roundtrip = serial_to_date(serial).unwrap();
assert_eq!(date, roundtrip);
```

> Note: Excel uses the 1900 date system with a known bug where it incorrectly treats 1900 as a leap year. Serial number 60 (February 29, 1900) does not correspond to a real date. These conversion functions account for this bug.

### `is_date_num_fmt(num_fmt_id)` (Rust only)

Check whether a built-in number format ID represents a date or time format. Returns `true` for IDs 14-22 and 45-47.

```rust
use sheetkit::is_date_num_fmt;

assert!(is_date_num_fmt(14));   // m/d/yyyy
assert!(is_date_num_fmt(22));   // m/d/yyyy h:mm
assert!(!is_date_num_fmt(0));   // General
assert!(!is_date_num_fmt(49));  // @
```

### `is_date_format_code(code)` (Rust only)

Check whether a custom number format string represents a date or time format. Returns `true` if the format code contains date/time tokens (y, m, d, h, s) outside of quoted strings and escaped characters.

```rust
use sheetkit::is_date_format_code;

assert!(is_date_format_code("yyyy-mm-dd"));
assert!(is_date_format_code("h:mm:ss AM/PM"));
assert!(!is_date_format_code("#,##0.00"));
assert!(!is_date_format_code("0%"));
```

---

## 26. Sparklines

Sparklines are mini-charts embedded in worksheet cells. SheetKit supports three sparkline types: Line, Column, and Win/Loss. Excel defines 36 style presets (indices 0-35).

Sparklines are stored as x14 worksheet extensions in the OOXML package and persist through save/open roundtrips.

### Types

#### `SparklineType` (Rust) / `sparklineType` (TypeScript)

| Value | Rust | TypeScript | OOXML |
|-------|------|------------|-------|
| Line | `SparklineType::Line` | `"line"` | (default, omitted) |
| Column | `SparklineType::Column` | `"column"` | `"column"` |
| Win/Loss | `SparklineType::WinLoss` | `"winloss"` or `"stacked"` | `"stacked"` |

#### `SparklineConfig` (Rust)

```rust
use sheetkit::SparklineConfig;

let config = SparklineConfig::new("Sheet1!A1:A10", "B1");
```

Fields:

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `data_range` | `String` | (required) | Data source range (e.g., `"Sheet1!A1:A10"`) |
| `location` | `String` | (required) | Cell where sparkline is rendered (e.g., `"B1"`) |
| `sparkline_type` | `SparklineType` | `Line` | Sparkline chart type |
| `markers` | `bool` | `false` | Show data markers |
| `high_point` | `bool` | `false` | Highlight highest point |
| `low_point` | `bool` | `false` | Highlight lowest point |
| `first_point` | `bool` | `false` | Highlight first point |
| `last_point` | `bool` | `false` | Highlight last point |
| `negative_points` | `bool` | `false` | Highlight negative values |
| `show_axis` | `bool` | `false` | Show horizontal axis |
| `line_weight` | `Option<f64>` | `None` | Line weight in points |
| `style` | `Option<u32>` | `None` | Style preset index (0-35) |

#### `JsSparklineConfig` (TypeScript)

```typescript
const config = {
  dataRange: 'Sheet1!A1:A10',
  location: 'B1',
  sparklineType: 'line',    // "line" | "column" | "winloss" | "stacked"
  markers: true,
  highPoint: false,
  lowPoint: false,
  firstPoint: false,
  lastPoint: false,
  negativePoints: false,
  showAxis: false,
  lineWeight: 0.75,
  style: 1,
};
```

### Workbook.addSparkline / Workbook::add_sparkline

Add a sparkline to a worksheet.

**Rust:**

```rust
use sheetkit::{Workbook, SparklineConfig, SparklineType};

let mut wb = Workbook::new();

let mut config = SparklineConfig::new("Sheet1!A1:A10", "B1");
config.sparkline_type = SparklineType::Column;
config.markers = true;
config.high_point = true;

wb.add_sparkline("Sheet1", &config).unwrap();
```

**TypeScript:**

```typescript
import { Workbook } from 'sheetkit';

const wb = new Workbook();
wb.addSparkline('Sheet1', {
  dataRange: 'Sheet1!A1:A10',
  location: 'B1',
  sparklineType: 'column',
  markers: true,
  highPoint: true,
});
```

### Workbook.getSparklines / Workbook::get_sparklines

Retrieve all sparklines for a worksheet.

**Rust:**

```rust
let sparklines = wb.get_sparklines("Sheet1").unwrap();
for s in &sparklines {
    println!("{} -> {}", s.data_range, s.location);
}
```

**TypeScript:**

```typescript
const sparklines = wb.getSparklines('Sheet1');
for (const s of sparklines) {
  console.log(`${s.dataRange} -> ${s.location}`);
}
```

### Workbook.removeSparkline / Workbook::remove_sparkline

Remove a sparkline by its location cell reference.

**Rust:**

```rust
wb.remove_sparkline("Sheet1", "B1").unwrap();
```

**TypeScript:**

```typescript
wb.removeSparkline('Sheet1', 'B1');
```

### Validation

The `validate_sparkline_config` function (Rust) checks that:
- `data_range` is not empty
- `location` is not empty
- `line_weight` (if set) is positive
- `style` (if set) is in range 0-35

Validation is automatically applied when calling `add_sparkline`.

```rust
use sheetkit_core::sparkline::{SparklineConfig, validate_sparkline_config};

let config = SparklineConfig::new("Sheet1!A1:A10", "B1");
validate_sparkline_config(&config).unwrap(); // Ok
```

## 27. Theme Colors

Resolve theme color slots (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink) with optional tint.

### Workbook.getThemeColor (Node.js) / Workbook::get_theme_color (Rust)

| Parameter | Type              | Description                                      |
| --------- | ----------------- | ------------------------------------------------ |
| index     | `u32` / `number`  | Theme color index (0-11)                         |
| tint      | `Option<f64>` / `number \| null` | Tint value: positive lightens, negative darkens |

**Returns:** ARGB hex string (e.g. `"FF4472C4"`) or `None`/`null` if out of range.

**Theme Color Indices:**

| Index | Slot Name | Default Color |
| ----- | --------- | ------------- |
| 0     | dk1       | FF000000      |
| 1     | lt1       | FFFFFFFF      |
| 2     | dk2       | FF44546A      |
| 3     | lt2       | FFE7E6E6      |
| 4     | accent1   | FF4472C4      |
| 5     | accent2   | FFED7D31      |
| 6     | accent3   | FFA5A5A5      |
| 7     | accent4   | FFFFC000      |
| 8     | accent5   | FF5B9BD5      |
| 9     | accent6   | FF70AD47      |
| 10    | hlink     | FF0563C1      |
| 11    | folHlink  | FF954F72      |

#### Node.js

```javascript
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

#### Rust

```rust
let wb = Workbook::new();

// Get accent1 color (no tint)
let color = wb.get_theme_color(4, None); // Some("FF4472C4")

// Apply tint
let tinted = wb.get_theme_color(0, Some(0.5)); // Some("FF7F7F7F")
```

### Gradient Fill

The `FillStyle` type supports gradient fills via the `gradient` field.

#### Types

```rust
pub struct GradientFillStyle {
    pub gradient_type: GradientType, // Linear or Path
    pub degree: Option<f64>,         // Rotation angle for linear gradients
    pub left: Option<f64>,           // Path gradient coordinates (0.0-1.0)
    pub right: Option<f64>,
    pub top: Option<f64>,
    pub bottom: Option<f64>,
    pub stops: Vec<GradientStop>,    // Color stops
}

pub struct GradientStop {
    pub position: f64,     // Position (0.0-1.0)
    pub color: StyleColor, // Color at this stop
}

pub enum GradientType {
    Linear,
    Path,
}
```

#### Rust Example

```rust
use sheetkit::*;

let mut wb = Workbook::new();
let style_id = wb.add_style(&Style {
    fill: Some(FillStyle {
        pattern: PatternType::None,
        fg_color: None,
        bg_color: None,
        gradient: Some(GradientFillStyle {
            gradient_type: GradientType::Linear,
            degree: Some(90.0),
            left: None,
            right: None,
            top: None,
            bottom: None,
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: StyleColor::Rgb("FFFFFFFF".to_string()),
                },
                GradientStop {
                    position: 1.0,
                    color: StyleColor::Rgb("FF4472C4".to_string()),
                },
            ],
        }),
    }),
    ..Style::default()
})?;
```

---

## 28. Rich Text

Rich text allows a single cell to contain multiple text segments (runs), each with independent formatting such as font, size, bold, italic, and color.

### `RichTextRun` Type

Each run in a rich text cell is described by a `RichTextRun`.

**Rust:**

```rust
pub struct RichTextRun {
    pub text: String,
    pub font: Option<String>,
    pub size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub color: Option<String>,
}
```

**TypeScript:**

```typescript
interface RichTextRun {
  text: string;
  font?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;  // RGB hex string, e.g. "#FF0000"
}
```

### `set_cell_rich_text` / `setCellRichText`

Set a cell value to rich text with multiple formatted runs.

**Rust:**

```rust
use sheetkit::{Workbook, RichTextRun};

let mut wb = Workbook::new();
let runs = vec![
    RichTextRun {
        text: "Bold text".to_string(),
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
];
wb.set_cell_rich_text("Sheet1", "A1", runs)?;
```

**TypeScript:**

```typescript
const wb = new Workbook();
wb.setCellRichText("Sheet1", "A1", [
  { text: "Bold text", font: "Arial", size: 14, bold: true, color: "#FF0000" },
  { text: " normal text" },
]);
```

### `get_cell_rich_text` / `getCellRichText`

Retrieve the rich text runs for a cell. Returns `None`/`null` for non-rich-text cells.

**Rust:**

```rust
let runs = wb.get_cell_rich_text("Sheet1", "A1")?;
if let Some(runs) = runs {
    for run in &runs {
        println!("Text: {:?}, Bold: {}", run.text, run.bold);
    }
}
```

**TypeScript:**

```typescript
const runs = wb.getCellRichText("Sheet1", "A1");
if (runs) {
  for (const run of runs) {
    console.log(`Text: ${run.text}, Bold: ${run.bold ?? false}`);
  }
}
```

### `CellValue::RichString` (Rust only)

Rich text cells use the `CellValue::RichString(Vec<RichTextRun>)` variant. When read through `get_cell_value`, the display value is the concatenation of all run texts.

```rust
match wb.get_cell_value("Sheet1", "A1")? {
    CellValue::RichString(runs) => {
        println!("Rich text with {} runs", runs.len());
    }
    _ => {}
}
```

### `rich_text_to_plain`

Utility function to extract the concatenated plain text from a slice of rich text runs.

**Rust:**

```rust
use sheetkit::rich_text_to_plain;

let plain = rich_text_to_plain(&runs);
```
