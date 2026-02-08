### Freeze/Split Panes

Freeze rows and columns so they remain visible while scrolling. The freeze cell identifies the top-left cell of the scrollable area: `"A2"` freezes row 1, `"B1"` freezes column A, `"B2"` freezes both.

#### Rust

```rust
let mut wb = Workbook::new();

// Freeze the top row
wb.set_panes("Sheet1", "A2")?;

// Freeze the first column
wb.set_panes("Sheet1", "B1")?;

// Read current pane setting
let pane: Option<String> = wb.get_panes("Sheet1")?;

// Remove freeze panes
wb.unset_panes("Sheet1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Freeze the top row
wb.setPanes('Sheet1', 'A2');

// Freeze the first column
wb.setPanes('Sheet1', 'B1');

// Read current pane setting
const pane: string | null = wb.getPanes('Sheet1');

// Remove freeze panes
wb.unsetPanes('Sheet1');
```

---

### Page Layout

Control how a sheet appears when printed: margins, paper size, orientation, print options, headers/footers, and page breaks.

#### Rust

```rust
use sheetkit::page_layout::{Orientation, PageMarginsConfig, PaperSize};

let mut wb = Workbook::new();

// Set margins (inches)
wb.set_page_margins("Sheet1", &PageMarginsConfig {
    left: 0.7, right: 0.7,
    top: 0.75, bottom: 0.75,
    header: 0.3, footer: 0.3,
})?;

// Set page setup
wb.set_page_setup("Sheet1",
    Some(Orientation::Landscape), Some(PaperSize::A4),
    Some(100), None, None)?;

// Header and footer
wb.set_header_footer("Sheet1",
    Some("&CMonthly Report"), Some("&LPage &P of &N"))?;

// Page breaks (1-based row number)
wb.insert_page_break("Sheet1", 20)?;
let breaks = wb.get_page_breaks("Sheet1")?;
wb.remove_page_break("Sheet1", 20)?;
```

#### TypeScript

```typescript
// Set margins (inches)
wb.setPageMargins('Sheet1', {
    left: 0.7, right: 0.7,
    top: 0.75, bottom: 0.75,
    header: 0.3, footer: 0.3,
});

// Set page setup
wb.setPageSetup('Sheet1', {
    paperSize: 'a4',
    orientation: 'landscape',
    scale: 100,
});

// Header and footer
wb.setHeaderFooter('Sheet1', '&CMonthly Report', '&LPage &P of &N');

// Page breaks (1-based row number)
wb.insertPageBreak('Sheet1', 20);
const breaks: number[] = wb.getPageBreaks('Sheet1');
wb.removePageBreak('Sheet1', 20);
```

---

### Data Validation

Add data validation rules to cell ranges. These rules restrict what values users can enter in the specified cells.

#### Validation Types

| Rust Variant             | TS String       | Description                    |
|--------------------------|-----------------|--------------------------------|
| `ValidationType::Whole`  | `"whole"`       | Whole number constraint        |
| `ValidationType::Decimal`| `"decimal"`     | Decimal number constraint      |
| `ValidationType::List`   | `"list"`        | Dropdown list                  |
| `ValidationType::Date`   | `"date"`        | Date constraint                |
| `ValidationType::Time`   | `"time"`        | Time constraint                |
| `ValidationType::TextLength` | `"textLength"` | Text length constraint      |
| `ValidationType::Custom` | `"custom"`      | Custom formula constraint      |

#### Validation Operators

`Between`, `NotBetween`, `Equal`, `NotEqual`, `LessThan`, `LessThanOrEqual`, `GreaterThan`, `GreaterThanOrEqual`.

In TypeScript, use lowercase strings: `"between"`, `"notBetween"`, `"equal"`, etc.

#### Error Styles

`Stop`, `Warning`, `Information` -- controls the severity of the error dialog shown on invalid input.

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
        formula1: Some("\"Option A,Option B,Option C\"".into()),
        formula2: None,
        allow_blank: true,
        show_input_message: true,
        prompt_title: Some("Select an option".into()),
        prompt_message: Some("Choose from the dropdown".into()),
        show_error_message: true,
        error_style: Some(ErrorStyle::Stop),
        error_title: Some("Invalid".into()),
        error_message: Some("Please select from the list".into()),
    },
)?;

// Read all validations on a sheet
let validations = wb.get_data_validations("Sheet1")?;

// Remove a validation by cell range reference
wb.remove_data_validation("Sheet1", "C2:C100")?;
```

#### TypeScript

```typescript
// Dropdown list validation
wb.addDataValidation('Sheet1', {
    sqref: 'C2:C100',
    validationType: 'list',
    formula1: '"Option A,Option B,Option C"',
    allowBlank: true,
    showInputMessage: true,
    promptTitle: 'Select an option',
    promptMessage: 'Choose from the dropdown',
    showErrorMessage: true,
    errorStyle: 'stop',
    errorTitle: 'Invalid',
    errorMessage: 'Please select from the list',
});

// Read all validations on a sheet
const validations = wb.getDataValidations('Sheet1');

// Remove a validation by cell range reference
wb.removeDataValidation('Sheet1', 'C2:C100');
```

---

### Comments

Add, read, and remove cell comments.

#### Rust

```rust
use sheetkit::{CommentConfig, Workbook};

let mut wb = Workbook::new();

// Add a comment
wb.add_comment(
    "Sheet1",
    &CommentConfig {
        cell: "A1".into(),
        author: "Admin".into(),
        text: "This cell contains the project name.".into(),
    },
)?;

// Get all comments on a sheet
let comments: Vec<CommentConfig> = wb.get_comments("Sheet1")?;

// Remove a comment from a specific cell
wb.remove_comment("Sheet1", "A1")?;
```

#### TypeScript

```typescript
// Add a comment
wb.addComment('Sheet1', {
    cell: 'A1',
    author: 'Admin',
    text: 'This cell contains the project name.',
});

// Get all comments on a sheet
const comments = wb.getComments('Sheet1');

// Remove a comment from a specific cell
wb.removeComment('Sheet1', 'A1');
```

---

### Auto-Filter

Apply or remove auto-filter dropdowns on a range of columns.

#### Rust

```rust
// Set auto-filter on a range
wb.set_auto_filter("Sheet1", "A1:D100")?;

// Remove auto-filter
wb.remove_auto_filter("Sheet1")?;
```

#### TypeScript

```typescript
// Set auto-filter on a range
wb.setAutoFilter('Sheet1', 'A1:D100');

// Remove auto-filter
wb.removeAutoFilter('Sheet1');
```

---

### Formula Evaluation

Evaluate formulas against workbook data. The evaluator supports 110+ Excel functions including SUM, AVERAGE, IF, VLOOKUP, INDEX, MATCH, DATE, and more. Function names are case-insensitive.

#### Rust

```rust
let mut wb = Workbook::new();

wb.set_cell_value("Sheet1", "A1", CellValue::from(10))?;
wb.set_cell_value("Sheet1", "A2", CellValue::from(20))?;
wb.set_cell_value("Sheet1", "A3", CellValue::from(30))?;

// Evaluate a single formula
let result = wb.evaluate_formula("Sheet1", "SUM(A1:A3)")?;

// Set formula cells and recalculate all at once
wb.set_cell_value("Sheet1", "A4",
    CellValue::Formula { expr: "SUM(A1:A3)".into(), result: None })?;
wb.calculate_all()?;
```

#### TypeScript

```typescript
const wb = new Workbook();

wb.setCellValue('Sheet1', 'A1', 10);
wb.setCellValue('Sheet1', 'A2', 20);
wb.setCellValue('Sheet1', 'A3', 30);

// Evaluate a single formula
const result = wb.evaluateFormula('Sheet1', 'SUM(A1:A3)');

// Recalculate all formula cells in dependency order
wb.calculateAll();
```

---

### Pivot Tables

Pivot tables summarize data from a source range into a structured report with row fields, column fields, and aggregated data fields.

#### Rust

```rust
use sheetkit::{CellValue, Workbook};
use sheetkit::pivot::{PivotTableConfig, PivotField, PivotDataField, AggregateFunction};

let mut wb = Workbook::new();

// Prepare source data
wb.set_cell_value("Data", "A1", CellValue::from("Region"))?;
wb.set_cell_value("Data", "B1", CellValue::from("Revenue"))?;
wb.set_cell_value("Data", "A2", CellValue::from("East"))?;
wb.set_cell_value("Data", "B2", CellValue::from(1000))?;
wb.set_cell_value("Data", "A3", CellValue::from("West"))?;
wb.set_cell_value("Data", "B3", CellValue::from(2000))?;

wb.add_pivot_table(&PivotTableConfig {
    name: "SalesPivot".into(),
    source_sheet: "Data".into(),
    source_range: "A1:B3".into(),
    target_sheet: "PivotSheet".into(),
    target_cell: "A1".into(),
    rows: vec![PivotField { name: "Region".into() }],
    columns: vec![],
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

// Prepare source data on "Data" sheet
wb.newSheet('Data');
wb.setCellValue('Data', 'A1', 'Region');
wb.setCellValue('Data', 'B1', 'Revenue');
wb.setCellValue('Data', 'A2', 'East');
wb.setCellValue('Data', 'B2', 1000);
wb.setCellValue('Data', 'A3', 'West');
wb.setCellValue('Data', 'B3', 2000);

wb.addPivotTable({
    name: 'SalesPivot',
    sourceSheet: 'Data',
    sourceRange: 'A1:B3',
    targetSheet: 'PivotSheet',
    targetCell: 'A1',
    rows: [{ name: 'Region' }],
    columns: [],
    data: [{
        name: 'Revenue',
        function: 'sum',
        displayName: 'Total Revenue',
    }],
});

const tables = wb.getPivotTables();
wb.deletePivotTable('SalesPivot');
```

Supported aggregate functions: `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNums`, `StdDev`, `StdDevP`, `Var`, `VarP`.

---

### StreamWriter

The StreamWriter provides a forward-only, streaming API for writing large sheets efficiently. It writes XML directly to an internal buffer, avoiding the need to build the entire worksheet in memory.

Rows must be written in ascending order. Column widths must be set before writing any rows.

#### Rust

```rust
use sheetkit::{CellValue, Workbook};

let mut wb = Workbook::new();

// Create a stream writer for a new sheet
let mut sw = wb.new_stream_writer("LargeSheet")?;

// Set column widths (must be done before writing rows)
sw.set_col_width(1, 20.0)?;     // column 1 (A)
sw.set_col_width(2, 15.0)?;     // column 2 (B)

// Write rows in ascending order (1-based)
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

// Optionally add merge cells
sw.add_merge_cell("A1:B1")?;

// Apply the stream writer to the workbook
wb.apply_stream_writer(sw)?;

wb.save("large_file.xlsx")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Create a stream writer for a new sheet
const sw = wb.newStreamWriter('LargeSheet');

// Set column widths (must be done before writing rows)
sw.setColWidth(1, 20);     // column 1 (A)
sw.setColWidth(2, 15);     // column 2 (B)

// Write rows in ascending order (1-based)
sw.writeRow(1, ['Name', 'Score']);
for (let i = 2; i <= 10000; i++) {
    sw.writeRow(i, [`User_${i - 1}`, i * 1.5]);
}

// Optionally add merge cells
sw.addMergeCell('A1:B1');

// Apply the stream writer to the workbook
wb.applyStreamWriter(sw);

await wb.save('large_file.xlsx');
```

#### StreamWriter API Summary

| Method                | Description                                     |
|-----------------------|-------------------------------------------------|
| `set_col_width`       | Set width for a single column (1-based number)  |
| `set_col_width_range` | Set width for a range of columns (Rust only)    |
| `write_row`           | Write a row of values at the given row number   |
| `add_merge_cell`      | Add a merge cell reference (e.g., `"A1:C3"`)    |

---

### Document Properties

Set and read document metadata: core properties (title, author, etc.), application properties, and custom properties.

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

// Custom properties (string, integer, float, boolean, or datetime)
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

// Custom properties (string, number, or boolean)
wb.setCustomProperty('Project', 'SheetKit');
wb.setCustomProperty('Version', 1);
wb.setCustomProperty('Released', false);

const val = wb.getCustomProperty('Project');       // string | number | boolean | null
const deleted: boolean = wb.deleteCustomProperty('Version');
```

#### DocProperties Fields

| Field              | Type             | Description                     |
|--------------------|------------------|---------------------------------|
| `title`            | `Option<String>` | Document title                  |
| `subject`          | `Option<String>` | Document subject                |
| `creator`          | `Option<String>` | Author name                     |
| `keywords`         | `Option<String>` | Keywords for search             |
| `description`      | `Option<String>` | Document description            |
| `last_modified_by` | `Option<String>` | Last editor                     |
| `revision`         | `Option<String>` | Revision number                 |
| `created`          | `Option<String>` | Creation timestamp              |
| `modified`         | `Option<String>` | Last modification timestamp     |
| `category`         | `Option<String>` | Category                        |
| `content_status`   | `Option<String>` | Content status                  |

#### AppProperties Fields

| Field          | Type             | Description                     |
|----------------|------------------|---------------------------------|
| `application`  | `Option<String>` | Application name                |
| `doc_security` | `Option<u32>`    | Document security level         |
| `company`      | `Option<String>` | Company name                    |
| `app_version`  | `Option<String>` | Application version             |
| `manager`      | `Option<String>` | Manager name                    |
| `template`     | `Option<String>` | Template name                   |

---

### Workbook Protection

Protect the workbook structure to prevent users from adding, deleting, or renaming sheets. An optional password can be set (legacy Excel hash -- not cryptographically secure).

#### Rust

```rust
use sheetkit::{Workbook, WorkbookProtectionConfig};

let mut wb = Workbook::new();

// Protect the workbook
wb.protect_workbook(WorkbookProtectionConfig {
    password: Some("secret".into()),
    lock_structure: true,    // prevent sheet add/delete/rename
    lock_windows: false,     // allow window resize
    lock_revision: false,    // allow revision tracking changes
});

// Check if protected
let is_protected: bool = wb.is_workbook_protected();

// Remove protection
wb.unprotect_workbook();
```

#### TypeScript

```typescript
// Protect the workbook
wb.protectWorkbook({
    password: 'secret',
    lockStructure: true,
    lockWindows: false,
    lockRevision: false,
});

// Check if protected
const isProtected: boolean = wb.isWorkbookProtected();

// Remove protection
wb.unprotectWorkbook();
```

---

### File Encryption

SheetKit supports file-level encryption for .xlsx files using the ECMA-376 standard. Encrypted files are stored in OLE/CFB compound containers rather than plain ZIP archives.

- **Reading**: Supports both Standard Encryption (Office 2007, AES-128-ECB) and Agile Encryption (Office 2010+, AES-256-CBC).
- **Writing**: Uses Agile Encryption (AES-256-CBC + SHA-512 with 100,000 iterations).

> Note: File encryption is different from workbook/sheet protection. Encryption prevents the file from being opened at all without the correct password, while protection only restricts editing operations.

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
import { Workbook } from 'sheetkit';

const wb = new Workbook();
wb.setCellValue('Sheet1', 'A1', 'Confidential');

// Save with password (Agile Encryption)
wb.saveWithPassword('encrypted.xlsx', 'mypassword');

// Open encrypted file (sync)
const wb2 = Workbook.openWithPasswordSync('encrypted.xlsx', 'mypassword');
const val = wb2.getCellValue('Sheet1', 'A1');

// Async variants
const wb3 = await Workbook.openWithPassword('encrypted.xlsx', 'mypassword');
await wb3.saveWithPassword('encrypted_copy.xlsx', 'newpassword');
```

> Note: The `encryption` feature must be enabled in Rust (`sheetkit = { features = ["encryption"] }`). Node.js bindings always include encryption support.

---

### Sparklines

Sparklines are mini-charts rendered inside individual cells. Three types are supported: line, column, and win/loss. Style presets (0-35) correspond to the built-in Excel sparkline styles.

#### Rust

```rust
use sheetkit::{SparklineConfig, SparklineType, Workbook};

let mut wb = Workbook::new();

// Populate data
for i in 1..=10 {
    wb.set_cell_value("Sheet1", &format!("A{i}"), CellValue::from(i as f64 * 1.5))?;
}

// Add a column sparkline in cell B1
let mut config = SparklineConfig::new("Sheet1!A1:A10", "B1");
config.sparkline_type = SparklineType::Column;
config.high_point = true;
config.low_point = true;
config.style = Some(5);

wb.add_sparkline("Sheet1", &config)?;

// Read sparklines
let sparklines = wb.get_sparklines("Sheet1")?;

// Remove a sparkline by location
wb.remove_sparkline("Sheet1", "B1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Populate data
for (let i = 1; i <= 10; i++) {
    wb.setCellValue('Sheet1', `A${i}`, i * 1.5);
}

// Add a column sparkline in cell B1
wb.addSparkline('Sheet1', {
    dataRange: 'Sheet1!A1:A10',
    location: 'B1',
    sparklineType: 'column',
    highPoint: true,
    lowPoint: true,
    style: 5,
});

// Read sparklines
const sparklines = wb.getSparklines('Sheet1');

// Remove a sparkline by location
wb.removeSparkline('Sheet1', 'B1');
```

#### SparklineConfig Fields

| Field | Rust Type | TS Type | Description |
|-------|-----------|---------|-------------|
| `data_range` / `dataRange` | `String` | `string` | Data source range (e.g., `"Sheet1!A1:A10"`) |
| `location` | `String` | `string` | Cell where sparkline is rendered |
| `sparkline_type` / `sparklineType` | `SparklineType` | `string?` | `"line"`, `"column"`, or `"stacked"` (win/loss) |
| `markers` | `bool` | `boolean?` | Show data markers |
| `high_point` / `highPoint` | `bool` | `boolean?` | Highlight highest point |
| `low_point` / `lowPoint` | `bool` | `boolean?` | Highlight lowest point |
| `first_point` / `firstPoint` | `bool` | `boolean?` | Highlight first point |
| `last_point` / `lastPoint` | `bool` | `boolean?` | Highlight last point |
| `negative_points` / `negativePoints` | `bool` | `boolean?` | Highlight negative values |
| `show_axis` / `showAxis` | `bool` | `boolean?` | Show horizontal axis |
| `line_weight` / `lineWeight` | `Option<f64>` | `number?` | Line weight in points |
| `style` | `Option<u32>` | `number?` | Style preset index (0-35) |

---

### Defined Names

Defined names assign a human-readable name to a cell range or formula. They can be workbook-scoped (visible everywhere) or sheet-scoped (visible only on the named sheet).

#### Rust

```rust
use sheetkit::Workbook;

let mut wb = Workbook::new();

// Workbook-scoped name
wb.set_defined_name("SalesTotal", "Sheet1!$B$10", None, None)?;

// Sheet-scoped name with comment
wb.set_defined_name(
    "LocalRange", "Sheet1!$A$1:$D$10",
    Some("Sheet1"), Some("Local data range"),
)?;

// Read a defined name
if let Some(info) = wb.get_defined_name("SalesTotal", None)? {
    println!("Value: {}", info.value);
}

// List all defined names
let names = wb.get_all_defined_names();

// Delete a defined name
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

// Sheet-scoped name with comment
wb.setDefinedName({
    name: 'LocalRange',
    value: 'Sheet1!$A$1:$D$10',
    scope: 'Sheet1',
    comment: 'Local data range',
});

// Read a defined name (null = workbook scope)
const info = wb.getDefinedName('SalesTotal', null);

// List all defined names
const names = wb.getDefinedNames();

// Delete a defined name
wb.deleteDefinedName('SalesTotal', null);
```

---

### Sheet Protection

Sheet protection restricts editing operations on individual worksheets. Unlike workbook protection (which prevents structural changes), sheet protection controls cell-level actions such as formatting, inserting, deleting, sorting, and filtering. An optional password can be set.

#### Rust

```rust
use sheetkit::Workbook;
use sheetkit::sheet::SheetProtectionConfig;

let mut wb = Workbook::new();

// Protect a sheet with a password and allow sorting
wb.protect_sheet("Sheet1", SheetProtectionConfig {
    password: Some("secret".into()),
    sort: true,
    auto_filter: true,
    ..Default::default()
})?;

// Check if a sheet is protected
let is_protected: bool = wb.is_sheet_protected("Sheet1")?;

// Remove protection
wb.unprotect_sheet("Sheet1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Protect a sheet with a password and allow sorting
wb.protectSheet('Sheet1', {
    password: 'secret',
    sort: true,
    autoFilter: true,
});

// Check if a sheet is protected
const isProtected: boolean = wb.isSheetProtected('Sheet1');

// Remove protection
wb.unprotectSheet('Sheet1');
```

#### SheetProtectionConfig Fields

| Field | Rust Type | TS Type | Description |
|-------|-----------|---------|-------------|
| `password` | `Option<String>` | `string?` | Optional password (legacy Excel hash) |
| `select_locked_cells` / `selectLockedCells` | `bool` | `boolean?` | Allow selecting locked cells |
| `select_unlocked_cells` / `selectUnlockedCells` | `bool` | `boolean?` | Allow selecting unlocked cells |
| `format_cells` / `formatCells` | `bool` | `boolean?` | Allow formatting cells |
| `format_columns` / `formatColumns` | `bool` | `boolean?` | Allow formatting columns |
| `format_rows` / `formatRows` | `bool` | `boolean?` | Allow formatting rows |
| `insert_columns` / `insertColumns` | `bool` | `boolean?` | Allow inserting columns |
| `insert_rows` / `insertRows` | `bool` | `boolean?` | Allow inserting rows |
| `insert_hyperlinks` / `insertHyperlinks` | `bool` | `boolean?` | Allow inserting hyperlinks |
| `delete_columns` / `deleteColumns` | `bool` | `boolean?` | Allow deleting columns |
| `delete_rows` / `deleteRows` | `bool` | `boolean?` | Allow deleting rows |
| `sort` | `bool` | `boolean?` | Allow sorting |
| `auto_filter` / `autoFilter` | `bool` | `boolean?` | Allow using auto-filter |
| `pivot_tables` / `pivotTables` | `bool` | `boolean?` | Allow using pivot tables |

---

## Examples

Complete example projects demonstrating all features are available in the repository:

- **Rust**: `examples/rust/` -- a standalone Cargo project (`cargo run` from within the directory)
- **Node.js**: `examples/node/` -- a TypeScript project (build the native module first, then run with `npx tsx index.ts`)

Each example walks through every feature: creating a workbook, setting cell values, managing sheets, applying styles, adding charts and images, data validation, comments, auto-filter, streaming large datasets, document properties, workbook protection, file encryption, sparklines, defined names, and sheet protection.

---

## Utility Functions

SheetKit also exposes helper functions for working with cell references:

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