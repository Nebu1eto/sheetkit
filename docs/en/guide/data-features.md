### Sheet Management

Create, delete, rename, copy, and navigate sheets.

#### Rust

```rust
let mut wb = Workbook::new();

// Create a new sheet (returns 0-based index)
let idx: usize = wb.new_sheet("Sales")?;

// Delete a sheet by name
wb.delete_sheet("Sales")?;

// Rename a sheet
wb.set_sheet_name("Sheet1", "Main")?;

// Copy a sheet (returns new sheet's 0-based index)
let idx: usize = wb.copy_sheet("Main", "Main_Copy")?;

// Get the index of a sheet (None if not found)
let idx: Option<usize> = wb.get_sheet_index("Main");

// Get/set the active sheet
let active: &str = wb.get_active_sheet();
wb.set_active_sheet("Main")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Create a new sheet (returns 0-based index)
const idx: number = wb.newSheet('Sales');

// Delete a sheet
wb.deleteSheet('Sales');

// Rename a sheet
wb.setSheetName('Sheet1', 'Main');

// Copy a sheet (returns new sheet's 0-based index)
const copyIdx: number = wb.copySheet('Main', 'Main_Copy');

// Get the index of a sheet (null if not found)
const sheetIdx: number | null = wb.getSheetIndex('Main');

// Get/set the active sheet
const active: string = wb.getActiveSheet();
wb.setActiveSheet('Main');
```

---

### Row and Column Operations

Insert, delete, and configure rows and columns.

#### Rust

```rust
let mut wb = Workbook::new();

// -- Rows (1-based row numbers) --

// Insert 3 empty rows starting at row 2
wb.insert_rows("Sheet1", 2, 3)?;

// Remove row 5
wb.remove_row("Sheet1", 5)?;

// Duplicate row 1 (inserts copy below)
wb.duplicate_row("Sheet1", 1)?;

// Set/get row height
wb.set_row_height("Sheet1", 1, 25.0)?;
let height: Option<f64> = wb.get_row_height("Sheet1", 1)?;

// Show/hide a row
wb.set_row_visible("Sheet1", 3, false)?;

// -- Columns (letter-based, e.g., "A", "B", "AA") --

// Set/get column width
wb.set_col_width("Sheet1", "A", 20.0)?;
let width: Option<f64> = wb.get_col_width("Sheet1", "A")?;

// Show/hide a column
wb.set_col_visible("Sheet1", "B", false)?;

// Insert 2 empty columns starting at column "C"
wb.insert_cols("Sheet1", "C", 2)?;

// Remove column "D"
wb.remove_col("Sheet1", "D")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// -- Rows (1-based row numbers) --
wb.insertRows('Sheet1', 2, 3);
wb.removeRow('Sheet1', 5);
wb.duplicateRow('Sheet1', 1);
wb.setRowHeight('Sheet1', 1, 25);
const height: number | null = wb.getRowHeight('Sheet1', 1);
wb.setRowVisible('Sheet1', 3, false);

// -- Columns (letter-based) --
wb.setColWidth('Sheet1', 'A', 20);
const width: number | null = wb.getColWidth('Sheet1', 'A');
wb.setColVisible('Sheet1', 'B', false);
wb.insertCols('Sheet1', 'C', 2);
wb.removeCol('Sheet1', 'D');
```

---

### Row/Column Iterators

Read all populated rows or columns from a sheet. Empty rows and columns are omitted.

#### Rust

```rust
let mut wb = Workbook::open("data.xlsx")?;

let rows = wb.get_rows("Sheet1")?;
for (row_num, cells) in &rows {
    for (col_name, value) in cells {
        println!("Row {row_num}, Col {col_name}: {value}");
    }
}

let cols = wb.get_cols("Sheet1")?;
```

#### TypeScript

```typescript
const rows = wb.getRows('Sheet1');
for (const row of rows) {
    for (const cell of row.cells) {
        console.log(`${cell.column}${row.row}: ${cell.value}`);
    }
}

const cols = wb.getCols('Sheet1');
```

---

### Row/Column Outline Levels and Styles

Group rows or columns into collapsible outline levels (0-7), and apply styles to entire rows or columns.

#### Rust

```rust
let mut wb = Workbook::new();

// Outline levels
wb.set_row_outline_level("Sheet1", 3, 1)?;
let level: u8 = wb.get_row_outline_level("Sheet1", 3)?;

wb.set_col_outline_level("Sheet1", "B", 2)?;
let col_level: u8 = wb.get_col_outline_level("Sheet1", "B")?;

// Row/column styles
let style_id = wb.add_style(&style)?;
wb.set_row_style("Sheet1", 1, style_id)?;
let current: u32 = wb.get_row_style("Sheet1", 1)?;

wb.set_col_style("Sheet1", "A", style_id)?;
let col_style: u32 = wb.get_col_style("Sheet1", "A")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

// Outline levels
wb.setRowOutlineLevel('Sheet1', 3, 1);
const level: number = wb.getRowOutlineLevel('Sheet1', 3);

wb.setColOutlineLevel('Sheet1', 'B', 2);
const colLevel: number = wb.getColOutlineLevel('Sheet1', 'B');

// Row/column styles
const styleId = wb.addStyle({ font: { bold: true } });
wb.setRowStyle('Sheet1', 1, styleId);
const current: number = wb.getRowStyle('Sheet1', 1);

wb.setColStyle('Sheet1', 'A', styleId);
const colStyle: number = wb.getColStyle('Sheet1', 'A');
```

---

### Styles

Styles control the visual presentation of cells. Register a style definition to get a style ID, then apply that ID to cells. Identical style definitions are deduplicated automatically.

A `Style` can include any combination of: font, fill, border, alignment, number format, and protection.

#### Rust

```rust
use sheetkit::{
    AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle,
    FillStyle, FontStyle, HorizontalAlign, PatternType, Style,
    StyleColor, VerticalAlign, Workbook,
};

let mut wb = Workbook::new();

// Register a style
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

// Apply the style to a cell
wb.set_cell_style("Sheet1", "A1", style_id)?;

// Read the style ID of a cell (None if default)
let current_style: Option<u32> = wb.get_cell_style("Sheet1", "A1")?;
```

#### TypeScript

```typescript
// Register a style
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

// Apply the style to a cell
wb.setCellStyle('Sheet1', 'A1', styleId);

// Read the style ID of a cell (null if default)
const currentStyle: number | null = wb.getCellStyle('Sheet1', 'A1');
```

#### Style Components Reference

**FontStyle**

| Field           | Rust Type          | TS Type    | Description                     |
|-----------------|--------------------|------------|---------------------------------|
| `name`          | `Option<String>`   | `string?`  | Font family (e.g., "Calibri")   |
| `size`          | `Option<f64>`      | `number?`  | Font size in points             |
| `bold`          | `bool`             | `boolean?` | Bold text                       |
| `italic`        | `bool`             | `boolean?` | Italic text                     |
| `underline`     | `bool`             | `boolean?` | Underline text                  |
| `strikethrough` | `bool`             | `boolean?` | Strikethrough text              |
| `color`         | `Option<StyleColor>` | `string?` | Font color (hex string in TS)  |

**FillStyle**

| Field      | Rust Type          | TS Type   | Description                             |
|------------|--------------------|-----------|-----------------------------------------|
| `pattern`  | `PatternType`      | `string?` | Pattern type (see values below)         |
| `fg_color` | `Option<StyleColor>` | `string?` | Foreground color                      |
| `bg_color` | `Option<StyleColor>` | `string?` | Background color                      |

PatternType values: `None`, `Solid`, `Gray125`, `DarkGray`, `MediumGray`, `LightGray`.
In TypeScript, use lowercase strings: `"none"`, `"solid"`, `"gray125"`, `"darkGray"`, `"mediumGray"`, `"lightGray"`.

**BorderStyle**

Each side (`left`, `right`, `top`, `bottom`, `diagonal`) accepts a `BorderSideStyle` with:
- `style`: one of `Thin`, `Medium`, `Thick`, `Dashed`, `Dotted`, `Double`, `Hair`, `MediumDashed`, `DashDot`, `MediumDashDot`, `DashDotDot`, `MediumDashDotDot`, `SlantDashDot`
- `color`: optional color

In TypeScript, use lowercase strings for border style: `"thin"`, `"medium"`, `"thick"`, etc.

**AlignmentStyle**

| Field           | Rust Type                | TS Type    | Description                 |
|-----------------|--------------------------|------------|-----------------------------|
| `horizontal`    | `Option<HorizontalAlign>`| `string?`  | Horizontal alignment        |
| `vertical`      | `Option<VerticalAlign>`  | `string?`  | Vertical alignment          |
| `wrap_text`     | `bool`                   | `boolean?` | Wrap text                   |
| `text_rotation` | `Option<u32>`            | `number?`  | Text rotation in degrees    |
| `indent`        | `Option<u32>`            | `number?`  | Indentation level           |
| `shrink_to_fit` | `bool`                   | `boolean?` | Shrink text to fit cell     |

HorizontalAlign values: `General`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterContinuous`, `Distributed`.
VerticalAlign values: `Top`, `Center`, `Bottom`, `Justify`, `Distributed`.

**NumFmtStyle** (Rust only)

```rust
use sheetkit::style::NumFmtStyle;

// Built-in format (e.g., percent, date, currency)
NumFmtStyle::Builtin(9)  // 0%

// Custom format string
NumFmtStyle::Custom("#,##0.00".to_string())
```

In TypeScript, use `numFmtId` (built-in format ID) or `customNumFmt` (custom format string) on the style object.

**ProtectionStyle**

| Field    | Rust Type | TS Type    | Description                     |
|----------|-----------|------------|---------------------------------|
| `locked` | `bool`    | `boolean?` | Lock the cell (default: true)   |
| `hidden` | `bool`    | `boolean?` | Hide formulas in protected view |

---

### Charts

Add charts to worksheets. Charts are anchored between two cells (top-left and bottom-right) and render data from specified cell ranges.

#### Supported Chart Types (43 types)

**Column charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::Col` | `"col"` | Clustered column |
| `ChartType::ColStacked` | `"colStacked"` | Stacked column |
| `ChartType::ColPercentStacked` | `"colPercentStacked"` | 100% stacked column |
| `ChartType::Col3D` | `"col3D"` | 3D clustered column |
| `ChartType::Col3DStacked` | `"col3DStacked"` | 3D stacked column |
| `ChartType::Col3DPercentStacked` | `"col3DPercentStacked"` | 3D 100% stacked column |

**Bar charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::Bar` | `"bar"` | Clustered bar |
| `ChartType::BarStacked` | `"barStacked"` | Stacked bar |
| `ChartType::BarPercentStacked` | `"barPercentStacked"` | 100% stacked bar |
| `ChartType::Bar3D` | `"bar3D"` | 3D clustered bar |
| `ChartType::Bar3DStacked` | `"bar3DStacked"` | 3D stacked bar |
| `ChartType::Bar3DPercentStacked` | `"bar3DPercentStacked"` | 3D 100% stacked bar |

**Line charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::Line` | `"line"` | Line |
| `ChartType::LineStacked` | `"lineStacked"` | Stacked line |
| `ChartType::LinePercentStacked` | `"linePercentStacked"` | 100% stacked line |
| `ChartType::Line3D` | `"line3D"` | 3D line |

**Pie charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::Pie` | `"pie"` | Pie |
| `ChartType::Pie3D` | `"pie3D"` | 3D pie |
| `ChartType::Doughnut` | `"doughnut"` | Doughnut |

**Area charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::Area` | `"area"` | Area |
| `ChartType::AreaStacked` | `"areaStacked"` | Stacked area |
| `ChartType::AreaPercentStacked` | `"areaPercentStacked"` | 100% stacked area |
| `ChartType::Area3D` | `"area3D"` | 3D area |
| `ChartType::Area3DStacked` | `"area3DStacked"` | 3D stacked area |
| `ChartType::Area3DPercentStacked` | `"area3DPercentStacked"` | 3D 100% stacked area |

**Scatter charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::Scatter` | `"scatter"` | Scatter (markers only) |
| `ChartType::ScatterSmooth` | `"scatterSmooth"` | Scatter with smooth lines |
| `ChartType::ScatterLine` | `"scatterLine"` | Scatter with straight lines |

**Radar charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::Radar` | `"radar"` | Radar |
| `ChartType::RadarFilled` | `"radarFilled"` | Filled radar |
| `ChartType::RadarMarker` | `"radarMarker"` | Radar with markers |

**Stock charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::StockHLC` | `"stockHLC"` | High-Low-Close |
| `ChartType::StockOHLC` | `"stockOHLC"` | Open-High-Low-Close |
| `ChartType::StockVHLC` | `"stockVHLC"` | Volume-High-Low-Close |
| `ChartType::StockVOHLC` | `"stockVOHLC"` | Volume-Open-High-Low-Close |

**Surface charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::Surface` | `"surface"` | 3D surface |
| `ChartType::Surface3D` | `"surface3D"` | 3D surface (top view) |
| `ChartType::SurfaceWireframe` | `"surfaceWireframe"` | Wireframe surface |
| `ChartType::SurfaceWireframe3D` | `"surfaceWireframe3D"` | Wireframe surface (top view) |

**Other charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::Bubble` | `"bubble"` | Bubble |

**Combo charts:**

| Rust Variant | TS String | Description |
|---|---|---|
| `ChartType::ColLine` | `"colLine"` | Column + line combo |
| `ChartType::ColLineStacked` | `"colLineStacked"` | Stacked column + line |
| `ChartType::ColLinePercentStacked` | `"colLinePercentStacked"` | 100% stacked column + line |

#### Rust

```rust
use sheetkit::{ChartConfig, ChartSeries, ChartType, Workbook};

let mut wb = Workbook::new();

// Populate data first...
wb.set_cell_value("Sheet1", "A1", CellValue::String("Q1".into()))?;
wb.set_cell_value("Sheet1", "B1", CellValue::Number(1500.0))?;
// ... more data rows ...

// Add a chart anchored from D1 to K15
wb.add_chart(
    "Sheet1",
    "D1",   // top-left anchor cell
    "K15",  // bottom-right anchor cell
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

### Images

Embed images into worksheets. Supports 11 formats: PNG, JPEG, GIF, BMP, ICO, TIFF, SVG, EMF, EMZ, WMF, WMZ. Images are anchored to a cell and sized by pixel dimensions.

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

### Data Validation

Add data validation rules to cell ranges to restrict what values users can enter.

#### Validation Types

| Rust Variant | TS String | Description |
|---|---|---|
| `ValidationType::None` | `"none"` | No constraint (prompt/message only) |
| `ValidationType::Whole` | `"whole"` | Whole number constraint |
| `ValidationType::Decimal` | `"decimal"` | Decimal number constraint |
| `ValidationType::List` | `"list"` | Dropdown list |
| `ValidationType::Date` | `"date"` | Date constraint |
| `ValidationType::Time` | `"time"` | Time constraint |
| `ValidationType::TextLength` | `"textLength"` | Text length constraint |
| `ValidationType::Custom` | `"custom"` | Custom formula constraint |

#### Operators

`Between`, `NotBetween`, `Equal`, `NotEqual`, `LessThan`, `LessThanOrEqual`, `GreaterThan`, `GreaterThanOrEqual`.

In TypeScript, input is case-insensitive; output uses camelCase: `"between"`, `"notBetween"`, `"lessThan"`, etc.

The `sqref` must be a valid cell range reference. For types other than `none`, `formula1` is required. For `between`/`notBetween` operators, `formula2` is also required.

#### Error Styles

`Stop`, `Warning`, `Information` -- controls the severity of the error dialog shown when invalid data is entered.

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

// Get all validations on a sheet
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

// Get all validations on a sheet
const validations = wb.getDataValidations('Sheet1');

// Remove a validation by cell range reference
wb.removeDataValidation('Sheet1', 'C2:C100');
```

---

### Merge Cells

Merge a rectangular range of cells into a single visual cell. The value of the top-left cell is displayed across the merged area.

#### Rust

```rust
let mut wb = Workbook::new();

wb.set_cell_value("Sheet1", "A1", CellValue::from("Report Header"))?;
wb.merge_cells("Sheet1", "A1", "C1")?;

let merged = wb.get_merge_cells("Sheet1")?;

wb.unmerge_cell("Sheet1", "A1:C1")?;
```

#### TypeScript

```typescript
const wb = new Workbook();

wb.setCellValue('Sheet1', 'A1', 'Report Header');
wb.mergeCells('Sheet1', 'A1', 'C1');

const merged: string[] = wb.getMergeCells('Sheet1');

wb.unmergeCell('Sheet1', 'A1:C1');
```

---

### Hyperlinks

Attach hyperlinks to cells. Three types are supported: external URLs, internal sheet references, and email addresses.

#### Rust

```rust
use sheetkit::hyperlink::HyperlinkType;

let mut wb = Workbook::new();

// External URL
wb.set_cell_hyperlink("Sheet1", "A1",
    HyperlinkType::External("https://example.com".into()),
    Some("Example"), Some("Click to visit"))?;

// Internal sheet reference
wb.set_cell_hyperlink("Sheet1", "A2",
    HyperlinkType::Internal("Sheet2!A1".into()), None, None)?;

// Email
wb.set_cell_hyperlink("Sheet1", "A3",
    HyperlinkType::Email("mailto:user@example.com".into()),
    None, None)?;

// Read and delete
let info = wb.get_cell_hyperlink("Sheet1", "A1")?;
wb.delete_cell_hyperlink("Sheet1", "A1")?;
```

#### TypeScript

```typescript
// External URL
wb.setCellHyperlink('Sheet1', 'A1', {
    linkType: 'external',
    target: 'https://example.com',
    display: 'Example',
    tooltip: 'Click to visit',
});

// Internal sheet reference
wb.setCellHyperlink('Sheet1', 'A2', {
    linkType: 'internal',
    target: 'Sheet2!A1',
});

// Email
wb.setCellHyperlink('Sheet1', 'A3', {
    linkType: 'email',
    target: 'mailto:user@example.com',
});

// Read and delete
const info = wb.getCellHyperlink('Sheet1', 'A1');
wb.deleteCellHyperlink('Sheet1', 'A1');
```

---

### Conditional Formatting

Change cell appearance based on rules applied to their values. Supports 17 rule types including cell value comparisons, color scales, data bars, and text matching.

#### Rust

```rust
use sheetkit::conditional::*;
use sheetkit::style::*;

let mut wb = Workbook::new();

// Highlight cells greater than 100
let rules = vec![ConditionalFormatRule {
    rule_type: ConditionalFormatType::CellIs {
        operator: CfOperator::GreaterThan,
        formula: "100".to_string(),
        formula2: None,
    },
    format: Some(ConditionalStyle {
        font: Some(FontStyle {
            color: Some(StyleColor::Rgb("#FF0000".into())),
            ..Default::default()
        }),
        fill: Some(FillStyle {
            pattern: PatternType::Solid,
            fg_color: Some(StyleColor::Rgb("#FFCCCC".into())),
            bg_color: None,
        }),
        border: None,
        num_fmt: None,
    }),
    priority: Some(1),
    stop_if_true: false,
}];
wb.set_conditional_format("Sheet1", "A1:A100", &rules)?;

let formats = wb.get_conditional_formats("Sheet1")?;
wb.delete_conditional_format("Sheet1", "A1:A100")?;
```

#### TypeScript

```typescript
// Highlight cells greater than 100
wb.setConditionalFormat('Sheet1', 'A1:A100', [{
    ruleType: 'cellIs',
    operator: 'greaterThan',
    formula: '100',
    format: {
        font: { color: '#FF0000' },
        fill: { pattern: 'solid', fgColor: '#FFCCCC' },
    },
    priority: 1,
}]);

// Color scale (green to red)
wb.setConditionalFormat('Sheet1', 'B2:B50', [{
    ruleType: 'colorScale',
    minType: 'min',
    minColor: '#63BE7B',
    maxType: 'max',
    maxColor: '#F8696B',
}]);

// Data bar
wb.setConditionalFormat('Sheet1', 'C2:C50', [{
    ruleType: 'dataBar',
    barColor: '#638EC6',
    showValue: true,
}]);

const formats = wb.getConditionalFormats('Sheet1');
wb.deleteConditionalFormat('Sheet1', 'A1:A100');
```

---

### Freeze/Split Panes