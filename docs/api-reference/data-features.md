## 8. Merge Cells

Merge cells combines a rectangular range of cells into a single visual cell.

### `merge_cells(sheet, top_left, bottom_right)` / `mergeCells(sheet, topLeft, bottomRight)`

Merge a range of cells defined by its top-left and bottom-right corners.

**Rust:**

```rust
wb.merge_cells("Sheet1", "A1", "C3")?;
```

**TypeScript:**

```typescript
wb.mergeCells("Sheet1", "A1", "C3");
```

### `unmerge_cell(sheet, reference)` / `unmergeCell(sheet, reference)`

Remove a merged cell range. The `reference` must match the exact range that was merged (e.g., "A1:C3").

**Rust:**

```rust
wb.unmerge_cell("Sheet1", "A1:C3")?;
```

**TypeScript:**

```typescript
wb.unmergeCell("Sheet1", "A1:C3");
```

### `get_merge_cells(sheet)` / `getMergeCells(sheet)`

Get all merged cell ranges on a sheet. Returns a list of range strings.

**Rust:**

```rust
let ranges: Vec<String> = wb.get_merge_cells("Sheet1")?;
// e.g., ["A1:C3", "E5:F6"]
```

**TypeScript:**

```typescript
const ranges: string[] = wb.getMergeCells("Sheet1");
```

---

## 9. Hyperlinks

Hyperlinks can link cells to external URLs, internal sheet references, or email addresses.

### `set_cell_hyperlink` / `setCellHyperlink`

Set a hyperlink on a cell.

**Rust:**

```rust
use sheetkit::hyperlink::HyperlinkType;

// External URL
wb.set_cell_hyperlink("Sheet1", "A1", HyperlinkType::External("https://example.com".into()), Some("Example"), Some("Click to visit"))?;

// Internal sheet reference
wb.set_cell_hyperlink("Sheet1", "B1", HyperlinkType::Internal("Sheet2!A1".into()), None, None)?;

// Email
wb.set_cell_hyperlink("Sheet1", "C1", HyperlinkType::Email("mailto:user@example.com".into()), None, None)?;
```

**TypeScript:**

```typescript
// External URL
wb.setCellHyperlink("Sheet1", "A1", {
    linkType: "external",
    target: "https://example.com",
    display: "Example",
    tooltip: "Click to visit",
});

// Internal sheet reference
wb.setCellHyperlink("Sheet1", "B1", {
    linkType: "internal",
    target: "Sheet2!A1",
});

// Email
wb.setCellHyperlink("Sheet1", "C1", {
    linkType: "email",
    target: "mailto:user@example.com",
});
```

### `get_cell_hyperlink` / `getCellHyperlink`

Get hyperlink information for a cell, or `None`/`null` if no hyperlink exists.

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
    // info.linkType, info.target, info.display, info.tooltip
}
```

### `delete_cell_hyperlink` / `deleteCellHyperlink`

Remove a hyperlink from a cell.

**Rust:**

```rust
wb.delete_cell_hyperlink("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.deleteCellHyperlink("Sheet1", "A1");
```

### HyperlinkOptions (TypeScript)

```typescript
interface JsHyperlinkOptions {
    linkType: string;    // "external" | "internal" | "email"
    target: string;      // URL, sheet reference, or mailto address
    display?: string;    // optional display text
    tooltip?: string;    // optional tooltip text
}
```

> Note: External and email hyperlinks are stored in the worksheet `.rels` file with `TargetMode="External"`. Internal hyperlinks use only a `location` attribute.

---

## 10. Charts

Charts render data from cell ranges and are anchored between two cells (top-left and bottom-right corners of the chart area).

### `add_chart` / `addChart`

Add a chart to a sheet.

**Rust:**

```rust
use sheetkit::{ChartConfig, ChartSeries, ChartType};

let config = ChartConfig {
    chart_type: ChartType::Col,
    title: Some("Quarterly Sales".to_string()),
    series: vec![ChartSeries {
        name: "Revenue".to_string(),
        categories: "Sheet1!$A$2:$A$5".to_string(),
        values: "Sheet1!$B$2:$B$5".to_string(),
        x_values: None,
        bubble_sizes: None,
    }],
    show_legend: true,
    view_3d: None,
};
wb.add_chart("Sheet1", "D1", "K15", &config)?;
```

**TypeScript:**

```typescript
wb.addChart("Sheet1", "D1", "K15", {
    chartType: "col",
    title: "Quarterly Sales",
    series: [{
        name: "Revenue",
        categories: "Sheet1!$A$2:$A$5",
        values: "Sheet1!$B$2:$B$5",
    }],
    showLegend: true,
});
```

### ChartConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `chart_type` | `ChartType` | `string` | Chart type (see table below) |
| `title` | `Option<String>` | `string?` | Chart title |
| `series` | `Vec<ChartSeries>` | `JsChartSeries[]` | Data series |
| `show_legend` | `bool` | `boolean?` | Show legend (default: true) |
| `view_3d` | `Option<View3DConfig>` | `JsView3DConfig?` | 3D rotation settings |

### ChartSeries

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `name` | `String` | `string` | Series name |
| `categories` | `String` | `string` | Category axis range (e.g., "Sheet1!$A$2:$A$5") |
| `values` | `String` | `string` | Value axis range (e.g., "Sheet1!$B$2:$B$5") |
| `x_values` | `Option<String>` | `string?` | X-axis values (scatter/bubble charts only) |
| `bubble_sizes` | `Option<String>` | `string?` | Bubble sizes (bubble charts only) |

### Supported Chart Types (43 types)

**Column charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Col` | `"col"` | Clustered column |
| `ChartType::ColStacked` | `"colStacked"` | Stacked column |
| `ChartType::ColPercentStacked` | `"colPercentStacked"` | 100% stacked column |
| `ChartType::Col3D` | `"col3D"` | 3D clustered column |
| `ChartType::Col3DStacked` | `"col3DStacked"` | 3D stacked column |
| `ChartType::Col3DPercentStacked` | `"col3DPercentStacked"` | 3D 100% stacked column |

**Bar charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Bar` | `"bar"` | Clustered bar |
| `ChartType::BarStacked` | `"barStacked"` | Stacked bar |
| `ChartType::BarPercentStacked` | `"barPercentStacked"` | 100% stacked bar |
| `ChartType::Bar3D` | `"bar3D"` | 3D clustered bar |
| `ChartType::Bar3DStacked` | `"bar3DStacked"` | 3D stacked bar |
| `ChartType::Bar3DPercentStacked` | `"bar3DPercentStacked"` | 3D 100% stacked bar |

**Line charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Line` | `"line"` | Line |
| `ChartType::LineStacked` | `"lineStacked"` | Stacked line |
| `ChartType::LinePercentStacked` | `"linePercentStacked"` | 100% stacked line |
| `ChartType::Line3D` | `"line3D"` | 3D line |

**Pie charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Pie` | `"pie"` | Pie |
| `ChartType::Pie3D` | `"pie3D"` | 3D pie |

**Area charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Area` | `"area"` | Area |
| `ChartType::AreaStacked` | `"areaStacked"` | Stacked area |
| `ChartType::AreaPercentStacked` | `"areaPercentStacked"` | 100% stacked area |
| `ChartType::Area3D` | `"area3D"` | 3D area |
| `ChartType::Area3DStacked` | `"area3DStacked"` | 3D stacked area |
| `ChartType::Area3DPercentStacked` | `"area3DPercentStacked"` | 3D 100% stacked area |

**Scatter charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Scatter` | `"scatter"` | Scatter (markers only) |
| `ChartType::ScatterSmooth` | `"scatterSmooth"` | Scatter with smooth lines |
| `ChartType::ScatterLine` | `"scatterLine"` | Scatter with straight lines |

**Radar charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Radar` | `"radar"` | Radar |
| `ChartType::RadarFilled` | `"radarFilled"` | Filled radar |
| `ChartType::RadarMarker` | `"radarMarker"` | Radar with markers |

**Stock charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::StockHLC` | `"stockHLC"` | High-Low-Close |
| `ChartType::StockOHLC` | `"stockOHLC"` | Open-High-Low-Close |
| `ChartType::StockVHLC` | `"stockVHLC"` | Volume-High-Low-Close |
| `ChartType::StockVOHLC` | `"stockVOHLC"` | Volume-Open-High-Low-Close |

**Surface charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Surface` | `"surface"` | 3D surface |
| `ChartType::Surface3D` | `"surface3D"` | 3D surface (top view) |
| `ChartType::SurfaceWireframe` | `"surfaceWireframe"` | Wireframe surface |
| `ChartType::SurfaceWireframe3D` | `"surfaceWireframe3D"` | Wireframe surface (top view) |

**Other charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::Doughnut` | `"doughnut"` | Doughnut |
| `ChartType::Bubble` | `"bubble"` | Bubble |

**Combo charts:**

| Rust | TypeScript | Description |
|---|---|---|
| `ChartType::ColLine` | `"colLine"` | Column + line combo |
| `ChartType::ColLineStacked` | `"colLineStacked"` | Stacked column + line |
| `ChartType::ColLinePercentStacked` | `"colLinePercentStacked"` | 100% stacked column + line |

### View3DConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `rot_x` | `Option<i32>` | `number?` | X-axis rotation angle |
| `rot_y` | `Option<i32>` | `number?` | Y-axis rotation angle |
| `depth_percent` | `Option<u32>` | `number?` | Depth as percentage of chart width |
| `right_angle_axes` | `Option<bool>` | `boolean?` | Use right-angle axes |
| `perspective` | `Option<u32>` | `number?` | Perspective field of view |

---

## 11. Images

Embed images into worksheets. Supports 11 formats: PNG, JPEG, GIF, BMP, ICO, TIFF, SVG, EMF, EMZ, WMF, WMZ.

### `add_image` / `addImage`

Add an image to a sheet at the specified cell position.

**Rust:**

```rust
use sheetkit::image::{ImageConfig, ImageFormat};

let image_data = std::fs::read("logo.png")?;
let config = ImageConfig {
    data: image_data,
    format: ImageFormat::Png,
    from_cell: "B2".to_string(),
    width_px: 200,
    height_px: 100,
};
wb.add_image("Sheet1", &config)?;
```

**TypeScript:**

```typescript
import { readFileSync } from "fs";

const imageData = readFileSync("logo.png");
wb.addImage("Sheet1", {
    data: imageData,
    format: "png",   // "png" | "jpeg" | "jpg" | "gif" | "bmp" | "ico" | "tiff" | "tif" | "svg" | "emf" | "emz" | "wmf" | "wmz"
    fromCell: "B2",
    widthPx: 200,
    heightPx: 100,
});
```

### ImageConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `data` | `Vec<u8>` | `Buffer` | Raw image bytes |
| `format` | `ImageFormat` | `string` | 11 formats supported (see [Images](./image.md#supported-formats)) |
| `from_cell` | `String` | `string` | Anchor cell (top-left corner) |
| `width_px` | `u32` | `number` | Image width in pixels |
| `height_px` | `u32` | `number` | Image height in pixels |

---

## 12. Data Validation

Data validation restricts what values users can enter in cells.

### `add_data_validation` / `addDataValidation`

Add a validation rule to a cell range.

**Rust:**

```rust
use sheetkit::validation::*;

// Dropdown list
let config = DataValidationConfig {
    sqref: "B2:B100".to_string(),
    validation_type: ValidationType::List,
    operator: None,
    formula1: Some("\"Red,Green,Blue\"".to_string()),
    formula2: None,
    allow_blank: true,
    error_style: Some(ErrorStyle::Stop),
    error_title: Some("Invalid color".to_string()),
    error_message: Some("Please select from the list.".to_string()),
    prompt_title: Some("Color".to_string()),
    prompt_message: Some("Choose a color.".to_string()),
    show_input_message: true,
    show_error_message: true,
};
wb.add_data_validation("Sheet1", &config)?;
```

**TypeScript:**

```typescript
wb.addDataValidation("Sheet1", {
    sqref: "B2:B100",
    validationType: "list",
    formula1: '"Red,Green,Blue"',
    allowBlank: true,
    errorStyle: "stop",
    errorTitle: "Invalid color",
    errorMessage: "Please select from the list.",
    promptTitle: "Color",
    promptMessage: "Choose a color.",
    showInputMessage: true,
    showErrorMessage: true,
});
```

> Note (Node.js): `validationType` must be a supported value (`none`, `list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom`). Unknown values return an error. The `sqref` must be a valid cell range (e.g. `"A1:B10"`). For types other than `none`, `formula1` is required. For `between`/`notBetween` operators, `formula2` is also required.

### `get_data_validations` / `getDataValidations`

Get all data validation rules on a sheet.

**Rust:**

```rust
let validations = wb.get_data_validations("Sheet1")?;
```

**TypeScript:**

```typescript
const validations = wb.getDataValidations("Sheet1");
```

### `remove_data_validation` / `removeDataValidation`

Remove a data validation rule by its cell range (sqref).

**Rust:**

```rust
wb.remove_data_validation("Sheet1", "B2:B100")?;
```

**TypeScript:**

```typescript
wb.removeDataValidation("Sheet1", "B2:B100");
```

### Validation Types

| Rust | TypeScript | Description |
|---|---|---|
| `ValidationType::None` | `"none"` | No constraint (prompt/message only) |
| `ValidationType::Whole` | `"whole"` | Whole number |
| `ValidationType::Decimal` | `"decimal"` | Decimal number |
| `ValidationType::List` | `"list"` | Dropdown list |
| `ValidationType::Date` | `"date"` | Date value |
| `ValidationType::Time` | `"time"` | Time value |
| `ValidationType::TextLength` | `"textLength"` | Text length constraint |
| `ValidationType::Custom` | `"custom"` | Custom formula |

### Validation Operators

Used with `Whole`, `Decimal`, `Date`, `Time`, and `TextLength` types. TypeScript input is case-insensitive; output uses camelCase matching the OOXML spec:

| Rust | TypeScript |
|---|---|
| `ValidationOperator::Between` | `"between"` |
| `ValidationOperator::NotBetween` | `"notBetween"` |
| `ValidationOperator::Equal` | `"equal"` |
| `ValidationOperator::NotEqual` | `"notEqual"` |
| `ValidationOperator::LessThan` | `"lessThan"` |
| `ValidationOperator::LessThanOrEqual` | `"lessThanOrEqual"` |
| `ValidationOperator::GreaterThan` | `"greaterThan"` |
| `ValidationOperator::GreaterThanOrEqual` | `"greaterThanOrEqual"` |

### Error Styles

| Rust | TypeScript | Description |
|---|---|---|
| `ErrorStyle::Stop` | `"stop"` | Reject invalid input |
| `ErrorStyle::Warning` | `"warning"` | Warn but allow |
| `ErrorStyle::Information` | `"information"` | Inform only |

---

## 13. Comments

Comments (also known as notes) attach text annotations to individual cells.

When a comment is added, SheetKit automatically generates a VML (Vector Markup Language) drawing part (`xl/drawings/vmlDrawingN.vml`) and a `<legacyDrawing>` reference in the worksheet XML. This ensures that comments render correctly in the Excel UI, which requires VML legacy drawing support for note pop-up boxes.

When opening an existing workbook that already contains VML comment drawings, SheetKit preserves the VML parts through the save/open cycle. When all comments on a sheet are removed, the associated VML part and relationships are cleaned up automatically.

### `add_comment` / `addComment`

Add a comment to a cell.

**Rust:**

```rust
use sheetkit::comment::CommentConfig;

wb.add_comment("Sheet1", &CommentConfig {
    cell: "A1".to_string(),
    author: "Alice".to_string(),
    text: "Please verify this value.".to_string(),
})?;
```

**TypeScript:**

```typescript
wb.addComment("Sheet1", {
    cell: "A1",
    author: "Alice",
    text: "Please verify this value.",
});
```

### `get_comments` / `getComments`

Get all comments on a sheet.

**Rust:**

```rust
let comments = wb.get_comments("Sheet1")?;
for c in &comments {
    println!("{}: {} (by {})", c.cell, c.text, c.author);
}
```

**TypeScript:**

```typescript
const comments = wb.getComments("Sheet1");
```

### `remove_comment` / `removeComment`

Remove a comment from a specific cell.

**Rust:**

```rust
wb.remove_comment("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.removeComment("Sheet1", "A1");
```

### VML Compatibility

Excel uses VML (Vector Markup Language) for rendering comment note boxes. SheetKit handles the following automatically:

- Generating minimal VML drawing parts when new comments are created.
- Preserving existing VML parts from workbooks opened from disk.
- Wiring `<legacyDrawing>` relationship references in worksheet XML.
- Adding the appropriate content type entries for VML parts.
- Cleaning up VML parts and relationships when all comments on a sheet are removed.

No additional API calls are needed. The VML handling is transparent to the user.

---

## 14. Auto-Filter

Auto-filter adds dropdown filter controls to a header row.

### `set_auto_filter` / `setAutoFilter`

Set an auto-filter on a cell range. The first row of the range becomes the filter header.

**Rust:**

```rust
wb.set_auto_filter("Sheet1", "A1:D100")?;
```

**TypeScript:**

```typescript
wb.setAutoFilter("Sheet1", "A1:D100");
```

### `remove_auto_filter` / `removeAutoFilter`

Remove the auto-filter from a sheet.

**Rust:**

```rust
wb.remove_auto_filter("Sheet1")?;
```

**TypeScript:**

```typescript
wb.removeAutoFilter("Sheet1");
```

---

## 15. Conditional Formatting

Conditional formatting changes the appearance of cells based on rules applied to their values.

### `set_conditional_format` / `setConditionalFormat`

Apply one or more conditional formatting rules to a cell range.

**Rust:**

```rust
use sheetkit::conditional::*;
use sheetkit::style::*;

// Highlight cells greater than 100 in red
let rules = vec![ConditionalFormatRule {
    rule_type: ConditionalFormatType::CellIs {
        operator: CfOperator::GreaterThan,
        formula: "100".to_string(),
        formula2: None,
    },
    format: Some(ConditionalStyle {
        font: Some(FontStyle {
            color: Some(StyleColor::Rgb("#FF0000".to_string())),
            ..Default::default()
        }),
        fill: Some(FillStyle {
            pattern: PatternType::Solid,
            fg_color: Some(StyleColor::Rgb("#FFCCCC".to_string())),
            bg_color: None,
        }),
        border: None,
        num_fmt: None,
    }),
    priority: Some(1),
    stop_if_true: false,
}];
wb.set_conditional_format("Sheet1", "A1:A100", &rules)?;
```

**TypeScript:**

```typescript
wb.setConditionalFormat("Sheet1", "A1:A100", [{
    ruleType: "cellIs",
    operator: "greaterThan",
    formula: "100",
    format: {
        font: { color: "#FF0000" },
        fill: { pattern: "solid", fgColor: "#FFCCCC" },
    },
    priority: 1,
}]);
```

### `get_conditional_formats` / `getConditionalFormats`

Get all conditional formatting rules for a sheet.

**Rust:**

```rust
let formats = wb.get_conditional_formats("Sheet1")?;
// Vec<(sqref: String, rules: Vec<ConditionalFormatRule>)>
```

**TypeScript:**

```typescript
const formats = wb.getConditionalFormats("Sheet1");
// JsConditionalFormatEntry[]
// { sqref: string, rules: JsConditionalFormatRule[] }[]
```

### `delete_conditional_format` / `deleteConditionalFormat`

Remove conditional formatting for a specific cell range.

**Rust:**

```rust
wb.delete_conditional_format("Sheet1", "A1:A100")?;
```

**TypeScript:**

```typescript
wb.deleteConditionalFormat("Sheet1", "A1:A100");
```

### Rule Types (17 types)

| Rule Type | Description | Key Fields |
|---|---|---|
| `cellIs` | Compare cell value against formula(s) | `operator`, `formula`, `formula2` |
| `expression` | Custom formula evaluates to true/false | `formula` |
| `colorScale` | Gradient color scale (2 or 3 colors) | `min/mid/max_type`, `min/mid/max_value`, `min/mid/max_color` |
| `dataBar` | Data bar proportional to value | `min/max_type`, `min/max_value`, `bar_color`, `show_value` |
| `duplicateValues` | Highlight duplicate values | (no extra fields) |
| `uniqueValues` | Highlight unique values | (no extra fields) |
| `top10` | Top N values | `rank`, `percent` |
| `bottom10` | Bottom N values | `rank`, `percent` |
| `aboveAverage` | Above/below average values | `above`, `equal_average` |
| `containsBlanks` | Cells that are blank | (no extra fields) |
| `notContainsBlanks` | Cells that are not blank | (no extra fields) |
| `containsErrors` | Cells containing errors | (no extra fields) |
| `notContainsErrors` | Cells without errors | (no extra fields) |
| `containsText` | Cells containing specific text | `text` |
| `notContainsText` | Cells not containing text | `text` |
| `beginsWith` | Cells beginning with text | `text` |
| `endsWith` | Cells ending with text | `text` |

### CfOperator Values

Used with `cellIs` rules: `lessThan`, `lessThanOrEqual`, `equal`, `notEqual`, `greaterThanOrEqual`, `greaterThan`, `between`, `notBetween`.

### CfValueType Values

Used with `colorScale` and `dataBar` for min/mid/max: `num`, `percent`, `min`, `max`, `percentile`, `formula`.

### Examples

**Color scale (green to red):**

```typescript
wb.setConditionalFormat("Sheet1", "B2:B50", [{
    ruleType: "colorScale",
    minType: "min",
    minColor: "#63BE7B",
    maxType: "max",
    maxColor: "#F8696B",
}]);
```

**Data bar:**

```typescript
wb.setConditionalFormat("Sheet1", "C2:C50", [{
    ruleType: "dataBar",
    barColor: "#638EC6",
    showValue: true,
}]);
```

**Contains text:**

```typescript
wb.setConditionalFormat("Sheet1", "A1:A100", [{
    ruleType: "containsText",
    text: "urgent",
    format: {
        font: { bold: true, color: "#FF0000" },
    },
}]);
```

---

## 16. Tables

Tables are structured data ranges with headers, styling, and optional auto-filter. Tables are stored as separate OOXML parts (`xl/tables/tableN.xml`) with full relationship and content-type wiring.

### `add_table(sheet, config)` / `addTable(sheet, config)`

Create a table on a sheet. The table name must be unique across the entire workbook. Returns an error if the name is already taken, the configuration is invalid, or the sheet does not exist.

**Rust:**

```rust
use sheetkit::table::{TableConfig, TableColumn};

let config = TableConfig {
    name: "Sales".to_string(),
    display_name: "Sales".to_string(),
    range: "A1:C10".to_string(),
    columns: vec![
        TableColumn { name: "Product".to_string(), totals_row_function: None, totals_row_label: None },
        TableColumn { name: "Quantity".to_string(), totals_row_function: None, totals_row_label: None },
        TableColumn { name: "Price".to_string(), totals_row_function: None, totals_row_label: None },
    ],
    show_header_row: true,
    style_name: Some("TableStyleMedium2".to_string()),
    auto_filter: true,
    ..TableConfig::default()
};
wb.add_table("Sheet1", &config)?;
```

**TypeScript:**

```typescript
wb.addTable("Sheet1", {
    name: "Sales",
    displayName: "Sales",
    range: "A1:C10",
    columns: [
        { name: "Product" },
        { name: "Quantity" },
        { name: "Price" },
    ],
    showHeaderRow: true,
    styleName: "TableStyleMedium2",
    autoFilter: true,
});
```

### `get_tables(sheet)` / `getTables(sheet)`

List all tables on a sheet. Returns metadata for each table.

**Rust:**

```rust
let tables = wb.get_tables("Sheet1")?;
for t in &tables {
    println!("{}: {} ({})", t.name, t.range, t.columns.join(", "));
}
```

**TypeScript:**

```typescript
const tables = wb.getTables("Sheet1");
for (const t of tables) {
    console.log(`${t.name}: ${t.range}`);
}
```

### `delete_table(sheet, name)` / `deleteTable(sheet, name)`

Remove a table from a sheet by name. Returns an error if the table is not found on the specified sheet.

**Rust:**

```rust
wb.delete_table("Sheet1", "Sales")?;
```

**TypeScript:**

```typescript
wb.deleteTable("Sheet1", "Sales");
```

### TableConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `name` | `String` | `string` | Internal table name (must be unique in workbook) |
| `display_name` | `String` | `string` | Display name shown in the UI |
| `range` | `String` | `string` | Cell range (e.g. "A1:D10") |
| `columns` | `Vec<TableColumn>` | `TableColumn[]` | Column definitions |
| `show_header_row` | `bool` | `boolean?` | Show header row (default: true) |
| `style_name` | `Option<String>` | `string?` | Table style (e.g. "TableStyleMedium2") |
| `auto_filter` | `bool` | `boolean?` | Enable auto-filter (default: true) |
| `show_first_column` | `bool` | `boolean?` | Highlight first column (default: false) |
| `show_last_column` | `bool` | `boolean?` | Highlight last column (default: false) |
| `show_row_stripes` | `bool` | `boolean?` | Show row stripes (default: true) |
| `show_column_stripes` | `bool` | `boolean?` | Show column stripes (default: false) |

### TableColumn

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `name` | `String` | `string` | Column header name |
| `totals_row_function` | `Option<String>` | `string?` | Totals row function (e.g. "sum", "count", "average") |
| `totals_row_label` | `Option<String>` | `string?` | Totals row label (for the first column) |

### TableInfo (returned by `get_tables`)

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `name` | `String` | `string` | Table name |
| `display_name` | `String` | `string` | Display name |
| `range` | `String` | `string` | Cell range |
| `show_header_row` | `bool` | `boolean` | Whether header row is shown |
| `auto_filter` | `bool` | `boolean` | Whether auto-filter is enabled |
| `columns` | `Vec<String>` | `string[]` | Column header names |
| `style_name` | `Option<String>` | `string \| null` | Table style name |

> Note: Table names must be unique across the entire workbook, not just within a single sheet. When a sheet is deleted, all tables on that sheet are removed automatically.

---

## 17. Data Conversion Utilities (Node.js only)

These convenience methods convert between sheet data and common formats. They are available only in the TypeScript/Node.js bindings.

### `toJSON(sheet, options?)`

Convert a sheet to an array of objects. Each object maps column headers (from the first row) to cell values.

```typescript
const wb = await Workbook.open("data.xlsx");
const records = wb.toJSON("Sheet1");
// [{ Name: "Alice", Age: 30, City: "Seoul" }, ...]

// With options
const records2 = wb.toJSON("Sheet1", { headerRow: 2, range: "A2:C100" });
```

### `toCSV(sheet, options?)`

Convert a sheet to a CSV string. Values are quoted as needed and separated by commas.

```typescript
const csv = wb.toCSV("Sheet1");
// "Name,Age,City\nAlice,30,Seoul\n..."

// With custom separator
const tsv = wb.toCSV("Sheet1", { separator: "\t" });
```

### `toHTML(sheet, options?)`

Convert a sheet to an HTML `<table>` string. All text content is XSS-safe (HTML-escaped).

```typescript
const html = wb.toHTML("Sheet1");
// "<table><thead><tr><th>Name</th>..."

// With CSS class
const html2 = wb.toHTML("Sheet1", { tableClass: "data-table" });
```

### `fromJSON(sheet, data, options?)`

Write an array of objects to a sheet. Keys become the header row, values fill the data rows.

```typescript
const wb = new Workbook();
wb.fromJSON("Sheet1", [
    { Name: "Alice", Age: 30, City: "Seoul" },
    { Name: "Bob", Age: 25, City: "Busan" },
]);
await wb.save("output.xlsx");

// With options
wb.fromJSON("Sheet1", data, { startCell: "B2", writeHeaders: true });
```

### Conversion Options

**ToJSONOptions:**

| Field | Type | Default | Description |
|---|---|---|---|
| `headerRow` | `number` | `1` | Row number to use as column headers (1-based) |
| `range` | `string?` | `undefined` | Limit to a specific cell range |

**ToCSVOptions:**

| Field | Type | Default | Description |
|---|---|---|---|
| `separator` | `string` | `","` | Field separator |
| `lineEnding` | `string` | `"\n"` | Line ending |
| `escapeFormulas` | `boolean` | `false` | Prefix cells starting with `=`, `+`, `-`, or `@` with a tab character to prevent formula injection in downstream tools |

**ToHTMLOptions:**

| Field | Type | Default | Description |
|---|---|---|---|
| `tableClass` | `string?` | `undefined` | CSS class for the `<table>` element |
| `includeHeaders` | `boolean` | `true` | Include a `<thead>` section |

**FromJSONOptions:**

| Field | Type | Default | Description |
|---|---|---|---|
| `startCell` | `string` | `"A1"` | Top-left cell to start writing |
| `writeHeaders` | `boolean` | `true` | Write object keys as the header row |

---
