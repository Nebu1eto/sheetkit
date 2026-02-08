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

### Supported Chart Types (41 types)

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

Embed images into worksheets. Supported formats: PNG, JPEG, and GIF.

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
    format: "png",   // "png" | "jpeg" | "gif"
    fromCell: "B2",
    widthPx: 200,
    heightPx: 100,
});
```

### ImageConfig

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `data` | `Vec<u8>` | `Buffer` | Raw image bytes |
| `format` | `ImageFormat` | `string` | `Png`/`"png"`, `Jpeg`/`"jpeg"`, `Gif`/`"gif"` |
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

> Note (Node.js): `validationType` must be a supported value (`list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom`). Unknown values return an error.

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
| `ValidationType::Whole` | `"whole"` | Whole number |
| `ValidationType::Decimal` | `"decimal"` | Decimal number |
| `ValidationType::List` | `"list"` | Dropdown list |
| `ValidationType::Date` | `"date"` | Date value |
| `ValidationType::Time` | `"time"` | Time value |
| `ValidationType::TextLength` | `"textlength"` | Text length constraint |
| `ValidationType::Custom` | `"custom"` | Custom formula |

### Validation Operators

Used with `Whole`, `Decimal`, `Date`, `Time`, and `TextLength` types:

| Rust | TypeScript |
|---|---|
| `ValidationOperator::Between` | `"between"` |
| `ValidationOperator::NotBetween` | `"notbetween"` |
| `ValidationOperator::Equal` | `"equal"` |
| `ValidationOperator::NotEqual` | `"notequal"` |
| `ValidationOperator::LessThan` | `"lessthan"` |
| `ValidationOperator::LessThanOrEqual` | `"lessthanorequal"` |
| `ValidationOperator::GreaterThan` | `"greaterthan"` |
| `ValidationOperator::GreaterThanOrEqual` | `"greaterthanorequal"` |

### Error Styles

| Rust | TypeScript | Description |
|---|---|---|
| `ErrorStyle::Stop` | `"stop"` | Reject invalid input |
| `ErrorStyle::Warning` | `"warning"` | Warn but allow |
| `ErrorStyle::Information` | `"information"` | Inform only |

---

## 13. Comments

Comments (also known as notes) attach text annotations to individual cells.

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

### Rule Types (18 types)

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
