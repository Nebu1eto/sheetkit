## 7. Styles

Styles control the visual formatting of cells. A style is registered once with `add_style`, which returns a numeric style ID. That ID is then applied to cells, rows, or columns.

### `add_style(style)` / `addStyle(style)`

Register a style definition and return its style ID. Identical styles are deduplicated: registering the same style twice returns the same ID.

**Rust:**

```rust
use sheetkit::style::*;

let style = Style {
    font: Some(FontStyle {
        name: Some("Arial".to_string()),
        size: Some(12.0),
        bold: true,
        color: Some(StyleColor::Rgb("#FF0000".to_string())),
        ..Default::default()
    }),
    fill: Some(FillStyle {
        pattern: PatternType::Solid,
        fg_color: Some(StyleColor::Rgb("#FFFF00".to_string())),
        bg_color: None,
    }),
    num_fmt: Some(NumFmtStyle::Custom("#,##0.00".to_string())),
    ..Default::default()
};
let style_id: u32 = wb.add_style(&style)?;
```

**TypeScript:**

```typescript
const styleId = wb.addStyle({
    font: {
        name: "Arial",
        size: 12,
        bold: true,
        color: "#FF0000",
    },
    fill: {
        pattern: "solid",
        fgColor: "#FFFF00",
    },
    customNumFmt: "#,##0.00",
});
```

### `set_cell_style` / `get_cell_style`

Apply a style ID to a single cell, or get the style ID currently applied to a cell. Returns `None`/`null` for the default style.

**Rust:**

```rust
wb.set_cell_style("Sheet1", "A1", style_id)?;
let current: Option<u32> = wb.get_cell_style("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.setCellStyle("Sheet1", "A1", styleId);
const current: number | null = wb.getCellStyle("Sheet1", "A1");
```

### Style Components Reference

#### FontStyle

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `name` | `Option<String>` | `string?` | Font family (e.g., "Calibri") |
| `size` | `Option<f64>` | `number?` | Font size in points |
| `bold` | `bool` | `boolean?` | Bold text |
| `italic` | `bool` | `boolean?` | Italic text |
| `underline` | `bool` | `boolean?` | Underline text |
| `strikethrough` | `bool` | `boolean?` | Strikethrough text |
| `color` | `Option<StyleColor>` | `string?` | Font color |

**StyleColor (Rust):** `StyleColor::Rgb("#FF0000".into())`, `StyleColor::Theme(1)`, `StyleColor::Indexed(8)`

**Color strings (TypeScript):** `"#FF0000"` (RGB hex), `"theme:1"` (theme color), `"indexed:8"` (indexed color)

#### FillStyle

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `pattern` | `PatternType` | `string?` | Fill pattern type |
| `fg_color` | `Option<StyleColor>` | `string?` | Foreground color |
| `bg_color` | `Option<StyleColor>` | `string?` | Background color |

**PatternType values:**

| Rust | TypeScript | Description |
|---|---|---|
| `PatternType::None` | `"none"` | No fill |
| `PatternType::Solid` | `"solid"` | Solid fill |
| `PatternType::Gray125` | `"gray125"` | 12.5% gray |
| `PatternType::DarkGray` | `"darkGray"` | Dark gray |
| `PatternType::MediumGray` | `"mediumGray"` | Medium gray |
| `PatternType::LightGray` | `"lightGray"` | Light gray |

#### BorderStyle

Each side (`left`, `right`, `top`, `bottom`, `diagonal`) accepts a `BorderSideStyle`:

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `style` | `BorderLineStyle` | `string?` | Line style |
| `color` | `Option<StyleColor>` | `string?` | Border color |

**BorderLineStyle values:**

| Rust | TypeScript |
|---|---|
| `BorderLineStyle::Thin` | `"thin"` |
| `BorderLineStyle::Medium` | `"medium"` |
| `BorderLineStyle::Thick` | `"thick"` |
| `BorderLineStyle::Dashed` | `"dashed"` |
| `BorderLineStyle::Dotted` | `"dotted"` |
| `BorderLineStyle::Double` | `"double"` |
| `BorderLineStyle::Hair` | `"hair"` |
| `BorderLineStyle::MediumDashed` | `"mediumDashed"` |
| `BorderLineStyle::DashDot` | `"dashDot"` |
| `BorderLineStyle::MediumDashDot` | `"mediumDashDot"` |
| `BorderLineStyle::DashDotDot` | `"dashDotDot"` |
| `BorderLineStyle::MediumDashDotDot` | `"mediumDashDotDot"` |
| `BorderLineStyle::SlantDashDot` | `"slantDashDot"` |

**Rust example:**

```rust
use sheetkit::style::*;

let style = Style {
    border: Some(BorderStyle {
        top: Some(BorderSideStyle {
            style: BorderLineStyle::Thin,
            color: Some(StyleColor::Rgb("#000000".to_string())),
        }),
        bottom: Some(BorderSideStyle {
            style: BorderLineStyle::Double,
            color: Some(StyleColor::Rgb("#0000FF".to_string())),
        }),
        ..Default::default()
    }),
    ..Default::default()
};
```

**TypeScript example:**

```typescript
const styleId = wb.addStyle({
    border: {
        top: { style: "thin", color: "#000000" },
        bottom: { style: "double", color: "#0000FF" },
    },
});
```

#### AlignmentStyle

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `horizontal` | `Option<HorizontalAlign>` | `string?` | Horizontal alignment |
| `vertical` | `Option<VerticalAlign>` | `string?` | Vertical alignment |
| `wrap_text` | `bool` | `boolean?` | Enable text wrapping |
| `text_rotation` | `Option<u32>` | `number?` | Rotation angle in degrees |
| `indent` | `Option<u32>` | `number?` | Indentation level |
| `shrink_to_fit` | `bool` | `boolean?` | Shrink text to fit cell width |

**HorizontalAlign values:** `General`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterContinuous`, `Distributed`
In TypeScript: `"general"`, `"left"`, `"center"`, `"right"`, `"fill"`, `"justify"`, `"centerContinuous"`, `"distributed"`

**VerticalAlign values:** `Top`, `Center`, `Bottom`, `Justify`, `Distributed`
In TypeScript: `"top"`, `"center"`, `"bottom"`, `"justify"`, `"distributed"`

#### NumFmtStyle

Number formats control how values are displayed.

**Rust:**

```rust
use sheetkit::style::NumFmtStyle;

// Built-in format by ID
NumFmtStyle::Builtin(9)  // 0%

// Custom format string
NumFmtStyle::Custom("#,##0.00".to_string())
```

**TypeScript:**

Use `numFmtId` for built-in formats or `customNumFmt` for custom format strings on the style object:

```typescript
// Built-in format
wb.addStyle({ numFmtId: 9 }); // 0%

// Custom format
wb.addStyle({ customNumFmt: "#,##0.00" });
```

**Common built-in format IDs:**

| ID | Format | Description |
|---|---|---|
| 0 | General | General |
| 1 | 0 | Integer |
| 2 | 0.00 | 2 decimal places |
| 3 | #,##0 | Thousands separator |
| 4 | #,##0.00 | Thousands with 2 decimals |
| 9 | 0% | Percentage |
| 10 | 0.00% | Percentage with 2 decimals |
| 11 | 0.00E+00 | Scientific notation |
| 14 | m/d/yyyy | Date |
| 15 | d-mmm-yy | Date |
| 20 | h:mm | Time |
| 21 | h:mm:ss | Time |
| 22 | m/d/yyyy h:mm | Date and time |
| 49 | @ | Text |

#### ProtectionStyle

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `locked` | `bool` | `boolean?` | Lock the cell (default: true) |
| `hidden` | `bool` | `boolean?` | Hide formulas in protected sheet view |

---
