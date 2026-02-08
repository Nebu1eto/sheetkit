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