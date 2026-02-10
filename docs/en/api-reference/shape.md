## Shapes

Insert preset geometry shapes into worksheets. Shapes are anchored between two cells (top-left and bottom-right) and can include text, fill colors, and line styling.

### Supported Shape Types

| Type | Preset Name | Description |
|------|------------|-------------|
| `rect` | `rect` | Rectangle |
| `roundRect` | `roundRect` | Rounded rectangle |
| `ellipse` | `ellipse` | Ellipse / circle |
| `triangle` | `triangle` | Triangle |
| `diamond` | `diamond` | Diamond |
| `pentagon` | `pentagon` | Pentagon |
| `hexagon` | `hexagon` | Hexagon |
| `octagon` | `octagon` | Octagon |
| `rightArrow` | `rightArrow` | Right arrow |
| `leftArrow` | `leftArrow` | Left arrow |
| `upArrow` | `upArrow` | Up arrow |
| `downArrow` | `downArrow` | Down arrow |
| `leftRightArrow` | `leftRightArrow` | Left-right arrow |
| `upDownArrow` | `upDownArrow` | Up-down arrow |
| `star4` | `star4` | 4-point star |
| `star5` | `star5` | 5-point star |
| `star6` | `star6` | 6-point star |
| `flowChartProcess` | `flowChartProcess` | Flowchart process |
| `flowChartDecision` | `flowChartDecision` | Flowchart decision |
| `flowChartTerminator` | `flowChartTerminator` | Flowchart terminator |
| `flowChartData` | `flowChartInputOutput` | Flowchart data (I/O) |
| `heart` | `heart` | Heart |
| `lightning` | `lightningBolt` | Lightning bolt |
| `plus` | `mathPlus` | Plus sign |
| `minus` | `mathMinus` | Minus sign |
| `cloud` | `cloud` | Cloud |
| `callout1` | `wedgeRectCallout` | Rectangular callout |
| `callout2` | `wedgeRoundRectCallout` | Rounded rectangular callout |

Shape type strings are case-insensitive. Aliases like `"rectangle"` for `"rect"`, `"circle"` for `"ellipse"`, and `"oval"` for `"ellipse"` are also accepted.

### `add_shape` / `addShape`

Add a shape to a sheet, anchored between two cells.

**Rust:**

```rust
use sheetkit::{ShapeConfig, ShapeType};

let config = ShapeConfig {
    shape_type: ShapeType::RoundRect,
    from_cell: "B2".to_string(),
    to_cell: "F10".to_string(),
    text: Some("Hello World".to_string()),
    fill_color: Some("4472C4".to_string()),
    line_color: Some("2F528F".to_string()),
    line_width: Some(1.5),
};
wb.add_shape("Sheet1", &config)?;
```

**TypeScript:**

```typescript
wb.addShape("Sheet1", {
    shapeType: "roundRect",
    fromCell: "B2",
    toCell: "F10",
    text: "Hello World",
    fillColor: "4472C4",
    lineColor: "2F528F",
    lineWidth: 1.5,
});
```

### ShapeConfig

| Field | Rust Type | TS Type | Required | Description |
|---|---|---|---|---|
| `shape_type` / `shapeType` | `ShapeType` | `string` | Yes | Preset geometry type (see table above) |
| `from_cell` / `fromCell` | `String` | `string` | Yes | Top-left anchor cell (e.g., `"B2"`) |
| `to_cell` / `toCell` | `String` | `string` | Yes | Bottom-right anchor cell (e.g., `"F10"`) |
| `text` | `Option<String>` | `string?` | No | Text content displayed inside the shape |
| `fill_color` / `fillColor` | `Option<String>` | `string?` | No | Fill color as hex (e.g., `"4472C4"`) |
| `line_color` / `lineColor` | `Option<String>` | `string?` | No | Line/border color as hex (e.g., `"2F528F"`) |
| `line_width` / `lineWidth` | `Option<f64>` | `number?` | No | Line width in points |

Shapes do not require external relationship entries (unlike charts and images). They are embedded directly in the drawing XML.

### Notes

- Shapes use the OOXML `<xdr:sp>` element inside a `<xdr:twoCellAnchor>`.
- Fill and line colors use sRGB hex values (6 characters, no `#` prefix).
- Line width is specified in points (1 point = 12700 EMU).
- Multiple shapes can be added to the same sheet.
- Shapes can coexist with charts and images on the same sheet.

---
