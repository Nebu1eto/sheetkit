# SVG Rendering

SheetKit can render a worksheet to SVG for visual preview, thumbnails, or embedding in web applications. The renderer produces a self-contained SVG string from the worksheet's cell values, styles, and layout.

## Basic Usage

```typescript
import { Workbook } from 'sheetkit';

const wb = Workbook.openSync('report.xlsx');
const svg = wb.renderToSvg({ sheetName: 'Sheet1' });

// Write to file
import { writeFileSync } from 'node:fs';
writeFileSync('preview.svg', svg);
```

## Render Options

The `renderToSvg` method accepts a `JsRenderOptions` object:

| Option             | Type               | Default   | Description                                       |
|--------------------|--------------------|-----------|----------------------------------------------------|
| `sheetName`        | `string`           | required  | Name of the sheet to render.                       |
| `range`            | `string \| null`   | `null`    | Cell range to render (e.g. `"A1:F20"`). Omit to render the used range. |
| `showGridlines`    | `boolean \| null`  | `true`    | Whether to draw gridlines between cells.           |
| `showHeaders`      | `boolean \| null`  | `true`    | Whether to draw row/column headers (A, B, 1, 2).  |
| `scale`            | `number \| null`   | `1.0`     | Scale factor (2.0 = double size).                  |
| `defaultFontFamily`| `string \| null`   | `"Arial"` | Default font family for cell text.                 |
| `defaultFontSize`  | `number \| null`   | `11.0`    | Default font size in points.                       |

## Rendering a Sub-Range

Render only a specific area of the sheet:

```typescript
const svg = wb.renderToSvg({
  sheetName: 'Sheet1',
  range: 'A1:D10',
});
```

## Controlling Visual Output

```typescript
// Minimal rendering without headers or gridlines
const svg = wb.renderToSvg({
  sheetName: 'Sheet1',
  showGridlines: false,
  showHeaders: false,
});

// High-resolution rendering (2x scale)
const svg2x = wb.renderToSvg({
  sheetName: 'Sheet1',
  scale: 2,
});
```

## Rust API

The Rust API is available through the `Workbook::render_to_svg` method:

```rust
use sheetkit::Workbook;
use sheetkit::RenderOptions;

let wb = Workbook::open("report.xlsx").unwrap();
let svg = wb.render_to_svg(&RenderOptions {
    sheet_name: "Sheet1".to_string(),
    ..RenderOptions::default()
}).unwrap();
```

The lower-level `render::render_to_svg` function is also available for direct use with `WorksheetXml`, `SharedStringTable`, and `StyleSheet` references.

## Supported Features

The SVG renderer supports the following visual features:

- Cell text values (string, number, boolean, date, formula cached results)
- Column widths and row heights (explicit and defaults)
- Font styles: bold, italic, underline, strikethrough, font color, font name, font size
- Cell fill colors (solid pattern fills)
- Cell borders (left, right, top, bottom) with line style and color
- Text alignment (horizontal: left, center, right; vertical: top, center, bottom)
- Row and column headers with background shading
- Gridlines with configurable visibility
- Scale factor for output dimensions
- Sub-range rendering

## Known Limitations

The following features are not yet supported by the renderer:

- Merged cells (rendered as individual cells)
- Conditional formatting (colors not applied in SVG)
- Images and charts
- Rich text (individual run formatting within a cell)
- Gradient fills
- Theme and indexed color resolution (defaults to black)
- Number format display (raw values shown)
- Text wrapping and overflow
- Diagonal borders
- Hidden rows and columns
- Outline/group collapse
