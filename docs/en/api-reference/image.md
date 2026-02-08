## 11. Images

Embed images into worksheets.

### Supported Formats

| Format | Extensions | MIME Type |
|--------|-----------|-----------|
| PNG | `png` | `image/png` |
| JPEG | `jpeg`, `jpg` | `image/jpeg` |
| GIF | `gif` | `image/gif` |
| BMP | `bmp` | `image/bmp` |
| ICO | `ico` | `image/x-icon` |
| TIFF | `tiff`, `tif` | `image/tiff` |
| SVG | `svg` | `image/svg+xml` |
| EMF | `emf` | `image/x-emf` |
| EMZ | `emz` | `image/x-emz` |
| WMF | `wmf` | `image/x-wmf` |
| WMZ | `wmz` | `image/x-wmz` |

Format strings are case-insensitive (e.g., `"PNG"`, `"Svg"`, `"TIFF"` all work).

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
| `format` | `ImageFormat` | `string` | See supported formats table above |
| `from_cell` | `String` | `string` | Anchor cell (top-left corner) |
| `width_px` | `u32` | `number` | Image width in pixels |
| `height_px` | `u32` | `number` | Image height in pixels |

Passing an unsupported format string returns an error with a message indicating the unrecognised format.

---
