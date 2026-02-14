## Images

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

### `delete_picture` / `deletePicture`

Delete a picture anchored at the given cell. Removes the drawing anchor, image data, relationship entry, and content type for the picture. Searches both one-cell and two-cell anchors.

Returns an error if no picture is found at the specified cell.

**Parameters:**

| Parameter | Rust Type | TS Type | Description |
|---|---|---|---|
| `sheet` | `&str` | `string` | Sheet name |
| `cell` | `&str` | `string` | Anchor cell of the picture (e.g., `"B2"`) |

**Rust:**

```rust
wb.delete_picture("Sheet1", "B2")?;
```

**TypeScript:**

```typescript
wb.deletePicture("Sheet1", "B2");
```

### `get_pictures` / `getPictures`

Get all pictures anchored at the given cell. Returns picture data, format, anchor cell reference, and dimensions for each picture found.

**Parameters:**

| Parameter | Rust Type | TS Type | Description |
|---|---|---|---|
| `sheet` | `&str` | `string` | Sheet name |
| `cell` | `&str` | `string` | Anchor cell to query (e.g., `"B2"`) |

**Returns:** `Vec<PictureInfo>` (Rust) / `JsPictureInfo[]` (TypeScript)

**Rust:**

```rust
let pictures = wb.get_pictures("Sheet1", "B2")?;
for pic in &pictures {
    println!("Format: {}, Cell: {}, Size: {}x{}",
        pic.format.extension(), pic.cell, pic.width_px, pic.height_px);
    // pic.data contains the raw image bytes
}
```

**TypeScript:**

```typescript
const pictures = wb.getPictures("Sheet1", "B2");
for (const pic of pictures) {
    console.log(`Format: ${pic.format}, Cell: ${pic.cell}, Size: ${pic.widthPx}x${pic.heightPx}`);
    // pic.data is a Buffer containing the raw image bytes
}
```

### `get_picture_cells` / `getPictureCells`

Get all cells that have pictures anchored to them on the given sheet.

**Parameters:**

| Parameter | Rust Type | TS Type | Description |
|---|---|---|---|
| `sheet` | `&str` | `string` | Sheet name |

**Returns:** `Vec<String>` (Rust) / `string[]` (TypeScript) -- list of cell references (e.g., `["A1", "B2", "D5"]`)

**Rust:**

```rust
let cells = wb.get_picture_cells("Sheet1")?;
for cell in &cells {
    println!("Picture at: {}", cell);
}
```

**TypeScript:**

```typescript
const cells = wb.getPictureCells("Sheet1");
for (const cell of cells) {
    console.log(`Picture at: ${cell}`);
}
```

### PictureInfo / JsPictureInfo

Information about a picture retrieved from a worksheet.

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `data` | `Vec<u8>` | `Buffer` | Raw image bytes |
| `format` | `ImageFormat` | `string` | Image format (Rust enum / format extension string) |
| `cell` | `String` | `string` | Anchor cell reference (e.g., `"B2"`) |
| `width_px` / `widthPx` | `u32` | `number` | Image width in pixels |
| `height_px` / `heightPx` | `u32` | `number` | Image height in pixels |

---
