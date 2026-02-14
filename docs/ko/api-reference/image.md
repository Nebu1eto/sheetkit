## 이미지

시트에 이미지를 삽입하는 기능을 다룹니다.

### 지원 형식

| 형식 | 확장자 | MIME 타입 |
|------|--------|-----------|
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

형식 문자열은 대소문자를 구분하지 않습니다 (예: `"PNG"`, `"Svg"`, `"TIFF"` 모두 허용).

### `add_image(sheet, config)` / `addImage(sheet, config)`

시트에 이미지를 추가합니다.

**Rust:**

```rust
use sheetkit::image::{ImageConfig, ImageFormat};

let data = std::fs::read("logo.png")?;
wb.add_image("Sheet1", &ImageConfig {
    data,
    format: ImageFormat::Png,
    from_cell: "A1".into(),
    width_px: 200,
    height_px: 100,
})?;
```

**TypeScript:**

```typescript
import { readFileSync } from "fs";

const data = readFileSync("logo.png");
wb.addImage("Sheet1", {
    data: data,
    format: "png",   // "png" | "jpeg" | "jpg" | "gif" | "bmp" | "ico" | "tiff" | "tif" | "svg" | "emf" | "emz" | "wmf" | "wmz"
    fromCell: "A1",
    widthPx: 200,
    heightPx: 100,
});
```

### ImageConfig 구조

| 속성 | 타입 | 설명 |
|------|------|------|
| `data` | `Vec<u8>` / `Buffer` | 이미지 바이너리 데이터 |
| `format` | `ImageFormat` / `string` | 위 지원 형식 표 참조 |
| `from_cell` / `fromCell` | `string` | 이미지 시작 위치 셀 |
| `width_px` / `widthPx` | `u32` / `number` | 너비 (픽셀) |
| `height_px` / `heightPx` | `u32` / `number` | 높이 (픽셀) |

지원하지 않는 형식 문자열을 전달하면 해당 형식명을 포함한 오류가 반환됩니다.

### `delete_picture(sheet, cell)` / `deletePicture(sheet, cell)`

지정된 셀에 고정된 이미지를 삭제합니다. drawing anchor, 이미지 데이터, relationship 항목, content type을 모두 제거합니다. one-cell anchor와 two-cell anchor를 모두 검색합니다.

해당 셀에 이미지가 없으면 오류가 반환됩니다.

**매개변수:**

| 매개변수 | 타입 | 설명 |
|----------|------|------|
| `sheet` | `&str` / `string` | 시트 이름 |
| `cell` | `&str` / `string` | 이미지의 anchor 셀 (예: `"B2"`) |

**Rust:**

```rust
wb.delete_picture("Sheet1", "B2")?;
```

**TypeScript:**

```typescript
wb.deletePicture("Sheet1", "B2");
```

### `get_pictures(sheet, cell)` / `getPictures(sheet, cell)`

지정된 셀에 고정된 모든 이미지를 가져옵니다. 이미지 데이터, 형식, anchor 셀 참조, 크기를 포함한 정보가 반환됩니다.

**매개변수:**

| 매개변수 | 타입 | 설명 |
|----------|------|------|
| `sheet` | `&str` / `string` | 시트 이름 |
| `cell` | `&str` / `string` | 조회할 anchor 셀 (예: `"B2"`) |

**반환값:** `Vec<PictureInfo>` (Rust) / `JsPictureInfo[]` (TypeScript)

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

### `get_picture_cells(sheet)` / `getPictureCells(sheet)`

지정된 시트에서 이미지가 고정된 모든 셀 목록을 가져옵니다.

**매개변수:**

| 매개변수 | 타입 | 설명 |
|----------|------|------|
| `sheet` | `&str` / `string` | 시트 이름 |

**반환값:** `Vec<String>` (Rust) / `string[]` (TypeScript) -- 셀 참조 목록 (예: `["A1", "B2", "D5"]`)

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

### PictureInfo / JsPictureInfo 구조

워크시트에서 가져온 이미지 정보를 나타냅니다.

| 속성 | 타입 | 설명 |
|------|------|------|
| `data` | `Vec<u8>` / `Buffer` | 이미지 바이너리 데이터 |
| `format` | `ImageFormat` / `string` | 이미지 형식 (Rust enum / 형식 확장자 문자열) |
| `cell` | `String` / `string` | anchor 셀 참조 (예: `"B2"`) |
| `width_px` / `widthPx` | `u32` / `number` | 이미지 너비 (픽셀) |
| `height_px` / `heightPx` | `u32` / `number` | 이미지 높이 (픽셀) |

---
