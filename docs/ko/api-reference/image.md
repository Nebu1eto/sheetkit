## 11. 이미지

시트에 이미지를 삽입하는 기능을 다룬다.

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

형식 문자열은 대소문자를 구분하지 않는다 (예: `"PNG"`, `"Svg"`, `"TIFF"` 모두 허용).

### `add_image(sheet, config)` / `addImage(sheet, config)`

시트에 이미지를 추가한다.

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

지원하지 않는 형식 문자열을 전달하면 해당 형식명을 포함한 오류가 반환된다.

---
