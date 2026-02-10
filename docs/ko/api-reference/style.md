## 7. 스타일

셀의 폰트, 채우기, 테두리, 정렬, 숫자 서식 및 보호 설정을 다룹니다. 스타일은 먼저 `add_style`로 등록한 후 셀에 적용하는 2단계 방식입니다. 동일한 스타일은 자동으로 중복 제거됩니다.

### `add_style(style)` / `addStyle(style)`

스타일을 등록하고 스타일 ID를 반환합니다.

**Rust:**

```rust
use sheetkit::style::*;

let style = Style {
    font: Some(FontStyle {
        name: Some("Arial".into()),
        size: Some(14.0),
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        color: Some(StyleColor::Rgb("#FF0000".into())),
    }),
    fill: Some(FillStyle {
        pattern: PatternType::Solid,
        fg_color: Some(StyleColor::Rgb("#FFFF00".into())),
        bg_color: None,
    }),
    border: Some(BorderStyle {
        left: Some(BorderSideStyle {
            style: BorderLineStyle::Thin,
            color: Some(StyleColor::Rgb("#000000".into())),
        }),
        right: None,
        top: None,
        bottom: None,
        diagonal: None,
    }),
    alignment: Some(AlignmentStyle {
        horizontal: Some(HorizontalAlign::Center),
        vertical: Some(VerticalAlign::Center),
        wrap_text: true,
        text_rotation: None,
        indent: None,
        shrink_to_fit: false,
    }),
    num_fmt: Some(NumFmtStyle::Custom("#,##0.00".into())),
    protection: Some(ProtectionStyle {
        locked: true,
        hidden: false,
    }),
};
let style_id = wb.add_style(&style)?;
```

**TypeScript:**

```typescript
const styleId = wb.addStyle({
    font: {
        name: "Arial",
        size: 14,
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        color: "#FF0000",
    },
    fill: {
        pattern: "solid",
        fgColor: "#FFFF00",
    },
    border: {
        left: { style: "thin", color: "#000000" },
    },
    alignment: {
        horizontal: "center",
        vertical: "center",
        wrapText: true,
    },
    customNumFmt: "#,##0.00",
    protection: {
        locked: true,
        hidden: false,
    },
});
```

### `set_cell_style` / `get_cell_style`

셀에 스타일 ID를 적용하거나 조회합니다.

**Rust:**

```rust
wb.set_cell_style("Sheet1", "A1", style_id)?;
let sid: Option<u32> = wb.get_cell_style("Sheet1", "A1")?;
```

**TypeScript:**

```typescript
wb.setCellStyle("Sheet1", "A1", styleId);
const sid: number | null = wb.getCellStyle("Sheet1", "A1");
```

### 스타일 구성 요소 테이블

#### Font (폰트)

| 속성 | 타입 | 설명 |
|------|------|------|
| `name` | `string?` | 폰트 이름 (예: "Calibri", "Arial") |
| `size` | `f64?` / `number?` | 폰트 크기 (포인트) |
| `bold` | `bool` / `boolean?` | 굵게 |
| `italic` | `bool` / `boolean?` | 기울임 |
| `underline` | `bool` / `boolean?` | 밑줄 |
| `strikethrough` | `bool` / `boolean?` | 취소선 |
| `color` | `StyleColor?` / `string?` | 폰트 색상 |

> TypeScript에서 색상은 문자열로 지정합니다: `"#RRGGBB"` (RGB), `"theme:N"` (테마), `"indexed:N"` (인덱스).

#### Fill (채우기)

| 속성 | 타입 | 설명 |
|------|------|------|
| `pattern` | `PatternType` / `string?` | 패턴 종류 |
| `fg_color` / `fgColor` | `StyleColor?` / `string?` | 전경색 |
| `bg_color` / `bgColor` | `StyleColor?` / `string?` | 배경색 |

**PatternType 값:**

| 값 | 설명 |
|----|------|
| `none` | 없음 |
| `solid` | 단색 채우기 |
| `gray125` | 12.5% 회색 |
| `darkGray` | 진한 회색 |
| `mediumGray` | 중간 회색 |
| `lightGray` | 연한 회색 |

#### Border (테두리)

| 속성 | 타입 | 설명 |
|------|------|------|
| `left` | `BorderSideStyle?` | 왼쪽 테두리 |
| `right` | `BorderSideStyle?` | 오른쪽 테두리 |
| `top` | `BorderSideStyle?` | 위쪽 테두리 |
| `bottom` | `BorderSideStyle?` | 아래쪽 테두리 |
| `diagonal` | `BorderSideStyle?` | 대각선 테두리 |

각 `BorderSideStyle`은 `style`과 `color`를 포함합니다.

**BorderLineStyle 값:**

| 값 | 설명 |
|----|------|
| `thin` | 가는 선 |
| `medium` | 중간 선 |
| `thick` | 굵은 선 |
| `dashed` | 파선 |
| `dotted` | 점선 |
| `double` | 이중선 |
| `hair` | 머리카락 선 |
| `mediumDashed` | 중간 파선 |
| `dashDot` | 일점쇄선 |
| `mediumDashDot` | 중간 일점쇄선 |
| `dashDotDot` | 이점쇄선 |
| `mediumDashDotDot` | 중간 이점쇄선 |
| `slantDashDot` | 기울어진 일점쇄선 |

#### Alignment (정렬)

| 속성 | 타입 | 설명 |
|------|------|------|
| `horizontal` | `HorizontalAlign?` / `string?` | 가로 정렬 |
| `vertical` | `VerticalAlign?` / `string?` | 세로 정렬 |
| `wrap_text` / `wrapText` | `bool` / `boolean?` | 텍스트 줄바꿈 |
| `text_rotation` / `textRotation` | `u32?` / `number?` | 텍스트 회전 각도 |
| `indent` | `u32?` / `number?` | 들여쓰기 수준 |
| `shrink_to_fit` / `shrinkToFit` | `bool` / `boolean?` | 셀에 맞춰 축소 |

**HorizontalAlign 값:** `general`, `left`, `center`, `right`, `fill`, `justify`, `centerContinuous`, `distributed`

**VerticalAlign 값:** `top`, `center`, `bottom`, `justify`, `distributed`

#### NumFmt (숫자 서식)

Rust에서는 `NumFmtStyle` 열거형을 사용합니다:
- `NumFmtStyle::Builtin(id)` -- 내장 서식 ID 사용
- `NumFmtStyle::Custom(code)` -- 사용자 정의 서식 코드

TypeScript에서는 두 가지 방식으로 지정합니다:
- `numFmtId: number` -- 내장 서식 ID
- `customNumFmt: string` -- 사용자 정의 서식 코드 (우선 적용)

**주요 내장 숫자 서식 ID:**

| ID | 서식 | 설명 |
|----|------|------|
| 0 | General | 일반 |
| 1 | 0 | 정수 |
| 2 | 0.00 | 소수 2자리 |
| 3 | #,##0 | 천 단위 구분 |
| 4 | #,##0.00 | 천 단위 구분 + 소수 |
| 9 | 0% | 백분율 |
| 10 | 0.00% | 소수 백분율 |
| 11 | 0.00E+00 | 과학적 표기 |
| 14 | m/d/yyyy | 날짜 |
| 15 | d-mmm-yy | 날짜 |
| 20 | h:mm | 시각 |
| 21 | h:mm:ss | 시각(초) |
| 22 | m/d/yyyy h:mm | 날짜+시각 |
| 49 | @ | 텍스트 |

#### Protection (보호)

| 속성 | 타입 | 설명 |
|------|------|------|
| `locked` | `bool` / `boolean?` | 셀 잠금 (기본값: true) |
| `hidden` | `bool` / `boolean?` | 수식 숨기기 |

> 셀 보호는 시트 보호가 활성화된 경우에만 효과가 있습니다.

---
