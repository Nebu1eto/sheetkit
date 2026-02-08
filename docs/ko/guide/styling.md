### 스타일

스타일은 셀의 시각적 표현을 제어합니다. 스타일 정의를 등록하면 스타일 ID를 받고, 이 ID를 셀에 적용합니다. 동일한 스타일 정의는 자동으로 중복 제거됩니다.

`Style`은 글꼴, 채우기, 테두리, 정렬, 숫자 형식, 보호 속성을 자유롭게 조합할 수 있습니다.

#### Rust

```rust
use sheetkit::{
    AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle,
    FillStyle, FontStyle, HorizontalAlign, PatternType, Style,
    StyleColor, VerticalAlign, Workbook,
};

let mut wb = Workbook::new();

// 스타일 등록
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

// 셀에 스타일 적용
wb.set_cell_style("Sheet1", "A1", style_id)?;

// 셀의 스타일 ID 조회 (기본 스타일이면 None)
let current_style: Option<u32> = wb.get_cell_style("Sheet1", "A1")?;
```

#### TypeScript

```typescript
// 스타일 등록
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

// 셀에 스타일 적용
wb.setCellStyle('Sheet1', 'A1', styleId);

// 셀의 스타일 ID 조회 (기본 스타일이면 null)
const currentStyle: number | null = wb.getCellStyle('Sheet1', 'A1');
```

#### 스타일 구성 요소 상세

**FontStyle (글꼴)**

| 필드             | Rust 타입           | TS 타입    | 설명                        |
|-----------------|---------------------|------------|----------------------------|
| `name`          | `Option<String>`    | `string?`  | 글꼴 이름 (예: "Calibri")   |
| `size`          | `Option<f64>`       | `number?`  | 글꼴 크기 (포인트)           |
| `bold`          | `bool`              | `boolean?` | 굵게                        |
| `italic`        | `bool`              | `boolean?` | 기울임꼴                     |
| `underline`     | `bool`              | `boolean?` | 밑줄                        |
| `strikethrough` | `bool`              | `boolean?` | 취소선                       |
| `color`         | `Option<StyleColor>` | `string?` | 글꼴 색상 (TS에서는 hex 문자열) |

**FillStyle (채우기)**

| 필드       | Rust 타입           | TS 타입   | 설명                        |
|-----------|---------------------|-----------|----------------------------|
| `pattern` | `PatternType`       | `string?` | 패턴 유형 (아래 값 참조)      |
| `fg_color`| `Option<StyleColor>` | `string?` | 전경 색상                   |
| `bg_color`| `Option<StyleColor>` | `string?` | 배경 색상                   |

PatternType 값: `None`, `Solid`, `Gray125`, `DarkGray`, `MediumGray`, `LightGray`.
TypeScript에서는 소문자 문자열 사용: `"none"`, `"solid"`, `"gray125"`, `"darkGray"`, `"mediumGray"`, `"lightGray"`.

**BorderStyle (테두리)**

각 면(`left`, `right`, `top`, `bottom`, `diagonal`)은 `BorderSideStyle`을 받으며 다음을 포함합니다:
- `style`: `Thin`, `Medium`, `Thick`, `Dashed`, `Dotted`, `Double`, `Hair`, `MediumDashed`, `DashDot`, `MediumDashDot`, `DashDotDot`, `MediumDashDotDot`, `SlantDashDot` 중 하나
- `color`: 선택적 색상

TypeScript에서는 소문자 문자열 사용: `"thin"`, `"medium"`, `"thick"` 등.

**AlignmentStyle (정렬)**

| 필드             | Rust 타입                 | TS 타입    | 설명                  |
|-----------------|--------------------------|------------|----------------------|
| `horizontal`    | `Option<HorizontalAlign>` | `string?` | 가로 정렬             |
| `vertical`      | `Option<VerticalAlign>`   | `string?` | 세로 정렬             |
| `wrap_text`     | `bool`                    | `boolean?`| 텍스트 줄 바꿈        |
| `text_rotation` | `Option<u32>`             | `number?` | 텍스트 회전 각도(도)   |
| `indent`        | `Option<u32>`             | `number?` | 들여쓰기 수준          |
| `shrink_to_fit` | `bool`                    | `boolean?`| 셀에 맞게 텍스트 축소  |

HorizontalAlign 값: `General`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterContinuous`, `Distributed`.
VerticalAlign 값: `Top`, `Center`, `Bottom`, `Justify`, `Distributed`.

**NumFmtStyle (숫자 형식, Rust 전용)**

```rust
use sheetkit::style::NumFmtStyle;

// 기본 제공 형식 (예: 퍼센트, 날짜, 통화)
NumFmtStyle::Builtin(9)  // 0%

// 사용자 정의 형식 문자열
NumFmtStyle::Custom("#,##0.00".to_string())
```

TypeScript에서는 스타일 객체의 `numFmtId`(기본 제공 형식 ID) 또는 `customNumFmt`(사용자 정의 형식 문자열)을 사용합니다.

**ProtectionStyle (보호)**

| 필드     | Rust 타입 | TS 타입    | 설명                              |
|---------|-----------|------------|----------------------------------|
| `locked`| `bool`    | `boolean?` | 셀 잠금 (기본값: true)             |
| `hidden`| `bool`    | `boolean?` | 보호 모드에서 수식 숨김             |

---

### 차트