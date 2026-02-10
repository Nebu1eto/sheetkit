## 도형 (Shapes)

시트에 preset geometry 도형을 삽입합니다. 도형은 두 셀(좌상단, 우하단) 사이에 anchor되며, 텍스트, 채우기 색상, 선 스타일을 포함할 수 있습니다.

### 지원 도형 타입

| 타입 | Preset 이름 | 설명 |
|------|------------|------|
| `rect` | `rect` | 직사각형 |
| `roundRect` | `roundRect` | 둥근 직사각형 |
| `ellipse` | `ellipse` | 타원 / 원 |
| `triangle` | `triangle` | 삼각형 |
| `diamond` | `diamond` | 다이아몬드 |
| `pentagon` | `pentagon` | 오각형 |
| `hexagon` | `hexagon` | 육각형 |
| `octagon` | `octagon` | 팔각형 |
| `rightArrow` | `rightArrow` | 오른쪽 화살표 |
| `leftArrow` | `leftArrow` | 왼쪽 화살표 |
| `upArrow` | `upArrow` | 위쪽 화살표 |
| `downArrow` | `downArrow` | 아래쪽 화살표 |
| `leftRightArrow` | `leftRightArrow` | 좌우 화살표 |
| `upDownArrow` | `upDownArrow` | 상하 화살표 |
| `star4` | `star4` | 4각 별 |
| `star5` | `star5` | 5각 별 |
| `star6` | `star6` | 6각 별 |
| `flowChartProcess` | `flowChartProcess` | 순서도 프로세스 |
| `flowChartDecision` | `flowChartDecision` | 순서도 판단 |
| `flowChartTerminator` | `flowChartTerminator` | 순서도 종료 |
| `flowChartData` | `flowChartInputOutput` | 순서도 데이터 (I/O) |
| `heart` | `heart` | 하트 |
| `lightning` | `lightningBolt` | 번개 |
| `plus` | `mathPlus` | 더하기 |
| `minus` | `mathMinus` | 빼기 |
| `cloud` | `cloud` | 구름 |
| `callout1` | `wedgeRectCallout` | 직사각형 말풍선 |
| `callout2` | `wedgeRoundRectCallout` | 둥근 직사각형 말풍선 |

도형 타입 문자열은 대소문자를 구분하지 않습니다. `"rectangle"` (`"rect"` 대체), `"circle"` 또는 `"oval"` (`"ellipse"` 대체) 등의 별칭도 허용됩니다.

### `add_shape(sheet, config)` / `addShape(sheet, config)`

시트에 도형을 추가합니다. 도형은 두 셀 사이에 anchor됩니다.

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

### ShapeConfig 구조

| 속성 | 타입 | 필수 | 설명 |
|------|------|------|------|
| `shape_type` / `shapeType` | `ShapeType` / `string` | 예 | Preset geometry 타입 (위 표 참조) |
| `from_cell` / `fromCell` | `string` | 예 | 좌상단 anchor 셀 (예: `"B2"`) |
| `to_cell` / `toCell` | `string` | 예 | 우하단 anchor 셀 (예: `"F10"`) |
| `text` | `string?` | 아니요 | 도형 내부에 표시되는 텍스트 |
| `fill_color` / `fillColor` | `string?` | 아니요 | 채우기 색상 hex 값 (예: `"4472C4"`) |
| `line_color` / `lineColor` | `string?` | 아니요 | 선/테두리 색상 hex 값 (예: `"2F528F"`) |
| `line_width` / `lineWidth` | `f64` / `number?` | 아니요 | 선 너비 (포인트 단위) |

도형은 차트나 이미지와 달리 외부 relationship 항목이 필요하지 않습니다. drawing XML에 직접 포함됩니다.

### 참고사항

- 도형은 OOXML `<xdr:sp>` 요소를 사용하며, `<xdr:twoCellAnchor>` 내부에 배치됩니다.
- 채우기 및 선 색상은 sRGB hex 값을 사용합니다 (6자리, `#` 접두사 없음).
- 선 너비는 포인트 단위로 지정됩니다 (1 포인트 = 12700 EMU).
- 같은 시트에 여러 도형을 추가할 수 있습니다.
- 도형은 같은 시트의 차트 및 이미지와 공존할 수 있습니다.

---
