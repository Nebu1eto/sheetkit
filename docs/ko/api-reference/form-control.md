## 양식 컨트롤 (Form Controls)

시트에 레거시 양식 컨트롤을 삽입합니다. 양식 컨트롤은 VML (Vector Markup Language) drawing 파트를 사용하며, 버튼, 체크 박스, 옵션 버튼, 스핀 버튼, 스크롤 바, 그룹 박스, 레이블을 지원합니다.

### 지원 컨트롤 타입

| 타입 문자열 | Rust Enum | 설명 |
|-------------|-----------|------|
| `button` | `FormControlType::Button` | 명령 버튼 |
| `checkbox` | `FormControlType::CheckBox` | 체크 박스 |
| `optionButton` | `FormControlType::OptionButton` | 옵션 (라디오) 버튼 |
| `spinButton` | `FormControlType::SpinButton` | 스핀 버튼 |
| `scrollBar` | `FormControlType::ScrollBar` | 스크롤 바 |
| `groupBox` | `FormControlType::GroupBox` | 그룹 박스 |
| `label` | `FormControlType::Label` | 레이블 |

컨트롤 타입 문자열은 대소문자를 구분하지 않습니다. `"radio"` (`"optionButton"` 대체), `"spin"` (`"spinButton"` 대체), `"scroll"` (`"scrollBar"` 대체), `"group"` (`"groupBox"` 대체) 등의 별칭도 허용됩니다.

### `add_form_control(sheet, config)` / `addFormControl(sheet, config)`

시트에 양식 컨트롤을 추가합니다.

**Rust:**

```rust
use sheetkit::{FormControlConfig, FormControlType};

// Button
let config = FormControlConfig::button("B2", "Click Me");
wb.add_form_control("Sheet1", config)?;

// Checkbox with cell link
let mut config = FormControlConfig::checkbox("B4", "Enable Feature");
config.cell_link = Some("$D$4".to_string());
config.checked = Some(true);
wb.add_form_control("Sheet1", config)?;

// Spin button
let config = FormControlConfig::spin_button("E2", 0, 100);
wb.add_form_control("Sheet1", config)?;

// Scroll bar
let config = FormControlConfig::scroll_bar("F2", 0, 200);
wb.add_form_control("Sheet1", config)?;
```

**TypeScript:**

```typescript
// Button
wb.addFormControl("Sheet1", {
    controlType: "button",
    cell: "B2",
    text: "Click Me",
});

// Checkbox with cell link
wb.addFormControl("Sheet1", {
    controlType: "checkbox",
    cell: "B4",
    text: "Enable Feature",
    cellLink: "$D$4",
    checked: true,
});

// Spin button
wb.addFormControl("Sheet1", {
    controlType: "spinButton",
    cell: "E2",
    minValue: 0,
    maxValue: 100,
    increment: 5,
    currentValue: 50,
});

// Scroll bar
wb.addFormControl("Sheet1", {
    controlType: "scrollBar",
    cell: "F2",
    minValue: 0,
    maxValue: 200,
    pageIncrement: 10,
});
```

### `get_form_controls(sheet)` / `getFormControls(sheet)`

시트의 모든 양식 컨트롤을 가져옵니다. `FormControlInfo` 객체 배열을 반환합니다.

**Rust:**

```rust
let controls = wb.get_form_controls("Sheet1")?;
for ctrl in &controls {
    println!("{:?}: {}", ctrl.control_type, ctrl.cell);
}
```

**TypeScript:**

```typescript
const controls = wb.getFormControls("Sheet1");
for (const ctrl of controls) {
    console.log(`${ctrl.controlType}: ${ctrl.cell}`);
}
```

### `delete_form_control(sheet, index)` / `deleteFormControl(sheet, index)`

0부터 시작하는 인덱스로 양식 컨트롤을 삭제합니다.

**Rust:**

```rust
wb.delete_form_control("Sheet1", 0)?;
```

**TypeScript:**

```typescript
wb.deleteFormControl("Sheet1", 0);
```

### FormControlConfig 구조

| 속성 | 타입 | 필수 | 설명 |
|------|------|------|------|
| `control_type` / `controlType` | `FormControlType` / `string` | 예 | 컨트롤 타입 (위 표 참조) |
| `cell` | `string` | 예 | anchor 셀 (좌상단 모서리, 예: `"B2"`) |
| `width` | `f64` / `number?` | 아니요 | 너비 (포인트 단위, 기본값 자동) |
| `height` | `f64` / `number?` | 아니요 | 높이 (포인트 단위, 기본값 자동) |
| `text` | `string?` | 아니요 | 표시 텍스트 (Button, CheckBox, OptionButton, GroupBox, Label) |
| `macro_name` / `macroName` | `string?` | 아니요 | VBA 매크로 이름 (Button만 해당) |
| `cell_link` / `cellLink` | `string?` | 아니요 | 값 바인딩을 위한 연결 셀 (CheckBox, OptionButton, SpinButton, ScrollBar) |
| `checked` | `boolean?` | 아니요 | 초기 선택 상태 (CheckBox, OptionButton) |
| `min_value` / `minValue` | `u32` / `number?` | 아니요 | 최솟값 (SpinButton, ScrollBar) |
| `max_value` / `maxValue` | `u32` / `number?` | 아니요 | 최댓값 (SpinButton, ScrollBar) |
| `increment` | `u32` / `number?` | 아니요 | 단계 증분 (SpinButton, ScrollBar) |
| `page_increment` / `pageIncrement` | `u32` / `number?` | 아니요 | 페이지 증분 (ScrollBar만 해당) |
| `current_value` / `currentValue` | `u32` / `number?` | 아니요 | 현재 값 (SpinButton, ScrollBar) |
| `three_d` / `threeD` | `boolean?` | 아니요 | 3D 음영 활성화 (기본값: true) |

### FormControlInfo 구조

| 속성 | 타입 | 설명 |
|------|------|------|
| `control_type` / `controlType` | `FormControlType` / `string` | 컨트롤 타입 |
| `cell` | `string` | anchor 셀 참조 |
| `text` | `string?` | 표시 텍스트 |
| `macro_name` / `macroName` | `string?` | VBA 매크로 이름 |
| `cell_link` / `cellLink` | `string?` | 연결 셀 참조 |
| `checked` | `boolean?` | 선택 상태 |
| `current_value` / `currentValue` | `u32` / `number?` | 현재 값 |
| `min_value` / `minValue` | `u32` / `number?` | 최솟값 |
| `max_value` / `maxValue` | `u32` / `number?` | 최댓값 |
| `increment` | `u32` / `number?` | 단계 증분 |
| `page_increment` / `pageIncrement` | `u32` / `number?` | 페이지 증분 |

### 참고사항

- 양식 컨트롤은 DrawingML이 아닌 레거시 VML drawing 파트(`xl/drawings/vmlDrawingN.vml`)를 사용합니다.
- 양식 컨트롤은 같은 시트의 주석(comment)과 공존할 수 있습니다. 둘 다 있는 경우 VML 콘텐츠가 하나의 파일로 병합됩니다.
- 양식 컨트롤은 차트, 이미지, 도형과도 같은 시트에서 공존할 수 있습니다.
- `cellLink` 속성은 컨트롤의 값을 시트 셀에 바인딩합니다. 체크 박스의 경우 연결된 셀에 TRUE/FALSE가 설정됩니다. 스핀 버튼과 스크롤 바의 경우 숫자 값이 설정됩니다.
- 버튼 컨트롤은 `macroName` 속성으로 VBA 매크로를 연결할 수 있습니다.
- 너비/높이를 지정하지 않으면 컨트롤 타입별로 기본 크기가 자동 적용됩니다.

---
