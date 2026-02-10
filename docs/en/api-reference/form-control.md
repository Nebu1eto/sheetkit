## Form Controls

Insert legacy form controls into worksheets. Form controls use VML (Vector Markup Language) drawing parts and support buttons, check boxes, option buttons, spin buttons, scroll bars, group boxes, and labels.

### Supported Control Types

| Type String | Rust Enum | Description |
|-------------|-----------|-------------|
| `button` | `FormControlType::Button` | Command button |
| `checkbox` | `FormControlType::CheckBox` | Check box |
| `optionButton` | `FormControlType::OptionButton` | Option (radio) button |
| `spinButton` | `FormControlType::SpinButton` | Spin button |
| `scrollBar` | `FormControlType::ScrollBar` | Scroll bar |
| `groupBox` | `FormControlType::GroupBox` | Group box |
| `label` | `FormControlType::Label` | Label |

Control type strings are case-insensitive. Aliases like `"radio"` for `"optionButton"`, `"spin"` for `"spinButton"`, `"scroll"` for `"scrollBar"`, and `"group"` for `"groupBox"` are also accepted.

### `add_form_control` / `addFormControl`

Add a form control to a sheet.

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

### `get_form_controls` / `getFormControls`

Get all form controls on a sheet. Returns an array of `FormControlInfo` objects.

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

### `delete_form_control` / `deleteFormControl`

Delete a form control by its 0-based index in the control list.

**Rust:**

```rust
wb.delete_form_control("Sheet1", 0)?;
```

**TypeScript:**

```typescript
wb.deleteFormControl("Sheet1", 0);
```

### FormControlConfig

| Field | Rust Type | TS Type | Required | Description |
|---|---|---|---|---|
| `control_type` / `controlType` | `FormControlType` | `string` | Yes | Control type (see table above) |
| `cell` | `String` | `string` | Yes | Anchor cell (top-left corner, e.g., `"B2"`) |
| `width` | `Option<f64>` | `number?` | No | Width in points (auto-sized by default) |
| `height` | `Option<f64>` | `number?` | No | Height in points (auto-sized by default) |
| `text` | `Option<String>` | `string?` | No | Display text (Button, CheckBox, OptionButton, GroupBox, Label) |
| `macro_name` / `macroName` | `Option<String>` | `string?` | No | VBA macro name (Button only) |
| `cell_link` / `cellLink` | `Option<String>` | `string?` | No | Linked cell for value binding (CheckBox, OptionButton, SpinButton, ScrollBar) |
| `checked` | `Option<bool>` | `boolean?` | No | Initial checked state (CheckBox, OptionButton) |
| `min_value` / `minValue` | `Option<u32>` | `number?` | No | Minimum value (SpinButton, ScrollBar) |
| `max_value` / `maxValue` | `Option<u32>` | `number?` | No | Maximum value (SpinButton, ScrollBar) |
| `increment` | `Option<u32>` | `number?` | No | Step increment (SpinButton, ScrollBar) |
| `page_increment` / `pageIncrement` | `Option<u32>` | `number?` | No | Page increment (ScrollBar only) |
| `current_value` / `currentValue` | `Option<u32>` | `number?` | No | Current value (SpinButton, ScrollBar) |
| `three_d` / `threeD` | `Option<bool>` | `boolean?` | No | Enable 3D shading (default: true) |

### FormControlInfo

| Field | Rust Type | TS Type | Description |
|---|---|---|---|
| `control_type` / `controlType` | `FormControlType` | `string` | Control type |
| `cell` | `String` | `string` | Anchor cell reference |
| `text` | `Option<String>` | `string?` | Display text |
| `macro_name` / `macroName` | `Option<String>` | `string?` | VBA macro name |
| `cell_link` / `cellLink` | `Option<String>` | `string?` | Linked cell reference |
| `checked` | `Option<bool>` | `boolean?` | Checked state |
| `current_value` / `currentValue` | `Option<u32>` | `number?` | Current value |
| `min_value` / `minValue` | `Option<u32>` | `number?` | Minimum value |
| `max_value` / `maxValue` | `Option<u32>` | `number?` | Maximum value |
| `increment` | `Option<u32>` | `number?` | Step increment |
| `page_increment` / `pageIncrement` | `Option<u32>` | `number?` | Page increment |

### Notes

- Form controls use legacy VML drawing parts (`xl/drawings/vmlDrawingN.vml`), not DrawingML.
- Form controls can coexist with comments on the same sheet. When both are present, the VML content is merged into a single file.
- Form controls can also coexist with charts, images, and shapes on the same sheet.
- The `cellLink` field binds a control's value to a worksheet cell. For check boxes, the linked cell receives TRUE/FALSE. For spin buttons and scroll bars, it receives the numeric value.
- Button controls support the `macroName` field to associate a VBA macro.
- Default dimensions are automatically applied per control type if width/height are not specified.

---
