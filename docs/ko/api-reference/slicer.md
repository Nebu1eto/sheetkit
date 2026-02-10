## 슬라이서 (Slicers)

슬라이서는 사용자가 테이블 데이터를 시각적으로 필터링할 수 있는 컨트롤입니다. SheetKit은 테이블 기반 슬라이서의 추가, 조회, 삭제를 지원합니다 (Excel 2010에서 도입되었습니다).

### `add_slicer(sheet, config)` / `addSlicer(sheet, config)`

워크시트에 테이블 열을 대상으로 하는 슬라이서를 추가합니다.

**Rust:**

```rust
use sheetkit::SlicerConfig;

let config = SlicerConfig {
    name: "StatusFilter".to_string(),
    cell: "F1".to_string(),
    table_name: "Table1".to_string(),
    column_name: "Status".to_string(),
    caption: Some("Status".to_string()),
    style: Some("SlicerStyleLight1".to_string()),
    width: Some(200),
    height: Some(200),
    show_caption: Some(true),
    column_count: None,
};
wb.add_slicer("Sheet1", &config)?;
```

**TypeScript:**

```typescript
wb.addSlicer("Sheet1", {
    name: "StatusFilter",
    cell: "F1",
    tableName: "Table1",
    columnName: "Status",
    caption: "Status",
    style: "SlicerStyleLight1",
    width: 200,
    height: 200,
    showCaption: true,
});
```

### SlicerConfig

| 필드 | Rust 타입 | TS 타입 | 설명 |
|---|---|---|---|
| `name` | `String` | `string` | 슬라이서 고유 이름입니다 |
| `cell` | `String` | `string` | 앵커 셀 (왼쪽 상단, 예: "F1")입니다 |
| `table_name` | `String` | `string` | 소스 테이블 이름입니다 |
| `column_name` | `String` | `string` | 필터링할 테이블 열 이름입니다 |
| `caption` | `Option<String>` | `string?` | 캡션 헤더 텍스트입니다. 기본값은 열 이름입니다 |
| `style` | `Option<String>` | `string?` | 시각적 스타일 (예: "SlicerStyleLight1")입니다 |
| `width` | `Option<u32>` | `number?` | 너비(픽셀)입니다. 기본값은 200입니다 |
| `height` | `Option<u32>` | `number?` | 높이(픽셀)입니다. 기본값은 200입니다 |
| `show_caption` | `Option<bool>` | `boolean?` | 캡션 헤더 표시 여부입니다 |
| `column_count` | `Option<u32>` | `number?` | 슬라이서 표시 열 수입니다 |

### `get_slicers(sheet)` / `getSlicers(sheet)`

워크시트의 모든 슬라이서 정보를 반환합니다.

**Rust:**

```rust
let slicers = wb.get_slicers("Sheet1")?;
for s in &slicers {
    println!("{}: filtering column '{}'", s.name, s.column_name);
}
```

**TypeScript:**

```typescript
const slicers = wb.getSlicers("Sheet1");
for (const s of slicers) {
    console.log(`${s.name}: filtering column '${s.columnName}'`);
}
```

### SlicerInfo

| 필드 | Rust 타입 | TS 타입 | 설명 |
|---|---|---|---|
| `name` | `String` | `string` | 슬라이서 이름입니다 |
| `caption` | `String` | `string` | 표시 캡션입니다 |
| `table_name` | `String` | `string` | 소스 테이블 이름입니다 |
| `column_name` | `String` | `string` | 필터링 중인 열 이름입니다 |
| `style` | `Option<String>` | `string \| null` | 시각적 스타일입니다 (설정된 경우) |

### `delete_slicer(sheet, name)` / `deleteSlicer(sheet, name)`

워크시트에서 이름으로 슬라이서를 삭제합니다. 슬라이서 정의, 캐시, content type, 관계가 모두 제거됩니다.

**Rust:**

```rust
wb.delete_slicer("Sheet1", "StatusFilter")?;
```

**TypeScript:**

```typescript
wb.deleteSlicer("Sheet1", "StatusFilter");
```

### 슬라이서 스타일

Excel은 기본 슬라이서 스타일을 제공합니다. `style` 필드에 다음 중 하나를 전달하세요:

- `SlicerStyleLight1` ~ `SlicerStyleLight6`
- `SlicerStyleDark1` ~ `SlicerStyleDark6`
- `SlicerStyleOther1` ~ `SlicerStyleOther2`

### 참고 사항

- 슬라이서는 Excel 2010 이상의 기능입니다. 슬라이서가 포함된 파일은 이전 버전의 스프레드시트 애플리케이션에서 올바르게 표시되지 않을 수 있습니다.
- 슬라이서는 OOXML 파트(`xl/slicers/`, `xl/slicerCaches/`)로 저장되며 적절한 content type과 관계가 설정됩니다.
- 동일한 워크시트에 여러 슬라이서를 추가할 수 있으며, 각각 다른 테이블 열을 필터링합니다.
- 슬라이서는 저장/열기 라운드트립을 지원합니다.

---
