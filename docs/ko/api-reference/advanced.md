## 16. 틀 고정

특정 행/열을 고정하여 스크롤 시에도 보이도록 하는 기능을 다룬다.

### `set_panes(sheet, cell)` / `setPanes(sheet, cell)`

틀 고정을 설정한다. 셀 참조는 스크롤 가능한 영역의 왼쪽 상단 셀을 나타낸다.

**Rust:**

```rust
// Freeze first row
wb.set_panes("Sheet1", "A2")?;

// Freeze first column
wb.set_panes("Sheet1", "B1")?;

// Freeze first row and first column
wb.set_panes("Sheet1", "B2")?;
```

**TypeScript:**

```typescript
wb.setPanes("Sheet1", "A2");  // freeze row 1
wb.setPanes("Sheet1", "B1");  // freeze column A
wb.setPanes("Sheet1", "B2");  // freeze row 1 + column A
```

### `unset_panes(sheet)` / `unsetPanes(sheet)`

틀 고정을 제거한다.

**Rust:**

```rust
wb.unset_panes("Sheet1")?;
```

**TypeScript:**

```typescript
wb.unsetPanes("Sheet1");
```

### `get_panes(sheet)` / `getPanes(sheet)`

현재 틀 고정 설정을 조회한다. 설정이 없으면 None / null을 반환한다.

**Rust:**

```rust
if let Some(cell) = wb.get_panes("Sheet1")? {
    println!("Frozen at: {}", cell);  // e.g., "B2"
}
```

**TypeScript:**

```typescript
const pane = wb.getPanes("Sheet1");
if (pane) {
    console.log(`Frozen at: ${pane}`);
}
```

---

## 17. 페이지 레이아웃

인쇄 관련 설정을 다룬다. 여백, 용지 크기, 방향, 배율, 머리글/바닥글, 인쇄 옵션, 페이지 나누기를 포함한다.

### 여백 (Margins)

`set_page_margins` / `get_page_margins`로 페이지 여백을 인치 단위로 설정하거나 조회한다.

**Rust:**

```rust
use sheetkit::page_layout::PageMarginsConfig;

wb.set_page_margins("Sheet1", &PageMarginsConfig {
    left: 0.7,
    right: 0.7,
    top: 0.75,
    bottom: 0.75,
    header: 0.3,
    footer: 0.3,
})?;

let margins = wb.get_page_margins("Sheet1")?;
```

**TypeScript:**

```typescript
wb.setPageMargins("Sheet1", {
    left: 0.7,
    right: 0.7,
    top: 0.75,
    bottom: 0.75,
    header: 0.3,
    footer: 0.3,
});

const margins = wb.getPageMargins("Sheet1");
```

### 페이지 설정 (Page Setup)

용지 크기, 방향, 배율, 페이지 맞춤 설정을 다룬다.

**TypeScript:**

```typescript
wb.setPageSetup("Sheet1", {
    paperSize: "a4",
    orientation: "landscape",
    scale: 80,
    fitToWidth: 1,
    fitToHeight: 0,
});

const setup = wb.getPageSetup("Sheet1");
```

**용지 크기 값:** `letter`, `tabloid`, `legal`, `a3`, `a4`, `a5`, `b4`, `b5`

**방향 값:** `portrait` (세로), `landscape` (가로)

### 인쇄 옵션 (Print Options)

**TypeScript:**

```typescript
wb.setPrintOptions("Sheet1", {
    gridLines: true,
    headings: true,
    horizontalCentered: true,
    verticalCentered: false,
});

const opts = wb.getPrintOptions("Sheet1");
```

| 속성 | 타입 | 설명 |
|------|------|------|
| `grid_lines` / `gridLines` | `bool?` / `boolean?` | 눈금선 인쇄 |
| `headings` | `bool?` / `boolean?` | 행/열 머리글 인쇄 |
| `horizontal_centered` / `horizontalCentered` | `bool?` / `boolean?` | 가로 가운데 정렬 |
| `vertical_centered` / `verticalCentered` | `bool?` / `boolean?` | 세로 가운데 정렬 |

### 머리글/바닥글 (Header/Footer)

```typescript
wb.setHeaderFooter("Sheet1", "&LLeft Text&CCenter Text&RRight Text", "&CPage &P of &N");

const hf = wb.getHeaderFooter("Sheet1");
// hf.header, hf.footer
```

> Excel 서식 코드: `&L` (왼쪽), `&C` (가운데), `&R` (오른쪽), `&P` (현재 페이지), `&N` (총 페이지 수)

### 페이지 나누기 (Page Breaks)

```typescript
wb.insertPageBreak("Sheet1", 20);  // insert break before row 20
wb.insertPageBreak("Sheet1", 40);

const breaks: number[] = wb.getPageBreaks("Sheet1");
// [20, 40]

wb.removePageBreak("Sheet1", 20);
```

**Rust:**

```rust
wb.insert_page_break("Sheet1", 20)?;
let breaks: Vec<u32> = wb.get_page_breaks("Sheet1")?;
wb.remove_page_break("Sheet1", 20)?;
```

---

## 18. 정의된 이름

워크북 내에서 셀 범위에 이름을 부여하는 기능을 다룬다. 워크북 범위(모든 시트에서 사용 가능) 또는 시트 범위(특정 시트에서만 사용 가능)로 정의할 수 있다.

### `set_defined_name` / `setDefinedName`

정의된 이름을 추가하거나 업데이트한다. 동일한 이름과 범위를 가진 항목이 이미 존재하면 값과 주석이 업데이트된다(중복 생성 없음).

**Rust:**

```rust
// Workbook-scoped name
wb.set_defined_name("SalesTotal", "Sheet1!$B$10", None, None)?;

// Sheet-scoped name with comment
wb.set_defined_name("LocalRange", "Sheet1!$A$1:$D$10", Some("Sheet1"), Some("Local data range"))?;
```

**TypeScript:**

```typescript
// Workbook-scoped name
wb.setDefinedName({ name: "SalesTotal", value: "Sheet1!$B$10" });

// Sheet-scoped name with comment
wb.setDefinedName({
    name: "LocalRange",
    value: "Sheet1!$A$1:$D$10",
    scope: "Sheet1",
    comment: "Local data range",
});
```

### `get_defined_name` / `getDefinedName`

이름과 선택적 범위로 정의된 이름을 조회한다. 없으면 `None`/`null`을 반환한다.

**Rust:**

```rust
if let Some(info) = wb.get_defined_name("SalesTotal", None)? {
    println!("Refers to: {}", info.value);
}

// Sheet-scoped name
if let Some(info) = wb.get_defined_name("LocalRange", Some("Sheet1"))? {
    println!("Sheet-scoped: {}", info.value);
}
```

**TypeScript:**

```typescript
const info = wb.getDefinedName("SalesTotal");
if (info) {
    console.log(`Refers to: ${info.value}`);
}

const local = wb.getDefinedName("LocalRange", "Sheet1");
```

### `get_all_defined_names` / `getDefinedNames`

워크북의 모든 정의된 이름을 반환한다.

**Rust:**

```rust
let names = wb.get_all_defined_names();
for dn in &names {
    println!("{}: {} (scope: {:?})", dn.name, dn.value, dn.scope);
}
```

**TypeScript:**

```typescript
const names = wb.getDefinedNames();
for (const dn of names) {
    console.log(`${dn.name}: ${dn.value} (scope: ${dn.scope ?? "workbook"})`);
}
```

### `delete_defined_name` / `deleteDefinedName`

이름과 선택적 범위로 정의된 이름을 삭제한다. 해당 이름이 없으면 오류를 반환한다.

**Rust:**

```rust
wb.delete_defined_name("SalesTotal", None)?;
wb.delete_defined_name("LocalRange", Some("Sheet1"))?;
```

**TypeScript:**

```typescript
wb.deleteDefinedName("SalesTotal");
wb.deleteDefinedName("LocalRange", "Sheet1");
```

### DefinedNameInfo

| 속성 | Rust 타입 | TypeScript 타입 | 설명 |
|------|-----------|-----------------|------|
| `name` | `String` | `string` | 정의된 이름 |
| `value` | `String` | `string` | 참조 또는 수식 |
| `scope` | `DefinedNameScope` | `string?` | 시트 이름(시트 범위) 또는 `None`/`undefined`(워크북 범위) |
| `comment` | `Option<String>` | `string?` | 선택적 주석 |

> 정의된 이름에는 `\ / ? * [ ]` 문자를 사용할 수 없으며, 앞뒤 공백도 허용되지 않는다.

---

## 19. 문서 속성

워크북의 메타데이터를 설정하고 조회하는 기능을 다룬다. 핵심 속성, 앱 속성, 사용자 정의 속성의 세 가지 유형이 있다.

### 핵심 속성 (Core Properties)

제목, 작성자 등 표준 문서 메타데이터를 다룬다.

**Rust:**

```rust
use sheetkit::doc_props::DocProperties;

wb.set_doc_props(DocProperties {
    title: Some("Annual Report".into()),
    subject: Some("Financial Data".into()),
    creator: Some("Finance Team".into()),
    keywords: Some("finance, annual, 2024".into()),
    description: Some("Annual financial report".into()),
    last_modified_by: Some("Admin".into()),
    revision: Some("3".into()),
    created: Some("2024-01-01T00:00:00Z".into()),
    modified: Some("2024-06-15T10:30:00Z".into()),
    category: Some("Reports".into()),
    content_status: Some("Final".into()),
});

let props = wb.get_doc_props();
```

**TypeScript:**

```typescript
wb.setDocProps({
    title: "Annual Report",
    subject: "Financial Data",
    creator: "Finance Team",
    keywords: "finance, annual, 2024",
    description: "Annual financial report",
    lastModifiedBy: "Admin",
    revision: "3",
    created: "2024-01-01T00:00:00Z",
    modified: "2024-06-15T10:30:00Z",
    category: "Reports",
    contentStatus: "Final",
});

const props = wb.getDocProps();
```

**DocProperties 속성:**

| 속성 | 타입 | 설명 |
|------|------|------|
| `title` | `string?` | 제목 |
| `subject` | `string?` | 주제 |
| `creator` | `string?` | 작성자 |
| `keywords` | `string?` | 키워드 |
| `description` | `string?` | 설명 |
| `last_modified_by` / `lastModifiedBy` | `string?` | 마지막 수정자 |
| `revision` | `string?` | 수정 번호 |
| `created` | `string?` | 생성 날짜 (ISO 8601) |
| `modified` | `string?` | 수정 날짜 (ISO 8601) |
| `category` | `string?` | 분류 |
| `content_status` / `contentStatus` | `string?` | 콘텐츠 상태 |

### 앱 속성 (App Properties)

애플리케이션 관련 메타데이터를 다룬다.

**Rust:**

```rust
use sheetkit::doc_props::AppProperties;

wb.set_app_props(AppProperties {
    application: Some("SheetKit".into()),
    doc_security: Some(0),
    company: Some("ACME Corp".into()),
    app_version: Some("1.0".into()),
    manager: Some("Department Lead".into()),
    template: None,
});

let app_props = wb.get_app_props();
```

**TypeScript:**

```typescript
wb.setAppProps({
    application: "SheetKit",
    docSecurity: 0,
    company: "ACME Corp",
    appVersion: "1.0",
    manager: "Department Lead",
});

const appProps = wb.getAppProps();
```

**AppProperties 속성:**

| 속성 | 타입 | 설명 |
|------|------|------|
| `application` | `string?` | 애플리케이션 이름 |
| `doc_security` / `docSecurity` | `u32?` / `number?` | 보안 수준 |
| `company` | `string?` | 회사 이름 |
| `app_version` / `appVersion` | `string?` | 앱 버전 |
| `manager` | `string?` | 관리자 |
| `template` | `string?` | 템플릿 |

### 사용자 정의 속성 (Custom Properties)

키-값 쌍으로 사용자 정의 메타데이터를 저장한다. 값은 문자열, 숫자, 불리언 타입을 지원한다.

**Rust:**

```rust
use sheetkit::doc_props::CustomPropertyValue;

wb.set_custom_property("Department", CustomPropertyValue::String("Engineering".into()));
wb.set_custom_property("Version", CustomPropertyValue::Int(3));
wb.set_custom_property("Approved", CustomPropertyValue::Bool(true));
wb.set_custom_property("Rating", CustomPropertyValue::Float(4.5));

let val = wb.get_custom_property("Department");
// Some(CustomPropertyValue::String("Engineering"))

let deleted = wb.delete_custom_property("Deprecated");
// true if existed
```

**TypeScript:**

```typescript
wb.setCustomProperty("Department", "Engineering");
wb.setCustomProperty("Version", 3);
wb.setCustomProperty("Approved", true);
wb.setCustomProperty("Rating", 4.5);

const val = wb.getCustomProperty("Department");
// "Engineering"

const deleted: boolean = wb.deleteCustomProperty("Deprecated");
```

> TypeScript에서 정수 값은 자동으로 Int로 변환되고, 소수점이 있는 숫자는 Float으로 저장된다.

---

## 20. 워크북 보호

워크북 구조(시트 추가/삭제/이름 변경)와 창 위치를 보호하는 기능을 다룬다. 선택적으로 비밀번호를 설정할 수 있다.

### `protect_workbook(config)` / `protectWorkbook(config)`

워크북 보호를 설정한다.

**Rust:**

```rust
use sheetkit::protection::WorkbookProtectionConfig;

wb.protect_workbook(WorkbookProtectionConfig {
    password: Some("secret123".into()),
    lock_structure: true,
    lock_windows: false,
    lock_revision: false,
});
```

**TypeScript:**

```typescript
wb.protectWorkbook({
    password: "secret123",
    lockStructure: true,
    lockWindows: false,
    lockRevision: false,
});
```

### `unprotect_workbook()` / `unprotectWorkbook()`

워크북 보호를 해제한다.

**Rust:**

```rust
wb.unprotect_workbook();
```

**TypeScript:**

```typescript
wb.unprotectWorkbook();
```

### `is_workbook_protected()` / `isWorkbookProtected()`

워크북이 보호되어 있는지 확인한다.

**Rust:**

```rust
let protected: bool = wb.is_workbook_protected();
```

**TypeScript:**

```typescript
const isProtected: boolean = wb.isWorkbookProtected();
```

### WorkbookProtectionConfig 속성

| 속성 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `password` | `string?` | `None` | 보호 비밀번호 (레거시 해시로 저장) |
| `lock_structure` / `lockStructure` | `bool` / `boolean?` | `false` | 시트 추가/삭제/이름 변경 차단 |
| `lock_windows` / `lockWindows` | `bool` / `boolean?` | `false` | 창 위치/크기 변경 차단 |
| `lock_revision` / `lockRevision` | `bool` / `boolean?` | `false` | 수정 추적 잠금 |

> 비밀번호는 Excel의 레거시 16비트 해시 알고리즘으로 저장된다. 이 해시는 암호학적으로 안전하지 않다.

---

## 21. 시트 보호

개별 시트의 편집을 제한하는 기능을 다룬다. 선택적으로 비밀번호를 설정하고 특정 작업을 허용할 수 있다.

### `protect_sheet` / `protectSheet`

시트를 보호한다. 모든 권한 불리언 값은 기본적으로 `false`(금지)이며, `true`로 설정하면 보호 상태에서도 해당 작업이 허용된다.

**Rust:**

```rust
use sheetkit::sheet::SheetProtectionConfig;

wb.protect_sheet("Sheet1", &SheetProtectionConfig {
    password: Some("mypass".to_string()),
    format_cells: true,
    insert_rows: true,
    sort: true,
    ..SheetProtectionConfig::default()
})?;
```

**TypeScript:**

```typescript
wb.protectSheet("Sheet1", {
    password: "mypass",
    formatCells: true,
    insertRows: true,
    sort: true,
});

// Protect with defaults (all actions forbidden, no password)
wb.protectSheet("Sheet1");
```

### `unprotect_sheet` / `unprotectSheet`

시트 보호를 해제한다.

**Rust:**

```rust
wb.unprotect_sheet("Sheet1")?;
```

**TypeScript:**

```typescript
wb.unprotectSheet("Sheet1");
```

### `is_sheet_protected` / `isSheetProtected`

시트가 보호되어 있는지 확인한다.

**Rust:**

```rust
let protected: bool = wb.is_sheet_protected("Sheet1")?;
```

**TypeScript:**

```typescript
const isProtected: boolean = wb.isSheetProtected("Sheet1");
```

### SheetProtectionConfig 속성

| 속성 | Rust 타입 | TypeScript 타입 | 기본값 | 설명 |
|------|-----------|-----------------|--------|------|
| `password` | `Option<String>` | `string?` | `None` | 보호 비밀번호 (레거시 해시로 저장) |
| `select_locked_cells` / `selectLockedCells` | `bool` | `boolean?` | `false` | 잠긴 셀 선택 허용 |
| `select_unlocked_cells` / `selectUnlockedCells` | `bool` | `boolean?` | `false` | 잠기지 않은 셀 선택 허용 |
| `format_cells` / `formatCells` | `bool` | `boolean?` | `false` | 셀 서식 변경 허용 |
| `format_columns` / `formatColumns` | `bool` | `boolean?` | `false` | 열 서식 변경 허용 |
| `format_rows` / `formatRows` | `bool` | `boolean?` | `false` | 행 서식 변경 허용 |
| `insert_columns` / `insertColumns` | `bool` | `boolean?` | `false` | 열 삽입 허용 |
| `insert_rows` / `insertRows` | `bool` | `boolean?` | `false` | 행 삽입 허용 |
| `insert_hyperlinks` / `insertHyperlinks` | `bool` | `boolean?` | `false` | 하이퍼링크 삽입 허용 |
| `delete_columns` / `deleteColumns` | `bool` | `boolean?` | `false` | 열 삭제 허용 |
| `delete_rows` / `deleteRows` | `bool` | `boolean?` | `false` | 행 삭제 허용 |
| `sort` | `bool` | `boolean?` | `false` | 정렬 허용 |
| `auto_filter` / `autoFilter` | `bool` | `boolean?` | `false` | 자동 필터 사용 허용 |
| `pivot_tables` / `pivotTables` | `bool` | `boolean?` | `false` | 피벗 테이블 사용 허용 |

> 비밀번호는 Excel의 레거시 16비트 해시 알고리즘으로 저장된다. 이 해시는 암호학적으로 안전하지 않다.

---

## 22. 수식 평가

셀 수식을 파싱하고 실행하는 기능을 다룬다. nom 파서로 수식을 AST로 변환한 후 평가 엔진이 결과를 계산한다.

### `evaluate_formula(sheet, formula)` / `evaluateFormula(sheet, formula)`

주어진 시트 컨텍스트에서 수식 문자열을 평가하여 결과를 반환한다. 워크북의 현재 셀 데이터를 참조할 수 있다.

**Rust:**

```rust
wb.set_cell_value("Sheet1", "A1", CellValue::Number(10.0))?;
wb.set_cell_value("Sheet1", "A2", CellValue::Number(20.0))?;

let result = wb.evaluate_formula("Sheet1", "SUM(A1:A2)")?;
// CellValue::Number(30.0)
```

**TypeScript:**

```typescript
wb.setCellValue("Sheet1", "A1", 10);
wb.setCellValue("Sheet1", "A2", 20);

const result = wb.evaluateFormula("Sheet1", "SUM(A1:A2)");
// 30
```

### `calculate_all()` / `calculateAll()`

워크북의 모든 수식 셀을 재계산한다. 의존성 그래프를 구축하고 위상 정렬을 수행하여 올바른 순서로 평가한다.

**Rust:**

```rust
wb.calculate_all()?;
```

**TypeScript:**

```typescript
wb.calculateAll();
```

> 순환 참조가 발견되면 오류가 발생한다. 최대 재귀 깊이는 256이다.

### 지원 함수 목록 (110개, 8개 카테고리)

#### 수학 함수 (Math) -- 20개

| 함수 | 설명 |
|------|------|
| `SUM` | 합계 |
| `PRODUCT` | 곱 |
| `ABS` | 절대값 |
| `INT` | 정수 변환 (내림) |
| `MOD` | 나머지 |
| `POWER` | 거듭제곱 |
| `SQRT` | 제곱근 |
| `ROUND` | 반올림 |
| `ROUNDUP` | 올림 |
| `ROUNDDOWN` | 내림 |
| `CEILING` | 올림 (배수) |
| `FLOOR` | 내림 (배수) |
| `SIGN` | 부호 |
| `RAND` | 난수 (0-1) |
| `RANDBETWEEN` | 정수 난수 (범위) |
| `PI` | 원주율 |
| `LOG` | 로그 |
| `LOG10` | 상용 로그 |
| `LN` | 자연 로그 |
| `EXP` | 지수 함수 |
| `QUOTIENT` | 정수 몫 |
| `FACT` | 팩토리얼 |
| `SUMIF` | 조건부 합계 |
| `SUMIFS` | 다중 조건부 합계 |

#### 통계 함수 (Statistical) -- 16개

| 함수 | 설명 |
|------|------|
| `AVERAGE` | 평균 |
| `COUNT` | 숫자 셀 개수 |
| `COUNTA` | 비어 있지 않은 셀 개수 |
| `COUNTBLANK` | 빈 셀 개수 |
| `COUNTIF` | 조건부 개수 |
| `COUNTIFS` | 다중 조건부 개수 |
| `MIN` | 최소값 |
| `MAX` | 최대값 |
| `MEDIAN` | 중앙값 |
| `MODE` | 최빈값 |
| `LARGE` | N번째 큰 값 |
| `SMALL` | N번째 작은 값 |
| `RANK` | 순위 |
| `AVERAGEIF` | 조건부 평균 |
| `AVERAGEIFS` | 다중 조건부 평균 |

#### 논리 함수 (Logical) -- 10개

| 함수 | 설명 |
|------|------|
| `IF` | 조건 분기 |
| `AND` | 논리곱 |
| `OR` | 논리합 |
| `NOT` | 부정 |
| `XOR` | 배타적 논리합 |
| `TRUE` | TRUE 상수 |
| `FALSE` | FALSE 상수 |
| `IFERROR` | 오류 시 대체값 |
| `IFNA` | #N/A 시 대체값 |
| `IFS` | 다중 조건 분기 |
| `SWITCH` | 값 기반 분기 |

#### 텍스트 함수 (Text) -- 15개

| 함수 | 설명 |
|------|------|
| `LEN` | 문자열 길이 |
| `LOWER` | 소문자 변환 |
| `UPPER` | 대문자 변환 |
| `TRIM` | 공백 제거 |
| `LEFT` | 왼쪽 문자 추출 |
| `RIGHT` | 오른쪽 문자 추출 |
| `MID` | 중간 문자 추출 |
| `CONCATENATE` | 문자열 연결 |
| `CONCAT` | 문자열 연결 (최신) |
| `FIND` | 문자열 찾기 (대소문자 구분) |
| `SEARCH` | 문자열 찾기 (대소문자 무시) |
| `SUBSTITUTE` | 문자열 치환 |
| `REPLACE` | 위치 기반 문자열 교체 |
| `REPT` | 문자열 반복 |
| `EXACT` | 완전 일치 비교 |
| `T` | 텍스트 변환 |
| `PROPER` | 단어 첫 글자 대문자 |

#### 정보 함수 (Information) -- 11개

| 함수 | 설명 |
|------|------|
| `ISNUMBER` | 숫자 여부 |
| `ISTEXT` | 텍스트 여부 |
| `ISBLANK` | 빈 셀 여부 |
| `ISERROR` | 오류 여부 (#N/A 포함) |
| `ISERR` | 오류 여부 (#N/A 제외) |
| `ISNA` | #N/A 여부 |
| `ISLOGICAL` | 논리값 여부 |
| `ISEVEN` | 짝수 여부 |
| `ISODD` | 홀수 여부 |
| `TYPE` | 값 유형 번호 |
| `N` | 숫자 변환 |
| `NA` | #N/A 생성 |
| `ERROR.TYPE` | 오류 유형 번호 |

#### 변환 함수 (Conversion) -- 2개

| 함수 | 설명 |
|------|------|
| `VALUE` | 텍스트를 숫자로 변환 |
| `TEXT` | 값을 서식 문자열로 변환 |

#### 날짜/시각 함수 (Date/Time) -- 17개

| 함수 | 설명 |
|------|------|
| `DATE` | 연/월/일로 날짜 생성 |
| `TODAY` | 오늘 날짜 |
| `NOW` | 현재 날짜 및 시각 |
| `YEAR` | 연도 추출 |
| `MONTH` | 월 추출 |
| `DAY` | 일 추출 |
| `HOUR` | 시 추출 |
| `MINUTE` | 분 추출 |
| `SECOND` | 초 추출 |
| `DATEDIF` | 날짜 차이 계산 |
| `EDATE` | N개월 후 날짜 |
| `EOMONTH` | N개월 후 월말 |
| `DATEVALUE` | 텍스트를 날짜로 변환 |
| `WEEKDAY` | 요일 번호 |
| `WEEKNUM` | 주차 번호 |
| `NETWORKDAYS` | 근무일수 계산 |
| `WORKDAY` | N 근무일 후 날짜 |

#### 찾기/참조 함수 (Lookup) -- 11개

| 함수 | 설명 |
|------|------|
| `VLOOKUP` | 세로 방향 조회 |
| `HLOOKUP` | 가로 방향 조회 |
| `INDEX` | 범위에서 값 추출 |
| `MATCH` | 위치 찾기 |
| `LOOKUP` | 벡터 조회 |
| `ROW` | 행 번호 |
| `COLUMN` | 열 번호 |
| `ROWS` | 범위의 행 수 |
| `COLUMNS` | 범위의 열 수 |
| `CHOOSE` | 인덱스로 값 선택 |
| `ADDRESS` | 셀 주소 문자열 생성 |

---

## 23. 피벗 테이블

피벗 테이블을 생성, 조회, 삭제하는 기능을 다룬다. 소스 데이터 범위로부터 행/열/데이터 필드를 지정하여 피벗 테이블을 구성한다.

### `add_pivot_table(config)` / `addPivotTable(config)`

피벗 테이블을 추가한다.

**Rust:**

```rust
use sheetkit::pivot::*;

wb.add_pivot_table(&PivotTableConfig {
    name: "SalesPivot".into(),
    source_sheet: "RawData".into(),
    source_range: "A1:D100".into(),
    target_sheet: "PivotSheet".into(),
    target_cell: "A3".into(),
    rows: vec![
        PivotField { name: "Region".into() },
        PivotField { name: "Product".into() },
    ],
    columns: vec![
        PivotField { name: "Quarter".into() },
    ],
    data: vec![
        PivotDataField {
            name: "Revenue".into(),
            function: AggregateFunction::Sum,
            display_name: Some("Total Revenue".into()),
        },
        PivotDataField {
            name: "Quantity".into(),
            function: AggregateFunction::Count,
            display_name: None,
        },
    ],
})?;
```

**TypeScript:**

```typescript
wb.addPivotTable({
    name: "SalesPivot",
    sourceSheet: "RawData",
    sourceRange: "A1:D100",
    targetSheet: "PivotSheet",
    targetCell: "A3",
    rows: [
        { name: "Region" },
        { name: "Product" },
    ],
    columns: [
        { name: "Quarter" },
    ],
    data: [
        { name: "Revenue", function: "sum", displayName: "Total Revenue" },
        { name: "Quantity", function: "count" },
    ],
});
```

> Node.js에서 `data[].function`은 지원되는 집계 함수(`sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevP`, `var`, `varP`)만 허용되며, 지원되지 않는 값은 오류를 반환한다.

### `get_pivot_tables()` / `getPivotTables()`

워크북의 모든 피벗 테이블 정보를 반환한다.

**Rust:**

```rust
let tables = wb.get_pivot_tables();
for t in &tables {
    println!("{}: {}!{} -> {}!{}", t.name, t.source_sheet, t.source_range, t.target_sheet, t.location);
}
```

**TypeScript:**

```typescript
const tables = wb.getPivotTables();
for (const t of tables) {
    console.log(`${t.name}: ${t.sourceSheet}!${t.sourceRange} -> ${t.targetSheet}!${t.location}`);
}
```

**PivotTableInfo 구조:**

| 속성 | 타입 | 설명 |
|------|------|------|
| `name` | `string` | 피벗 테이블 이름 |
| `source_sheet` / `sourceSheet` | `string` | 소스 데이터 시트 이름 |
| `source_range` / `sourceRange` | `string` | 소스 데이터 범위 |
| `target_sheet` / `targetSheet` | `string` | 대상 시트 이름 |
| `location` | `string` | 피벗 테이블 위치 |

### `delete_pivot_table(name)` / `deletePivotTable(name)`

이름으로 피벗 테이블을 삭제한다.

**Rust:**

```rust
wb.delete_pivot_table("SalesPivot")?;
```

**TypeScript:**

```typescript
wb.deletePivotTable("SalesPivot");
```

### PivotTableConfig 구조

| 속성 | 타입 | 설명 |
|------|------|------|
| `name` | `string` | 피벗 테이블 이름 |
| `source_sheet` / `sourceSheet` | `string` | 소스 데이터가 있는 시트 |
| `source_range` / `sourceRange` | `string` | 소스 데이터 범위 (예: "A1:D100") |
| `target_sheet` / `targetSheet` | `string` | 피벗 테이블을 배치할 시트 |
| `target_cell` / `targetCell` | `string` | 피벗 테이블 시작 셀 |
| `rows` | `PivotField[]` | 행 필드 |
| `columns` | `PivotField[]` | 열 필드 |
| `data` | `PivotDataField[]` | 데이터(값) 필드 |

### AggregateFunction (집계 함수)

| 값 | 설명 |
|----|------|
| `sum` | 합계 |
| `count` | 개수 |
| `average` | 평균 |
| `max` | 최대값 |
| `min` | 최소값 |
| `product` | 곱 |
| `countNums` | 숫자 개수 |
| `stdDev` | 표준편차 |
| `stdDevP` | 모표준편차 |
| `var` | 분산 |
| `varP` | 모분산 |

---

## 24. 스트림 라이터

대용량 데이터를 메모리 효율적으로 쓰기 위한 스트리밍 API이다. 행은 오름차순으로만 쓸 수 있으며, 전체 워크시트 XML을 메모리에 구축하지 않고 직접 버퍼에 기록한다.

### 사용 흐름

1. `new_stream_writer`로 스트림 라이터 생성
2. 열 너비, 셀 병합 등 설정
3. `write_row`로 행 데이터를 순서대로 기록
4. `apply_stream_writer`로 워크북에 적용

**Rust:**

```rust
use sheetkit::cell::CellValue;

let mut sw = wb.new_stream_writer("LargeData")?;

// Set column widths
sw.set_col_width(1, 15.0)?;   // Column A
sw.set_col_width(2, 20.0)?;   // Column B

// Add merge cells
sw.add_merge_cell("A1:B1")?;

// Write header row
sw.write_row(1, &[
    CellValue::from("Name"),
    CellValue::from("Value"),
])?;

// Write data rows (must be in ascending order)
for i in 2..=10000 {
    sw.write_row(i, &[
        CellValue::from(format!("Item {}", i - 1)),
        CellValue::Number(i as f64 * 1.5),
    ])?;
}

// Apply to workbook
let sheet_index = wb.apply_stream_writer(sw)?;
wb.save("large_data.xlsx")?;
```

**TypeScript:**

```typescript
const sw = wb.newStreamWriter("LargeData");

// Set column widths
sw.setColWidth(1, 15);   // Column A
sw.setColWidth(2, 20);   // Column B

// Set width for a range of columns
sw.setColWidthRange(3, 10, 12);  // Columns C-J

// Add merge cells
sw.addMergeCell("A1:B1");

// Write header row
sw.writeRow(1, ["Name", "Value"]);

// Write data rows
for (let i = 2; i <= 10000; i++) {
    sw.writeRow(i, [`Item ${i - 1}`, i * 1.5]);
}

// Apply to workbook
const sheetIndex: number = wb.applyStreamWriter(sw);
await wb.save("large_data.xlsx");
```

### StreamWriter API

#### `new_stream_writer(sheet_name)` / `newStreamWriter(sheetName)`

새 시트를 위한 스트림 라이터를 생성한다.

#### `write_row(row, values)` / `writeRow(row, values)`

행 데이터를 기록한다. 행 번호는 1부터 시작하며 반드시 오름차순이어야 한다.

#### `set_col_width(col, width)` / `setColWidth(col, width)`

열 너비를 설정한다. `col`은 1부터 시작하는 열 번호이다.

#### `set_col_width_range(min_col, max_col, width)` / `setColWidthRange(minCol, maxCol, width)`

열 범위의 너비를 한 번에 설정한다.

#### `add_merge_cell(reference)` / `addMergeCell(reference)`

셀 병합을 추가한다 (예: "A1:C3").

#### `apply_stream_writer(writer)` / `applyStreamWriter(writer)`

스트림 라이터의 결과를 워크북에 적용한다. 시트 인덱스를 반환한다. 적용 후 스트림 라이터는 소비(consumed)되어 더 이상 사용할 수 없다.

### StreamRowOptions (Rust 전용)

Rust에서는 `write_row_with_options`를 사용하여 행별 옵션을 지정할 수 있다.

| 속성 | 타입 | 설명 |
|------|------|------|
| `height` | `Option<f64>` | 행 높이 (포인트) |
| `visible` | `Option<bool>` | 행 표시 여부 |
| `outline_level` | `Option<u8>` | 아웃라인 수준 (0-7) |
| `style_id` | `Option<u32>` | 행 스타일 ID |

---

## 25. 유틸리티 함수

셀 참조 변환에 사용되는 유틸리티 함수들이다. Rust에서는 `sheetkit_core::utils::cell_ref` 모듈에서 제공한다.

### `cell_name_to_coordinates`

A1 형식의 셀 참조를 (열, 행) 좌표로 변환한다. 열과 행 모두 1부터 시작한다.

**Rust:**

```rust
use sheetkit_core::utils::cell_ref::cell_name_to_coordinates;

let (col, row) = cell_name_to_coordinates("B3")?;
// col = 2, row = 3

let (col, row) = cell_name_to_coordinates("$AB$100")?;
// col = 28, row = 100 (absolute references are supported)
```

### `coordinates_to_cell_name`

(열, 행) 좌표를 A1 형식의 셀 참조로 변환한다.

**Rust:**

```rust
use sheetkit_core::utils::cell_ref::coordinates_to_cell_name;

let name = coordinates_to_cell_name(2, 3)?;
// "B3"

let name = coordinates_to_cell_name(28, 100)?;
// "AB100"
```

### `column_name_to_number`

열 이름을 1부터 시작하는 열 번호로 변환한다.

**Rust:**

```rust
use sheetkit_core::utils::cell_ref::column_name_to_number;

assert_eq!(column_name_to_number("A")?, 1);
assert_eq!(column_name_to_number("Z")?, 26);
assert_eq!(column_name_to_number("AA")?, 27);
assert_eq!(column_name_to_number("XFD")?, 16384);  // maximum column
```

### `column_number_to_name`

1부터 시작하는 열 번호를 열 이름으로 변환한다.

**Rust:**

```rust
use sheetkit_core::utils::cell_ref::column_number_to_name;

assert_eq!(column_number_to_name(1)?, "A");
assert_eq!(column_number_to_name(26)?, "Z");
assert_eq!(column_number_to_name(27)?, "AA");
assert_eq!(column_number_to_name(16384)?, "XFD");
```

> 유틸리티 함수는 현재 Rust 전용으로 제공된다. TypeScript에서는 문자열 기반 셀 참조("A1", "B2" 등)를 직접 사용한다.

### `is_date_num_fmt(num_fmt_id)` (Rust 전용)

내장 숫자 서식 ID가 날짜/시간 서식인지 확인한다. ID 14-22 및 45-47에 대해 `true`를 반환한다.

```rust
use sheetkit::is_date_num_fmt;

assert!(is_date_num_fmt(14));   // m/d/yyyy
assert!(is_date_num_fmt(22));   // m/d/yyyy h:mm
assert!(!is_date_num_fmt(0));   // General
assert!(!is_date_num_fmt(49));  // @
```

### `is_date_format_code(code)` (Rust 전용)

사용자 정의 숫자 서식 문자열이 날짜/시간 서식인지 확인한다. 따옴표로 감싸진 문자열과 이스케이프된 문자를 제외하고, 서식 코드에 날짜/시간 토큰(y, m, d, h, s)이 포함되어 있으면 `true`를 반환한다.

```rust
use sheetkit::is_date_format_code;

assert!(is_date_format_code("yyyy-mm-dd"));
assert!(is_date_format_code("h:mm:ss AM/PM"));
assert!(!is_date_format_code("#,##0.00"));
assert!(!is_date_format_code("0%"));
```

---

## 부록: 제한 사항

| 항목 | 제한 |
|------|------|
| 최대 열 수 | 16,384 (XFD) |
| 최대 행 수 | 1,048,576 |
| 최대 셀 문자 수 | 32,767 |
| 최대 행 높이 | 409 포인트 |
| 최대 아웃라인 수준 | 7 |
| 최대 스타일 XF 수 | 65,430 |
| 수식 최대 재귀 깊이 | 256 |
| 지원 수식 함수 수 | 110 / 456 |

---

## 26. 스파크라인

스파크라인은 워크시트 셀에 삽입되는 미니 차트이다. SheetKit은 Line, Column, Win/Loss 세 가지 스파크라인 유형을 지원한다. Excel은 36가지 스타일 프리셋(인덱스 0-35)을 정의한다.

> **참고:** 스파크라인 타입, 설정, XML 변환이 구현되어 있다. 워크북 통합(`addSparkline` / `getSparklines`)은 향후 릴리스에서 추가될 예정이다.

### 타입

#### `SparklineType` (Rust) / `sparklineType` (TypeScript)

| 값 | Rust | TypeScript | OOXML |
|------|------|------------|-------|
| Line | `SparklineType::Line` | `"line"` | (default, omitted) |
| Column | `SparklineType::Column` | `"column"` | `"column"` |
| Win/Loss | `SparklineType::WinLoss` | `"winloss"` or `"stacked"` | `"stacked"` |

#### `SparklineConfig` (Rust)

```rust
use sheetkit::SparklineConfig;

let config = SparklineConfig::new("Sheet1!A1:A10", "B1");
```

필드:

| 필드 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `data_range` | `String` | (필수) | 데이터 소스 범위 (예: `"Sheet1!A1:A10"`) |
| `location` | `String` | (필수) | 스파크라인이 렌더링되는 셀 (예: `"B1"`) |
| `sparkline_type` | `SparklineType` | `Line` | 스파크라인 차트 유형 |
| `markers` | `bool` | `false` | 데이터 마커 표시 |
| `high_point` | `bool` | `false` | 최고점 강조 |
| `low_point` | `bool` | `false` | 최저점 강조 |
| `first_point` | `bool` | `false` | 첫 번째 점 강조 |
| `last_point` | `bool` | `false` | 마지막 점 강조 |
| `negative_points` | `bool` | `false` | 음수 값 강조 |
| `show_axis` | `bool` | `false` | 가로축 표시 |
| `line_weight` | `Option<f64>` | `None` | 선 두께 (포인트) |
| `style` | `Option<u32>` | `None` | 스타일 프리셋 인덱스 (0-35) |

#### `JsSparklineConfig` (TypeScript)

```typescript
const config = {
  dataRange: 'Sheet1!A1:A10',
  location: 'B1',
  sparklineType: 'line',    // "line" | "column" | "winloss" | "stacked"
  markers: true,
  highPoint: false,
  lowPoint: false,
  firstPoint: false,
  lastPoint: false,
  negativePoints: false,
  showAxis: false,
  lineWeight: 0.75,
  style: 1,
};
```

### 유효성 검사

`validate_sparkline_config` 함수(Rust)는 다음을 확인한다:
- `data_range`가 비어 있지 않음
- `location`이 비어 있지 않음
- `line_weight`(설정된 경우) 양수 여부
- `style`(설정된 경우) 0-35 범위 내 여부

```rust
use sheetkit_core::sparkline::{SparklineConfig, validate_sparkline_config};

let config = SparklineConfig::new("Sheet1!A1:A10", "B1");
validate_sparkline_config(&config).unwrap(); // Ok
```

## 27. Theme Colors

Theme color slot (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink) resolve with optional tint.

### Workbook.getThemeColor (Node.js) / Workbook::get_theme_color (Rust)

| Parameter | Type              | Description                                      |
| --------- | ----------------- | ------------------------------------------------ |
| index     | `u32` / `number`  | Theme color index (0-11)                         |
| tint      | `Option<f64>` / `number \| null` | Tint value: positive lightens, negative darkens |

**Returns:** ARGB hex string (e.g. `"FF4472C4"`) or `None`/`null` if out of range.

**Theme Color Indices:**

| Index | Slot Name | Default Color |
| ----- | --------- | ------------- |
| 0     | dk1       | FF000000      |
| 1     | lt1       | FFFFFFFF      |
| 2     | dk2       | FF44546A      |
| 3     | lt2       | FFE7E6E6      |
| 4     | accent1   | FF4472C4      |
| 5     | accent2   | FFED7D31      |
| 6     | accent3   | FFA5A5A5      |
| 7     | accent4   | FFFFC000      |
| 8     | accent5   | FF5B9BD5      |
| 9     | accent6   | FF70AD47      |
| 10    | hlink     | FF0563C1      |
| 11    | folHlink  | FF954F72      |

#### Node.js

```javascript
const wb = new Workbook();

// Get accent1 color (no tint)
const color = wb.getThemeColor(4, null); // "FF4472C4"

// Lighten black by 50%
const lightened = wb.getThemeColor(0, 0.5); // "FF7F7F7F"

// Darken white by 50%
const darkened = wb.getThemeColor(1, -0.5); // "FF7F7F7F"

// Out of range returns null
const invalid = wb.getThemeColor(99, null); // null
```

#### Rust

```rust
let wb = Workbook::new();

// Get accent1 color (no tint)
let color = wb.get_theme_color(4, None); // Some("FF4472C4")

// Apply tint
let tinted = wb.get_theme_color(0, Some(0.5)); // Some("FF7F7F7F")
```

### Gradient Fill

`FillStyle` type supports gradient fills via the `gradient` field.

#### Types

```rust
pub struct GradientFillStyle {
    pub gradient_type: GradientType, // Linear or Path
    pub degree: Option<f64>,         // Rotation angle for linear gradients
    pub left: Option<f64>,           // Path gradient coordinates (0.0-1.0)
    pub right: Option<f64>,
    pub top: Option<f64>,
    pub bottom: Option<f64>,
    pub stops: Vec<GradientStop>,    // Color stops
}

pub struct GradientStop {
    pub position: f64,     // Position (0.0-1.0)
    pub color: StyleColor, // Color at this stop
}

pub enum GradientType {
    Linear,
    Path,
}
```

#### Rust Example

```rust
use sheetkit::*;

let mut wb = Workbook::new();
let style_id = wb.add_style(&Style {
    fill: Some(FillStyle {
        pattern: PatternType::None,
        fg_color: None,
        bg_color: None,
        gradient: Some(GradientFillStyle {
            gradient_type: GradientType::Linear,
            degree: Some(90.0),
            left: None,
            right: None,
            top: None,
            bottom: None,
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: StyleColor::Rgb("FFFFFFFF".to_string()),
                },
                GradientStop {
                    position: 1.0,
                    color: StyleColor::Rgb("FF4472C4".to_string()),
                },
            ],
        }),
    }),
    ..Style::default()
})?;
```

---

## 28. 서식 있는 텍스트

서식 있는 텍스트(Rich Text)를 사용하면 하나의 셀에 글꼴, 크기, 굵게, 기울임, 색상 등 서로 다른 서식을 가진 여러 텍스트 조각(run)을 넣을 수 있다.

### `RichTextRun` 타입

각 run은 `RichTextRun`으로 기술된다.

**Rust:**

```rust
pub struct RichTextRun {
    pub text: String,
    pub font: Option<String>,
    pub size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub color: Option<String>,
}
```

**TypeScript:**

```typescript
interface RichTextRun {
  text: string;
  font?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;  // RGB hex string, e.g. "#FF0000"
}
```

### `set_cell_rich_text` / `setCellRichText`

셀에 여러 서식 run으로 구성된 서식 있는 텍스트를 설정한다.

**Rust:**

```rust
use sheetkit::{Workbook, RichTextRun};

let mut wb = Workbook::new();
let runs = vec![
    RichTextRun {
        text: "Bold text".to_string(),
        font: Some("Arial".to_string()),
        size: Some(14.0),
        bold: true,
        italic: false,
        color: Some("#FF0000".to_string()),
    },
    RichTextRun {
        text: " normal text".to_string(),
        font: None,
        size: None,
        bold: false,
        italic: false,
        color: None,
    },
];
wb.set_cell_rich_text("Sheet1", "A1", runs)?;
```

**TypeScript:**

```typescript
const wb = new Workbook();
wb.setCellRichText("Sheet1", "A1", [
  { text: "Bold text", font: "Arial", size: 14, bold: true, color: "#FF0000" },
  { text: " normal text" },
]);
```

### `get_cell_rich_text` / `getCellRichText`

셀의 서식 있는 텍스트 run을 가져온다. 서식 있는 텍스트가 아닌 셀은 `None`/`null`을 반환한다.

**Rust:**

```rust
let runs = wb.get_cell_rich_text("Sheet1", "A1")?;
if let Some(runs) = runs {
    for run in &runs {
        println!("Text: {:?}, Bold: {}", run.text, run.bold);
    }
}
```

**TypeScript:**

```typescript
const runs = wb.getCellRichText("Sheet1", "A1");
if (runs) {
  for (const run of runs) {
    console.log(`Text: ${run.text}, Bold: ${run.bold ?? false}`);
  }
}
```

### `CellValue::RichString` (Rust 전용)

서식 있는 텍스트 셀은 `CellValue::RichString(Vec<RichTextRun>)` variant를 사용한다. `get_cell_value`로 읽으면 모든 run의 텍스트가 연결된 문자열로 표시된다.

```rust
match wb.get_cell_value("Sheet1", "A1")? {
    CellValue::RichString(runs) => {
        println!("Rich text with {} runs", runs.len());
    }
    _ => {}
}
```

### `rich_text_to_plain`

서식 있는 텍스트 run 슬라이스에서 연결된 일반 텍스트를 추출하는 유틸리티 함수이다.

**Rust:**

```rust
use sheetkit::rich_text_to_plain;

let plain = rich_text_to_plain(&runs);
```
