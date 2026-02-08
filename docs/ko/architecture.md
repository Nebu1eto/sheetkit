# SheetKit 아키텍처

## 1. 개요

SheetKit은 Excel (.xlsx) 파일을 읽고 쓰기 위한 Go Excelize 라이브러리의 Rust 재작성 버전이다. .xlsx 형식은 XML 파트를 포함하는 ZIP 아카이브인 OOXML (Office Open XML)이다. SheetKit은 ZIP을 읽고, 각 XML 파트를 타입이 지정된 Rust 구조체로 역직렬화하며, 조작을 위한 고수준 API를 제공하고, 저장 시 모든 것을 유효한 .xlsx 파일로 직렬화한다.

## 2. 크레이트 구조

```
crates/
  sheetkit-xml/     # XML 스키마 타입 (serde 기반)
  sheetkit-core/    # 비즈니스 로직
  sheetkit/         # 공개 파사드 (재내보내기)
packages/
  sheetkit/         # Node.js 바인딩 (napi-rs)
```

### sheetkit-xml

OOXML 스키마에 매핑되는 저수준 XML 데이터 구조체. 각 파일은 주요 OOXML 파트에 대응된다:

| 파일 | OOXML 파트 |
|---|---|
| `worksheet.rs` | Worksheet (시트 데이터, 셀 병합, 조건부 서식, 유효성 검사) |
| `shared_strings.rs` | SharedStrings (SST) |
| `styles.rs` | Stylesheet (글꼴, 채우기, 테두리, 숫자 형식, XF 레코드, DXF 레코드) |
| `workbook.rs` | Workbook (시트, 정의된 이름, 계산 속성, 피벗 캐시) |
| `content_types.rs` | `[Content_Types].xml` |
| `relationships.rs` | `.rels` 관계 파일 |
| `chart.rs` | 차트 정의 (DrawingML 차트) |
| `drawing.rs` | DrawingML (앵커, 도형, 이미지 참조) |
| `comments.rs` | 코멘트 데이터 및 작성자 |
| `doc_props.rs` | Core, App, Custom 문서 속성 |
| `pivot_table.rs` | 피벗 테이블 정의 |
| `pivot_cache.rs` | 피벗 캐시 정의 및 레코드 |
| `namespaces.rs` | OOXML 네임스페이스 상수 |

모든 타입은 XML 요소/속성 매핑을 위해 `quick-xml` 속성과 함께 `serde::Deserialize` 및 `serde::Serialize` 파생 매크로를 사용한다.

### sheetkit-core

모든 비즈니스 로직이 이 크레이트에 위치한다. 중심 타입은 `workbook.rs`의 `Workbook`으로, 역직렬화된 XML 상태를 소유하고 공개 API를 제공한다.

**핵심 모듈:**

| 모듈 | 책임 |
|---|---|
| `workbook.rs` | ZIP 열기, XML 파트 역직렬화, 가변 상태 관리, 직렬화 및 저장 |
| `cell.rs` | `CellValue` 열거형 (String, Number, Bool, Empty, Date, Formula, Error), chrono를 통한 날짜 시리얼 넘버 변환 |
| `sst.rs` | O(1) 문자열 중복 제거를 위한 HashMap 기반 공유 문자열 테이블 런타임 |
| `sheet.rs` | 시트 관리: 생성, 삭제, 이름 변경, 복사, 활성 시트 설정, 틀 고정/분할, 시트 속성, 시트 보호 |
| `row.rs` | 행 작업: 삽입, 삭제, 복제, 높이 설정, 가시성, 아웃라인 수준, 행 스타일, 반복자 |
| `col.rs` | 열 작업: 너비 설정, 가시성, 삽입, 삭제, 아웃라인 수준, 열 스타일 |
| `style.rs` | 스타일 시스템: 글꼴, 채우기, 테두리, 정렬, 숫자 형식, 셀 보호. 자동 XF 중복 제거를 포함한 StyleBuilder API |
| `conditional.rs` | 조건부 서식: DXF 레코드를 사용하는 18가지 규칙 유형 (셀 값, 색상 스케일, 데이터 막대, 아이콘 집합, 상위/하위 등) |
| `chart.rs` | 41가지 차트 유형 생성 (bar, line, pie, area, scatter, radar, stock, surface, doughnut, combo, 3D 변형). DrawingML 앵커 및 관계 관리 |
| `image.rs` | 이미지 삽입 (PNG, JPEG, GIF), DrawingML의 두 셀 앵커 사용 |
| `validation.rs` | 데이터 유효성 검사 규칙 (드롭다운, 정수, 소수, 텍스트 길이, 날짜, 시간, 사용자 정의 수식) |
| `comment.rs` | VML 드로잉 형식을 사용하는 셀 코멘트 |
| `table.rs` | 테이블 및 자동 필터 지원 |
| `hyperlink.rs` | 하이퍼링크: 외부 및 이메일은 워크시트 .rels 사용, 내부 (시트 간)는 location 속성만 사용 |
| `merge.rs` | 셀 범위 병합 및 해제 |
| `doc_props.rs` | Core (dc:, dcterms:, cp:), App, Custom 문서 속성. DC 네임스페이스는 수동 quick-xml Writer/Reader 필요 |
| `protection.rs` | 레거시 비밀번호 해시를 사용한 워크북 수준 보호 |
| `page_layout.rs` | 페이지 여백, 페이지 설정, 인쇄 옵션, 머리글/바닥글 |
| `defined_names.rs` | 이름이 지정된 범위 (워크북 범위 및 시트 범위) |
| `pivot.rs` | 피벗 테이블: 캐시 정의, 캐시 레코드, 테이블 정의, 워크북 피벗 캐시 컬렉션 |
| `stream.rs` | StreamWriter: 전체 시트를 메모리에 보관하지 않고 대용량 파일을 생성하기 위한 전방 전용 XML 작성기. SST 병합, 틀 고정, 행 옵션, 열 스타일/가시성/아웃라인 지원 |
| `error.rs` | thiserror를 사용한 에러 타입 |

**수식 서브시스템** (`formula/`):

| 파일 | 책임 |
|---|---|
| `parser.rs` | AST를 생성하는 nom 기반 수식 파서. 연산자 우선순위, 셀 참조, 범위 참조, 함수 호출, 문자열/숫자/불리언 리터럴 처리 |
| `ast.rs` | AST 노드 타입 (BinaryOp, UnaryOp, FunctionCall, CellRef, RangeRef, Literal 등) |
| `eval.rs` | 수식 평가기. 워크북 데이터 접근을 위한 `CellDataProvider` 트레이트 사용, 한 셀의 수식이 다른 셀을 참조할 때 가변/불변 대여 충돌을 피하기 위한 `CellSnapshot` (HashMap) 사용. `calculate_all()`은 의존성 그래프를 구축하고 Kahn 알고리즘으로 위상 정렬 수행 |
| `functions/mod.rs` | 함수 이름을 구현에 매핑하는 함수 디스패치 테이블 |
| `functions/math.rs` | 수학 함수 (SUM, AVERAGE, ABS, ROUND 등) |
| `functions/statistical.rs` | 통계 함수 (COUNT, COUNTA, MAX, MIN, STDEV 등) |
| `functions/text.rs` | 텍스트 함수 (CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM 등) |
| `functions/logical.rs` | 논리 함수 (IF, AND, OR, NOT, IFERROR 등) |
| `functions/information.rs` | 정보 함수 (ISBLANK, ISERROR, ISNUMBER, TYPE 등) |
| `functions/date_time.rs` | 날짜/시간 함수 (DATE, TODAY, NOW, YEAR, MONTH, DAY 등) |
| `functions/lookup.rs` | 조회 함수 (VLOOKUP, HLOOKUP, INDEX, MATCH 등) |

**유틸리티** (`utils/`):

| 파일 | 책임 |
|---|---|
| `cell_ref.rs` | 셀 참조 파싱: "A1"에서 (row, col)로 변환, 열 문자 변환, 범위 파싱 |
| `constants.rs` | 공유 상수 |

### sheetkit (파사드)

얇은 재내보내기 크레이트. `lib.rs`에 `pub use sheetkit_core::*;`가 포함되어 있어 최종 사용자는 `sheetkit`에 의존하면 전체 공개 API를 사용할 수 있다.

### sheetkit-node (packages/sheetkit)

napi-rs (v3, compat-mode 없음)를 통한 Node.js 바인딩.

- `src/lib.rs` -- 모든 바인딩을 포함하는 단일 파일. `#[napi]` `Workbook` 클래스는 `sheetkit_core::workbook::Workbook`을 `inner` 필드로 래핑한다. 메서드는 `inner`에 위임하고 Rust 타입과 napi 호환 타입 간 변환을 수행한다.
- `#[napi(object)]` 구조체는 JS 친화적 데이터 전송 타입을 정의한다 (예: `JsStyle`, `JsChartConfig`, `JsPivotTableOption`).
- napi v3의 `Either` 열거형은 다형 값을 처리한다 (예: 문자열, 숫자 또는 불리언이 될 수 있는 셀 값).
- `index.js` -- `createRequire`를 통해 `binding.cjs`를 로드하는 ESM 래퍼. 이 파일은 저장소에 체크인되어 있으며 napi 빌드 출력으로 덮어쓰면 안 된다.
- `index.d.ts` -- napi-derive가 생성한 TypeScript 타입 정의.

## 3. 주요 설계 결정

### XML 처리

대부분의 XML 파트는 `quick-xml`과 함께 `serde` 파생 매크로로 처리된다. 그러나 일부 OOXML 파트는 serde가 올바르게 처리할 수 없는 네임스페이스 접두사를 사용한다:

- **DC 네임스페이스** (코어 속성의 `dc:`, `dcterms:`, `cp:`) -- 수동 quick-xml Writer/Reader로 직렬화 및 역직렬화.
- **vt: 네임스페이스** (사용자 정의 속성의 변형 타입) -- 마찬가지로 수동 처리.

XML 선언 `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`은 ZIP에 쓰기 전에 각 직렬화된 XML 파트 앞에 수동으로 추가된다.

### ZIP 아카이브 처리

`zip` 크레이트가 .xlsx 아카이브의 읽기와 쓰기를 담당한다. 모든 ZIP 항목은 `SimpleFileOptions::default().compression_method(CompressionMethod::Deflated)`를 사용한다. `open()` 시 모든 XML 파트가 메모리로 읽혀 역직렬화된다. `save()` 시 모든 파트가 재직렬화되어 새 ZIP 아카이브에 원자적으로 기록된다.

### SharedStrings (SST)

공유 문자열 테이블은 .xlsx 파일에서 선택 사항이다. 아카이브에 `sharedStrings.xml`이 없으면 `Sst::default()`가 사용된다. 런타임에 SST는 O(1) 중복 제거를 위해 `HashMap<String, usize>`를 유지한다: 문자열 셀 값이 설정되면, SST는 문자열이 이미 존재하는 경우 기존 인덱스를 반환하고, 없으면 삽입 후 새 인덱스를 반환한다.

### 스타일 중복 제거

`add_style()`이 호출되면, 스타일 구성 요소 (글꼴, 채우기, 테두리, 정렬, 숫자 형식, 보호)가 각각 기존 레코드와 비교된다. 모든 구성 요소가 기존 XF (셀 형식) 레코드와 일치하면, 중복을 생성하지 않고 해당 레코드의 인덱스가 반환된다. 이를 통해 styles.xml을 간결하게 유지한다.

### CellValue와 날짜 감지

`CellValue`는 String, Number, Bool, Empty, Date, Formula, Error 변형을 가진 열거형이다.

읽기 시, 숫자 셀은 적용된 숫자 형식과 비교된다. 형식 ID가 알려진 날짜 형식 범위 (내장 ID 14-22, 27-36, 45-47)에 해당하거나 사용자 정의 형식 문자열이 날짜/시간 패턴을 포함하면, 셀은 `CellValue::Number` 대신 `CellValue::Date(serial_number)`로 반환된다. chrono 크레이트가 Excel 시리얼 넘버와 `NaiveDate`/`NaiveDateTime` 간의 변환을 처리한다.

### 수식 시스템

수식 시스템은 두 가지 독립적인 부분으로 구성된다:

1. **파서**: `nom` 크레이트를 사용하여 수식 문자열을 AST로 파싱한다. 파서는 연산자 우선순위, 중첩 함수 호출, 절대/상대 셀 참조, 범위 참조 및 모든 리터럴 타입을 포함하는 Excel 수식 구문을 처리한다.

2. **평가기**: AST를 순회하며 결과를 계산한다. `CellDataProvider` 트레이트가 워크북 데이터 접근을 추상화하여 평가기가 워크북을 직접 대여하지 않는다. 평가 전에, 한 셀의 수식이 다른 셀을 참조할 때 가변/불변 대여 충돌을 피하기 위해 셀 값이 `CellSnapshot` (`HashMap<(String, u32, u32), CellValue>`)에 스냅샷된다.

   `calculate_all()`은 워크북의 모든 수식 셀을 의존성 순서대로 평가한다. 셀 의존성의 방향 그래프를 구축하고, Kahn 알고리즘으로 위상 정렬을 수행하여 리프에서 루트 순서로 셀을 평가한다. 순환 참조는 감지되어 에러로 보고된다.

### 조건부 서식

조건부 서식 규칙은 전체 XF 레코드가 아닌 styles.xml의 DXF (Differential Formatting) 레코드를 참조한다. DXF 레코드는 셀의 기본 스타일과 다른 서식 속성만 포함한다. SheetKit은 셀 값 비교, 색상 스케일, 데이터 막대, 아이콘 집합, 상위/하위 N, 평균 이상/이하, 중복, 빈 셀, 에러, 수식 기반 규칙을 포함하는 18가지 규칙 유형을 지원한다.

### 피벗 테이블

OOXML의 피벗 테이블은 4개의 상호 연결된 파트로 구성된다:

1. **pivotTable 정의** (`pivotTable{n}.xml`) -- 필드 레이아웃, 행/열/데이터/페이지 필드
2. **pivotCacheDefinition** (`pivotCacheDefinition{n}.xml`) -- 소스 범위, 필드 정의
3. **pivotCacheRecords** (`pivotCacheRecords{n}.xml`) -- 캐시된 소스 데이터
4. **워크북 피벗 캐시** -- workbook.xml에서 캐시 ID를 정의에 연결하는 컬렉션

각 피벗 테이블은 전용 캐시를 갖는다. 캐시 정의는 소스 데이터 범위를 지정하고, 캐시 레코드는 해당 데이터의 스냅샷을 저장한다.

### StreamWriter

StreamWriter는 대용량 .xlsx 파일을 생성하기 위한 전방 전용 API를 제공한다. 전체 워크시트 XML을 메모리에 구축하는 대신, 행이 추가될 때 XML 요소를 출력에 직접 기록한다. 제약 사항:

- 행은 오름차순으로 작성해야 한다 (임의 접근 불가).
- `flush()`가 호출되면 시트 XML이 완료된다.
- 스트림의 SST 항목은 flush 시 워크북 SST에 병합된다.
- 틀 고정, 행 옵션 (높이, 가시성, 아웃라인), 열 수준 스타일/가시성/아웃라인을 지원한다.

### napi 바인딩 설계

Node.js 바인딩은 `inner` 필드 패턴을 따른다: napi `Workbook` 클래스는 `sheetkit_core::workbook::Workbook`을 `inner` 필드로 포함한다. 각 napi 메서드는 JS 타입에서 인수를 언래핑하고, `inner`의 해당 메서드를 호출하며, 결과를 JS 호환 타입으로 다시 변환한다.

napi v3의 `Either` 타입은 `JsUnknown` 대신 다형 값에 사용되어 Rust와 TypeScript 양쪽에서 타입 안전성을 제공한다.

## 4. 데이터 흐름

### .xlsx 파일 읽기

```
.xlsx file (ZIP archive)
  |
  v
zip::ZipArchive::new(reader)
  |
  v
For each known XML part path (e.g., "xl/workbook.xml", "xl/worksheets/sheet1.xml"):
  zip_archive.by_name(path) -> raw XML bytes
  |
  v
quick_xml::de::from_str() or manual Reader -> sheetkit-xml typed struct
  |
  v
All deserialized parts assembled into Workbook struct
  (sheets, styles, SST, relationships, drawings, charts, etc.)
```

### 데이터 조작

```
User code calls Workbook methods:
  wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))
  |
  v
Workbook locates the target worksheet (by name -> sheet index)
  |
  v
SST.get_or_insert("Hello") -> string index (O(1) via HashMap)
  |
  v
Worksheet row/cell data updated with SST index reference
```

### .xlsx 파일 쓰기

```
Workbook.save_as("output.xlsx")
  |
  v
For each XML part:
  quick_xml::se::to_string() or manual Writer -> XML string
  Prepend XML declaration
  |
  v
zip::ZipWriter::start_file(path, options)
  writer.write_all(xml_bytes)
  |
  v
zip::ZipWriter::finish() -> .xlsx file
```

## 5. 테스트 전략

- **단위 테스트**: `#[cfg(test)]` 인라인 테스트 블록을 사용하여 모듈과 함께 배치. 각 모듈은 자체 기능을 독립적으로 테스트한다.
- **Node.js 테스트**: `packages/sheetkit/__test__/index.spec.ts`에 위치. vitest를 사용하여 napi 바인딩을 엔드투엔드로 테스트한다.
- **테스트 커버리지**: 프로젝트는 모든 모듈에 걸쳐 1,100개 이상의 Rust 테스트와 100개 이상의 Node.js 테스트를 유지한다.
- **테스트 출력 파일**: 테스트 중 생성된 `.xlsx` 파일은 저장소를 깨끗하게 유지하기 위해 gitignore 처리된다.
