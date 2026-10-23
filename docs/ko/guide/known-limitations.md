# 알려진 제한 사항

이 문서는 사용자가 관찰할 수 있는 SheetKit 1.0의 알려진 제한 사항을 정리합니다. 각 항목에는 영향, 가능한 경우 우회 방법, 계획 상태를 적습니다. 현재 계획에 포함된 항목은 없으므로 모든 항목의 상태는 예정된 작업이 없음을 나타냅니다. Excel과 결과가 다른 수식은 [Excel 호환성](./excel-compatibility.md)에 정리되어 있으며, 여기에는 큰 수식 engine 제한만 다시 적습니다.

## 수식 engine

### 열 전체와 행 전체 참조는 평가되지 않습니다

`SUM(A:A)`와 `SUM(1:1)`은 cell에 저장할 수 있고, 삽입과 삭제 작업에서 Excel과 같은 방식으로 이동합니다. 하지만 계산기는 이 참조를 parse할 수 없습니다. `evaluate_formula("Sheet1", "SUM(A:A)")`는 수식 parse 오류를 반환하고, `calculate_all`은 cached result로 `#VALUE!`를 저장합니다.

- **영향**: Excel 같은 다른 application이 workbook을 다시 계산하기 전까지 SheetKit이 저장한 cached value가 틀립니다. 수식 text 자체는 올바릅니다.
- **우회 방법**: SheetKit이 값을 계산해야 한다면 `SUM(A1:A1000)`처럼 범위가 정해진 range를 사용하세요.
- **상태**: 예정된 작업이 없습니다.

### Array constant는 평가되지 않습니다

`SUM({1,2;3,4})`처럼 array constant가 있는 수식은 `evaluate_formula`에서 수식 parse 오류를 반환하며, `calculate_all`도 계산하지 못합니다.

- **영향**: 위와 같습니다. 수식 text는 그대로 저장됩니다.
- **우회 방법**: 상수를 cell에 넣고 그 range를 참조하세요.
- **상태**: 예정된 작업이 없습니다.

### 오류 인수와 일부 날짜, 텍스트 경계 입력

텍스트 함수는 텍스트 인수로 받은 오류를 그대로 전파하지만, `LEFT("abc",1/0)`처럼 숫자 인수에 들어온 오류는 오류 값이 아닌 평가 오류로 나타납니다. `COUNT`와 `SUM`은 SheetKit의 기본 숫자 표기만 숫자 텍스트로 인식하므로, Excel이 세는 `" 3 "`, `"50%"`, `"1,000"`, `"(5)"`, `"12:00"`은 세지 않습니다. `EDATE(1/0,1)`과 `YEAR(1/0)`은 `#DIV/0!` 대신 `#VALUE!`를 반환합니다. `NETWORKDAYS`, `WEEKDAY`, `WEEKNUM`은 아직 calendar 산술을 사용하므로 serial 0을 거부하고 1900년 3월 이전의 요일이 Excel과 다릅니다. Excel 결과를 포함한 전체 목록은 [Excel 호환성](./excel-compatibility.md)에 있습니다.

- **영향**: Excel에서 workbook을 다시 계산하기 전까지 이 입력의 cached value가 Excel과 다릅니다.
- **우회 방법**: SheetKit에는 없습니다. SheetKit의 cached value가 중요하다면 이 입력을 피하세요.
- **상태**: 예정된 작업이 없습니다.

## 구조 편집

### Pivot table은 행과 열의 삽입, 삭제에 맞춰 갱신되지 않습니다

`insert_rows`, `remove_row`, `insert_cols`, `remove_col`은 pivot table 위치와 pivot cache source range를 바꾸지 않습니다.

- **영향**: Pivot table이나 source data의 위쪽 또는 왼쪽을 편집하면 pivot table이 이전 좌표를 가리킵니다.
- **우회 방법**: 편집 후 `delete_pivot_table`로 pivot table을 삭제하고 `add_pivot_table`로 다시 추가하거나, Excel에서 새로 고치세요.
- **상태**: 예정된 작업이 없습니다.

### Table 안에 열을 삽입해도 table column이 추가되지 않습니다

Table 안에서 `insert_cols`를 호출하면 table range는 넓어지지만 `tableColumn` 항목은 추가되지 않습니다. `E1:F3`의 두 열 table에서 `insert_cols("Sheet1", "F", 1)`을 호출하면 range는 `E1:G3`이 되고 column은 두 개로 남습니다. Excel 16.114는 이 결과를 열 때 "We found a problem with some content"를 표시하고 파일을 복구했습니다.

- **영향**: 저장된 파일은 Excel에서 복구가 필요합니다.
- **우회 방법**: Table 밖에 열을 삽입하거나, `delete_table`로 table을 삭제한 뒤 새 column 목록으로 다시 추가하세요.
- **상태**: 예정된 작업이 없습니다.

### Shared formula의 master cell을 삭제해도 dependant가 다시 쓰이지 않습니다

Excel이 쓴 파일에서는 아래나 옆으로 채운 수식이 master cell 하나에 저장되고 다른 cell이 이를 공유하는 경우가 많습니다. `remove_row`나 `remove_col`이 master cell을 삭제하면 나머지 cell은 여전히 shared formula를 참조하지만, 그 수식을 정의하는 cell은 더 이상 없습니다.

- **영향**: SheetKit으로 읽을 때 dependant cell의 수식 text가 사라지며, Excel이 파일 복구가 필요하다고 보고할 수 있습니다.
- **우회 방법**: 편집 전에 `set_cell_formula`로 dependant cell의 수식을 명시적으로 설정하세요. SheetKit은 이를 일반 수식으로 씁니다.
- **상태**: 예정된 작업이 없습니다.

### 삭제된 구간 안의 위치는 `#REF!`가 되지 않고 이동합니다

삭제된 행이나 열을 가리키는 고정 pane, selection, drawing anchor(chart, image, shape), sparkline 위치, 내부 hyperlink 대상은 삭제 구간 바로 다음 행이나 열로 이동합니다. 반면 수식은 `#REF!`가 됩니다.

- **영향**: 내부 hyperlink나 sparkline이 편집 전과 다른 cell을 가리킬 수 있습니다.
- **우회 방법**: 삭제 영역을 가리키던 내부 hyperlink와 sparkline을 확인하고 편집 후 다시 설정하세요.
- **상태**: 예정된 작업이 없습니다.

### Table을 삭제해도 totals row 내용이 남습니다

`delete_table`은 table 정의를 제거하지만 `SUBTOTAL(…,[Column])` 수식을 포함한 totals row cell은 남깁니다. Excel 16.114에서 이렇게 남은 수식은 `=SUBTOTAL(102,#REF!)`로 표시되었습니다.

- **영향**: 남은 수식이 Excel에서 `#REF!`를 표시합니다.
- **우회 방법**: `delete_table` 후 totals row cell을 지우세요.
- **상태**: 예정된 작업이 없습니다.

## Table, slicer, comment

### Calculated column 수식과 custom totals 수식은 설정할 수 없습니다

기존 table의 `calculatedColumnFormula`와 `totalsRowFormula`는 저장할 때 보존되지만 `TableConfig`에는 이에 해당하는 field가 없습니다. 그래서 `custom` totals 함수는 거부됩니다.

- **영향**: 새 table에 calculated column이나 custom totals 수식을 만들 수 없습니다.
- **우회 방법**: `set_cell_formula`로 cell에 수식을 쓰세요. Excel은 이를 calculated column으로 취급하지 않습니다.
- **상태**: 예정된 작업이 없습니다.

### Pivot table에 연결된 slicer는 저장할 때 pivot data를 잃습니다

SheetKit은 table slicer의 slicer cache만 model로 다룹니다. Pivot table에 연결된 slicer가 있는 workbook을 열고 저장하면, slicer cache가 pivot table 목록과 item data 없이 저장됩니다.

- **영향**: SheetKit을 거쳐 저장하면 pivot slicer의 연결과 item이 사라집니다.
- **우회 방법**: 없습니다. 이런 workbook은 SheetKit으로 저장하지 않거나 Excel에서 slicer를 다시 만드세요.
- **상태**: 예정된 작업이 없습니다.

### 같은 cell의 note 작업과 thread 작업은 함께 쓸 수 없습니다

Excel은 각 threaded comment를 `[Threaded comment]` placeholder가 들어 있는 legacy note와 함께 저장하며, 한 cell에는 note 또는 thread 중 하나만 둡니다. SheetKit도 같은 규칙을 적용합니다. Thread가 있는 cell에 `add_comment`나 `remove_comment`를 호출하면 `InvalidArgument`를 반환하고, 일반 note가 있는 cell에 `add_threaded_comment`를 호출해도 `InvalidArgument`를 반환합니다. Thread root를 삭제하면 Excel처럼 reply도 함께 삭제됩니다.

- **영향**: note를 thread로 덮어쓰거나 placeholder note를 직접 수정하던 코드는 이제 Excel에서 thread를 잃는 대신 오류를 받습니다.
- **우회 방법**: Thread를 추가하기 전에 note를 제거하세요. Thread가 있는 cell을 비우려면 `delete_threaded_comment`를 사용하세요. `get_comments`는 placeholder note를 반환하지 않으므로 thread가 있는 cell은 `get_threaded_comments`로 찾으세요.
- **상태**: 설계로 해결되었습니다.

## VBA

### 동아시아 codepage의 VBA source는 손실을 감수하고 decode됩니다

`get_vba_modules`는 codepage 1252와 65001(UTF-8)의 module source를 완전히 decode합니다. Codepage 932(일본어), 949(한국어), 936(중국어 간체)에서는 ASCII와 CP932 반각 가타카나만 유지되고, 다른 2바이트 문자는 U+FFFD로 바뀝니다. Module마다 `VbaProject::warnings` 항목이 바뀐 sequence 수를 보고합니다. 그 밖의 codepage는 warning과 함께 손실이 있는 UTF-8 decode로 대체됩니다.

- **영향**: 이 module의 비ASCII identifier, 문자열, comment를 읽을 수 없습니다. Workbook을 저장할 때 VBA project byte는 그대로 보존됩니다.
- **우회 방법**: `get_vba_project`로 원본 project를 읽고 codepage를 지원하는 도구로 decode하세요.
- **상태**: 예정된 작업이 없습니다.

## 파일 열기

### 압축 비율 검사는 크기 하한을 넘는 entry에만 적용됩니다

Entry별 압축 비율 제한(기본값 100:1)은 압축을 풀었을 때 하한(기본값 8 MiB)보다 큰 ZIP entry에만 적용됩니다. 더 작은 entry는 비율 검사를 받지 않지만 entry 크기 제한과 전체 크기 제한에는 계속 포함됩니다. 하한을 넘으면 비율은 entry 전체에 적용됩니다. 압축된 byte가 0인 비어 있지 않은 entry는 항상 거부되며, 하한을 `0`으로 지정하면 모든 entry를 검사합니다.

- **영향**: 8 MiB보다 크고 비율이 100:1을 넘는 반복적인 part가 있는 정상 파일은 기본값에서 거부됩니다.
- **우회 방법**: 신뢰할 수 있는 파일에는 `OpenOptions::compression_ratio_floor_bytes`(Node: `compressionRatioFloorBytes`)로 하한을 높이거나 `OpenOptions::max_compression_ratio`(Rust 전용)로 비율을 더 큰 유한 값으로 높이세요.
- **상태**: 설계대로 동작합니다. 변경 예정은 없습니다.

### 해결됨: prefix가 있는 SpreadsheetML namespace

Root가 prefix가 있는 main namespace를 쓰는 worksheet, workbook part, shared string, style(예: `<x:worksheet xmlns:x="…">`)은 eager, lazy, stream 읽기 mode에서 모두 열리고 올바르게 읽히고 저장됩니다. Relationship namespace 선언이 빠진 part도 열립니다. 이 동작은 namespace 호환성 regression test로 검증됩니다.

## 기여자 참고 사항

### `Cargo.lock`은 commit되지 않습니다

`Cargo.lock`은 `.gitignore`에 포함되어 있으므로 각 checkout은 처음 build할 때 dependency version을 결정합니다. 따라서 같은 revision을 두 번 build해도 서로 다른 dependency version을 사용할 수 있습니다.

- **영향**: 같은 revision의 build가 재현된다고 보장할 수 없으며, revision 간 benchmark 비교에 dependency 변경과 code 변경이 섞일 수 있습니다.
- **우회 방법**: 정상 동작이 확인된 build의 `Cargo.lock`을 보관해 두고 다른 checkout에 복사하세요. Revision 비교 script는 이 작업을 자동으로 수행합니다. `benchmarks/COMPARISON.md`를 참고하세요.
- **상태**: 예정된 작업이 없습니다.
