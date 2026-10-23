# Excel 호환성

SheetKit은 Excel parity를 목표로 합니다. SheetKit이 쓰거나 편집한 workbook은 Microsoft Excel에서 복구 없이 열려야 하고, 계산과 구조 편집은 같은 입력에 대해 Excel과 같은 결과를 내야 합니다. SheetKit 1.0은 보안과 정확성을 위해 여러 동작을 강화했으며, 그중 일부는 이전 SheetKit release와의 호환성을 깨뜨립니다. 이러한 변경은 Excel과 일치할 때만 허용되므로, 이 문서는 변경된 각 동작을 같은 입력에 대한 Excel의 결과와 나란히 기록합니다.

## 검증 방법

각 동작은 같은 입력으로 두 번 실행했습니다. 한 번은 공개 SheetKit API(Rust crate, 현재 1.0 branch)로, 한 번은 2026-10-23에 AppleScript로 구동한 Microsoft Excel 16.114 for Mac으로 실행했습니다. 수식은 임시 workbook에 입력하고 Excel이 계산한 값이나 오류를 다시 읽었습니다. 구조 편집은 SheetKit이 생성한 파일에 Excel의 `delete`, `insert` 명령으로 적용한 뒤 Excel이 저장한 XML을 같은 편집에 대한 SheetKit의 출력과 비교했습니다. Table 이름과 totals row는 Excel의 list object model로 설정했습니다.

상태 값은 다음과 같습니다.

- **MATCH**: SheetKit과 Excel의 결과가 같습니다.
- **DIFFERS**: 결과가 다릅니다. Notes 열에 의도된 차이인지, 보안 제한인지, 결함인지 적습니다.
- **NOT VERIFIABLE**: AppleScript로 Excel에 같은 입력을 줄 수 없었습니다. 이유를 함께 적습니다.
- **N/A**: Excel에서 볼 수 있는 효과가 없거나 Node.js API에만 해당합니다.

AppleScript object model은 Excel 사용자 인터페이스와 같지 않습니다. 둘이 다를 수 있는 항목(예: table 이름)은 어느 쪽을 검증했는지 표에 적습니다.

Node.js method 이름은 Rust 이름의 camelCase 형태입니다(`remove_row`는 `removeRow`, `add_table`은 `addTable`). 수식 결과는 두 binding에서 같습니다.

## 집계 함수

Test data는 `A2` = `TRUE`, `A3` = `FALSE`, `A4` = text `"3"`, `A5` = `5`이며, `A1`과 `B1:B3`은 비어 있습니다.

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| `A1:A5`에 대한 `AVERAGE`, `SUM`, `MIN`, `MAX`, `COUNT` | `calculate_all` / `evaluate_formula`: 5, 5, 5, 5, 1 | 5, 5, 5, 5, 1 | MATCH | Range 인수 안의 boolean, text, 빈 cell은 무시됩니다. |
| `AVERAGE(B1:B3)`(모두 비어 있음), `AVERAGE(A2:A4)`(boolean과 text만 있음) | `#DIV/0!` | `#DIV/0!` | MATCH | |
| `B1:B3`에 대한 `MIN`, `MAX`, `SUM` | 0 | 0 | MATCH | SheetKit은 빈 range의 `SUM`을 음의 0(`-0.0`)으로 반환하며, 이 값은 0과 같다고 비교됩니다. |
| 직접 입력한 scalar 인수: `AVERAGE(TRUE)`, `SUM("3")`, `SUM(TRUE)` | 1, 3, 1 | 1, 3, 1 | MATCH | 인수로 직접 입력한 boolean과 숫자 text는 계산에 포함됩니다. |
| `AVERAGE("x")` | `calculate_all`에서 `#VALUE!` | `#VALUE!` | MATCH | `evaluate_formula`는 오류 값 대신 `Err`("cannot convert")를 반환합니다. |
| `COUNT(TRUE,"3","x")`, `COUNT("3")`, `COUNT(TRUE)` | 2, 1, 1 | 2, 1, 1 | MATCH | `COUNT` 인수로 직접 입력한 boolean과 숫자 text를 셉니다. 참조를 통해서는 숫자만 셉니다. |

## 날짜 함수

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| `DATE(10000,1,1)`, `DATE(-1,1,1)` | `#NUM!` | `#NUM!` | MATCH | |
| `DATE(0,1,1)`, `DATE(9999,12,31)` | 1, 2958465 | 1 (1900-01-01), 2958465 | MATCH | |
| `DATE(2024,2147483647,1)` | `#NUM!` | `#NUM!` | MATCH | |
| `day` = -32768, -5, 32766, 32767인 `DATE(2024,1,day)` | 12523, 45286, 78057, 78058 | 12523, 45286, 78057, 78058 | MATCH | |
| `day` = 32768, 40000, 65536, -40000인 `DATE(2024,1,day)` | 78059, 85291, 110827, 5291 | 모두 78058입니다. | DIFFERS | Excel은 -32768..32767 밖에서 시험한 모든 day 값에 대해 serial 78058(2113-09-17, day 32767의 결과)을 반환했습니다. SheetKit은 day offset을 산술적으로 적용하며 이 clamping을 재현하지 않습니다. |
| `day` = 2958101, 2147483647, 4294967296, 1E10, -2147483648인 `DATE(2024,1,day)` | `#NUM!` | 78058 | DIFFERS | 위와 같은 Excel clamping입니다. SheetKit은 결과가 1900-01-01..9999-12-31 밖이면 `#NUM!`을 반환합니다. |
| `EDATE(45000,2147483647)`, `EDATE(45000,-600)` | `#NUM!`, 26738 | `#NUM!`, 26738 | MATCH | |
| `EDATE(0,0)`, `EDATE(-1,0)` | 0, `#NUM!` | 0, `#NUM!` | MATCH | Serial 0은 1900-01-00이며, 음수 serial은 `#NUM!`을 반환합니다. |
| `EOMONTH(45000,2147483647)`, `EOMONTH(2958465,0)`, `EOMONTH(2958465,1)` | `#NUM!`, 2958465, `#NUM!` | `#NUM!`, 2958465, `#NUM!` | MATCH | |
| `EOMONTH(1,-1)`, `EOMONTH(31,-1)` | `#NUM!`, `#NUM!` | `#NUM!`, `#NUM!` | MATCH | 1900-01-01 이전 결과는 오류입니다. |
| `EOMONTH(0,0)`, `EOMONTH(-1,0)`, `EOMONTH(60,0)` | 31, `#NUM!`, 59 | 31, `#NUM!`, 59 | MATCH | |
| `WORKDAY(45000,1E17)`, `WORKDAY(45000,-1E17)`, `WORKDAY(1,1)` | `#NUM!`, `#NUM!`, 2 | `#NUM!`, `#NUM!`, 2 | MATCH | SheetKit은 반복 전에 일수를 제한합니다. |
| `WORKDAY(0,1)`, `WORKDAY(-1,1)` | 2, `#NUM!` | 2, `#NUM!` | MATCH | |

## 텍스트 함수

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| 음수 개수: `LEFT("abc",-1)`, `RIGHT("abc",-1)`, `MID("abc",1,-1)`, `REPT("a",-1)`, `REPLACE("abc",1,-1,"x")` | `#VALUE!` | `#VALUE!` | MATCH | |
| 1보다 작은 시작 위치: `MID("abc",0,1)`, `REPLACE("abc",0,1,"x")`, `REPLACE("abc",-1,1,"x")` | `#VALUE!` | `#VALUE!` | MATCH | |
| 매우 큰 개수: `LEFT("abc",1E10)`, `LEFT("abc",1E300)`, `RIGHT("abc",2147483648)` | `"abc"` | `"abc"` | MATCH | |
| 끝을 넘는 `MID`: `MID("abc",4,1)`, `MID("abc",32768,1)`, `MID("abc",2147483648,1)` | `""` | `""` | MATCH | |
| `MID("abc",1,2147483647)` | `"abc"` | `"abc"` | MATCH | |
| `MID("abc",1,2147483648)`, `MID("abc",1,1E10)`, `MID("abc",1E10,1)` | `#VALUE!` | `#VALUE!` | MATCH | 2147483648보다 큰 시작 위치와 2147483647보다 큰 길이를 거부합니다. |
| `REPT` 결과 길이: `LEN(REPT("a",32767))`, `REPT("a",32768)`, `REPT("ab",1E10)` | 32767, `#VALUE!`, `#VALUE!` | 32767, `#VALUE!`, `#VALUE!` | MATCH | 32767자는 Excel의 cell text 한도입니다. |
| `REPLACE("abc",4,1,"x")` | `"abcx"` | `"abcx"` | MATCH | |
| 시작 위치나 길이가 32767보다 큰 `REPLACE`: 시작 32768, 2147483647, 1E10, 길이 2147483647, 2147483648, 1E10 | `#VALUE!` | `#VALUE!` | MATCH | |
| 오류 인수의 `LEN`: `LEN(1/0)`, `LEN(REPT("a",32768))` | `#DIV/0!`, `#VALUE!` | `#DIV/0!`, `#VALUE!` | MATCH | Text 함수는 오류 인수를 전파합니다. |

## 반올림

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| `ROUND(1,309)`, `ROUNDDOWN(1,-309)`, `ROUND(123.456,-2)` | 1, 0, 100 | 1, 0, 100 | MATCH | |
| `ROUNDUP(123.456,-400)` | `#NUM!` | `#NUM!` | MATCH | |
| `ROUND(1.7E308,-308)` | `#NUM!` | 수식을 입력할 수 없습니다. | NOT VERIFIABLE | Excel은 literal `1.7E308`을 거부하며(가장 큰 숫자 literal은 약 9.99999999999999E+307입니다), AppleScript 실행 후 cell은 비어 있었습니다. SheetKit은 literal을 받아들이고 반올림 결과가 overflow되므로 `#NUM!`을 반환합니다. |
| `ROUND(9E307,-307)`, `ROUND(9.9E307,-307)` | 9.000000000000004E+307, 1.0000000000000004E+308 | 9E+307, 1E+308 | DIFFERS | 극단적인 크기에서 생기는 부동소수점 오차입니다. 상대 오차는 약 4E-16입니다. |

## TEXT 숫자 서식

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| `TEXT(1.25,"# ?/8")` | `"1 2/8"` | `"1 2/8"` | MATCH | |
| `TEXT(1,"0E+0")`, `TEXT(1,"ŉ0")` | `"1E+0"`, `"ŉ1"` | `"1E+0"`, `"ŉ1"` | MATCH | |
| `TEXT(1,"ŉE+é")` | `"1E+0"` | `#VALUE!` | DIFFERS | 보안 제한입니다. SheetKit은 모든 format code에 대해 유한하고 크기가 제한된 출력을 만들며, Excel이 잘못되었다고 보는 code를 모두 거부하지는 않습니다. |
| `TEXT(42949672964294967296,"/?")`, `TEXT(1E15,"/?")` | `"42949672964294970000"`, `"1000000000000000    "` | `#VALUE!`, `#VALUE!` | DIFFERS | 위와 같은 보안 제한입니다. |
| `TEXT(1.5,"?/18446744073709551615")`, `TEXT(1.5,"?/1000000000")` | `"3/2"`, `"3/2"` | `#VALUE!`, `#VALUE!` | DIFFERS | 보안 제한입니다. SheetKit은 사용할 수 없는 분모에 대해 제한된 분수 탐색으로 대체합니다. |
| `TEXT(1.5,"?/99999")` | `"149999/99999"` | `"3/9    "` | DIFFERS | Excel 출력을 반환된 그대로 기록했습니다. SheetKit은 고정 분모를 적용합니다. |
| `TEXT(12345678901234567890,"0")` | `"12345678901234567168"` | `"12345678901234500000"` | DIFFERS | Excel은 유효 숫자 15자리를 표시하고 SheetKit은 double 값 전체를 출력합니다. |

## 다른 sheet의 수식 평가

Test data: `Sheet2!A5` = 1, `Sheet2!B5` = `=A5+ROW()+COLUMN()*100`, `Sheet2!C5` = `=SUM(A5:B5)`, `Sheet1!A5` = 5.

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| Sheet1의 `=Sheet2!B5` | 206 | 206 | MATCH | 한정자가 없는 `A5`, `ROW()`, `COLUMN()`은 수식이 있는 Sheet2를 기준으로 해석됩니다. |
| Sheet1의 `=Sheet2!C5`, `=SUM(Sheet2!A5:B5)` | 207, 207 | 207, 207 | MATCH | |

## 행 삭제

Sheet1의 5행과 6행을 삭제했습니다(SheetKit에서는 `remove_row`를 두 번, Excel에서는 `A5:A6`의 `delete entire row`). Sheet1과 다른 sheet의 수식이 Sheet1을 참조했습니다.

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| 삭제된 cell에 대한 단일 참조: `=A5`, `=Sheet1!A5` | `=#REF!`, `=Sheet1!#REF!` | `=#REF!`, `=Sheet1!#REF!` | MATCH | |
| 일부가 삭제된 range: `SUM(A4:A7)`, `SUM(Sheet1!A6:A9)` | `SUM(A4:A5)`, `SUM(Sheet1!A5:A7)` | 같습니다. | MATCH | Range가 줄어듭니다. |
| 전부 삭제된 range: `SUM(A5:A6)`, `SUM(Sheet1!5:5)` | `SUM(#REF!)`, `SUM(Sheet1!#REF!)` | 같습니다. | MATCH | |
| 삭제 구간 아래의 참조: `Sheet1!A10`, `SUM(Sheet1!4:7)` | `Sheet1!A8`, `SUM(Sheet1!4:5)` | 같습니다. | MATCH | |
| 열 전체 참조: `SUM(A:A)` | 변경되지 않습니다. | 변경되지 않습니다. | MATCH | |
| 병합 cell `C5:D6`, `E4:E7`, `F8:G9` | 제거, `E4:E5`, `F6:G7` | 같습니다. | MATCH | |
| 데이터 유효성 `H5`, `H8`, `H3:H7`(같은 규칙) | 제거, `H6`, `H3:H5` | `H3:H6`의 규칙 하나 | MATCH | 같은 cell을 포함합니다. Excel은 저장할 때 같은 규칙을 `sqref` 하나로 합칩니다. |
| 조건부 서식 `I5:I6`, `I4:I8` | 제거, `I4:I6` | 같습니다. | MATCH | |
| `J5`, `J8`의 hyperlink | 제거, `J6` | 같습니다. | MATCH | |
| `L5`, `L8`의 comment | 제거, `L6` | 같습니다. | MATCH | |
| 5행과 8행 아래의 행 page break | 제거, 6행 아래 | 같습니다. | MATCH | |

## 행과 열 삽입

위 삭제 후 A열 앞에 열 하나, 1행 앞에 행 하나를 삽입했습니다(SheetKit에서는 `insert_cols`, `insert_rows`, Excel에서는 열 전체와 행 전체에 `insert`).

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| 열 전체 참조: `SUM(A:A)`, `SUM(Sheet1!B:B)`, `SUM(Sheet1!A:B)` | `SUM(B:B)`, `SUM(Sheet1!C:C)`, `SUM(Sheet1!B:C)` | 같습니다. | MATCH | |
| 행 전체 참조: `SUM(2:2)`, `SUM(Sheet1!1:1)`, `SUM(Sheet1!4:5)` | `SUM(3:3)`, `SUM(Sheet1!2:2)`, `SUM(Sheet1!5:6)` | 같습니다. | MATCH | 열 전체와 행 전체 참조는 이동하지만 평가되지는 않습니다. [알려진 제한 사항](./known-limitations.md)을 참고하세요. |
| 병합, 조건부 서식, hyperlink, comment, page break | `F5:F6`, `G7:H8`; `J5:J7`; `K7`; `M7`; 7행 아래 | 같습니다. | MATCH | 데이터 유효성은 같은 cell을 포함하며, Excel은 이번에도 같은 규칙을 합칩니다. |

## Table

SheetKit에서는 `add_table`로 table 이름을 지정했습니다. Excel에서는 AppleScript로 list object의 `name`을 설정하고 workbook을 저장한 뒤, 저장된 `name`과 `displayName`을 읽었습니다. Excel 사용자 인터페이스의 table 이름 상자는 시험하지 않았습니다.

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| 기존 table과 대소문자만 다른 이름(`Sales`가 있을 때 `sales`, `SALES`) | `add_table`이 오류("table 'sales' already exists")를 반환합니다 | 허용합니다. `name="sales"`, `displayName="sales_1"`로 저장합니다. | DIFFERS | 의도된 차이입니다. Excel object model은 display name을 바꾸고, SheetKit은 조용히 이름을 바꾸지 않고 거부합니다. |
| Cell 참조와 같은 이름(`A1`, `XFD1048576`, `R1C1`, `R`, `C`, `RC`, `R1`) | 오류("must not be a cell reference") | `A1`은 `displayName="A1_"`, `R1C1`은 `"_R1C1"`로 저장합니다. | DIFFERS | 위와 같은 의도된 차이입니다. 마지막 열을 넘는 `XFE1`은 두 쪽 모두 허용합니다. |
| 공백이 있는 이름(`Bad Name`) | 오류("contains invalid character") | `displayName="Bad_Name"`으로 저장합니다. | DIFFERS | 위와 같은 의도된 차이입니다. |
| 숫자로 시작하는 이름(`1abc`) | 오류("must start with a letter, underscore, or backslash") | `displayName="_1abc"`로 저장합니다. | DIFFERS | 위와 같은 의도된 차이입니다. |
| 256자 이름, 255자 이름 | 오류, 허용 | 오류 -10006("Can't set name"), 허용 | MATCH | |
| `_ok`, `\back`, `a.b` | 허용 | 허용 | MATCH | |
| Totals row 존재 | `totals_row_function`이나 `totals_row_label`이 있는 `TableConfig` column은 `totalsRowCount="1"`을 씁니다. Table `ref`는 totals row를 포함하고 `autoFilter` ref는 제외합니다 | `totalsRowCount="1"`, `ref="A1:B4"`, `autoFilter ref="A1:B3"` | MATCH | |
| `sum`, `average`, `count`, `countNums`, `min`, `max`, `stdDev`, `var`의 `SUBTOTAL` 함수 번호 | 109, 101, 103, 102, 105, 104, 107, 110 | 109, 101, 103, 102, 105, 104, 107, 110 | MATCH | |
| 허용되는 totals 함수 이름 | `none`, `sum`, `average`, `count`, `countNums`, `min`, `max`, `stdDev`, `var`. `custom`은 거부됩니다 | 같은 이름과 `custom` | DIFFERS | 지원하지 않습니다. Custom totals 수식에는 `totalsRowFormula`가 필요한데 `TableConfig`로 설정할 수 없습니다. [알려진 제한 사항](./known-limitations.md)을 참고하세요. |
| Totals cell 수식 text | `SUBTOTAL(109,[Qty])` | Cell에는 `=SUBTOTAL(109,[Qty])`로 표시하고 `SUBTOTAL(109,Sales[Qty])`로 저장합니다. | MATCH | Excel은 SheetKit 형식을 복구 없이 열고 계산합니다. |

## 병합 cell

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| `A1:B2`가 이미 병합된 상태에서 `B2:C3` 병합 | `Workbook::merge_cells`와 `StreamWriter::add_merge_cell`이 병합 겹침 오류를 반환합니다 | 허용합니다. 병합이 `A1:C3`이 됩니다. | DIFFERS | 의도된 차이입니다. SheetKit은 기존 병합을 조용히 넓히지 않고 겹치는 병합을 거부합니다. |
| 같은 range를 다시 병합(`E1:F2` 두 번) | `StreamWriter::add_merge_cell`은 아무 작업도 하지 않습니다. `Workbook::merge_cells`는 겹침 오류를 반환합니다. | 아무 작업도 하지 않습니다. | MATCH (stream), DIFFERS (`Workbook`) | `Workbook`의 차이는 의도된 것입니다. 기존 병합에 `merge_cells`를 다시 호출하면 무시하지 않고 보고합니다. |

## Comment와 shared string

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| Note 하나와 threaded comment 하나가 있는 sheet의 note 목록 | `get_comments`가 note만 반환합니다 | `Excel comments` 개수는 1(note)입니다. | MATCH | Excel은 `[Threaded comment]` legacy twin을 note로 나열하지 않습니다. |
| Excel에 표시되는 threaded comment 내용 | `get_threaded_comments`가 thread를 반환합니다 | 읽을 수 없습니다. | NOT VERIFIABLE | Excel의 AppleScript dictionary는 threaded comment에 접근할 수 없습니다. |
| `A1`의 일반 `Hello`와 `A2`의 굵은 rich text `Hello` | Shared string 항목 두 개 | `A1` 문자는 굵지 않고 `A2` 문자는 굵습니다. | MATCH | 두 cell은 서식을 따로 유지합니다. |

## Excel에서 보이지 않는 동작

| 동작 | SheetKit | Excel | 상태 | Notes |
| --- | --- | --- | --- | --- |
| Relationship을 공유하는 두 cell 중 하나의 hyperlink 삭제 | `delete_cell_hyperlink`는 남은 hyperlink가 relationship을 쓰지 않을 때만 relationship을 제거합니다 | 해당하지 않습니다. | N/A | Package 관리 동작입니다. |
| 새 relationship id | 사용하지 않는 가장 작은 `rIdN`을 쓰므로 빈 번호를 재사용합니다 | 해당하지 않습니다. | N/A | Package 관리 동작입니다. |
| Node.js 정수 인수(행, 열, 개수, style id, outline level, 숫자 style과 option field) | 유한하지 않은 값, 소수, 안전하지 않은 값(`Number.MAX_SAFE_INTEGER` 초과), 범위 밖의 값은 "`<name>` must be an integer from `<min>` to `<max>`" 오류를 던집니다 | 해당하지 않습니다. | N/A | API 계약입니다. 이전 release에서는 napi가 값을 자르거나, wrap하거나, `NaN`을 0으로 바꿨습니다. |

## 요약

위 74개 행 중 54개는 MATCH, 14개는 DIFFERS, 1개는 분할(같은 range 재병합은 `StreamWriter`에서는 일치하고 `Workbook`에서는 다름), 2개는 NOT VERIFIABLE, 3개는 N/A입니다.

- **보안 제한**: `TEXT` format code 행입니다. SheetKit은 Excel이 거부하는 format code를 거부하는 대신 서식 결과의 크기를 제한합니다.
- **의도된 차이**: table 이름(이름을 바꾸지 않고 거부)과 병합(넓히지 않고 겹침을 거부)입니다.
- **재현하지 않는 동작**: 범위를 벗어난 `DATE` day 인수에 대한 Excel의 clamping, `TEXT`의 Excel 15자리 표시 정밀도, 1E308 근처 `ROUND`의 부동소수점 오차입니다.
