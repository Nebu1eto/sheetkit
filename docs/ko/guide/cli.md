# CLI 도구

SheetKit은 터미널에서 Excel (.xlsx) 파일을 직접 다룰 수 있는 커맨드라인 도구를 포함합니다. CLI를 통해 코드를 작성하지 않고도 일반적인 워크북 작업을 빠르게 수행할 수 있습니다.

## 설치

`cli` feature를 활성화하여 소스에서 빌드합니다:

```bash
cargo install sheetkit --features cli
```

또는 로컬에서 빌드합니다:

```bash
cargo build --release --package sheetkit --features cli
```

바이너리는 `target/release/sheetkit` 경로에 생성됩니다.

## 명령어

### info

시트 이름, 활성 시트, 문서 속성 등 워크북 정보를 표시합니다.

```bash
sheetkit info report.xlsx
```

출력 예시:

```
File: report.xlsx
Sheets: 3
  1: Summary (active)
  2: Data
  3: Charts
Creator: John Smith
Modified: 2025-01-15T10:30:00Z
```

### sheets

워크북의 모든 시트 이름을 한 줄에 하나씩 출력합니다.

```bash
sheetkit sheets report.xlsx
```

출력 예시:

```
Summary
Data
Charts
```

스크립팅에 활용할 수 있습니다:

```bash
for sheet in $(sheetkit sheets report.xlsx); do
  sheetkit convert report.xlsx -f csv -o "${sheet}.csv" --sheet "$sheet"
done
```

### read

시트 데이터를 읽어서 표시합니다. 기본적으로 활성 시트를 탭 구분 형식으로 출력합니다.

```bash
# 활성 시트를 테이블로 읽기
sheetkit read report.xlsx

# 특정 시트 읽기
sheetkit read report.xlsx --sheet Data

# CSV 형식으로 출력
sheetkit read report.xlsx --format csv
```

옵션:

| 플래그 | 단축 | 설명 |
|--------|------|------|
| `--sheet <name>` | `-s` | 읽을 시트 (기본값: 활성 시트) |
| `--format <fmt>` | `-f` | 출력 형식: `table` (기본값) 또는 `csv` |

### get

단일 셀 값을 가져옵니다.

```bash
sheetkit get report.xlsx Sheet1 A1
```

셀 값이 stdout으로 출력됩니다. 빈 셀은 출력이 없습니다. 정수인 숫자 값은 소수점 없이 표시됩니다. 불리언 값은 `TRUE` 또는 `FALSE`로 표시됩니다.

### set

셀 값을 설정하고 새 파일로 저장합니다.

```bash
sheetkit set report.xlsx Sheet1 A1 "New Title" -o updated.xlsx
```

값은 자동으로 해석됩니다:
- `TRUE` / `FALSE` (대소문자 무관)는 불리언으로 저장됩니다.
- 유효한 숫자는 숫자 값으로 저장됩니다.
- 그 외의 값은 문자열로 저장됩니다.

옵션:

| 플래그 | 단축 | 설명 |
|--------|------|------|
| `--output <path>` | `-o` | 출력 파일 경로 (필수) |

### convert

시트를 다른 형식으로 변환합니다.

```bash
# 활성 시트를 CSV로 변환
sheetkit convert report.xlsx -f csv -o output.csv

# 특정 시트를 변환
sheetkit convert report.xlsx -f csv -o data.csv --sheet Data
```

옵션:

| 플래그 | 단축 | 설명 |
|--------|------|------|
| `--format <fmt>` | `-f` | 대상 형식: `csv` (필수) |
| `--output <path>` | `-o` | 출력 파일 경로 (필수) |
| `--sheet <name>` | `-s` | 변환할 시트 (기본값: 활성 시트) |

## 종료 코드

| 코드 | 의미 |
|------|------|
| 0 | 성공 |
| 1 | 오류 (잘못된 파일, 존재하지 않는 시트, 잘못된 셀 참조 등) |

오류 메시지는 stderr로 출력됩니다.

## 사용 예시

```bash
# 알 수 없는 스프레드시트를 빠르게 확인
sheetkit info data.xlsx
sheetkit sheets data.xlsx
sheetkit read data.xlsx | head -5

# 스크립트에서 특정 값 추출
total=$(sheetkit get data.xlsx Summary B10)
echo "Total: $total"

# 셀을 업데이트하고 내보내기
sheetkit set template.xlsx Report A1 "Q4 2025" -o report.xlsx
sheetkit convert report.xlsx -f csv -o report.csv

# 모든 시트를 CSV로 내보내기
for sheet in $(sheetkit sheets workbook.xlsx); do
  sheetkit convert workbook.xlsx -f csv -o "${sheet}.csv" --sheet "$sheet"
done
```
