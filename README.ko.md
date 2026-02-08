# SheetKit

Excel(.xlsx) 파일을 읽고 쓰는 Rust 라이브러리. napi-rs를 통한 Node.js 바인딩 지원.

> For the English version, see [README.md](README.md).

> **주의**: SheetKit은 실험적 프로젝트입니다. API가 예고 없이 변경될 수 있습니다. 현재 활발히 개발 중입니다.

## 주요 기능

- .xlsx 파일 읽기/쓰기
- Rust 코어 + Node.js 바인딩 (napi-rs)
- 셀 조작 (문자열, 숫자, 불리언, 날짜, 수식)
- 시트 관리 (생성, 삭제, 이름 변경, 복사, 활성 시트)
- 행/열 조작 (삽입, 삭제, 크기 조정, 숨기기, 아웃라인)
- 스타일 시스템 (글꼴, 채우기, 테두리, 정렬, 숫자 서식, 보호)
- 41가지 차트 유형 (3D 지원)
- 이미지 (PNG, JPEG, GIF)
- 조건부 서식 (18가지 규칙 유형)
- 데이터 유효성 검사, 메모, 자동 필터
- 수식 계산 (110개 이상 함수)
- 대용량 데이터를 위한 스트리밍 작성기
- 셀 병합, 하이퍼링크, 틀 고정/분할
- 페이지 레이아웃 및 인쇄 설정
- 문서 속성, 통합 문서/시트 보호
- 피벗 테이블
- 정의된 이름 (명명된 범위)

## 빠른 시작

**Rust:**

```rust
use sheetkit::{CellValue, Workbook};

fn main() -> sheetkit::Result<()> {
    let mut wb = Workbook::new();
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))?;
    wb.set_cell_value("Sheet1", "B1", CellValue::Number(42.0))?;
    wb.save("output.xlsx")?;
    Ok(())
}
```

**TypeScript:**

```typescript
import { Workbook } from "sheetkit";

const wb = new Workbook();
wb.setCellValue("Sheet1", "A1", "Hello");
wb.setCellValue("Sheet1", "B1", 42);
wb.save("output.xlsx");
```

## 설치

**Rust** -- `Cargo.toml`에 추가:

```toml
[dependencies]
sheetkit = "0.1"
```

**Node.js:**

```bash
npm install sheetkit
```

## 문서

자세한 문서는 [docs/ko/](docs/ko/) 폴더를 참조하세요.

## 감사의 글

SheetKit은 Go 언어로 작성된 Excel 라이브러리인 [Excelize](https://github.com/qax-os/excelize)의 구현에서 깊은 영감을 받아 Rust와 TypeScript 생태계를 위해 만들어진 프로젝트입니다. Excelize 팀과 기여자분들의 훌륭한 작업에 깊은 존경과 감사를 표합니다.

## 라이선스

MIT OR Apache-2.0
