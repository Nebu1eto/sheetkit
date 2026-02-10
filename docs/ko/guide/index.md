# SheetKit 사용자 가이드

SheetKit은 Excel (.xlsx) 파일을 읽고 쓰는 Rust 라이브러리이며, napi-rs를 통한 Node.js 바인딩을 제공합니다.

---

## 목차

- [기본 작업](./basic-operations.md)
  - 설치
  - 빠른 시작
  - 워크북 I/O
  - 셀 조작
  - 워크북 형식 및 VBA 보존
- [스타일](./styling.md)
  - 스타일
- [데이터 기능](./data-features.md)
  - 시트 관리
  - 행과 열 조작
  - 행/열 반복자
  - 행/열 아웃라인 및 스타일
  - 차트
  - 이미지
  - 셀 병합
  - 하이퍼링크
  - 조건부 서식
  - 테이블
  - 데이터 변환 유틸리티
- [고급](./advanced.md)
  - 틀 고정/분할
  - 페이지 레이아웃
  - 데이터 유효성 검사
  - 코멘트
  - 자동 필터
  - 수식 계산
  - 피벗 테이블
  - StreamWriter
  - 문서 속성
  - 워크북 보호
  - 스파크라인
  - 정의된 이름
  - 시트 보호
  - 시트 보기 옵션
  - 시트 표시 여부
  - 예제 프로젝트
  - 유틸리티 함수
  - 테마 색상
  - 서식 있는 텍스트
  - 파일 암호화

---

## 시작하기

[기본 작업](./basic-operations.md)부터 시작하여 워크북을 생성하고 조작하는 방법을 배운 다음, [스타일](./styling.md)과 [데이터 기능](./data-features.md)을 탐색하여 더 고급 기능을 사용해보세요.

모든 API 메서드의 포괄적인 참고 자료는 [API 레퍼런스](../api-reference/index.md)를 참조하세요.

---

## 설치

### Rust

`Cargo.toml`에 `sheetkit`을 추가하세요:

```toml
[dependencies]
sheetkit = "0.1"
```

### Node.js

```bash
npm install @sheetkit/node
```

> Node.js 패키지는 napi-rs로 빌드된 네이티브 애드온입니다. 설치 중에 네이티브 모듈을 컴파일하려면 Rust 빌드 도구 체인(rustc, cargo)이 필요합니다.

---

## 빠른 예제

### Rust

```rust
use sheetkit::{CellValue, Workbook};

fn main() -> sheetkit::Result<()> {
    let mut wb = Workbook::new();
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))?;
    wb.save("output.xlsx")?;
    Ok(())
}
```

### TypeScript / Node.js

```typescript
import { Workbook } from '@sheetkit/node';

const wb = new Workbook();
wb.setCellValue('Sheet1', 'A1', 'Hello');
await wb.save('output.xlsx');
```

---

## 다음 단계

- [기본 작업](./basic-operations.md)을 자세한 예제로 배워보세요
- [스타일](./styling.md)을 적용하여 스프레드시트를 전문적으로 만들어보세요
- [데이터 기능](./data-features.md)을 추가하여 차트, 유효성 검사, 주석 등을 포함시키세요
- [고급](./advanced.md) 기능을 탐색하여 복잡한 워크북을 만들어보세요
- 저장소의 `examples/rust/` 및 `examples/node/`에서 완전한 예제 프로젝트를 확인하세요
