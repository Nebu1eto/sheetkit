# SheetKit 기여 가이드

## 1. 사전 요구 사항

- **Rust 툴체인** (rustc, cargo) -- 최신 안정 릴리스
- **Node.js** >= 18
- **pnpm** >= 9

## 2. 저장소 설정

```bash
git clone <repo-url>
cd sheetkit
pnpm install
cargo build --workspace
```

클론 후 전체 워크스페이스가 빌드되고 모든 테스트가 통과하는지 확인한다:

```bash
cargo test --workspace
cargo clippy --workspace
cargo fmt --check
```

## 3. 프로젝트 구조

```
sheetkit/
  crates/
    sheetkit-xml/      # XML 스키마 타입 (serde 기반 OOXML 매핑)
    sheetkit-core/     # 핵심 비즈니스 로직 (모든 기능이 이 크레이트에 구현됨)
    sheetkit/          # 공개 파사드 크레이트 (sheetkit-core에서 재내보내기)
  packages/
    sheetkit/          # napi-rs를 통한 Node.js 바인딩
  examples/
    rust/              # Rust 사용 예제
    node/              # Node.js 사용 예제
  docs/                # 문서
```

## 4. 빌드 명령어

### Rust 워크스페이스

```bash
cargo build --workspace        # 모든 크레이트 빌드
cargo test --workspace         # 모든 Rust 테스트 실행
cargo clippy --workspace       # 린트 (경고가 없어야 함)
cargo fmt --check              # 포매팅 확인
```

### Node.js 바인딩

```bash
cd packages/sheetkit

# 네이티브 애드온 빌드 (Rust cdylib 컴파일, index.js와 index.d.ts 생성)
pnpm build

# Node.js 테스트 스위트 실행
pnpm test
```

`pnpm build` 명령은 `napi build --platform --release --esm`을 실행하며, ESM 출력을 직접 생성한다. 수동 파일 이름 변경이 필요 없다.

## 5. 개발 워크플로우

SheetKit은 TDD (테스트 주도 개발) 접근 방식을 따른다:

1. 기대하는 동작을 설명하는 **테스트를 먼저 작성**한다.
2. 모든 테스트가 통과할 때까지 **기능을 구현**한다.
3. 작업이 완료되기 전에 **전체 검증 체크리스트를 실행**한다.

### 검증 체크리스트

모든 변경 사항은 제출 전에 다음을 모두 통과해야 한다:

- [ ] `cargo build --workspace` -- 에러 없이 컴파일
- [ ] `cargo test --workspace` -- 모든 테스트 통과
- [ ] `cargo clippy --workspace` -- 경고 없음
- [ ] `cargo fmt --check` -- 포매팅 올바름
- [ ] `cd packages/sheetkit && npx vitest run` -- Node.js 테스트 통과 (바인딩이 변경된 경우)

## 6. 코드 스타일

### Rust

- 표준 Rust 관례를 따른다. 자동 포매팅에는 `cargo fmt`를 사용한다.
- `cargo fmt`가 직접 수정하지 않은 크레이트의 파일을 재포매팅할 수 있다는 점에 유의한다. 이러한 재포매팅된 파일도 커밋에 포함한다.

### TypeScript / JavaScript

- 포매팅과 린팅에 Biome을 사용한다.
- 모든 JavaScript 및 TypeScript 코드에 ESM만 사용한다.

### 일반 규칙

- **코드 내 모든 곳에서 영어 사용**: 모든 변수 이름, 문자열 리터럴, 주석, 예제 데이터 값은 영어여야 한다. 데모나 테스트 데이터도 영어 문자열을 사용해야 한다 (예: "Name", "Sales", "Employee List").
- **문서 주석**: Rust에는 `///`, TypeScript에는 `/** */`를 사용한다. 입력, 동작, 출력을 간결하게 설명한다.
- **인라인 주석**: 코드 자체에서 자명하지 않은 로직에만 사용한다.
- **섹션 마커나 장식적 주석 금지**: 주석 배너, 구분선 또는 장식적 마커를 추가하지 않는다.

## 7. 새 기능 추가

새 기능을 구현할 때 다음 단계를 따른다:

### 1단계: XML 타입 (필요한 경우)

기능에 새로운 OOXML XML 구조가 필요한 경우 `crates/sheetkit-xml/src/`에 serde 기반 타입을 추가한다. 관련된 OOXML 파트에 따라 새 파일을 만들거나 기존 파일을 확장한다.

### 2단계: 핵심 비즈니스 로직

`crates/sheetkit-core/src/`에 기능을 구현한다. 여러 기여자가 동시에 작업할 때 병합 충돌을 최소화하기 위해 로직을 자체 모듈 파일 (예: `feature_name.rs`)에 배치한다.

같은 파일 내의 `#[cfg(test)]` 인라인 테스트 모듈에 테스트를 작성한다. 테스트는 사소한 속성이 아닌 필수적인 동작을 검증해야 한다.

`crates/sheetkit-core/src/lib.rs`에 새 모듈을 등록한다.

### 3단계: 파사드 재내보내기 (필요한 경우)

최종 사용자가 접근해야 하는 새로운 공개 타입이 도입된 경우 `crates/sheetkit/src/lib.rs`를 통해 재내보내기되는지 확인한다.

### 4단계: Node.js 바인딩

`packages/sheetkit/src/lib.rs`에 napi 바인딩을 추가한다:

- 핵심 구현에 위임하는 `#[napi]` 메서드를 `Workbook` 클래스에 추가한다.
- 새로운 설정이나 결과 타입에 대해 `#[napi(object)]` 구조체를 정의한다.
- 다형 매개변수나 반환 값에 `Either` 타입을 사용한다.

### 5단계: Node.js 테스트

`packages/sheetkit/__test__/index.spec.ts`에 새 바인딩을 다루는 테스트 케이스를 추가한다.

### 6단계: 재빌드 및 검증

napi 바인딩을 재빌드하고 전체 검증 체크리스트를 실행한다 (5절 참조).

## 8. 워크스페이스 레이아웃

### Cargo 워크스페이스

Cargo 워크스페이스에는 다음이 포함된다:

- `crates/sheetkit-xml`
- `crates/sheetkit-core`
- `crates/sheetkit`
- `packages/sheetkit` (napi-rs 크레이트, Cargo.toml에서 이름은 `sheetkit-node`)
- `examples/rust`

### pnpm 워크스페이스

pnpm 워크스페이스에는 다음이 포함된다:

- `packages/*`
- `examples/*`

## 9. 주요 의존성

| 크레이트 | 용도 |
|---|---|
| `quick-xml` | XML 파싱 및 직렬화 (`serialize` 및 `overlapped-lists` 기능 포함) |
| `serde` | XML 타입을 위한 파생 기반 (역)직렬화 |
| `zip` | .xlsx 파일을 위한 ZIP 아카이브 처리 (`deflate` 기능 포함) |
| `thiserror` | 인체공학적 에러 타입 정의 |
| `nom` | 수식 문자열 파싱 (AST 생성) |
| `napi` / `napi-derive` | Node.js 네이티브 애드온 바인딩 (v3, compat-mode 없음) |
| `chrono` | 날짜/시간 처리 및 Excel 시리얼 넘버 변환 |
| `tempfile` | 테스트에서 임시 파일 생성 |
| `pretty_assertions` | 테스트에서 개선된 어설션 diff 출력 |

## 10. 자주 발생하는 문제

### cargo fmt 부작용

`cargo fmt`가 직접 변경하지 않은 크레이트의 파일을 재포매팅할 수 있다. 포매팅 후 항상 `git diff`를 확인하고 재포매팅된 파일을 커밋에 포함한다.

### ZIP 압축 옵션

ZIP 항목을 쓸 때 항상 다음을 사용한다:

```rust
SimpleFileOptions::default().compression_method(CompressionMethod::Deflated)
```

### napi 빌드 출력

napi v3와 `--esm` 플래그를 사용하면 ESM 출력이 직접 생성된다. `packages/sheetkit`에서 `pnpm build`를 실행하면 `index.js`와 `index.d.ts`가 바로 사용 가능하다.

### 파일 구성

새 기능은 기존의 큰 파일 (특히 `workbook.rs`)에 추가하지 말고 자체 소스 파일에 배치한다. 이렇게 하면 여러 사람이 동시에 다른 기능을 작업할 때 병합 충돌이 줄어든다.

### 커밋을 위한 파일 스테이징

변경한 특정 파일만 스테이징한다. `git add -A`나 `git add .`를 사용하지 않는다. 이는 생성된 파일, 테스트 출력 또는 기타 의도하지 않은 변경 사항을 실수로 포함할 수 있다.

### 테스트 출력 파일

테스트에서 생성된 `.xlsx` 파일은 gitignore 처리되며 저장소에 커밋해서는 안 된다.

### 네임스페이스 접두사 처리

일부 OOXML 네임스페이스 접두사 (`dc:`, `dcterms:`, `cp:`, `vt:`)는 serde로 처리할 수 없으며 수동 quick-xml Writer/Reader 코드가 필요하다. 기능이 문서 속성이나 유사한 네임스페이스 접두사 요소를 포함하는 경우 수동 직렬화/역직렬화 로직을 작성해야 한다.

### 수식 파서

수식 파서는 수동 작성 파서가 아닌 `nom` 크레이트를 사용한다. 수식 지원을 확장해야 하는 경우 `crates/sheetkit-core/src/formula/parser.rs`의 nom 조합자에 먼저 익숙해져야 한다.
