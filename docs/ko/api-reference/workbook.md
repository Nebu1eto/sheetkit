## 1. 워크북 입출력

워크북의 생성, 열기, 저장 및 시트 이름 조회를 다루는 기본 API이다.

### `Workbook::new()` / `new Workbook()`

빈 워크북을 생성한다. 기본적으로 "Sheet1"이라는 시트 하나가 포함된다.

**Rust:**

```rust
use sheetkit::Workbook;

let wb = Workbook::new();
```

**TypeScript:**

```typescript
import { Workbook } from "sheetkit";

const wb = new Workbook();
```

### `Workbook::open(path)` / `Workbook.open(path)`

기존 .xlsx 파일을 열어 메모리에 로드한다.

**Rust:**

```rust
let wb = Workbook::open("report.xlsx")?;
```

**TypeScript:**

```typescript
const wb = await Workbook.open("report.xlsx");
```

> 파일이 존재하지 않거나 유효한 .xlsx 형식이 아니면 오류가 발생한다.
> Node.js에서 `Workbook.open(path)`는 비동기이며 `Promise<Workbook>`을 반환한다. 동기 동작이 필요하면 `Workbook.openSync(path)`를 사용한다.

### `wb.save(path)`

워크북을 .xlsx 파일로 저장한다.

**Rust:**

```rust
wb.save("output.xlsx")?;
```

**TypeScript:**

```typescript
await wb.save("output.xlsx");
```

> ZIP 압축은 Deflate 방식을 사용한다.
> Node.js에서 `wb.save(path)`는 비동기이며 `Promise<void>`를 반환한다. 동기 동작이 필요하면 `wb.saveSync(path)`를 사용한다.

### `wb.sheet_names()` / `wb.sheetNames`

워크북에 포함된 모든 시트의 이름을 순서대로 반환한다.

**Rust:**

```rust
let names: Vec<&str> = wb.sheet_names();
// ["Sheet1", "Sheet2"]
```

**TypeScript:**

```typescript
const names: string[] = wb.sheetNames;
// ["Sheet1", "Sheet2"]
```

> TypeScript에서는 getter 프로퍼티로 접근한다.

---
