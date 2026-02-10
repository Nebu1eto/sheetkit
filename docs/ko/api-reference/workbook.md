## 1. 워크북 입출력

워크북의 생성, 열기, 저장 및 시트 이름 조회를 다루는 기본 API입니다.

### `Workbook::new()` / `new Workbook()`

빈 워크북을 생성합니다. 기본적으로 "Sheet1"이라는 시트 하나가 포함됩니다.

**Rust:**

```rust
use sheetkit::Workbook;

let wb = Workbook::new();
```

**TypeScript:**

```typescript
import { Workbook } from "@sheetkit/node";

const wb = new Workbook();
```

### `Workbook::open(path)` / `Workbook.open(path)`

기존 .xlsx 파일을 열어 메모리에 로드합니다.

**Rust:**

```rust
let wb = Workbook::open("report.xlsx")?;
```

**TypeScript:**

```typescript
const wb = await Workbook.open("report.xlsx");
```

> 파일이 존재하지 않거나 유효한 .xlsx 형식이 아니면 오류가 발생합니다.
> Node.js에서 `Workbook.open(path)`는 비동기이며 `Promise<Workbook>`을 반환합니다. 동기 동작이 필요하면 `Workbook.openSync(path)`를 사용합니다.

### `wb.save(path)`

워크북을 .xlsx 파일로 저장합니다.

**Rust:**

```rust
wb.save("output.xlsx")?;
```

**TypeScript:**

```typescript
await wb.save("output.xlsx");
```

> ZIP 압축은 Deflate 방식을 사용합니다.
> Node.js에서 `wb.save(path)`는 비동기이며 `Promise<void>`를 반환합니다. 동기 동작이 필요하면 `wb.saveSync(path)`를 사용합니다.

### `Workbook::open_from_buffer(data)` / `Workbook.openBufferSync(data)`

파일 경로 대신 메모리 내 바이트 버퍼에서 워크북을 엽니다. 업로드된 파일이나 네트워크를 통해 수신한 데이터를 처리할 때 유용합니다.

**Rust:**

```rust
let data: Vec<u8> = std::fs::read("report.xlsx")?;
let wb = Workbook::open_from_buffer(&data)?;
```

**TypeScript:**

```typescript
// 동기
const data: Buffer = fs.readFileSync("report.xlsx");
const wb = Workbook.openBufferSync(data);

// 비동기
const wb2 = await Workbook.openBuffer(data);
```

> Node.js에서 `Workbook.openBuffer(data)`는 비동기이며 `Promise<Workbook>`을 반환합니다. 동기 동작이 필요하면 `Workbook.openBufferSync(data)`를 사용합니다.

### `wb.save_to_buffer()` / `wb.writeBufferSync()`

워크북을 디스크에 쓰지 않고 메모리 내 바이트 버퍼로 직렬화합니다. HTTP 응답으로 파일을 전송하거나 서비스 간 데이터를 전달할 때 유용합니다.

**Rust:**

```rust
let buf: Vec<u8> = wb.save_to_buffer()?;
// buf에 유효한 .xlsx 데이터가 들어 있다
```

**TypeScript:**

```typescript
// 동기
const buf: Buffer = wb.writeBufferSync();

// 비동기
const buf2: Buffer = await wb.writeBuffer();
```

> Node.js에서 `wb.writeBuffer()`는 비동기이며 `Promise<Buffer>`를 반환합니다. 동기 동작이 필요하면 `wb.writeBufferSync()`를 사용합니다.

### `wb.sheet_names()` / `wb.sheetNames`

워크북에 포함된 모든 시트의 이름을 순서대로 반환합니다.

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

> TypeScript에서는 getter 프로퍼티로 접근합니다.

---
