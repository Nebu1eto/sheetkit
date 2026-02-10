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
// Sync
const data: Buffer = fs.readFileSync("report.xlsx");
const wb = Workbook.openBufferSync(data);

// Async
const wb2 = await Workbook.openBuffer(data);
```

> Node.js에서 `Workbook.openBuffer(data)`는 비동기이며 `Promise<Workbook>`을 반환합니다. 동기 동작이 필요하면 `Workbook.openBufferSync(data)`를 사용합니다.

### `wb.save_to_buffer()` / `wb.writeBufferSync()`

워크북을 디스크에 쓰지 않고 메모리 내 바이트 버퍼로 직렬화합니다. HTTP 응답으로 파일을 전송하거나 서비스 간 데이터를 전달할 때 유용합니다.

**Rust:**

```rust
let buf: Vec<u8> = wb.save_to_buffer()?;
// buf contains valid .xlsx data
```

**TypeScript:**

```typescript
// Sync
const buf: Buffer = wb.writeBufferSync();

// Async
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

### `wb.format()` / `wb.format`

워크북 형식을 반환합니다. 파일을 열 때 패키지의 콘텐츠 타입에서 자동으로 감지됩니다. 이 형식은 저장 시 `xl/workbook.xml`에 사용되는 OOXML 콘텐츠 타입을 결정합니다.

**Rust:**

```rust
use sheetkit::{Workbook, WorkbookFormat};

let wb = Workbook::open("report.xlsm")?;
let fmt: WorkbookFormat = wb.format();
assert_eq!(fmt, WorkbookFormat::Xlsm);
```

**TypeScript:**

```typescript
const wb = await Workbook.open("report.xlsm");
const fmt: string = wb.format; // "xlsm"
```

### `wb.set_format(format)` / `wb.format = ...`

워크북 형식을 명시적으로 설정합니다. 자동 감지된 형식을 덮어쓰며 저장 시 출력되는 콘텐츠 타입을 제어합니다.

**Rust:**

```rust
use sheetkit::{Workbook, WorkbookFormat};

let mut wb = Workbook::new();
wb.set_format(WorkbookFormat::Xlsm);
wb.save("macros.xlsm")?;
```

**TypeScript:**

```typescript
const wb = new Workbook();
wb.format = "xlsm";
await wb.save("macros.xlsm");
```

### WorkbookFormat

| Rust | TypeScript | 확장자 | 설명 |
|---|---|---|---|
| `WorkbookFormat::Xlsx` | `"xlsx"` | `.xlsx` | 표준 스프레드시트 (기본값) |
| `WorkbookFormat::Xlsm` | `"xlsm"` | `.xlsm` | 매크로 사용 스프레드시트 |
| `WorkbookFormat::Xltx` | `"xltx"` | `.xltx` | 템플릿 |
| `WorkbookFormat::Xltm` | `"xltm"` | `.xltm` | 매크로 사용 템플릿 |
| `WorkbookFormat::Xlam` | `"xlam"` | `.xlam` | 매크로 사용 추가 기능 |

### 확장자 기반 저장

저장 시 파일 확장자에서 대상 형식이 자동으로 유추됩니다. 인식되는 확장자(`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xlam`)가 있으면 워크북 형식이 쓰기 전에 업데이트됩니다. 인식되지 않는 확장자는 오류를 반환합니다.

**Rust:**

```rust
let mut wb = Workbook::new();
// Format is inferred from ".xlsm" extension
wb.save("output.xlsm")?;
assert_eq!(wb.format(), WorkbookFormat::Xlsm);
```

**TypeScript:**

```typescript
const wb = new Workbook();
await wb.save("output.xlsm"); // Format is automatically set to xlsm
```

`save_to_buffer()` / `writeBufferSync()` 사용 시에는 파일 확장자가 없으므로 저장된 형식이 그대로 사용됩니다. Buffer 저장 전에 `set_format()`으로 형식을 명시적으로 설정하세요.

### VBA 보존

매크로 사용 워크북(`.xlsm`, `.xltm`)에는 VBA 프로젝트 blob(`xl/vbaProject.bin`)이 포함됩니다. 이러한 파일을 열고 다시 저장하면 VBA 프로젝트가 투명하게 보존됩니다. 추가 API 호출이 필요하지 않습니다.

```rust
// Open macro-enabled file, modify data, and save -- VBA macros are preserved
let mut wb = Workbook::open("with_macros.xlsm")?;
wb.set_cell_value("Sheet1", "A1", CellValue::String("Updated".into()))?;
wb.save("with_macros.xlsm")?;
```

```typescript
const wb = await Workbook.open("with_macros.xlsm");
wb.setCellValue("Sheet1", "A1", "Updated");
await wb.save("with_macros.xlsm"); // VBA is preserved
```

---
