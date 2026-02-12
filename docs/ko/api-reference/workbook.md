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
// ".xlsm" 확장자에서 형식이 유추됩니다
wb.save("output.xlsm")?;
assert_eq!(wb.format(), WorkbookFormat::Xlsm);
```

**TypeScript:**

```typescript
const wb = new Workbook();
await wb.save("output.xlsm"); // 형식이 자동으로 xlsm으로 설정됩니다
```

`save_to_buffer()` / `writeBufferSync()` 사용 시에는 파일 확장자가 없으므로 저장된 형식이 그대로 사용됩니다. Buffer 저장 전에 `set_format()`으로 형식을 명시적으로 설정하세요.

### VBA 보존

매크로 사용 워크북(`.xlsm`, `.xltm`)에는 VBA 프로젝트 blob(`xl/vbaProject.bin`)이 포함됩니다. 이러한 파일을 열고 다시 저장하면 VBA 프로젝트가 투명하게 보존됩니다. 추가 API 호출이 필요하지 않습니다.

```rust
// 매크로 사용 파일을 열고 데이터를 수정한 후 저장하면 VBA 매크로가 보존됩니다
let mut wb = Workbook::open("with_macros.xlsm")?;
wb.set_cell_value("Sheet1", "A1", CellValue::String("Updated".into()))?;
wb.save("with_macros.xlsm")?;
```

```typescript
const wb = await Workbook.open("with_macros.xlsm");
wb.setCellValue("Sheet1", "A1", "Updated");
await wb.save("with_macros.xlsm"); // VBA가 보존됩니다
```

### `OpenOptions`

워크북을 열 때 파싱 방식을 제어하는 옵션입니다. 모든 필드는 선택 사항입니다.

| 필드 | Rust 타입 | TypeScript 타입 | 기본값 | 설명 |
|---|---|---|---|---|
| `read_mode` / `readMode` | `Option<ReadMode>` | `'lazy' \| 'eager' \| 'stream'?` | `'lazy'` | open 시 파싱 범위를 제어합니다. |
| `aux_parts` / `auxParts` | `Option<AuxParts>` | `'deferred' \| 'eager'?` | `'deferred'` | 보조 파트(comments, charts, images)의 파싱 시점을 제어합니다. |
| `sheet_rows` / `sheetRows` | `Option<u32>` | `number?` | 무제한 | 시트당 읽을 최대 행 수입니다. 초과 행은 무시됩니다. |
| `sheets` | `Option<Vec<String>>` | `string[]?` | 전체 | 이 목록에 포함된 시트만 파싱합니다. 선택되지 않은 시트는 워크북에 존재하지만 데이터가 없습니다. |
| `max_unzip_size` / `maxUnzipSize` | `Option<u64>` | `number?` | 무제한 | ZIP 아카이브의 전체 압축 해제 크기 제한(바이트)입니다. zip bomb을 방지합니다. |
| `max_zip_entries` / `maxZipEntries` | `Option<usize>` | `number?` | 무제한 | ZIP 아카이브의 최대 엔트리 수입니다. zip bomb을 방지합니다. |

#### ReadMode

| 값 | 설명 |
|------|------|
| `'lazy'` | ZIP 인덱스와 메타데이터만 파싱합니다. 시트 XML은 첫 접근 시 파싱됩니다. Node.js 기본값입니다. |
| `'eager'` | open 시 모든 시트와 보조 파트를 파싱합니다. 이전 버전과 동일한 동작입니다. |
| `'stream'` | 최소한의 파싱만 수행합니다. `openSheetReader()`와 함께 사용하여 순방향 메모리 제한 반복에 사용합니다. |

#### AuxParts

| 값 | 설명 |
|------|------|
| `'deferred'` | 보조 파트는 첫 접근 시 로드됩니다. 기본값입니다. |
| `'eager'` | open 시 모든 보조 파트를 파싱합니다. |

### `Workbook::open_with_options(path, options)` / `Workbook.open(path, options?)`

커스텀 파싱 옵션으로 `.xlsx` 파일을 엽니다. 옵션을 생략하면 `Workbook::open`과 동일하게 동작합니다.

**Rust:**

```rust
use sheetkit::{Workbook, OpenOptions};

// "Sales" 시트의 처음 100행만 읽기
let opts = OpenOptions::new()
    .sheet_rows(100)
    .sheets(vec!["Sales".to_string()]);
let wb = Workbook::open_with_options("report.xlsx", &opts)?;
```

**TypeScript:**

```typescript
// "Sales" 시트의 처음 100행만 읽기
const wb = Workbook.openSync("report.xlsx", {
  sheetRows: 100,
  sheets: ["Sales"],
});

// ZIP 안전 제한 설정
const wb2 = await Workbook.open("untrusted.xlsx", {
  maxUnzipSize: 500_000_000,  // 500 MB
  maxZipEntries: 5000,
});
```

### `Workbook::open_from_buffer_with_options(data, options)` / `Workbook.openBufferSync(data, options?)`

메모리 내 버퍼에서 커스텀 파싱 옵션으로 워크북을 엽니다.

**Rust:**

```rust
let data = std::fs::read("report.xlsx")?;
let opts = OpenOptions::new().sheet_rows(50);
let wb = Workbook::open_from_buffer_with_options(&data, &opts)?;
```

**TypeScript:**

```typescript
const data = fs.readFileSync("report.xlsx");
const wb = Workbook.openBufferSync(data, { sheetRows: 50 });
```

> Node.js에서 옵션 매개변수는 모든 open 메서드에서 선택 사항입니다. 생략하면 기존 동작과 동일합니다.

### `wb.openSheetReader(sheet, opts?)` (TypeScript 전용)

지정한 시트에 대한 순방향 전용 streaming reader를 엽니다. 전체 시트를 메모리에 로드하지 않고 배치 단위로 행을 읽습니다. `readMode: 'stream'`과 함께 사용하는 것이 가장 좋습니다.

```typescript
const wb = await Workbook.open("large.xlsx", { readMode: "stream" });
const reader = await wb.openSheetReader("Sheet1", { batchSize: 500 });

// 비동기 반복자: JsRowData[] 배치를 yield합니다
for await (const batch of reader) {
  for (const row of batch) {
    console.log(row);
  }
}
```

**옵션:**

| 필드 | 타입 | 기본값 | 설명 |
|---|---|---|---|
| `batchSize` | `number?` | `1000` | 배치당 행 수입니다. |

**반환값:** `Promise<SheetStreamReader>`

### `SheetStreamReader`

워크시트 데이터를 위한 순방향 전용 streaming reader입니다. `Workbook.openSheetReader()`를 통해 생성됩니다.

**메서드:**

| 메서드 | 반환 타입 | 설명 |
|---|---|---|
| `next(batchSize?)` | `Promise<JsRowData[] \| null>` | 다음 배치를 읽습니다. 완료되면 `null`을 반환합니다. |
| `close()` | `Promise<void>` | 리소스를 해제합니다. `for await`에서 자동으로 호출됩니다. |
| `[Symbol.asyncIterator]()` | `AsyncGenerator<JsRowData[]>` | 배치를 yield하는 비동기 반복자입니다. |

### `wb.getRowsBufferV2(sheet)` (TypeScript 전용)

시트의 셀 데이터를 인라인 문자열이 포함된 v2 바이너리 buffer로 직렬화합니다. v1 형식(`getRowsBuffer`)과 달리, v2 형식은 전역 문자열 테이블을 제거하여 점진적인 행 단위 디코딩을 가능하게 합니다.

```typescript
const bufV2 = wb.getRowsBufferV2("Sheet1");
```

**반환값:** `Buffer`

---
