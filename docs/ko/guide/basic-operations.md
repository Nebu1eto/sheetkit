## 설치

### Rust

`Cargo.toml`에 `sheetkit`을 추가합니다:

```toml
[dependencies]
sheetkit = "0.1"
```

### Node.js

```bash
npm install @sheetkit/node
```

> Node.js 패키지는 napi-rs로 빌드된 네이티브 애드온입니다. 설치 시 네이티브 모듈을 컴파일하기 위해 Rust 빌드 도구(rustc, cargo)가 필요합니다.

---

## 빠른 시작

### Rust

```rust
use sheetkit::{CellValue, Workbook};

fn main() -> sheetkit::Result<()> {
    // Create new workbook (includes "Sheet1" by default)
    let mut wb = Workbook::new();

    // Write cell values
    wb.set_cell_value("Sheet1", "A1", CellValue::String("Name".into()))?;
    wb.set_cell_value("Sheet1", "B1", CellValue::String("Age".into()))?;
    wb.set_cell_value("Sheet1", "A2", CellValue::String("John Doe".into()))?;
    wb.set_cell_value("Sheet1", "B2", CellValue::Number(30.0))?;

    // Read cell value
    let val = wb.get_cell_value("Sheet1", "A1")?;
    println!("A1 = {:?}", val);

    // Save to file
    wb.save("output.xlsx")?;

    // Open existing file
    let wb2 = Workbook::open("output.xlsx")?;
    println!("Sheets: {:?}", wb2.sheet_names());

    Ok(())
}
```

### TypeScript / Node.js

```typescript
import { Workbook } from '@sheetkit/node';

// Create new workbook (includes "Sheet1" by default)
const wb = new Workbook();

// Write cell values
wb.setCellValue('Sheet1', 'A1', 'Name');
wb.setCellValue('Sheet1', 'B1', 'Age');
wb.setCellValue('Sheet1', 'A2', 'John Doe');
wb.setCellValue('Sheet1', 'B2', 30);

// Read cell value
const val = wb.getCellValue('Sheet1', 'A1');
console.log('A1 =', val);

// Save to file
await wb.save('output.xlsx');

// Open existing file
const wb2 = await Workbook.open('output.xlsx');
console.log('Sheets:', wb2.sheetNames);
```

---

## API 레퍼런스

### 워크북 I/O

워크북을 생성, 열기, 저장하는 기본 기능입니다.

#### Rust

```rust
use sheetkit::Workbook;

// Create empty workbook with "Sheet1"
let mut wb = Workbook::new();

// Open existing .xlsx file
let wb = Workbook::open("input.xlsx")?;

// Save as .xlsx file
wb.save("output.xlsx")?;

// Get all sheet names
let names: Vec<&str> = wb.sheet_names();
```

#### TypeScript

```typescript
import { Workbook } from '@sheetkit/node';

// Create empty workbook with "Sheet1"
const wb = new Workbook();

// Open existing .xlsx file
const wb2 = await Workbook.open('input.xlsx');

// Save as .xlsx file
await wb.save('output.xlsx');

// Get all sheet names
const names: string[] = wb.sheetNames;
```

#### Buffer I/O

파일 시스템을 거치지 않고 메모리 내 버퍼로 읽고 쓸 수 있습니다.

**Rust:**

```rust
// Save to buffer
let buf: Vec<u8> = wb.save_to_buffer()?;

// Open from buffer
let wb2 = Workbook::open_from_buffer(&buf)?;
```

**TypeScript:**

```typescript
// Save to buffer
const buf: Buffer = wb.writeBufferSync();

// Open from buffer
const wb2 = Workbook.openBufferSync(buf);

// Async version
const buf2: Buffer = await wb.writeBuffer();
const wb3 = await Workbook.openBuffer(buf2);
```

---

### 워크북 형식 및 VBA 보존

SheetKit은 표준 `.xlsx` 외에도 다양한 Excel 파일 형식을 지원합니다. 파일을 열 때 패키지 콘텐츠 타입에서 형식이 자동으로 감지되며, 저장 시 파일 확장자에서 형식이 유추됩니다.

#### 지원 형식

| 확장자 | 설명 |
|--------|------|
| `.xlsx` | 표준 스프레드시트 (기본값) |
| `.xlsm` | 매크로 사용 스프레드시트 |
| `.xltx` | 템플릿 |
| `.xltm` | 매크로 사용 템플릿 |
| `.xlam` | 매크로 사용 추가 기능 |

#### Rust

```rust
use sheetkit::{Workbook, WorkbookFormat};

// Format is auto-detected when opening
let wb = Workbook::open("macros.xlsm")?;
assert_eq!(wb.format(), WorkbookFormat::Xlsm);

// Format is inferred from file extension when saving
let mut wb2 = Workbook::new();
wb2.save("template.xltx")?;  // Saved as template format

// Explicit format control
let mut wb3 = Workbook::new();
wb3.set_format(WorkbookFormat::Xlsm);
wb3.save_to_buffer()?;  // Buffer uses xlsm content type
```

#### TypeScript

```typescript
// Format is auto-detected when opening
const wb = await Workbook.open("macros.xlsm");
console.log(wb.format);  // "xlsm"

// Format is inferred from file extension when saving
const wb2 = new Workbook();
await wb2.save("template.xltx");  // Saved as template format

// Explicit format control
const wb3 = new Workbook();
wb3.format = "xlsm";
const buf = wb3.writeBufferSync();  // Buffer uses xlsm content type
```

#### VBA 보존

매크로 사용 파일(`.xlsm`, `.xltm`)은 열기/저장 라운드트립을 통해 VBA 프로젝트를 보존합니다. 별도의 코드가 필요하지 않으며, VBA 바이너리 blob이 메모리에 유지되어 저장 시 자동으로 다시 기록됩니다.

```typescript
const wb = await Workbook.open("with_macros.xlsm");
wb.setCellValue("Sheet1", "A1", "Updated");
await wb.save("with_macros.xlsm");  // Macros are preserved
```

자세한 API 설명은 [API 레퍼런스](../api-reference/workbook.md)를 참조하세요.

---

### 셀 조작

셀 값을 읽고 씁니다. 셀은 시트 이름과 셀 참조(예: `"A1"`, `"B2"`, `"AA100"`)로 식별합니다.

#### CellValue 타입

| Rust 변형               | TypeScript 타입 | 설명                                |
|--------------------------|-----------------|-------------------------------------|
| `CellValue::String(s)`  | `string`        | 텍스트 값                            |
| `CellValue::Number(n)`  | `number`        | 숫자 값 (내부적으로 f64로 저장)       |
| `CellValue::Bool(b)`    | `boolean`       | 불리언 값                            |
| `CellValue::Empty`      | `null`          | 빈 셀 / 값 지우기                    |
| `CellValue::Formula{..}`| --              | 수식 (Rust 전용)                     |
| `CellValue::Error(e)`   | --              | `#DIV/0!` 같은 에러 값 (Rust 전용)   |

#### Rust

```rust
use sheetkit::{CellValue, Workbook};

let mut wb = Workbook::new();

// Set values of various types
wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".into()))?;
wb.set_cell_value("Sheet1", "B1", CellValue::Number(42.0))?;
wb.set_cell_value("Sheet1", "C1", CellValue::Bool(true))?;
wb.set_cell_value("Sheet1", "D1", CellValue::Empty)?;

// Convenient conversion using From trait
wb.set_cell_value("Sheet1", "A2", CellValue::from("Text"))?;
wb.set_cell_value("Sheet1", "B2", CellValue::from(100i32))?;
wb.set_cell_value("Sheet1", "C2", CellValue::from(3.14))?;

// Read cell value
let val = wb.get_cell_value("Sheet1", "A1")?;
match val {
    CellValue::String(s) => println!("String: {}", s),
    CellValue::Number(n) => println!("Number: {}", n),
    CellValue::Bool(b) => println!("Bool: {}", b),
    CellValue::Empty => println!("(empty)"),
    _ => {}
}
```

#### TypeScript

```typescript
// Set values -- type is automatically determined from JavaScript value type
wb.setCellValue('Sheet1', 'A1', 'Hello');       // string
wb.setCellValue('Sheet1', 'B1', 42);            // number
wb.setCellValue('Sheet1', 'C1', true);          // boolean
wb.setCellValue('Sheet1', 'D1', null);          // Clear cell

// Read cell value -- returns string | number | boolean | null
const val = wb.getCellValue('Sheet1', 'A1');
```

---

### 시트 관리