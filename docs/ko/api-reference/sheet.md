## 시트 관리

시트의 생성, 삭제, 이름 변경, 복사 등 시트 단위 조작을 다룹니다.

### `new_sheet(name)` / `newSheet(name)`

빈 시트를 추가합니다. 0부터 시작하는 시트 인덱스를 반환합니다.

**Rust:**

```rust
let index = wb.new_sheet("Data")?;
```

**TypeScript:**

```typescript
const index: number = wb.newSheet("Data");
```

### `delete_sheet(name)` / `deleteSheet(name)`

시트를 삭제합니다. 마지막 시트는 삭제할 수 없습니다.

**Rust:**

```rust
wb.delete_sheet("Data")?;
```

**TypeScript:**

```typescript
wb.deleteSheet("Data");
```

### `set_sheet_name(old, new)` / `setSheetName(old, new)`

시트 이름을 변경합니다.

**Rust:**

```rust
wb.set_sheet_name("Sheet1", "Summary")?;
```

**TypeScript:**

```typescript
wb.setSheetName("Sheet1", "Summary");
```

### `copy_sheet(source, target)` / `copySheet(source, target)`

기존 시트를 복사하여 새 시트를 생성합니다. 새 시트의 0부터 시작하는 인덱스를 반환합니다.

**Rust:**

```rust
let index = wb.copy_sheet("Sheet1", "Sheet1_Copy")?;
```

**TypeScript:**

```typescript
const index: number = wb.copySheet("Sheet1", "Sheet1_Copy");
```

### `get_sheet_index(name)` / `getSheetIndex(name)`

시트의 0부터 시작하는 인덱스를 반환합니다. 존재하지 않으면 None / null을 반환합니다.

**Rust:**

```rust
let index: Option<usize> = wb.get_sheet_index("Sheet1");
```

**TypeScript:**

```typescript
const index: number | null = wb.getSheetIndex("Sheet1");
```

### `get_active_sheet()` / `getActiveSheet()`

현재 활성 시트의 이름을 반환합니다.

**Rust:**

```rust
let name: &str = wb.get_active_sheet();
```

**TypeScript:**

```typescript
const name: string = wb.getActiveSheet();
```

### `set_active_sheet(name)` / `setActiveSheet(name)`

활성 시트를 변경합니다.

**Rust:**

```rust
wb.set_active_sheet("Data")?;
```

**TypeScript:**

```typescript
wb.setActiveSheet("Data");
```

---
