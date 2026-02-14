## Sheet Management

### `new_sheet(name)` / `newSheet(name)`

Create a new empty sheet. Returns the 0-based sheet index.

**Rust:**

```rust
let index: usize = wb.new_sheet("Sales")?;
```

**TypeScript:**

```typescript
const index: number = wb.newSheet("Sales");
```

### `delete_sheet(name)` / `deleteSheet(name)`

Delete a sheet by name. Returns an error if the sheet does not exist or if it is the last remaining sheet (a workbook must always have at least one sheet).

**Rust:**

```rust
wb.delete_sheet("Sheet2")?;
```

**TypeScript:**

```typescript
wb.deleteSheet("Sheet2");
```

### `set_sheet_name(old, new)` / `setSheetName(old, new)`

Rename a sheet. Returns an error if the old name does not exist or the new name is invalid or already taken.

**Rust:**

```rust
wb.set_sheet_name("Sheet1", "Summary")?;
```

**TypeScript:**

```typescript
wb.setSheetName("Sheet1", "Summary");
```

### `copy_sheet(src, dst)` / `copySheet(src, dst)`

Copy a sheet. Creates a new sheet named `dst` with the same content as `src`. Returns the 0-based index of the new sheet.

**Rust:**

```rust
let index: usize = wb.copy_sheet("Sheet1", "Sheet1_Copy")?;
```

**TypeScript:**

```typescript
const index: number = wb.copySheet("Sheet1", "Sheet1_Copy");
```

### `get_sheet_index(name)` / `getSheetIndex(name)`

Get the 0-based index of a sheet by name, or `None`/`null` if not found.

**Rust:**

```rust
let idx: Option<usize> = wb.get_sheet_index("Sales");
```

**TypeScript:**

```typescript
const idx: number | null = wb.getSheetIndex("Sales");
```

### `get_active_sheet()` / `getActiveSheet()`

Get the name of the currently active sheet.

**Rust:**

```rust
let name: &str = wb.get_active_sheet();
```

**TypeScript:**

```typescript
const name: string = wb.getActiveSheet();
```

### `set_active_sheet(name)` / `setActiveSheet(name)`

Set the active sheet by name. Returns an error if the sheet does not exist.

**Rust:**

```rust
wb.set_active_sheet("Sales")?;
```

**TypeScript:**

```typescript
wb.setActiveSheet("Sales");
```

### Sheet Name Rules

Sheet names must:
- Be non-empty
- Be at most 31 characters
- Not contain `: \ / ? * [ ]`
- Not start or end with a single quote (`'`)

---
