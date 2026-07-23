# SheetKit for Node.js

SheetKit provides Node.js bindings for reading and writing Excel `.xlsx` files.
It supports Node.js 24 and is published as an ESM-only package.

## Install

```bash
npm install @sheetkit/node
```

## Quick start

```typescript
import { Workbook } from "@sheetkit/node";

const workbook = new Workbook();
workbook.setCellValue("Sheet1", "A1", "Hello");
await workbook.save("output.xlsx");
```

## Documentation

- [Documentation](https://sheetkit.dev/)
- [1.0 migration guide](https://sheetkit.dev/guide/migration-v1)
- [Repository](https://github.com/Nebu1eto/sheetkit)

SheetKit is licensed under MIT or Apache-2.0.
