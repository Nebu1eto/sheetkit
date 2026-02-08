# SheetKit API Reference

SheetKit is a Rust library for reading and writing Excel (.xlsx) files, with Node.js bindings via napi-rs. This document covers every public API method available in both the Rust crate and the TypeScript/Node.js package.

---

## Table of Contents

- [Workbook I/O](./workbook.md) - Create, open, and save workbooks
- [Cell Operations](./cell.md) - Get and set cell values, cell value types
- [Sheet Management](./sheet.md) - Create, delete, rename, and copy sheets
- [Row and Column Operations](./row-column.md) - Insert, delete, duplicate rows and columns; manage heights, widths, visibility, and outline levels
- [Styles](./style.md) - Font, fill, border, alignment, number format, and protection styles; style builder and deduplication
- [Charts](./chart.md) - Create and manage charts (6+ chart types)
- [Images](./image.md) - Insert and manage images
- [Data Features](./data-features.md) - Merge cells, hyperlinks, data validation, comments, auto-filter, conditional formatting
- [Advanced](./advanced.md) - Freeze/split panes, page layout, defined names, document properties, workbook and sheet protection, formula evaluation, pivot tables, streaming writer, utilities, sparklines, theme colors, and rich text

---

## Quick Navigation by Topic

**Core Operations:**
- [Workbook I/O](./workbook.md)
- [Cell Operations](./cell.md)
- [Sheet Management](./sheet.md)

**Row and Column Management:**
- [Row and Column Operations](./row-column.md)

**Styling and Formatting:**
- [Styles](./style.md)
- [Conditional Formatting](./data-features.md#15-conditional-formatting) (in Data Features)

**Data and Content:**
- [Cell Operations](./cell.md)
- [Data Features](./data-features.md) (merge cells, hyperlinks, validation, comments, filters)
- [Charts](./chart.md)
- [Images](./image.md)

**Advanced Features:**
- [Freeze/Split Panes](./advanced.md#16-freezesplit-panes)
- [Page Layout](./advanced.md#17-page-layout)
- [Defined Names](./advanced.md#18-defined-names)
- [Document Properties](./advanced.md#19-document-properties)
- [Workbook Protection](./advanced.md#20-workbook-protection)
- [Sheet Protection](./advanced.md#21-sheet-protection)
- [Formula Evaluation](./advanced.md#22-formula-evaluation)
- [Pivot Tables](./advanced.md#23-pivot-tables)
- [StreamWriter](./advanced.md#24-streamwriter)
- [Utility Functions](./advanced.md#25-utility-functions)
- [Sparklines](./advanced.md#26-sparklines)
- [Theme Colors](./advanced.md#27-theme-colors)
- [Rich Text](./advanced.md#28-rich-text)

---

## API Overview

Every section in this reference includes code examples for both **Rust** and **TypeScript/Node.js**. Follow the tabs or code block headers to find the implementation for your language.

For a gentler introduction, see the [User Guide](../guide/index.md).
