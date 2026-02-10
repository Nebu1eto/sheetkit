# SheetKit API Reference

SheetKit is a high-performance SpreadsheetML library for Rust and TypeScript. This document covers every public API method available in both the Rust crate and the TypeScript package.

---

## Table of Contents

### Core Operations

- **[Workbook I/O](./workbook.md)**
  Create new workbooks, open existing files (.xlsx, .xlsm, .xltx, .xltm, .xlam), save with various options, detect file format, preserve VBA projects, handle encrypted files, and use partial reading options for large files.

- **[Cell Operations](./cell.md)**
  Read and write cell values (string, number, boolean, date, formula, error, empty), batch get/set operations, and understand cell value type conversions between Rust and TypeScript.

- **[Sheet Management](./sheet.md)**
  Create new sheets, delete existing sheets, rename sheets, copy sheets within or across workbooks, reorder sheets, get/set active sheet, and list all sheet names.

- **[Row and Column Operations](./row-column.md)**
  Insert and delete rows/columns, duplicate ranges, set heights and widths, hide/unhide rows and columns, manage outline levels for grouping, and apply styles to entire rows or columns.

### Styling and Formatting

- **[Styles](./style.md)**
  Define and apply cell styles including fonts (bold, italic, size, color), fills (solid, gradient, pattern), borders (style, color, thickness), alignment (horizontal, vertical, rotation, wrap text), number formats (built-in and custom), and cell protection.

### Content and Visualization

- **[Charts](./chart.md)**
  Create and configure 43 chart types (column, bar, line, pie, scatter, area, doughnut, radar, surface, bubble, stock, combo, and more), customize titles, legends, axes, data series, and 3D view options.

- **[Images](./image.md)**
  Insert images in 11 formats (PNG, JPEG, GIF, BMP, ICO, TIFF, SVG, EMF, EMZ, WMF, WMZ), position and size images, manage image anchoring, and retrieve image metadata.

- **[Shapes](./shape.md)**
  Add preset geometric shapes (rectangles, circles, arrows, callouts, stars, and more) with customizable position, size, fill, border, and text content.

- **[Slicers](./slicer.md)**
  Create visual filters for Excel tables, configure slicer appearance, manage slicer items and selection, and control slicer position and size.

- **[Form Controls](./form-control.md)**
  Add interactive form controls: buttons, check boxes, option buttons (radio buttons), spin buttons, scroll bars, group boxes, and labels with cell linking and macro assignment.

### Data Features

- **[Data Features](./data-features.md)**
  Merge cells, create hyperlinks, apply data validation rules (list, number, date, text length, custom), add comments and threaded comments (Excel 2019+), enable auto-filter, apply conditional formatting (17 rule types), create and manage Excel tables, and use data conversion utilities (to/from JSON, CSV, HTML, SVG).

### Advanced Features

- **[Advanced](./advanced.md)**
  Freeze and split panes, configure page layout and print settings, define named ranges, set document properties (author, title, subject, keywords, etc.), protect workbooks and sheets with passwords, evaluate formulas (110+ functions), create pivot tables, use StreamWriter for memory-efficient large file generation, access utility functions, add sparklines, work with theme colors, create rich text with inline formatting, encrypt and decrypt files, configure sheet view options (gridlines, zoom, formulas), and control sheet visibility (visible, hidden, very hidden).

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
- [Data Features](./data-features.md) (merge cells, hyperlinks, validation, comments, filters, tables, data conversion)
- [Charts](./chart.md)
- [Images](./image.md)
- [Form Controls](./form-control.md)

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
- [File Encryption](./advanced.md#29-file-encryption)
- [Sheet View Options](./advanced.md#31-sheet-view-options)
- [Sheet Visibility](./advanced.md#32-sheet-visibility)

---

## API Overview

Every section in this reference includes code examples for both **Rust** and **TypeScript/Node.js**. Follow the tabs or code block headers to find the implementation for your language.

For a gentler introduction, see the [User Guide](../guide/index.md).
