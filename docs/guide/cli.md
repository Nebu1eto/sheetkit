# CLI Tool

SheetKit includes a command-line tool for working with Excel (.xlsx) files directly from the terminal. The CLI provides quick access to common workbook operations without writing code.

## Installation

Build from source with the `cli` feature:

```bash
cargo install sheetkit --features cli
```

Or build locally:

```bash
cargo build --release --package sheetkit --features cli
```

The binary will be at `target/release/sheetkit`.

## Commands

### info

Show workbook information including sheet names, active sheet, and document properties.

```bash
sheetkit info report.xlsx
```

Example output:

```
File: report.xlsx
Sheets: 3
  1: Summary (active)
  2: Data
  3: Charts
Creator: John Smith
Modified: 2025-01-15T10:30:00Z
```

### sheets

List all sheet names in the workbook, one per line.

```bash
sheetkit sheets report.xlsx
```

Example output:

```
Summary
Data
Charts
```

This is useful for scripting:

```bash
for sheet in $(sheetkit sheets report.xlsx); do
  sheetkit convert report.xlsx -f csv -o "${sheet}.csv" --sheet "$sheet"
done
```

### read

Read and display sheet data. Defaults to the active sheet with tab-separated output.

```bash
# Read the active sheet as a table
sheetkit read report.xlsx

# Read a specific sheet
sheetkit read report.xlsx --sheet Data

# Output as CSV
sheetkit read report.xlsx --format csv
```

Options:

| Flag | Short | Description |
|------|-------|-------------|
| `--sheet <name>` | `-s` | Sheet to read (default: active sheet) |
| `--format <fmt>` | `-f` | Output format: `table` (default) or `csv` |

### get

Get a single cell value.

```bash
sheetkit get report.xlsx Sheet1 A1
```

Returns the cell value to stdout. Empty cells produce no output. Numeric values are displayed without trailing zeros for integers. Boolean values are displayed as `TRUE` or `FALSE`.

### set

Set a cell value and save to a new file.

```bash
sheetkit set report.xlsx Sheet1 A1 "New Title" -o updated.xlsx
```

The value is automatically interpreted:
- `TRUE` / `FALSE` (case-insensitive) are stored as booleans.
- Valid numbers are stored as numeric values.
- Everything else is stored as a string.

Options:

| Flag | Short | Description |
|------|-------|-------------|
| `--output <path>` | `-o` | Output file path (required) |

### convert

Convert a sheet to another format.

```bash
# Convert active sheet to CSV
sheetkit convert report.xlsx -f csv -o output.csv

# Convert a specific sheet
sheetkit convert report.xlsx -f csv -o data.csv --sheet Data
```

Options:

| Flag | Short | Description |
|------|-------|-------------|
| `--format <fmt>` | `-f` | Target format: `csv` (required) |
| `--output <path>` | `-o` | Output file path (required) |
| `--sheet <name>` | `-s` | Sheet to convert (default: active sheet) |

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Error (invalid file, missing sheet, bad cell reference, etc.) |

Error messages are printed to stderr.

## Examples

```bash
# Quick inspection of an unknown spreadsheet
sheetkit info data.xlsx
sheetkit sheets data.xlsx
sheetkit read data.xlsx | head -5

# Extract a specific value for use in a script
total=$(sheetkit get data.xlsx Summary B10)
echo "Total: $total"

# Batch update a cell and export
sheetkit set template.xlsx Report A1 "Q4 2025" -o report.xlsx
sheetkit convert report.xlsx -f csv -o report.csv

# Export all sheets to CSV
for sheet in $(sheetkit sheets workbook.xlsx); do
  sheetkit convert workbook.xlsx -f csv -o "${sheet}.csv" --sheet "$sheet"
done
```
