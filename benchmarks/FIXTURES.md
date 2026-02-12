# Benchmark Fixture Matrix

This document describes the benchmark fixture files used for performance testing.
All fixtures are generated deterministically by `benchmarks/node/generate-fixtures.ts`
and stored in `benchmarks/node/fixtures/`.

## Fixture Overview

| File | Size (bytes) | Rows | Columns | Sheets | Size Category |
|------|-------------|------|---------|--------|---------------|
| `scale-1k.xlsx` | 76,556 | 1,001 | 10 | 1 | Small |
| `scale-10k.xlsx` | 744,616 | 10,001 | 10 | 1 | Medium |
| `scale-100k.xlsx` | 7,582,314 | 100,001 | 10 | 1 | Large |
| `large-data.xlsx` | 7,545,321 | 50,001 | 20 | 1 | Large |
| `strings.xlsx` | 1,402,734 | 20,001 | 10 | 1 | Medium |
| `multi-sheet.xlsx` | 3,771,540 | 50,010 | 10 | 10 | Large |
| `comments.xlsx` | 176,212 | 2,001 | 5 | 1 | Small |
| `formulas.xlsx` | 577,635 | 10,001 | 7 | 1 | Medium |
| `heavy-styles.xlsx` | 364,058 | 5,001 | 10 | 1 | Medium |
| `data-validation.xlsx` | 289,242 | 5,001 | 8 | 1 | Medium |
| `merged-cells.xlsx` | 22,701 | 500 | 8 | 1 | Small |
| `mixed-workload.xlsx` | 412,710 | ~5,012 | 15 | 3 | Medium |

Row counts include the header row. Size categories: Small (<100K), Medium (100K-5M), Large (>5M).

## Detailed Fixture Descriptions

### scale-1k.xlsx

Scaling series baseline. 1,000 data rows + 1 header row, 10 columns (A-J).

- **Content pattern**: 3-column repeating cycle. Column mod 3 = 0: integer (`row * (col + 1)`). Column mod 3 = 1: string (`R{row}C{col}`). Column mod 3 = 2: float (`(row * col) / 100`).
- **Header row**: `Col_1` through `Col_10`.
- **Cell density**: Dense (every cell populated).
- **Cell types**: Integer, string, float.
- **Features**: None (no styles, comments, formulas, validation, or merged cells).
- **Primary use**: Read/write scaling baseline, overhead measurement.

### scale-10k.xlsx

Scaling series midpoint. Same structure as scale-1k, 10,000 data rows.

- **Content pattern**: Identical to scale-1k.
- **Cell density**: Dense.
- **Cell types**: Integer, string, float.
- **Features**: None.
- **Primary use**: Read/write scaling midpoint, linear scaling validation.

### scale-100k.xlsx

Scaling series upper bound. Same structure as scale-1k, 100,000 data rows.

- **Content pattern**: Identical to scale-1k.
- **Cell density**: Dense.
- **Cell types**: Integer, string, float.
- **Features**: None.
- **Primary use**: Read/write scaling upper bound, large-file behavior, memory pressure testing.

### large-data.xlsx

High-volume mixed-type workload. 50,000 data rows + 1 header row, 20 columns (A-T).

- **Content pattern**: 4-column repeating cycle. Column mod 4 = 0: integer (`row * (col + 1)`). Column mod 4 = 1: string (`Row{r}_Col{c+1}`). Column mod 4 = 2: float (`(row * (col + 1)) / 100`). Column mod 4 = 3: boolean (`row % 2 == 0`).
- **Header row**: `Column_1` through `Column_20`.
- **Cell density**: Dense (1,000,000 data cells).
- **Cell types**: Integer, string, float, boolean.
- **Features**: None.
- **Primary use**: Large read/write throughput, random-access cell lookup, buffer round-trip baseline.

### strings.xlsx

String-heavy SST (Shared String Table) stress test. 20,000 data rows + 1 header row, 10 columns (A-J).

- **Content pattern**: All string values derived from 20-word NATO phonetic alphabet list. Columns: Full_Name, Email, Department, Address, Description, City, Country, Phone, Title, Bio. Each row generates unique strings via deterministic word selection (`words[(r * N) % 20]`).
- **Header row**: `Full_Name`, `Email`, `Department`, `Address`, `Description`, `City`, `Country`, `Phone`, `Title`, `Bio`.
- **Cell density**: Dense (200,000 string cells).
- **Cell types**: String only.
- **Features**: None. SST is the stress target.
- **Primary use**: Shared string table performance, string deserialization throughput, SST deduplication efficiency.

### multi-sheet.xlsx

Multi-sheet workload. 10 sheets, each with 5,000 data rows + 1 header row, 10 columns (A-J).

- **Sheets**: `Sheet1` through `Sheet10`.
- **Content pattern**: Alternating number/string per column. Even columns: integer (`row * (col + 1) + sheet * 1000`). Odd columns: string (`S{sheet+1}_R{row}_C{col+1}`).
- **Header row**: `Header_1` through `Header_10` (identical across all sheets).
- **Cell density**: Dense (500,000 total cells across all sheets).
- **Cell types**: Integer, string.
- **Features**: None.
- **Primary use**: Multi-sheet iteration overhead, per-sheet memory allocation, sheet switching cost.

### comments.xlsx

Comment-heavy workbook. 2,000 data rows + 1 header row, 5 columns (A-E).

- **Content pattern**: ID (integer), Name (string), Score (integer 0-99), Status ("Flagged" or "OK"), Review_Notes (string).
- **Comment placement**: Every row has a comment on the Score cell (column C, 2,000 comments). Rows where `r % 3 == 0` also have a comment on the Status cell (column D, ~667 comments). Total: ~2,667 comments.
- **Comment structure**: Author field + text body. Authors: "Reviewer" (score comments), "Manager" (status comments).
- **Cell density**: Dense.
- **Cell types**: Integer, string.
- **Features**: Comments (2,667 total).
- **Primary use**: Comment read/write performance, comment XML parsing overhead.

### formulas.xlsx

Formula-heavy workbook. 10,000 data rows + 1 header row, 7 columns (A-G).

- **Content pattern**: 2 data columns (A: float `r * 1.5`, B: float `(r % 100) + 0.5`) and 5 formula columns (C: `A+B`, D: `A*B`, E: `AVERAGE(A,B)`, F: `MAX(A,B)`, G: `IF(A>B,"A","B")`).
- **Header row**: `Value_A`, `Value_B`, `Sum`, `Product`, `Average`, `Max`, `Condition`.
- **Cell density**: Dense (70,000 cells: 20,000 data + 50,000 formula).
- **Cell types**: Float (data columns), formula (5 columns).
- **Formula types**: Arithmetic (`+`, `*`), statistical (`AVERAGE`, `MAX`), logical (`IF`).
- **Features**: Formulas (50,000 formula cells).
- **Primary use**: Formula storage/retrieval performance, formula XML serialization cost.

### heavy-styles.xlsx

Style-heavy workbook. 5,000 data rows + 1 header row, 10 columns (A-J).

- **Content pattern**: 5-column repeating cycle. Column mod 5 = 0: integer (ID). Column mod 5 = 1: string (name). Column mod 5 = 2: float (amount). Column mod 5 = 3: float 0.00-0.99 (rate). Column mod 5 = 4: string (notes).
- **Header row**: `ID`, `Name`, `Amount`, `Rate`, `Date`, `Category`, `Score`, `Percent`, `Notes`, `Status`.
- **Style definitions** (5 styles, cycled per column):
  - Bold: font bold + size 12 + Arial, solid fill (blue), thin borders, center alignment.
  - Number: numFmtId 4 (#,##0.00), Calibri 11, hair bottom border.
  - Percent: numFmtId 10 (0.00%), Calibri 11 italic, solid fill (yellow).
  - Date: custom format `yyyy-mm-dd`, Calibri 11.
  - Wrap: wrapText alignment, Calibri 10, solid fill (gray).
- **Cell density**: Dense (50,000 styled cells + 10 styled header cells).
- **Cell types**: Integer, string, float.
- **Features**: Styles (5 distinct style definitions applied to every cell).
- **Primary use**: Style XML parsing/serialization cost, style deduplication performance, styles.xml size impact.

### data-validation.xlsx

Data validation workbook. 5,000 data rows + 1 header row, 8 columns (A-H).

- **Content pattern**: Status (string from list), Priority (string from list), Score (integer 0-100), Rate (float 0.00-0.99), Quantity (integer), Department (string), Code (string `X{padded_number}`), Notes (string).
- **Header row**: `Status`, `Priority`, `Score`, `Rate`, `Quantity`, `Department`, `Code`, `Notes`.
- **Validation rules** (8 rules, one per column):
  - A: List validation (`Active,Inactive,Pending,Closed,Archived`), error message.
  - B: List validation (`Critical,High,Medium,Low`).
  - C: Whole number, between 0-100, error message.
  - D: Decimal, between 0-1, input prompt.
  - E: Whole number, >= 0.
  - F: Text length, <= 30 characters, error message.
  - G: Custom formula (`ISNUMBER(FIND(LEFT(G2,1),"ABC...Z"))`).
  - H: Text length, between 0-500.
- **Cell density**: Dense.
- **Cell types**: String, integer, float.
- **Features**: Data validation (8 rules spanning 5,000 rows each).
- **Primary use**: Validation rule read/write overhead, worksheet extensions XML cost.

### merged-cells.xlsx

Merged-cell workbook. 500 total rows, 8 columns (A-H).

- **Content pattern**: 100 sections, each containing:
  - 1 title row: string value, merged A:H (1 merge region per section).
  - 1 sub-header row: 4 string values, each merged across 2 columns (A:B, C:D, E:F, G:H = 4 merge regions per section).
  - 3 data rows: 8 integer values per row.
- **Total merge regions**: 500 (100 title merges + 400 sub-header merges).
- **Cell density**: Mixed. Title and sub-header rows are sparse (merged). Data rows are dense.
- **Cell types**: String (titles, sub-headers), integer (data).
- **Features**: Merged cells (500 merge regions).
- **Primary use**: Merge region read/write performance, sparse row handling.

### mixed-workload.xlsx

Realistic ERP-style document. 3 sheets with mixed features.

- **Sheet 1 ("Sheet1" -- Invoice list)**: 3,000 data rows + 2 header rows, 8 columns (A-H).
  - Content: Invoice_ID (string), Customer (string from 10-name list), Amount (float), Tax (formula `C*0.1`), Total (formula `C+D`), Status (string from 4-value list), Due_Date (serial number), Notes (string, every 5th row).
  - Styles: Bold header with fill/border/alignment. Currency format on Amount/Tax/Total columns.
  - Validation: List validation on Status column.
  - Comments: On every overdue invoice's Status cell (~750 comments).
  - Merged cells: Title row merged A1:H1.

- **Sheet 2 ("Employees" -- Employee directory)**: 2,000 data rows + 1 header row, 15 columns (A-O).
  - Content: EmpID, First_Name, Last_Name, Email, Department, Title, Salary, Bonus_Rate, Total_Comp (formula `G*(1+H)`), Start_Date, Manager, Location, Phone, Status, Notes.
  - Styles: Bold header, currency format on Salary/Total_Comp, percent format on Bonus_Rate.
  - Validation: List on Department, decimal range on Salary (30k-500k), decimal range on Bonus_Rate (0-0.5).

- **Sheet 3 ("Summary" -- Dashboard)**: 9 rows, 4 columns (A-D).
  - Content: 7 metrics with cross-sheet formulas (COUNTA, SUM, AVERAGE, MAX referencing Sheet1 and Employees).
  - Merged title row A1:D1.

- **Total features**: 3 sheets, ~5,012 rows, styles (3 definitions), formulas (~5,007 formula cells), validation (4 rules), comments (~750), merged cells (2 regions).
- **Primary use**: Realistic mixed-feature workload, combined feature overhead measurement.

## Fixture Categories

### Scaling Series

Fixtures with identical column structure at different row counts, for measuring
how performance scales with data volume.

| Fixture | Data Rows | Total Cells |
|---------|-----------|-------------|
| `scale-1k.xlsx` | 1,000 | 10,000 |
| `scale-10k.xlsx` | 10,000 | 100,000 |
| `scale-100k.xlsx` | 100,000 | 1,000,000 |

All three use the same 10-column structure with 3-type repeating pattern (integer, string, float).
No features beyond raw data. Suitable for linear scaling analysis.

### Feature-Specific

Each fixture isolates a single Excel feature for targeted benchmarking.

| Fixture | Target Feature | Feature Volume |
|---------|---------------|----------------|
| `comments.xlsx` | Comments | ~2,667 comments |
| `formulas.xlsx` | Formulas | 50,000 formula cells |
| `heavy-styles.xlsx` | Styles | 5 style definitions, 50,000 styled cells |
| `data-validation.xlsx` | Data validation | 8 validation rules x 5,000 rows |
| `merged-cells.xlsx` | Merged cells | 500 merge regions |

### Workload Profiles

Fixtures representing realistic usage patterns.

| Fixture | Profile | Key Characteristic |
|---------|---------|-------------------|
| `large-data.xlsx` | High-volume data dump | 1M cells, 4 data types, no features |
| `strings.xlsx` | Text-heavy export | 200K string cells, SST stress |
| `multi-sheet.xlsx` | Multi-tab report | 10 sheets, 500K total cells |
| `mixed-workload.xlsx` | ERP document | All features combined across 3 sheets |

## Content Type Matrix

| Fixture | Integer | Float | String | Boolean | Formula | Sparse |
|---------|---------|-------|--------|---------|---------|--------|
| `scale-1k.xlsx` | Yes | Yes | Yes | No | No | No |
| `scale-10k.xlsx` | Yes | Yes | Yes | No | No | No |
| `scale-100k.xlsx` | Yes | Yes | Yes | No | No | No |
| `large-data.xlsx` | Yes | Yes | Yes | Yes | No | No |
| `strings.xlsx` | No | No | Yes | No | No | No |
| `multi-sheet.xlsx` | Yes | No | Yes | No | No | No |
| `comments.xlsx` | Yes | No | Yes | No | No | No |
| `formulas.xlsx` | No | Yes | No | No | Yes | No |
| `heavy-styles.xlsx` | Yes | Yes | Yes | No | No | No |
| `data-validation.xlsx` | Yes | Yes | Yes | No | No | No |
| `merged-cells.xlsx` | Yes | No | Yes | No | No | Yes |
| `mixed-workload.xlsx` | Yes | Yes | Yes | No | Yes | Partial |

## Feature Presence Matrix

| Fixture | Styles | Comments | Formulas | Validation | Merged Cells | Multi-Sheet |
|---------|--------|----------|----------|------------|-------------|-------------|
| `scale-1k.xlsx` | - | - | - | - | - | - |
| `scale-10k.xlsx` | - | - | - | - | - | - |
| `scale-100k.xlsx` | - | - | - | - | - | - |
| `large-data.xlsx` | - | - | - | - | - | - |
| `strings.xlsx` | - | - | - | - | - | - |
| `multi-sheet.xlsx` | - | - | - | - | - | Yes (10) |
| `comments.xlsx` | - | Yes (2,667) | - | - | - | - |
| `formulas.xlsx` | - | - | Yes (50,000) | - | - | - |
| `heavy-styles.xlsx` | Yes (5 defs) | - | - | - | - | - |
| `data-validation.xlsx` | - | - | - | Yes (8 rules) | - | - |
| `merged-cells.xlsx` | - | - | - | - | Yes (500) | - |
| `mixed-workload.xlsx` | Yes (3 defs) | Yes (~750) | Yes (~5,007) | Yes (4 rules) | Yes (2) | Yes (3) |

## Benchmark Scenario Mapping

Which fixtures are used in which benchmark scenarios (from `benchmark.ts`):

| Benchmark Scenario | Fixture Used |
|-------------------|-------------|
| Read Large Data | `large-data.xlsx` |
| Read Heavy Styles | `heavy-styles.xlsx` |
| Read Multi-Sheet | `multi-sheet.xlsx` |
| Read Formulas | `formulas.xlsx` |
| Read Strings | `strings.xlsx` |
| Read Data Validation | `data-validation.xlsx` |
| Read Comments | `comments.xlsx` |
| Read Merged Cells | `merged-cells.xlsx` |
| Read Mixed Workload | `mixed-workload.xlsx` |
| Read Scale 1k/10k/100k | `scale-1k.xlsx`, `scale-10k.xlsx`, `scale-100k.xlsx` |
| Random-Access Read | `large-data.xlsx` |
| Write benchmarks | Generated in-memory (no fixture file read) |
| Buffer Round-Trip | Generated in-memory |
| Streaming Write | Generated in-memory |
| Mixed Workload Write | Generated in-memory |

## Fixture Generation

### How to regenerate

```bash
cd benchmarks/node
pnpm generate
```

This runs `tsx generate-fixtures.ts`, which uses `@sheetkit/node` (SheetKit's own Node.js bindings)
to create all 12 fixture files. The generator requires a working SheetKit build.

### Prerequisites

1. SheetKit native module must be built: `cd packages/sheetkit && pnpm build`
2. Benchmark dependencies must be installed: `cd benchmarks/node && pnpm install`

### Determinism

All fixtures are generated deterministically. The same code always produces files with
identical cell content. However, file-level byte equality is not guaranteed across
regenerations because:

- ZIP compression may produce different byte sequences across library versions.
- XML serialization order within the ZIP archive may vary.
- Timestamps in ZIP entry metadata may differ.

For benchmark reproducibility, the fixture content (cell values, formulas, styles,
comments, validation rules, merged regions) is stable. The on-disk file size may
fluctuate slightly between regenerations but should remain within a few percent of
the documented sizes.

### Adding new fixtures

1. Add a generator function to `generate-fixtures.ts`.
2. Call it from the main block at the bottom of the file.
3. Add the fixture to `.gitignore` if fixtures are not checked in, or commit it.
4. Update this document with the new fixture's metadata.
5. Add read/write benchmark scenarios in `benchmark.ts` as needed.
