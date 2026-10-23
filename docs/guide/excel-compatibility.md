# Excel Compatibility

SheetKit aims for Excel parity: a workbook that SheetKit writes or edits should open in Microsoft Excel without repair, and a calculation or structural edit should produce what Excel would produce for the same input. SheetKit 1.0 tightened several behaviors for security and correctness. Some of those changes break compatibility with earlier SheetKit releases. They are acceptable only where they match Excel, so this page records each changed behavior next to what Excel did when given the same input.

## How this was verified

Each behavior was run twice on the same input: once through the public SheetKit API (Rust crate, current 1.0 branch), and once in Microsoft Excel 16.114 for Mac, driven through AppleScript on 2026-10-23. Formulas were entered into a scratch workbook, and the value or error Excel calculated was read back. Structural edits were made with Excel's own `delete` and `insert` commands on a file SheetKit generated, and Excel then saved the result. The saved XML was compared with SheetKit's output for the same edit. Table names and totals rows were set through Excel's list object model.

Status values:

- **MATCH**: SheetKit and Excel produced the same result.
- **DIFFERS**: the results differ. The Notes column says whether the difference is deliberate, a security limit, or a defect.
- **NOT VERIFIABLE**: Excel could not be driven to the same input through AppleScript. The reason is given.
- **N/A**: the behavior has no Excel-visible effect, or applies only to the Node.js API.

The AppleScript object model is not the Excel user interface. Where the two can disagree (table naming, for example), the table says which one was tested.

Node.js methods follow the Rust names in camelCase (`remove_row` is `removeRow`, `add_table` is `addTable`). Formula results are the same in both bindings.

## Aggregate functions

Test data: `A1` empty, `A2` = `TRUE`, `A3` = `FALSE`, `A4` = text `"3"`, `A5` = `5`. `B1:B3` empty.

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| `AVERAGE`, `SUM`, `MIN`, `MAX`, `COUNT` over `A1:A5` | `calculate_all` / `evaluate_formula`: 5, 5, 5, 5, 1 | 5, 5, 5, 5, 1 | MATCH | Booleans, text, and empty cells inside a range argument are ignored. |
| `AVERAGE(B1:B3)` (all empty) and `AVERAGE(A2:A4)` (only booleans and text) | `#DIV/0!` | `#DIV/0!` | MATCH | |
| `MIN`, `MAX`, `SUM` over `B1:B3` | 0 | 0 | MATCH | SheetKit returns `SUM` of an empty range as negative zero (`-0.0`), which compares equal to 0. |
| Direct scalar arguments: `AVERAGE(TRUE)`, `SUM("3")`, `SUM(TRUE)` | 1, 3, 1 | 1, 3, 1 | MATCH | Booleans and numeric text typed directly as arguments are counted. |
| `AVERAGE("x")` | `#VALUE!` from `calculate_all` | `#VALUE!` | MATCH | `evaluate_formula` returns an `Err` ("cannot convert") instead of an error value. |
| `COUNT(TRUE,"3","x")`, `COUNT("3")`, `COUNT(TRUE)` | 2, 1, 1 | 2, 1, 1 | MATCH | Booleans and numeric text typed directly as `COUNT` arguments are counted; through references only numbers are counted. |

## Date functions

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| `DATE(10000,1,1)`, `DATE(-1,1,1)` | `#NUM!` | `#NUM!` | MATCH | |
| `DATE(0,1,1)`, `DATE(9999,12,31)` | 1, 2958465 | 1 (1900-01-01), 2958465 | MATCH | |
| `DATE(2024,2147483647,1)` | `#NUM!` | `#NUM!` | MATCH | |
| `DATE(2024,1,day)` with `day` = -32768, -5, 32766, 32767 | 12523, 45286, 78057, 78058 | 12523, 45286, 78057, 78058 | MATCH | |
| `DATE(2024,1,day)` with `day` = 32768, 40000, 65536, -40000 | 78059, 85291, 110827, 5291 | 78058 for every one | DIFFERS | Excel returned serial 78058 (2113-09-17, the result for day 32767) for every day value tested outside -32768..32767. SheetKit applies the day offset arithmetically and does not reproduce this clamping. |
| `DATE(2024,1,day)` with `day` = 2958101, 2147483647, 4294967296, 1E10, -2147483648 | `#NUM!` | 78058 | DIFFERS | Same Excel clamping as above. SheetKit returns `#NUM!` when the result falls outside 1900-01-01..9999-12-31. |
| `EDATE(45000,2147483647)`, `EDATE(45000,-600)` | `#NUM!`, 26738 | `#NUM!`, 26738 | MATCH | |
| `EDATE(0,0)`, `EDATE(-1,0)` | 0, `#NUM!` | 0, `#NUM!` | MATCH | Serial 0 is 1900-01-00; negative serials return `#NUM!`. |
| `EOMONTH(45000,2147483647)`, `EOMONTH(2958465,0)`, `EOMONTH(2958465,1)` | `#NUM!`, 2958465, `#NUM!` | `#NUM!`, 2958465, `#NUM!` | MATCH | |
| `EOMONTH(1,-1)`, `EOMONTH(31,-1)` | `#NUM!`, `#NUM!` | `#NUM!`, `#NUM!` | MATCH | A result before 1900-01-01 is an error. |
| `EOMONTH(0,0)`, `EOMONTH(-1,0)`, `EOMONTH(60,0)` | 31, `#NUM!`, 59 | 31, `#NUM!`, 59 | MATCH | |
| `WORKDAY(45000,1E17)`, `WORKDAY(45000,-1E17)`, `WORKDAY(1,1)` | `#NUM!`, `#NUM!`, 2 | `#NUM!`, `#NUM!`, 2 | MATCH | SheetKit bounds the day count before iterating. |
| `WORKDAY(0,1)`, `WORKDAY(-1,1)` | 2, `#NUM!` | 2, `#NUM!` | MATCH | |

## Text functions

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| Negative counts: `LEFT("abc",-1)`, `RIGHT("abc",-1)`, `MID("abc",1,-1)`, `REPT("a",-1)`, `REPLACE("abc",1,-1,"x")` | `#VALUE!` | `#VALUE!` | MATCH | |
| Start below 1: `MID("abc",0,1)`, `REPLACE("abc",0,1,"x")`, `REPLACE("abc",-1,1,"x")` | `#VALUE!` | `#VALUE!` | MATCH | |
| Huge counts: `LEFT("abc",1E10)`, `LEFT("abc",1E300)`, `RIGHT("abc",2147483648)` | `"abc"` | `"abc"` | MATCH | |
| `MID` past the end: `MID("abc",4,1)`, `MID("abc",32768,1)`, `MID("abc",2147483648,1)` | `""` | `""` | MATCH | |
| `MID("abc",1,2147483647)` | `"abc"` | `"abc"` | MATCH | |
| `MID("abc",1,2147483648)`, `MID("abc",1,1E10)`, `MID("abc",1E10,1)` | `#VALUE!` | `#VALUE!` | MATCH | A start above 2147483648 or a length above 2147483647 is rejected. |
| `REPT` result length: `LEN(REPT("a",32767))`, `REPT("a",32768)`, `REPT("ab",1E10)` | 32767, `#VALUE!`, `#VALUE!` | 32767, `#VALUE!`, `#VALUE!` | MATCH | 32767 characters is Excel's cell text limit. |
| `REPLACE("abc",4,1,"x")` | `"abcx"` | `"abcx"` | MATCH | |
| `REPLACE` with start or length above 32767: start 32768, 2147483647, 1E10; length 2147483647, 2147483648, 1E10 | `#VALUE!` | `#VALUE!` | MATCH | |
| `LEN` of an error argument: `LEN(1/0)`, `LEN(REPT("a",32768))` | `#DIV/0!`, `#VALUE!` | `#DIV/0!`, `#VALUE!` | MATCH | Text functions propagate an error argument. |

## Rounding

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| `ROUND(1,309)`, `ROUNDDOWN(1,-309)`, `ROUND(123.456,-2)` | 1, 0, 100 | 1, 0, 100 | MATCH | |
| `ROUNDUP(123.456,-400)` | `#NUM!` | `#NUM!` | MATCH | |
| `ROUND(1.7E308,-308)` | `#NUM!` | Formula cannot be entered | NOT VERIFIABLE | Excel rejects the literal `1.7E308` (its largest numeric literal is about 9.99999999999999E+307), and AppleScript left the cell empty. SheetKit accepts the literal and returns `#NUM!` because the rounded result overflows. |
| `ROUND(9E307,-307)`, `ROUND(9.9E307,-307)` | 9.000000000000004E+307, 1.0000000000000004E+308 | 9E+307, 1E+308 | DIFFERS | Floating-point error at extreme magnitudes. The relative error is about 4E-16. |

## TEXT number formats

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| `TEXT(1.25,"# ?/8")` | `"1 2/8"` | `"1 2/8"` | MATCH | |
| `TEXT(1,"0E+0")`, `TEXT(1,"ŉ0")` | `"1E+0"`, `"ŉ1"` | `"1E+0"`, `"ŉ1"` | MATCH | |
| `TEXT(1,"ŉE+é")` | `"1E+0"` | `#VALUE!` | DIFFERS | Security limit. SheetKit always produces finite, bounded output for any format code and does not reject every code Excel considers invalid. |
| `TEXT(42949672964294967296,"/?")`, `TEXT(1E15,"/?")` | `"42949672964294970000"`, `"1000000000000000    "` | `#VALUE!`, `#VALUE!` | DIFFERS | Security limit, as above. |
| `TEXT(1.5,"?/18446744073709551615")`, `TEXT(1.5,"?/1000000000")` | `"3/2"`, `"3/2"` | `#VALUE!`, `#VALUE!` | DIFFERS | Security limit. SheetKit falls back to a bounded fraction search for denominators it cannot use. |
| `TEXT(1.5,"?/99999")` | `"149999/99999"` | `"3/9    "` | DIFFERS | Raw Excel output recorded as returned. SheetKit applies the fixed denominator. |
| `TEXT(12345678901234567890,"0")` | `"12345678901234567168"` | `"12345678901234500000"` | DIFFERS | Excel shows 15 significant digits; SheetKit prints the full double value. |

## Cross-sheet evaluation

Test data: `Sheet2!A5` = 1, `Sheet2!B5` = `=A5+ROW()+COLUMN()*100`, `Sheet2!C5` = `=SUM(A5:B5)`, `Sheet1!A5` = 5.

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| `=Sheet2!B5` on Sheet1 | 206 | 206 | MATCH | Unqualified `A5`, `ROW()`, and `COLUMN()` resolve against Sheet2, where the formula lives. |
| `=Sheet2!C5`, `=SUM(Sheet2!A5:B5)` on Sheet1 | 207, 207 | 207, 207 | MATCH | |

## Deleting rows

Rows 5 and 6 of Sheet1 were deleted (`remove_row` twice in SheetKit; `delete entire row` of `A5:A6` in Excel). Formulas on Sheet1 and on a second sheet referenced Sheet1.

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| Single reference to a deleted cell: `=A5`, `=Sheet1!A5` | `=#REF!`, `=Sheet1!#REF!` | `=#REF!`, `=Sheet1!#REF!` | MATCH | |
| Range partly deleted: `SUM(A4:A7)`, `SUM(Sheet1!A6:A9)` | `SUM(A4:A5)`, `SUM(Sheet1!A5:A7)` | same | MATCH | The range shrinks. |
| Range fully deleted: `SUM(A5:A6)`, `SUM(Sheet1!5:5)` | `SUM(#REF!)`, `SUM(Sheet1!#REF!)` | same | MATCH | |
| References below the band: `Sheet1!A10`, `SUM(Sheet1!4:7)` | `Sheet1!A8`, `SUM(Sheet1!4:5)` | same | MATCH | |
| Whole-column reference: `SUM(A:A)` | unchanged | unchanged | MATCH | |
| Merged cells `C5:D6`, `E4:E7`, `F8:G9` | removed, `E4:E5`, `F6:G7` | same | MATCH | |
| Data validations `H5`, `H8`, `H3:H7` (identical rules) | removed, `H6`, `H3:H5` | one rule on `H3:H6` | MATCH | The same cells are covered. Excel merges identical rules into one `sqref` when it saves. |
| Conditional formats `I5:I6`, `I4:I8` | removed, `I4:I6` | same | MATCH | |
| Hyperlinks on `J5`, `J8` | removed, `J6` | same | MATCH | |
| Comments on `L5`, `L8` | removed, `L6` | same | MATCH | |
| Row page breaks below rows 5 and 8 | removed, below row 6 | same | MATCH | |

## Inserting rows and columns

After the deletion above, one column was inserted before column A and one row before row 1 (`insert_cols` and `insert_rows` in SheetKit; `insert` on an entire column and an entire row in Excel).

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| Whole-column references: `SUM(A:A)`, `SUM(Sheet1!B:B)`, `SUM(Sheet1!A:B)` | `SUM(B:B)`, `SUM(Sheet1!C:C)`, `SUM(Sheet1!B:C)` | same | MATCH | |
| Whole-row references: `SUM(2:2)`, `SUM(Sheet1!1:1)`, `SUM(Sheet1!4:5)` | `SUM(3:3)`, `SUM(Sheet1!2:2)`, `SUM(Sheet1!5:6)` | same | MATCH | Whole-column and whole-row references shift but are not evaluated. See [Known Limitations](./known-limitations.md). |
| Merges, conditional formats, hyperlinks, comments, page breaks | `F5:F6`, `G7:H8`; `J5:J7`; `K7`; `M7`; below row 7 | same | MATCH | Data validations cover the same cells; Excel again merges identical rules. |

## Tables

Table names were set with `add_table` in SheetKit. In Excel they were set by assigning `name` on a list object through AppleScript, saving the workbook, and reading the saved `name` and `displayName`. The Excel user interface's table name box was not tested.

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| Name that differs only in case from an existing table (`sales`, `SALES` when `Sales` exists) | `add_table` returns an error ("table 'sales' already exists") | Accepted; saved `name="sales"`, `displayName="sales_1"` | DIFFERS | Deliberate. Excel's object model renames the display name; SheetKit rejects the name instead of renaming silently. |
| Name that is a cell reference (`A1`, `XFD1048576`, `R1C1`, `R`, `C`, `RC`, `R1`) | Error ("must not be a cell reference") | `A1` saved as `displayName="A1_"`; `R1C1` as `"_R1C1"` | DIFFERS | Deliberate, as above. `XFE1` (beyond the last column) is accepted by both. |
| Name with a space (`Bad Name`) | Error ("contains invalid character") | Saved as `displayName="Bad_Name"` | DIFFERS | Deliberate, as above. |
| Name starting with a digit (`1abc`) | Error ("must start with a letter, underscore, or backslash") | Saved as `displayName="_1abc"` | DIFFERS | Deliberate, as above. |
| 256-character name; 255-character name | Error; accepted | Error -10006 ("Can't set name"); accepted | MATCH | |
| `_ok`, `\back`, `a.b` | Accepted | Accepted | MATCH | |
| Totals row present | `TableConfig` column with `totals_row_function` or `totals_row_label` writes `totalsRowCount="1"`; the table `ref` includes the totals row and the `autoFilter` ref excludes it | `totalsRowCount="1"`, `ref="A1:B4"`, `autoFilter ref="A1:B3"` | MATCH | |
| `SUBTOTAL` function numbers for `sum`, `average`, `count`, `countNums`, `min`, `max`, `stdDev`, `var` | 109, 101, 103, 102, 105, 104, 107, 110 | 109, 101, 103, 102, 105, 104, 107, 110 | MATCH | |
| Accepted totals function names | `none`, `sum`, `average`, `count`, `countNums`, `min`, `max`, `stdDev`, `var`; `custom` is rejected | The same names plus `custom` | DIFFERS | Not supported. A custom totals formula needs `totalsRowFormula`, which `TableConfig` cannot set. See [Known Limitations](./known-limitations.md). |
| Totals cell formula text | `SUBTOTAL(109,[Qty])` | Displays `=SUBTOTAL(109,[Qty])` in the cell; saves `SUBTOTAL(109,Sales[Qty])` | MATCH | Excel opens SheetKit's form without repair and calculates it. |

## Merged cells

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| Merging `B2:C3` when `A1:B2` is already merged | `Workbook::merge_cells` and `StreamWriter::add_merge_cell` return a merge overlap error | Accepted; the merge becomes `A1:C3` | DIFFERS | Deliberate. SheetKit rejects overlapping merges instead of silently enlarging the existing merge. |
| Merging the same range again (`E1:F2` twice) | `StreamWriter::add_merge_cell`: no-op. `Workbook::merge_cells`: overlap error | No-op | MATCH (stream), DIFFERS (`Workbook`) | The `Workbook` difference is deliberate. Calling `merge_cells` again on an existing merge is reported rather than ignored. |

## Comments and shared strings

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| Notes list on a sheet with one note and one threaded comment | `get_comments` returns only the note | `Excel comments` count is 1 (the note) | MATCH | Excel does not list the `[Threaded comment]` legacy twin as a note. |
| Threaded comment content as shown in Excel | `get_threaded_comments` returns the thread | Not readable | NOT VERIFIABLE | Excel's AppleScript dictionary has no access to threaded comments. |
| Plain `Hello` in `A1` and bold rich text `Hello` in `A2` | Two shared string entries | `A1` characters not bold, `A2` characters bold | MATCH | The two cells keep separate formatting. |

## Not Excel-visible

| Behavior | SheetKit | Excel | Status | Notes |
| --- | --- | --- | --- | --- |
| Deleting the hyperlink on one of two cells that share a relationship | `delete_cell_hyperlink` removes the relationship only when no remaining hyperlink uses it | n/a | N/A | Package bookkeeping. |
| New relationship ids | The lowest unused `rIdN` is used, so gaps are reused | n/a | N/A | Package bookkeeping. |
| Node.js integer arguments (rows, columns, counts, style ids, outline levels, numeric style and option fields) | Non-finite, fractional, unsafe (beyond `Number.MAX_SAFE_INTEGER`), and out-of-range values throw "`<name>` must be an integer from `<min>` to `<max>`" | n/a | N/A | API contract. Earlier releases let napi truncate, wrap, or turn `NaN` into 0. |

## Summary

Of the 74 rows above, 54 MATCH, 14 DIFFER, 1 is split (merging the same range again matches for `StreamWriter` and differs for `Workbook`), 2 are NOT VERIFIABLE, and 3 are N/A.

- **Security limits**: the `TEXT` format code rows. SheetKit keeps formatting bounded instead of rejecting format codes Excel rejects.
- **Deliberate**: table naming (reject instead of rename) and merges (reject overlaps instead of enlarging).
- **Not reproduced**: Excel's clamping of out-of-range `DATE` day arguments, Excel's 15-digit display precision in `TEXT`, and floating-point error in `ROUND` near 1E308.
