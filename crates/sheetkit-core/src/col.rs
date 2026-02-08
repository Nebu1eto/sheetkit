//! Column operations for worksheet manipulation.
//!
//! All functions operate directly on a [`WorksheetXml`] structure, keeping the
//! business logic decoupled from the [`Workbook`](crate::workbook::Workbook)
//! wrapper.

use std::collections::BTreeMap;

use sheetkit_xml::worksheet::{Col, Cols, WorksheetXml};

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::row::get_rows;
use crate::sst::SharedStringTable;
use crate::utils::cell_ref::{
    cell_name_to_coordinates, column_name_to_number, coordinates_to_cell_name,
};
use crate::utils::constants::{MAX_COLUMNS, MAX_COLUMN_WIDTH};

/// Get all columns with their data from a worksheet.
///
/// Returns a Vec of `(column_name, Vec<(row_number, CellValue)>)` tuples.
/// Only columns that have data are included (sparse). The columns are sorted
/// alphabetically (A, B, C, ... Z, AA, AB, ...).
#[allow(clippy::type_complexity)]
pub fn get_cols(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
) -> Result<Vec<(String, Vec<(u32, CellValue)>)>> {
    let rows = get_rows(ws, sst)?;

    // Transpose row-based data into column-based data using a BTreeMap
    // to keep columns in sorted order.
    let mut col_map: BTreeMap<String, Vec<(u32, CellValue)>> = BTreeMap::new();

    for (row_num, cells) in rows {
        for (col_name, value) in cells {
            col_map.entry(col_name).or_default().push((row_num, value));
        }
    }

    // BTreeMap sorts by String, but column names need a special sort:
    // "A" < "B" < ... < "Z" < "AA" < "AB" etc.
    // We sort by (length, name) to get the correct order.
    let mut result: Vec<(String, Vec<(u32, CellValue)>)> = col_map.into_iter().collect();
    result.sort_by(|(a, _), (b, _)| a.len().cmp(&b.len()).then_with(|| a.cmp(b)));

    Ok(result)
}

/// Set the width of a column. Creates the `Cols` container and/or a `Col`
/// entry if they do not yet exist.
///
/// Valid range: `0.0 ..= 255.0`.
pub fn set_col_width(ws: &mut WorksheetXml, col: &str, width: f64) -> Result<()> {
    let col_num = column_name_to_number(col)?;
    if !(0.0..=MAX_COLUMN_WIDTH).contains(&width) {
        return Err(Error::ColumnWidthExceeded {
            width,
            max: MAX_COLUMN_WIDTH,
        });
    }

    let col_entry = find_or_create_col(ws, col_num);
    col_entry.width = Some(width);
    col_entry.custom_width = Some(true);
    Ok(())
}

/// Get the width of a column. Returns `None` when there is no explicit width
/// defined for the column.
pub fn get_col_width(ws: &WorksheetXml, col: &str) -> Option<f64> {
    let col_num = column_name_to_number(col).ok()?;
    ws.cols
        .as_ref()
        .and_then(|cols| {
            cols.cols
                .iter()
                .find(|c| col_num >= c.min && col_num <= c.max)
        })
        .and_then(|c| c.width)
}

/// Set the visibility of a column.
pub fn set_col_visible(ws: &mut WorksheetXml, col: &str, visible: bool) -> Result<()> {
    let col_num = column_name_to_number(col)?;
    let col_entry = find_or_create_col(ws, col_num);
    col_entry.hidden = if visible { None } else { Some(true) };
    Ok(())
}

/// Get the visibility of a column. Returns true if visible (not hidden).
///
/// Columns are visible by default, so this returns true if no `Col` entry
/// exists or if it has no `hidden` attribute set.
pub fn get_col_visible(ws: &WorksheetXml, col: &str) -> Result<bool> {
    let col_num = column_name_to_number(col)?;
    let hidden = ws
        .cols
        .as_ref()
        .and_then(|cols| {
            cols.cols
                .iter()
                .find(|c| col_num >= c.min && col_num <= c.max)
        })
        .and_then(|c| c.hidden)
        .unwrap_or(false);
    Ok(!hidden)
}

/// Set the outline (grouping) level of a column.
///
/// Valid range: `0..=7` (Excel supports up to 7 outline levels).
pub fn set_col_outline_level(ws: &mut WorksheetXml, col: &str, level: u8) -> Result<()> {
    let col_num = column_name_to_number(col)?;
    if level > 7 {
        return Err(Error::Internal(format!(
            "outline level {level} exceeds maximum 7"
        )));
    }

    let col_entry = find_or_create_col(ws, col_num);
    col_entry.outline_level = if level == 0 { None } else { Some(level) };
    Ok(())
}

/// Get the outline (grouping) level of a column. Returns 0 if not set.
pub fn get_col_outline_level(ws: &WorksheetXml, col: &str) -> Result<u8> {
    let col_num = column_name_to_number(col)?;
    let level = ws
        .cols
        .as_ref()
        .and_then(|cols| {
            cols.cols
                .iter()
                .find(|c| col_num >= c.min && col_num <= c.max)
        })
        .and_then(|c| c.outline_level)
        .unwrap_or(0);
    Ok(level)
}

/// Insert `count` columns starting at `col`, shifting existing columns at
/// and to the right of `col` further right.
///
/// Cell references are updated: e.g. when inserting 2 columns at "B", a cell
/// "C3" becomes "E3".
pub fn insert_cols(ws: &mut WorksheetXml, col: &str, count: u32) -> Result<()> {
    let start_col = column_name_to_number(col)?;
    if count == 0 {
        return Ok(());
    }

    // Validate we won't exceed max columns.
    let max_existing = ws
        .sheet_data
        .rows
        .iter()
        .flat_map(|r| r.cells.iter())
        .filter_map(|c| cell_name_to_coordinates(&c.r).ok())
        .map(|(col_n, _)| col_n)
        .max()
        .unwrap_or(0);
    let furthest = max_existing.max(start_col);
    if furthest.checked_add(count).is_none_or(|v| v > MAX_COLUMNS) {
        return Err(Error::InvalidColumnNumber(furthest + count));
    }

    // Shift cell references.
    for row in ws.sheet_data.rows.iter_mut() {
        for cell in row.cells.iter_mut() {
            let (c, r) = cell_name_to_coordinates(&cell.r)?;
            if c >= start_col {
                cell.r = coordinates_to_cell_name(c + count, r)?;
            }
        }
    }

    // Shift Col definitions.
    if let Some(ref mut cols) = ws.cols {
        for c in cols.cols.iter_mut() {
            if c.min >= start_col {
                c.min += count;
            }
            if c.max >= start_col {
                c.max += count;
            }
        }
    }

    Ok(())
}

/// Remove a single column, shifting columns to its right leftward by one.
pub fn remove_col(ws: &mut WorksheetXml, col: &str) -> Result<()> {
    let col_num = column_name_to_number(col)?;

    // Remove cells in the target column and shift cells to the right.
    for row in ws.sheet_data.rows.iter_mut() {
        // Remove cells at the target column.
        row.cells.retain(|cell| {
            cell_name_to_coordinates(&cell.r)
                .map(|(c, _)| c != col_num)
                .unwrap_or(true)
        });

        // Shift cells that are to the right of the removed column.
        for cell in row.cells.iter_mut() {
            let (c, r) = cell_name_to_coordinates(&cell.r).unwrap();
            if c > col_num {
                cell.r = coordinates_to_cell_name(c - 1, r).unwrap();
            }
        }
    }

    // Shift Col definitions.
    if let Some(ref mut cols) = ws.cols {
        // Remove col entries that exactly span the removed column only.
        cols.cols
            .retain(|c| !(c.min == col_num && c.max == col_num));

        for c in cols.cols.iter_mut() {
            if c.min > col_num {
                c.min -= 1;
            }
            if c.max >= col_num {
                // Only shrink max if it's above the removed column.
                if c.max > col_num {
                    c.max -= 1;
                }
            }
        }

        // Remove the Cols wrapper if it's now empty.
        if cols.cols.is_empty() {
            ws.cols = None;
        }
    }

    Ok(())
}

/// Find an existing Col entry that covers exactly `col_num`, or create a
/// new single-column entry for it.
fn find_or_create_col(ws: &mut WorksheetXml, col_num: u32) -> &mut Col {
    // Ensure the Cols container exists.
    if ws.cols.is_none() {
        ws.cols = Some(Cols { cols: vec![] });
    }
    let cols = ws.cols.as_mut().unwrap();

    // Look for an existing entry that spans exactly this column.
    let existing = cols
        .cols
        .iter()
        .position(|c| c.min == col_num && c.max == col_num);

    if let Some(idx) = existing {
        return &mut cols.cols[idx];
    }

    // Create a new single-column entry.
    cols.cols.push(Col {
        min: col_num,
        max: col_num,
        width: None,
        style: None,
        hidden: None,
        custom_width: None,
        outline_level: None,
    });
    let last = cols.cols.len() - 1;
    &mut cols.cols[last]
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetkit_xml::worksheet::{Cell, Row, SheetData};

    /// Helper: build a worksheet with some cells for column tests.
    fn sample_ws() -> WorksheetXml {
        let mut ws = WorksheetXml::default();
        ws.sheet_data = SheetData {
            rows: vec![
                Row {
                    r: 1,
                    spans: None,
                    s: None,
                    custom_format: None,
                    ht: None,
                    hidden: None,
                    custom_height: None,
                    outline_level: None,
                    cells: vec![
                        Cell {
                            r: "A1".to_string(),
                            s: None,
                            t: None,
                            v: Some("10".to_string()),
                            f: None,
                            is: None,
                        },
                        Cell {
                            r: "B1".to_string(),
                            s: None,
                            t: None,
                            v: Some("20".to_string()),
                            f: None,
                            is: None,
                        },
                        Cell {
                            r: "D1".to_string(),
                            s: None,
                            t: None,
                            v: Some("40".to_string()),
                            f: None,
                            is: None,
                        },
                    ],
                },
                Row {
                    r: 2,
                    spans: None,
                    s: None,
                    custom_format: None,
                    ht: None,
                    hidden: None,
                    custom_height: None,
                    outline_level: None,
                    cells: vec![
                        Cell {
                            r: "A2".to_string(),
                            s: None,
                            t: None,
                            v: Some("100".to_string()),
                            f: None,
                            is: None,
                        },
                        Cell {
                            r: "C2".to_string(),
                            s: None,
                            t: None,
                            v: Some("300".to_string()),
                            f: None,
                            is: None,
                        },
                    ],
                },
            ],
        };
        ws
    }

    #[test]
    fn test_set_and_get_col_width() {
        let mut ws = WorksheetXml::default();
        set_col_width(&mut ws, "A", 15.0).unwrap();
        assert_eq!(get_col_width(&ws, "A"), Some(15.0));
    }

    #[test]
    fn test_set_col_width_creates_cols_container() {
        let mut ws = WorksheetXml::default();
        assert!(ws.cols.is_none());
        set_col_width(&mut ws, "B", 20.0).unwrap();
        assert!(ws.cols.is_some());
        let cols = ws.cols.as_ref().unwrap();
        assert_eq!(cols.cols.len(), 1);
        assert_eq!(cols.cols[0].min, 2);
        assert_eq!(cols.cols[0].max, 2);
        assert_eq!(cols.cols[0].custom_width, Some(true));
    }

    #[test]
    fn test_set_col_width_zero_is_valid() {
        let mut ws = WorksheetXml::default();
        set_col_width(&mut ws, "A", 0.0).unwrap();
        assert_eq!(get_col_width(&ws, "A"), Some(0.0));
    }

    #[test]
    fn test_set_col_width_max_is_valid() {
        let mut ws = WorksheetXml::default();
        set_col_width(&mut ws, "A", 255.0).unwrap();
        assert_eq!(get_col_width(&ws, "A"), Some(255.0));
    }

    #[test]
    fn test_set_col_width_exceeds_max_returns_error() {
        let mut ws = WorksheetXml::default();
        let result = set_col_width(&mut ws, "A", 256.0);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::ColumnWidthExceeded { .. }
        ));
    }

    #[test]
    fn test_set_col_width_negative_returns_error() {
        let mut ws = WorksheetXml::default();
        let result = set_col_width(&mut ws, "A", -1.0);
        assert!(result.is_err());
    }

    #[test]
    fn test_get_col_width_nonexistent_returns_none() {
        let ws = WorksheetXml::default();
        assert_eq!(get_col_width(&ws, "Z"), None);
    }

    #[test]
    fn test_set_col_width_invalid_column_returns_error() {
        let mut ws = WorksheetXml::default();
        let result = set_col_width(&mut ws, "XFE", 10.0);
        assert!(result.is_err());
    }

    #[test]
    fn test_set_col_hidden() {
        let mut ws = WorksheetXml::default();
        set_col_visible(&mut ws, "A", false).unwrap();

        let col = &ws.cols.as_ref().unwrap().cols[0];
        assert_eq!(col.hidden, Some(true));
    }

    #[test]
    fn test_set_col_visible_clears_hidden() {
        let mut ws = WorksheetXml::default();
        set_col_visible(&mut ws, "A", false).unwrap();
        set_col_visible(&mut ws, "A", true).unwrap();

        let col = &ws.cols.as_ref().unwrap().cols[0];
        assert_eq!(col.hidden, None);
    }

    #[test]
    fn test_insert_cols_shifts_cells_right() {
        let mut ws = sample_ws();
        insert_cols(&mut ws, "B", 2).unwrap();

        // Row 1: A1 stays, B1->D1, D1->F1.
        let r1 = &ws.sheet_data.rows[0];
        assert_eq!(r1.cells[0].r, "A1");
        assert_eq!(r1.cells[1].r, "D1"); // B shifted by 2
        assert_eq!(r1.cells[2].r, "F1"); // D shifted by 2

        // Row 2: A2 stays, C2->E2.
        let r2 = &ws.sheet_data.rows[1];
        assert_eq!(r2.cells[0].r, "A2");
        assert_eq!(r2.cells[1].r, "E2"); // C shifted by 2
    }

    #[test]
    fn test_insert_cols_at_column_a() {
        let mut ws = sample_ws();
        insert_cols(&mut ws, "A", 1).unwrap();

        // All cells shift right by 1.
        let r1 = &ws.sheet_data.rows[0];
        assert_eq!(r1.cells[0].r, "B1"); // A->B
        assert_eq!(r1.cells[1].r, "C1"); // B->C
        assert_eq!(r1.cells[2].r, "E1"); // D->E
    }

    #[test]
    fn test_insert_cols_count_zero_is_noop() {
        let mut ws = sample_ws();
        insert_cols(&mut ws, "B", 0).unwrap();
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");
        assert_eq!(ws.sheet_data.rows[0].cells[1].r, "B1");
    }

    #[test]
    fn test_insert_cols_on_empty_sheet() {
        let mut ws = WorksheetXml::default();
        insert_cols(&mut ws, "A", 5).unwrap();
        assert!(ws.sheet_data.rows.is_empty());
    }

    #[test]
    fn test_insert_cols_shifts_col_definitions() {
        let mut ws = WorksheetXml::default();
        set_col_width(&mut ws, "C", 20.0).unwrap();
        insert_cols(&mut ws, "B", 2).unwrap();

        // Col C (3) should now be at col 5 (E).
        let col = &ws.cols.as_ref().unwrap().cols[0];
        assert_eq!(col.min, 5);
        assert_eq!(col.max, 5);
    }

    #[test]
    fn test_remove_col_shifts_cells_left() {
        let mut ws = sample_ws();
        remove_col(&mut ws, "B").unwrap();

        // Row 1: A1 stays, B1 removed, D1->C1.
        let r1 = &ws.sheet_data.rows[0];
        assert_eq!(r1.cells.len(), 2);
        assert_eq!(r1.cells[0].r, "A1");
        assert_eq!(r1.cells[1].r, "C1"); // D shifted left
        assert_eq!(r1.cells[1].v, Some("40".to_string()));

        // Row 2: A2 stays, C2->B2.
        let r2 = &ws.sheet_data.rows[1];
        assert_eq!(r2.cells[0].r, "A2");
        assert_eq!(r2.cells[1].r, "B2"); // C shifted left
    }

    #[test]
    fn test_remove_first_col() {
        let mut ws = sample_ws();
        remove_col(&mut ws, "A").unwrap();

        // Row 1: A1 removed, B1->A1, D1->C1.
        let r1 = &ws.sheet_data.rows[0];
        assert_eq!(r1.cells.len(), 2);
        assert_eq!(r1.cells[0].r, "A1");
        assert_eq!(r1.cells[0].v, Some("20".to_string())); // was B1
        assert_eq!(r1.cells[1].r, "C1");
        assert_eq!(r1.cells[1].v, Some("40".to_string())); // was D1
    }

    #[test]
    fn test_remove_col_with_col_definitions() {
        let mut ws = WorksheetXml::default();
        set_col_width(&mut ws, "B", 20.0).unwrap();
        remove_col(&mut ws, "B").unwrap();

        // The Col entry for B should be removed.
        assert!(ws.cols.is_none());
    }

    #[test]
    fn test_remove_col_invalid_column_returns_error() {
        let mut ws = WorksheetXml::default();
        let result = remove_col(&mut ws, "XFE");
        assert!(result.is_err());
    }

    #[test]
    fn test_set_multiple_col_widths() {
        let mut ws = WorksheetXml::default();
        set_col_width(&mut ws, "A", 10.0).unwrap();
        set_col_width(&mut ws, "C", 30.0).unwrap();

        assert_eq!(get_col_width(&ws, "A"), Some(10.0));
        assert_eq!(get_col_width(&ws, "B"), None);
        assert_eq!(get_col_width(&ws, "C"), Some(30.0));
    }

    #[test]
    fn test_overwrite_col_width() {
        let mut ws = WorksheetXml::default();
        set_col_width(&mut ws, "A", 10.0).unwrap();
        set_col_width(&mut ws, "A", 25.0).unwrap();

        assert_eq!(get_col_width(&ws, "A"), Some(25.0));
    }

    #[test]
    fn test_get_col_visible_default_is_true() {
        let ws = WorksheetXml::default();
        assert!(get_col_visible(&ws, "A").unwrap());
    }

    #[test]
    fn test_get_col_visible_after_hide() {
        let mut ws = WorksheetXml::default();
        set_col_visible(&mut ws, "B", false).unwrap();
        assert!(!get_col_visible(&ws, "B").unwrap());
    }

    #[test]
    fn test_get_col_visible_after_hide_then_show() {
        let mut ws = WorksheetXml::default();
        set_col_visible(&mut ws, "A", false).unwrap();
        set_col_visible(&mut ws, "A", true).unwrap();
        assert!(get_col_visible(&ws, "A").unwrap());
    }

    #[test]
    fn test_get_col_visible_invalid_column_returns_error() {
        let ws = WorksheetXml::default();
        let result = get_col_visible(&ws, "XFE");
        assert!(result.is_err());
    }

    #[test]
    fn test_set_col_outline_level() {
        let mut ws = WorksheetXml::default();
        set_col_outline_level(&mut ws, "A", 3).unwrap();

        let col = &ws.cols.as_ref().unwrap().cols[0];
        assert_eq!(col.outline_level, Some(3));
    }

    #[test]
    fn test_set_col_outline_level_zero_clears() {
        let mut ws = WorksheetXml::default();
        set_col_outline_level(&mut ws, "A", 3).unwrap();
        set_col_outline_level(&mut ws, "A", 0).unwrap();

        let col = &ws.cols.as_ref().unwrap().cols[0];
        assert_eq!(col.outline_level, None);
    }

    #[test]
    fn test_set_col_outline_level_exceeds_max_returns_error() {
        let mut ws = WorksheetXml::default();
        let result = set_col_outline_level(&mut ws, "A", 8);
        assert!(result.is_err());
    }

    #[test]
    fn test_set_col_outline_level_max_valid() {
        let mut ws = WorksheetXml::default();
        set_col_outline_level(&mut ws, "A", 7).unwrap();

        let col = &ws.cols.as_ref().unwrap().cols[0];
        assert_eq!(col.outline_level, Some(7));
    }

    #[test]
    fn test_get_col_outline_level_default_is_zero() {
        let ws = WorksheetXml::default();
        assert_eq!(get_col_outline_level(&ws, "A").unwrap(), 0);
    }

    #[test]
    fn test_get_col_outline_level_after_set() {
        let mut ws = WorksheetXml::default();
        set_col_outline_level(&mut ws, "B", 5).unwrap();
        assert_eq!(get_col_outline_level(&ws, "B").unwrap(), 5);
    }

    #[test]
    fn test_get_col_outline_level_after_clear() {
        let mut ws = WorksheetXml::default();
        set_col_outline_level(&mut ws, "C", 4).unwrap();
        set_col_outline_level(&mut ws, "C", 0).unwrap();
        assert_eq!(get_col_outline_level(&ws, "C").unwrap(), 0);
    }

    #[test]
    fn test_get_col_outline_level_invalid_column_returns_error() {
        let ws = WorksheetXml::default();
        let result = get_col_outline_level(&ws, "XFE");
        assert!(result.is_err());
    }

    // -- get_cols tests --

    #[test]
    fn test_get_cols_empty_sheet() {
        let ws = WorksheetXml::default();
        let sst = SharedStringTable::new();
        let cols = get_cols(&ws, &sst).unwrap();
        assert!(cols.is_empty());
    }

    #[test]
    fn test_get_cols_transposes_row_data() {
        let ws = sample_ws();
        let sst = SharedStringTable::new();
        let cols = get_cols(&ws, &sst).unwrap();

        // sample_ws has:
        //   Row 1: A1=10, B1=20, D1=40
        //   Row 2: A2=100, C2=300
        // Columns should be: A, B, C, D

        assert_eq!(cols.len(), 4);

        // Column A: (1, 10.0), (2, 100.0)
        assert_eq!(cols[0].0, "A");
        assert_eq!(cols[0].1.len(), 2);
        assert_eq!(cols[0].1[0], (1, CellValue::Number(10.0)));
        assert_eq!(cols[0].1[1], (2, CellValue::Number(100.0)));

        // Column B: (1, 20.0)
        assert_eq!(cols[1].0, "B");
        assert_eq!(cols[1].1.len(), 1);
        assert_eq!(cols[1].1[0], (1, CellValue::Number(20.0)));

        // Column C: (2, 300.0)
        assert_eq!(cols[2].0, "C");
        assert_eq!(cols[2].1.len(), 1);
        assert_eq!(cols[2].1[0], (2, CellValue::Number(300.0)));

        // Column D: (1, 40.0)
        assert_eq!(cols[3].0, "D");
        assert_eq!(cols[3].1.len(), 1);
        assert_eq!(cols[3].1[0], (1, CellValue::Number(40.0)));
    }

    #[test]
    fn test_get_cols_with_shared_strings() {
        let mut sst = SharedStringTable::new();
        sst.add("Name");
        sst.add("Age");
        sst.add("Alice");

        let mut ws = WorksheetXml::default();
        ws.sheet_data = SheetData {
            rows: vec![
                Row {
                    r: 1,
                    spans: None,
                    s: None,
                    custom_format: None,
                    ht: None,
                    hidden: None,
                    custom_height: None,
                    outline_level: None,
                    cells: vec![
                        Cell {
                            r: "A1".to_string(),
                            s: None,
                            t: Some("s".to_string()),
                            v: Some("0".to_string()),
                            f: None,
                            is: None,
                        },
                        Cell {
                            r: "B1".to_string(),
                            s: None,
                            t: Some("s".to_string()),
                            v: Some("1".to_string()),
                            f: None,
                            is: None,
                        },
                    ],
                },
                Row {
                    r: 2,
                    spans: None,
                    s: None,
                    custom_format: None,
                    ht: None,
                    hidden: None,
                    custom_height: None,
                    outline_level: None,
                    cells: vec![
                        Cell {
                            r: "A2".to_string(),
                            s: None,
                            t: Some("s".to_string()),
                            v: Some("2".to_string()),
                            f: None,
                            is: None,
                        },
                        Cell {
                            r: "B2".to_string(),
                            s: None,
                            t: None,
                            v: Some("30".to_string()),
                            f: None,
                            is: None,
                        },
                    ],
                },
            ],
        };

        let cols = get_cols(&ws, &sst).unwrap();
        assert_eq!(cols.len(), 2);

        // Column A: "Name", "Alice"
        assert_eq!(cols[0].0, "A");
        assert_eq!(cols[0].1[0].1, CellValue::String("Name".to_string()));
        assert_eq!(cols[0].1[1].1, CellValue::String("Alice".to_string()));

        // Column B: "Age", 30
        assert_eq!(cols[1].0, "B");
        assert_eq!(cols[1].1[0].1, CellValue::String("Age".to_string()));
        assert_eq!(cols[1].1[1].1, CellValue::Number(30.0));
    }

    #[test]
    fn test_get_cols_sorted_correctly() {
        // Verify that columns are sorted by length then alphabetically:
        // A, B, ..., Z, AA, AB, ...
        let mut ws = WorksheetXml::default();
        ws.sheet_data = SheetData {
            rows: vec![Row {
                r: 1,
                spans: None,
                s: None,
                custom_format: None,
                ht: None,
                hidden: None,
                custom_height: None,
                outline_level: None,
                cells: vec![
                    Cell {
                        r: "AA1".to_string(),
                        s: None,
                        t: None,
                        v: Some("1".to_string()),
                        f: None,
                        is: None,
                    },
                    Cell {
                        r: "B1".to_string(),
                        s: None,
                        t: None,
                        v: Some("2".to_string()),
                        f: None,
                        is: None,
                    },
                    Cell {
                        r: "A1".to_string(),
                        s: None,
                        t: None,
                        v: Some("3".to_string()),
                        f: None,
                        is: None,
                    },
                ],
            }],
        };

        let sst = SharedStringTable::new();
        let cols = get_cols(&ws, &sst).unwrap();

        assert_eq!(cols.len(), 3);
        assert_eq!(cols[0].0, "A");
        assert_eq!(cols[1].0, "B");
        assert_eq!(cols[2].0, "AA");
    }
}
