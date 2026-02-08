//! Column operations for worksheet manipulation.
//!
//! All functions operate directly on a [`WorksheetXml`] structure, keeping the
//! business logic decoupled from the [`Workbook`](crate::workbook::Workbook)
//! wrapper.

use sheetkit_xml::worksheet::{Col, Cols, WorksheetXml};

use crate::error::{Error, Result};
use crate::utils::cell_ref::{
    cell_name_to_coordinates, column_name_to_number, coordinates_to_cell_name,
};
use crate::utils::constants::{MAX_COLUMNS, MAX_COLUMN_WIDTH};

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

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
    if furthest
        .checked_add(count)
        .is_none_or(|v| v > MAX_COLUMNS)
    {
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

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

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

// ===========================================================================
// Tests
// ===========================================================================

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

    // ----- set_col_width / get_col_width --------------------------------

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

    // ----- set_col_visible ----------------------------------------------

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

    // ----- insert_cols --------------------------------------------------

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

    // ----- remove_col ---------------------------------------------------

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

    // ----- Multiple column widths ---------------------------------------

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
}
