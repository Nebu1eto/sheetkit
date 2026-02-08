//! Row operations for worksheet manipulation.
//!
//! All functions operate directly on a [`WorksheetXml`] structure, keeping the
//! business logic decoupled from the [`Workbook`](crate::workbook::Workbook)
//! wrapper.

use sheetkit_xml::worksheet::{Row, WorksheetXml};

use crate::error::{Error, Result};
use crate::utils::cell_ref::{cell_name_to_coordinates, coordinates_to_cell_name};
use crate::utils::constants::{MAX_ROWS, MAX_ROW_HEIGHT};

/// Insert `count` empty rows starting at `start_row`, shifting existing rows
/// at and below `start_row` downward.
///
/// Cell references inside shifted rows are updated so that e.g. "B5" becomes
/// "B8" when 3 rows are inserted at row 5.
pub fn insert_rows(ws: &mut WorksheetXml, start_row: u32, count: u32) -> Result<()> {
    if start_row == 0 {
        return Err(Error::InvalidRowNumber(0));
    }
    if count == 0 {
        return Ok(());
    }
    // Validate that shifting won't exceed MAX_ROWS.
    let max_existing = ws.sheet_data.rows.iter().map(|r| r.r).max().unwrap_or(0);
    let furthest = max_existing.max(start_row);
    if furthest.checked_add(count).is_none_or(|v| v > MAX_ROWS) {
        return Err(Error::InvalidRowNumber(furthest + count));
    }

    // Shift rows that are >= start_row downward by `count`.
    // Iterate in reverse to avoid overwriting.
    for row in ws.sheet_data.rows.iter_mut().rev() {
        if row.r >= start_row {
            let new_row_num = row.r + count;
            shift_row_cells(row, new_row_num)?;
            row.r = new_row_num;
        }
    }

    Ok(())
}

/// Remove a single row, shifting rows below it upward by one.
pub fn remove_row(ws: &mut WorksheetXml, row: u32) -> Result<()> {
    if row == 0 {
        return Err(Error::InvalidRowNumber(0));
    }

    // Remove the target row.
    ws.sheet_data.rows.retain(|r| r.r != row);

    // Shift rows above `row` upward.
    for r in ws.sheet_data.rows.iter_mut() {
        if r.r > row {
            let new_row_num = r.r - 1;
            shift_row_cells(r, new_row_num)?;
            r.r = new_row_num;
        }
    }

    Ok(())
}

/// Duplicate a row, inserting the copy directly below the source row.
pub fn duplicate_row(ws: &mut WorksheetXml, row: u32) -> Result<()> {
    duplicate_row_to(ws, row, row + 1)
}

/// Duplicate a row to a specific target row number. Existing rows at and
/// below `target` are shifted down to make room.
pub fn duplicate_row_to(ws: &mut WorksheetXml, row: u32, target: u32) -> Result<()> {
    if row == 0 {
        return Err(Error::InvalidRowNumber(0));
    }
    if target == 0 {
        return Err(Error::InvalidRowNumber(0));
    }
    if target > MAX_ROWS {
        return Err(Error::InvalidRowNumber(target));
    }

    // Find and clone the source row.
    let source = ws
        .sheet_data
        .rows
        .iter()
        .find(|r| r.r == row)
        .cloned()
        .ok_or(Error::InvalidRowNumber(row))?;

    // Shift existing rows at target downward.
    insert_rows(ws, target, 1)?;

    // Build the duplicated row with updated cell references.
    let mut new_row = source;
    shift_row_cells(&mut new_row, target)?;
    new_row.r = target;

    // Insert the new row in sorted position.
    let pos = ws
        .sheet_data
        .rows
        .iter()
        .position(|r| r.r > target)
        .unwrap_or(ws.sheet_data.rows.len());
    // Check if there's already a row at target (shouldn't be, but be safe).
    if let Some(existing) = ws.sheet_data.rows.iter().position(|r| r.r == target) {
        ws.sheet_data.rows[existing] = new_row;
    } else {
        ws.sheet_data.rows.insert(pos, new_row);
    }

    Ok(())
}

/// Set the height of a row in points. Creates the row if it does not exist.
///
/// Valid range: `0.0 ..= 409.0`.
pub fn set_row_height(ws: &mut WorksheetXml, row: u32, height: f64) -> Result<()> {
    if row == 0 || row > MAX_ROWS {
        return Err(Error::InvalidRowNumber(row));
    }
    if !(0.0..=MAX_ROW_HEIGHT).contains(&height) {
        return Err(Error::RowHeightExceeded {
            height,
            max: MAX_ROW_HEIGHT,
        });
    }

    let r = find_or_create_row(ws, row);
    r.ht = Some(height);
    r.custom_height = Some(true);
    Ok(())
}

/// Get the height of a row. Returns `None` if the row does not exist or has
/// no explicit height set.
pub fn get_row_height(ws: &WorksheetXml, row: u32) -> Option<f64> {
    ws.sheet_data
        .rows
        .iter()
        .find(|r| r.r == row)
        .and_then(|r| r.ht)
}

/// Set the visibility of a row. Creates the row if it does not exist.
pub fn set_row_visible(ws: &mut WorksheetXml, row: u32, visible: bool) -> Result<()> {
    if row == 0 || row > MAX_ROWS {
        return Err(Error::InvalidRowNumber(row));
    }

    let r = find_or_create_row(ws, row);
    r.hidden = if visible { None } else { Some(true) };
    Ok(())
}

/// Set the outline (grouping) level of a row.
///
/// Valid range: `0..=7` (Excel supports up to 7 outline levels).
pub fn set_row_outline_level(ws: &mut WorksheetXml, row: u32, level: u8) -> Result<()> {
    if row == 0 || row > MAX_ROWS {
        return Err(Error::InvalidRowNumber(row));
    }
    if level > 7 {
        return Err(Error::Internal(format!(
            "outline level {level} exceeds maximum 7"
        )));
    }

    let r = find_or_create_row(ws, row);
    r.outline_level = if level == 0 { None } else { Some(level) };
    Ok(())
}

/// Update all cell references in a row to point to `new_row_num`.
fn shift_row_cells(row: &mut Row, new_row_num: u32) -> Result<()> {
    for cell in row.cells.iter_mut() {
        let (col, _) = cell_name_to_coordinates(&cell.r)?;
        cell.r = coordinates_to_cell_name(col, new_row_num)?;
    }
    Ok(())
}

/// Find an existing row or create a new empty one, keeping rows sorted.
fn find_or_create_row(ws: &mut WorksheetXml, row: u32) -> &mut Row {
    // Check if row exists already.
    let exists = ws.sheet_data.rows.iter().position(|r| r.r == row);
    if let Some(idx) = exists {
        return &mut ws.sheet_data.rows[idx];
    }

    // Insert in sorted order.
    let pos = ws
        .sheet_data
        .rows
        .iter()
        .position(|r| r.r > row)
        .unwrap_or(ws.sheet_data.rows.len());
    ws.sheet_data.rows.insert(
        pos,
        Row {
            r: row,
            spans: None,
            s: None,
            custom_format: None,
            ht: None,
            hidden: None,
            custom_height: None,
            outline_level: None,
            cells: vec![],
        },
    );
    &mut ws.sheet_data.rows[pos]
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetkit_xml::worksheet::{Cell, SheetData};

    /// Helper: build a minimal worksheet with some pre-populated rows.
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
                    cells: vec![Cell {
                        r: "A2".to_string(),
                        s: None,
                        t: None,
                        v: Some("30".to_string()),
                        f: None,
                        is: None,
                    }],
                },
                Row {
                    r: 5,
                    spans: None,
                    s: None,
                    custom_format: None,
                    ht: None,
                    hidden: None,
                    custom_height: None,
                    outline_level: None,
                    cells: vec![Cell {
                        r: "C5".to_string(),
                        s: None,
                        t: None,
                        v: Some("50".to_string()),
                        f: None,
                        is: None,
                    }],
                },
            ],
        };
        ws
    }

    #[test]
    fn test_insert_rows_shifts_cells_down() {
        let mut ws = sample_ws();
        insert_rows(&mut ws, 2, 3).unwrap();

        // Row 1 should be untouched.
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");

        // Row 2 -> 5 (shifted by 3).
        assert_eq!(ws.sheet_data.rows[1].r, 5);
        assert_eq!(ws.sheet_data.rows[1].cells[0].r, "A5");

        // Row 5 -> 8 (shifted by 3).
        assert_eq!(ws.sheet_data.rows[2].r, 8);
        assert_eq!(ws.sheet_data.rows[2].cells[0].r, "C8");
    }

    #[test]
    fn test_insert_rows_at_row_1() {
        let mut ws = sample_ws();
        insert_rows(&mut ws, 1, 2).unwrap();

        // All rows shift by 2.
        assert_eq!(ws.sheet_data.rows[0].r, 3);
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A3");
        assert_eq!(ws.sheet_data.rows[1].r, 4);
        assert_eq!(ws.sheet_data.rows[2].r, 7);
    }

    #[test]
    fn test_insert_rows_count_zero_is_noop() {
        let mut ws = sample_ws();
        insert_rows(&mut ws, 1, 0).unwrap();
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[1].r, 2);
        assert_eq!(ws.sheet_data.rows[2].r, 5);
    }

    #[test]
    fn test_insert_rows_row_zero_returns_error() {
        let mut ws = sample_ws();
        let result = insert_rows(&mut ws, 0, 1);
        assert!(result.is_err());
    }

    #[test]
    fn test_insert_rows_beyond_max_returns_error() {
        let mut ws = WorksheetXml::default();
        ws.sheet_data.rows.push(Row {
            r: MAX_ROWS,
            spans: None,
            s: None,
            custom_format: None,
            ht: None,
            hidden: None,
            custom_height: None,
            outline_level: None,
            cells: vec![],
        });
        let result = insert_rows(&mut ws, 1, 1);
        assert!(result.is_err());
    }

    #[test]
    fn test_insert_rows_on_empty_sheet() {
        let mut ws = WorksheetXml::default();
        insert_rows(&mut ws, 1, 5).unwrap();
        assert!(ws.sheet_data.rows.is_empty());
    }

    #[test]
    fn test_remove_row_shifts_up() {
        let mut ws = sample_ws();
        remove_row(&mut ws, 2).unwrap();

        // Row 1 untouched.
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");

        // Original row 2 is gone; row 5 shifted to 4.
        assert_eq!(ws.sheet_data.rows.len(), 2);
        assert_eq!(ws.sheet_data.rows[1].r, 4);
        assert_eq!(ws.sheet_data.rows[1].cells[0].r, "C4");
    }

    #[test]
    fn test_remove_first_row() {
        let mut ws = sample_ws();
        remove_row(&mut ws, 1).unwrap();

        // Remaining rows shift up.
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");
        assert_eq!(ws.sheet_data.rows[1].r, 4);
    }

    #[test]
    fn test_remove_nonexistent_row_still_shifts() {
        let mut ws = sample_ws();
        // Row 3 doesn't exist, but rows below should shift.
        remove_row(&mut ws, 3).unwrap();
        assert_eq!(ws.sheet_data.rows.len(), 3); // no row removed
        assert_eq!(ws.sheet_data.rows[2].r, 4); // row 5 -> 4
    }

    #[test]
    fn test_remove_row_zero_returns_error() {
        let mut ws = sample_ws();
        let result = remove_row(&mut ws, 0);
        assert!(result.is_err());
    }

    #[test]
    fn test_duplicate_row_inserts_copy_below() {
        let mut ws = sample_ws();
        duplicate_row(&mut ws, 1).unwrap();

        // Row 1 stays.
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");
        assert_eq!(ws.sheet_data.rows[0].cells[0].v, Some("10".to_string()));

        // Row 2 is the duplicate with updated refs.
        assert_eq!(ws.sheet_data.rows[1].r, 2);
        assert_eq!(ws.sheet_data.rows[1].cells[0].r, "A2");
        assert_eq!(ws.sheet_data.rows[1].cells[0].v, Some("10".to_string()));
        assert_eq!(ws.sheet_data.rows[1].cells.len(), 2);

        // Original row 2 shifted to 3.
        assert_eq!(ws.sheet_data.rows[2].r, 3);
        assert_eq!(ws.sheet_data.rows[2].cells[0].r, "A3");
    }

    #[test]
    fn test_duplicate_row_to_specific_target() {
        let mut ws = sample_ws();
        duplicate_row_to(&mut ws, 1, 5).unwrap();

        // Row 1 unchanged.
        assert_eq!(ws.sheet_data.rows[0].r, 1);

        // Target row 5 is the copy.
        let row5 = ws.sheet_data.rows.iter().find(|r| r.r == 5).unwrap();
        assert_eq!(row5.cells[0].r, "A5");
        assert_eq!(row5.cells[0].v, Some("10".to_string()));

        // Original row 5 shifted to 6.
        let row6 = ws.sheet_data.rows.iter().find(|r| r.r == 6).unwrap();
        assert_eq!(row6.cells[0].r, "C6");
    }

    #[test]
    fn test_duplicate_nonexistent_row_returns_error() {
        let mut ws = sample_ws();
        let result = duplicate_row(&mut ws, 99);
        assert!(result.is_err());
    }

    #[test]
    fn test_set_and_get_row_height() {
        let mut ws = sample_ws();
        set_row_height(&mut ws, 1, 25.5).unwrap();

        assert_eq!(get_row_height(&ws, 1), Some(25.5));
        let row = ws.sheet_data.rows.iter().find(|r| r.r == 1).unwrap();
        assert_eq!(row.custom_height, Some(true));
    }

    #[test]
    fn test_set_row_height_creates_row_if_missing() {
        let mut ws = WorksheetXml::default();
        set_row_height(&mut ws, 10, 30.0).unwrap();

        assert_eq!(get_row_height(&ws, 10), Some(30.0));
        assert_eq!(ws.sheet_data.rows.len(), 1);
        assert_eq!(ws.sheet_data.rows[0].r, 10);
    }

    #[test]
    fn test_set_row_height_zero_is_valid() {
        let mut ws = WorksheetXml::default();
        set_row_height(&mut ws, 1, 0.0).unwrap();
        assert_eq!(get_row_height(&ws, 1), Some(0.0));
    }

    #[test]
    fn test_set_row_height_max_is_valid() {
        let mut ws = WorksheetXml::default();
        set_row_height(&mut ws, 1, 409.0).unwrap();
        assert_eq!(get_row_height(&ws, 1), Some(409.0));
    }

    #[test]
    fn test_set_row_height_exceeds_max_returns_error() {
        let mut ws = WorksheetXml::default();
        let result = set_row_height(&mut ws, 1, 410.0);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::RowHeightExceeded { .. }
        ));
    }

    #[test]
    fn test_set_row_height_negative_returns_error() {
        let mut ws = WorksheetXml::default();
        let result = set_row_height(&mut ws, 1, -1.0);
        assert!(result.is_err());
    }

    #[test]
    fn test_set_row_height_row_zero_returns_error() {
        let mut ws = WorksheetXml::default();
        let result = set_row_height(&mut ws, 0, 15.0);
        assert!(result.is_err());
    }

    #[test]
    fn test_get_row_height_nonexistent_returns_none() {
        let ws = WorksheetXml::default();
        assert_eq!(get_row_height(&ws, 99), None);
    }

    #[test]
    fn test_set_row_hidden() {
        let mut ws = sample_ws();
        set_row_visible(&mut ws, 1, false).unwrap();

        let row = ws.sheet_data.rows.iter().find(|r| r.r == 1).unwrap();
        assert_eq!(row.hidden, Some(true));
    }

    #[test]
    fn test_set_row_visible_clears_hidden() {
        let mut ws = sample_ws();
        set_row_visible(&mut ws, 1, false).unwrap();
        set_row_visible(&mut ws, 1, true).unwrap();

        let row = ws.sheet_data.rows.iter().find(|r| r.r == 1).unwrap();
        assert_eq!(row.hidden, None);
    }

    #[test]
    fn test_set_row_visible_creates_row_if_missing() {
        let mut ws = WorksheetXml::default();
        set_row_visible(&mut ws, 3, false).unwrap();
        assert_eq!(ws.sheet_data.rows.len(), 1);
        assert_eq!(ws.sheet_data.rows[0].r, 3);
        assert_eq!(ws.sheet_data.rows[0].hidden, Some(true));
    }

    #[test]
    fn test_set_row_visible_row_zero_returns_error() {
        let mut ws = WorksheetXml::default();
        let result = set_row_visible(&mut ws, 0, true);
        assert!(result.is_err());
    }

    #[test]
    fn test_set_row_outline_level() {
        let mut ws = sample_ws();
        set_row_outline_level(&mut ws, 1, 3).unwrap();

        let row = ws.sheet_data.rows.iter().find(|r| r.r == 1).unwrap();
        assert_eq!(row.outline_level, Some(3));
    }

    #[test]
    fn test_set_row_outline_level_zero_clears() {
        let mut ws = sample_ws();
        set_row_outline_level(&mut ws, 1, 3).unwrap();
        set_row_outline_level(&mut ws, 1, 0).unwrap();

        let row = ws.sheet_data.rows.iter().find(|r| r.r == 1).unwrap();
        assert_eq!(row.outline_level, None);
    }

    #[test]
    fn test_set_row_outline_level_exceeds_max_returns_error() {
        let mut ws = sample_ws();
        let result = set_row_outline_level(&mut ws, 1, 8);
        assert!(result.is_err());
    }

    #[test]
    fn test_set_row_outline_level_row_zero_returns_error() {
        let mut ws = WorksheetXml::default();
        let result = set_row_outline_level(&mut ws, 0, 1);
        assert!(result.is_err());
    }
}
