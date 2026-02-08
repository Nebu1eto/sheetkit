//! Merge cell operations.
//!
//! Provides functions for merging and unmerging ranges of cells in a worksheet.

use crate::error::{Error, Result};
use crate::utils::cell_ref::cell_name_to_coordinates;
use sheetkit_xml::worksheet::{MergeCell, MergeCells, WorksheetXml};

/// Parse a range reference like "A1:C3" into ((col1, row1), (col2, row2)) coordinates,
/// both 1-based. Ensures the returned rectangle is normalized so that
/// (col1, row1) is the top-left and (col2, row2) is the bottom-right.
fn parse_range(reference: &str) -> Result<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = reference.split(':').collect();
    if parts.len() != 2 {
        return Err(Error::InvalidCellReference(format!(
            "expected range like 'A1:C3', got '{reference}'"
        )));
    }
    let (c1, r1) = cell_name_to_coordinates(parts[0])?;
    let (c2, r2) = cell_name_to_coordinates(parts[1])?;
    let min_col = c1.min(c2);
    let max_col = c1.max(c2);
    let min_row = r1.min(r2);
    let max_row = r1.max(r2);
    Ok((min_col, min_row, max_col, max_row))
}

/// Check whether two rectangular ranges overlap.
fn ranges_overlap(a: (u32, u32, u32, u32), b: (u32, u32, u32, u32)) -> bool {
    let (a_min_col, a_min_row, a_max_col, a_max_row) = a;
    let (b_min_col, b_min_row, b_max_col, b_max_row) = b;
    a_min_col <= b_max_col
        && a_max_col >= b_min_col
        && a_min_row <= b_max_row
        && a_max_row >= b_min_row
}

/// Merge a range of cells on the given worksheet.
///
/// `top_left` and `bottom_right` are cell references like "A1" and "C3".
/// Returns an error if the new range overlaps with any existing merge region.
pub fn merge_cells(ws: &mut WorksheetXml, top_left: &str, bottom_right: &str) -> Result<()> {
    let (tl_col, tl_row) = cell_name_to_coordinates(top_left)?;
    let (br_col, br_row) = cell_name_to_coordinates(bottom_right)?;

    let min_col = tl_col.min(br_col);
    let max_col = tl_col.max(br_col);
    let min_row = tl_row.min(br_row);
    let max_row = tl_row.max(br_row);
    let new_range = (min_col, min_row, max_col, max_row);

    let reference = format!("{top_left}:{bottom_right}");

    // Check for overlaps with existing merge regions.
    if let Some(ref mc) = ws.merge_cells {
        for existing in &mc.merge_cells {
            let existing_range = parse_range(&existing.reference)?;
            if ranges_overlap(new_range, existing_range) {
                return Err(Error::MergeCellOverlap {
                    new: reference,
                    existing: existing.reference.clone(),
                });
            }
        }
    }

    // Add the merge cell entry.
    let merge_cells = ws.merge_cells.get_or_insert_with(|| MergeCells {
        count: None,
        merge_cells: Vec::new(),
    });
    merge_cells.merge_cells.push(MergeCell { reference });
    merge_cells.count = Some(merge_cells.merge_cells.len() as u32);

    Ok(())
}

/// Remove a specific merge cell range from the worksheet.
///
/// `reference` is the exact range string like "A1:C3" that was previously merged.
/// Returns an error if the range is not found.
pub fn unmerge_cell(ws: &mut WorksheetXml, reference: &str) -> Result<()> {
    let mc = ws
        .merge_cells
        .as_mut()
        .ok_or_else(|| Error::MergeCellNotFound(reference.to_string()))?;

    let initial_len = mc.merge_cells.len();
    mc.merge_cells.retain(|m| m.reference != reference);

    if mc.merge_cells.len() == initial_len {
        return Err(Error::MergeCellNotFound(reference.to_string()));
    }

    if mc.merge_cells.is_empty() {
        ws.merge_cells = None;
    } else {
        mc.count = Some(mc.merge_cells.len() as u32);
    }

    Ok(())
}

/// Get all merge cell references in the worksheet.
///
/// Returns a list of range strings like `["A1:B2", "D1:F3"]`.
pub fn get_merge_cells(ws: &WorksheetXml) -> Vec<String> {
    ws.merge_cells
        .as_ref()
        .map(|mc| mc.merge_cells.iter().map(|m| m.reference.clone()).collect())
        .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn new_ws() -> WorksheetXml {
        WorksheetXml::default()
    }

    #[test]
    fn test_merge_cells_basic() {
        let mut ws = new_ws();
        merge_cells(&mut ws, "A1", "B2").unwrap();
        let merged = get_merge_cells(&ws);
        assert_eq!(merged, vec!["A1:B2"]);
        assert_eq!(ws.merge_cells.as_ref().unwrap().count, Some(1));
    }

    #[test]
    fn test_merge_cells_multiple() {
        let mut ws = new_ws();
        merge_cells(&mut ws, "A1", "B2").unwrap();
        merge_cells(&mut ws, "D1", "F3").unwrap();
        merge_cells(&mut ws, "A5", "C7").unwrap();
        let merged = get_merge_cells(&ws);
        assert_eq!(merged.len(), 3);
        assert_eq!(merged[0], "A1:B2");
        assert_eq!(merged[1], "D1:F3");
        assert_eq!(merged[2], "A5:C7");
        assert_eq!(ws.merge_cells.as_ref().unwrap().count, Some(3));
    }

    #[test]
    fn test_merge_cells_overlap_detection() {
        let mut ws = new_ws();
        merge_cells(&mut ws, "A1", "C3").unwrap();

        // Exact overlap.
        let err = merge_cells(&mut ws, "A1", "C3").unwrap_err();
        assert!(err.to_string().contains("overlaps"));

        // Partial overlap -- B2:D4 overlaps with A1:C3.
        let err = merge_cells(&mut ws, "B2", "D4").unwrap_err();
        assert!(err.to_string().contains("overlaps"));

        // Fully contained -- B2:B2 is inside A1:C3.
        let err = merge_cells(&mut ws, "B2", "B2").unwrap_err();
        assert!(err.to_string().contains("overlaps"));

        // Non-overlapping should succeed.
        merge_cells(&mut ws, "D1", "F3").unwrap();
    }

    #[test]
    fn test_merge_cells_overlap_adjacent_no_overlap() {
        let mut ws = new_ws();
        merge_cells(&mut ws, "A1", "B2").unwrap();
        // C1:D2 is adjacent but does not overlap with A1:B2.
        merge_cells(&mut ws, "C1", "D2").unwrap();
        // A3:B4 is below and does not overlap.
        merge_cells(&mut ws, "A3", "B4").unwrap();
        assert_eq!(get_merge_cells(&ws).len(), 3);
    }

    #[test]
    fn test_unmerge_cell() {
        let mut ws = new_ws();
        merge_cells(&mut ws, "A1", "B2").unwrap();
        merge_cells(&mut ws, "D1", "F3").unwrap();

        unmerge_cell(&mut ws, "A1:B2").unwrap();
        let merged = get_merge_cells(&ws);
        assert_eq!(merged, vec!["D1:F3"]);
        assert_eq!(ws.merge_cells.as_ref().unwrap().count, Some(1));
    }

    #[test]
    fn test_unmerge_cell_last_removes_element() {
        let mut ws = new_ws();
        merge_cells(&mut ws, "A1", "B2").unwrap();
        unmerge_cell(&mut ws, "A1:B2").unwrap();
        assert!(ws.merge_cells.is_none());
        assert!(get_merge_cells(&ws).is_empty());
    }

    #[test]
    fn test_unmerge_cell_not_found() {
        let mut ws = new_ws();
        let err = unmerge_cell(&mut ws, "A1:B2").unwrap_err();
        assert!(err.to_string().contains("not found"));

        // Add one range, then try to unmerge a different range.
        merge_cells(&mut ws, "A1", "B2").unwrap();
        let err = unmerge_cell(&mut ws, "C1:D2").unwrap_err();
        assert!(err.to_string().contains("not found"));
    }

    #[test]
    fn test_get_merge_cells_empty() {
        let ws = new_ws();
        assert!(get_merge_cells(&ws).is_empty());
    }

    #[test]
    fn test_merge_cells_invalid_reference() {
        let mut ws = new_ws();
        let err = merge_cells(&mut ws, "!!!", "B2").unwrap_err();
        assert!(err.to_string().contains("invalid cell reference"));

        let err = merge_cells(&mut ws, "A1", "ZZZ").unwrap_err();
        assert!(err.to_string().contains("no row number"));
    }

    #[test]
    fn test_parse_range_valid() {
        let (c1, r1, c2, r2) = parse_range("A1:C3").unwrap();
        assert_eq!((c1, r1, c2, r2), (1, 1, 3, 3));
    }

    #[test]
    fn test_parse_range_reversed() {
        // Even if cells are given in reversed order, we normalize.
        let (c1, r1, c2, r2) = parse_range("C3:A1").unwrap();
        assert_eq!((c1, r1, c2, r2), (1, 1, 3, 3));
    }

    #[test]
    fn test_parse_range_invalid() {
        assert!(parse_range("A1").is_err());
        assert!(parse_range("A1:B2:C3").is_err());
        assert!(parse_range("").is_err());
    }

    #[test]
    fn test_ranges_overlap_function() {
        // Overlapping rectangles.
        assert!(ranges_overlap((1, 1, 3, 3), (2, 2, 4, 4)));
        // Identical.
        assert!(ranges_overlap((1, 1, 3, 3), (1, 1, 3, 3)));
        // Contained.
        assert!(ranges_overlap((1, 1, 5, 5), (2, 2, 3, 3)));
        // Adjacent horizontally -- no overlap.
        assert!(!ranges_overlap((1, 1, 2, 2), (3, 1, 4, 2)));
        // Adjacent vertically -- no overlap.
        assert!(!ranges_overlap((1, 1, 2, 2), (1, 3, 2, 4)));
        // Completely disjoint.
        assert!(!ranges_overlap((1, 1, 2, 2), (5, 5, 6, 6)));
    }
}
