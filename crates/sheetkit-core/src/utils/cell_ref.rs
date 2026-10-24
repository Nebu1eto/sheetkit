//! Cell reference conversion utilities.
//!
//! Provides functions for converting between A1-style cell references
//! (e.g. `"A1"`, `"$AB$100"`, `"XFD1048576"`) and 1-based `(col, row)`
//! numeric coordinates.

use crate::error::{Error, Result};
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS};
use sheetkit_xml::cell_ref::parse_a1_reference;
use sheetkit_xml::worksheet::CompactCellRef;

/// Create a validated compact A1 reference from 1-based coordinates.
pub(crate) fn compact_cell_ref(col: u32, row: u32) -> Result<CompactCellRef> {
    CompactCellRef::from_coordinates(col, row)
        .map_err(|error| Error::InvalidCellReference(error.to_string()))
}

/// Decode a [`CompactCellRef`] into `(column, row)`.
///
/// Every constructor, including deserialization, admits only canonical,
/// in-range A1 references, so syntax and range checks are not repeated.
fn compact_coordinates(reference: &CompactCellRef) -> (u32, u32) {
    let mut column = 0u32;
    let mut row = 0u32;
    for &byte in reference.as_str().as_bytes() {
        if byte.is_ascii_uppercase() {
            column = column * 26 + u32::from(byte - b'A') + 1;
        } else {
            row = row * 10 + u32::from(byte - b'0');
        }
    }
    debug_assert!(parse_a1_reference(reference.as_str()).is_ok_and(|parsed| {
        parsed.is_canonical() && parsed.column() == column && parsed.row() == row
    }));
    (column, row)
}

/// Validate worksheet cell structure, then sort rows and cells, cache cell
/// columns for binary search, and release excess cell capacity.
pub(crate) fn validate_and_index_worksheet(
    worksheet: &mut sheetkit_xml::worksheet::WorksheetXml,
) -> Result<()> {
    const COLUMN_WORDS: usize = (MAX_COLUMNS as usize).div_ceil(u64::BITS as usize);
    let mut seen_rows = vec![0u64; (MAX_ROWS as usize).div_ceil(u64::BITS as usize)];
    // Shared across rows; only rows whose columns are not strictly
    // increasing use it, and they clear the words they set.
    let mut seen_columns = [0u64; COLUMN_WORDS];
    let column_bit = |column: u32| {
        let index = (column - 1) as usize;
        (
            index / u64::BITS as usize,
            1u64 << (index % u64::BITS as usize),
        )
    };

    for row in &mut worksheet.sheet_data.rows {
        if !(1..=MAX_ROWS).contains(&row.r) {
            return Err(Error::InvalidRowNumber(row.r));
        }

        let row_index = (row.r - 1) as usize;
        let row_word = row_index / u64::BITS as usize;
        let row_mask = 1u64 << (row_index % u64::BITS as usize);
        if seen_rows[row_word] & row_mask != 0 {
            return Err(Error::InvalidCellReference(format!(
                "duplicate worksheet row {}",
                row.r
            )));
        }
        seen_rows[row_word] |= row_mask;

        let mut last_column = 0u32;
        let mut uses_bitset = false;
        for index in 0..row.cells.len() {
            let cell = &row.cells[index];
            let (column, cell_row) = compact_coordinates(&cell.r);
            if cell_row != row.r {
                return Err(Error::InvalidCellReference(format!(
                    "cell '{}' does not belong to worksheet row {}",
                    cell.r, row.r
                )));
            }

            if !uses_bitset && column <= last_column {
                // Strictly increasing columns cannot repeat. On the first
                // out-of-order cell, record every earlier column of the row.
                uses_bitset = true;
                for earlier in &row.cells[..index] {
                    let (word, mask) = column_bit(earlier.col);
                    seen_columns[word] |= mask;
                }
            }
            if uses_bitset {
                let (word, mask) = column_bit(column);
                if seen_columns[word] & mask != 0 {
                    return Err(Error::InvalidCellReference(format!(
                        "duplicate worksheet cell reference '{}'",
                        row.cells[index].r
                    )));
                }
                seen_columns[word] |= mask;
            }
            last_column = column;
            // Callers discard the worksheet on error, so caching the column
            // before the whole sheet is validated cannot leak partial state.
            row.cells[index].col = column;
        }
        if uses_bitset {
            for cell in &row.cells {
                seen_columns[column_bit(cell.col).0] = 0;
            }
        }
    }

    worksheet.sheet_data.rows.sort_unstable_by_key(|row| row.r);
    for row in &mut worksheet.sheet_data.rows {
        row.cells.sort_unstable_by_key(|cell| cell.col);
        row.cells.shrink_to_fit();
    }
    Ok(())
}

/// Convert a column name (e.g. `"A"`, `"Z"`, `"AA"`, `"XFD"`) to a 1-based
/// column number.
///
/// # Examples
///
/// ```
/// use sheetkit_core::utils::cell_ref::column_name_to_number;
///
/// assert_eq!(column_name_to_number("A").unwrap(), 1);
/// assert_eq!(column_name_to_number("Z").unwrap(), 26);
/// assert_eq!(column_name_to_number("AA").unwrap(), 27);
/// assert_eq!(column_name_to_number("XFD").unwrap(), 16384);
/// ```
pub fn column_name_to_number(name: &str) -> Result<u32> {
    if name.is_empty() {
        return Err(Error::InvalidCellReference("empty column name".to_string()));
    }

    let mut result: u32 = 0;

    for c in name.chars() {
        if !c.is_ascii_alphabetic() {
            return Err(Error::InvalidCellReference(format!(
                "non-alphabetic character in column name: '{c}'"
            )));
        }

        let digit = (c.to_ascii_uppercase() as u32) - ('A' as u32) + 1;

        result = result
            .checked_mul(26)
            .and_then(|r| r.checked_add(digit))
            .ok_or(Error::InvalidColumnNumber(0))?;
    }

    if result > MAX_COLUMNS {
        return Err(Error::InvalidColumnNumber(result));
    }

    Ok(result)
}

/// Convert a 1-based column number to its letter name.
///
/// # Examples
///
/// ```
/// use sheetkit_core::utils::cell_ref::column_number_to_name;
///
/// assert_eq!(column_number_to_name(1).unwrap(), "A");
/// assert_eq!(column_number_to_name(26).unwrap(), "Z");
/// assert_eq!(column_number_to_name(27).unwrap(), "AA");
/// assert_eq!(column_number_to_name(16384).unwrap(), "XFD");
/// ```
pub fn column_number_to_name(num: u32) -> Result<String> {
    if !(1..=MAX_COLUMNS).contains(&num) {
        return Err(Error::InvalidColumnNumber(num));
    }

    let mut col = num;
    let mut result = String::with_capacity(3);

    while col > 0 {
        col -= 1; // adjust to 0-indexed
        let remainder = (col % 26) as u8;
        result.insert(0, (b'A' + remainder) as char);
        col /= 26;
    }

    Ok(result)
}

/// Parse an A1-style cell reference into `(col, row)` coordinates (both
/// 1-based).
///
/// Absolute-reference markers (`$`) are stripped before parsing.
///
/// # Examples
///
/// ```
/// use sheetkit_core::utils::cell_ref::cell_name_to_coordinates;
///
/// assert_eq!(cell_name_to_coordinates("A1").unwrap(), (1, 1));
/// assert_eq!(cell_name_to_coordinates("$B$2").unwrap(), (2, 2));
/// assert_eq!(cell_name_to_coordinates("AA100").unwrap(), (27, 100));
/// ```
pub fn cell_name_to_coordinates(cell: &str) -> Result<(u32, u32)> {
    let parsed =
        parse_a1_reference(cell).map_err(|error| Error::InvalidCellReference(error.to_string()))?;
    Ok((parsed.column(), parsed.row()))
}

/// Convert 1-based `(col, row)` coordinates to an A1-style cell reference.
///
/// # Examples
///
/// ```
/// use sheetkit_core::utils::cell_ref::coordinates_to_cell_name;
///
/// assert_eq!(coordinates_to_cell_name(1, 1).unwrap(), "A1");
/// assert_eq!(coordinates_to_cell_name(27, 100).unwrap(), "AA100");
/// ```
pub fn coordinates_to_cell_name(col: u32, row: u32) -> Result<String> {
    if !(1..=MAX_COLUMNS).contains(&col) {
        return Err(Error::InvalidColumnNumber(col));
    }
    if !(1..=MAX_ROWS).contains(&row) {
        return Err(Error::InvalidRowNumber(row));
    }

    let col_name = column_number_to_name(col)?;
    Ok(format!("{col_name}{row}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_name_a() {
        assert_eq!(column_name_to_number("A").unwrap(), 1);
    }

    #[test]
    fn test_column_name_z() {
        assert_eq!(column_name_to_number("Z").unwrap(), 26);
    }

    #[test]
    fn test_column_name_aa() {
        assert_eq!(column_name_to_number("AA").unwrap(), 27);
    }

    #[test]
    fn test_column_name_az() {
        assert_eq!(column_name_to_number("AZ").unwrap(), 52);
    }

    #[test]
    fn test_column_name_ba() {
        assert_eq!(column_name_to_number("BA").unwrap(), 53);
    }

    #[test]
    fn test_column_name_xfd() {
        assert_eq!(column_name_to_number("XFD").unwrap(), 16384);
    }

    #[test]
    fn test_column_name_lowercase() {
        assert_eq!(column_name_to_number("a").unwrap(), 1);
        assert_eq!(column_name_to_number("xfd").unwrap(), 16384);
    }

    #[test]
    fn test_column_name_xfe_out_of_range() {
        assert!(column_name_to_number("XFE").is_err());
    }

    #[test]
    fn test_column_name_empty() {
        assert!(column_name_to_number("").is_err());
    }

    #[test]
    fn test_column_name_with_digit() {
        assert!(column_name_to_number("A1").is_err());
    }

    #[test]
    fn test_column_number_1() {
        assert_eq!(column_number_to_name(1).unwrap(), "A");
    }

    #[test]
    fn test_column_number_26() {
        assert_eq!(column_number_to_name(26).unwrap(), "Z");
    }

    #[test]
    fn test_column_number_27() {
        assert_eq!(column_number_to_name(27).unwrap(), "AA");
    }

    #[test]
    fn test_column_number_52() {
        assert_eq!(column_number_to_name(52).unwrap(), "AZ");
    }

    #[test]
    fn test_column_number_53() {
        assert_eq!(column_number_to_name(53).unwrap(), "BA");
    }

    #[test]
    fn test_column_number_16384() {
        assert_eq!(column_number_to_name(16384).unwrap(), "XFD");
    }

    #[test]
    fn test_column_number_0_err() {
        assert!(column_number_to_name(0).is_err());
    }

    #[test]
    fn test_column_number_16385_err() {
        assert!(column_number_to_name(16385).is_err());
    }

    #[test]
    fn test_column_roundtrip_all() {
        for n in 1..=MAX_COLUMNS {
            let name = column_number_to_name(n).unwrap();
            let back = column_name_to_number(&name).unwrap();
            assert_eq!(n, back, "roundtrip failed for column {n} (name={name})");
        }
    }

    #[test]
    fn test_cell_a1() {
        assert_eq!(cell_name_to_coordinates("A1").unwrap(), (1, 1));
    }

    #[test]
    fn test_cell_z10() {
        assert_eq!(cell_name_to_coordinates("Z10").unwrap(), (26, 10));
    }

    #[test]
    fn test_cell_aa1() {
        assert_eq!(cell_name_to_coordinates("AA1").unwrap(), (27, 1));
    }

    #[test]
    fn test_cell_absolute_a1() {
        assert_eq!(cell_name_to_coordinates("$A$1").unwrap(), (1, 1));
    }

    #[test]
    fn test_cell_absolute_ab100() {
        assert_eq!(cell_name_to_coordinates("$AB$100").unwrap(), (28, 100));
    }

    #[test]
    fn test_cell_mixed_absolute() {
        assert_eq!(cell_name_to_coordinates("$A1").unwrap(), (1, 1));
        assert_eq!(cell_name_to_coordinates("A$1").unwrap(), (1, 1));
    }

    #[test]
    fn test_cell_max() {
        assert_eq!(
            cell_name_to_coordinates("XFD1048576").unwrap(),
            (16384, 1_048_576)
        );
    }

    #[test]
    fn test_cell_empty_err() {
        assert!(cell_name_to_coordinates("").is_err());
    }

    #[test]
    fn test_cell_only_letters_err() {
        assert!(cell_name_to_coordinates("ABC").is_err());
    }

    #[test]
    fn test_cell_only_digits_err() {
        assert!(cell_name_to_coordinates("123").is_err());
    }

    #[test]
    fn test_cell_row_zero_err() {
        assert!(cell_name_to_coordinates("A0").is_err());
    }

    #[test]
    fn test_cell_row_too_large_err() {
        assert!(cell_name_to_coordinates("A1048577").is_err());
    }

    #[test]
    fn test_cell_col_too_large_err() {
        assert!(cell_name_to_coordinates("XFE1").is_err());
    }

    #[test]
    fn test_cell_rejects_misplaced_absolute_markers() {
        for reference in ["A$1$", "$A1$", "$$A1", "A$$1"] {
            assert!(
                cell_name_to_coordinates(reference).is_err(),
                "accepted {reference}"
            );
        }
    }

    #[test]
    fn test_cell_rejects_trailing_junk_overflow_and_utf8() {
        for reference in ["A1junk", "A42949672960", "A1é"] {
            assert!(
                cell_name_to_coordinates(reference).is_err(),
                "accepted {reference}"
            );
        }
    }

    #[test]
    fn test_coords_1_1() {
        assert_eq!(coordinates_to_cell_name(1, 1).unwrap(), "A1");
    }

    #[test]
    fn test_coords_27_100() {
        assert_eq!(coordinates_to_cell_name(27, 100).unwrap(), "AA100");
    }

    #[test]
    fn test_coords_max() {
        assert_eq!(
            coordinates_to_cell_name(16384, 1_048_576).unwrap(),
            "XFD1048576"
        );
    }

    #[test]
    fn test_coords_col_0_err() {
        assert!(coordinates_to_cell_name(0, 1).is_err());
    }

    #[test]
    fn test_coords_row_0_err() {
        assert!(coordinates_to_cell_name(1, 0).is_err());
    }

    #[test]
    fn test_coords_col_too_large_err() {
        assert!(coordinates_to_cell_name(16385, 1).is_err());
    }

    #[test]
    fn test_coords_row_too_large_err() {
        assert!(coordinates_to_cell_name(1, 1_048_577).is_err());
    }

    #[test]
    fn test_cell_roundtrip() {
        let cases = vec![
            (1, 1, "A1"),
            (26, 1, "Z1"),
            (27, 1, "AA1"),
            (52, 10, "AZ10"),
            (16384, 1_048_576, "XFD1048576"),
        ];

        for (col, row, expected_name) in cases {
            let name = coordinates_to_cell_name(col, row).unwrap();
            assert_eq!(name, expected_name);

            let (c, r) = cell_name_to_coordinates(&name).unwrap();
            assert_eq!((c, r), (col, row));
        }
    }

    fn worksheet_cell(reference: &str) -> sheetkit_xml::worksheet::Cell {
        sheetkit_xml::worksheet::Cell {
            r: CompactCellRef::new(reference).unwrap(),
            col: 0,
            s: None,
            t: sheetkit_xml::worksheet::CellTypeTag::None,
            v: None,
            f: None,
            is: None,
        }
    }

    fn worksheet_row(row: u32, references: &[&str]) -> sheetkit_xml::worksheet::Row {
        sheetkit_xml::worksheet::Row {
            r: row,
            spans: None,
            s: None,
            custom_format: None,
            ht: None,
            custom_height: None,
            hidden: None,
            outline_level: None,
            cells: references
                .iter()
                .map(|reference| worksheet_cell(reference))
                .collect(),
        }
    }

    #[test]
    fn worksheet_indexing_rejects_row_mismatch_and_duplicate_coordinates() {
        let mut mismatch = sheetkit_xml::worksheet::WorksheetXml::default();
        mismatch.sheet_data.rows.push(worksheet_row(2, &["A1"]));
        assert!(validate_and_index_worksheet(&mut mismatch).is_err());

        let mut duplicate_cells = sheetkit_xml::worksheet::WorksheetXml::default();
        duplicate_cells
            .sheet_data
            .rows
            .push(worksheet_row(1, &["A1", "A1"]));
        assert!(validate_and_index_worksheet(&mut duplicate_cells).is_err());

        let mut duplicate_rows = sheetkit_xml::worksheet::WorksheetXml::default();
        duplicate_rows
            .sheet_data
            .rows
            .extend([worksheet_row(1, &["A1"]), worksheet_row(1, &["B1"])]);
        assert!(validate_and_index_worksheet(&mut duplicate_rows).is_err());
    }

    #[test]
    fn worksheet_indexing_detects_out_of_order_duplicates_without_leaking_rows() {
        let duplicate_message = |references: &[&str]| {
            let mut worksheet = sheetkit_xml::worksheet::WorksheetXml::default();
            worksheet.sheet_data.rows.push(worksheet_row(1, references));
            validate_and_index_worksheet(&mut worksheet)
                .unwrap_err()
                .to_string()
        };
        assert!(duplicate_message(&["A1", "C1", "B1", "A1"]).contains("duplicate"));
        assert!(duplicate_message(&["C1", "A1", "C1"]).contains("duplicate"));
        // The first error in document order wins over later invalid cells.
        assert!(duplicate_message(&["B1", "B1", "C2"]).contains("duplicate"));
        assert!(duplicate_message(&["B1", "C2", "B1"]).contains("does not belong"));

        // Columns recorded for one unordered row must not leak into the next.
        let mut worksheet = sheetkit_xml::worksheet::WorksheetXml::default();
        worksheet.sheet_data.rows.extend([
            worksheet_row(1, &["B1", "A1"]),
            worksheet_row(2, &["B2", "A2"]),
            worksheet_row(3, &["A3", "B3"]),
        ]);
        validate_and_index_worksheet(&mut worksheet).unwrap();
        let columns: Vec<Vec<u32>> = worksheet
            .sheet_data
            .rows
            .iter()
            .map(|row| row.cells.iter().map(|cell| cell.col).collect())
            .collect();
        assert_eq!(columns, [[1, 2], [1, 2], [1, 2]]);
    }

    #[test]
    fn worksheet_indexing_validates_before_sorting_and_populates_columns() {
        let mut worksheet = sheetkit_xml::worksheet::WorksheetXml::default();
        worksheet
            .sheet_data
            .rows
            .extend([worksheet_row(2, &["B2", "A2"]), worksheet_row(1, &["C1"])]);

        validate_and_index_worksheet(&mut worksheet).unwrap();

        assert_eq!(
            worksheet
                .sheet_data
                .rows
                .iter()
                .map(|row| row.r)
                .collect::<Vec<_>>(),
            [1, 2]
        );
        assert_eq!(
            worksheet.sheet_data.rows[1]
                .cells
                .iter()
                .map(|cell| (cell.r.as_str(), cell.col))
                .collect::<Vec<_>>(),
            [("A2", 1), ("B2", 2)]
        );
    }
}
