//! Cell reference conversion utilities.
//!
//! Provides functions for converting between A1-style cell references
//! (e.g. `"A1"`, `"$AB$100"`, `"XFD1048576"`) and 1-based `(col, row)`
//! numeric coordinates.

use crate::error::{Error, Result};
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS};

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
    // Strip absolute-reference markers.
    let cell = cell.replace('$', "");

    if cell.is_empty() {
        return Err(Error::InvalidCellReference(
            "empty cell reference".to_string(),
        ));
    }

    // Split into column letters and row digits.
    let mut col_end = 0;
    for (i, c) in cell.char_indices() {
        if c.is_ascii_alphabetic() {
            col_end = i + c.len_utf8();
        } else {
            break;
        }
    }

    if col_end == 0 {
        return Err(Error::InvalidCellReference(format!(
            "no column letters in '{cell}'"
        )));
    }

    let col_str = &cell[..col_end];
    let row_str = &cell[col_end..];

    if row_str.is_empty() {
        return Err(Error::InvalidCellReference(format!(
            "no row number in '{cell}'"
        )));
    }

    let col = column_name_to_number(col_str)?;

    let row: u32 = row_str
        .parse()
        .map_err(|_| Error::InvalidCellReference(format!("invalid row number in '{cell}'")))?;

    if !(1..=MAX_ROWS).contains(&row) {
        return Err(Error::InvalidRowNumber(row));
    }

    Ok((col, row))
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

// =========================================================================
// Tests
// =========================================================================

#[cfg(test)]
mod tests {
    use super::*;

    // ----- column_name_to_number -----------------------------------------

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

    // ----- column_number_to_name -----------------------------------------

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

    // ----- roundtrip column name <-> number ------------------------------

    #[test]
    fn test_column_roundtrip_all() {
        for n in 1..=MAX_COLUMNS {
            let name = column_number_to_name(n).unwrap();
            let back = column_name_to_number(&name).unwrap();
            assert_eq!(n, back, "roundtrip failed for column {n} (name={name})");
        }
    }

    // ----- cell_name_to_coordinates --------------------------------------

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

    // ----- coordinates_to_cell_name --------------------------------------

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

    // ----- roundtrip cell name <-> coordinates ---------------------------

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
}
