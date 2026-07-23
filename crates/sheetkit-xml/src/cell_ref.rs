use std::fmt;

use serde::de::Deserializer;
use serde::ser::Serializer;
use serde::{Deserialize, Serialize};
use thiserror::Error;

/// Maximum number of columns supported by an Excel worksheet.
pub const MAX_COLUMNS: u32 = 16_384;
/// Maximum number of rows supported by an Excel worksheet.
pub const MAX_ROWS: u32 = 1_048_576;

/// Error returned for an invalid or noncanonical A1 cell reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
pub enum CellRefError {
    /// The reference is empty.
    #[error("cell reference is empty")]
    Empty,
    /// The reference contains non-ASCII input.
    #[error("cell reference must contain ASCII characters only")]
    NonAscii,
    /// The reference does not match A1 syntax.
    #[error("cell reference has invalid A1 syntax")]
    InvalidSyntax,
    /// The reference denotes a column outside Excel's worksheet grid.
    #[error("cell reference column is outside 1..=16384")]
    ColumnOutOfRange,
    /// The reference denotes a row outside Excel's worksheet grid.
    #[error("cell reference row is outside 1..=1048576")]
    RowOutOfRange,
    /// Worksheet cell references must use one canonical spelling per cell.
    #[error("worksheet cell reference is not canonical")]
    NonCanonical,
}

/// A validated A1 reference and its 1-based coordinates.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParsedCellRef {
    column: u32,
    row: u32,
    absolute_column: bool,
    absolute_row: bool,
    canonical: bool,
}

impl ParsedCellRef {
    /// Returns the 1-based column number.
    pub const fn column(self) -> u32 {
        self.column
    }

    /// Returns the 1-based row number.
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Returns whether the column has an absolute marker.
    pub const fn absolute_column(self) -> bool {
        self.absolute_column
    }

    /// Returns whether the row has an absolute marker.
    pub const fn absolute_row(self) -> bool {
        self.absolute_row
    }

    /// Returns whether the input is the canonical worksheet spelling.
    pub const fn is_canonical(self) -> bool {
        self.canonical
    }
}

/// Parses a complete A1 reference without allocation.
///
/// Optional absolute markers are accepted only in their grammar positions.
/// Callers that persist worksheet cell references must additionally require
/// [`ParsedCellRef::is_canonical`].
pub fn parse_a1_reference(input: &str) -> Result<ParsedCellRef, CellRefError> {
    if input.is_empty() {
        return Err(CellRefError::Empty);
    }
    if !input.is_ascii() {
        return Err(CellRefError::NonAscii);
    }

    let bytes = input.as_bytes();
    let mut index = 0usize;
    let absolute_column = bytes.get(index) == Some(&b'$');
    if absolute_column {
        index += 1;
    }

    let column_start = index;
    let mut column = 0u32;
    let mut uppercase_column = true;
    while let Some(byte) = bytes.get(index).copied() {
        if !byte.is_ascii_alphabetic() {
            break;
        }
        uppercase_column &= byte.is_ascii_uppercase();
        let digit = u32::from(byte.to_ascii_uppercase() - b'A') + 1;
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add(digit))
            .ok_or(CellRefError::ColumnOutOfRange)?;
        if column > MAX_COLUMNS {
            return Err(CellRefError::ColumnOutOfRange);
        }
        index += 1;
    }
    if index == column_start {
        return Err(CellRefError::InvalidSyntax);
    }

    let absolute_row = bytes.get(index) == Some(&b'$');
    if absolute_row {
        index += 1;
    }

    let row_start = index;
    let mut row = 0u32;
    while let Some(byte) = bytes.get(index).copied() {
        if !byte.is_ascii_digit() {
            break;
        }
        row = row
            .checked_mul(10)
            .and_then(|value| value.checked_add(u32::from(byte - b'0')))
            .ok_or(CellRefError::RowOutOfRange)?;
        if row > MAX_ROWS {
            return Err(CellRefError::RowOutOfRange);
        }
        index += 1;
    }
    if index == row_start || index != bytes.len() {
        return Err(CellRefError::InvalidSyntax);
    }
    if row == 0 {
        return Err(CellRefError::RowOutOfRange);
    }

    let canonical =
        !absolute_column && !absolute_row && uppercase_column && bytes[row_start] != b'0';
    Ok(ParsedCellRef {
        column,
        row,
        absolute_column,
        absolute_row,
        canonical,
    })
}

/// Inline canonical worksheet cell reference such as `A1` or `XFD1048576`.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct CompactCellRef {
    buf: [u8; 10],
    len: u8,
}

impl CompactCellRef {
    /// Creates a validated canonical worksheet cell reference.
    pub fn new(input: &str) -> Result<Self, CellRefError> {
        let parsed = parse_a1_reference(input)?;
        if !parsed.is_canonical() {
            return Err(CellRefError::NonCanonical);
        }
        let mut buf = [0u8; 10];
        buf[..input.len()].copy_from_slice(input.as_bytes());
        Ok(Self {
            buf,
            len: input.len() as u8,
        })
    }

    /// Returns the canonical reference as a string slice.
    pub fn as_str(&self) -> &str {
        // Safety: constructors and deserialization accept canonical ASCII only.
        unsafe { std::str::from_utf8_unchecked(&self.buf[..usize::from(self.len)]) }
    }

    /// Creates a canonical reference from checked 1-based coordinates.
    pub fn from_coordinates(column: u32, row: u32) -> Result<Self, CellRefError> {
        if !(1..=MAX_COLUMNS).contains(&column) {
            return Err(CellRefError::ColumnOutOfRange);
        }
        if !(1..=MAX_ROWS).contains(&row) {
            return Err(CellRefError::RowOutOfRange);
        }

        let mut buf = [0u8; 10];
        let mut position = 0usize;
        let mut column_buf = [0u8; 3];
        let mut column_len = 0usize;
        let mut remaining_column = column;
        while remaining_column > 0 {
            remaining_column -= 1;
            column_buf[column_len] = b'A' + (remaining_column % 26) as u8;
            column_len += 1;
            remaining_column /= 26;
        }
        for index in (0..column_len).rev() {
            buf[position] = column_buf[index];
            position += 1;
        }

        let mut row_buf = [0u8; 7];
        let mut row_len = 0usize;
        let mut remaining_row = row;
        while remaining_row > 0 {
            row_buf[row_len] = b'0' + (remaining_row % 10) as u8;
            row_len += 1;
            remaining_row /= 10;
        }
        for index in (0..row_len).rev() {
            buf[position] = row_buf[index];
            position += 1;
        }

        Ok(Self {
            buf,
            len: position as u8,
        })
    }
}

impl fmt::Display for CompactCellRef {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl fmt::Debug for CompactCellRef {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "CompactCellRef(\"{}\")", self.as_str())
    }
}

impl TryFrom<&str> for CompactCellRef {
    type Error = CellRefError;

    fn try_from(input: &str) -> Result<Self, Self::Error> {
        Self::new(input)
    }
}

impl TryFrom<String> for CompactCellRef {
    type Error = CellRefError;

    fn try_from(input: String) -> Result<Self, Self::Error> {
        Self::new(&input)
    }
}

impl AsRef<str> for CompactCellRef {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl PartialEq<&str> for CompactCellRef {
    fn eq(&self, other: &&str) -> bool {
        self.as_str() == *other
    }
}

impl PartialEq<str> for CompactCellRef {
    fn eq(&self, other: &str) -> bool {
        self.as_str() == other
    }
}

impl Serialize for CompactCellRef {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        serializer.serialize_str(self.as_str())
    }
}

impl<'de> Deserialize<'de> for CompactCellRef {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        struct CompactCellRefVisitor;

        impl serde::de::Visitor<'_> for CompactCellRefVisitor {
            type Value = CompactCellRef;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("a canonical A1 cell reference")
            }

            fn visit_str<E: serde::de::Error>(self, value: &str) -> Result<CompactCellRef, E> {
                CompactCellRef::new(value).map_err(E::custom)
            }

            fn visit_string<E: serde::de::Error>(self, value: String) -> Result<CompactCellRef, E> {
                self.visit_str(&value)
            }
        }

        deserializer.deserialize_str(CompactCellRefVisitor)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parser_accepts_excel_boundaries_and_absolute_core_refs() {
        let min = parse_a1_reference("A1").unwrap();
        assert_eq!((min.column(), min.row()), (1, 1));

        let max = parse_a1_reference("$XFD$1048576").unwrap();
        assert_eq!((max.column(), max.row()), (16_384, 1_048_576));
    }

    #[test]
    fn parser_rejects_bounds_zero_overflow_utf8_and_trailing_junk() {
        for invalid in [
            "",
            "A0",
            "XFE1",
            "A1048577",
            "A42949672960",
            "A1junk",
            "A1$",
            "$$A1",
            "한글A1",
        ] {
            assert!(parse_a1_reference(invalid).is_err(), "accepted {invalid:?}");
        }
    }

    #[test]
    fn compact_ref_rejects_aliases_and_arbitrary_invalid_strings_without_panic() {
        for invalid in ["a1", "A01", "$A$1", "A1junk", "한글", "ABCDEFGHIJK"] {
            assert!(
                CompactCellRef::new(invalid).is_err(),
                "accepted {invalid:?}"
            );
        }
        assert_eq!(
            CompactCellRef::new("XFD1048576").unwrap().as_str(),
            "XFD1048576"
        );
    }

    #[test]
    fn parser_and_compact_ref_never_panic_for_varied_utf8() {
        let alphabet = ['A', 'z', '$', '0', '9', '_', '/', '\0', 'é', '한', '💾'];
        let mut state = 0x9e37_79b9u32;
        for length in 0..=128 {
            let mut input = String::new();
            for _ in 0..length {
                state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
                input.push(alphabet[state as usize % alphabet.len()]);
            }
            let _ = parse_a1_reference(&input);
            let _ = CompactCellRef::new(&input);
        }

        let oversized_digits = "0".repeat(16 * 1024);
        assert!(parse_a1_reference(&oversized_digits).is_err());
        assert!(CompactCellRef::new(&oversized_digits).is_err());
    }

    #[test]
    fn compact_ref_coordinates_are_fallible_at_exact_boundaries() {
        assert_eq!(
            CompactCellRef::from_coordinates(16_384, 1_048_576)
                .unwrap()
                .as_str(),
            "XFD1048576"
        );
        for (column, row) in [(0, 1), (1, 0), (16_385, 1), (1, 1_048_577)] {
            assert!(CompactCellRef::from_coordinates(column, row).is_err());
        }
    }

    #[test]
    fn compact_ref_deserialization_rejects_noncanonical_aliases() {
        for invalid in ["\"a1\"", "\"A01\"", "\"$A$1\"", "\"A1junk\""] {
            assert!(quick_xml::de::from_str::<CompactCellRef>(invalid).is_err());
        }
    }
}
