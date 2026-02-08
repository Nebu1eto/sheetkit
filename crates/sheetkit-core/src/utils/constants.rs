//! Excel limit constants and default values.
//!
//! These constants mirror the hard limits enforced by Microsoft Excel 2007+
//! (OOXML / `.xlsx` format).

/// Maximum number of columns (XFD = 16 384 = 2^14).
pub const MAX_COLUMNS: u32 = 16_384;

/// Maximum number of rows (1 048 576 = 2^20).
pub const MAX_ROWS: u32 = 1_048_576;

/// Maximum number of cell styles that can be stored in a workbook.
pub const MAX_CELL_STYLES: usize = 65_430;

/// Maximum column width in character-width units.
pub const MAX_COLUMN_WIDTH: f64 = 255.0;

/// Maximum row height in points.
pub const MAX_ROW_HEIGHT: f64 = 409.0;

/// Maximum font size in points.
pub const MAX_FONT_SIZE: f64 = 409.0;

/// Maximum length (in characters) of a sheet name.
pub const MAX_SHEET_NAME_LENGTH: usize = 31;

/// Maximum number of characters that a single cell can contain.
pub const MAX_CELL_CHARS: usize = 32_767;

/// Characters that are not allowed in Excel sheet names.
pub const SHEET_NAME_INVALID_CHARS: &[char] = &[':', '\\', '/', '?', '*', '[', ']'];

/// Default column width used when no explicit width is set (character-width units).
pub const DEFAULT_COL_WIDTH: f64 = 9.140625;

/// Default row height in points.
pub const DEFAULT_ROW_HEIGHT: f64 = 15.0;

/// Chunk size used by the streaming writer (16 MiB).
pub const STREAM_CHUNK_SIZE: usize = 16 * 1024 * 1024;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_max_columns() {
        assert_eq!(MAX_COLUMNS, 16_384);
    }

    #[test]
    fn test_max_rows() {
        assert_eq!(MAX_ROWS, 1_048_576);
    }

    #[test]
    fn test_max_cell_styles() {
        assert_eq!(MAX_CELL_STYLES, 65_430);
    }

    #[test]
    fn test_max_column_width() {
        assert!((MAX_COLUMN_WIDTH - 255.0).abs() < f64::EPSILON);
    }

    #[test]
    fn test_max_row_height() {
        assert!((MAX_ROW_HEIGHT - 409.0).abs() < f64::EPSILON);
    }

    #[test]
    fn test_max_font_size() {
        assert!((MAX_FONT_SIZE - 409.0).abs() < f64::EPSILON);
    }

    #[test]
    fn test_max_sheet_name_length() {
        assert_eq!(MAX_SHEET_NAME_LENGTH, 31);
    }

    #[test]
    fn test_max_cell_chars() {
        assert_eq!(MAX_CELL_CHARS, 32_767);
    }

    #[test]
    fn test_default_col_width() {
        assert!((DEFAULT_COL_WIDTH - 9.140625).abs() < f64::EPSILON);
    }

    #[test]
    fn test_default_row_height() {
        assert!((DEFAULT_ROW_HEIGHT - 15.0).abs() < f64::EPSILON);
    }

    #[test]
    fn test_stream_chunk_size() {
        assert_eq!(STREAM_CHUNK_SIZE, 16 * 1024 * 1024);
    }

    #[test]
    fn test_sheet_name_invalid_chars() {
        assert_eq!(SHEET_NAME_INVALID_CHARS.len(), 7);
        assert!(SHEET_NAME_INVALID_CHARS.contains(&':'));
        assert!(SHEET_NAME_INVALID_CHARS.contains(&'\\'));
        assert!(SHEET_NAME_INVALID_CHARS.contains(&'/'));
        assert!(SHEET_NAME_INVALID_CHARS.contains(&'?'));
        assert!(SHEET_NAME_INVALID_CHARS.contains(&'*'));
        assert!(SHEET_NAME_INVALID_CHARS.contains(&'['));
        assert!(SHEET_NAME_INVALID_CHARS.contains(&']'));
    }
}
