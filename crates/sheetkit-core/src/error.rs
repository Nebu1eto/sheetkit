//! Error types for the SheetKit core library.
//!
//! Provides a comprehensive [`Error`] enum covering all failure modes
//! encountered when reading, writing, and manipulating Excel workbooks.

use thiserror::Error;

/// The top-level error type for SheetKit.
#[derive(Error, Debug)]
pub enum Error {
    /// The given string is not a valid A1-style cell reference.
    #[error("invalid cell reference: {0}")]
    InvalidCellReference(String),

    /// The row number is out of the allowed range (1..=1_048_576).
    #[error("invalid row number: {0}")]
    InvalidRowNumber(u32),

    /// The column number is out of the allowed range (1..=16_384).
    #[error("invalid column number: {0}")]
    InvalidColumnNumber(u32),

    /// No sheet with the given name exists in the workbook.
    #[error("sheet '{name}' does not exist")]
    SheetNotFound { name: String },

    /// A sheet with the given name already exists.
    #[error("sheet '{name}' already exists")]
    SheetAlreadyExists { name: String },

    /// The sheet name violates Excel naming rules.
    #[error("invalid sheet name: {0}")]
    InvalidSheetName(String),

    /// An underlying I/O error.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    /// An error originating from the ZIP layer.
    #[error("ZIP error: {0}")]
    Zip(String),

    /// An error encountered while parsing XML.
    #[error("XML parse error: {0}")]
    XmlParse(String),

    /// An error encountered while deserializing XML into typed structures.
    #[error("XML deserialization error: {0}")]
    XmlDeserialize(String),

    /// Column width exceeds the allowed maximum (255).
    #[error("column width {width} exceeds maximum {max}")]
    ColumnWidthExceeded { width: f64, max: f64 },

    /// Row height exceeds the allowed maximum (409).
    #[error("row height {height} exceeds maximum {max}")]
    RowHeightExceeded { height: f64, max: f64 },

    /// A cell value exceeds the maximum character limit.
    #[error("cell value too long: {length} characters (max {max})")]
    CellValueTooLong { length: usize, max: usize },

    /// The style ID was not found in the stylesheet.
    #[error("style not found: {id}")]
    StyleNotFound { id: u32 },

    /// Too many cell styles have been registered.
    #[error("cell styles exceeded maximum ({max})")]
    CellStylesExceeded { max: usize },

    /// A row has already been written; rows must be written in ascending order.
    #[error("row {row} has already been written (must write rows in ascending order)")]
    StreamRowAlreadyWritten { row: u32 },

    /// The stream writer has already been finished.
    #[error("stream writer already finished")]
    StreamAlreadyFinished,

    /// Column widths cannot be set after rows have been written.
    #[error("cannot set column width after rows have been written")]
    StreamColumnsAfterRows,

    /// Merge cell ranges overlap.
    #[error("merge cell range '{new}' overlaps with existing range '{existing}'")]
    MergeCellOverlap { new: String, existing: String },

    /// The specified merge cell range was not found.
    #[error("merge cell range '{0}' not found")]
    MergeCellNotFound(String),

    /// The defined name is invalid.
    #[error("invalid defined name: {0}")]
    InvalidDefinedName(String),

    /// The specified defined name was not found.
    #[error("defined name '{name}' not found")]
    DefinedNameNotFound { name: String },

    /// A circular reference was detected during formula evaluation.
    #[error("circular reference detected at {cell}")]
    CircularReference { cell: String },

    /// The formula references an unknown function.
    #[error("unknown function: {name}")]
    UnknownFunction { name: String },

    /// A function received the wrong number of arguments.
    #[error("function {name} expects {expected} arguments, got {got}")]
    WrongArgCount {
        name: String,
        expected: String,
        got: usize,
    },

    /// A general formula evaluation error.
    #[error("formula evaluation error: {0}")]
    FormulaError(String),

    /// The specified pivot table was not found.
    #[error("pivot table '{name}' not found")]
    PivotTableNotFound { name: String },

    /// A pivot table with the given name already exists.
    #[error("pivot table '{name}' already exists")]
    PivotTableAlreadyExists { name: String },

    /// The specified table was not found.
    #[error("table '{name}' not found")]
    TableNotFound { name: String },

    /// A table with the given name already exists.
    #[error("table '{name}' already exists")]
    TableAlreadyExists { name: String },

    /// The source data range for a pivot table is invalid.
    #[error("invalid source range: {0}")]
    InvalidSourceRange(String),

    /// The specified slicer was not found.
    #[error("slicer '{name}' not found")]
    SlicerNotFound { name: String },

    /// A slicer with the given name already exists.
    #[error("slicer '{name}' already exists")]
    SlicerAlreadyExists { name: String },

    /// The specified table was not found.
    #[error("table '{name}' not found")]
    TableNotFound { name: String },

    /// The specified column was not found in the table.
    #[error("column '{column}' not found in table '{table}'")]
    TableColumnNotFound { table: String, column: String },

    /// The image format is not supported.
    #[error("unsupported image format: {format}")]
    UnsupportedImageFormat { format: String },

    /// The file is encrypted and requires a password to open.
    #[error("file is encrypted, password required")]
    FileEncrypted,

    /// The provided password is incorrect.
    #[error("incorrect password")]
    IncorrectPassword,

    /// The encryption method is not supported.
    #[error("unsupported encryption method: {0}")]
    UnsupportedEncryption(String),

    /// The outline level exceeds the allowed maximum (7).
    #[error("outline level {level} exceeds maximum {max}")]
    OutlineLevelExceeded { level: u8, max: u8 },

    /// A merge cell reference format is invalid.
    #[error("invalid merge cell reference: {0}")]
    InvalidMergeCellReference(String),

    /// A cell range reference (sqref) is invalid.
    #[error("invalid reference: {reference}")]
    InvalidReference { reference: String },

    /// A function argument or configuration value is invalid.
    #[error("invalid argument: {0}")]
    InvalidArgument(String),

    /// The file extension is not a supported OOXML spreadsheet format.
    #[error("unsupported file extension: {0}")]
    UnsupportedFileExtension(String),

    /// The total decompressed size of the ZIP archive exceeds the safety limit.
    #[error("ZIP decompressed size {size} bytes exceeds limit of {limit} bytes")]
    ZipSizeExceeded { size: u64, limit: u64 },

    /// The number of entries in the ZIP archive exceeds the safety limit.
    #[error("ZIP entry count {count} exceeds limit of {limit}")]
    ZipEntryCountExceeded { count: usize, limit: usize },

    /// An internal or otherwise unclassified error.
    #[error("internal error: {0}")]
    Internal(String),
}

/// A convenience alias used throughout the crate.
pub type Result<T> = std::result::Result<T, Error>;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_error_display_invalid_cell_reference() {
        let err = Error::InvalidCellReference("XYZ0".to_string());
        assert_eq!(err.to_string(), "invalid cell reference: XYZ0");
    }

    #[test]
    fn test_error_display_sheet_not_found() {
        let err = Error::SheetNotFound {
            name: "Missing".to_string(),
        };
        assert_eq!(err.to_string(), "sheet 'Missing' does not exist");
    }

    #[test]
    fn test_error_display_sheet_already_exists() {
        let err = Error::SheetAlreadyExists {
            name: "Sheet1".to_string(),
        };
        assert_eq!(err.to_string(), "sheet 'Sheet1' already exists");
    }

    #[test]
    fn test_error_display_invalid_sheet_name() {
        let err = Error::InvalidSheetName("bad[name".to_string());
        assert_eq!(err.to_string(), "invalid sheet name: bad[name");
    }

    #[test]
    fn test_error_display_invalid_row_number() {
        let err = Error::InvalidRowNumber(0);
        assert_eq!(err.to_string(), "invalid row number: 0");
    }

    #[test]
    fn test_error_display_invalid_column_number() {
        let err = Error::InvalidColumnNumber(99999);
        assert_eq!(err.to_string(), "invalid column number: 99999");
    }

    #[test]
    fn test_error_display_io() {
        let io_err = std::io::Error::new(std::io::ErrorKind::NotFound, "gone");
        let err = Error::Io(io_err);
        assert_eq!(err.to_string(), "I/O error: gone");
    }

    #[test]
    fn test_error_display_zip() {
        let err = Error::Zip("corrupted archive".to_string());
        assert_eq!(err.to_string(), "ZIP error: corrupted archive");
    }

    #[test]
    fn test_error_display_xml_parse() {
        let err = Error::XmlParse("unexpected EOF".to_string());
        assert_eq!(err.to_string(), "XML parse error: unexpected EOF");
    }

    #[test]
    fn test_error_display_xml_deserialize() {
        let err = Error::XmlDeserialize("missing attribute".to_string());
        assert_eq!(
            err.to_string(),
            "XML deserialization error: missing attribute"
        );
    }

    #[test]
    fn test_error_display_cell_value_too_long() {
        let err = Error::CellValueTooLong {
            length: 40000,
            max: 32767,
        };
        assert_eq!(
            err.to_string(),
            "cell value too long: 40000 characters (max 32767)"
        );
    }

    #[test]
    fn test_error_display_outline_level_exceeded() {
        let err = Error::OutlineLevelExceeded { level: 8, max: 7 };
        assert_eq!(err.to_string(), "outline level 8 exceeds maximum 7");
    }

    #[test]
    fn test_error_display_invalid_merge_cell_reference() {
        let err = Error::InvalidMergeCellReference("bad ref".to_string());
        assert_eq!(err.to_string(), "invalid merge cell reference: bad ref");
    }

    #[test]
    fn test_error_display_unsupported_file_extension() {
        let err = Error::UnsupportedFileExtension("csv".to_string());
        assert_eq!(err.to_string(), "unsupported file extension: csv");
    }

    #[test]
    fn test_error_display_internal() {
        let err = Error::Internal("something went wrong".to_string());
        assert_eq!(err.to_string(), "internal error: something went wrong");
    }

    #[test]
    fn test_from_io_error() {
        let io_err = std::io::Error::new(std::io::ErrorKind::PermissionDenied, "denied");
        let err: Error = io_err.into();
        assert!(matches!(err, Error::Io(_)));
    }

    #[test]
    fn test_error_is_send_and_sync() {
        fn assert_send<T: Send>() {}
        fn assert_sync<T: Sync>() {}
        assert_send::<Error>();
        assert_sync::<Error>();
    }
}
