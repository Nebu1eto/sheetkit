//! Error types for the SheetKit core library.
//!
//! Provides a comprehensive [`Error`] enum covering all failure modes
//! encountered when reading, writing, and manipulating Excel workbooks.

use thiserror::Error;

/// The top-level error type for SheetKit.
#[derive(Error, Debug)]
pub enum Error {
    // ===== Cell reference errors =====
    /// The given string is not a valid A1-style cell reference.
    #[error("invalid cell reference: {0}")]
    InvalidCellReference(String),

    /// The row number is out of the allowed range (1..=1_048_576).
    #[error("invalid row number: {0}")]
    InvalidRowNumber(u32),

    /// The column number is out of the allowed range (1..=16_384).
    #[error("invalid column number: {0}")]
    InvalidColumnNumber(u32),

    // ===== Sheet errors =====
    /// No sheet with the given name exists in the workbook.
    #[error("sheet '{name}' does not exist")]
    SheetNotFound { name: String },

    /// A sheet with the given name already exists.
    #[error("sheet '{name}' already exists")]
    SheetAlreadyExists { name: String },

    /// The sheet name violates Excel naming rules.
    #[error("invalid sheet name: {0}")]
    InvalidSheetName(String),

    // ===== I/O errors =====
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

    // ===== Value errors =====
    /// A cell value exceeds the maximum character limit.
    #[error("cell value too long: {length} characters (max {max})")]
    CellValueTooLong { length: usize, max: usize },

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
        assert_eq!(err.to_string(), "XML deserialization error: missing attribute");
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
