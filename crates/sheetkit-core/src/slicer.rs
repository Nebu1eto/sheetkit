//! Slicer configuration and management.
//!
//! Provides types and logic for adding, retrieving, and deleting slicers
//! on tables. Slicers are visual filter controls introduced in Excel 2010.

use crate::error::Result;

/// Configuration for adding a slicer to a table.
#[derive(Debug, Clone)]
pub struct SlicerConfig {
    /// Unique slicer name.
    pub name: String,
    /// Anchor cell (top-left corner of the slicer, e.g. "F1").
    pub cell: String,
    /// Source table name.
    pub table_name: String,
    /// Column name from the table to filter.
    pub column_name: String,
    /// Caption displayed on the slicer header. Defaults to column_name.
    pub caption: Option<String>,
    /// Slicer visual style (e.g. "SlicerStyleLight1").
    pub style: Option<String>,
    /// Width in pixels. Defaults to 200.
    pub width: Option<u32>,
    /// Height in pixels. Defaults to 200.
    pub height: Option<u32>,
    /// Whether to show the caption header.
    pub show_caption: Option<bool>,
    /// Number of columns in the slicer item display.
    pub column_count: Option<u32>,
}

/// Information about an existing slicer, returned by `get_slicers`.
#[derive(Debug, Clone)]
pub struct SlicerInfo {
    /// The slicer's unique name.
    pub name: String,
    /// The display caption.
    pub caption: String,
    /// The source table name.
    pub table_name: String,
    /// The column name being filtered.
    pub column_name: String,
    /// The visual style name, if set.
    pub style: Option<String>,
}

/// Default slicer row height in EMU (241300 EMU = approx 19pt).
pub const DEFAULT_ROW_HEIGHT_EMU: u32 = 241300;

/// Default slicer width in pixels.
pub const DEFAULT_WIDTH_PX: u32 = 200;

/// Default slicer height in pixels: u32.
pub const DEFAULT_HEIGHT_PX: u32 = 200;

/// Pixels to EMU conversion factor (1 pixel = 9525 EMU at 96 DPI).
pub const PX_TO_EMU: u64 = 9525;

/// Validate a slicer config.
pub fn validate_slicer_config(config: &SlicerConfig) -> Result<()> {
    if config.name.is_empty() {
        return Err(crate::error::Error::InvalidArgument(
            "slicer name cannot be empty".to_string(),
        ));
    }
    if config.cell.is_empty() {
        return Err(crate::error::Error::InvalidArgument(
            "slicer anchor cell cannot be empty".to_string(),
        ));
    }
    if config.table_name.is_empty() {
        return Err(crate::error::Error::InvalidArgument(
            "slicer table_name cannot be empty".to_string(),
        ));
    }
    if config.column_name.is_empty() {
        return Err(crate::error::Error::InvalidArgument(
            "slicer column_name cannot be empty".to_string(),
        ));
    }
    // Validate cell reference.
    crate::utils::cell_ref::cell_name_to_coordinates(&config.cell)?;
    Ok(())
}

/// Generate the sanitized cache name from a slicer name.
pub fn slicer_cache_name(name: &str) -> String {
    format!("Slicer_{}", name.replace(' ', "_"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_validate_slicer_config_valid() {
        let config = SlicerConfig {
            name: "StatusFilter".to_string(),
            cell: "F1".to_string(),
            table_name: "Table1".to_string(),
            column_name: "Status".to_string(),
            caption: None,
            style: None,
            width: None,
            height: None,
            show_caption: None,
            column_count: None,
        };
        assert!(validate_slicer_config(&config).is_ok());
    }

    #[test]
    fn test_validate_slicer_config_empty_name() {
        let config = SlicerConfig {
            name: "".to_string(),
            cell: "A1".to_string(),
            table_name: "T".to_string(),
            column_name: "C".to_string(),
            caption: None,
            style: None,
            width: None,
            height: None,
            show_caption: None,
            column_count: None,
        };
        let err = validate_slicer_config(&config).unwrap_err();
        assert!(err.to_string().contains("name cannot be empty"));
    }

    #[test]
    fn test_validate_slicer_config_empty_cell() {
        let config = SlicerConfig {
            name: "S1".to_string(),
            cell: "".to_string(),
            table_name: "T".to_string(),
            column_name: "C".to_string(),
            caption: None,
            style: None,
            width: None,
            height: None,
            show_caption: None,
            column_count: None,
        };
        let err = validate_slicer_config(&config).unwrap_err();
        assert!(err.to_string().contains("cell cannot be empty"));
    }

    #[test]
    fn test_validate_slicer_config_empty_table() {
        let config = SlicerConfig {
            name: "S1".to_string(),
            cell: "A1".to_string(),
            table_name: "".to_string(),
            column_name: "C".to_string(),
            caption: None,
            style: None,
            width: None,
            height: None,
            show_caption: None,
            column_count: None,
        };
        let err = validate_slicer_config(&config).unwrap_err();
        assert!(err.to_string().contains("table_name cannot be empty"));
    }

    #[test]
    fn test_validate_slicer_config_empty_column() {
        let config = SlicerConfig {
            name: "S1".to_string(),
            cell: "A1".to_string(),
            table_name: "T".to_string(),
            column_name: "".to_string(),
            caption: None,
            style: None,
            width: None,
            height: None,
            show_caption: None,
            column_count: None,
        };
        let err = validate_slicer_config(&config).unwrap_err();
        assert!(err.to_string().contains("column_name cannot be empty"));
    }

    #[test]
    fn test_validate_slicer_config_invalid_cell() {
        let config = SlicerConfig {
            name: "S1".to_string(),
            cell: "XYZ0".to_string(),
            table_name: "T".to_string(),
            column_name: "C".to_string(),
            caption: None,
            style: None,
            width: None,
            height: None,
            show_caption: None,
            column_count: None,
        };
        assert!(validate_slicer_config(&config).is_err());
    }

    #[test]
    fn test_slicer_cache_name() {
        assert_eq!(slicer_cache_name("Category"), "Slicer_Category");
        assert_eq!(slicer_cache_name("My Filter"), "Slicer_My_Filter");
    }

    #[test]
    fn test_slicer_info_struct() {
        let info = SlicerInfo {
            name: "StatusFilter".to_string(),
            caption: "Status".to_string(),
            table_name: "Table1".to_string(),
            column_name: "Status".to_string(),
            style: Some("SlicerStyleLight1".to_string()),
        };
        assert_eq!(info.name, "StatusFilter");
        assert_eq!(info.column_name, "Status");
    }
}
