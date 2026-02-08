//! Table and auto-filter management utilities.
//!
//! Provides functions for setting and removing auto-filters on worksheets.

use crate::error::Result;
use sheetkit_xml::worksheet::{AutoFilter, WorksheetXml};

/// Configuration for a table.
#[derive(Debug, Clone)]
pub struct TableConfig {
    /// The table name (used internally).
    pub name: String,
    /// The display name shown in the UI.
    pub display_name: String,
    /// The cell range (e.g. "A1:D10").
    pub range: String,
    /// Column definitions.
    pub columns: Vec<TableColumn>,
    /// Whether to show the header row.
    pub show_header_row: bool,
    /// The table style name (e.g. "TableStyleMedium2").
    pub style_name: Option<String>,
    /// Whether to enable auto-filter on the table.
    pub auto_filter: bool,
}

/// A column within a table.
#[derive(Debug, Clone)]
pub struct TableColumn {
    /// The column header name.
    pub name: String,
}

/// Set an auto-filter on a worksheet for the given cell range.
pub fn set_auto_filter(ws: &mut WorksheetXml, range: &str) -> Result<()> {
    ws.auto_filter = Some(AutoFilter {
        reference: range.to_string(),
    });
    Ok(())
}

/// Remove any auto-filter from a worksheet.
pub fn remove_auto_filter(ws: &mut WorksheetXml) {
    ws.auto_filter = None;
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_set_auto_filter() {
        let mut ws = WorksheetXml::default();
        set_auto_filter(&mut ws, "A1:D10").unwrap();

        assert!(ws.auto_filter.is_some());
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:D10");
    }

    #[test]
    fn test_remove_auto_filter() {
        let mut ws = WorksheetXml::default();
        set_auto_filter(&mut ws, "A1:D10").unwrap();
        remove_auto_filter(&mut ws);

        assert!(ws.auto_filter.is_none());
    }

    #[test]
    fn test_auto_filter_xml_roundtrip() {
        let mut ws = WorksheetXml::default();
        set_auto_filter(&mut ws, "A1:C100").unwrap();

        let xml = quick_xml::se::to_string(&ws).unwrap();
        assert!(xml.contains("autoFilter"));
        assert!(xml.contains("A1:C100"));

        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.auto_filter.is_some());
        assert_eq!(parsed.auto_filter.as_ref().unwrap().reference, "A1:C100");
    }

    #[test]
    fn test_remove_auto_filter_when_none() {
        let mut ws = WorksheetXml::default();
        // Should not panic when removing a non-existent filter.
        remove_auto_filter(&mut ws);
        assert!(ws.auto_filter.is_none());
    }

    #[test]
    fn test_overwrite_auto_filter() {
        let mut ws = WorksheetXml::default();
        set_auto_filter(&mut ws, "A1:B10").unwrap();
        set_auto_filter(&mut ws, "A1:D20").unwrap();

        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:D20");
    }

    #[test]
    fn test_table_config_creation() {
        let config = TableConfig {
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            range: "A1:D10".to_string(),
            columns: vec![
                TableColumn {
                    name: "Name".to_string(),
                },
                TableColumn {
                    name: "Age".to_string(),
                },
                TableColumn {
                    name: "City".to_string(),
                },
                TableColumn {
                    name: "Score".to_string(),
                },
            ],
            show_header_row: true,
            style_name: Some("TableStyleMedium2".to_string()),
            auto_filter: true,
        };

        assert_eq!(config.name, "Table1");
        assert_eq!(config.columns.len(), 4);
        assert!(config.auto_filter);
    }
}
