//! Table and auto-filter management.
//!
//! Provides functions for creating, listing, and deleting OOXML table parts on
//! worksheets, as well as the existing auto-filter helpers.

use crate::error::{Error, Result};
use sheetkit_xml::table::{
    TableAutoFilter, TableColumnXml, TableColumnsXml, TableStyleInfoXml, TableXml,
};
use sheetkit_xml::worksheet::{AutoFilter, WorksheetXml};

/// Configuration for creating a table.
#[derive(Debug, Clone)]
pub struct TableConfig {
    /// The table name (used internally, must be unique within the workbook).
    pub name: String,
    /// The display name shown in the UI.
    pub display_name: String,
    /// The cell range (e.g. "A1:D10").
    pub range: String,
    /// Column definitions.
    pub columns: Vec<TableColumn>,
    /// Whether to show the header row. Defaults to true.
    pub show_header_row: bool,
    /// The table style name (e.g. "TableStyleMedium2").
    pub style_name: Option<String>,
    /// Whether to enable auto-filter on the table.
    pub auto_filter: bool,
    /// Whether to show first column formatting.
    pub show_first_column: bool,
    /// Whether to show last column formatting.
    pub show_last_column: bool,
    /// Whether to show row stripes.
    pub show_row_stripes: bool,
    /// Whether to show column stripes.
    pub show_column_stripes: bool,
}

impl Default for TableConfig {
    fn default() -> Self {
        Self {
            name: String::new(),
            display_name: String::new(),
            range: String::new(),
            columns: Vec::new(),
            show_header_row: true,
            style_name: None,
            auto_filter: true,
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        }
    }
}

/// A column within a table.
#[derive(Debug, Clone)]
pub struct TableColumn {
    /// The column header name.
    pub name: String,
    /// Optional totals row function (e.g., "sum", "count", "average").
    pub totals_row_function: Option<String>,
    /// Optional totals row label (used for the first column in totals row).
    pub totals_row_label: Option<String>,
}

/// Metadata about an existing table, returned by list/get operations.
#[derive(Debug, Clone, PartialEq)]
pub struct TableInfo {
    /// The table name.
    pub name: String,
    /// The display name.
    pub display_name: String,
    /// The cell range (e.g. "A1:D10").
    pub range: String,
    /// Whether the table has a header row.
    pub show_header_row: bool,
    /// Whether auto-filter is enabled.
    pub auto_filter: bool,
    /// Column names.
    pub columns: Vec<String>,
    /// The style name, if any.
    pub style_name: Option<String>,
}

/// Build a `TableXml` from a `TableConfig` and a unique table ID.
pub(crate) fn build_table_xml(config: &TableConfig, table_id: u32) -> TableXml {
    let columns: Vec<TableColumnXml> = config
        .columns
        .iter()
        .enumerate()
        .map(|(i, col)| TableColumnXml {
            id: (i + 1) as u32,
            name: col.name.clone(),
            totals_row_function: col.totals_row_function.clone(),
            totals_row_label: col.totals_row_label.clone(),
        })
        .collect();

    let auto_filter = if config.auto_filter {
        Some(TableAutoFilter {
            reference: config.range.clone(),
        })
    } else {
        None
    };

    let style_info = if config.style_name.is_some()
        || config.show_first_column
        || config.show_last_column
        || config.show_row_stripes
        || config.show_column_stripes
    {
        Some(TableStyleInfoXml {
            name: config.style_name.clone(),
            show_first_column: Some(config.show_first_column),
            show_last_column: Some(config.show_last_column),
            show_row_stripes: Some(config.show_row_stripes),
            show_column_stripes: Some(config.show_column_stripes),
        })
    } else {
        None
    };

    let header_row_count = if !config.show_header_row {
        Some(0)
    } else {
        None
    };

    TableXml {
        xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
        id: table_id,
        name: config.name.clone(),
        display_name: config.display_name.clone(),
        reference: config.range.clone(),
        totals_row_count: None,
        totals_row_shown: None,
        header_row_count,
        auto_filter,
        table_columns: TableColumnsXml {
            count: columns.len() as u32,
            columns,
        },
        table_style_info: style_info,
    }
}

/// Convert a `TableXml` into a `TableInfo` for external consumption.
pub(crate) fn table_xml_to_info(table_xml: &TableXml) -> TableInfo {
    let columns = table_xml
        .table_columns
        .columns
        .iter()
        .map(|c| c.name.clone())
        .collect();
    let show_header_row = table_xml.header_row_count != Some(0);
    let auto_filter = table_xml.auto_filter.is_some();
    let style_name = table_xml
        .table_style_info
        .as_ref()
        .and_then(|s| s.name.clone());

    TableInfo {
        name: table_xml.name.clone(),
        display_name: table_xml.display_name.clone(),
        range: table_xml.reference.clone(),
        show_header_row,
        auto_filter,
        columns,
        style_name,
    }
}

/// Validate that a table name is non-empty and the range is non-empty.
pub(crate) fn validate_table_config(config: &TableConfig) -> Result<()> {
    if config.name.is_empty() {
        return Err(Error::InvalidArgument("table name cannot be empty".into()));
    }
    if config.range.is_empty() {
        return Err(Error::InvalidArgument("table range cannot be empty".into()));
    }
    if config.columns.is_empty() {
        return Err(Error::InvalidArgument(
            "table must have at least one column".into(),
        ));
    }
    Ok(())
}

/// Internal table registry entry stored in the workbook.
///
/// Tracks the table ID, name, owning sheet index, and column names so that
/// slicer metadata can be wired to real table definitions without requiring
/// full OOXML table part serialization.
#[derive(Debug, Clone)]
pub(crate) struct TableEntry {
    /// Auto-assigned 1-based table ID.
    pub id: u32,
    /// The table name.
    pub name: String,
    /// The sheet index where the table resides.
    pub sheet_index: usize,
    /// The cell range of the table (e.g. "A1:D10").
    pub range: String,
    /// Column names in order.
    pub columns: Vec<String>,
}

/// Information about a table, returned by `get_tables`.
#[derive(Debug, Clone)]
pub struct TableInfo {
    /// The table name.
    pub name: String,
    /// The cell range.
    pub range: String,
    /// Column names.
    pub columns: Vec<String>,
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
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Age".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "City".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Score".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
            ],
            show_header_row: true,
            style_name: Some("TableStyleMedium2".to_string()),
            auto_filter: true,
            ..TableConfig::default()
        };

        assert_eq!(config.name, "Table1");
        assert_eq!(config.columns.len(), 4);
        assert!(config.auto_filter);
    }

    #[test]
    fn test_build_table_xml() {
        let config = TableConfig {
            name: "Sales".to_string(),
            display_name: "Sales".to_string(),
            range: "A1:C5".to_string(),
            columns: vec![
                TableColumn {
                    name: "Product".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Quantity".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Price".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
            ],
            show_header_row: true,
            style_name: Some("TableStyleMedium2".to_string()),
            auto_filter: true,
            show_row_stripes: true,
            ..TableConfig::default()
        };

        let table_xml = build_table_xml(&config, 1);
        assert_eq!(table_xml.id, 1);
        assert_eq!(table_xml.name, "Sales");
        assert_eq!(table_xml.reference, "A1:C5");
        assert_eq!(table_xml.table_columns.count, 3);
        assert!(table_xml.auto_filter.is_some());
        assert!(table_xml.table_style_info.is_some());
        assert!(table_xml.header_row_count.is_none());
    }

    #[test]
    fn test_build_table_xml_no_header() {
        let config = TableConfig {
            name: "T1".to_string(),
            display_name: "T1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col1".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            show_header_row: false,
            ..TableConfig::default()
        };

        let table_xml = build_table_xml(&config, 2);
        assert_eq!(table_xml.header_row_count, Some(0));
    }

    #[test]
    fn test_table_xml_to_info() {
        let table_xml = TableXml {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            id: 1,
            name: "Inventory".to_string(),
            display_name: "Inventory".to_string(),
            reference: "A1:D20".to_string(),
            totals_row_count: None,
            totals_row_shown: None,
            header_row_count: None,
            auto_filter: Some(TableAutoFilter {
                reference: "A1:D20".to_string(),
            }),
            table_columns: TableColumnsXml {
                count: 4,
                columns: vec![
                    TableColumnXml {
                        id: 1,
                        name: "Item".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumnXml {
                        id: 2,
                        name: "Stock".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumnXml {
                        id: 3,
                        name: "Price".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumnXml {
                        id: 4,
                        name: "Supplier".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                ],
            },
            table_style_info: Some(TableStyleInfoXml {
                name: Some("TableStyleLight1".to_string()),
                show_first_column: Some(false),
                show_last_column: Some(false),
                show_row_stripes: Some(true),
                show_column_stripes: Some(false),
            }),
        };

        let info = table_xml_to_info(&table_xml);
        assert_eq!(info.name, "Inventory");
        assert_eq!(info.display_name, "Inventory");
        assert_eq!(info.range, "A1:D20");
        assert!(info.show_header_row);
        assert!(info.auto_filter);
        assert_eq!(info.columns, vec!["Item", "Stock", "Price", "Supplier"]);
        assert_eq!(info.style_name, Some("TableStyleLight1".to_string()));
    }

    #[test]
    fn test_validate_table_config_empty_name() {
        let config = TableConfig {
            name: String::new(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        assert!(validate_table_config(&config).is_err());
    }

    #[test]
    fn test_validate_table_config_empty_range() {
        let config = TableConfig {
            name: "T1".to_string(),
            range: String::new(),
            columns: vec![TableColumn {
                name: "Col".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        assert!(validate_table_config(&config).is_err());
    }

    #[test]
    fn test_validate_table_config_no_columns() {
        let config = TableConfig {
            name: "T1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![],
            ..TableConfig::default()
        };
        assert!(validate_table_config(&config).is_err());
    }

    #[test]
    fn test_validate_table_config_valid() {
        let config = TableConfig {
            name: "T1".to_string(),
            display_name: "T1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        assert!(validate_table_config(&config).is_ok());
    }

    #[test]
    fn test_table_config_default() {
        let config = TableConfig::default();
        assert!(config.show_header_row);
        assert!(config.auto_filter);
        assert!(config.show_row_stripes);
        assert!(!config.show_first_column);
        assert!(!config.show_last_column);
        assert!(!config.show_column_stripes);
    }
}
