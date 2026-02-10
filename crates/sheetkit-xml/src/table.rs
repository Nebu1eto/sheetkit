//! Table XML schema structures.
//!
//! Represents `xl/tables/table{N}.xml` in the OOXML package.

use serde::{Deserialize, Serialize};

use crate::namespaces;

/// Root element for a table definition part.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "table")]
pub struct TableXml {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "@id")]
    pub id: u32,

    #[serde(rename = "@name")]
    pub name: String,

    #[serde(rename = "@displayName")]
    pub display_name: String,

    #[serde(rename = "@ref")]
    pub reference: String,

    #[serde(rename = "@totalsRowCount", skip_serializing_if = "Option::is_none")]
    pub totals_row_count: Option<u32>,

    #[serde(rename = "@totalsRowShown", skip_serializing_if = "Option::is_none")]
    pub totals_row_shown: Option<bool>,

    #[serde(rename = "@headerRowCount", skip_serializing_if = "Option::is_none")]
    pub header_row_count: Option<u32>,

    #[serde(rename = "autoFilter", skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<TableAutoFilter>,

    #[serde(rename = "tableColumns")]
    pub table_columns: TableColumnsXml,

    #[serde(rename = "tableStyleInfo", skip_serializing_if = "Option::is_none")]
    pub table_style_info: Option<TableStyleInfoXml>,
}

/// Auto-filter reference within a table definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TableAutoFilter {
    #[serde(rename = "@ref")]
    pub reference: String,
}

/// Container for table column definitions.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TableColumnsXml {
    #[serde(rename = "@count")]
    pub count: u32,

    #[serde(rename = "tableColumn")]
    pub columns: Vec<TableColumnXml>,
}

/// A single column within a table definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TableColumnXml {
    #[serde(rename = "@id")]
    pub id: u32,

    #[serde(rename = "@name")]
    pub name: String,

    #[serde(rename = "@totalsRowFunction", skip_serializing_if = "Option::is_none")]
    pub totals_row_function: Option<String>,

    #[serde(rename = "@totalsRowLabel", skip_serializing_if = "Option::is_none")]
    pub totals_row_label: Option<String>,
}

/// Style information for a table.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TableStyleInfoXml {
    #[serde(rename = "@name", skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,

    #[serde(rename = "@showFirstColumn", skip_serializing_if = "Option::is_none")]
    pub show_first_column: Option<bool>,

    #[serde(rename = "@showLastColumn", skip_serializing_if = "Option::is_none")]
    pub show_last_column: Option<bool>,

    #[serde(rename = "@showRowStripes", skip_serializing_if = "Option::is_none")]
    pub show_row_stripes: Option<bool>,

    #[serde(rename = "@showColumnStripes", skip_serializing_if = "Option::is_none")]
    pub show_column_stripes: Option<bool>,
}

impl Default for TableXml {
    fn default() -> Self {
        Self {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            id: 1,
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            reference: "A1:A1".to_string(),
            totals_row_count: None,
            totals_row_shown: None,
            header_row_count: None,
            auto_filter: None,
            table_columns: TableColumnsXml {
                count: 1,
                columns: vec![TableColumnXml {
                    id: 1,
                    name: "Column1".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                }],
            },
            table_style_info: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_xml_default() {
        let table = TableXml::default();
        assert_eq!(table.xmlns, namespaces::SPREADSHEET_ML);
        assert_eq!(table.name, "Table1");
        assert_eq!(table.display_name, "Table1");
        assert_eq!(table.reference, "A1:A1");
        assert_eq!(table.table_columns.count, 1);
        assert_eq!(table.table_columns.columns.len(), 1);
    }

    #[test]
    fn test_table_xml_serialize_roundtrip() {
        let table = TableXml {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            id: 1,
            name: "SalesTable".to_string(),
            display_name: "SalesTable".to_string(),
            reference: "A1:D10".to_string(),
            totals_row_count: None,
            totals_row_shown: None,
            header_row_count: None,
            auto_filter: Some(TableAutoFilter {
                reference: "A1:D10".to_string(),
            }),
            table_columns: TableColumnsXml {
                count: 4,
                columns: vec![
                    TableColumnXml {
                        id: 1,
                        name: "Name".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumnXml {
                        id: 2,
                        name: "Region".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumnXml {
                        id: 3,
                        name: "Sales".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumnXml {
                        id: 4,
                        name: "Profit".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                ],
            },
            table_style_info: Some(TableStyleInfoXml {
                name: Some("TableStyleMedium2".to_string()),
                show_first_column: Some(false),
                show_last_column: Some(false),
                show_row_stripes: Some(true),
                show_column_stripes: Some(false),
            }),
        };

        let xml = quick_xml::se::to_string(&table).unwrap();
        assert!(xml.contains("SalesTable"));
        assert!(xml.contains("A1:D10"));
        assert!(xml.contains("autoFilter"));
        assert!(xml.contains("TableStyleMedium2"));

        let parsed: TableXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.name, "SalesTable");
        assert_eq!(parsed.display_name, "SalesTable");
        assert_eq!(parsed.reference, "A1:D10");
        assert_eq!(parsed.table_columns.count, 4);
        assert_eq!(parsed.table_columns.columns.len(), 4);
        assert_eq!(parsed.table_columns.columns[0].name, "Name");
        assert!(parsed.auto_filter.is_some());
        assert_eq!(parsed.auto_filter.unwrap().reference, "A1:D10");
        let style = parsed.table_style_info.unwrap();
        assert_eq!(style.name, Some("TableStyleMedium2".to_string()));
        assert_eq!(style.show_row_stripes, Some(true));
    }

    #[test]
    fn test_table_xml_without_optional_fields() {
        let table = TableXml {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            id: 2,
            name: "Table2".to_string(),
            display_name: "Table2".to_string(),
            reference: "B1:C5".to_string(),
            totals_row_count: None,
            totals_row_shown: None,
            header_row_count: None,
            auto_filter: None,
            table_columns: TableColumnsXml {
                count: 2,
                columns: vec![
                    TableColumnXml {
                        id: 1,
                        name: "Col1".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumnXml {
                        id: 2,
                        name: "Col2".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                ],
            },
            table_style_info: None,
        };

        let xml = quick_xml::se::to_string(&table).unwrap();
        assert!(!xml.contains("autoFilter"));
        assert!(!xml.contains("tableStyleInfo"));

        let parsed: TableXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.id, 2);
        assert!(parsed.auto_filter.is_none());
        assert!(parsed.table_style_info.is_none());
    }

    #[test]
    fn test_table_xml_with_totals_row() {
        let table = TableXml {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            id: 3,
            name: "Table3".to_string(),
            display_name: "Table3".to_string(),
            reference: "A1:B5".to_string(),
            totals_row_count: Some(1),
            totals_row_shown: Some(true),
            header_row_count: None,
            auto_filter: None,
            table_columns: TableColumnsXml {
                count: 2,
                columns: vec![
                    TableColumnXml {
                        id: 1,
                        name: "Label".to_string(),
                        totals_row_function: None,
                        totals_row_label: Some("Total".to_string()),
                    },
                    TableColumnXml {
                        id: 2,
                        name: "Amount".to_string(),
                        totals_row_function: Some("sum".to_string()),
                        totals_row_label: None,
                    },
                ],
            },
            table_style_info: None,
        };

        let xml = quick_xml::se::to_string(&table).unwrap();
        assert!(xml.contains("totalsRowCount"));
        assert!(xml.contains("totalsRowFunction"));
        assert!(xml.contains("totalsRowLabel"));

        let parsed: TableXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.totals_row_count, Some(1));
        assert_eq!(
            parsed.table_columns.columns[0].totals_row_label,
            Some("Total".to_string())
        );
        assert_eq!(
            parsed.table_columns.columns[1].totals_row_function,
            Some("sum".to_string())
        );
    }

    #[test]
    fn test_parse_real_excel_table_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1" name="Table1" displayName="Table1" ref="A1:C4">
  <autoFilter ref="A1:C4"/>
  <tableColumns count="3">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="City"/>
    <tableColumn id="3" name="Score"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0"
                  showRowStripes="1" showColumnStripes="0"/>
</table>"#;

        let parsed: TableXml = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(parsed.id, 1);
        assert_eq!(parsed.name, "Table1");
        assert_eq!(parsed.display_name, "Table1");
        assert_eq!(parsed.reference, "A1:C4");
        assert!(parsed.auto_filter.is_some());
        assert_eq!(parsed.auto_filter.unwrap().reference, "A1:C4");
        assert_eq!(parsed.table_columns.count, 3);
        assert_eq!(parsed.table_columns.columns[0].name, "Name");
        assert_eq!(parsed.table_columns.columns[1].name, "City");
        assert_eq!(parsed.table_columns.columns[2].name, "Score");
        let style = parsed.table_style_info.unwrap();
        assert_eq!(style.name, Some("TableStyleMedium9".to_string()));
        assert_eq!(style.show_first_column, Some(false));
        assert_eq!(style.show_row_stripes, Some(true));
    }
}
