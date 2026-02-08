//! Pivot table configuration and builder.

use crate::error::{Error, Result};

/// Aggregate function for pivot table data fields.
#[derive(Debug, Clone, PartialEq)]
pub enum AggregateFunction {
    Sum,
    Count,
    Average,
    Max,
    Min,
    Product,
    CountNums,
    StdDev,
    StdDevP,
    Var,
    VarP,
}

impl AggregateFunction {
    /// Returns the XML string representation of this aggregate function.
    pub fn to_xml_str(&self) -> &str {
        match self {
            Self::Sum => "sum",
            Self::Count => "count",
            Self::Average => "average",
            Self::Max => "max",
            Self::Min => "min",
            Self::Product => "product",
            Self::CountNums => "countNums",
            Self::StdDev => "stdDev",
            Self::StdDevP => "stdDevp",
            Self::Var => "var",
            Self::VarP => "varp",
        }
    }

    /// Parses an aggregate function from its XML string representation.
    pub fn from_xml_str(s: &str) -> Option<Self> {
        match s {
            "sum" => Some(Self::Sum),
            "count" => Some(Self::Count),
            "average" => Some(Self::Average),
            "max" => Some(Self::Max),
            "min" => Some(Self::Min),
            "product" => Some(Self::Product),
            "countNums" => Some(Self::CountNums),
            "stdDev" => Some(Self::StdDev),
            "stdDevp" => Some(Self::StdDevP),
            "var" => Some(Self::Var),
            "varp" => Some(Self::VarP),
            _ => None,
        }
    }
}

/// Configuration for adding a pivot table.
#[derive(Debug, Clone)]
pub struct PivotTableConfig {
    /// Name of the pivot table.
    pub name: String,
    /// Source data sheet name.
    pub source_sheet: String,
    /// Source data range (e.g., "A1:D10").
    pub source_range: String,
    /// Target sheet name where the pivot table will be placed.
    pub target_sheet: String,
    /// Target cell (top-left corner of pivot table, e.g., "A1").
    pub target_cell: String,
    /// Row fields (column names from source data).
    pub rows: Vec<PivotField>,
    /// Column fields.
    pub columns: Vec<PivotField>,
    /// Data/value fields.
    pub data: Vec<PivotDataField>,
}

/// A field used as row or column in the pivot table.
#[derive(Debug, Clone)]
pub struct PivotField {
    /// Column name from the source data header row.
    pub name: String,
}

/// A data/value field in the pivot table.
#[derive(Debug, Clone)]
pub struct PivotDataField {
    /// Column name from the source data header row.
    pub name: String,
    /// Aggregate function to apply.
    pub function: AggregateFunction,
    /// Optional custom display name.
    pub display_name: Option<String>,
}

/// Information about an existing pivot table.
#[derive(Debug, Clone)]
pub struct PivotTableInfo {
    pub name: String,
    pub source_sheet: String,
    pub source_range: String,
    pub target_sheet: String,
    pub location: String,
}

/// Builds the pivot table definition XML from config and field names.
pub fn build_pivot_table_xml(
    config: &PivotTableConfig,
    cache_id: u32,
    field_names: &[String],
) -> Result<sheetkit_xml::pivot_table::PivotTableDefinition> {
    use sheetkit_xml::pivot_table::*;

    let ns = sheetkit_xml::namespaces::SPREADSHEET_ML;

    let find_field_index = |name: &str| -> Result<usize> {
        field_names.iter().position(|n| n == name).ok_or_else(|| {
            Error::Internal(format!("pivot field '{}' not found in source data", name))
        })
    };

    let mut pivot_field_defs = Vec::new();
    for field_name in field_names {
        let is_row = config.rows.iter().any(|r| r.name == *field_name);
        let is_col = config.columns.iter().any(|c| c.name == *field_name);
        let is_data = config.data.iter().any(|d| d.name == *field_name);

        let axis = if is_row {
            Some("axisRow".to_string())
        } else if is_col {
            Some("axisCol".to_string())
        } else {
            None
        };

        pivot_field_defs.push(PivotFieldDef {
            axis,
            data_field: if is_data { Some(true) } else { None },
            show_all: Some(false),
            items: None,
        });
    }

    let row_fields = if config.rows.is_empty() {
        None
    } else {
        let fields: Result<Vec<FieldRef>> = config
            .rows
            .iter()
            .map(|r| find_field_index(&r.name).map(|i| FieldRef { index: i as i32 }))
            .collect();
        Some(FieldList {
            count: Some(config.rows.len() as u32),
            fields: fields?,
        })
    };

    let col_fields = if config.columns.is_empty() {
        None
    } else {
        let fields: Result<Vec<FieldRef>> = config
            .columns
            .iter()
            .map(|c| find_field_index(&c.name).map(|i| FieldRef { index: i as i32 }))
            .collect();
        Some(FieldList {
            count: Some(config.columns.len() as u32),
            fields: fields?,
        })
    };

    let data_fields = if config.data.is_empty() {
        None
    } else {
        let fields: Result<Vec<DataFieldDef>> = config
            .data
            .iter()
            .map(|d| {
                let idx = find_field_index(&d.name)?;
                Ok(DataFieldDef {
                    name: d.display_name.clone().or_else(|| {
                        Some(format!(
                            "{} of {}",
                            capitalize_first(d.function.to_xml_str()),
                            d.name
                        ))
                    }),
                    field_index: idx as u32,
                    subtotal: Some(d.function.to_xml_str().to_string()),
                    base_field: Some(0),
                    base_item: Some(0),
                })
            })
            .collect();
        Some(DataFields {
            count: Some(config.data.len() as u32),
            fields: fields?,
        })
    };

    Ok(PivotTableDefinition {
        xmlns: ns.to_string(),
        name: config.name.clone(),
        cache_id,
        data_on_rows: Some(false),
        apply_number_formats: Some(false),
        apply_border_formats: Some(false),
        apply_font_formats: Some(false),
        apply_pattern_formats: Some(false),
        apply_alignment_formats: Some(false),
        apply_width_height_formats: Some(true),
        location: PivotLocation {
            reference: config.target_cell.clone(),
            first_header_row: 1,
            first_data_row: 1,
            first_data_col: 1,
        },
        pivot_fields: PivotFields {
            count: Some(field_names.len() as u32),
            fields: pivot_field_defs,
        },
        row_fields,
        col_fields,
        data_fields,
    })
}

/// Builds the pivot cache definition XML.
pub fn build_pivot_cache_definition(
    source_sheet: &str,
    source_range: &str,
    field_names: &[String],
) -> sheetkit_xml::pivot_cache::PivotCacheDefinition {
    use sheetkit_xml::pivot_cache::*;

    let cache_fields = CacheFields {
        count: Some(field_names.len() as u32),
        fields: field_names
            .iter()
            .map(|name| CacheField {
                name: name.clone(),
                num_fmt_id: Some(0),
                shared_items: Some(SharedItems {
                    contains_semi_mixed_types: None,
                    contains_string: None,
                    contains_number: None,
                    contains_blank: None,
                    count: Some(0),
                    string_items: vec![],
                    number_items: vec![],
                }),
            })
            .collect(),
    };

    PivotCacheDefinition {
        xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
        xmlns_r: sheetkit_xml::namespaces::RELATIONSHIPS.to_string(),
        r_id: None,
        record_count: Some(0),
        cache_source: CacheSource {
            source_type: "worksheet".to_string(),
            worksheet_source: Some(WorksheetSource {
                reference: source_range.to_string(),
                sheet: source_sheet.to_string(),
            }),
        },
        cache_fields,
    }
}

fn capitalize_first(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        None => String::new(),
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_aggregate_function_to_xml_str() {
        assert_eq!(AggregateFunction::Sum.to_xml_str(), "sum");
        assert_eq!(AggregateFunction::Count.to_xml_str(), "count");
        assert_eq!(AggregateFunction::Average.to_xml_str(), "average");
        assert_eq!(AggregateFunction::Max.to_xml_str(), "max");
        assert_eq!(AggregateFunction::Min.to_xml_str(), "min");
        assert_eq!(AggregateFunction::Product.to_xml_str(), "product");
        assert_eq!(AggregateFunction::CountNums.to_xml_str(), "countNums");
        assert_eq!(AggregateFunction::StdDev.to_xml_str(), "stdDev");
        assert_eq!(AggregateFunction::StdDevP.to_xml_str(), "stdDevp");
        assert_eq!(AggregateFunction::Var.to_xml_str(), "var");
        assert_eq!(AggregateFunction::VarP.to_xml_str(), "varp");
    }

    #[test]
    fn test_aggregate_function_from_xml_str() {
        assert_eq!(
            AggregateFunction::from_xml_str("sum"),
            Some(AggregateFunction::Sum)
        );
        assert_eq!(
            AggregateFunction::from_xml_str("count"),
            Some(AggregateFunction::Count)
        );
        assert_eq!(
            AggregateFunction::from_xml_str("average"),
            Some(AggregateFunction::Average)
        );
        assert_eq!(
            AggregateFunction::from_xml_str("max"),
            Some(AggregateFunction::Max)
        );
        assert_eq!(
            AggregateFunction::from_xml_str("min"),
            Some(AggregateFunction::Min)
        );
        assert_eq!(
            AggregateFunction::from_xml_str("product"),
            Some(AggregateFunction::Product)
        );
        assert_eq!(
            AggregateFunction::from_xml_str("countNums"),
            Some(AggregateFunction::CountNums)
        );
        assert_eq!(
            AggregateFunction::from_xml_str("stdDev"),
            Some(AggregateFunction::StdDev)
        );
        assert_eq!(
            AggregateFunction::from_xml_str("stdDevp"),
            Some(AggregateFunction::StdDevP)
        );
        assert_eq!(
            AggregateFunction::from_xml_str("var"),
            Some(AggregateFunction::Var)
        );
        assert_eq!(
            AggregateFunction::from_xml_str("varp"),
            Some(AggregateFunction::VarP)
        );
    }

    #[test]
    fn test_aggregate_function_from_xml_str_unknown() {
        assert_eq!(AggregateFunction::from_xml_str("unknown"), None);
        assert_eq!(AggregateFunction::from_xml_str(""), None);
        assert_eq!(AggregateFunction::from_xml_str("SUM"), None);
    }

    #[test]
    fn test_aggregate_function_roundtrip() {
        let functions = vec![
            AggregateFunction::Sum,
            AggregateFunction::Count,
            AggregateFunction::Average,
            AggregateFunction::Max,
            AggregateFunction::Min,
            AggregateFunction::Product,
            AggregateFunction::CountNums,
            AggregateFunction::StdDev,
            AggregateFunction::StdDevP,
            AggregateFunction::Var,
            AggregateFunction::VarP,
        ];
        for func in functions {
            let xml_str = func.to_xml_str();
            let parsed = AggregateFunction::from_xml_str(xml_str).unwrap();
            assert_eq!(func, parsed);
        }
    }

    #[test]
    fn test_capitalize_first() {
        assert_eq!(capitalize_first("sum"), "Sum");
        assert_eq!(capitalize_first("count"), "Count");
        assert_eq!(capitalize_first("average"), "Average");
        assert_eq!(capitalize_first(""), "");
        assert_eq!(capitalize_first("a"), "A");
    }

    #[test]
    fn test_build_pivot_table_xml_basic() {
        let config = PivotTableConfig {
            name: "PivotTable1".to_string(),
            source_sheet: "Data".to_string(),
            source_range: "A1:C5".to_string(),
            target_sheet: "Pivot".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![PivotField {
                name: "Region".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        let field_names = vec![
            "Region".to_string(),
            "Product".to_string(),
            "Sales".to_string(),
        ];

        let def = build_pivot_table_xml(&config, 0, &field_names).unwrap();
        assert_eq!(def.name, "PivotTable1");
        assert_eq!(def.cache_id, 0);
        assert_eq!(def.pivot_fields.count, Some(3));
        assert_eq!(def.pivot_fields.fields.len(), 3);

        // Region is axisRow
        assert_eq!(def.pivot_fields.fields[0].axis, Some("axisRow".to_string()));
        assert_eq!(def.pivot_fields.fields[0].data_field, None);

        // Product has no axis
        assert_eq!(def.pivot_fields.fields[1].axis, None);

        // Sales is data field
        assert_eq!(def.pivot_fields.fields[2].axis, None);
        assert_eq!(def.pivot_fields.fields[2].data_field, Some(true));

        // Row fields
        let row_fields = def.row_fields.unwrap();
        assert_eq!(row_fields.count, Some(1));
        assert_eq!(row_fields.fields[0].index, 0);

        // No col fields
        assert!(def.col_fields.is_none());

        // Data fields
        let data_fields = def.data_fields.unwrap();
        assert_eq!(data_fields.count, Some(1));
        assert_eq!(data_fields.fields[0].field_index, 2);
        assert_eq!(data_fields.fields[0].subtotal, Some("sum".to_string()));
        assert_eq!(data_fields.fields[0].name, Some("Sum of Sales".to_string()));
    }

    #[test]
    fn test_build_pivot_table_xml_with_columns() {
        let config = PivotTableConfig {
            name: "SalesReport".to_string(),
            source_sheet: "Data".to_string(),
            source_range: "A1:D10".to_string(),
            target_sheet: "Report".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![PivotField {
                name: "Region".to_string(),
            }],
            columns: vec![PivotField {
                name: "Quarter".to_string(),
            }],
            data: vec![PivotDataField {
                name: "Revenue".to_string(),
                function: AggregateFunction::Average,
                display_name: Some("Avg Revenue".to_string()),
            }],
        };
        let field_names = vec![
            "Region".to_string(),
            "Quarter".to_string(),
            "Revenue".to_string(),
        ];

        let def = build_pivot_table_xml(&config, 1, &field_names).unwrap();
        assert_eq!(def.cache_id, 1);

        // Region = axisRow, Quarter = axisCol
        assert_eq!(def.pivot_fields.fields[0].axis, Some("axisRow".to_string()));
        assert_eq!(def.pivot_fields.fields[1].axis, Some("axisCol".to_string()));

        let col_fields = def.col_fields.unwrap();
        assert_eq!(col_fields.count, Some(1));
        assert_eq!(col_fields.fields[0].index, 1);

        let data_fields = def.data_fields.unwrap();
        assert_eq!(data_fields.fields[0].name, Some("Avg Revenue".to_string()));
        assert_eq!(data_fields.fields[0].subtotal, Some("average".to_string()));
    }

    #[test]
    fn test_build_pivot_table_xml_unknown_field() {
        let config = PivotTableConfig {
            name: "Bad".to_string(),
            source_sheet: "Data".to_string(),
            source_range: "A1:B2".to_string(),
            target_sheet: "Pivot".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![PivotField {
                name: "NonExistent".to_string(),
            }],
            columns: vec![],
            data: vec![],
        };
        let field_names = vec!["Actual".to_string()];

        let result = build_pivot_table_xml(&config, 0, &field_names);
        assert!(result.is_err());
        let err = result.unwrap_err().to_string();
        assert!(err.contains("NonExistent"));
    }

    #[test]
    fn test_build_pivot_table_xml_no_rows_or_cols() {
        let config = PivotTableConfig {
            name: "DataOnly".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:B5".to_string(),
            target_sheet: "Pivot".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Amount".to_string(),
                function: AggregateFunction::Count,
                display_name: None,
            }],
        };
        let field_names = vec!["Amount".to_string()];

        let def = build_pivot_table_xml(&config, 0, &field_names).unwrap();
        assert!(def.row_fields.is_none());
        assert!(def.col_fields.is_none());
        assert!(def.data_fields.is_some());
    }

    #[test]
    fn test_build_pivot_cache_definition() {
        let field_names = vec![
            "Name".to_string(),
            "Region".to_string(),
            "Sales".to_string(),
        ];
        let def = build_pivot_cache_definition("Sheet1", "A1:C10", &field_names);

        assert_eq!(def.xmlns, sheetkit_xml::namespaces::SPREADSHEET_ML);
        assert_eq!(def.cache_source.source_type, "worksheet");
        let ws = def.cache_source.worksheet_source.unwrap();
        assert_eq!(ws.sheet, "Sheet1");
        assert_eq!(ws.reference, "A1:C10");

        assert_eq!(def.cache_fields.count, Some(3));
        assert_eq!(def.cache_fields.fields.len(), 3);
        assert_eq!(def.cache_fields.fields[0].name, "Name");
        assert_eq!(def.cache_fields.fields[1].name, "Region");
        assert_eq!(def.cache_fields.fields[2].name, "Sales");

        // Each field should have empty shared items
        for field in &def.cache_fields.fields {
            assert!(field.shared_items.is_some());
            let items = field.shared_items.as_ref().unwrap();
            assert_eq!(items.count, Some(0));
        }

        assert_eq!(def.record_count, Some(0));
        assert!(def.r_id.is_none());
    }

    #[test]
    fn test_build_pivot_cache_definition_empty_fields() {
        let field_names: Vec<String> = vec![];
        let def = build_pivot_cache_definition("Sheet1", "A1:A1", &field_names);
        assert_eq!(def.cache_fields.count, Some(0));
        assert!(def.cache_fields.fields.is_empty());
    }

    #[test]
    fn test_pivot_table_info_struct() {
        let info = PivotTableInfo {
            name: "PT1".to_string(),
            source_sheet: "Data".to_string(),
            source_range: "A1:D10".to_string(),
            target_sheet: "Report".to_string(),
            location: "A3:E20".to_string(),
        };
        assert_eq!(info.name, "PT1");
        assert_eq!(info.source_sheet, "Data");
        assert_eq!(info.source_range, "A1:D10");
        assert_eq!(info.target_sheet, "Report");
        assert_eq!(info.location, "A3:E20");
    }

    #[test]
    fn test_build_pivot_table_xml_generates_default_display_name() {
        let config = PivotTableConfig {
            name: "PT".to_string(),
            source_sheet: "S".to_string(),
            source_range: "A1:B2".to_string(),
            target_sheet: "T".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![],
            columns: vec![],
            data: vec![
                PivotDataField {
                    name: "Amount".to_string(),
                    function: AggregateFunction::Sum,
                    display_name: None,
                },
                PivotDataField {
                    name: "Count".to_string(),
                    function: AggregateFunction::Count,
                    display_name: Some("Total Count".to_string()),
                },
            ],
        };
        let field_names = vec!["Amount".to_string(), "Count".to_string()];

        let def = build_pivot_table_xml(&config, 0, &field_names).unwrap();
        let data_fields = def.data_fields.unwrap();

        // No display_name -> auto-generated
        assert_eq!(
            data_fields.fields[0].name,
            Some("Sum of Amount".to_string())
        );
        // Custom display_name preserved
        assert_eq!(data_fields.fields[1].name, Some("Total Count".to_string()));
    }

    #[test]
    fn test_error_pivot_table_not_found() {
        let err = Error::PivotTableNotFound {
            name: "Missing".to_string(),
        };
        assert_eq!(err.to_string(), "pivot table 'Missing' not found");
    }

    #[test]
    fn test_error_pivot_table_already_exists() {
        let err = Error::PivotTableAlreadyExists {
            name: "PT1".to_string(),
        };
        assert_eq!(err.to_string(), "pivot table 'PT1' already exists");
    }

    #[test]
    fn test_error_invalid_source_range() {
        let err = Error::InvalidSourceRange("bad range".to_string());
        assert_eq!(err.to_string(), "invalid source range: bad range");
    }
}
