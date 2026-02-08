//! Pivot table XML schema structures.
//!
//! Represents `xl/pivotTables/pivotTable{N}.xml` in the OOXML package.

use serde::{Deserialize, Serialize};

/// Root element for a pivot table definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "pivotTableDefinition")]
pub struct PivotTableDefinition {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "@name")]
    pub name: String,

    #[serde(rename = "@cacheId")]
    pub cache_id: u32,

    #[serde(rename = "@dataOnRows", skip_serializing_if = "Option::is_none")]
    pub data_on_rows: Option<bool>,

    #[serde(
        rename = "@applyNumberFormats",
        skip_serializing_if = "Option::is_none"
    )]
    pub apply_number_formats: Option<bool>,

    #[serde(
        rename = "@applyBorderFormats",
        skip_serializing_if = "Option::is_none"
    )]
    pub apply_border_formats: Option<bool>,

    #[serde(rename = "@applyFontFormats", skip_serializing_if = "Option::is_none")]
    pub apply_font_formats: Option<bool>,

    #[serde(
        rename = "@applyPatternFormats",
        skip_serializing_if = "Option::is_none"
    )]
    pub apply_pattern_formats: Option<bool>,

    #[serde(
        rename = "@applyAlignmentFormats",
        skip_serializing_if = "Option::is_none"
    )]
    pub apply_alignment_formats: Option<bool>,

    #[serde(
        rename = "@applyWidthHeightFormats",
        skip_serializing_if = "Option::is_none"
    )]
    pub apply_width_height_formats: Option<bool>,

    #[serde(rename = "location")]
    pub location: PivotLocation,

    #[serde(rename = "pivotFields")]
    pub pivot_fields: PivotFields,

    #[serde(rename = "rowFields", skip_serializing_if = "Option::is_none")]
    pub row_fields: Option<FieldList>,

    #[serde(rename = "colFields", skip_serializing_if = "Option::is_none")]
    pub col_fields: Option<FieldList>,

    #[serde(rename = "dataFields", skip_serializing_if = "Option::is_none")]
    pub data_fields: Option<DataFields>,
}

/// Location of the pivot table within the worksheet.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PivotLocation {
    #[serde(rename = "@ref")]
    pub reference: String,

    #[serde(rename = "@firstHeaderRow")]
    pub first_header_row: u32,

    #[serde(rename = "@firstDataRow")]
    pub first_data_row: u32,

    #[serde(rename = "@firstDataCol")]
    pub first_data_col: u32,
}

/// Container for pivot field definitions.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PivotFields {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "pivotField", default)]
    pub fields: Vec<PivotFieldDef>,
}

/// Individual pivot field definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PivotFieldDef {
    #[serde(rename = "@axis", skip_serializing_if = "Option::is_none")]
    pub axis: Option<String>,

    #[serde(rename = "@dataField", skip_serializing_if = "Option::is_none")]
    pub data_field: Option<bool>,

    #[serde(rename = "@showAll", skip_serializing_if = "Option::is_none")]
    pub show_all: Option<bool>,

    #[serde(rename = "items", skip_serializing_if = "Option::is_none")]
    pub items: Option<FieldItems>,
}

/// Container for field items.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FieldItems {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "item", default)]
    pub items: Vec<FieldItem>,
}

/// Individual field item entry.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FieldItem {
    #[serde(rename = "@t", skip_serializing_if = "Option::is_none")]
    pub item_type: Option<String>,

    #[serde(rename = "@x", skip_serializing_if = "Option::is_none")]
    pub index: Option<u32>,
}

/// List of field references (used for row and column fields).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FieldList {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "field", default)]
    pub fields: Vec<FieldRef>,
}

/// Reference to a pivot field by index.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FieldRef {
    #[serde(rename = "@x")]
    pub index: i32,
}

/// Container for data fields.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DataFields {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "dataField", default)]
    pub fields: Vec<DataFieldDef>,
}

/// Individual data field definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DataFieldDef {
    #[serde(rename = "@name", skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,

    #[serde(rename = "@fld")]
    pub field_index: u32,

    #[serde(rename = "@subtotal", skip_serializing_if = "Option::is_none")]
    pub subtotal: Option<String>,

    #[serde(rename = "@baseField", skip_serializing_if = "Option::is_none")]
    pub base_field: Option<i32>,

    #[serde(rename = "@baseItem", skip_serializing_if = "Option::is_none")]
    pub base_item: Option<u32>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_pivot_location_roundtrip() {
        let loc = PivotLocation {
            reference: "A3:D20".to_string(),
            first_header_row: 1,
            first_data_row: 2,
            first_data_col: 1,
        };
        let xml = quick_xml::se::to_string(&loc).unwrap();
        let parsed: PivotLocation = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(loc, parsed);
    }

    #[test]
    fn test_field_item_roundtrip() {
        let item = FieldItem {
            item_type: Some("default".to_string()),
            index: Some(0),
        };
        let xml = quick_xml::se::to_string(&item).unwrap();
        let parsed: FieldItem = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(item, parsed);
    }

    #[test]
    fn test_field_item_optional_fields_skipped() {
        let item = FieldItem {
            item_type: None,
            index: Some(3),
        };
        let xml = quick_xml::se::to_string(&item).unwrap();
        assert!(!xml.contains("t="));
        assert!(xml.contains("x=\"3\""));
    }

    #[test]
    fn test_field_items_roundtrip() {
        let items = FieldItems {
            count: Some(2),
            items: vec![
                FieldItem {
                    item_type: None,
                    index: Some(0),
                },
                FieldItem {
                    item_type: Some("default".to_string()),
                    index: None,
                },
            ],
        };
        let xml = quick_xml::se::to_string(&items).unwrap();
        let parsed: FieldItems = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(items, parsed);
    }

    #[test]
    fn test_pivot_field_def_roundtrip() {
        let field = PivotFieldDef {
            axis: Some("axisRow".to_string()),
            data_field: None,
            show_all: Some(false),
            items: Some(FieldItems {
                count: Some(1),
                items: vec![FieldItem {
                    item_type: Some("default".to_string()),
                    index: None,
                }],
            }),
        };
        let xml = quick_xml::se::to_string(&field).unwrap();
        let parsed: PivotFieldDef = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(field, parsed);
    }

    #[test]
    fn test_pivot_field_def_no_axis() {
        let field = PivotFieldDef {
            axis: None,
            data_field: Some(true),
            show_all: Some(false),
            items: None,
        };
        let xml = quick_xml::se::to_string(&field).unwrap();
        assert!(!xml.contains("axis="));
        assert!(xml.contains("dataField=\"true\""));
    }

    #[test]
    fn test_pivot_fields_roundtrip() {
        let fields = PivotFields {
            count: Some(2),
            fields: vec![
                PivotFieldDef {
                    axis: Some("axisRow".to_string()),
                    data_field: None,
                    show_all: Some(false),
                    items: None,
                },
                PivotFieldDef {
                    axis: None,
                    data_field: Some(true),
                    show_all: Some(false),
                    items: None,
                },
            ],
        };
        let xml = quick_xml::se::to_string(&fields).unwrap();
        let parsed: PivotFields = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(fields, parsed);
    }

    #[test]
    fn test_field_ref_roundtrip() {
        let field_ref = FieldRef { index: 2 };
        let xml = quick_xml::se::to_string(&field_ref).unwrap();
        let parsed: FieldRef = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(field_ref, parsed);
    }

    #[test]
    fn test_field_ref_negative_index() {
        let field_ref = FieldRef { index: -2 };
        let xml = quick_xml::se::to_string(&field_ref).unwrap();
        let parsed: FieldRef = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.index, -2);
    }

    #[test]
    fn test_field_list_roundtrip() {
        let list = FieldList {
            count: Some(2),
            fields: vec![FieldRef { index: 0 }, FieldRef { index: 3 }],
        };
        let xml = quick_xml::se::to_string(&list).unwrap();
        let parsed: FieldList = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(list, parsed);
    }

    #[test]
    fn test_data_field_def_roundtrip() {
        let data_field = DataFieldDef {
            name: Some("Sum of Sales".to_string()),
            field_index: 2,
            subtotal: Some("sum".to_string()),
            base_field: Some(0),
            base_item: Some(0),
        };
        let xml = quick_xml::se::to_string(&data_field).unwrap();
        let parsed: DataFieldDef = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(data_field, parsed);
    }

    #[test]
    fn test_data_field_def_optional_fields_skipped() {
        let data_field = DataFieldDef {
            name: None,
            field_index: 1,
            subtotal: None,
            base_field: None,
            base_item: None,
        };
        let xml = quick_xml::se::to_string(&data_field).unwrap();
        assert!(!xml.contains("name="));
        assert!(!xml.contains("subtotal="));
        assert!(!xml.contains("baseField="));
        assert!(!xml.contains("baseItem="));
        assert!(xml.contains("fld=\"1\""));
    }

    #[test]
    fn test_data_fields_roundtrip() {
        let data_fields = DataFields {
            count: Some(1),
            fields: vec![DataFieldDef {
                name: Some("Count of Items".to_string()),
                field_index: 0,
                subtotal: Some("count".to_string()),
                base_field: Some(0),
                base_item: Some(0),
            }],
        };
        let xml = quick_xml::se::to_string(&data_fields).unwrap();
        let parsed: DataFields = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(data_fields, parsed);
    }

    #[test]
    fn test_pivot_table_definition_minimal_roundtrip() {
        let def = PivotTableDefinition {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            name: "PivotTable1".to_string(),
            cache_id: 0,
            data_on_rows: None,
            apply_number_formats: None,
            apply_border_formats: None,
            apply_font_formats: None,
            apply_pattern_formats: None,
            apply_alignment_formats: None,
            apply_width_height_formats: None,
            location: PivotLocation {
                reference: "A3:C20".to_string(),
                first_header_row: 1,
                first_data_row: 1,
                first_data_col: 1,
            },
            pivot_fields: PivotFields {
                count: Some(2),
                fields: vec![
                    PivotFieldDef {
                        axis: Some("axisRow".to_string()),
                        data_field: None,
                        show_all: Some(false),
                        items: None,
                    },
                    PivotFieldDef {
                        axis: None,
                        data_field: Some(true),
                        show_all: Some(false),
                        items: None,
                    },
                ],
            },
            row_fields: Some(FieldList {
                count: Some(1),
                fields: vec![FieldRef { index: 0 }],
            }),
            col_fields: None,
            data_fields: Some(DataFields {
                count: Some(1),
                fields: vec![DataFieldDef {
                    name: Some("Sum of Amount".to_string()),
                    field_index: 1,
                    subtotal: Some("sum".to_string()),
                    base_field: Some(0),
                    base_item: Some(0),
                }],
            }),
        };
        let xml = quick_xml::se::to_string(&def).unwrap();
        let parsed: PivotTableDefinition = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(def, parsed);
    }

    #[test]
    fn test_pivot_table_definition_full_roundtrip() {
        let def = PivotTableDefinition {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            name: "SalesReport".to_string(),
            cache_id: 1,
            data_on_rows: Some(false),
            apply_number_formats: Some(false),
            apply_border_formats: Some(false),
            apply_font_formats: Some(false),
            apply_pattern_formats: Some(false),
            apply_alignment_formats: Some(false),
            apply_width_height_formats: Some(true),
            location: PivotLocation {
                reference: "A1:E30".to_string(),
                first_header_row: 1,
                first_data_row: 2,
                first_data_col: 1,
            },
            pivot_fields: PivotFields {
                count: Some(3),
                fields: vec![
                    PivotFieldDef {
                        axis: Some("axisRow".to_string()),
                        data_field: None,
                        show_all: Some(false),
                        items: Some(FieldItems {
                            count: Some(2),
                            items: vec![
                                FieldItem {
                                    item_type: None,
                                    index: Some(0),
                                },
                                FieldItem {
                                    item_type: Some("default".to_string()),
                                    index: None,
                                },
                            ],
                        }),
                    },
                    PivotFieldDef {
                        axis: Some("axisCol".to_string()),
                        data_field: None,
                        show_all: Some(false),
                        items: None,
                    },
                    PivotFieldDef {
                        axis: None,
                        data_field: Some(true),
                        show_all: Some(false),
                        items: None,
                    },
                ],
            },
            row_fields: Some(FieldList {
                count: Some(1),
                fields: vec![FieldRef { index: 0 }],
            }),
            col_fields: Some(FieldList {
                count: Some(1),
                fields: vec![FieldRef { index: 1 }],
            }),
            data_fields: Some(DataFields {
                count: Some(1),
                fields: vec![DataFieldDef {
                    name: Some("Sum of Revenue".to_string()),
                    field_index: 2,
                    subtotal: Some("sum".to_string()),
                    base_field: Some(0),
                    base_item: Some(0),
                }],
            }),
        };
        let xml = quick_xml::se::to_string(&def).unwrap();
        let parsed: PivotTableDefinition = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(def, parsed);
    }

    #[test]
    fn test_pivot_table_definition_serialization_structure() {
        let def = PivotTableDefinition {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            name: "TestPivot".to_string(),
            cache_id: 0,
            data_on_rows: Some(false),
            apply_number_formats: None,
            apply_border_formats: None,
            apply_font_formats: None,
            apply_pattern_formats: None,
            apply_alignment_formats: None,
            apply_width_height_formats: None,
            location: PivotLocation {
                reference: "A1".to_string(),
                first_header_row: 1,
                first_data_row: 1,
                first_data_col: 1,
            },
            pivot_fields: PivotFields {
                count: Some(0),
                fields: vec![],
            },
            row_fields: None,
            col_fields: None,
            data_fields: None,
        };
        let xml = quick_xml::se::to_string(&def).unwrap();
        assert!(xml.contains("<pivotTableDefinition"));
        assert!(xml.contains("name=\"TestPivot\""));
        assert!(xml.contains("cacheId=\"0\""));
        assert!(xml.contains("<location"));
        assert!(xml.contains("<pivotFields"));
        assert!(!xml.contains("<rowFields"));
        assert!(!xml.contains("<colFields"));
        assert!(!xml.contains("<dataFields"));
    }

    #[test]
    fn test_empty_pivot_fields() {
        let fields = PivotFields {
            count: Some(0),
            fields: vec![],
        };
        let xml = quick_xml::se::to_string(&fields).unwrap();
        let parsed: PivotFields = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.count, Some(0));
        assert!(parsed.fields.is_empty());
    }

    #[test]
    fn test_empty_field_list() {
        let list = FieldList {
            count: Some(0),
            fields: vec![],
        };
        let xml = quick_xml::se::to_string(&list).unwrap();
        let parsed: FieldList = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.count, Some(0));
        assert!(parsed.fields.is_empty());
    }

    #[test]
    fn test_empty_data_fields() {
        let data_fields = DataFields {
            count: Some(0),
            fields: vec![],
        };
        let xml = quick_xml::se::to_string(&data_fields).unwrap();
        let parsed: DataFields = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.count, Some(0));
        assert!(parsed.fields.is_empty());
    }
}
