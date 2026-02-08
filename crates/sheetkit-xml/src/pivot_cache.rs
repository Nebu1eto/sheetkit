//! Pivot cache XML schema structures.
//!
//! Represents `xl/pivotCache/pivotCacheDefinition{N}.xml` and
//! `xl/pivotCache/pivotCacheRecords{N}.xml`.

use serde::{Deserialize, Serialize};

/// Root element for pivot cache definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "pivotCacheDefinition")]
pub struct PivotCacheDefinition {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "@xmlns:r")]
    pub xmlns_r: String,

    #[serde(rename = "@r:id", skip_serializing_if = "Option::is_none")]
    pub r_id: Option<String>,

    #[serde(rename = "@recordCount", skip_serializing_if = "Option::is_none")]
    pub record_count: Option<u32>,

    #[serde(rename = "cacheSource")]
    pub cache_source: CacheSource,

    #[serde(rename = "cacheFields")]
    pub cache_fields: CacheFields,
}

/// Source of data for the pivot cache.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CacheSource {
    #[serde(rename = "@type")]
    pub source_type: String,

    #[serde(rename = "worksheetSource", skip_serializing_if = "Option::is_none")]
    pub worksheet_source: Option<WorksheetSource>,
}

/// Worksheet-based data source.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct WorksheetSource {
    #[serde(rename = "@ref")]
    pub reference: String,

    #[serde(rename = "@sheet")]
    pub sheet: String,
}

/// Container for cache field definitions.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CacheFields {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "cacheField", default)]
    pub fields: Vec<CacheField>,
}

/// Individual cache field definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CacheField {
    #[serde(rename = "@name")]
    pub name: String,

    #[serde(rename = "@numFmtId", skip_serializing_if = "Option::is_none")]
    pub num_fmt_id: Option<u32>,

    #[serde(rename = "sharedItems", skip_serializing_if = "Option::is_none")]
    pub shared_items: Option<SharedItems>,
}

/// Shared items within a cache field.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SharedItems {
    #[serde(
        rename = "@containsSemiMixedTypes",
        skip_serializing_if = "Option::is_none"
    )]
    pub contains_semi_mixed_types: Option<bool>,

    #[serde(rename = "@containsString", skip_serializing_if = "Option::is_none")]
    pub contains_string: Option<bool>,

    #[serde(rename = "@containsNumber", skip_serializing_if = "Option::is_none")]
    pub contains_number: Option<bool>,

    #[serde(rename = "@containsBlank", skip_serializing_if = "Option::is_none")]
    pub contains_blank: Option<bool>,

    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "s", default)]
    pub string_items: Vec<StringItem>,

    #[serde(rename = "n", default)]
    pub number_items: Vec<NumberItem>,
}

/// String value in shared items.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct StringItem {
    #[serde(rename = "@v")]
    pub value: String,
}

/// Numeric value in shared items.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct NumberItem {
    #[serde(rename = "@v")]
    pub value: f64,
}

/// Root element for pivot cache records.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "pivotCacheRecords")]
pub struct PivotCacheRecords {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "@xmlns:r")]
    pub xmlns_r: String,

    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "r", default)]
    pub records: Vec<CacheRecord>,
}

/// Individual cache record containing field values.
///
/// Each record uses separate optional vectors for each value type because
/// serde + quick-xml does not reliably roundtrip internally-tagged enums
/// within a mixed-content element. This flat representation is simpler and
/// still captures the data faithfully.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CacheRecord {
    #[serde(rename = "x", default)]
    pub index_fields: Vec<IndexField>,

    #[serde(rename = "n", default)]
    pub number_fields: Vec<NumberField>,

    #[serde(rename = "s", default)]
    pub string_fields: Vec<StringField>,

    #[serde(rename = "b", default)]
    pub bool_fields: Vec<BoolField>,
}

/// Index reference field in a cache record.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct IndexField {
    #[serde(rename = "@v")]
    pub v: u32,
}

/// Number field in a cache record.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct NumberField {
    #[serde(rename = "@v")]
    pub v: f64,
}

/// String field in a cache record.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct StringField {
    #[serde(rename = "@v")]
    pub v: String,
}

/// Boolean field in a cache record.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BoolField {
    #[serde(rename = "@v")]
    pub v: bool,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_worksheet_source_roundtrip() {
        let src = WorksheetSource {
            reference: "A1:D10".to_string(),
            sheet: "Sheet1".to_string(),
        };
        let xml = quick_xml::se::to_string(&src).unwrap();
        let parsed: WorksheetSource = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(src, parsed);
    }

    #[test]
    fn test_cache_source_roundtrip() {
        let src = CacheSource {
            source_type: "worksheet".to_string(),
            worksheet_source: Some(WorksheetSource {
                reference: "A1:D10".to_string(),
                sheet: "Data".to_string(),
            }),
        };
        let xml = quick_xml::se::to_string(&src).unwrap();
        let parsed: CacheSource = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(src, parsed);
    }

    #[test]
    fn test_cache_source_without_worksheet() {
        let src = CacheSource {
            source_type: "external".to_string(),
            worksheet_source: None,
        };
        let xml = quick_xml::se::to_string(&src).unwrap();
        assert!(!xml.contains("worksheetSource"));
        let parsed: CacheSource = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(src, parsed);
    }

    #[test]
    fn test_string_item_roundtrip() {
        let item = StringItem {
            value: "North".to_string(),
        };
        let xml = quick_xml::se::to_string(&item).unwrap();
        let parsed: StringItem = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(item, parsed);
    }

    #[test]
    fn test_number_item_roundtrip() {
        let item = NumberItem { value: 42.5 };
        let xml = quick_xml::se::to_string(&item).unwrap();
        let parsed: NumberItem = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(item, parsed);
    }

    #[test]
    fn test_shared_items_with_strings_roundtrip() {
        let items = SharedItems {
            contains_semi_mixed_types: Some(false),
            contains_string: Some(true),
            contains_number: Some(false),
            contains_blank: None,
            count: Some(3),
            string_items: vec![
                StringItem {
                    value: "North".to_string(),
                },
                StringItem {
                    value: "South".to_string(),
                },
                StringItem {
                    value: "East".to_string(),
                },
            ],
            number_items: vec![],
        };
        let xml = quick_xml::se::to_string(&items).unwrap();
        let parsed: SharedItems = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(items, parsed);
    }

    #[test]
    fn test_shared_items_with_numbers_roundtrip() {
        let items = SharedItems {
            contains_semi_mixed_types: None,
            contains_string: Some(false),
            contains_number: Some(true),
            contains_blank: None,
            count: Some(2),
            string_items: vec![],
            number_items: vec![NumberItem { value: 100.0 }, NumberItem { value: 200.0 }],
        };
        let xml = quick_xml::se::to_string(&items).unwrap();
        let parsed: SharedItems = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(items, parsed);
    }

    #[test]
    fn test_shared_items_empty_roundtrip() {
        let items = SharedItems {
            contains_semi_mixed_types: None,
            contains_string: None,
            contains_number: None,
            contains_blank: None,
            count: Some(0),
            string_items: vec![],
            number_items: vec![],
        };
        let xml = quick_xml::se::to_string(&items).unwrap();
        let parsed: SharedItems = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(items, parsed);
    }

    #[test]
    fn test_cache_field_roundtrip() {
        let field = CacheField {
            name: "Region".to_string(),
            num_fmt_id: Some(0),
            shared_items: Some(SharedItems {
                contains_semi_mixed_types: None,
                contains_string: Some(true),
                contains_number: None,
                contains_blank: None,
                count: Some(2),
                string_items: vec![
                    StringItem {
                        value: "North".to_string(),
                    },
                    StringItem {
                        value: "South".to_string(),
                    },
                ],
                number_items: vec![],
            }),
        };
        let xml = quick_xml::se::to_string(&field).unwrap();
        let parsed: CacheField = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(field, parsed);
    }

    #[test]
    fn test_cache_field_no_shared_items() {
        let field = CacheField {
            name: "Amount".to_string(),
            num_fmt_id: None,
            shared_items: None,
        };
        let xml = quick_xml::se::to_string(&field).unwrap();
        assert!(!xml.contains("sharedItems"));
        assert!(!xml.contains("numFmtId"));
        let parsed: CacheField = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(field, parsed);
    }

    #[test]
    fn test_cache_fields_roundtrip() {
        let fields = CacheFields {
            count: Some(2),
            fields: vec![
                CacheField {
                    name: "Region".to_string(),
                    num_fmt_id: Some(0),
                    shared_items: None,
                },
                CacheField {
                    name: "Sales".to_string(),
                    num_fmt_id: Some(0),
                    shared_items: None,
                },
            ],
        };
        let xml = quick_xml::se::to_string(&fields).unwrap();
        let parsed: CacheFields = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(fields, parsed);
    }

    #[test]
    fn test_pivot_cache_definition_roundtrip() {
        let def = PivotCacheDefinition {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            xmlns_r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                .to_string(),
            r_id: None,
            record_count: Some(5),
            cache_source: CacheSource {
                source_type: "worksheet".to_string(),
                worksheet_source: Some(WorksheetSource {
                    reference: "A1:C6".to_string(),
                    sheet: "Data".to_string(),
                }),
            },
            cache_fields: CacheFields {
                count: Some(3),
                fields: vec![
                    CacheField {
                        name: "Name".to_string(),
                        num_fmt_id: Some(0),
                        shared_items: None,
                    },
                    CacheField {
                        name: "Region".to_string(),
                        num_fmt_id: Some(0),
                        shared_items: None,
                    },
                    CacheField {
                        name: "Sales".to_string(),
                        num_fmt_id: Some(0),
                        shared_items: None,
                    },
                ],
            },
        };
        let xml = quick_xml::se::to_string(&def).unwrap();
        let parsed: PivotCacheDefinition = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(def, parsed);
    }

    #[test]
    fn test_pivot_cache_definition_structure() {
        let def = PivotCacheDefinition {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            xmlns_r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                .to_string(),
            r_id: Some("rId1".to_string()),
            record_count: Some(10),
            cache_source: CacheSource {
                source_type: "worksheet".to_string(),
                worksheet_source: Some(WorksheetSource {
                    reference: "A1:D11".to_string(),
                    sheet: "Sheet1".to_string(),
                }),
            },
            cache_fields: CacheFields {
                count: Some(1),
                fields: vec![CacheField {
                    name: "Col1".to_string(),
                    num_fmt_id: None,
                    shared_items: None,
                }],
            },
        };
        let xml = quick_xml::se::to_string(&def).unwrap();
        assert!(xml.contains("<pivotCacheDefinition"));
        assert!(xml.contains("recordCount=\"10\""));
        assert!(xml.contains("<cacheSource"));
        assert!(xml.contains("type=\"worksheet\""));
        assert!(xml.contains("<worksheetSource"));
        assert!(xml.contains("<cacheFields"));
    }

    #[test]
    fn test_index_field_roundtrip() {
        let field = IndexField { v: 3 };
        let xml = quick_xml::se::to_string(&field).unwrap();
        let parsed: IndexField = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(field, parsed);
    }

    #[test]
    fn test_number_field_roundtrip() {
        let field = NumberField { v: 99.5 };
        let xml = quick_xml::se::to_string(&field).unwrap();
        let parsed: NumberField = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(field, parsed);
    }

    #[test]
    fn test_string_field_roundtrip() {
        let field = StringField {
            v: "hello".to_string(),
        };
        let xml = quick_xml::se::to_string(&field).unwrap();
        let parsed: StringField = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(field, parsed);
    }

    #[test]
    fn test_bool_field_roundtrip() {
        let field = BoolField { v: true };
        let xml = quick_xml::se::to_string(&field).unwrap();
        let parsed: BoolField = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(field, parsed);
    }

    #[test]
    fn test_cache_record_roundtrip() {
        let record = CacheRecord {
            index_fields: vec![IndexField { v: 0 }, IndexField { v: 1 }],
            number_fields: vec![NumberField { v: 150.0 }],
            string_fields: vec![],
            bool_fields: vec![],
        };
        let xml = quick_xml::se::to_string(&record).unwrap();
        let parsed: CacheRecord = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(record, parsed);
    }

    #[test]
    fn test_cache_record_with_strings() {
        let record = CacheRecord {
            index_fields: vec![],
            number_fields: vec![],
            string_fields: vec![
                StringField {
                    v: "alpha".to_string(),
                },
                StringField {
                    v: "beta".to_string(),
                },
            ],
            bool_fields: vec![BoolField { v: false }],
        };
        let xml = quick_xml::se::to_string(&record).unwrap();
        let parsed: CacheRecord = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(record, parsed);
    }

    #[test]
    fn test_pivot_cache_records_roundtrip() {
        let records = PivotCacheRecords {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            xmlns_r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                .to_string(),
            count: Some(2),
            records: vec![
                CacheRecord {
                    index_fields: vec![IndexField { v: 0 }],
                    number_fields: vec![NumberField { v: 100.0 }],
                    string_fields: vec![],
                    bool_fields: vec![],
                },
                CacheRecord {
                    index_fields: vec![IndexField { v: 1 }],
                    number_fields: vec![NumberField { v: 200.0 }],
                    string_fields: vec![],
                    bool_fields: vec![],
                },
            ],
        };
        let xml = quick_xml::se::to_string(&records).unwrap();
        let parsed: PivotCacheRecords = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(records, parsed);
    }

    #[test]
    fn test_pivot_cache_records_empty() {
        let records = PivotCacheRecords {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            xmlns_r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                .to_string(),
            count: Some(0),
            records: vec![],
        };
        let xml = quick_xml::se::to_string(&records).unwrap();
        let parsed: PivotCacheRecords = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(records, parsed);
    }

    #[test]
    fn test_pivot_cache_records_structure() {
        let records = PivotCacheRecords {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            xmlns_r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                .to_string(),
            count: Some(1),
            records: vec![CacheRecord {
                index_fields: vec![IndexField { v: 0 }],
                number_fields: vec![],
                string_fields: vec![],
                bool_fields: vec![],
            }],
        };
        let xml = quick_xml::se::to_string(&records).unwrap();
        assert!(xml.contains("<pivotCacheRecords"));
        assert!(xml.contains("count=\"1\""));
    }
}
