//! Slicer XML schema structures.
//!
//! Represents `xl/slicers/slicer{N}.xml` and `xl/slicerCaches/slicerCache{N}.xml`
//! in the OOXML package. These are Office 2010+ extensions for visual filter controls.
//!
//! The slicer cache XML uses namespace-prefixed elements (`x15:tableSlicerCache`)
//! which cannot round-trip through serde. The `SlicerCacheDefinition` is therefore
//! serialized manually during workbook save.

use serde::{Deserialize, Serialize};

/// Root element for a slicer definition part (`xl/slicers/slicer{N}.xml`).
///
/// Contains one or more slicer definitions that describe visual filter controls
/// linked to table or pivot table columns.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "slicers")]
pub struct SlicerDefinitions {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "@xmlns:mc", skip_serializing_if = "Option::is_none")]
    pub xmlns_mc: Option<String>,

    #[serde(rename = "slicer", default)]
    pub slicers: Vec<SlicerDefinition>,
}

/// A single slicer definition within the slicers part.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SlicerDefinition {
    #[serde(rename = "@name")]
    pub name: String,

    #[serde(rename = "@cache")]
    pub cache: String,

    #[serde(rename = "@caption", skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,

    #[serde(rename = "@startItem", skip_serializing_if = "Option::is_none")]
    pub start_item: Option<u32>,

    #[serde(rename = "@columnCount", skip_serializing_if = "Option::is_none")]
    pub column_count: Option<u32>,

    #[serde(rename = "@showCaption", skip_serializing_if = "Option::is_none")]
    pub show_caption: Option<bool>,

    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,

    #[serde(rename = "@lockedPosition", skip_serializing_if = "Option::is_none")]
    pub locked_position: Option<bool>,

    #[serde(rename = "@rowHeight")]
    pub row_height: u32,
}

/// In-memory representation of a slicer cache definition.
///
/// Serialized manually (not via serde) because the XML uses namespace-prefixed
/// child elements (`x15:tableSlicerCache`) that serde cannot handle.
#[derive(Debug, Clone, PartialEq)]
pub struct SlicerCacheDefinition {
    pub name: String,
    pub source_name: String,
    pub table_slicer_cache: Option<TableSlicerCache>,
}

/// Tabular slicer cache linking the slicer to a table column.
#[derive(Debug, Clone, PartialEq)]
pub struct TableSlicerCache {
    pub table_id: u32,
    pub column: u32,
}

/// Serialize a `SlicerCacheDefinition` to XML manually.
///
/// Produces the full `slicerCacheDefinition` element with proper namespace
/// prefixes for the table slicer cache extension.
pub fn serialize_slicer_cache(def: &SlicerCacheDefinition) -> String {
    use std::fmt::Write;

    let ns_x14 = crate::namespaces::SLICER_2009;
    let ns_x15 = crate::namespaces::SLICER_2010;
    let ns_mc = crate::namespaces::MC;

    let mut xml = String::new();
    let _ = write!(
        xml,
        "<slicerCacheDefinition \
         xmlns=\"{ns_x14}\" \
         xmlns:mc=\"{ns_mc}\" \
         name=\"{}\" \
         sourceName=\"{}\"",
        escape_xml_attr(&def.name),
        escape_xml_attr(&def.source_name),
    );

    if let Some(ref tsc) = def.table_slicer_cache {
        let _ = write!(xml, ">");
        let _ = write!(
            xml,
            "<extLst>\
             <ext xmlns:x15=\"{ns_x15}\" \
             uri=\"{{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}}\">\
             <x15:tableSlicerCache tableId=\"{}\" column=\"{}\"/>\
             </ext>\
             </extLst>",
            tsc.table_id, tsc.column,
        );
        let _ = write!(xml, "</slicerCacheDefinition>");
    } else {
        let _ = write!(xml, "/>");
    }

    xml
}

/// Parse a `SlicerCacheDefinition` from XML.
///
/// Handles the namespace-prefixed `x15:tableSlicerCache` element that serde
/// cannot deserialize. Uses simple string matching for robustness.
pub fn parse_slicer_cache(xml: &str) -> Option<SlicerCacheDefinition> {
    let name = extract_attr(xml, "name")?;
    let source_name = extract_attr(xml, "sourceName")?;

    let table_slicer_cache = if let Some(tsc_start) = xml.find("tableSlicerCache") {
        let remainder = &xml[tsc_start..];
        let table_id = extract_attr(remainder, "tableId").and_then(|s| s.parse::<u32>().ok());
        let column = extract_attr(remainder, "column").and_then(|s| s.parse::<u32>().ok());
        match (table_id, column) {
            (Some(tid), Some(col)) => Some(TableSlicerCache {
                table_id: tid,
                column: col,
            }),
            _ => None,
        }
    } else {
        None
    };

    Some(SlicerCacheDefinition {
        name,
        source_name,
        table_slicer_cache,
    })
}

/// Extract a named XML attribute value from an element string.
fn extract_attr(xml: &str, attr_name: &str) -> Option<String> {
    let pattern = format!("{}=\"", attr_name);
    let start = xml.find(&pattern)?;
    let after_eq = start + pattern.len();
    let end = xml[after_eq..].find('"')?;
    Some(xml[after_eq..after_eq + end].to_string())
}

/// Escape special characters for use in XML attribute values.
fn escape_xml_attr(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_slicer_definition_roundtrip() {
        let defs = SlicerDefinitions {
            xmlns: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main".to_string(),
            xmlns_mc: None,
            slicers: vec![SlicerDefinition {
                name: "Slicer_Category".to_string(),
                cache: "Slicer_Category".to_string(),
                caption: Some("Category".to_string()),
                start_item: None,
                column_count: None,
                show_caption: Some(true),
                style: Some("SlicerStyleLight1".to_string()),
                locked_position: None,
                row_height: 241300,
            }],
        };

        let xml = quick_xml::se::to_string(&defs).unwrap();
        assert!(xml.contains("Slicer_Category"));
        assert!(xml.contains("SlicerStyleLight1"));

        let parsed: SlicerDefinitions = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.slicers.len(), 1);
        assert_eq!(parsed.slicers[0].name, "Slicer_Category");
        assert_eq!(parsed.slicers[0].row_height, 241300);
    }

    #[test]
    fn test_slicer_cache_serialize_and_parse() {
        let cache = SlicerCacheDefinition {
            name: "Slicer_Category".to_string(),
            source_name: "Category".to_string(),
            table_slicer_cache: Some(TableSlicerCache {
                table_id: 1,
                column: 2,
            }),
        };

        let xml = serialize_slicer_cache(&cache);
        assert!(xml.contains("Slicer_Category"));
        assert!(xml.contains("sourceName=\"Category\""));
        assert!(xml.contains("x15:tableSlicerCache"));
        assert!(xml.contains("tableId=\"1\""));
        assert!(xml.contains("column=\"2\""));

        let parsed = parse_slicer_cache(&xml).unwrap();
        assert_eq!(parsed.name, "Slicer_Category");
        assert_eq!(parsed.source_name, "Category");
        let tsc = parsed.table_slicer_cache.unwrap();
        assert_eq!(tsc.table_id, 1);
        assert_eq!(tsc.column, 2);
    }

    #[test]
    fn test_slicer_definition_minimal() {
        let def = SlicerDefinition {
            name: "S1".to_string(),
            cache: "S1".to_string(),
            caption: None,
            start_item: None,
            column_count: None,
            show_caption: None,
            style: None,
            locked_position: None,
            row_height: 241300,
        };

        let xml = quick_xml::se::to_string(&def).unwrap();
        assert!(!xml.contains("caption="));
        assert!(!xml.contains("style="));
        assert!(xml.contains("rowHeight=\"241300\""));
    }

    #[test]
    fn test_multiple_slicers_in_definitions() {
        let defs = SlicerDefinitions {
            xmlns: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main".to_string(),
            xmlns_mc: None,
            slicers: vec![
                SlicerDefinition {
                    name: "Slicer_A".to_string(),
                    cache: "Slicer_A".to_string(),
                    caption: Some("Column A".to_string()),
                    start_item: None,
                    column_count: Some(2),
                    show_caption: Some(true),
                    style: Some("SlicerStyleDark1".to_string()),
                    locked_position: None,
                    row_height: 241300,
                },
                SlicerDefinition {
                    name: "Slicer_B".to_string(),
                    cache: "Slicer_B".to_string(),
                    caption: Some("Column B".to_string()),
                    start_item: None,
                    column_count: None,
                    show_caption: Some(false),
                    style: None,
                    locked_position: Some(true),
                    row_height: 241300,
                },
            ],
        };

        let xml = quick_xml::se::to_string(&defs).unwrap();
        let parsed: SlicerDefinitions = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.slicers.len(), 2);
        assert_eq!(parsed.slicers[0].name, "Slicer_A");
        assert_eq!(parsed.slicers[1].name, "Slicer_B");
    }

    #[test]
    fn test_slicer_cache_without_table_cache() {
        let cache = SlicerCacheDefinition {
            name: "Slicer_X".to_string(),
            source_name: "ColumnX".to_string(),
            table_slicer_cache: None,
        };

        let xml = serialize_slicer_cache(&cache);
        assert!(!xml.contains("tableSlicerCache"));
        assert!(xml.contains("Slicer_X"));

        let parsed = parse_slicer_cache(&xml).unwrap();
        assert_eq!(parsed.name, "Slicer_X");
        assert!(parsed.table_slicer_cache.is_none());
    }

    #[test]
    fn test_slicer_cache_escapes_special_chars() {
        let cache = SlicerCacheDefinition {
            name: "Slicer_A&B".to_string(),
            source_name: "Col<1>".to_string(),
            table_slicer_cache: None,
        };

        let xml = serialize_slicer_cache(&cache);
        assert!(xml.contains("Slicer_A&amp;B"));
        assert!(xml.contains("Col&lt;1&gt;"));
    }

    #[test]
    fn test_extract_attr() {
        let xml = r#"<elem name="hello" id="42">"#;
        assert_eq!(extract_attr(xml, "name"), Some("hello".to_string()));
        assert_eq!(extract_attr(xml, "id"), Some("42".to_string()));
        assert_eq!(extract_attr(xml, "missing"), None);
    }
}
