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
/// Uses namespace-aware event parsing, so element prefixes are irrelevant and
/// attribute values are entity-decoded with either quote style. The root must
/// be a `slicerCacheDefinition` element with unprefixed `name` and
/// `sourceName` attributes. The first `tableSlicerCache` element that is not
/// bound to a foreign namespace supplies the table link. Returns `None` for
/// malformed XML or a missing required attribute.
pub fn parse_slicer_cache(xml: &str) -> Option<SlicerCacheDefinition> {
    use quick_xml::events::Event;
    use quick_xml::name::{Namespace, ResolveResult};
    use quick_xml::NsReader;

    let x15 = Namespace(crate::namespaces::SLICER_2010.as_bytes());
    let mut reader = NsReader::from_str(xml);
    let mut root: Option<(String, String)> = None;
    let mut table_slicer_cache = None;

    loop {
        let (ns, event) = reader.read_resolved_event().ok()?;
        let element = match event {
            Event::Start(e) | Event::Empty(e) => e,
            Event::Eof => break,
            _ => continue,
        };
        let local = element.local_name();
        if root.is_none() {
            if local.as_ref() != b"slicerCacheDefinition" {
                return None;
            }
            let name = unprefixed_attr(&element, b"name")?;
            let source_name = unprefixed_attr(&element, b"sourceName")?;
            root = Some((name, source_name));
            continue;
        }
        let in_x15_or_unbound = match ns {
            ResolveResult::Bound(bound) => bound == x15,
            ResolveResult::Unbound | ResolveResult::Unknown(_) => true,
        };
        if table_slicer_cache.is_none()
            && local.as_ref() == b"tableSlicerCache"
            && in_x15_or_unbound
        {
            let number =
                |key: &[u8]| -> Option<u32> { unprefixed_attr(&element, key)?.trim().parse().ok() };
            if let (Some(table_id), Some(column)) = (number(b"tableId"), number(b"column")) {
                table_slicer_cache = Some(TableSlicerCache { table_id, column });
            }
        }
    }

    let (name, source_name) = root?;
    Some(SlicerCacheDefinition {
        name,
        source_name,
        table_slicer_cache,
    })
}

/// Look up an unprefixed attribute by name and entity-decode its value.
///
/// Returns `None` when the attribute is absent or its value contains an
/// invalid entity reference.
fn unprefixed_attr(element: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<String> {
    let attr = element
        .attributes()
        .flatten()
        .find(|attr| attr.key.as_ref() == key)?;
    attr.unescape_value().ok().map(|v| v.into_owned())
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
    fn test_slicer_cache_special_chars_roundtrip_multi_cycle() {
        let cache = SlicerCacheDefinition {
            name: "Slicer_A&B".to_string(),
            source_name: "A&B <c> \"d\" 'e'".to_string(),
            table_slicer_cache: Some(TableSlicerCache {
                table_id: 3,
                column: 4,
            }),
        };

        let mut current = cache.clone();
        for _ in 0..3 {
            let xml = serialize_slicer_cache(&current);
            current = parse_slicer_cache(&xml).unwrap();
            assert_eq!(current, cache);
        }
    }

    #[test]
    fn test_parse_slicer_cache_single_quoted_attributes() {
        let xml = r#"<slicerCacheDefinition xmlns='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main' name='Slicer_Q' sourceName='Col &amp; Q'><extLst><ext uri='{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}' xmlns:x15='http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'><x15:tableSlicerCache tableId='5' column='6'/></ext></extLst></slicerCacheDefinition>"#;

        let parsed = parse_slicer_cache(xml).unwrap();
        assert_eq!(parsed.name, "Slicer_Q");
        assert_eq!(parsed.source_name, "Col & Q");
        assert_eq!(
            parsed.table_slicer_cache,
            Some(TableSlicerCache {
                table_id: 5,
                column: 6
            })
        );
    }

    #[test]
    fn test_parse_slicer_cache_alternate_namespace_prefix() {
        let xml = r#"<sc:slicerCacheDefinition xmlns:sc="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" sourceName="Region"
    name="Slicer_Region"><sc:extLst><sc:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:t="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><t:tableSlicerCache column="2" tableId="9"/></sc:ext></sc:extLst></sc:slicerCacheDefinition>"#;

        let parsed = parse_slicer_cache(xml).unwrap();
        assert_eq!(parsed.name, "Slicer_Region");
        assert_eq!(parsed.source_name, "Region");
        assert_eq!(
            parsed.table_slicer_cache,
            Some(TableSlicerCache {
                table_id: 9,
                column: 2
            })
        );
    }

    #[test]
    fn test_parse_slicer_cache_ignores_lookalike_attributes() {
        // `xname` and a prefixed `p:name` must not satisfy the `name` lookup,
        // and an element in a foreign namespace is not a table slicer cache.
        let xml = r#"<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:p="urn:other" xname="wrong" p:name="wrong" name="right" sourceName="S"><extLst><ext xmlns:o="urn:other"><o:tableSlicerCache tableId="1" column="1"/></ext></extLst></slicerCacheDefinition>"#;

        let parsed = parse_slicer_cache(xml).unwrap();
        assert_eq!(parsed.name, "right");
        assert!(parsed.table_slicer_cache.is_none());
    }

    #[test]
    fn test_parse_slicer_cache_rejects_malformed_input() {
        assert!(parse_slicer_cache("").is_none());
        assert!(parse_slicer_cache("<slicerCacheDefinition sourceName=\"S\"/>").is_none());
        assert!(parse_slicer_cache("<other name=\"N\" sourceName=\"S\"/>").is_none());
        assert!(
            parse_slicer_cache("<slicerCacheDefinition name=\"a&bogus;\" sourceName=\"S\"/>")
                .is_none()
        );
    }
}
