//! Relationships XML schema structures.
//!
//! Used in `_rels/.rels`, `xl/_rels/workbook.xml.rels`, and other relationship files.

use serde::{Deserialize, Serialize};

use crate::namespaces;

/// Relationships root element.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "Relationships")]
pub struct Relationships {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "Relationship", default)]
    pub relationships: Vec<Relationship>,
}

/// Individual relationship entry.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Relationship {
    #[serde(rename = "@Id")]
    pub id: String,

    #[serde(rename = "@Type")]
    pub rel_type: String,

    #[serde(rename = "@Target")]
    pub target: String,

    #[serde(rename = "@TargetMode", skip_serializing_if = "Option::is_none")]
    pub target_mode: Option<String>,
}

/// Creates the package-level relationships (`_rels/.rels`).
///
/// Contains relationships from the package root to the workbook, core
/// properties, and extended properties parts.
pub fn package_rels() -> Relationships {
    Relationships {
        xmlns: namespaces::PACKAGE_RELATIONSHIPS.to_string(),
        relationships: vec![
            Relationship {
                id: "rId1".to_string(),
                rel_type: rel_types::OFFICE_DOCUMENT.to_string(),
                target: "xl/workbook.xml".to_string(),
                target_mode: None,
            },
            Relationship {
                id: "rId2".to_string(),
                rel_type: rel_types::CORE_PROPERTIES.to_string(),
                target: "docProps/core.xml".to_string(),
                target_mode: None,
            },
            Relationship {
                id: "rId3".to_string(),
                rel_type: rel_types::EXTENDED_PROPERTIES.to_string(),
                target: "docProps/app.xml".to_string(),
                target_mode: None,
            },
        ],
    }
}

/// Creates the workbook-level relationships (`xl/_rels/workbook.xml.rels`).
///
/// Contains relationships to worksheets, styles, and shared strings.
pub fn workbook_rels() -> Relationships {
    Relationships {
        xmlns: namespaces::PACKAGE_RELATIONSHIPS.to_string(),
        relationships: vec![
            Relationship {
                id: "rId1".to_string(),
                rel_type: rel_types::WORKSHEET.to_string(),
                target: "worksheets/sheet1.xml".to_string(),
                target_mode: None,
            },
            Relationship {
                id: "rId2".to_string(),
                rel_type: rel_types::STYLES.to_string(),
                target: "styles.xml".to_string(),
                target_mode: None,
            },
            Relationship {
                id: "rId3".to_string(),
                rel_type: rel_types::SHARED_STRINGS.to_string(),
                target: "sharedStrings.xml".to_string(),
                target_mode: None,
            },
        ],
    }
}

/// Relationship type URI constants.
pub mod rel_types {
    // Package level
    pub const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

    // Workbook level
    pub const WORKSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    pub const SHARED_STRINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    pub const CHARTSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    pub const CALC_CHAIN: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
    pub const EXTERNAL_LINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
    pub const PIVOT_CACHE_DEF: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
    pub const PIVOT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
    pub const PIVOT_CACHE_RECORDS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";

    // Worksheet level
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const DRAWING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    pub const TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const PRINTER_SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";

    // Drawing level
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    // Custom properties
    pub const CUSTOM_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_package_rels_factory() {
        let rels = package_rels();
        assert_eq!(rels.xmlns, namespaces::PACKAGE_RELATIONSHIPS);
        assert_eq!(rels.relationships.len(), 3);

        // Office document relationship
        assert_eq!(rels.relationships[0].id, "rId1");
        assert_eq!(rels.relationships[0].rel_type, rel_types::OFFICE_DOCUMENT);
        assert_eq!(rels.relationships[0].target, "xl/workbook.xml");
        assert!(rels.relationships[0].target_mode.is_none());

        // Core properties relationship
        assert_eq!(rels.relationships[1].id, "rId2");
        assert_eq!(rels.relationships[1].rel_type, rel_types::CORE_PROPERTIES);
        assert_eq!(rels.relationships[1].target, "docProps/core.xml");

        // Extended properties relationship
        assert_eq!(rels.relationships[2].id, "rId3");
        assert_eq!(
            rels.relationships[2].rel_type,
            rel_types::EXTENDED_PROPERTIES
        );
        assert_eq!(rels.relationships[2].target, "docProps/app.xml");
    }

    #[test]
    fn test_workbook_rels_factory() {
        let rels = workbook_rels();
        assert_eq!(rels.xmlns, namespaces::PACKAGE_RELATIONSHIPS);
        assert_eq!(rels.relationships.len(), 3);

        // Verify worksheet relationship
        assert_eq!(rels.relationships[0].id, "rId1");
        assert_eq!(rels.relationships[0].rel_type, rel_types::WORKSHEET);
        assert_eq!(rels.relationships[0].target, "worksheets/sheet1.xml");

        // Verify styles relationship
        assert_eq!(rels.relationships[1].id, "rId2");
        assert_eq!(rels.relationships[1].rel_type, rel_types::STYLES);
        assert_eq!(rels.relationships[1].target, "styles.xml");

        // Verify shared strings relationship
        assert_eq!(rels.relationships[2].id, "rId3");
        assert_eq!(rels.relationships[2].rel_type, rel_types::SHARED_STRINGS);
        assert_eq!(rels.relationships[2].target, "sharedStrings.xml");
    }

    #[test]
    fn test_relationships_roundtrip() {
        let rels = package_rels();
        let xml = quick_xml::se::to_string(&rels).unwrap();
        let parsed: Relationships = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(rels.xmlns, parsed.xmlns);
        assert_eq!(rels.relationships.len(), parsed.relationships.len());
        assert_eq!(rels.relationships[0].id, parsed.relationships[0].id);
        assert_eq!(
            rels.relationships[0].rel_type,
            parsed.relationships[0].rel_type
        );
        assert_eq!(rels.relationships[0].target, parsed.relationships[0].target);
    }

    #[test]
    fn test_workbook_rels_roundtrip() {
        let rels = workbook_rels();
        let xml = quick_xml::se::to_string(&rels).unwrap();
        let parsed: Relationships = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(rels.relationships.len(), parsed.relationships.len());
        for (orig, parsed) in rels.relationships.iter().zip(parsed.relationships.iter()) {
            assert_eq!(orig.id, parsed.id);
            assert_eq!(orig.rel_type, parsed.rel_type);
            assert_eq!(orig.target, parsed.target);
        }
    }

    #[test]
    fn test_relationship_with_target_mode() {
        let rel = Relationship {
            id: "rId1".to_string(),
            rel_type: rel_types::HYPERLINK.to_string(),
            target: "https://example.com".to_string(),
            target_mode: Some("External".to_string()),
        };
        let xml = quick_xml::se::to_string(&rel).unwrap();
        assert!(xml.contains("TargetMode=\"External\""));

        let parsed: Relationship = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.target_mode, Some("External".to_string()));
    }

    #[test]
    fn test_relationship_without_target_mode_omits_attr() {
        let rel = Relationship {
            id: "rId1".to_string(),
            rel_type: rel_types::WORKSHEET.to_string(),
            target: "worksheets/sheet1.xml".to_string(),
            target_mode: None,
        };
        let xml = quick_xml::se::to_string(&rel).unwrap();
        assert!(!xml.contains("TargetMode"));
    }

    #[test]
    fn test_parse_real_excel_rels() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

        let parsed: Relationships = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(parsed.xmlns, namespaces::PACKAGE_RELATIONSHIPS);
        assert_eq!(parsed.relationships.len(), 1);
        assert_eq!(parsed.relationships[0].id, "rId1");
        assert_eq!(parsed.relationships[0].target, "xl/workbook.xml");
    }

    #[test]
    fn test_parse_real_excel_workbook_rels() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#;

        let parsed: Relationships = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(parsed.relationships.len(), 3);
        assert_eq!(parsed.relationships[0].rel_type, rel_types::WORKSHEET);
        assert_eq!(parsed.relationships[1].rel_type, rel_types::STYLES);
        assert_eq!(parsed.relationships[2].rel_type, rel_types::SHARED_STRINGS);
    }

    #[test]
    fn test_serialize_structure() {
        let rels = package_rels();
        let xml = quick_xml::se::to_string(&rels).unwrap();
        assert!(xml.contains("<Relationships"));
        assert!(xml.contains("<Relationship"));
        assert!(xml.contains("Id="));
        assert!(xml.contains("Type="));
        assert!(xml.contains("Target="));
    }

    #[test]
    fn test_empty_relationships() {
        let rels = Relationships {
            xmlns: namespaces::PACKAGE_RELATIONSHIPS.to_string(),
            relationships: vec![],
        };
        let xml = quick_xml::se::to_string(&rels).unwrap();
        let parsed: Relationships = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.relationships.is_empty());
    }
}
