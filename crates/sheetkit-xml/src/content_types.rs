//! [Content_Types].xml schema structures.
//!
//! Defines the content types for all parts in the OOXML package.

use serde::{Deserialize, Serialize};

use crate::namespaces;

/// `[Content_Types].xml` root element.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "Types")]
pub struct ContentTypes {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "Default", default)]
    pub defaults: Vec<ContentTypeDefault>,

    #[serde(rename = "Override", default)]
    pub overrides: Vec<ContentTypeOverride>,
}

/// Extension-based default content type mapping.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ContentTypeDefault {
    #[serde(rename = "@Extension")]
    pub extension: String,

    #[serde(rename = "@ContentType")]
    pub content_type: String,
}

/// Path-specific content type override.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ContentTypeOverride {
    #[serde(rename = "@PartName")]
    pub part_name: String,

    #[serde(rename = "@ContentType")]
    pub content_type: String,
}

impl Default for ContentTypes {
    fn default() -> Self {
        Self {
            xmlns: namespaces::CONTENT_TYPES.to_string(),
            defaults: vec![
                ContentTypeDefault {
                    extension: "rels".to_string(),
                    content_type: mime_types::RELS.to_string(),
                },
                ContentTypeDefault {
                    extension: "xml".to_string(),
                    content_type: mime_types::XML.to_string(),
                },
            ],
            overrides: vec![
                ContentTypeOverride {
                    part_name: "/xl/workbook.xml".to_string(),
                    content_type: mime_types::WORKBOOK.to_string(),
                },
                ContentTypeOverride {
                    part_name: "/xl/worksheets/sheet1.xml".to_string(),
                    content_type: mime_types::WORKSHEET.to_string(),
                },
                ContentTypeOverride {
                    part_name: "/xl/styles.xml".to_string(),
                    content_type: mime_types::STYLES.to_string(),
                },
                ContentTypeOverride {
                    part_name: "/xl/sharedStrings.xml".to_string(),
                    content_type: mime_types::SHARED_STRINGS.to_string(),
                },
            ],
        }
    }
}

/// Standard content type MIME string constants.
pub mod mime_types {
    // Default extensions
    pub const RELS: &str = "application/vnd.openxmlformats-package.relationships+xml";
    pub const XML: &str = "application/xml";
    pub const PNG: &str = "image/png";
    pub const JPEG: &str = "image/jpeg";

    // Workbook
    pub const WORKBOOK: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    pub const WORKBOOK_MACRO: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
    pub const WORKBOOK_TEMPLATE: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";

    // Worksheet
    pub const WORKSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    pub const CHARTSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

    // Shared elements
    pub const SHARED_STRINGS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    pub const STYLES: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    pub const THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";

    // Charts and drawings
    pub const CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    pub const DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";

    // Table
    pub const TABLE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

    // Comments
    pub const COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";

    // Document properties
    pub const CORE_PROPERTIES: &str = "application/vnd.openxmlformats-package.core-properties+xml";
    pub const EXTENDED_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    pub const CUSTOM_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.custom-properties+xml";
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_content_types_default() {
        let ct = ContentTypes::default();
        assert_eq!(ct.xmlns, namespaces::CONTENT_TYPES);
        assert_eq!(ct.defaults.len(), 2);
        assert_eq!(ct.overrides.len(), 4);

        // Check default extensions
        assert_eq!(ct.defaults[0].extension, "rels");
        assert_eq!(ct.defaults[1].extension, "xml");

        // Check overrides contain workbook, worksheet, styles, shared strings
        let part_names: Vec<&str> = ct.overrides.iter().map(|o| o.part_name.as_str()).collect();
        assert!(part_names.contains(&"/xl/workbook.xml"));
        assert!(part_names.contains(&"/xl/worksheets/sheet1.xml"));
        assert!(part_names.contains(&"/xl/styles.xml"));
        assert!(part_names.contains(&"/xl/sharedStrings.xml"));
    }

    #[test]
    fn test_content_types_roundtrip() {
        let ct = ContentTypes::default();
        let xml = quick_xml::se::to_string(&ct).unwrap();
        let parsed: ContentTypes = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(ct.defaults.len(), parsed.defaults.len());
        assert_eq!(ct.overrides.len(), parsed.overrides.len());
        assert_eq!(ct.xmlns, parsed.xmlns);
    }

    #[test]
    fn test_content_types_serialize_structure() {
        let ct = ContentTypes::default();
        let xml = quick_xml::se::to_string(&ct).unwrap();

        // Should contain Types as root element
        assert!(xml.contains("<Types"));
        assert!(xml.contains("xmlns="));
        // Should contain Default elements
        assert!(xml.contains("<Default"));
        assert!(xml.contains("Extension="));
        assert!(xml.contains("ContentType="));
        // Should contain Override elements
        assert!(xml.contains("<Override"));
        assert!(xml.contains("PartName="));
    }

    #[test]
    fn test_parse_real_excel_content_types() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#;

        let parsed: ContentTypes = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(parsed.defaults.len(), 2);
        assert_eq!(parsed.overrides.len(), 4);
        assert_eq!(parsed.defaults[0].extension, "rels");
        assert_eq!(parsed.overrides[0].part_name, "/xl/workbook.xml");
    }

    #[test]
    fn test_content_type_default_fields() {
        let default = ContentTypeDefault {
            extension: "png".to_string(),
            content_type: mime_types::PNG.to_string(),
        };
        let xml = quick_xml::se::to_string(&default).unwrap();
        assert!(xml.contains("Extension=\"png\""));
        assert!(xml.contains("ContentType=\"image/png\""));
    }
}
