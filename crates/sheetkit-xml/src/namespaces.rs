//! OOXML namespace definitions.
//! Standard namespaces used across all XML documents.

// Core spreadsheet namespace
pub const SPREADSHEET_ML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

// Relationship namespaces
pub const RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub const PACKAGE_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";

// Content Types
pub const CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

// DrawingML namespaces
pub const DRAWING_ML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub const DRAWING_ML_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
pub const DRAWING_ML_SPREADSHEET: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

// Markup Compatibility
pub const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

// Dublin Core (document properties)
pub const DC: &str = "http://purl.org/dc/elements/1.1/";
pub const DC_TERMS: &str = "http://purl.org/dc/terms/";

// Extended Properties
pub const EXTENDED_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
pub const CORE_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

// Dublin Core DCMI Type
pub const DC_MITYPE: &str = "http://purl.org/dc/dcmitype/";

// VT Types (docProps)
pub const VT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

// Custom Properties
pub const CUSTOM_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

// XML standard
pub const XML: &str = "http://www.w3.org/XML/1998/namespace";
pub const XSI: &str = "http://www.w3.org/2001/XMLSchema-instance";

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_namespace_constants_are_valid_uris() {
        // All namespace constants should be non-empty strings starting with http or urn
        let namespaces = [
            SPREADSHEET_ML,
            RELATIONSHIPS,
            PACKAGE_RELATIONSHIPS,
            CONTENT_TYPES,
            DRAWING_ML,
            MC,
            EXTENDED_PROPERTIES,
            CORE_PROPERTIES,
        ];
        for ns in namespaces {
            assert!(!ns.is_empty());
            assert!(
                ns.starts_with("http://") || ns.starts_with("urn:"),
                "Namespace should start with http:// or urn: but got: {ns}"
            );
        }
    }

    #[test]
    fn test_spreadsheet_ml_namespace() {
        assert_eq!(
            SPREADSHEET_ML,
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        );
    }

    #[test]
    fn test_relationships_namespace() {
        assert_eq!(
            RELATIONSHIPS,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        );
    }

    #[test]
    fn test_content_types_namespace() {
        assert_eq!(
            CONTENT_TYPES,
            "http://schemas.openxmlformats.org/package/2006/content-types"
        );
    }
}
