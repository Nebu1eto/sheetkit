//! Workbook file I/O: reading and writing `.xlsx` files.
//!
//! An `.xlsx` file is a ZIP archive containing XML parts. This module provides
//! [`Workbook`] which holds the parsed XML structures in memory and can
//! serialize them back to a valid `.xlsx` file.

use std::io::{Read as _, Write as _};
use std::path::Path;

use serde::Serialize;
use sheetkit_xml::content_types::ContentTypes;
use sheetkit_xml::relationships::{self, rel_types, Relationships};
use sheetkit_xml::shared_strings::Sst;
use sheetkit_xml::styles::StyleSheet;
use sheetkit_xml::workbook::WorkbookXml;
use sheetkit_xml::worksheet::WorksheetXml;
use zip::write::SimpleFileOptions;
use zip::CompressionMethod;

use crate::error::{Error, Result};

/// XML declaration prepended to every XML part in the package.
const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// In-memory representation of an `.xlsx` workbook.
pub struct Workbook {
    content_types: ContentTypes,
    package_rels: Relationships,
    workbook_xml: WorkbookXml,
    workbook_rels: Relationships,
    worksheets: Vec<(String, WorksheetXml)>,
    stylesheet: StyleSheet,
    shared_strings: Sst,
}

impl Workbook {
    /// Create a new empty workbook containing a single empty sheet named "Sheet1".
    pub fn new() -> Self {
        Self {
            content_types: ContentTypes::default(),
            package_rels: relationships::package_rels(),
            workbook_xml: WorkbookXml::default(),
            workbook_rels: relationships::workbook_rels(),
            worksheets: vec![("Sheet1".to_string(), WorksheetXml::default())],
            stylesheet: StyleSheet::default(),
            shared_strings: Sst::default(),
        }
    }

    /// Open an existing `.xlsx` file from disk.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let mut archive = zip::ZipArchive::new(file).map_err(|e| Error::Zip(e.to_string()))?;

        // Parse [Content_Types].xml
        let content_types: ContentTypes = read_xml_part(&mut archive, "[Content_Types].xml")?;

        // Parse _rels/.rels
        let package_rels: Relationships = read_xml_part(&mut archive, "_rels/.rels")?;

        // Parse xl/workbook.xml
        let workbook_xml: WorkbookXml = read_xml_part(&mut archive, "xl/workbook.xml")?;

        // Parse xl/_rels/workbook.xml.rels
        let workbook_rels: Relationships =
            read_xml_part(&mut archive, "xl/_rels/workbook.xml.rels")?;

        // Parse each worksheet referenced in the workbook
        let mut worksheets = Vec::new();
        for sheet_entry in &workbook_xml.sheets.sheets {
            // Find the relationship target for this sheet's rId
            let rel = workbook_rels
                .relationships
                .iter()
                .find(|r| r.id == sheet_entry.r_id && r.rel_type == rel_types::WORKSHEET);

            if let Some(rel) = rel {
                let sheet_path = format!("xl/{}", rel.target);
                let ws: WorksheetXml = read_xml_part(&mut archive, &sheet_path)?;
                worksheets.push((sheet_entry.name.clone(), ws));
            }
        }

        // Parse xl/styles.xml
        let stylesheet: StyleSheet = read_xml_part(&mut archive, "xl/styles.xml")?;

        // Parse xl/sharedStrings.xml (optional -- may not exist for workbooks with no strings)
        let shared_strings: Sst =
            read_xml_part(&mut archive, "xl/sharedStrings.xml").unwrap_or_default();

        Ok(Self {
            content_types,
            package_rels,
            workbook_xml,
            workbook_rels,
            worksheets,
            stylesheet,
            shared_strings,
        })
    }

    /// Save the workbook to a `.xlsx` file at the given path.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let file = std::fs::File::create(path)?;
        let mut zip = zip::ZipWriter::new(file);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

        // [Content_Types].xml
        write_xml_part(
            &mut zip,
            "[Content_Types].xml",
            &self.content_types,
            options,
        )?;

        // _rels/.rels
        write_xml_part(&mut zip, "_rels/.rels", &self.package_rels, options)?;

        // xl/workbook.xml
        write_xml_part(&mut zip, "xl/workbook.xml", &self.workbook_xml, options)?;

        // xl/_rels/workbook.xml.rels
        write_xml_part(
            &mut zip,
            "xl/_rels/workbook.xml.rels",
            &self.workbook_rels,
            options,
        )?;

        // xl/worksheets/sheet{N}.xml
        for (i, (_name, ws)) in self.worksheets.iter().enumerate() {
            let entry_name = format!("xl/worksheets/sheet{}.xml", i + 1);
            write_xml_part(&mut zip, &entry_name, ws, options)?;
        }

        // xl/styles.xml
        write_xml_part(&mut zip, "xl/styles.xml", &self.stylesheet, options)?;

        // xl/sharedStrings.xml
        write_xml_part(
            &mut zip,
            "xl/sharedStrings.xml",
            &self.shared_strings,
            options,
        )?;

        zip.finish().map_err(|e| Error::Zip(e.to_string()))?;
        Ok(())
    }

    /// Return the names of all sheets in workbook order.
    pub fn sheet_names(&self) -> Vec<&str> {
        self.worksheets
            .iter()
            .map(|(name, _)| name.as_str())
            .collect()
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/// Serialize a value to XML with the standard XML declaration prepended.
fn serialize_xml<T: Serialize>(value: &T) -> Result<String> {
    let body = quick_xml::se::to_string(value).map_err(|e| Error::XmlParse(e.to_string()))?;
    Ok(format!("{XML_DECLARATION}\n{body}"))
}

/// Read a ZIP entry and deserialize it from XML.
fn read_xml_part<T: serde::de::DeserializeOwned>(
    archive: &mut zip::ZipArchive<std::fs::File>,
    name: &str,
) -> Result<T> {
    let mut entry = archive
        .by_name(name)
        .map_err(|e| Error::Zip(e.to_string()))?;
    let mut content = String::new();
    entry
        .read_to_string(&mut content)
        .map_err(|e| Error::Zip(e.to_string()))?;
    quick_xml::de::from_str(&content).map_err(|e| Error::XmlDeserialize(e.to_string()))
}

/// Serialize a value to XML and write it as a ZIP entry.
fn write_xml_part<T: Serialize, W: std::io::Write + std::io::Seek>(
    zip: &mut zip::ZipWriter<W>,
    name: &str,
    value: &T,
    options: SimpleFileOptions,
) -> Result<()> {
    let xml = serialize_xml(value)?;
    zip.start_file(name, options)
        .map_err(|e| Error::Zip(e.to_string()))?;
    zip.write_all(xml.as_bytes())?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn test_new_workbook_has_sheet1() {
        let wb = Workbook::new();
        assert_eq!(wb.sheet_names(), vec!["Sheet1"]);
    }

    #[test]
    fn test_new_workbook_save_creates_file() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("test.xlsx");
        let wb = Workbook::new();
        wb.save(&path).unwrap();
        assert!(path.exists());
    }

    #[test]
    fn test_save_and_open_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("roundtrip.xlsx");

        let wb = Workbook::new();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1"]);
    }

    #[test]
    fn test_saved_file_is_valid_zip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("valid.xlsx");
        let wb = Workbook::new();
        wb.save(&path).unwrap();

        // Verify it's a valid ZIP with expected entries
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        let expected_files = [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/worksheets/sheet1.xml",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
        ];

        for name in &expected_files {
            assert!(archive.by_name(name).is_ok(), "Missing ZIP entry: {}", name);
        }
    }

    #[test]
    fn test_open_nonexistent_file_returns_error() {
        let result = Workbook::open("/nonexistent/path.xlsx");
        assert!(result.is_err());
    }

    #[test]
    fn test_saved_xml_has_declarations() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("decl.xlsx");
        let wb = Workbook::new();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        let mut content = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("[Content_Types].xml").unwrap(),
            &mut content,
        )
        .unwrap();
        assert!(content.starts_with("<?xml"));
    }

    #[test]
    fn test_default_trait() {
        let wb = Workbook::default();
        assert_eq!(wb.sheet_names(), vec!["Sheet1"]);
    }

    #[test]
    fn test_serialize_xml_helper() {
        let ct = ContentTypes::default();
        let xml = serialize_xml(&ct).unwrap();
        assert!(xml.starts_with("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"));
        assert!(xml.contains("<Types"));
    }
}
