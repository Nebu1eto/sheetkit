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
use sheetkit_xml::worksheet::{Cell, CellFormula, Row, WorksheetXml};
use zip::write::SimpleFileOptions;
use zip::CompressionMethod;

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::sst::SharedStringTable;
use crate::utils::cell_ref::cell_name_to_coordinates;
use crate::utils::constants::MAX_CELL_CHARS;

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
    #[allow(dead_code)]
    shared_strings: Sst,
    sst_runtime: SharedStringTable,
}

impl Workbook {
    /// Create a new empty workbook containing a single empty sheet named "Sheet1".
    pub fn new() -> Self {
        let shared_strings = Sst::default();
        let sst_runtime = SharedStringTable::from_sst(&shared_strings);
        Self {
            content_types: ContentTypes::default(),
            package_rels: relationships::package_rels(),
            workbook_xml: WorkbookXml::default(),
            workbook_rels: relationships::workbook_rels(),
            worksheets: vec![("Sheet1".to_string(), WorksheetXml::default())],
            stylesheet: StyleSheet::default(),
            shared_strings,
            sst_runtime,
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

        let sst_runtime = SharedStringTable::from_sst(&shared_strings);

        Ok(Self {
            content_types,
            package_rels,
            workbook_xml,
            workbook_rels,
            worksheets,
            stylesheet,
            shared_strings,
            sst_runtime,
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

        // xl/sharedStrings.xml -- write from the runtime SST
        let sst_xml = self.sst_runtime.to_sst();
        write_xml_part(&mut zip, "xl/sharedStrings.xml", &sst_xml, options)?;

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

    // -----------------------------------------------------------------------
    // Cell operations
    // -----------------------------------------------------------------------

    /// Get the value of a cell.
    ///
    /// Returns [`CellValue::Empty`] for cells that have no value or do not
    /// exist in the sheet data.
    pub fn get_cell_value(&self, sheet: &str, cell: &str) -> Result<CellValue> {
        let ws = self
            .worksheets
            .iter()
            .find(|(name, _)| name == sheet)
            .map(|(_, ws)| ws)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })?;

        let (col, row) = cell_name_to_coordinates(cell)?;
        let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(col, row)?;

        // Find the row.
        let xml_row = ws.sheet_data.rows.iter().find(|r| r.r == row);
        let xml_row = match xml_row {
            Some(r) => r,
            None => return Ok(CellValue::Empty),
        };

        // Find the cell.
        let xml_cell = xml_row.cells.iter().find(|c| c.r == cell_ref);
        let xml_cell = match xml_cell {
            Some(c) => c,
            None => return Ok(CellValue::Empty),
        };

        self.xml_cell_to_value(xml_cell)
    }

    /// Set the value of a cell.
    ///
    /// The value can be any type that implements `Into<CellValue>`, including
    /// `&str`, `String`, `f64`, `i32`, `i64`, and `bool`.
    ///
    /// Setting a cell to [`CellValue::Empty`] removes the cell from the row.
    pub fn set_cell_value(
        &mut self,
        sheet: &str,
        cell: &str,
        value: impl Into<CellValue>,
    ) -> Result<()> {
        let value = value.into();

        // Validate string length.
        if let CellValue::String(ref s) = value {
            if s.len() > MAX_CELL_CHARS {
                return Err(Error::CellValueTooLong {
                    length: s.len(),
                    max: MAX_CELL_CHARS,
                });
            }
        }

        let ws = self
            .worksheets
            .iter_mut()
            .find(|(name, _)| name == sheet)
            .map(|(_, ws)| ws)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })?;

        let (col, row_num) = cell_name_to_coordinates(cell)?;
        let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(col, row_num)?;

        // Find or create the row (keep rows sorted by row number).
        let row_idx = match ws.sheet_data.rows.iter().position(|r| r.r >= row_num) {
            Some(idx) if ws.sheet_data.rows[idx].r == row_num => idx,
            Some(idx) => {
                ws.sheet_data.rows.insert(idx, new_row(row_num));
                idx
            }
            None => {
                ws.sheet_data.rows.push(new_row(row_num));
                ws.sheet_data.rows.len() - 1
            }
        };

        let row = &mut ws.sheet_data.rows[row_idx];

        // Handle Empty: remove the cell if present.
        if value == CellValue::Empty {
            row.cells.retain(|c| c.r != cell_ref);
            return Ok(());
        }

        // Find or create the cell.
        let cell_idx = match row.cells.iter().position(|c| c.r == cell_ref) {
            Some(idx) => idx,
            None => {
                // Insert in column order.
                let insert_pos = row
                    .cells
                    .iter()
                    .position(|c| {
                        cell_name_to_coordinates(&c.r)
                            .map(|(c_col, _)| c_col > col)
                            .unwrap_or(false)
                    })
                    .unwrap_or(row.cells.len());
                row.cells.insert(
                    insert_pos,
                    Cell {
                        r: cell_ref,
                        s: None,
                        t: None,
                        v: None,
                        f: None,
                        is: None,
                    },
                );
                insert_pos
            }
        };

        let xml_cell = &mut row.cells[cell_idx];
        value_to_xml_cell(&mut self.sst_runtime, xml_cell, value);

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Sheet management
    // -----------------------------------------------------------------------

    /// Create a new empty sheet with the given name. Returns the 0-based sheet index.
    pub fn new_sheet(&mut self, name: &str) -> Result<usize> {
        crate::sheet::add_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            name,
            WorksheetXml::default(),
        )
    }

    /// Delete a sheet by name.
    pub fn delete_sheet(&mut self, name: &str) -> Result<()> {
        crate::sheet::delete_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            name,
        )
    }

    /// Rename a sheet.
    pub fn set_sheet_name(&mut self, old_name: &str, new_name: &str) -> Result<()> {
        crate::sheet::rename_sheet(
            &mut self.workbook_xml,
            &mut self.worksheets,
            old_name,
            new_name,
        )
    }

    /// Copy a sheet, returning the 0-based index of the new copy.
    pub fn copy_sheet(&mut self, source: &str, target: &str) -> Result<usize> {
        crate::sheet::copy_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            source,
            target,
        )
    }

    /// Get a sheet's 0-based index by name. Returns `None` if not found.
    pub fn get_sheet_index(&self, name: &str) -> Option<usize> {
        crate::sheet::find_sheet_index(&self.worksheets, name)
    }

    /// Get the name of the active sheet.
    pub fn get_active_sheet(&self) -> &str {
        let idx = crate::sheet::active_sheet_index(&self.workbook_xml);
        self.worksheets
            .get(idx)
            .map(|(n, _)| n.as_str())
            .unwrap_or_else(|| self.worksheets[0].0.as_str())
    }

    /// Set the active sheet by name.
    pub fn set_active_sheet(&mut self, name: &str) -> Result<()> {
        let idx = crate::sheet::find_sheet_index(&self.worksheets, name).ok_or_else(|| {
            Error::SheetNotFound {
                name: name.to_string(),
            }
        })?;
        crate::sheet::set_active_sheet_index(&mut self.workbook_xml, idx as u32);
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Row operations
    // -----------------------------------------------------------------------

    /// Insert `count` empty rows starting at `start_row` in the named sheet.
    pub fn insert_rows(&mut self, sheet: &str, start_row: u32, count: u32) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::row::insert_rows(ws, start_row, count)
    }

    /// Remove a single row from the named sheet, shifting rows below it up.
    pub fn remove_row(&mut self, sheet: &str, row: u32) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::row::remove_row(ws, row)
    }

    /// Duplicate a row, inserting the copy directly below.
    pub fn duplicate_row(&mut self, sheet: &str, row: u32) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::row::duplicate_row(ws, row)
    }

    /// Set the height of a row in points.
    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::row::set_row_height(ws, row, height)
    }

    /// Get the height of a row.
    pub fn get_row_height(&self, sheet: &str, row: u32) -> Result<Option<f64>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::row::get_row_height(ws, row))
    }

    /// Set the visibility of a row.
    pub fn set_row_visible(&mut self, sheet: &str, row: u32, visible: bool) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::row::set_row_visible(ws, row, visible)
    }

    // -----------------------------------------------------------------------
    // Column operations
    // -----------------------------------------------------------------------

    /// Set the width of a column.
    pub fn set_col_width(&mut self, sheet: &str, col: &str, width: f64) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::col::set_col_width(ws, col, width)
    }

    /// Get the width of a column.
    pub fn get_col_width(&self, sheet: &str, col: &str) -> Result<Option<f64>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::col::get_col_width(ws, col))
    }

    /// Set the visibility of a column.
    pub fn set_col_visible(&mut self, sheet: &str, col: &str, visible: bool) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::col::set_col_visible(ws, col, visible)
    }

    /// Insert `count` columns starting at `col` in the named sheet.
    pub fn insert_cols(&mut self, sheet: &str, col: &str, count: u32) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::col::insert_cols(ws, col, count)
    }

    /// Remove a single column from the named sheet.
    pub fn remove_col(&mut self, sheet: &str, col: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::col::remove_col(ws, col)
    }

    // -----------------------------------------------------------------------
    // Private helpers
    // -----------------------------------------------------------------------

    /// Get a mutable reference to the worksheet XML for the named sheet.
    fn worksheet_mut(&mut self, sheet: &str) -> Result<&mut WorksheetXml> {
        self.worksheets
            .iter_mut()
            .find(|(name, _)| name == sheet)
            .map(|(_, ws)| ws)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })
    }

    /// Get an immutable reference to the worksheet XML for the named sheet.
    fn worksheet_ref(&self, sheet: &str) -> Result<&WorksheetXml> {
        self.worksheets
            .iter()
            .find(|(name, _)| name == sheet)
            .map(|(_, ws)| ws)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })
    }

    // -----------------------------------------------------------------------
    // Private helpers for cell conversion
    // -----------------------------------------------------------------------

    /// Convert an XML Cell to a CellValue.
    fn xml_cell_to_value(&self, xml_cell: &Cell) -> Result<CellValue> {
        // Check for formula first.
        if let Some(ref formula) = xml_cell.f {
            let expr = formula.value.clone().unwrap_or_default();
            let result = match (&xml_cell.t, &xml_cell.v) {
                (Some(t), Some(v)) if t == "b" => Some(Box::new(CellValue::Bool(v == "1"))),
                (Some(t), Some(v)) if t == "e" => Some(Box::new(CellValue::Error(v.clone()))),
                (_, Some(v)) => v
                    .parse::<f64>()
                    .ok()
                    .map(|n| Box::new(CellValue::Number(n))),
                _ => None,
            };
            return Ok(CellValue::Formula { expr, result });
        }

        let cell_type = xml_cell.t.as_deref();
        let cell_value = xml_cell.v.as_deref();

        match (cell_type, cell_value) {
            // Shared string
            (Some("s"), Some(v)) => {
                let idx: usize = v
                    .parse()
                    .map_err(|_| Error::Internal(format!("invalid SST index: {v}")))?;
                let s = self.sst_runtime.get(idx).unwrap_or("").to_string();
                Ok(CellValue::String(s))
            }
            // Boolean
            (Some("b"), Some(v)) => Ok(CellValue::Bool(v == "1")),
            // Error
            (Some("e"), Some(v)) => Ok(CellValue::Error(v.to_string())),
            // Inline string
            (Some("inlineStr"), _) => {
                let s = xml_cell
                    .is
                    .as_ref()
                    .and_then(|is| is.t.clone())
                    .unwrap_or_default();
                Ok(CellValue::String(s))
            }
            // Formula string (cached string result)
            (Some("str"), Some(v)) => Ok(CellValue::String(v.to_string())),
            // Number (explicit or default type)
            (None | Some("n"), Some(v)) => {
                let n: f64 = v
                    .parse()
                    .map_err(|_| Error::Internal(format!("invalid number: {v}")))?;
                Ok(CellValue::Number(n))
            }
            // No value
            _ => Ok(CellValue::Empty),
        }
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

/// Write a CellValue into an XML Cell (mutating it in place).
fn value_to_xml_cell(sst: &mut SharedStringTable, xml_cell: &mut Cell, value: CellValue) {
    // Clear previous values.
    xml_cell.t = None;
    xml_cell.v = None;
    xml_cell.f = None;
    xml_cell.is = None;

    match value {
        CellValue::String(s) => {
            let idx = sst.add(&s);
            xml_cell.t = Some("s".to_string());
            xml_cell.v = Some(idx.to_string());
        }
        CellValue::Number(n) => {
            xml_cell.v = Some(n.to_string());
        }
        CellValue::Bool(b) => {
            xml_cell.t = Some("b".to_string());
            xml_cell.v = Some(if b { "1" } else { "0" }.to_string());
        }
        CellValue::Formula { expr, .. } => {
            xml_cell.f = Some(CellFormula {
                t: None,
                reference: None,
                si: None,
                value: Some(expr),
            });
        }
        CellValue::Error(e) => {
            xml_cell.t = Some("e".to_string());
            xml_cell.v = Some(e);
        }
        CellValue::Empty => {
            // Already cleared above; the caller should have removed the cell.
        }
    }
}

/// Create a new empty row with the given 1-based row number.
fn new_row(row_num: u32) -> Row {
    Row {
        r: row_num,
        spans: None,
        s: None,
        custom_format: None,
        ht: None,
        hidden: None,
        custom_height: None,
        outline_level: None,
        cells: vec![],
    }
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

    // -----------------------------------------------------------------------
    // Cell operation tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_set_and_get_string_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Hello").unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::String("Hello".to_string()));
    }

    #[test]
    fn test_set_and_get_number_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "B2", 42.5f64).unwrap();
        let val = wb.get_cell_value("Sheet1", "B2").unwrap();
        assert_eq!(val, CellValue::Number(42.5));
    }

    #[test]
    fn test_set_and_get_bool_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "C3", true).unwrap();
        let val = wb.get_cell_value("Sheet1", "C3").unwrap();
        assert_eq!(val, CellValue::Bool(true));

        wb.set_cell_value("Sheet1", "D4", false).unwrap();
        let val = wb.get_cell_value("Sheet1", "D4").unwrap();
        assert_eq!(val, CellValue::Bool(false));
    }

    #[test]
    fn test_set_value_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.set_cell_value("NoSuchSheet", "A1", "test");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_get_value_sheet_not_found() {
        let wb = Workbook::new();
        let result = wb.get_cell_value("NoSuchSheet", "A1");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_get_empty_cell_returns_empty() {
        let wb = Workbook::new();
        let val = wb.get_cell_value("Sheet1", "Z99").unwrap();
        assert_eq!(val, CellValue::Empty);
    }

    #[test]
    fn test_cell_value_roundtrip_save_open() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("cell_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Hello").unwrap();
        wb.set_cell_value("Sheet1", "B1", 42.0f64).unwrap();
        wb.set_cell_value("Sheet1", "C1", true).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Hello".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Sheet1", "B1").unwrap(),
            CellValue::Number(42.0)
        );
        assert_eq!(
            wb2.get_cell_value("Sheet1", "C1").unwrap(),
            CellValue::Bool(true)
        );
    }

    #[test]
    fn test_set_empty_value_clears_cell() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "test").unwrap();
        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("test".to_string())
        );

        wb.set_cell_value("Sheet1", "A1", CellValue::Empty).unwrap();
        assert_eq!(wb.get_cell_value("Sheet1", "A1").unwrap(), CellValue::Empty);
    }

    #[test]
    fn test_string_too_long_returns_error() {
        let mut wb = Workbook::new();
        let long_string = "x".repeat(MAX_CELL_CHARS + 1);
        let result = wb.set_cell_value("Sheet1", "A1", long_string.as_str());
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::CellValueTooLong { .. }
        ));
    }

    #[test]
    fn test_set_multiple_cells_same_row() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "first").unwrap();
        wb.set_cell_value("Sheet1", "B1", "second").unwrap();
        wb.set_cell_value("Sheet1", "C1", "third").unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("first".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "B1").unwrap(),
            CellValue::String("second".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "C1").unwrap(),
            CellValue::String("third".to_string())
        );
    }

    #[test]
    fn test_overwrite_cell_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "original").unwrap();
        wb.set_cell_value("Sheet1", "A1", "updated").unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("updated".to_string())
        );
    }

    #[test]
    fn test_set_and_get_error_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::Error("#DIV/0!".to_string()))
            .unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::Error("#DIV/0!".to_string()));
    }

    #[test]
    fn test_set_and_get_formula_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "A1",
            CellValue::Formula {
                expr: "SUM(B1:B10)".to_string(),
                result: None,
            },
        )
        .unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        match val {
            CellValue::Formula { expr, .. } => {
                assert_eq!(expr, "SUM(B1:B10)");
            }
            other => panic!("expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_set_i32_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", 100i32).unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::Number(100.0));
    }

    #[test]
    fn test_set_string_at_max_length() {
        let mut wb = Workbook::new();
        let max_string = "x".repeat(MAX_CELL_CHARS);
        wb.set_cell_value("Sheet1", "A1", max_string.as_str())
            .unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::String(max_string));
    }

    #[test]
    fn test_set_cells_different_rows() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "row1").unwrap();
        wb.set_cell_value("Sheet1", "A3", "row3").unwrap();
        wb.set_cell_value("Sheet1", "A2", "row2").unwrap(); // inserted between

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("row1".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "A2").unwrap(),
            CellValue::String("row2".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "A3").unwrap(),
            CellValue::String("row3".to_string())
        );
    }

    #[test]
    fn test_string_deduplication_in_sst() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "same").unwrap();
        wb.set_cell_value("Sheet1", "A2", "same").unwrap();
        wb.set_cell_value("Sheet1", "A3", "different").unwrap();

        // Both A1 and A2 should point to the same SST index
        assert_eq!(wb.sst_runtime.len(), 2);
        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("same".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "A2").unwrap(),
            CellValue::String("same".to_string())
        );
    }

    // -----------------------------------------------------------------------
    // Sheet management tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_new_sheet_basic() {
        let mut wb = Workbook::new();
        let idx = wb.new_sheet("Sheet2").unwrap();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_names(), vec!["Sheet1", "Sheet2"]);
    }

    #[test]
    fn test_new_sheet_duplicate_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.new_sheet("Sheet1");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::SheetAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_new_sheet_invalid_name_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.new_sheet("Bad/Name");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::InvalidSheetName(_)));
    }

    #[test]
    fn test_delete_sheet_basic() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.delete_sheet("Sheet1").unwrap();
        assert_eq!(wb.sheet_names(), vec!["Sheet2"]);
    }

    #[test]
    fn test_delete_last_sheet_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.delete_sheet("Sheet1");
        assert!(result.is_err());
    }

    #[test]
    fn test_delete_nonexistent_sheet_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.delete_sheet("NoSuchSheet");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_set_sheet_name_basic() {
        let mut wb = Workbook::new();
        wb.set_sheet_name("Sheet1", "Renamed").unwrap();
        assert_eq!(wb.sheet_names(), vec!["Renamed"]);
    }

    #[test]
    fn test_set_sheet_name_to_existing_returns_error() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        let result = wb.set_sheet_name("Sheet1", "Sheet2");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::SheetAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_copy_sheet_basic() {
        let mut wb = Workbook::new();
        let idx = wb.copy_sheet("Sheet1", "Sheet1 Copy").unwrap();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_names(), vec!["Sheet1", "Sheet1 Copy"]);
    }

    #[test]
    fn test_get_sheet_index() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        assert_eq!(wb.get_sheet_index("Sheet1"), Some(0));
        assert_eq!(wb.get_sheet_index("Sheet2"), Some(1));
        assert_eq!(wb.get_sheet_index("Nonexistent"), None);
    }

    #[test]
    fn test_get_active_sheet_default() {
        let wb = Workbook::new();
        assert_eq!(wb.get_active_sheet(), "Sheet1");
    }

    #[test]
    fn test_set_active_sheet() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.set_active_sheet("Sheet2").unwrap();
        assert_eq!(wb.get_active_sheet(), "Sheet2");
    }

    #[test]
    fn test_set_active_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.set_active_sheet("NoSuchSheet");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_sheet_management_roundtrip_save_open() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("sheet_mgmt.xlsx");

        let mut wb = Workbook::new();
        wb.new_sheet("Data").unwrap();
        wb.new_sheet("Summary").unwrap();
        wb.set_sheet_name("Sheet1", "Overview").unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Overview", "Data", "Summary"]);
    }

    // -----------------------------------------------------------------------
    // Row operation wrapper tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_workbook_insert_rows() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "stay").unwrap();
        wb.set_cell_value("Sheet1", "A2", "shift").unwrap();
        wb.insert_rows("Sheet1", 2, 1).unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("stay".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "A3").unwrap(),
            CellValue::String("shift".to_string())
        );
        assert_eq!(wb.get_cell_value("Sheet1", "A2").unwrap(), CellValue::Empty);
    }

    #[test]
    fn test_workbook_insert_rows_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.insert_rows("NoSheet", 1, 1);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_remove_row() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "first").unwrap();
        wb.set_cell_value("Sheet1", "A2", "second").unwrap();
        wb.set_cell_value("Sheet1", "A3", "third").unwrap();
        wb.remove_row("Sheet1", 2).unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("first".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "A2").unwrap(),
            CellValue::String("third".to_string())
        );
    }

    #[test]
    fn test_workbook_remove_row_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.remove_row("NoSheet", 1);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_duplicate_row() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "original").unwrap();
        wb.duplicate_row("Sheet1", 1).unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("original".to_string())
        );
        // The duplicated row at row 2 has the same SST index.
        assert_eq!(
            wb.get_cell_value("Sheet1", "A2").unwrap(),
            CellValue::String("original".to_string())
        );
    }

    #[test]
    fn test_workbook_set_and_get_row_height() {
        let mut wb = Workbook::new();
        wb.set_row_height("Sheet1", 3, 25.0).unwrap();
        assert_eq!(wb.get_row_height("Sheet1", 3).unwrap(), Some(25.0));
    }

    #[test]
    fn test_workbook_get_row_height_sheet_not_found() {
        let wb = Workbook::new();
        let result = wb.get_row_height("NoSheet", 1);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_set_row_visible() {
        let mut wb = Workbook::new();
        wb.set_row_visible("Sheet1", 1, false).unwrap();
    }

    #[test]
    fn test_workbook_set_row_visible_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.set_row_visible("NoSheet", 1, false);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    // -----------------------------------------------------------------------
    // Column operation wrapper tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_workbook_set_and_get_col_width() {
        let mut wb = Workbook::new();
        wb.set_col_width("Sheet1", "A", 18.0).unwrap();
        assert_eq!(wb.get_col_width("Sheet1", "A").unwrap(), Some(18.0));
    }

    #[test]
    fn test_workbook_get_col_width_sheet_not_found() {
        let wb = Workbook::new();
        let result = wb.get_col_width("NoSheet", "A");
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_set_col_visible() {
        let mut wb = Workbook::new();
        wb.set_col_visible("Sheet1", "B", false).unwrap();
    }

    #[test]
    fn test_workbook_set_col_visible_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.set_col_visible("NoSheet", "A", false);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_insert_cols() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "a").unwrap();
        wb.set_cell_value("Sheet1", "B1", "b").unwrap();
        wb.insert_cols("Sheet1", "B", 1).unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("a".to_string())
        );
        assert_eq!(wb.get_cell_value("Sheet1", "B1").unwrap(), CellValue::Empty);
        assert_eq!(
            wb.get_cell_value("Sheet1", "C1").unwrap(),
            CellValue::String("b".to_string())
        );
    }

    #[test]
    fn test_workbook_insert_cols_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.insert_cols("NoSheet", "A", 1);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_remove_col() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "a").unwrap();
        wb.set_cell_value("Sheet1", "B1", "b").unwrap();
        wb.set_cell_value("Sheet1", "C1", "c").unwrap();
        wb.remove_col("Sheet1", "B").unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("a".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "B1").unwrap(),
            CellValue::String("c".to_string())
        );
    }

    #[test]
    fn test_workbook_remove_col_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.remove_col("NoSheet", "A");
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }
}
