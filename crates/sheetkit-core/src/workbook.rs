//! Workbook file I/O: reading and writing `.xlsx` files.
//!
//! An `.xlsx` file is a ZIP archive containing XML parts. This module provides
//! [`Workbook`] which holds the parsed XML structures in memory and can
//! serialize them back to a valid `.xlsx` file.

use std::collections::HashMap;
use std::io::{Read as _, Write as _};
use std::path::Path;

use serde::Serialize;
use sheetkit_xml::chart::ChartSpace;
use sheetkit_xml::comments::Comments;
use sheetkit_xml::content_types::{
    mime_types, ContentTypeDefault, ContentTypeOverride, ContentTypes,
};
use sheetkit_xml::drawing::{MarkerType, WsDr};
use sheetkit_xml::relationships::{self, rel_types, Relationship, Relationships};
use sheetkit_xml::shared_strings::Sst;
use sheetkit_xml::styles::StyleSheet;
use sheetkit_xml::workbook::{WorkbookProtection, WorkbookXml};
use sheetkit_xml::worksheet::{Cell, CellFormula, DrawingRef, Row, WorksheetXml};
use zip::write::SimpleFileOptions;
use zip::CompressionMethod;

use crate::cell::CellValue;
use crate::chart::ChartConfig;
use crate::comment::CommentConfig;
use crate::error::{Error, Result};
use crate::image::ImageConfig;
use crate::protection::WorkbookProtectionConfig;
use crate::sst::SharedStringTable;
use crate::utils::cell_ref::cell_name_to_coordinates;
use crate::utils::constants::MAX_CELL_CHARS;
use crate::validation::DataValidationConfig;

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
    /// Per-sheet comments, parallel to the `worksheets` vector.
    sheet_comments: Vec<Option<Comments>>,
    /// Chart parts: (zip path like "xl/charts/chart1.xml", ChartSpace data).
    charts: Vec<(String, ChartSpace)>,
    /// Drawing parts: (zip path like "xl/drawings/drawing1.xml", WsDr data).
    drawings: Vec<(String, WsDr)>,
    /// Image parts: (zip path like "xl/media/image1.png", raw bytes).
    images: Vec<(String, Vec<u8>)>,
    /// Maps sheet index -> drawing index in `drawings`.
    #[allow(dead_code)]
    worksheet_drawings: HashMap<usize, usize>,
    /// Per-sheet worksheet relationship files.
    worksheet_rels: HashMap<usize, Relationships>,
    /// Per-drawing relationship files: drawing_index -> Relationships.
    drawing_rels: HashMap<usize, Relationships>,
    /// Core document properties (docProps/core.xml).
    core_properties: Option<sheetkit_xml::doc_props::CoreProperties>,
    /// Extended/application properties (docProps/app.xml).
    app_properties: Option<sheetkit_xml::doc_props::ExtendedProperties>,
    /// Custom properties (docProps/custom.xml).
    custom_properties: Option<sheetkit_xml::doc_props::CustomProperties>,
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
            sheet_comments: vec![None],
            charts: vec![],
            drawings: vec![],
            images: vec![],
            worksheet_drawings: HashMap::new(),
            worksheet_rels: HashMap::new(),
            drawing_rels: HashMap::new(),
            core_properties: None,
            app_properties: None,
            custom_properties: None,
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

        // Initialize per-sheet comments (one entry per worksheet).
        let sheet_comments: Vec<Option<Comments>> = vec![None; worksheets.len()];

        // Parse docProps/core.xml (optional - uses manual XML parsing)
        let core_properties = read_string_part(&mut archive, "docProps/core.xml")
            .ok()
            .and_then(|xml_str| {
                sheetkit_xml::doc_props::deserialize_core_properties(&xml_str).ok()
            });

        // Parse docProps/app.xml (optional - uses serde)
        let app_properties: Option<sheetkit_xml::doc_props::ExtendedProperties> =
            read_xml_part(&mut archive, "docProps/app.xml").ok();

        // Parse docProps/custom.xml (optional - uses manual XML parsing)
        let custom_properties = read_string_part(&mut archive, "docProps/custom.xml")
            .ok()
            .and_then(|xml_str| {
                sheetkit_xml::doc_props::deserialize_custom_properties(&xml_str).ok()
            });

        Ok(Self {
            content_types,
            package_rels,
            workbook_xml,
            workbook_rels,
            worksheets,
            stylesheet,
            shared_strings,
            sst_runtime,
            sheet_comments,
            charts: vec![],
            drawings: vec![],
            images: vec![],
            worksheet_drawings: HashMap::new(),
            worksheet_rels: HashMap::new(),
            drawing_rels: HashMap::new(),
            core_properties,
            app_properties,
            custom_properties,
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

        // xl/comments{N}.xml -- write per-sheet comments
        for (i, comments) in self.sheet_comments.iter().enumerate() {
            if let Some(ref c) = comments {
                let entry_name = format!("xl/comments{}.xml", i + 1);
                write_xml_part(&mut zip, &entry_name, c, options)?;
            }
        }

        // xl/drawings/drawing{N}.xml -- write drawing parts
        for (path, drawing) in &self.drawings {
            write_xml_part(&mut zip, path, drawing, options)?;
        }

        // xl/charts/chart{N}.xml -- write chart parts
        for (path, chart) in &self.charts {
            write_xml_part(&mut zip, path, chart, options)?;
        }

        // xl/media/image{N}.{ext} -- write image data
        for (path, data) in &self.images {
            zip.start_file(path, options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(data)?;
        }

        // xl/worksheets/_rels/sheet{N}.xml.rels -- write worksheet relationships
        for (sheet_idx, rels) in &self.worksheet_rels {
            let path = format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_idx + 1);
            write_xml_part(&mut zip, &path, rels, options)?;
        }

        // xl/drawings/_rels/drawing{N}.xml.rels -- write drawing relationships
        for (drawing_idx, rels) in &self.drawing_rels {
            let path = format!("xl/drawings/_rels/drawing{}.xml.rels", drawing_idx + 1);
            write_xml_part(&mut zip, &path, rels, options)?;
        }

        // docProps/core.xml
        if let Some(ref props) = self.core_properties {
            let xml_str = sheetkit_xml::doc_props::serialize_core_properties(props);
            zip.start_file("docProps/core.xml", options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(xml_str.as_bytes())?;
        }

        // docProps/app.xml
        if let Some(ref props) = self.app_properties {
            write_xml_part(&mut zip, "docProps/app.xml", props, options)?;
        }

        // docProps/custom.xml
        if let Some(ref props) = self.custom_properties {
            let xml_str = sheetkit_xml::doc_props::serialize_custom_properties(props);
            zip.start_file("docProps/custom.xml", options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(xml_str.as_bytes())?;
        }

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
    // Streaming
    // -----------------------------------------------------------------------

    /// Create a [`StreamWriter`](crate::stream::StreamWriter) for a new sheet.
    ///
    /// The sheet will be added to the workbook when the StreamWriter is applied
    /// via [`apply_stream_writer`](Self::apply_stream_writer).
    pub fn new_stream_writer(&self, sheet_name: &str) -> Result<crate::stream::StreamWriter> {
        crate::sheet::validate_sheet_name(sheet_name)?;
        if self.worksheets.iter().any(|(n, _)| n == sheet_name) {
            return Err(Error::SheetAlreadyExists {
                name: sheet_name.to_string(),
            });
        }
        Ok(crate::stream::StreamWriter::new(sheet_name))
    }

    /// Apply a completed [`StreamWriter`](crate::stream::StreamWriter) to the
    /// workbook, adding it as a new sheet.
    ///
    /// Returns the 0-based index of the new sheet.
    pub fn apply_stream_writer(&mut self, writer: crate::stream::StreamWriter) -> Result<usize> {
        let sheet_name = writer.sheet_name().to_string();
        let (xml_bytes, sst) = writer.into_parts()?;

        // Parse the XML back into WorksheetXml
        let mut ws: WorksheetXml = quick_xml::de::from_str(
            &String::from_utf8(xml_bytes).map_err(|e| Error::Internal(e.to_string()))?,
        )
        .map_err(|e| Error::XmlDeserialize(e.to_string()))?;

        // Merge SST entries and build index mapping (old_index -> new_index)
        let mut sst_remap: Vec<usize> = Vec::with_capacity(sst.len());
        for i in 0..sst.len() {
            if let Some(s) = sst.get(i) {
                let new_idx = self.sst_runtime.add(s);
                sst_remap.push(new_idx);
            }
        }

        // Remap SST indices in the worksheet cells
        for row in &mut ws.sheet_data.rows {
            for cell in &mut row.cells {
                if cell.t.as_deref() == Some("s") {
                    if let Some(ref v) = cell.v {
                        if let Ok(old_idx) = v.parse::<usize>() {
                            if let Some(&new_idx) = sst_remap.get(old_idx) {
                                cell.v = Some(new_idx.to_string());
                            }
                        }
                    }
                }
            }
        }

        // Add the sheet
        crate::sheet::add_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            &sheet_name,
            ws,
        )
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
    // Style operations
    // -----------------------------------------------------------------------

    /// Register a new style and return its ID.
    ///
    /// The style is deduplicated: if an identical style already exists in
    /// the stylesheet, the existing ID is returned.
    pub fn add_style(&mut self, style: &crate::style::Style) -> Result<u32> {
        crate::style::add_style(&mut self.stylesheet, style)
    }

    /// Get the style ID applied to a cell.
    ///
    /// Returns `None` if the cell does not exist or has no explicit style
    /// (i.e. uses the default style 0).
    pub fn get_cell_style(&self, sheet: &str, cell: &str) -> Result<Option<u32>> {
        let ws = self.worksheet_ref(sheet)?;

        let (col, row) = cell_name_to_coordinates(cell)?;
        let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(col, row)?;

        // Find the row.
        let xml_row = match ws.sheet_data.rows.iter().find(|r| r.r == row) {
            Some(r) => r,
            None => return Ok(None),
        };

        // Find the cell.
        let xml_cell = match xml_row.cells.iter().find(|c| c.r == cell_ref) {
            Some(c) => c,
            None => return Ok(None),
        };

        Ok(xml_cell.s)
    }

    /// Set the style ID for a cell.
    ///
    /// If the cell does not exist, an empty cell with just the style is created.
    /// The `style_id` must be a valid index in cellXfs.
    pub fn set_cell_style(&mut self, sheet: &str, cell: &str, style_id: u32) -> Result<()> {
        // Validate the style_id.
        if style_id as usize >= self.stylesheet.cell_xfs.xfs.len() {
            return Err(Error::StyleNotFound { id: style_id });
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

        row.cells[cell_idx].s = Some(style_id);
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Data validation operations
    // -----------------------------------------------------------------------

    /// Add a data validation rule to a sheet.
    pub fn add_data_validation(
        &mut self,
        sheet: &str,
        config: &DataValidationConfig,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::validation::add_validation(ws, config)
    }

    /// Get all data validation rules for a sheet.
    pub fn get_data_validations(&self, sheet: &str) -> Result<Vec<DataValidationConfig>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::validation::get_validations(ws))
    }

    /// Remove a data validation rule matching the given cell range from a sheet.
    pub fn remove_data_validation(&mut self, sheet: &str, sqref: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::validation::remove_validation(ws, sqref)
    }

    // -----------------------------------------------------------------------
    // Comment operations
    // -----------------------------------------------------------------------

    /// Add a comment to a cell on the given sheet.
    pub fn add_comment(&mut self, sheet: &str, config: &CommentConfig) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        crate::comment::add_comment(&mut self.sheet_comments[idx], config);
        Ok(())
    }

    /// Get all comments for a sheet.
    pub fn get_comments(&self, sheet: &str) -> Result<Vec<CommentConfig>> {
        let idx = self.sheet_index(sheet)?;
        Ok(crate::comment::get_all_comments(&self.sheet_comments[idx]))
    }

    /// Remove a comment from a cell on the given sheet.
    pub fn remove_comment(&mut self, sheet: &str, cell: &str) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        crate::comment::remove_comment(&mut self.sheet_comments[idx], cell);
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Auto-filter operations
    // -----------------------------------------------------------------------

    /// Set an auto-filter on a sheet for the given cell range.
    pub fn set_auto_filter(&mut self, sheet: &str, range: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::table::set_auto_filter(ws, range)
    }

    /// Remove the auto-filter from a sheet.
    pub fn remove_auto_filter(&mut self, sheet: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::table::remove_auto_filter(ws);
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Chart operations
    // -----------------------------------------------------------------------

    /// Add a chart to a sheet, anchored between two cells.
    ///
    /// The chart spans from `from_cell` (e.g., `"B2"`) to `to_cell`
    /// (e.g., `"J15"`). The `config` specifies the chart type, series data,
    /// title, and legend visibility.
    pub fn add_chart(
        &mut self,
        sheet: &str,
        from_cell: &str,
        to_cell: &str,
        config: &ChartConfig,
    ) -> Result<()> {
        let sheet_idx =
            crate::sheet::find_sheet_index(&self.worksheets, sheet).ok_or_else(|| {
                Error::SheetNotFound {
                    name: sheet.to_string(),
                }
            })?;

        // Parse cell references to marker coordinates (0-based).
        let (from_col, from_row) = cell_name_to_coordinates(from_cell)?;
        let (to_col, to_row) = cell_name_to_coordinates(to_cell)?;

        let from_marker = MarkerType {
            col: from_col - 1,
            col_off: 0,
            row: from_row - 1,
            row_off: 0,
        };
        let to_marker = MarkerType {
            col: to_col - 1,
            col_off: 0,
            row: to_row - 1,
            row_off: 0,
        };

        // Allocate chart part.
        let chart_num = self.charts.len() + 1;
        let chart_path = format!("xl/charts/chart{}.xml", chart_num);
        let chart_space = crate::chart::build_chart_xml(config);
        self.charts.push((chart_path, chart_space));

        // Get or create drawing for this sheet.
        let drawing_idx = self.ensure_drawing_for_sheet(sheet_idx);

        // Add chart reference to the drawing's relationships.
        let chart_rid = self.next_drawing_rid(drawing_idx);
        let chart_rel_target = format!("../charts/chart{}.xml", chart_num);

        let dr_rels = self
            .drawing_rels
            .entry(drawing_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        dr_rels.relationships.push(Relationship {
            id: chart_rid.clone(),
            rel_type: rel_types::CHART.to_string(),
            target: chart_rel_target,
            target_mode: None,
        });

        // Build the chart anchor and add it to the drawing.
        let drawing = &mut self.drawings[drawing_idx].1;
        let anchor = crate::chart::build_drawing_with_chart(&chart_rid, from_marker, to_marker);
        drawing.two_cell_anchors.extend(anchor.two_cell_anchors);

        // Add content type for the chart.
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/xl/charts/chart{}.xml", chart_num),
            content_type: mime_types::CHART.to_string(),
        });

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Image operations
    // -----------------------------------------------------------------------

    /// Add an image to a sheet from bytes.
    ///
    /// The image is anchored to the cell specified in `config.from_cell`.
    /// Dimensions are specified in pixels via `config.width_px` and
    /// `config.height_px`.
    pub fn add_image(&mut self, sheet: &str, config: &ImageConfig) -> Result<()> {
        crate::image::validate_image_config(config)?;

        let sheet_idx =
            crate::sheet::find_sheet_index(&self.worksheets, sheet).ok_or_else(|| {
                Error::SheetNotFound {
                    name: sheet.to_string(),
                }
            })?;

        // Allocate image media part.
        let image_num = self.images.len() + 1;
        let image_path = format!("xl/media/image{}.{}", image_num, config.format.extension());
        self.images.push((image_path, config.data.clone()));

        // Ensure the image extension has a default content type.
        let ext = config.format.extension().to_string();
        if !self
            .content_types
            .defaults
            .iter()
            .any(|d| d.extension == ext)
        {
            self.content_types.defaults.push(ContentTypeDefault {
                extension: ext,
                content_type: config.format.content_type().to_string(),
            });
        }

        // Get or create drawing for this sheet.
        let drawing_idx = self.ensure_drawing_for_sheet(sheet_idx);

        // Add image reference to the drawing's relationships.
        let image_rid = self.next_drawing_rid(drawing_idx);
        let image_rel_target = format!("../media/image{}.{}", image_num, config.format.extension());

        let dr_rels = self
            .drawing_rels
            .entry(drawing_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        dr_rels.relationships.push(Relationship {
            id: image_rid.clone(),
            rel_type: rel_types::IMAGE.to_string(),
            target: image_rel_target,
            target_mode: None,
        });

        // Count existing objects in the drawing to assign a unique ID.
        let drawing = &mut self.drawings[drawing_idx].1;
        let pic_id = (drawing.one_cell_anchors.len() + drawing.two_cell_anchors.len() + 2) as u32;

        // Add image anchor to the drawing.
        crate::image::add_image_to_drawing(drawing, &image_rid, config, pic_id)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Workbook Protection
    // -----------------------------------------------------------------------

    /// Protect the workbook structure and/or windows.
    pub fn protect_workbook(&mut self, config: WorkbookProtectionConfig) {
        let password_hash = config.password.as_ref().map(|p| {
            let hash = crate::protection::legacy_password_hash(p);
            format!("{:04X}", hash)
        });
        self.workbook_xml.workbook_protection = Some(WorkbookProtection {
            workbook_password: password_hash,
            lock_structure: if config.lock_structure {
                Some(true)
            } else {
                None
            },
            lock_windows: if config.lock_windows {
                Some(true)
            } else {
                None
            },
            revisions_password: None,
            lock_revision: if config.lock_revision {
                Some(true)
            } else {
                None
            },
        });
    }

    /// Remove workbook protection.
    pub fn unprotect_workbook(&mut self) {
        self.workbook_xml.workbook_protection = None;
    }

    /// Check if the workbook is protected.
    pub fn is_workbook_protected(&self) -> bool {
        self.workbook_xml.workbook_protection.is_some()
    }

    // -----------------------------------------------------------------------
    // Private helpers
    // -----------------------------------------------------------------------

    /// Ensure a drawing exists for the given sheet index, creating one if needed.
    /// Returns the drawing index.
    fn ensure_drawing_for_sheet(&mut self, sheet_idx: usize) -> usize {
        if let Some(&idx) = self.worksheet_drawings.get(&sheet_idx) {
            return idx;
        }

        let idx = self.drawings.len();
        let drawing_path = format!("xl/drawings/drawing{}.xml", idx + 1);
        self.drawings.push((drawing_path, WsDr::default()));
        self.worksheet_drawings.insert(sheet_idx, idx);

        // Add drawing reference to the worksheet.
        let ws_rid = self.next_worksheet_rid(sheet_idx);
        self.worksheets[sheet_idx].1.drawing = Some(DrawingRef {
            r_id: ws_rid.clone(),
        });

        // Add worksheet->drawing relationship.
        let drawing_rel_target = format!("../drawings/drawing{}.xml", idx + 1);
        let ws_rels = self
            .worksheet_rels
            .entry(sheet_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        ws_rels.relationships.push(Relationship {
            id: ws_rid,
            rel_type: rel_types::DRAWING.to_string(),
            target: drawing_rel_target,
            target_mode: None,
        });

        // Add content type for the drawing.
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/xl/drawings/drawing{}.xml", idx + 1),
            content_type: mime_types::DRAWING.to_string(),
        });

        idx
    }

    /// Generate the next relationship ID for a worksheet's rels.
    fn next_worksheet_rid(&self, sheet_idx: usize) -> String {
        let existing = self
            .worksheet_rels
            .get(&sheet_idx)
            .map(|r| r.relationships.as_slice())
            .unwrap_or(&[]);
        crate::sheet::next_rid(existing)
    }

    /// Generate the next relationship ID for a drawing's rels.
    fn next_drawing_rid(&self, drawing_idx: usize) -> String {
        let existing = self
            .drawing_rels
            .get(&drawing_idx)
            .map(|r| r.relationships.as_slice())
            .unwrap_or(&[]);
        crate::sheet::next_rid(existing)
    }

    /// Get the 0-based index of a sheet by name.
    fn sheet_index(&self, sheet: &str) -> Result<usize> {
        self.worksheets
            .iter()
            .position(|(name, _)| name == sheet)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })
    }

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

    // -----------------------------------------------------------------------
    // Document properties
    // -----------------------------------------------------------------------

    /// Set the core document properties (title, author, etc.).
    pub fn set_doc_props(&mut self, props: crate::doc_props::DocProperties) {
        self.core_properties = Some(props.to_core_properties());
        // Ensure content types includes core properties
        self.ensure_doc_props_content_types();
    }

    /// Get the core document properties.
    pub fn get_doc_props(&self) -> crate::doc_props::DocProperties {
        self.core_properties
            .as_ref()
            .map(crate::doc_props::DocProperties::from)
            .unwrap_or_default()
    }

    /// Set the application properties (company, app version, etc.).
    pub fn set_app_props(&mut self, props: crate::doc_props::AppProperties) {
        self.app_properties = Some(props.to_extended_properties());
        // Ensure content types includes extended properties
        self.ensure_doc_props_content_types();
    }

    /// Get the application properties.
    pub fn get_app_props(&self) -> crate::doc_props::AppProperties {
        self.app_properties
            .as_ref()
            .map(crate::doc_props::AppProperties::from)
            .unwrap_or_default()
    }

    /// Set a custom property by name. If a property with the same name already
    /// exists, its value is replaced.
    pub fn set_custom_property(
        &mut self,
        name: &str,
        value: crate::doc_props::CustomPropertyValue,
    ) {
        let props = self
            .custom_properties
            .get_or_insert_with(sheetkit_xml::doc_props::CustomProperties::default);
        crate::doc_props::set_custom_property(props, name, value);
        self.ensure_custom_props_content_types();
    }

    /// Get a custom property value by name, or `None` if it does not exist.
    pub fn get_custom_property(&self, name: &str) -> Option<crate::doc_props::CustomPropertyValue> {
        self.custom_properties
            .as_ref()
            .and_then(|p| crate::doc_props::find_custom_property(p, name))
    }

    /// Remove a custom property by name. Returns `true` if a property was
    /// found and removed.
    pub fn delete_custom_property(&mut self, name: &str) -> bool {
        if let Some(ref mut props) = self.custom_properties {
            crate::doc_props::delete_custom_property(props, name)
        } else {
            false
        }
    }

    /// Ensure content types contains entries for core and extended properties.
    fn ensure_doc_props_content_types(&mut self) {
        let core_part = "/docProps/core.xml";
        let app_part = "/docProps/app.xml";

        let has_core = self
            .content_types
            .overrides
            .iter()
            .any(|o| o.part_name == core_part);
        if !has_core {
            self.content_types.overrides.push(ContentTypeOverride {
                part_name: core_part.to_string(),
                content_type: mime_types::CORE_PROPERTIES.to_string(),
            });
        }

        let has_app = self
            .content_types
            .overrides
            .iter()
            .any(|o| o.part_name == app_part);
        if !has_app {
            self.content_types.overrides.push(ContentTypeOverride {
                part_name: app_part.to_string(),
                content_type: mime_types::EXTENDED_PROPERTIES.to_string(),
            });
        }
    }

    /// Ensure content types and package rels contain entries for custom properties.
    fn ensure_custom_props_content_types(&mut self) {
        // Ensure doc props content types exist first
        self.ensure_doc_props_content_types();

        let custom_part = "/docProps/custom.xml";
        let has_custom = self
            .content_types
            .overrides
            .iter()
            .any(|o| o.part_name == custom_part);
        if !has_custom {
            self.content_types.overrides.push(ContentTypeOverride {
                part_name: custom_part.to_string(),
                content_type: mime_types::CUSTOM_PROPERTIES.to_string(),
            });
        }

        // Ensure package rels contains custom properties relationship
        let has_custom_rel = self
            .package_rels
            .relationships
            .iter()
            .any(|r| r.rel_type == rel_types::CUSTOM_PROPERTIES);
        if !has_custom_rel {
            let next_id = self.package_rels.relationships.len() + 1;
            self.package_rels.relationships.push(Relationship {
                id: format!("rId{next_id}"),
                rel_type: rel_types::CUSTOM_PROPERTIES.to_string(),
                target: "docProps/custom.xml".to_string(),
                target_mode: None,
            });
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

/// Read a ZIP entry as a raw string (no serde deserialization).
fn read_string_part(archive: &mut zip::ZipArchive<std::fs::File>, name: &str) -> Result<String> {
    let mut entry = archive
        .by_name(name)
        .map_err(|e| Error::Zip(e.to_string()))?;
    let mut content = String::new();
    entry
        .read_to_string(&mut content)
        .map_err(|e| Error::Zip(e.to_string()))?;
    Ok(content)
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
    // Style operation tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_add_style_returns_id() {
        use crate::style::{FontStyle, Style};

        let mut wb = Workbook::new();
        let style = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };
        let id = wb.add_style(&style).unwrap();
        assert!(id > 0);
    }

    #[test]
    fn test_get_cell_style_unstyled_cell_returns_none() {
        let wb = Workbook::new();
        let result = wb.get_cell_style("Sheet1", "A1").unwrap();
        assert!(result.is_none());
    }

    #[test]
    fn test_set_cell_style_on_existing_value() {
        use crate::style::{FontStyle, Style};

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Hello").unwrap();

        let style = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };
        let style_id = wb.add_style(&style).unwrap();
        wb.set_cell_style("Sheet1", "A1", style_id).unwrap();

        let retrieved_id = wb.get_cell_style("Sheet1", "A1").unwrap();
        assert_eq!(retrieved_id, Some(style_id));

        // The value should still be there.
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::String("Hello".to_string()));
    }

    #[test]
    fn test_set_cell_style_on_empty_cell_creates_cell() {
        use crate::style::{FontStyle, Style};

        let mut wb = Workbook::new();
        let style = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };
        let style_id = wb.add_style(&style).unwrap();

        // Set style on a cell that doesn't exist yet.
        wb.set_cell_style("Sheet1", "B5", style_id).unwrap();

        let retrieved_id = wb.get_cell_style("Sheet1", "B5").unwrap();
        assert_eq!(retrieved_id, Some(style_id));

        // The cell value should be empty.
        let val = wb.get_cell_value("Sheet1", "B5").unwrap();
        assert_eq!(val, CellValue::Empty);
    }

    #[test]
    fn test_set_cell_style_invalid_id() {
        let mut wb = Workbook::new();
        let result = wb.set_cell_style("Sheet1", "A1", 999);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StyleNotFound { .. }));
    }

    #[test]
    fn test_set_cell_style_sheet_not_found() {
        let mut wb = Workbook::new();
        let style = crate::style::Style::default();
        let style_id = wb.add_style(&style).unwrap();
        let result = wb.set_cell_style("NoSuchSheet", "A1", style_id);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_get_cell_style_sheet_not_found() {
        let wb = Workbook::new();
        let result = wb.get_cell_style("NoSuchSheet", "A1");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_style_roundtrip_save_open() {
        use crate::style::{
            AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle, FillStyle, FontStyle,
            HorizontalAlign, NumFmtStyle, PatternType, Style, StyleColor, VerticalAlign,
        };

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("style_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Styled").unwrap();

        let style = Style {
            font: Some(FontStyle {
                name: Some("Arial".to_string()),
                size: Some(14.0),
                bold: true,
                italic: true,
                color: Some(StyleColor::Rgb("FFFF0000".to_string())),
                ..FontStyle::default()
            }),
            fill: Some(FillStyle {
                pattern: PatternType::Solid,
                fg_color: Some(StyleColor::Rgb("FFFFFF00".to_string())),
                bg_color: None,
            }),
            border: Some(BorderStyle {
                left: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                right: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                top: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                bottom: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                diagonal: None,
            }),
            alignment: Some(AlignmentStyle {
                horizontal: Some(HorizontalAlign::Center),
                vertical: Some(VerticalAlign::Center),
                wrap_text: true,
                ..AlignmentStyle::default()
            }),
            num_fmt: Some(NumFmtStyle::Custom("#,##0.00".to_string())),
            protection: None,
        };
        let style_id = wb.add_style(&style).unwrap();
        wb.set_cell_style("Sheet1", "A1", style_id).unwrap();
        wb.save(&path).unwrap();

        // Re-open and verify.
        let wb2 = Workbook::open(&path).unwrap();
        let retrieved_id = wb2.get_cell_style("Sheet1", "A1").unwrap();
        assert_eq!(retrieved_id, Some(style_id));

        // Verify the value is still there.
        let val = wb2.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::String("Styled".to_string()));

        // Reverse-lookup the style to verify components survived the roundtrip.
        let retrieved_style = crate::style::get_style(&wb2.stylesheet, style_id).unwrap();
        assert!(retrieved_style.font.is_some());
        let font = retrieved_style.font.unwrap();
        assert!(font.bold);
        assert!(font.italic);
        assert_eq!(font.name, Some("Arial".to_string()));

        assert!(retrieved_style.fill.is_some());
        let fill = retrieved_style.fill.unwrap();
        assert_eq!(fill.pattern, PatternType::Solid);

        assert!(retrieved_style.alignment.is_some());
        let align = retrieved_style.alignment.unwrap();
        assert_eq!(align.horizontal, Some(HorizontalAlign::Center));
        assert_eq!(align.vertical, Some(VerticalAlign::Center));
        assert!(align.wrap_text);
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

    // -----------------------------------------------------------------------
    // StreamWriter integration tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_new_stream_writer_validates_name() {
        let wb = Workbook::new();
        let result = wb.new_stream_writer("Bad[Name");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::InvalidSheetName(_)));
    }

    #[test]
    fn test_new_stream_writer_rejects_duplicate() {
        let wb = Workbook::new();
        let result = wb.new_stream_writer("Sheet1");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::SheetAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_new_stream_writer_valid_name() {
        let wb = Workbook::new();
        let sw = wb.new_stream_writer("StreamSheet").unwrap();
        assert_eq!(sw.sheet_name(), "StreamSheet");
    }

    #[test]
    fn test_apply_stream_writer_adds_sheet() {
        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("StreamSheet").unwrap();
        sw.write_row(1, &[CellValue::from("Hello"), CellValue::from(42)])
            .unwrap();
        let idx = wb.apply_stream_writer(sw).unwrap();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_names(), vec!["Sheet1", "StreamSheet"]);
    }

    #[test]
    fn test_apply_stream_writer_merges_sst() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Existing").unwrap();

        let mut sw = wb.new_stream_writer("StreamSheet").unwrap();
        sw.write_row(1, &[CellValue::from("New"), CellValue::from("Existing")])
            .unwrap();
        wb.apply_stream_writer(sw).unwrap();

        assert!(wb.sst_runtime.len() >= 2);
    }

    #[test]
    fn test_stream_writer_save_and_reopen() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_test.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Normal").unwrap();

        let mut sw = wb.new_stream_writer("Streamed").unwrap();
        sw.write_row(1, &[CellValue::from("Name"), CellValue::from("Value")])
            .unwrap();
        sw.write_row(2, &[CellValue::from("Alice"), CellValue::from(100)])
            .unwrap();
        sw.write_row(3, &[CellValue::from("Bob"), CellValue::from(200)])
            .unwrap();
        wb.apply_stream_writer(sw).unwrap();

        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1", "Streamed"]);
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Normal".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Streamed", "A1").unwrap(),
            CellValue::String("Name".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Streamed", "B2").unwrap(),
            CellValue::Number(100.0)
        );
        assert_eq!(
            wb2.get_cell_value("Streamed", "A3").unwrap(),
            CellValue::String("Bob".to_string())
        );
    }

    // -----------------------------------------------------------------------
    // Data validation tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_workbook_add_data_validation() {
        let mut wb = Workbook::new();
        let config =
            crate::validation::DataValidationConfig::dropdown("A1:A100", &["Yes", "No", "Maybe"]);
        wb.add_data_validation("Sheet1", &config).unwrap();

        let validations = wb.get_data_validations("Sheet1").unwrap();
        assert_eq!(validations.len(), 1);
        assert_eq!(validations[0].sqref, "A1:A100");
    }

    #[test]
    fn test_workbook_remove_data_validation() {
        let mut wb = Workbook::new();
        let config1 = crate::validation::DataValidationConfig::dropdown("A1:A100", &["Yes", "No"]);
        let config2 = crate::validation::DataValidationConfig::whole_number("B1:B100", 1, 100);
        wb.add_data_validation("Sheet1", &config1).unwrap();
        wb.add_data_validation("Sheet1", &config2).unwrap();

        wb.remove_data_validation("Sheet1", "A1:A100").unwrap();

        let validations = wb.get_data_validations("Sheet1").unwrap();
        assert_eq!(validations.len(), 1);
        assert_eq!(validations[0].sqref, "B1:B100");
    }

    #[test]
    fn test_workbook_data_validation_sheet_not_found() {
        let mut wb = Workbook::new();
        let config = crate::validation::DataValidationConfig::dropdown("A1:A100", &["Yes", "No"]);
        let result = wb.add_data_validation("NoSheet", &config);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_data_validation_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("validation_roundtrip.xlsx");

        let mut wb = Workbook::new();
        let config =
            crate::validation::DataValidationConfig::dropdown("A1:A50", &["Red", "Blue", "Green"]);
        wb.add_data_validation("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let validations = wb2.get_data_validations("Sheet1").unwrap();
        assert_eq!(validations.len(), 1);
        assert_eq!(validations[0].sqref, "A1:A50");
        assert_eq!(
            validations[0].validation_type,
            crate::validation::ValidationType::List
        );
    }

    // -----------------------------------------------------------------------
    // Comment tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_workbook_add_comment() {
        let mut wb = Workbook::new();
        let config = crate::comment::CommentConfig {
            cell: "A1".to_string(),
            author: "Alice".to_string(),
            text: "Test comment".to_string(),
        };
        wb.add_comment("Sheet1", &config).unwrap();

        let comments = wb.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell, "A1");
        assert_eq!(comments[0].author, "Alice");
        assert_eq!(comments[0].text, "Test comment");
    }

    #[test]
    fn test_workbook_remove_comment() {
        let mut wb = Workbook::new();
        let config = crate::comment::CommentConfig {
            cell: "A1".to_string(),
            author: "Alice".to_string(),
            text: "Test comment".to_string(),
        };
        wb.add_comment("Sheet1", &config).unwrap();
        wb.remove_comment("Sheet1", "A1").unwrap();

        let comments = wb.get_comments("Sheet1").unwrap();
        assert!(comments.is_empty());
    }

    #[test]
    fn test_workbook_multiple_comments() {
        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Alice".to_string(),
                text: "First".to_string(),
            },
        )
        .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "B2".to_string(),
                author: "Bob".to_string(),
                text: "Second".to_string(),
            },
        )
        .unwrap();

        let comments = wb.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 2);
    }

    #[test]
    fn test_workbook_comment_sheet_not_found() {
        let mut wb = Workbook::new();
        let config = crate::comment::CommentConfig {
            cell: "A1".to_string(),
            author: "Alice".to_string(),
            text: "Test".to_string(),
        };
        let result = wb.add_comment("NoSheet", &config);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_comment_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("comment_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "A saved comment".to_string(),
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        // Verify the comments XML was written to the ZIP.
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(
            archive.by_name("xl/comments1.xml").is_ok(),
            "comments1.xml should be present in the ZIP"
        );
    }

    // -----------------------------------------------------------------------
    // Auto-filter tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_workbook_set_auto_filter() {
        let mut wb = Workbook::new();
        wb.set_auto_filter("Sheet1", "A1:D10").unwrap();

        let ws = wb.worksheet_ref("Sheet1").unwrap();
        assert!(ws.auto_filter.is_some());
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:D10");
    }

    #[test]
    fn test_workbook_remove_auto_filter() {
        let mut wb = Workbook::new();
        wb.set_auto_filter("Sheet1", "A1:D10").unwrap();
        wb.remove_auto_filter("Sheet1").unwrap();

        let ws = wb.worksheet_ref("Sheet1").unwrap();
        assert!(ws.auto_filter.is_none());
    }

    #[test]
    fn test_workbook_auto_filter_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.set_auto_filter("NoSheet", "A1:D10");
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_auto_filter_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("autofilter_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.set_auto_filter("Sheet1", "A1:C50").unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let ws = wb2.worksheet_ref("Sheet1").unwrap();
        assert!(ws.auto_filter.is_some());
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:C50");
    }

    // -----------------------------------------------------------------------
    // Chart tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_add_chart_basic() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: Some("Test Chart".to_string()),
            series: vec![ChartSeries {
                name: "Sales".to_string(),
                categories: "Sheet1!$A$1:$A$5".to_string(),
                values: "Sheet1!$B$1:$B$5".to_string(),
            }],
            show_legend: true,
        };
        wb.add_chart("Sheet1", "E1", "L15", &config).unwrap();

        assert_eq!(wb.charts.len(), 1);
        assert_eq!(wb.drawings.len(), 1);
        assert!(wb.worksheet_drawings.contains_key(&0));
        assert!(wb.drawing_rels.contains_key(&0));
        assert!(wb.worksheets[0].1.drawing.is_some());
    }

    #[test]
    fn test_add_chart_sheet_not_found() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Line,
            title: None,
            series: vec![ChartSeries {
                name: String::new(),
                categories: "Sheet1!$A$1:$A$5".to_string(),
                values: "Sheet1!$B$1:$B$5".to_string(),
            }],
            show_legend: false,
        };
        let result = wb.add_chart("NoSheet", "A1", "H10", &config);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_add_multiple_charts_same_sheet() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        let config1 = ChartConfig {
            chart_type: ChartType::Col,
            title: Some("Chart 1".to_string()),
            series: vec![ChartSeries {
                name: "S1".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
            }],
            show_legend: true,
        };
        let config2 = ChartConfig {
            chart_type: ChartType::Line,
            title: Some("Chart 2".to_string()),
            series: vec![ChartSeries {
                name: "S2".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$C$1:$C$3".to_string(),
            }],
            show_legend: false,
        };
        wb.add_chart("Sheet1", "A1", "F10", &config1).unwrap();
        wb.add_chart("Sheet1", "A12", "F22", &config2).unwrap();

        assert_eq!(wb.charts.len(), 2);
        assert_eq!(wb.drawings.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 2);
    }

    #[test]
    fn test_add_charts_different_sheets() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();

        let config = ChartConfig {
            chart_type: ChartType::Pie,
            title: None,
            series: vec![ChartSeries {
                name: String::new(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
            }],
            show_legend: true,
        };
        wb.add_chart("Sheet1", "A1", "F10", &config).unwrap();
        wb.add_chart("Sheet2", "A1", "F10", &config).unwrap();

        assert_eq!(wb.charts.len(), 2);
        assert_eq!(wb.drawings.len(), 2);
    }

    #[test]
    fn test_save_with_chart() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("with_chart.xlsx");

        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Bar,
            title: Some("Bar Chart".to_string()),
            series: vec![ChartSeries {
                name: "Data".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
            }],
            show_legend: true,
        };
        wb.add_chart("Sheet1", "E2", "L15", &config).unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        assert!(archive.by_name("xl/charts/chart1.xml").is_ok());
        assert!(archive.by_name("xl/drawings/drawing1.xml").is_ok());
        assert!(archive
            .by_name("xl/worksheets/_rels/sheet1.xml.rels")
            .is_ok());
        assert!(archive
            .by_name("xl/drawings/_rels/drawing1.xml.rels")
            .is_ok());
    }

    // -----------------------------------------------------------------------
    // Image tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_add_image_basic() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x89, 0x50, 0x4E, 0x47],
            format: ImageFormat::Png,
            from_cell: "B2".to_string(),
            width_px: 400,
            height_px: 300,
        };
        wb.add_image("Sheet1", &config).unwrap();

        assert_eq!(wb.images.len(), 1);
        assert_eq!(wb.drawings.len(), 1);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 1);
        assert!(wb.worksheet_drawings.contains_key(&0));
    }

    #[test]
    fn test_add_image_sheet_not_found() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x89],
            format: ImageFormat::Png,
            from_cell: "A1".to_string(),
            width_px: 100,
            height_px: 100,
        };
        let result = wb.add_image("NoSheet", &config);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_add_image_invalid_config() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![],
            format: ImageFormat::Png,
            from_cell: "A1".to_string(),
            width_px: 100,
            height_px: 100,
        };
        assert!(wb.add_image("Sheet1", &config).is_err());

        let config = ImageConfig {
            data: vec![1],
            format: ImageFormat::Jpeg,
            from_cell: "A1".to_string(),
            width_px: 0,
            height_px: 100,
        };
        assert!(wb.add_image("Sheet1", &config).is_err());
    }

    #[test]
    fn test_save_with_image() {
        use crate::image::{ImageConfig, ImageFormat};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("with_image.xlsx");

        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A],
            format: ImageFormat::Png,
            from_cell: "C3".to_string(),
            width_px: 200,
            height_px: 150,
        };
        wb.add_image("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        assert!(archive.by_name("xl/media/image1.png").is_ok());
        assert!(archive.by_name("xl/drawings/drawing1.xml").is_ok());
        assert!(archive
            .by_name("xl/worksheets/_rels/sheet1.xml.rels")
            .is_ok());
        assert!(archive
            .by_name("xl/drawings/_rels/drawing1.xml.rels")
            .is_ok());
    }

    #[test]
    fn test_save_with_jpeg_image() {
        use crate::image::{ImageConfig, ImageFormat};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("with_jpeg.xlsx");

        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0xFF, 0xD8, 0xFF, 0xE0],
            format: ImageFormat::Jpeg,
            from_cell: "A1".to_string(),
            width_px: 640,
            height_px: 480,
        };
        wb.add_image("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/media/image1.jpeg").is_ok());
    }

    #[test]
    fn test_add_chart_and_image_same_sheet() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        use crate::image::{ImageConfig, ImageFormat};

        let mut wb = Workbook::new();

        let chart_config = ChartConfig {
            chart_type: ChartType::Col,
            title: Some("My Chart".to_string()),
            series: vec![ChartSeries {
                name: "Series 1".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
            }],
            show_legend: true,
        };
        wb.add_chart("Sheet1", "E1", "L10", &chart_config).unwrap();

        let image_config = ImageConfig {
            data: vec![0x89, 0x50, 0x4E, 0x47],
            format: ImageFormat::Png,
            from_cell: "E12".to_string(),
            width_px: 300,
            height_px: 200,
        };
        wb.add_image("Sheet1", &image_config).unwrap();

        assert_eq!(wb.drawings.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 1);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 1);
        assert_eq!(wb.charts.len(), 1);
        assert_eq!(wb.images.len(), 1);
    }

    #[test]
    fn test_save_with_chart_roundtrip_drawing_ref() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("chart_drawref.xlsx");

        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: None,
            series: vec![ChartSeries {
                name: "Series 1".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
            }],
            show_legend: false,
        };
        wb.add_chart("Sheet1", "A1", "F10", &config).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let ws = wb2.worksheet_ref("Sheet1").unwrap();
        assert!(ws.drawing.is_some());
    }

    #[test]
    fn test_protect_unprotect_workbook() {
        let mut wb = Workbook::new();
        assert!(!wb.is_workbook_protected());

        wb.protect_workbook(crate::protection::WorkbookProtectionConfig {
            password: None,
            lock_structure: true,
            lock_windows: false,
            lock_revision: false,
        });
        assert!(wb.is_workbook_protected());

        wb.unprotect_workbook();
        assert!(!wb.is_workbook_protected());
    }

    #[test]
    fn test_protect_workbook_with_password() {
        let mut wb = Workbook::new();
        wb.protect_workbook(crate::protection::WorkbookProtectionConfig {
            password: Some("secret".to_string()),
            lock_structure: true,
            lock_windows: false,
            lock_revision: false,
        });

        let prot = wb.workbook_xml.workbook_protection.as_ref().unwrap();
        assert!(prot.workbook_password.is_some());
        let hash_str = prot.workbook_password.as_ref().unwrap();
        // Should be a 4-character uppercase hex string
        assert_eq!(hash_str.len(), 4);
        assert!(hash_str.chars().all(|c| c.is_ascii_hexdigit()));
        assert_eq!(prot.lock_structure, Some(true));
    }

    #[test]
    fn test_protect_workbook_structure_only() {
        let mut wb = Workbook::new();
        wb.protect_workbook(crate::protection::WorkbookProtectionConfig {
            password: None,
            lock_structure: true,
            lock_windows: false,
            lock_revision: false,
        });

        let prot = wb.workbook_xml.workbook_protection.as_ref().unwrap();
        assert!(prot.workbook_password.is_none());
        assert_eq!(prot.lock_structure, Some(true));
        assert!(prot.lock_windows.is_none());
        assert!(prot.lock_revision.is_none());
    }

    #[test]
    fn test_protect_workbook_save_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("protected.xlsx");

        let mut wb = Workbook::new();
        wb.protect_workbook(crate::protection::WorkbookProtectionConfig {
            password: Some("hello".to_string()),
            lock_structure: true,
            lock_windows: true,
            lock_revision: false,
        });
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert!(wb2.is_workbook_protected());
        let prot = wb2.workbook_xml.workbook_protection.as_ref().unwrap();
        assert!(prot.workbook_password.is_some());
        assert_eq!(prot.lock_structure, Some(true));
        assert_eq!(prot.lock_windows, Some(true));
    }

    #[test]
    fn test_is_workbook_protected() {
        let wb = Workbook::new();
        assert!(!wb.is_workbook_protected());

        let mut wb2 = Workbook::new();
        wb2.protect_workbook(crate::protection::WorkbookProtectionConfig {
            password: None,
            lock_structure: false,
            lock_windows: false,
            lock_revision: false,
        });
        // Even with no locks, the protection element is present
        assert!(wb2.is_workbook_protected());
    }

    #[test]
    fn test_unprotect_already_unprotected() {
        let mut wb = Workbook::new();
        assert!(!wb.is_workbook_protected());
        // Should be a no-op, not panic
        wb.unprotect_workbook();
        assert!(!wb.is_workbook_protected());
    }

    // -----------------------------------------------------------------------
    // Document properties tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_set_get_doc_props() {
        let mut wb = Workbook::new();
        let props = crate::doc_props::DocProperties {
            title: Some("My Title".to_string()),
            subject: Some("My Subject".to_string()),
            creator: Some("Author".to_string()),
            keywords: Some("rust, excel".to_string()),
            description: Some("A test workbook".to_string()),
            last_modified_by: Some("Editor".to_string()),
            revision: Some("2".to_string()),
            created: Some("2024-01-01T00:00:00Z".to_string()),
            modified: Some("2024-06-01T12:00:00Z".to_string()),
            category: Some("Testing".to_string()),
            content_status: Some("Draft".to_string()),
        };
        wb.set_doc_props(props);

        let got = wb.get_doc_props();
        assert_eq!(got.title.as_deref(), Some("My Title"));
        assert_eq!(got.subject.as_deref(), Some("My Subject"));
        assert_eq!(got.creator.as_deref(), Some("Author"));
        assert_eq!(got.keywords.as_deref(), Some("rust, excel"));
        assert_eq!(got.description.as_deref(), Some("A test workbook"));
        assert_eq!(got.last_modified_by.as_deref(), Some("Editor"));
        assert_eq!(got.revision.as_deref(), Some("2"));
        assert_eq!(got.created.as_deref(), Some("2024-01-01T00:00:00Z"));
        assert_eq!(got.modified.as_deref(), Some("2024-06-01T12:00:00Z"));
        assert_eq!(got.category.as_deref(), Some("Testing"));
        assert_eq!(got.content_status.as_deref(), Some("Draft"));
    }

    #[test]
    fn test_set_get_app_props() {
        let mut wb = Workbook::new();
        let props = crate::doc_props::AppProperties {
            application: Some("SheetKit".to_string()),
            doc_security: Some(0),
            company: Some("Acme Corp".to_string()),
            app_version: Some("1.0.0".to_string()),
            manager: Some("Boss".to_string()),
            template: Some("default.xltx".to_string()),
        };
        wb.set_app_props(props);

        let got = wb.get_app_props();
        assert_eq!(got.application.as_deref(), Some("SheetKit"));
        assert_eq!(got.doc_security, Some(0));
        assert_eq!(got.company.as_deref(), Some("Acme Corp"));
        assert_eq!(got.app_version.as_deref(), Some("1.0.0"));
        assert_eq!(got.manager.as_deref(), Some("Boss"));
        assert_eq!(got.template.as_deref(), Some("default.xltx"));
    }

    #[test]
    fn test_custom_property_crud() {
        let mut wb = Workbook::new();

        // Set
        wb.set_custom_property(
            "Project",
            crate::doc_props::CustomPropertyValue::String("SheetKit".to_string()),
        );

        // Get
        let val = wb.get_custom_property("Project");
        assert_eq!(
            val,
            Some(crate::doc_props::CustomPropertyValue::String(
                "SheetKit".to_string()
            ))
        );

        // Update
        wb.set_custom_property(
            "Project",
            crate::doc_props::CustomPropertyValue::String("Updated".to_string()),
        );
        let val = wb.get_custom_property("Project");
        assert_eq!(
            val,
            Some(crate::doc_props::CustomPropertyValue::String(
                "Updated".to_string()
            ))
        );

        // Delete
        assert!(wb.delete_custom_property("Project"));
        assert!(wb.get_custom_property("Project").is_none());
        assert!(!wb.delete_custom_property("Project")); // already gone
    }

    #[test]
    fn test_doc_props_save_open_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("doc_props.xlsx");

        let mut wb = Workbook::new();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Test Title".to_string()),
            creator: Some("Test Author".to_string()),
            created: Some("2024-01-01T00:00:00Z".to_string()),
            ..Default::default()
        });
        wb.set_app_props(crate::doc_props::AppProperties {
            application: Some("SheetKit".to_string()),
            company: Some("TestCorp".to_string()),
            ..Default::default()
        });
        wb.set_custom_property("Version", crate::doc_props::CustomPropertyValue::Int(42));
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let doc = wb2.get_doc_props();
        assert_eq!(doc.title.as_deref(), Some("Test Title"));
        assert_eq!(doc.creator.as_deref(), Some("Test Author"));
        assert_eq!(doc.created.as_deref(), Some("2024-01-01T00:00:00Z"));

        let app = wb2.get_app_props();
        assert_eq!(app.application.as_deref(), Some("SheetKit"));
        assert_eq!(app.company.as_deref(), Some("TestCorp"));

        let custom = wb2.get_custom_property("Version");
        assert_eq!(custom, Some(crate::doc_props::CustomPropertyValue::Int(42)));
    }

    #[test]
    fn test_open_without_doc_props() {
        // A newly created workbook saved without setting doc props should
        // still open gracefully (core/app/custom properties are all None).
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("no_props.xlsx");

        let wb = Workbook::new();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let doc = wb2.get_doc_props();
        assert!(doc.title.is_none());
        assert!(doc.creator.is_none());

        let app = wb2.get_app_props();
        assert!(app.application.is_none());

        assert!(wb2.get_custom_property("anything").is_none());
    }

    #[test]
    fn test_custom_property_multiple_types() {
        let mut wb = Workbook::new();

        wb.set_custom_property(
            "StringProp",
            crate::doc_props::CustomPropertyValue::String("hello".to_string()),
        );
        wb.set_custom_property("IntProp", crate::doc_props::CustomPropertyValue::Int(-7));
        wb.set_custom_property(
            "FloatProp",
            crate::doc_props::CustomPropertyValue::Float(3.14),
        );
        wb.set_custom_property(
            "BoolProp",
            crate::doc_props::CustomPropertyValue::Bool(true),
        );
        wb.set_custom_property(
            "DateProp",
            crate::doc_props::CustomPropertyValue::DateTime("2024-01-01T00:00:00Z".to_string()),
        );

        assert_eq!(
            wb.get_custom_property("StringProp"),
            Some(crate::doc_props::CustomPropertyValue::String(
                "hello".to_string()
            ))
        );
        assert_eq!(
            wb.get_custom_property("IntProp"),
            Some(crate::doc_props::CustomPropertyValue::Int(-7))
        );
        assert_eq!(
            wb.get_custom_property("FloatProp"),
            Some(crate::doc_props::CustomPropertyValue::Float(3.14))
        );
        assert_eq!(
            wb.get_custom_property("BoolProp"),
            Some(crate::doc_props::CustomPropertyValue::Bool(true))
        );
        assert_eq!(
            wb.get_custom_property("DateProp"),
            Some(crate::doc_props::CustomPropertyValue::DateTime(
                "2024-01-01T00:00:00Z".to_string()
            ))
        );
    }

    #[test]
    fn test_doc_props_default_values() {
        let wb = Workbook::new();
        let doc = wb.get_doc_props();
        assert!(doc.title.is_none());
        assert!(doc.subject.is_none());
        assert!(doc.creator.is_none());
        assert!(doc.keywords.is_none());
        assert!(doc.description.is_none());
        assert!(doc.last_modified_by.is_none());
        assert!(doc.revision.is_none());
        assert!(doc.created.is_none());
        assert!(doc.modified.is_none());
        assert!(doc.category.is_none());
        assert!(doc.content_status.is_none());

        let app = wb.get_app_props();
        assert!(app.application.is_none());
        assert!(app.doc_security.is_none());
        assert!(app.company.is_none());
        assert!(app.app_version.is_none());
        assert!(app.manager.is_none());
        assert!(app.template.is_none());
    }
}
