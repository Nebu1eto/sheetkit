//! Workbook file I/O: reading and writing `.xlsx` files.
//!
//! An `.xlsx` file is a ZIP archive containing XML parts. This module provides
//! [`Workbook`] which holds the parsed XML structures in memory and can
//! serialize them back to a valid `.xlsx` file.

use std::collections::{HashMap, HashSet};
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
use crate::cell_ref_shift::shift_cell_references_in_text;
use crate::chart::ChartConfig;
use crate::comment::CommentConfig;
use crate::conditional::ConditionalFormatRule;
use crate::error::{Error, Result};
use crate::image::ImageConfig;
use crate::pivot::{PivotTableConfig, PivotTableInfo};
use crate::protection::WorkbookProtectionConfig;
use crate::sst::SharedStringTable;
use crate::utils::cell_ref::{cell_name_to_coordinates, column_name_to_number};
use crate::utils::constants::MAX_CELL_CHARS;
use crate::validation::DataValidationConfig;
use crate::workbook_paths::{
    default_relationships, relationship_part_path, relative_relationship_target,
    resolve_relationship_target,
};

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
    /// Chart parts preserved as raw XML when typed parsing is not supported.
    raw_charts: Vec<(String, Vec<u8>)>,
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
    /// Pivot table parts: (zip path, PivotTableDefinition data).
    pivot_tables: Vec<(String, sheetkit_xml::pivot_table::PivotTableDefinition)>,
    /// Pivot cache definition parts: (zip path, PivotCacheDefinition data).
    pivot_cache_defs: Vec<(String, sheetkit_xml::pivot_cache::PivotCacheDefinition)>,
    /// Pivot cache records parts: (zip path, PivotCacheRecords data).
    pivot_cache_records: Vec<(String, sheetkit_xml::pivot_cache::PivotCacheRecords)>,
    /// Raw theme XML bytes from xl/theme/theme1.xml (preserved for round-trip).
    theme_xml: Option<Vec<u8>>,
    /// Parsed theme colors from the theme XML.
    theme_colors: sheetkit_xml::theme::ThemeColors,
    /// Per-sheet sparkline configurations, parallel to the `worksheets` vector.
    sheet_sparklines: Vec<Vec<crate::sparkline::SparklineConfig>>,
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
            raw_charts: vec![],
            drawings: vec![],
            images: vec![],
            worksheet_drawings: HashMap::new(),
            worksheet_rels: HashMap::new(),
            drawing_rels: HashMap::new(),
            core_properties: None,
            app_properties: None,
            custom_properties: None,
            pivot_tables: vec![],
            pivot_cache_defs: vec![],
            pivot_cache_records: vec![],
            theme_xml: None,
            theme_colors: crate::theme::default_theme_colors(),
            sheet_sparklines: vec![vec![]],
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

        // Parse each worksheet referenced in the workbook.
        let mut worksheets = Vec::new();
        let mut worksheet_paths = Vec::new();
        for sheet_entry in &workbook_xml.sheets.sheets {
            // Find the relationship target for this sheet's rId.
            let rel = workbook_rels
                .relationships
                .iter()
                .find(|r| r.id == sheet_entry.r_id && r.rel_type == rel_types::WORKSHEET);

            let rel = rel.ok_or_else(|| {
                Error::Internal(format!(
                    "missing worksheet relationship for sheet '{}'",
                    sheet_entry.name
                ))
            })?;

            let sheet_path = resolve_relationship_target("xl/workbook.xml", &rel.target);
            let ws: WorksheetXml = read_xml_part(&mut archive, &sheet_path)?;
            worksheets.push((sheet_entry.name.clone(), ws));
            worksheet_paths.push(sheet_path);
        }

        // Parse xl/styles.xml
        let stylesheet: StyleSheet = read_xml_part(&mut archive, "xl/styles.xml")?;

        // Parse xl/sharedStrings.xml (optional -- may not exist for workbooks with no strings)
        let shared_strings: Sst =
            read_xml_part(&mut archive, "xl/sharedStrings.xml").unwrap_or_default();

        let sst_runtime = SharedStringTable::from_sst(&shared_strings);

        // Parse xl/theme/theme1.xml (optional -- preserved as raw bytes for round-trip).
        let (theme_xml, theme_colors) = match read_bytes_part(&mut archive, "xl/theme/theme1.xml") {
            Ok(bytes) => {
                let colors = sheetkit_xml::theme::parse_theme_colors(&bytes);
                (Some(bytes), colors)
            }
            Err(_) => (None, crate::theme::default_theme_colors()),
        };

        // Parse per-sheet worksheet relationship files (optional).
        let mut worksheet_rels: HashMap<usize, Relationships> = HashMap::new();
        for (i, sheet_path) in worksheet_paths.iter().enumerate() {
            let rels_path = relationship_part_path(sheet_path);
            if let Ok(rels) = read_xml_part::<Relationships>(&mut archive, &rels_path) {
                worksheet_rels.insert(i, rels);
            }
        }

        // Parse comments, drawings, drawing rels, charts, and images.
        let mut sheet_comments: Vec<Option<Comments>> = vec![None; worksheets.len()];
        let mut drawings: Vec<(String, WsDr)> = Vec::new();
        let mut worksheet_drawings: HashMap<usize, usize> = HashMap::new();
        let mut drawing_path_to_idx: HashMap<String, usize> = HashMap::new();

        for (sheet_idx, sheet_path) in worksheet_paths.iter().enumerate() {
            let Some(rels) = worksheet_rels.get(&sheet_idx) else {
                continue;
            };

            if let Some(comment_rel) = rels
                .relationships
                .iter()
                .find(|r| r.rel_type == rel_types::COMMENTS)
            {
                let comment_path = resolve_relationship_target(sheet_path, &comment_rel.target);
                if let Ok(comments) = read_xml_part::<Comments>(&mut archive, &comment_path) {
                    sheet_comments[sheet_idx] = Some(comments);
                }
            }

            if let Some(drawing_rel) = rels
                .relationships
                .iter()
                .find(|r| r.rel_type == rel_types::DRAWING)
            {
                let drawing_path = resolve_relationship_target(sheet_path, &drawing_rel.target);
                let drawing_idx = if let Some(idx) = drawing_path_to_idx.get(&drawing_path) {
                    *idx
                } else if let Ok(drawing) = read_xml_part::<WsDr>(&mut archive, &drawing_path) {
                    let idx = drawings.len();
                    drawings.push((drawing_path.clone(), drawing));
                    drawing_path_to_idx.insert(drawing_path.clone(), idx);
                    idx
                } else {
                    continue;
                };
                worksheet_drawings.insert(sheet_idx, drawing_idx);
            }
        }

        // Fallback: load drawing parts listed in content types even when they
        // are not discoverable via worksheet rel parsing.
        for ovr in &content_types.overrides {
            if ovr.content_type != mime_types::DRAWING {
                continue;
            }
            let drawing_path = ovr.part_name.trim_start_matches('/').to_string();
            if drawing_path_to_idx.contains_key(&drawing_path) {
                continue;
            }
            if let Ok(drawing) = read_xml_part::<WsDr>(&mut archive, &drawing_path) {
                let idx = drawings.len();
                drawings.push((drawing_path.clone(), drawing));
                drawing_path_to_idx.insert(drawing_path, idx);
            }
        }

        let mut drawing_rels: HashMap<usize, Relationships> = HashMap::new();
        let mut charts: Vec<(String, ChartSpace)> = Vec::new();
        let mut raw_charts: Vec<(String, Vec<u8>)> = Vec::new();
        let mut images: Vec<(String, Vec<u8>)> = Vec::new();
        let mut seen_chart_paths: HashSet<String> = HashSet::new();
        let mut seen_image_paths: HashSet<String> = HashSet::new();

        for (drawing_idx, (drawing_path, _)) in drawings.iter().enumerate() {
            let drawing_rels_path = relationship_part_path(drawing_path);
            let Ok(rels) = read_xml_part::<Relationships>(&mut archive, &drawing_rels_path) else {
                continue;
            };

            for rel in &rels.relationships {
                if rel.rel_type == rel_types::CHART {
                    let chart_path = resolve_relationship_target(drawing_path, &rel.target);
                    if seen_chart_paths.insert(chart_path.clone()) {
                        match read_xml_part::<ChartSpace>(&mut archive, &chart_path) {
                            Ok(chart) => charts.push((chart_path, chart)),
                            Err(_) => {
                                if let Ok(bytes) = read_bytes_part(&mut archive, &chart_path) {
                                    raw_charts.push((chart_path, bytes));
                                }
                            }
                        }
                    }
                } else if rel.rel_type == rel_types::IMAGE {
                    let image_path = resolve_relationship_target(drawing_path, &rel.target);
                    if seen_image_paths.insert(image_path.clone()) {
                        if let Ok(bytes) = read_bytes_part(&mut archive, &image_path) {
                            images.push((image_path, bytes));
                        }
                    }
                }
            }

            drawing_rels.insert(drawing_idx, rels);
        }

        // Fallback: load chart parts listed in content types even when no
        // drawing relationship was read.
        for ovr in &content_types.overrides {
            if ovr.content_type != mime_types::CHART {
                continue;
            }
            let chart_path = ovr.part_name.trim_start_matches('/').to_string();
            if seen_chart_paths.insert(chart_path.clone()) {
                match read_xml_part::<ChartSpace>(&mut archive, &chart_path) {
                    Ok(chart) => charts.push((chart_path, chart)),
                    Err(_) => {
                        if let Ok(bytes) = read_bytes_part(&mut archive, &chart_path) {
                            raw_charts.push((chart_path, bytes));
                        }
                    }
                }
            }
        }

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

        // Parse pivot cache definitions, pivot tables, and pivot cache records.
        let mut pivot_cache_defs = Vec::new();
        let mut pivot_tables = Vec::new();
        let mut pivot_cache_records = Vec::new();
        for ovr in &content_types.overrides {
            let path = ovr.part_name.trim_start_matches('/');
            if ovr.content_type == mime_types::PIVOT_CACHE_DEFINITION {
                if let Ok(pcd) = read_xml_part::<sheetkit_xml::pivot_cache::PivotCacheDefinition>(
                    &mut archive,
                    path,
                ) {
                    pivot_cache_defs.push((path.to_string(), pcd));
                }
            } else if ovr.content_type == mime_types::PIVOT_TABLE {
                if let Ok(pt) = read_xml_part::<sheetkit_xml::pivot_table::PivotTableDefinition>(
                    &mut archive,
                    path,
                ) {
                    pivot_tables.push((path.to_string(), pt));
                }
            } else if ovr.content_type == mime_types::PIVOT_CACHE_RECORDS {
                if let Ok(pcr) = read_xml_part::<sheetkit_xml::pivot_cache::PivotCacheRecords>(
                    &mut archive,
                    path,
                ) {
                    pivot_cache_records.push((path.to_string(), pcr));
                }
            }
        }

        // Parse sparklines from worksheet extension lists.
        let mut sheet_sparklines: Vec<Vec<crate::sparkline::SparklineConfig>> =
            vec![vec![]; worksheets.len()];
        for (i, ws_path) in worksheet_paths.iter().enumerate() {
            if let Ok(raw) = read_string_part(&mut archive, ws_path) {
                let parsed = parse_sparklines_from_xml(&raw);
                if !parsed.is_empty() {
                    sheet_sparklines[i] = parsed;
                }
            }
        }

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
            charts,
            raw_charts,
            drawings,
            images,
            worksheet_drawings,
            worksheet_rels,
            drawing_rels,
            core_properties,
            app_properties,
            custom_properties,
            pivot_tables,
            pivot_cache_defs,
            pivot_cache_records,
            theme_xml,
            theme_colors,
            sheet_sparklines,
        })
    }

    /// Save the workbook to a `.xlsx` file at the given path.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let file = std::fs::File::create(path)?;
        let mut zip = zip::ZipWriter::new(file);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        let mut content_types = self.content_types.clone();
        let mut worksheet_rels = self.worksheet_rels.clone();

        // Synchronize comment parts with worksheet relationships/content types.
        for sheet_idx in 0..self.worksheets.len() {
            let has_comments = self
                .sheet_comments
                .get(sheet_idx)
                .and_then(|c| c.as_ref())
                .is_some();
            if let Some(rels) = worksheet_rels.get_mut(&sheet_idx) {
                rels.relationships
                    .retain(|r| r.rel_type != rel_types::COMMENTS);
            }
            if !has_comments {
                continue;
            }

            let comment_path = format!("xl/comments{}.xml", sheet_idx + 1);
            let part_name = format!("/{}", comment_path);
            if !content_types
                .overrides
                .iter()
                .any(|o| o.part_name == part_name && o.content_type == mime_types::COMMENTS)
            {
                content_types.overrides.push(ContentTypeOverride {
                    part_name,
                    content_type: mime_types::COMMENTS.to_string(),
                });
            }

            let sheet_path = self.sheet_part_path(sheet_idx);
            let target = relative_relationship_target(&sheet_path, &comment_path);
            let rels = worksheet_rels
                .entry(sheet_idx)
                .or_insert_with(default_relationships);
            let rid = crate::sheet::next_rid(&rels.relationships);
            rels.relationships.push(Relationship {
                id: rid,
                rel_type: rel_types::COMMENTS.to_string(),
                target,
                target_mode: None,
            });
        }

        // [Content_Types].xml
        write_xml_part(&mut zip, "[Content_Types].xml", &content_types, options)?;

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
            let entry_name = self.sheet_part_path(i);
            let sparklines = self.sheet_sparklines.get(i).cloned().unwrap_or_default();
            if sparklines.is_empty() {
                write_xml_part(&mut zip, &entry_name, ws, options)?;
            } else {
                let xml = serialize_worksheet_with_sparklines(ws, &sparklines)?;
                zip.start_file(&entry_name, options)
                    .map_err(|e| Error::Zip(e.to_string()))?;
                zip.write_all(xml.as_bytes())?;
            }
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
        for (path, data) in &self.raw_charts {
            if self.charts.iter().any(|(p, _)| p == path) {
                continue;
            }
            zip.start_file(path, options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(data)?;
        }

        // xl/media/image{N}.{ext} -- write image data
        for (path, data) in &self.images {
            zip.start_file(path, options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(data)?;
        }

        // xl/worksheets/_rels/sheet{N}.xml.rels -- write worksheet relationships
        for (sheet_idx, rels) in &worksheet_rels {
            let sheet_path = self.sheet_part_path(*sheet_idx);
            let path = relationship_part_path(&sheet_path);
            write_xml_part(&mut zip, &path, rels, options)?;
        }

        // xl/drawings/_rels/drawing{N}.xml.rels -- write drawing relationships
        for (drawing_idx, rels) in &self.drawing_rels {
            if let Some((drawing_path, _)) = self.drawings.get(*drawing_idx) {
                let path = relationship_part_path(drawing_path);
                write_xml_part(&mut zip, &path, rels, options)?;
            }
        }

        // xl/pivotTables/pivotTable{N}.xml
        for (path, pt) in &self.pivot_tables {
            write_xml_part(&mut zip, path, pt, options)?;
        }

        // xl/pivotCache/pivotCacheDefinition{N}.xml
        for (path, pcd) in &self.pivot_cache_defs {
            write_xml_part(&mut zip, path, pcd, options)?;
        }

        // xl/pivotCache/pivotCacheRecords{N}.xml
        for (path, pcr) in &self.pivot_cache_records {
            write_xml_part(&mut zip, path, pcr, options)?;
        }

        // xl/theme/theme1.xml
        {
            let default_theme = crate::theme::default_theme_xml();
            let theme_bytes = self.theme_xml.as_deref().unwrap_or(&default_theme);
            zip.start_file("xl/theme/theme1.xml", options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(theme_bytes)?;
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

    /// Create a new empty sheet with the given name. Returns the 0-based sheet index.
    pub fn new_sheet(&mut self, name: &str) -> Result<usize> {
        let idx = crate::sheet::add_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            name,
            WorksheetXml::default(),
        )?;
        if self.sheet_comments.len() < self.worksheets.len() {
            self.sheet_comments.push(None);
        }
        if self.sheet_sparklines.len() < self.worksheets.len() {
            self.sheet_sparklines.push(vec![]);
        }
        Ok(idx)
    }

    /// Delete a sheet by name.
    pub fn delete_sheet(&mut self, name: &str) -> Result<()> {
        let idx = self.sheet_index(name)?;
        crate::sheet::delete_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            name,
        )?;

        if idx < self.sheet_comments.len() {
            self.sheet_comments.remove(idx);
        }
        if idx < self.sheet_sparklines.len() {
            self.sheet_sparklines.remove(idx);
        }
        self.reindex_sheet_maps_after_delete(idx);
        Ok(())
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
        let idx = crate::sheet::copy_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            source,
            target,
        )?;
        if self.sheet_comments.len() < self.worksheets.len() {
            self.sheet_comments.push(None);
        }
        let source_sparklines = {
            let src_idx = self.sheet_index(source).unwrap_or(0);
            self.sheet_sparklines
                .get(src_idx)
                .cloned()
                .unwrap_or_default()
        };
        if self.sheet_sparklines.len() < self.worksheets.len() {
            self.sheet_sparklines.push(source_sparklines);
        }
        Ok(idx)
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
        let idx = crate::sheet::add_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            &sheet_name,
            ws,
        )?;
        if self.sheet_comments.len() < self.worksheets.len() {
            self.sheet_comments.push(None);
        }
        Ok(idx)
    }

    /// Insert `count` empty rows starting at `start_row` in the named sheet.
    pub fn insert_rows(&mut self, sheet: &str, start_row: u32, count: u32) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        {
            let ws = &mut self.worksheets[sheet_idx].1;
            crate::row::insert_rows(ws, start_row, count)?;
        }
        self.apply_reference_shift_for_sheet(sheet_idx, |col, row| {
            if row >= start_row {
                (col, row + count)
            } else {
                (col, row)
            }
        })
    }

    /// Remove a single row from the named sheet, shifting rows below it up.
    pub fn remove_row(&mut self, sheet: &str, row: u32) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        {
            let ws = &mut self.worksheets[sheet_idx].1;
            crate::row::remove_row(ws, row)?;
        }
        self.apply_reference_shift_for_sheet(sheet_idx, |col, r| {
            if r > row {
                (col, r - 1)
            } else {
                (col, r)
            }
        })
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

    /// Get the visibility of a row. Returns true if visible (not hidden).
    pub fn get_row_visible(&self, sheet: &str, row: u32) -> Result<bool> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::row::get_row_visible(ws, row))
    }

    /// Set the outline level of a row.
    pub fn set_row_outline_level(&mut self, sheet: &str, row: u32, level: u8) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::row::set_row_outline_level(ws, row, level)
    }

    /// Get the outline level of a row. Returns 0 if not set.
    pub fn get_row_outline_level(&self, sheet: &str, row: u32) -> Result<u8> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::row::get_row_outline_level(ws, row))
    }

    /// Set the style for an entire row.
    ///
    /// The `style_id` must be a valid index in cellXfs (returned by `add_style`).
    pub fn set_row_style(&mut self, sheet: &str, row: u32, style_id: u32) -> Result<()> {
        if style_id as usize >= self.stylesheet.cell_xfs.xfs.len() {
            return Err(Error::StyleNotFound { id: style_id });
        }
        let ws = self.worksheet_mut(sheet)?;
        crate::row::set_row_style(ws, row, style_id)
    }

    /// Get the style ID for a row. Returns 0 (default) if not set.
    pub fn get_row_style(&self, sheet: &str, row: u32) -> Result<u32> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::row::get_row_style(ws, row))
    }

    /// Get all rows with their data from a sheet.
    ///
    /// Returns a Vec of `(row_number, Vec<(column_name, CellValue)>)` tuples.
    /// Only rows that contain at least one cell are included (sparse).
    #[allow(clippy::type_complexity)]
    pub fn get_rows(&self, sheet: &str) -> Result<Vec<(u32, Vec<(String, CellValue)>)>> {
        let ws = self.worksheet_ref(sheet)?;
        crate::row::get_rows(ws, &self.sst_runtime)
    }

    /// Get all columns with their data from a sheet.
    ///
    /// Returns a Vec of `(column_name, Vec<(row_number, CellValue)>)` tuples.
    /// Only columns that have data are included (sparse).
    #[allow(clippy::type_complexity)]
    pub fn get_cols(&self, sheet: &str) -> Result<Vec<(String, Vec<(u32, CellValue)>)>> {
        let ws = self.worksheet_ref(sheet)?;
        crate::col::get_cols(ws, &self.sst_runtime)
    }

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

    /// Get the visibility of a column. Returns true if visible (not hidden).
    pub fn get_col_visible(&self, sheet: &str, col: &str) -> Result<bool> {
        let ws = self.worksheet_ref(sheet)?;
        crate::col::get_col_visible(ws, col)
    }

    /// Set the outline level of a column.
    pub fn set_col_outline_level(&mut self, sheet: &str, col: &str, level: u8) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::col::set_col_outline_level(ws, col, level)
    }

    /// Get the outline level of a column. Returns 0 if not set.
    pub fn get_col_outline_level(&self, sheet: &str, col: &str) -> Result<u8> {
        let ws = self.worksheet_ref(sheet)?;
        crate::col::get_col_outline_level(ws, col)
    }

    /// Set the style for an entire column.
    ///
    /// The `style_id` must be a valid index in cellXfs (returned by `add_style`).
    pub fn set_col_style(&mut self, sheet: &str, col: &str, style_id: u32) -> Result<()> {
        if style_id as usize >= self.stylesheet.cell_xfs.xfs.len() {
            return Err(Error::StyleNotFound { id: style_id });
        }
        let ws = self.worksheet_mut(sheet)?;
        crate::col::set_col_style(ws, col, style_id)
    }

    /// Get the style ID for a column. Returns 0 (default) if not set.
    pub fn get_col_style(&self, sheet: &str, col: &str) -> Result<u32> {
        let ws = self.worksheet_ref(sheet)?;
        crate::col::get_col_style(ws, col)
    }

    /// Insert `count` columns starting at `col` in the named sheet.
    pub fn insert_cols(&mut self, sheet: &str, col: &str, count: u32) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let start_col = column_name_to_number(col)?;
        {
            let ws = &mut self.worksheets[sheet_idx].1;
            crate::col::insert_cols(ws, col, count)?;
        }
        self.apply_reference_shift_for_sheet(sheet_idx, |c, row| {
            if c >= start_col {
                (c + count, row)
            } else {
                (c, row)
            }
        })
    }

    /// Remove a single column from the named sheet.
    pub fn remove_col(&mut self, sheet: &str, col: &str) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let col_num = column_name_to_number(col)?;
        {
            let ws = &mut self.worksheets[sheet_idx].1;
            crate::col::remove_col(ws, col)?;
        }
        self.apply_reference_shift_for_sheet(sheet_idx, |c, row| {
            if c > col_num {
                (c - 1, row)
            } else {
                (c, row)
            }
        })
    }

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

    /// Merge a range of cells on the given sheet.
    ///
    /// `top_left` and `bottom_right` are cell references like "A1" and "C3".
    /// Returns an error if the range overlaps with an existing merge region.
    pub fn merge_cells(&mut self, sheet: &str, top_left: &str, bottom_right: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::merge::merge_cells(ws, top_left, bottom_right)
    }

    /// Remove a merged cell range from the given sheet.
    ///
    /// `reference` is the exact range string like "A1:C3".
    pub fn unmerge_cell(&mut self, sheet: &str, reference: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::merge::unmerge_cell(ws, reference)
    }

    /// Get all merged cell ranges on the given sheet.
    ///
    /// Returns a list of range strings like `["A1:B2", "D1:F3"]`.
    pub fn get_merge_cells(&self, sheet: &str) -> Result<Vec<String>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::merge::get_merge_cells(ws))
    }

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

    /// Set conditional formatting rules on a cell range of a sheet.
    pub fn set_conditional_format(
        &mut self,
        sheet: &str,
        sqref: &str,
        rules: &[ConditionalFormatRule],
    ) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        let ws = &mut self.worksheets[idx].1;
        crate::conditional::set_conditional_format(ws, &mut self.stylesheet, sqref, rules)
    }

    /// Get all conditional formatting rules for a sheet.
    ///
    /// Returns a list of `(sqref, rules)` pairs.
    pub fn get_conditional_formats(
        &self,
        sheet: &str,
    ) -> Result<Vec<(String, Vec<ConditionalFormatRule>)>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::conditional::get_conditional_formats(
            ws,
            &self.stylesheet,
        ))
    }

    /// Delete conditional formatting rules for a specific cell range on a sheet.
    pub fn delete_conditional_format(&mut self, sheet: &str, sqref: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::conditional::delete_conditional_format(ws, sqref)
    }

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

    /// Set freeze panes on a sheet.
    ///
    /// The cell reference indicates the top-left cell of the scrollable area.
    /// For example, `"A2"` freezes row 1, `"B1"` freezes column A, and `"B2"`
    /// freezes both row 1 and column A.
    pub fn set_panes(&mut self, sheet: &str, cell: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::sheet::set_panes(ws, cell)
    }

    /// Remove any freeze or split panes from a sheet.
    pub fn unset_panes(&mut self, sheet: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::sheet::unset_panes(ws);
        Ok(())
    }

    /// Get the current freeze pane cell reference for a sheet, if any.
    ///
    /// Returns the top-left cell of the unfrozen area (e.g., `"A2"` if row 1
    /// is frozen), or `None` if no panes are configured.
    pub fn get_panes(&self, sheet: &str) -> Result<Option<String>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::sheet::get_panes(ws))
    }

    // -- Page layout methods --

    /// Set page margins on a sheet.
    pub fn set_page_margins(
        &mut self,
        sheet: &str,
        margins: &crate::page_layout::PageMarginsConfig,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::set_page_margins(ws, margins)
    }

    /// Get page margins for a sheet, returning Excel defaults if not set.
    pub fn get_page_margins(&self, sheet: &str) -> Result<crate::page_layout::PageMarginsConfig> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_page_margins(ws))
    }

    /// Set page setup options (orientation, paper size, scale, fit-to-page).
    ///
    /// Only non-`None` parameters are applied; existing values for `None`
    /// parameters are preserved.
    pub fn set_page_setup(
        &mut self,
        sheet: &str,
        orientation: Option<crate::page_layout::Orientation>,
        paper_size: Option<crate::page_layout::PaperSize>,
        scale: Option<u32>,
        fit_to_width: Option<u32>,
        fit_to_height: Option<u32>,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::set_page_setup(
            ws,
            orientation,
            paper_size,
            scale,
            fit_to_width,
            fit_to_height,
        )
    }

    /// Get the page orientation for a sheet.
    pub fn get_orientation(&self, sheet: &str) -> Result<Option<crate::page_layout::Orientation>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_orientation(ws))
    }

    /// Get the paper size for a sheet.
    pub fn get_paper_size(&self, sheet: &str) -> Result<Option<crate::page_layout::PaperSize>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_paper_size(ws))
    }

    /// Get scale, fit-to-width, and fit-to-height values for a sheet.
    ///
    /// Returns `(scale, fit_to_width, fit_to_height)`, each `None` if not set.
    pub fn get_page_setup_details(
        &self,
        sheet: &str,
    ) -> Result<(Option<u32>, Option<u32>, Option<u32>)> {
        let ws = self.worksheet_ref(sheet)?;
        Ok((
            crate::page_layout::get_scale(ws),
            crate::page_layout::get_fit_to_width(ws),
            crate::page_layout::get_fit_to_height(ws),
        ))
    }

    /// Set header and footer text for printing.
    pub fn set_header_footer(
        &mut self,
        sheet: &str,
        header: Option<&str>,
        footer: Option<&str>,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::set_header_footer(ws, header, footer)
    }

    /// Get the header and footer text for a sheet.
    pub fn get_header_footer(&self, sheet: &str) -> Result<(Option<String>, Option<String>)> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_header_footer(ws))
    }

    /// Set print options on a sheet.
    pub fn set_print_options(
        &mut self,
        sheet: &str,
        grid_lines: Option<bool>,
        headings: Option<bool>,
        h_centered: Option<bool>,
        v_centered: Option<bool>,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::set_print_options(ws, grid_lines, headings, h_centered, v_centered)
    }

    /// Get print options for a sheet.
    ///
    /// Returns `(grid_lines, headings, horizontal_centered, vertical_centered)`.
    #[allow(clippy::type_complexity)]
    pub fn get_print_options(
        &self,
        sheet: &str,
    ) -> Result<(Option<bool>, Option<bool>, Option<bool>, Option<bool>)> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_print_options(ws))
    }

    /// Insert a horizontal page break before the given 1-based row.
    pub fn insert_page_break(&mut self, sheet: &str, row: u32) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::insert_page_break(ws, row)
    }

    /// Remove a horizontal page break at the given 1-based row.
    pub fn remove_page_break(&mut self, sheet: &str, row: u32) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::remove_page_break(ws, row)
    }

    /// Get all row page break positions (1-based row numbers).
    pub fn get_page_breaks(&self, sheet: &str) -> Result<Vec<u32>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_page_breaks(ws))
    }

    /// Set a hyperlink on a cell.
    ///
    /// For external URLs and email links, a relationship entry is created in
    /// the worksheet's `.rels` file. Internal sheet references use only the
    /// `location` attribute without a relationship.
    pub fn set_cell_hyperlink(
        &mut self,
        sheet: &str,
        cell: &str,
        link: crate::hyperlink::HyperlinkType,
        display: Option<&str>,
        tooltip: Option<&str>,
    ) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let ws = &mut self.worksheets[sheet_idx].1;
        let rels = self
            .worksheet_rels
            .entry(sheet_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        crate::hyperlink::set_cell_hyperlink(ws, rels, cell, &link, display, tooltip)
    }

    /// Get hyperlink information for a cell.
    ///
    /// Returns `None` if the cell has no hyperlink.
    pub fn get_cell_hyperlink(
        &self,
        sheet: &str,
        cell: &str,
    ) -> Result<Option<crate::hyperlink::HyperlinkInfo>> {
        let sheet_idx = self.sheet_index(sheet)?;
        let ws = &self.worksheets[sheet_idx].1;
        let empty_rels = Relationships {
            xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
            relationships: vec![],
        };
        let rels = self.worksheet_rels.get(&sheet_idx).unwrap_or(&empty_rels);
        crate::hyperlink::get_cell_hyperlink(ws, rels, cell)
    }

    /// Delete a hyperlink from a cell.
    ///
    /// Removes both the hyperlink element from the worksheet XML and any
    /// associated relationship entry.
    pub fn delete_cell_hyperlink(&mut self, sheet: &str, cell: &str) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let ws = &mut self.worksheets[sheet_idx].1;
        let rels = self
            .worksheet_rels
            .entry(sheet_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        crate::hyperlink::delete_cell_hyperlink(ws, rels, cell)
    }

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

    /// Add a pivot table to the workbook.
    ///
    /// The pivot table summarizes data from `config.source_sheet` /
    /// `config.source_range` and places its output on `config.target_sheet`
    /// starting at `config.target_cell`.
    pub fn add_pivot_table(&mut self, config: &PivotTableConfig) -> Result<()> {
        // Validate source sheet exists.
        let _src_idx = self.sheet_index(&config.source_sheet)?;

        // Validate target sheet exists.
        let target_idx = self.sheet_index(&config.target_sheet)?;

        // Check for duplicate name.
        if self
            .pivot_tables
            .iter()
            .any(|(_, pt)| pt.name == config.name)
        {
            return Err(Error::PivotTableAlreadyExists {
                name: config.name.clone(),
            });
        }

        // Read header row from the source data.
        let field_names = self.read_header_row(&config.source_sheet, &config.source_range)?;
        if field_names.is_empty() {
            return Err(Error::InvalidSourceRange(
                "source range header row is empty".to_string(),
            ));
        }

        // Assign a cache ID (next available).
        let cache_id = self
            .pivot_tables
            .iter()
            .map(|(_, pt)| pt.cache_id)
            .max()
            .map(|m| m + 1)
            .unwrap_or(0);

        // Build XML structures.
        let pt_def = crate::pivot::build_pivot_table_xml(config, cache_id, &field_names)?;
        let pcd = crate::pivot::build_pivot_cache_definition(
            &config.source_sheet,
            &config.source_range,
            &field_names,
        );
        let pcr = sheetkit_xml::pivot_cache::PivotCacheRecords {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            xmlns_r: sheetkit_xml::namespaces::RELATIONSHIPS.to_string(),
            count: Some(0),
            records: vec![],
        };

        // Determine part numbers.
        let pt_num = self.pivot_tables.len() + 1;
        let cache_num = self.pivot_cache_defs.len() + 1;

        let pt_path = format!("xl/pivotTables/pivotTable{}.xml", pt_num);
        let pcd_path = format!("xl/pivotCache/pivotCacheDefinition{}.xml", cache_num);
        let pcr_path = format!("xl/pivotCache/pivotCacheRecords{}.xml", cache_num);

        // Store parts.
        self.pivot_tables.push((pt_path.clone(), pt_def));
        self.pivot_cache_defs.push((pcd_path.clone(), pcd));
        self.pivot_cache_records.push((pcr_path.clone(), pcr));

        // Add content type overrides.
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/{}", pt_path),
            content_type: mime_types::PIVOT_TABLE.to_string(),
        });
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/{}", pcd_path),
            content_type: mime_types::PIVOT_CACHE_DEFINITION.to_string(),
        });
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/{}", pcr_path),
            content_type: mime_types::PIVOT_CACHE_RECORDS.to_string(),
        });

        // Add workbook relationship for pivot cache definition.
        let wb_rid = crate::sheet::next_rid(&self.workbook_rels.relationships);
        self.workbook_rels.relationships.push(Relationship {
            id: wb_rid.clone(),
            rel_type: rel_types::PIVOT_CACHE_DEF.to_string(),
            target: format!("pivotCache/pivotCacheDefinition{}.xml", cache_num),
            target_mode: None,
        });

        // Update workbook_xml.pivot_caches.
        let pivot_caches = self
            .workbook_xml
            .pivot_caches
            .get_or_insert_with(|| sheetkit_xml::workbook::PivotCaches { caches: vec![] });
        pivot_caches
            .caches
            .push(sheetkit_xml::workbook::PivotCacheEntry {
                cache_id,
                r_id: wb_rid,
            });

        // Add worksheet relationship for pivot table on the target sheet.
        let ws_rid = self.next_worksheet_rid(target_idx);
        let ws_rels = self
            .worksheet_rels
            .entry(target_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        ws_rels.relationships.push(Relationship {
            id: ws_rid,
            rel_type: rel_types::PIVOT_TABLE.to_string(),
            target: format!("../pivotTables/pivotTable{}.xml", pt_num),
            target_mode: None,
        });

        Ok(())
    }

    /// Get information about all pivot tables in the workbook.
    pub fn get_pivot_tables(&self) -> Vec<PivotTableInfo> {
        self.pivot_tables
            .iter()
            .map(|(_path, pt)| {
                // Find the matching cache definition by cache_id.
                let (source_sheet, source_range) = self
                    .pivot_cache_defs
                    .iter()
                    .enumerate()
                    .find(|(i, _)| {
                        self.workbook_xml
                            .pivot_caches
                            .as_ref()
                            .and_then(|pc| pc.caches.iter().find(|e| e.cache_id == pt.cache_id))
                            .is_some()
                            || *i == pt.cache_id as usize
                    })
                    .and_then(|(_, (_, pcd))| {
                        pcd.cache_source
                            .worksheet_source
                            .as_ref()
                            .map(|ws| (ws.sheet.clone(), ws.reference.clone()))
                    })
                    .unwrap_or_default();

                // Determine target sheet from the pivot table path.
                let target_sheet = self.find_pivot_table_target_sheet(pt).unwrap_or_default();

                PivotTableInfo {
                    name: pt.name.clone(),
                    source_sheet,
                    source_range,
                    target_sheet,
                    location: pt.location.reference.clone(),
                }
            })
            .collect()
    }

    /// Delete a pivot table by name.
    pub fn delete_pivot_table(&mut self, name: &str) -> Result<()> {
        // Find the pivot table.
        let pt_idx = self
            .pivot_tables
            .iter()
            .position(|(_, pt)| pt.name == name)
            .ok_or_else(|| Error::PivotTableNotFound {
                name: name.to_string(),
            })?;

        let (pt_path, pt_def) = self.pivot_tables.remove(pt_idx);
        let cache_id = pt_def.cache_id;

        // Remove the matching pivot cache definition and records.
        // Find the workbook_xml pivot cache entry for this cache_id.
        let mut wb_cache_rid = None;
        if let Some(ref mut pivot_caches) = self.workbook_xml.pivot_caches {
            if let Some(pos) = pivot_caches
                .caches
                .iter()
                .position(|e| e.cache_id == cache_id)
            {
                wb_cache_rid = Some(pivot_caches.caches[pos].r_id.clone());
                pivot_caches.caches.remove(pos);
            }
            if pivot_caches.caches.is_empty() {
                self.workbook_xml.pivot_caches = None;
            }
        }

        // Remove the workbook relationship for this cache.
        if let Some(ref rid) = wb_cache_rid {
            // Find the target to determine which cache def to remove.
            if let Some(rel) = self
                .workbook_rels
                .relationships
                .iter()
                .find(|r| r.id == *rid)
            {
                let target_path = format!("xl/{}", rel.target);
                self.pivot_cache_defs.retain(|(p, _)| *p != target_path);

                // Remove matching cache records (same numbering).
                let records_path = target_path.replace("pivotCacheDefinition", "pivotCacheRecords");
                self.pivot_cache_records.retain(|(p, _)| *p != records_path);
            }
            self.workbook_rels.relationships.retain(|r| r.id != *rid);
        }

        // Remove content type overrides for the removed parts.
        let pt_part = format!("/{}", pt_path);
        self.content_types
            .overrides
            .retain(|o| o.part_name != pt_part);

        // Also remove cache def and records content types if the paths were removed.
        self.content_types.overrides.retain(|o| {
            let p = o.part_name.trim_start_matches('/');
            // Keep if it is still in our live lists.
            if o.content_type == mime_types::PIVOT_CACHE_DEFINITION {
                return self.pivot_cache_defs.iter().any(|(path, _)| path == p);
            }
            if o.content_type == mime_types::PIVOT_CACHE_RECORDS {
                return self.pivot_cache_records.iter().any(|(path, _)| path == p);
            }
            if o.content_type == mime_types::PIVOT_TABLE {
                return self.pivot_tables.iter().any(|(path, _)| path == p);
            }
            true
        });

        // Remove worksheet relationship for this pivot table.
        for (_idx, rels) in self.worksheet_rels.iter_mut() {
            rels.relationships.retain(|r| {
                if r.rel_type != rel_types::PIVOT_TABLE {
                    return true;
                }
                // Check if the target matches the removed pivot table.
                let full_target = format!(
                    "xl/pivotTables/{}",
                    r.target.trim_start_matches("../pivotTables/")
                );
                full_target != pt_path
            });
        }

        Ok(())
    }

    /// Add a sparkline to a worksheet.
    pub fn add_sparkline(
        &mut self,
        sheet: &str,
        config: &crate::sparkline::SparklineConfig,
    ) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        crate::sparkline::validate_sparkline_config(config)?;
        while self.sheet_sparklines.len() <= idx {
            self.sheet_sparklines.push(vec![]);
        }
        self.sheet_sparklines[idx].push(config.clone());
        Ok(())
    }

    /// Get all sparklines for a worksheet.
    pub fn get_sparklines(&self, sheet: &str) -> Result<Vec<crate::sparkline::SparklineConfig>> {
        let idx = self.sheet_index(sheet)?;
        Ok(self.sheet_sparklines.get(idx).cloned().unwrap_or_default())
    }

    /// Remove a sparkline by its location cell reference.
    pub fn remove_sparkline(&mut self, sheet: &str, location: &str) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        let sparklines = self
            .sheet_sparklines
            .get_mut(idx)
            .ok_or_else(|| Error::Internal(format!("no sparkline data for sheet '{sheet}'")))?;
        let pos = sparklines
            .iter()
            .position(|s| s.location == location)
            .ok_or_else(|| {
                Error::Internal(format!(
                    "sparkline at location '{location}' not found on sheet '{sheet}'"
                ))
            })?;
        sparklines.remove(pos);
        Ok(())
    }

    /// Resolve a theme color by index (0-11) with optional tint.
    /// Returns the ARGB hex string (e.g. "FF4472C4") or None if the index is out of range.
    pub fn get_theme_color(&self, index: u32, tint: Option<f64>) -> Option<String> {
        crate::theme::resolve_theme_color(&self.theme_colors, index, tint)
    }

    /// Read the header row (first row) of a range from a sheet, returning cell
    /// values as strings.
    fn read_header_row(&self, sheet: &str, range: &str) -> Result<Vec<String>> {
        let parts: Vec<&str> = range.split(':').collect();
        if parts.len() != 2 {
            return Err(Error::InvalidSourceRange(range.to_string()));
        }
        let (start_col, start_row) = cell_name_to_coordinates(parts[0])
            .map_err(|_| Error::InvalidSourceRange(range.to_string()))?;
        let (end_col, _end_row) = cell_name_to_coordinates(parts[1])
            .map_err(|_| Error::InvalidSourceRange(range.to_string()))?;

        let mut headers = Vec::new();
        for col in start_col..=end_col {
            let cell_name = crate::utils::cell_ref::coordinates_to_cell_name(col, start_row)?;
            let val = self.get_cell_value(sheet, &cell_name)?;
            let s = match val {
                CellValue::String(s) => s,
                CellValue::Number(n) => n.to_string(),
                CellValue::Bool(b) => b.to_string(),
                CellValue::RichString(runs) => crate::rich_text::rich_text_to_plain(&runs),
                _ => String::new(),
            };
            headers.push(s);
        }
        Ok(headers)
    }

    /// Find the target sheet name for a pivot table by looking at worksheet
    /// relationships that reference its path.
    fn find_pivot_table_target_sheet(
        &self,
        pt: &sheetkit_xml::pivot_table::PivotTableDefinition,
    ) -> Option<String> {
        // Find the pivot table path.
        let pt_path = self
            .pivot_tables
            .iter()
            .find(|(_, p)| p.name == pt.name)
            .map(|(path, _)| path.as_str())?;

        // Find which worksheet has a relationship pointing to this pivot table.
        for (sheet_idx, rels) in &self.worksheet_rels {
            for r in &rels.relationships {
                if r.rel_type == rel_types::PIVOT_TABLE {
                    let full_target = format!(
                        "xl/pivotTables/{}",
                        r.target.trim_start_matches("../pivotTables/")
                    );
                    if full_target == pt_path {
                        return self
                            .worksheets
                            .get(*sheet_idx)
                            .map(|(name, _)| name.clone());
                    }
                }
            }
        }
        None
    }

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

    /// Return `(col, row)` pairs for all occupied cells on the named sheet.
    pub fn get_occupied_cells(&self, sheet: &str) -> Result<Vec<(u32, u32)>> {
        let ws = self
            .worksheets
            .iter()
            .find(|(name, _)| name == sheet)
            .map(|(_, ws)| ws)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })?;
        let mut cells = Vec::new();
        for row in &ws.sheet_data.rows {
            for cell in &row.cells {
                if let Ok((c, r)) = cell_name_to_coordinates(&cell.r) {
                    cells.push((c, r));
                }
            }
        }
        Ok(cells)
    }

    /// Evaluate a single formula string in the context of `sheet`.
    ///
    /// A [`CellSnapshot`] is built from the current workbook state so
    /// that cell references within the formula can be resolved.
    pub fn evaluate_formula(&self, sheet: &str, formula: &str) -> Result<CellValue> {
        // Validate the sheet exists.
        let _ = self.sheet_index(sheet)?;
        let parsed = crate::formula::parser::parse_formula(formula)?;
        let snapshot = self.build_cell_snapshot(sheet)?;
        crate::formula::eval::evaluate(&parsed, &snapshot)
    }

    /// Recalculate every formula cell across all sheets and store the
    /// computed result back into each cell. Uses a dependency graph and
    /// topological sort so formulas are evaluated after their dependencies.
    pub fn calculate_all(&mut self) -> Result<()> {
        use crate::formula::eval::{build_dependency_graph, topological_sort, CellCoord};

        let sheet_names: Vec<String> = self.sheet_names().iter().map(|s| s.to_string()).collect();

        // Collect all formula cells with their coordinates and formula strings.
        let mut formula_cells: Vec<(CellCoord, String)> = Vec::new();
        for sn in &sheet_names {
            let ws = self
                .worksheets
                .iter()
                .find(|(name, _)| name == sn)
                .map(|(_, ws)| ws)
                .ok_or_else(|| Error::SheetNotFound {
                    name: sn.to_string(),
                })?;
            for row in &ws.sheet_data.rows {
                for cell in &row.cells {
                    if let Some(ref f) = cell.f {
                        let formula_str = f.value.clone().unwrap_or_default();
                        if !formula_str.is_empty() {
                            if let Ok((c, r)) = cell_name_to_coordinates(&cell.r) {
                                formula_cells.push((
                                    CellCoord {
                                        sheet: sn.clone(),
                                        col: c,
                                        row: r,
                                    },
                                    formula_str,
                                ));
                            }
                        }
                    }
                }
            }
        }

        if formula_cells.is_empty() {
            return Ok(());
        }

        // Build dependency graph and determine evaluation order.
        let deps = build_dependency_graph(&formula_cells)?;
        let coords: Vec<CellCoord> = formula_cells.iter().map(|(c, _)| c.clone()).collect();
        let eval_order = topological_sort(&coords, &deps)?;

        // Build a lookup from coord to formula string.
        let formula_map: HashMap<CellCoord, String> = formula_cells.into_iter().collect();

        // Build a snapshot of all cell data.
        let first_sheet = sheet_names.first().cloned().unwrap_or_default();
        let mut snapshot = self.build_cell_snapshot(&first_sheet)?;

        // Evaluate in dependency order, updating the snapshot progressively
        // so later formulas see already-computed results.
        let mut results: Vec<(CellCoord, String, CellValue)> = Vec::new();
        for coord in &eval_order {
            if let Some(formula_str) = formula_map.get(coord) {
                snapshot.set_current_sheet(&coord.sheet);
                let parsed = crate::formula::parser::parse_formula(formula_str)?;
                let mut evaluator = crate::formula::eval::Evaluator::new(&snapshot);
                let result = evaluator.eval_expr(&parsed)?;
                snapshot.set_cell(&coord.sheet, coord.col, coord.row, result.clone());
                results.push((coord.clone(), formula_str.clone(), result));
            }
        }

        // Write results back directly to the XML cells, preserving the
        // formula element and storing the computed value in the v/t fields.
        for (coord, _formula_str, result) in results {
            let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(coord.col, coord.row)?;
            if let Some((_, ws)) = self.worksheets.iter_mut().find(|(n, _)| *n == coord.sheet) {
                if let Some(row) = ws.sheet_data.rows.iter_mut().find(|r| r.r == coord.row) {
                    if let Some(cell) = row.cells.iter_mut().find(|c| c.r == cell_ref) {
                        match &result {
                            CellValue::Number(n) => {
                                cell.v = Some(n.to_string());
                                cell.t = None;
                            }
                            CellValue::String(s) => {
                                cell.v = Some(s.clone());
                                cell.t = Some("str".to_string());
                            }
                            CellValue::Bool(b) => {
                                cell.v = Some(if *b { "1".to_string() } else { "0".to_string() });
                                cell.t = Some("b".to_string());
                            }
                            CellValue::Error(e) => {
                                cell.v = Some(e.clone());
                                cell.t = Some("e".to_string());
                            }
                            CellValue::Date(n) => {
                                cell.v = Some(n.to_string());
                                cell.t = None;
                            }
                            _ => {}
                        }
                    }
                }
            }
        }

        Ok(())
    }

    /// Build a [`CellSnapshot`] for formula evaluation, with the given
    /// sheet as the current-sheet context.
    fn build_cell_snapshot(
        &self,
        current_sheet: &str,
    ) -> Result<crate::formula::eval::CellSnapshot> {
        let mut snapshot = crate::formula::eval::CellSnapshot::new(current_sheet.to_string());
        for (sn, ws) in &self.worksheets {
            for row in &ws.sheet_data.rows {
                for cell in &row.cells {
                    if let Ok((c, r)) = cell_name_to_coordinates(&cell.r) {
                        let cv = self.xml_cell_to_value(cell)?;
                        snapshot.set_cell(sn, c, r, cv);
                    }
                }
            }
        }
        Ok(snapshot)
    }

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

    /// Resolve the part path for a sheet index from workbook relationships.
    /// Falls back to the default `xl/worksheets/sheet{N}.xml` naming.
    fn sheet_part_path(&self, sheet_idx: usize) -> String {
        if let Some(sheet_entry) = self.workbook_xml.sheets.sheets.get(sheet_idx) {
            if let Some(rel) = self
                .workbook_rels
                .relationships
                .iter()
                .find(|r| r.id == sheet_entry.r_id && r.rel_type == rel_types::WORKSHEET)
            {
                return resolve_relationship_target("xl/workbook.xml", &rel.target);
            }
        }
        format!("xl/worksheets/sheet{}.xml", sheet_idx + 1)
    }

    /// Reindex per-sheet maps after deleting a sheet.
    fn reindex_sheet_maps_after_delete(&mut self, removed_idx: usize) {
        self.worksheet_rels = self
            .worksheet_rels
            .iter()
            .filter_map(|(idx, rels)| {
                if *idx == removed_idx {
                    None
                } else if *idx > removed_idx {
                    Some((idx - 1, rels.clone()))
                } else {
                    Some((*idx, rels.clone()))
                }
            })
            .collect();

        self.worksheet_drawings = self
            .worksheet_drawings
            .iter()
            .filter_map(|(idx, drawing_idx)| {
                if *idx == removed_idx {
                    None
                } else if *idx > removed_idx {
                    Some((idx - 1, *drawing_idx))
                } else {
                    Some((*idx, *drawing_idx))
                }
            })
            .collect();
    }

    /// Apply a cell-reference shift transformation to sheet-scoped structures.
    fn apply_reference_shift_for_sheet<F>(&mut self, sheet_idx: usize, shift_cell: F) -> Result<()>
    where
        F: Fn(u32, u32) -> (u32, u32) + Copy,
    {
        {
            let ws = &mut self.worksheets[sheet_idx].1;

            // Cell formulas.
            for row in &mut ws.sheet_data.rows {
                for cell in &mut row.cells {
                    if let Some(ref mut f) = cell.f {
                        if let Some(ref mut expr) = f.value {
                            *expr = shift_cell_references_in_text(expr, shift_cell)?;
                        }
                    }
                }
            }

            // Merged ranges.
            if let Some(ref mut merges) = ws.merge_cells {
                for mc in &mut merges.merge_cells {
                    mc.reference = shift_cell_references_in_text(&mc.reference, shift_cell)?;
                }
            }

            // Auto-filter.
            if let Some(ref mut af) = ws.auto_filter {
                af.reference = shift_cell_references_in_text(&af.reference, shift_cell)?;
            }

            // Data validations.
            if let Some(ref mut dvs) = ws.data_validations {
                for dv in &mut dvs.data_validations {
                    dv.sqref = shift_cell_references_in_text(&dv.sqref, shift_cell)?;
                    if let Some(ref mut f1) = dv.formula1 {
                        *f1 = shift_cell_references_in_text(f1, shift_cell)?;
                    }
                    if let Some(ref mut f2) = dv.formula2 {
                        *f2 = shift_cell_references_in_text(f2, shift_cell)?;
                    }
                }
            }

            // Conditional formatting ranges/formulas.
            for cf in &mut ws.conditional_formatting {
                cf.sqref = shift_cell_references_in_text(&cf.sqref, shift_cell)?;
                for rule in &mut cf.cf_rules {
                    for f in &mut rule.formulas {
                        *f = shift_cell_references_in_text(f, shift_cell)?;
                    }
                }
            }

            // Hyperlinks.
            if let Some(ref mut hyperlinks) = ws.hyperlinks {
                for hl in &mut hyperlinks.hyperlinks {
                    hl.reference = shift_cell_references_in_text(&hl.reference, shift_cell)?;
                    if let Some(ref mut loc) = hl.location {
                        *loc = shift_cell_references_in_text(loc, shift_cell)?;
                    }
                }
            }

            // Pane/selection references.
            if let Some(ref mut views) = ws.sheet_views {
                for view in &mut views.sheet_views {
                    if let Some(ref mut pane) = view.pane {
                        if let Some(ref mut top_left) = pane.top_left_cell {
                            *top_left = shift_cell_references_in_text(top_left, shift_cell)?;
                        }
                    }
                    for sel in &mut view.selection {
                        if let Some(ref mut ac) = sel.active_cell {
                            *ac = shift_cell_references_in_text(ac, shift_cell)?;
                        }
                        if let Some(ref mut sqref) = sel.sqref {
                            *sqref = shift_cell_references_in_text(sqref, shift_cell)?;
                        }
                    }
                }
            }
        }

        // Drawing anchors attached to this sheet.
        if let Some(&drawing_idx) = self.worksheet_drawings.get(&sheet_idx) {
            if let Some((_, drawing)) = self.drawings.get_mut(drawing_idx) {
                for anchor in &mut drawing.one_cell_anchors {
                    let (new_col, new_row) = shift_cell(anchor.from.col + 1, anchor.from.row + 1);
                    anchor.from.col = new_col - 1;
                    anchor.from.row = new_row - 1;
                }
                for anchor in &mut drawing.two_cell_anchors {
                    let (from_col, from_row) = shift_cell(anchor.from.col + 1, anchor.from.row + 1);
                    anchor.from.col = from_col - 1;
                    anchor.from.row = from_row - 1;
                    let (to_col, to_row) = shift_cell(anchor.to.col + 1, anchor.to.row + 1);
                    anchor.to.col = to_col - 1;
                    anchor.to.row = to_row - 1;
                }
            }
        }

        Ok(())
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
            // Number (explicit or default type) -- may be a date if styled.
            (None | Some("n"), Some(v)) => {
                let n: f64 = v
                    .parse()
                    .map_err(|_| Error::Internal(format!("invalid number: {v}")))?;
                // Check whether this cell has a date number format.
                if self.is_date_styled_cell(xml_cell) {
                    return Ok(CellValue::Date(n));
                }
                Ok(CellValue::Number(n))
            }
            // No value
            _ => Ok(CellValue::Empty),
        }
    }

    /// Check whether a cell's style indicates a date/time number format.
    fn is_date_styled_cell(&self, xml_cell: &Cell) -> bool {
        let style_idx = match xml_cell.s {
            Some(idx) => idx as usize,
            None => return false,
        };
        let xf = match self.stylesheet.cell_xfs.xfs.get(style_idx) {
            Some(xf) => xf,
            None => return false,
        };
        let num_fmt_id = xf.num_fmt_id.unwrap_or(0);
        // Check built-in date format IDs.
        if crate::cell::is_date_num_fmt(num_fmt_id) {
            return true;
        }
        // Check custom number formats for date patterns.
        if num_fmt_id >= 164 {
            if let Some(ref num_fmts) = self.stylesheet.num_fmts {
                if let Some(nf) = num_fmts
                    .num_fmts
                    .iter()
                    .find(|nf| nf.num_fmt_id == num_fmt_id)
                {
                    return crate::cell::is_date_format_code(&nf.format_code);
                }
            }
        }
        false
    }

    /// Set the core document properties (title, author, etc.).
    pub fn set_doc_props(&mut self, props: crate::doc_props::DocProperties) {
        self.core_properties = Some(props.to_core_properties());
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

    /// Set a cell to a rich text value (multiple formatted runs).
    pub fn set_cell_rich_text(
        &mut self,
        sheet: &str,
        cell: &str,
        runs: Vec<crate::rich_text::RichTextRun>,
    ) -> Result<()> {
        self.set_cell_value(sheet, cell, CellValue::RichString(runs))
    }

    /// Get rich text runs for a cell, if it contains rich text.
    ///
    /// Returns `None` if the cell is empty, contains a plain string, or holds
    /// a non-string value.
    pub fn get_cell_rich_text(
        &self,
        sheet: &str,
        cell: &str,
    ) -> Result<Option<Vec<crate::rich_text::RichTextRun>>> {
        let (col, row) = cell_name_to_coordinates(cell)?;
        let sheet_idx = self
            .worksheets
            .iter()
            .position(|(name, _)| name == sheet)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })?;
        let ws = &self.worksheets[sheet_idx].1;

        for xml_row in &ws.sheet_data.rows {
            if xml_row.r == row {
                for xml_cell in &xml_row.cells {
                    let (cc, cr) = cell_name_to_coordinates(&xml_cell.r)?;
                    if cc == col && cr == row {
                        if xml_cell.t.as_deref() == Some("s") {
                            if let Some(ref v) = xml_cell.v {
                                if let Ok(idx) = v.parse::<usize>() {
                                    return Ok(self.sst_runtime.get_rich_text(idx));
                                }
                            }
                        }
                        return Ok(None);
                    }
                }
            }
        }
        Ok(None)
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

    /// Add or update a defined name in the workbook.
    ///
    /// If `scope` is `None`, the name is workbook-scoped (visible from all sheets).
    /// If `scope` is `Some(sheet_name)`, it is sheet-scoped using the sheet's 0-based index.
    /// If a name with the same name and scope already exists, its value and comment are updated.
    pub fn set_defined_name(
        &mut self,
        name: &str,
        value: &str,
        scope: Option<&str>,
        comment: Option<&str>,
    ) -> Result<()> {
        let dn_scope = self.resolve_defined_name_scope(scope)?;
        crate::defined_names::set_defined_name(
            &mut self.workbook_xml,
            name,
            value,
            dn_scope,
            comment,
        )
    }

    /// Get a defined name by name and scope.
    ///
    /// If `scope` is `None`, looks for a workbook-scoped name.
    /// If `scope` is `Some(sheet_name)`, looks for a sheet-scoped name.
    /// Returns `None` if no matching defined name is found.
    pub fn get_defined_name(
        &self,
        name: &str,
        scope: Option<&str>,
    ) -> Result<Option<crate::defined_names::DefinedNameInfo>> {
        let dn_scope = self.resolve_defined_name_scope(scope)?;
        Ok(crate::defined_names::get_defined_name(
            &self.workbook_xml,
            name,
            dn_scope,
        ))
    }

    /// List all defined names in the workbook.
    pub fn get_all_defined_names(&self) -> Vec<crate::defined_names::DefinedNameInfo> {
        crate::defined_names::get_all_defined_names(&self.workbook_xml)
    }

    /// Delete a defined name by name and scope.
    ///
    /// Returns an error if the name does not exist for the given scope.
    pub fn delete_defined_name(&mut self, name: &str, scope: Option<&str>) -> Result<()> {
        let dn_scope = self.resolve_defined_name_scope(scope)?;
        crate::defined_names::delete_defined_name(&mut self.workbook_xml, name, dn_scope)
    }

    /// Protect a sheet with optional password and permission settings.
    ///
    /// Delegates to [`crate::sheet::protect_sheet`] after looking up the sheet.
    pub fn protect_sheet(
        &mut self,
        sheet: &str,
        config: &crate::sheet::SheetProtectionConfig,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::sheet::protect_sheet(ws, config)
    }

    /// Remove sheet protection.
    pub fn unprotect_sheet(&mut self, sheet: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::sheet::unprotect_sheet(ws)
    }

    /// Check if a sheet is protected.
    pub fn is_sheet_protected(&self, sheet: &str) -> Result<bool> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::sheet::is_sheet_protected(ws))
    }

    /// Resolve an optional sheet name to a [`DefinedNameScope`](crate::defined_names::DefinedNameScope).
    fn resolve_defined_name_scope(
        &self,
        scope: Option<&str>,
    ) -> Result<crate::defined_names::DefinedNameScope> {
        match scope {
            None => Ok(crate::defined_names::DefinedNameScope::Workbook),
            Some(sheet_name) => {
                let idx = self.sheet_index(sheet_name)?;
                Ok(crate::defined_names::DefinedNameScope::Sheet(idx as u32))
            }
        }
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

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

/// Read a ZIP entry as raw bytes.
fn read_bytes_part(archive: &mut zip::ZipArchive<std::fs::File>, name: &str) -> Result<Vec<u8>> {
    let mut entry = archive
        .by_name(name)
        .map_err(|e| Error::Zip(e.to_string()))?;
    let mut content = Vec::new();
    entry
        .read_to_end(&mut content)
        .map_err(|e| Error::Zip(e.to_string()))?;
    Ok(content)
}

/// Serialize a worksheet with sparkline extension list appended.
fn serialize_worksheet_with_sparklines(
    ws: &WorksheetXml,
    sparklines: &[crate::sparkline::SparklineConfig],
) -> Result<String> {
    let body = quick_xml::se::to_string(ws).map_err(|e| Error::XmlParse(e.to_string()))?;

    let closing = "</worksheet>";
    let ext_xml = build_sparkline_ext_xml(sparklines);
    if let Some(pos) = body.rfind(closing) {
        let mut result =
            String::with_capacity(XML_DECLARATION.len() + 1 + body.len() + ext_xml.len());
        result.push_str(XML_DECLARATION);
        result.push('\n');
        result.push_str(&body[..pos]);
        result.push_str(&ext_xml);
        result.push_str(closing);
        Ok(result)
    } else {
        Ok(format!("{XML_DECLARATION}\n{body}"))
    }
}

/// Build the extLst XML block for sparklines using manual string construction.
fn build_sparkline_ext_xml(sparklines: &[crate::sparkline::SparklineConfig]) -> String {
    use std::fmt::Write;
    let mut xml = String::new();
    let _ = write!(
        xml,
        "<extLst>\
         <ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" \
         uri=\"{{05C60535-1F16-4fd2-B633-F4F36F0B64E0}}\">\
         <x14:sparklineGroups \
         xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
    );
    for config in sparklines {
        let group = crate::sparkline::config_to_xml_group(config);
        let _ = write!(xml, "<x14:sparklineGroup");
        if let Some(ref t) = group.sparkline_type {
            let _ = write!(xml, " type=\"{t}\"");
        }
        if group.markers == Some(true) {
            let _ = write!(xml, " markers=\"1\"");
        }
        if group.high == Some(true) {
            let _ = write!(xml, " high=\"1\"");
        }
        if group.low == Some(true) {
            let _ = write!(xml, " low=\"1\"");
        }
        if group.first == Some(true) {
            let _ = write!(xml, " first=\"1\"");
        }
        if group.last == Some(true) {
            let _ = write!(xml, " last=\"1\"");
        }
        if group.negative == Some(true) {
            let _ = write!(xml, " negative=\"1\"");
        }
        if group.display_x_axis == Some(true) {
            let _ = write!(xml, " displayXAxis=\"1\"");
        }
        if let Some(w) = group.line_weight {
            let _ = write!(xml, " lineWeight=\"{w}\"");
        }
        let _ = write!(xml, "><x14:sparklines>");
        for sp in &group.sparklines.items {
            let _ = write!(
                xml,
                "<x14:sparkline><xm:f>{}</xm:f><xm:sqref>{}</xm:sqref></x14:sparkline>",
                sp.formula, sp.sqref
            );
        }
        let _ = write!(xml, "</x14:sparklines></x14:sparklineGroup>");
    }
    let _ = write!(xml, "</x14:sparklineGroups></ext></extLst>");
    xml
}

/// Parse sparkline configurations from raw worksheet XML content.
fn parse_sparklines_from_xml(xml: &str) -> Vec<crate::sparkline::SparklineConfig> {
    use crate::sparkline::{SparklineConfig, SparklineType};

    let mut sparklines = Vec::new();

    // Find all sparklineGroup elements and parse their attributes and children.
    let mut search_from = 0;
    while let Some(group_start) = xml[search_from..].find("<x14:sparklineGroup") {
        let abs_start = search_from + group_start;
        let group_end_tag = "</x14:sparklineGroup>";
        let abs_end = match xml[abs_start..].find(group_end_tag) {
            Some(pos) => abs_start + pos + group_end_tag.len(),
            None => break,
        };
        let group_xml = &xml[abs_start..abs_end];

        // Parse group-level attributes.
        let sparkline_type = extract_xml_attr(group_xml, "type")
            .and_then(|s| SparklineType::parse(&s))
            .unwrap_or_default();
        let markers = extract_xml_bool_attr(group_xml, "markers");
        let high_point = extract_xml_bool_attr(group_xml, "high");
        let low_point = extract_xml_bool_attr(group_xml, "low");
        let first_point = extract_xml_bool_attr(group_xml, "first");
        let last_point = extract_xml_bool_attr(group_xml, "last");
        let negative_points = extract_xml_bool_attr(group_xml, "negative");
        let show_axis = extract_xml_bool_attr(group_xml, "displayXAxis");
        let line_weight =
            extract_xml_attr(group_xml, "lineWeight").and_then(|s| s.parse::<f64>().ok());

        // Parse individual sparkline entries within this group.
        let mut sp_from = 0;
        while let Some(sp_start) = group_xml[sp_from..].find("<x14:sparkline>") {
            let sp_abs = sp_from + sp_start;
            let sp_end_tag = "</x14:sparkline>";
            let sp_abs_end = match group_xml[sp_abs..].find(sp_end_tag) {
                Some(pos) => sp_abs + pos + sp_end_tag.len(),
                None => break,
            };
            let sp_xml = &group_xml[sp_abs..sp_abs_end];

            let formula = extract_xml_element(sp_xml, "xm:f").unwrap_or_default();
            let sqref = extract_xml_element(sp_xml, "xm:sqref").unwrap_or_default();

            if !formula.is_empty() && !sqref.is_empty() {
                sparklines.push(SparklineConfig {
                    data_range: formula,
                    location: sqref,
                    sparkline_type: sparkline_type.clone(),
                    markers,
                    high_point,
                    low_point,
                    first_point,
                    last_point,
                    negative_points,
                    show_axis,
                    line_weight,
                    style: None,
                });
            }
            sp_from = sp_abs_end;
        }
        search_from = abs_end;
    }
    sparklines
}

/// Extract an XML attribute value from an element's opening tag.
fn extract_xml_attr(xml: &str, attr: &str) -> Option<String> {
    // Look for attr="value" or attr='value' patterns.
    let patterns = [format!(" {attr}=\""), format!(" {attr}='")];
    for pat in &patterns {
        if let Some(start) = xml.find(pat.as_str()) {
            let val_start = start + pat.len();
            let quote = pat.chars().last().unwrap();
            if let Some(end) = xml[val_start..].find(quote) {
                return Some(xml[val_start..val_start + end].to_string());
            }
        }
    }
    None
}

/// Extract a boolean attribute from an XML element (true for "1" or "true").
fn extract_xml_bool_attr(xml: &str, attr: &str) -> bool {
    extract_xml_attr(xml, attr)
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        .unwrap_or(false)
}

/// Extract the text content of an XML element like `<tag>content</tag>`.
fn extract_xml_element(xml: &str, tag: &str) -> Option<String> {
    let open = format!("<{tag}>");
    let close = format!("</{tag}>");
    let start = xml.find(&open)?;
    let content_start = start + open.len();
    let end = xml[content_start..].find(&close)?;
    Some(xml[content_start..content_start + end].to_string())
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
        CellValue::Date(serial) => {
            // Dates are stored as numbers in Excel. The style must apply a
            // date number format for correct display.
            xml_cell.v = Some(serial.to_string());
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
        CellValue::RichString(runs) => {
            let idx = sst.add_rich_text(&runs);
            xml_cell.t = Some("s".to_string());
            xml_cell.v = Some(idx.to_string());
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
    fn test_set_and_get_date_value() {
        use crate::style::{builtin_num_fmts, NumFmtStyle, Style};

        let mut wb = Workbook::new();
        // Create a date style.
        let style_id = wb
            .add_style(&Style {
                num_fmt: Some(NumFmtStyle::Builtin(builtin_num_fmts::DATE_MDY)),
                ..Style::default()
            })
            .unwrap();

        // Set a date value.
        let date_serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 6, 15).unwrap());
        wb.set_cell_value("Sheet1", "A1", CellValue::Date(date_serial))
            .unwrap();
        wb.set_cell_style("Sheet1", "A1", style_id).unwrap();

        // Get the value back -- it should be Date because the cell has a date style.
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::Date(date_serial));
    }

    #[test]
    fn test_date_value_without_style_returns_number() {
        let mut wb = Workbook::new();
        // Set a date value without a date style.
        let date_serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 6, 15).unwrap());
        wb.set_cell_value("Sheet1", "A1", CellValue::Date(date_serial))
            .unwrap();

        // Without a date style, the value is read back as Number.
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::Number(date_serial));
    }

    #[test]
    fn test_date_value_roundtrip_through_save() {
        use crate::style::{builtin_num_fmts, NumFmtStyle, Style};

        let mut wb = Workbook::new();
        let style_id = wb
            .add_style(&Style {
                num_fmt: Some(NumFmtStyle::Builtin(builtin_num_fmts::DATETIME)),
                ..Style::default()
            })
            .unwrap();

        let dt = chrono::NaiveDate::from_ymd_opt(2024, 3, 15)
            .unwrap()
            .and_hms_opt(14, 30, 0)
            .unwrap();
        let serial = crate::cell::datetime_to_serial(dt);
        wb.set_cell_value("Sheet1", "A1", CellValue::Date(serial))
            .unwrap();
        wb.set_cell_style("Sheet1", "A1", style_id).unwrap();

        let tmp = tempfile::NamedTempFile::new().unwrap();
        let path = tmp.path().to_str().unwrap();
        wb.save(path).unwrap();

        let wb2 = Workbook::open(path).unwrap();
        let val = wb2.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::Date(serial));
    }

    #[test]
    fn test_date_from_naive_date_conversion() {
        let date = chrono::NaiveDate::from_ymd_opt(2024, 1, 1).unwrap();
        let cv: CellValue = date.into();
        match cv {
            CellValue::Date(s) => {
                let roundtripped = crate::cell::serial_to_date(s).unwrap();
                assert_eq!(roundtripped, date);
            }
            _ => panic!("expected Date variant"),
        }
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
                gradient: None,
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
    fn test_workbook_insert_rows_updates_formula_and_ranges() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "C1",
            CellValue::Formula {
                expr: "SUM(A2:B2)".to_string(),
                result: None,
            },
        )
        .unwrap();
        wb.add_data_validation(
            "Sheet1",
            &crate::validation::DataValidationConfig::whole_number("A2:A5", 1, 9),
        )
        .unwrap();
        wb.set_auto_filter("Sheet1", "A2:B10").unwrap();
        wb.merge_cells("Sheet1", "A2", "B3").unwrap();

        wb.insert_rows("Sheet1", 2, 1).unwrap();

        match wb.get_cell_value("Sheet1", "C1").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "SUM(A3:B3)"),
            other => panic!("expected formula, got {other:?}"),
        }

        let validations = wb.get_data_validations("Sheet1").unwrap();
        assert_eq!(validations.len(), 1);
        assert_eq!(validations[0].sqref, "A3:A6");

        let merges = wb.get_merge_cells("Sheet1").unwrap();
        assert_eq!(merges, vec!["A3:B4".to_string()]);

        let ws = wb.worksheet_ref("Sheet1").unwrap();
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A3:B11");
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
    fn test_workbook_insert_cols_updates_formula_and_ranges() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "D1",
            CellValue::Formula {
                expr: "SUM(A1:B1)".to_string(),
                result: None,
            },
        )
        .unwrap();
        wb.add_data_validation(
            "Sheet1",
            &crate::validation::DataValidationConfig::whole_number("B2:C3", 1, 9),
        )
        .unwrap();
        wb.set_auto_filter("Sheet1", "A1:C10").unwrap();
        wb.merge_cells("Sheet1", "B3", "C4").unwrap();

        wb.insert_cols("Sheet1", "B", 2).unwrap();

        match wb.get_cell_value("Sheet1", "F1").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "SUM(A1:D1)"),
            other => panic!("expected formula, got {other:?}"),
        }

        let validations = wb.get_data_validations("Sheet1").unwrap();
        assert_eq!(validations.len(), 1);
        assert_eq!(validations[0].sqref, "D2:E3");

        let merges = wb.get_merge_cells("Sheet1").unwrap();
        assert_eq!(merges, vec!["D3:E4".to_string()]);

        let ws = wb.worksheet_ref("Sheet1").unwrap();
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:E10");
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

    #[test]
    fn test_workbook_comment_roundtrip_open() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("comment_roundtrip_open.xlsx");

        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "Persist me".to_string(),
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let comments = wb2.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell, "A1");
        assert_eq!(comments[0].author, "Author");
        assert_eq!(comments[0].text, "Persist me");
    }

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
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
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
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
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
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
        };
        let config2 = ChartConfig {
            chart_type: ChartType::Line,
            title: Some("Chart 2".to_string()),
            series: vec![ChartSeries {
                name: "S2".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$C$1:$C$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
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
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
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
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
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
    fn test_save_with_new_image_formats() {
        use crate::image::{ImageConfig, ImageFormat};

        let formats = [
            (ImageFormat::Bmp, "bmp"),
            (ImageFormat::Ico, "ico"),
            (ImageFormat::Tiff, "tiff"),
            (ImageFormat::Svg, "svg"),
            (ImageFormat::Emf, "emf"),
            (ImageFormat::Emz, "emz"),
            (ImageFormat::Wmf, "wmf"),
            (ImageFormat::Wmz, "wmz"),
        ];

        for (format, ext) in &formats {
            let dir = TempDir::new().unwrap();
            let path = dir.path().join(format!("with_{ext}.xlsx"));

            let mut wb = Workbook::new();
            let config = ImageConfig {
                data: vec![0x00, 0x01, 0x02, 0x03],
                format: format.clone(),
                from_cell: "A1".to_string(),
                width_px: 100,
                height_px: 100,
            };
            wb.add_image("Sheet1", &config).unwrap();
            wb.save(&path).unwrap();

            let file = std::fs::File::open(&path).unwrap();
            let mut archive = zip::ZipArchive::new(file).unwrap();
            let media_path = format!("xl/media/image1.{ext}");
            assert!(
                archive.by_name(&media_path).is_ok(),
                "expected {media_path} in archive for format {ext}"
            );
        }
    }

    #[test]
    fn test_add_image_new_format_content_type_default() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x42, 0x4D],
            format: ImageFormat::Bmp,
            from_cell: "A1".to_string(),
            width_px: 100,
            height_px: 100,
        };
        wb.add_image("Sheet1", &config).unwrap();

        let has_bmp_default = wb
            .content_types
            .defaults
            .iter()
            .any(|d| d.extension == "bmp" && d.content_type == "image/bmp");
        assert!(has_bmp_default, "content types should have bmp default");
    }

    #[test]
    fn test_add_image_svg_content_type_default() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x3C, 0x73, 0x76, 0x67],
            format: ImageFormat::Svg,
            from_cell: "B3".to_string(),
            width_px: 200,
            height_px: 200,
        };
        wb.add_image("Sheet1", &config).unwrap();

        let has_svg_default = wb
            .content_types
            .defaults
            .iter()
            .any(|d| d.extension == "svg" && d.content_type == "image/svg+xml");
        assert!(has_svg_default, "content types should have svg default");
    }

    #[test]
    fn test_add_image_emf_content_type_and_path() {
        use crate::image::{ImageConfig, ImageFormat};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("with_emf.xlsx");

        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x01, 0x00, 0x00, 0x00],
            format: ImageFormat::Emf,
            from_cell: "A1".to_string(),
            width_px: 150,
            height_px: 150,
        };
        wb.add_image("Sheet1", &config).unwrap();

        let has_emf_default = wb
            .content_types
            .defaults
            .iter()
            .any(|d| d.extension == "emf" && d.content_type == "image/x-emf");
        assert!(has_emf_default);

        wb.save(&path).unwrap();
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/media/image1.emf").is_ok());
    }

    #[test]
    fn test_add_multiple_new_format_images() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();

        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x42, 0x4D],
                format: ImageFormat::Bmp,
                from_cell: "A1".to_string(),
                width_px: 100,
                height_px: 100,
            },
        )
        .unwrap();

        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x3C, 0x73],
                format: ImageFormat::Svg,
                from_cell: "C1".to_string(),
                width_px: 100,
                height_px: 100,
            },
        )
        .unwrap();

        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x01, 0x00],
                format: ImageFormat::Wmf,
                from_cell: "E1".to_string(),
                width_px: 100,
                height_px: 100,
            },
        )
        .unwrap();

        assert_eq!(wb.images.len(), 3);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 3);

        let ext_defaults: Vec<&str> = wb
            .content_types
            .defaults
            .iter()
            .map(|d| d.extension.as_str())
            .collect();
        assert!(ext_defaults.contains(&"bmp"));
        assert!(ext_defaults.contains(&"svg"));
        assert!(ext_defaults.contains(&"wmf"));
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
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
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
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "A1", "F10", &config).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let ws = wb2.worksheet_ref("Sheet1").unwrap();
        assert!(ws.drawing.is_some());
    }

    #[test]
    fn test_open_save_preserves_existing_drawing_chart_and_image_parts() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        use crate::image::{ImageConfig, ImageFormat};
        let dir = TempDir::new().unwrap();
        let path1 = dir.path().join("source_with_parts.xlsx");
        let path2 = dir.path().join("resaved_with_parts.xlsx");

        let mut wb = Workbook::new();
        wb.add_chart(
            "Sheet1",
            "E1",
            "L10",
            &ChartConfig {
                chart_type: ChartType::Col,
                title: Some("Chart".to_string()),
                series: vec![ChartSeries {
                    name: "Series 1".to_string(),
                    categories: "Sheet1!$A$1:$A$3".to_string(),
                    values: "Sheet1!$B$1:$B$3".to_string(),
                    x_values: None,
                    bubble_sizes: None,
                }],
                show_legend: true,
                view_3d: None,
            },
        )
        .unwrap();
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x89, 0x50, 0x4E, 0x47],
                format: ImageFormat::Png,
                from_cell: "E12".to_string(),
                width_px: 120,
                height_px: 80,
            },
        )
        .unwrap();
        wb.save(&path1).unwrap();

        let wb2 = Workbook::open(&path1).unwrap();
        assert_eq!(wb2.charts.len() + wb2.raw_charts.len(), 1);
        assert_eq!(wb2.drawings.len(), 1);
        assert_eq!(wb2.images.len(), 1);
        assert_eq!(wb2.drawing_rels.len(), 1);
        assert_eq!(wb2.worksheet_drawings.len(), 1);

        wb2.save(&path2).unwrap();

        let file = std::fs::File::open(&path2).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/charts/chart1.xml").is_ok());
        assert!(archive.by_name("xl/drawings/drawing1.xml").is_ok());
        assert!(archive.by_name("xl/media/image1.png").is_ok());
        assert!(archive
            .by_name("xl/worksheets/_rels/sheet1.xml.rels")
            .is_ok());
        assert!(archive
            .by_name("xl/drawings/_rels/drawing1.xml.rels")
            .is_ok());
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

    #[test]
    fn test_set_and_get_external_hyperlink() {
        use crate::hyperlink::HyperlinkType;

        let mut wb = Workbook::new();
        wb.set_cell_hyperlink(
            "Sheet1",
            "A1",
            HyperlinkType::External("https://example.com".to_string()),
            Some("Example"),
            Some("Visit Example"),
        )
        .unwrap();

        let info = wb.get_cell_hyperlink("Sheet1", "A1").unwrap().unwrap();
        assert_eq!(
            info.link_type,
            HyperlinkType::External("https://example.com".to_string())
        );
        assert_eq!(info.display, Some("Example".to_string()));
        assert_eq!(info.tooltip, Some("Visit Example".to_string()));
    }

    #[test]
    fn test_set_and_get_internal_hyperlink() {
        use crate::hyperlink::HyperlinkType;

        let mut wb = Workbook::new();
        wb.new_sheet("Data").unwrap();
        wb.set_cell_hyperlink(
            "Sheet1",
            "B2",
            HyperlinkType::Internal("Data!A1".to_string()),
            Some("Go to Data"),
            None,
        )
        .unwrap();

        let info = wb.get_cell_hyperlink("Sheet1", "B2").unwrap().unwrap();
        assert_eq!(
            info.link_type,
            HyperlinkType::Internal("Data!A1".to_string())
        );
        assert_eq!(info.display, Some("Go to Data".to_string()));
    }

    #[test]
    fn test_set_and_get_email_hyperlink() {
        use crate::hyperlink::HyperlinkType;

        let mut wb = Workbook::new();
        wb.set_cell_hyperlink(
            "Sheet1",
            "C3",
            HyperlinkType::Email("mailto:user@example.com".to_string()),
            None,
            None,
        )
        .unwrap();

        let info = wb.get_cell_hyperlink("Sheet1", "C3").unwrap().unwrap();
        assert_eq!(
            info.link_type,
            HyperlinkType::Email("mailto:user@example.com".to_string())
        );
    }

    #[test]
    fn test_delete_hyperlink_via_workbook() {
        use crate::hyperlink::HyperlinkType;

        let mut wb = Workbook::new();
        wb.set_cell_hyperlink(
            "Sheet1",
            "A1",
            HyperlinkType::External("https://example.com".to_string()),
            None,
            None,
        )
        .unwrap();

        wb.delete_cell_hyperlink("Sheet1", "A1").unwrap();

        let info = wb.get_cell_hyperlink("Sheet1", "A1").unwrap();
        assert!(info.is_none());
    }

    #[test]
    fn test_hyperlink_roundtrip_save_open() {
        use crate::hyperlink::HyperlinkType;

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("hyperlink.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_hyperlink(
            "Sheet1",
            "A1",
            HyperlinkType::External("https://rust-lang.org".to_string()),
            Some("Rust"),
            Some("Rust Homepage"),
        )
        .unwrap();
        wb.set_cell_hyperlink(
            "Sheet1",
            "B1",
            HyperlinkType::Internal("Sheet1!C1".to_string()),
            Some("Go to C1"),
            None,
        )
        .unwrap();
        wb.set_cell_hyperlink(
            "Sheet1",
            "C1",
            HyperlinkType::Email("mailto:hello@example.com".to_string()),
            Some("Email"),
            None,
        )
        .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();

        // External link roundtrip.
        let a1 = wb2.get_cell_hyperlink("Sheet1", "A1").unwrap().unwrap();
        assert_eq!(
            a1.link_type,
            HyperlinkType::External("https://rust-lang.org".to_string())
        );
        assert_eq!(a1.display, Some("Rust".to_string()));
        assert_eq!(a1.tooltip, Some("Rust Homepage".to_string()));

        // Internal link roundtrip.
        let b1 = wb2.get_cell_hyperlink("Sheet1", "B1").unwrap().unwrap();
        assert_eq!(
            b1.link_type,
            HyperlinkType::Internal("Sheet1!C1".to_string())
        );
        assert_eq!(b1.display, Some("Go to C1".to_string()));

        // Email link roundtrip.
        let c1 = wb2.get_cell_hyperlink("Sheet1", "C1").unwrap().unwrap();
        assert_eq!(
            c1.link_type,
            HyperlinkType::Email("mailto:hello@example.com".to_string())
        );
        assert_eq!(c1.display, Some("Email".to_string()));
    }

    #[test]
    fn test_hyperlink_on_nonexistent_sheet() {
        use crate::hyperlink::HyperlinkType;

        let mut wb = Workbook::new();
        let result = wb.set_cell_hyperlink(
            "NoSheet",
            "A1",
            HyperlinkType::External("https://example.com".to_string()),
            None,
            None,
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_workbook_get_rows_empty_sheet() {
        let wb = Workbook::new();
        let rows = wb.get_rows("Sheet1").unwrap();
        assert!(rows.is_empty());
    }

    #[test]
    fn test_workbook_get_rows_with_data() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Name").unwrap();
        wb.set_cell_value("Sheet1", "B1", 42.0).unwrap();
        wb.set_cell_value("Sheet1", "A2", "Alice").unwrap();
        wb.set_cell_value("Sheet1", "B2", true).unwrap();

        let rows = wb.get_rows("Sheet1").unwrap();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].0, 1);
        assert_eq!(rows[0].1.len(), 2);
        assert_eq!(rows[0].1[0].0, "A");
        assert_eq!(rows[0].1[0].1, CellValue::String("Name".to_string()));
        assert_eq!(rows[0].1[1].0, "B");
        assert_eq!(rows[0].1[1].1, CellValue::Number(42.0));
        assert_eq!(rows[1].0, 2);
        assert_eq!(rows[1].1[0].1, CellValue::String("Alice".to_string()));
        assert_eq!(rows[1].1[1].1, CellValue::Bool(true));
    }

    #[test]
    fn test_workbook_get_rows_sheet_not_found() {
        let wb = Workbook::new();
        assert!(wb.get_rows("NoSheet").is_err());
    }

    #[test]
    fn test_workbook_get_cols_empty_sheet() {
        let wb = Workbook::new();
        let cols = wb.get_cols("Sheet1").unwrap();
        assert!(cols.is_empty());
    }

    #[test]
    fn test_workbook_get_cols_with_data() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Name").unwrap();
        wb.set_cell_value("Sheet1", "B1", 42.0).unwrap();
        wb.set_cell_value("Sheet1", "A2", "Alice").unwrap();
        wb.set_cell_value("Sheet1", "B2", 30.0).unwrap();

        let cols = wb.get_cols("Sheet1").unwrap();
        assert_eq!(cols.len(), 2);
        assert_eq!(cols[0].0, "A");
        assert_eq!(cols[0].1.len(), 2);
        assert_eq!(cols[0].1[0], (1, CellValue::String("Name".to_string())));
        assert_eq!(cols[0].1[1], (2, CellValue::String("Alice".to_string())));
        assert_eq!(cols[1].0, "B");
        assert_eq!(cols[1].1[0], (1, CellValue::Number(42.0)));
        assert_eq!(cols[1].1[1], (2, CellValue::Number(30.0)));
    }

    #[test]
    fn test_workbook_get_cols_sheet_not_found() {
        let wb = Workbook::new();
        assert!(wb.get_cols("NoSheet").is_err());
    }

    #[test]
    fn test_workbook_get_rows_roundtrip_save_open() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "hello").unwrap();
        wb.set_cell_value("Sheet1", "B1", 99.0).unwrap();
        wb.set_cell_value("Sheet1", "A2", true).unwrap();

        let tmp = std::env::temp_dir().join("test_get_rows_roundtrip.xlsx");
        wb.save(&tmp).unwrap();

        let wb2 = Workbook::open(&tmp).unwrap();
        let rows = wb2.get_rows("Sheet1").unwrap();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].1[0].1, CellValue::String("hello".to_string()));
        assert_eq!(rows[0].1[1].1, CellValue::Number(99.0));
        assert_eq!(rows[1].1[0].1, CellValue::Bool(true));

        let _ = std::fs::remove_file(&tmp);
    }

    // -- Pivot table tests --

    fn make_pivot_workbook() -> Workbook {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Name").unwrap();
        wb.set_cell_value("Sheet1", "B1", "Region").unwrap();
        wb.set_cell_value("Sheet1", "C1", "Sales").unwrap();
        wb.set_cell_value("Sheet1", "A2", "Alice").unwrap();
        wb.set_cell_value("Sheet1", "B2", "North").unwrap();
        wb.set_cell_value("Sheet1", "C2", 100.0).unwrap();
        wb.set_cell_value("Sheet1", "A3", "Bob").unwrap();
        wb.set_cell_value("Sheet1", "B3", "South").unwrap();
        wb.set_cell_value("Sheet1", "C3", 200.0).unwrap();
        wb.set_cell_value("Sheet1", "A4", "Carol").unwrap();
        wb.set_cell_value("Sheet1", "B4", "North").unwrap();
        wb.set_cell_value("Sheet1", "C4", 150.0).unwrap();
        wb
    }

    fn basic_pivot_config() -> PivotTableConfig {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        PivotTableConfig {
            name: "PivotTable1".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "E1".to_string(),
            rows: vec![PivotField {
                name: "Name".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        }
    }

    #[test]
    fn test_add_pivot_table_basic() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        assert_eq!(wb.pivot_tables.len(), 1);
        assert_eq!(wb.pivot_cache_defs.len(), 1);
        assert_eq!(wb.pivot_cache_records.len(), 1);
        assert_eq!(wb.pivot_tables[0].1.name, "PivotTable1");
        assert_eq!(wb.pivot_tables[0].1.cache_id, 0);
    }

    #[test]
    fn test_add_pivot_table_with_columns() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();
        let config = PivotTableConfig {
            name: "PT2".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "E1".to_string(),
            rows: vec![PivotField {
                name: "Name".to_string(),
            }],
            columns: vec![PivotField {
                name: "Region".to_string(),
            }],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Average,
                display_name: Some("Avg Sales".to_string()),
            }],
        };
        wb.add_pivot_table(&config).unwrap();

        let pt = &wb.pivot_tables[0].1;
        assert!(pt.row_fields.is_some());
        assert!(pt.col_fields.is_some());
        assert!(pt.data_fields.is_some());
    }

    #[test]
    fn test_add_pivot_table_source_sheet_not_found() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = Workbook::new();
        let config = PivotTableConfig {
            name: "PT".to_string(),
            source_sheet: "NonExistent".to_string(),
            source_range: "A1:B2".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![PivotField {
                name: "Col1".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Col2".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        let result = wb.add_pivot_table(&config);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_add_pivot_table_target_sheet_not_found() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();
        let config = PivotTableConfig {
            name: "PT".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Report".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![PivotField {
                name: "Name".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        let result = wb.add_pivot_table(&config);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_add_pivot_table_duplicate_name() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        let result = wb.add_pivot_table(&config);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::PivotTableAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_get_pivot_tables_empty() {
        let wb = Workbook::new();
        let pts = wb.get_pivot_tables();
        assert!(pts.is_empty());
    }

    #[test]
    fn test_get_pivot_tables_after_add() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        let pts = wb.get_pivot_tables();
        assert_eq!(pts.len(), 1);
        assert_eq!(pts[0].name, "PivotTable1");
        assert_eq!(pts[0].source_sheet, "Sheet1");
        assert_eq!(pts[0].source_range, "A1:C4");
        assert_eq!(pts[0].target_sheet, "Sheet1");
        assert_eq!(pts[0].location, "E1");
    }

    #[test]
    fn test_delete_pivot_table() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();
        assert_eq!(wb.pivot_tables.len(), 1);

        wb.delete_pivot_table("PivotTable1").unwrap();
        assert!(wb.pivot_tables.is_empty());
        assert!(wb.pivot_cache_defs.is_empty());
        assert!(wb.pivot_cache_records.is_empty());
        assert!(wb.workbook_xml.pivot_caches.is_none());

        // Content type overrides for pivot parts should be gone.
        let pivot_overrides: Vec<_> = wb
            .content_types
            .overrides
            .iter()
            .filter(|o| {
                o.content_type == mime_types::PIVOT_TABLE
                    || o.content_type == mime_types::PIVOT_CACHE_DEFINITION
                    || o.content_type == mime_types::PIVOT_CACHE_RECORDS
            })
            .collect();
        assert!(pivot_overrides.is_empty());
    }

    #[test]
    fn test_delete_pivot_table_not_found() {
        let wb_result = Workbook::new().delete_pivot_table("NonExistent");
        assert!(wb_result.is_err());
        assert!(matches!(
            wb_result.unwrap_err(),
            Error::PivotTableNotFound { .. }
        ));
    }

    #[test]
    fn test_pivot_table_save_open_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("pivot_roundtrip.xlsx");

        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        wb.save(&path).unwrap();

        // Verify the ZIP contains pivot parts.
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/pivotTables/pivotTable1.xml").is_ok());
        assert!(archive
            .by_name("xl/pivotCache/pivotCacheDefinition1.xml")
            .is_ok());
        assert!(archive
            .by_name("xl/pivotCache/pivotCacheRecords1.xml")
            .is_ok());

        // Re-open and verify pivot table is parsed.
        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.pivot_tables.len(), 1);
        assert_eq!(wb2.pivot_tables[0].1.name, "PivotTable1");
        assert_eq!(wb2.pivot_cache_defs.len(), 1);
        assert_eq!(wb2.pivot_cache_records.len(), 1);
    }

    #[test]
    fn test_add_multiple_pivot_tables() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();

        let config1 = basic_pivot_config();
        wb.add_pivot_table(&config1).unwrap();

        let config2 = PivotTableConfig {
            name: "PivotTable2".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "H1".to_string(),
            rows: vec![PivotField {
                name: "Region".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Count,
                display_name: None,
            }],
        };
        wb.add_pivot_table(&config2).unwrap();

        assert_eq!(wb.pivot_tables.len(), 2);
        assert_eq!(wb.pivot_cache_defs.len(), 2);
        assert_eq!(wb.pivot_tables[0].1.cache_id, 0);
        assert_eq!(wb.pivot_tables[1].1.cache_id, 1);

        let pts = wb.get_pivot_tables();
        assert_eq!(pts.len(), 2);
        assert_eq!(pts[0].name, "PivotTable1");
        assert_eq!(pts[1].name, "PivotTable2");
    }

    #[test]
    fn test_add_pivot_table_content_types_added() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        let has_pt_ct = wb.content_types.overrides.iter().any(|o| {
            o.content_type == mime_types::PIVOT_TABLE
                && o.part_name == "/xl/pivotTables/pivotTable1.xml"
        });
        assert!(has_pt_ct);

        let has_pcd_ct = wb.content_types.overrides.iter().any(|o| {
            o.content_type == mime_types::PIVOT_CACHE_DEFINITION
                && o.part_name == "/xl/pivotCache/pivotCacheDefinition1.xml"
        });
        assert!(has_pcd_ct);

        let has_pcr_ct = wb.content_types.overrides.iter().any(|o| {
            o.content_type == mime_types::PIVOT_CACHE_RECORDS
                && o.part_name == "/xl/pivotCache/pivotCacheRecords1.xml"
        });
        assert!(has_pcr_ct);
    }

    #[test]
    fn test_add_pivot_table_workbook_rels_and_pivot_caches() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        // Workbook rels should have a pivot cache definition relationship.
        let cache_rel = wb
            .workbook_rels
            .relationships
            .iter()
            .find(|r| r.rel_type == rel_types::PIVOT_CACHE_DEF);
        assert!(cache_rel.is_some());
        let cache_rel = cache_rel.unwrap();
        assert_eq!(cache_rel.target, "pivotCache/pivotCacheDefinition1.xml");

        // Workbook XML should have pivot caches.
        let pivot_caches = wb.workbook_xml.pivot_caches.as_ref().unwrap();
        assert_eq!(pivot_caches.caches.len(), 1);
        assert_eq!(pivot_caches.caches[0].cache_id, 0);
        assert_eq!(pivot_caches.caches[0].r_id, cache_rel.id);
    }

    #[test]
    fn test_add_pivot_table_worksheet_rels_added() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        // Sheet1 is index 0; its rels should have a pivot table relationship.
        let ws_rels = wb.worksheet_rels.get(&0).unwrap();
        let pt_rel = ws_rels
            .relationships
            .iter()
            .find(|r| r.rel_type == rel_types::PIVOT_TABLE);
        assert!(pt_rel.is_some());
        assert_eq!(pt_rel.unwrap().target, "../pivotTables/pivotTable1.xml");
    }

    #[test]
    fn test_add_pivot_table_on_separate_target_sheet() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();
        wb.new_sheet("Report").unwrap();

        let config = PivotTableConfig {
            name: "CrossSheet".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Report".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![PivotField {
                name: "Name".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        wb.add_pivot_table(&config).unwrap();

        let pts = wb.get_pivot_tables();
        assert_eq!(pts.len(), 1);
        assert_eq!(pts[0].target_sheet, "Report");
        assert_eq!(pts[0].source_sheet, "Sheet1");

        // Worksheet rels should be on the Report sheet (index 1).
        let ws_rels = wb.worksheet_rels.get(&1).unwrap();
        let pt_rel = ws_rels
            .relationships
            .iter()
            .find(|r| r.rel_type == rel_types::PIVOT_TABLE);
        assert!(pt_rel.is_some());
    }

    #[test]
    fn test_pivot_table_invalid_source_range() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();
        let config = PivotTableConfig {
            name: "BadRange".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "INVALID".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "E1".to_string(),
            rows: vec![PivotField {
                name: "Name".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        let result = wb.add_pivot_table(&config);
        assert!(result.is_err());
    }

    #[test]
    fn test_delete_pivot_table_then_add_another() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();
        let config1 = basic_pivot_config();
        wb.add_pivot_table(&config1).unwrap();
        wb.delete_pivot_table("PivotTable1").unwrap();

        let config2 = PivotTableConfig {
            name: "PivotTable2".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "E1".to_string(),
            rows: vec![PivotField {
                name: "Region".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Max,
                display_name: None,
            }],
        };
        wb.add_pivot_table(&config2).unwrap();

        assert_eq!(wb.pivot_tables.len(), 1);
        assert_eq!(wb.pivot_tables[0].1.name, "PivotTable2");
    }

    #[test]
    fn test_pivot_table_cache_definition_stores_source_info() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        let pcd = &wb.pivot_cache_defs[0].1;
        let ws_source = pcd.cache_source.worksheet_source.as_ref().unwrap();
        assert_eq!(ws_source.sheet, "Sheet1");
        assert_eq!(ws_source.reference, "A1:C4");
        assert_eq!(pcd.cache_fields.fields.len(), 3);
        assert_eq!(pcd.cache_fields.fields[0].name, "Name");
        assert_eq!(pcd.cache_fields.fields[1].name, "Region");
        assert_eq!(pcd.cache_fields.fields[2].name, "Sales");
    }

    #[test]
    fn test_pivot_table_field_names_from_data() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        let pt = &wb.pivot_tables[0].1;
        assert_eq!(pt.pivot_fields.fields.len(), 3);
        // Name is a row field.
        assert_eq!(pt.pivot_fields.fields[0].axis, Some("axisRow".to_string()));
        // Region is not used.
        assert_eq!(pt.pivot_fields.fields[1].axis, None);
        // Sales is a data field.
        assert_eq!(pt.pivot_fields.fields[2].data_field, Some(true));
    }

    #[test]
    fn test_pivot_table_empty_header_row_error() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = Workbook::new();
        // No data set in the sheet.
        let config = PivotTableConfig {
            name: "Empty".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:B1".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "D1".to_string(),
            rows: vec![PivotField {
                name: "X".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Y".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        let result = wb.add_pivot_table(&config);
        assert!(result.is_err());
    }

    #[test]
    fn test_pivot_table_multiple_save_roundtrip() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("multi_pivot.xlsx");

        let mut wb = make_pivot_workbook();
        let config1 = basic_pivot_config();
        wb.add_pivot_table(&config1).unwrap();

        let config2 = PivotTableConfig {
            name: "PT2".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "H1".to_string(),
            rows: vec![PivotField {
                name: "Region".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Min,
                display_name: None,
            }],
        };
        wb.add_pivot_table(&config2).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.pivot_tables.len(), 2);
        let names: Vec<&str> = wb2
            .pivot_tables
            .iter()
            .map(|(_, pt)| pt.name.as_str())
            .collect();
        assert!(names.contains(&"PivotTable1"));
        assert!(names.contains(&"PT2"));
    }

    #[test]
    fn test_calculate_all_with_dependency_order() {
        let mut wb = Workbook::new();
        // A1 = 10 (value)
        wb.set_cell_value("Sheet1", "A1", 10.0).unwrap();
        // A2 = A1 * 2 (formula depends on A1)
        wb.set_cell_value(
            "Sheet1",
            "A2",
            CellValue::Formula {
                expr: "A1*2".to_string(),
                result: None,
            },
        )
        .unwrap();
        // A3 = A2 + A1 (formula depends on A2 and A1)
        wb.set_cell_value(
            "Sheet1",
            "A3",
            CellValue::Formula {
                expr: "A2+A1".to_string(),
                result: None,
            },
        )
        .unwrap();

        wb.calculate_all().unwrap();

        // A2 should be 20 (10 * 2)
        let a2 = wb.get_cell_value("Sheet1", "A2").unwrap();
        match a2 {
            CellValue::Formula { result, .. } => {
                assert_eq!(*result.unwrap(), CellValue::Number(20.0));
            }
            _ => panic!("A2 should be a formula cell"),
        }

        // A3 should be 30 (20 + 10)
        let a3 = wb.get_cell_value("Sheet1", "A3").unwrap();
        match a3 {
            CellValue::Formula { result, .. } => {
                assert_eq!(*result.unwrap(), CellValue::Number(30.0));
            }
            _ => panic!("A3 should be a formula cell"),
        }
    }

    #[test]
    fn test_calculate_all_no_formulas() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", 10.0).unwrap();
        wb.set_cell_value("Sheet1", "B1", 20.0).unwrap();
        // Should succeed without error when there are no formulas.
        wb.calculate_all().unwrap();
    }

    #[test]
    fn test_calculate_all_cycle_detection() {
        let mut wb = Workbook::new();
        // A1 = B1, B1 = A1
        wb.set_cell_value(
            "Sheet1",
            "A1",
            CellValue::Formula {
                expr: "B1".to_string(),
                result: None,
            },
        )
        .unwrap();
        wb.set_cell_value(
            "Sheet1",
            "B1",
            CellValue::Formula {
                expr: "A1".to_string(),
                result: None,
            },
        )
        .unwrap();

        let result = wb.calculate_all();
        assert!(result.is_err());
        let err_str = result.unwrap_err().to_string();
        assert!(
            err_str.contains("circular reference"),
            "expected circular reference error, got: {err_str}"
        );
    }

    #[test]
    fn test_set_and_get_cell_rich_text() {
        use crate::rich_text::RichTextRun;

        let mut wb = Workbook::new();
        let runs = vec![
            RichTextRun {
                text: "Bold".to_string(),
                font: None,
                size: None,
                bold: true,
                italic: false,
                color: None,
            },
            RichTextRun {
                text: " Normal".to_string(),
                font: None,
                size: None,
                bold: false,
                italic: false,
                color: None,
            },
        ];
        wb.set_cell_rich_text("Sheet1", "A1", runs.clone()).unwrap();

        // The cell value should be a shared string whose plain text is "Bold Normal".
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val.to_string(), "Bold Normal");

        // get_cell_rich_text should return the runs.
        let got = wb.get_cell_rich_text("Sheet1", "A1").unwrap();
        assert!(got.is_some());
        let got_runs = got.unwrap();
        assert_eq!(got_runs.len(), 2);
        assert_eq!(got_runs[0].text, "Bold");
        assert!(got_runs[0].bold);
        assert_eq!(got_runs[1].text, " Normal");
        assert!(!got_runs[1].bold);
    }

    #[test]
    fn test_get_cell_rich_text_returns_none_for_plain() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("plain".to_string()))
            .unwrap();
        let got = wb.get_cell_rich_text("Sheet1", "A1").unwrap();
        assert!(got.is_none());
    }

    #[test]
    fn test_rich_text_roundtrip_save_open() {
        use crate::rich_text::RichTextRun;

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("rich_text.xlsx");

        // Note: quick-xml's serde deserializer trims leading and trailing
        // whitespace from text content. To avoid false failures, test text
        // values must not rely on boundary whitespace being preserved.
        let mut wb = Workbook::new();
        let runs = vec![
            RichTextRun {
                text: "Hello".to_string(),
                font: Some("Arial".to_string()),
                size: Some(14.0),
                bold: true,
                italic: false,
                color: Some("#FF0000".to_string()),
            },
            RichTextRun {
                text: "World".to_string(),
                font: None,
                size: None,
                bold: false,
                italic: true,
                color: None,
            },
        ];
        wb.set_cell_rich_text("Sheet1", "B2", runs).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let val = wb2.get_cell_value("Sheet1", "B2").unwrap();
        assert_eq!(val.to_string(), "HelloWorld");

        let got = wb2.get_cell_rich_text("Sheet1", "B2").unwrap();
        assert!(got.is_some());
        let got_runs = got.unwrap();
        assert_eq!(got_runs.len(), 2);
        assert_eq!(got_runs[0].text, "Hello");
        assert!(got_runs[0].bold);
        assert_eq!(got_runs[0].font.as_deref(), Some("Arial"));
        assert_eq!(got_runs[0].size, Some(14.0));
        assert_eq!(got_runs[0].color.as_deref(), Some("#FF0000"));
        assert_eq!(got_runs[1].text, "World");
        assert!(got_runs[1].italic);
        assert!(!got_runs[1].bold);
    }

    // Defined names workbook API tests

    #[test]
    fn test_set_defined_name_workbook_scope() {
        let mut wb = Workbook::new();
        wb.set_defined_name("SalesData", "Sheet1!$A$1:$D$10", None, None)
            .unwrap();

        let info = wb.get_defined_name("SalesData", None).unwrap().unwrap();
        assert_eq!(info.name, "SalesData");
        assert_eq!(info.value, "Sheet1!$A$1:$D$10");
        assert_eq!(info.scope, crate::defined_names::DefinedNameScope::Workbook);
        assert!(info.comment.is_none());
    }

    #[test]
    fn test_set_defined_name_sheet_scope() {
        let mut wb = Workbook::new();
        wb.set_defined_name("LocalRange", "Sheet1!$B$2:$C$5", Some("Sheet1"), None)
            .unwrap();

        let info = wb
            .get_defined_name("LocalRange", Some("Sheet1"))
            .unwrap()
            .unwrap();
        assert_eq!(info.name, "LocalRange");
        assert_eq!(info.value, "Sheet1!$B$2:$C$5");
        assert_eq!(info.scope, crate::defined_names::DefinedNameScope::Sheet(0));
    }

    #[test]
    fn test_update_existing_defined_name() {
        let mut wb = Workbook::new();
        wb.set_defined_name("DataRange", "Sheet1!$A$1:$A$10", None, None)
            .unwrap();

        wb.set_defined_name("DataRange", "Sheet1!$A$1:$A$50", None, Some("Updated"))
            .unwrap();

        let all = wb.get_all_defined_names();
        assert_eq!(all.len(), 1, "should not duplicate the entry");
        assert_eq!(all[0].value, "Sheet1!$A$1:$A$50");
        assert_eq!(all[0].comment, Some("Updated".to_string()));
    }

    #[test]
    fn test_get_all_defined_names() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();

        wb.set_defined_name("Alpha", "Sheet1!$A$1", None, None)
            .unwrap();
        wb.set_defined_name("Beta", "Sheet1!$B$1", Some("Sheet1"), None)
            .unwrap();
        wb.set_defined_name("Gamma", "Sheet2!$C$1", Some("Sheet2"), None)
            .unwrap();

        let all = wb.get_all_defined_names();
        assert_eq!(all.len(), 3);
        assert_eq!(all[0].name, "Alpha");
        assert_eq!(all[1].name, "Beta");
        assert_eq!(all[2].name, "Gamma");
    }

    #[test]
    fn test_delete_defined_name() {
        let mut wb = Workbook::new();
        wb.set_defined_name("ToDelete", "Sheet1!$A$1", None, None)
            .unwrap();
        assert!(wb.get_defined_name("ToDelete", None).unwrap().is_some());

        wb.delete_defined_name("ToDelete", None).unwrap();
        assert!(wb.get_defined_name("ToDelete", None).unwrap().is_none());
    }

    #[test]
    fn test_delete_nonexistent_defined_name_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.delete_defined_name("Ghost", None);
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("Ghost"));
    }

    #[test]
    fn test_defined_name_sheet_scope_requires_existing_sheet() {
        let mut wb = Workbook::new();
        let result = wb.set_defined_name("TestName", "Sheet1!$A$1", Some("NonExistent"), None);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_defined_name_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("defined_names.xlsx");

        let mut wb = Workbook::new();
        wb.set_defined_name("Revenue", "Sheet1!$E$1:$E$100", None, Some("Total revenue"))
            .unwrap();
        wb.set_defined_name("LocalName", "Sheet1!$A$1", Some("Sheet1"), None)
            .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let all = wb2.get_all_defined_names();
        assert_eq!(all.len(), 2);
        assert_eq!(all[0].name, "Revenue");
        assert_eq!(all[0].value, "Sheet1!$E$1:$E$100");
        assert_eq!(all[0].comment, Some("Total revenue".to_string()));
        assert_eq!(all[1].name, "LocalName");
        assert_eq!(all[1].value, "Sheet1!$A$1");
        assert_eq!(
            all[1].scope,
            crate::defined_names::DefinedNameScope::Sheet(0)
        );
    }

    // Sheet protection workbook API tests

    #[test]
    fn test_protect_sheet_via_workbook() {
        let mut wb = Workbook::new();
        let config = crate::sheet::SheetProtectionConfig::default();
        wb.protect_sheet("Sheet1", &config).unwrap();

        assert!(wb.is_sheet_protected("Sheet1").unwrap());
    }

    #[test]
    fn test_unprotect_sheet_via_workbook() {
        let mut wb = Workbook::new();
        let config = crate::sheet::SheetProtectionConfig::default();
        wb.protect_sheet("Sheet1", &config).unwrap();
        assert!(wb.is_sheet_protected("Sheet1").unwrap());

        wb.unprotect_sheet("Sheet1").unwrap();
        assert!(!wb.is_sheet_protected("Sheet1").unwrap());
    }

    #[test]
    fn test_protect_sheet_nonexistent_returns_error() {
        let mut wb = Workbook::new();
        let config = crate::sheet::SheetProtectionConfig::default();
        let result = wb.protect_sheet("NoSuchSheet", &config);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_is_sheet_protected_nonexistent_returns_error() {
        let wb = Workbook::new();
        let result = wb.is_sheet_protected("NoSuchSheet");
        assert!(result.is_err());
    }

    #[test]
    fn test_protect_sheet_with_password_and_permissions() {
        let mut wb = Workbook::new();
        let config = crate::sheet::SheetProtectionConfig {
            password: Some("secret".to_string()),
            format_cells: true,
            insert_rows: true,
            sort: true,
            ..crate::sheet::SheetProtectionConfig::default()
        };
        wb.protect_sheet("Sheet1", &config).unwrap();
        assert!(wb.is_sheet_protected("Sheet1").unwrap());
    }

    #[test]
    fn test_sheet_protection_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("sheet_protection.xlsx");

        let mut wb = Workbook::new();
        let config = crate::sheet::SheetProtectionConfig {
            password: Some("pass".to_string()),
            format_cells: true,
            ..crate::sheet::SheetProtectionConfig::default()
        };
        wb.protect_sheet("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert!(wb2.is_sheet_protected("Sheet1").unwrap());
    }

    #[test]
    fn test_add_sparkline_and_get_sparklines() {
        let mut wb = Workbook::new();
        let config = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1");
        wb.add_sparkline("Sheet1", &config).unwrap();

        let sparklines = wb.get_sparklines("Sheet1").unwrap();
        assert_eq!(sparklines.len(), 1);
        assert_eq!(sparklines[0].data_range, "Sheet1!A1:A10");
        assert_eq!(sparklines[0].location, "B1");
    }

    #[test]
    fn test_add_multiple_sparklines_to_same_sheet() {
        let mut wb = Workbook::new();
        let config1 = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1");
        let config2 = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B2");
        let mut config3 = crate::sparkline::SparklineConfig::new("Sheet1!C1:C10", "D1");
        config3.sparkline_type = crate::sparkline::SparklineType::Column;

        wb.add_sparkline("Sheet1", &config1).unwrap();
        wb.add_sparkline("Sheet1", &config2).unwrap();
        wb.add_sparkline("Sheet1", &config3).unwrap();

        let sparklines = wb.get_sparklines("Sheet1").unwrap();
        assert_eq!(sparklines.len(), 3);
        assert_eq!(
            sparklines[2].sparkline_type,
            crate::sparkline::SparklineType::Column
        );
    }

    #[test]
    fn test_remove_sparkline_by_location() {
        let mut wb = Workbook::new();
        let config1 = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1");
        let config2 = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B2");
        wb.add_sparkline("Sheet1", &config1).unwrap();
        wb.add_sparkline("Sheet1", &config2).unwrap();

        wb.remove_sparkline("Sheet1", "B1").unwrap();

        let sparklines = wb.get_sparklines("Sheet1").unwrap();
        assert_eq!(sparklines.len(), 1);
        assert_eq!(sparklines[0].location, "B2");
    }

    #[test]
    fn test_remove_nonexistent_sparkline_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.remove_sparkline("Sheet1", "Z99");
        assert!(result.is_err());
    }

    #[test]
    fn test_sparkline_on_nonexistent_sheet_returns_error() {
        let mut wb = Workbook::new();
        let config = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1");
        let result = wb.add_sparkline("NoSuchSheet", &config);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));

        let result = wb.get_sparklines("NoSuchSheet");
        assert!(result.is_err());
    }

    #[test]
    fn test_sparkline_save_open_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("sparkline_roundtrip.xlsx");

        let mut wb = Workbook::new();
        for i in 1..=10 {
            wb.set_cell_value(
                "Sheet1",
                &format!("A{i}"),
                CellValue::Number(i as f64 * 10.0),
            )
            .unwrap();
        }

        let mut config = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1");
        config.sparkline_type = crate::sparkline::SparklineType::Column;
        config.markers = true;
        config.high_point = true;
        config.line_weight = Some(1.5);

        wb.add_sparkline("Sheet1", &config).unwrap();

        let config2 = crate::sparkline::SparklineConfig::new("Sheet1!A1:A5", "C1");
        wb.add_sparkline("Sheet1", &config2).unwrap();

        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let sparklines = wb2.get_sparklines("Sheet1").unwrap();
        assert_eq!(sparklines.len(), 2);
        assert_eq!(sparklines[0].data_range, "Sheet1!A1:A10");
        assert_eq!(sparklines[0].location, "B1");
        assert_eq!(
            sparklines[0].sparkline_type,
            crate::sparkline::SparklineType::Column
        );
        assert!(sparklines[0].markers);
        assert!(sparklines[0].high_point);
        assert_eq!(sparklines[0].line_weight, Some(1.5));
        assert_eq!(sparklines[1].data_range, "Sheet1!A1:A5");
        assert_eq!(sparklines[1].location, "C1");
    }

    #[test]
    fn test_sparkline_empty_sheet_returns_empty_vec() {
        let wb = Workbook::new();
        let sparklines = wb.get_sparklines("Sheet1").unwrap();
        assert!(sparklines.is_empty());
    }
}
