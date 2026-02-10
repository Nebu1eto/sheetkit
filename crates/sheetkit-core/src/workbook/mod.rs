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

/// The OOXML package format, determined by the workbook content type in
/// `[Content_Types].xml`. Controls which content type string is emitted for
/// `xl/workbook.xml` on save.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub enum WorkbookFormat {
    /// Standard spreadsheet (.xlsx).
    #[default]
    Xlsx,
    /// Macro-enabled spreadsheet (.xlsm).
    Xlsm,
    /// Template (.xltx).
    Xltx,
    /// Macro-enabled template (.xltm).
    Xltm,
    /// Macro-enabled add-in (.xlam).
    Xlam,
}

impl WorkbookFormat {
    /// Infer the format from a workbook content type string found in
    /// `[Content_Types].xml`.
    pub fn from_content_type(ct: &str) -> Option<Self> {
        match ct {
            mime_types::WORKBOOK => Some(Self::Xlsx),
            mime_types::WORKBOOK_MACRO => Some(Self::Xlsm),
            mime_types::WORKBOOK_TEMPLATE => Some(Self::Xltx),
            mime_types::WORKBOOK_TEMPLATE_MACRO => Some(Self::Xltm),
            mime_types::WORKBOOK_ADDIN_MACRO => Some(Self::Xlam),
            _ => None,
        }
    }

    /// Infer the format from a file extension (case-insensitive, without the
    /// leading dot). Returns `None` for unrecognized extensions.
    pub fn from_extension(ext: &str) -> Option<Self> {
        match ext.to_ascii_lowercase().as_str() {
            "xlsx" => Some(Self::Xlsx),
            "xlsm" => Some(Self::Xlsm),
            "xltx" => Some(Self::Xltx),
            "xltm" => Some(Self::Xltm),
            "xlam" => Some(Self::Xlam),
            _ => None,
        }
    }

    /// Return the OOXML content type string for this format.
    pub fn content_type(self) -> &'static str {
        match self {
            Self::Xlsx => mime_types::WORKBOOK,
            Self::Xlsm => mime_types::WORKBOOK_MACRO,
            Self::Xltx => mime_types::WORKBOOK_TEMPLATE,
            Self::Xltm => mime_types::WORKBOOK_TEMPLATE_MACRO,
            Self::Xlam => mime_types::WORKBOOK_ADDIN_MACRO,
        }
    }
}

use sheetkit_xml::drawing::{MarkerType, WsDr};
use sheetkit_xml::relationships::{self, rel_types, Relationship, Relationships};
use sheetkit_xml::shared_strings::Sst;
use sheetkit_xml::styles::StyleSheet;
use sheetkit_xml::workbook::{WorkbookProtection, WorkbookXml};
use sheetkit_xml::worksheet::{Cell, CellFormula, CellTypeTag, DrawingRef, Row, WorksheetXml};
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

mod cell_ops;
mod data;
mod drawing;
mod features;
mod io;
mod sheet_ops;

/// XML declaration prepended to every XML part in the package.
const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// In-memory representation of an `.xlsx` workbook.
pub struct Workbook {
    format: WorkbookFormat,
    content_types: ContentTypes,
    package_rels: Relationships,
    workbook_xml: WorkbookXml,
    workbook_rels: Relationships,
    worksheets: Vec<(String, WorksheetXml)>,
    stylesheet: StyleSheet,
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
    /// Per-sheet VML drawing bytes (for legacy comment rendering), parallel to `worksheets`.
    /// `None` means no VML part exists for that sheet.
    sheet_vml: Vec<Option<Vec<u8>>>,
    /// Raw VBA project binary blob (`xl/vbaProject.bin`), preserved opaquely for round-trip.
    /// `None` for non-macro workbooks.
    vba_blob: Option<Vec<u8>>,
    /// Table parts: (zip path like "xl/tables/table1.xml", TableXml data, sheet_index).
    tables: Vec<(String, sheetkit_xml::table::TableXml, usize)>,
    /// O(1) sheet name -> index lookup cache. Must be kept in sync with
    /// `worksheets` via [`rebuild_sheet_index`].
    sheet_name_index: HashMap<String, usize>,
}

impl Workbook {
    /// Return the detected or assigned workbook format.
    pub fn format(&self) -> WorkbookFormat {
        self.format
    }

    /// Set the workbook format. This determines the content type written for
    /// `xl/workbook.xml` on save.
    pub fn set_format(&mut self, format: WorkbookFormat) {
        self.format = format;
    }

    /// Get the 0-based index of a sheet by name. O(1) via HashMap.
    pub(crate) fn sheet_index(&self, sheet: &str) -> Result<usize> {
        self.sheet_name_index
            .get(sheet)
            .copied()
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })
    }

    /// Get a mutable reference to the worksheet XML for the named sheet.
    pub(crate) fn worksheet_mut(&mut self, sheet: &str) -> Result<&mut WorksheetXml> {
        let idx = self.sheet_index(sheet)?;
        Ok(&mut self.worksheets[idx].1)
    }

    /// Get an immutable reference to the worksheet XML for the named sheet.
    pub(crate) fn worksheet_ref(&self, sheet: &str) -> Result<&WorksheetXml> {
        let idx = self.sheet_index(sheet)?;
        Ok(&self.worksheets[idx].1)
    }

    /// Public immutable reference to a worksheet's XML by sheet name.
    pub fn worksheet_xml_ref(&self, sheet: &str) -> Result<&WorksheetXml> {
        self.worksheet_ref(sheet)
    }

    /// Public immutable reference to the shared string table.
    pub fn sst_ref(&self) -> &SharedStringTable {
        &self.sst_runtime
    }

    /// Rebuild the sheet name -> index lookup after any structural change
    /// to the worksheets vector.
    pub(crate) fn rebuild_sheet_index(&mut self) {
        self.sheet_name_index.clear();
        for (i, (name, _)) in self.worksheets.iter().enumerate() {
            self.sheet_name_index.insert(name.clone(), i);
        }
    }

    /// Resolve the part path for a sheet index from workbook relationships.
    /// Falls back to the default `xl/worksheets/sheet{N}.xml` naming.
    pub(crate) fn sheet_part_path(&self, sheet_idx: usize) -> String {
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
}
