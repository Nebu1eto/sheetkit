//! Workbook file I/O: reading and writing `.xlsx` files.
//!
//! An `.xlsx` file is a ZIP archive containing XML parts. This module provides
//! [`Workbook`] which holds the parsed XML structures in memory and can
//! serialize them back to a valid `.xlsx` file.

use std::collections::{HashMap, HashSet};
use std::io::{Read as _, Write as _};
use std::path::Path;
use std::sync::OnceLock;

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
use crate::threaded_comment::{PersonData, PersonInput, ThreadedCommentData, ThreadedCommentInput};
use crate::utils::cell_ref::{cell_name_to_coordinates, column_name_to_number};
use crate::utils::constants::MAX_CELL_CHARS;
use crate::validation::DataValidationConfig;
use crate::workbook_paths::{
    default_relationships, relationship_part_path, relative_relationship_target,
    resolve_relationship_target,
};

pub(crate) mod aux;
mod cell_ops;
mod data;
mod drawing;
mod features;
mod io;
mod open_options;
mod sheet_ops;
mod source;

pub use open_options::{AuxParts, OpenOptions, ReadMode};
pub(crate) use source::PackageSource;

/// Helper to initialize an `OnceLock<WorksheetXml>` with a value at
/// construction time. Avoids repeating the `set`+`unwrap` pattern.
pub(crate) fn initialized_lock(ws: WorksheetXml) -> OnceLock<WorksheetXml> {
    let lock = OnceLock::new();
    let _ = lock.set(ws);
    lock
}

/// XML declaration prepended to every XML part in the package.
const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// In-memory representation of an `.xlsx` workbook.
pub struct Workbook {
    format: WorkbookFormat,
    content_types: ContentTypes,
    package_rels: Relationships,
    workbook_xml: WorkbookXml,
    workbook_rels: Relationships,
    /// Per-sheet worksheet XML, stored as `(name, OnceLock<WorksheetXml>)`.
    /// When a sheet is eagerly parsed, the `OnceLock` is initialized at open
    /// time. When a sheet is deferred (lazy mode or filtered out), the lock
    /// is empty and `raw_sheet_xml[i]` holds the raw bytes; the first call
    /// to [`worksheet_ref`] or [`worksheet_mut`] hydrates the lock on demand.
    worksheets: Vec<(String, OnceLock<WorksheetXml>)>,
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
    /// ZIP entries not recognized by the parser, preserved for round-trip fidelity.
    /// Each entry is (zip_path, raw_bytes).
    unknown_parts: Vec<(String, Vec<u8>)>,
    /// Typed index of auxiliary parts deferred during Lazy/Stream open.
    /// Stores raw bytes grouped by category (comments, charts, doc props, etc.)
    /// and supports on-demand hydration with dirty tracking.
    deferred_parts: aux::DeferredAuxParts,
    /// Raw VBA project binary blob (`xl/vbaProject.bin`), preserved for round-trip
    /// and used for VBA module extraction. `None` for non-macro workbooks.
    vba_blob: Option<Vec<u8>>,
    /// Table parts: (zip path like "xl/tables/table1.xml", TableXml data, sheet_index).
    tables: Vec<(String, sheetkit_xml::table::TableXml, usize)>,
    /// Raw XML bytes for sheets that were not parsed during open.
    /// Parallel to `worksheets`. `Some(bytes)` means the sheet XML has not
    /// been deserialized: either filtered out by the `sheets` option, or
    /// deferred in Lazy/Stream mode. The bytes are written directly on save
    /// if the corresponding `OnceLock` in `worksheets` was never initialized.
    raw_sheet_xml: Vec<Option<Vec<u8>>>,
    /// Slicer definition parts: (zip path, SlicerDefinitions data).
    slicer_defs: Vec<(String, sheetkit_xml::slicer::SlicerDefinitions)>,
    /// Slicer cache definition parts: (zip path, raw XML string).
    slicer_caches: Vec<(String, sheetkit_xml::slicer::SlicerCacheDefinition)>,
    /// Per-sheet threaded comments (Excel 2019+), parallel to the `worksheets` vector.
    sheet_threaded_comments: Vec<Option<sheetkit_xml::threaded_comment::ThreadedComments>>,
    /// Person list shared across all sheets (for threaded comment authors).
    person_list: sheetkit_xml::threaded_comment::PersonList,
    /// Per-sheet form control configurations, parallel to `worksheets`.
    sheet_form_controls: Vec<Vec<crate::control::FormControlConfig>>,
    /// O(1) sheet name -> index lookup cache. Must be kept in sync with
    /// `worksheets` via [`rebuild_sheet_index`].
    sheet_name_index: HashMap<String, usize>,
    /// Streamed sheet data keyed by sheet index. During save, these sheets
    /// are written by streaming from their temp files instead of serializing
    /// the (empty placeholder) WorksheetXml.
    streamed_sheets: HashMap<usize, crate::stream::StreamedSheetData>,
    /// Backing storage for the xlsx package, retained for lazy part access.
    #[allow(dead_code)]
    package_source: Option<PackageSource>,
    /// Read mode used when this workbook was opened.
    read_mode: ReadMode,
    /// Optional row limit from `OpenOptions::sheet_rows`, applied during
    /// on-demand hydration of deferred sheets.
    sheet_rows_limit: Option<u32>,
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

    /// Invalidate streamed data for a sheet by index. Must be called before
    /// any mutation to a sheet that may have been created via StreamWriter,
    /// so that the normal WorksheetXml serialization path is used on save.
    pub(crate) fn invalidate_streamed(&mut self, idx: usize) {
        self.streamed_sheets.remove(&idx);
    }

    /// Get a mutable reference to the worksheet XML for the named sheet.
    ///
    /// If the sheet has streamed data (from [`apply_stream_writer`]), the
    /// streamed entry is removed so that subsequent edits are not silently
    /// ignored on save. Deferred sheets are hydrated on demand.
    pub(crate) fn worksheet_mut(&mut self, sheet: &str) -> Result<&mut WorksheetXml> {
        let idx = self.sheet_index(sheet)?;
        self.invalidate_streamed(idx);
        self.ensure_hydrated(idx)?;
        Ok(self.worksheets[idx].1.get_mut().unwrap())
    }

    /// Get an immutable reference to the worksheet XML for the named sheet.
    /// Deferred sheets are hydrated lazily via `OnceLock`.
    pub(crate) fn worksheet_ref(&self, sheet: &str) -> Result<&WorksheetXml> {
        let idx = self.sheet_index(sheet)?;
        self.worksheet_ref_by_index(idx)
    }

    /// Get an immutable reference to the worksheet XML by index.
    /// Deferred sheets are hydrated lazily via `OnceLock`.
    pub(crate) fn worksheet_ref_by_index(&self, idx: usize) -> Result<&WorksheetXml> {
        if let Some(ws) = self.worksheets[idx].1.get() {
            return Ok(ws);
        }
        // Hydrate from raw_sheet_xml on first access.
        if let Some(Some(bytes)) = self.raw_sheet_xml.get(idx) {
            let mut ws = io::deserialize_worksheet_xml(bytes)?;
            if let Some(max_rows) = self.sheet_rows_limit {
                ws.sheet_data.rows.truncate(max_rows as usize);
            }
            Ok(self.worksheets[idx].1.get_or_init(|| ws))
        } else {
            Err(Error::Internal(format!(
                "sheet at index {} has no materialized or deferred data",
                idx
            )))
        }
    }

    /// Public immutable reference to a worksheet's XML by sheet name.
    /// Deferred sheets are hydrated lazily on first access.
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
        for (i, (name, _ws_lock)) in self.worksheets.iter().enumerate() {
            self.sheet_name_index.insert(name.clone(), i);
        }
    }

    /// Ensure the sheet at the given index is hydrated (parsed from raw XML).
    /// This is used by `&mut self` methods that need a mutable `OnceLock`
    /// reference via `get_mut()`, which requires the lock to be initialized.
    fn ensure_hydrated(&mut self, idx: usize) -> Result<()> {
        if self.worksheets[idx].1.get().is_some() {
            // OnceLock is set. If raw bytes are still present, this is a
            // placeholder (filtered-out sheet with WorksheetXml::default()).
            // Replace the placeholder with properly parsed data.
            if let Some(Some(bytes)) = self.raw_sheet_xml.get(idx) {
                let mut ws = io::deserialize_worksheet_xml(bytes)?;
                if let Some(max_rows) = self.sheet_rows_limit {
                    ws.sheet_data.rows.truncate(max_rows as usize);
                }
                *self.worksheets[idx].1.get_mut().unwrap() = ws;
                self.raw_sheet_xml[idx] = None;
            }
            return Ok(());
        }
        if let Some(Some(bytes)) = self.raw_sheet_xml.get(idx) {
            let mut ws = io::deserialize_worksheet_xml(bytes)?;
            if let Some(max_rows) = self.sheet_rows_limit {
                ws.sheet_data.rows.truncate(max_rows as usize);
            }
            let _ = self.worksheets[idx].1.set(ws);
            self.raw_sheet_xml[idx] = None;
            Ok(())
        } else {
            Err(Error::Internal(format!(
                "sheet at index {} has no materialized or deferred data",
                idx
            )))
        }
    }

    /// Hydrate if needed and return a mutable reference to the worksheet
    /// at the given index. Callers must hold `&mut self`.
    pub(crate) fn worksheet_mut_by_index(&mut self, idx: usize) -> Result<&mut WorksheetXml> {
        self.ensure_hydrated(idx)?;
        Ok(self.worksheets[idx].1.get_mut().unwrap())
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

    /// Create a forward-only streaming reader for the named sheet.
    ///
    /// The reader processes worksheet XML row-by-row without materializing the
    /// full DOM, enabling bounded-memory processing of large worksheets. The
    /// workbook's shared string table and optional `sheet_rows` limit are
    /// passed through to the reader.
    ///
    /// The XML bytes come from `raw_sheet_xml` (deferred sheets in Lazy/Stream
    /// mode) or from a freshly hydrated worksheet serialized back to bytes.
    pub fn open_sheet_reader(
        &self,
        sheet: &str,
    ) -> Result<
        crate::stream_reader::SheetStreamReader<'_, std::io::BufReader<std::io::Cursor<Vec<u8>>>>,
    > {
        let idx = self.sheet_index(sheet)?;
        let xml_bytes = self.sheet_xml_bytes(idx)?;
        let cursor = std::io::Cursor::new(xml_bytes);
        let buf_reader = std::io::BufReader::new(cursor);
        Ok(crate::stream_reader::SheetStreamReader::new(
            buf_reader,
            &self.sst_runtime,
            self.sheet_rows_limit,
        ))
    }

    /// Get the raw XML bytes for a sheet by index.
    ///
    /// When the OnceLock is uninitialised (Lazy/Stream deferred), raw bytes
    /// from `raw_sheet_xml` are used so the DOM is never materialised. When
    /// the OnceLock IS initialised (Eager parse or filtered-out sheet), the
    /// parsed worksheet is serialised back so that `sheets(...)` filtering is
    /// respected (filtered sheets have an empty worksheet placeholder).
    ///
    /// The returned bytes are cloned because the `SheetStreamReader` takes
    /// ownership of its `BufRead` source.
    fn sheet_xml_bytes(&self, idx: usize) -> Result<Vec<u8>> {
        // If the OnceLock is already initialised (eager parse OR filtered-out
        // placeholder), serialise whatever is stored there. This ensures
        // filtered-out sheets yield an empty worksheet.
        if let Some(ws) = self.worksheets[idx].1.get() {
            let xml = quick_xml::se::to_string(ws)
                .map_err(|e| Error::Internal(format!("failed to serialize worksheet: {e}")))?;
            return Ok(xml.into_bytes());
        }
        // Lazy/Stream deferred: OnceLock not yet initialised, use raw bytes.
        if let Some(Some(bytes)) = self.raw_sheet_xml.get(idx) {
            return Ok(bytes.clone());
        }
        Err(Error::Internal(format!(
            "sheet at index {} has no materialized or deferred data",
            idx
        )))
    }
}
