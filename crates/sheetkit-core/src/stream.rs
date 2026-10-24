//! Streaming worksheet writer.
//!
//! The [`StreamWriter`] writes row data directly to a temporary file instead of
//! accumulating it in memory. This enables writing sheets with millions of rows
//! without holding the entire worksheet in memory.
//!
//! Strings are written as inline strings (`<is><t>...</t></is>`) rather than
//! shared string references, eliminating the need for SST index remapping and
//! allowing each row to be serialized independently.
//!
//! Rows must be written in ascending order.
//!
//! # Example
//!
//! ```
//! use sheetkit_core::stream::StreamWriter;
//! use sheetkit_core::cell::CellValue;
//!
//! let mut sw = StreamWriter::new("Sheet1");
//! sw.set_col_width(1, 20.0).unwrap();
//! sw.write_row(1, &[CellValue::from("Name"), CellValue::from("Age")]).unwrap();
//! sw.write_row(2, &[CellValue::from("Alice"), CellValue::from(30)]).unwrap();
//! // StreamWriter data is applied to a workbook via apply_stream_writer().
//! ```

use std::fmt::Write as _;
use std::io::{BufWriter, Write as _};

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::resource::{ResourceCategory, ResourceDimension, ResourceOperation, ResourceStage};
use crate::save::{resource_limit, StreamLimits};
use crate::utils::cell_ref::{cell_name_to_coordinates, coordinates_to_cell_name};
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS, MAX_ROW_HEIGHT};

use sheetkit_xml::worksheet::{Col, Cols, MergeCell, MergeCells, Pane, SheetView, SheetViews};

/// Maximum outline level allowed by Excel.
const MAX_OUTLINE_LEVEL: u8 = 7;

/// Capacity of the buffered writer in front of the row temp file.
const TEMP_FILE_BUFFER_BYTES: usize = 64 * 1024;

/// Largest row XML buffer capacity retained between rows.
const RETAINED_ROW_BUFFER_BYTES: usize = 1024 * 1024;

/// Options for a streamed row.
#[derive(Debug, Clone, Default)]
pub struct StreamRowOptions {
    /// Custom row height in points.
    pub height: Option<f64>,
    /// Row visibility (false = hidden).
    pub visible: Option<bool>,
    /// Outline level (0-7).
    pub outline_level: Option<u8>,
    /// Style ID for the row.
    pub style_id: Option<u32>,
}

/// Internal column configuration beyond width.
#[derive(Debug, Clone, Default)]
struct StreamColOptions {
    width: Option<f64>,
    style_id: Option<u32>,
    hidden: Option<bool>,
    outline_level: Option<u8>,
}

/// Metadata for a streamed sheet stored in the workbook after applying.
///
/// Contains the temporary file with pre-serialized row XML and the
/// configuration needed to compose the full worksheet XML on save.
pub struct StreamedSheetData {
    /// Temporary file containing `<row>...</row>` XML fragments.
    pub(crate) temp_file: tempfile::NamedTempFile,
    /// Number of bytes written to the temp file (used for diagnostics/testing).
    #[allow(dead_code)]
    pub(crate) data_len: u64,
    /// Number of rows written to this streamed sheet.
    pub(crate) row_count: u64,
    /// Number of cell slots visited while writing this streamed sheet.
    pub(crate) cell_count: u64,
    /// Pre-built `<sheetViews>` XML fragment (if freeze panes are set).
    pub(crate) sheet_views_xml: Option<String>,
    /// Pre-built `<cols>` XML fragment (if column settings exist).
    pub(crate) cols_xml: Option<String>,
    /// Pre-built `<mergeCells>` XML fragment (if merge cells exist).
    pub(crate) merge_cells_xml: Option<String>,
}

impl StreamedSheetData {
    /// Create a deep copy of this streamed sheet data, duplicating the
    /// temp file contents so the copy is fully independent.
    pub(crate) fn try_clone(&self) -> Result<Self> {
        let mut new_temp = tempfile::NamedTempFile::new()?;
        let mut src = std::fs::File::open(self.temp_file.path())?;
        std::io::copy(&mut src, new_temp.as_file_mut())?;
        Ok(StreamedSheetData {
            temp_file: new_temp,
            data_len: self.data_len,
            row_count: self.row_count,
            cell_count: self.cell_count,
            sheet_views_xml: self.sheet_views_xml.clone(),
            cols_xml: self.cols_xml.clone(),
            merge_cells_xml: self.merge_cells_xml.clone(),
        })
    }
}

#[derive(Debug, Clone, Copy, Default)]
struct StreamFormulaUsage {
    cells: u64,
    input_bytes: u64,
    output_bytes: u64,
}

/// A streaming worksheet writer that writes rows directly to a temp file.
///
/// Rows must be written in ascending order. Each row is serialized to XML
/// and written to a temporary file immediately, keeping memory usage constant
/// regardless of the number of rows.
pub struct StreamWriter {
    sheet_name: String,
    writer: Option<BufWriter<tempfile::NamedTempFile>>,
    bytes_written: u64,
    row_count: u64,
    cell_count: u64,
    limits: StreamLimits,
    formula_usage: StreamFormulaUsage,
    last_row: u32,
    started: bool,
    finished: bool,
    failed: bool,
    config_count: u64,
    config_bytes: u64,
    write_failure_after_bytes: Option<u64>,
    /// Reusable buffer holding one rendered row before it reaches the file.
    row_buf: String,
    col_widths: Vec<(u32, u32, f64)>,
    col_options: Vec<(u32, StreamColOptions)>,
    /// Canonical merge references with their normalized
    /// (min_col, min_row, max_col, max_row) rectangles.
    merge_cells: Vec<(String, (u32, u32, u32, u32))>,
    /// Freeze pane: (x_split, y_split, top_left_cell).
    freeze_pane: Option<(u32, u32, String)>,
}

// Compile-time guard: StreamWriter must stay Send without an unsafe impl.
const _: () = {
    fn assert_send<T: Send>() {}
    let _ = assert_send::<StreamWriter>;
};

impl std::fmt::Debug for StreamWriter {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("StreamWriter")
            .field("sheet_name", &self.sheet_name)
            .field("bytes_written", &self.bytes_written)
            .field("last_row", &self.last_row)
            .field("started", &self.started)
            .field("finished", &self.finished)
            .finish()
    }
}

impl StreamWriter {
    /// Create a new StreamWriter for the given sheet name.
    pub fn new(sheet_name: &str) -> Self {
        Self {
            sheet_name: sheet_name.to_string(),
            writer: None,
            bytes_written: 0,
            row_count: 0,
            cell_count: 0,
            limits: StreamLimits::default(),
            formula_usage: StreamFormulaUsage::default(),
            last_row: 0,
            started: false,
            finished: false,
            failed: false,
            config_count: 0,
            config_bytes: 0,
            write_failure_after_bytes: None,
            row_buf: String::new(),
            col_widths: Vec::new(),
            col_options: Vec::new(),
            merge_cells: Vec::new(),
            freeze_pane: None,
        }
    }

    /// Create a new StreamWriter and initialize its temporary file eagerly.
    pub fn try_new(sheet_name: &str) -> Result<Self> {
        let mut writer = Self::new(sheet_name);
        writer.ensure_temp_file()?;
        Ok(writer)
    }

    /// Create a StreamWriter with explicit finite operation limits.
    pub fn new_with_limits(sheet_name: &str, limits: StreamLimits) -> Result<Self> {
        let limits = limits.validate()?;
        Ok(Self {
            sheet_name: sheet_name.to_string(),
            writer: None,
            bytes_written: 0,
            row_count: 0,
            cell_count: 0,
            limits,
            formula_usage: StreamFormulaUsage::default(),
            last_row: 0,
            started: false,
            finished: false,
            failed: false,
            config_count: 0,
            config_bytes: 0,
            write_failure_after_bytes: None,
            row_buf: String::new(),
            col_widths: Vec::new(),
            col_options: Vec::new(),
            merge_cells: Vec::new(),
            freeze_pane: None,
        })
    }

    /// Create a limited StreamWriter and initialize its temporary file eagerly.
    pub fn try_new_with_limits(sheet_name: &str, limits: StreamLimits) -> Result<Self> {
        let mut writer = Self::new_with_limits(sheet_name, limits)?;
        writer.ensure_temp_file()?;
        Ok(writer)
    }

    fn ensure_temp_file(&mut self) -> Result<&mut BufWriter<tempfile::NamedTempFile>> {
        if self.writer.is_none() {
            self.writer = Some(BufWriter::with_capacity(
                TEMP_FILE_BUFFER_BYTES,
                tempfile::NamedTempFile::new()?,
            ));
        }
        self.writer
            .as_mut()
            .ok_or_else(|| Error::Internal("stream writer initialization failed".to_string()))
    }

    /// Get the sheet name.
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Set column width for a single column (1-based).
    /// Must be called before any write_row() calls.
    pub fn set_col_width(&mut self, col: u32, width: f64) -> Result<()> {
        self.set_col_width_range(col, col, width)
    }

    /// Set column width for a range (min_col..=max_col), both 1-based.
    /// Must be called before any write_row() calls.
    pub fn set_col_width_range(&mut self, min_col: u32, max_col: u32, width: f64) -> Result<()> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        if self.started {
            return Err(Error::StreamColumnsAfterRows);
        }
        if min_col == 0 || max_col == 0 || min_col > MAX_COLUMNS || max_col > MAX_COLUMNS {
            return Err(Error::InvalidColumnNumber(if min_col == 0 {
                min_col
            } else {
                max_col
            }));
        }
        if let Some(existing) = self
            .col_widths
            .iter_mut()
            .find(|(min, max, _)| *min == min_col && *max == max_col)
        {
            existing.2 = width;
            return Ok(());
        }
        self.charge_config(1, std::mem::size_of::<(u32, u32, f64)>() as u64)?;
        self.col_widths.push((min_col, max_col, width));
        Ok(())
    }

    /// Set column style for a single column (1-based).
    /// Must be called before any write_row() calls.
    pub fn set_col_style(&mut self, col: u32, style_id: u32) -> Result<()> {
        self.ensure_col_configurable(col)?;
        self.get_or_create_col_options(col)?.style_id = Some(style_id);
        Ok(())
    }

    /// Set column visibility for a single column (1-based).
    /// Must be called before any write_row() calls.
    pub fn set_col_visible(&mut self, col: u32, visible: bool) -> Result<()> {
        self.ensure_col_configurable(col)?;
        self.get_or_create_col_options(col)?.hidden = Some(!visible);
        Ok(())
    }

    /// Set column outline level for a single column (1-based).
    /// Must be called before any write_row() calls. Level must be 0-7.
    pub fn set_col_outline_level(&mut self, col: u32, level: u8) -> Result<()> {
        self.ensure_col_configurable(col)?;
        if level > MAX_OUTLINE_LEVEL {
            return Err(Error::OutlineLevelExceeded {
                level,
                max: MAX_OUTLINE_LEVEL,
            });
        }
        self.get_or_create_col_options(col)?.outline_level = Some(level);
        Ok(())
    }

    /// Set freeze panes for the streamed sheet.
    /// Must be called before any write_row() calls.
    /// `top_left_cell` is the cell below and to the right of the frozen area,
    /// e.g., "A2" freezes row 1, "B1" freezes column A, "C3" freezes rows 1-2
    /// and columns A-B.
    pub fn set_freeze_panes(&mut self, top_left_cell: &str) -> Result<()> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        if self.started {
            return Err(Error::StreamColumnsAfterRows);
        }
        let (col, row) = cell_name_to_coordinates(top_left_cell)?;
        if col == 1 && row == 1 {
            return Err(Error::InvalidCellReference(
                "freeze pane at A1 has no effect".to_string(),
            ));
        }
        let additional_count = u64::from(self.freeze_pane.is_none());
        let previous_bytes = self
            .freeze_pane
            .as_ref()
            .map(|(_, _, value)| value.len() as u64)
            .unwrap_or_default();
        let next_bytes = top_left_cell.len() as u64;
        self.replace_config(additional_count, previous_bytes, next_bytes)?;
        self.freeze_pane = Some((col - 1, row - 1, top_left_cell.to_string()));
        Ok(())
    }

    /// Write a row of values with row-level options.
    pub fn write_row_with_options(
        &mut self,
        row: u32,
        values: &[CellValue],
        options: &StreamRowOptions,
    ) -> Result<()> {
        self.write_row_impl(row, values, None, Some(options))
    }

    /// Write a row of values. Rows must be written in ascending order.
    /// Row numbers are 1-based.
    pub fn write_row(&mut self, row: u32, values: &[CellValue]) -> Result<()> {
        self.write_row_impl(row, values, None, None)
    }

    /// Write multiple rows of values starting at the given row number.
    /// Each element in the slice becomes a row. Rows are numbered consecutively
    /// from `start_row`. All rows must be in ascending order relative to any
    /// previously written row.
    pub fn write_rows(&mut self, start_row: u32, rows: &[Vec<CellValue>]) -> Result<()> {
        for (i, values) in rows.iter().enumerate() {
            let offset = u32::try_from(i).map_err(|_| Error::InvalidRowNumber(u32::MAX))?;
            let row_num = start_row
                .checked_add(offset)
                .ok_or(Error::InvalidRowNumber(u32::MAX))?;
            self.validate_row_structure(row_num, values, None)?;
        }
        let next_formula_usage =
            self.preflight_formula_values(rows.iter().flat_map(|row| row.iter()))?;

        let mut next_row_count = self.row_count;
        let mut next_cell_count = self.cell_count;
        let mut next_bytes = self.bytes_written;
        for (i, values) in rows.iter().enumerate() {
            let offset = u32::try_from(i).map_err(|_| Error::InvalidRowNumber(u32::MAX))?;
            let row_num = start_row
                .checked_add(offset)
                .ok_or(Error::InvalidRowNumber(u32::MAX))?;
            self.preflight_row(
                row_num,
                values,
                None,
                None,
                &mut next_row_count,
                &mut next_cell_count,
                &mut next_bytes,
            )?;
        }
        self.formula_usage = next_formula_usage;
        let mut row_xml = self.take_row_buffer();
        let mut result = Ok(());
        for (i, values) in rows.iter().enumerate() {
            let offset = u32::try_from(i).map_err(|_| Error::InvalidRowNumber(u32::MAX))?;
            let row_num = start_row
                .checked_add(offset)
                .ok_or(Error::InvalidRowNumber(u32::MAX))?;
            // Every row was admitted above, so rendering needs no limit here.
            row_xml.clear();
            render_row_xml(&mut row_xml, row_num, values, None, None, u64::MAX)?;
            result = self.commit_row_xml(row_num, values.len(), &row_xml);
            if result.is_err() {
                break;
            }
        }
        self.restore_row_buffer(row_xml);
        result
    }

    /// Write a row with a specific style ID applied to all cells.
    pub fn write_row_with_style(
        &mut self,
        row: u32,
        values: &[CellValue],
        style_id: u32,
    ) -> Result<()> {
        self.write_row_impl(row, values, Some(style_id), None)
    }

    /// Add a merge cell reference (e.g., "A1:B2").
    ///
    /// Must be called before finish(). The reference must be a valid range
    /// in the form "A1:B2". It is stored in canonical form: absolute markers
    /// are dropped, letters are uppercased, and inverted endpoints are
    /// normalized to top-left:bottom-right. Re-adding an equivalent range is
    /// a no-op; a range overlapping a different existing merge returns
    /// [`Error::MergeCellOverlap`].
    pub fn add_merge_cell(&mut self, reference: &str) -> Result<()> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        let parts: Vec<&str> = reference.split(':').collect();
        if parts.len() != 2 {
            return Err(Error::InvalidMergeCellReference(reference.to_string()));
        }
        let (c1, r1) = cell_name_to_coordinates(parts[0])
            .map_err(|_| Error::InvalidMergeCellReference(reference.to_string()))?;
        let (c2, r2) = cell_name_to_coordinates(parts[1])
            .map_err(|_| Error::InvalidMergeCellReference(reference.to_string()))?;
        let range = (c1.min(c2), r1.min(r2), c1.max(c2), r1.max(r2));
        let canonical = format!(
            "{}:{}",
            coordinates_to_cell_name(range.0, range.1)?,
            coordinates_to_cell_name(range.2, range.3)?
        );
        for (existing, coords) in &self.merge_cells {
            if *coords == range {
                return Ok(());
            }
            if range.0 <= coords.2
                && range.2 >= coords.0
                && range.1 <= coords.3
                && range.3 >= coords.1
            {
                return Err(Error::MergeCellOverlap {
                    new: canonical,
                    existing: existing.clone(),
                });
            }
        }
        self.charge_config(1, u64::try_from(canonical.len()).unwrap_or(u64::MAX))?;
        self.merge_cells.push((canonical, range));
        Ok(())
    }

    /// Consume the writer and return a [`StreamedSheetData`] containing the
    /// temp file and pre-built XML fragments for the worksheet header/footer.
    pub fn into_streamed_data(mut self) -> Result<(String, StreamedSheetData)> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        if self.failed {
            return Err(Error::Internal(
                "stream writer is unusable after an output failure".to_string(),
            ));
        }
        self.finished = true;

        // Build XML fragments before consuming the writer.
        let sheet_views_xml = self.build_sheet_views_xml();
        let cols_xml = self.build_cols_xml();
        let merge_cells_xml = self.build_merge_cells_xml();
        let config_xml_bytes = sheet_views_xml
            .as_ref()
            .map(|value| value.len() as u64)
            .unwrap_or_default()
            .saturating_add(
                cols_xml
                    .as_ref()
                    .map(|value| value.len() as u64)
                    .unwrap_or_default(),
            )
            .saturating_add(
                merge_cells_xml
                    .as_ref()
                    .map(|value| value.len() as u64)
                    .unwrap_or_default(),
            )
            .saturating_add(512);
        checked_stream_add(
            self.bytes_written,
            config_xml_bytes,
            self.limits.output_bytes,
            ResourceDimension::OutputBytes,
        )?;
        let bytes_written = self.bytes_written;
        let sheet_name = self.sheet_name;

        // Flush the writer and recover the temp file.
        let mut writer = self
            .writer
            .map_or_else(|| tempfile::NamedTempFile::new().map(BufWriter::new), Ok)?;
        writer.flush()?;
        let temp_file = writer.into_inner().map_err(|e| Error::Io(e.into_error()))?;

        let data = StreamedSheetData {
            temp_file,
            data_len: bytes_written,
            row_count: self.row_count,
            cell_count: self.cell_count,
            sheet_views_xml,
            cols_xml,
            merge_cells_xml,
        };
        Ok((sheet_name, data))
    }

    /// Validate that column configuration is allowed (not finished, not started,
    /// valid column number).
    fn ensure_col_configurable(&self, col: u32) -> Result<()> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        if self.started {
            return Err(Error::StreamColumnsAfterRows);
        }
        if col == 0 || col > MAX_COLUMNS {
            return Err(Error::InvalidColumnNumber(col));
        }
        Ok(())
    }

    /// Get or create a StreamColOptions entry for the given column.
    fn get_or_create_col_options(&mut self, col: u32) -> Result<&mut StreamColOptions> {
        if let Some(pos) = self.col_options.iter().position(|(c, _)| *c == col) {
            Ok(&mut self.col_options[pos].1)
        } else {
            self.charge_config(
                1,
                u64::try_from(std::mem::size_of::<(u32, StreamColOptions)>()).unwrap_or(u64::MAX),
            )?;
            self.col_options.push((col, StreamColOptions::default()));
            let last = self.col_options.len() - 1;
            Ok(&mut self.col_options[last].1)
        }
    }

    fn charge_config(&mut self, count: u64, bytes: u64) -> Result<()> {
        let next_count = checked_stream_config_add(
            self.config_count,
            count,
            self.limits.config_count,
            ResourceDimension::CollectionCount,
        )?;
        let next_bytes = checked_stream_config_add(
            self.config_bytes,
            bytes,
            self.limits.config_bytes,
            ResourceDimension::RetainedBytes,
        )?;
        self.config_count = next_count;
        self.config_bytes = next_bytes;
        Ok(())
    }

    fn replace_config(
        &mut self,
        additional_count: u64,
        previous_bytes: u64,
        next_bytes: u64,
    ) -> Result<()> {
        let next_count = checked_stream_config_add(
            self.config_count,
            additional_count,
            self.limits.config_count,
            ResourceDimension::CollectionCount,
        )?;
        let retained = self
            .config_bytes
            .saturating_sub(previous_bytes)
            .saturating_add(next_bytes);
        if retained > self.limits.config_bytes {
            return Err(resource_limit(
                ResourceOperation::StreamWrite,
                ResourceCategory::Streaming,
                ResourceDimension::RetainedBytes,
                self.limits.config_bytes,
                retained,
                ResourceStage::Validate,
            ));
        }
        self.config_count = next_count;
        self.config_bytes = retained;
        Ok(())
    }

    /// Build the `<sheetViews>` XML fragment for freeze panes.
    fn build_sheet_views_xml(&self) -> Option<String> {
        let (x_split, y_split, top_left_cell) = self.freeze_pane.as_ref()?;
        let active_pane = match (*x_split > 0, *y_split > 0) {
            (true, true) => "bottomRight",
            (true, false) => "topRight",
            (false, true) => "bottomLeft",
            (false, false) => unreachable!(),
        };

        let sheet_views = SheetViews {
            sheet_views: vec![SheetView {
                tab_selected: Some(true),
                show_grid_lines: None,
                show_formulas: None,
                show_row_col_headers: None,
                zoom_scale: None,
                view: None,
                top_left_cell: None,
                workbook_view_id: 0,
                pane: Some(Pane {
                    x_split: if *x_split > 0 { Some(*x_split) } else { None },
                    y_split: if *y_split > 0 { Some(*y_split) } else { None },
                    top_left_cell: Some(top_left_cell.clone()),
                    active_pane: Some(active_pane.to_string()),
                    state: Some("frozen".to_string()),
                }),
                selection: vec![],
            }],
        };

        quick_xml::se::to_string_with_root("sheetViews", &sheet_views).ok()
    }

    /// Build the `<cols>` XML fragment.
    fn build_cols_xml(&self) -> Option<String> {
        let has_widths = !self.col_widths.is_empty();
        let has_options = !self.col_options.is_empty();
        if !has_widths && !has_options {
            return None;
        }

        let mut col_defs = Vec::new();
        for &(min, max, width) in &self.col_widths {
            col_defs.push(Col {
                min,
                max,
                width: Some(width),
                style: None,
                hidden: None,
                custom_width: Some(true),
                outline_level: None,
            });
        }
        for &(col_num, ref opts) in &self.col_options {
            col_defs.push(Col {
                min: col_num,
                max: col_num,
                width: opts.width,
                style: opts.style_id,
                hidden: opts.hidden,
                custom_width: if opts.width.is_some() {
                    Some(true)
                } else {
                    None
                },
                outline_level: opts.outline_level.filter(|&l| l > 0),
            });
        }

        let cols = Cols { cols: col_defs };
        quick_xml::se::to_string_with_root("cols", &cols).ok()
    }

    /// Build the `<mergeCells>` XML fragment.
    fn build_merge_cells_xml(&self) -> Option<String> {
        if self.merge_cells.is_empty() {
            return None;
        }

        let mc = MergeCells {
            count: Some(self.merge_cells.len() as u32),
            merge_cells: self
                .merge_cells
                .iter()
                .map(|(r, _)| MergeCell {
                    reference: r.clone(),
                })
                .collect(),
            cached_coords: Vec::new(),
        };
        quick_xml::se::to_string_with_root("mergeCells", &mc).ok()
    }

    /// Internal unified implementation for writing a row.
    /// Renders the row XML once into a reusable buffer, admits it against
    /// every stream limit, then writes it to the temp file.
    fn write_row_impl(
        &mut self,
        row: u32,
        values: &[CellValue],
        cell_style_id: Option<u32>,
        options: Option<&StreamRowOptions>,
    ) -> Result<()> {
        self.validate_row_structure(row, values, options)?;
        let next_formula_usage = self.preflight_formula_values(values.iter())?;
        validate_stream_values(values)?;

        let mut row_xml = self.take_row_buffer();
        let result = render_row_xml(
            &mut row_xml,
            row,
            values,
            cell_style_id,
            options,
            self.limits.row_xml_bytes,
        )
        .and_then(|row_xml_bytes| {
            let mut next_row_count = self.row_count;
            let mut next_cell_count = self.cell_count;
            let mut next_bytes = self.bytes_written;
            self.charge_row_totals(
                values,
                row_xml_bytes,
                &mut next_row_count,
                &mut next_cell_count,
                &mut next_bytes,
            )
        })
        .and_then(|()| {
            self.formula_usage = next_formula_usage;
            self.commit_row_xml(row, values.len(), &row_xml)
        });
        self.restore_row_buffer(row_xml);
        result
    }

    fn take_row_buffer(&mut self) -> String {
        let mut buffer = std::mem::take(&mut self.row_buf);
        buffer.clear();
        buffer
    }

    fn restore_row_buffer(&mut self, buffer: String) {
        if buffer.capacity() <= RETAINED_ROW_BUFFER_BYTES {
            self.row_buf = buffer;
        }
    }

    fn preflight_formula_values<'a>(
        &self,
        values: impl IntoIterator<Item = &'a CellValue>,
    ) -> Result<StreamFormulaUsage> {
        let budget = crate::formula::FormulaBudget::secure_mutation();
        let mut formula_cells = 0u64;
        let mut formula_input_bytes = 0u64;
        let mut formula_output_bytes = 0u64;
        for value in values {
            collect_formula_usage(
                value,
                &budget,
                &mut formula_cells,
                &mut formula_input_bytes,
                &mut formula_output_bytes,
            )?;
        }
        let next = StreamFormulaUsage {
            cells: self.formula_usage.cells.saturating_add(formula_cells),
            input_bytes: self
                .formula_usage
                .input_bytes
                .saturating_add(formula_input_bytes),
            output_bytes: self
                .formula_usage
                .output_bytes
                .saturating_add(formula_output_bytes),
        };
        budget.charge_mutation_cells(next.cells)?;
        budget.charge_mutation_input_total(next.input_bytes)?;
        budget.charge_mutation_output(next.output_bytes)?;
        Ok(next)
    }

    #[allow(clippy::too_many_arguments)]
    fn preflight_row(
        &self,
        row: u32,
        values: &[CellValue],
        cell_style_id: Option<u32>,
        options: Option<&StreamRowOptions>,
        row_count: &mut u64,
        cell_count: &mut u64,
        output_bytes: &mut u64,
    ) -> Result<()> {
        self.validate_row_structure(row, values, options)?;
        validate_stream_values(values)?;
        let row_xml_bytes = measure_row_xml(
            row,
            values,
            cell_style_id,
            options,
            self.limits.row_xml_bytes,
        )?;
        self.charge_row_totals(values, row_xml_bytes, row_count, cell_count, output_bytes)
    }

    /// Checks the cumulative row, cell, and byte limits for one admitted row.
    fn charge_row_totals(
        &self,
        values: &[CellValue],
        row_xml_bytes: u64,
        row_count: &mut u64,
        cell_count: &mut u64,
        output_bytes: &mut u64,
    ) -> Result<()> {
        *row_count = checked_stream_add(
            *row_count,
            1,
            self.limits.row_count,
            ResourceDimension::RowCount,
        )?;
        *cell_count = checked_stream_add(
            *cell_count,
            u64::try_from(values.len()).unwrap_or(u64::MAX),
            self.limits.cell_count,
            ResourceDimension::CellCount,
        )?;
        *output_bytes = checked_stream_add(
            *output_bytes,
            row_xml_bytes,
            self.limits.output_bytes,
            ResourceDimension::OutputBytes,
        )?;
        Ok(())
    }

    fn validate_row_structure(
        &self,
        row: u32,
        values: &[CellValue],
        options: Option<&StreamRowOptions>,
    ) -> Result<()> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        if self.failed {
            return Err(Error::Internal(
                "stream writer is unusable after an output failure".to_string(),
            ));
        }

        // Validate row number.
        if row == 0 || row > MAX_ROWS {
            return Err(Error::InvalidRowNumber(row));
        }

        // Enforce ascending row order.
        if row <= self.last_row {
            return Err(Error::StreamRowAlreadyWritten { row });
        }

        // Validate column count.
        if values.len() > MAX_COLUMNS as usize {
            return Err(Error::InvalidColumnNumber(values.len() as u32));
        }

        // Validate row options if provided.
        if let Some(opts) = options {
            if let Some(height) = opts.height {
                if height > MAX_ROW_HEIGHT {
                    return Err(Error::RowHeightExceeded {
                        height,
                        max: MAX_ROW_HEIGHT,
                    });
                }
            }
            if let Some(level) = opts.outline_level {
                if level > MAX_OUTLINE_LEVEL {
                    return Err(Error::OutlineLevelExceeded {
                        level,
                        max: MAX_OUTLINE_LEVEL,
                    });
                }
            }
        }

        Ok(())
    }

    /// Writes one admitted row's XML to the temp file and records it.
    fn commit_row_xml(&mut self, row: u32, cell_slots: usize, row_xml: &str) -> Result<()> {
        let failure_after = self.write_failure_after_bytes;
        let bytes = row_xml.as_bytes();
        let writer = self.ensure_temp_file()?;
        let written = match failure_after {
            Some(failure_after)
                if u64::try_from(bytes.len()).unwrap_or(u64::MAX) > failure_after =>
            {
                let prefix = usize::try_from(failure_after)
                    .unwrap_or(usize::MAX)
                    .min(bytes.len());
                if prefix > 0 {
                    let _ = writer.write_all(&bytes[..prefix]);
                }
                Err(std::io::Error::other("injected stream output failure"))
            }
            _ => writer.write_all(bytes),
        };
        if let Err(error) = written {
            self.failed = true;
            return Err(Error::Io(error));
        }
        let row_bytes = u64::try_from(bytes.len()).unwrap_or(u64::MAX);

        self.row_count = self
            .row_count
            .checked_add(1)
            .ok_or_else(|| Error::Internal("stream row count overflow".to_string()))?;
        self.cell_count = self
            .cell_count
            .checked_add(u64::try_from(cell_slots).unwrap_or(u64::MAX))
            .ok_or_else(|| Error::Internal("stream cell count overflow".to_string()))?;
        self.bytes_written = self
            .bytes_written
            .checked_add(row_bytes)
            .ok_or_else(|| Error::Internal("stream byte count overflow".to_string()))?;
        self.started = true;
        self.last_row = row;
        Ok(())
    }
}

/// Build the XML for the worksheet opening, including the XML declaration,
/// worksheet namespace, sheetViews, cols, and the opening `<sheetData>` tag.
pub(crate) fn build_worksheet_header(streamed: &StreamedSheetData) -> String {
    let mut xml = String::with_capacity(512);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');
    xml.push_str(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );

    if let Some(ref views) = streamed.sheet_views_xml {
        xml.push_str(views);
    }
    if let Some(ref cols) = streamed.cols_xml {
        xml.push_str(cols);
    }

    xml.push_str("<sheetData>");
    xml
}

/// Build the XML for the worksheet closing: `</sheetData>`, mergeCells, and
/// `</worksheet>`.
pub(crate) fn build_worksheet_footer(streamed: &StreamedSheetData) -> String {
    let mut xml = String::with_capacity(256);
    xml.push_str("</sheetData>");

    if let Some(ref mc) = streamed.merge_cells_xml {
        xml.push_str(mc);
    }

    xml.push_str("</worksheet>");
    xml
}

/// Write streamed sheet data to a ZIP writer. Composes the full worksheet
/// XML by writing the header, streaming rows from the temp file, and
/// appending the footer.
pub(crate) fn write_streamed_sheet<W: std::io::Write + std::io::Seek>(
    zip: &mut zip::ZipWriter<W>,
    entry_name: &str,
    streamed: &StreamedSheetData,
    options: zip::write::SimpleFileOptions,
    xml_limit: u64,
) -> Result<()> {
    let header = build_worksheet_header(streamed);
    let footer = build_worksheet_footer(streamed);
    let observed = u64::try_from(header.len())
        .unwrap_or(u64::MAX)
        .saturating_add(streamed.data_len)
        .saturating_add(u64::try_from(footer.len()).unwrap_or(u64::MAX));
    if observed > xml_limit {
        return Err(resource_limit(
            ResourceOperation::WorkbookSave,
            ResourceCategory::Output,
            ResourceDimension::XmlBytes,
            xml_limit,
            observed,
            ResourceStage::Serialize,
        ));
    }
    zip.start_file(entry_name, options)
        .map_err(|e| Error::Zip(e.to_string()))?;

    // Write the header (XML declaration, <worksheet>, <sheetViews>, <cols>, <sheetData>).
    zip.write_all(header.as_bytes())?;

    // Open the temp file by path for reading (starts at position 0).
    let file = std::fs::File::open(streamed.temp_file.path())?;
    let mut reader = std::io::BufReader::new(file);
    std::io::copy(&mut reader, zip)?;

    // Write the footer (</sheetData>, <mergeCells>, </worksheet>).
    zip.write_all(footer.as_bytes())?;

    Ok(())
}

/// Destination for rendered row XML that enforces the row byte limit.
///
/// `push_str` charges its argument as one unit and `push_chars` charges each
/// character separately, so the size reported on rejection does not depend on
/// how plain text is grouped.
trait RowXmlSink {
    fn push_str(&mut self, value: &str);
    fn push_chars(&mut self, value: &str);

    fn push(&mut self, value: char) {
        let mut encoded = [0u8; 4];
        self.push_str(value.encode_utf8(&mut encoded));
    }
}

/// Running size of one row's XML measured against the row limit.
struct RowXmlCounter {
    written: u64,
    limit: u64,
    exceeded: bool,
}

impl RowXmlCounter {
    fn new(limit: u64) -> Self {
        Self {
            written: 0,
            limit,
            exceeded: false,
        }
    }

    /// Charges `amount` bytes as one unit and returns whether they fit.
    fn charge(&mut self, amount: usize) -> bool {
        if self.exceeded {
            return false;
        }
        let observed = self
            .written
            .saturating_add(u64::try_from(amount).unwrap_or(u64::MAX));
        self.written = observed;
        if observed > self.limit {
            self.exceeded = true;
            return false;
        }
        true
    }

    /// Charges `value` one character at a time and returns whether it fits.
    fn charge_chars(&mut self, value: &str) -> bool {
        if self.exceeded {
            return false;
        }
        let remaining = self.limit.saturating_sub(self.written);
        let length = u64::try_from(value.len()).unwrap_or(u64::MAX);
        if length <= remaining {
            self.written = self.written.saturating_add(length);
            return true;
        }
        // Record the size at the first character that crosses the limit.
        let crossing = value
            .char_indices()
            .map(|(start, ch)| u64::try_from(start + ch.len_utf8()).unwrap_or(u64::MAX))
            .find(|end| *end > remaining)
            .unwrap_or(length);
        self.written = self.written.saturating_add(crossing);
        self.exceeded = true;
        false
    }

    /// Returns the admitted row size, or the row XML limit error.
    fn finish(&self) -> Result<u64> {
        if self.exceeded {
            return Err(resource_limit(
                ResourceOperation::StreamWrite,
                ResourceCategory::Streaming,
                ResourceDimension::XmlBytes,
                self.limit,
                self.written,
                ResourceStage::Serialize,
            ));
        }
        Ok(self.written)
    }
}

impl RowXmlSink for RowXmlCounter {
    fn push_str(&mut self, value: &str) {
        self.charge(value.len());
    }

    fn push_chars(&mut self, value: &str) {
        self.charge_chars(value);
    }
}

/// Row XML sink that keeps the admitted bytes in memory.
struct RowXmlBuffer<'a> {
    counter: RowXmlCounter,
    out: &'a mut String,
}

impl RowXmlSink for RowXmlBuffer<'_> {
    fn push_str(&mut self, value: &str) {
        if self.counter.charge(value.len()) {
            self.out.push_str(value);
        }
    }

    fn push_chars(&mut self, value: &str) {
        if self.counter.charge_chars(value) {
            self.out.push_str(value);
        }
    }

    fn push(&mut self, value: char) {
        if self.counter.charge(value.len_utf8()) {
            self.out.push(value);
        }
    }
}

fn measure_row_xml(
    row: u32,
    values: &[CellValue],
    cell_style_id: Option<u32>,
    options: Option<&StreamRowOptions>,
    limit: u64,
) -> Result<u64> {
    let mut counter = RowXmlCounter::new(limit);
    write_row_xml(&mut counter, row, values, cell_style_id, options);
    counter.finish()
}

/// Render one row's XML into `out`, rejecting rows above `limit` bytes.
/// Returns the rendered size, which equals the bytes appended to `out`.
fn render_row_xml(
    out: &mut String,
    row: u32,
    values: &[CellValue],
    cell_style_id: Option<u32>,
    options: Option<&StreamRowOptions>,
    limit: u64,
) -> Result<u64> {
    let mut sink = RowXmlBuffer {
        counter: RowXmlCounter::new(limit),
        out,
    };
    write_row_xml(&mut sink, row, values, cell_style_id, options);
    sink.counter.finish()
}

fn checked_stream_add(
    current: u64,
    additional: u64,
    limit: u64,
    dimension: ResourceDimension,
) -> Result<u64> {
    let observed = current.saturating_add(additional);
    if observed > limit {
        return Err(resource_limit(
            ResourceOperation::StreamWrite,
            ResourceCategory::Streaming,
            dimension,
            limit,
            observed,
            ResourceStage::Serialize,
        ));
    }
    Ok(observed)
}

fn checked_stream_config_add(
    current: u64,
    additional: u64,
    limit: u64,
    dimension: ResourceDimension,
) -> Result<u64> {
    let observed = current.saturating_add(additional);
    if observed > limit {
        return Err(resource_limit(
            ResourceOperation::StreamWrite,
            ResourceCategory::Streaming,
            dimension,
            limit,
            observed,
            ResourceStage::Validate,
        ));
    }
    Ok(observed)
}

fn collect_formula_usage(
    value: &CellValue,
    budget: &crate::formula::FormulaBudget,
    formula_cells: &mut u64,
    formula_input_bytes: &mut u64,
    formula_output_bytes: &mut u64,
) -> Result<()> {
    if let CellValue::Formula { expr, result } = value {
        let bytes = u64::try_from(expr.len()).unwrap_or(u64::MAX);
        budget.check_mutation_input(bytes)?;
        *formula_cells = formula_cells.saturating_add(1);
        *formula_input_bytes = formula_input_bytes.saturating_add(bytes);
        *formula_output_bytes = formula_output_bytes.saturating_add(bytes);
        if let Some(result) = result {
            *formula_output_bytes =
                formula_output_bytes.saturating_add(cached_formula_payload_bytes(result)?);
        }
    }
    Ok(())
}

fn cached_formula_payload_bytes(value: &CellValue) -> Result<u64> {
    Ok(match value {
        CellValue::String(value) | CellValue::Error(value) => {
            u64::try_from(value.len()).unwrap_or(u64::MAX)
        }
        CellValue::RichString(_) | CellValue::Formula { .. } => {
            return Err(resource_limit(
                ResourceOperation::StreamWrite,
                ResourceCategory::Structure,
                ResourceDimension::CollectionCount,
                0,
                1,
                ResourceStage::Validate,
            ));
        }
        CellValue::Number(_) | CellValue::Date(_) => 32,
        CellValue::Bool(_) => 5,
        CellValue::Empty => 0,
    })
}

/// Write an XML `<row>` element with inline strings to a pre-admitted sink.
fn write_row_xml<W: RowXmlSink>(
    xml: &mut W,
    row: u32,
    values: &[CellValue],
    cell_style_id: Option<u32>,
    options: Option<&StreamRowOptions>,
) {
    // Row number and style are formatted once per row; numbers and cell
    // references reuse one scratch buffer instead of allocating per cell.
    let row_text = row.to_string();
    let style_text = cell_style_id.map(|sid| sid.to_string());
    let mut scratch = String::new();

    xml.push_str("<row r=\"");
    xml.push_str(&row_text);
    xml.push('"');

    // Row-level attributes from options.
    if let Some(opts) = options {
        if let Some(sid) = opts.style_id {
            xml.push_str(" s=\"");
            xml.push_str(&sid.to_string());
            xml.push_str("\" customFormat=\"1\"");
        }
        if let Some(height) = opts.height {
            xml.push_str(" ht=\"");
            xml.push_str(&height.to_string());
            xml.push_str("\" customHeight=\"1\"");
        }
        if let Some(false) = opts.visible {
            xml.push_str(" hidden=\"1\"");
        }
        if let Some(level) = opts.outline_level {
            if level > 0 {
                xml.push_str(" outlineLevel=\"");
                xml.push_str(&level.to_string());
                xml.push('"');
            }
        }
    }

    xml.push('>');

    for (i, value) in values.iter().enumerate() {
        let col = (i as u32) + 1;
        if matches!(value, CellValue::Empty) {
            continue;
        }

        scratch.clear();
        push_cell_name(&mut scratch, col, row, &row_text);

        xml.push_str("<c r=\"");
        xml.push_str(&scratch);
        xml.push('"');

        // Style attribute.
        if let Some(sid) = &style_text {
            xml.push_str(" s=\"");
            xml.push_str(sid);
            xml.push('"');
        }

        match value {
            CellValue::String(s) => {
                xml.push_str(" t=\"inlineStr\"><is><t>");
                xml_escape_into(xml, s);
                xml.push_str("</t></is></c>");
            }
            CellValue::Number(n) => {
                xml.push_str("><v>");
                push_display(xml, &mut scratch, n);
                xml.push_str("</v></c>");
            }
            CellValue::Date(serial) => {
                xml.push_str(" t=\"d\"><v>");
                if let Some(datetime) = crate::cell::serial_to_datetime(*serial) {
                    xml.push_str(&datetime.format("%Y-%m-%dT%H:%M:%S").to_string());
                } else {
                    xml.push_str(&serial.to_string());
                }
                xml.push_str("</v></c>");
            }
            CellValue::Bool(b) => {
                xml.push_str(" t=\"b\"><v>");
                xml.push_str(if *b { "1" } else { "0" });
                xml.push_str("</v></c>");
            }
            CellValue::Formula { expr, result } => {
                // Set the cell type attribute based on the cached result type.
                // Without the correct t attribute, xml_cell_to_value cannot
                // decode the result on reopen.
                if let Some(res) = result {
                    match res.as_ref() {
                        CellValue::String(_) => xml.push_str(" t=\"str\""),
                        CellValue::Bool(_) => xml.push_str(" t=\"b\""),
                        CellValue::Error(_) => xml.push_str(" t=\"e\""),
                        CellValue::Date(_) => xml.push_str(" t=\"d\""),
                        _ => {} // Number: default (no t) is numeric
                    }
                }
                xml.push_str("><f>");
                xml_escape_into(xml, expr);
                xml.push_str("</f>");
                if let Some(res) = result {
                    match res.as_ref() {
                        CellValue::String(s) => {
                            xml.push_str("<v>");
                            xml_escape_into(xml, s);
                            xml.push_str("</v>");
                        }
                        CellValue::Number(n) => {
                            xml.push_str("<v>");
                            push_display(xml, &mut scratch, n);
                            xml.push_str("</v>");
                        }
                        CellValue::Bool(b) => {
                            xml.push_str("<v>");
                            xml.push_str(if *b { "1" } else { "0" });
                            xml.push_str("</v>");
                        }
                        CellValue::Date(d) => {
                            xml.push_str("<v>");
                            if let Some(datetime) = crate::cell::serial_to_datetime(*d) {
                                xml.push_str(&datetime.format("%Y-%m-%dT%H:%M:%S").to_string());
                            } else {
                                xml.push_str(&d.to_string());
                            }
                            xml.push_str("</v>");
                        }
                        CellValue::Error(e) => {
                            xml.push_str("<v>");
                            xml_escape_into(xml, e);
                            xml.push_str("</v>");
                        }
                        _ => {}
                    }
                }
                xml.push_str("</c>");
            }
            CellValue::Error(e) => {
                xml.push_str(" t=\"e\"><v>");
                xml_escape_into(xml, e);
                xml.push_str("</v></c>");
            }
            CellValue::RichString(runs) => {
                xml.push_str(" t=\"inlineStr\"><is><t>");
                for run in runs {
                    xml_escape_into(xml, &run.text);
                }
                xml.push_str("</t></is></c>");
            }
            CellValue::Empty => unreachable!(),
        }
    }

    xml.push_str("</row>");
}

/// Escape XML special characters into an existing string buffer, dropping
/// characters XML cannot represent. Unchanged text is pushed in runs.
fn xml_escape_into<W: RowXmlSink>(buf: &mut W, s: &str) {
    let mut run_start = 0;
    for (index, ch) in s.char_indices() {
        let replacement = match ch {
            '&' => "&amp;",
            '<' => "&lt;",
            '>' => "&gt;",
            '"' => "&quot;",
            '\'' => "&apos;",
            _ if !crate::utils::is_xml_char(ch) => "",
            _ => continue,
        };
        if run_start < index {
            buf.push_chars(&s[run_start..index]);
        }
        if !replacement.is_empty() {
            buf.push_str(replacement);
        }
        run_start = index + ch.len_utf8();
    }
    if run_start < s.len() {
        buf.push_chars(&s[run_start..]);
    }
}

/// Append the A1 name of (`col`, `row`) to `out`, reusing the row's
/// formatted number. Matches `coordinates_to_cell_name`, including its
/// fallback text for coordinates outside the sheet.
fn push_cell_name(out: &mut String, col: u32, row: u32, row_text: &str) {
    if !(1..=MAX_COLUMNS).contains(&col) || !(1..=MAX_ROWS).contains(&row) {
        let _ = write!(out, "{col}:{row}");
        return;
    }
    // MAX_COLUMNS (XFD) needs at most three letters.
    let mut letters = [0u8; 3];
    let mut length = 0;
    let mut remaining = col;
    while remaining > 0 {
        remaining -= 1;
        letters[length] = b'A' + (remaining % 26) as u8;
        length += 1;
        remaining /= 26;
    }
    for &letter in letters[..length].iter().rev() {
        out.push(char::from(letter));
    }
    out.push_str(row_text);
}

/// Push `value` formatted with `Display` as one token, using `scratch`.
fn push_display<W: RowXmlSink>(xml: &mut W, scratch: &mut String, value: impl std::fmt::Display) {
    scratch.clear();
    let _ = write!(scratch, "{value}");
    xml.push_str(scratch);
}

fn validate_stream_values(values: &[CellValue]) -> Result<()> {
    for value in values {
        validate_stream_value(value)?;
    }
    Ok(())
}

fn validate_stream_value(value: &CellValue) -> Result<()> {
    match value {
        CellValue::Number(n) | CellValue::Date(n) if !n.is_finite() => {
            Err(Error::StreamNonFiniteValue)
        }
        CellValue::Formula {
            result: Some(result),
            ..
        } => match result.as_ref() {
            CellValue::Formula { .. } => Err(resource_limit(
                ResourceOperation::StreamWrite,
                ResourceCategory::Structure,
                ResourceDimension::NestingDepth,
                1,
                2,
                ResourceStage::Validate,
            )),
            other => validate_stream_value(other),
        },
        _ => Ok(()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetkit_xml::worksheet::WorksheetXml;
    use std::io::{Read as _, Seek as _};

    /// Helper to parse a streamed writer into a WorksheetXml for assertion.
    fn finish_and_parse(sw: StreamWriter) -> WorksheetXml {
        let (_, mut streamed) = sw.into_streamed_data().unwrap();
        let header = build_worksheet_header(&streamed);
        let footer = build_worksheet_footer(&streamed);

        let file = streamed.temp_file.as_file_mut();
        file.seek(std::io::SeekFrom::Start(0)).unwrap();
        let mut row_data = String::new();
        file.read_to_string(&mut row_data).unwrap();

        let full_xml = format!("{header}{row_data}{footer}");
        quick_xml::de::from_str(&full_xml).unwrap()
    }

    /// Helper to get the raw XML from a stream writer.
    fn finish_and_get_xml(sw: StreamWriter) -> String {
        let (_, mut streamed) = sw.into_streamed_data().unwrap();
        let header = build_worksheet_header(&streamed);
        let footer = build_worksheet_footer(&streamed);

        let file = streamed.temp_file.as_file_mut();
        file.seek(std::io::SeekFrom::Start(0)).unwrap();
        let mut row_data = String::new();
        file.read_to_string(&mut row_data).unwrap();

        format!("{header}{row_data}{footer}")
    }

    #[test]
    fn test_new_defers_temp_file_creation() {
        let sw = StreamWriter::new("Sheet1");

        assert!(sw.writer.is_none());
    }

    #[test]
    fn test_try_new_initializes_temp_file_fallibly() {
        let sw = StreamWriter::try_new("Sheet1").unwrap();

        assert!(sw.writer.is_some());
    }

    #[test]
    fn test_basic_write_and_finish() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Name"), CellValue::from("Age")])
            .unwrap();
        sw.write_row(2, &[CellValue::from("Alice"), CellValue::from(30)])
            .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("<?xml version=\"1.0\""));
        assert!(xml.contains("<worksheet"));
        assert!(xml.contains("<sheetData>"));
        assert!(xml.contains("</sheetData>"));
        assert!(xml.contains("</worksheet>"));
    }

    #[test]
    fn test_parse_output_xml_back() {
        let mut sw = StreamWriter::new("TestSheet");
        sw.write_row(1, &[CellValue::from("Hello"), CellValue::from(42)])
            .unwrap();
        let ws = finish_and_parse(sw);

        assert_eq!(ws.sheet_data.rows.len(), 1);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[0].cells.len(), 2);
        // First cell is an inline string (t="inlineStr")
        assert_eq!(
            ws.sheet_data.rows[0].cells[0].t,
            sheetkit_xml::worksheet::CellTypeTag::InlineString
        );
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");
        // Second cell is a number
        assert_eq!(
            ws.sheet_data.rows[0].cells[1].t,
            sheetkit_xml::worksheet::CellTypeTag::None
        );
        assert_eq!(ws.sheet_data.rows[0].cells[1].v, Some("42".to_string()));
        assert_eq!(ws.sheet_data.rows[0].cells[1].r, "B1");
    }

    #[test]
    fn test_inline_strings_used() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Hello"), CellValue::from("World")])
            .unwrap();
        let xml = finish_and_get_xml(sw);

        // Should use inline strings, not shared string references
        assert!(xml.contains("t=\"inlineStr\""));
        assert!(xml.contains("<is><t>Hello</t></is>"));
        assert!(xml.contains("<is><t>World</t></is>"));
        // Should NOT contain shared string references
        assert!(!xml.contains("t=\"s\""));
    }

    #[test]
    fn test_number_value() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from(3.15)]).unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("<v>3.15</v>"));
    }

    #[test]
    fn test_bool_values() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from(true), CellValue::from(false)])
            .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("<v>1</v>"));
        assert!(xml.contains("<v>0</v>"));
    }

    #[test]
    fn test_formula_value() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::Formula {
                expr: "SUM(A2:A10)".to_string(),
                result: None,
            }],
        )
        .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("SUM(A2:A10)"));
    }

    #[test]
    fn test_error_value() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::Error("#DIV/0!".to_string())])
            .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("#DIV/0!"));
    }

    #[test]
    fn test_empty_values_are_skipped() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::from("A"), CellValue::Empty, CellValue::from("C")],
        )
        .unwrap();
        let ws = finish_and_parse(sw);

        // 2 cells (Empty is skipped)
        assert_eq!(ws.sheet_data.rows[0].cells.len(), 2);
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");
        assert_eq!(ws.sheet_data.rows[0].cells[1].r, "C1");
    }

    #[test]
    fn test_write_row_with_style() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row_with_style(1, &[CellValue::from("Styled")], 5)
            .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("s=\"5\""));
    }

    #[test]
    fn test_set_col_width_before_rows() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width(1, 20.0).unwrap();
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let ws = finish_and_parse(sw);

        let cols = ws.cols.unwrap();
        assert_eq!(cols.cols.len(), 1);
        assert_eq!(cols.cols[0].min, 1);
        assert_eq!(cols.cols[0].max, 1);
        assert_eq!(cols.cols[0].width, Some(20.0));
        assert_eq!(cols.cols[0].custom_width, Some(true));
    }

    #[test]
    fn test_set_col_width_range() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width_range(1, 3, 15.5).unwrap();
        let ws = finish_and_parse(sw);

        let cols = ws.cols.unwrap();
        assert_eq!(cols.cols.len(), 1);
        assert_eq!(cols.cols[0].min, 1);
        assert_eq!(cols.cols[0].max, 3);
        assert_eq!(cols.cols[0].width, Some(15.5));
    }

    #[test]
    fn test_col_width_in_output_xml() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width(2, 25.0).unwrap();
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let xml = finish_and_get_xml(sw);

        // Verify cols section appears before sheetData
        let cols_pos = xml.find("<cols>").unwrap();
        let sheet_data_pos = xml.find("<sheetData").unwrap();
        assert!(cols_pos < sheet_data_pos);
    }

    #[test]
    fn test_col_width_after_rows_returns_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let result = sw.set_col_width(1, 20.0);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamColumnsAfterRows));
    }

    #[test]
    fn test_rows_in_order_succeeds() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        sw.write_row(2, &[CellValue::from("b")]).unwrap();
        sw.write_row(3, &[CellValue::from("c")]).unwrap();
        let _data = sw.into_streamed_data().unwrap();
    }

    #[test]
    fn test_rows_with_gaps_succeeds() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        sw.write_row(3, &[CellValue::from("b")]).unwrap();
        sw.write_row(5, &[CellValue::from("c")]).unwrap();
        let ws = finish_and_parse(sw);

        assert_eq!(ws.sheet_data.rows.len(), 3);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[1].r, 3);
        assert_eq!(ws.sheet_data.rows[2].r, 5);
    }

    #[test]
    fn test_duplicate_row_number_fails() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        let result = sw.write_row(1, &[CellValue::from("b")]);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::StreamRowAlreadyWritten { row: 1 }
        ));
    }

    #[test]
    fn test_row_zero_fails() {
        let mut sw = StreamWriter::new("Sheet1");
        let result = sw.write_row(0, &[CellValue::from("a")]);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::InvalidRowNumber(0)));
    }

    #[test]
    fn test_row_out_of_order_fails() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(5, &[CellValue::from("a")]).unwrap();
        let result = sw.write_row(3, &[CellValue::from("b")]);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::StreamRowAlreadyWritten { row: 3 }
        ));
    }

    #[test]
    fn test_merge_cells_in_output() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Merged")]).unwrap();
        sw.add_merge_cell("A1:B1").unwrap();
        let ws = finish_and_parse(sw);

        let mc = ws.merge_cells.unwrap();
        assert_eq!(mc.merge_cells.len(), 1);
        assert_eq!(mc.merge_cells[0].reference, "A1:B1");
    }

    #[test]
    fn test_multiple_merge_cells() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        sw.add_merge_cell("A1:B1").unwrap();
        sw.add_merge_cell("C1:D1").unwrap();
        let ws = finish_and_parse(sw);

        let mc = ws.merge_cells.unwrap();
        assert_eq!(mc.merge_cells.len(), 2);
        assert_eq!(mc.merge_cells[0].reference, "A1:B1");
        assert_eq!(mc.merge_cells[1].reference, "C1:D1");
    }

    #[test]
    fn test_finish_twice_returns_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        // First consume
        sw.into_streamed_data().unwrap();
        // Cannot call again -- consumed by move.
    }

    #[test]
    fn test_write_after_finish_returns_error() {
        // StreamWriter is consumed by into_streamed_data(), so this test
        // verifies that the finished flag works for the legacy path.
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        sw.finished = true;
        let result = sw.write_row(2, &[CellValue::from("b")]);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamAlreadyFinished));
    }

    #[test]
    fn test_finish_with_no_rows() {
        let sw = StreamWriter::new("Sheet1");
        let ws = finish_and_parse(sw);

        // Should still produce valid XML
        assert!(ws.sheet_data.rows.is_empty());
    }

    #[test]
    fn test_finish_with_cols_and_no_rows() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width(1, 20.0).unwrap();
        let ws = finish_and_parse(sw);

        assert!(ws.cols.is_some());
        assert!(ws.sheet_data.rows.is_empty());
    }

    #[test]
    fn test_add_merge_cell_invalid_reference() {
        let mut sw = StreamWriter::new("Sheet1");
        // Missing colon separator.
        let result = sw.add_merge_cell("A1B2");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::InvalidMergeCellReference(_)
        ));
    }

    #[test]
    fn test_add_merge_cell_invalid_cell_name() {
        let mut sw = StreamWriter::new("Sheet1");
        // Invalid cell name before colon.
        let result = sw.add_merge_cell("ZZZ:B2");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::InvalidMergeCellReference(_)
        ));
    }

    #[test]
    fn test_add_merge_cell_empty_reference() {
        let mut sw = StreamWriter::new("Sheet1");
        let result = sw.add_merge_cell("");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::InvalidMergeCellReference(_)
        ));
    }

    #[test]
    fn test_add_merge_cell_after_finish_fails() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.finished = true;
        let result = sw.add_merge_cell("A1:B1");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamAlreadyFinished));
    }

    #[test]
    fn add_merge_cell_canonicalizes_absolute_lowercase_and_inverted_ranges() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.add_merge_cell("$a$1:b2").unwrap();
        sw.add_merge_cell("D4:C3").unwrap();
        let ws = finish_and_parse(sw);
        let refs: Vec<String> = ws
            .merge_cells
            .unwrap()
            .merge_cells
            .into_iter()
            .map(|m| m.reference)
            .collect();
        assert_eq!(refs, vec!["A1:B2".to_string(), "C3:D4".to_string()]);
    }

    #[test]
    fn add_merge_cell_rejects_overlaps_and_dedupes_equivalent_ranges() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.add_merge_cell("A1:B2").unwrap();
        // Equivalent spelling of an existing range is an idempotent no-op.
        sw.add_merge_cell("B2:A1").unwrap();
        assert_eq!(sw.config_count, 1);
        let err = sw.add_merge_cell("B2:C3").unwrap_err();
        match err {
            Error::MergeCellOverlap { new, existing } => {
                assert_eq!(new, "B2:C3");
                assert_eq!(existing, "A1:B2");
            }
            other => panic!("expected MergeCellOverlap, got {other:?}"),
        }
        assert_eq!(sw.config_count, 1);
        sw.add_merge_cell("C3:D4").unwrap();
        let ws = finish_and_parse(sw);
        assert_eq!(ws.merge_cells.unwrap().merge_cells.len(), 2);
    }

    #[test]
    fn stream_writer_is_send() {
        fn assert_send<T: Send>() {}
        assert_send::<StreamWriter>();
    }

    #[test]
    fn test_set_col_width_after_finish_fails() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.finished = true;
        let result = sw.set_col_width(1, 10.0);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamAlreadyFinished));
    }

    #[test]
    fn test_sheet_name_getter() {
        let sw = StreamWriter::new("MySheet");
        assert_eq!(sw.sheet_name(), "MySheet");
    }

    #[test]
    fn test_all_value_types_in_single_row() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[
                CellValue::from("text"),
                CellValue::from(42),
                CellValue::from(3.15),
                CellValue::from(true),
                CellValue::from(false),
                CellValue::Formula {
                    expr: "A1+B1".to_string(),
                    result: None,
                },
                CellValue::Error("#N/A".to_string()),
                CellValue::Empty,
            ],
        )
        .unwrap();
        let ws = finish_and_parse(sw);

        // 7 cells (Empty is skipped)
        assert_eq!(ws.sheet_data.rows[0].cells.len(), 7);
    }

    #[test]
    fn test_col_width_invalid_column_zero() {
        let mut sw = StreamWriter::new("Sheet1");
        let result = sw.set_col_width(0, 10.0);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::InvalidColumnNumber(0)));
    }

    #[test]
    fn test_write_row_with_options_height() {
        let mut sw = StreamWriter::new("Sheet1");
        let opts = StreamRowOptions {
            height: Some(30.0),
            ..Default::default()
        };
        sw.write_row_with_options(1, &[CellValue::from("tall row")], &opts)
            .unwrap();
        let ws = finish_and_parse(sw);

        assert_eq!(ws.sheet_data.rows[0].ht, Some(30.0));
        assert_eq!(ws.sheet_data.rows[0].custom_height, Some(true));
    }

    #[test]
    fn test_write_row_with_options_hidden() {
        let mut sw = StreamWriter::new("Sheet1");
        let opts = StreamRowOptions {
            visible: Some(false),
            ..Default::default()
        };
        sw.write_row_with_options(1, &[CellValue::from("hidden")], &opts)
            .unwrap();
        let ws = finish_and_parse(sw);

        assert_eq!(ws.sheet_data.rows[0].hidden, Some(true));
    }

    #[test]
    fn test_write_row_with_options_outline_level() {
        let mut sw = StreamWriter::new("Sheet1");
        let opts = StreamRowOptions {
            outline_level: Some(3),
            ..Default::default()
        };
        sw.write_row_with_options(1, &[CellValue::from("grouped")], &opts)
            .unwrap();
        let ws = finish_and_parse(sw);

        assert_eq!(ws.sheet_data.rows[0].outline_level, Some(3));
    }

    #[test]
    fn test_write_row_with_options_style() {
        let mut sw = StreamWriter::new("Sheet1");
        let opts = StreamRowOptions {
            style_id: Some(2),
            ..Default::default()
        };
        sw.write_row_with_options(1, &[CellValue::from("styled row")], &opts)
            .unwrap();
        let ws = finish_and_parse(sw);

        assert_eq!(ws.sheet_data.rows[0].s, Some(2));
        assert_eq!(ws.sheet_data.rows[0].custom_format, Some(true));
    }

    #[test]
    fn test_write_row_with_options_all() {
        let mut sw = StreamWriter::new("Sheet1");
        let opts = StreamRowOptions {
            height: Some(25.5),
            visible: Some(false),
            outline_level: Some(2),
            style_id: Some(7),
        };
        sw.write_row_with_options(1, &[CellValue::from("all options")], &opts)
            .unwrap();
        let ws = finish_and_parse(sw);

        let row = &ws.sheet_data.rows[0];
        assert_eq!(row.ht, Some(25.5));
        assert_eq!(row.custom_height, Some(true));
        assert_eq!(row.hidden, Some(true));
        assert_eq!(row.outline_level, Some(2));
        assert_eq!(row.s, Some(7));
        assert_eq!(row.custom_format, Some(true));
    }

    #[test]
    fn test_write_row_with_options_height_exceeded() {
        let mut sw = StreamWriter::new("Sheet1");
        let opts = StreamRowOptions {
            height: Some(500.0),
            ..Default::default()
        };
        let result = sw.write_row_with_options(1, &[CellValue::from("too tall")], &opts);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::RowHeightExceeded { .. }
        ));
    }

    #[test]
    fn test_write_row_with_options_outline_level_exceeded() {
        let mut sw = StreamWriter::new("Sheet1");
        let opts = StreamRowOptions {
            outline_level: Some(8),
            ..Default::default()
        };
        let result = sw.write_row_with_options(1, &[CellValue::from("bad level")], &opts);
        assert!(result.is_err());
    }

    #[test]
    fn test_write_row_with_options_visible_true_no_hidden_attr() {
        let mut sw = StreamWriter::new("Sheet1");
        let opts = StreamRowOptions {
            visible: Some(true),
            ..Default::default()
        };
        sw.write_row_with_options(1, &[CellValue::from("visible")], &opts)
            .unwrap();
        let xml = finish_and_get_xml(sw);

        // visible=true should NOT produce hidden attribute
        assert!(!xml.contains("hidden="));
    }

    #[test]
    fn test_col_style() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_style(1, 3).unwrap();
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let ws = finish_and_parse(sw);

        let cols = ws.cols.unwrap();
        let styled_col = cols.cols.iter().find(|c| c.style.is_some()).unwrap();
        assert_eq!(styled_col.style, Some(3));
        assert_eq!(styled_col.min, 1);
        assert_eq!(styled_col.max, 1);
    }

    #[test]
    fn test_col_visible() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_visible(2, false).unwrap();
        sw.write_row(1, &[CellValue::from("a"), CellValue::from("b")])
            .unwrap();
        let ws = finish_and_parse(sw);

        let cols = ws.cols.unwrap();
        let hidden_col = cols.cols.iter().find(|c| c.hidden == Some(true)).unwrap();
        assert_eq!(hidden_col.min, 2);
        assert_eq!(hidden_col.max, 2);
    }

    #[test]
    fn test_col_outline_level() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_outline_level(3, 2).unwrap();
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        let ws = finish_and_parse(sw);

        let cols = ws.cols.unwrap();
        let outlined_col = cols
            .cols
            .iter()
            .find(|c| c.outline_level.is_some())
            .unwrap();
        assert_eq!(outlined_col.outline_level, Some(2));
        assert_eq!(outlined_col.min, 3);
        assert_eq!(outlined_col.max, 3);
    }

    #[test]
    fn test_col_outline_level_exceeded() {
        let mut sw = StreamWriter::new("Sheet1");
        let result = sw.set_col_outline_level(1, 8);
        assert!(result.is_err());
    }

    #[test]
    fn test_col_style_after_rows_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let result = sw.set_col_style(1, 1);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamColumnsAfterRows));
    }

    #[test]
    fn test_col_visible_after_rows_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let result = sw.set_col_visible(1, false);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamColumnsAfterRows));
    }

    #[test]
    fn test_col_outline_after_rows_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let result = sw.set_col_outline_level(1, 1);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamColumnsAfterRows));
    }

    #[test]
    fn test_freeze_panes_rows() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_freeze_panes("A2").unwrap();
        sw.write_row(1, &[CellValue::from("header")]).unwrap();
        sw.write_row(2, &[CellValue::from("data")]).unwrap();
        let ws = finish_and_parse(sw);

        let views = ws.sheet_views.unwrap();
        let pane = views.sheet_views[0].pane.as_ref().unwrap();
        assert_eq!(pane.y_split, Some(1));
        assert_eq!(pane.top_left_cell, Some("A2".to_string()));
        assert_eq!(pane.active_pane, Some("bottomLeft".to_string()));
        assert_eq!(pane.state, Some("frozen".to_string()));
        // xSplit should not appear when only rows are frozen
        assert_eq!(pane.x_split, None);
    }

    #[test]
    fn test_freeze_panes_cols() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_freeze_panes("B1").unwrap();
        sw.write_row(1, &[CellValue::from("a"), CellValue::from("b")])
            .unwrap();
        let ws = finish_and_parse(sw);

        let views = ws.sheet_views.unwrap();
        let pane = views.sheet_views[0].pane.as_ref().unwrap();
        assert_eq!(pane.x_split, Some(1));
        assert_eq!(pane.top_left_cell, Some("B1".to_string()));
        assert_eq!(pane.active_pane, Some("topRight".to_string()));
        assert_eq!(pane.state, Some("frozen".to_string()));
        // ySplit should not appear when only cols are frozen
        assert_eq!(pane.y_split, None);
    }

    #[test]
    fn test_freeze_panes_both() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_freeze_panes("C3").unwrap();
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        let ws = finish_and_parse(sw);

        let views = ws.sheet_views.unwrap();
        let pane = views.sheet_views[0].pane.as_ref().unwrap();
        assert_eq!(pane.x_split, Some(2));
        assert_eq!(pane.y_split, Some(2));
        assert_eq!(pane.top_left_cell, Some("C3".to_string()));
        assert_eq!(pane.active_pane, Some("bottomRight".to_string()));
        assert_eq!(pane.state, Some("frozen".to_string()));
    }

    #[test]
    fn test_freeze_panes_after_rows_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let result = sw.set_freeze_panes("A2");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamColumnsAfterRows));
    }

    #[test]
    fn test_freeze_panes_a1_error() {
        let mut sw = StreamWriter::new("Sheet1");
        let result = sw.set_freeze_panes("A1");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::InvalidCellReference(_)
        ));
    }

    #[test]
    fn test_freeze_panes_invalid_cell_error() {
        let mut sw = StreamWriter::new("Sheet1");
        let result = sw.set_freeze_panes("ZZZZ1");
        assert!(result.is_err());
    }

    #[test]
    fn test_freeze_panes_appears_before_cols() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_freeze_panes("A2").unwrap();
        sw.set_col_width(1, 20.0).unwrap();
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let xml = finish_and_get_xml(sw);

        let views_pos = xml.find("<sheetView").unwrap();
        let cols_pos = xml.find("<cols>").unwrap();
        let data_pos = xml.find("<sheetData").unwrap();
        assert!(views_pos < cols_pos);
        assert!(cols_pos < data_pos);
    }

    #[test]
    fn test_freeze_panes_after_finish_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.finished = true;
        let result = sw.set_freeze_panes("A2");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamAlreadyFinished));
    }

    #[test]
    fn test_no_freeze_panes_no_sheet_views() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let xml = finish_and_get_xml(sw);

        // When no freeze panes are set, sheetViews should not appear
        assert!(!xml.contains("<sheetView"));
    }

    #[test]
    fn test_write_row_backward_compat() {
        // Ensure the original write_row still works exactly as before.
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("hello")]).unwrap();
        let ws = finish_and_parse(sw);

        let row = &ws.sheet_data.rows[0];
        assert_eq!(row.r, 1);
        // Should not have any row-level attributes beyond r
        assert!(row.ht.is_none());
        assert!(row.hidden.is_none());
        assert!(row.outline_level.is_none());
        assert!(row.custom_format.is_none());
    }

    #[test]
    fn test_col_options_combined_with_widths() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width(1, 20.0).unwrap();
        sw.set_col_style(2, 5).unwrap();
        sw.set_col_visible(3, false).unwrap();
        sw.write_row(
            1,
            &[
                CellValue::from("a"),
                CellValue::from("b"),
                CellValue::from("c"),
            ],
        )
        .unwrap();
        let ws = finish_and_parse(sw);

        let cols = ws.cols.unwrap();
        assert_eq!(cols.cols.len(), 3);
    }

    #[test]
    fn test_large_scale_10000_rows() {
        let mut sw = StreamWriter::new("BigSheet");
        for i in 1..=10_000u32 {
            sw.write_row(
                i,
                &[
                    CellValue::from(format!("Row {}", i)),
                    CellValue::from(i as i32),
                    CellValue::from(i as f64 * 1.5),
                ],
            )
            .unwrap();
        }
        let ws = finish_and_parse(sw);

        assert_eq!(ws.sheet_data.rows.len(), 10_000);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[9999].r, 10_000);
    }

    #[test]
    fn test_xml_escape_in_formula() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::Formula {
                expr: "IF(A1>0,\"yes\",\"no\")".to_string(),
                result: None,
            }],
        )
        .unwrap();
        let xml = finish_and_get_xml(sw);

        // The formula should be XML-escaped
        assert!(xml.contains("&gt;"));
    }

    #[test]
    fn test_xml_escape_in_string() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Tom & Jerry <friends>")])
            .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("Tom &amp; Jerry &lt;friends&gt;"));
    }

    #[test]
    fn test_into_streamed_data_returns_valid_data() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width(1, 20.0).unwrap();
        sw.write_row(1, &[CellValue::from("Hello"), CellValue::from(42)])
            .unwrap();
        sw.write_row(2, &[CellValue::from("World"), CellValue::from(99)])
            .unwrap();
        sw.add_merge_cell("A1:B1").unwrap();

        let (name, streamed) = sw.into_streamed_data().unwrap();
        assert_eq!(name, "Sheet1");
        assert!(streamed.data_len > 0);
        assert!(streamed.cols_xml.is_some());
        assert!(streamed.merge_cells_xml.is_some());
    }

    #[test]
    fn test_streamed_sheet_data_try_clone() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width(1, 20.0).unwrap();
        sw.set_freeze_panes("A2").unwrap();
        sw.write_row(1, &[CellValue::from("Hello")]).unwrap();
        sw.write_row(2, &[CellValue::from("World")]).unwrap();
        sw.add_merge_cell("A1:B1").unwrap();

        let (_, original) = sw.into_streamed_data().unwrap();
        let cloned = original.try_clone().unwrap();

        // Verify metadata is cloned.
        assert_eq!(cloned.data_len, original.data_len);
        assert_eq!(cloned.cols_xml, original.cols_xml);
        assert_eq!(cloned.sheet_views_xml, original.sheet_views_xml);
        assert_eq!(cloned.merge_cells_xml, original.merge_cells_xml);

        // Verify temp file contents are identical but independent.
        assert_ne!(original.temp_file.path(), cloned.temp_file.path());
        let orig_bytes = std::fs::read(original.temp_file.path()).unwrap();
        let clone_bytes = std::fs::read(cloned.temp_file.path()).unwrap();
        assert_eq!(orig_bytes, clone_bytes);
        assert!(!orig_bytes.is_empty());
    }

    #[test]
    fn test_date_value() {
        let mut sw = StreamWriter::new("Sheet1");
        // Excel serial number for 2024-01-15 = 45306
        sw.write_row(1, &[CellValue::Date(45306.0)]).unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("t=\"d\""));
        assert!(xml.contains("<v>2024-01-15T00:00:00</v>"));
    }

    #[test]
    fn test_date_value_with_time() {
        let mut sw = StreamWriter::new("Sheet1");
        // Excel serial with fractional time (noon)
        sw.write_row(1, &[CellValue::Date(45306.5)]).unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("<v>2024-01-15T12:00:00</v>"));
    }

    #[test]
    fn test_rich_string_to_inline_plain_text() {
        use crate::rich_text::RichTextRun;
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::RichString(vec![
                RichTextRun {
                    text: "Hello ".to_string(),
                    font: None,
                    size: None,
                    bold: true,
                    italic: false,
                    color: None,
                },
                RichTextRun {
                    text: "World".to_string(),
                    font: None,
                    size: None,
                    bold: false,
                    italic: true,
                    color: None,
                },
            ])],
        )
        .unwrap();
        let xml = finish_and_get_xml(sw);

        // RichString is flattened to plain inline string
        assert!(xml.contains("t=\"inlineStr\""));
        assert!(xml.contains("<is><t>Hello World</t></is>"));
    }

    #[test]
    fn test_rich_string_xml_escaping() {
        use crate::rich_text::RichTextRun;
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::RichString(vec![RichTextRun {
                text: "A & B <C>".to_string(),
                font: None,
                size: None,
                bold: false,
                italic: false,
                color: None,
            }])],
        )
        .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("A &amp; B &lt;C&gt;"));
    }

    #[test]
    fn test_formula_with_numeric_result() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::Formula {
                expr: "SUM(A2:A10)".to_string(),
                result: Some(Box::new(CellValue::Number(55.0))),
            }],
        )
        .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("<f>SUM(A2:A10)</f>"));
        assert!(xml.contains("<v>55</v>"));
        // Numeric results should not have a t attribute (default is numeric)
        assert!(!xml.contains("t=\"str\""));
        assert!(!xml.contains("t=\"b\""));
        assert!(!xml.contains("t=\"e\""));
    }

    #[test]
    fn test_formula_with_error_result() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::Formula {
                expr: "1/0".to_string(),
                result: Some(Box::new(CellValue::Error("#DIV/0!".to_string()))),
            }],
        )
        .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("t=\"e\""));
        assert!(xml.contains("<f>1/0</f>"));
        assert!(xml.contains("<v>#DIV/0!</v>"));
    }

    #[test]
    fn test_formula_with_string_result() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::Formula {
                expr: "IF(A1>0,\"yes\",\"no\")".to_string(),
                result: Some(Box::new(CellValue::String("yes".to_string()))),
            }],
        )
        .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("t=\"str\""));
        assert!(xml.contains("<f>"));
        assert!(xml.contains("<v>yes</v>"));
    }

    #[test]
    fn test_formula_with_bool_result() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::Formula {
                expr: "A1>0".to_string(),
                result: Some(Box::new(CellValue::Bool(true))),
            }],
        )
        .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("t=\"b\""));
        assert!(xml.contains("<f>A1&gt;0</f>"));
        assert!(xml.contains("<v>1</v>"));
    }

    #[test]
    fn test_formula_without_result() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::Formula {
                expr: "A1+B1".to_string(),
                result: None,
            }],
        )
        .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("<f>A1+B1</f></c>"));
        // No <v> tag when result is None
        assert!(!xml.contains("<v>"));
    }

    #[test]
    fn test_error_value_xml_escaping() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::Error("#VALUE!".to_string())])
            .unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("t=\"e\""));
        assert!(xml.contains("#VALUE!"));
    }

    #[test]
    fn test_empty_row_produces_valid_xml() {
        let mut sw = StreamWriter::new("Sheet1");
        // Row with only empty values
        sw.write_row(1, &[CellValue::Empty, CellValue::Empty])
            .unwrap();
        let ws = finish_and_parse(sw);

        assert_eq!(ws.sheet_data.rows.len(), 1);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        // All empties are skipped, row has no cells
        assert_eq!(ws.sheet_data.rows[0].cells.len(), 0);
    }

    #[test]
    fn test_memory_efficiency() {
        // This test validates that large data doesn't accumulate in memory.
        // The StreamWriter writes to a temp file, so memory should remain
        // roughly constant regardless of row count.
        let mut sw = StreamWriter::new("Sheet1");
        for i in 1..=100_000u32 {
            sw.write_row(
                i,
                &[
                    CellValue::from(format!("Row number {}", i)),
                    CellValue::from(i as f64),
                ],
            )
            .unwrap();
        }
        // Verify substantial data was written to disk, not held in memory.
        assert!(sw.bytes_written > 1_000_000);

        let (_, streamed) = sw.into_streamed_data().unwrap();
        assert!(streamed.data_len > 1_000_000);
    }

    #[test]
    fn test_write_rows_basic() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_rows(
            1,
            &[
                vec![CellValue::from("A"), CellValue::from(1)],
                vec![CellValue::from("B"), CellValue::from(2)],
                vec![CellValue::from("C"), CellValue::from(3)],
            ],
        )
        .unwrap();
        let ws = finish_and_parse(sw);

        assert_eq!(ws.sheet_data.rows.len(), 3);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[1].r, 2);
        assert_eq!(ws.sheet_data.rows[2].r, 3);
        assert_eq!(ws.sheet_data.rows[0].cells.len(), 2);
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");
        assert_eq!(ws.sheet_data.rows[2].cells[0].r, "A3");
    }

    #[test]
    fn test_write_rows_empty() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_rows(1, &[]).unwrap();
        let ws = finish_and_parse(sw);

        assert!(ws.sheet_data.rows.is_empty());
    }

    #[test]
    fn test_write_rows_mixed_with_single() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("first")]).unwrap();
        sw.write_rows(
            2,
            &[
                vec![CellValue::from("second")],
                vec![CellValue::from("third")],
                vec![CellValue::from("fourth")],
            ],
        )
        .unwrap();
        let ws = finish_and_parse(sw);

        assert_eq!(ws.sheet_data.rows.len(), 4);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[1].r, 2);
        assert_eq!(ws.sheet_data.rows[2].r, 3);
        assert_eq!(ws.sheet_data.rows[3].r, 4);
    }

    #[test]
    fn test_write_rows_overflow() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("ok")]).unwrap();
        let result = sw.write_rows(
            u32::MAX - 1,
            &[
                vec![CellValue::from("a")],
                vec![CellValue::from("b")],
                vec![CellValue::from("c")],
            ],
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_write_rows_backward_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(5, &[CellValue::from("row5")]).unwrap();
        let result = sw.write_rows(3, &[vec![CellValue::from("a")], vec![CellValue::from("b")]]);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::StreamRowAlreadyWritten { row: 3 }
        ));
    }

    #[test]
    fn test_rejects_non_finite_values_without_advancing_stream_state() {
        let mut sw = StreamWriter::new("Sheet1");
        let error = sw.write_row(1, &[CellValue::Number(f64::NAN)]).unwrap_err();
        assert!(matches!(error, Error::StreamNonFiniteValue));
        assert_eq!(sw.last_row, 0);
        assert_eq!(sw.bytes_written, 0);

        let error = sw
            .write_row(
                1,
                &[CellValue::Formula {
                    expr: "1/0".to_string(),
                    result: Some(Box::new(CellValue::Date(f64::INFINITY))),
                }],
            )
            .unwrap_err();
        assert!(matches!(error, Error::StreamNonFiniteValue));
        assert_eq!(sw.last_row, 0);
    }

    #[test]
    fn test_sanitizes_illegal_xml_controls_and_preserves_whitespace() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[
                CellValue::String("a\u{0001}b\t\n\rc".to_string()),
                CellValue::Formula {
                    expr: "A1&\"\u{0002}x\"".to_string(),
                    result: Some(Box::new(CellValue::String("r\u{0003}s".to_string()))),
                },
            ],
        )
        .unwrap();
        let xml = finish_and_get_xml(sw);
        assert!(!xml.contains('\u{0001}'));
        assert!(!xml.contains('\u{0002}'));
        assert!(!xml.contains('\u{0003}'));
        assert!(xml.contains("a b\t\n\rc") || xml.contains("ab\t\n\rc"));
        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.sheet_data.rows.len(), 1);
    }

    #[test]
    fn stream_limits_accept_exact_row_xml_and_reject_plus_one_without_output() {
        let value = CellValue::String("<&payload>".repeat(64));
        let mut probe = StreamWriter::new("Probe");
        probe.write_row(1, std::slice::from_ref(&value)).unwrap();
        let exact = probe.bytes_written;

        let exact_limits = crate::save::StreamLimits {
            row_xml_bytes: exact,
            output_bytes: exact,
            ..crate::save::StreamLimits::default()
        };
        let mut accepted = StreamWriter::new_with_limits("Accepted", exact_limits).unwrap();
        accepted.write_row(1, std::slice::from_ref(&value)).unwrap();
        assert_eq!(accepted.bytes_written, exact);

        let rejected_limits = crate::save::StreamLimits {
            row_xml_bytes: exact - 1,
            output_bytes: exact,
            ..crate::save::StreamLimits::default()
        };
        let mut rejected = StreamWriter::new_with_limits("Rejected", rejected_limits).unwrap();
        let error = rejected.write_row(1, &[value]).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(
            metadata.operation(),
            crate::resource::ResourceOperation::StreamWrite
        );
        assert_eq!(
            metadata.dimension(),
            crate::resource::ResourceDimension::XmlBytes
        );
        assert_eq!(metadata.limit(), exact - 1);
        assert_eq!(metadata.observed(), exact);
        assert_eq!(rejected.bytes_written, 0);
        assert_eq!(rejected.last_row, 0);
        assert!(rejected.writer.is_none());
    }

    #[test]
    fn stream_limits_reject_huge_string_formula_and_rich_text_before_output() {
        let limits = crate::save::StreamLimits {
            row_xml_bytes: 1024,
            output_bytes: 4096,
            ..crate::save::StreamLimits::default()
        };
        let cases = [
            CellValue::String("s".repeat(4096)),
            CellValue::Formula {
                expr: "A".repeat(4096),
                result: None,
            },
            CellValue::RichString(vec![crate::rich_text::RichTextRun {
                text: "r".repeat(4096),
                font: None,
                size: None,
                bold: false,
                italic: false,
                color: None,
            }]),
        ];

        for value in cases {
            let mut writer = StreamWriter::new_with_limits("Rejected", limits).unwrap();
            assert!(matches!(
                writer.write_row(1, &[value]),
                Err(Error::ResourceLimitExceeded(_))
            ));
            assert_eq!(writer.bytes_written, 0);
            assert_eq!(writer.last_row, 0);
            assert!(writer.writer.is_none());
        }
    }

    #[test]
    fn stream_batch_policy_failure_is_atomic_and_counts_rows_and_cells() {
        let mut probe = StreamWriter::new("Probe");
        probe
            .write_rows(
                1,
                &[
                    vec![CellValue::String("one".to_string())],
                    vec![CellValue::String("two".to_string())],
                ],
            )
            .unwrap();
        let exact_bytes = probe.bytes_written;

        let limits = crate::save::StreamLimits {
            row_count: 2,
            cell_count: 2,
            output_bytes: exact_bytes,
            ..crate::save::StreamLimits::default()
        };
        let mut accepted = StreamWriter::new_with_limits("Accepted", limits).unwrap();
        accepted
            .write_rows(
                1,
                &[
                    vec![CellValue::String("one".to_string())],
                    vec![CellValue::String("two".to_string())],
                ],
            )
            .unwrap();

        for exceeded in [
            crate::save::StreamLimits {
                row_count: 1,
                ..limits
            },
            crate::save::StreamLimits {
                cell_count: 1,
                ..limits
            },
            crate::save::StreamLimits {
                output_bytes: exact_bytes - 1,
                ..limits
            },
        ] {
            let mut rejected = StreamWriter::new_with_limits("Rejected", exceeded).unwrap();
            let error = rejected
                .write_rows(
                    1,
                    &[
                        vec![CellValue::String("one".to_string())],
                        vec![CellValue::String("two".to_string())],
                    ],
                )
                .unwrap_err();
            assert!(matches!(error, Error::ResourceLimitExceeded(_)));
            assert_eq!(rejected.bytes_written, 0);
            assert_eq!(rejected.last_row, 0);
            assert!(rejected.writer.is_none());
        }
    }

    #[test]
    fn stream_formula_budget_enforces_per_formula_and_cumulative_input_before_output() {
        let mut oversized = StreamWriter::new("Oversized");
        let error = oversized
            .write_row(
                1,
                &[CellValue::Formula {
                    expr: "A".repeat(32_769),
                    result: None,
                }],
            )
            .unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(
            metadata.operation(),
            crate::resource::ResourceOperation::FormulaMutate
        );
        assert_eq!(
            metadata.dimension(),
            crate::resource::ResourceDimension::InputBytes
        );
        assert_eq!(oversized.bytes_written, 0);
        assert!(oversized.writer.is_none());

        let formula = || CellValue::Formula {
            expr: "A".repeat(32_768),
            result: None,
        };
        let exact = (0..128).map(|_| formula()).collect::<Vec<_>>();
        let mut accepted = StreamWriter::new("Exact");
        accepted.write_row(1, &exact).unwrap();

        let plus_one = (0..129).map(|_| formula()).collect::<Vec<_>>();
        let mut rejected = StreamWriter::new("PlusOne");
        let error = rejected.write_row(1, &plus_one).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(
            metadata.dimension(),
            crate::resource::ResourceDimension::InputBytes
        );
        assert_eq!(metadata.limit(), 4 * 1024 * 1024);
        assert!(metadata.observed() > metadata.limit());
        assert_eq!(rejected.bytes_written, 0);
        assert!(rejected.writer.is_none());
    }

    #[test]
    fn stream_policy_rejection_does_not_consume_formula_budget() {
        let limits = crate::save::StreamLimits {
            cell_count: 1,
            ..crate::save::StreamLimits::default()
        };
        let formula = || CellValue::Formula {
            expr: "A".repeat(32_768),
            result: None,
        };
        let rejected_values = (0..128).map(|_| formula()).collect::<Vec<_>>();
        let mut writer = StreamWriter::new_with_limits("Retry", limits).unwrap();

        let error = writer.write_row(1, &rejected_values).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension(), ResourceDimension::CellCount);
        assert_eq!(writer.bytes_written, 0);
        assert_eq!(writer.last_row, 0);
        assert!(writer.writer.is_none());

        writer
            .write_row(
                1,
                &[CellValue::Formula {
                    expr: "1".to_string(),
                    result: None,
                }],
            )
            .unwrap();
    }

    #[test]
    fn stream_config_limits_are_atomic_and_deduplicate_updates() {
        let limits = crate::save::StreamLimits {
            config_count: 1,
            config_bytes: std::mem::size_of::<(u32, u32, f64)>() as u64,
            ..crate::save::StreamLimits::default()
        };
        let mut writer = StreamWriter::new_with_limits("Config", limits).unwrap();
        writer.set_col_width_range(1, 2, 10.0).unwrap();
        writer.set_col_width_range(1, 2, 20.0).unwrap();
        assert_eq!(writer.config_count, 1);

        let error = writer.add_merge_cell("A1:B1").unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension(), ResourceDimension::CollectionCount);
        assert_eq!(metadata.stage(), Some(ResourceStage::Validate));
        assert_eq!(writer.config_count, 1);
        assert_eq!(
            writer.config_bytes,
            std::mem::size_of::<(u32, u32, f64)>() as u64
        );

        let mut retained = StreamWriter::new_with_limits(
            "Retained",
            crate::save::StreamLimits {
                config_bytes: 1,
                ..crate::save::StreamLimits::default()
            },
        )
        .unwrap();
        let error = retained.set_col_width(1, 10.0).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension(),
            ResourceDimension::RetainedBytes
        );
        assert_eq!(retained.config_count, 0);
        assert_eq!(retained.config_bytes, 0);
        assert!(retained.col_widths.is_empty());
    }

    #[test]
    fn stream_cached_formula_results_reject_nested_and_unserializable_values_without_recursion() {
        let mut nested = CellValue::String("cached".to_string());
        for _ in 0..20_000 {
            nested = CellValue::Formula {
                expr: "A1".to_string(),
                result: Some(Box::new(nested)),
            };
        }
        let values = vec![CellValue::Formula {
            expr: "A1".to_string(),
            result: Some(Box::new(nested)),
        }];
        let mut writer = StreamWriter::new("Nested");
        let error = writer.write_row(1, &values).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().category(),
            ResourceCategory::Structure
        );
        assert!(writer.writer.is_none());
        std::mem::forget(values);

        let mut rich = StreamWriter::new("Rich");
        let error = rich
            .write_row(
                1,
                &[CellValue::Formula {
                    expr: "A1".to_string(),
                    result: Some(Box::new(CellValue::RichString(vec![]))),
                }],
            )
            .unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().category(),
            ResourceCategory::Structure
        );
        assert!(rich.writer.is_none());
    }

    #[test]
    fn partial_stream_output_failure_poison_prevents_retry_and_publication() {
        let mut writer = StreamWriter::new("Poisoned");
        writer.write_failure_after_bytes = Some(5);
        let error = writer
            .write_row(1, &[CellValue::String("payload".to_string())])
            .unwrap_err();
        assert!(matches!(error, Error::Io(_)));
        assert!(writer.failed);
        let path = writer
            .writer
            .as_ref()
            .unwrap()
            .get_ref()
            .path()
            .to_path_buf();
        assert!(path.exists());
        assert!(writer
            .write_row(1, &[CellValue::String("retry".to_string())])
            .is_err());
        assert!(writer.into_streamed_data().is_err());
        assert!(!path.exists());
    }

    /// Charges every character of a run separately, as the escape loop did
    /// before plain text was pushed in runs.
    struct PerCharCounter(RowXmlCounter);

    impl RowXmlSink for PerCharCounter {
        fn push_str(&mut self, value: &str) {
            self.0.charge(value.len());
        }

        fn push_chars(&mut self, value: &str) {
            for ch in value.chars() {
                self.0.charge(ch.len_utf8());
            }
        }
    }

    #[test]
    fn cell_names_match_coordinates_to_cell_name() {
        for row in [1, 9, 10, 1_048_576] {
            let row_text = row.to_string();
            for col in 1..=MAX_COLUMNS {
                let mut name = String::new();
                push_cell_name(&mut name, col, row, &row_text);
                assert_eq!(name, coordinates_to_cell_name(col, row).unwrap());
            }
        }
        let mut fallback = String::new();
        push_cell_name(&mut fallback, MAX_COLUMNS + 1, 1, "1");
        assert_eq!(fallback, format!("{}:1", MAX_COLUMNS + 1));
    }

    #[test]
    fn rendered_row_limit_matches_per_character_accounting() {
        let values = vec![
            CellValue::String("ab\u{d55c}c&d\u{1}e\u{1f600}f<g".to_string()),
            CellValue::Number(12.5),
            CellValue::Formula {
                expr: "A1&\"\u{e9}\"".to_string(),
                result: Some(Box::new(CellValue::String("x\u{ae00}y".to_string()))),
            },
        ];
        let full = measure_row_xml(7, &values, Some(3), None, u64::MAX).unwrap();
        for limit in 1..=full + 2 {
            let mut reference = PerCharCounter(RowXmlCounter::new(limit));
            write_row_xml(&mut reference, 7, &values, Some(3), None);
            let expected = reference.0.finish();

            let mut rendered = String::new();
            let actual = render_row_xml(&mut rendered, 7, &values, Some(3), None, limit);
            let measured = measure_row_xml(7, &values, Some(3), None, limit);
            for result in [&actual, &measured] {
                match (&expected, result) {
                    (Ok(e), Ok(a)) => assert_eq!(e, a),
                    (Err(e), Err(a)) => {
                        let (e, a) = (
                            e.resource_metadata().unwrap(),
                            a.resource_metadata().unwrap(),
                        );
                        assert_eq!(e.observed(), a.observed(), "limit {limit}");
                        assert_eq!(e.limit(), a.limit());
                    }
                    _ => panic!("limit {limit}: outcome differs"),
                }
            }
            if let Ok(size) = actual {
                assert_eq!(size, rendered.len() as u64);
            }
        }
    }

    #[test]
    fn write_row_reuses_one_render_for_limit_and_output() {
        let values = vec![CellValue::String("a&b".to_string()), CellValue::Number(1.5)];
        let mut expected = String::new();
        let size = render_row_xml(&mut expected, 1, &values, None, None, u64::MAX).unwrap();

        let limits = crate::save::StreamLimits {
            row_xml_bytes: size,
            ..crate::save::StreamLimits::default()
        };
        let mut exact = StreamWriter::new_with_limits("Exact", limits).unwrap();
        exact.write_row(1, &values).unwrap();
        assert_eq!(exact.bytes_written, size);
        let xml = finish_and_get_xml(exact);
        assert!(xml.contains(&expected));

        let limits = crate::save::StreamLimits {
            row_xml_bytes: size - 1,
            ..crate::save::StreamLimits::default()
        };
        let mut rejected = StreamWriter::new_with_limits("Rejected", limits).unwrap();
        assert!(rejected.write_row(1, &values).is_err());
        assert_eq!(rejected.bytes_written, 0);
        assert!(rejected.writer.is_none());
    }
}
