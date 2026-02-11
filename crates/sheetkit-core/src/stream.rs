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

use std::io::{BufWriter, Write as _};

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::utils::cell_ref::cell_name_to_coordinates;
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS, MAX_ROW_HEIGHT};

use sheetkit_xml::worksheet::{Col, Cols, MergeCell, MergeCells, Pane, SheetView, SheetViews};

/// Maximum outline level allowed by Excel.
const MAX_OUTLINE_LEVEL: u8 = 7;

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
            sheet_views_xml: self.sheet_views_xml.clone(),
            cols_xml: self.cols_xml.clone(),
            merge_cells_xml: self.merge_cells_xml.clone(),
        })
    }
}

/// A streaming worksheet writer that writes rows directly to a temp file.
///
/// Rows must be written in ascending order. Each row is serialized to XML
/// and written to a temporary file immediately, keeping memory usage constant
/// regardless of the number of rows.
pub struct StreamWriter {
    sheet_name: String,
    writer: BufWriter<tempfile::NamedTempFile>,
    bytes_written: u64,
    last_row: u32,
    started: bool,
    finished: bool,
    col_widths: Vec<(u32, u32, f64)>,
    col_options: Vec<(u32, StreamColOptions)>,
    merge_cells: Vec<String>,
    /// Freeze pane: (x_split, y_split, top_left_cell).
    freeze_pane: Option<(u32, u32, String)>,
}

// StreamWriter contains a BufWriter<NamedTempFile> which is not Send by
// default on all platforms. The temp file is exclusively owned by the writer
// so this is safe.
unsafe impl Send for StreamWriter {}

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
        let temp = tempfile::NamedTempFile::new().expect("failed to create temp file");
        Self {
            sheet_name: sheet_name.to_string(),
            writer: BufWriter::new(temp),
            bytes_written: 0,
            last_row: 0,
            started: false,
            finished: false,
            col_widths: Vec::new(),
            col_options: Vec::new(),
            merge_cells: Vec::new(),
            freeze_pane: None,
        }
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
        self.col_widths.push((min_col, max_col, width));
        Ok(())
    }

    /// Set column style for a single column (1-based).
    /// Must be called before any write_row() calls.
    pub fn set_col_style(&mut self, col: u32, style_id: u32) -> Result<()> {
        self.ensure_col_configurable(col)?;
        self.get_or_create_col_options(col).style_id = Some(style_id);
        Ok(())
    }

    /// Set column visibility for a single column (1-based).
    /// Must be called before any write_row() calls.
    pub fn set_col_visible(&mut self, col: u32, visible: bool) -> Result<()> {
        self.ensure_col_configurable(col)?;
        self.get_or_create_col_options(col).hidden = Some(!visible);
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
        self.get_or_create_col_options(col).outline_level = Some(level);
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
    /// Must be called before finish(). The reference must be a valid range
    /// in the form "A1:B2".
    pub fn add_merge_cell(&mut self, reference: &str) -> Result<()> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        // Validate the range format.
        let parts: Vec<&str> = reference.split(':').collect();
        if parts.len() != 2 {
            return Err(Error::InvalidMergeCellReference(reference.to_string()));
        }
        cell_name_to_coordinates(parts[0])
            .map_err(|_| Error::InvalidMergeCellReference(reference.to_string()))?;
        cell_name_to_coordinates(parts[1])
            .map_err(|_| Error::InvalidMergeCellReference(reference.to_string()))?;
        self.merge_cells.push(reference.to_string());
        Ok(())
    }

    /// Consume the writer and return a [`StreamedSheetData`] containing the
    /// temp file and pre-built XML fragments for the worksheet header/footer.
    pub fn into_streamed_data(mut self) -> Result<(String, StreamedSheetData)> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        self.finished = true;

        // Build XML fragments before consuming the writer.
        let sheet_views_xml = self.build_sheet_views_xml();
        let cols_xml = self.build_cols_xml();
        let merge_cells_xml = self.build_merge_cells_xml();
        let bytes_written = self.bytes_written;
        let sheet_name = self.sheet_name.clone();

        // Flush the writer and recover the temp file.
        self.writer.flush()?;
        let temp_file = self
            .writer
            .into_inner()
            .map_err(|e| Error::Io(e.into_error()))?;

        let data = StreamedSheetData {
            temp_file,
            data_len: bytes_written,
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
    fn get_or_create_col_options(&mut self, col: u32) -> &mut StreamColOptions {
        if let Some(pos) = self.col_options.iter().position(|(c, _)| *c == col) {
            &mut self.col_options[pos].1
        } else {
            self.col_options.push((col, StreamColOptions::default()));
            let last = self.col_options.len() - 1;
            &mut self.col_options[last].1
        }
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
                .map(|r| MergeCell {
                    reference: r.clone(),
                })
                .collect(),
            cached_coords: Vec::new(),
        };
        quick_xml::se::to_string_with_root("mergeCells", &mc).ok()
    }

    /// Internal unified implementation for writing a row.
    /// Serializes the row to XML and writes it directly to the temp file.
    fn write_row_impl(
        &mut self,
        row: u32,
        values: &[CellValue],
        cell_style_id: Option<u32>,
        options: Option<&StreamRowOptions>,
    ) -> Result<()> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
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

        self.started = true;
        self.last_row = row;

        // Build row XML directly and write to temp file.
        let xml = build_row_xml(row, values, cell_style_id, options);
        let bytes = xml.as_bytes();
        self.writer.write_all(bytes)?;
        self.bytes_written += bytes.len() as u64;

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
) -> Result<()> {
    zip.start_file(entry_name, options)
        .map_err(|e| Error::Zip(e.to_string()))?;

    // Write the header (XML declaration, <worksheet>, <sheetViews>, <cols>, <sheetData>).
    let header = build_worksheet_header(streamed);
    zip.write_all(header.as_bytes())?;

    // Open the temp file by path for reading (starts at position 0).
    let file = std::fs::File::open(streamed.temp_file.path())?;
    let mut reader = std::io::BufReader::new(file);
    std::io::copy(&mut reader, zip)?;

    // Write the footer (</sheetData>, <mergeCells>, </worksheet>).
    let footer = build_worksheet_footer(streamed);
    zip.write_all(footer.as_bytes())?;

    Ok(())
}

/// Build an XML `<row>` element with inline strings for a single row.
fn build_row_xml(
    row: u32,
    values: &[CellValue],
    cell_style_id: Option<u32>,
    options: Option<&StreamRowOptions>,
) -> String {
    let mut xml = String::with_capacity(128 + values.len() * 64);
    xml.push_str("<row r=\"");
    xml.push_str(&row.to_string());
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

        let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(col, row)
            .unwrap_or_else(|_| format!("{col}:{row}"));

        xml.push_str("<c r=\"");
        xml.push_str(&cell_ref);
        xml.push('"');

        // Style attribute.
        if let Some(sid) = cell_style_id {
            xml.push_str(" s=\"");
            xml.push_str(&sid.to_string());
            xml.push('"');
        }

        match value {
            CellValue::String(s) => {
                xml.push_str(" t=\"inlineStr\"><is><t>");
                xml_escape_into(&mut xml, s);
                xml.push_str("</t></is></c>");
            }
            CellValue::Number(n) => {
                xml.push_str("><v>");
                xml.push_str(&n.to_string());
                xml.push_str("</v></c>");
            }
            CellValue::Date(serial) => {
                xml.push_str("><v>");
                xml.push_str(&serial.to_string());
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
                        _ => {} // Number/Date: default (no t) is numeric
                    }
                }
                xml.push_str("><f>");
                xml_escape_into(&mut xml, expr);
                xml.push_str("</f>");
                if let Some(res) = result {
                    match res.as_ref() {
                        CellValue::String(s) => {
                            xml.push_str("<v>");
                            xml_escape_into(&mut xml, s);
                            xml.push_str("</v>");
                        }
                        CellValue::Number(n) => {
                            xml.push_str("<v>");
                            xml.push_str(&n.to_string());
                            xml.push_str("</v>");
                        }
                        CellValue::Bool(b) => {
                            xml.push_str("<v>");
                            xml.push_str(if *b { "1" } else { "0" });
                            xml.push_str("</v>");
                        }
                        CellValue::Date(d) => {
                            xml.push_str("<v>");
                            xml.push_str(&d.to_string());
                            xml.push_str("</v>");
                        }
                        CellValue::Error(e) => {
                            xml.push_str("<v>");
                            xml_escape_into(&mut xml, e);
                            xml.push_str("</v>");
                        }
                        _ => {}
                    }
                }
                xml.push_str("</c>");
            }
            CellValue::Error(e) => {
                xml.push_str(" t=\"e\"><v>");
                xml_escape_into(&mut xml, e);
                xml.push_str("</v></c>");
            }
            CellValue::RichString(runs) => {
                let plain = crate::rich_text::rich_text_to_plain(runs);
                xml.push_str(" t=\"inlineStr\"><is><t>");
                xml_escape_into(&mut xml, &plain);
                xml.push_str("</t></is></c>");
            }
            CellValue::Empty => unreachable!(),
        }
    }

    xml.push_str("</row>");
    xml
}

/// Escape XML special characters into an existing string buffer.
fn xml_escape_into(buf: &mut String, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => buf.push_str("&amp;"),
            '<' => buf.push_str("&lt;"),
            '>' => buf.push_str("&gt;"),
            '"' => buf.push_str("&quot;"),
            '\'' => buf.push_str("&apos;"),
            _ => buf.push(ch),
        }
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

        // Date is serialized as a plain numeric <v> with no type tag
        assert!(xml.contains("<v>45306</v>"));
        // Should not have t="..." attribute (same as Number)
        assert!(!xml.contains("t=\"inlineStr\""));
        assert!(!xml.contains("t=\"b\""));
    }

    #[test]
    fn test_date_value_with_time() {
        let mut sw = StreamWriter::new("Sheet1");
        // Excel serial with fractional time (noon)
        sw.write_row(1, &[CellValue::Date(45306.5)]).unwrap();
        let xml = finish_and_get_xml(sw);

        assert!(xml.contains("<v>45306.5</v>"));
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
}
