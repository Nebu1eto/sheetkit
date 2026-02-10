//! Streaming worksheet writer.
//!
//! The [`StreamWriter`] builds `WorksheetXml` structs directly instead of
//! writing raw XML text. Rows must be written in ascending order.
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
//! let xml_bytes = sw.finish().unwrap();
//! assert!(!xml_bytes.is_empty());
//! ```

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::sst::SharedStringTable;
use crate::utils::cell_ref::cell_name_to_coordinates;
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS, MAX_ROW_HEIGHT};

use sheetkit_xml::worksheet::{
    Cell, CellFormula, CellTypeTag, Col, Cols, CompactCellRef, MergeCell, MergeCells, Pane, Row,
    SheetData, SheetView, SheetViews, WorksheetXml,
};

/// XML declaration prepended to the worksheet XML.
const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

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

/// A streaming worksheet writer that writes rows sequentially.
///
/// Rows must be written in ascending order. The StreamWriter builds
/// `Row` structs directly, then assembles them into a `WorksheetXml`
/// on finish.
#[derive(Debug)]
pub struct StreamWriter {
    sheet_name: String,
    rows: Vec<Row>,
    last_row: u32,
    started: bool,
    finished: bool,
    col_widths: Vec<(u32, u32, f64)>,
    col_options: Vec<(u32, StreamColOptions)>,
    sst: SharedStringTable,
    merge_cells: Vec<String>,
    /// Freeze pane: (x_split, y_split, top_left_cell).
    freeze_pane: Option<(u32, u32, String)>,
}

impl StreamWriter {
    /// Create a new StreamWriter for the given sheet name.
    pub fn new(sheet_name: &str) -> Self {
        Self {
            sheet_name: sheet_name.to_string(),
            rows: Vec::new(),
            last_row: 0,
            started: false,
            finished: false,
            col_widths: Vec::new(),
            col_options: Vec::new(),
            sst: SharedStringTable::new(),
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

    /// Finish writing and return the complete worksheet XML bytes.
    pub fn finish(&mut self) -> Result<Vec<u8>> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        self.finished = true;
        let ws = self.build_worksheet_xml();
        let body = quick_xml::se::to_string(&ws).map_err(|e| Error::Internal(e.to_string()))?;
        let mut xml = String::with_capacity(body.len() + 60);
        xml.push_str(XML_DECLARATION);
        xml.push('\n');
        xml.push_str(&body);
        Ok(xml.into_bytes())
    }

    /// Get a reference to the shared string table.
    pub fn sst(&self) -> &SharedStringTable {
        &self.sst
    }

    /// Consume the writer and return (xml_bytes, shared_string_table).
    pub fn into_parts(mut self) -> Result<(Vec<u8>, SharedStringTable)> {
        let xml = self.finish()?;
        Ok((xml, self.sst))
    }

    /// Consume the writer and return the worksheet XML struct and shared string table.
    /// This is the optimized path that avoids XML serialization/deserialization.
    pub fn into_worksheet_parts(mut self) -> Result<(WorksheetXml, SharedStringTable)> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        self.finished = true;
        let ws = self.build_worksheet_xml_owned();
        Ok((ws, self.sst))
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

    /// Build a `WorksheetXml` by cloning the accumulated rows.
    fn build_worksheet_xml(&self) -> WorksheetXml {
        self.build_worksheet_xml_inner(self.rows.clone())
    }

    /// Build a `WorksheetXml` by taking ownership of the accumulated rows.
    fn build_worksheet_xml_owned(&mut self) -> WorksheetXml {
        let rows = std::mem::take(&mut self.rows);
        self.build_worksheet_xml_inner(rows)
    }

    /// Assemble the final `WorksheetXml` from the given rows and accumulated config.
    fn build_worksheet_xml_inner(&self, rows: Vec<Row>) -> WorksheetXml {
        let sheet_views = self
            .freeze_pane
            .as_ref()
            .map(|(x_split, y_split, top_left_cell)| {
                let active_pane = match (*x_split > 0, *y_split > 0) {
                    (true, true) => "bottomRight",
                    (true, false) => "topRight",
                    (false, true) => "bottomLeft",
                    (false, false) => unreachable!(),
                };
                SheetViews {
                    sheet_views: vec![SheetView {
                        tab_selected: Some(true),
                        zoom_scale: None,
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
                }
            });

        let cols = {
            let has_widths = !self.col_widths.is_empty();
            let has_options = !self.col_options.is_empty();
            if has_widths || has_options {
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
                Some(Cols { cols: col_defs })
            } else {
                None
            }
        };

        let merge_cells = if self.merge_cells.is_empty() {
            None
        } else {
            Some(MergeCells {
                count: Some(self.merge_cells.len() as u32),
                merge_cells: self
                    .merge_cells
                    .iter()
                    .map(|r| MergeCell {
                        reference: r.clone(),
                    })
                    .collect(),
            })
        };

        WorksheetXml {
            sheet_views,
            cols,
            sheet_data: SheetData { rows },
            merge_cells,
            ..WorksheetXml::default()
        }
    }

    /// Internal unified implementation for writing a row.
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

        let mut cells = Vec::new();

        for (i, value) in values.iter().enumerate() {
            let col = (i as u32) + 1;

            match value {
                CellValue::Empty => continue,
                CellValue::String(s) => {
                    let idx = self.sst.add(s);
                    cells.push(Cell {
                        r: CompactCellRef::from_coordinates(col, row),
                        col,
                        s: cell_style_id,
                        t: CellTypeTag::SharedString,
                        v: Some(idx.to_string()),
                        f: None,
                        is: None,
                    });
                }
                CellValue::Number(n) => {
                    cells.push(Cell {
                        r: CompactCellRef::from_coordinates(col, row),
                        col,
                        s: cell_style_id,
                        t: CellTypeTag::None,
                        v: Some(n.to_string()),
                        f: None,
                        is: None,
                    });
                }
                CellValue::Date(serial) => {
                    cells.push(Cell {
                        r: CompactCellRef::from_coordinates(col, row),
                        col,
                        s: cell_style_id,
                        t: CellTypeTag::None,
                        v: Some(serial.to_string()),
                        f: None,
                        is: None,
                    });
                }
                CellValue::Bool(b) => {
                    let val = if *b { "1" } else { "0" };
                    cells.push(Cell {
                        r: CompactCellRef::from_coordinates(col, row),
                        col,
                        s: cell_style_id,
                        t: CellTypeTag::Boolean,
                        v: Some(val.to_string()),
                        f: None,
                        is: None,
                    });
                }
                CellValue::Formula { expr, .. } => {
                    cells.push(Cell {
                        r: CompactCellRef::from_coordinates(col, row),
                        col,
                        s: cell_style_id,
                        t: CellTypeTag::None,
                        v: None,
                        f: Some(CellFormula {
                            t: None,
                            reference: None,
                            si: None,
                            value: Some(expr.clone()),
                        }),
                        is: None,
                    });
                }
                CellValue::Error(e) => {
                    cells.push(Cell {
                        r: CompactCellRef::from_coordinates(col, row),
                        col,
                        s: cell_style_id,
                        t: CellTypeTag::Error,
                        v: Some(e.clone()),
                        f: None,
                        is: None,
                    });
                }
                CellValue::RichString(runs) => {
                    let plain = crate::rich_text::rich_text_to_plain(runs);
                    let idx = self.sst.add(&plain);
                    cells.push(Cell {
                        r: CompactCellRef::from_coordinates(col, row),
                        col,
                        s: cell_style_id,
                        t: CellTypeTag::SharedString,
                        v: Some(idx.to_string()),
                        f: None,
                        is: None,
                    });
                }
            }
        }

        let mut xml_row = Row {
            r: row,
            spans: None,
            s: None,
            custom_format: None,
            ht: None,
            hidden: None,
            custom_height: None,
            outline_level: None,
            cells,
        };

        if let Some(opts) = options {
            if let Some(height) = opts.height {
                xml_row.ht = Some(height);
                xml_row.custom_height = Some(true);
            }
            if let Some(false) = opts.visible {
                xml_row.hidden = Some(true);
            }
            if let Some(level) = opts.outline_level {
                if level > 0 {
                    xml_row.outline_level = Some(level);
                }
            }
            if let Some(sid) = opts.style_id {
                xml_row.s = Some(sid);
                xml_row.custom_format = Some(true);
            }
        }

        self.rows.push(xml_row);
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetkit_xml::worksheet::WorksheetXml;

    #[test]
    fn test_basic_write_and_finish() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Name"), CellValue::from("Age")])
            .unwrap();
        sw.write_row(2, &[CellValue::from("Alice"), CellValue::from(30)])
            .unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        // Parse back into WorksheetXml
        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
        assert_eq!(ws.sheet_data.rows.len(), 1);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[0].cells.len(), 2);
        // First cell is a shared string (t="s")
        assert_eq!(
            ws.sheet_data.rows[0].cells[0].t,
            sheetkit_xml::worksheet::CellTypeTag::SharedString
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
    fn test_string_values_populate_sst() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Hello"), CellValue::from("World")])
            .unwrap();

        assert_eq!(sw.sst().len(), 2);
        assert_eq!(sw.sst().get(0), Some("Hello"));
        assert_eq!(sw.sst().get(1), Some("World"));
    }

    #[test]
    fn test_number_value() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from(3.15)]).unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        assert!(xml.contains("<v>3.15</v>"));
    }

    #[test]
    fn test_bool_values() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from(true), CellValue::from(false)])
            .unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        assert!(xml.contains("SUM(A2:A10)"));
    }

    #[test]
    fn test_error_value() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::Error("#DIV/0!".to_string())])
            .unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        // Parse and verify only 2 cells present (A1 and C1)
        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
        assert_eq!(ws.sheet_data.rows[0].cells.len(), 2);
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");
        assert_eq!(ws.sheet_data.rows[0].cells[1].r, "C1");
    }

    #[test]
    fn test_write_row_with_style() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row_with_style(1, &[CellValue::from("Styled")], 5)
            .unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        assert!(xml.contains("s=\"5\""));
    }

    #[test]
    fn test_set_col_width_before_rows() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width(1, 20.0).unwrap();
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        // Verify cols section appears before sheetData
        let cols_pos = xml_str.find("<cols>").unwrap();
        let sheet_data_pos = xml_str.find("<sheetData").unwrap();
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
        sw.finish().unwrap();
    }

    #[test]
    fn test_rows_with_gaps_succeeds() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        sw.write_row(3, &[CellValue::from("b")]).unwrap();
        sw.write_row(5, &[CellValue::from("c")]).unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
        let mc = ws.merge_cells.unwrap();
        assert_eq!(mc.merge_cells.len(), 2);
        assert_eq!(mc.merge_cells[0].reference, "A1:B1");
        assert_eq!(mc.merge_cells[1].reference, "C1:D1");
    }

    #[test]
    fn test_finish_twice_returns_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        sw.finish().unwrap();
        let result = sw.finish();
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamAlreadyFinished));
    }

    #[test]
    fn test_write_after_finish_returns_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        sw.finish().unwrap();
        let result = sw.write_row(2, &[CellValue::from("b")]);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamAlreadyFinished));
    }

    #[test]
    fn test_finish_with_no_rows() {
        let mut sw = StreamWriter::new("Sheet1");
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        // Should still produce valid XML
        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
        assert!(ws.sheet_data.rows.is_empty());
    }

    #[test]
    fn test_finish_with_cols_and_no_rows() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width(1, 20.0).unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
        assert!(ws.cols.is_some());
        assert!(ws.sheet_data.rows.is_empty());
    }

    #[test]
    fn test_into_parts() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Hello")]).unwrap();
        let (xml_bytes, sst) = sw.into_parts().unwrap();

        assert!(!xml_bytes.is_empty());
        assert_eq!(sst.len(), 1);
        assert_eq!(sst.get(0), Some("Hello"));
    }

    #[test]
    fn test_sst_deduplication() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("same"), CellValue::from("same")])
            .unwrap();
        sw.write_row(2, &[CellValue::from("same"), CellValue::from("other")])
            .unwrap();

        assert_eq!(sw.sst().len(), 2); // "same" and "other"
        assert_eq!(sw.sst().get(0), Some("same"));
        assert_eq!(sw.sst().get(1), Some("other"));
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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        // Verify it parses
        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
        assert_eq!(ws.sheet_data.rows.len(), 10_000);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[9999].r, 10_000);
    }

    #[test]
    fn test_large_sst_dedup() {
        let mut sw = StreamWriter::new("Sheet1");
        // Write 1000 rows, each with the same 3 string values
        for i in 1..=1000u32 {
            sw.write_row(
                i,
                &[
                    CellValue::from("Alpha"),
                    CellValue::from("Beta"),
                    CellValue::from("Gamma"),
                ],
            )
            .unwrap();
        }
        // SST should only have 3 unique strings
        assert_eq!(sw.sst().len(), 3);
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();
        // quick_xml escapes > in text content; quotes don't require escaping in text content
        assert!(xml.contains("&gt;"));
        // Verify the formula round-trips through XML
        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        let formula = ws.sheet_data.rows[0].cells[0].f.as_ref().unwrap();
        assert_eq!(formula.value, Some("IF(A1>0,\"yes\",\"no\")".to_string()));
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
        sw.finish().unwrap();
        let result = sw.add_merge_cell("A1:B1");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamAlreadyFinished));
    }

    #[test]
    fn test_set_col_width_after_finish_fails() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.finish().unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        // Should parse without error
        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        // Parse back and verify
        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        // visible=true should NOT produce hidden attribute
        assert!(!xml.contains("hidden="));
    }

    #[test]
    fn test_col_style() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_style(1, 3).unwrap();
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        let views_pos = xml.find("<sheetView").unwrap();
        let cols_pos = xml.find("<cols>").unwrap();
        let data_pos = xml.find("<sheetData").unwrap();
        assert!(views_pos < cols_pos);
        assert!(cols_pos < data_pos);
    }

    #[test]
    fn test_freeze_panes_after_finish_error() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.finish().unwrap();
        let result = sw.set_freeze_panes("A2");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StreamAlreadyFinished));
    }

    #[test]
    fn test_no_freeze_panes_no_sheet_views() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        // When no freeze panes are set, sheetViews should not appear
        assert!(!xml.contains("<sheetView"));
    }

    #[test]
    fn test_write_row_backward_compat() {
        // Ensure the original write_row still works exactly as before.
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("hello")]).unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        let row = &ws.sheet_data.rows[0];
        assert_eq!(row.r, 1);
        // Should not have any row-level attributes beyond r
        assert!(row.ht.is_none());
        assert!(row.hidden.is_none());
        assert!(row.outline_level.is_none());
        assert!(row.custom_format.is_none());
    }

    #[test]
    fn test_write_row_with_style_backward_compat() {
        // Ensure the original write_row_with_style still works.
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row_with_style(1, &[CellValue::from("styled")], 5)
            .unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        // Cell should have s="5" attribute
        assert!(xml.contains("s=\"5\""));
        // Row should not have row-level style attributes
        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        assert!(ws.sheet_data.rows[0].custom_format.is_none());
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
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        let ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        let cols = ws.cols.unwrap();
        assert_eq!(cols.cols.len(), 3);
    }

    #[test]
    fn test_into_worksheet_parts_basic() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Hello"), CellValue::from(42)])
            .unwrap();
        sw.write_row(2, &[CellValue::from("World"), CellValue::from(99)])
            .unwrap();
        let (ws, sst) = sw.into_worksheet_parts().unwrap();
        assert_eq!(ws.sheet_data.rows.len(), 2);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[0].cells.len(), 2);
        assert_eq!(ws.sheet_data.rows[1].r, 2);
        assert_eq!(sst.len(), 2);
    }

    #[test]
    fn test_into_worksheet_parts_sst() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Alpha"), CellValue::from("Beta")])
            .unwrap();
        let (ws, sst) = sw.into_worksheet_parts().unwrap();
        assert_eq!(sst.len(), 2);
        assert_eq!(sst.get(0), Some("Alpha"));
        assert_eq!(sst.get(1), Some("Beta"));
        // Cells should reference SST indices
        assert_eq!(ws.sheet_data.rows[0].cells[0].v, Some("0".to_string()));
        assert_eq!(ws.sheet_data.rows[0].cells[1].v, Some("1".to_string()));
        assert_eq!(
            ws.sheet_data.rows[0].cells[0].t,
            sheetkit_xml::worksheet::CellTypeTag::SharedString
        );
    }

    #[test]
    fn test_into_worksheet_parts_cell_col_set() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(
            1,
            &[CellValue::from("A"), CellValue::Empty, CellValue::from("C")],
        )
        .unwrap();
        let (ws, _) = sw.into_worksheet_parts().unwrap();
        assert_eq!(ws.sheet_data.rows[0].cells.len(), 2);
        assert_eq!(ws.sheet_data.rows[0].cells[0].col, 1);
        assert_eq!(ws.sheet_data.rows[0].cells[0].r.as_str(), "A1");
        assert_eq!(ws.sheet_data.rows[0].cells[1].col, 3);
        assert_eq!(ws.sheet_data.rows[0].cells[1].r.as_str(), "C1");
    }

    #[test]
    fn test_into_worksheet_parts_freeze_panes() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_freeze_panes("C3").unwrap();
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let (ws, _) = sw.into_worksheet_parts().unwrap();
        assert!(ws.sheet_views.is_some());
        let views = ws.sheet_views.unwrap();
        let pane = views.sheet_views[0].pane.as_ref().unwrap();
        assert_eq!(pane.x_split, Some(2));
        assert_eq!(pane.y_split, Some(2));
        assert_eq!(pane.top_left_cell, Some("C3".to_string()));
        assert_eq!(pane.state, Some("frozen".to_string()));
    }

    #[test]
    fn test_into_worksheet_parts_cols() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width(1, 20.0).unwrap();
        sw.set_col_style(2, 5).unwrap();
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let (ws, _) = sw.into_worksheet_parts().unwrap();
        assert!(ws.cols.is_some());
        let cols = ws.cols.unwrap();
        assert_eq!(cols.cols.len(), 2);
    }

    #[test]
    fn test_into_worksheet_parts_merge_cells() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Merged")]).unwrap();
        sw.add_merge_cell("A1:B1").unwrap();
        let (ws, _) = sw.into_worksheet_parts().unwrap();
        assert!(ws.merge_cells.is_some());
        let mc = ws.merge_cells.unwrap();
        assert_eq!(mc.merge_cells.len(), 1);
        assert_eq!(mc.merge_cells[0].reference, "A1:B1");
    }
}
