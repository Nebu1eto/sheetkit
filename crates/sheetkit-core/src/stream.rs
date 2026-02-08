//! Streaming worksheet writer.
//!
//! The [`StreamWriter`] writes XML directly to an internal buffer, avoiding
//! the need to build the entire [`WorksheetXml`] in memory. Rows must be
//! written in ascending order.
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

use std::fmt::Write as _;

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::sst::SharedStringTable;
use crate::utils::cell_ref::coordinates_to_cell_name;
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS};

/// XML declaration prepended to the worksheet XML.
const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// SpreadsheetML namespace.
const NS_SPREADSHEET: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// Relationships namespace.
const NS_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// A streaming worksheet writer that writes rows sequentially.
///
/// Rows must be written in ascending order. The StreamWriter writes
/// XML directly to an internal buffer, avoiding the need to build
/// the entire worksheet in memory.
#[derive(Debug)]
pub struct StreamWriter {
    sheet_name: String,
    buffer: String,
    last_row: u32,
    started: bool,
    finished: bool,
    col_widths: Vec<(u32, u32, f64)>,
    sst: SharedStringTable,
    merge_cells: Vec<String>,
}

impl StreamWriter {
    /// Create a new StreamWriter for the given sheet name.
    pub fn new(sheet_name: &str) -> Self {
        Self {
            sheet_name: sheet_name.to_string(),
            buffer: String::new(),
            last_row: 0,
            started: false,
            finished: false,
            col_widths: Vec::new(),
            sst: SharedStringTable::new(),
            merge_cells: Vec::new(),
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

    /// Write a row of values. Rows must be written in ascending order.
    /// Row numbers are 1-based.
    pub fn write_row(&mut self, row: u32, values: &[CellValue]) -> Result<()> {
        self.write_row_with_style_impl(row, values, None)
    }

    /// Write a row with a specific style ID applied to all cells.
    pub fn write_row_with_style(
        &mut self,
        row: u32,
        values: &[CellValue],
        style_id: u32,
    ) -> Result<()> {
        self.write_row_with_style_impl(row, values, Some(style_id))
    }

    /// Add a merge cell reference (e.g., "A1:B2").
    /// Must be called before finish().
    pub fn add_merge_cell(&mut self, reference: &str) -> Result<()> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        self.merge_cells.push(reference.to_string());
        Ok(())
    }

    /// Finish writing and return the complete worksheet XML bytes.
    pub fn finish(&mut self) -> Result<Vec<u8>> {
        if self.finished {
            return Err(Error::StreamAlreadyFinished);
        }
        self.finished = true;

        if !self.started {
            // No rows were written; write the full header now.
            self.write_header();
        }

        // Close sheetData.
        self.buffer.push_str("</sheetData>");

        // Write mergeCells if any.
        if !self.merge_cells.is_empty() {
            write!(
                self.buffer,
                "<mergeCells count=\"{}\">",
                self.merge_cells.len()
            )
            .unwrap();
            for mc in &self.merge_cells {
                write!(self.buffer, "<mergeCell ref=\"{}\"/>", xml_escape(mc)).unwrap();
            }
            self.buffer.push_str("</mergeCells>");
        }

        // Close worksheet.
        self.buffer.push_str("</worksheet>");

        Ok(self.buffer.as_bytes().to_vec())
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

    // --- Internal helpers ---

    /// Ensure the XML header and sheetData opening tag have been written.
    fn ensure_started(&mut self) {
        if !self.started {
            self.started = true;
            self.write_header();
        }
    }

    /// Write the XML declaration, worksheet opening tag, cols, and sheetData opening tag.
    fn write_header(&mut self) {
        write!(
            self.buffer,
            "{}\n<worksheet xmlns=\"{}\" xmlns:r=\"{}\">",
            XML_DECLARATION, NS_SPREADSHEET, NS_RELATIONSHIPS
        )
        .unwrap();

        // Write cols section if any.
        if !self.col_widths.is_empty() {
            self.buffer.push_str("<cols>");
            for &(min, max, width) in &self.col_widths {
                write!(
                    self.buffer,
                    "<col min=\"{}\" max=\"{}\" width=\"{}\" customWidth=\"1\"/>",
                    min, max, width
                )
                .unwrap();
            }
            self.buffer.push_str("</cols>");
        }

        self.buffer.push_str("<sheetData>");
    }

    /// Internal implementation for writing a row.
    fn write_row_with_style_impl(
        &mut self,
        row: u32,
        values: &[CellValue],
        style_id: Option<u32>,
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

        self.ensure_started();

        self.last_row = row;

        // Write the row opening tag.
        write!(self.buffer, "<row r=\"{}\">", row).unwrap();

        // Write each cell.
        for (i, value) in values.iter().enumerate() {
            let col = (i as u32) + 1;
            let cell_ref = coordinates_to_cell_name(col, row)?;
            self.write_cell(&cell_ref, value, style_id);
        }

        self.buffer.push_str("</row>");

        Ok(())
    }

    /// Write a single cell value as XML.
    fn write_cell(&mut self, cell_ref: &str, value: &CellValue, style_id: Option<u32>) {
        match value {
            CellValue::Empty => {
                // Skip empty cells entirely.
            }
            CellValue::String(s) => {
                let idx = self.sst.add(s);
                self.buffer.push_str("<c r=\"");
                self.buffer.push_str(cell_ref);
                self.buffer.push('"');
                if let Some(sid) = style_id {
                    write!(self.buffer, " s=\"{}\"", sid).unwrap();
                }
                write!(self.buffer, " t=\"s\"><v>{}</v></c>", idx).unwrap();
            }
            CellValue::Number(n) => {
                self.buffer.push_str("<c r=\"");
                self.buffer.push_str(cell_ref);
                self.buffer.push('"');
                if let Some(sid) = style_id {
                    write!(self.buffer, " s=\"{}\"", sid).unwrap();
                }
                write!(self.buffer, "><v>{}</v></c>", n).unwrap();
            }
            CellValue::Bool(b) => {
                let val = if *b { "1" } else { "0" };
                self.buffer.push_str("<c r=\"");
                self.buffer.push_str(cell_ref);
                self.buffer.push('"');
                if let Some(sid) = style_id {
                    write!(self.buffer, " s=\"{}\"", sid).unwrap();
                }
                write!(self.buffer, " t=\"b\"><v>{}</v></c>", val).unwrap();
            }
            CellValue::Formula { expr, .. } => {
                self.buffer.push_str("<c r=\"");
                self.buffer.push_str(cell_ref);
                self.buffer.push('"');
                if let Some(sid) = style_id {
                    write!(self.buffer, " s=\"{}\"", sid).unwrap();
                }
                write!(self.buffer, "><f>{}</f></c>", xml_escape(expr)).unwrap();
            }
            CellValue::Error(e) => {
                self.buffer.push_str("<c r=\"");
                self.buffer.push_str(cell_ref);
                self.buffer.push('"');
                if let Some(sid) = style_id {
                    write!(self.buffer, " s=\"{}\"", sid).unwrap();
                }
                write!(self.buffer, " t=\"e\"><v>{}</v></c>", xml_escape(e)).unwrap();
            }
        }
    }
}

/// Escape XML special characters in a string.
fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetkit_xml::worksheet::WorksheetXml;

    // -----------------------------------------------------------------------
    // Basic functionality
    // -----------------------------------------------------------------------

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
        assert!(xml.contains("<row r=\"1\">"));
        assert!(xml.contains("<row r=\"2\">"));
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
        assert_eq!(ws.sheet_data.rows[0].cells[0].t, Some("s".to_string()));
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");
        // Second cell is a number
        assert_eq!(ws.sheet_data.rows[0].cells[1].t, None);
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
        sw.write_row(1, &[CellValue::from(3.14)]).unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        assert!(xml.contains("<v>3.14</v>"));
    }

    #[test]
    fn test_bool_values() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from(true), CellValue::from(false)])
            .unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        assert!(xml.contains("t=\"b\""));
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

        assert!(xml.contains("<f>SUM(A2:A10)</f>"));
    }

    #[test]
    fn test_error_value() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::Error("#DIV/0!".to_string())])
            .unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        assert!(xml.contains("t=\"e\""));
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

    // -----------------------------------------------------------------------
    // Column widths
    // -----------------------------------------------------------------------

    #[test]
    fn test_set_col_width_before_rows() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.set_col_width(1, 20.0).unwrap();
        sw.write_row(1, &[CellValue::from("data")]).unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml = String::from_utf8(xml_bytes).unwrap();

        assert!(xml.contains("<cols>"));
        assert!(xml.contains("min=\"1\""));
        assert!(xml.contains("max=\"1\""));
        assert!(xml.contains("width=\"20\""));
        assert!(xml.contains("customWidth=\"1\""));
        assert!(xml.contains("</cols>"));
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
        let sheet_data_pos = xml_str.find("<sheetData>").unwrap();
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

    // -----------------------------------------------------------------------
    // Row ordering
    // -----------------------------------------------------------------------

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

    // -----------------------------------------------------------------------
    // Merge cells
    // -----------------------------------------------------------------------

    #[test]
    fn test_merge_cells_in_output() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Merged")]).unwrap();
        sw.add_merge_cell("A1:B1").unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        assert!(xml_str.contains("<mergeCells count=\"1\">"));
        assert!(xml_str.contains("<mergeCell ref=\"A1:B1\"/>"));
        assert!(xml_str.contains("</mergeCells>"));
    }

    #[test]
    fn test_multiple_merge_cells() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("a")]).unwrap();
        sw.add_merge_cell("A1:B1").unwrap();
        sw.add_merge_cell("C1:D1").unwrap();
        let xml_bytes = sw.finish().unwrap();
        let xml_str = String::from_utf8(xml_bytes).unwrap();

        assert!(xml_str.contains("count=\"2\""));

        let ws: WorksheetXml = quick_xml::de::from_str(&xml_str).unwrap();
        let mc = ws.merge_cells.unwrap();
        assert_eq!(mc.merge_cells.len(), 2);
        assert_eq!(mc.merge_cells[0].reference, "A1:B1");
        assert_eq!(mc.merge_cells[1].reference, "C1:D1");
    }

    // -----------------------------------------------------------------------
    // finish() behavior
    // -----------------------------------------------------------------------

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

    // -----------------------------------------------------------------------
    // into_parts
    // -----------------------------------------------------------------------

    #[test]
    fn test_into_parts() {
        let mut sw = StreamWriter::new("Sheet1");
        sw.write_row(1, &[CellValue::from("Hello")]).unwrap();
        let (xml_bytes, sst) = sw.into_parts().unwrap();

        assert!(!xml_bytes.is_empty());
        assert_eq!(sst.len(), 1);
        assert_eq!(sst.get(0), Some("Hello"));
    }

    // -----------------------------------------------------------------------
    // SST deduplication
    // -----------------------------------------------------------------------

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

    // -----------------------------------------------------------------------
    // Large-scale test
    // -----------------------------------------------------------------------

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

    // -----------------------------------------------------------------------
    // XML escaping
    // -----------------------------------------------------------------------

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
        // The quotes and > should be escaped
        assert!(xml.contains("&gt;"));
        assert!(xml.contains("&quot;"));
    }

    // -----------------------------------------------------------------------
    // Edge cases
    // -----------------------------------------------------------------------

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
                CellValue::from(3.14),
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
}
