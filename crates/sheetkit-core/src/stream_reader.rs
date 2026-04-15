//! Forward-only streaming worksheet reader.
//!
//! [`SheetStreamReader`] reads worksheet XML row-by-row using event-driven
//! parsing (`quick_xml::Reader`) without materializing the full DOM. This
//! enables processing large worksheets with bounded memory by reading rows
//! in batches.
//!
//! Shared string indices are resolved through a reference to the workbook's
//! [`SharedStringTable`]. Cell types (string, number, boolean, date, formula,
//! error, inline string) are handled according to the OOXML specification.

use std::io::BufRead;

use quick_xml::events::Event;
use quick_xml::name::QName;

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::sst::SharedStringTable;
use crate::utils::cell_ref::cell_name_to_coordinates;
use crate::workbook::open_options::DateInterpretation;

/// A single row produced by the streaming reader.
#[derive(Debug, Clone)]
pub struct StreamRow {
    /// 1-based row number.
    pub row_number: u32,
    /// Cells in this row as (1-based column index, value) pairs.
    pub cells: Vec<(u32, CellValue)>,
}

/// Forward-only streaming reader for worksheet XML.
///
/// Reads rows in batches without deserializing the entire worksheet into
/// memory. The reader borrows the shared string table for resolving string
/// cell references.
pub struct SheetStreamReader<'a, R: BufRead> {
    reader: quick_xml::Reader<R>,
    sst: &'a SharedStringTable,
    done: bool,
    row_limit: Option<u32>,
    rows_emitted: u32,
    date_interpretation: DateInterpretation,
    style_is_date: Vec<bool>,
}

impl<'a, R: BufRead> SheetStreamReader<'a, R> {
    /// Create a new streaming reader over the given `BufRead` source.
    ///
    /// `sst` is a reference to the shared string table for resolving
    /// shared string cell values. Optional configuration (row limit, date
    /// interpretation) is applied via the builder methods.
    pub fn new(source: R, sst: &'a SharedStringTable) -> Self {
        let mut reader = quick_xml::Reader::from_reader(source);
        reader.config_mut().trim_text(false);
        Self {
            reader,
            sst,
            done: false,
            row_limit: None,
            rows_emitted: 0,
            date_interpretation: DateInterpretation::default(),
            style_is_date: Vec::new(),
        }
    }

    /// Cap the number of rows produced by this reader. `None` means no
    /// limit (the default).
    pub fn row_limit(mut self, limit: Option<u32>) -> Self {
        self.row_limit = limit;
        self
    }

    /// Configure how date-formatted number cells are interpreted.
    ///
    /// `style_is_date` is a per-style boolean produced by
    /// [`crate::style::compute_style_is_date`]; it is indexed by each
    /// cell's `s` attribute. Empty vectors are equivalent to "no style
    /// was a date format", which is the safe default.
    pub fn date_promotion(
        mut self,
        interpretation: DateInterpretation,
        style_is_date: Vec<bool>,
    ) -> Self {
        self.date_interpretation = interpretation;
        self.style_is_date = style_is_date;
        self
    }

    /// Returns `true` when `style_idx` refers to a style whose number
    /// format is a date format and the reader is configured with
    /// [`DateInterpretation::NumFmt`].
    fn should_promote_to_date(&self, style_idx: Option<u32>) -> bool {
        if !matches!(self.date_interpretation, DateInterpretation::NumFmt) {
            return false;
        }
        let idx = match style_idx {
            Some(i) => i as usize,
            None => return false,
        };
        self.style_is_date.get(idx).copied().unwrap_or(false)
    }

    /// Read the next batch of rows. Returns an empty `Vec` when there are no
    /// more rows to read.
    pub fn next_batch(&mut self, batch_size: usize) -> Result<Vec<StreamRow>> {
        if self.done {
            return Ok(Vec::new());
        }

        let mut rows = Vec::with_capacity(batch_size);
        let mut buf = Vec::with_capacity(4096);

        loop {
            if rows.len() >= batch_size {
                break;
            }
            if let Some(limit) = self.row_limit {
                if self.rows_emitted >= limit {
                    self.done = true;
                    break;
                }
            }

            buf.clear();
            match self
                .reader
                .read_event_into(&mut buf)
                .map_err(|e| Error::XmlParse(e.to_string()))?
            {
                Event::Start(ref e) if e.name() == QName(b"row") => {
                    let row_number = extract_row_number(e)?;
                    let row = self.parse_row_body(row_number)?;
                    self.rows_emitted += 1;
                    if !row.cells.is_empty() {
                        rows.push(row);
                    }
                }
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }
        }

        Ok(rows)
    }

    /// Returns `true` if there are potentially more rows to read.
    pub fn has_more(&self) -> bool {
        !self.done
    }

    /// Close the reader and release resources.
    pub fn close(self) {
        drop(self);
    }

    /// Parse the body of a `<row>` element (its child `<c>` elements) after
    /// the row number has been extracted from the opening tag.
    fn parse_row_body(&mut self, row_number: u32) -> Result<StreamRow> {
        let mut cells = Vec::new();
        let mut buf = Vec::with_capacity(1024);

        loop {
            buf.clear();
            match self
                .reader
                .read_event_into(&mut buf)
                .map_err(|e| Error::XmlParse(e.to_string()))?
            {
                Event::Start(ref e) if e.name() == QName(b"c") => {
                    let (col, cell_type, style_idx) = extract_cell_attrs(e)?;
                    if let Some(col) = col {
                        let promote = self.should_promote_to_date(style_idx);
                        let cv = self.parse_cell_body(cell_type.as_deref(), promote)?;
                        cells.push((col, cv));
                    } else {
                        self.skip_to_end_of(b"c")?;
                    }
                }
                Event::Empty(ref e) if e.name() == QName(b"c") => {
                    let (col, cell_type, style_idx) = extract_cell_attrs(e)?;
                    if let Some(col) = col {
                        let promote = self.should_promote_to_date(style_idx);
                        let cv = resolve_cell_value(
                            self.sst,
                            cell_type.as_deref(),
                            None,
                            None,
                            None,
                            promote,
                        )?;
                        cells.push((col, cv));
                    }
                }
                Event::End(ref e) if e.name() == QName(b"row") => break,
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }
        }

        Ok(StreamRow { row_number, cells })
    }

    /// Parse the body of a `<c>` element (its child `<v>`, `<f>`, `<is>`
    /// elements) after the cell attributes have been extracted.
    fn parse_cell_body(
        &mut self,
        cell_type: Option<&str>,
        promote_to_date: bool,
    ) -> Result<CellValue> {
        let mut value_text: Option<String> = None;
        let mut formula_text: Option<String> = None;
        let mut inline_string: Option<String> = None;
        let mut buf = Vec::with_capacity(512);
        let mut in_is = false;

        loop {
            buf.clear();
            match self
                .reader
                .read_event_into(&mut buf)
                .map_err(|e| Error::XmlParse(e.to_string()))?
            {
                Event::Start(ref e) => {
                    let local = e.local_name();
                    if local.as_ref() == b"v" {
                        value_text = Some(self.read_text_content(b"v")?);
                    } else if local.as_ref() == b"f" {
                        formula_text = Some(self.read_text_content(b"f")?);
                    } else if local.as_ref() == b"is" {
                        in_is = true;
                        inline_string = Some(String::new());
                    } else if local.as_ref() == b"t" && in_is {
                        let t = self.read_text_content(b"t")?;
                        if let Some(ref mut is) = inline_string {
                            is.push_str(&t);
                        }
                    }
                }
                Event::End(ref e) => {
                    let local = e.local_name();
                    if local.as_ref() == b"c" {
                        break;
                    }
                    if local.as_ref() == b"is" {
                        in_is = false;
                    }
                }
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }
        }

        resolve_cell_value(
            self.sst,
            cell_type,
            value_text.as_deref(),
            formula_text,
            inline_string,
            promote_to_date,
        )
    }

    /// Read text content between an opening tag and its matching closing tag.
    fn read_text_content(&mut self, end_tag: &[u8]) -> Result<String> {
        let mut text = String::new();
        let mut buf = Vec::with_capacity(256);
        loop {
            buf.clear();
            match self
                .reader
                .read_event_into(&mut buf)
                .map_err(|e| Error::XmlParse(e.to_string()))?
            {
                Event::Text(ref e) => {
                    let decoded = e.unescape().map_err(|e| Error::XmlParse(e.to_string()))?;
                    text.push_str(&decoded);
                }
                Event::End(ref e) if e.local_name().as_ref() == end_tag => break,
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }
        }
        Ok(text)
    }

    /// Skip all events until the matching end tag for the given element.
    fn skip_to_end_of(&mut self, tag: &[u8]) -> Result<()> {
        let mut buf = Vec::with_capacity(256);
        let mut depth: u32 = 1;
        loop {
            buf.clear();
            match self
                .reader
                .read_event_into(&mut buf)
                .map_err(|e| Error::XmlParse(e.to_string()))?
            {
                Event::Start(ref e) if e.local_name().as_ref() == tag => {
                    depth += 1;
                }
                Event::End(ref e) if e.local_name().as_ref() == tag => {
                    depth -= 1;
                    if depth == 0 {
                        break;
                    }
                }
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }
        }
        Ok(())
    }
}

/// Owning variant of [`SheetStreamReader`] for use in FFI contexts where
/// lifetime parameters are not supported (e.g., napi classes).
///
/// Stores its own copy of the shared string table and the XML byte source,
/// avoiding any borrowed references. The parsing logic delegates to the same
/// free functions used by the borrowed reader.
pub struct OwnedSheetStreamReader {
    reader: quick_xml::Reader<std::io::BufReader<std::io::Cursor<Vec<u8>>>>,
    sst: SharedStringTable,
    done: bool,
    row_limit: Option<u32>,
    rows_emitted: u32,
    date_interpretation: DateInterpretation,
    style_is_date: Vec<bool>,
}

impl OwnedSheetStreamReader {
    /// Create a new owned streaming reader.
    ///
    /// `xml_bytes` is the raw worksheet XML. `sst` is a read-only clone of
    /// the shared string table. Configuration is applied via builder
    /// methods ([`Self::row_limit`], [`Self::date_promotion`]).
    pub fn new(xml_bytes: Vec<u8>, sst: SharedStringTable) -> Self {
        let cursor = std::io::Cursor::new(xml_bytes);
        let buf_reader = std::io::BufReader::new(cursor);
        let mut reader = quick_xml::Reader::from_reader(buf_reader);
        reader.config_mut().trim_text(false);
        Self {
            reader,
            sst,
            done: false,
            row_limit: None,
            rows_emitted: 0,
            date_interpretation: DateInterpretation::default(),
            style_is_date: Vec::new(),
        }
    }

    /// Cap the number of rows produced by this reader. `None` means no
    /// limit (the default).
    pub fn row_limit(mut self, limit: Option<u32>) -> Self {
        self.row_limit = limit;
        self
    }

    /// Configure how date-formatted number cells are interpreted.
    /// See [`SheetStreamReader::date_promotion`] for details.
    pub fn date_promotion(
        mut self,
        interpretation: DateInterpretation,
        style_is_date: Vec<bool>,
    ) -> Self {
        self.date_interpretation = interpretation;
        self.style_is_date = style_is_date;
        self
    }

    fn should_promote_to_date(&self, style_idx: Option<u32>) -> bool {
        if !matches!(self.date_interpretation, DateInterpretation::NumFmt) {
            return false;
        }
        let idx = match style_idx {
            Some(i) => i as usize,
            None => return false,
        };
        self.style_is_date.get(idx).copied().unwrap_or(false)
    }

    /// Read the next batch of rows. Returns an empty `Vec` when there are no
    /// more rows to read.
    pub fn next_batch(&mut self, batch_size: usize) -> Result<Vec<StreamRow>> {
        if self.done {
            return Ok(Vec::new());
        }

        let mut rows = Vec::with_capacity(batch_size);
        let mut buf = Vec::with_capacity(4096);

        loop {
            if rows.len() >= batch_size {
                break;
            }
            if let Some(limit) = self.row_limit {
                if self.rows_emitted >= limit {
                    self.done = true;
                    break;
                }
            }

            buf.clear();
            match self
                .reader
                .read_event_into(&mut buf)
                .map_err(|e| Error::XmlParse(e.to_string()))?
            {
                Event::Start(ref e) if e.name() == QName(b"row") => {
                    let row_number = extract_row_number(e)?;
                    let row = self.parse_row_body(row_number)?;
                    self.rows_emitted += 1;
                    if !row.cells.is_empty() {
                        rows.push(row);
                    }
                }
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }
        }

        Ok(rows)
    }

    /// Returns `true` if there are potentially more rows to read.
    pub fn has_more(&self) -> bool {
        !self.done
    }

    /// Close the reader and release resources.
    pub fn close(self) {
        drop(self);
    }

    fn parse_row_body(&mut self, row_number: u32) -> Result<StreamRow> {
        let mut cells = Vec::new();
        let mut buf = Vec::with_capacity(1024);

        loop {
            buf.clear();
            match self
                .reader
                .read_event_into(&mut buf)
                .map_err(|e| Error::XmlParse(e.to_string()))?
            {
                Event::Start(ref e) if e.name() == QName(b"c") => {
                    let (col, cell_type, style_idx) = extract_cell_attrs(e)?;
                    if let Some(col) = col {
                        let promote = self.should_promote_to_date(style_idx);
                        let cv = self.parse_cell_body(cell_type.as_deref(), promote)?;
                        cells.push((col, cv));
                    } else {
                        self.skip_to_end_of(b"c")?;
                    }
                }
                Event::Empty(ref e) if e.name() == QName(b"c") => {
                    let (col, cell_type, style_idx) = extract_cell_attrs(e)?;
                    if let Some(col) = col {
                        let promote = self.should_promote_to_date(style_idx);
                        let cv = resolve_cell_value(
                            &self.sst,
                            cell_type.as_deref(),
                            None,
                            None,
                            None,
                            promote,
                        )?;
                        cells.push((col, cv));
                    }
                }
                Event::End(ref e) if e.name() == QName(b"row") => break,
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }
        }

        Ok(StreamRow { row_number, cells })
    }

    fn parse_cell_body(
        &mut self,
        cell_type: Option<&str>,
        promote_to_date: bool,
    ) -> Result<CellValue> {
        let mut value_text: Option<String> = None;
        let mut formula_text: Option<String> = None;
        let mut inline_string: Option<String> = None;
        let mut buf = Vec::with_capacity(512);
        let mut in_is = false;

        loop {
            buf.clear();
            match self
                .reader
                .read_event_into(&mut buf)
                .map_err(|e| Error::XmlParse(e.to_string()))?
            {
                Event::Start(ref e) => {
                    let local = e.local_name();
                    if local.as_ref() == b"v" {
                        value_text = Some(self.read_text_content(b"v")?);
                    } else if local.as_ref() == b"f" {
                        formula_text = Some(self.read_text_content(b"f")?);
                    } else if local.as_ref() == b"is" {
                        in_is = true;
                        inline_string = Some(String::new());
                    } else if local.as_ref() == b"t" && in_is {
                        let t = self.read_text_content(b"t")?;
                        if let Some(ref mut is) = inline_string {
                            is.push_str(&t);
                        }
                    }
                }
                Event::End(ref e) => {
                    let local = e.local_name();
                    if local.as_ref() == b"c" {
                        break;
                    }
                    if local.as_ref() == b"is" {
                        in_is = false;
                    }
                }
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }
        }

        resolve_cell_value(
            &self.sst,
            cell_type,
            value_text.as_deref(),
            formula_text,
            inline_string,
            promote_to_date,
        )
    }

    fn read_text_content(&mut self, end_tag: &[u8]) -> Result<String> {
        let mut text = String::new();
        let mut buf = Vec::with_capacity(256);
        loop {
            buf.clear();
            match self
                .reader
                .read_event_into(&mut buf)
                .map_err(|e| Error::XmlParse(e.to_string()))?
            {
                Event::Text(ref e) => {
                    let decoded = e.unescape().map_err(|e| Error::XmlParse(e.to_string()))?;
                    text.push_str(&decoded);
                }
                Event::End(ref e) if e.local_name().as_ref() == end_tag => break,
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }
        }
        Ok(text)
    }

    fn skip_to_end_of(&mut self, tag: &[u8]) -> Result<()> {
        let mut buf = Vec::with_capacity(256);
        let mut depth: u32 = 1;
        loop {
            buf.clear();
            match self
                .reader
                .read_event_into(&mut buf)
                .map_err(|e| Error::XmlParse(e.to_string()))?
            {
                Event::Start(ref e) if e.local_name().as_ref() == tag => {
                    depth += 1;
                }
                Event::End(ref e) if e.local_name().as_ref() == tag => {
                    depth -= 1;
                    if depth == 0 {
                        break;
                    }
                }
                Event::Eof => {
                    self.done = true;
                    break;
                }
                _ => {}
            }
        }
        Ok(())
    }
}

/// Extract the `r` (row number) attribute from a `<row>` element.
fn extract_row_number(start: &quick_xml::events::BytesStart<'_>) -> Result<u32> {
    for attr in start.attributes().flatten() {
        if attr.key == QName(b"r") {
            let val =
                std::str::from_utf8(&attr.value).map_err(|e| Error::XmlParse(e.to_string()))?;
            return val
                .parse::<u32>()
                .map_err(|e| Error::XmlParse(format!("invalid row number: {e}")));
        }
    }
    Err(Error::XmlParse(
        "row element missing r attribute".to_string(),
    ))
}

/// Extract the cell reference (column index), type attribute, and style
/// index from a `<c>` element.
fn extract_cell_attrs(
    start: &quick_xml::events::BytesStart<'_>,
) -> Result<(Option<u32>, Option<String>, Option<u32>)> {
    let mut cell_ref: Option<String> = None;
    let mut cell_type: Option<String> = None;
    let mut style_idx: Option<u32> = None;

    for attr in start.attributes().flatten() {
        match attr.key {
            QName(b"r") => {
                cell_ref = Some(
                    std::str::from_utf8(&attr.value)
                        .map_err(|e| Error::XmlParse(e.to_string()))?
                        .to_string(),
                );
            }
            QName(b"t") => {
                cell_type = Some(
                    std::str::from_utf8(&attr.value)
                        .map_err(|e| Error::XmlParse(e.to_string()))?
                        .to_string(),
                );
            }
            QName(b"s") => {
                let raw =
                    std::str::from_utf8(&attr.value).map_err(|e| Error::XmlParse(e.to_string()))?;
                style_idx = raw.parse::<u32>().ok();
            }
            _ => {}
        }
    }

    let col = match &cell_ref {
        Some(r) => Some(cell_name_to_coordinates(r)?.0),
        None => None,
    };

    Ok((col, cell_type, style_idx))
}

/// Resolve cell type, value text, formula, and inline string into a `CellValue`.
///
/// `promote_to_date` is honored only for non-formula numeric cells (`t="n"`
/// or untyped). When true, such cells are returned as [`CellValue::Date`]
/// instead of [`CellValue::Number`], used by the
/// [`DateInterpretation::NumFmt`] path of the streaming reader.
fn resolve_cell_value(
    sst: &SharedStringTable,
    cell_type: Option<&str>,
    value_text: Option<&str>,
    formula_text: Option<String>,
    inline_string: Option<String>,
    promote_to_date: bool,
) -> Result<CellValue> {
    if let Some(formula) = formula_text {
        let cached = match (cell_type, value_text) {
            (Some("b"), Some(v)) => Some(Box::new(CellValue::Bool(v == "1"))),
            (Some("e"), Some(v)) => Some(Box::new(CellValue::Error(v.to_string()))),
            (Some("str"), Some(v)) => Some(Box::new(CellValue::String(v.to_string()))),
            (_, Some(v)) => v
                .parse::<f64>()
                .ok()
                .map(|n| Box::new(CellValue::Number(n))),
            _ => None,
        };
        return Ok(CellValue::Formula {
            expr: formula,
            result: cached,
        });
    }

    match (cell_type, value_text) {
        (Some("s"), Some(v)) => {
            let idx: usize = v
                .parse()
                .map_err(|_| Error::Internal(format!("invalid SST index: {v}")))?;
            let s = sst
                .get(idx)
                .ok_or_else(|| Error::Internal(format!("SST index {idx} out of bounds")))?;
            Ok(CellValue::String(s.to_string()))
        }
        (Some("b"), Some(v)) => Ok(CellValue::Bool(v == "1")),
        (Some("e"), Some(v)) => Ok(CellValue::Error(v.to_string())),
        (Some("inlineStr"), _) => Ok(CellValue::String(inline_string.unwrap_or_default())),
        (Some("str"), Some(v)) => Ok(CellValue::String(v.to_string())),
        (Some("d"), Some(v)) => {
            let n: f64 = v
                .parse()
                .map_err(|_| Error::Internal(format!("invalid date value: {v}")))?;
            Ok(CellValue::Date(n))
        }
        (Some("n") | None, Some(v)) => {
            let n: f64 = v
                .parse()
                .map_err(|_| Error::Internal(format!("invalid number: {v}")))?;
            if promote_to_date {
                Ok(CellValue::Date(n))
            } else {
                Ok(CellValue::Number(n))
            }
        }
        _ => Ok(CellValue::Empty),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    fn make_sst(strings: &[&str]) -> SharedStringTable {
        let mut sst = SharedStringTable::new();
        for s in strings {
            sst.add(s);
        }
        sst
    }

    fn worksheet_xml(sheet_data: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
{sheet_data}
</sheetData>
</worksheet>"#
        )
    }

    fn read_all(xml: &str, sst: &SharedStringTable, row_limit: Option<u32>) -> Vec<StreamRow> {
        let cursor = Cursor::new(xml.as_bytes().to_vec());
        let mut reader = SheetStreamReader::new(cursor, sst).row_limit(row_limit);
        let mut all = Vec::new();
        loop {
            let batch = reader.next_batch(100).unwrap();
            if batch.is_empty() {
                break;
            }
            all.extend(batch);
        }
        all
    }

    #[test]
    fn test_basic_batch_reading() {
        let sst = make_sst(&["Name", "Age"]);
        let xml = worksheet_xml(
            r#"
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
<row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2"><v>30</v></c></row>
<row r="3"><c r="A3" t="s"><v>0</v></c><c r="B3"><v>25</v></c></row>
"#,
        );

        let cursor = Cursor::new(xml.as_bytes().to_vec());
        let mut reader = SheetStreamReader::new(cursor, &sst);

        let batch1 = reader.next_batch(2).unwrap();
        assert_eq!(batch1.len(), 2);
        assert!(reader.has_more());

        let batch2 = reader.next_batch(2).unwrap();
        assert_eq!(batch2.len(), 1);

        let batch3 = reader.next_batch(2).unwrap();
        assert!(batch3.is_empty());
        assert!(!reader.has_more());
    }

    #[test]
    fn test_sparse_rows() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(
            r#"
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="5"><c r="C5"><v>5</v></c></row>
<row r="100"><c r="A100"><v>100</v></c></row>
"#,
        );

        let rows = read_all(&xml, &sst, None);
        assert_eq!(rows.len(), 3);
        assert_eq!(rows[0].row_number, 1);
        assert_eq!(rows[1].row_number, 5);
        assert_eq!(rows[1].cells[0].0, 3);
        assert_eq!(rows[2].row_number, 100);
    }

    #[test]
    fn test_all_cell_types() {
        let sst = make_sst(&["Hello"]);
        let xml = worksheet_xml(
            r#"
<row r="1">
  <c r="A1" t="s"><v>0</v></c>
  <c r="B1"><v>42.5</v></c>
  <c r="C1" t="b"><v>1</v></c>
  <c r="D1" t="e"><v>#DIV/0!</v></c>
  <c r="E1" t="inlineStr"><is><t>Inline</t></is></c>
  <c r="F1" t="n"><v>99</v></c>
  <c r="G1" t="d"><v>45000</v></c>
</row>
"#,
        );

        let rows = read_all(&xml, &sst, None);
        assert_eq!(rows.len(), 1);
        let cells = &rows[0].cells;

        assert_eq!(cells[0], (1, CellValue::String("Hello".to_string())));
        assert_eq!(cells[1], (2, CellValue::Number(42.5)));
        assert_eq!(cells[2], (3, CellValue::Bool(true)));
        assert_eq!(cells[3], (4, CellValue::Error("#DIV/0!".to_string())));
        assert_eq!(cells[4], (5, CellValue::String("Inline".to_string())));
        assert_eq!(cells[5], (6, CellValue::Number(99.0)));
        assert_eq!(cells[6], (7, CellValue::Date(45000.0)));
    }

    #[test]
    fn test_boolean_false() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(r#"<row r="1"><c r="A1" t="b"><v>0</v></c></row>"#);
        let rows = read_all(&xml, &sst, None);
        assert_eq!(rows[0].cells[0].1, CellValue::Bool(false));
    }

    #[test]
    fn test_shared_string_resolution() {
        let sst = make_sst(&["First", "Second", "Third"]);
        let xml = worksheet_xml(
            r#"
<row r="1">
  <c r="A1" t="s"><v>0</v></c>
  <c r="B1" t="s"><v>1</v></c>
  <c r="C1" t="s"><v>2</v></c>
</row>
"#,
        );

        let rows = read_all(&xml, &sst, None);
        assert_eq!(rows[0].cells[0].1, CellValue::String("First".to_string()));
        assert_eq!(rows[0].cells[1].1, CellValue::String("Second".to_string()));
        assert_eq!(rows[0].cells[2].1, CellValue::String("Third".to_string()));
    }

    #[test]
    fn test_shared_string_out_of_bounds() {
        let sst = make_sst(&["Only"]);
        let xml = worksheet_xml(r#"<row r="1"><c r="A1" t="s"><v>999</v></c></row>"#);

        let cursor = Cursor::new(xml.as_bytes().to_vec());
        let mut reader = SheetStreamReader::new(cursor, &sst);
        let result = reader.next_batch(10);
        assert!(result.is_err());
    }

    #[test]
    fn test_row_limit() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(
            r#"
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="2"><c r="A2"><v>2</v></c></row>
<row r="3"><c r="A3"><v>3</v></c></row>
<row r="4"><c r="A4"><v>4</v></c></row>
<row r="5"><c r="A5"><v>5</v></c></row>
"#,
        );

        let rows = read_all(&xml, &sst, Some(3));
        assert_eq!(rows.len(), 3);
        assert_eq!(rows[0].row_number, 1);
        assert_eq!(rows[2].row_number, 3);
    }

    #[test]
    fn test_row_limit_zero() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#);

        let rows = read_all(&xml, &sst, Some(0));
        assert!(rows.is_empty());
    }

    #[test]
    fn test_empty_sheet() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml("");

        let rows = read_all(&xml, &sst, None);
        assert!(rows.is_empty());
    }

    #[test]
    fn test_empty_rows_are_skipped() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(
            r#"
<row r="1"></row>
<row r="2"><c r="A2"><v>42</v></c></row>
<row r="3"></row>
"#,
        );

        let rows = read_all(&xml, &sst, None);
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0].row_number, 2);
    }

    #[test]
    fn test_empty_rows_count_against_limit() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(
            r#"
<row r="1"></row>
<row r="2"></row>
<row r="3"><c r="A3"><v>3</v></c></row>
<row r="4"><c r="A4"><v>4</v></c></row>
"#,
        );

        let rows = read_all(&xml, &sst, Some(2));
        assert!(
            rows.is_empty(),
            "with limit=2 and 2 empty rows, no data rows should be returned"
        );

        let rows2 = read_all(&xml, &sst, Some(3));
        assert_eq!(rows2.len(), 1);
        assert_eq!(rows2[0].row_number, 3);
    }

    #[test]
    fn test_formula_with_cached_number() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(r#"<row r="1"><c r="A1"><f>SUM(B1:B10)</f><v>42</v></c></row>"#);

        let rows = read_all(&xml, &sst, None);
        match &rows[0].cells[0].1 {
            CellValue::Formula { expr, result } => {
                assert_eq!(expr, "SUM(B1:B10)");
                assert_eq!(result.as_deref(), Some(&CellValue::Number(42.0)));
            }
            other => panic!("expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_formula_with_cached_string() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(
            r#"<row r="1"><c r="A1" t="str"><f>CONCAT("a","b")</f><v>ab</v></c></row>"#,
        );

        let rows = read_all(&xml, &sst, None);
        match &rows[0].cells[0].1 {
            CellValue::Formula { expr, result } => {
                assert_eq!(expr, r#"CONCAT("a","b")"#);
                assert_eq!(
                    result.as_deref(),
                    Some(&CellValue::String("ab".to_string()))
                );
            }
            other => panic!("expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_formula_with_cached_boolean() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(r#"<row r="1"><c r="A1" t="b"><f>TRUE()</f><v>1</v></c></row>"#);

        let rows = read_all(&xml, &sst, None);
        match &rows[0].cells[0].1 {
            CellValue::Formula { expr, result } => {
                assert_eq!(expr, "TRUE()");
                assert_eq!(result.as_deref(), Some(&CellValue::Bool(true)));
            }
            other => panic!("expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_formula_with_cached_error() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(r#"<row r="1"><c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c></row>"#);

        let rows = read_all(&xml, &sst, None);
        match &rows[0].cells[0].1 {
            CellValue::Formula { expr, result } => {
                assert_eq!(expr, "1/0");
                assert_eq!(
                    result.as_deref(),
                    Some(&CellValue::Error("#DIV/0!".to_string()))
                );
            }
            other => panic!("expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_formula_without_cached_value() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(r#"<row r="1"><c r="A1"><f>A2+A3</f></c></row>"#);

        let rows = read_all(&xml, &sst, None);
        match &rows[0].cells[0].1 {
            CellValue::Formula { expr, result } => {
                assert_eq!(expr, "A2+A3");
                assert!(result.is_none());
            }
            other => panic!("expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_inline_string_with_rich_text_runs() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(
            r#"<row r="1"><c r="A1" t="inlineStr"><is><r><t>Bold</t></r><r><t> Normal</t></r></is></c></row>"#,
        );

        let rows = read_all(&xml, &sst, None);
        assert_eq!(
            rows[0].cells[0].1,
            CellValue::String("Bold Normal".to_string())
        );
    }

    #[test]
    fn test_reader_close() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#);
        let cursor = Cursor::new(xml.as_bytes().to_vec());
        let reader = SheetStreamReader::new(cursor, &sst);
        reader.close();
    }

    #[test]
    fn test_reader_drop_without_reading_all() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(
            r#"
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="2"><c r="A2"><v>2</v></c></row>
"#,
        );
        let cursor = Cursor::new(xml.as_bytes().to_vec());
        let mut reader = SheetStreamReader::new(cursor, &sst);
        let batch = reader.next_batch(1).unwrap();
        assert_eq!(batch.len(), 1);
        drop(reader);
    }

    #[test]
    fn test_has_more_transitions() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#);

        let cursor = Cursor::new(xml.as_bytes().to_vec());
        let mut reader = SheetStreamReader::new(cursor, &sst);
        assert!(reader.has_more());

        let batch = reader.next_batch(100).unwrap();
        assert_eq!(batch.len(), 1);

        let batch2 = reader.next_batch(100).unwrap();
        assert!(batch2.is_empty());
        assert!(!reader.has_more());
    }

    #[test]
    fn test_batch_size_one() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(
            r#"
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="2"><c r="A2"><v>2</v></c></row>
<row r="3"><c r="A3"><v>3</v></c></row>
"#,
        );

        let cursor = Cursor::new(xml.as_bytes().to_vec());
        let mut reader = SheetStreamReader::new(cursor, &sst);

        for expected_row in 1..=3 {
            let batch = reader.next_batch(1).unwrap();
            assert_eq!(batch.len(), 1);
            assert_eq!(batch[0].row_number, expected_row);
        }

        let batch = reader.next_batch(1).unwrap();
        assert!(batch.is_empty());
    }

    #[test]
    fn test_cell_with_no_value() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(r#"<row r="1"><c r="A1"></c><c r="B1"><v>42</v></c></row>"#);

        let rows = read_all(&xml, &sst, None);
        assert_eq!(rows[0].cells.len(), 2);
        assert_eq!(rows[0].cells[0].1, CellValue::Empty);
        assert_eq!(rows[0].cells[1].1, CellValue::Number(42.0));
    }

    #[test]
    fn test_self_closing_cell_element() {
        let sst = SharedStringTable::new();
        let xml = worksheet_xml(
            r#"<row r="1"><c r="A1"/><c r="B1"><v>42</v></c><c r="C1" t="b"/></row>"#,
        );

        let rows = read_all(&xml, &sst, None);
        assert_eq!(rows[0].cells.len(), 3);
        assert_eq!(rows[0].cells[0], (1, CellValue::Empty));
        assert_eq!(rows[0].cells[1], (2, CellValue::Number(42.0)));
        assert_eq!(rows[0].cells[2], (3, CellValue::Empty));
    }

    #[test]
    fn test_integration_with_saved_workbook() {
        let mut wb = crate::workbook::Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Name").unwrap();
        wb.set_cell_value("Sheet1", "B1", "Score").unwrap();
        wb.set_cell_value("Sheet1", "A2", "Alice").unwrap();
        wb.set_cell_value("Sheet1", "B2", 95.5f64).unwrap();
        wb.set_cell_value("Sheet1", "A3", "Bob").unwrap();
        wb.set_cell_value("Sheet1", "B3", 87.0f64).unwrap();

        let dir = tempfile::TempDir::new().unwrap();
        let path = dir.path().join("stream_reader_test.xlsx");
        wb.save(&path).unwrap();

        let wb2 = crate::workbook::Workbook::open_with_options(
            &path,
            &crate::workbook::OpenOptions::new().read_mode(crate::workbook::ReadMode::Lazy),
        )
        .unwrap();

        let mut reader = wb2.open_sheet_reader("Sheet1").unwrap();
        let rows = reader.next_batch(100).unwrap();

        assert_eq!(rows.len(), 3);
        assert_eq!(rows[0].row_number, 1);
        assert_eq!(rows[0].cells[0].1, CellValue::String("Name".to_string()));
        assert_eq!(rows[0].cells[1].1, CellValue::String("Score".to_string()));
        assert_eq!(rows[1].cells[0].1, CellValue::String("Alice".to_string()));
        assert_eq!(rows[1].cells[1].1, CellValue::Number(95.5));
        assert_eq!(rows[2].cells[0].1, CellValue::String("Bob".to_string()));
        assert_eq!(rows[2].cells[1].1, CellValue::Number(87.0));
    }

    #[test]
    fn test_integration_with_row_limit() {
        let mut wb = crate::workbook::Workbook::new();
        for i in 1..=10 {
            let cell = format!("A{i}");
            wb.set_cell_value("Sheet1", &cell, i as f64).unwrap();
        }

        let dir = tempfile::TempDir::new().unwrap();
        let path = dir.path().join("stream_limit_test.xlsx");
        wb.save(&path).unwrap();

        let wb2 = crate::workbook::Workbook::open_with_options(
            &path,
            &crate::workbook::OpenOptions::new()
                .read_mode(crate::workbook::ReadMode::Lazy)
                .sheet_rows(5),
        )
        .unwrap();

        let mut reader = wb2.open_sheet_reader("Sheet1").unwrap();
        let mut all_rows = Vec::new();
        loop {
            let batch = reader.next_batch(3).unwrap();
            if batch.is_empty() {
                break;
            }
            all_rows.extend(batch);
        }

        assert_eq!(all_rows.len(), 5);
        assert_eq!(all_rows[4].row_number, 5);
    }

    #[test]
    fn test_integration_sheet_not_found() {
        let wb = crate::workbook::Workbook::new();
        let result = wb.open_sheet_reader("NonExistent");
        assert!(result.is_err());
    }

    /// End-to-end date interpretation test harness. Builds a workbook with
    /// mixed styled cells, saves it, reopens under both policies, and lets
    /// the caller assert on the returned cell values.
    fn write_mixed_style_workbook(path: &std::path::Path) {
        use crate::style::{builtin_num_fmts, NumFmtStyle, Style};
        let mut wb = crate::workbook::Workbook::new();
        let builtin_date_style = wb
            .add_style(&Style {
                num_fmt: Some(NumFmtStyle::Builtin(builtin_num_fmts::DATE_MDY)),
                ..Style::default()
            })
            .unwrap();
        let custom_date_style = wb
            .add_style(&Style {
                num_fmt: Some(NumFmtStyle::Custom("yyyy-mm-dd hh:mm".to_string())),
                ..Style::default()
            })
            .unwrap();
        let decimal_style = wb
            .add_style(&Style {
                num_fmt: Some(NumFmtStyle::Builtin(builtin_num_fmts::DECIMAL_2)),
                ..Style::default()
            })
            .unwrap();

        // A1: number + built-in date format.
        wb.set_cell_value("Sheet1", "A1", 46127.0_f64).unwrap();
        wb.set_cell_style("Sheet1", "A1", builtin_date_style)
            .unwrap();
        // B1: number + custom date-like format.
        wb.set_cell_value("Sheet1", "B1", 46127.9993_f64).unwrap();
        wb.set_cell_style("Sheet1", "B1", custom_date_style)
            .unwrap();
        // C1: number + non-date format (decimal 2).
        wb.set_cell_value("Sheet1", "C1", 2.5_f64).unwrap();
        wb.set_cell_style("Sheet1", "C1", decimal_style).unwrap();
        // D1: number with no explicit style.
        wb.set_cell_value("Sheet1", "D1", 42.0_f64).unwrap();

        wb.save(path).unwrap();
    }

    #[test]
    fn test_integration_date_interpretation_cell_type_default() {
        let dir = tempfile::TempDir::new().unwrap();
        let path = dir.path().join("dates_cell_type.xlsx");
        write_mixed_style_workbook(&path);

        let wb = crate::workbook::Workbook::open_with_options(
            &path,
            &crate::workbook::OpenOptions::new().read_mode(crate::workbook::ReadMode::Lazy),
        )
        .unwrap();
        let mut reader = wb.open_sheet_reader("Sheet1").unwrap();
        let rows = reader.next_batch(10).unwrap();

        // All number cells remain Number under the default policy, regardless
        // of whether their style references a date format.
        assert_eq!(rows[0].cells[0].1, CellValue::Number(46127.0));
        match &rows[0].cells[1].1 {
            CellValue::Number(v) => assert!((*v - 46127.9993).abs() < 1e-9),
            other => panic!("expected Number, got {:?}", other),
        }
        assert_eq!(rows[0].cells[2].1, CellValue::Number(2.5));
        assert_eq!(rows[0].cells[3].1, CellValue::Number(42.0));
    }

    #[test]
    fn test_integration_date_interpretation_num_fmt_promotes_date_styles() {
        let dir = tempfile::TempDir::new().unwrap();
        let path = dir.path().join("dates_num_fmt.xlsx");
        write_mixed_style_workbook(&path);

        let wb = crate::workbook::Workbook::open_with_options(
            &path,
            &crate::workbook::OpenOptions::new()
                .read_mode(crate::workbook::ReadMode::Lazy)
                .date_interpretation(crate::workbook::DateInterpretation::NumFmt),
        )
        .unwrap();
        let mut reader = wb.open_sheet_reader("Sheet1").unwrap();
        let rows = reader.next_batch(10).unwrap();

        // A1: built-in date format → promoted to Date.
        assert_eq!(rows[0].cells[0].1, CellValue::Date(46127.0));
        // B1: custom date format code → promoted to Date.
        match &rows[0].cells[1].1 {
            CellValue::Date(v) => assert!((*v - 46127.9993).abs() < 1e-9),
            other => panic!("expected Date, got {:?}", other),
        }
        // C1: non-date format → stays Number.
        assert_eq!(rows[0].cells[2].1, CellValue::Number(2.5));
        // D1: no style → stays Number.
        assert_eq!(rows[0].cells[3].1, CellValue::Number(42.0));
    }
}
