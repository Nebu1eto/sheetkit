//! Read-direction buffer serializer v2 for bulk data transfer.
//!
//! Converts a `WorksheetXml` and `SharedStringTable` into a compact binary
//! buffer with inline strings, eliminating the global string table that v1
//! required. This allows incremental (row-by-row) decoding on the JS side
//! without eagerly materializing all strings upfront.
//!
//! Binary format (little-endian throughout):
//!
//! ```text
//! HEADER (16 bytes)
//!   magic:     u32  = 0x534B5232 ("SKR2")
//!   version:   u16  = 2
//!   row_count: u32  = number of rows
//!   col_count: u16  = number of columns
//!   flags:     u32  = bit 0: always 0 (v2 is always sparse)
//!                      bits 16..31: min_col (1-based)
//!
//! ROW INDEX (row_count * 8 bytes)
//!   per row: row_number (u32) + byte_offset (u32) into CELL DATA section
//!   offset = 0xFFFFFFFF for empty rows
//!
//! CELL DATA (variable length)
//!   per row: cell_count (u16) + cells
//!   per cell: col (u16) + type (u8) + payload
//!     0x00 (empty):  0 bytes
//!     0x01 (number): 8 bytes (f64 LE)
//!     0x02 (string): 4 bytes (len u32 LE) + N bytes (UTF-8)
//!     0x03 (bool):   1 byte
//!     0x04 (date):   8 bytes (f64 LE, Excel serial)
//!     0x05 (error):  4 bytes (len u32 LE) + N bytes (UTF-8)
//!     0x06 (formula): 4 bytes (len u32 LE) + N bytes (UTF-8, cached formula text)
//!     0x07 (rich):   4 bytes (len u32 LE) + N bytes (UTF-8)
//! ```

use crate::error::Result;
use crate::sst::SharedStringTable;
use crate::utils::cell_ref::cell_name_to_coordinates;
use sheetkit_xml::worksheet::{CellTypeTag, WorksheetXml};

pub const MAGIC_V2: u32 = 0x534B5232;
pub const VERSION_V2: u16 = 2;
pub const HEADER_SIZE: usize = 16;

pub const TYPE_EMPTY: u8 = 0x00;
pub const TYPE_NUMBER: u8 = 0x01;
pub const TYPE_STRING: u8 = 0x02;
pub const TYPE_BOOL: u8 = 0x03;
pub const TYPE_DATE: u8 = 0x04;
pub const TYPE_ERROR: u8 = 0x05;
pub const TYPE_FORMULA: u8 = 0x06;
pub const TYPE_RICH_STRING: u8 = 0x07;

const EMPTY_ROW_OFFSET: u32 = 0xFFFF_FFFF;

/// A cell entry with inline string data (variable-length payload).
struct CellEntryV2 {
    col: u16,
    type_tag: u8,
    payload: CellPayload,
}

enum CellPayload {
    Empty,
    Number(f64),
    Bool(u8),
    Str(String),
}

impl CellPayload {
    fn byte_size(&self) -> usize {
        match self {
            CellPayload::Empty => 0,
            CellPayload::Number(_) => 8,
            CellPayload::Bool(_) => 1,
            CellPayload::Str(s) => 4 + s.len(),
        }
    }

    fn write_to(&self, buf: &mut Vec<u8>) {
        match self {
            CellPayload::Empty => {}
            CellPayload::Number(n) => {
                buf.extend_from_slice(&n.to_le_bytes());
            }
            CellPayload::Bool(b) => {
                buf.push(*b);
            }
            CellPayload::Str(s) => {
                buf.extend_from_slice(&(s.len() as u32).to_le_bytes());
                buf.extend_from_slice(s.as_bytes());
            }
        }
    }
}

struct RowEntriesV2 {
    row_num: u32,
    cells: Vec<CellEntryV2>,
}

impl RowEntriesV2 {
    /// Byte size of the cell data for this row (cell_count u16 + all cells).
    fn byte_size(&self) -> usize {
        // cell_count (2 bytes) + each cell (col u16 + type u8 + payload)
        let mut size = 2usize;
        for cell in &self.cells {
            size += 2 + 1 + cell.payload.byte_size();
        }
        size
    }
}

/// Serialize a worksheet's cell data into a compact binary buffer using v2 format.
///
/// The v2 format inlines all string data with each cell, eliminating the global
/// string table. This enables incremental row-by-row decoding without eagerly
/// decoding all strings.
pub fn sheet_to_raw_buffer_v2(ws: &WorksheetXml, sst: &SharedStringTable) -> Result<Vec<u8>> {
    let rows = &ws.sheet_data.rows;

    if rows.is_empty() {
        return Ok(write_empty_buffer());
    }

    let (min_row, max_row, min_col, max_col, total_cells) = scan_dimensions(ws)?;

    if total_cells == 0 {
        return Ok(write_empty_buffer());
    }

    let row_count = (max_row - min_row + 1) as usize;
    let col_count = (max_col - min_col + 1) as usize;

    let cell_entries = collect_cell_entries_v2(ws, sst, min_col)?;

    let flags: u32 = (min_col & 0xFFFF) << 16;

    let row_index_size = row_count * 8;

    let cell_data = encode_cell_data_v2(&cell_entries, min_row, max_row);
    let row_index = build_row_index_v2(&cell_entries, min_row, max_row);

    let total_size = HEADER_SIZE + row_index_size + cell_data.len();
    let mut buf = Vec::with_capacity(total_size);

    buf.extend_from_slice(&MAGIC_V2.to_le_bytes());
    buf.extend_from_slice(&VERSION_V2.to_le_bytes());
    buf.extend_from_slice(&(row_count as u32).to_le_bytes());
    buf.extend_from_slice(&(col_count as u16).to_le_bytes());
    buf.extend_from_slice(&flags.to_le_bytes());

    buf.extend_from_slice(&row_index);
    buf.extend_from_slice(&cell_data);

    Ok(buf)
}

fn write_empty_buffer() -> Vec<u8> {
    let mut buf = Vec::with_capacity(HEADER_SIZE);
    buf.extend_from_slice(&MAGIC_V2.to_le_bytes());
    buf.extend_from_slice(&VERSION_V2.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf.extend_from_slice(&0u16.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf
}

fn scan_dimensions(ws: &WorksheetXml) -> Result<(u32, u32, u32, u32, usize)> {
    let mut min_row = u32::MAX;
    let mut max_row = 0u32;
    let mut min_col = u32::MAX;
    let mut max_col = 0u32;
    let mut total_cells = 0usize;

    for row in &ws.sheet_data.rows {
        if row.cells.is_empty() {
            continue;
        }
        min_row = min_row.min(row.r);
        max_row = max_row.max(row.r);

        for cell in &row.cells {
            let col = resolve_col(cell)?;
            min_col = min_col.min(col);
            max_col = max_col.max(col);
            total_cells += 1;
        }
    }

    if total_cells == 0 {
        return Ok((1, 1, 1, 1, 0));
    }

    Ok((min_row, max_row, min_col, max_col, total_cells))
}

fn resolve_col(cell: &sheetkit_xml::worksheet::Cell) -> Result<u32> {
    if cell.col > 0 {
        return Ok(cell.col);
    }
    let (col, _row) = cell_name_to_coordinates(cell.r.as_str())?;
    Ok(col)
}

fn collect_cell_entries_v2(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    min_col: u32,
) -> Result<Vec<RowEntriesV2>> {
    let mut result = Vec::with_capacity(ws.sheet_data.rows.len());

    for row in &ws.sheet_data.rows {
        if row.cells.is_empty() {
            continue;
        }

        let mut cells = Vec::with_capacity(row.cells.len());
        for cell in &row.cells {
            let col = resolve_col(cell)?;
            let relative_col = (col - min_col) as u16;
            let (type_tag, payload) = encode_cell_value_v2(cell, sst)?;
            cells.push(CellEntryV2 {
                col: relative_col,
                type_tag,
                payload,
            });
        }

        result.push(RowEntriesV2 {
            row_num: row.r,
            cells,
        });
    }

    Ok(result)
}

fn encode_cell_value_v2(
    cell: &sheetkit_xml::worksheet::Cell,
    sst: &SharedStringTable,
) -> Result<(u8, CellPayload)> {
    if cell.f.is_some() {
        let formula_expr = cell
            .f
            .as_ref()
            .and_then(|f| f.value.as_deref())
            .unwrap_or("");
        return Ok((TYPE_FORMULA, CellPayload::Str(formula_expr.to_string())));
    }

    match cell.t {
        CellTypeTag::SharedString => {
            if let Some(ref v) = cell.v {
                if let Ok(sst_idx) = v.parse::<usize>() {
                    let text = sst.get(sst_idx).unwrap_or("");
                    if sst.get_rich_text(sst_idx).is_some() {
                        return Ok((TYPE_RICH_STRING, CellPayload::Str(text.to_string())));
                    }
                    return Ok((TYPE_STRING, CellPayload::Str(text.to_string())));
                }
            }
            Ok((TYPE_EMPTY, CellPayload::Empty))
        }
        CellTypeTag::Boolean => {
            let b = if let Some(ref v) = cell.v {
                if v == "1" || v.eq_ignore_ascii_case("true") {
                    1u8
                } else {
                    0u8
                }
            } else {
                0u8
            };
            Ok((TYPE_BOOL, CellPayload::Bool(b)))
        }
        CellTypeTag::Error => {
            let error_text = cell.v.as_deref().unwrap_or("#VALUE!");
            Ok((TYPE_ERROR, CellPayload::Str(error_text.to_string())))
        }
        CellTypeTag::InlineString => {
            let text = cell
                .is
                .as_ref()
                .and_then(|is| is.t.as_deref())
                .or(cell.v.as_deref())
                .unwrap_or("");
            Ok((TYPE_STRING, CellPayload::Str(text.to_string())))
        }
        CellTypeTag::Date => {
            if let Some(ref v) = cell.v {
                if let Ok(n) = v.parse::<f64>() {
                    return Ok((TYPE_DATE, CellPayload::Number(n)));
                }
            }
            Ok((TYPE_EMPTY, CellPayload::Empty))
        }
        CellTypeTag::FormulaString => {
            if let Some(ref v) = cell.v {
                return Ok((TYPE_STRING, CellPayload::Str(v.clone())));
            }
            Ok((TYPE_EMPTY, CellPayload::Empty))
        }
        CellTypeTag::None | CellTypeTag::Number => {
            if let Some(ref v) = cell.v {
                if let Ok(n) = v.parse::<f64>() {
                    return Ok((TYPE_NUMBER, CellPayload::Number(n)));
                }
            }
            Ok((TYPE_EMPTY, CellPayload::Empty))
        }
    }
}

fn encode_cell_data_v2(row_entries: &[RowEntriesV2], min_row: u32, max_row: u32) -> Vec<u8> {
    let total_rows = (max_row - min_row + 1) as usize;
    let mut entries_by_row: Vec<Option<&RowEntriesV2>> = vec![None; total_rows];
    for re in row_entries {
        let idx = (re.row_num - min_row) as usize;
        entries_by_row[idx] = Some(re);
    }

    let mut buf = Vec::new();
    for entry in &entries_by_row {
        match entry {
            Some(re) => {
                let count = re.cells.len() as u16;
                buf.extend_from_slice(&count.to_le_bytes());
                for cell in &re.cells {
                    buf.extend_from_slice(&cell.col.to_le_bytes());
                    buf.push(cell.type_tag);
                    cell.payload.write_to(&mut buf);
                }
            }
            None => {
                buf.extend_from_slice(&0u16.to_le_bytes());
            }
        }
    }

    buf
}

fn build_row_index_v2(row_entries: &[RowEntriesV2], min_row: u32, max_row: u32) -> Vec<u8> {
    let total_rows = (max_row - min_row + 1) as usize;
    let mut index = Vec::with_capacity(total_rows * 8);

    let mut entries_by_row: Vec<Option<&RowEntriesV2>> = vec![None; total_rows];
    for re in row_entries {
        let idx = (re.row_num - min_row) as usize;
        entries_by_row[idx] = Some(re);
    }

    let mut byte_offset = 0u32;
    for (i, entry) in entries_by_row.iter().enumerate() {
        let row_num = min_row + i as u32;
        index.extend_from_slice(&row_num.to_le_bytes());
        match entry {
            Some(re) => {
                if re.cells.is_empty() {
                    index.extend_from_slice(&EMPTY_ROW_OFFSET.to_le_bytes());
                    byte_offset += 2; // just the cell_count u16
                } else {
                    index.extend_from_slice(&byte_offset.to_le_bytes());
                    byte_offset += re.byte_size() as u32;
                }
            }
            None => {
                index.extend_from_slice(&EMPTY_ROW_OFFSET.to_le_bytes());
                byte_offset += 2; // empty row writes cell_count=0
            }
        }
    }

    index
}

#[cfg(test)]
#[allow(clippy::field_reassign_with_default)]
mod tests {
    use super::*;
    use sheetkit_xml::worksheet::{
        Cell, CellFormula, CellTypeTag, CompactCellRef, InlineString, Row, SheetData, WorksheetXml,
    };

    fn make_cell(col_ref: &str, col_num: u32, t: CellTypeTag, v: Option<&str>) -> Cell {
        Cell {
            r: CompactCellRef::new(col_ref),
            col: col_num,
            s: None,
            t,
            v: v.map(|s| s.to_string()),
            f: None,
            is: None,
        }
    }

    fn make_row(row_num: u32, cells: Vec<Cell>) -> Row {
        Row {
            r: row_num,
            spans: None,
            s: None,
            custom_format: None,
            ht: None,
            hidden: None,
            custom_height: None,
            outline_level: None,
            cells,
        }
    }

    fn make_worksheet(rows: Vec<Row>) -> WorksheetXml {
        let mut ws = WorksheetXml::default();
        ws.sheet_data = SheetData { rows };
        ws
    }

    fn make_sst(strings: &[&str]) -> SharedStringTable {
        let mut sst = SharedStringTable::new();
        for s in strings {
            sst.add(s);
        }
        sst
    }

    fn read_u32_le(buf: &[u8], offset: usize) -> u32 {
        u32::from_le_bytes(buf[offset..offset + 4].try_into().unwrap())
    }

    fn read_u16_le(buf: &[u8], offset: usize) -> u16 {
        u16::from_le_bytes(buf[offset..offset + 2].try_into().unwrap())
    }

    fn read_f64_le(buf: &[u8], offset: usize) -> f64 {
        f64::from_le_bytes(buf[offset..offset + 8].try_into().unwrap())
    }

    fn read_inline_string(buf: &[u8], offset: usize) -> (String, usize) {
        let len = read_u32_le(buf, offset) as usize;
        let s = std::str::from_utf8(&buf[offset + 4..offset + 4 + len])
            .unwrap()
            .to_string();
        (s, 4 + len)
    }

    /// Parse v2 buffer cell data for a single row at the given absolute offset.
    /// Returns vec of (col, type_tag, value_description).
    fn read_v2_row_cells(buf: &[u8], abs_offset: usize) -> Vec<(u16, u8, String)> {
        let cell_count = read_u16_le(buf, abs_offset) as usize;
        let mut pos = abs_offset + 2;
        let mut result = Vec::with_capacity(cell_count);

        for _ in 0..cell_count {
            let col = read_u16_le(buf, pos);
            let type_tag = buf[pos + 2];
            pos += 3;

            let val = match type_tag {
                TYPE_EMPTY => String::new(),
                TYPE_NUMBER | TYPE_DATE => {
                    let n = read_f64_le(buf, pos);
                    pos += 8;
                    format!("{n}")
                }
                TYPE_BOOL => {
                    let b = buf[pos];
                    pos += 1;
                    format!("{b}")
                }
                TYPE_STRING | TYPE_ERROR | TYPE_FORMULA | TYPE_RICH_STRING => {
                    let (s, consumed) = read_inline_string(buf, pos);
                    pos += consumed;
                    s
                }
                _ => panic!("unknown type tag: {type_tag}"),
            };

            result.push((col, type_tag, val));
        }

        result
    }

    #[test]
    fn test_empty_sheet() {
        let ws = make_worksheet(vec![]);
        let sst = SharedStringTable::new();
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        assert_eq!(buf.len(), HEADER_SIZE);
        assert_eq!(read_u32_le(&buf, 0), MAGIC_V2);
        assert_eq!(read_u16_le(&buf, 4), VERSION_V2);
        assert_eq!(read_u32_le(&buf, 6), 0);
        assert_eq!(read_u16_le(&buf, 10), 0);
        assert_eq!(read_u32_le(&buf, 12), 0);
    }

    #[test]
    fn test_single_number_cell() {
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::None, Some("42.5"))],
        )]);
        let sst = SharedStringTable::new();
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        assert_eq!(read_u32_le(&buf, 0), MAGIC_V2);
        assert_eq!(read_u32_le(&buf, 6), 1); // row_count
        assert_eq!(read_u16_le(&buf, 10), 1); // col_count

        let row_index_start = HEADER_SIZE;
        let row1_num = read_u32_le(&buf, row_index_start);
        let row1_offset = read_u32_le(&buf, row_index_start + 4);
        assert_eq!(row1_num, 1);
        assert_ne!(row1_offset, EMPTY_ROW_OFFSET);

        let cell_data_start = HEADER_SIZE + 8;
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);
        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].0, 0); // col 0 (relative)
        assert_eq!(cells[0].1, TYPE_NUMBER);
        assert_eq!(cells[0].2, "42.5");
    }

    #[test]
    fn test_string_cell_sst() {
        let sst = make_sst(&["Hello", "World"]);
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::SharedString, Some("1"))],
        )]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let cell_data_start = HEADER_SIZE + 8;
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].1, TYPE_STRING);
        assert_eq!(cells[0].2, "World"); // SST index 1 = "World"
    }

    #[test]
    fn test_bool_cell() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![
                make_cell("A1", 1, CellTypeTag::Boolean, Some("1")),
                make_cell("B1", 2, CellTypeTag::Boolean, Some("0")),
            ],
        )]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let cell_data_start = HEADER_SIZE + 8;
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0].1, TYPE_BOOL);
        assert_eq!(cells[0].2, "1");
        assert_eq!(cells[1].1, TYPE_BOOL);
        assert_eq!(cells[1].2, "0");
    }

    #[test]
    fn test_error_cell() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::Error, Some("#DIV/0!"))],
        )]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let cell_data_start = HEADER_SIZE + 8;
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].1, TYPE_ERROR);
        assert_eq!(cells[0].2, "#DIV/0!");
    }

    #[test]
    fn test_formula_cell() {
        let sst = SharedStringTable::new();
        let mut cell = make_cell("A1", 1, CellTypeTag::None, Some("84"));
        cell.f = Some(Box::new(CellFormula {
            t: None,
            reference: None,
            si: None,
            value: Some("A2+B2".to_string()),
        }));
        let ws = make_worksheet(vec![make_row(1, vec![cell])]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let cell_data_start = HEADER_SIZE + 8;
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].1, TYPE_FORMULA);
        assert_eq!(cells[0].2, "A2+B2");
    }

    #[test]
    fn test_inline_string_cell() {
        let sst = SharedStringTable::new();
        let mut cell = make_cell("A1", 1, CellTypeTag::InlineString, None);
        cell.is = Some(Box::new(InlineString {
            t: Some("Inline Text".to_string()),
        }));
        let ws = make_worksheet(vec![make_row(1, vec![cell])]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let cell_data_start = HEADER_SIZE + 8;
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].1, TYPE_STRING);
        assert_eq!(cells[0].2, "Inline Text");
    }

    #[test]
    fn test_date_cell() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::Date, Some("44927.0"))],
        )]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let cell_data_start = HEADER_SIZE + 8;
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].1, TYPE_DATE);
        let val: f64 = cells[0].2.parse().unwrap();
        assert!((val - 44927.0).abs() < f64::EPSILON);
    }

    #[test]
    fn test_mixed_types_row() {
        let sst = make_sst(&["Hello"]);
        let ws = make_worksheet(vec![make_row(
            1,
            vec![
                make_cell("A1", 1, CellTypeTag::None, Some("3.14")),
                make_cell("B1", 2, CellTypeTag::SharedString, Some("0")),
                make_cell("C1", 3, CellTypeTag::Boolean, Some("1")),
            ],
        )]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let cell_data_start = HEADER_SIZE + 8;
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells.len(), 3);
        assert_eq!(cells[0].1, TYPE_NUMBER);
        assert_eq!(cells[1].1, TYPE_STRING);
        assert_eq!(cells[1].2, "Hello");
        assert_eq!(cells[2].1, TYPE_BOOL);
    }

    #[test]
    fn test_sparse_rows_with_gaps() {
        let sst = SharedStringTable::new();
        let rows = vec![
            make_row(1, vec![make_cell("A1", 1, CellTypeTag::None, Some("1"))]),
            make_row(
                100,
                vec![make_cell("T100", 20, CellTypeTag::None, Some("2"))],
            ),
        ];
        let ws = make_worksheet(rows);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let row_count = read_u32_le(&buf, 6) as usize;
        assert_eq!(row_count, 100);
        let col_count = read_u16_le(&buf, 10) as usize;
        assert_eq!(col_count, 20);

        // Verify row 1 has data
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        assert_ne!(row1_offset, EMPTY_ROW_OFFSET);

        // Verify row 2 is empty
        let row2_offset = read_u32_le(&buf, HEADER_SIZE + 8 + 4);
        assert_eq!(row2_offset, EMPTY_ROW_OFFSET);

        // Verify row 100 has data
        let row100_offset = read_u32_le(&buf, HEADER_SIZE + 99 * 8 + 4);
        assert_ne!(row100_offset, EMPTY_ROW_OFFSET);
    }

    #[test]
    fn test_row_index_entries() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![
            make_row(1, vec![make_cell("A1", 1, CellTypeTag::None, Some("1"))]),
            make_row(3, vec![make_cell("A3", 1, CellTypeTag::None, Some("3"))]),
        ]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let row_count = read_u32_le(&buf, 6) as usize;
        assert_eq!(row_count, 3);

        let row1_num = read_u32_le(&buf, HEADER_SIZE);
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        assert_eq!(row1_num, 1);
        assert_ne!(row1_offset, EMPTY_ROW_OFFSET);

        let row2_num = read_u32_le(&buf, HEADER_SIZE + 8);
        let row2_offset = read_u32_le(&buf, HEADER_SIZE + 12);
        assert_eq!(row2_num, 2);
        assert_eq!(row2_offset, EMPTY_ROW_OFFSET);

        let row3_num = read_u32_le(&buf, HEADER_SIZE + 16);
        let row3_offset = read_u32_le(&buf, HEADER_SIZE + 20);
        assert_eq!(row3_num, 3);
        assert_ne!(row3_offset, EMPTY_ROW_OFFSET);
    }

    #[test]
    fn test_header_format() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![
            make_row(
                2,
                vec![
                    make_cell("B2", 2, CellTypeTag::None, Some("10")),
                    make_cell("D2", 4, CellTypeTag::None, Some("20")),
                ],
            ),
            make_row(5, vec![make_cell("C5", 3, CellTypeTag::None, Some("30"))]),
        ]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        assert_eq!(read_u32_le(&buf, 0), MAGIC_V2);
        assert_eq!(read_u16_le(&buf, 4), VERSION_V2);
        let row_count = read_u32_le(&buf, 6);
        assert_eq!(row_count, 4, "rows 2-5 = 4 rows");
        let col_count = read_u16_le(&buf, 10);
        assert_eq!(col_count, 3, "cols B-D = 3 columns");
    }

    #[test]
    fn test_no_string_table_section() {
        let sst = make_sst(&["Alpha", "Beta"]);
        let ws = make_worksheet(vec![make_row(
            1,
            vec![
                make_cell("A1", 1, CellTypeTag::SharedString, Some("0")),
                make_cell("B1", 2, CellTypeTag::SharedString, Some("1")),
            ],
        )]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        // v2 has no string table section: after header + row_index comes cell data directly
        let cell_data_start = HEADER_SIZE + 8; // 1 row * 8 bytes
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0].1, TYPE_STRING);
        assert_eq!(cells[0].2, "Alpha");
        assert_eq!(cells[1].1, TYPE_STRING);
        assert_eq!(cells[1].2, "Beta");
    }

    #[test]
    fn test_large_string_inline() {
        let sst = SharedStringTable::new();
        let long_str = "A".repeat(10000);
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell(
                "A1",
                1,
                CellTypeTag::InlineString,
                Some(&long_str),
            )],
        )]);
        // InlineString uses `is` field; set it via the cell is field
        // Actually InlineString with v.as_deref() fallback will use `v` if `is` is None
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let cell_data_start = HEADER_SIZE + 8;
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].1, TYPE_STRING);
        assert_eq!(cells[0].2, long_str);
    }

    #[test]
    fn test_formula_string_type() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell(
                "A1",
                1,
                CellTypeTag::FormulaString,
                Some("computed"),
            )],
        )]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let cell_data_start = HEADER_SIZE + 8;
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells[0].1, TYPE_STRING);
        assert_eq!(cells[0].2, "computed");
    }

    #[test]
    fn test_rich_string_from_sst() {
        use crate::rich_text::RichTextRun;

        let mut sst = SharedStringTable::new();
        sst.add("plain");
        sst.add_rich_text(&[
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
        ]);

        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::SharedString, Some("1"))],
        )]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let cell_data_start = HEADER_SIZE + 8;
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);

        assert_eq!(cells[0].1, TYPE_RICH_STRING);
        assert_eq!(cells[0].2, "Bold Normal");
    }

    #[test]
    fn test_multiple_rows_multiple_types() {
        let sst = make_sst(&["Hello", "World"]);
        let ws = make_worksheet(vec![
            make_row(
                1,
                vec![
                    make_cell("A1", 1, CellTypeTag::None, Some("42")),
                    make_cell("B1", 2, CellTypeTag::SharedString, Some("0")),
                ],
            ),
            make_row(
                2,
                vec![
                    make_cell("A2", 1, CellTypeTag::Boolean, Some("1")),
                    make_cell("B2", 2, CellTypeTag::Error, Some("#N/A")),
                ],
            ),
        ]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let row_count = read_u32_le(&buf, 6) as usize;
        assert_eq!(row_count, 2);

        let cell_data_start = HEADER_SIZE + row_count * 8;

        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let cells1 = read_v2_row_cells(&buf, cell_data_start + row1_offset as usize);
        assert_eq!(cells1.len(), 2);
        assert_eq!(cells1[0].1, TYPE_NUMBER);
        assert_eq!(cells1[1].1, TYPE_STRING);
        assert_eq!(cells1[1].2, "Hello");

        let row2_offset = read_u32_le(&buf, HEADER_SIZE + 8 + 4);
        let cells2 = read_v2_row_cells(&buf, cell_data_start + row2_offset as usize);
        assert_eq!(cells2.len(), 2);
        assert_eq!(cells2[0].1, TYPE_BOOL);
        assert_eq!(cells2[1].1, TYPE_ERROR);
        assert_eq!(cells2[1].2, "#N/A");
    }

    #[test]
    fn test_cell_without_col_uses_ref_parsing() {
        let sst = SharedStringTable::new();
        let mut cell = make_cell("C5", 0, CellTypeTag::None, Some("42"));
        cell.col = 0;
        let ws = make_worksheet(vec![make_row(5, vec![cell])]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        let row_count = read_u32_le(&buf, 6);
        assert_eq!(row_count, 1);
        let col_count = read_u16_le(&buf, 10);
        assert_eq!(col_count, 1);
    }
}
