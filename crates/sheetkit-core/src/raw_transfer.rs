//! Read-direction buffer serializer for bulk data transfer.
//!
//! Converts a `WorksheetXml` and `SharedStringTable` into a compact binary
//! buffer that can be transferred to JavaScript as a single `Buffer` object,
//! avoiding per-cell napi object creation overhead.
//!
//! Binary format (little-endian throughout):
//!
//! ```text
//! HEADER (16 bytes)
//!   magic:     u32  = 0x534B5244 ("SKRD")
//!   version:   u16  = 1
//!   row_count: u32  = number of rows
//!   col_count: u16  = number of columns
//!   flags:     u32  = bit 0: 1=sparse, 0=dense
//!
//! ROW INDEX (row_count * 8 bytes)
//!   per row: row_number (u32) + offset (u32) into CELL DATA
//!   offset = 0xFFFFFFFF for empty rows
//!
//! STRING TABLE
//!   count:     u32
//!   blob_size: u32
//!   offsets:   u32[count] (byte offset within blob)
//!   blob:      concatenated UTF-8 strings (blob_size bytes)
//!
//! CELL DATA
//!   Dense:  row_count * col_count * 9 bytes
//!     per cell: type (u8) + payload (8 bytes)
//!   Sparse: variable length
//!     per row: cell_count (u16) + cell_count * 11 bytes
//!       per cell: col (u16) + type (u8) + payload (8 bytes)
//! ```

use std::collections::HashMap;

use crate::error::Result;
use crate::sst::SharedStringTable;
use crate::utils::cell_ref::cell_name_to_coordinates;
use sheetkit_xml::worksheet::{CellTypeTag, WorksheetXml};

pub const MAGIC: u32 = 0x534B5244;
pub const VERSION: u16 = 1;
pub const HEADER_SIZE: usize = 16;
pub const CELL_STRIDE: usize = 9;

pub const TYPE_EMPTY: u8 = 0x00;
pub const TYPE_NUMBER: u8 = 0x01;
pub const TYPE_STRING: u8 = 0x02;
pub const TYPE_BOOL: u8 = 0x03;
pub const TYPE_DATE: u8 = 0x04;
pub const TYPE_ERROR: u8 = 0x05;
pub const TYPE_FORMULA: u8 = 0x06;
pub const TYPE_RICH_STRING: u8 = 0x07;

pub const FLAG_SPARSE: u32 = 0x01;

const SPARSE_DENSITY_THRESHOLD: f64 = 0.30;
const EMPTY_ROW_OFFSET: u32 = 0xFFFF_FFFF;
const SPARSE_CELL_STRIDE: usize = 11;

/// Serialize a worksheet's cell data into a compact binary buffer.
///
/// Reads cell data directly from `WorksheetXml` sheet data, resolving shared
/// string references via `sst`. The resulting buffer uses either dense or
/// sparse layout depending on cell density relative to the bounding rectangle.
pub fn sheet_to_raw_buffer(ws: &WorksheetXml, sst: &SharedStringTable) -> Result<Vec<u8>> {
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

    let total_grid = row_count * col_count;
    let density = total_cells as f64 / total_grid as f64;
    let sparse = density < SPARSE_DENSITY_THRESHOLD;

    let mut string_table = StringTableBuilder::from_sst(sst);

    let cell_entries = collect_cell_entries(ws, sst, min_col, &mut string_table)?;

    let flags: u32 = if sparse { FLAG_SPARSE } else { 0 };

    let row_index_size = row_count * 8;
    let string_section = string_table.encode();
    let cell_data = if sparse {
        encode_sparse_cells(&cell_entries, min_row, max_row)
    } else {
        encode_dense_cells(&cell_entries, min_row, row_count, col_count)
    };
    let row_index = build_row_index(&cell_entries, min_row, max_row, col_count, sparse);

    let total_size = HEADER_SIZE + row_index_size + string_section.len() + cell_data.len();
    let mut buf = Vec::with_capacity(total_size);

    buf.extend_from_slice(&MAGIC.to_le_bytes());
    buf.extend_from_slice(&VERSION.to_le_bytes());
    buf.extend_from_slice(&(row_count as u32).to_le_bytes());
    buf.extend_from_slice(&(col_count as u16).to_le_bytes());
    buf.extend_from_slice(&flags.to_le_bytes());

    buf.extend_from_slice(&row_index);
    buf.extend_from_slice(&string_section);
    buf.extend_from_slice(&cell_data);

    Ok(buf)
}

fn write_empty_buffer() -> Vec<u8> {
    let mut buf = Vec::with_capacity(HEADER_SIZE);
    buf.extend_from_slice(&MAGIC.to_le_bytes());
    buf.extend_from_slice(&VERSION.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf.extend_from_slice(&0u16.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf
}

struct CellEntry {
    col: u32,
    type_tag: u8,
    payload: [u8; 8],
}

struct RowEntries {
    row_num: u32,
    cells: Vec<CellEntry>,
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

fn collect_cell_entries(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    min_col: u32,
    string_table: &mut StringTableBuilder,
) -> Result<Vec<RowEntries>> {
    let mut result = Vec::with_capacity(ws.sheet_data.rows.len());

    for row in &ws.sheet_data.rows {
        if row.cells.is_empty() {
            continue;
        }

        let mut cells = Vec::with_capacity(row.cells.len());
        for cell in &row.cells {
            let col = resolve_col(cell)?;
            let relative_col = col - min_col;
            let (type_tag, payload) = encode_cell_value(cell, sst, string_table)?;
            cells.push(CellEntry {
                col: relative_col,
                type_tag,
                payload,
            });
        }

        result.push(RowEntries {
            row_num: row.r,
            cells,
        });
    }

    Ok(result)
}

fn encode_cell_value(
    cell: &sheetkit_xml::worksheet::Cell,
    sst: &SharedStringTable,
    string_table: &mut StringTableBuilder,
) -> Result<(u8, [u8; 8])> {
    let mut payload = [0u8; 8];

    if cell.f.is_some() {
        let formula_expr = cell
            .f
            .as_ref()
            .and_then(|f| f.value.as_deref())
            .unwrap_or("");
        let idx = string_table.intern(formula_expr);
        payload[..4].copy_from_slice(&(idx as u32).to_le_bytes());
        return Ok((TYPE_FORMULA, payload));
    }

    match cell.t {
        CellTypeTag::SharedString => {
            if let Some(ref v) = cell.v {
                if let Ok(sst_idx) = v.parse::<usize>() {
                    let text = sst.get(sst_idx).unwrap_or("");
                    let idx = string_table.intern(text);
                    payload[..4].copy_from_slice(&(idx as u32).to_le_bytes());
                    if sst.get_rich_text(sst_idx).is_some() {
                        return Ok((TYPE_RICH_STRING, payload));
                    }
                    return Ok((TYPE_STRING, payload));
                }
            }
            Ok((TYPE_EMPTY, payload))
        }
        CellTypeTag::Boolean => {
            if let Some(ref v) = cell.v {
                payload[0] = if v == "1" || v.eq_ignore_ascii_case("true") {
                    1
                } else {
                    0
                };
            }
            Ok((TYPE_BOOL, payload))
        }
        CellTypeTag::Error => {
            let error_text = cell.v.as_deref().unwrap_or("#VALUE!");
            let idx = string_table.intern(error_text);
            payload[..4].copy_from_slice(&(idx as u32).to_le_bytes());
            Ok((TYPE_ERROR, payload))
        }
        CellTypeTag::InlineString => {
            let text = cell
                .is
                .as_ref()
                .and_then(|is| is.t.as_deref())
                .or(cell.v.as_deref())
                .unwrap_or("");
            let idx = string_table.intern(text);
            payload[..4].copy_from_slice(&(idx as u32).to_le_bytes());
            Ok((TYPE_STRING, payload))
        }
        CellTypeTag::Date => {
            if let Some(ref v) = cell.v {
                if let Ok(n) = v.parse::<f64>() {
                    payload.copy_from_slice(&n.to_le_bytes());
                    return Ok((TYPE_DATE, payload));
                }
            }
            Ok((TYPE_EMPTY, payload))
        }
        CellTypeTag::FormulaString => {
            if let Some(ref v) = cell.v {
                let idx = string_table.intern(v);
                payload[..4].copy_from_slice(&(idx as u32).to_le_bytes());
                return Ok((TYPE_STRING, payload));
            }
            Ok((TYPE_EMPTY, payload))
        }
        CellTypeTag::None | CellTypeTag::Number => {
            if let Some(ref v) = cell.v {
                if let Ok(n) = v.parse::<f64>() {
                    payload.copy_from_slice(&n.to_le_bytes());
                    return Ok((TYPE_NUMBER, payload));
                }
            }
            Ok((TYPE_EMPTY, payload))
        }
    }
}

struct StringTableBuilder {
    strings: Vec<String>,
    index_map: HashMap<String, usize>,
}

impl StringTableBuilder {
    fn from_sst(sst: &SharedStringTable) -> Self {
        let count = sst.len();
        let mut strings = Vec::with_capacity(count);
        let mut index_map = HashMap::with_capacity(count);

        for i in 0..count {
            if let Some(s) = sst.get(i) {
                let owned = s.to_string();
                index_map.entry(owned.clone()).or_insert(i);
                strings.push(owned);
            }
        }

        Self { strings, index_map }
    }

    fn intern(&mut self, s: &str) -> usize {
        if let Some(&idx) = self.index_map.get(s) {
            return idx;
        }
        let idx = self.strings.len();
        self.strings.push(s.to_string());
        self.index_map.insert(s.to_string(), idx);
        idx
    }

    /// Encode the string table section: count(u32) + blob_size(u32) + offsets(u32[count]) + blob.
    fn encode(&self) -> Vec<u8> {
        let count = self.strings.len() as u32;
        if count == 0 {
            let mut buf = Vec::with_capacity(8);
            buf.extend_from_slice(&0u32.to_le_bytes());
            buf.extend_from_slice(&0u32.to_le_bytes());
            return buf;
        }

        let mut blob = Vec::new();
        let mut offsets = Vec::with_capacity(self.strings.len());
        for s in &self.strings {
            offsets.push(blob.len() as u32);
            blob.extend_from_slice(s.as_bytes());
        }
        let blob_size = blob.len() as u32;

        let total = 4 + 4 + self.strings.len() * 4 + blob.len();
        let mut buf = Vec::with_capacity(total);
        buf.extend_from_slice(&count.to_le_bytes());
        buf.extend_from_slice(&blob_size.to_le_bytes());
        for off in &offsets {
            buf.extend_from_slice(&off.to_le_bytes());
        }
        buf.extend_from_slice(&blob);
        buf
    }
}

fn encode_dense_cells(
    row_entries: &[RowEntries],
    min_row: u32,
    row_count: usize,
    col_count: usize,
) -> Vec<u8> {
    let total = row_count * col_count * CELL_STRIDE;
    let mut buf = vec![0u8; total];

    for re in row_entries {
        let row_offset = (re.row_num - min_row) as usize;
        for ce in &re.cells {
            let cell_idx = row_offset * col_count + ce.col as usize;
            let pos = cell_idx * CELL_STRIDE;
            buf[pos] = ce.type_tag;
            buf[pos + 1..pos + 9].copy_from_slice(&ce.payload);
        }
    }

    buf
}

fn encode_sparse_cells(row_entries: &[RowEntries], min_row: u32, max_row: u32) -> Vec<u8> {
    let total_rows = (max_row - min_row + 1) as usize;
    let mut entries_by_row: Vec<Option<&RowEntries>> = vec![None; total_rows];
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
                for ce in &re.cells {
                    buf.extend_from_slice(&(ce.col as u16).to_le_bytes());
                    buf.push(ce.type_tag);
                    buf.extend_from_slice(&ce.payload);
                }
            }
            None => {
                buf.extend_from_slice(&0u16.to_le_bytes());
            }
        }
    }

    buf
}

fn build_row_index(
    row_entries: &[RowEntries],
    min_row: u32,
    max_row: u32,
    col_count: usize,
    sparse: bool,
) -> Vec<u8> {
    let total_rows = (max_row - min_row + 1) as usize;
    let mut index = Vec::with_capacity(total_rows * 8);

    let mut entries_map: HashMap<u32, &RowEntries> = HashMap::new();
    for re in row_entries {
        entries_map.insert(re.row_num, re);
    }

    if sparse {
        let mut sparse_offset = 0u32;
        for row_num in min_row..=max_row {
            index.extend_from_slice(&row_num.to_le_bytes());
            if let Some(re) = entries_map.get(&row_num) {
                if re.cells.is_empty() {
                    index.extend_from_slice(&EMPTY_ROW_OFFSET.to_le_bytes());
                    sparse_offset += 2;
                } else {
                    index.extend_from_slice(&sparse_offset.to_le_bytes());
                    sparse_offset += 2 + (re.cells.len() as u32) * SPARSE_CELL_STRIDE as u32;
                }
            } else {
                index.extend_from_slice(&EMPTY_ROW_OFFSET.to_le_bytes());
                sparse_offset += 2;
            }
        }
    } else {
        for row_num in min_row..=max_row {
            index.extend_from_slice(&row_num.to_le_bytes());
            if entries_map.contains_key(&row_num) {
                let offset = ((row_num - min_row) as usize * col_count * CELL_STRIDE) as u32;
                index.extend_from_slice(&offset.to_le_bytes());
            } else {
                index.extend_from_slice(&EMPTY_ROW_OFFSET.to_le_bytes());
            }
        }
    }

    index
}

#[cfg(test)]
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

    /// Parse a buffer and return (row_index_end, string_section_end, cell_data_start, flags).
    fn parse_sections(buf: &[u8]) -> (usize, usize, usize, u32) {
        let row_count = read_u32_le(buf, 6) as usize;
        let flags = read_u32_le(buf, 12);
        let row_index_end = HEADER_SIZE + row_count * 8;
        let string_count = read_u32_le(buf, row_index_end) as usize;
        let blob_size = read_u32_le(buf, row_index_end + 4) as usize;
        let string_section_end = row_index_end + 8 + string_count * 4 + blob_size;
        (row_index_end, string_section_end, string_section_end, flags)
    }

    /// Read a string from the string table by index.
    fn read_string(buf: &[u8], string_section_start: usize, idx: usize) -> String {
        let count = read_u32_le(buf, string_section_start) as usize;
        let blob_size = read_u32_le(buf, string_section_start + 4) as usize;
        assert!(
            idx < count,
            "string index {idx} out of range (count={count})"
        );
        let offsets_start = string_section_start + 8;
        let blob_start = offsets_start + count * 4;

        let start = read_u32_le(buf, offsets_start + idx * 4) as usize;
        let end = if idx + 1 < count {
            read_u32_le(buf, offsets_start + (idx + 1) * 4) as usize
        } else {
            blob_size
        };
        String::from_utf8(buf[blob_start + start..blob_start + end].to_vec()).unwrap()
    }

    /// Read the cell type tag and payload from cell data at a given position.
    fn read_cell_at(
        buf: &[u8],
        cell_data_start: usize,
        is_sparse: bool,
        cell_index: usize,
    ) -> (u8, &[u8]) {
        if is_sparse {
            panic!("use read_sparse_row for sparse format");
        }
        let pos = cell_data_start + cell_index * CELL_STRIDE;
        (buf[pos], &buf[pos + 1..pos + 9])
    }

    /// Read sparse row: returns vec of (col, type_tag, payload_slice).
    fn read_sparse_row<'a>(buf: &'a [u8], row_offset: usize) -> Vec<(u16, u8, &'a [u8])> {
        let cell_count = read_u16_le(buf, row_offset) as usize;
        let mut result = Vec::with_capacity(cell_count);
        let mut pos = row_offset + 2;
        for _ in 0..cell_count {
            let col = read_u16_le(buf, pos);
            let type_tag = buf[pos + 2];
            let payload = &buf[pos + 3..pos + 11];
            result.push((col, type_tag, payload));
            pos += SPARSE_CELL_STRIDE;
        }
        result
    }

    #[test]
    fn test_empty_sheet() {
        let ws = make_worksheet(vec![]);
        let sst = SharedStringTable::new();
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        assert_eq!(buf.len(), HEADER_SIZE);
        assert_eq!(read_u32_le(&buf, 0), MAGIC);
        assert_eq!(read_u16_le(&buf, 4), VERSION);
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
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        assert_eq!(read_u32_le(&buf, 0), MAGIC);
        let row_count = read_u32_le(&buf, 6);
        let col_count = read_u16_le(&buf, 10) as usize;
        assert_eq!(row_count, 1);
        assert_eq!(col_count, 1);

        let (st_start, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            assert_eq!(cells.len(), 1);
            assert_eq!(cells[0].1, TYPE_NUMBER);
            let val = f64::from_le_bytes(cells[0].2.try_into().unwrap());
            assert!((val - 42.5).abs() < f64::EPSILON);
        } else {
            let (tag, payload) = read_cell_at(&buf, cd_start, false, 0);
            assert_eq!(tag, TYPE_NUMBER);
            let val = f64::from_le_bytes(payload.try_into().unwrap());
            assert!((val - 42.5).abs() < f64::EPSILON);
        }

        let _ = st_start;
    }

    #[test]
    fn test_string_cell_sst() {
        let sst = make_sst(&["Hello", "World"]);
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::SharedString, Some("1"))],
        )]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let (st_start, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        let string_count = read_u32_le(&buf, st_start);
        assert!(string_count >= 2);

        let s0 = read_string(&buf, st_start, 0);
        assert_eq!(s0, "Hello");
        let s1 = read_string(&buf, st_start, 1);
        assert_eq!(s1, "World");

        let str_idx = if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            assert_eq!(cells[0].1, TYPE_STRING);
            u32::from_le_bytes(cells[0].2[..4].try_into().unwrap()) as usize
        } else {
            let (tag, payload) = read_cell_at(&buf, cd_start, false, 0);
            assert_eq!(tag, TYPE_STRING);
            u32::from_le_bytes(payload[..4].try_into().unwrap()) as usize
        };
        assert_eq!(str_idx, 1, "should reference SST index 1 = 'World'");
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
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let col_count = read_u16_le(&buf, 10) as usize;
        assert_eq!(col_count, 2);

        let (_, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            assert_eq!(cells.len(), 2);
            assert_eq!(cells[0].1, TYPE_BOOL);
            assert_eq!(cells[0].2[0], 1);
            assert_eq!(cells[1].1, TYPE_BOOL);
            assert_eq!(cells[1].2[0], 0);
        } else {
            let (tag0, payload0) = read_cell_at(&buf, cd_start, false, 0);
            assert_eq!(tag0, TYPE_BOOL);
            assert_eq!(payload0[0], 1);
            let (tag1, payload1) = read_cell_at(&buf, cd_start, false, 1);
            assert_eq!(tag1, TYPE_BOOL);
            assert_eq!(payload1[0], 0);
        }
    }

    #[test]
    fn test_error_cell() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::Error, Some("#DIV/0!"))],
        )]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let (st_start, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        let string_count = read_u32_le(&buf, st_start);
        assert!(string_count >= 1);

        let type_tag = if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            cells[0].1
        } else {
            let (tag, _) = read_cell_at(&buf, cd_start, false, 0);
            tag
        };
        assert_eq!(type_tag, TYPE_ERROR);

        let error_str = read_string(&buf, st_start, 0);
        assert_eq!(error_str, "#DIV/0!");
    }

    #[test]
    fn test_formula_cell() {
        let sst = SharedStringTable::new();
        let mut cell = make_cell("A1", 1, CellTypeTag::None, Some("84"));
        cell.f = Some(CellFormula {
            t: None,
            reference: None,
            si: None,
            value: Some("A2+B2".to_string()),
        });
        let ws = make_worksheet(vec![make_row(1, vec![cell])]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let (st_start, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        let (type_tag, str_idx) = if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            let idx = u32::from_le_bytes(cells[0].2[..4].try_into().unwrap()) as usize;
            (cells[0].1, idx)
        } else {
            let (tag, payload) = read_cell_at(&buf, cd_start, false, 0);
            let idx = u32::from_le_bytes(payload[..4].try_into().unwrap()) as usize;
            (tag, idx)
        };

        assert_eq!(type_tag, TYPE_FORMULA);
        let formula = read_string(&buf, st_start, str_idx);
        assert_eq!(formula, "A2+B2");
    }

    #[test]
    fn test_inline_string_cell() {
        let sst = SharedStringTable::new();
        let mut cell = make_cell("A1", 1, CellTypeTag::InlineString, None);
        cell.is = Some(InlineString {
            t: Some("Inline Text".to_string()),
        });
        let ws = make_worksheet(vec![make_row(1, vec![cell])]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let (st_start, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        let (type_tag, str_idx) = if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            let idx = u32::from_le_bytes(cells[0].2[..4].try_into().unwrap()) as usize;
            (cells[0].1, idx)
        } else {
            let (tag, payload) = read_cell_at(&buf, cd_start, false, 0);
            let idx = u32::from_le_bytes(payload[..4].try_into().unwrap()) as usize;
            (tag, idx)
        };

        assert_eq!(type_tag, TYPE_STRING);
        let text = read_string(&buf, st_start, str_idx);
        assert_eq!(text, "Inline Text");
    }

    #[test]
    fn test_date_cell() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::Date, Some("44927.0"))],
        )]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let (_, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            assert_eq!(cells[0].1, TYPE_DATE);
            let val = f64::from_le_bytes(cells[0].2.try_into().unwrap());
            assert!((val - 44927.0).abs() < f64::EPSILON);
        } else {
            let (tag, payload) = read_cell_at(&buf, cd_start, false, 0);
            assert_eq!(tag, TYPE_DATE);
            let val = f64::from_le_bytes(payload.try_into().unwrap());
            assert!((val - 44927.0).abs() < f64::EPSILON);
        }
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
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let col_count = read_u16_le(&buf, 10) as usize;
        assert_eq!(col_count, 3);

        let (_, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            assert_eq!(cells.len(), 3);
            assert_eq!(cells[0].1, TYPE_NUMBER);
            assert_eq!(cells[1].1, TYPE_STRING);
            assert_eq!(cells[2].1, TYPE_BOOL);
        } else {
            let (t0, p0) = read_cell_at(&buf, cd_start, false, 0);
            assert_eq!(t0, TYPE_NUMBER);
            let val = f64::from_le_bytes(p0.try_into().unwrap());
            assert!((val - 3.14).abs() < f64::EPSILON);

            let (t1, _) = read_cell_at(&buf, cd_start, false, 1);
            assert_eq!(t1, TYPE_STRING);

            let (t2, p2) = read_cell_at(&buf, cd_start, false, 2);
            assert_eq!(t2, TYPE_BOOL);
            assert_eq!(p2[0], 1);
        }
    }

    #[test]
    fn test_dense_format() {
        let sst = SharedStringTable::new();
        let mut rows = Vec::new();
        for r in 1..=5u32 {
            let mut cells = Vec::new();
            for c in 1..=5u32 {
                let col_letter = match c {
                    1 => "A",
                    2 => "B",
                    3 => "C",
                    4 => "D",
                    5 => "E",
                    _ => unreachable!(),
                };
                let cell_ref = format!("{col_letter}{r}");
                cells.push(make_cell(
                    &cell_ref,
                    c,
                    CellTypeTag::None,
                    Some(&format!("{}", r * 10 + c)),
                ));
            }
            rows.push(make_row(r, cells));
        }
        let ws = make_worksheet(rows);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let flags = read_u32_le(&buf, 12);
        assert_eq!(
            flags & FLAG_SPARSE,
            0,
            "5x5 fully populated should be dense"
        );

        let row_count = read_u32_le(&buf, 6) as usize;
        let col_count = read_u16_le(&buf, 10) as usize;
        assert_eq!(row_count, 5);
        assert_eq!(col_count, 5);

        let (_, _, cd_start, _) = parse_sections(&buf);
        let cell_data_size = row_count * col_count * CELL_STRIDE;
        assert_eq!(
            buf.len() - cd_start,
            cell_data_size,
            "dense cell data should be exactly row_count * col_count * CELL_STRIDE"
        );

        for r in 0..5usize {
            for c in 0..5usize {
                let idx = r * col_count + c;
                let (tag, payload) = read_cell_at(&buf, cd_start, false, idx);
                assert_eq!(tag, TYPE_NUMBER);
                let val = f64::from_le_bytes(payload.try_into().unwrap());
                let expected = ((r + 1) * 10 + (c + 1)) as f64;
                assert!(
                    (val - expected).abs() < f64::EPSILON,
                    "cell ({r},{c}) expected {expected}, got {val}"
                );
            }
        }
    }

    #[test]
    fn test_sparse_format() {
        let sst = SharedStringTable::new();
        let rows = vec![
            make_row(1, vec![make_cell("A1", 1, CellTypeTag::None, Some("1"))]),
            make_row(
                100,
                vec![make_cell("T100", 20, CellTypeTag::None, Some("2"))],
            ),
        ];
        let ws = make_worksheet(rows);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let flags = read_u32_le(&buf, 12);
        assert_ne!(
            flags & FLAG_SPARSE,
            0,
            "2 cells in 100x20 grid should be sparse"
        );

        let row_count = read_u32_le(&buf, 6) as usize;
        assert_eq!(row_count, 100);
        let col_count = read_u16_le(&buf, 10) as usize;
        assert_eq!(col_count, 20);
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
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        assert_eq!(read_u32_le(&buf, 0), MAGIC);
        assert_eq!(read_u16_le(&buf, 4), VERSION);
        let row_count = read_u32_le(&buf, 6);
        assert_eq!(row_count, 4, "rows 2-5 = 4 rows");
        let col_count = read_u16_le(&buf, 10);
        assert_eq!(col_count, 3, "cols B-D = 3 columns");
    }

    #[test]
    fn test_string_table_format() {
        let sst = make_sst(&["Alpha", "Beta", "Gamma"]);
        let ws = make_worksheet(vec![make_row(
            1,
            vec![
                make_cell("A1", 1, CellTypeTag::SharedString, Some("0")),
                make_cell("B1", 2, CellTypeTag::SharedString, Some("2")),
            ],
        )]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let (st_start, _, _, _) = parse_sections(&buf);
        let string_count = read_u32_le(&buf, st_start) as usize;
        assert_eq!(string_count, 3, "all SST strings should be in table");

        let s0 = read_string(&buf, st_start, 0);
        let s1 = read_string(&buf, st_start, 1);
        let s2 = read_string(&buf, st_start, 2);
        assert_eq!(s0, "Alpha");
        assert_eq!(s1, "Beta");
        assert_eq!(s2, "Gamma");
    }

    #[test]
    fn test_large_sheet_dimensions() {
        let sst = SharedStringTable::new();
        let mut rows = Vec::new();
        for r in [1u32, 500, 1000] {
            let mut cells = Vec::new();
            for c in [1u32, 10, 50] {
                let col_name = crate::utils::cell_ref::column_number_to_name(c).unwrap();
                let cell_ref = format!("{col_name}{r}");
                cells.push(make_cell(
                    &cell_ref,
                    c,
                    CellTypeTag::None,
                    Some(&format!("{}", r + c)),
                ));
            }
            rows.push(make_row(r, cells));
        }
        let ws = make_worksheet(rows);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let row_count = read_u32_le(&buf, 6);
        assert_eq!(row_count, 1000, "rows 1-1000 = 1000 rows");
        let col_count = read_u16_le(&buf, 10);
        assert_eq!(col_count, 50, "cols A-AX = 50 columns");

        let flags = read_u32_le(&buf, 12);
        assert_ne!(
            flags & FLAG_SPARSE,
            0,
            "9 cells in 1000x50 should be sparse"
        );
    }

    #[test]
    fn test_row_index_entries() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![
            make_row(1, vec![make_cell("A1", 1, CellTypeTag::None, Some("1"))]),
            make_row(3, vec![make_cell("A3", 1, CellTypeTag::None, Some("3"))]),
        ]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let row_count = read_u32_le(&buf, 6) as usize;
        assert_eq!(row_count, 3);

        let row_index_start = HEADER_SIZE;
        let row1_num = read_u32_le(&buf, row_index_start);
        let row1_offset = read_u32_le(&buf, row_index_start + 4);
        assert_eq!(row1_num, 1);
        assert_ne!(row1_offset, EMPTY_ROW_OFFSET);

        let row2_num = read_u32_le(&buf, row_index_start + 8);
        let row2_offset = read_u32_le(&buf, row_index_start + 12);
        assert_eq!(row2_num, 2);
        assert_eq!(row2_offset, EMPTY_ROW_OFFSET, "row 2 has no data");

        let row3_num = read_u32_le(&buf, row_index_start + 16);
        let row3_offset = read_u32_le(&buf, row_index_start + 20);
        assert_eq!(row3_num, 3);
        assert_ne!(row3_offset, EMPTY_ROW_OFFSET);
    }

    #[test]
    fn test_string_deduplication() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![
                make_cell("A1", 1, CellTypeTag::Error, Some("#N/A")),
                make_cell("B1", 2, CellTypeTag::Error, Some("#N/A")),
            ],
        )]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let (st_start, _, _, _) = parse_sections(&buf);
        let string_count = read_u32_le(&buf, st_start) as usize;
        assert_eq!(
            string_count, 1,
            "duplicate error strings should be deduplicated"
        );
        let s = read_string(&buf, st_start, 0);
        assert_eq!(s, "#N/A");
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
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let (_, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            assert_eq!(cells[0].1, TYPE_STRING);
        } else {
            let (tag, _) = read_cell_at(&buf, cd_start, false, 0);
            assert_eq!(tag, TYPE_STRING);
        }
    }

    #[test]
    fn test_number_with_explicit_type() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::Number, Some("99.9"))],
        )]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let (_, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            assert_eq!(cells[0].1, TYPE_NUMBER);
            let val = f64::from_le_bytes(cells[0].2.try_into().unwrap());
            assert!((val - 99.9).abs() < f64::EPSILON);
        } else {
            let (tag, payload) = read_cell_at(&buf, cd_start, false, 0);
            assert_eq!(tag, TYPE_NUMBER);
            let val = f64::from_le_bytes(payload.try_into().unwrap());
            assert!((val - 99.9).abs() < f64::EPSILON);
        }
    }

    #[test]
    fn test_cell_without_col_uses_ref_parsing() {
        let sst = SharedStringTable::new();
        let mut cell = make_cell("C5", 0, CellTypeTag::None, Some("42"));
        cell.col = 0;
        let ws = make_worksheet(vec![make_row(5, vec![cell])]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let row_count = read_u32_le(&buf, 6);
        assert_eq!(row_count, 1);
        let col_count = read_u16_le(&buf, 10);
        assert_eq!(col_count, 1);
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
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();

        let (st_start, _, cd_start, flags) = parse_sections(&buf);
        let is_sparse = flags & FLAG_SPARSE != 0;

        let (type_tag, str_idx) = if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            let idx = u32::from_le_bytes(cells[0].2[..4].try_into().unwrap()) as usize;
            (cells[0].1, idx)
        } else {
            let (tag, payload) = read_cell_at(&buf, cd_start, false, 0);
            let idx = u32::from_le_bytes(payload[..4].try_into().unwrap()) as usize;
            (tag, idx)
        };

        assert_eq!(type_tag, TYPE_RICH_STRING);
        let text = read_string(&buf, st_start, str_idx);
        assert_eq!(text, "Bold Normal");
    }
}
