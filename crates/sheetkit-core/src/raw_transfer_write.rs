//! Write-direction buffer serialization and deserialization for raw FFI transfer.
//!
//! Provides [`cells_to_raw_buffer`] to encode structured cell data into a
//! compact binary buffer, and [`raw_buffer_to_cells`] to decode a buffer
//! back into cell values. These are used for the JS-to-Rust write path
//! (e.g., `setSheetData`) and for round-trip testing.
//!
//! The binary format matches the specification in `raw_transfer.rs` so that
//! buffers produced by either module can be consumed by the other.

use std::collections::HashMap;

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::rich_text;

const MAGIC: u32 = 0x534B5244;
const VERSION: u16 = 1;
const HEADER_SIZE: usize = 16;
const ROW_INDEX_ENTRY_SIZE: usize = 8;
const CELL_STRIDE: usize = 9;
const SPARSE_ENTRY_SIZE: usize = 11;
const EMPTY_ROW_SENTINEL: u32 = 0xFFFF_FFFF;
const FLAG_SPARSE: u32 = 1;
const DENSITY_THRESHOLD: f64 = 0.3;

const TYPE_EMPTY: u8 = 0x00;
const TYPE_NUMBER: u8 = 0x01;
const TYPE_STRING: u8 = 0x02;
const TYPE_BOOL: u8 = 0x03;
const TYPE_DATE: u8 = 0x04;
const TYPE_ERROR: u8 = 0x05;
const TYPE_FORMULA: u8 = 0x06;
const TYPE_RICH_STRING: u8 = 0x07;

/// A row of cell data: (1-based row number, cells in that row).
type CellRow = (u32, Vec<(u32, CellValue)>);

/// Intermediate encoded cell: (0-based column index, type tag, 8-byte payload).
type EncodedCell = (u16, u8, [u8; 8]);

/// Intermediate encoded row: (1-based row number, encoded cells).
type EncodedRow = (u32, Vec<EncodedCell>);

struct BufferHeader {
    _version: u16,
    row_count: u32,
    col_count: u16,
    flags: u32,
}

fn read_header(buf: &[u8]) -> Result<BufferHeader> {
    if buf.len() < HEADER_SIZE {
        return Err(Error::Internal(format!(
            "buffer too short for header: {} bytes (need {})",
            buf.len(),
            HEADER_SIZE
        )));
    }
    let magic = u32::from_le_bytes(buf[0..4].try_into().unwrap());
    if magic != MAGIC {
        return Err(Error::Internal(format!(
            "invalid buffer magic: expected 0x{MAGIC:08X}, got 0x{magic:08X}"
        )));
    }
    let version = u16::from_le_bytes(buf[4..6].try_into().unwrap());
    let row_count = u32::from_le_bytes(buf[6..10].try_into().unwrap());
    let col_count = u16::from_le_bytes(buf[10..12].try_into().unwrap());
    let flags = u32::from_le_bytes(buf[12..16].try_into().unwrap());
    Ok(BufferHeader {
        _version: version,
        row_count,
        col_count,
        flags,
    })
}

fn read_row_index(buf: &[u8], row_count: u32) -> Result<Vec<(u32, u32)>> {
    let start = HEADER_SIZE;
    let end = start + row_count as usize * ROW_INDEX_ENTRY_SIZE;
    if buf.len() < end {
        return Err(Error::Internal(format!(
            "buffer too short for row index: {} bytes (need {})",
            buf.len(),
            end
        )));
    }
    let mut entries = Vec::with_capacity(row_count as usize);
    for i in 0..row_count as usize {
        let offset = start + i * ROW_INDEX_ENTRY_SIZE;
        let row_num = u32::from_le_bytes(buf[offset..offset + 4].try_into().unwrap());
        let row_off = u32::from_le_bytes(buf[offset + 4..offset + 8].try_into().unwrap());
        entries.push((row_num, row_off));
    }
    Ok(entries)
}

/// Read the string table. Returns (strings, byte position after string table).
fn read_string_table(buf: &[u8], offset: usize) -> Result<(Vec<String>, usize)> {
    if buf.len() < offset + 8 {
        return Err(Error::Internal(
            "buffer too short for string table header".to_string(),
        ));
    }
    let count = u32::from_le_bytes(buf[offset..offset + 4].try_into().unwrap()) as usize;
    let blob_size = u32::from_le_bytes(buf[offset + 4..offset + 8].try_into().unwrap()) as usize;

    let offsets_start = offset + 8;
    let offsets_end = offsets_start + count * 4;
    let blob_start = offsets_end;
    let blob_end = blob_start + blob_size;

    if buf.len() < blob_end {
        return Err(Error::Internal(format!(
            "buffer too short for string table: {} bytes (need {})",
            buf.len(),
            blob_end
        )));
    }

    let mut string_offsets = Vec::with_capacity(count);
    for i in 0..count {
        let pos = offsets_start + i * 4;
        let off = u32::from_le_bytes(buf[pos..pos + 4].try_into().unwrap()) as usize;
        string_offsets.push(off);
    }

    let mut strings = Vec::with_capacity(count);
    for i in 0..count {
        let start = blob_start + string_offsets[i];
        let end = if i + 1 < count {
            blob_start + string_offsets[i + 1]
        } else {
            blob_end
        };
        let s = std::str::from_utf8(&buf[start..end])
            .map_err(|e| Error::Internal(format!("invalid UTF-8 in string table: {e}")))?;
        strings.push(s.to_string());
    }

    Ok((strings, blob_end))
}

fn decode_cell_payload(type_tag: u8, payload: &[u8], strings: &[String]) -> Result<CellValue> {
    match type_tag {
        TYPE_EMPTY => Ok(CellValue::Empty),
        TYPE_NUMBER => {
            let n = f64::from_le_bytes(payload[0..8].try_into().unwrap());
            Ok(CellValue::Number(n))
        }
        TYPE_STRING => {
            let idx = u32::from_le_bytes(payload[0..4].try_into().unwrap()) as usize;
            let s = strings
                .get(idx)
                .cloned()
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            Ok(CellValue::String(s))
        }
        TYPE_BOOL => Ok(CellValue::Bool(payload[0] != 0)),
        TYPE_DATE => {
            let n = f64::from_le_bytes(payload[0..8].try_into().unwrap());
            Ok(CellValue::Date(n))
        }
        TYPE_ERROR => {
            let idx = u32::from_le_bytes(payload[0..4].try_into().unwrap()) as usize;
            let s = strings
                .get(idx)
                .cloned()
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            Ok(CellValue::Error(s))
        }
        TYPE_FORMULA => {
            let idx = u32::from_le_bytes(payload[0..4].try_into().unwrap()) as usize;
            let expr = strings
                .get(idx)
                .cloned()
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            Ok(CellValue::Formula { expr, result: None })
        }
        TYPE_RICH_STRING => {
            let idx = u32::from_le_bytes(payload[0..4].try_into().unwrap()) as usize;
            let s = strings
                .get(idx)
                .cloned()
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            Ok(CellValue::String(s))
        }
        _ => Ok(CellValue::Empty),
    }
}

fn read_dense_cells(
    buf: &[u8],
    cell_data_start: usize,
    row_index: &[(u32, u32)],
    col_count: u16,
    strings: &[String],
) -> Result<Vec<CellRow>> {
    let mut result = Vec::new();
    for &(row_num, offset) in row_index {
        if offset == EMPTY_ROW_SENTINEL {
            continue;
        }
        let row_start = cell_data_start + offset as usize;
        let row_end = row_start + col_count as usize * CELL_STRIDE;
        if buf.len() < row_end {
            return Err(Error::Internal(format!(
                "buffer too short for dense row data at offset {}",
                row_start
            )));
        }
        let mut cells = Vec::new();
        for c in 0..col_count as usize {
            let cell_offset = row_start + c * CELL_STRIDE;
            let type_tag = buf[cell_offset];
            if type_tag == TYPE_EMPTY {
                continue;
            }
            let payload = &buf[cell_offset + 1..cell_offset + 9];
            let value = decode_cell_payload(type_tag, payload, strings)?;
            cells.push((c as u32 + 1, value));
        }
        if !cells.is_empty() {
            result.push((row_num, cells));
        }
    }
    Ok(result)
}

fn read_sparse_cells(
    buf: &[u8],
    cell_data_start: usize,
    row_index: &[(u32, u32)],
    strings: &[String],
) -> Result<Vec<CellRow>> {
    let mut result = Vec::new();
    for &(row_num, offset) in row_index {
        if offset == EMPTY_ROW_SENTINEL {
            continue;
        }
        let pos = cell_data_start + offset as usize;
        if buf.len() < pos + 2 {
            return Err(Error::Internal(
                "buffer too short for sparse row cell count".to_string(),
            ));
        }
        let cell_count = u16::from_le_bytes(buf[pos..pos + 2].try_into().unwrap()) as usize;
        let entries_start = pos + 2;
        let entries_end = entries_start + cell_count * SPARSE_ENTRY_SIZE;
        if buf.len() < entries_end {
            return Err(Error::Internal(format!(
                "buffer too short for sparse row entries at offset {}",
                entries_start
            )));
        }
        let mut cells = Vec::with_capacity(cell_count);
        for i in 0..cell_count {
            let entry_off = entries_start + i * SPARSE_ENTRY_SIZE;
            let col = u16::from_le_bytes(buf[entry_off..entry_off + 2].try_into().unwrap());
            let type_tag = buf[entry_off + 2];
            let payload = &buf[entry_off + 3..entry_off + 11];
            let value = decode_cell_payload(type_tag, payload, strings)?;
            cells.push((col as u32 + 1, value));
        }
        if !cells.is_empty() {
            result.push((row_num, cells));
        }
    }
    Ok(result)
}

/// Decode a raw buffer into cell values for applying to a worksheet.
///
/// Returns rows as `(row_number, cells)` where each cell is
/// `(col_number, CellValue)`. Both row and column numbers are 1-based.
pub fn raw_buffer_to_cells(buf: &[u8]) -> Result<Vec<CellRow>> {
    let header = read_header(buf)?;
    if header.row_count == 0 {
        return Ok(Vec::new());
    }

    let row_index = read_row_index(buf, header.row_count)?;
    let string_table_offset = HEADER_SIZE + header.row_count as usize * ROW_INDEX_ENTRY_SIZE;
    let (strings, cell_data_start) = read_string_table(buf, string_table_offset)?;

    let is_sparse = header.flags & FLAG_SPARSE != 0;
    if is_sparse {
        read_sparse_cells(buf, cell_data_start, &row_index, &strings)
    } else {
        read_dense_cells(buf, cell_data_start, &row_index, header.col_count, &strings)
    }
}

struct StringTable {
    strings: Vec<String>,
    index_map: HashMap<String, u32>,
}

impl StringTable {
    fn new() -> Self {
        Self {
            strings: Vec::new(),
            index_map: HashMap::new(),
        }
    }

    fn intern(&mut self, s: &str) -> u32 {
        if let Some(&idx) = self.index_map.get(s) {
            return idx;
        }
        let idx = self.strings.len() as u32;
        self.strings.push(s.to_string());
        self.index_map.insert(s.to_string(), idx);
        idx
    }
}

fn cell_type_tag(value: &CellValue) -> u8 {
    match value {
        CellValue::Empty => TYPE_EMPTY,
        CellValue::Number(_) => TYPE_NUMBER,
        CellValue::String(_) => TYPE_STRING,
        CellValue::Bool(_) => TYPE_BOOL,
        CellValue::Date(_) => TYPE_DATE,
        CellValue::Error(_) => TYPE_ERROR,
        CellValue::Formula { .. } => TYPE_FORMULA,
        CellValue::RichString(_) => TYPE_RICH_STRING,
    }
}

fn encode_cell_payload(value: &CellValue, st: &mut StringTable) -> [u8; 8] {
    let mut payload = [0u8; 8];
    match value {
        CellValue::Empty => {}
        CellValue::Number(n) => {
            payload[0..8].copy_from_slice(&n.to_le_bytes());
        }
        CellValue::String(s) => {
            let idx = st.intern(s);
            payload[0..4].copy_from_slice(&idx.to_le_bytes());
        }
        CellValue::Bool(b) => {
            payload[0] = u8::from(*b);
        }
        CellValue::Date(n) => {
            payload[0..8].copy_from_slice(&n.to_le_bytes());
        }
        CellValue::Error(s) => {
            let idx = st.intern(s);
            payload[0..4].copy_from_slice(&idx.to_le_bytes());
        }
        CellValue::Formula { expr, .. } => {
            let idx = st.intern(expr);
            payload[0..4].copy_from_slice(&idx.to_le_bytes());
        }
        CellValue::RichString(runs) => {
            let plain = rich_text::rich_text_to_plain(runs);
            let idx = st.intern(&plain);
            payload[0..4].copy_from_slice(&idx.to_le_bytes());
        }
    }
    payload
}

/// Encode cell values into a raw buffer for transfer.
///
/// Takes rows as `(row_number, cells)` where each cell is
/// `(col_number, CellValue)`. Both row and column numbers are 1-based.
/// Returns the encoded binary buffer.
pub fn cells_to_raw_buffer(rows: &[(u32, Vec<(u32, CellValue)>)]) -> Result<Vec<u8>> {
    if rows.is_empty() {
        return write_empty_buffer();
    }

    let mut max_col: u32 = 0;
    let mut total_cells: usize = 0;
    for (_, cells) in rows {
        for &(col, _) in cells {
            if col > max_col {
                max_col = col;
            }
        }
        total_cells += cells.len();
    }

    let row_count = rows.len() as u32;
    let col_count = max_col as u16;

    let grid_size = row_count as usize * col_count as usize;
    let density = if grid_size > 0 {
        total_cells as f64 / grid_size as f64
    } else {
        0.0
    };
    let is_sparse = density < DENSITY_THRESHOLD;

    let mut st = StringTable::new();
    let mut row_payloads: Vec<EncodedRow> = Vec::with_capacity(rows.len());
    for &(row_num, ref cells) in rows {
        let mut encoded_cells = Vec::with_capacity(cells.len());
        for &(col, ref value) in cells {
            let tag = cell_type_tag(value);
            let payload = encode_cell_payload(value, &mut st);
            encoded_cells.push((col as u16 - 1, tag, payload));
        }
        row_payloads.push((row_num, encoded_cells));
    }

    let row_index_size = row_count as usize * ROW_INDEX_ENTRY_SIZE;
    let string_table_size = compute_string_table_size(&st);
    let cell_data_size = if is_sparse {
        compute_sparse_size(&row_payloads)
    } else {
        compute_dense_size(row_count, col_count)
    };

    let total_size = HEADER_SIZE + row_index_size + string_table_size + cell_data_size;
    let mut buf = vec![0u8; total_size];

    write_header(
        &mut buf,
        row_count,
        col_count,
        if is_sparse { FLAG_SPARSE } else { 0 },
    );

    let cell_data_start = HEADER_SIZE + row_index_size + string_table_size;

    if is_sparse {
        write_sparse_data(&mut buf, &row_payloads, cell_data_start);
    } else {
        write_dense_data(&mut buf, &row_payloads, col_count, cell_data_start);
    }

    write_row_index(&mut buf, &row_payloads, is_sparse, col_count);
    write_string_table(&mut buf, HEADER_SIZE + row_index_size, &st);

    Ok(buf)
}

fn write_empty_buffer() -> Result<Vec<u8>> {
    let st_size = 8; // count(4) + blob_size(4), both zero
    let total = HEADER_SIZE + st_size;
    let mut buf = vec![0u8; total];
    write_header(&mut buf, 0, 0, 0);
    // String table: count=0, blob_size=0
    buf[HEADER_SIZE..HEADER_SIZE + 4].copy_from_slice(&0u32.to_le_bytes());
    buf[HEADER_SIZE + 4..HEADER_SIZE + 8].copy_from_slice(&0u32.to_le_bytes());
    Ok(buf)
}

fn write_header(buf: &mut [u8], row_count: u32, col_count: u16, flags: u32) {
    buf[0..4].copy_from_slice(&MAGIC.to_le_bytes());
    buf[4..6].copy_from_slice(&VERSION.to_le_bytes());
    buf[6..10].copy_from_slice(&row_count.to_le_bytes());
    buf[10..12].copy_from_slice(&col_count.to_le_bytes());
    buf[12..16].copy_from_slice(&flags.to_le_bytes());
}

fn compute_string_table_size(st: &StringTable) -> usize {
    let blob_size: usize = st.strings.iter().map(|s| s.len()).sum();
    8 + st.strings.len() * 4 + blob_size // count(4) + blob_size(4) + offsets + blob
}

fn write_string_table(buf: &mut [u8], offset: usize, st: &StringTable) {
    let count = st.strings.len() as u32;
    let blob_size: usize = st.strings.iter().map(|s| s.len()).sum();

    buf[offset..offset + 4].copy_from_slice(&count.to_le_bytes());
    buf[offset + 4..offset + 8].copy_from_slice(&(blob_size as u32).to_le_bytes());

    let offsets_start = offset + 8;
    let blob_start = offsets_start + st.strings.len() * 4;

    let mut blob_offset: u32 = 0;
    for (i, s) in st.strings.iter().enumerate() {
        let pos = offsets_start + i * 4;
        buf[pos..pos + 4].copy_from_slice(&blob_offset.to_le_bytes());
        let dst = blob_start + blob_offset as usize;
        buf[dst..dst + s.len()].copy_from_slice(s.as_bytes());
        blob_offset += s.len() as u32;
    }
}

fn compute_dense_size(row_count: u32, col_count: u16) -> usize {
    row_count as usize * col_count as usize * CELL_STRIDE
}

fn compute_sparse_size(row_payloads: &[EncodedRow]) -> usize {
    let mut size = 0;
    for (_, cells) in row_payloads {
        size += 2 + cells.len() * SPARSE_ENTRY_SIZE; // cell_count(u16) + entries
    }
    size
}

fn write_row_index(buf: &mut [u8], row_payloads: &[EncodedRow], is_sparse: bool, col_count: u16) {
    let base = HEADER_SIZE;
    if is_sparse {
        let mut data_offset: u32 = 0;
        for (i, (row_num, cells)) in row_payloads.iter().enumerate() {
            let pos = base + i * ROW_INDEX_ENTRY_SIZE;
            buf[pos..pos + 4].copy_from_slice(&row_num.to_le_bytes());
            if cells.is_empty() {
                buf[pos + 4..pos + 8].copy_from_slice(&EMPTY_ROW_SENTINEL.to_le_bytes());
            } else {
                buf[pos + 4..pos + 8].copy_from_slice(&data_offset.to_le_bytes());
            }
            let row_size = 2 + cells.len() * SPARSE_ENTRY_SIZE;
            data_offset += row_size as u32;
        }
    } else {
        for (i, (row_num, _)) in row_payloads.iter().enumerate() {
            let pos = base + i * ROW_INDEX_ENTRY_SIZE;
            buf[pos..pos + 4].copy_from_slice(&row_num.to_le_bytes());
            let offset = i as u32 * col_count as u32 * CELL_STRIDE as u32;
            buf[pos + 4..pos + 8].copy_from_slice(&offset.to_le_bytes());
        }
    }
}

fn write_dense_data(
    buf: &mut [u8],
    row_payloads: &[EncodedRow],
    col_count: u16,
    cell_data_start: usize,
) {
    for (i, (_, cells)) in row_payloads.iter().enumerate() {
        let row_start = cell_data_start + i * col_count as usize * CELL_STRIDE;
        for &(col_idx, tag, ref payload) in cells {
            let cell_off = row_start + col_idx as usize * CELL_STRIDE;
            buf[cell_off] = tag;
            buf[cell_off + 1..cell_off + 9].copy_from_slice(payload);
        }
    }
}

fn write_sparse_data(buf: &mut [u8], row_payloads: &[EncodedRow], cell_data_start: usize) {
    let mut offset = cell_data_start;
    for (_, cells) in row_payloads {
        let cell_count = cells.len() as u16;
        buf[offset..offset + 2].copy_from_slice(&cell_count.to_le_bytes());
        offset += 2;
        for &(col_idx, tag, ref payload) in cells {
            buf[offset..offset + 2].copy_from_slice(&col_idx.to_le_bytes());
            buf[offset + 2] = tag;
            buf[offset + 3..offset + 11].copy_from_slice(payload);
            offset += SPARSE_ENTRY_SIZE;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::rich_text::RichTextRun;

    #[test]
    fn test_decode_empty_buffer() {
        let buf = cells_to_raw_buffer(&[]).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert!(result.is_empty());
    }

    #[test]
    fn test_decode_single_number() {
        let rows = vec![(1, vec![(1, CellValue::Number(42.5))])];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].0, 1);
        assert_eq!(result[0].1.len(), 1);
        assert_eq!(result[0].1[0].0, 1);
        assert_eq!(result[0].1[0].1, CellValue::Number(42.5));
    }

    #[test]
    fn test_decode_string_with_table() {
        let rows = vec![(1, vec![(1, CellValue::String("hello world".to_string()))])];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result.len(), 1);
        assert_eq!(
            result[0].1[0].1,
            CellValue::String("hello world".to_string())
        );
    }

    #[test]
    fn test_decode_bool_true_false() {
        let rows = vec![(
            1,
            vec![(1, CellValue::Bool(true)), (2, CellValue::Bool(false))],
        )];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result[0].1[0].1, CellValue::Bool(true));
        assert_eq!(result[0].1[1].1, CellValue::Bool(false));
    }

    #[test]
    fn test_decode_error() {
        let rows = vec![(1, vec![(1, CellValue::Error("#DIV/0!".to_string()))])];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result[0].1[0].1, CellValue::Error("#DIV/0!".to_string()));
    }

    #[test]
    fn test_decode_formula() {
        let rows = vec![(
            1,
            vec![(
                1,
                CellValue::Formula {
                    expr: "SUM(A1:A10)".to_string(),
                    result: None,
                },
            )],
        )];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(
            result[0].1[0].1,
            CellValue::Formula {
                expr: "SUM(A1:A10)".to_string(),
                result: None,
            }
        );
    }

    #[test]
    fn test_decode_date() {
        let serial = 44927.0; // 2023-01-01
        let rows = vec![(1, vec![(1, CellValue::Date(serial))])];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result[0].1[0].1, CellValue::Date(serial));
    }

    #[test]
    fn test_decode_mixed_row() {
        let rows = vec![(
            3,
            vec![
                (1, CellValue::Number(1.0)),
                (2, CellValue::String("text".to_string())),
                (3, CellValue::Bool(true)),
                (4, CellValue::Date(44927.0)),
                (5, CellValue::Error("#N/A".to_string())),
                (
                    6,
                    CellValue::Formula {
                        expr: "A3+B3".to_string(),
                        result: None,
                    },
                ),
            ],
        )];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();

        assert_eq!(result.len(), 1);
        assert_eq!(result[0].0, 3);
        let cells = &result[0].1;
        assert_eq!(cells.len(), 6);
        assert_eq!(cells[0], (1, CellValue::Number(1.0)));
        assert_eq!(cells[1], (2, CellValue::String("text".to_string())));
        assert_eq!(cells[2], (3, CellValue::Bool(true)));
        assert_eq!(cells[3], (4, CellValue::Date(44927.0)));
        assert_eq!(cells[4], (5, CellValue::Error("#N/A".to_string())));
        assert_eq!(
            cells[5],
            (
                6,
                CellValue::Formula {
                    expr: "A3+B3".to_string(),
                    result: None,
                }
            )
        );
    }

    #[test]
    fn test_round_trip_cells_to_buffer() {
        let rows = vec![
            (
                1,
                vec![
                    (1, CellValue::String("Name".to_string())),
                    (2, CellValue::String("Age".to_string())),
                    (3, CellValue::String("Active".to_string())),
                ],
            ),
            (
                2,
                vec![
                    (1, CellValue::String("Alice".to_string())),
                    (2, CellValue::Number(30.0)),
                    (3, CellValue::Bool(true)),
                ],
            ),
            (
                3,
                vec![
                    (1, CellValue::String("Bob".to_string())),
                    (2, CellValue::Number(25.0)),
                    (3, CellValue::Bool(false)),
                ],
            ),
        ];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result, rows);
    }

    #[test]
    fn test_invalid_magic() {
        let mut buf = vec![0u8; 24];
        buf[0..4].copy_from_slice(&0xDEADBEEFu32.to_le_bytes());
        let err = raw_buffer_to_cells(&buf).unwrap_err();
        assert!(err.to_string().contains("invalid buffer magic"));
    }

    #[test]
    fn test_buffer_too_short() {
        let buf = vec![0u8; 4];
        let err = raw_buffer_to_cells(&buf).unwrap_err();
        assert!(err.to_string().contains("buffer too short"));
    }

    #[test]
    fn test_rich_string_degrades_to_string() {
        let runs = vec![
            RichTextRun {
                text: "bold ".to_string(),
                font: None,
                size: None,
                bold: true,
                italic: false,
                color: None,
            },
            RichTextRun {
                text: "text".to_string(),
                font: None,
                size: None,
                bold: false,
                italic: false,
                color: None,
            },
        ];
        let rows = vec![(1, vec![(1, CellValue::RichString(runs))])];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result[0].1[0].1, CellValue::String("bold text".to_string()));
    }

    #[test]
    fn test_multiple_rows_and_columns() {
        let rows = vec![
            (
                1,
                vec![(1, CellValue::Number(1.0)), (5, CellValue::Number(5.0))],
            ),
            (10, vec![(3, CellValue::String("mid".to_string()))]),
            (
                100,
                vec![(1, CellValue::Bool(true)), (5, CellValue::Date(45000.0))],
            ),
        ];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result.len(), 3);
        assert_eq!(result[0].0, 1);
        assert_eq!(result[1].0, 10);
        assert_eq!(result[2].0, 100);
        assert_eq!(result[0].1.len(), 2);
        assert_eq!(result[0].1[0], (1, CellValue::Number(1.0)));
        assert_eq!(result[0].1[1], (5, CellValue::Number(5.0)));
        assert_eq!(result[1].1[0], (3, CellValue::String("mid".to_string())));
        assert_eq!(result[2].1[0], (1, CellValue::Bool(true)));
        assert_eq!(result[2].1[1], (5, CellValue::Date(45000.0)));
    }

    #[test]
    fn test_sparse_format_selected_for_sparse_data() {
        // 10 rows with 1 cell each, but col ranges up to 100 -> density = 10/(10*100) = 1%
        let mut rows = Vec::new();
        for i in 1..=10 {
            rows.push((i, vec![(100, CellValue::Number(i as f64))]));
        }
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let header = read_header(&buf).unwrap();
        assert_ne!(header.flags & FLAG_SPARSE, 0, "sparse flag should be set");

        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result.len(), 10);
        for (i, (row_num, cells)) in result.iter().enumerate() {
            assert_eq!(*row_num, i as u32 + 1);
            assert_eq!(cells[0], (100, CellValue::Number((i + 1) as f64)));
        }
    }

    #[test]
    fn test_dense_format_selected_for_dense_data() {
        let mut rows = Vec::new();
        for r in 1..=5 {
            let cells: Vec<(u32, CellValue)> = (1..=5)
                .map(|c| (c, CellValue::Number((r * 10 + c) as f64)))
                .collect();
            rows.push((r, cells));
        }
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let header = read_header(&buf).unwrap();
        assert_eq!(
            header.flags & FLAG_SPARSE,
            0,
            "sparse flag should not be set"
        );

        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result.len(), 5);
        for r in 0..5 {
            assert_eq!(result[r].0, (r + 1) as u32);
            assert_eq!(result[r].1.len(), 5);
            for c in 0..5 {
                let expected = ((r + 1) * 10 + (c + 1)) as f64;
                assert_eq!(
                    result[r].1[c],
                    ((c + 1) as u32, CellValue::Number(expected))
                );
            }
        }
    }

    #[test]
    fn test_string_deduplication() {
        let rows = vec![(
            1,
            vec![
                (1, CellValue::String("repeated".to_string())),
                (2, CellValue::String("repeated".to_string())),
                (3, CellValue::String("unique".to_string())),
            ],
        )];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let header = read_header(&buf).unwrap();
        let st_offset = HEADER_SIZE + header.row_count as usize * ROW_INDEX_ENTRY_SIZE;
        let count = u32::from_le_bytes(buf[st_offset..st_offset + 4].try_into().unwrap());
        assert_eq!(count, 2, "string table should have 2 unique strings, not 3");

        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result[0].1[0].1, CellValue::String("repeated".to_string()));
        assert_eq!(result[0].1[1].1, CellValue::String("repeated".to_string()));
        assert_eq!(result[0].1[2].1, CellValue::String("unique".to_string()));
    }

    #[test]
    fn test_header_fields() {
        let rows = vec![
            (
                1,
                vec![(1, CellValue::Number(1.0)), (3, CellValue::Number(3.0))],
            ),
            (2, vec![(2, CellValue::Number(2.0))]),
        ];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let header = read_header(&buf).unwrap();
        assert_eq!(header._version, VERSION);
        assert_eq!(header.row_count, 2);
        assert_eq!(header.col_count, 3);
    }

    #[test]
    fn test_formula_result_not_preserved() {
        let rows = vec![(
            1,
            vec![(
                1,
                CellValue::Formula {
                    expr: "1+1".to_string(),
                    result: Some(Box::new(CellValue::Number(2.0))),
                },
            )],
        )];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(
            result[0].1[0].1,
            CellValue::Formula {
                expr: "1+1".to_string(),
                result: None,
            }
        );
    }

    #[test]
    fn test_hand_constructed_dense_buffer() {
        // Manually construct a buffer with 1 row, 2 cols, dense, 1 number + 1 bool
        let row_count: u32 = 1;
        let col_count: u16 = 2;

        let st_size = 8; // count(4) + blob_size(4) + 0 offsets + 0 blob
        let cell_data_size = 2 * CELL_STRIDE; // 2 cols * 9 bytes
        let total = HEADER_SIZE + ROW_INDEX_ENTRY_SIZE + st_size + cell_data_size;

        let mut buf = vec![0u8; total];
        // Header
        buf[0..4].copy_from_slice(&MAGIC.to_le_bytes());
        buf[4..6].copy_from_slice(&1u16.to_le_bytes()); // version
        buf[6..10].copy_from_slice(&row_count.to_le_bytes());
        buf[10..12].copy_from_slice(&col_count.to_le_bytes());
        buf[12..16].copy_from_slice(&0u32.to_le_bytes()); // flags (dense)

        // Row index: row 1 at offset 0
        let ri_start = HEADER_SIZE;
        buf[ri_start..ri_start + 4].copy_from_slice(&1u32.to_le_bytes());
        buf[ri_start + 4..ri_start + 8].copy_from_slice(&0u32.to_le_bytes());

        // String table: count=0, blob_size=0
        let st_start = ri_start + ROW_INDEX_ENTRY_SIZE;
        buf[st_start..st_start + 4].copy_from_slice(&0u32.to_le_bytes());
        buf[st_start + 4..st_start + 8].copy_from_slice(&0u32.to_le_bytes());

        // Cell data
        let cd_start = st_start + st_size;
        // Col 0: Number 99.0
        buf[cd_start] = TYPE_NUMBER;
        buf[cd_start + 1..cd_start + 9].copy_from_slice(&99.0f64.to_le_bytes());
        // Col 1: Bool true
        buf[cd_start + CELL_STRIDE] = TYPE_BOOL;
        buf[cd_start + CELL_STRIDE + 1] = 1;

        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].0, 1);
        assert_eq!(result[0].1[0], (1, CellValue::Number(99.0)));
        assert_eq!(result[0].1[1], (2, CellValue::Bool(true)));
    }

    #[test]
    fn test_hand_constructed_sparse_buffer() {
        // Manually construct a sparse buffer: 1 row, col_count=100, 1 cell at col 50
        let row_count: u32 = 1;
        let col_count: u16 = 100;

        let st_size = 8; // count=0, blob_size=0
        let cell_data_size = 2 + SPARSE_ENTRY_SIZE; // cell_count(2) + 1 entry(11)
        let total = HEADER_SIZE + ROW_INDEX_ENTRY_SIZE + st_size + cell_data_size;

        let mut buf = vec![0u8; total];
        // Header
        buf[0..4].copy_from_slice(&MAGIC.to_le_bytes());
        buf[4..6].copy_from_slice(&1u16.to_le_bytes());
        buf[6..10].copy_from_slice(&row_count.to_le_bytes());
        buf[10..12].copy_from_slice(&col_count.to_le_bytes());
        buf[12..16].copy_from_slice(&FLAG_SPARSE.to_le_bytes());

        // Row index
        let ri_start = HEADER_SIZE;
        buf[ri_start..ri_start + 4].copy_from_slice(&5u32.to_le_bytes()); // row 5
        buf[ri_start + 4..ri_start + 8].copy_from_slice(&0u32.to_le_bytes()); // offset 0

        // String table
        let st_start = ri_start + ROW_INDEX_ENTRY_SIZE;
        buf[st_start..st_start + 4].copy_from_slice(&0u32.to_le_bytes());
        buf[st_start + 4..st_start + 8].copy_from_slice(&0u32.to_le_bytes());

        // Sparse cell data
        let cd_start = st_start + st_size;
        buf[cd_start..cd_start + 2].copy_from_slice(&1u16.to_le_bytes()); // 1 cell
        let entry = cd_start + 2;
        buf[entry..entry + 2].copy_from_slice(&49u16.to_le_bytes()); // col index 49 (0-based)
        buf[entry + 2] = TYPE_NUMBER;
        buf[entry + 3..entry + 11].copy_from_slice(&7.77f64.to_le_bytes());

        let result = raw_buffer_to_cells(&buf).unwrap();
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].0, 5);
        assert_eq!(result[0].1[0], (50, CellValue::Number(7.77))); // 1-based col 50
    }
}
