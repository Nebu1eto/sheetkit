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
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS};

const MAGIC: u32 = 0x534B5244;
const VERSION: u16 = 1;
const MAGIC_V2: u32 = crate::raw_transfer_v2::MAGIC_V2;
const VERSION_V2: u16 = crate::raw_transfer_v2::VERSION_V2;
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
    let magic = read_u32(buf, 0, "header")?;
    if magic != MAGIC {
        return Err(Error::Internal(format!(
            "invalid buffer magic: expected 0x{MAGIC:08X}, got 0x{magic:08X}"
        )));
    }
    let version = read_u16(buf, 4, "header")?;
    if version != VERSION {
        return Err(Error::Internal(format!(
            "unsupported buffer version: {version}"
        )));
    }
    let row_count = read_u32(buf, 6, "header")?;
    let col_count = read_u16(buf, 10, "header")?;
    let flags = read_u32(buf, 12, "header")?;
    if flags & 0x0000_FFFE != 0 {
        return Err(Error::Internal("invalid v1 buffer flags".to_string()));
    }
    Ok(BufferHeader {
        _version: version,
        row_count,
        col_count,
        flags,
    })
}

fn checked_end(start: usize, len: usize, context: &str) -> Result<usize> {
    start
        .checked_add(len)
        .ok_or_else(|| Error::Internal(format!("{context} exceeds addressable buffer size")))
}

fn read_u16(buf: &[u8], offset: usize, context: &str) -> Result<u16> {
    let end = checked_end(offset, 2, context)?;
    let bytes = buf
        .get(offset..end)
        .ok_or_else(|| Error::Internal(format!("buffer too short for {context}")))?;
    Ok(u16::from_le_bytes(
        bytes.try_into().expect("fixed-size slice"),
    ))
}

fn read_u32(buf: &[u8], offset: usize, context: &str) -> Result<u32> {
    let end = checked_end(offset, 4, context)?;
    let bytes = buf
        .get(offset..end)
        .ok_or_else(|| Error::Internal(format!("buffer too short for {context}")))?;
    Ok(u32::from_le_bytes(
        bytes.try_into().expect("fixed-size slice"),
    ))
}

fn read_f64(buf: &[u8], offset: usize, context: &str) -> Result<f64> {
    let end = checked_end(offset, 8, context)?;
    let bytes = buf
        .get(offset..end)
        .ok_or_else(|| Error::Internal(format!("buffer too short for {context}")))?;
    Ok(f64::from_le_bytes(
        bytes.try_into().expect("fixed-size slice"),
    ))
}

fn validate_coordinate(row: u32, col: u32) -> Result<()> {
    if !(1..=MAX_ROWS).contains(&row) {
        return Err(Error::Internal("row number is out of range".to_string()));
    }
    if !(1..=MAX_COLUMNS).contains(&col) {
        return Err(Error::Internal("cell column is out of range".to_string()));
    }
    Ok(())
}

fn read_row_index(buf: &[u8], row_count: u32) -> Result<Vec<(u32, u32)>> {
    let start = HEADER_SIZE;
    let size = (row_count as usize)
        .checked_mul(ROW_INDEX_ENTRY_SIZE)
        .ok_or_else(|| Error::Internal("row index exceeds addressable buffer size".to_string()))?;
    let end = checked_end(start, size, "row index")?;
    if buf.len() < end {
        return Err(Error::Internal(format!(
            "buffer too short for row index: {} bytes (need {})",
            buf.len(),
            end
        )));
    }
    let mut entries = Vec::with_capacity(row_count as usize);
    for i in 0..row_count as usize {
        let offset = checked_end(
            start,
            i.checked_mul(ROW_INDEX_ENTRY_SIZE).ok_or_else(|| {
                Error::Internal("row index exceeds addressable buffer size".to_string())
            })?,
            "row index",
        )?;
        let row_num = read_u32(buf, offset, "row index entry")?;
        if !(1..=MAX_ROWS).contains(&row_num) {
            return Err(Error::Internal("row number is out of range".to_string()));
        }
        let row_off = read_u32(
            buf,
            checked_end(offset, 4, "row index entry")?,
            "row index entry",
        )?;
        entries.push((row_num, row_off));
    }
    Ok(entries)
}

/// Read the string table. Returns (strings, byte position after string table).
fn read_string_table(buf: &[u8], offset: usize) -> Result<(Vec<String>, usize)> {
    let header_end = offset.checked_add(8).ok_or_else(|| {
        Error::Internal("string table header offset overflows buffer bounds".to_string())
    })?;
    if buf.len() < header_end {
        return Err(Error::Internal(
            "buffer too short for string table header".to_string(),
        ));
    }
    let count = read_u32(buf, offset, "string table header")? as usize;
    let blob_size = read_u32(
        buf,
        checked_end(offset, 4, "string table header")?,
        "string table header",
    )? as usize;

    let offsets_start = header_end;
    let offsets_end = offsets_start
        .checked_add(count.checked_mul(4).ok_or_else(|| {
            Error::Internal("string table offsets exceed addressable buffer size".to_string())
        })?)
        .ok_or_else(|| {
            Error::Internal("string table offsets exceed addressable buffer size".to_string())
        })?;
    let blob_start = offsets_end;
    let blob_end = blob_start.checked_add(blob_size).ok_or_else(|| {
        Error::Internal("string table blob exceeds addressable buffer size".to_string())
    })?;

    if buf.len() < blob_end {
        return Err(Error::Internal(format!(
            "buffer too short for string table: {} bytes (need {})",
            buf.len(),
            blob_end
        )));
    }

    let mut string_offsets = Vec::with_capacity(count);
    for i in 0..count {
        let pos = checked_end(
            offsets_start,
            i.checked_mul(4).ok_or_else(|| {
                Error::Internal("string table offsets exceed addressable buffer size".to_string())
            })?,
            "string table offsets",
        )?;
        let off = read_u32(buf, pos, "string table offset")? as usize;
        string_offsets.push(off);
    }

    let mut strings = Vec::with_capacity(count);
    let mut previous_offset = 0;
    for i in 0..count {
        let offset = string_offsets[i];
        if offset > blob_size || offset < previous_offset {
            return Err(Error::Internal(
                "string table offsets must be monotonic and within the blob".to_string(),
            ));
        }
        let next_offset = string_offsets.get(i + 1).copied().unwrap_or(blob_size);
        if next_offset > blob_size || next_offset < offset {
            return Err(Error::Internal(
                "string table offsets must be monotonic and within the blob".to_string(),
            ));
        }
        let start = checked_end(blob_start, offset, "string table blob")?;
        let end = checked_end(blob_start, next_offset, "string table blob")?;
        let s = std::str::from_utf8(&buf[start..end])
            .map_err(|e| Error::Internal(format!("invalid UTF-8 in string table: {e}")))?;
        strings.push(s.to_string());
        previous_offset = offset;
    }

    Ok((strings, blob_end))
}

fn decode_cell_payload(type_tag: u8, payload: &[u8], strings: &[String]) -> Result<CellValue> {
    match type_tag {
        TYPE_EMPTY => Ok(CellValue::Empty),
        TYPE_NUMBER => {
            let n = read_f64(payload, 0, "cell payload")?;
            Ok(CellValue::Number(n))
        }
        TYPE_STRING => {
            let idx = read_u32(payload, 0, "cell payload")? as usize;
            let s = strings
                .get(idx)
                .cloned()
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            Ok(CellValue::String(s))
        }
        TYPE_BOOL => Ok(CellValue::Bool(
            *payload
                .first()
                .ok_or_else(|| Error::Internal("buffer too short for cell payload".to_string()))?
                != 0,
        )),
        TYPE_DATE => {
            let n = read_f64(payload, 0, "cell payload")?;
            Ok(CellValue::Date(n))
        }
        TYPE_ERROR => {
            let idx = read_u32(payload, 0, "cell payload")? as usize;
            let s = strings
                .get(idx)
                .cloned()
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            Ok(CellValue::Error(s))
        }
        TYPE_FORMULA => {
            let idx = read_u32(payload, 0, "cell payload")? as usize;
            let expr = strings
                .get(idx)
                .cloned()
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            Ok(CellValue::Formula { expr, result: None })
        }
        TYPE_RICH_STRING => {
            let idx = read_u32(payload, 0, "cell payload")? as usize;
            let s = strings
                .get(idx)
                .cloned()
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            Ok(CellValue::String(s))
        }
        _ => Err(Error::Internal(format!(
            "unknown cell type tag: {type_tag}"
        ))),
    }
}

fn read_dense_cells(
    buf: &[u8],
    cell_data_start: usize,
    row_index: &[(u32, u32)],
    col_count: u16,
    strings: &[String],
    min_col: u32,
) -> Result<Vec<CellRow>> {
    let row_size = (col_count as usize)
        .checked_mul(CELL_STRIDE)
        .ok_or_else(|| Error::Internal("dense row exceeds addressable buffer size".to_string()))?;
    let mut result = Vec::new();
    for (i, &(row_num, offset)) in row_index.iter().enumerate() {
        if offset == EMPTY_ROW_SENTINEL {
            return Err(Error::Internal(
                "dense row index cannot be empty".to_string(),
            ));
        }
        let expected_offset = i.checked_mul(row_size).ok_or_else(|| {
            Error::Internal("dense row offset exceeds addressable buffer size".to_string())
        })?;
        if offset as usize != expected_offset {
            return Err(Error::Internal(
                "dense row offsets are inconsistent".to_string(),
            ));
        }
        let row_start = checked_end(cell_data_start, expected_offset, "dense row data")?;
        let row_end = checked_end(row_start, row_size, "dense row data")?;
        if buf.len() < row_end {
            return Err(Error::Internal(format!(
                "buffer too short for dense row data at offset {}",
                row_start
            )));
        }
        let mut cells = Vec::new();
        for c in 0..col_count as usize {
            let cell_offset = checked_end(
                row_start,
                c.checked_mul(CELL_STRIDE).ok_or_else(|| {
                    Error::Internal("dense cell offset exceeds addressable buffer size".to_string())
                })?,
                "dense cell data",
            )?;
            let type_tag = buf[cell_offset];
            if type_tag == TYPE_EMPTY {
                continue;
            }
            let payload = &buf[cell_offset + 1..cell_offset + 9];
            let value = decode_cell_payload(type_tag, payload, strings)?;
            let col = min_col
                .checked_add(c as u32)
                .ok_or_else(|| Error::Internal("cell column is out of range".to_string()))?;
            validate_coordinate(row_num, col)?;
            cells.push((col, value));
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
    min_col: u32,
    col_count: u16,
) -> Result<Vec<CellRow>> {
    let mut result = Vec::new();
    let mut expected_offset = 0usize;
    for &(row_num, offset) in row_index {
        if offset == EMPTY_ROW_SENTINEL {
            let empty_pos = checked_end(cell_data_start, expected_offset, "sparse row data")?;
            if read_u16(buf, empty_pos, "sparse empty row cell count")? != 0 {
                return Err(Error::Internal(
                    "sparse empty row must have zero cells".to_string(),
                ));
            }
            expected_offset = checked_end(expected_offset, 2, "sparse row data")?;
            continue;
        }
        if offset as usize != expected_offset {
            return Err(Error::Internal(
                "sparse row offsets are inconsistent".to_string(),
            ));
        }
        let pos = checked_end(cell_data_start, expected_offset, "sparse row data")?;
        let cell_count = read_u16(buf, pos, "sparse row cell count")? as usize;
        let entries_start = checked_end(pos, 2, "sparse row data")?;
        let entries_len = cell_count.checked_mul(SPARSE_ENTRY_SIZE).ok_or_else(|| {
            Error::Internal("sparse row entries exceed addressable buffer size".to_string())
        })?;
        let entries_end = checked_end(entries_start, entries_len, "sparse row entries")?;
        if buf.len() < entries_end {
            return Err(Error::Internal(format!(
                "buffer too short for sparse row entries at offset {}",
                entries_start
            )));
        }
        let mut cells = Vec::with_capacity(cell_count);
        for i in 0..cell_count {
            let entry_off = checked_end(
                entries_start,
                i.checked_mul(SPARSE_ENTRY_SIZE).ok_or_else(|| {
                    Error::Internal("sparse row entries exceed addressable buffer size".to_string())
                })?,
                "sparse row entry",
            )?;
            let col = read_u16(buf, entry_off, "sparse cell column")?;
            if col >= col_count {
                return Err(Error::Internal(
                    "sparse cell column is out of range".to_string(),
                ));
            }
            let type_tag = *buf
                .get(checked_end(entry_off, 2, "sparse row entry")?)
                .ok_or_else(|| {
                    Error::Internal("buffer too short for sparse cell type".to_string())
                })?;
            let payload_start = checked_end(entry_off, 3, "sparse row entry")?;
            let payload_end = checked_end(entry_off, SPARSE_ENTRY_SIZE, "sparse row entry")?;
            let payload = &buf[payload_start..payload_end];
            let value = decode_cell_payload(type_tag, payload, strings)?;
            let col = min_col
                .checked_add(col as u32)
                .ok_or_else(|| Error::Internal("cell column is out of range".to_string()))?;
            validate_coordinate(row_num, col)?;
            cells.push((col, value));
        }
        expected_offset = checked_end(
            expected_offset,
            2usize.checked_add(entries_len).ok_or_else(|| {
                Error::Internal("sparse row entries exceed addressable buffer size".to_string())
            })?,
            "sparse row data",
        )?;
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
    if buf.len() < HEADER_SIZE {
        return Err(Error::Internal(format!(
            "buffer too short for header: {} bytes (need {})",
            buf.len(),
            HEADER_SIZE
        )));
    }
    if read_u32(buf, 0, "header")? == MAGIC_V2 {
        return raw_buffer_v2_to_cells(buf);
    }
    let header = read_header(buf)?;
    if header.row_count == 0
        && header.col_count == 0
        && header.flags == 0
        && buf.len() == HEADER_SIZE
    {
        return Ok(Vec::new());
    }
    let row_index = read_row_index(buf, header.row_count)?;
    let row_index_size = (header.row_count as usize)
        .checked_mul(ROW_INDEX_ENTRY_SIZE)
        .ok_or_else(|| Error::Internal("row index exceeds addressable buffer size".to_string()))?;
    let string_table_offset = checked_end(HEADER_SIZE, row_index_size, "row index")?;
    let (strings, cell_data_start) = read_string_table(buf, string_table_offset)?;
    let min_col = (header.flags >> 16).max(1);
    let is_sparse = header.flags & FLAG_SPARSE != 0;
    let rows = if is_sparse {
        read_sparse_cells(
            buf,
            cell_data_start,
            &row_index,
            &strings,
            min_col,
            header.col_count,
        )?
    } else {
        read_dense_cells(
            buf,
            cell_data_start,
            &row_index,
            header.col_count,
            &strings,
            min_col,
        )?
    };
    let expected_cell_data_len = if is_sparse {
        let mut size = 0usize;
        for &(_, offset) in &row_index {
            let pos = checked_end(cell_data_start, size, "sparse row data")?;
            let cell_count = read_u16(buf, pos, "sparse row cell count")? as usize;
            let row_size = 2usize
                .checked_add(cell_count.checked_mul(SPARSE_ENTRY_SIZE).ok_or_else(|| {
                    Error::Internal("sparse row entries exceed addressable buffer size".to_string())
                })?)
                .ok_or_else(|| {
                    Error::Internal("sparse row entries exceed addressable buffer size".to_string())
                })?;
            if offset != EMPTY_ROW_SENTINEL && offset as usize != size {
                return Err(Error::Internal(
                    "sparse row offsets are inconsistent".to_string(),
                ));
            }
            size = checked_end(size, row_size, "sparse row data")?;
        }
        size
    } else {
        (header.row_count as usize)
            .checked_mul(header.col_count as usize)
            .and_then(|n| n.checked_mul(CELL_STRIDE))
            .ok_or_else(|| {
                Error::Internal("dense cell data exceeds addressable buffer size".to_string())
            })?
    };
    let expected_end = checked_end(cell_data_start, expected_cell_data_len, "cell data")?;
    if expected_end != buf.len() {
        return Err(Error::Internal(
            "buffer has trailing data or inconsistent cell section".to_string(),
        ));
    }
    Ok(rows)
}

fn raw_buffer_v2_to_cells(buf: &[u8]) -> Result<Vec<CellRow>> {
    let version = read_u16(buf, 4, "header")?;
    if version != VERSION_V2 {
        return Err(Error::Internal(format!(
            "unsupported buffer version: {version}"
        )));
    }
    let row_count = read_u32(buf, 6, "header")?;
    let col_count = read_u16(buf, 10, "header")?;
    let flags = read_u32(buf, 12, "header")?;
    if flags & 0x0000_FFFF != 0 {
        return Err(Error::Internal("invalid v2 buffer flags".to_string()));
    }
    if row_count == 0 && col_count == 0 && flags == 0 && buf.len() == HEADER_SIZE {
        return Ok(Vec::new());
    }
    let min_col = (flags >> 16).max(1);
    let row_index = read_row_index(buf, row_count)?;
    let row_index_size = (row_count as usize)
        .checked_mul(ROW_INDEX_ENTRY_SIZE)
        .ok_or_else(|| Error::Internal("row index exceeds addressable buffer size".to_string()))?;
    let cell_data_start = checked_end(HEADER_SIZE, row_index_size, "row index")?;
    let mut expected_offset = 0usize;
    let mut result = Vec::new();

    for &(row_num, offset) in &row_index {
        let pos = checked_end(cell_data_start, expected_offset, "v2 row data")?;
        let cell_count = read_u16(buf, pos, "v2 row cell count")? as usize;
        if offset == EMPTY_ROW_SENTINEL {
            if cell_count != 0 {
                return Err(Error::Internal(
                    "v2 empty row must have zero cells".to_string(),
                ));
            }
            expected_offset = checked_end(expected_offset, 2, "v2 row data")?;
            continue;
        }
        if offset as usize != expected_offset {
            return Err(Error::Internal(
                "v2 row offsets are inconsistent".to_string(),
            ));
        }
        let mut cells = Vec::with_capacity(cell_count);
        let mut cursor = checked_end(pos, 2, "v2 row data")?;
        for _ in 0..cell_count {
            let col = read_u16(buf, cursor, "v2 cell column")?;
            if col >= col_count {
                return Err(Error::Internal(
                    "v2 cell column is out of range".to_string(),
                ));
            }
            let type_tag = *buf
                .get(checked_end(cursor, 2, "v2 cell")?)
                .ok_or_else(|| Error::Internal("buffer too short for v2 cell type".to_string()))?;
            let payload_start = checked_end(cursor, 3, "v2 cell")?;
            let (value, payload_size) = match type_tag {
                TYPE_EMPTY => (CellValue::Empty, 0),
                TYPE_NUMBER => (
                    CellValue::Number(read_f64(buf, payload_start, "v2 number payload")?),
                    8,
                ),
                TYPE_DATE => (
                    CellValue::Date(read_f64(buf, payload_start, "v2 date payload")?),
                    8,
                ),
                TYPE_BOOL => (
                    CellValue::Bool(
                        *buf.get(payload_start).ok_or_else(|| {
                            Error::Internal("buffer too short for v2 bool payload".to_string())
                        })? != 0,
                    ),
                    1,
                ),
                TYPE_STRING | TYPE_ERROR | TYPE_FORMULA | TYPE_RICH_STRING => {
                    let byte_len = read_u32(buf, payload_start, "v2 string length")? as usize;
                    let text_start = checked_end(payload_start, 4, "v2 string payload")?;
                    let text_end = checked_end(text_start, byte_len, "v2 string payload")?;
                    let text =
                        std::str::from_utf8(buf.get(text_start..text_end).ok_or_else(|| {
                            Error::Internal("buffer too short for v2 string payload".to_string())
                        })?)
                        .map_err(|e| {
                            Error::Internal(format!("invalid UTF-8 in v2 string payload: {e}"))
                        })?
                        .to_string();
                    let value = match type_tag {
                        TYPE_STRING | TYPE_RICH_STRING => CellValue::String(text),
                        TYPE_ERROR => CellValue::Error(text),
                        TYPE_FORMULA => CellValue::Formula {
                            expr: text,
                            result: None,
                        },
                        _ => unreachable!(),
                    };
                    (
                        value,
                        4usize.checked_add(byte_len).ok_or_else(|| {
                            Error::Internal(
                                "v2 string payload exceeds addressable buffer size".to_string(),
                            )
                        })?,
                    )
                }
                _ => {
                    return Err(Error::Internal(format!(
                        "unknown cell type tag: {type_tag}"
                    )))
                }
            };
            let cell_size = 3usize.checked_add(payload_size).ok_or_else(|| {
                Error::Internal("v2 cell exceeds addressable buffer size".to_string())
            })?;
            cursor = checked_end(cursor, cell_size, "v2 cell")?;
            if !matches!(value, CellValue::Empty) {
                let col = min_col
                    .checked_add(col as u32)
                    .ok_or_else(|| Error::Internal("cell column is out of range".to_string()))?;
                validate_coordinate(row_num, col)?;
                cells.push((col, value));
            }
        }
        expected_offset = cursor
            .checked_sub(cell_data_start)
            .ok_or_else(|| Error::Internal("v2 cell offset precedes data section".to_string()))?;
        if !cells.is_empty() {
            result.push((row_num, cells));
        }
    }
    let expected_end = checked_end(cell_data_start, expected_offset, "v2 cell data")?;
    if expected_end != buf.len() {
        return Err(Error::Internal(
            "buffer has trailing data or inconsistent cell section".to_string(),
        ));
    }
    Ok(result)
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
#[allow(clippy::needless_range_loop)]
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
    fn test_decode_header_only_empty_buffers() {
        let mut v1 = vec![0u8; HEADER_SIZE];
        v1[0..4].copy_from_slice(&MAGIC.to_le_bytes());
        v1[4..6].copy_from_slice(&VERSION.to_le_bytes());
        assert!(raw_buffer_to_cells(&v1).unwrap().is_empty());

        let mut v2 = vec![0u8; HEADER_SIZE];
        v2[0..4].copy_from_slice(&MAGIC_V2.to_le_bytes());
        v2[4..6].copy_from_slice(&VERSION_V2.to_le_bytes());
        assert!(raw_buffer_to_cells(&v2).unwrap().is_empty());
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
    fn test_rejects_non_monotonic_string_offsets() {
        let rows = vec![(
            (1),
            vec![
                (1, CellValue::String("first".to_string())),
                (2, CellValue::String("second".to_string())),
            ],
        )];
        let mut buf = cells_to_raw_buffer(&rows).unwrap();
        let string_table = HEADER_SIZE + ROW_INDEX_ENTRY_SIZE;
        let offsets = string_table + 8;
        buf[offsets..offsets + 4].copy_from_slice(&4u32.to_le_bytes());
        buf[offsets + 4..offsets + 8].copy_from_slice(&0u32.to_le_bytes());

        assert!(raw_buffer_to_cells(&buf).is_err());
    }

    #[test]
    fn test_rejects_string_offset_past_blob() {
        let rows = vec![(1, vec![(1, CellValue::String("value".to_string()))])];
        let mut buf = cells_to_raw_buffer(&rows).unwrap();
        let string_table = HEADER_SIZE + ROW_INDEX_ENTRY_SIZE;
        let offsets = string_table + 8;
        buf[offsets..offsets + 4].copy_from_slice(&u32::MAX.to_le_bytes());

        assert!(raw_buffer_to_cells(&buf).is_err());
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
    fn rejects_unknown_versions_and_trailing_data() {
        let rows = vec![(1, vec![(1, CellValue::Number(1.0))])];
        let mut unknown_version = cells_to_raw_buffer(&rows).unwrap();
        unknown_version[4..6].copy_from_slice(&99u16.to_le_bytes());
        assert!(raw_buffer_to_cells(&unknown_version)
            .unwrap_err()
            .to_string()
            .contains("unsupported buffer version"));

        let mut trailing = cells_to_raw_buffer(&rows).unwrap();
        trailing.push(0);
        assert!(raw_buffer_to_cells(&trailing)
            .unwrap_err()
            .to_string()
            .contains("trailing data"));
    }

    #[test]
    fn rejects_truncated_and_inconsistent_sections() {
        let rows = vec![(1, vec![(1, CellValue::String("value".to_string()))])];
        let buf = cells_to_raw_buffer(&rows).unwrap();
        for length in 0..buf.len() {
            assert!(
                raw_buffer_to_cells(&buf[..length]).is_err(),
                "length {length}"
            );
        }

        let mut descending_offsets = buf.clone();
        let string_table = HEADER_SIZE + ROW_INDEX_ENTRY_SIZE;
        descending_offsets[string_table..string_table + 4].copy_from_slice(&2u32.to_le_bytes());
        descending_offsets[string_table + 8..string_table + 12]
            .copy_from_slice(&4u32.to_le_bytes());
        descending_offsets[string_table + 12..string_table + 16]
            .copy_from_slice(&3u32.to_le_bytes());
        assert!(raw_buffer_to_cells(&descending_offsets).is_err());

        let mut oversized_count = buf.clone();
        oversized_count[6..10].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(raw_buffer_to_cells(&oversized_count).is_err());
    }

    #[test]
    fn decodes_v1_min_col_and_v2_coordinates() {
        let rows = vec![(1, vec![(1, CellValue::String("v1".to_string()))])];
        let mut v1 = cells_to_raw_buffer(&rows).unwrap();
        let flags =
            (3u32 << 16) | (u32::from_le_bytes(v1[12..16].try_into().unwrap()) & FLAG_SPARSE);
        v1[12..16].copy_from_slice(&flags.to_le_bytes());
        assert_eq!(
            raw_buffer_to_cells(&v1).unwrap(),
            vec![(1, vec![(3, CellValue::String("v1".to_string()))])]
        );

        let mut v2 = Vec::new();
        v2.extend_from_slice(&crate::raw_transfer_v2::MAGIC_V2.to_le_bytes());
        v2.extend_from_slice(&crate::raw_transfer_v2::VERSION_V2.to_le_bytes());
        v2.extend_from_slice(&1u32.to_le_bytes());
        v2.extend_from_slice(&1u16.to_le_bytes());
        v2.extend_from_slice(&(3u32 << 16).to_le_bytes());
        v2.extend_from_slice(&1u32.to_le_bytes());
        v2.extend_from_slice(&0u32.to_le_bytes());
        v2.extend_from_slice(&1u16.to_le_bytes());
        v2.extend_from_slice(&0u16.to_le_bytes());
        v2.push(TYPE_NUMBER);
        v2.extend_from_slice(&42.0f64.to_le_bytes());
        assert_eq!(
            raw_buffer_to_cells(&v2).unwrap(),
            vec![(1, vec![(3, CellValue::Number(42.0))])],
        );
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
