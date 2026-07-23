//! Write-direction buffer serialization and deserialization for raw FFI transfer.
//!
//! Provides [`cells_to_raw_buffer`] to encode structured cell data into a
//! compact binary buffer, and [`raw_buffer_to_cells`] to decode a buffer
//! back into cell values. These are used for the JS-to-Rust write path
//! (e.g., `setSheetData`) and for round-trip testing.
//!
//! The binary format matches the specification in `raw_transfer.rs` so that
//! buffers produced by either module can be consumed by the other.

use std::collections::hash_map::DefaultHasher;
use std::collections::HashMap;
use std::hash::Hasher;

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::formula::FormulaBudget;
use crate::resource::ResourceDimension;
use crate::traversal::{check_limit, checked_output_add, checked_output_mul, TraversalLimits};
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

/// Opaque raw-transfer values admitted under one decode/mutation budget.
pub struct ValidatedCellTransfer {
    pub(crate) entries: Vec<(String, CellValue)>,
}

/// Validate a workbook target before decoding a raw transfer.
pub fn validate_cell_transfer_target(
    workbook: &crate::workbook::Workbook,
    sheet: &str,
) -> Result<()> {
    let sheet_idx = workbook.sheet_index(sheet)?;
    workbook.reject_streamed_mutation(sheet_idx)
}

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

fn validate_decode_layout(
    row_count: u32,
    col_count: u16,
    flags: u32,
    is_sparse: bool,
    limits: TraversalLimits,
) -> Result<u32> {
    check_limit(
        ResourceDimension::RowCount,
        limits.max_rows(),
        u64::from(row_count),
    )?;
    check_limit(
        ResourceDimension::ColumnCount,
        limits.max_columns(),
        u64::from(col_count),
    )?;
    if (row_count == 0) != (col_count == 0) {
        return Err(Error::Internal(
            "row and column counts must both be zero or nonzero".to_string(),
        ));
    }
    if row_count == 0 {
        if flags != 0 {
            return Err(Error::Internal(
                "empty raw transfer flags must be zero".to_string(),
            ));
        }
        return Ok(1);
    }
    let min_col = (flags >> 16).max(1);
    let max_col = min_col
        .checked_add(u32::from(col_count) - 1)
        .ok_or_else(|| Error::Internal("cell column span overflows".to_string()))?;
    if max_col > MAX_COLUMNS {
        return Err(Error::Internal(
            "cell column span is out of range".to_string(),
        ));
    }
    if !is_sparse {
        let cell_count = u64::from(row_count)
            .checked_mul(u64::from(col_count))
            .ok_or_else(|| Error::Internal("dense cell count overflow".to_string()))?;
        check_limit(ResourceDimension::CellCount, limits.max_cells(), cell_count)?;
    }
    Ok(min_col)
}

fn read_row_index(
    buf: &[u8],
    row_count: u32,
    budget: &mut DecodeBudget,
) -> Result<Vec<(u32, u32)>> {
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
    let mut entries = Vec::new();
    budget.charge_array::<(u32, u32)>(row_count as usize)?;
    entries
        .try_reserve_exact(row_count as usize)
        .map_err(|_| Error::Internal("row index allocation failed".to_string()))?;
    let mut previous_row = 0u32;
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
        if row_num <= previous_row {
            return Err(Error::Internal(
                "row numbers must be strictly increasing".to_string(),
            ));
        }
        previous_row = row_num;
        let row_off = read_u32(
            buf,
            checked_end(offset, 4, "row index entry")?,
            "row index entry",
        )?;
        entries.push((row_num, row_off));
    }
    Ok(entries)
}

fn validate_string_table_layout(buf: &[u8], offset: usize) -> Result<usize> {
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
    if count == 0 && blob_size != 0 {
        return Err(Error::Internal(
            "empty string table must have an empty blob".to_string(),
        ));
    }

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

    let mut previous_offset = 0;
    for i in 0..count {
        let pos = checked_end(
            offsets_start,
            i.checked_mul(4).ok_or_else(|| {
                Error::Internal("string table offsets exceed addressable buffer size".to_string())
            })?,
            "string table offsets",
        )?;
        let current = read_u32(buf, pos, "string table offset")? as usize;
        if i == 0 && current != 0 {
            return Err(Error::Internal(
                "first string table offset must be zero".to_string(),
            ));
        }
        let next = if i + 1 < count {
            read_u32(
                buf,
                checked_end(pos, 4, "string table offsets")?,
                "string table offset",
            )? as usize
        } else {
            blob_size
        };
        if current > blob_size || current < previous_offset || next > blob_size || next < current {
            return Err(Error::Internal(
                "string table offsets must be monotonic and within the blob".to_string(),
            ));
        }
        let start = checked_end(blob_start, current, "string table blob")?;
        let end = checked_end(blob_start, next, "string table blob")?;
        std::str::from_utf8(&buf[start..end])
            .map_err(|e| Error::Internal(format!("invalid UTF-8 in string table: {e}")))?;
        previous_offset = current;
    }
    Ok(blob_end)
}

/// Read the string table. Returns (strings, byte position after string table).
fn read_string_table<'a>(
    buf: &'a [u8],
    offset: usize,
    budget: &mut DecodeBudget,
) -> Result<(Vec<&'a str>, usize)> {
    let blob_end = validate_string_table_layout(buf, offset)?;
    let header_end = checked_end(offset, 8, "string table header")?;
    let count = read_u32(buf, offset, "string table header")? as usize;
    let blob_size = read_u32(
        buf,
        checked_end(offset, 4, "string table header")?,
        "string table header",
    )? as usize;
    let offsets_start = header_end;
    let offsets_end = checked_end(
        offsets_start,
        count.checked_mul(4).ok_or_else(|| {
            Error::Internal("string table offsets exceed addressable buffer size".to_string())
        })?,
        "string table offsets",
    )?;
    let blob_start = offsets_end;

    let mut string_offsets = Vec::new();
    budget.charge_array::<usize>(count)?;
    string_offsets
        .try_reserve_exact(count)
        .map_err(|_| Error::Internal("string table offset allocation failed".to_string()))?;
    for i in 0..count {
        let pos = checked_end(
            offsets_start,
            i.checked_mul(4).ok_or_else(|| {
                Error::Internal("string table offsets exceed addressable buffer size".to_string())
            })?,
            "string table offsets",
        )?;
        let off = read_u32(buf, pos, "string table offset")? as usize;
        if i == 0 && off != 0 {
            return Err(Error::Internal(
                "first string table offset must be zero".to_string(),
            ));
        }
        string_offsets.push(off);
    }

    let mut strings = Vec::new();
    budget.charge_array::<&str>(count)?;
    strings
        .try_reserve_exact(count)
        .map_err(|_| Error::Internal("string table allocation failed".to_string()))?;
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
        strings.push(s);
        previous_offset = offset;
    }

    Ok((strings, blob_end))
}

fn preflight_sparse_cell_count(
    buf: &[u8],
    cell_data_start: usize,
    row_index: &[(u32, u32)],
    limits: TraversalLimits,
) -> Result<u64> {
    let mut expected_offset = 0usize;
    let mut total_cells = 0u64;
    for &(_, offset) in row_index {
        let pos = checked_end(cell_data_start, expected_offset, "sparse row data")?;
        let cell_count = read_u16(buf, pos, "sparse row cell count")? as usize;
        if offset == EMPTY_ROW_SENTINEL && cell_count != 0 {
            return Err(Error::Internal(
                "sparse empty row must have zero cells".to_string(),
            ));
        }
        if offset != EMPTY_ROW_SENTINEL && offset as usize != expected_offset {
            return Err(Error::Internal(
                "sparse row offsets are inconsistent".to_string(),
            ));
        }
        let entries_len = cell_count.checked_mul(SPARSE_ENTRY_SIZE).ok_or_else(|| {
            Error::Internal("sparse row entries exceed addressable buffer size".to_string())
        })?;
        let row_size = 2usize.checked_add(entries_len).ok_or_else(|| {
            Error::Internal("sparse row entries exceed addressable buffer size".to_string())
        })?;
        let row_end = checked_end(pos, row_size, "sparse row entries")?;
        if row_end > buf.len() {
            return Err(Error::Internal(
                "buffer too short for sparse row entries".to_string(),
            ));
        }
        total_cells = total_cells
            .checked_add(cell_count as u64)
            .ok_or_else(|| Error::Internal("decoded cell count overflow".to_string()))?;
        check_limit(
            ResourceDimension::CellCount,
            limits.max_cells(),
            total_cells,
        )?;
        expected_offset = checked_end(expected_offset, row_size, "sparse row data")?;
    }
    Ok(total_cells)
}

struct DecodeBudget {
    output_bytes: u64,
    cells: u64,
    limits: TraversalLimits,
    formula: FormulaBudget,
}

impl DecodeBudget {
    fn new(limits: TraversalLimits) -> Self {
        Self {
            output_bytes: 0,
            cells: 0,
            limits,
            formula: FormulaBudget::secure_mutation(),
        }
    }

    fn charge_output(&mut self, bytes: u64) -> Result<()> {
        self.output_bytes = checked_output_add(self.output_bytes, bytes, self.limits)?;
        Ok(())
    }

    fn charge_array<T>(&mut self, count: usize) -> Result<()> {
        let count = u64::try_from(count).map_err(|_| {
            Error::Internal("decoded collection count is not addressable".to_string())
        })?;
        let element_size = u64::try_from(std::mem::size_of::<T>())
            .map_err(|_| Error::Internal("decoded element size is not addressable".to_string()))?;
        let bytes = checked_output_mul(count, element_size, self.limits)?;
        self.charge_output(bytes)
    }

    fn copy(&mut self, value: &str) -> Result<String> {
        let bytes = u64::try_from(value.len())
            .map_err(|_| Error::Internal("raw transfer string is not addressable".to_string()))?;
        let next = checked_output_add(self.output_bytes, bytes, self.limits)?;
        let copy = fallible_string_copy(value)?;
        self.output_bytes = next;
        Ok(copy)
    }

    fn charge_cells(&mut self, count: usize) -> Result<()> {
        let next = self
            .cells
            .checked_add(count as u64)
            .ok_or_else(|| Error::Internal("decoded cell count overflow".to_string()))?;
        check_limit(ResourceDimension::CellCount, self.limits.max_cells(), next)?;
        self.cells = next;
        Ok(())
    }

    fn charge_formula(&self, bytes: usize) -> Result<()> {
        let bytes = u64::try_from(bytes)
            .map_err(|_| Error::Internal("raw transfer formula is not addressable".to_string()))?;
        self.formula.charge_mutation_cells(1)?;
        self.formula.charge_mutation_formula(bytes, bytes)
    }
}

fn decode_cell_payload(
    type_tag: u8,
    payload: &[u8],
    strings: &[&str],
    budget: &mut DecodeBudget,
) -> Result<CellValue> {
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
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            let s = budget.copy(s)?;
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
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            let s = budget.copy(s)?;
            Ok(CellValue::Error(s))
        }
        TYPE_FORMULA => {
            let idx = read_u32(payload, 0, "cell payload")? as usize;
            let expr = strings
                .get(idx)
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            budget.charge_formula(expr.len())?;
            let expr = budget.copy(expr)?;
            Ok(CellValue::Formula { expr, result: None })
        }
        TYPE_RICH_STRING => {
            let idx = read_u32(payload, 0, "cell payload")? as usize;
            let s = strings
                .get(idx)
                .ok_or_else(|| Error::Internal(format!("string index {idx} out of range")))?;
            let s = budget.copy(s)?;
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
    strings: &[&str],
    min_col: u32,
    budget: &mut DecodeBudget,
) -> Result<Vec<CellRow>> {
    let row_size = (col_count as usize)
        .checked_mul(CELL_STRIDE)
        .ok_or_else(|| Error::Internal("dense row exceeds addressable buffer size".to_string()))?;
    let mut result = Vec::new();
    budget.charge_array::<CellRow>(row_index.len())?;
    result
        .try_reserve_exact(row_index.len())
        .map_err(|_| Error::Internal("decoded row allocation failed".to_string()))?;
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
        budget.charge_array::<(u32, CellValue)>(col_count as usize)?;
        cells
            .try_reserve_exact(col_count as usize)
            .map_err(|_| Error::Internal("dense cell allocation failed".to_string()))?;
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
            let value = decode_cell_payload(type_tag, payload, strings, budget)?;
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
    strings: &[&str],
    min_col: u32,
    col_count: u16,
    budget: &mut DecodeBudget,
) -> Result<Vec<CellRow>> {
    let mut result = Vec::new();
    budget.charge_array::<CellRow>(row_index.len())?;
    result
        .try_reserve_exact(row_index.len())
        .map_err(|_| Error::Internal("decoded row allocation failed".to_string()))?;
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
        budget.charge_cells(cell_count)?;
        let mut cells = Vec::new();
        budget.charge_array::<(u32, CellValue)>(cell_count)?;
        cells
            .try_reserve_exact(cell_count)
            .map_err(|_| Error::Internal("sparse cell allocation failed".to_string()))?;
        let mut previous_col = None;
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
            if previous_col.is_some_and(|previous| col <= previous) {
                return Err(Error::Internal(
                    "sparse cell columns must be strictly increasing".to_string(),
                ));
            }
            previous_col = Some(col);
            let type_tag = *buf
                .get(checked_end(entry_off, 2, "sparse row entry")?)
                .ok_or_else(|| {
                    Error::Internal("buffer too short for sparse cell type".to_string())
                })?;
            let payload_start = checked_end(entry_off, 3, "sparse row entry")?;
            let payload_end = checked_end(entry_off, SPARSE_ENTRY_SIZE, "sparse row entry")?;
            let payload = &buf[payload_start..payload_end];
            let value = decode_cell_payload(type_tag, payload, strings, budget)?;
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
    raw_buffer_to_cells_with_limits(buf, TraversalLimits::default())
}

fn preflight_shifted_transfer_bounds(
    buf: &[u8],
    start_row: u32,
    start_col: u32,
    limits: TraversalLimits,
) -> Result<()> {
    crate::utils::cell_ref::compact_cell_ref(start_col, start_row)?;
    if buf.len() < HEADER_SIZE {
        return Err(Error::Internal(format!(
            "buffer too short for header: {} bytes (need {})",
            buf.len(),
            HEADER_SIZE
        )));
    }

    let magic = read_u32(buf, 0, "header")?;
    let (row_count, col_count, flags, is_sparse) = if magic == MAGIC_V2 {
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
        (row_count, col_count, flags, true)
    } else {
        let header = read_header(buf)?;
        (
            header.row_count,
            header.col_count,
            header.flags,
            header.flags & FLAG_SPARSE != 0,
        )
    };
    let min_col = validate_decode_layout(row_count, col_count, flags, is_sparse, limits)?;
    if row_count == 0 {
        return Ok(());
    }

    let row_index_size = (row_count as usize)
        .checked_mul(ROW_INDEX_ENTRY_SIZE)
        .ok_or_else(|| Error::Internal("row index exceeds addressable buffer size".to_string()))?;
    let row_index_end = checked_end(HEADER_SIZE, row_index_size, "row index")?;
    if buf.len() < row_index_end {
        return Err(Error::Internal(format!(
            "buffer too short for row index: {} bytes (need {})",
            buf.len(),
            row_index_end
        )));
    }
    let last_row_offset = row_index_end - ROW_INDEX_ENTRY_SIZE;
    let max_row = read_u32(buf, last_row_offset, "row index entry")?;
    validate_coordinate(max_row, min_col)?;
    max_row
        .checked_sub(1)
        .and_then(|offset| start_row.checked_add(offset))
        .filter(|row| *row <= MAX_ROWS)
        .ok_or(Error::InvalidRowNumber(u32::MAX))?;

    let max_col = min_col
        .checked_add(u32::from(col_count) - 1)
        .ok_or(Error::InvalidColumnNumber(u32::MAX))?;
    max_col
        .checked_sub(1)
        .and_then(|offset| start_col.checked_add(offset))
        .filter(|col| *col <= MAX_COLUMNS)
        .ok_or(Error::InvalidColumnNumber(u32::MAX))?;
    Ok(())
}

/// Decode and shift raw-transfer values for one atomic workbook mutation.
///
/// Formula limits are consumed during decoding and are not reset when the
/// opaque result is applied to a workbook.
pub fn raw_buffer_to_validated_cell_transfer(
    buf: &[u8],
    start_row: u32,
    start_col: u32,
) -> Result<ValidatedCellTransfer> {
    raw_buffer_to_validated_cell_transfer_with_limits(
        buf,
        start_row,
        start_col,
        TraversalLimits::default(),
    )
}

fn raw_buffer_to_validated_cell_transfer_with_limits(
    buf: &[u8],
    start_row: u32,
    start_col: u32,
    limits: TraversalLimits,
) -> Result<ValidatedCellTransfer> {
    preflight_shifted_transfer_bounds(buf, start_row, start_col, limits)?;
    let mut budget = DecodeBudget::new(limits);
    let rows = raw_buffer_to_cells_with_budget(buf, &mut budget)?;
    let mut cell_count = 0u64;
    let mut previous = None;

    for (row, cells) in &rows {
        let target_row = row
            .checked_sub(1)
            .and_then(|offset| start_row.checked_add(offset))
            .filter(|row| *row <= MAX_ROWS)
            .ok_or(Error::InvalidRowNumber(u32::MAX))?;
        for (col, value) in cells {
            let target_col = col
                .checked_sub(1)
                .and_then(|offset| start_col.checked_add(offset))
                .filter(|col| *col <= MAX_COLUMNS)
                .ok_or(Error::InvalidColumnNumber(u32::MAX))?;
            let coordinate = (target_row, target_col);
            if previous.is_some_and(|previous| coordinate <= previous) {
                return Err(Error::InvalidArgument(
                    "raw transfer destination coordinates must be unique and ordered".to_string(),
                ));
            }
            previous = Some(coordinate);

            if let CellValue::String(value) = value {
                if value.len() > crate::utils::constants::MAX_CELL_CHARS {
                    return Err(Error::CellValueTooLong {
                        length: value.len(),
                        max: crate::utils::constants::MAX_CELL_CHARS,
                    });
                }
            }

            cell_count = cell_count.checked_add(1).ok_or_else(|| {
                Error::InvalidArgument("raw transfer cell count overflows".to_string())
            })?;
            check_limit(ResourceDimension::CellCount, limits.max_cells(), cell_count)?;
            budget.charge_output(
                u64::try_from(std::mem::size_of::<(String, CellValue)>()).unwrap_or(u64::MAX),
            )?;
            budget.charge_output(
                u64::from(transfer_column_name_len(target_col))
                    + u64::from(transfer_decimal_digits(target_row)),
            )?;
        }
    }

    let mut entries = Vec::new();
    entries
        .try_reserve_exact(usize::try_from(cell_count).map_err(|_| {
            Error::InvalidArgument("raw transfer cell count is not addressable".to_string())
        })?)
        .map_err(|_| Error::Internal("raw transfer destination allocation failed".to_string()))?;
    for (row, cells) in rows {
        let target_row = start_row + row - 1;
        for (col, value) in cells {
            let target_col = start_col + col - 1;
            let cell = crate::utils::cell_ref::coordinates_to_cell_name(target_col, target_row)?;
            entries.push((cell, value));
        }
    }
    Ok(ValidatedCellTransfer { entries })
}

const fn transfer_column_name_len(column: u32) -> u32 {
    if column <= 26 {
        1
    } else if column <= 26 * 27 {
        2
    } else {
        3
    }
}

const fn transfer_decimal_digits(mut value: u32) -> u32 {
    let mut digits = 1;
    while value >= 10 {
        value /= 10;
        digits += 1;
    }
    digits
}

fn raw_buffer_to_cells_with_limits(buf: &[u8], limits: TraversalLimits) -> Result<Vec<CellRow>> {
    let mut budget = DecodeBudget::new(limits);
    raw_buffer_to_cells_with_budget(buf, &mut budget)
}

fn raw_buffer_to_cells_with_budget(buf: &[u8], budget: &mut DecodeBudget) -> Result<Vec<CellRow>> {
    if buf.len() < HEADER_SIZE {
        return Err(Error::Internal(format!(
            "buffer too short for header: {} bytes (need {})",
            buf.len(),
            HEADER_SIZE
        )));
    }
    if read_u32(buf, 0, "header")? == MAGIC_V2 {
        return raw_buffer_v2_to_cells(buf, budget);
    }
    let limits = budget.limits;
    let header = read_header(buf)?;
    if header.row_count == 0
        && header.col_count == 0
        && header.flags == 0
        && buf.len() == HEADER_SIZE
    {
        return Ok(Vec::new());
    }
    let is_sparse = header.flags & FLAG_SPARSE != 0;
    let min_col = validate_decode_layout(
        header.row_count,
        header.col_count,
        header.flags,
        is_sparse,
        limits,
    )?;
    let row_index = read_row_index(buf, header.row_count, budget)?;
    let row_index_size = (header.row_count as usize)
        .checked_mul(ROW_INDEX_ENTRY_SIZE)
        .ok_or_else(|| Error::Internal("row index exceeds addressable buffer size".to_string()))?;
    let string_table_offset = checked_end(HEADER_SIZE, row_index_size, "row index")?;
    let cell_data_start = validate_string_table_layout(buf, string_table_offset)?;
    if is_sparse {
        preflight_sparse_cell_count(buf, cell_data_start, &row_index, limits)?;
    }
    let (strings, materialized_cell_data_start) =
        read_string_table(buf, string_table_offset, budget)?;
    debug_assert_eq!(cell_data_start, materialized_cell_data_start);
    if !is_sparse {
        let dense_cells = (header.row_count as usize)
            .checked_mul(header.col_count as usize)
            .ok_or_else(|| Error::Internal("dense cell count overflow".to_string()))?;
        budget.charge_cells(dense_cells)?;
    }
    let rows = if is_sparse {
        read_sparse_cells(
            buf,
            cell_data_start,
            &row_index,
            &strings,
            min_col,
            header.col_count,
            budget,
        )?
    } else {
        read_dense_cells(
            buf,
            cell_data_start,
            &row_index,
            header.col_count,
            &strings,
            min_col,
            budget,
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

fn raw_buffer_v2_to_cells(buf: &[u8], budget: &mut DecodeBudget) -> Result<Vec<CellRow>> {
    let limits = budget.limits;
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
    let min_col = validate_decode_layout(row_count, col_count, flags, true, limits)?;
    let row_index = read_row_index(buf, row_count, budget)?;
    let row_index_size = (row_count as usize)
        .checked_mul(ROW_INDEX_ENTRY_SIZE)
        .ok_or_else(|| Error::Internal("row index exceeds addressable buffer size".to_string()))?;
    let cell_data_start = checked_end(HEADER_SIZE, row_index_size, "row index")?;
    let mut expected_offset = 0usize;
    let mut result = Vec::new();
    budget.charge_array::<CellRow>(row_index.len())?;
    result
        .try_reserve_exact(row_index.len())
        .map_err(|_| Error::Internal("decoded row allocation failed".to_string()))?;

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
        if cell_count > usize::from(col_count) {
            return Err(Error::Internal(
                "v2 cell count exceeds the column span".to_string(),
            ));
        }
        let mut cursor = checked_end(pos, 2, "v2 row data")?;
        let minimum_cell_bytes = cell_count.checked_mul(3).ok_or_else(|| {
            Error::Internal("declared v2 cell count exceeds addressable size".to_string())
        })?;
        let minimum_end = checked_end(cursor, minimum_cell_bytes, "v2 cell headers")?;
        if minimum_end > buf.len() {
            return Err(Error::Internal(
                "declared v2 cell count exceeds remaining payload".to_string(),
            ));
        }
        budget.charge_cells(cell_count)?;
        let mut cells = Vec::new();
        budget.charge_array::<(u32, CellValue)>(cell_count)?;
        cells
            .try_reserve_exact(cell_count)
            .map_err(|_| Error::Internal("v2 cell result allocation failed".to_string()))?;
        let mut previous_col = None;
        for _ in 0..cell_count {
            let col = read_u16(buf, cursor, "v2 cell column")?;
            if col >= col_count {
                return Err(Error::Internal(
                    "v2 cell column is out of range".to_string(),
                ));
            }
            if previous_col.is_some_and(|previous| col <= previous) {
                return Err(Error::Internal(
                    "v2 cell columns must be strictly increasing".to_string(),
                ));
            }
            previous_col = Some(col);
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
                    if type_tag == TYPE_FORMULA {
                        budget.charge_formula(byte_len)?;
                    }
                    let text =
                        std::str::from_utf8(buf.get(text_start..text_end).ok_or_else(|| {
                            Error::Internal("buffer too short for v2 string payload".to_string())
                        })?)
                        .map_err(|e| {
                            Error::Internal(format!("invalid UTF-8 in v2 string payload: {e}"))
                        })?;
                    let text = budget.copy(text)?;
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
    hash_index: HashMap<(u64, usize), Vec<u32>>,
    base_output_size: u64,
    encoded_size: u64,
    limits: TraversalLimits,
}

impl StringTable {
    fn new(base_output_size: u64, limits: TraversalLimits) -> Result<Self> {
        checked_output_add(base_output_size, 8, limits)?;
        Ok(Self {
            strings: Vec::new(),
            index_map: HashMap::new(),
            hash_index: HashMap::new(),
            base_output_size,
            encoded_size: 8,
            limits,
        })
    }

    fn ensure_candidate_fits(&self, len: u64) -> Result<u64> {
        let entry_size = checked_output_add(4, len, self.limits)?;
        let encoded_size = checked_output_add(self.encoded_size, entry_size, self.limits)?;
        checked_output_add(self.base_output_size, encoded_size, self.limits)?;
        Ok(encoded_size)
    }

    fn intern(&mut self, s: &str) -> Result<u32> {
        if let Some(&idx) = self.index_map.get(s) {
            return Ok(idx);
        }
        let encoded_size = self.ensure_candidate_fits(s.len() as u64)?;
        let idx = u32::try_from(self.strings.len())
            .map_err(|_| Error::Internal("string table count exceeds format".to_string()))?;
        self.strings
            .try_reserve(1)
            .map_err(|_| Error::Internal("string table allocation failed".to_string()))?;
        self.index_map
            .try_reserve(1)
            .map_err(|_| Error::Internal("string index allocation failed".to_string()))?;
        let hash_key = (hash_chunks([s.as_bytes()]), s.len());
        let new_hash_candidates = if let Some(candidates) = self.hash_index.get_mut(&hash_key) {
            candidates
                .try_reserve(1)
                .map_err(|_| Error::Internal("string hash index allocation failed".to_string()))?;
            None
        } else {
            self.hash_index
                .try_reserve(1)
                .map_err(|_| Error::Internal("string hash index allocation failed".to_string()))?;
            let mut candidates = Vec::new();
            candidates
                .try_reserve_exact(1)
                .map_err(|_| Error::Internal("string hash index allocation failed".to_string()))?;
            Some(candidates)
        };
        let stored = fallible_string_copy(s)?;
        let key = fallible_string_copy(s)?;
        self.strings.push(stored);
        self.index_map.insert(key, idx);
        if let Some(mut candidates) = new_hash_candidates {
            candidates.push(idx);
            self.hash_index.insert(hash_key, candidates);
        } else {
            self.hash_index
                .get_mut(&hash_key)
                .expect("existing hash candidate list")
                .push(idx);
        }
        self.encoded_size = encoded_size;
        Ok(idx)
    }
}

fn hash_chunks<'a>(chunks: impl IntoIterator<Item = &'a [u8]>) -> u64 {
    let mut hasher = DefaultHasher::new();
    for chunk in chunks {
        hasher.write(chunk);
    }
    hasher.finish()
}

fn fallible_string_copy(value: &str) -> Result<String> {
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|_| Error::Internal("raw transfer string allocation failed".to_string()))?;
    copy.push_str(value);
    Ok(copy)
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

fn encode_cell_payload(value: &CellValue, st: &mut StringTable) -> Result<[u8; 8]> {
    let mut payload = [0u8; 8];
    match value {
        CellValue::Empty => {}
        CellValue::Number(n) => {
            payload[0..8].copy_from_slice(&n.to_le_bytes());
        }
        CellValue::String(s) => {
            let idx = st.intern(s)?;
            payload[0..4].copy_from_slice(&idx.to_le_bytes());
        }
        CellValue::Bool(b) => {
            payload[0] = u8::from(*b);
        }
        CellValue::Date(n) => {
            payload[0..8].copy_from_slice(&n.to_le_bytes());
        }
        CellValue::Error(s) => {
            let idx = st.intern(s)?;
            payload[0..4].copy_from_slice(&idx.to_le_bytes());
        }
        CellValue::Formula { expr, .. } => {
            let idx = st.intern(expr)?;
            payload[0..4].copy_from_slice(&idx.to_le_bytes());
        }
        CellValue::RichString(runs) => {
            let plain_len = runs.iter().try_fold(0usize, |total, run| {
                total
                    .checked_add(run.text.len())
                    .ok_or_else(|| Error::Internal("rich text size overflow".to_string()))
            })?;
            let hash_key = (
                hash_chunks(runs.iter().map(|run| run.text.as_bytes())),
                plain_len,
            );
            let existing = st.hash_index.get(&hash_key).and_then(|candidates| {
                candidates.iter().copied().find(|&index| {
                    st.strings[index as usize]
                        .as_bytes()
                        .iter()
                        .copied()
                        .eq(runs
                            .iter()
                            .flat_map(|run| run.text.as_bytes().iter().copied()))
                })
            });
            let idx = if let Some(index) = existing {
                index
            } else {
                st.ensure_candidate_fits(plain_len as u64)?;
                let mut plain = String::new();
                plain
                    .try_reserve_exact(plain_len)
                    .map_err(|_| Error::Internal("rich text allocation failed".to_string()))?;
                for run in runs {
                    plain.push_str(&run.text);
                }
                st.intern(&plain)?
            };
            payload[0..4].copy_from_slice(&idx.to_le_bytes());
        }
    }
    Ok(payload)
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

    let limits = TraversalLimits::default();
    let row_count = u32::try_from(rows.len())
        .map_err(|_| Error::Internal("row count exceeds transfer format".to_string()))?;
    check_limit(
        ResourceDimension::RowCount,
        limits.max_rows(),
        u64::from(row_count),
    )?;

    let mut max_col = 0u32;
    let mut total_cells = 0u64;
    let mut previous_row = 0u32;
    for (row_num, cells) in rows {
        validate_coordinate(*row_num, 1)?;
        if *row_num <= previous_row {
            return Err(Error::Internal(
                "row numbers must be strictly increasing".to_string(),
            ));
        }
        previous_row = *row_num;
        let mut previous_col = 0u32;
        for (col, _) in cells {
            let col = *col;
            validate_coordinate(*row_num, col)?;
            if col <= previous_col {
                return Err(Error::Internal(
                    "cell columns must be strictly increasing".to_string(),
                ));
            }
            previous_col = col;
            max_col = max_col.max(col);
        }
        total_cells = total_cells
            .checked_add(cells.len() as u64)
            .ok_or_else(|| Error::Internal("cell count exceeds addressable size".to_string()))?;
    }
    check_limit(
        ResourceDimension::CellCount,
        limits.max_cells(),
        total_cells,
    )?;
    if total_cells == 0 {
        return write_empty_buffer();
    }

    let col_count = u16::try_from(max_col)
        .map_err(|_| Error::Internal("cell column is out of range".to_string()))?;
    check_limit(
        ResourceDimension::ColumnCount,
        limits.max_columns(),
        u64::from(col_count),
    )?;

    let grid_size = u64::from(row_count)
        .checked_mul(u64::from(col_count))
        .ok_or_else(|| Error::Internal("transfer grid size overflow".to_string()))?;
    let density = if grid_size > 0 {
        total_cells as f64 / grid_size as f64
    } else {
        0.0
    };
    let is_sparse = density < DENSITY_THRESHOLD;

    let row_index_size =
        checked_output_mul(u64::from(row_count), ROW_INDEX_ENTRY_SIZE as u64, limits)?;
    let cell_data_size = if is_sparse {
        compute_sparse_input_size(rows, limits)?
    } else {
        compute_dense_size(row_count, col_count, limits)?
    };
    let mut base_output_size = checked_output_add(HEADER_SIZE as u64, row_index_size, limits)?;
    base_output_size = checked_output_add(base_output_size, cell_data_size, limits)?;

    let mut st = StringTable::new(base_output_size, limits)?;
    let mut row_payloads: Vec<EncodedRow> = Vec::new();
    row_payloads
        .try_reserve_exact(rows.len())
        .map_err(|_| Error::Internal("row payload allocation failed".to_string()))?;
    for &(row_num, ref cells) in rows {
        let mut encoded_cells = Vec::new();
        encoded_cells
            .try_reserve_exact(cells.len())
            .map_err(|_| Error::Internal("cell payload allocation failed".to_string()))?;
        for &(col, ref value) in cells {
            let tag = cell_type_tag(value);
            let payload = encode_cell_payload(value, &mut st)?;
            let zero_based_col = col
                .checked_sub(1)
                .and_then(|value| u16::try_from(value).ok())
                .ok_or_else(|| Error::Internal("cell column is out of range".to_string()))?;
            encoded_cells.push((zero_based_col, tag, payload));
        }
        row_payloads.push((row_num, encoded_cells));
    }

    let string_table_size = compute_string_table_size(&st, limits)?;
    let exact_cell_data_size = if is_sparse {
        compute_sparse_size(&row_payloads, limits)?
    } else {
        cell_data_size
    };

    let mut total_size = checked_output_add(HEADER_SIZE as u64, row_index_size, limits)?;
    total_size = checked_output_add(total_size, string_table_size, limits)?;
    total_size = checked_output_add(total_size, exact_cell_data_size, limits)?;
    let total_size = usize::try_from(total_size)
        .map_err(|_| Error::Internal("raw transfer output is not addressable".to_string()))?;
    let row_index_size = usize::try_from(row_index_size)
        .map_err(|_| Error::Internal("row index is not addressable".to_string()))?;
    let string_table_size = usize::try_from(string_table_size)
        .map_err(|_| Error::Internal("string table is not addressable".to_string()))?;
    let mut buf = Vec::new();
    buf.try_reserve_exact(total_size)
        .map_err(|_| Error::Internal("raw transfer output allocation failed".to_string()))?;
    buf.resize(total_size, 0);

    write_header(
        &mut buf,
        row_count,
        col_count,
        if is_sparse { FLAG_SPARSE } else { 0 },
    );

    let string_table_start = checked_end(HEADER_SIZE, row_index_size, "raw transfer output")?;
    let cell_data_start =
        checked_end(string_table_start, string_table_size, "raw transfer output")?;

    if is_sparse {
        write_sparse_data(&mut buf, &row_payloads, cell_data_start)?;
    } else {
        write_dense_data(&mut buf, &row_payloads, col_count, cell_data_start)?;
    }

    write_row_index(&mut buf, &row_payloads, is_sparse, col_count)?;
    write_string_table(&mut buf, string_table_start, &st)?;

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

fn compute_string_table_size(st: &StringTable, limits: TraversalLimits) -> Result<u64> {
    let blob_size = st.strings.iter().try_fold(0u64, |total, string| {
        total
            .checked_add(string.len() as u64)
            .ok_or_else(|| Error::Internal("string table size overflow".to_string()))
    })?;
    let offsets_size = checked_output_mul(st.strings.len() as u64, 4, limits)?;
    let size = checked_output_add(8, offsets_size, limits)?;
    checked_output_add(size, blob_size, limits)
}

fn write_string_table(buf: &mut [u8], offset: usize, st: &StringTable) -> Result<()> {
    let count = u32::try_from(st.strings.len())
        .map_err(|_| Error::Internal("string table count exceeds format".to_string()))?;
    let blob_size = st.strings.iter().try_fold(0usize, |total, string| {
        total
            .checked_add(string.len())
            .ok_or_else(|| Error::Internal("string table size overflow".to_string()))
    })?;
    let blob_size_u32 = u32::try_from(blob_size)
        .map_err(|_| Error::Internal("string table blob exceeds format".to_string()))?;

    buf[offset..offset + 4].copy_from_slice(&count.to_le_bytes());
    buf[offset + 4..offset + 8].copy_from_slice(&blob_size_u32.to_le_bytes());

    let offsets_start = checked_end(offset, 8, "string table")?;
    let offsets_size = st
        .strings
        .len()
        .checked_mul(4)
        .ok_or_else(|| Error::Internal("string table offsets overflow".to_string()))?;
    let blob_start = checked_end(offsets_start, offsets_size, "string table")?;

    let mut blob_offset: u32 = 0;
    for (i, s) in st.strings.iter().enumerate() {
        let pos = checked_end(
            offsets_start,
            i.checked_mul(4)
                .ok_or_else(|| Error::Internal("string table offsets overflow".to_string()))?,
            "string table offsets",
        )?;
        buf[pos..pos + 4].copy_from_slice(&blob_offset.to_le_bytes());
        let dst = checked_end(blob_start, blob_offset as usize, "string table blob")?;
        buf[dst..dst + s.len()].copy_from_slice(s.as_bytes());
        blob_offset =
            blob_offset
                .checked_add(u32::try_from(s.len()).map_err(|_| {
                    Error::Internal("string table entry exceeds format".to_string())
                })?)
                .ok_or_else(|| Error::Internal("string table blob exceeds format".to_string()))?;
    }
    Ok(())
}

fn compute_dense_size(row_count: u32, col_count: u16, limits: TraversalLimits) -> Result<u64> {
    let cells = u64::from(row_count)
        .checked_mul(u64::from(col_count))
        .ok_or_else(|| Error::Internal("dense transfer grid size overflow".to_string()))?;
    checked_output_mul(cells, CELL_STRIDE as u64, limits)
}

fn compute_sparse_size(row_payloads: &[EncodedRow], limits: TraversalLimits) -> Result<u64> {
    let mut size = 0u64;
    for (_, cells) in row_payloads {
        let entries = checked_output_mul(cells.len() as u64, SPARSE_ENTRY_SIZE as u64, limits)?;
        size = checked_output_add(size, 2, limits)?;
        size = checked_output_add(size, entries, limits)?;
    }
    Ok(size)
}

fn compute_sparse_input_size(
    rows: &[(u32, Vec<(u32, CellValue)>)],
    limits: TraversalLimits,
) -> Result<u64> {
    let mut size = 0u64;
    for (_, cells) in rows {
        let entries = checked_output_mul(cells.len() as u64, SPARSE_ENTRY_SIZE as u64, limits)?;
        size = checked_output_add(size, 2, limits)?;
        size = checked_output_add(size, entries, limits)?;
    }
    Ok(size)
}

fn write_row_index(
    buf: &mut [u8],
    row_payloads: &[EncodedRow],
    is_sparse: bool,
    col_count: u16,
) -> Result<()> {
    let base = HEADER_SIZE;
    if is_sparse {
        let mut data_offset = 0u32;
        for (i, (row_num, cells)) in row_payloads.iter().enumerate() {
            let pos = checked_end(
                base,
                i.checked_mul(ROW_INDEX_ENTRY_SIZE)
                    .ok_or_else(|| Error::Internal("row index offset overflow".to_string()))?,
                "row index",
            )?;
            buf[pos..pos + 4].copy_from_slice(&row_num.to_le_bytes());
            if cells.is_empty() {
                buf[pos + 4..pos + 8].copy_from_slice(&EMPTY_ROW_SENTINEL.to_le_bytes());
            } else {
                buf[pos + 4..pos + 8].copy_from_slice(&data_offset.to_le_bytes());
            }
            let row_size = cells
                .len()
                .checked_mul(SPARSE_ENTRY_SIZE)
                .and_then(|size| size.checked_add(2))
                .ok_or_else(|| Error::Internal("sparse row size overflow".to_string()))?;
            data_offset =
                data_offset
                    .checked_add(u32::try_from(row_size).map_err(|_| {
                        Error::Internal("sparse row size exceeds format".to_string())
                    })?)
                    .ok_or_else(|| Error::Internal("sparse offsets exceed format".to_string()))?;
        }
    } else {
        for (i, (row_num, _)) in row_payloads.iter().enumerate() {
            let pos = checked_end(
                base,
                i.checked_mul(ROW_INDEX_ENTRY_SIZE)
                    .ok_or_else(|| Error::Internal("row index offset overflow".to_string()))?,
                "row index",
            )?;
            buf[pos..pos + 4].copy_from_slice(&row_num.to_le_bytes());
            let offset = (i as u64)
                .checked_mul(u64::from(col_count))
                .and_then(|value| value.checked_mul(CELL_STRIDE as u64))
                .and_then(|value| u32::try_from(value).ok())
                .ok_or_else(|| Error::Internal("dense row offset exceeds format".to_string()))?;
            buf[pos + 4..pos + 8].copy_from_slice(&offset.to_le_bytes());
        }
    }
    Ok(())
}

fn write_dense_data(
    buf: &mut [u8],
    row_payloads: &[EncodedRow],
    col_count: u16,
    cell_data_start: usize,
) -> Result<()> {
    for (i, (_, cells)) in row_payloads.iter().enumerate() {
        let row_offset = i
            .checked_mul(col_count as usize)
            .and_then(|value| value.checked_mul(CELL_STRIDE))
            .ok_or_else(|| Error::Internal("dense row offset overflow".to_string()))?;
        let row_start = checked_end(cell_data_start, row_offset, "dense row data")?;
        for &(col_idx, tag, ref payload) in cells {
            let cell_offset = (col_idx as usize)
                .checked_mul(CELL_STRIDE)
                .ok_or_else(|| Error::Internal("dense cell offset overflow".to_string()))?;
            let cell_off = checked_end(row_start, cell_offset, "dense cell data")?;
            buf[cell_off] = tag;
            buf[cell_off + 1..cell_off + 9].copy_from_slice(payload);
        }
    }
    Ok(())
}

fn write_sparse_data(
    buf: &mut [u8],
    row_payloads: &[EncodedRow],
    cell_data_start: usize,
) -> Result<()> {
    let mut offset = cell_data_start;
    for (_, cells) in row_payloads {
        let cell_count = u16::try_from(cells.len())
            .map_err(|_| Error::Internal("sparse row cell count exceeds format".to_string()))?;
        buf[offset..offset + 2].copy_from_slice(&cell_count.to_le_bytes());
        offset = checked_end(offset, 2, "sparse row data")?;
        for &(col_idx, tag, ref payload) in cells {
            buf[offset..offset + 2].copy_from_slice(&col_idx.to_le_bytes());
            buf[offset + 2] = tag;
            buf[offset + 3..offset + 11].copy_from_slice(payload);
            offset = checked_end(offset, SPARSE_ENTRY_SIZE, "sparse cell data")?;
        }
    }
    Ok(())
}

#[cfg(test)]
#[allow(clippy::needless_range_loop)]
mod tests {
    use super::*;
    use crate::rich_text::RichTextRun;

    fn v2_formula_buffer(formula: &str) -> Vec<u8> {
        v2_formula_row(&[formula.to_string()])
    }

    fn v2_formula_row(formulas: &[String]) -> Vec<u8> {
        let mut buffer = Vec::new();
        buffer.extend_from_slice(&MAGIC_V2.to_le_bytes());
        buffer.extend_from_slice(&VERSION_V2.to_le_bytes());
        buffer.extend_from_slice(&1u32.to_le_bytes());
        buffer.extend_from_slice(&(formulas.len() as u16).to_le_bytes());
        buffer.extend_from_slice(&0u32.to_le_bytes());
        buffer.extend_from_slice(&1u32.to_le_bytes());
        buffer.extend_from_slice(&0u32.to_le_bytes());
        buffer.extend_from_slice(&(formulas.len() as u16).to_le_bytes());
        for (col, formula) in formulas.iter().enumerate() {
            buffer.extend_from_slice(&(col as u16).to_le_bytes());
            buffer.push(TYPE_FORMULA);
            buffer.extend_from_slice(&(formula.len() as u32).to_le_bytes());
            buffer.extend_from_slice(formula.as_bytes());
        }
        buffer
    }

    fn assert_formula_input_limit(error: Error, observed: u64) {
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected a structured formula mutation limit error");
        };
        assert_eq!(metadata.operation_code(), "formula_mutate");
        assert_eq!(metadata.category_code(), "formula");
        assert_eq!(metadata.dimension_code(), "input_bytes");
        assert_eq!(metadata.limit(), 32_768);
        assert_eq!(metadata.observed(), observed);
        assert_eq!(metadata.stage_code(), Some("validate"));
    }

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
    fn formula_decode_limit_is_exact_for_v1_and_v2() {
        let exact = "x".repeat(32_768);
        let oversized = "x".repeat(32_769);

        for buffer in [
            cells_to_raw_buffer(&[(
                1,
                vec![(
                    1,
                    CellValue::Formula {
                        expr: exact.clone(),
                        result: None,
                    },
                )],
            )])
            .unwrap(),
            v2_formula_buffer(&exact),
        ] {
            let decoded = raw_buffer_to_cells(&buffer).unwrap();
            assert_eq!(
                decoded[0].1[0].1,
                CellValue::Formula {
                    expr: exact.clone(),
                    result: None,
                }
            );
        }

        for buffer in [
            cells_to_raw_buffer(&[(
                1,
                vec![(
                    1,
                    CellValue::Formula {
                        expr: oversized.clone(),
                        result: None,
                    },
                )],
            )])
            .unwrap(),
            v2_formula_buffer(&oversized),
        ] {
            assert_formula_input_limit(raw_buffer_to_cells(&buffer).unwrap_err(), 32_769);
        }
    }

    #[test]
    fn formula_decode_cumulative_input_limit_is_exact_for_v1_and_v2() {
        let formula = "x".repeat(32_768);
        let exact_formulas = vec![formula.clone(); 128];
        let mut plus_one_formulas = exact_formulas.clone();
        plus_one_formulas.push("x".to_string());
        let v1_rows = |formulas: &[String]| {
            vec![(
                1,
                formulas
                    .iter()
                    .enumerate()
                    .map(|(col, formula)| {
                        (
                            col as u32 + 1,
                            CellValue::Formula {
                                expr: formula.clone(),
                                result: None,
                            },
                        )
                    })
                    .collect(),
            )]
        };

        for buffer in [
            cells_to_raw_buffer(&v1_rows(&exact_formulas)).unwrap(),
            v2_formula_row(&exact_formulas),
        ] {
            let decoded = raw_buffer_to_cells(&buffer).unwrap();
            assert_eq!(decoded[0].1.len(), 128);
        }

        for buffer in [
            cells_to_raw_buffer(&v1_rows(&plus_one_formulas)).unwrap(),
            v2_formula_row(&plus_one_formulas),
        ] {
            let Error::ResourceLimitExceeded(metadata) = raw_buffer_to_cells(&buffer).unwrap_err()
            else {
                panic!("expected a structured cumulative formula limit error");
            };
            assert_eq!(metadata.operation_code(), "formula_mutate");
            assert_eq!(metadata.dimension_code(), "input_bytes");
            assert_eq!(metadata.limit(), 4 * 1024 * 1024);
            assert_eq!(metadata.observed(), 4 * 1024 * 1024 + 1);
        }
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
    fn v2_impossible_cell_count_is_rejected_before_reserve() {
        let mut buffer = Vec::new();
        buffer.extend_from_slice(&MAGIC_V2.to_le_bytes());
        buffer.extend_from_slice(&VERSION_V2.to_le_bytes());
        buffer.extend_from_slice(&1u32.to_le_bytes());
        buffer.extend_from_slice(&1u16.to_le_bytes());
        buffer.extend_from_slice(&0u32.to_le_bytes());
        buffer.extend_from_slice(&1u32.to_le_bytes());
        buffer.extend_from_slice(&0u32.to_le_bytes());
        buffer.extend_from_slice(&u16::MAX.to_le_bytes());

        let error = raw_buffer_to_cells(&buffer).unwrap_err();
        assert!(error
            .to_string()
            .contains("v2 cell count exceeds the column span"));
    }

    #[test]
    fn encoder_rejects_noncanonical_and_out_of_grid_coordinates() {
        for rows in [
            vec![(0, vec![(1, CellValue::Number(1.0))])],
            vec![(MAX_ROWS + 1, vec![(1, CellValue::Number(1.0))])],
            vec![(1, vec![(0, CellValue::Number(1.0))])],
            vec![(1, vec![(MAX_COLUMNS + 1, CellValue::Number(1.0))])],
            vec![(1, vec![(u32::MAX, CellValue::Number(1.0))])],
            vec![
                (2, vec![(1, CellValue::Number(1.0))]),
                (1, vec![(1, CellValue::Number(2.0))]),
            ],
            vec![
                (1, vec![(1, CellValue::Number(1.0))]),
                (1, vec![(2, CellValue::Number(2.0))]),
            ],
            vec![(
                1,
                vec![(2, CellValue::Number(1.0)), (1, CellValue::Number(2.0))],
            )],
            vec![(
                1,
                vec![(1, CellValue::Number(1.0)), (1, CellValue::Number(2.0))],
            )],
        ] {
            assert!(cells_to_raw_buffer(&rows).is_err());
        }

        let boundary = vec![(MAX_ROWS, vec![(MAX_COLUMNS, CellValue::Number(1.0))])];
        assert_eq!(
            raw_buffer_to_cells(&cells_to_raw_buffer(&boundary).unwrap()).unwrap(),
            boundary
        );
    }

    #[test]
    fn encoder_preserves_canonical_v1_wire_bytes() {
        let rows = vec![(1, vec![(1, CellValue::Number(42.0))])];
        let encoded = cells_to_raw_buffer(&rows).unwrap();

        let mut expected = Vec::new();
        expected.extend_from_slice(&MAGIC.to_le_bytes());
        expected.extend_from_slice(&VERSION.to_le_bytes());
        expected.extend_from_slice(&1u32.to_le_bytes());
        expected.extend_from_slice(&1u16.to_le_bytes());
        expected.extend_from_slice(&0u32.to_le_bytes());
        expected.extend_from_slice(&1u32.to_le_bytes());
        expected.extend_from_slice(&0u32.to_le_bytes());
        expected.extend_from_slice(&0u32.to_le_bytes());
        expected.extend_from_slice(&0u32.to_le_bytes());
        expected.push(TYPE_NUMBER);
        expected.extend_from_slice(&42.0f64.to_le_bytes());

        assert_eq!(encoded, expected);
    }

    #[test]
    fn encoder_size_preflight_rejects_huge_dense_output() {
        let error = compute_dense_size(
            MAX_ROWS,
            u16::try_from(MAX_COLUMNS).unwrap(),
            TraversalLimits::default(),
        )
        .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected a structured output limit error");
        };
        assert_eq!(metadata.dimension_code(), "output_bytes");
    }

    #[test]
    fn decoder_rejects_noncanonical_row_and_column_order() {
        let rows = vec![
            (
                1,
                vec![(1, CellValue::Number(1.0)), (100, CellValue::Number(2.0))],
            ),
            (2, vec![(1, CellValue::Number(3.0))]),
        ];
        let v1 = cells_to_raw_buffer(&rows).unwrap();

        let mut duplicate_row = v1.clone();
        duplicate_row[HEADER_SIZE + ROW_INDEX_ENTRY_SIZE..HEADER_SIZE + ROW_INDEX_ENTRY_SIZE + 4]
            .copy_from_slice(&1u32.to_le_bytes());
        assert!(raw_buffer_to_cells(&duplicate_row)
            .unwrap_err()
            .to_string()
            .contains("row numbers must be strictly increasing"));

        let mut duplicate_v1_column = v1;
        let cell_data = HEADER_SIZE + 2 * ROW_INDEX_ENTRY_SIZE + 8;
        let first_col = duplicate_v1_column[cell_data + 2..cell_data + 4].to_vec();
        duplicate_v1_column[cell_data + 2 + SPARSE_ENTRY_SIZE..cell_data + 4 + SPARSE_ENTRY_SIZE]
            .copy_from_slice(&first_col);
        assert!(raw_buffer_to_cells(&duplicate_v1_column)
            .unwrap_err()
            .to_string()
            .contains("sparse cell columns must be strictly increasing"));

        let mut duplicate_v2_column = Vec::new();
        duplicate_v2_column.extend_from_slice(&MAGIC_V2.to_le_bytes());
        duplicate_v2_column.extend_from_slice(&VERSION_V2.to_le_bytes());
        duplicate_v2_column.extend_from_slice(&1u32.to_le_bytes());
        duplicate_v2_column.extend_from_slice(&2u16.to_le_bytes());
        duplicate_v2_column.extend_from_slice(&0u32.to_le_bytes());
        duplicate_v2_column.extend_from_slice(&1u32.to_le_bytes());
        duplicate_v2_column.extend_from_slice(&0u32.to_le_bytes());
        duplicate_v2_column.extend_from_slice(&2u16.to_le_bytes());
        duplicate_v2_column.extend_from_slice(&0u16.to_le_bytes());
        duplicate_v2_column.push(TYPE_EMPTY);
        duplicate_v2_column.extend_from_slice(&0u16.to_le_bytes());
        duplicate_v2_column.push(TYPE_EMPTY);
        assert!(raw_buffer_to_cells(&duplicate_v2_column)
            .unwrap_err()
            .to_string()
            .contains("v2 cell columns must be strictly increasing"));
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
    fn string_table_budget_accounts_for_unique_strings_only() {
        let limits = TraversalLimits::new(1, 1, 1, 34).unwrap();
        let mut table = StringTable::new(16, limits).unwrap();

        let first = table.intern("repeat").unwrap();
        for _ in 0..100 {
            assert_eq!(table.intern("repeat").unwrap(), first);
        }
        let error = table.intern("unique").unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected a structured output limit error");
        };
        assert_eq!(metadata.dimension_code(), "output_bytes");
    }

    #[test]
    fn decoded_string_copies_obey_a_cumulative_budget() {
        let limits = TraversalLimits::new(1, 1, 1, 5).unwrap();
        let mut budget = DecodeBudget::new(limits);

        assert_eq!(budget.copy("abc").unwrap(), "abc");
        let error = budget.copy("def").unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected a structured output limit error");
        };
        assert_eq!(metadata.dimension_code(), "output_bytes");
        assert_eq!(budget.output_bytes, 3);
    }

    #[test]
    fn decoder_accounts_for_all_collections_and_string_copies_cumulatively() {
        let encoded =
            cells_to_raw_buffer(&[(1, vec![(1, CellValue::String("abc".to_string()))])]).unwrap();
        let exact_output_bytes = u64::try_from(
            std::mem::size_of::<(u32, u32)>()
                + std::mem::size_of::<usize>()
                + std::mem::size_of::<&str>()
                + std::mem::size_of::<CellRow>()
                + std::mem::size_of::<(u32, CellValue)>()
                + "abc".len(),
        )
        .unwrap();

        let exact_limits = TraversalLimits::new(1, 1, 1, exact_output_bytes).unwrap();
        assert_eq!(
            raw_buffer_to_cells_with_limits(&encoded, exact_limits).unwrap(),
            vec![(1, vec![(1, CellValue::String("abc".to_string()))])]
        );

        let plus_one_limits = TraversalLimits::new(1, 1, 1, exact_output_bytes - 1).unwrap();
        let Error::ResourceLimitExceeded(metadata) =
            raw_buffer_to_cells_with_limits(&encoded, plus_one_limits).unwrap_err()
        else {
            panic!("expected a structured cumulative output limit error");
        };
        assert_eq!(metadata.dimension_code(), "output_bytes");
        assert_eq!(metadata.observed(), exact_output_bytes);
    }

    #[test]
    fn validated_transfer_shares_decode_and_destination_output_budget() {
        let encoded =
            cells_to_raw_buffer(&[(1, vec![(1, CellValue::String("abc".to_string()))])]).unwrap();
        let exact_output_bytes = u64::try_from(
            std::mem::size_of::<(u32, u32)>()
                + std::mem::size_of::<usize>()
                + std::mem::size_of::<&str>()
                + std::mem::size_of::<CellRow>()
                + std::mem::size_of::<(u32, CellValue)>()
                + "abc".len()
                + std::mem::size_of::<(String, CellValue)>()
                + 2,
        )
        .unwrap();

        let exact_limits = TraversalLimits::new(1, 1, 1, exact_output_bytes).unwrap();
        let transfer =
            raw_buffer_to_validated_cell_transfer_with_limits(&encoded, 1, 1, exact_limits)
                .unwrap();
        assert_eq!(
            transfer.entries,
            vec![("A1".to_string(), CellValue::String("abc".to_string()))]
        );

        let plus_one_limits = TraversalLimits::new(1, 1, 1, exact_output_bytes - 1).unwrap();
        let error = match raw_buffer_to_validated_cell_transfer_with_limits(
            &encoded,
            1,
            1,
            plus_one_limits,
        ) {
            Ok(_) => panic!("expected a structured cumulative output limit error"),
            Err(error) => error,
        };
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected a structured cumulative output limit error");
        };
        assert_eq!(metadata.dimension_code(), "output_bytes");
        assert_eq!(metadata.observed(), exact_output_bytes);
    }

    #[test]
    fn decoded_cells_obey_a_cumulative_budget() {
        let limits = TraversalLimits::new(1, 1, 2, 1).unwrap();
        let mut budget = DecodeBudget::new(limits);

        budget.charge_cells(2).unwrap();
        let error = budget.charge_cells(1).unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected a structured cell limit error");
        };
        assert_eq!(metadata.dimension_code(), "cell_count");
        assert_eq!(budget.cells, 2);
    }

    #[test]
    fn decoder_rejects_invalid_column_spans_before_sections() {
        for magic in [MAGIC, MAGIC_V2] {
            let mut buffer = vec![0u8; HEADER_SIZE];
            buffer[0..4].copy_from_slice(&magic.to_le_bytes());
            buffer[4..6].copy_from_slice(
                &(if magic == MAGIC { VERSION } else { VERSION_V2 }).to_le_bytes(),
            );
            buffer[6..10].copy_from_slice(&1u32.to_le_bytes());
            buffer[10..12].copy_from_slice(&2u16.to_le_bytes());
            buffer[12..16].copy_from_slice(&(MAX_COLUMNS << 16).to_le_bytes());

            assert!(raw_buffer_to_cells(&buffer)
                .unwrap_err()
                .to_string()
                .contains("cell column span is out of range"));
        }
    }

    #[test]
    fn decoder_preflights_dense_and_sparse_cell_counts_before_materialization() {
        let mut dense = vec![0u8; HEADER_SIZE];
        dense[0..4].copy_from_slice(&MAGIC.to_le_bytes());
        dense[4..6].copy_from_slice(&VERSION.to_le_bytes());
        dense[6..10].copy_from_slice(&611u32.to_le_bytes());
        dense[10..12].copy_from_slice(&(MAX_COLUMNS as u16).to_le_bytes());
        let Error::ResourceLimitExceeded(metadata) = raw_buffer_to_cells(&dense).unwrap_err()
        else {
            panic!("expected a structured dense cell limit error");
        };
        assert_eq!(metadata.dimension_code(), "cell_count");

        let mut sparse = vec![0u8; 2 + 3 * SPARSE_ENTRY_SIZE];
        sparse[0..2].copy_from_slice(&3u16.to_le_bytes());
        let limits = TraversalLimits::new(1, 3, 2, 1).unwrap();
        let error = preflight_sparse_cell_count(&sparse, 0, &[(1, 0)], limits).unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected a structured sparse cell limit error");
        };
        assert_eq!(metadata.dimension_code(), "cell_count");
    }

    #[test]
    fn v2_sparse_layout_is_not_charged_as_a_dense_grid() {
        let row_count = 611u32;
        let mut buffer = Vec::new();
        buffer.extend_from_slice(&MAGIC_V2.to_le_bytes());
        buffer.extend_from_slice(&VERSION_V2.to_le_bytes());
        buffer.extend_from_slice(&row_count.to_le_bytes());
        buffer.extend_from_slice(&(MAX_COLUMNS as u16).to_le_bytes());
        buffer.extend_from_slice(&0u32.to_le_bytes());
        for row in 1..=row_count {
            buffer.extend_from_slice(&row.to_le_bytes());
            buffer.extend_from_slice(&((row - 1) * 13).to_le_bytes());
        }
        for _ in 0..row_count {
            buffer.extend_from_slice(&1u16.to_le_bytes());
            buffer.extend_from_slice(&0u16.to_le_bytes());
            buffer.push(TYPE_NUMBER);
            buffer.extend_from_slice(&1.0f64.to_le_bytes());
        }

        let rows = raw_buffer_to_cells(&buffer).unwrap();
        assert_eq!(rows.len(), row_count as usize);
        assert_eq!(rows.last().unwrap().0, row_count);
    }

    #[test]
    fn rich_text_admission_preserves_deduplication() {
        let limits = TraversalLimits::new(1, 1, 1, 34).unwrap();
        let mut table = StringTable::new(16, limits).unwrap();
        let expected = table.intern("repeat").unwrap();
        let runs = vec![
            RichTextRun {
                text: "re".to_string(),
                font: None,
                size: None,
                bold: false,
                italic: false,
                color: None,
            },
            RichTextRun {
                text: "peat".to_string(),
                font: None,
                size: None,
                bold: true,
                italic: false,
                color: None,
            },
        ];

        let payload = encode_cell_payload(&CellValue::RichString(runs), &mut table).unwrap();
        assert_eq!(
            u32::from_le_bytes(payload[0..4].try_into().unwrap()),
            expected
        );
        assert_eq!(table.strings.len(), 1);
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
