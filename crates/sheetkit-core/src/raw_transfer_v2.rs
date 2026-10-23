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
//!   row_count: u32  = number of occupied rows
//!   col_count: u16  = number of columns
//!   flags:     u32  = bit 0: always 0 (v2 is always sparse)
//!                      bits 16..31: min_col (1-based)
//!
//! ROW INDEX (row_count * 8 bytes; unoccupied rows are omitted)
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

use std::borrow::Cow;

use crate::error::{Error, Result};
use crate::raw_transfer::{inline_string_visible_text, RawValueBudget, SharedFormulaResolver};
use crate::resource::ResourceDimension;
use crate::sst::SharedStringTable;
use crate::traversal::{
    check_limit, checked_output_add, checked_output_mul, CheckedRect, TraversalLimits,
    TraversalMode, TraversalOptions,
};
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

#[cfg(test)]
const EMPTY_ROW_OFFSET: u32 = 0xFFFF_FFFF;

/// A cell entry with inline string data (variable-length payload).
struct CellEntryV2<'a> {
    col: u16,
    type_tag: u8,
    payload: CellPayload<'a>,
}

/// Cell payload. Strings borrow from the worksheet or shared string table
/// whenever possible; only derived text (shifted shared formulas, rich
/// inline strings) is owned.
enum CellPayload<'a> {
    Empty,
    Number(f64),
    Bool(u8),
    Str(Cow<'a, str>),
}

impl CellPayload<'_> {
    fn byte_size(&self, limits: TraversalLimits) -> Result<u64> {
        match self {
            CellPayload::Empty => Ok(0),
            CellPayload::Number(_) => Ok(8),
            CellPayload::Bool(_) => Ok(1),
            CellPayload::Str(s) => checked_output_add(4, s.len() as u64, limits),
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

struct RowEntriesV2<'a> {
    row_num: u32,
    cells: Vec<CellEntryV2<'a>>,
}

impl RowEntriesV2<'_> {
    /// Byte size of the cell data for this row (cell_count u16 + all cells).
    fn byte_size(&self, limits: TraversalLimits) -> Result<u64> {
        let mut size = 2u64;
        for cell in &self.cells {
            size = checked_output_add(size, 3, limits)?;
            size = checked_output_add(size, cell.payload.byte_size(limits)?, limits)?;
        }
        Ok(size)
    }
}

/// Serialize a worksheet's cell data into a compact binary buffer using v2 format.
///
/// The v2 format inlines all string data with each cell, eliminating the global
/// string table. This enables incremental row-by-row decoding without eagerly
/// decoding all strings.
pub fn sheet_to_raw_buffer_v2(ws: &WorksheetXml, sst: &SharedStringTable) -> Result<Vec<u8>> {
    sheet_to_raw_buffer_v2_with_options(ws, sst, TraversalOptions::default())
}

/// Serializes a worksheet in canonical v2 sparse format with finite options.
pub fn sheet_to_raw_buffer_v2_with_options(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    options: TraversalOptions,
) -> Result<Vec<u8>> {
    sheet_to_raw_buffer_v2_with_date_styles_and_options(ws, sst, &[], options)
}

/// Serialize a worksheet while promoting numeric cells that use date styles.
pub fn sheet_to_raw_buffer_v2_with_date_styles(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    style_is_date: &[bool],
) -> Result<Vec<u8>> {
    sheet_to_raw_buffer_v2_with_date_styles_and_options(
        ws,
        sst,
        style_is_date,
        TraversalOptions::default(),
    )
}

/// Serializes a worksheet in canonical v2 sparse format with date styles.
pub fn sheet_to_raw_buffer_v2_with_date_styles_and_options(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    style_is_date: &[bool],
    options: TraversalOptions,
) -> Result<Vec<u8>> {
    if options.mode() != TraversalMode::Sparse {
        return Err(Error::Internal(
            "raw transfer v2 supports sparse traversal only".to_string(),
        ));
    }
    let rows = &ws.sheet_data.rows;
    let limits = options.limits();

    if rows.is_empty() {
        return write_empty_buffer(limits);
    }

    let (min_row, max_row, min_col, max_col, total_cells) = scan_dimensions(ws, limits)?;

    if total_cells == 0 {
        return write_empty_buffer(limits);
    }

    let rect = CheckedRect::new_sparse(min_row, max_row, min_col, max_col, limits)?;
    let occupied_rows = rows.iter().filter(|row| !row.cells.is_empty()).count() as u64;
    check_limit(
        ResourceDimension::RowCount,
        limits.max_rows(),
        occupied_rows,
    )?;
    check_limit(
        ResourceDimension::CellCount,
        limits.max_cells(),
        total_cells,
    )?;
    let row_count = occupied_rows;
    let col_count = rect.column_count();

    let mut raw_value_budget = RawValueBudget::new(limits);
    let shared_formulas = SharedFormulaResolver::new(ws, &mut raw_value_budget)?;
    let mut cell_entries = collect_cell_entries_v2(
        ws,
        sst,
        &shared_formulas,
        &mut raw_value_budget,
        min_col,
        style_is_date,
        limits,
        occupied_rows,
    )?;
    cell_entries.sort_unstable_by_key(|row| row.row_num);

    let flags: u32 = (min_col & 0xFFFF) << 16;

    let row_index_size = checked_output_mul(row_count, 8, limits)?;
    let mut row_sizes = Vec::new();
    row_sizes
        .try_reserve_exact(cell_entries.len())
        .map_err(|_| Error::Internal("raw transfer v2 row size allocation failed".to_string()))?;
    let cell_data_size = cell_entries.iter().try_fold(0u64, |total, row| {
        let row_size = row.byte_size(limits)?;
        row_sizes.push(row_size);
        checked_output_add(total, row_size, limits)
    })?;
    let total_size = checked_output_add(
        checked_output_add(HEADER_SIZE as u64, row_index_size, limits)?,
        cell_data_size,
        limits,
    )?;
    let total_size = usize::try_from(total_size)
        .map_err(|_| Error::Internal("raw transfer v2 output is not addressable".to_string()))?;

    validate_cell_data_v2(&cell_entries, cell_data_size)?;
    validate_row_index_v2(&row_sizes, row_index_size)?;

    let mut buf = Vec::new();
    buf.try_reserve_exact(total_size)
        .map_err(|_| Error::Internal("raw transfer v2 output allocation failed".to_string()))?;

    buf.extend_from_slice(&MAGIC_V2.to_le_bytes());
    buf.extend_from_slice(&VERSION_V2.to_le_bytes());
    buf.extend_from_slice(&(row_count as u32).to_le_bytes());
    buf.extend_from_slice(&(col_count as u16).to_le_bytes());
    buf.extend_from_slice(&flags.to_le_bytes());

    write_row_index_v2(&mut buf, &cell_entries, &row_sizes);
    write_cell_data_v2(&mut buf, &cell_entries);
    debug_assert_eq!(buf.len(), total_size);

    Ok(buf)
}

fn write_empty_buffer(limits: TraversalLimits) -> Result<Vec<u8>> {
    checked_output_add(0, HEADER_SIZE as u64, limits)?;
    let mut buf = Vec::new();
    buf.try_reserve_exact(HEADER_SIZE)
        .map_err(|_| Error::Internal("raw transfer v2 output allocation failed".to_string()))?;
    buf.extend_from_slice(&MAGIC_V2.to_le_bytes());
    buf.extend_from_slice(&VERSION_V2.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf.extend_from_slice(&0u16.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    Ok(buf)
}

fn scan_dimensions(
    ws: &WorksheetXml,
    limits: TraversalLimits,
) -> Result<(u32, u32, u32, u32, u64)> {
    let mut min_row = u32::MAX;
    let mut max_row = 0u32;
    let mut min_col = u32::MAX;
    let mut max_col = 0u32;
    let mut total_cells = 0u64;

    for row in &ws.sheet_data.rows {
        if row.cells.is_empty() {
            continue;
        }
        min_row = min_row.min(row.r);
        max_row = max_row.max(row.r);
        total_cells = total_cells
            .checked_add(row.cells.len() as u64)
            .ok_or_else(|| Error::Internal("raw transfer v2 cell count overflow".to_string()))?;
        check_limit(
            ResourceDimension::CellCount,
            limits.max_cells(),
            total_cells,
        )?;
        for cell in &row.cells {
            let col = resolve_col(cell)?;
            min_col = min_col.min(col);
            max_col = max_col.max(col);
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

#[allow(clippy::too_many_arguments)]
fn collect_cell_entries_v2<'a>(
    ws: &'a WorksheetXml,
    sst: &'a SharedStringTable,
    shared_formulas: &'a SharedFormulaResolver,
    raw_value_budget: &mut RawValueBudget,
    min_col: u32,
    style_is_date: &[bool],
    limits: TraversalLimits,
    occupied_rows: u64,
) -> Result<Vec<RowEntriesV2<'a>>> {
    let mut result = Vec::new();
    let occupied_rows = usize::try_from(occupied_rows)
        .map_err(|_| Error::Internal("raw transfer v2 row count is not addressable".to_string()))?;
    result
        .try_reserve_exact(occupied_rows)
        .map_err(|_| Error::Internal("raw transfer v2 row allocation failed".to_string()))?;
    let mut payload_bytes = 0u64;

    for row in &ws.sheet_data.rows {
        if row.cells.is_empty() {
            continue;
        }

        check_limit(
            ResourceDimension::ColumnCount,
            limits.max_columns(),
            row.cells.len() as u64,
        )?;
        let mut cells = Vec::new();
        cells
            .try_reserve_exact(row.cells.len())
            .map_err(|_| Error::Internal("raw transfer v2 cell allocation failed".to_string()))?;
        for cell in &row.cells {
            let col = resolve_col(cell)?;
            let relative_col = (col - min_col) as u16;
            let (type_tag, payload) = encode_cell_value_v2(
                cell,
                sst,
                shared_formulas,
                style_is_date,
                limits,
                &mut payload_bytes,
                raw_value_budget,
            )?;
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

fn encode_cell_value_v2<'a>(
    cell: &'a sheetkit_xml::worksheet::Cell,
    sst: &'a SharedStringTable,
    shared_formulas: &'a SharedFormulaResolver,
    style_is_date: &[bool],
    limits: TraversalLimits,
    payload_bytes: &mut u64,
    raw_value_budget: &mut RawValueBudget,
) -> Result<(u8, CellPayload<'a>)> {
    if cell.f.is_some() {
        if let Some(formula_expr) = shared_formulas.formula_for(cell, raw_value_budget)? {
            let formula_expr = bounded_string(formula_expr, limits, payload_bytes)?;
            return Ok((TYPE_FORMULA, CellPayload::Str(formula_expr)));
        }
    }

    match cell.t {
        CellTypeTag::SharedString => {
            if let Some(ref v) = cell.v {
                if let Ok(sst_idx) = v.parse::<usize>() {
                    let text = sst.get(sst_idx).unwrap_or("");
                    if sst.get_rich_text(sst_idx).is_some() {
                        return Ok((
                            TYPE_RICH_STRING,
                            CellPayload::Str(bounded_string(
                                Cow::Borrowed(text),
                                limits,
                                payload_bytes,
                            )?),
                        ));
                    }
                    return Ok((
                        TYPE_STRING,
                        CellPayload::Str(bounded_string(
                            Cow::Borrowed(text),
                            limits,
                            payload_bytes,
                        )?),
                    ));
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
            Ok((
                TYPE_ERROR,
                CellPayload::Str(bounded_string(
                    Cow::Borrowed(error_text),
                    limits,
                    payload_bytes,
                )?),
            ))
        }
        CellTypeTag::InlineString => {
            let text = if let Some(inline) = cell.is.as_deref() {
                inline_string_visible_text(inline, raw_value_budget)?
            } else {
                Cow::Borrowed(cell.v.as_deref().unwrap_or(""))
            };
            Ok((
                TYPE_STRING,
                CellPayload::Str(bounded_string(text, limits, payload_bytes)?),
            ))
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
                return Ok((
                    TYPE_STRING,
                    CellPayload::Str(bounded_string(Cow::Borrowed(v), limits, payload_bytes)?),
                ));
            }
            Ok((TYPE_EMPTY, CellPayload::Empty))
        }
        CellTypeTag::None | CellTypeTag::Number => {
            if let Some(ref v) = cell.v {
                if let Ok(n) = v.parse::<f64>() {
                    let is_date = cell
                        .s
                        .and_then(|idx| style_is_date.get(idx as usize).copied())
                        .unwrap_or(false);
                    return Ok((
                        if is_date { TYPE_DATE } else { TYPE_NUMBER },
                        CellPayload::Number(n),
                    ));
                }
            }
            Ok((TYPE_EMPTY, CellPayload::Empty))
        }
    }
}

/// Charge one string payload against the output budget. The payload is
/// borrowed, so no per-cell copy is made before the final buffer write.
fn bounded_string<'a>(
    value: Cow<'a, str>,
    limits: TraversalLimits,
    payload_bytes: &mut u64,
) -> Result<Cow<'a, str>> {
    *payload_bytes = checked_output_add(*payload_bytes, value.len() as u64, limits)?;
    Ok(value)
}

/// Check every cell-data conversion before the single output write.
fn validate_cell_data_v2(row_entries: &[RowEntriesV2<'_>], total: u64) -> Result<()> {
    usize::try_from(total)
        .map_err(|_| Error::Internal("raw transfer v2 cell data is too large".to_string()))?;
    for row in row_entries {
        u16::try_from(row.cells.len())
            .map_err(|_| Error::Internal("raw transfer v2 row has too many cells".to_string()))?;
    }
    Ok(())
}

/// Check every row-index conversion before the single output write.
fn validate_row_index_v2(row_sizes: &[u64], total: u64) -> Result<()> {
    usize::try_from(total)
        .map_err(|_| Error::Internal("raw transfer v2 row index is too large".to_string()))?;
    let mut byte_offset = 0u32;
    for row_size in row_sizes {
        let row_size = u32::try_from(*row_size)
            .map_err(|_| Error::Internal("raw transfer v2 row is too large".to_string()))?;
        byte_offset = byte_offset
            .checked_add(row_size)
            .ok_or_else(|| Error::Internal("raw transfer v2 offset overflow".to_string()))?;
    }
    Ok(())
}

/// Write the row index. Conversions were checked by [`validate_row_index_v2`].
fn write_row_index_v2(buf: &mut Vec<u8>, row_entries: &[RowEntriesV2<'_>], row_sizes: &[u64]) {
    let mut byte_offset = 0u32;
    for (row, row_size) in row_entries.iter().zip(row_sizes) {
        buf.extend_from_slice(&row.row_num.to_le_bytes());
        buf.extend_from_slice(&byte_offset.to_le_bytes());
        byte_offset = byte_offset.wrapping_add(*row_size as u32);
    }
}

/// Write the cell data. Conversions were checked by [`validate_cell_data_v2`].
fn write_cell_data_v2(buf: &mut Vec<u8>, row_entries: &[RowEntriesV2<'_>]) {
    for row in row_entries {
        buf.extend_from_slice(&(row.cells.len() as u16).to_le_bytes());
        for cell in &row.cells {
            buf.extend_from_slice(&cell.col.to_le_bytes());
            buf.push(cell.type_tag);
            cell.payload.write_to(buf);
        }
    }
}

#[cfg(test)]
#[allow(clippy::field_reassign_with_default)]
mod tests {
    use super::*;
    use sheetkit_xml::shared_strings::{R, T};
    use sheetkit_xml::worksheet::{
        Cell, CellFormula, CellTypeTag, CompactCellRef, InlineString, Row, SheetData, WorksheetXml,
    };

    fn make_cell(col_ref: &str, col_num: u32, t: CellTypeTag, v: Option<&str>) -> Cell {
        Cell {
            r: CompactCellRef::new(col_ref).unwrap(),
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
    fn test_shared_formula_followers_shift_relative_references() {
        let sst = SharedStringTable::new();
        let mut master = make_cell("B2", 2, CellTypeTag::None, Some("1"));
        master.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: Some("B2:C3".to_string()),
            si: Some(7),
            value: Some("A2+$A$1".to_string()),
        }));
        let mut across = make_cell("C2", 3, CellTypeTag::None, Some("2"));
        across.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: None,
            si: Some(7),
            value: None,
        }));
        let mut down = make_cell("B3", 2, CellTypeTag::None, Some("3"));
        down.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: None,
            si: Some(7),
            value: None,
        }));
        let ws = make_worksheet(vec![
            make_row(2, vec![master, across]),
            make_row(3, vec![down]),
        ]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();
        let cell_data_start = HEADER_SIZE + 16;
        let row2_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        let row3_offset = read_u32_le(&buf, HEADER_SIZE + 12);
        let row2 = read_v2_row_cells(&buf, cell_data_start + row2_offset as usize);
        let row3 = read_v2_row_cells(&buf, cell_data_start + row3_offset as usize);

        assert_eq!(
            row2,
            vec![
                (0, TYPE_FORMULA, "A2+$A$1".to_string()),
                (1, TYPE_FORMULA, "B2+$A$1".to_string())
            ]
        );
        assert_eq!(row3, vec![(0, TYPE_FORMULA, "A3+$A$1".to_string())]);
    }

    #[test]
    fn v1_and_v2_match_for_budgeted_shared_formulas_and_inline_rich_text() {
        let sst = SharedStringTable::new();
        let mut master = make_cell("A1", 1, CellTypeTag::None, Some("1"));
        master.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: Some("A1:B1".to_string()),
            si: Some(12),
            value: Some("A1+$A$1".to_string()),
        }));
        let mut follower = make_cell("B1", 2, CellTypeTag::None, Some("2"));
        follower.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: None,
            si: Some(12),
            value: None,
        }));
        let mut rich = make_cell("C1", 3, CellTypeTag::InlineString, None);
        rich.is = Some(Box::new(InlineString {
            t: None,
            r: vec![
                R {
                    r_pr: None,
                    t: T {
                        xml_space: None,
                        value: "Rich".to_string(),
                    },
                },
                R {
                    r_pr: None,
                    t: T {
                        xml_space: None,
                        value: " Text".to_string(),
                    },
                },
            ],
            r_ph: vec![],
        }));
        let ws = make_worksheet(vec![make_row(1, vec![master, follower, rich])]);

        let v1 = crate::raw_transfer::sheet_to_raw_buffer(&ws, &sst).unwrap();
        let v2 = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();

        assert_eq!(
            crate::raw_transfer_write::raw_buffer_to_cells(&v1).unwrap(),
            crate::raw_transfer_write::raw_buffer_to_cells(&v2).unwrap()
        );
    }

    #[test]
    fn test_shared_formula_without_master_uses_cached_value_or_empty() {
        let sst = SharedStringTable::new();
        let mut cached = make_cell("A1", 1, CellTypeTag::None, Some("17"));
        cached.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: None,
            si: Some(99),
            value: None,
        }));
        let mut missing = make_cell("B1", 2, CellTypeTag::None, None);
        missing.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: None,
            si: Some(99),
            value: None,
        }));
        let ws = make_worksheet(vec![make_row(1, vec![cached, missing])]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();
        let cells = read_v2_row_cells(&buf, HEADER_SIZE + 8);

        assert_eq!(cells[0], (0, TYPE_NUMBER, "17".to_string()));
        assert_eq!(cells[1], (1, TYPE_EMPTY, String::new()));
    }

    #[test]
    fn test_inline_string_cell() {
        let sst = SharedStringTable::new();
        let mut cell = make_cell("A1", 1, CellTypeTag::InlineString, None);
        cell.is = Some(Box::new(InlineString {
            t: Some("Inline Text".to_string()),
            r: vec![],
            r_ph: vec![],
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
    fn test_inline_string_preserves_escaped_visible_text() {
        let sst = SharedStringTable::new();
        let mut cell = make_cell("A1", 1, CellTypeTag::InlineString, None);
        cell.is = Some(Box::new(InlineString {
            t: Some("A < B & C".to_string()),
            r: vec![],
            r_ph: vec![],
        }));
        let ws = make_worksheet(vec![make_row(1, vec![cell])]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();
        let cells = read_v2_row_cells(&buf, HEADER_SIZE + 8);

        assert_eq!(cells, vec![(0, TYPE_STRING, "A < B & C".to_string())]);
    }

    #[test]
    fn test_inline_rich_text_concatenates_runs_without_phonetics() {
        let sst = SharedStringTable::new();
        let mut cell = make_cell("A1", 1, CellTypeTag::InlineString, None);
        cell.is = Some(Box::new(InlineString {
            t: None,
            r: vec![
                R {
                    r_pr: None,
                    t: T {
                        xml_space: None,
                        value: "Bold".to_string(),
                    },
                },
                R {
                    r_pr: None,
                    t: T {
                        xml_space: Some("preserve".to_string()),
                        value: " Text & More".to_string(),
                    },
                },
            ],
            r_ph: vec![sheetkit_xml::worksheet::PhoneticRun {
                sb: Some(0),
                eb: Some(4),
                value: "phonetic".to_string(),
            }],
        }));
        let ws = make_worksheet(vec![make_row(1, vec![cell])]);
        let buf = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();
        let cells = read_v2_row_cells(&buf, HEADER_SIZE + 8);

        assert_eq!(
            cells,
            vec![(0, TYPE_STRING, "Bold Text & More".to_string())]
        );
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
        assert_eq!(row_count, 2);
        let col_count = read_u16_le(&buf, 10) as usize;
        assert_eq!(col_count, 20);

        // Verify row 1 has data
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        assert_ne!(row1_offset, EMPTY_ROW_OFFSET);

        // Verify the second entry preserves absolute row 100.
        assert_eq!(read_u32_le(&buf, HEADER_SIZE + 8), 100);
        let row100_offset = read_u32_le(&buf, HEADER_SIZE + 12);
        assert_ne!(row100_offset, EMPTY_ROW_OFFSET);
    }

    #[test]
    fn canonical_sparse_size_depends_on_occupied_rows_not_row_distance() {
        let sst = SharedStringTable::new();
        let near = make_worksheet(vec![
            make_row(1, vec![make_cell("A1", 1, CellTypeTag::None, Some("1"))]),
            make_row(2, vec![make_cell("A2", 1, CellTypeTag::None, Some("2"))]),
        ]);
        let far = make_worksheet(vec![
            make_row(1, vec![make_cell("A1", 1, CellTypeTag::None, Some("1"))]),
            make_row(
                1_048_576,
                vec![make_cell("A1048576", 1, CellTypeTag::None, Some("2"))],
            ),
        ]);

        let near_buffer = sheet_to_raw_buffer_v2(&near, &sst).unwrap();
        let far_buffer = sheet_to_raw_buffer_v2(&far, &sst).unwrap();

        assert_eq!(read_u32_le(&near_buffer, 6), 2);
        assert_eq!(read_u32_le(&far_buffer, 6), 2);
        assert_eq!(near_buffer.len(), far_buffer.len());
        assert_eq!(read_u32_le(&far_buffer, HEADER_SIZE + 8), 1_048_576);
        let decoded = crate::raw_transfer_write::raw_buffer_to_cells(&far_buffer).unwrap();
        assert_eq!(decoded[0].0, 1);
        assert_eq!(decoded[1].0, 1_048_576);
    }

    #[test]
    fn canonical_sparse_enforces_exact_output_boundary() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell(
                "A1",
                1,
                CellTypeTag::InlineString,
                Some("payload"),
            )],
        )]);
        let baseline = sheet_to_raw_buffer_v2(&ws, &sst).unwrap();
        let exact = TraversalLimits::new(1, 1, 1, baseline.len() as u64).unwrap();
        sheet_to_raw_buffer_v2_with_options(
            &ws,
            &sst,
            TraversalOptions::default().with_limits(exact),
        )
        .unwrap();

        let error = sheet_to_raw_buffer_v2_with_options(
            &ws,
            &sst,
            TraversalOptions::default()
                .with_limits(TraversalLimits::new(1, 1, 1, baseline.len() as u64 - 1).unwrap()),
        )
        .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
        assert_eq!(metadata.observed(), baseline.len() as u64);
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
        assert_eq!(row_count, 2);

        let row1_num = read_u32_le(&buf, HEADER_SIZE);
        let row1_offset = read_u32_le(&buf, HEADER_SIZE + 4);
        assert_eq!(row1_num, 1);
        assert_ne!(row1_offset, EMPTY_ROW_OFFSET);

        let row3_num = read_u32_le(&buf, HEADER_SIZE + 8);
        let row3_offset = read_u32_le(&buf, HEADER_SIZE + 12);
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
        assert_eq!(row_count, 2, "only occupied rows are indexed");
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
