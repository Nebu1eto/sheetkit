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
//!   row_count: u32  = occupied rows (sparse) or rectangle rows (dense)
//!   col_count: u16  = number of columns
//!   flags:     u32  = bit 0: 1=sparse, 0=dense
//!
//! ROW INDEX (row_count * 8 bytes; sparse omits unoccupied rows)
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

use std::borrow::Cow;
use std::collections::HashMap;

use crate::cell_ref_shift::{
    shift_cell_references_with_abs_preallocated, shifted_cell_references_len_with_abs,
};
use crate::error::{Error, Result};
use crate::resource::ResourceDimension;
use crate::sst::SharedStringTable;
use crate::traversal::{
    check_limit, checked_output_add, checked_output_mul, CheckedRect, TraversalLimits,
    TraversalMode, TraversalOptions,
};
use crate::utils::cell_ref::cell_name_to_coordinates;
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS};
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

#[cfg(test)]
const EMPTY_ROW_OFFSET: u32 = 0xFFFF_FFFF;
const SPARSE_CELL_STRIDE: usize = 11;

/// Serialize a worksheet's cell data into a compact binary buffer.
///
/// Reads cell data directly from `WorksheetXml` sheet data, resolving shared
/// string references via `sst`. The default layout is sparse.
pub fn sheet_to_raw_buffer(ws: &WorksheetXml, sst: &SharedStringTable) -> Result<Vec<u8>> {
    sheet_to_raw_buffer_with_options(ws, sst, TraversalOptions::default())
}

/// Serializes a worksheet using an explicit finite traversal policy.
pub fn sheet_to_raw_buffer_with_options(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    options: TraversalOptions,
) -> Result<Vec<u8>> {
    sheet_to_raw_buffer_with_date_styles_and_options(ws, sst, &[], options)
}

/// Serialize a worksheet while promoting numeric cells that use date styles.
pub fn sheet_to_raw_buffer_with_date_styles(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    style_is_date: &[bool],
) -> Result<Vec<u8>> {
    sheet_to_raw_buffer_with_date_styles_and_options(
        ws,
        sst,
        style_is_date,
        TraversalOptions::default(),
    )
}

/// Serializes a worksheet with date styles and an explicit traversal policy.
pub fn sheet_to_raw_buffer_with_date_styles_and_options(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    style_is_date: &[bool],
    options: TraversalOptions,
) -> Result<Vec<u8>> {
    let rows = &ws.sheet_data.rows;
    let limits = options.limits();

    if rows.is_empty() {
        return write_empty_buffer(limits);
    }

    let (min_row, max_row, min_col, max_col, total_cells) = scan_dimensions(ws, limits)?;

    if total_cells == 0 {
        return write_empty_buffer(limits);
    }

    let sparse = options.mode() == TraversalMode::Sparse;
    let rect = if sparse {
        CheckedRect::new_sparse(min_row, max_row, min_col, max_col, limits)?
    } else {
        CheckedRect::new(min_row, max_row, min_col, max_col, limits)?
    };
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
    let row_count = if sparse {
        occupied_rows
    } else {
        rect.row_count()
    };
    let col_count = rect.column_count();

    let mut string_table = StringTableBuilder::new();

    let mut cell_entries = collect_cell_entries(
        ws,
        sst,
        min_col,
        &mut string_table,
        style_is_date,
        limits,
        occupied_rows,
    )?;
    cell_entries.sort_unstable_by_key(|row| row.row_num);

    let flags: u32 = (if sparse { FLAG_SPARSE } else { 0 }) | ((min_col & 0xFFFF) << 16);

    let row_index_size = checked_output_mul(row_count, 8, limits)?;
    let string_section_size = string_table.encoded_size(limits)?;
    let cell_data_size = if sparse {
        sparse_cell_data_size(&cell_entries, limits)?
    } else {
        checked_output_mul(rect.cell_count(), CELL_STRIDE as u64, limits)?
    };
    let mut total_size = checked_output_add(HEADER_SIZE as u64, row_index_size, limits)?;
    total_size = checked_output_add(total_size, string_section_size, limits)?;
    total_size = checked_output_add(total_size, cell_data_size, limits)?;
    let total_size = usize::try_from(total_size)
        .map_err(|_| Error::Internal("raw transfer output size is not addressable".to_string()))?;

    let string_section = string_table.encode(limits)?;
    let cell_data = if sparse {
        encode_sparse_cells(&cell_entries, cell_data_size, limits)?
    } else {
        encode_dense_cells(
            &cell_entries,
            min_row,
            rect.row_count(),
            col_count,
            cell_data_size,
        )?
    };
    let row_index = build_row_index(
        &cell_entries,
        min_row,
        rect.row_count(),
        col_count,
        sparse,
        row_index_size,
    )?;

    let mut buf = Vec::new();
    buf.try_reserve_exact(total_size)
        .map_err(|_| Error::Internal("raw transfer output allocation failed".to_string()))?;

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

fn write_empty_buffer(limits: TraversalLimits) -> Result<Vec<u8>> {
    checked_output_add(0, HEADER_SIZE as u64, limits)?;
    let mut buf = Vec::new();
    buf.try_reserve_exact(HEADER_SIZE)
        .map_err(|_| Error::Internal("raw transfer output allocation failed".to_string()))?;
    buf.extend_from_slice(&MAGIC.to_le_bytes());
    buf.extend_from_slice(&VERSION.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf.extend_from_slice(&0u16.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    Ok(buf)
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

#[derive(Debug)]
struct SharedFormulaMaster {
    formula: String,
    col: u32,
    row: u32,
}

/// Resolves shared-formula followers to their shifted master expression.
///
/// A follower with no usable master is encoded through its cached cell value.
/// If it has no cached value, the encoders emit `TYPE_EMPTY` rather than an
/// empty formula record.
#[derive(Debug)]
pub(crate) struct SharedFormulaResolver {
    masters: HashMap<u32, SharedFormulaMaster>,
}

pub(crate) struct RawValueBudget {
    limits: TraversalLimits,
    allocated_bytes: u64,
}

impl RawValueBudget {
    pub(crate) const fn new(limits: TraversalLimits) -> Self {
        Self {
            limits,
            allocated_bytes: 0,
        }
    }

    fn charge(&mut self, bytes: usize) -> Result<()> {
        self.allocated_bytes = checked_output_add(self.allocated_bytes, bytes as u64, self.limits)?;
        Ok(())
    }
}

pub(crate) fn inline_string_visible_text<'a>(
    inline: &'a sheetkit_xml::worksheet::InlineString,
    budget: &mut RawValueBudget,
) -> Result<Cow<'a, str>> {
    if inline.r.is_empty() {
        return Ok(Cow::Borrowed(inline.t.as_deref().unwrap_or("")));
    }
    budget.charge(inline.r.len())?;
    let visible_len = inline.r.iter().try_fold(0usize, |total, run| {
        total
            .checked_add(run.t.value.len())
            .ok_or_else(|| Error::Internal("inline rich text length overflow".to_string()))
    })?;
    budget.charge(visible_len)?;
    let mut visible = String::new();
    visible
        .try_reserve_exact(visible_len)
        .map_err(|_| Error::Internal("inline rich text allocation failed".to_string()))?;
    for run in &inline.r {
        visible.push_str(&run.t.value);
    }
    Ok(Cow::Owned(visible))
}

impl SharedFormulaResolver {
    pub(crate) fn new(ws: &WorksheetXml, budget: &mut RawValueBudget) -> Result<Self> {
        let mut masters = HashMap::new();
        for row in &ws.sheet_data.rows {
            for cell in &row.cells {
                let Some(formula) = cell.f.as_deref() else {
                    continue;
                };
                if formula.t.as_deref() != Some("shared") {
                    continue;
                }
                let (Some(si), Some(value)) = (formula.si, formula.value.as_deref()) else {
                    continue;
                };
                if value.is_empty() {
                    continue;
                }
                if let std::collections::hash_map::Entry::Vacant(entry) = masters.entry(si) {
                    let col = resolve_col(cell)?;
                    budget.charge(value.len())?;
                    let mut master_formula = String::new();
                    master_formula.try_reserve_exact(value.len()).map_err(|_| {
                        Error::Internal("shared formula allocation failed".to_string())
                    })?;
                    master_formula.push_str(value);
                    entry.insert(SharedFormulaMaster {
                        formula: master_formula,
                        col,
                        row: row.r,
                    });
                }
            }
        }
        Ok(Self { masters })
    }

    pub(crate) fn formula_for<'a>(
        &'a self,
        cell: &'a sheetkit_xml::worksheet::Cell,
        budget: &mut RawValueBudget,
    ) -> Result<Option<Cow<'a, str>>> {
        let Some(formula) = cell.f.as_deref() else {
            return Ok(None);
        };
        if formula.t.as_deref() != Some("shared")
            || formula.value.as_deref().is_some_and(|v| !v.is_empty())
        {
            return Ok(formula.value.as_deref().map(Cow::Borrowed));
        }

        let Some(master) = formula.si.and_then(|si| self.masters.get(&si)) else {
            return Ok(None);
        };
        let col = resolve_col(cell)?;
        let (_, row) = cell_name_to_coordinates(cell.r.as_str())?;
        let col_delta = i64::from(col) - i64::from(master.col);
        let row_delta = i64::from(row) - i64::from(master.row);
        let shift = |ref_col: u32, ref_row: u32, abs_col: bool, abs_row: bool| {
            let shifted_col = if abs_col {
                ref_col
            } else {
                (i64::from(ref_col) + col_delta).clamp(1, i64::from(MAX_COLUMNS)) as u32
            };
            let shifted_row = if abs_row {
                ref_row
            } else {
                (i64::from(ref_row) + row_delta).clamp(1, i64::from(MAX_ROWS)) as u32
            };
            (shifted_col, shifted_row)
        };
        budget.charge(master.formula.len())?;
        let shifted_len = shifted_cell_references_len_with_abs(&master.formula, shift)?;
        budget.charge(shifted_len)?;
        let shifted =
            shift_cell_references_with_abs_preallocated(&master.formula, shifted_len, shift)?;
        debug_assert_eq!(shifted.len(), shifted_len);
        Ok(Some(Cow::Owned(shifted)))
    }
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
            .ok_or_else(|| Error::Internal("raw transfer cell count overflow".to_string()))?;
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

fn collect_cell_entries(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    min_col: u32,
    string_table: &mut StringTableBuilder,
    style_is_date: &[bool],
    limits: TraversalLimits,
    occupied_rows: u64,
) -> Result<Vec<RowEntries>> {
    let mut result = Vec::new();
    let occupied_rows = usize::try_from(occupied_rows)
        .map_err(|_| Error::Internal("raw transfer row count is not addressable".to_string()))?;
    result
        .try_reserve_exact(occupied_rows)
        .map_err(|_| Error::Internal("raw transfer row allocation failed".to_string()))?;
    let mut raw_value_budget = RawValueBudget::new(limits);
    let shared_formulas = SharedFormulaResolver::new(ws, &mut raw_value_budget)?;

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
            .map_err(|_| Error::Internal("raw transfer cell allocation failed".to_string()))?;
        for cell in &row.cells {
            let col = resolve_col(cell)?;
            let relative_col = col - min_col;
            let (type_tag, payload) = encode_cell_value(
                cell,
                sst,
                string_table,
                &shared_formulas,
                style_is_date,
                limits,
                &mut raw_value_budget,
            )?;
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
    shared_formulas: &SharedFormulaResolver,
    style_is_date: &[bool],
    limits: TraversalLimits,
    raw_value_budget: &mut RawValueBudget,
) -> Result<(u8, [u8; 8])> {
    let mut payload = [0u8; 8];

    if cell.f.is_some() {
        if let Some(formula_expr) = shared_formulas.formula_for(cell, raw_value_budget)? {
            let idx = string_table.intern(formula_expr.as_ref(), limits)?;
            payload[..4].copy_from_slice(&(idx as u32).to_le_bytes());
            return Ok((TYPE_FORMULA, payload));
        }
    }

    match cell.t {
        CellTypeTag::SharedString => {
            if let Some(ref v) = cell.v {
                if let Ok(sst_idx) = v.parse::<usize>() {
                    let text = sst.get(sst_idx).unwrap_or("");
                    let idx = string_table.intern(text, limits)?;
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
            let idx = string_table.intern(error_text, limits)?;
            payload[..4].copy_from_slice(&(idx as u32).to_le_bytes());
            Ok((TYPE_ERROR, payload))
        }
        CellTypeTag::InlineString => {
            let text = if let Some(inline) = cell.is.as_deref() {
                inline_string_visible_text(inline, raw_value_budget)?
            } else {
                Cow::Borrowed(cell.v.as_deref().unwrap_or(""))
            };
            let idx = string_table.intern(text.as_ref(), limits)?;
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
                let idx = string_table.intern(v, limits)?;
                payload[..4].copy_from_slice(&(idx as u32).to_le_bytes());
                return Ok((TYPE_STRING, payload));
            }
            Ok((TYPE_EMPTY, payload))
        }
        CellTypeTag::None | CellTypeTag::Number => {
            if let Some(ref v) = cell.v {
                if let Ok(n) = v.parse::<f64>() {
                    payload.copy_from_slice(&n.to_le_bytes());
                    let is_date = cell
                        .s
                        .and_then(|idx| style_is_date.get(idx as usize).copied())
                        .unwrap_or(false);
                    return Ok((if is_date { TYPE_DATE } else { TYPE_NUMBER }, payload));
                }
            }
            Ok((TYPE_EMPTY, payload))
        }
    }
}

struct StringTableBuilder {
    strings: Vec<String>,
    index_map: HashMap<String, usize>,
    blob_bytes: u64,
}

impl StringTableBuilder {
    fn new() -> Self {
        Self {
            strings: Vec::new(),
            index_map: HashMap::new(),
            blob_bytes: 0,
        }
    }

    fn intern(&mut self, s: &str, limits: TraversalLimits) -> Result<usize> {
        if let Some(&idx) = self.index_map.get(s) {
            return Ok(idx);
        }
        let new_count = self.strings.len() as u64 + 1;
        check_limit(
            ResourceDimension::CollectionCount,
            limits.max_cells(),
            new_count,
        )?;
        let new_blob_bytes = checked_output_add(self.blob_bytes, s.len() as u64, limits)?;
        let offsets = checked_output_mul(new_count, 4, limits)?;
        checked_output_add(
            checked_output_add(8, offsets, limits)?,
            new_blob_bytes,
            limits,
        )?;
        let idx = self.strings.len();
        self.strings
            .try_reserve(1)
            .map_err(|_| Error::Internal("raw transfer string allocation failed".to_string()))?;
        self.index_map.try_reserve(1).map_err(|_| {
            Error::Internal("raw transfer string index allocation failed".to_string())
        })?;
        let mut owned = String::new();
        owned
            .try_reserve_exact(s.len())
            .map_err(|_| Error::Internal("raw transfer string allocation failed".to_string()))?;
        owned.push_str(s);
        self.strings.push(owned.clone());
        self.index_map.insert(owned, idx);
        self.blob_bytes = new_blob_bytes;
        Ok(idx)
    }

    fn encoded_size(&self, limits: TraversalLimits) -> Result<u64> {
        let count = self.strings.len() as u64;
        check_limit(
            ResourceDimension::CollectionCount,
            limits.max_cells(),
            count,
        )?;
        let offsets = checked_output_mul(count, 4, limits)?;
        let header_and_offsets = checked_output_add(8, offsets, limits)?;
        checked_output_add(header_and_offsets, self.blob_bytes, limits)
    }

    /// Encode the string table section: count(u32) + blob_size(u32) + offsets(u32[count]) + blob.
    fn encode(&self, limits: TraversalLimits) -> Result<Vec<u8>> {
        let total = usize::try_from(self.encoded_size(limits)?)
            .map_err(|_| Error::Internal("raw transfer string table is too large".to_string()))?;
        let count = self.strings.len() as u32;
        if count == 0 {
            let mut buf = Vec::new();
            buf.try_reserve_exact(8).map_err(|_| {
                Error::Internal("raw transfer string table allocation failed".to_string())
            })?;
            buf.extend_from_slice(&0u32.to_le_bytes());
            buf.extend_from_slice(&0u32.to_le_bytes());
            return Ok(buf);
        }

        let mut blob = Vec::new();
        let blob_len = usize::try_from(self.blob_bytes)
            .map_err(|_| Error::Internal("raw transfer string blob is too large".to_string()))?;
        blob.try_reserve_exact(blob_len).map_err(|_| {
            Error::Internal("raw transfer string blob allocation failed".to_string())
        })?;
        let mut offsets = Vec::new();
        offsets.try_reserve_exact(self.strings.len()).map_err(|_| {
            Error::Internal("raw transfer string offsets allocation failed".to_string())
        })?;
        for s in &self.strings {
            offsets.push(blob.len() as u32);
            blob.extend_from_slice(s.as_bytes());
        }
        let blob_size = blob.len() as u32;

        let mut buf = Vec::new();
        buf.try_reserve_exact(total).map_err(|_| {
            Error::Internal("raw transfer string table allocation failed".to_string())
        })?;
        buf.extend_from_slice(&count.to_le_bytes());
        buf.extend_from_slice(&blob_size.to_le_bytes());
        for off in &offsets {
            buf.extend_from_slice(&off.to_le_bytes());
        }
        buf.extend_from_slice(&blob);
        Ok(buf)
    }
}

fn encode_dense_cells(
    row_entries: &[RowEntries],
    min_row: u32,
    row_count: u64,
    col_count: u64,
    total: u64,
) -> Result<Vec<u8>> {
    let total = usize::try_from(total)
        .map_err(|_| Error::Internal("dense raw transfer output is too large".to_string()))?;
    let row_count = usize::try_from(row_count)
        .map_err(|_| Error::Internal("dense raw transfer row count is too large".to_string()))?;
    let col_count = usize::try_from(col_count)
        .map_err(|_| Error::Internal("dense raw transfer column count is too large".to_string()))?;
    let mut buf = Vec::new();
    buf.try_reserve_exact(total)
        .map_err(|_| Error::Internal("dense raw transfer allocation failed".to_string()))?;
    buf.resize(total, 0);

    for re in row_entries {
        let row_offset = (re.row_num - min_row) as usize;
        for ce in &re.cells {
            let cell_idx = row_offset * col_count + ce.col as usize;
            let pos = cell_idx * CELL_STRIDE;
            buf[pos] = ce.type_tag;
            buf[pos + 1..pos + 9].copy_from_slice(&ce.payload);
        }
    }

    debug_assert_eq!(buf.len(), row_count * col_count * CELL_STRIDE);
    Ok(buf)
}

fn sparse_cell_data_size(row_entries: &[RowEntries], limits: TraversalLimits) -> Result<u64> {
    row_entries.iter().try_fold(0u64, |total, row| {
        let cells = checked_output_mul(row.cells.len() as u64, SPARSE_CELL_STRIDE as u64, limits)?;
        checked_output_add(total, checked_output_add(2, cells, limits)?, limits)
    })
}

fn encode_sparse_cells(
    row_entries: &[RowEntries],
    total: u64,
    _limits: TraversalLimits,
) -> Result<Vec<u8>> {
    let total = usize::try_from(total)
        .map_err(|_| Error::Internal("sparse raw transfer output is too large".to_string()))?;
    let mut buf = Vec::new();
    buf.try_reserve_exact(total)
        .map_err(|_| Error::Internal("sparse raw transfer allocation failed".to_string()))?;
    for row in row_entries {
        let count = u16::try_from(row.cells.len())
            .map_err(|_| Error::Internal("raw transfer row has too many cells".to_string()))?;
        buf.extend_from_slice(&count.to_le_bytes());
        for cell in &row.cells {
            let col = u16::try_from(cell.col).map_err(|_| {
                Error::Internal("raw transfer relative column is too large".to_string())
            })?;
            buf.extend_from_slice(&col.to_le_bytes());
            buf.push(cell.type_tag);
            buf.extend_from_slice(&cell.payload);
        }
    }

    Ok(buf)
}

fn build_row_index(
    row_entries: &[RowEntries],
    min_row: u32,
    dense_row_count: u64,
    col_count: u64,
    sparse: bool,
    total: u64,
) -> Result<Vec<u8>> {
    let total = usize::try_from(total)
        .map_err(|_| Error::Internal("raw transfer row index is too large".to_string()))?;
    let mut index = Vec::new();
    index
        .try_reserve_exact(total)
        .map_err(|_| Error::Internal("raw transfer row index allocation failed".to_string()))?;

    if sparse {
        let mut sparse_offset = 0u32;
        for row in row_entries {
            index.extend_from_slice(&row.row_num.to_le_bytes());
            index.extend_from_slice(&sparse_offset.to_le_bytes());
            let row_size = 2u32
                .checked_add(
                    u32::try_from(row.cells.len())
                        .map_err(|_| Error::Internal("raw transfer row is too large".to_string()))?
                        .checked_mul(SPARSE_CELL_STRIDE as u32)
                        .ok_or_else(|| {
                            Error::Internal("raw transfer row is too large".to_string())
                        })?,
                )
                .ok_or_else(|| Error::Internal("raw transfer offset overflow".to_string()))?;
            sparse_offset = sparse_offset
                .checked_add(row_size)
                .ok_or_else(|| Error::Internal("raw transfer offset overflow".to_string()))?;
        }
    } else {
        for row_offset in 0..dense_row_count {
            let row_num = min_row
                .checked_add(u32::try_from(row_offset).map_err(|_| {
                    Error::Internal("raw transfer row offset is too large".to_string())
                })?)
                .ok_or_else(|| Error::Internal("raw transfer row overflow".to_string()))?;
            index.extend_from_slice(&row_num.to_le_bytes());
            let offset = row_offset
                .checked_mul(col_count)
                .and_then(|cells| cells.checked_mul(CELL_STRIDE as u64))
                .and_then(|bytes| u32::try_from(bytes).ok())
                .ok_or_else(|| Error::Internal("raw transfer offset overflow".to_string()))?;
            index.extend_from_slice(&offset.to_le_bytes());
        }
    }

    Ok(index)
}

#[cfg(test)]
#[allow(
    clippy::field_reassign_with_default,
    clippy::needless_lifetimes,
    clippy::approx_constant
)]
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
        assert_eq!(string_count, 1);

        let s0 = read_string(&buf, st_start, 0);
        assert_eq!(s0, "World");

        let str_idx = if is_sparse {
            let cells = read_sparse_row(&buf, cd_start);
            assert_eq!(cells[0].1, TYPE_STRING);
            u32::from_le_bytes(cells[0].2[..4].try_into().unwrap()) as usize
        } else {
            let (tag, payload) = read_cell_at(&buf, cd_start, false, 0);
            assert_eq!(tag, TYPE_STRING);
            u32::from_le_bytes(payload[..4].try_into().unwrap()) as usize
        };
        assert_eq!(
            str_idx, 0,
            "should reference local string index 0 = 'World'"
        );
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
        cell.f = Some(Box::new(CellFormula {
            t: None,
            reference: None,
            si: None,
            value: Some("A2+B2".to_string()),
        }));
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
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();
        let (st_start, _, cd_start, flags) = parse_sections(&buf);
        assert_ne!(flags & FLAG_SPARSE, 0);
        let row2 = read_sparse_row(&buf, cd_start);
        let row3_offset = read_u32_le(&buf, HEADER_SIZE + 12) as usize;
        let row3 = read_sparse_row(&buf, cd_start + row3_offset);

        for ((_, tag, payload), expected) in row2
            .into_iter()
            .chain(row3)
            .zip(["A2+$A$1", "B2+$A$1", "A3+$A$1"])
        {
            assert_eq!(tag, TYPE_FORMULA);
            let string_idx = u32::from_le_bytes(payload[..4].try_into().unwrap()) as usize;
            assert_eq!(read_string(&buf, st_start, string_idx), expected);
        }
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
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();
        let (_, _, cd_start, flags) = parse_sections(&buf);
        assert_ne!(flags & FLAG_SPARSE, 0);
        let cells = read_sparse_row(&buf, cd_start);
        let (_, cached_tag, cached_payload) = cells[0];
        assert_eq!(cached_tag, TYPE_NUMBER);
        assert_eq!(f64::from_le_bytes(cached_payload.try_into().unwrap()), 17.0);
        let (_, missing_tag, _) = cells[1];
        assert_eq!(missing_tag, TYPE_EMPTY);
    }

    #[test]
    fn shared_formula_master_clone_obeys_exact_preallocation_budget() {
        let formula = "SUM(A1:B2)+C3";
        let mut master = make_cell("A1", 1, CellTypeTag::None, Some("1"));
        master.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: Some("A1:A2".to_string()),
            si: Some(3),
            value: Some(formula.to_string()),
        }));
        let ws = make_worksheet(vec![make_row(1, vec![master])]);

        let exact_limits = TraversalLimits::new(1, 1, 1, formula.len() as u64).unwrap();
        let mut exact_budget = RawValueBudget::new(exact_limits);
        SharedFormulaResolver::new(&ws, &mut exact_budget).unwrap();

        let under_limits = TraversalLimits::new(1, 1, 1, formula.len() as u64 - 1).unwrap();
        let mut under_budget = RawValueBudget::new(under_limits);
        let error = SharedFormulaResolver::new(&ws, &mut under_budget).unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
        assert_eq!(metadata.observed(), formula.len() as u64);
    }

    #[test]
    fn shared_formula_rewrites_share_one_cumulative_budget() {
        let formula = "A1+B1";
        let mut master = make_cell("A1", 1, CellTypeTag::None, Some("1"));
        master.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: Some("A1:B1".to_string()),
            si: Some(4),
            value: Some(formula.to_string()),
        }));
        let mut follower = make_cell("B1", 2, CellTypeTag::None, Some("2"));
        follower.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: None,
            si: Some(4),
            value: None,
        }));
        let ws = make_worksheet(vec![make_row(1, vec![master, follower.clone()])]);
        let total = (formula.len() * 3) as u64;

        let exact_limits = TraversalLimits::new(1, 2, 2, total).unwrap();
        let mut exact_budget = RawValueBudget::new(exact_limits);
        let resolver = SharedFormulaResolver::new(&ws, &mut exact_budget).unwrap();
        assert_eq!(
            resolver
                .formula_for(&follower, &mut exact_budget)
                .unwrap()
                .as_deref(),
            Some("B1+C1")
        );

        let under_limits = TraversalLimits::new(1, 2, 2, total - 1).unwrap();
        let mut under_budget = RawValueBudget::new(under_limits);
        let resolver = SharedFormulaResolver::new(&ws, &mut under_budget).unwrap();
        let error = resolver
            .formula_for(&follower, &mut under_budget)
            .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.observed(), total);

        let preflight_limit = (formula.len() * 2 - 1) as u64;
        let preflight_limits = TraversalLimits::new(1, 2, 2, preflight_limit).unwrap();
        let mut preflight_budget = RawValueBudget::new(preflight_limits);
        let resolver = SharedFormulaResolver::new(&ws, &mut preflight_budget).unwrap();
        let error = resolver
            .formula_for(&follower, &mut preflight_budget)
            .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.observed(), (formula.len() * 2) as u64);
    }

    #[test]
    fn inline_rich_text_collection_obeys_exact_and_cumulative_budgets() {
        let inline = InlineString {
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
                        xml_space: None,
                        value: " Text".to_string(),
                    },
                },
            ],
            r_ph: vec![],
        };
        let charged_per_call = (inline.r.len() + "Bold Text".len()) as u64;

        let exact_limits = TraversalLimits::new(1, 1, 1, charged_per_call * 2).unwrap();
        let mut exact_budget = RawValueBudget::new(exact_limits);
        assert_eq!(
            inline_string_visible_text(&inline, &mut exact_budget)
                .unwrap()
                .as_ref(),
            "Bold Text"
        );
        inline_string_visible_text(&inline, &mut exact_budget).unwrap();

        let under_limits = TraversalLimits::new(1, 1, 1, charged_per_call * 2 - 1).unwrap();
        let mut under_budget = RawValueBudget::new(under_limits);
        inline_string_visible_text(&inline, &mut under_budget).unwrap();
        let error = inline_string_visible_text(&inline, &mut under_budget).unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.observed(), charged_per_call * 2);
    }

    #[test]
    fn empty_inline_rich_runs_are_charged_before_iteration() {
        let run_count = 16usize;
        let inline = InlineString {
            t: None,
            r: (0..run_count)
                .map(|_| R {
                    r_pr: None,
                    t: T {
                        xml_space: None,
                        value: String::new(),
                    },
                })
                .collect(),
            r_ph: vec![],
        };

        let exact_limits = TraversalLimits::new(1, 1, 1, (run_count * 2) as u64).unwrap();
        let mut exact_budget = RawValueBudget::new(exact_limits);
        assert_eq!(
            inline_string_visible_text(&inline, &mut exact_budget)
                .unwrap()
                .as_ref(),
            ""
        );
        inline_string_visible_text(&inline, &mut exact_budget).unwrap();

        let under_limits = TraversalLimits::new(1, 1, 1, (run_count * 2 - 1) as u64).unwrap();
        let mut under_budget = RawValueBudget::new(under_limits);
        inline_string_visible_text(&inline, &mut under_budget).unwrap();
        let error = inline_string_visible_text(&inline, &mut under_budget).unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.observed(), (run_count * 2) as u64);
    }

    #[test]
    fn shared_formula_extreme_offsets_stay_inside_the_excel_grid() {
        let mut master = make_cell("A1", 1, CellTypeTag::None, Some("1"));
        master.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: Some("A1:XFD1048576".to_string()),
            si: Some(8),
            value: Some("A1+XFD1048576".to_string()),
        }));
        let mut follower = make_cell("XFD1048576", 16_384, CellTypeTag::None, Some("2"));
        follower.f = Some(Box::new(CellFormula {
            t: Some("shared".to_string()),
            reference: None,
            si: Some(8),
            value: None,
        }));
        let ws = make_worksheet(vec![
            make_row(1, vec![master]),
            make_row(1_048_576, vec![follower.clone()]),
        ]);
        let mut budget = RawValueBudget::new(TraversalLimits::default());
        let resolver = SharedFormulaResolver::new(&ws, &mut budget).unwrap();

        assert_eq!(
            resolver
                .formula_for(&follower, &mut budget)
                .unwrap()
                .as_deref(),
            Some("XFD1048576+XFD1048576")
        );
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
    fn test_inline_string_preserves_escaped_visible_text() {
        let sst = SharedStringTable::new();
        let mut cell = make_cell("A1", 1, CellTypeTag::InlineString, None);
        cell.is = Some(Box::new(InlineString {
            t: Some("A < B & C".to_string()),
            r: vec![],
            r_ph: vec![],
        }));
        let ws = make_worksheet(vec![make_row(1, vec![cell])]);
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();
        let (st_start, _, cd_start, _) = parse_sections(&buf);
        let cells = read_sparse_row(&buf, cd_start);
        let (_, tag, payload) = cells[0];
        assert_eq!(tag, TYPE_STRING);
        let string_idx = u32::from_le_bytes(payload[..4].try_into().unwrap()) as usize;
        assert_eq!(read_string(&buf, st_start, string_idx), "A < B & C");
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
        let buf = sheet_to_raw_buffer(&ws, &sst).unwrap();
        let (st_start, _, cd_start, _) = parse_sections(&buf);
        let cells = read_sparse_row(&buf, cd_start);
        let (_, tag, payload) = cells[0];
        assert_eq!(tag, TYPE_STRING);
        let string_idx = u32::from_le_bytes(payload[..4].try_into().unwrap()) as usize;
        assert_eq!(read_string(&buf, st_start, string_idx), "Bold Text & More");
    }

    #[test]
    fn test_date_cell() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::Date, Some("44927.0"))],
        )]);
        let buf = sheet_to_raw_buffer_with_options(
            &ws,
            &sst,
            TraversalOptions::new(TraversalMode::Dense),
        )
        .unwrap();

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
        let buf = sheet_to_raw_buffer_with_options(
            &ws,
            &sst,
            TraversalOptions::new(TraversalMode::Dense),
        )
        .unwrap();

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
        assert_eq!(row_count, 2);
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
        assert_eq!(row_count, 2, "only occupied rows are indexed");
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
        assert_eq!(
            string_count, 2,
            "only strings used by this sheet are retained"
        );

        let s0 = read_string(&buf, st_start, 0);
        let s1 = read_string(&buf, st_start, 1);
        assert_eq!(s0, "Alpha");
        assert_eq!(s1, "Gamma");
    }

    #[test]
    fn test_string_table_ignores_unused_global_sst_entries() {
        let stale = "x".repeat(100);
        let sst = make_sst(&[stale.as_str(), "x"]);
        let ws = make_worksheet(vec![make_row(
            1,
            vec![make_cell("A1", 1, CellTypeTag::SharedString, Some("1"))],
        )]);

        for max_cells in [1, 2] {
            let limits = TraversalLimits::new(1, 1, max_cells, 85).unwrap();
            let options = TraversalOptions::new(TraversalMode::Dense).with_limits(limits);
            let buf = sheet_to_raw_buffer_with_options(&ws, &sst, options).unwrap();
            let (st_start, _, _, _) = parse_sections(&buf);

            assert_eq!(read_u32_le(&buf, st_start), 1);
            assert_eq!(read_string(&buf, st_start, 0), "x");
            assert!(buf.len() <= 85);
        }
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
        assert_eq!(row_count, 3, "only occupied rows are indexed");
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
    fn sparse_default_size_depends_on_occupied_rows_not_row_distance() {
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

        let near_buffer = sheet_to_raw_buffer(&near, &sst).unwrap();
        let far_buffer = sheet_to_raw_buffer(&far, &sst).unwrap();

        assert_eq!(read_u32_le(&near_buffer, 6), 2);
        assert_eq!(read_u32_le(&far_buffer, 6), 2);
        assert_eq!(near_buffer.len(), far_buffer.len());
        assert_eq!(read_u32_le(&far_buffer, HEADER_SIZE + 8), 1_048_576);
        let decoded = crate::raw_transfer_write::raw_buffer_to_cells(&far_buffer).unwrap();
        assert_eq!(decoded[0].0, 1);
        assert_eq!(decoded[1].0, 1_048_576);
    }

    #[test]
    fn explicit_dense_mode_is_bounded_by_the_rectangle_cell_limit() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![
            make_row(1, vec![make_cell("A1", 1, CellTypeTag::None, Some("1"))]),
            make_row(3, vec![make_cell("B3", 2, CellTypeTag::None, Some("2"))]),
        ]);
        let exact = TraversalLimits::new(3, 2, 6, 1024).unwrap();
        let buffer = sheet_to_raw_buffer_with_options(
            &ws,
            &sst,
            TraversalOptions::new(TraversalMode::Dense).with_limits(exact),
        )
        .unwrap();
        assert_eq!(read_u32_le(&buffer, 6), 3);
        assert_eq!(read_u32_le(&buffer, 12) & FLAG_SPARSE, 0);

        let too_small = TraversalLimits::new(3, 2, 5, 1024).unwrap();
        let error = sheet_to_raw_buffer_with_options(
            &ws,
            &sst,
            TraversalOptions::new(TraversalMode::Dense).with_limits(too_small),
        )
        .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.dimension(), ResourceDimension::CellCount);
        assert_eq!(metadata.limit(), 5);
        assert_eq!(metadata.observed(), 6);
    }

    #[test]
    fn sparse_mode_enforces_occupied_cell_and_output_exact_boundaries() {
        let sst = SharedStringTable::new();
        let ws = make_worksheet(vec![make_row(
            1,
            vec![
                make_cell("A1", 1, CellTypeTag::None, Some("1")),
                make_cell("B1", 2, CellTypeTag::None, Some("2")),
            ],
        )]);
        let baseline = sheet_to_raw_buffer(&ws, &sst).unwrap();
        let exact = TraversalLimits::new(1, 2, 2, baseline.len() as u64).unwrap();
        sheet_to_raw_buffer_with_options(
            &ws,
            &sst,
            TraversalOptions::new(TraversalMode::Sparse).with_limits(exact),
        )
        .unwrap();

        let cell_error = sheet_to_raw_buffer_with_options(
            &ws,
            &sst,
            TraversalOptions::new(TraversalMode::Sparse)
                .with_limits(TraversalLimits::new(1, 2, 1, 1024).unwrap()),
        )
        .unwrap_err();
        let Error::ResourceLimitExceeded(cell_metadata) = cell_error else {
            panic!("expected resource limit error");
        };
        assert_eq!(cell_metadata.dimension(), ResourceDimension::CellCount);
        assert_eq!(cell_metadata.observed(), 2);

        let output_error = sheet_to_raw_buffer_with_options(
            &ws,
            &sst,
            TraversalOptions::new(TraversalMode::Sparse)
                .with_limits(TraversalLimits::new(1, 2, 2, baseline.len() as u64 - 1).unwrap()),
        )
        .unwrap_err();
        let Error::ResourceLimitExceeded(output_metadata) = output_error else {
            panic!("expected resource limit error");
        };
        assert_eq!(output_metadata.dimension(), ResourceDimension::OutputBytes);
        assert_eq!(output_metadata.observed(), baseline.len() as u64);
    }

    #[test]
    fn empty_output_still_obeys_the_output_limit() {
        let limits = TraversalLimits::new(1, 1, 1, HEADER_SIZE as u64 - 1).unwrap();
        let error = sheet_to_raw_buffer_with_options(
            &make_worksheet(vec![]),
            &SharedStringTable::new(),
            TraversalOptions::default().with_limits(limits),
        )
        .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
        assert_eq!(metadata.observed(), HEADER_SIZE as u64);
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
        assert_eq!(row_count, 2);

        let row_index_start = HEADER_SIZE;
        let row1_num = read_u32_le(&buf, row_index_start);
        let row1_offset = read_u32_le(&buf, row_index_start + 4);
        assert_eq!(row1_num, 1);
        assert_ne!(row1_offset, EMPTY_ROW_OFFSET);

        let row3_num = read_u32_le(&buf, row_index_start + 8);
        let row3_offset = read_u32_le(&buf, row_index_start + 12);
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
