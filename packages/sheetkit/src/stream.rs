use napi::bindgen_prelude::*;
use napi_derive::napi;

use sheetkit_core::cell::CellValue;
use sheetkit_core::stream::StreamWriter;

/// Forward-only streaming writer for large sheets.
#[derive(Default)]
#[napi]
pub struct JsStreamWriter {
    pub(crate) inner: Option<StreamWriter>,
}

#[napi]
impl JsStreamWriter {
    /// Get the sheet name.
    #[napi(getter)]
    pub fn sheet_name(&self) -> Result<String> {
        let writer = self
            .inner
            .as_ref()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        Ok(writer.sheet_name().to_string())
    }

    /// Set column width (1-based column number).
    #[napi]
    pub fn set_col_width(&mut self, col: u32, width: f64) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        writer
            .set_col_width(col, width)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set column width for a range of columns.
    #[napi]
    pub fn set_col_width_range(&mut self, min_col: u32, max_col: u32, width: f64) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        writer
            .set_col_width_range(min_col, max_col, width)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Write a row of values. Rows must be written in ascending order.
    #[napi]
    pub fn write_row(
        &mut self,
        row: u32,
        values: Vec<Either4<String, f64, bool, Null>>,
    ) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        let cell_values: Vec<CellValue> = values
            .into_iter()
            .map(|v| match v {
                Either4::A(s) => CellValue::String(s),
                Either4::B(n) => CellValue::Number(n),
                Either4::C(b) => CellValue::Bool(b),
                Either4::D(_) => CellValue::Empty,
            })
            .collect();
        writer
            .write_row(row, &cell_values)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Write multiple rows at once starting at the given row number.
    /// More efficient than calling writeRow in a loop because it crosses
    /// the FFI boundary only once.
    #[napi]
    pub fn write_rows(
        &mut self,
        start_row: u32,
        rows: Vec<Vec<Either4<String, f64, bool, Null>>>,
    ) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        let cell_rows: Vec<Vec<CellValue>> = rows
            .into_iter()
            .map(|row| {
                row.into_iter()
                    .map(|v| match v {
                        Either4::A(s) => CellValue::String(s),
                        Either4::B(n) => CellValue::Number(n),
                        Either4::C(b) => CellValue::Bool(b),
                        Either4::D(_) => CellValue::Empty,
                    })
                    .collect()
            })
            .collect();
        writer
            .write_rows(start_row, &cell_rows)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Write a row with a specific style ID applied to all cells.
    #[napi]
    pub fn write_row_with_style(
        &mut self,
        row: u32,
        values: Vec<Either4<String, f64, bool, Null>>,
        style_id: u32,
    ) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        let cell_values: Vec<CellValue> = values
            .into_iter()
            .map(|v| match v {
                Either4::A(s) => CellValue::String(s),
                Either4::B(n) => CellValue::Number(n),
                Either4::C(b) => CellValue::Bool(b),
                Either4::D(_) => CellValue::Empty,
            })
            .collect();
        writer
            .write_row_with_style(row, &cell_values, style_id)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Add a merge cell reference (e.g., "A1:C3").
    #[napi]
    pub fn add_merge_cell(&mut self, reference: String) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        writer
            .add_merge_cell(&reference)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set column style for a single column (1-based).
    #[napi]
    pub fn set_col_style(&mut self, col: u32, style_id: u32) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        writer
            .set_col_style(col, style_id)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set column visibility (1-based).
    #[napi]
    pub fn set_col_visible(&mut self, col: u32, visible: bool) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        writer
            .set_col_visible(col, visible)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set column outline level (1-based, level 0-7).
    #[napi]
    pub fn set_col_outline_level(&mut self, col: u32, level: u8) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        writer
            .set_col_outline_level(col, level)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set freeze panes. The top_left_cell is the cell below and to the right
    /// of the frozen area (e.g., "A2" freezes row 1).
    #[napi]
    pub fn set_freeze_panes(&mut self, top_left_cell: String) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        writer
            .set_freeze_panes(&top_left_cell)
            .map_err(|e| Error::from_reason(e.to_string()))
    }
}
