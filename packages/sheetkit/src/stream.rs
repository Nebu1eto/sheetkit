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
}
