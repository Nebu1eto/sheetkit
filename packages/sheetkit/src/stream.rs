use napi::bindgen_prelude::*;
use napi_derive::napi;

use sheetkit_core::stream::StreamWriter;

use crate::conversions::{js_value_to_cell_value, JsCellInputValue};

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
    pub fn write_row(&mut self, row: u32, values: Vec<JsCellInputValue>) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        let cell_values = values
            .into_iter()
            .map(js_value_to_cell_value)
            .collect::<Result<Vec<_>>>()?;
        writer
            .write_row(row, &cell_values)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Write multiple rows at once starting at the given row number.
    /// More efficient than calling writeRow in a loop because it crosses
    /// the FFI boundary only once.
    #[napi]
    pub fn write_rows(&mut self, start_row: u32, rows: Vec<Vec<JsCellInputValue>>) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        let cell_rows = rows
            .into_iter()
            .map(|row| row.into_iter().map(js_value_to_cell_value).collect())
            .collect::<Result<Vec<Vec<_>>>>()?;
        writer
            .write_rows(start_row, &cell_rows)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Write a row with a specific style ID applied to all cells.
    #[napi]
    pub fn write_row_with_style(
        &mut self,
        row: u32,
        values: Vec<JsCellInputValue>,
        style_id: u32,
    ) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        let cell_values = values
            .into_iter()
            .map(js_value_to_cell_value)
            .collect::<Result<Vec<_>>>()?;
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::{DateValue, FormulaResultValue, FormulaValue};

    fn date(serial: f64) -> JsCellInputValue {
        Either6::D(DateValue {
            kind: "date".to_string(),
            serial,
            iso: None,
        })
    }

    fn formula(formula: &str, result: FormulaResultValue) -> JsCellInputValue {
        Either6::E(FormulaValue {
            kind: "formula".to_string(),
            formula: formula.to_string(),
            result: Some(result),
        })
    }

    fn number_result(value: f64) -> FormulaResultValue {
        FormulaResultValue {
            value_type: "number".to_string(),
            value: None,
            number_value: Some(value),
            bool_value: None,
            date: None,
        }
    }

    #[test]
    fn stream_writes_dates_and_formula_cached_results() {
        let path = std::env::temp_dir().join(format!(
            "sheetkit-stream-binding-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));

        let mut workbook = sheetkit_core::workbook::Workbook::new();
        let mut writer = JsStreamWriter {
            inner: Some(workbook.new_stream_writer("Streamed").unwrap()),
        };

        writer
            .write_row(1, vec![date(45306.0), formula("1+1", number_result(2.0))])
            .unwrap();
        writer
            .write_rows(
                2,
                vec![vec![
                    date(45307.0),
                    formula(
                        "TRUE",
                        FormulaResultValue {
                            value_type: "boolean".to_string(),
                            value: None,
                            number_value: None,
                            bool_value: Some(true),
                            date: None,
                        },
                    ),
                ]],
            )
            .unwrap();
        writer
            .write_row_with_style(
                3,
                vec![
                    date(45308.0),
                    formula(
                        "\"cached\"",
                        FormulaResultValue {
                            value_type: "string".to_string(),
                            value: Some("cached".to_string()),
                            number_value: None,
                            bool_value: None,
                            date: None,
                        },
                    ),
                ],
                0,
            )
            .unwrap();
        writer
            .write_row(
                4,
                vec![formula(
                    "TODAY()",
                    FormulaResultValue {
                        value_type: "date".to_string(),
                        value: None,
                        number_value: None,
                        bool_value: None,
                        date: Some(DateValue {
                            kind: "date".to_string(),
                            serial: 45309.0,
                            iso: None,
                        }),
                    },
                )],
            )
            .unwrap();

        workbook
            .apply_stream_writer(writer.inner.take().unwrap())
            .unwrap();
        workbook.save(&path).unwrap();

        let reopened = sheetkit_core::workbook::Workbook::open(&path).unwrap();
        assert_eq!(
            reopened.get_cell_value("Streamed", "A1").unwrap(),
            sheetkit_core::cell::CellValue::Date(45306.0)
        );
        assert_eq!(
            reopened.get_cell_value("Streamed", "A2").unwrap(),
            sheetkit_core::cell::CellValue::Date(45307.0)
        );
        assert_eq!(
            reopened.get_cell_value("Streamed", "A3").unwrap(),
            sheetkit_core::cell::CellValue::Date(45308.0)
        );
        assert_eq!(
            reopened.get_cell_value("Streamed", "B1").unwrap(),
            sheetkit_core::cell::CellValue::Formula {
                expr: "1+1".to_string(),
                result: Some(Box::new(sheetkit_core::cell::CellValue::Number(2.0))),
            }
        );
        assert_eq!(
            reopened.get_cell_value("Streamed", "B2").unwrap(),
            sheetkit_core::cell::CellValue::Formula {
                expr: "TRUE".to_string(),
                result: Some(Box::new(sheetkit_core::cell::CellValue::Bool(true))),
            }
        );
        assert_eq!(
            reopened.get_cell_value("Streamed", "B3").unwrap(),
            sheetkit_core::cell::CellValue::Formula {
                expr: "\"cached\"".to_string(),
                result: Some(Box::new(sheetkit_core::cell::CellValue::String(
                    "cached".to_string(),
                ))),
            }
        );
        assert_eq!(
            reopened.get_cell_value("Streamed", "A4").unwrap(),
            sheetkit_core::cell::CellValue::Formula {
                expr: "TODAY()".to_string(),
                result: Some(Box::new(sheetkit_core::cell::CellValue::Date(45309.0))),
            }
        );

        std::fs::remove_file(path).unwrap();
    }
}
