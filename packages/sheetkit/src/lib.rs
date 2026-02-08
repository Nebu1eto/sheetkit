#![deny(clippy::all)]

mod conversions;
mod stream;
mod types;

use napi::bindgen_prelude::*;
use napi_derive::napi;

use sheetkit_core::cell::CellValue;
use sheetkit_core::chart::{ChartConfig, ChartSeries, View3DConfig};
use sheetkit_core::comment::CommentConfig;
use sheetkit_core::conditional::ConditionalFormatRule;
use sheetkit_core::doc_props::CustomPropertyValue;
use sheetkit_core::image::ImageConfig;
use sheetkit_core::page_layout::PageMarginsConfig;
use sheetkit_core::pivot::{PivotDataField, PivotField, PivotTableConfig};
use sheetkit_core::protection::WorkbookProtectionConfig;
use sheetkit_core::validation::DataValidationConfig;

use crate::conversions::*;
use crate::stream::JsStreamWriter;
use crate::types::*;

/// Excel workbook for reading and writing .xlsx files.
#[napi]
pub struct Workbook {
    inner: sheetkit_core::workbook::Workbook,
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

#[napi]
impl Workbook {
    /// Create a new empty workbook with a single sheet named "Sheet1".
    #[napi(constructor)]
    pub fn new() -> Self {
        Self {
            inner: sheetkit_core::workbook::Workbook::new(),
        }
    }

    /// Open an existing .xlsx file from disk.
    #[napi(factory)]
    pub fn open(path: String) -> Result<Self> {
        let inner = sheetkit_core::workbook::Workbook::open(&path)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(Self { inner })
    }

    /// Save the workbook to a .xlsx file.
    #[napi]
    pub fn save(&self, path: String) -> Result<()> {
        self.inner
            .save(&path)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the names of all sheets in workbook order.
    #[napi(getter)]
    pub fn sheet_names(&self) -> Vec<String> {
        self.inner
            .sheet_names()
            .into_iter()
            .map(|s| s.to_string())
            .collect()
    }

    /// Get the value of a cell. Returns string, number, boolean, DateValue, or null.
    #[napi]
    pub fn get_cell_value(
        &self,
        sheet: String,
        cell: String,
    ) -> Result<Either5<Null, bool, f64, String, DateValue>> {
        let value = self
            .inner
            .get_cell_value(&sheet, &cell)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        cell_value_to_either(value)
    }

    /// Set the value of a cell. Pass string, number, boolean, DateValue, or null to clear.
    #[napi]
    pub fn set_cell_value(
        &mut self,
        sheet: String,
        cell: String,
        value: Either5<String, f64, bool, DateValue, Null>,
    ) -> Result<()> {
        let cell_value = match value {
            Either5::A(s) => CellValue::String(s),
            Either5::B(n) => CellValue::Number(n),
            Either5::C(b) => CellValue::Bool(b),
            Either5::D(d) => CellValue::Date(d.serial),
            Either5::E(_) => CellValue::Empty,
        };
        self.inner
            .set_cell_value(&sheet, &cell, cell_value)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Create a new empty sheet. Returns the 0-based sheet index.
    #[napi]
    pub fn new_sheet(&mut self, name: String) -> Result<u32> {
        self.inner
            .new_sheet(&name)
            .map(|i| i as u32)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Delete a sheet by name.
    #[napi]
    pub fn delete_sheet(&mut self, name: String) -> Result<()> {
        self.inner
            .delete_sheet(&name)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Rename a sheet.
    #[napi]
    pub fn set_sheet_name(&mut self, old_name: String, new_name: String) -> Result<()> {
        self.inner
            .set_sheet_name(&old_name, &new_name)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Copy a sheet. Returns the new sheet's 0-based index.
    #[napi]
    pub fn copy_sheet(&mut self, source: String, target: String) -> Result<u32> {
        self.inner
            .copy_sheet(&source, &target)
            .map(|i| i as u32)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the 0-based index of a sheet, or null if not found.
    #[napi]
    pub fn get_sheet_index(&self, name: String) -> Option<u32> {
        self.inner.get_sheet_index(&name).map(|i| i as u32)
    }

    /// Get the name of the active sheet.
    #[napi]
    pub fn get_active_sheet(&self) -> String {
        self.inner.get_active_sheet().to_string()
    }

    /// Set the active sheet by name.
    #[napi]
    pub fn set_active_sheet(&mut self, name: String) -> Result<()> {
        self.inner
            .set_active_sheet(&name)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Insert empty rows starting at the given 1-based row number.
    #[napi]
    pub fn insert_rows(&mut self, sheet: String, start_row: u32, count: u32) -> Result<()> {
        self.inner
            .insert_rows(&sheet, start_row, count)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Remove a row (1-based).
    #[napi]
    pub fn remove_row(&mut self, sheet: String, row: u32) -> Result<()> {
        self.inner
            .remove_row(&sheet, row)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Duplicate a row (1-based).
    #[napi]
    pub fn duplicate_row(&mut self, sheet: String, row: u32) -> Result<()> {
        self.inner
            .duplicate_row(&sheet, row)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set the height of a row (1-based).
    #[napi]
    pub fn set_row_height(&mut self, sheet: String, row: u32, height: f64) -> Result<()> {
        self.inner
            .set_row_height(&sheet, row, height)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the height of a row, or null if not explicitly set.
    #[napi]
    pub fn get_row_height(&self, sheet: String, row: u32) -> Result<Option<f64>> {
        self.inner
            .get_row_height(&sheet, row)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set whether a row is visible.
    #[napi]
    pub fn set_row_visible(&mut self, sheet: String, row: u32, visible: bool) -> Result<()> {
        self.inner
            .set_row_visible(&sheet, row, visible)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get whether a row is visible. Returns true if visible (not hidden).
    #[napi]
    pub fn get_row_visible(&self, sheet: String, row: u32) -> Result<bool> {
        self.inner
            .get_row_visible(&sheet, row)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set the outline level of a row (0-7).
    #[napi]
    pub fn set_row_outline_level(&mut self, sheet: String, row: u32, level: u8) -> Result<()> {
        self.inner
            .set_row_outline_level(&sheet, row, level)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the outline level of a row. Returns 0 if not set.
    #[napi]
    pub fn get_row_outline_level(&self, sheet: String, row: u32) -> Result<u8> {
        self.inner
            .get_row_outline_level(&sheet, row)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set the width of a column (e.g., "A", "B", "AA").
    #[napi]
    pub fn set_col_width(&mut self, sheet: String, col: String, width: f64) -> Result<()> {
        self.inner
            .set_col_width(&sheet, &col, width)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the width of a column, or null if not explicitly set.
    #[napi]
    pub fn get_col_width(&self, sheet: String, col: String) -> Result<Option<f64>> {
        self.inner
            .get_col_width(&sheet, &col)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set whether a column is visible.
    #[napi]
    pub fn set_col_visible(&mut self, sheet: String, col: String, visible: bool) -> Result<()> {
        self.inner
            .set_col_visible(&sheet, &col, visible)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get whether a column is visible. Returns true if visible (not hidden).
    #[napi]
    pub fn get_col_visible(&self, sheet: String, col: String) -> Result<bool> {
        self.inner
            .get_col_visible(&sheet, &col)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set the outline level of a column (0-7).
    #[napi]
    pub fn set_col_outline_level(&mut self, sheet: String, col: String, level: u8) -> Result<()> {
        self.inner
            .set_col_outline_level(&sheet, &col, level)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the outline level of a column. Returns 0 if not set.
    #[napi]
    pub fn get_col_outline_level(&self, sheet: String, col: String) -> Result<u8> {
        self.inner
            .get_col_outline_level(&sheet, &col)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Insert empty columns starting at the given column letter.
    #[napi]
    pub fn insert_cols(&mut self, sheet: String, col: String, count: u32) -> Result<()> {
        self.inner
            .insert_cols(&sheet, &col, count)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Remove a column by letter.
    #[napi]
    pub fn remove_col(&mut self, sheet: String, col: String) -> Result<()> {
        self.inner
            .remove_col(&sheet, &col)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Add a style definition. Returns the style ID for use with setCellStyle.
    #[napi]
    pub fn add_style(&mut self, style: JsStyle) -> Result<u32> {
        let core_style = js_style_to_core(&style);
        self.inner
            .add_style(&core_style)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the style ID applied to a cell, or null if default.
    #[napi]
    pub fn get_cell_style(&self, sheet: String, cell: String) -> Result<Option<u32>> {
        self.inner
            .get_cell_style(&sheet, &cell)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Apply a style ID to a cell.
    #[napi]
    pub fn set_cell_style(&mut self, sheet: String, cell: String, style_id: u32) -> Result<()> {
        self.inner
            .set_cell_style(&sheet, &cell, style_id)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Apply a style ID to an entire row.
    #[napi]
    pub fn set_row_style(&mut self, sheet: String, row: u32, style_id: u32) -> Result<()> {
        self.inner
            .set_row_style(&sheet, row, style_id)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the style ID for a row. Returns 0 if not set.
    #[napi]
    pub fn get_row_style(&self, sheet: String, row: u32) -> Result<u32> {
        self.inner
            .get_row_style(&sheet, row)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Apply a style ID to an entire column.
    #[napi]
    pub fn set_col_style(&mut self, sheet: String, col: String, style_id: u32) -> Result<()> {
        self.inner
            .set_col_style(&sheet, &col, style_id)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the style ID for a column. Returns 0 if not set.
    #[napi]
    pub fn get_col_style(&self, sheet: String, col: String) -> Result<u32> {
        self.inner
            .get_col_style(&sheet, &col)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Add a chart to a sheet.
    #[napi]
    pub fn add_chart(
        &mut self,
        sheet: String,
        from_cell: String,
        to_cell: String,
        config: JsChartConfig,
    ) -> Result<()> {
        let core_config = ChartConfig {
            chart_type: parse_chart_type(&config.chart_type),
            title: config.title,
            series: config
                .series
                .iter()
                .map(|s| ChartSeries {
                    name: s.name.clone(),
                    categories: s.categories.clone(),
                    values: s.values.clone(),
                    x_values: s.x_values.clone(),
                    bubble_sizes: s.bubble_sizes.clone(),
                })
                .collect(),
            show_legend: config.show_legend.unwrap_or(true),
            view_3d: config.view_3d.map(|v| View3DConfig {
                rot_x: v.rot_x,
                rot_y: v.rot_y,
                depth_percent: v.depth_percent,
                right_angle_axes: v.right_angle_axes,
                perspective: v.perspective,
            }),
        };
        self.inner
            .add_chart(&sheet, &from_cell, &to_cell, &core_config)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Add an image to a sheet.
    #[napi]
    pub fn add_image(&mut self, sheet: String, config: JsImageConfig) -> Result<()> {
        let core_config = ImageConfig {
            data: config.data.to_vec(),
            format: parse_image_format(&config.format)?,
            from_cell: config.from_cell,
            width_px: config.width_px,
            height_px: config.height_px,
        };
        self.inner
            .add_image(&sheet, &core_config)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Merge a range of cells on a sheet.
    #[napi]
    pub fn merge_cells(
        &mut self,
        sheet: String,
        top_left: String,
        bottom_right: String,
    ) -> Result<()> {
        self.inner
            .merge_cells(&sheet, &top_left, &bottom_right)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Remove a merged cell range from a sheet.
    #[napi]
    pub fn unmerge_cell(&mut self, sheet: String, reference: String) -> Result<()> {
        self.inner
            .unmerge_cell(&sheet, &reference)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get all merged cell ranges on a sheet.
    #[napi]
    pub fn get_merge_cells(&self, sheet: String) -> Result<Vec<String>> {
        self.inner
            .get_merge_cells(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Add a data validation rule to a sheet.
    #[napi]
    pub fn add_data_validation(
        &mut self,
        sheet: String,
        config: JsDataValidationConfig,
    ) -> Result<()> {
        let core_config = DataValidationConfig {
            sqref: config.sqref,
            validation_type: parse_validation_type(&config.validation_type),
            operator: config
                .operator
                .as_ref()
                .and_then(|s| parse_validation_operator(s)),
            formula1: config.formula1,
            formula2: config.formula2,
            allow_blank: config.allow_blank.unwrap_or(true),
            error_style: config
                .error_style
                .as_ref()
                .and_then(|s| parse_error_style(s)),
            error_title: config.error_title,
            error_message: config.error_message,
            prompt_title: config.prompt_title,
            prompt_message: config.prompt_message,
            show_input_message: config.show_input_message.unwrap_or(false),
            show_error_message: config.show_error_message.unwrap_or(false),
        };
        self.inner
            .add_data_validation(&sheet, &core_config)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get all data validations on a sheet.
    #[napi]
    pub fn get_data_validations(&self, sheet: String) -> Result<Vec<JsDataValidationConfig>> {
        let validations = self
            .inner
            .get_data_validations(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(validations.iter().map(core_validation_to_js).collect())
    }

    /// Remove a data validation by sqref.
    #[napi]
    pub fn remove_data_validation(&mut self, sheet: String, sqref: String) -> Result<()> {
        self.inner
            .remove_data_validation(&sheet, &sqref)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set conditional formatting rules on a cell range.
    #[napi]
    pub fn set_conditional_format(
        &mut self,
        sheet: String,
        sqref: String,
        rules: Vec<JsConditionalFormatRule>,
    ) -> Result<()> {
        let core_rules: Vec<ConditionalFormatRule> = rules
            .iter()
            .map(js_cf_rule_to_core)
            .collect::<Result<Vec<_>>>()?;
        self.inner
            .set_conditional_format(&sheet, &sqref, &core_rules)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get all conditional formatting rules for a sheet.
    #[napi]
    pub fn get_conditional_formats(&self, sheet: String) -> Result<Vec<JsConditionalFormatEntry>> {
        let formats = self
            .inner
            .get_conditional_formats(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(formats
            .iter()
            .map(|(sqref, rules)| JsConditionalFormatEntry {
                sqref: sqref.clone(),
                rules: rules.iter().map(core_cf_rule_to_js).collect(),
            })
            .collect())
    }

    /// Delete conditional formatting for a specific cell range.
    #[napi]
    pub fn delete_conditional_format(&mut self, sheet: String, sqref: String) -> Result<()> {
        self.inner
            .delete_conditional_format(&sheet, &sqref)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Add a comment to a cell.
    #[napi]
    pub fn add_comment(&mut self, sheet: String, config: JsCommentConfig) -> Result<()> {
        let core_config = CommentConfig {
            cell: config.cell,
            author: config.author,
            text: config.text,
        };
        self.inner
            .add_comment(&sheet, &core_config)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get all comments on a sheet.
    #[napi]
    pub fn get_comments(&self, sheet: String) -> Result<Vec<JsCommentConfig>> {
        let comments = self
            .inner
            .get_comments(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(comments
            .iter()
            .map(|c| JsCommentConfig {
                cell: c.cell.clone(),
                author: c.author.clone(),
                text: c.text.clone(),
            })
            .collect())
    }

    /// Remove a comment from a cell.
    #[napi]
    pub fn remove_comment(&mut self, sheet: String, cell: String) -> Result<()> {
        self.inner
            .remove_comment(&sheet, &cell)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set an auto-filter on a sheet.
    #[napi]
    pub fn set_auto_filter(&mut self, sheet: String, range: String) -> Result<()> {
        self.inner
            .set_auto_filter(&sheet, &range)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Remove the auto-filter from a sheet.
    #[napi]
    pub fn remove_auto_filter(&mut self, sheet: String) -> Result<()> {
        self.inner
            .remove_auto_filter(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Create a new stream writer for a new sheet.
    #[napi]
    pub fn new_stream_writer(&self, sheet_name: String) -> Result<JsStreamWriter> {
        let writer = self
            .inner
            .new_stream_writer(&sheet_name)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(JsStreamWriter {
            inner: Some(writer),
        })
    }

    /// Apply a stream writer's output to the workbook. Returns the sheet index.
    #[napi]
    pub fn apply_stream_writer(&mut self, writer: &mut JsStreamWriter) -> Result<u32> {
        let inner_writer = writer
            .inner
            .take()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        let index = self
            .inner
            .apply_stream_writer(inner_writer)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(index as u32)
    }

    /// Set core document properties (title, creator, etc.).
    #[napi]
    pub fn set_doc_props(&mut self, props: JsDocProperties) {
        self.inner.set_doc_props(js_doc_props_to_core(&props));
    }

    /// Get core document properties.
    #[napi]
    pub fn get_doc_props(&self) -> JsDocProperties {
        core_doc_props_to_js(&self.inner.get_doc_props())
    }

    /// Set application properties (company, app version, etc.).
    #[napi]
    pub fn set_app_props(&mut self, props: JsAppProperties) {
        self.inner.set_app_props(js_app_props_to_core(&props));
    }

    /// Get application properties.
    #[napi]
    pub fn get_app_props(&self) -> JsAppProperties {
        core_app_props_to_js(&self.inner.get_app_props())
    }

    /// Set a custom property. Value can be string, number, or boolean.
    #[napi]
    pub fn set_custom_property(
        &mut self,
        name: String,
        value: Either3<String, f64, bool>,
    ) -> Result<()> {
        let prop_value = match value {
            Either3::A(s) => CustomPropertyValue::String(s),
            Either3::B(n) => {
                if n.fract() == 0.0 && n >= i32::MIN as f64 && n <= i32::MAX as f64 {
                    CustomPropertyValue::Int(n as i32)
                } else {
                    CustomPropertyValue::Float(n)
                }
            }
            Either3::C(b) => CustomPropertyValue::Bool(b),
        };
        self.inner.set_custom_property(&name, prop_value);
        Ok(())
    }

    /// Get a custom property value, or null if not found.
    #[napi]
    pub fn get_custom_property(&self, name: String) -> Option<Either3<String, f64, bool>> {
        match self.inner.get_custom_property(&name) {
            Some(CustomPropertyValue::String(s)) => Some(Either3::A(s)),
            Some(CustomPropertyValue::Int(i)) => Some(Either3::B(i as f64)),
            Some(CustomPropertyValue::Float(f)) => Some(Either3::B(f)),
            Some(CustomPropertyValue::Bool(b)) => Some(Either3::C(b)),
            Some(CustomPropertyValue::DateTime(s)) => Some(Either3::A(s)),
            None => None,
        }
    }

    /// Delete a custom property. Returns true if it existed.
    #[napi]
    pub fn delete_custom_property(&mut self, name: String) -> bool {
        self.inner.delete_custom_property(&name)
    }

    /// Protect the workbook structure/windows with optional password.
    #[napi]
    pub fn protect_workbook(&mut self, config: JsWorkbookProtectionConfig) {
        self.inner.protect_workbook(WorkbookProtectionConfig {
            password: config.password,
            lock_structure: config.lock_structure.unwrap_or(false),
            lock_windows: config.lock_windows.unwrap_or(false),
            lock_revision: config.lock_revision.unwrap_or(false),
        });
    }

    /// Remove workbook protection.
    #[napi]
    pub fn unprotect_workbook(&mut self) {
        self.inner.unprotect_workbook();
    }

    /// Check if the workbook is protected.
    #[napi]
    pub fn is_workbook_protected(&self) -> bool {
        self.inner.is_workbook_protected()
    }

    /// Set freeze panes on a sheet.
    /// The cell reference indicates the top-left cell of the scrollable area.
    /// For example, "A2" freezes row 1, "B1" freezes column A.
    #[napi]
    pub fn set_panes(&mut self, sheet: String, cell: String) -> Result<()> {
        self.inner
            .set_panes(&sheet, &cell)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Remove any freeze or split panes from a sheet.
    #[napi]
    pub fn unset_panes(&mut self, sheet: String) -> Result<()> {
        self.inner
            .unset_panes(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the current freeze pane cell reference for a sheet, or null if none.
    #[napi]
    pub fn get_panes(&self, sheet: String) -> Result<Option<String>> {
        self.inner
            .get_panes(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set page margins on a sheet (values in inches).
    #[napi]
    pub fn set_page_margins(&mut self, sheet: String, margins: JsPageMargins) -> Result<()> {
        let config = PageMarginsConfig {
            left: margins.left,
            right: margins.right,
            top: margins.top,
            bottom: margins.bottom,
            header: margins.header,
            footer: margins.footer,
        };
        self.inner
            .set_page_margins(&sheet, &config)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get page margins for a sheet. Returns defaults if not explicitly set.
    #[napi]
    pub fn get_page_margins(&self, sheet: String) -> Result<JsPageMargins> {
        let m = self
            .inner
            .get_page_margins(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(JsPageMargins {
            left: m.left,
            right: m.right,
            top: m.top,
            bottom: m.bottom,
            header: m.header,
            footer: m.footer,
        })
    }

    /// Set page setup options (paper size, orientation, scale, fit-to-page).
    #[napi]
    pub fn set_page_setup(&mut self, sheet: String, setup: JsPageSetup) -> Result<()> {
        let orientation = setup
            .orientation
            .as_ref()
            .and_then(|s| parse_orientation(s));
        let paper_size = setup.paper_size.as_ref().and_then(|s| parse_paper_size(s));
        self.inner
            .set_page_setup(
                &sheet,
                orientation,
                paper_size,
                setup.scale,
                setup.fit_to_width,
                setup.fit_to_height,
            )
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the page setup for a sheet.
    #[napi]
    pub fn get_page_setup(&self, sheet: String) -> Result<JsPageSetup> {
        let orientation = self
            .inner
            .get_orientation(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        let paper_size = self
            .inner
            .get_paper_size(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        let (scale, fit_to_width, fit_to_height) = self
            .inner
            .get_page_setup_details(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(JsPageSetup {
            paper_size: paper_size.as_ref().map(paper_size_to_string),
            orientation: orientation.as_ref().map(orientation_to_string),
            scale,
            fit_to_width,
            fit_to_height,
        })
    }

    /// Set header and footer text for printing.
    #[napi]
    pub fn set_header_footer(
        &mut self,
        sheet: String,
        header: Option<String>,
        footer: Option<String>,
    ) -> Result<()> {
        self.inner
            .set_header_footer(&sheet, header.as_deref(), footer.as_deref())
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get the header and footer text for a sheet.
    /// Returns an object with `header` and `footer` fields, each possibly null.
    #[napi]
    pub fn get_header_footer(&self, sheet: String) -> Result<JsHeaderFooter> {
        let (header, footer) = self
            .inner
            .get_header_footer(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(JsHeaderFooter { header, footer })
    }

    /// Set print options on a sheet.
    #[napi]
    pub fn set_print_options(&mut self, sheet: String, opts: JsPrintOptions) -> Result<()> {
        self.inner
            .set_print_options(
                &sheet,
                opts.grid_lines,
                opts.headings,
                opts.horizontal_centered,
                opts.vertical_centered,
            )
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get print options for a sheet.
    #[napi]
    pub fn get_print_options(&self, sheet: String) -> Result<JsPrintOptions> {
        let (gl, hd, hc, vc) = self
            .inner
            .get_print_options(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(JsPrintOptions {
            grid_lines: gl,
            headings: hd,
            horizontal_centered: hc,
            vertical_centered: vc,
        })
    }

    /// Insert a horizontal page break before the given 1-based row.
    #[napi]
    pub fn insert_page_break(&mut self, sheet: String, row: u32) -> Result<()> {
        self.inner
            .insert_page_break(&sheet, row)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Remove a horizontal page break at the given 1-based row.
    #[napi]
    pub fn remove_page_break(&mut self, sheet: String, row: u32) -> Result<()> {
        self.inner
            .remove_page_break(&sheet, row)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get all row page break positions (1-based row numbers).
    #[napi]
    pub fn get_page_breaks(&self, sheet: String) -> Result<Vec<u32>> {
        self.inner
            .get_page_breaks(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set a hyperlink on a cell.
    #[napi]
    pub fn set_cell_hyperlink(
        &mut self,
        sheet: String,
        cell: String,
        opts: JsHyperlinkOptions,
    ) -> Result<()> {
        let link = parse_hyperlink_type(&opts)?;
        self.inner
            .set_cell_hyperlink(
                &sheet,
                &cell,
                link,
                opts.display.as_deref(),
                opts.tooltip.as_deref(),
            )
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get hyperlink information for a cell, or null if no hyperlink exists.
    #[napi]
    pub fn get_cell_hyperlink(
        &self,
        sheet: String,
        cell: String,
    ) -> Result<Option<JsHyperlinkInfo>> {
        let info = self
            .inner
            .get_cell_hyperlink(&sheet, &cell)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(info.as_ref().map(hyperlink_info_to_js))
    }

    /// Delete a hyperlink from a cell.
    #[napi]
    pub fn delete_cell_hyperlink(&mut self, sheet: String, cell: String) -> Result<()> {
        self.inner
            .delete_cell_hyperlink(&sheet, &cell)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get all rows with their data from a sheet.
    /// Only rows that contain at least one cell are included.
    #[napi]
    pub fn get_rows(&self, sheet: String) -> Result<Vec<JsRowData>> {
        let rows = self
            .inner
            .get_rows(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;

        Ok(rows
            .into_iter()
            .map(|(row_num, cells)| JsRowData {
                row: row_num,
                cells: cells
                    .into_iter()
                    .map(|(col, val)| cell_value_to_row_cell(col, val))
                    .collect(),
            })
            .collect())
    }

    /// Get all columns with their data from a sheet.
    /// Only columns that have data are included.
    #[napi]
    pub fn get_cols(&self, sheet: String) -> Result<Vec<JsColData>> {
        let cols = self
            .inner
            .get_cols(&sheet)
            .map_err(|e| Error::from_reason(e.to_string()))?;

        Ok(cols
            .into_iter()
            .map(|(col_name, cells)| JsColData {
                column: col_name,
                cells: cells
                    .into_iter()
                    .map(|(row, val)| cell_value_to_col_cell(row, val))
                    .collect(),
            })
            .collect())
    }

    /// Evaluate a formula string against the current workbook data.
    #[napi]
    pub fn evaluate_formula(
        &self,
        sheet: String,
        formula: String,
    ) -> Result<Either5<Null, bool, f64, String, DateValue>> {
        let result = self
            .inner
            .evaluate_formula(&sheet, &formula)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        cell_value_to_either(result)
    }

    /// Recalculate all formula cells in the workbook.
    #[napi]
    pub fn calculate_all(&mut self) -> Result<()> {
        self.inner
            .calculate_all()
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Add a pivot table to the workbook.
    #[napi]
    pub fn add_pivot_table(&mut self, config: JsPivotTableConfig) -> Result<()> {
        let core_config = PivotTableConfig {
            name: config.name,
            source_sheet: config.source_sheet,
            source_range: config.source_range,
            target_sheet: config.target_sheet,
            target_cell: config.target_cell,
            rows: config
                .rows
                .iter()
                .map(|f| PivotField {
                    name: f.name.clone(),
                })
                .collect(),
            columns: config
                .columns
                .iter()
                .map(|f| PivotField {
                    name: f.name.clone(),
                })
                .collect(),
            data: config
                .data
                .iter()
                .map(|f| PivotDataField {
                    name: f.name.clone(),
                    function: parse_aggregate_function(&f.function),
                    display_name: f.display_name.clone(),
                })
                .collect(),
        };
        self.inner
            .add_pivot_table(&core_config)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get all pivot tables in the workbook.
    #[napi]
    pub fn get_pivot_tables(&self) -> Vec<JsPivotTableInfo> {
        self.inner
            .get_pivot_tables()
            .into_iter()
            .map(|info| JsPivotTableInfo {
                name: info.name,
                source_sheet: info.source_sheet,
                source_range: info.source_range,
                target_sheet: info.target_sheet,
                location: info.location,
            })
            .collect()
    }

    /// Delete a pivot table by name.
    #[napi]
    pub fn delete_pivot_table(&mut self, name: String) -> Result<()> {
        self.inner
            .delete_pivot_table(&name)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Set a cell to a rich text value with multiple formatted runs.
    #[napi]
    pub fn set_cell_rich_text(
        &mut self,
        sheet: String,
        cell: String,
        runs: Vec<JsRichTextRun>,
    ) -> Result<()> {
        let core_runs: Vec<sheetkit_core::rich_text::RichTextRun> = runs
            .iter()
            .map(|r| sheetkit_core::rich_text::RichTextRun {
                text: r.text.clone(),
                font: r.font.clone(),
                size: r.size,
                bold: r.bold.unwrap_or(false),
                italic: r.italic.unwrap_or(false),
                color: r.color.clone(),
            })
            .collect();
        self.inner
            .set_cell_rich_text(&sheet, &cell, core_runs)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    /// Get rich text runs for a cell, or null if not rich text.
    #[napi]
    pub fn get_cell_rich_text(
        &self,
        sheet: String,
        cell: String,
    ) -> Result<Option<Vec<JsRichTextRun>>> {
        let runs = self
            .inner
            .get_cell_rich_text(&sheet, &cell)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        Ok(runs.map(|rs| {
            rs.iter()
                .map(|r| JsRichTextRun {
                    text: r.text.clone(),
                    font: r.font.clone(),
                    size: r.size,
                    bold: if r.bold { Some(true) } else { None },
                    italic: if r.italic { Some(true) } else { None },
                    color: r.color.clone(),
                })
                .collect()
        }))
    }

    /// Resolve a theme color by index (0-11) with optional tint.
    /// Returns the ARGB hex string (e.g. "FF4472C4") or null if out of range.
    #[napi]
    pub fn get_theme_color(&self, index: u32, tint: Option<f64>) -> Option<String> {
        self.inner.get_theme_color(index, tint)
    }
}
