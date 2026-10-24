#![deny(clippy::all)]

mod conversions;
mod error;
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
use sheetkit_core::table::{TableColumn, TableConfig};
use sheetkit_core::threaded_comment::{PersonInput, ThreadedCommentInput};
use sheetkit_core::validation::DataValidationConfig;
use sheetkit_core::workbook::WorkbookFormat;

use crate::conversions::*;
use crate::error::{
    map_core_error, OpenBufferPasswordTask, OpenBufferTask, OpenPasswordTask, OpenPathTask,
    SavePasswordTask, SavePathTask, WriteBufferTask,
};
use crate::stream::JsStreamWriter;
use crate::types::*;

fn outer_input_error(
    operation: sheetkit_core::ResourceOperation,
    category: sheetkit_core::ResourceCategory,
    limit: u64,
    observed: u64,
    stage: sheetkit_core::ResourceStage,
) -> sheetkit_core::error::Error {
    sheetkit_core::ResourceLimitError::exceeded(
        operation,
        category,
        sheetkit_core::ResourceDimension::InputBytes,
        limit,
        observed,
        Some(stage),
    )
    .map(sheetkit_core::error::Error::from)
    .unwrap_or_else(|| {
        sheetkit_core::error::Error::Internal("invalid outer input limit state".to_string())
    })
}

fn check_buffer_outer_input(
    options: &sheetkit_core::workbook::OpenOptions,
    observed: u64,
    encrypted: bool,
) -> std::result::Result<(), sheetkit_core::error::Error> {
    let open_ceiling = sheetkit_core::workbook::OpenLimits::hard_ceiling().outer_input_bytes;
    let open_limit = options.limits.outer_input_bytes;
    if open_limit == 0 || open_limit > open_ceiling {
        return Err(outer_input_error(
            sheetkit_core::ResourceOperation::WorkbookOpen,
            sheetkit_core::ResourceCategory::Archive,
            open_ceiling,
            open_limit.max(open_ceiling.saturating_add(1)),
            sheetkit_core::ResourceStage::Validate,
        ));
    }
    if observed > open_limit {
        return Err(outer_input_error(
            sheetkit_core::ResourceOperation::WorkbookOpen,
            sheetkit_core::ResourceCategory::Archive,
            open_limit,
            observed,
            sheetkit_core::ResourceStage::Inspect,
        ));
    }

    if encrypted {
        let crypt_ceiling = sheetkit_core::crypt::CryptLimits::hard_ceiling().outer_cfb_bytes;
        let crypt_limit = options.crypt_limits.outer_cfb_bytes;
        if crypt_limit == 0 || crypt_limit > crypt_ceiling {
            return Err(outer_input_error(
                sheetkit_core::ResourceOperation::Decrypt,
                sheetkit_core::ResourceCategory::Container,
                crypt_ceiling,
                crypt_limit.max(crypt_ceiling.saturating_add(1)),
                sheetkit_core::ResourceStage::Validate,
            ));
        }
        if observed > crypt_limit {
            return Err(outer_input_error(
                sheetkit_core::ResourceOperation::Decrypt,
                sheetkit_core::ResourceCategory::Container,
                crypt_limit,
                observed,
                sheetkit_core::ResourceStage::Inspect,
            ));
        }
    }
    Ok(())
}

fn precharge_pivot_config(
    config: &JsPivotTableConfig,
    limits: sheetkit_core::traversal::TraversalLimits,
) -> std::result::Result<(), sheetkit_core::error::Error> {
    let mut used = 0u64;
    let mut charge = |bytes: usize| {
        let remaining = limits.max_output_bytes().saturating_sub(used);
        let bytes = u64::try_from(bytes).unwrap_or(u64::MAX);
        let admitted = bytes.min(remaining.saturating_add(1));
        used = sheetkit_core::traversal::checked_output_add(used, admitted, limits)?;
        Ok::<(), sheetkit_core::error::Error>(())
    };

    charge(std::mem::size_of::<PivotTableConfig>())?;
    for value in [
        &config.name,
        &config.source_sheet,
        &config.source_range,
        &config.target_sheet,
        &config.target_cell,
    ] {
        charge(value.len())?;
    }

    charge(
        config
            .rows
            .len()
            .saturating_mul(std::mem::size_of::<PivotField>()),
    )?;
    for field in &config.rows {
        charge(field.name.len())?;
    }

    charge(
        config
            .columns
            .len()
            .saturating_mul(std::mem::size_of::<PivotField>()),
    )?;
    for field in &config.columns {
        charge(field.name.len())?;
    }

    charge(
        config
            .data
            .len()
            .saturating_mul(std::mem::size_of::<PivotDataField>()),
    )?;
    for field in &config.data {
        charge(field.name.len())?;
        charge(field.function.len())?;
        if let Some(display_name) = &field.display_name {
            charge(display_name.len())?;
        }
    }
    Ok(())
}

fn validate_stream_batch_size(
    batch_size: Option<f64>,
) -> std::result::Result<usize, sheetkit_core::error::Error> {
    const MAX_JS_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;

    let batch_size = batch_size.unwrap_or(1000.0);
    if batch_size == 0.0 {
        return Err(sheetkit_core::error::Error::InvalidArgument(
            "stream batch size must be greater than zero".to_string(),
        ));
    }
    if !batch_size.is_finite()
        || batch_size < 0.0
        || batch_size.fract() != 0.0
        || batch_size > MAX_JS_SAFE_INTEGER
    {
        return Err(sheetkit_core::error::Error::InvalidArgument(
            "stream batch size must be a positive safe integer".to_string(),
        ));
    }
    let batch_size = u64::try_from(batch_size as u128)
        .ok()
        .and_then(|value| usize::try_from(value).ok())
        .ok_or_else(|| {
            sheetkit_core::error::Error::InvalidArgument(
                "stream batch size is not representable on this platform".to_string(),
            )
        })?;
    sheetkit_core::stream_reader::validate_stream_batch_size(batch_size)
}

/// Excel workbook for reading and writing .xlsx files.
#[napi]
pub struct Workbook {
    inner: std::sync::Arc<sheetkit_core::workbook::Workbook>,
}

impl Workbook {
    fn from_core(inner: sheetkit_core::workbook::Workbook) -> Self {
        Self {
            inner: std::sync::Arc::new(inner),
        }
    }

    fn inner_mut(&mut self) -> Result<&mut sheetkit_core::workbook::Workbook> {
        std::sync::Arc::get_mut(&mut self.inner)
            .ok_or_else(|| Error::from_reason("workbook has an async save operation in progress"))
    }
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
        Self::from_core(sheetkit_core::workbook::Workbook::new())
    }

    /// Open an existing .xlsx file from disk.
    #[napi(factory, js_name = "openSync")]
    pub fn open_sync(env: Env, path: String, options: Option<JsOpenOptions>) -> Result<Self> {
        let opts = js_open_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let inner = sheetkit_core::workbook::Workbook::open_with_options(&path, &opts)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(Self::from_core(inner))
    }

    /// Open an existing .xlsx file from disk asynchronously.
    #[allow(private_interfaces)]
    #[napi]
    pub fn open(
        env: Env,
        path: String,
        options: Option<JsOpenOptions>,
    ) -> Result<AsyncTask<OpenPathTask>> {
        let opts = js_open_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        Ok(OpenPathTask::new(path, opts).into_async_task())
    }

    /// Save the workbook to a .xlsx file.
    #[napi(js_name = "saveSync")]
    pub fn save_sync(
        &self,
        env: Env,
        path: String,
        limits: Option<JsSaveLimits>,
        options: Option<JsSaveOptions>,
    ) -> Result<()> {
        let limits =
            js_save_limits_to_core(limits.as_ref()).map_err(|error| map_core_error(&env, error))?;
        let options = js_save_options_to_core(options.as_ref());
        self.inner
            .save_with_options(&path, limits, options)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Save the workbook to a .xlsx file using destination-preserving atomic replacement.
    #[allow(private_interfaces)]
    #[napi]
    pub fn save(
        &self,
        env: Env,
        path: String,
        limits: Option<JsSaveLimits>,
        options: Option<JsSaveOptions>,
    ) -> Result<AsyncTask<SavePathTask>> {
        let limits =
            js_save_limits_to_core(limits.as_ref()).map_err(|error| map_core_error(&env, error))?;
        let options = js_save_options_to_core(options.as_ref());
        Ok(SavePathTask::new(self.inner.clone(), path, limits, options).into_async_task())
    }

    /// Open a workbook from an in-memory Buffer.
    #[napi(factory, js_name = "openBufferSync")]
    pub fn open_buffer_sync(
        env: Env,
        data: Buffer,
        options: Option<JsOpenOptions>,
    ) -> Result<Self> {
        let opts = js_open_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let inner = sheetkit_core::workbook::Workbook::open_from_buffer_with_options(&data, &opts)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(Self::from_core(inner))
    }

    /// Open a workbook from an in-memory Buffer asynchronously.
    #[allow(private_interfaces)]
    #[napi]
    pub fn open_buffer(
        env: Env,
        data: Buffer,
        options: Option<JsOpenOptions>,
    ) -> Result<AsyncTask<OpenBufferTask>> {
        let opts = js_open_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        check_buffer_outer_input(&opts, u64::try_from(data.len()).unwrap_or(u64::MAX), false)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(OpenBufferTask::new(data.to_vec(), opts).into_async_task())
    }

    /// Open an encrypted .xlsx file using a password.
    #[napi(factory, js_name = "openWithPasswordSync")]
    pub fn open_with_password_sync(
        env: Env,
        path: String,
        password: String,
        options: Option<JsOpenOptions>,
    ) -> Result<Self> {
        let options = js_open_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let inner = sheetkit_core::workbook::Workbook::open_with_password_with_options(
            &path, &password, &options,
        )
        .map_err(|error| map_core_error(&env, error))?;
        Ok(Self::from_core(inner))
    }

    /// Open an encrypted .xlsx file using a password asynchronously.
    #[allow(private_interfaces)]
    #[napi]
    pub fn open_with_password(
        env: Env,
        path: String,
        password: String,
        options: Option<JsOpenOptions>,
    ) -> Result<AsyncTask<OpenPasswordTask>> {
        let options = js_open_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        Ok(OpenPasswordTask::new(path, password, options).into_async_task())
    }

    /// Open an encrypted workbook from an in-memory Buffer.
    #[napi(factory, js_name = "openBufferWithPasswordSync")]
    pub fn open_buffer_with_password_sync(
        env: Env,
        data: Buffer,
        password: String,
        options: Option<JsOpenOptions>,
    ) -> Result<Self> {
        let options = js_open_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let inner = sheetkit_core::workbook::Workbook::open_from_buffer_with_password_with_options(
            &data, &password, &options,
        )
        .map_err(|error| map_core_error(&env, error))?;
        Ok(Self::from_core(inner))
    }

    /// Open an encrypted workbook from an in-memory Buffer asynchronously.
    #[allow(private_interfaces)]
    #[napi]
    pub fn open_buffer_with_password(
        env: Env,
        data: Buffer,
        password: String,
        options: Option<JsOpenOptions>,
    ) -> Result<AsyncTask<OpenBufferPasswordTask>> {
        let options = js_open_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        check_buffer_outer_input(
            &options,
            u64::try_from(data.len()).unwrap_or(u64::MAX),
            true,
        )
        .map_err(|error| map_core_error(&env, error))?;
        Ok(OpenBufferPasswordTask::new(data.to_vec(), password, options).into_async_task())
    }

    /// Serialize the workbook to an in-memory Buffer.
    #[napi(js_name = "writeBufferSync")]
    pub fn write_buffer_sync(&self, env: Env, limits: Option<JsSaveLimits>) -> Result<Buffer> {
        let limits =
            js_save_limits_to_core(limits.as_ref()).map_err(|error| map_core_error(&env, error))?;
        let buf = self
            .inner
            .save_to_buffer_with_limits(limits)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(Buffer::from(buf))
    }

    /// Serialize the workbook to an in-memory Buffer asynchronously.
    #[allow(private_interfaces)]
    #[napi]
    pub fn write_buffer(
        &self,
        env: Env,
        limits: Option<JsSaveLimits>,
    ) -> Result<AsyncTask<WriteBufferTask>> {
        let limits =
            js_save_limits_to_core(limits.as_ref()).map_err(|error| map_core_error(&env, error))?;
        Ok(WriteBufferTask::new(self.inner.clone(), limits).into_async_task())
    }

    /// Save the workbook as an encrypted .xlsx file.
    #[napi(js_name = "saveWithPasswordSync")]
    pub fn save_with_password_sync(
        &self,
        env: Env,
        path: String,
        password: String,
        limits: Option<JsCryptLimits>,
        options: Option<JsSaveOptions>,
    ) -> Result<()> {
        let limits = js_crypt_limits_to_core(limits.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let options = js_save_options_to_core(options.as_ref());
        self.inner
            .save_with_password_with_options(
                &path,
                &password,
                limits,
                sheetkit_core::save::SaveLimits::default(),
                options,
            )
            .map_err(|error| map_core_error(&env, error))
    }

    /// Save an encrypted workbook for the TypeScript async rejection bridge.
    #[allow(private_interfaces)]
    #[napi]
    pub fn save_with_password(
        &self,
        env: Env,
        path: String,
        password: String,
        limits: Option<JsCryptLimits>,
        options: Option<JsSaveOptions>,
    ) -> Result<AsyncTask<SavePasswordTask>> {
        let limits = js_crypt_limits_to_core(limits.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let options = js_save_options_to_core(options.as_ref());
        Ok(
            SavePasswordTask::new(self.inner.clone(), path, password, limits, options)
                .into_async_task(),
        )
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

    /// Get the value of a cell. Returns string, number, boolean, DateValue,
    /// FormulaValue (expression and cached result), or null.
    #[napi]
    pub fn get_cell_value(
        &self,
        env: Env,
        sheet: String,
        cell: String,
    ) -> Result<Either6<Null, bool, f64, String, DateValue, FormulaValue>> {
        let value = self
            .inner
            .get_cell_value(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))?;
        cell_value_to_read_value(value)
    }

    /// Get the formatted display text for a cell, applying its number format.
    /// Returns the value formatted according to the cell's style.
    #[napi]
    pub fn get_cell_formatted_value(
        &self,
        env: Env,
        sheet: String,
        cell: String,
    ) -> Result<String> {
        self.inner
            .get_cell_formatted_value(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Format a numeric value using a format code string.
    /// This is a standalone utility that does not require cell context.
    #[napi]
    pub fn format_number(value: f64, format_code: String) -> String {
        sheetkit_core::numfmt::format_number(value, &format_code)
    }

    /// Get the format code string for a built-in number format ID (0-49).
    /// Returns null if the ID is not recognized.
    #[napi]
    pub fn builtin_format_code(id: f64) -> Result<Option<String>> {
        let id = checked_u32("id", id).map_err(invalid_arg)?;
        Ok(sheetkit_core::numfmt::builtin_format_code(id).map(|s| s.to_string()))
    }

    /// Set the value of a cell. Pass string, number, boolean, DateValue, or null to clear.
    #[napi]
    pub fn set_cell_value(
        &mut self,
        env: Env,
        sheet: String,
        cell: String,
        value: Either6<String, f64, bool, DateValue, FormulaValue, Null>,
    ) -> Result<()> {
        let cell_value = js_value_to_cell_value(value)?;
        self.inner_mut()?
            .set_cell_value(&sheet, &cell, cell_value)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set multiple cell values at once. More efficient than calling
    /// setCellValue repeatedly because it crosses the FFI boundary only once.
    #[napi]
    pub fn set_cell_values(
        &mut self,
        env: Env,
        sheet: String,
        cells: Vec<JsCellEntry>,
    ) -> Result<()> {
        let entries: Vec<(String, CellValue)> = cells
            .into_iter()
            .map(|entry| Ok((entry.cell, js_value_to_cell_value(entry.value)?)))
            .collect::<Result<_>>()?;
        self.inner_mut()?
            .set_cell_values(&sheet, entries)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set values in a single row starting from the given column.
    /// Values are placed left-to-right starting at startCol (e.g., "A").
    #[napi]
    #[allow(clippy::type_complexity)]
    pub fn set_row_values(
        &mut self,
        env: Env,
        sheet: String,
        row: f64,
        start_col: String,
        values: Vec<Either6<String, f64, bool, DateValue, FormulaValue, Null>>,
    ) -> Result<()> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        let col_num = crate::conversions::parse_column_name(&start_col)?;
        let cell_values: Vec<CellValue> = values
            .into_iter()
            .map(js_value_to_cell_value)
            .collect::<Result<_>>()?;
        self.inner_mut()?
            .set_row_values(&sheet, row, col_num, cell_values)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set a block of cell values from a 2D array.
    /// Each inner array is a row, each element is a cell value.
    /// Optionally specify a start cell (default "A1").
    #[napi]
    #[allow(clippy::type_complexity)]
    pub fn set_sheet_data(
        &mut self,
        env: Env,
        sheet: String,
        data: Vec<Vec<Either6<String, f64, bool, DateValue, FormulaValue, Null>>>,
        start_cell: Option<String>,
    ) -> Result<()> {
        let (start_col, start_row) = if let Some(ref cell) = start_cell {
            sheetkit_core::utils::cell_ref::cell_name_to_coordinates(cell)
                .map_err(|error| map_core_error(&env, error))?
        } else {
            (1, 1)
        };
        let cell_data: Vec<Vec<CellValue>> = data
            .into_iter()
            .map(|row| row.into_iter().map(js_value_to_cell_value).collect())
            .collect::<Result<_>>()?;
        self.inner_mut()?
            .set_sheet_data(&sheet, cell_data, start_row, start_col)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Create a new empty sheet. Returns the 0-based sheet index.
    #[napi]
    pub fn new_sheet(&mut self, env: Env, name: String) -> Result<u32> {
        self.inner_mut()?
            .new_sheet(&name)
            .map(|i| i as u32)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Delete a sheet by name.
    #[napi]
    pub fn delete_sheet(&mut self, env: Env, name: String) -> Result<()> {
        self.inner_mut()?
            .delete_sheet(&name)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Rename a sheet.
    #[napi]
    pub fn set_sheet_name(&mut self, env: Env, old_name: String, new_name: String) -> Result<()> {
        self.inner_mut()?
            .set_sheet_name(&old_name, &new_name)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Copy a sheet. Returns the new sheet's 0-based index.
    #[napi]
    pub fn copy_sheet(&mut self, env: Env, source: String, target: String) -> Result<u32> {
        self.inner_mut()?
            .copy_sheet(&source, &target)
            .map(|i| i as u32)
            .map_err(|error| map_core_error(&env, error))
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
    pub fn set_active_sheet(&mut self, env: Env, name: String) -> Result<()> {
        self.inner_mut()?
            .set_active_sheet(&name)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Insert empty rows starting at the given 1-based row number.
    #[napi]
    pub fn insert_rows(
        &mut self,
        env: Env,
        sheet: String,
        start_row: f64,
        count: f64,
    ) -> Result<()> {
        let start_row = checked_u32("startRow", start_row).map_err(invalid_arg)?;
        let count = checked_u32("count", count).map_err(invalid_arg)?;
        self.inner_mut()?
            .insert_rows(&sheet, start_row, count)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Remove a row (1-based).
    #[napi]
    pub fn remove_row(&mut self, env: Env, sheet: String, row: f64) -> Result<()> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        self.inner_mut()?
            .remove_row(&sheet, row)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Duplicate a row (1-based).
    #[napi]
    pub fn duplicate_row(&mut self, env: Env, sheet: String, row: f64) -> Result<()> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        self.inner_mut()?
            .duplicate_row(&sheet, row)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set the height of a row (1-based).
    #[napi]
    pub fn set_row_height(&mut self, env: Env, sheet: String, row: f64, height: f64) -> Result<()> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        self.inner_mut()?
            .set_row_height(&sheet, row, height)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the height of a row, or null if not explicitly set.
    #[napi]
    pub fn get_row_height(&self, env: Env, sheet: String, row: f64) -> Result<Option<f64>> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        self.inner
            .get_row_height(&sheet, row)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set whether a row is visible.
    #[napi]
    pub fn set_row_visible(
        &mut self,
        env: Env,
        sheet: String,
        row: f64,
        visible: bool,
    ) -> Result<()> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        self.inner_mut()?
            .set_row_visible(&sheet, row, visible)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get whether a row is visible. Returns true if visible (not hidden).
    #[napi]
    pub fn get_row_visible(&self, env: Env, sheet: String, row: f64) -> Result<bool> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        self.inner
            .get_row_visible(&sheet, row)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set the outline level of a row (0-7).
    #[napi]
    pub fn set_row_outline_level(
        &mut self,
        env: Env,
        sheet: String,
        row: f64,
        level: f64,
    ) -> Result<()> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        let level = checked_u8("level", level).map_err(invalid_arg)?;
        self.inner_mut()?
            .set_row_outline_level(&sheet, row, level)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the outline level of a row. Returns 0 if not set.
    #[napi]
    pub fn get_row_outline_level(&self, env: Env, sheet: String, row: f64) -> Result<u8> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        self.inner
            .get_row_outline_level(&sheet, row)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set the width of a column (e.g., "A", "B", "AA").
    #[napi]
    pub fn set_col_width(
        &mut self,
        env: Env,
        sheet: String,
        col: String,
        width: f64,
    ) -> Result<()> {
        self.inner_mut()?
            .set_col_width(&sheet, &col, width)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the width of a column, or null if not explicitly set.
    #[napi]
    pub fn get_col_width(&self, env: Env, sheet: String, col: String) -> Result<Option<f64>> {
        self.inner
            .get_col_width(&sheet, &col)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set whether a column is visible.
    #[napi]
    pub fn set_col_visible(
        &mut self,
        env: Env,
        sheet: String,
        col: String,
        visible: bool,
    ) -> Result<()> {
        self.inner_mut()?
            .set_col_visible(&sheet, &col, visible)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get whether a column is visible. Returns true if visible (not hidden).
    #[napi]
    pub fn get_col_visible(&self, env: Env, sheet: String, col: String) -> Result<bool> {
        self.inner
            .get_col_visible(&sheet, &col)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set the outline level of a column (0-7).
    #[napi]
    pub fn set_col_outline_level(
        &mut self,
        env: Env,
        sheet: String,
        col: String,
        level: f64,
    ) -> Result<()> {
        let level = checked_u8("level", level).map_err(invalid_arg)?;
        self.inner_mut()?
            .set_col_outline_level(&sheet, &col, level)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the outline level of a column. Returns 0 if not set.
    #[napi]
    pub fn get_col_outline_level(&self, env: Env, sheet: String, col: String) -> Result<u8> {
        self.inner
            .get_col_outline_level(&sheet, &col)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Insert empty columns starting at the given column letter.
    #[napi]
    pub fn insert_cols(&mut self, env: Env, sheet: String, col: String, count: f64) -> Result<()> {
        let count = checked_u32("count", count).map_err(invalid_arg)?;
        self.inner_mut()?
            .insert_cols(&sheet, &col, count)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Remove a column by letter.
    #[napi]
    pub fn remove_col(&mut self, env: Env, sheet: String, col: String) -> Result<()> {
        self.inner_mut()?
            .remove_col(&sheet, &col)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a style definition. Returns the style ID for use with setCellStyle.
    #[napi]
    pub fn add_style(&mut self, env: Env, style: JsStyle) -> Result<u32> {
        let core_style = js_style_to_core(&style)?;
        self.inner_mut()?
            .add_style(&core_style)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the style ID applied to a cell, or null if default.
    #[napi]
    pub fn get_cell_style(&self, env: Env, sheet: String, cell: String) -> Result<Option<u32>> {
        self.inner
            .get_cell_style(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Apply a style ID to a cell.
    #[napi]
    pub fn set_cell_style(
        &mut self,
        env: Env,
        sheet: String,
        cell: String,
        style_id: f64,
    ) -> Result<()> {
        let style_id = checked_u32("styleId", style_id).map_err(invalid_arg)?;
        self.inner_mut()?
            .set_cell_style(&sheet, &cell, style_id)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Apply a style ID to an entire row.
    #[napi]
    pub fn set_row_style(
        &mut self,
        env: Env,
        sheet: String,
        row: f64,
        style_id: f64,
    ) -> Result<()> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        let style_id = checked_u32("styleId", style_id).map_err(invalid_arg)?;
        self.inner_mut()?
            .set_row_style(&sheet, row, style_id)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the style ID for a row. Returns 0 if not set.
    #[napi]
    pub fn get_row_style(&self, env: Env, sheet: String, row: f64) -> Result<u32> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        self.inner
            .get_row_style(&sheet, row)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Apply a style ID to an entire column.
    #[napi]
    pub fn set_col_style(
        &mut self,
        env: Env,
        sheet: String,
        col: String,
        style_id: f64,
    ) -> Result<()> {
        let style_id = checked_u32("styleId", style_id).map_err(invalid_arg)?;
        self.inner_mut()?
            .set_col_style(&sheet, &col, style_id)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the style ID for a column. Returns 0 if not set.
    #[napi]
    pub fn get_col_style(&self, env: Env, sheet: String, col: String) -> Result<u32> {
        self.inner
            .get_col_style(&sheet, &col)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a chart to a sheet.
    #[napi]
    pub fn add_chart(
        &mut self,
        env: Env,
        sheet: String,
        from_cell: String,
        to_cell: String,
        config: JsChartConfig,
    ) -> Result<()> {
        let core_config = ChartConfig {
            chart_type: parse_chart_type(&config.chart_type)?,
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
            view_3d: config
                .view_3d
                .map(|v| {
                    Ok::<_, sheetkit_core::error::Error>(View3DConfig {
                        rot_x: checked_opt_i32("view3d.rotX", v.rot_x)?,
                        rot_y: checked_opt_i32("view3d.rotY", v.rot_y)?,
                        depth_percent: checked_opt_u32("view3d.depthPercent", v.depth_percent)?,
                        right_angle_axes: v.right_angle_axes,
                        perspective: checked_opt_u32("view3d.perspective", v.perspective)?,
                    })
                })
                .transpose()
                .map_err(invalid_arg)?,
        };
        self.inner_mut()?
            .add_chart(&sheet, &from_cell, &to_cell, &core_config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a shape to a sheet.
    #[napi]
    pub fn add_shape(&mut self, env: Env, sheet: String, config: JsShapeConfig) -> Result<()> {
        let core_config = sheetkit_core::shape::ShapeConfig {
            shape_type: parse_shape_type(&config.shape_type)?,
            from_cell: config.from_cell,
            to_cell: config.to_cell,
            text: config.text,
            fill_color: config.fill_color,
            line_color: config.line_color,
            line_width: config.line_width,
        };
        self.inner_mut()?
            .add_shape(&sheet, &core_config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a form control to a sheet.
    #[napi]
    pub fn add_form_control(
        &mut self,
        env: Env,
        sheet: String,
        config: JsFormControlConfig,
    ) -> Result<()> {
        let core_config = js_form_control_to_core(&config)?;
        self.inner_mut()?
            .add_form_control(&sheet, core_config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all form controls on a sheet.
    #[napi]
    pub fn get_form_controls(&mut self, env: Env, sheet: String) -> Result<Vec<JsFormControlInfo>> {
        let controls = self
            .inner_mut()?
            .get_form_controls(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(controls.iter().map(core_form_control_info_to_js).collect())
    }

    /// Delete a form control by index.
    #[napi]
    pub fn delete_form_control(&mut self, env: Env, sheet: String, index: f64) -> Result<()> {
        let index = checked_u32("index", index).map_err(invalid_arg)?;
        self.inner_mut()?
            .delete_form_control(&sheet, index as usize)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add an image to a sheet.
    #[napi]
    pub fn add_image(&mut self, env: Env, sheet: String, config: JsImageConfig) -> Result<()> {
        let core_config = ImageConfig {
            data: config.data.to_vec(),
            format: parse_image_format(&config.format)?,
            from_cell: config.from_cell,
            width_px: checked_u32("widthPx", config.width_px).map_err(invalid_arg)?,
            height_px: checked_u32("heightPx", config.height_px).map_err(invalid_arg)?,
        };
        self.inner_mut()?
            .add_image(&sheet, &core_config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Delete a chart anchored at the given cell.
    #[napi]
    pub fn delete_chart(&mut self, env: Env, sheet: String, cell: String) -> Result<()> {
        self.inner_mut()?
            .delete_chart(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Delete a picture anchored at the given cell.
    #[napi]
    pub fn delete_picture(&mut self, env: Env, sheet: String, cell: String) -> Result<()> {
        self.inner_mut()?
            .delete_picture(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all pictures anchored at the given cell.
    #[napi]
    pub fn get_pictures(
        &self,
        env: Env,
        sheet: String,
        cell: String,
    ) -> Result<Vec<JsPictureInfo>> {
        let pics = self
            .inner
            .get_pictures(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(pics
            .into_iter()
            .map(|p| JsPictureInfo {
                data: p.data.into(),
                format: p.format.extension().to_string(),
                cell: p.cell,
                width_px: p.width_px,
                height_px: p.height_px,
            })
            .collect())
    }

    /// Get all cells that have pictures anchored to them on the given sheet.
    #[napi]
    pub fn get_picture_cells(&self, env: Env, sheet: String) -> Result<Vec<String>> {
        self.inner
            .get_picture_cells(&sheet)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Merge a range of cells on a sheet.
    #[napi]
    pub fn merge_cells(
        &mut self,
        env: Env,
        sheet: String,
        top_left: String,
        bottom_right: String,
    ) -> Result<()> {
        self.inner_mut()?
            .merge_cells(&sheet, &top_left, &bottom_right)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Remove a merged cell range from a sheet.
    #[napi]
    pub fn unmerge_cell(&mut self, env: Env, sheet: String, reference: String) -> Result<()> {
        self.inner_mut()?
            .unmerge_cell(&sheet, &reference)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all merged cell ranges on a sheet.
    #[napi]
    pub fn get_merge_cells(&self, env: Env, sheet: String) -> Result<Vec<String>> {
        self.inner
            .get_merge_cells(&sheet)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a data validation rule to a sheet.
    #[napi]
    pub fn add_data_validation(
        &mut self,
        env: Env,
        sheet: String,
        config: JsDataValidationConfig,
    ) -> Result<()> {
        let core_config = DataValidationConfig {
            sqref: config.sqref,
            validation_type: parse_validation_type(&config.validation_type)?,
            operator: config
                .operator
                .as_ref()
                .and_then(|s| parse_validation_operator(s)),
            formula1: config.formula1,
            formula2: config.formula2,
            allow_blank: config.allow_blank.unwrap_or(false),
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
        self.inner_mut()?
            .add_data_validation(&sheet, &core_config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all data validations on a sheet.
    #[napi]
    pub fn get_data_validations(
        &self,
        env: Env,
        sheet: String,
    ) -> Result<Vec<JsDataValidationConfig>> {
        let validations = self
            .inner
            .get_data_validations(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(validations.iter().map(core_validation_to_js).collect())
    }

    /// Remove a data validation by sqref.
    #[napi]
    pub fn remove_data_validation(&mut self, env: Env, sheet: String, sqref: String) -> Result<()> {
        self.inner_mut()?
            .remove_data_validation(&sheet, &sqref)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set conditional formatting rules on a cell range.
    #[napi]
    pub fn set_conditional_format(
        &mut self,
        env: Env,
        sheet: String,
        sqref: String,
        rules: Vec<JsConditionalFormatRule>,
    ) -> Result<()> {
        let core_rules: Vec<ConditionalFormatRule> = rules
            .iter()
            .map(js_cf_rule_to_core)
            .collect::<Result<Vec<_>>>()?;
        self.inner_mut()?
            .set_conditional_format(&sheet, &sqref, &core_rules)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all conditional formatting rules for a sheet.
    #[napi]
    pub fn get_conditional_formats(
        &self,
        env: Env,
        sheet: String,
    ) -> Result<Vec<JsConditionalFormatEntry>> {
        let formats = self
            .inner
            .get_conditional_formats(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
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
    pub fn delete_conditional_format(
        &mut self,
        env: Env,
        sheet: String,
        sqref: String,
    ) -> Result<()> {
        self.inner_mut()?
            .delete_conditional_format(&sheet, &sqref)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a comment to a cell.
    #[napi]
    pub fn add_comment(&mut self, env: Env, sheet: String, config: JsCommentConfig) -> Result<()> {
        let core_config = CommentConfig {
            cell: config.cell,
            author: config.author,
            text: config.text,
        };
        self.inner_mut()?
            .add_comment(&sheet, &core_config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all comments on a sheet.
    #[napi]
    pub fn get_comments(&mut self, env: Env, sheet: String) -> Result<Vec<JsCommentConfig>> {
        let comments = self
            .inner_mut()?
            .get_comments(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
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
    pub fn remove_comment(&mut self, env: Env, sheet: String, cell: String) -> Result<()> {
        self.inner_mut()?
            .remove_comment(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a threaded comment to a cell. Returns the comment ID.
    #[napi]
    pub fn add_threaded_comment(
        &mut self,
        env: Env,
        sheet: String,
        cell: String,
        input: JsThreadedCommentInput,
    ) -> Result<String> {
        let core_input = ThreadedCommentInput {
            author: input.author,
            text: input.text,
            parent_id: input.parent_id,
        };
        self.inner_mut()?
            .add_threaded_comment(&sheet, &cell, &core_input)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all threaded comments for a sheet.
    #[napi]
    pub fn get_threaded_comments(
        &self,
        env: Env,
        sheet: String,
    ) -> Result<Vec<JsThreadedCommentData>> {
        let comments = self
            .inner
            .get_threaded_comments(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(comments
            .into_iter()
            .map(|c| JsThreadedCommentData {
                id: c.id,
                cell_ref: c.cell_ref,
                text: c.text,
                author: c.author,
                person_id: c.person_id,
                date_time: c.date_time,
                parent_id: c.parent_id,
                done: c.done,
            })
            .collect())
    }

    /// Get threaded comments for a specific cell on a sheet.
    #[napi]
    pub fn get_threaded_comments_by_cell(
        &self,
        env: Env,
        sheet: String,
        cell: String,
    ) -> Result<Vec<JsThreadedCommentData>> {
        let comments = self
            .inner
            .get_threaded_comments_by_cell(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(comments
            .into_iter()
            .map(|c| JsThreadedCommentData {
                id: c.id,
                cell_ref: c.cell_ref,
                text: c.text,
                author: c.author,
                person_id: c.person_id,
                date_time: c.date_time,
                parent_id: c.parent_id,
                done: c.done,
            })
            .collect())
    }

    /// Delete a threaded comment by ID.
    #[napi]
    pub fn delete_threaded_comment(
        &mut self,
        env: Env,
        sheet: String,
        comment_id: String,
    ) -> Result<()> {
        self.inner_mut()?
            .delete_threaded_comment(&sheet, &comment_id)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set the resolved (done) state of a threaded comment.
    #[napi]
    pub fn resolve_threaded_comment(
        &mut self,
        env: Env,
        sheet: String,
        comment_id: String,
        done: bool,
    ) -> Result<()> {
        self.inner_mut()?
            .resolve_threaded_comment(&sheet, &comment_id, done)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a person to the person list. Returns the person ID.
    #[napi]
    pub fn add_person(&mut self, input: JsPersonInput) -> Result<String> {
        let core_input = PersonInput {
            display_name: input.display_name,
            user_id: input.user_id,
            provider_id: input.provider_id,
        };
        Ok(self.inner_mut()?.add_person(&core_input))
    }

    /// Get all persons in the person list.
    #[napi]
    pub fn get_persons(&self) -> Vec<JsPersonData> {
        self.inner
            .get_persons()
            .into_iter()
            .map(|p| JsPersonData {
                id: p.id,
                display_name: p.display_name,
                user_id: p.user_id,
                provider_id: p.provider_id,
            })
            .collect()
    }

    /// Set an auto-filter on a sheet.
    #[napi]
    pub fn set_auto_filter(&mut self, env: Env, sheet: String, range: String) -> Result<()> {
        self.inner_mut()?
            .set_auto_filter(&sheet, &range)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Remove the auto-filter from a sheet.
    #[napi]
    pub fn remove_auto_filter(&mut self, env: Env, sheet: String) -> Result<()> {
        self.inner_mut()?
            .remove_auto_filter(&sheet)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Create a new stream writer for a new sheet.
    #[napi]
    pub fn new_stream_writer(
        &self,
        env: Env,
        sheet_name: String,
        limits: Option<JsStreamLimits>,
    ) -> Result<JsStreamWriter> {
        let limits = js_stream_limits_to_core(limits.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let writer = self
            .inner
            .new_stream_writer_with_limits(&sheet_name, limits)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(JsStreamWriter {
            inner: Some(writer),
        })
    }

    /// Apply a stream writer's output to the workbook. Returns the sheet index.
    #[napi]
    pub fn apply_stream_writer(&mut self, env: Env, writer: &mut JsStreamWriter) -> Result<u32> {
        let workbook = self.inner_mut()?;
        let inner_writer = writer
            .inner
            .take()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        let index = workbook
            .apply_stream_writer(inner_writer)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(index as u32)
    }

    /// Set core document properties (title, creator, etc.).
    #[napi]
    pub fn set_doc_props(&mut self, props: JsDocProperties) -> Result<()> {
        self.inner_mut()?
            .set_doc_props(js_doc_props_to_core(&props));
        Ok(())
    }

    /// Get core document properties.
    #[napi]
    pub fn get_doc_props(&self) -> JsDocProperties {
        core_doc_props_to_js(&self.inner.get_doc_props())
    }

    /// Set application properties (company, app version, etc.).
    #[napi]
    pub fn set_app_props(&mut self, props: JsAppProperties) -> Result<()> {
        self.inner_mut()?
            .set_app_props(js_app_props_to_core(&props)?);
        Ok(())
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
        self.inner_mut()?.set_custom_property(&name, prop_value);
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
    pub fn delete_custom_property(&mut self, name: String) -> Result<bool> {
        Ok(self.inner_mut()?.delete_custom_property(&name))
    }

    /// Protect the workbook structure/windows with optional password.
    #[napi]
    pub fn protect_workbook(&mut self, config: JsWorkbookProtectionConfig) -> Result<()> {
        self.inner_mut()?
            .protect_workbook(WorkbookProtectionConfig {
                password: config.password,
                lock_structure: config.lock_structure.unwrap_or(false),
                lock_windows: config.lock_windows.unwrap_or(false),
                lock_revision: config.lock_revision.unwrap_or(false),
            });
        Ok(())
    }

    /// Remove workbook protection.
    #[napi]
    pub fn unprotect_workbook(&mut self) -> Result<()> {
        self.inner_mut()?.unprotect_workbook();
        Ok(())
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
    pub fn set_panes(&mut self, env: Env, sheet: String, cell: String) -> Result<()> {
        self.inner_mut()?
            .set_panes(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Remove any freeze or split panes from a sheet.
    #[napi]
    pub fn unset_panes(&mut self, env: Env, sheet: String) -> Result<()> {
        self.inner_mut()?
            .unset_panes(&sheet)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the current freeze pane cell reference for a sheet, or null if none.
    #[napi]
    pub fn get_panes(&self, env: Env, sheet: String) -> Result<Option<String>> {
        self.inner
            .get_panes(&sheet)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set page margins on a sheet (values in inches).
    #[napi]
    pub fn set_page_margins(
        &mut self,
        env: Env,
        sheet: String,
        margins: JsPageMargins,
    ) -> Result<()> {
        let config = PageMarginsConfig {
            left: margins.left,
            right: margins.right,
            top: margins.top,
            bottom: margins.bottom,
            header: margins.header,
            footer: margins.footer,
        };
        self.inner_mut()?
            .set_page_margins(&sheet, &config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get page margins for a sheet. Returns defaults if not explicitly set.
    #[napi]
    pub fn get_page_margins(&self, env: Env, sheet: String) -> Result<JsPageMargins> {
        let m = self
            .inner
            .get_page_margins(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
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
    pub fn set_page_setup(&mut self, env: Env, sheet: String, setup: JsPageSetup) -> Result<()> {
        let orientation = setup
            .orientation
            .as_ref()
            .and_then(|s| parse_orientation(s));
        let paper_size = setup.paper_size.as_ref().and_then(|s| parse_paper_size(s));
        self.inner_mut()?
            .set_page_setup(
                &sheet,
                orientation,
                paper_size,
                checked_opt_u32("scale", setup.scale).map_err(invalid_arg)?,
                checked_opt_u32("fitToWidth", setup.fit_to_width).map_err(invalid_arg)?,
                checked_opt_u32("fitToHeight", setup.fit_to_height).map_err(invalid_arg)?,
            )
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the page setup for a sheet.
    #[napi]
    pub fn get_page_setup(&self, env: Env, sheet: String) -> Result<JsPageSetup> {
        let orientation = self
            .inner
            .get_orientation(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        let paper_size = self
            .inner
            .get_paper_size(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        let (scale, fit_to_width, fit_to_height) = self
            .inner
            .get_page_setup_details(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(JsPageSetup {
            paper_size: paper_size.as_ref().map(paper_size_to_string),
            orientation: orientation.as_ref().map(orientation_to_string),
            scale: scale.map(f64::from),
            fit_to_width: fit_to_width.map(f64::from),
            fit_to_height: fit_to_height.map(f64::from),
        })
    }

    /// Set header and footer text for printing.
    #[napi]
    pub fn set_header_footer(
        &mut self,
        env: Env,
        sheet: String,
        header: Option<String>,
        footer: Option<String>,
    ) -> Result<()> {
        self.inner_mut()?
            .set_header_footer(&sheet, header.as_deref(), footer.as_deref())
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the header and footer text for a sheet.
    /// Returns an object with `header` and `footer` fields, each possibly null.
    #[napi]
    pub fn get_header_footer(&self, env: Env, sheet: String) -> Result<JsHeaderFooter> {
        let (header, footer) = self
            .inner
            .get_header_footer(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(JsHeaderFooter { header, footer })
    }

    /// Set print options on a sheet.
    #[napi]
    pub fn set_print_options(
        &mut self,
        env: Env,
        sheet: String,
        opts: JsPrintOptions,
    ) -> Result<()> {
        self.inner_mut()?
            .set_print_options(
                &sheet,
                opts.grid_lines,
                opts.headings,
                opts.horizontal_centered,
                opts.vertical_centered,
            )
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get print options for a sheet.
    #[napi]
    pub fn get_print_options(&self, env: Env, sheet: String) -> Result<JsPrintOptions> {
        let (gl, hd, hc, vc) = self
            .inner
            .get_print_options(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(JsPrintOptions {
            grid_lines: gl,
            headings: hd,
            horizontal_centered: hc,
            vertical_centered: vc,
        })
    }

    /// Insert a horizontal page break before the given 1-based row.
    #[napi]
    pub fn insert_page_break(&mut self, env: Env, sheet: String, row: f64) -> Result<()> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        self.inner_mut()?
            .insert_page_break(&sheet, row)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Remove a horizontal page break at the given 1-based row.
    #[napi]
    pub fn remove_page_break(&mut self, env: Env, sheet: String, row: f64) -> Result<()> {
        let row = checked_u32("row", row).map_err(invalid_arg)?;
        self.inner_mut()?
            .remove_page_break(&sheet, row)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all row page break positions (1-based row numbers).
    #[napi]
    pub fn get_page_breaks(&self, env: Env, sheet: String) -> Result<Vec<u32>> {
        self.inner
            .get_page_breaks(&sheet)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set a hyperlink on a cell.
    #[napi]
    pub fn set_cell_hyperlink(
        &mut self,
        env: Env,
        sheet: String,
        cell: String,
        opts: JsHyperlinkOptions,
    ) -> Result<()> {
        let link = parse_hyperlink_type(&opts)?;
        self.inner_mut()?
            .set_cell_hyperlink(
                &sheet,
                &cell,
                link,
                opts.display.as_deref(),
                opts.tooltip.as_deref(),
            )
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get hyperlink information for a cell, or null if no hyperlink exists.
    #[napi]
    pub fn get_cell_hyperlink(
        &self,
        env: Env,
        sheet: String,
        cell: String,
    ) -> Result<Option<JsHyperlinkInfo>> {
        let info = self
            .inner
            .get_cell_hyperlink(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(info.as_ref().map(hyperlink_info_to_js))
    }

    /// Delete a hyperlink from a cell.
    #[napi]
    pub fn delete_cell_hyperlink(&mut self, env: Env, sheet: String, cell: String) -> Result<()> {
        self.inner_mut()?
            .delete_cell_hyperlink(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all rows with their data from a sheet.
    /// Only rows that contain at least one cell are included.
    #[napi]
    pub fn get_rows(
        &self,
        env: Env,
        sheet: String,
        options: Option<JsTraversalOptions>,
    ) -> Result<Vec<JsRowData>> {
        let options = js_traversal_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        if options.mode() != sheetkit_core::traversal::TraversalMode::Sparse {
            return Err(Error::from_reason(
                "native getRows supports sparse traversal only; use getRowsBuffer for dense traversal",
            ));
        }
        let rows = self
            .inner
            .get_rows_with_limits(&sheet, options.limits())
            .map_err(|error| map_core_error(&env, error))?;

        Ok(rows
            .into_iter()
            .map(|(row_num, cells)| JsRowData {
                row: row_num,
                cells: cells
                    .into_iter()
                    .map(|(col_num, val)| {
                        let col_name =
                            sheetkit_core::utils::cell_ref::column_number_to_name(col_num)
                                .unwrap_or_default();
                        cell_value_to_row_cell(col_name, val)
                    })
                    .collect(),
            })
            .collect())
    }

    /// Serialize a sheet's cell data into a compact binary buffer (v1 format).
    /// Returns the raw bytes suitable for efficient JS-side decoding.
    #[napi]
    pub fn get_rows_buffer(
        &self,
        env: Env,
        sheet: String,
        options: Option<JsTraversalOptions>,
    ) -> Result<Buffer> {
        let options = js_traversal_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let ws = self
            .inner
            .worksheet_xml_ref(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        let sst = self.inner.sst_ref();
        let style_is_date = self.inner.transfer_date_style_lookup();
        let buf = sheetkit_core::raw_transfer::sheet_to_raw_buffer_with_date_styles_and_options(
            ws,
            sst,
            &style_is_date,
            options,
        )
        .map_err(|error| map_core_error(&env, error))?;
        Ok(Buffer::from(buf))
    }

    /// Serialize a sheet's cell data into a compact binary buffer (v2 format).
    /// The v2 format inlines strings with each cell, eliminating the global
    /// string table and enabling incremental row-by-row decoding.
    #[napi]
    pub fn get_rows_buffer_v2(
        &self,
        env: Env,
        sheet: String,
        options: Option<JsTraversalOptions>,
    ) -> Result<Buffer> {
        let options = js_traversal_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let ws = self
            .inner
            .worksheet_xml_ref(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        let sst = self.inner.sst_ref();
        let style_is_date = self.inner.transfer_date_style_lookup();
        let buf =
            sheetkit_core::raw_transfer_v2::sheet_to_raw_buffer_v2_with_date_styles_and_options(
                ws,
                sst,
                &style_is_date,
                options,
            )
            .map_err(|error| map_core_error(&env, error))?;
        Ok(Buffer::from(buf))
    }

    /// List the formula cells of a sheet with their cached results.
    ///
    /// Expressions match the formula records in `getRowsBuffer` and
    /// `getRowsBufferV2` under the same options. `consumedOutputBytes` is the
    /// size of a row buffer already produced with those options; it counts
    /// against `limits.maxOutputBytes` before this list (default 0).
    #[napi]
    pub fn get_formula_cells(
        &self,
        env: Env,
        sheet: String,
        options: Option<JsTraversalOptions>,
        consumed_output_bytes: Option<f64>,
    ) -> Result<Vec<JsFormulaCell>> {
        let options = js_traversal_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let consumed_output_bytes = consumed_output_bytes
            .map(|value| finite_limit(value, "consumedOutputBytes"))
            .transpose()
            .map_err(invalid_arg)?
            .unwrap_or(0);
        let ws = self
            .inner
            .worksheet_xml_ref(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        let style_is_date = self.inner.transfer_date_style_lookup();
        let cells = sheetkit_core::raw_transfer::sheet_formula_cells(
            ws,
            self.inner.sst_ref(),
            &style_is_date,
            options,
            consumed_output_bytes,
        )
        .map_err(|error| map_core_error(&env, error))?;
        let mut result = Vec::with_capacity(cells.len());
        for cell in cells {
            let column = sheetkit_core::utils::cell_ref::column_number_to_name(cell.col)
                .map_err(|error| map_core_error(&env, error))?;
            if let Some(formula) = cell_value_to_formula_value(&cell.value) {
                result.push(JsFormulaCell {
                    row: cell.row,
                    column,
                    formula,
                });
            }
        }
        Ok(result)
    }

    /// Apply cell data from a binary buffer to a sheet.
    /// The buffer must follow the raw transfer binary format.
    /// Optionally specify a start cell (default "A1").
    #[napi]
    pub fn set_sheet_data_buffer(
        &mut self,
        env: Env,
        sheet: String,
        buf: Buffer,
        start_cell: Option<String>,
    ) -> Result<()> {
        let (start_col, start_row) = if let Some(ref cell) = start_cell {
            sheetkit_core::utils::cell_ref::cell_name_to_coordinates(cell)
                .map_err(|error| map_core_error(&env, error))?
        } else {
            (1, 1)
        };
        let workbook = self.inner_mut()?;
        sheetkit_core::raw_transfer_write::validate_cell_transfer_target(workbook, &sheet)
            .map_err(|error| map_core_error(&env, error))?;
        let transfer = sheetkit_core::raw_transfer_write::raw_buffer_to_validated_cell_transfer(
            &buf, start_row, start_col,
        )
        .map_err(|error| map_core_error(&env, error))?;
        workbook
            .set_validated_cell_transfer(&sheet, transfer)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all columns with their data from a sheet.
    /// Only columns that have data are included.
    #[napi]
    pub fn get_cols(
        &self,
        env: Env,
        sheet: String,
        options: Option<JsTraversalOptions>,
    ) -> Result<Vec<JsColData>> {
        let options = js_traversal_options_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        if options.mode() != sheetkit_core::traversal::TraversalMode::Sparse {
            return Err(Error::from_reason("getCols supports sparse traversal only"));
        }
        let cols = self
            .inner
            .get_cols_with_limits(&sheet, options.limits())
            .map_err(|error| map_core_error(&env, error))?;

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

    /// Set a formula on a cell.
    #[napi]
    pub fn set_cell_formula(
        &mut self,
        env: Env,
        sheet: String,
        cell: String,
        formula: String,
        limits: Option<JsFormulaLimits>,
    ) -> Result<()> {
        let limits = js_formula_limits_to_core(limits.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        self.inner_mut()?
            .set_cell_formula_with_limits(&sheet, &cell, &formula, limits)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Fill a single-column range with a formula, adjusting row references
    /// for each row relative to the first cell.
    #[napi]
    pub fn fill_formula(
        &mut self,
        env: Env,
        sheet: String,
        range: String,
        formula: String,
        limits: Option<JsFormulaLimits>,
    ) -> Result<()> {
        let limits = js_formula_limits_to_core(limits.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        self.inner_mut()?
            .fill_formula_with_limits(&sheet, &range, &formula, limits)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Evaluate a formula string against the current workbook data.
    #[napi]
    pub fn evaluate_formula(
        &self,
        env: Env,
        sheet: String,
        formula: String,
        limits: Option<JsFormulaLimits>,
    ) -> Result<Either5<Null, bool, f64, String, DateValue>> {
        let limits = js_formula_limits_to_core(limits.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let result = self
            .inner
            .evaluate_formula_with_limits(&sheet, &formula, limits)
            .map_err(|error| map_core_error(&env, error))?;
        cell_value_to_either(result)
    }

    /// Recalculate all formula cells in the workbook.
    #[napi]
    pub fn calculate_all(&mut self, env: Env, limits: Option<JsFormulaLimits>) -> Result<()> {
        let limits = js_formula_limits_to_core(limits.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        self.inner_mut()?
            .calculate_all_with_limits(limits)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a pivot table to the workbook.
    #[napi]
    pub fn add_pivot_table(
        &mut self,
        env: Env,
        config: JsPivotTableConfig,
        options: Option<JsTraversalOptions>,
    ) -> Result<()> {
        let limits = js_dense_traversal_limits_to_core(options.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        precharge_pivot_config(&config, limits).map_err(|error| map_core_error(&env, error))?;

        let JsPivotTableConfig {
            name,
            source_sheet,
            source_range,
            target_sheet,
            target_cell,
            rows,
            columns,
            data,
        } = config;
        let mut core_rows = Vec::new();
        core_rows
            .try_reserve_exact(rows.len())
            .map_err(|_| Error::from_reason("failed to allocate admitted pivot rows"))?;
        core_rows.extend(
            rows.into_iter()
                .map(|field| PivotField { name: field.name }),
        );
        let mut core_columns = Vec::new();
        core_columns
            .try_reserve_exact(columns.len())
            .map_err(|_| Error::from_reason("failed to allocate admitted pivot columns"))?;
        core_columns.extend(
            columns
                .into_iter()
                .map(|field| PivotField { name: field.name }),
        );
        let mut core_data = Vec::new();
        core_data
            .try_reserve_exact(data.len())
            .map_err(|_| Error::from_reason("failed to allocate admitted pivot data fields"))?;
        for field in data {
            core_data.push(PivotDataField {
                name: field.name,
                function: parse_aggregate_function(&field.function)?,
                display_name: field.display_name,
            });
        }
        let core_config = PivotTableConfig {
            name,
            source_sheet,
            source_range,
            target_sheet,
            target_cell,
            rows: core_rows,
            columns: core_columns,
            data: core_data,
        };
        self.inner_mut()?
            .add_pivot_table_with_limits(&core_config, limits)
            .map_err(|error| map_core_error(&env, error))
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
    pub fn delete_pivot_table(&mut self, env: Env, name: String) -> Result<()> {
        self.inner_mut()?
            .delete_pivot_table(&name)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a sparkline to a worksheet.
    #[napi]
    pub fn add_sparkline(
        &mut self,
        env: Env,
        sheet: String,
        config: JsSparklineConfig,
    ) -> Result<()> {
        let core_config = js_sparkline_to_core(&config)?;
        self.inner_mut()?
            .add_sparkline(&sheet, &core_config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all sparklines for a worksheet.
    #[napi]
    pub fn get_sparklines(&self, env: Env, sheet: String) -> Result<Vec<JsSparklineConfig>> {
        let sparklines = self
            .inner
            .get_sparklines(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(sparklines.iter().map(core_sparkline_to_js).collect())
    }

    /// Remove a sparkline by its location cell reference.
    #[napi]
    pub fn remove_sparkline(&mut self, env: Env, sheet: String, location: String) -> Result<()> {
        self.inner_mut()?
            .remove_sparkline(&sheet, &location)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Add a slicer to a sheet targeting a table column.
    #[napi]
    pub fn add_slicer(&mut self, env: Env, sheet: String, config: JsSlicerConfig) -> Result<()> {
        let core_config = sheetkit_core::slicer::SlicerConfig {
            name: config.name,
            cell: config.cell,
            table_name: config.table_name,
            column_name: config.column_name,
            caption: config.caption,
            style: config.style,
            width: checked_opt_u32("width", config.width).map_err(invalid_arg)?,
            height: checked_opt_u32("height", config.height).map_err(invalid_arg)?,
            show_caption: config.show_caption,
            column_count: checked_opt_u32("columnCount", config.column_count)
                .map_err(invalid_arg)?,
        };
        self.inner_mut()?
            .add_slicer(&sheet, &core_config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all slicers on a sheet.
    #[napi]
    pub fn get_slicers(&self, env: Env, sheet: String) -> Result<Vec<JsSlicerInfo>> {
        let slicers = self
            .inner
            .get_slicers(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(slicers
            .into_iter()
            .map(|info| JsSlicerInfo {
                name: info.name,
                caption: info.caption,
                table_name: info.table_name,
                column_name: info.column_name,
                style: info.style,
            })
            .collect())
    }

    /// Delete a slicer by name from a sheet.
    #[napi]
    pub fn delete_slicer(&mut self, env: Env, sheet: String, name: String) -> Result<()> {
        self.inner_mut()?
            .delete_slicer(&sheet, &name)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set a cell to a rich text value with multiple formatted runs.
    #[napi]
    pub fn set_cell_rich_text(
        &mut self,
        env: Env,
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
        self.inner_mut()?
            .set_cell_rich_text(&sheet, &cell, core_runs)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get rich text runs for a cell, or null if not rich text.
    #[napi]
    pub fn get_cell_rich_text(
        &self,
        env: Env,
        sheet: String,
        cell: String,
    ) -> Result<Option<Vec<JsRichTextRun>>> {
        let runs = self
            .inner
            .get_cell_rich_text(&sheet, &cell)
            .map_err(|error| map_core_error(&env, error))?;
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
    pub fn get_theme_color(&self, index: f64, tint: Option<f64>) -> Result<Option<String>> {
        let index = checked_u32("index", index).map_err(invalid_arg)?;
        Ok(self.inner.get_theme_color(index, tint))
    }

    /// Add or update a defined name. If a name with the same name and scope
    /// already exists, its value and comment are updated.
    #[napi]
    pub fn set_defined_name(&mut self, env: Env, config: JsDefinedNameConfig) -> Result<()> {
        self.inner_mut()?
            .set_defined_name(
                &config.name,
                &config.value,
                config.scope.as_deref(),
                config.comment.as_deref(),
            )
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get a defined name by name and optional scope (sheet name).
    /// Returns null if no matching defined name is found.
    #[napi]
    pub fn get_defined_name(
        &self,
        env: Env,
        name: String,
        scope: Option<String>,
    ) -> Result<Option<JsDefinedNameInfo>> {
        let info = self
            .inner
            .get_defined_name(&name, scope.as_deref())
            .map_err(|error| map_core_error(&env, error))?;
        Ok(info.map(|i| self.defined_name_info_to_js(&i)))
    }

    /// Get all defined names in the workbook.
    #[napi]
    pub fn get_defined_names(&self) -> Vec<JsDefinedNameInfo> {
        self.inner
            .get_all_defined_names()
            .iter()
            .map(|i| self.defined_name_info_to_js(i))
            .collect()
    }

    /// Delete a defined name by name and optional scope (sheet name).
    #[napi]
    pub fn delete_defined_name(
        &mut self,
        env: Env,
        name: String,
        scope: Option<String>,
    ) -> Result<()> {
        self.inner_mut()?
            .delete_defined_name(&name, scope.as_deref())
            .map_err(|error| map_core_error(&env, error))
    }

    /// Protect a sheet with optional password and permission settings.
    #[napi]
    pub fn protect_sheet(
        &mut self,
        env: Env,
        sheet: String,
        config: Option<JsSheetProtectionConfig>,
    ) -> Result<()> {
        let core_config = config
            .as_ref()
            .map(js_sheet_protection_to_core)
            .unwrap_or_default();
        self.inner_mut()?
            .protect_sheet(&sheet, &core_config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Remove sheet protection.
    #[napi]
    pub fn unprotect_sheet(&mut self, env: Env, sheet: String) -> Result<()> {
        self.inner_mut()?
            .unprotect_sheet(&sheet)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Check if a sheet is protected.
    #[napi]
    pub fn is_sheet_protected(&self, env: Env, sheet: String) -> Result<bool> {
        self.inner
            .is_sheet_protected(&sheet)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set sheet view options (gridlines, zoom, view mode, etc.).
    #[napi]
    pub fn set_sheet_view_options(
        &mut self,
        env: Env,
        sheet: String,
        opts: JsSheetViewOptions,
    ) -> Result<()> {
        use sheetkit_core::sheet::ViewMode;
        let core_opts = sheetkit_core::sheet::SheetViewOptions {
            show_gridlines: opts.show_gridlines,
            show_formulas: opts.show_formulas,
            show_row_col_headers: opts.show_row_col_headers,
            zoom_scale: checked_opt_u32("zoomScale", opts.zoom_scale).map_err(invalid_arg)?,
            view_mode: opts.view_mode.as_deref().map(|s| match s {
                "pageBreak" | "pageBreakPreview" => ViewMode::PageBreak,
                "pageLayout" => ViewMode::PageLayout,
                _ => ViewMode::Normal,
            }),
            top_left_cell: opts.top_left_cell,
        };
        self.inner_mut()?
            .set_sheet_view_options(&sheet, &core_opts)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get sheet view options.
    #[napi]
    pub fn get_sheet_view_options(&self, env: Env, sheet: String) -> Result<JsSheetViewOptions> {
        let opts = self
            .inner
            .get_sheet_view_options(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(JsSheetViewOptions {
            show_gridlines: opts.show_gridlines,
            show_formulas: opts.show_formulas,
            show_row_col_headers: opts.show_row_col_headers,
            zoom_scale: opts.zoom_scale.map(f64::from),
            view_mode: opts.view_mode.map(|v| v.as_str().to_string()),
            top_left_cell: opts.top_left_cell,
        })
    }

    /// Get the workbook format ("xlsx", "xlsm", "xltx", "xltm", "xlam").
    #[napi]
    pub fn get_format(&self) -> String {
        match self.inner.format() {
            WorkbookFormat::Xlsx => "xlsx".to_string(),
            WorkbookFormat::Xlsm => "xlsm".to_string(),
            WorkbookFormat::Xltx => "xltx".to_string(),
            WorkbookFormat::Xltm => "xltm".to_string(),
            WorkbookFormat::Xlam => "xlam".to_string(),
        }
    }

    /// Add a table to a sheet.
    #[napi]
    pub fn add_table(&mut self, env: Env, sheet: String, config: JsTableConfig) -> Result<()> {
        let core_config = TableConfig {
            name: config.name,
            display_name: config.display_name,
            range: config.range,
            columns: config
                .columns
                .into_iter()
                .map(|c| TableColumn {
                    name: c.name,
                    totals_row_function: c.totals_row_function,
                    totals_row_label: c.totals_row_label,
                })
                .collect(),
            show_header_row: config.show_header_row.unwrap_or(true),
            style_name: config.style_name,
            auto_filter: config.auto_filter.unwrap_or(true),
            show_first_column: config.show_first_column.unwrap_or(false),
            show_last_column: config.show_last_column.unwrap_or(false),
            show_row_stripes: config.show_row_stripes.unwrap_or(true),
            show_column_stripes: config.show_column_stripes.unwrap_or(false),
        };
        self.inner_mut()?
            .add_table(&sheet, &core_config)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get all tables on a sheet.
    #[napi]
    pub fn get_tables(&self, env: Env, sheet: String) -> Result<Vec<JsTableInfo>> {
        let tables = self
            .inner
            .get_tables(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(tables
            .into_iter()
            .map(|t| JsTableInfo {
                name: t.name,
                display_name: t.display_name,
                range: t.range,
                show_header_row: t.show_header_row,
                auto_filter: t.auto_filter,
                columns: t.columns,
                style_name: t.style_name,
            })
            .collect())
    }

    /// Delete a table from a sheet by name.
    #[napi]
    pub fn delete_table(&mut self, env: Env, sheet: String, name: String) -> Result<()> {
        self.inner_mut()?
            .delete_table(&sheet, &name)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Set sheet visibility ("visible", "hidden", or "veryHidden").
    #[napi]
    pub fn set_sheet_visibility(
        &mut self,
        env: Env,
        sheet: String,
        visibility: String,
    ) -> Result<()> {
        use sheetkit_core::sheet::SheetVisibility;
        let vis = match visibility.as_str() {
            "visible" => SheetVisibility::Visible,
            "hidden" => SheetVisibility::Hidden,
            "veryHidden" => SheetVisibility::VeryHidden,
            other => {
                return Err(Error::from_reason(format!(
                    "Invalid visibility: \"{other}\". Must be \"visible\", \"hidden\", or \"veryHidden\""
                )));
            }
        };
        self.inner_mut()?
            .set_sheet_visibility(&sheet, vis)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get sheet visibility. Returns "visible", "hidden", or "veryHidden".
    #[napi]
    pub fn get_sheet_visibility(&self, env: Env, sheet: String) -> Result<String> {
        let vis = self
            .inner
            .get_sheet_visibility(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(match vis {
            sheetkit_core::sheet::SheetVisibility::Visible => "visible".to_string(),
            sheetkit_core::sheet::SheetVisibility::Hidden => "hidden".to_string(),
            sheetkit_core::sheet::SheetVisibility::VeryHidden => "veryHidden".to_string(),
        })
    }

    /// Render a worksheet to an SVG string.
    #[napi]
    pub fn render_to_svg(
        &self,
        env: Env,
        options: JsRenderOptions,
        traversal: Option<JsTraversalOptions>,
    ) -> Result<String> {
        let limits = js_dense_traversal_limits_to_core(traversal.as_ref())
            .map_err(|error| map_core_error(&env, error))?;
        let render_opts = sheetkit_core::render::RenderOptions {
            sheet_name: options.sheet_name,
            range: options.range,
            show_gridlines: options.show_gridlines.unwrap_or(true),
            show_headers: options.show_headers.unwrap_or(true),
            scale: options.scale.unwrap_or(1.0),
            default_font_family: options
                .default_font_family
                .unwrap_or_else(|| "Arial".to_string()),
            default_font_size: options.default_font_size.unwrap_or(11.0),
        };
        self.inner
            .render_to_svg_with_limits(&render_opts, limits)
            .map_err(|error| map_core_error(&env, error))
    }

    /// Get the raw VBA project binary (xl/vbaProject.bin), or null if not present.
    #[napi]
    pub fn get_vba_project(&self) -> Option<Buffer> {
        self.inner
            .get_vba_project()
            .map(|b| Buffer::from(b.to_vec()))
    }

    /// Create an owned forward-only streaming reader for the named sheet.
    /// The reader owns an independent bounded source and shared SST snapshot,
    /// enabling a lifetime independent from the workbook.
    #[napi]
    pub fn open_sheet_reader(
        &self,
        env: Env,
        sheet: String,
        batch_size: Option<f64>,
    ) -> Result<NativeSheetStreamReader> {
        validate_stream_batch_size(batch_size).map_err(|error| map_core_error(&env, error))?;
        let reader = self
            .inner
            .open_sheet_reader_owned(&sheet)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(NativeSheetStreamReader {
            inner: Some(reader),
        })
    }

    /// Extract VBA module source code from the workbook's VBA project.
    /// Returns null if no VBA project is present.
    #[napi]
    pub fn get_vba_modules(
        &self,
        env: Env,
        limits: Option<JsVbaLimits>,
    ) -> Result<Option<JsVbaProject>> {
        let limits =
            js_vba_limits_to_core(limits.as_ref()).map_err(|error| map_core_error(&env, error))?;
        let project = self
            .inner
            .get_vba_modules_with_limits(limits)
            .map_err(|error| map_core_error(&env, error))?;
        Ok(project.map(|p| {
            let modules = p
                .modules
                .into_iter()
                .map(|m| {
                    let module_type = match m.module_type {
                        sheetkit_core::vba::VbaModuleType::Standard => "standard",
                        sheetkit_core::vba::VbaModuleType::Class => "class",
                        sheetkit_core::vba::VbaModuleType::Form => "form",
                        sheetkit_core::vba::VbaModuleType::Document => "document",
                        sheetkit_core::vba::VbaModuleType::ThisWorkbook => "thisWorkbook",
                    };
                    JsVbaModule {
                        name: m.name,
                        source_code: m.source_code,
                        module_type: module_type.to_string(),
                    }
                })
                .collect();
            JsVbaProject {
                modules,
                warnings: p.warnings,
            }
        }))
    }
}

impl Workbook {
    /// Convert a core DefinedNameInfo to a JS-facing JsDefinedNameInfo,
    /// resolving the sheet index back to a sheet name.
    fn defined_name_info_to_js(
        &self,
        info: &sheetkit_core::defined_names::DefinedNameInfo,
    ) -> JsDefinedNameInfo {
        use sheetkit_core::defined_names::DefinedNameScope;
        let scope = match &info.scope {
            DefinedNameScope::Workbook => None,
            DefinedNameScope::Sheet(idx) => {
                let names = self.inner.sheet_names();
                names.get(*idx as usize).map(|s| s.to_string())
            }
        };
        JsDefinedNameInfo {
            name: info.name.clone(),
            value: info.value.clone(),
            scope,
            comment: info.comment.clone(),
        }
    }
}

/// Forward-only streaming reader for worksheet data.
///
/// Reads rows in batches without loading the entire sheet into memory.
/// Created via `Workbook.openSheetReader()`.
#[napi]
pub struct NativeSheetStreamReader {
    inner: Option<sheetkit_core::stream_reader::OwnedSheetStreamReader>,
}

#[napi]
impl NativeSheetStreamReader {
    /// Read the next batch of rows. Returns null when there are no more rows.
    #[napi]
    pub fn next_batch(
        &mut self,
        env: Env,
        batch_size: Option<f64>,
    ) -> Result<Option<Vec<JsRowData>>> {
        let size =
            validate_stream_batch_size(batch_size).map_err(|error| map_core_error(&env, error))?;
        let reader = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("SheetStreamReader already closed"))?;
        let rows = reader
            .next_batch(size)
            .map_err(|error| map_core_error(&env, error))?;
        if rows.is_empty() {
            return Ok(None);
        }
        Ok(Some(
            rows.into_iter()
                .map(|sr| JsRowData {
                    row: sr.row_number,
                    cells: sr
                        .cells
                        .into_iter()
                        .map(|(col_num, val)| {
                            let col_name =
                                sheetkit_core::utils::cell_ref::column_number_to_name(col_num)
                                    .unwrap_or_default();
                            cell_value_to_row_cell(col_name, val)
                        })
                        .collect(),
                })
                .collect(),
        ))
    }

    /// Returns true if there are potentially more rows to read.
    #[napi]
    pub fn has_more(&self) -> Result<bool> {
        let reader = self
            .inner
            .as_ref()
            .ok_or_else(|| Error::from_reason("SheetStreamReader already closed"))?;
        Ok(reader.has_more())
    }

    /// Read the next batch of rows as a compact buffer that the TypeScript
    /// wrapper decodes into the same objects `nextBatch` returns.
    /// Returns null when there are no more rows.
    #[napi]
    pub fn next_batch_buffer(
        &mut self,
        env: Env,
        batch_size: Option<f64>,
    ) -> Result<Option<Buffer>> {
        let size =
            validate_stream_batch_size(batch_size).map_err(|error| map_core_error(&env, error))?;
        let reader = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("SheetStreamReader already closed"))?;
        let rows = reader
            .next_batch(size)
            .map_err(|error| map_core_error(&env, error))?;
        if rows.is_empty() {
            return Ok(None);
        }
        let bytes = encode_stream_batch(&rows).map_err(Error::from_reason)?;
        Ok(Some(Buffer::from(bytes)))
    }

    /// Close the reader and release resources.
    #[napi]
    pub fn close(&mut self) -> Result<()> {
        self.inner.take();
        Ok(())
    }
}
