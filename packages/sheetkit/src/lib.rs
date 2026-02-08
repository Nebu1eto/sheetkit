#![deny(clippy::all)]

use napi::bindgen_prelude::*;
use napi_derive::napi;

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
}
