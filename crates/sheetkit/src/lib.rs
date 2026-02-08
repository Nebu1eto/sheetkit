//! SheetKit: High-level API for reading and writing Excel (.xlsx) files.
//!
//! # Quick Start
//!
//! ```no_run
//! use sheetkit::Workbook;
//!
//! let wb = Workbook::new();
//! wb.save("output.xlsx").unwrap();
//! ```

// Re-export core types
pub use sheetkit_core::error::{Error, Result};
pub use sheetkit_core::workbook::Workbook;

/// Utility functions for cell reference conversion.
pub mod utils {
    pub use sheetkit_core::utils::cell_ref::{
        cell_name_to_coordinates, column_name_to_number, column_number_to_name,
        coordinates_to_cell_name,
    };
    pub use sheetkit_core::utils::constants;
}
