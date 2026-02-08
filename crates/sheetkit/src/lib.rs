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

pub use sheetkit_core::doc_props::{AppProperties, CustomPropertyValue, DocProperties};
pub use sheetkit_core::error::{Error, Result};
pub use sheetkit_core::protection::WorkbookProtectionConfig;
pub use sheetkit_core::stream::StreamWriter;
pub use sheetkit_core::workbook::Workbook;

pub use sheetkit_core::cell::{
    date_to_serial, datetime_to_serial, is_date_format_code, is_date_num_fmt, serial_to_date,
    serial_to_datetime, CellValue,
};
pub use sheetkit_core::chart::{ChartConfig, ChartSeries, ChartType, View3DConfig};
pub use sheetkit_core::comment::CommentConfig;
pub use sheetkit_core::conditional::{
    CfOperator, CfValueType, ConditionalFormatRule, ConditionalFormatType, ConditionalStyle,
};
pub use sheetkit_core::hyperlink::{HyperlinkInfo, HyperlinkType};
pub use sheetkit_core::image::{ImageConfig, ImageFormat};
pub use sheetkit_core::page_layout::{Orientation, PageMarginsConfig, PaperSize};
pub use sheetkit_core::style::{
    AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle, FillStyle, FontStyle,
    HorizontalAlign, NumFmtStyle, PatternType, ProtectionStyle, Style, StyleColor, VerticalAlign,
};
pub use sheetkit_core::validation::{
    DataValidationConfig, ErrorStyle, ValidationOperator, ValidationType,
};

/// Utility functions for cell reference conversion.
pub mod utils {
    pub use sheetkit_core::utils::cell_ref::{
        cell_name_to_coordinates, column_name_to_number, column_number_to_name,
        coordinates_to_cell_name,
    };
    pub use sheetkit_core::utils::constants;
}
