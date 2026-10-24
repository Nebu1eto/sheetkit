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

#[cfg(feature = "encryption")]
pub use sheetkit_core::crypt::CryptLimits;
pub use sheetkit_core::defined_names::{DefinedNameInfo, DefinedNameScope};
pub use sheetkit_core::doc_props::{AppProperties, CustomPropertyValue, DocProperties};
pub use sheetkit_core::error::{Error, Result};
pub use sheetkit_core::formula::FormulaLimits;
pub use sheetkit_core::protection::WorkbookProtectionConfig;
pub use sheetkit_core::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};
pub use sheetkit_core::save::{SaveLimits, SaveOptions, StreamLimits};
pub use sheetkit_core::sheet::SheetProtectionConfig;
pub use sheetkit_core::stream::StreamWriter;
pub use sheetkit_core::workbook::{
    AuxParts, DateInterpretation, OpenLimits, OpenOptions, ReadMode, Workbook, WorkbookFormat,
};

pub use sheetkit_core::cell::{
    date_to_serial, datetime_to_serial, is_date_format_code, is_date_num_fmt, serial_to_date,
    serial_to_datetime, CellValue,
};
pub use sheetkit_core::chart::{ChartConfig, ChartSeries, ChartType, View3DConfig};
pub use sheetkit_core::comment::CommentConfig;
pub use sheetkit_core::conditional::{
    CfOperator, CfValueType, ConditionalFormatRule, ConditionalFormatType, ConditionalStyle,
};
pub use sheetkit_core::control::{FormControlConfig, FormControlInfo, FormControlType};
pub use sheetkit_core::hyperlink::{HyperlinkInfo, HyperlinkType};
pub use sheetkit_core::image::{ImageConfig, ImageFormat, PictureInfo};
pub use sheetkit_core::numfmt::{builtin_format_code, format_number, format_with_builtin};
pub use sheetkit_core::output::{
    serialize_csv, serialize_csv_with_policy, CsvFormulaPolicy, MAX_CSV_OUTPUT_BYTES,
};
pub use sheetkit_core::page_layout::{Orientation, PageMarginsConfig, PaperSize};
pub use sheetkit_core::pivot::{
    AggregateFunction, PivotDataField, PivotField, PivotTableConfig, PivotTableInfo,
};
pub use sheetkit_core::render::RenderOptions;
pub use sheetkit_core::rich_text::{rich_text_to_plain, RichTextRun};
pub use sheetkit_core::shape::{ShapeConfig, ShapeType};
pub use sheetkit_core::slicer::{SlicerConfig, SlicerInfo};
pub use sheetkit_core::sparkline::{SparklineConfig, SparklineType};
pub use sheetkit_core::style::{
    AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle, FillStyle, FontStyle,
    GradientFillStyle, GradientStop, GradientType, HorizontalAlign, NumFmtStyle, PatternType,
    ProtectionStyle, Style, StyleColor, VerticalAlign,
};
pub use sheetkit_core::table::{TableColumn, TableConfig, TableInfo};
pub use sheetkit_core::threaded_comment::{
    PersonData, PersonInput, ThreadedCommentData, ThreadedCommentInput,
};
pub use sheetkit_core::traversal::{TraversalLimits, TraversalMode, TraversalOptions};
pub use sheetkit_core::validation::{
    DataValidationConfig, ErrorStyle, ValidationOperator, ValidationType,
};
pub use sheetkit_core::vba::{VbaLimits, VbaModule, VbaModuleType, VbaProject};

/// Utility functions for cell reference conversion.
pub mod utils {
    pub use sheetkit_core::utils::cell_ref::{
        cell_name_to_coordinates, column_name_to_number, column_number_to_name,
        coordinates_to_cell_name,
    };
    pub use sheetkit_core::utils::constants;

    /// Cell reference conversion functions, grouped by module.
    pub mod cell_ref {
        pub use sheetkit_core::utils::cell_ref::{
            cell_name_to_coordinates, column_name_to_number, column_number_to_name,
            coordinates_to_cell_name,
        };
    }
}

/// Table configuration types, also available at the crate root.
pub mod table {
    pub use sheetkit_core::table::{TableColumn, TableConfig, TableInfo};
}

/// Cell value types and date serial helpers, also available at the crate root.
pub mod cell {
    pub use sheetkit_core::cell::{
        date_to_serial, datetime_to_serial, is_date_format_code, is_date_num_fmt, serial_to_date,
        serial_to_datetime, CellValue,
    };
}

/// Chart configuration types, also available at the crate root.
pub mod chart {
    pub use sheetkit_core::chart::{ChartConfig, ChartSeries, ChartType, View3DConfig};
}

/// Comment configuration types, also available at the crate root.
pub mod comment {
    pub use sheetkit_core::comment::CommentConfig;
}

/// Conditional formatting types, also available at the crate root.
pub mod conditional {
    pub use sheetkit_core::conditional::{
        CfOperator, CfValueType, ConditionalFormatRule, ConditionalFormatType, ConditionalStyle,
    };
}

/// Document property types, also available at the crate root.
pub mod doc_props {
    pub use sheetkit_core::doc_props::{AppProperties, CustomPropertyValue, DocProperties};
}

/// Hyperlink types, also available at the crate root.
pub mod hyperlink {
    pub use sheetkit_core::hyperlink::{HyperlinkInfo, HyperlinkType};
}

/// Image types, also available at the crate root.
pub mod image {
    pub use sheetkit_core::image::{ImageConfig, ImageFormat, PictureInfo};
}

/// Page layout types, also available at the crate root.
pub mod page_layout {
    pub use sheetkit_core::page_layout::{Orientation, PageMarginsConfig, PaperSize};
}

/// Pivot table types, also available at the crate root.
pub mod pivot {
    pub use sheetkit_core::pivot::{
        AggregateFunction, PivotDataField, PivotField, PivotTableConfig, PivotTableInfo,
    };
}

/// Workbook protection types, also available at the crate root.
pub mod protection {
    pub use sheetkit_core::protection::WorkbookProtectionConfig;
}

/// Sheet protection, view, and visibility types.
pub mod sheet {
    pub use sheetkit_core::sheet::{
        SheetProtectionConfig, SheetViewOptions, SheetVisibility, ViewMode,
    };
}

/// Streaming writer types.
pub mod stream {
    pub use sheetkit_core::stream::{StreamRowOptions, StreamWriter};
}

/// Style types, also available at the crate root.
pub mod style {
    pub use sheetkit_core::style::{
        AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle, FillStyle, FontStyle,
        GradientFillStyle, GradientStop, GradientType, HorizontalAlign, NumFmtStyle, PatternType,
        ProtectionStyle, Style, StyleColor, VerticalAlign,
    };
}

/// Data validation types, also available at the crate root.
pub mod validation {
    pub use sheetkit_core::validation::{
        DataValidationConfig, ErrorStyle, ValidationOperator, ValidationType,
    };
}

/// VBA project types, also available at the crate root.
pub mod vba {
    pub use sheetkit_core::vba::{VbaLimits, VbaModule, VbaModuleType, VbaProject};
}

#[cfg(test)]
mod tests {
    #[test]
    fn table_types_are_reachable_through_the_facade() {
        use crate::table::{TableColumn, TableConfig};

        let mut wb = crate::Workbook::new();
        let config = TableConfig {
            name: "Sales".to_string(),
            display_name: "Sales".to_string(),
            range: "A1:A3".to_string(),
            columns: vec![TableColumn {
                name: "Qty".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        wb.add_table("Sheet1", &config).unwrap();
        let tables: Vec<crate::TableInfo> = wb.get_tables("Sheet1").unwrap();
        assert_eq!(tables[0].name, "Sales");
        let _: crate::TableColumn = config.columns[0].clone();
    }
}
