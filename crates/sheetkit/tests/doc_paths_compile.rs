//! Compile-time guard for the `sheetkit::` import paths used in `docs/`.
//!
//! Every path below appears in a documentation code block. If a re-export is
//! removed or renamed, this file stops compiling and CI fails.

#[allow(unused_imports)]
use sheetkit::{
    builtin_format_code,
    cell::CellValue as _,
    chart::{ChartConfig as _, ChartSeries as _, ChartType as _},
    comment::CommentConfig as _,
    conditional::{
        CfOperator as _, ConditionalFormatRule as _, ConditionalFormatType as _,
        ConditionalStyle as _,
    },
    doc_props::{AppProperties as _, CustomPropertyValue as _, DocProperties as _},
    format_number,
    hyperlink::HyperlinkType as _,
    image::{ImageConfig as _, ImageFormat as _},
    is_date_format_code, is_date_num_fmt,
    page_layout::{Orientation as _, PageMarginsConfig as _, PaperSize as _},
    pivot::{AggregateFunction as _, PivotDataField as _, PivotField as _, PivotTableConfig as _},
    protection::WorkbookProtectionConfig as _,
    rich_text_to_plain,
    sheet::{
        SheetProtectionConfig as _, SheetViewOptions as _, SheetVisibility as _, ViewMode as _,
    },
    stream::StreamRowOptions as _,
    style::{
        BorderLineStyle as _, BorderSideStyle as _, BorderStyle as _, FillStyle as _,
        FontStyle as _, NumFmtStyle as _, PatternType as _, Style as _, StyleColor as _,
    },
    table::{TableColumn as _, TableConfig as _},
    utils::cell_ref,
    validation::{
        DataValidationConfig as _, ErrorStyle as _, ValidationOperator as _, ValidationType as _,
    },
    vba::{VbaModule as _, VbaModuleType as _},
    AppProperties, AuxParts, ChartConfig, ChartSeries, ChartType, CommentConfig,
    CustomPropertyValue, DataValidationConfig, DateInterpretation, DocProperties, Error,
    ErrorStyle, FormControlConfig, FormControlType, ImageConfig, ImageFormat, NumFmtStyle,
    OpenOptions, PersonInput, ReadMode, RenderOptions, Result, RichTextRun, SaveLimits,
    ShapeConfig, ShapeType, SlicerConfig, SparklineConfig, SparklineType, Style,
    ThreadedCommentInput, ValidationType, Workbook, WorkbookFormat, WorkbookProtectionConfig,
};

#[test]
fn documented_cell_ref_module_resolves() {
    assert_eq!(cell_ref::cell_name_to_coordinates("B3").unwrap(), (2, 3));
    assert_eq!(cell_ref::coordinates_to_cell_name(2, 3).unwrap(), "B3");
    assert_eq!(cell_ref::column_name_to_number("AA").unwrap(), 27);
    assert_eq!(cell_ref::column_number_to_name(27).unwrap(), "AA");
}
