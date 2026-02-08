#![deny(clippy::all)]

use napi::bindgen_prelude::*;
use napi::{Env, JsUnknown, ValueType};
use napi_derive::napi;

use sheetkit_core::cell::CellValue;
use sheetkit_core::chart::{ChartConfig, ChartSeries, ChartType};
use sheetkit_core::comment::CommentConfig;
use sheetkit_core::conditional::{
    CfOperator, CfValueType, ConditionalFormatRule, ConditionalFormatType, ConditionalStyle,
};
use sheetkit_core::doc_props::{AppProperties, CustomPropertyValue, DocProperties};
use sheetkit_core::hyperlink::{HyperlinkInfo, HyperlinkType};
use sheetkit_core::image::{ImageConfig, ImageFormat};
use sheetkit_core::page_layout::{Orientation, PageMarginsConfig, PaperSize};
use sheetkit_core::protection::WorkbookProtectionConfig;
use sheetkit_core::stream::StreamWriter;
use sheetkit_core::style::{
    AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle, FillStyle, FontStyle,
    HorizontalAlign, NumFmtStyle, PatternType, ProtectionStyle, Style, StyleColor, VerticalAlign,
};
use sheetkit_core::validation::{
    DataValidationConfig, ErrorStyle, ValidationOperator, ValidationType,
};

#[napi(object)]
pub struct JsFontStyle {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strikethrough: Option<bool>,
    pub color: Option<String>,
}

#[napi(object)]
pub struct JsFillStyle {
    pub pattern: Option<String>,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
}

#[napi(object)]
pub struct JsBorderSideStyle {
    pub style: Option<String>,
    pub color: Option<String>,
}

#[napi(object)]
pub struct JsBorderStyle {
    pub left: Option<JsBorderSideStyle>,
    pub right: Option<JsBorderSideStyle>,
    pub top: Option<JsBorderSideStyle>,
    pub bottom: Option<JsBorderSideStyle>,
    pub diagonal: Option<JsBorderSideStyle>,
}

#[napi(object)]
pub struct JsAlignmentStyle {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: Option<bool>,
    pub text_rotation: Option<u32>,
    pub indent: Option<u32>,
    pub shrink_to_fit: Option<bool>,
}

#[napi(object)]
pub struct JsProtectionStyle {
    pub locked: Option<bool>,
    pub hidden: Option<bool>,
}

#[napi(object)]
pub struct JsStyle {
    pub font: Option<JsFontStyle>,
    pub fill: Option<JsFillStyle>,
    pub border: Option<JsBorderStyle>,
    pub alignment: Option<JsAlignmentStyle>,
    pub num_fmt_id: Option<u32>,
    pub custom_num_fmt: Option<String>,
    pub protection: Option<JsProtectionStyle>,
}

#[napi(object)]
pub struct JsChartSeries {
    pub name: String,
    pub categories: String,
    pub values: String,
}

#[napi(object)]
pub struct JsChartConfig {
    pub chart_type: String,
    pub title: Option<String>,
    pub series: Vec<JsChartSeries>,
    pub show_legend: Option<bool>,
}

#[napi(object)]
pub struct JsImageConfig {
    pub data: Buffer,
    pub format: String,
    pub from_cell: String,
    pub width_px: u32,
    pub height_px: u32,
}

#[napi(object)]
pub struct JsCommentConfig {
    pub cell: String,
    pub author: String,
    pub text: String,
}

#[napi(object)]
pub struct JsDataValidationConfig {
    pub sqref: String,
    pub validation_type: String,
    pub operator: Option<String>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub allow_blank: Option<bool>,
    pub error_style: Option<String>,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
    pub prompt_title: Option<String>,
    pub prompt_message: Option<String>,
    pub show_input_message: Option<bool>,
    pub show_error_message: Option<bool>,
}

#[napi(object)]
pub struct JsDocProperties {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub keywords: Option<String>,
    pub description: Option<String>,
    pub last_modified_by: Option<String>,
    pub revision: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
    pub category: Option<String>,
    pub content_status: Option<String>,
}

#[napi(object)]
pub struct JsAppProperties {
    pub application: Option<String>,
    pub doc_security: Option<u32>,
    pub company: Option<String>,
    pub app_version: Option<String>,
    pub manager: Option<String>,
    pub template: Option<String>,
}

#[napi(object)]
pub struct JsWorkbookProtectionConfig {
    pub password: Option<String>,
    pub lock_structure: Option<bool>,
    pub lock_windows: Option<bool>,
    pub lock_revision: Option<bool>,
}

/// Page margins configuration in inches.
#[napi(object)]
pub struct JsPageMargins {
    /// Left margin in inches (default 0.7).
    pub left: f64,
    /// Right margin in inches (default 0.7).
    pub right: f64,
    /// Top margin in inches (default 0.75).
    pub top: f64,
    /// Bottom margin in inches (default 0.75).
    pub bottom: f64,
    /// Header margin in inches (default 0.3).
    pub header: f64,
    /// Footer margin in inches (default 0.3).
    pub footer: f64,
}

/// Page setup configuration.
#[napi(object)]
pub struct JsPageSetup {
    /// Paper size: "letter", "tabloid", "legal", "a3", "a4", "a5", "b4", "b5".
    pub paper_size: Option<String>,
    /// Orientation: "portrait" or "landscape".
    pub orientation: Option<String>,
    /// Print scale percentage (10-400).
    pub scale: Option<u32>,
    /// Fit to this many pages wide.
    pub fit_to_width: Option<u32>,
    /// Fit to this many pages tall.
    pub fit_to_height: Option<u32>,
}

/// Print options configuration.
#[napi(object)]
pub struct JsPrintOptions {
    /// Print gridlines.
    pub grid_lines: Option<bool>,
    /// Print row/column headings.
    pub headings: Option<bool>,
    /// Center horizontally on page.
    pub horizontal_centered: Option<bool>,
    /// Center vertically on page.
    pub vertical_centered: Option<bool>,
}

/// Header and footer text.
#[napi(object)]
pub struct JsHeaderFooter {
    /// Header text (may use Excel formatting codes like &L, &C, &R).
    pub header: Option<String>,
    /// Footer text (may use Excel formatting codes like &L, &C, &R).
    pub footer: Option<String>,
}

fn parse_paper_size(s: &str) -> Option<PaperSize> {
    match s.to_lowercase().as_str() {
        "letter" => Some(PaperSize::Letter),
        "tabloid" => Some(PaperSize::Tabloid),
        "legal" => Some(PaperSize::Legal),
        "a3" => Some(PaperSize::A3),
        "a4" => Some(PaperSize::A4),
        "a5" => Some(PaperSize::A5),
        "b4" => Some(PaperSize::B4),
        "b5" => Some(PaperSize::B5),
        _ => None,
    }
}

fn paper_size_to_string(ps: &PaperSize) -> String {
    match ps {
        PaperSize::Letter => "letter".to_string(),
        PaperSize::Tabloid => "tabloid".to_string(),
        PaperSize::Legal => "legal".to_string(),
        PaperSize::A3 => "a3".to_string(),
        PaperSize::A4 => "a4".to_string(),
        PaperSize::A5 => "a5".to_string(),
        PaperSize::B4 => "b4".to_string(),
        PaperSize::B5 => "b5".to_string(),
    }
}

fn parse_orientation(s: &str) -> Option<Orientation> {
    match s.to_lowercase().as_str() {
        "portrait" => Some(Orientation::Portrait),
        "landscape" => Some(Orientation::Landscape),
        _ => None,
    }
}

fn orientation_to_string(o: &Orientation) -> String {
    match o {
        Orientation::Portrait => "portrait".to_string(),
        Orientation::Landscape => "landscape".to_string(),
    }
}

#[napi(object)]
pub struct JsHyperlinkOptions {
    /// Type of hyperlink: "external", "internal", or "email".
    pub link_type: String,
    /// The target URL, sheet reference, or email address.
    pub target: String,
    /// Optional display text.
    pub display: Option<String>,
    /// Optional tooltip text.
    pub tooltip: Option<String>,
}

#[napi(object)]
pub struct JsHyperlinkInfo {
    /// Type of hyperlink: "external", "internal", or "email".
    pub link_type: String,
    /// The target URL, sheet reference, or email address.
    pub target: String,
    /// Optional display text.
    pub display: Option<String>,
    /// Optional tooltip text.
    pub tooltip: Option<String>,
}

/// Conditional formatting style (differential format).
#[napi(object)]
pub struct JsConditionalStyle {
    pub font: Option<JsFontStyle>,
    pub fill: Option<JsFillStyle>,
    pub border: Option<JsBorderStyle>,
    pub custom_num_fmt: Option<String>,
}

/// Conditional formatting rule configuration.
#[napi(object)]
pub struct JsConditionalFormatRule {
    /// Rule type: "cellIs", "expression", "colorScale", "dataBar",
    /// "duplicateValues", "uniqueValues", "top10", "bottom10",
    /// "aboveAverage", "containsBlanks", "notContainsBlanks",
    /// "containsErrors", "notContainsErrors", "containsText",
    /// "notContainsText", "beginsWith", "endsWith".
    pub rule_type: String,
    /// Comparison operator for cellIs rules.
    pub operator: Option<String>,
    /// First formula/value.
    pub formula: Option<String>,
    /// Second formula/value (for between/notBetween).
    pub formula2: Option<String>,
    /// Text for text-based rules.
    pub text: Option<String>,
    /// Rank for top10/bottom10 rules.
    pub rank: Option<u32>,
    /// Whether rank is a percentage.
    pub percent: Option<bool>,
    /// Whether rule is above average (for aboveAverage rules).
    pub above: Option<bool>,
    /// Whether equal values count as matching (for aboveAverage rules).
    pub equal_average: Option<bool>,
    /// Color scale minimum value type.
    pub min_type: Option<String>,
    /// Color scale minimum value.
    pub min_value: Option<String>,
    /// Color scale minimum color (ARGB hex).
    pub min_color: Option<String>,
    /// Color scale middle value type.
    pub mid_type: Option<String>,
    /// Color scale middle value.
    pub mid_value: Option<String>,
    /// Color scale middle color (ARGB hex).
    pub mid_color: Option<String>,
    /// Color scale maximum value type.
    pub max_type: Option<String>,
    /// Color scale maximum value.
    pub max_value: Option<String>,
    /// Color scale maximum color (ARGB hex).
    pub max_color: Option<String>,
    /// Data bar color (ARGB hex).
    pub bar_color: Option<String>,
    /// Whether to show the cell value alongside the data bar.
    pub show_value: Option<bool>,
    /// Differential style to apply.
    pub format: Option<JsConditionalStyle>,
    /// Rule priority (lower = higher precedence).
    pub priority: Option<u32>,
    /// If true, no lower-priority rules apply when this matches.
    pub stop_if_true: Option<bool>,
}

/// Result of getting conditional formats from a sheet.
#[napi(object)]
pub struct JsConditionalFormatEntry {
    /// Cell range (e.g., "A1:A100").
    pub sqref: String,
    /// Rules applied to this range.
    pub rules: Vec<JsConditionalFormatRule>,
}

/// A single cell entry with its column name and value.
#[napi(object)]
pub struct JsRowCell {
    /// Column name (e.g., "A", "B", "AA").
    pub column: String,
    /// Cell value type: "string", "number", "boolean", "date", "empty", "error", "formula".
    pub value_type: String,
    /// String representation of the cell value.
    pub value: Option<String>,
    /// Numeric value (only set when value_type is "number").
    pub number_value: Option<f64>,
    /// Boolean value (only set when value_type is "boolean").
    pub bool_value: Option<bool>,
}

/// A row with its 1-based row number and cell data.
#[napi(object)]
pub struct JsRowData {
    /// 1-based row number.
    pub row: u32,
    /// Cells with data in this row.
    pub cells: Vec<JsRowCell>,
}

/// A single cell entry with its row number and value.
#[napi(object)]
pub struct JsColCell {
    /// 1-based row number.
    pub row: u32,
    /// Cell value type: "string", "number", "boolean", "date", "empty", "error", "formula".
    pub value_type: String,
    /// String representation of the cell value.
    pub value: Option<String>,
    /// Numeric value (only set when value_type is "number").
    pub number_value: Option<f64>,
    /// Boolean value (only set when value_type is "boolean").
    pub bool_value: Option<bool>,
}

/// A column with its name and cell data.
#[napi(object)]
pub struct JsColData {
    /// Column name (e.g., "A", "B", "AA").
    pub column: String,
    /// Cells with data in this column.
    pub cells: Vec<JsColCell>,
}

fn cell_value_to_row_cell(column: String, value: CellValue) -> JsRowCell {
    let (value_type, str_val, num_val, bool_val) = cell_value_to_parts(&value);
    JsRowCell {
        column,
        value_type,
        value: str_val,
        number_value: num_val,
        bool_value: bool_val,
    }
}

fn cell_value_to_col_cell(row: u32, value: CellValue) -> JsColCell {
    let (value_type, str_val, num_val, bool_val) = cell_value_to_parts(&value);
    JsColCell {
        row,
        value_type,
        value: str_val,
        number_value: num_val,
        bool_value: bool_val,
    }
}

fn cell_value_to_parts(value: &CellValue) -> (String, Option<String>, Option<f64>, Option<bool>) {
    match value {
        CellValue::Empty => ("empty".to_string(), None, None, None),
        CellValue::String(s) => ("string".to_string(), Some(s.clone()), None, None),
        CellValue::Number(n) => ("number".to_string(), None, Some(*n), None),
        CellValue::Bool(b) => ("boolean".to_string(), None, None, Some(*b)),
        CellValue::Error(e) => ("error".to_string(), Some(e.clone()), None, None),
        CellValue::Formula { expr, .. } => ("formula".to_string(), Some(expr.clone()), None, None),
        CellValue::Date(serial) => {
            let iso = sheetkit_core::cell::serial_to_datetime(*serial).map(|dt| {
                if serial.fract() == 0.0 {
                    dt.format("%Y-%m-%d").to_string()
                } else {
                    dt.format("%Y-%m-%dT%H:%M:%S").to_string()
                }
            });
            ("date".to_string(), iso, Some(*serial), None)
        }
    }
}

fn parse_hyperlink_type(opts: &JsHyperlinkOptions) -> Result<HyperlinkType> {
    match opts.link_type.to_lowercase().as_str() {
        "external" => Ok(HyperlinkType::External(opts.target.clone())),
        "internal" => Ok(HyperlinkType::Internal(opts.target.clone())),
        "email" => Ok(HyperlinkType::Email(opts.target.clone())),
        _ => Err(Error::from_reason(format!(
            "unknown hyperlink type: {}",
            opts.link_type
        ))),
    }
}

fn hyperlink_info_to_js(info: &HyperlinkInfo) -> JsHyperlinkInfo {
    let (link_type, target) = match &info.link_type {
        HyperlinkType::External(url) => ("external".to_string(), url.clone()),
        HyperlinkType::Internal(loc) => ("internal".to_string(), loc.clone()),
        HyperlinkType::Email(email) => ("email".to_string(), email.clone()),
    };
    JsHyperlinkInfo {
        link_type,
        target,
        display: info.display.clone(),
        tooltip: info.tooltip.clone(),
    }
}

fn js_to_cell_value(value: JsUnknown) -> Result<CellValue> {
    match value.get_type()? {
        ValueType::Null | ValueType::Undefined => Ok(CellValue::Empty),
        ValueType::Boolean => {
            let b = value.coerce_to_bool()?.get_value()?;
            Ok(CellValue::Bool(b))
        }
        ValueType::Number => {
            let n = value.coerce_to_number()?.get_double()?;
            Ok(CellValue::Number(n))
        }
        ValueType::String => {
            let s = value.coerce_to_string()?.into_utf8()?.as_str()?.to_string();
            Ok(CellValue::String(s))
        }
        ValueType::Object => {
            // Accept { type: "date", serial: number } objects for date values.
            let obj = value.coerce_to_object()?;
            let type_prop: JsUnknown = obj.get_named_property("type")?;
            if type_prop.get_type()? == ValueType::String {
                let type_str = type_prop
                    .coerce_to_string()?
                    .into_utf8()?
                    .as_str()?
                    .to_string();
                if type_str == "date" {
                    let serial_prop: JsUnknown = obj.get_named_property("serial")?;
                    let serial = serial_prop.coerce_to_number()?.get_double()?;
                    return Ok(CellValue::Date(serial));
                }
            }
            Err(Error::from_reason("unsupported cell value type"))
        }
        _ => Err(Error::from_reason("unsupported cell value type")),
    }
}

fn cell_value_to_js(env: Env, value: CellValue) -> Result<JsUnknown> {
    match value {
        CellValue::Empty => env.get_null().map(|v| v.into_unknown()),
        CellValue::Bool(b) => env.get_boolean(b).map(|v| v.into_unknown()),
        CellValue::Number(n) => env.create_double(n).map(|v| v.into_unknown()),
        CellValue::Date(serial) => {
            // Return dates as an object with type and serial number so JS
            // callers can distinguish dates from plain numbers.
            let mut obj = env.create_object()?;
            obj.set("type", env.create_string("date")?)?;
            obj.set("serial", env.create_double(serial)?)?;
            // Also provide an ISO string when possible.
            if let Some(dt) = sheetkit_core::cell::serial_to_datetime(serial) {
                let iso = if serial.fract() == 0.0 {
                    dt.format("%Y-%m-%d").to_string()
                } else {
                    dt.format("%Y-%m-%dT%H:%M:%S").to_string()
                };
                obj.set("iso", env.create_string(&iso)?)?;
            }
            Ok(obj.into_unknown())
        }
        CellValue::String(s) => env.create_string(&s).map(|v| v.into_unknown()),
        CellValue::Formula { expr, .. } => env.create_string(&expr).map(|v| v.into_unknown()),
        CellValue::Error(e) => env.create_string(&e).map(|v| v.into_unknown()),
    }
}

fn parse_style_color(s: &str) -> Option<StyleColor> {
    if s.starts_with('#') && s.len() == 7 {
        Some(StyleColor::Rgb(s.to_string()))
    } else if let Some(theme_str) = s.strip_prefix("theme:") {
        theme_str.parse::<u32>().ok().map(StyleColor::Theme)
    } else if let Some(indexed_str) = s.strip_prefix("indexed:") {
        indexed_str.parse::<u32>().ok().map(StyleColor::Indexed)
    } else {
        None
    }
}

fn parse_pattern_type(s: &str) -> PatternType {
    match s.to_lowercase().as_str() {
        "none" => PatternType::None,
        "solid" => PatternType::Solid,
        "gray125" => PatternType::Gray125,
        "darkgray" => PatternType::DarkGray,
        "mediumgray" => PatternType::MediumGray,
        "lightgray" => PatternType::LightGray,
        _ => PatternType::None,
    }
}

fn parse_border_line_style(s: &str) -> BorderLineStyle {
    match s.to_lowercase().as_str() {
        "thin" => BorderLineStyle::Thin,
        "medium" => BorderLineStyle::Medium,
        "thick" => BorderLineStyle::Thick,
        "dashed" => BorderLineStyle::Dashed,
        "dotted" => BorderLineStyle::Dotted,
        "double" => BorderLineStyle::Double,
        "hair" => BorderLineStyle::Hair,
        "mediumdashed" => BorderLineStyle::MediumDashed,
        "dashdot" => BorderLineStyle::DashDot,
        "mediumdashdot" => BorderLineStyle::MediumDashDot,
        "dashdotdot" => BorderLineStyle::DashDotDot,
        "mediumdashdotdot" => BorderLineStyle::MediumDashDotDot,
        "slantdashdot" => BorderLineStyle::SlantDashDot,
        _ => BorderLineStyle::Thin,
    }
}

fn parse_horizontal_align(s: &str) -> HorizontalAlign {
    match s.to_lowercase().as_str() {
        "general" => HorizontalAlign::General,
        "left" => HorizontalAlign::Left,
        "center" => HorizontalAlign::Center,
        "right" => HorizontalAlign::Right,
        "fill" => HorizontalAlign::Fill,
        "justify" => HorizontalAlign::Justify,
        "centercontinuous" => HorizontalAlign::CenterContinuous,
        "distributed" => HorizontalAlign::Distributed,
        _ => HorizontalAlign::General,
    }
}

fn parse_vertical_align(s: &str) -> VerticalAlign {
    match s.to_lowercase().as_str() {
        "top" => VerticalAlign::Top,
        "center" => VerticalAlign::Center,
        "bottom" => VerticalAlign::Bottom,
        "justify" => VerticalAlign::Justify,
        "distributed" => VerticalAlign::Distributed,
        _ => VerticalAlign::Bottom,
    }
}

fn js_style_to_core(js: &JsStyle) -> Style {
    Style {
        font: js.font.as_ref().map(|f| FontStyle {
            name: f.name.clone(),
            size: f.size,
            bold: f.bold.unwrap_or(false),
            italic: f.italic.unwrap_or(false),
            underline: f.underline.unwrap_or(false),
            strikethrough: f.strikethrough.unwrap_or(false),
            color: f.color.as_ref().and_then(|s| parse_style_color(s)),
        }),
        fill: js.fill.as_ref().map(|f| FillStyle {
            pattern: f
                .pattern
                .as_ref()
                .map(|s| parse_pattern_type(s))
                .unwrap_or(PatternType::None),
            fg_color: f.fg_color.as_ref().and_then(|s| parse_style_color(s)),
            bg_color: f.bg_color.as_ref().and_then(|s| parse_style_color(s)),
        }),
        border: js.border.as_ref().map(|b| {
            let side = |s: &JsBorderSideStyle| BorderSideStyle {
                style: s
                    .style
                    .as_ref()
                    .map(|s| parse_border_line_style(s))
                    .unwrap_or(BorderLineStyle::Thin),
                color: s.color.as_ref().and_then(|s| parse_style_color(s)),
            };
            BorderStyle {
                left: b.left.as_ref().map(&side),
                right: b.right.as_ref().map(&side),
                top: b.top.as_ref().map(&side),
                bottom: b.bottom.as_ref().map(&side),
                diagonal: b.diagonal.as_ref().map(&side),
            }
        }),
        alignment: js.alignment.as_ref().map(|a| AlignmentStyle {
            horizontal: a.horizontal.as_ref().map(|s| parse_horizontal_align(s)),
            vertical: a.vertical.as_ref().map(|s| parse_vertical_align(s)),
            wrap_text: a.wrap_text.unwrap_or(false),
            text_rotation: a.text_rotation,
            indent: a.indent,
            shrink_to_fit: a.shrink_to_fit.unwrap_or(false),
        }),
        num_fmt: if let Some(custom) = &js.custom_num_fmt {
            Some(NumFmtStyle::Custom(custom.clone()))
        } else {
            js.num_fmt_id.map(NumFmtStyle::Builtin)
        },
        protection: js.protection.as_ref().map(|p| ProtectionStyle {
            locked: p.locked.unwrap_or(true),
            hidden: p.hidden.unwrap_or(false),
        }),
    }
}

fn parse_chart_type(s: &str) -> ChartType {
    match s.to_lowercase().as_str() {
        "col" => ChartType::Col,
        "colstacked" => ChartType::ColStacked,
        "colpercentstacked" => ChartType::ColPercentStacked,
        "bar" => ChartType::Bar,
        "barstacked" => ChartType::BarStacked,
        "barpercentstacked" => ChartType::BarPercentStacked,
        "line" => ChartType::Line,
        "pie" => ChartType::Pie,
        _ => ChartType::Col,
    }
}

fn parse_image_format(s: &str) -> Result<ImageFormat> {
    match s.to_lowercase().as_str() {
        "png" => Ok(ImageFormat::Png),
        "jpeg" | "jpg" => Ok(ImageFormat::Jpeg),
        "gif" => Ok(ImageFormat::Gif),
        _ => Err(Error::from_reason(format!("unknown image format: {s}"))),
    }
}

fn parse_validation_type(s: &str) -> ValidationType {
    match s.to_lowercase().as_str() {
        "whole" => ValidationType::Whole,
        "decimal" => ValidationType::Decimal,
        "list" => ValidationType::List,
        "date" => ValidationType::Date,
        "time" => ValidationType::Time,
        "textlength" => ValidationType::TextLength,
        "custom" => ValidationType::Custom,
        _ => ValidationType::List,
    }
}

fn parse_validation_operator(s: &str) -> Option<ValidationOperator> {
    match s.to_lowercase().as_str() {
        "between" => Some(ValidationOperator::Between),
        "notbetween" => Some(ValidationOperator::NotBetween),
        "equal" => Some(ValidationOperator::Equal),
        "notequal" => Some(ValidationOperator::NotEqual),
        "lessthan" => Some(ValidationOperator::LessThan),
        "lessthanorequal" => Some(ValidationOperator::LessThanOrEqual),
        "greaterthan" => Some(ValidationOperator::GreaterThan),
        "greaterthanorequal" => Some(ValidationOperator::GreaterThanOrEqual),
        _ => None,
    }
}

fn parse_error_style(s: &str) -> Option<ErrorStyle> {
    match s.to_lowercase().as_str() {
        "stop" => Some(ErrorStyle::Stop),
        "warning" => Some(ErrorStyle::Warning),
        "information" => Some(ErrorStyle::Information),
        _ => None,
    }
}

fn validation_type_to_string(vt: &ValidationType) -> String {
    match vt {
        ValidationType::Whole => "whole".to_string(),
        ValidationType::Decimal => "decimal".to_string(),
        ValidationType::List => "list".to_string(),
        ValidationType::Date => "date".to_string(),
        ValidationType::Time => "time".to_string(),
        ValidationType::TextLength => "textlength".to_string(),
        ValidationType::Custom => "custom".to_string(),
    }
}

fn validation_operator_to_string(vo: &ValidationOperator) -> String {
    match vo {
        ValidationOperator::Between => "between".to_string(),
        ValidationOperator::NotBetween => "notbetween".to_string(),
        ValidationOperator::Equal => "equal".to_string(),
        ValidationOperator::NotEqual => "notequal".to_string(),
        ValidationOperator::LessThan => "lessthan".to_string(),
        ValidationOperator::LessThanOrEqual => "lessthanorequal".to_string(),
        ValidationOperator::GreaterThan => "greaterthan".to_string(),
        ValidationOperator::GreaterThanOrEqual => "greaterthanorequal".to_string(),
    }
}

fn error_style_to_string(es: &ErrorStyle) -> String {
    match es {
        ErrorStyle::Stop => "stop".to_string(),
        ErrorStyle::Warning => "warning".to_string(),
        ErrorStyle::Information => "information".to_string(),
    }
}

fn core_validation_to_js(v: &DataValidationConfig) -> JsDataValidationConfig {
    JsDataValidationConfig {
        sqref: v.sqref.clone(),
        validation_type: validation_type_to_string(&v.validation_type),
        operator: v.operator.as_ref().map(validation_operator_to_string),
        formula1: v.formula1.clone(),
        formula2: v.formula2.clone(),
        allow_blank: Some(v.allow_blank),
        error_style: v.error_style.as_ref().map(error_style_to_string),
        error_title: v.error_title.clone(),
        error_message: v.error_message.clone(),
        prompt_title: v.prompt_title.clone(),
        prompt_message: v.prompt_message.clone(),
        show_input_message: Some(v.show_input_message),
        show_error_message: Some(v.show_error_message),
    }
}

fn js_doc_props_to_core(js: &JsDocProperties) -> DocProperties {
    DocProperties {
        title: js.title.clone(),
        subject: js.subject.clone(),
        creator: js.creator.clone(),
        keywords: js.keywords.clone(),
        description: js.description.clone(),
        last_modified_by: js.last_modified_by.clone(),
        revision: js.revision.clone(),
        created: js.created.clone(),
        modified: js.modified.clone(),
        category: js.category.clone(),
        content_status: js.content_status.clone(),
    }
}

fn core_doc_props_to_js(props: &DocProperties) -> JsDocProperties {
    JsDocProperties {
        title: props.title.clone(),
        subject: props.subject.clone(),
        creator: props.creator.clone(),
        keywords: props.keywords.clone(),
        description: props.description.clone(),
        last_modified_by: props.last_modified_by.clone(),
        revision: props.revision.clone(),
        created: props.created.clone(),
        modified: props.modified.clone(),
        category: props.category.clone(),
        content_status: props.content_status.clone(),
    }
}

fn js_app_props_to_core(js: &JsAppProperties) -> AppProperties {
    AppProperties {
        application: js.application.clone(),
        doc_security: js.doc_security,
        company: js.company.clone(),
        app_version: js.app_version.clone(),
        manager: js.manager.clone(),
        template: js.template.clone(),
    }
}

fn core_app_props_to_js(props: &AppProperties) -> JsAppProperties {
    JsAppProperties {
        application: props.application.clone(),
        doc_security: props.doc_security,
        company: props.company.clone(),
        app_version: props.app_version.clone(),
        manager: props.manager.clone(),
        template: props.template.clone(),
    }
}

fn parse_cf_operator(s: &str) -> Option<CfOperator> {
    match s.to_lowercase().as_str() {
        "lessthan" => Some(CfOperator::LessThan),
        "lessthanorequal" => Some(CfOperator::LessThanOrEqual),
        "equal" => Some(CfOperator::Equal),
        "notequal" => Some(CfOperator::NotEqual),
        "greaterthanorequal" => Some(CfOperator::GreaterThanOrEqual),
        "greaterthan" => Some(CfOperator::GreaterThan),
        "between" => Some(CfOperator::Between),
        "notbetween" => Some(CfOperator::NotBetween),
        _ => None,
    }
}

fn cf_operator_to_string(op: &CfOperator) -> String {
    match op {
        CfOperator::LessThan => "lessThan".to_string(),
        CfOperator::LessThanOrEqual => "lessThanOrEqual".to_string(),
        CfOperator::Equal => "equal".to_string(),
        CfOperator::NotEqual => "notEqual".to_string(),
        CfOperator::GreaterThanOrEqual => "greaterThanOrEqual".to_string(),
        CfOperator::GreaterThan => "greaterThan".to_string(),
        CfOperator::Between => "between".to_string(),
        CfOperator::NotBetween => "notBetween".to_string(),
    }
}

fn parse_cf_value_type(s: &str) -> CfValueType {
    match s.to_lowercase().as_str() {
        "num" => CfValueType::Num,
        "percent" => CfValueType::Percent,
        "min" => CfValueType::Min,
        "max" => CfValueType::Max,
        "percentile" => CfValueType::Percentile,
        "formula" => CfValueType::Formula,
        _ => CfValueType::Num,
    }
}

fn cf_value_type_to_string(vt: &CfValueType) -> String {
    match vt {
        CfValueType::Num => "num".to_string(),
        CfValueType::Percent => "percent".to_string(),
        CfValueType::Min => "min".to_string(),
        CfValueType::Max => "max".to_string(),
        CfValueType::Percentile => "percentile".to_string(),
        CfValueType::Formula => "formula".to_string(),
    }
}

fn js_conditional_style_to_core(js: &JsConditionalStyle) -> ConditionalStyle {
    ConditionalStyle {
        font: js.font.as_ref().map(|f| FontStyle {
            name: f.name.clone(),
            size: f.size,
            bold: f.bold.unwrap_or(false),
            italic: f.italic.unwrap_or(false),
            underline: f.underline.unwrap_or(false),
            strikethrough: f.strikethrough.unwrap_or(false),
            color: f.color.as_ref().and_then(|s| parse_style_color(s)),
        }),
        fill: js.fill.as_ref().map(|f| FillStyle {
            pattern: f
                .pattern
                .as_ref()
                .map(|s| parse_pattern_type(s))
                .unwrap_or(PatternType::None),
            fg_color: f.fg_color.as_ref().and_then(|s| parse_style_color(s)),
            bg_color: f.bg_color.as_ref().and_then(|s| parse_style_color(s)),
        }),
        border: js.border.as_ref().map(|b| {
            let side = |s: &JsBorderSideStyle| BorderSideStyle {
                style: s
                    .style
                    .as_ref()
                    .map(|s| parse_border_line_style(s))
                    .unwrap_or(BorderLineStyle::Thin),
                color: s.color.as_ref().and_then(|s| parse_style_color(s)),
            };
            BorderStyle {
                left: b.left.as_ref().map(&side),
                right: b.right.as_ref().map(&side),
                top: b.top.as_ref().map(&side),
                bottom: b.bottom.as_ref().map(&side),
                diagonal: b.diagonal.as_ref().map(&side),
            }
        }),
        num_fmt: js
            .custom_num_fmt
            .as_ref()
            .map(|s| NumFmtStyle::Custom(s.clone())),
    }
}

fn js_cf_rule_to_core(js: &JsConditionalFormatRule) -> Result<ConditionalFormatRule> {
    let rule_type = match js.rule_type.as_str() {
        "cellIs" => {
            let operator = js
                .operator
                .as_ref()
                .and_then(|s| parse_cf_operator(s))
                .unwrap_or(CfOperator::Equal);
            ConditionalFormatType::CellIs {
                operator,
                formula: js.formula.clone().unwrap_or_default(),
                formula2: js.formula2.clone(),
            }
        }
        "expression" => ConditionalFormatType::Expression {
            formula: js.formula.clone().unwrap_or_default(),
        },
        "colorScale" => ConditionalFormatType::ColorScale {
            min_type: js
                .min_type
                .as_ref()
                .map(|s| parse_cf_value_type(s))
                .unwrap_or(CfValueType::Min),
            min_value: js.min_value.clone(),
            min_color: js.min_color.clone().unwrap_or_default(),
            mid_type: js.mid_type.as_ref().map(|s| parse_cf_value_type(s)),
            mid_value: js.mid_value.clone(),
            mid_color: js.mid_color.clone(),
            max_type: js
                .max_type
                .as_ref()
                .map(|s| parse_cf_value_type(s))
                .unwrap_or(CfValueType::Max),
            max_value: js.max_value.clone(),
            max_color: js.max_color.clone().unwrap_or_default(),
        },
        "dataBar" => ConditionalFormatType::DataBar {
            min_type: js
                .min_type
                .as_ref()
                .map(|s| parse_cf_value_type(s))
                .unwrap_or(CfValueType::Min),
            min_value: js.min_value.clone(),
            max_type: js
                .max_type
                .as_ref()
                .map(|s| parse_cf_value_type(s))
                .unwrap_or(CfValueType::Max),
            max_value: js.max_value.clone(),
            color: js.bar_color.clone().unwrap_or_default(),
            show_value: js.show_value.unwrap_or(true),
        },
        "duplicateValues" => ConditionalFormatType::DuplicateValues,
        "uniqueValues" => ConditionalFormatType::UniqueValues,
        "top10" => ConditionalFormatType::Top10 {
            rank: js.rank.unwrap_or(10),
            percent: js.percent.unwrap_or(false),
        },
        "bottom10" => ConditionalFormatType::Bottom10 {
            rank: js.rank.unwrap_or(10),
            percent: js.percent.unwrap_or(false),
        },
        "aboveAverage" => ConditionalFormatType::AboveAverage {
            above: js.above.unwrap_or(true),
            equal_average: js.equal_average.unwrap_or(false),
        },
        "containsBlanks" => ConditionalFormatType::ContainsBlanks,
        "notContainsBlanks" => ConditionalFormatType::NotContainsBlanks,
        "containsErrors" => ConditionalFormatType::ContainsErrors,
        "notContainsErrors" => ConditionalFormatType::NotContainsErrors,
        "containsText" => ConditionalFormatType::ContainsText {
            text: js.text.clone().unwrap_or_default(),
        },
        "notContainsText" => ConditionalFormatType::NotContainsText {
            text: js.text.clone().unwrap_or_default(),
        },
        "beginsWith" => ConditionalFormatType::BeginsWith {
            text: js.text.clone().unwrap_or_default(),
        },
        "endsWith" => ConditionalFormatType::EndsWith {
            text: js.text.clone().unwrap_or_default(),
        },
        other => {
            return Err(Error::from_reason(format!(
                "unknown conditional format rule type: {other}"
            )));
        }
    };

    let format = js.format.as_ref().map(js_conditional_style_to_core);

    Ok(ConditionalFormatRule {
        rule_type,
        format,
        priority: js.priority,
        stop_if_true: js.stop_if_true.unwrap_or(false),
    })
}

fn core_cf_rule_to_js(rule: &ConditionalFormatRule) -> JsConditionalFormatRule {
    let (rule_type, operator, formula, formula2, text, rank, percent, above, equal_average) =
        match &rule.rule_type {
            ConditionalFormatType::CellIs {
                operator,
                formula,
                formula2,
            } => (
                "cellIs".to_string(),
                Some(cf_operator_to_string(operator)),
                Some(formula.clone()),
                formula2.clone(),
                None,
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::Expression { formula } => (
                "expression".to_string(),
                None,
                Some(formula.clone()),
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::ColorScale { .. } => (
                "colorScale".to_string(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::DataBar { .. } => (
                "dataBar".to_string(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::DuplicateValues => (
                "duplicateValues".to_string(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::UniqueValues => (
                "uniqueValues".to_string(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::Top10 { rank, percent } => (
                "top10".to_string(),
                None,
                None,
                None,
                None,
                Some(*rank),
                Some(*percent),
                None,
                None,
            ),
            ConditionalFormatType::Bottom10 { rank, percent } => (
                "bottom10".to_string(),
                None,
                None,
                None,
                None,
                Some(*rank),
                Some(*percent),
                None,
                None,
            ),
            ConditionalFormatType::AboveAverage {
                above,
                equal_average,
            } => (
                "aboveAverage".to_string(),
                None,
                None,
                None,
                None,
                None,
                None,
                Some(*above),
                Some(*equal_average),
            ),
            ConditionalFormatType::ContainsBlanks => (
                "containsBlanks".to_string(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::NotContainsBlanks => (
                "notContainsBlanks".to_string(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::ContainsErrors => (
                "containsErrors".to_string(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::NotContainsErrors => (
                "notContainsErrors".to_string(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::ContainsText { text } => (
                "containsText".to_string(),
                None,
                None,
                None,
                Some(text.clone()),
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::NotContainsText { text } => (
                "notContainsText".to_string(),
                None,
                None,
                None,
                Some(text.clone()),
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::BeginsWith { text } => (
                "beginsWith".to_string(),
                None,
                None,
                None,
                Some(text.clone()),
                None,
                None,
                None,
                None,
            ),
            ConditionalFormatType::EndsWith { text } => (
                "endsWith".to_string(),
                None,
                None,
                None,
                Some(text.clone()),
                None,
                None,
                None,
                None,
            ),
        };

    let (
        min_type,
        min_value,
        min_color,
        mid_type,
        mid_value,
        mid_color,
        max_type,
        max_value,
        max_color,
        bar_color,
        show_value,
    ) = match &rule.rule_type {
        ConditionalFormatType::ColorScale {
            min_type,
            min_value,
            min_color,
            mid_type,
            mid_value,
            mid_color,
            max_type,
            max_value,
            max_color,
        } => (
            Some(cf_value_type_to_string(min_type)),
            min_value.clone(),
            Some(min_color.clone()),
            mid_type.as_ref().map(cf_value_type_to_string),
            mid_value.clone(),
            mid_color.clone(),
            Some(cf_value_type_to_string(max_type)),
            max_value.clone(),
            Some(max_color.clone()),
            None,
            None,
        ),
        ConditionalFormatType::DataBar {
            min_type,
            min_value,
            max_type,
            max_value,
            color,
            show_value,
        } => (
            Some(cf_value_type_to_string(min_type)),
            min_value.clone(),
            None,
            None,
            None,
            None,
            Some(cf_value_type_to_string(max_type)),
            max_value.clone(),
            None,
            Some(color.clone()),
            Some(*show_value),
        ),
        _ => (
            None, None, None, None, None, None, None, None, None, None, None,
        ),
    };

    let format = rule.format.as_ref().map(|s| {
        JsConditionalStyle {
            font: s.font.as_ref().map(|f| JsFontStyle {
                name: f.name.clone(),
                size: f.size,
                bold: if f.bold { Some(true) } else { None },
                italic: if f.italic { Some(true) } else { None },
                underline: if f.underline { Some(true) } else { None },
                strikethrough: if f.strikethrough { Some(true) } else { None },
                color: f.color.as_ref().map(|c| match c {
                    StyleColor::Rgb(rgb) => rgb.clone(),
                    StyleColor::Theme(t) => format!("theme:{t}"),
                    StyleColor::Indexed(i) => format!("indexed:{i}"),
                }),
            }),
            fill: s.fill.as_ref().map(|f| JsFillStyle {
                pattern: Some(match f.pattern {
                    PatternType::None => "none".to_string(),
                    PatternType::Solid => "solid".to_string(),
                    PatternType::Gray125 => "gray125".to_string(),
                    PatternType::DarkGray => "darkGray".to_string(),
                    PatternType::MediumGray => "mediumGray".to_string(),
                    PatternType::LightGray => "lightGray".to_string(),
                }),
                fg_color: f.fg_color.as_ref().map(|c| match c {
                    StyleColor::Rgb(rgb) => rgb.clone(),
                    StyleColor::Theme(t) => format!("theme:{t}"),
                    StyleColor::Indexed(i) => format!("indexed:{i}"),
                }),
                bg_color: f.bg_color.as_ref().map(|c| match c {
                    StyleColor::Rgb(rgb) => rgb.clone(),
                    StyleColor::Theme(t) => format!("theme:{t}"),
                    StyleColor::Indexed(i) => format!("indexed:{i}"),
                }),
            }),
            border: None, // Simplified: border conversion omitted for now
            custom_num_fmt: s.num_fmt.as_ref().and_then(|nf| match nf {
                NumFmtStyle::Custom(code) => Some(code.clone()),
                _ => None,
            }),
        }
    });

    JsConditionalFormatRule {
        rule_type,
        operator,
        formula,
        formula2,
        text,
        rank,
        percent,
        above,
        equal_average,
        min_type,
        min_value,
        min_color,
        mid_type,
        mid_value,
        mid_color,
        max_type,
        max_value,
        max_color,
        bar_color,
        show_value,
        format,
        priority: rule.priority,
        stop_if_true: if rule.stop_if_true { Some(true) } else { None },
    }
}

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
    #[napi(ts_return_type = "string | number | boolean | DateValue | null")]
    pub fn get_cell_value(&self, env: Env, sheet: String, cell: String) -> Result<JsUnknown> {
        let value = self
            .inner
            .get_cell_value(&sheet, &cell)
            .map_err(|e| Error::from_reason(e.to_string()))?;
        cell_value_to_js(env, value)
    }

    /// Set the value of a cell. Pass string, number, boolean, DateValue, or null to clear.
    #[napi(
        ts_args_type = "sheet: string, cell: string, value: string | number | boolean | { type: 'date', serial: number } | null"
    )]
    pub fn set_cell_value(&mut self, sheet: String, cell: String, value: JsUnknown) -> Result<()> {
        let cell_value = js_to_cell_value(value)?;
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
                })
                .collect(),
            show_legend: config.show_legend.unwrap_or(true),
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
    #[napi(ts_args_type = "name: string, value: string | number | boolean")]
    pub fn set_custom_property(&mut self, name: String, value: JsUnknown) -> Result<()> {
        let prop_value = match value.get_type()? {
            ValueType::Boolean => {
                let b = value.coerce_to_bool()?.get_value()?;
                CustomPropertyValue::Bool(b)
            }
            ValueType::Number => {
                let n = value.coerce_to_number()?.get_double()?;
                if n.fract() == 0.0 && n >= i32::MIN as f64 && n <= i32::MAX as f64 {
                    CustomPropertyValue::Int(n as i32)
                } else {
                    CustomPropertyValue::Float(n)
                }
            }
            ValueType::String => {
                let s = value.coerce_to_string()?.into_utf8()?.as_str()?.to_string();
                CustomPropertyValue::String(s)
            }
            _ => return Err(Error::from_reason("unsupported custom property value type")),
        };
        self.inner.set_custom_property(&name, prop_value);
        Ok(())
    }

    /// Get a custom property value, or null if not found.
    #[napi(ts_return_type = "string | number | boolean | null")]
    pub fn get_custom_property(&self, env: Env, name: String) -> Result<JsUnknown> {
        match self.inner.get_custom_property(&name) {
            Some(CustomPropertyValue::String(s)) => env.create_string(&s).map(|v| v.into_unknown()),
            Some(CustomPropertyValue::Int(i)) => env.create_int32(i).map(|v| v.into_unknown()),
            Some(CustomPropertyValue::Float(f)) => env.create_double(f).map(|v| v.into_unknown()),
            Some(CustomPropertyValue::Bool(b)) => env.get_boolean(b).map(|v| v.into_unknown()),
            Some(CustomPropertyValue::DateTime(s)) => {
                env.create_string(&s).map(|v| v.into_unknown())
            }
            None => env.get_null().map(|v| v.into_unknown()),
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
}

/// Forward-only streaming writer for large sheets.
#[derive(Default)]
#[napi]
pub struct JsStreamWriter {
    inner: Option<StreamWriter>,
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
    #[napi(ts_args_type = "row: number, values: Array<string | number | boolean | null>")]
    pub fn write_row(&mut self, row: u32, values: Vec<JsUnknown>) -> Result<()> {
        let writer = self
            .inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("StreamWriter already consumed"))?;
        let cell_values: Vec<CellValue> = values
            .into_iter()
            .map(js_to_cell_value)
            .collect::<Result<Vec<_>>>()?;
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
