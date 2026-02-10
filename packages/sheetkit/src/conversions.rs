use napi::bindgen_prelude::*;

use sheetkit_core::cell::CellValue;
use sheetkit_core::chart::ChartType;
use sheetkit_core::conditional::{
    CfOperator, CfValueType, ConditionalFormatRule, ConditionalFormatType, ConditionalStyle,
};
use sheetkit_core::doc_props::{AppProperties, DocProperties};
use sheetkit_core::hyperlink::{HyperlinkInfo, HyperlinkType};
use sheetkit_core::image::ImageFormat;
use sheetkit_core::page_layout::{Orientation, PaperSize};
use sheetkit_core::pivot::AggregateFunction;
use sheetkit_core::style::{
    AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle, FillStyle, FontStyle,
    HorizontalAlign, NumFmtStyle, PatternType, ProtectionStyle, Style, StyleColor, VerticalAlign,
};
use sheetkit_core::validation::{
    DataValidationConfig, ErrorStyle, ValidationOperator, ValidationType,
};

use crate::types::*;

pub(crate) fn js_value_to_cell_value(
    v: napi::bindgen_prelude::Either5<
        String,
        f64,
        bool,
        crate::types::DateValue,
        napi::bindgen_prelude::Null,
    >,
) -> CellValue {
    match v {
        napi::bindgen_prelude::Either5::A(s) => CellValue::String(s),
        napi::bindgen_prelude::Either5::B(n) => CellValue::Number(n),
        napi::bindgen_prelude::Either5::C(b) => CellValue::Bool(b),
        napi::bindgen_prelude::Either5::D(d) => CellValue::Date(d.serial),
        napi::bindgen_prelude::Either5::E(_) => CellValue::Empty,
    }
}

pub(crate) fn parse_column_name(col: &str) -> napi::Result<u32> {
    sheetkit_core::utils::cell_ref::column_name_to_number(col)
        .map_err(|e| napi::Error::from_reason(e.to_string()))
}

pub(crate) fn parse_paper_size(s: &str) -> Option<PaperSize> {
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

pub(crate) fn paper_size_to_string(ps: &PaperSize) -> String {
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

pub(crate) fn parse_orientation(s: &str) -> Option<Orientation> {
    match s.to_lowercase().as_str() {
        "portrait" => Some(Orientation::Portrait),
        "landscape" => Some(Orientation::Landscape),
        _ => None,
    }
}

pub(crate) fn orientation_to_string(o: &Orientation) -> String {
    match o {
        Orientation::Portrait => "portrait".to_string(),
        Orientation::Landscape => "landscape".to_string(),
    }
}

pub(crate) fn cell_value_to_row_cell(column: String, value: CellValue) -> JsRowCell {
    let (value_type, str_val, num_val, bool_val) = cell_value_to_parts(&value);
    JsRowCell {
        column,
        value_type,
        value: str_val,
        number_value: num_val,
        bool_value: bool_val,
    }
}

pub(crate) fn cell_value_to_col_cell(row: u32, value: CellValue) -> JsColCell {
    let (value_type, str_val, num_val, bool_val) = cell_value_to_parts(&value);
    JsColCell {
        row,
        value_type,
        value: str_val,
        number_value: num_val,
        bool_value: bool_val,
    }
}

pub(crate) fn cell_value_to_parts(
    value: &CellValue,
) -> (String, Option<String>, Option<f64>, Option<bool>) {
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
        CellValue::RichString(runs) => {
            let plain = sheetkit_core::rich_text::rich_text_to_plain(runs);
            ("string".to_string(), Some(plain), None, None)
        }
    }
}

pub(crate) fn parse_hyperlink_type(opts: &JsHyperlinkOptions) -> Result<HyperlinkType> {
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

pub(crate) fn hyperlink_info_to_js(info: &HyperlinkInfo) -> JsHyperlinkInfo {
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

pub(crate) fn cell_value_to_either(
    value: CellValue,
) -> Result<Either5<Null, bool, f64, String, DateValue>> {
    Ok(match value {
        CellValue::Empty => Either5::A(Null),
        CellValue::Bool(b) => Either5::B(b),
        CellValue::Number(n) => Either5::C(n),
        CellValue::String(s) => Either5::D(s),
        CellValue::Date(serial) => Either5::E(DateValue {
            kind: "date".to_string(),
            serial,
            iso: sheetkit_core::cell::serial_to_datetime(serial).map(|dt| {
                if serial.fract() == 0.0 {
                    dt.format("%Y-%m-%d").to_string()
                } else {
                    dt.format("%Y-%m-%dT%H:%M:%S").to_string()
                }
            }),
        }),
        CellValue::Formula { expr, .. } => Either5::D(expr),
        CellValue::Error(e) => Either5::D(e),
        CellValue::RichString(runs) => {
            Either5::D(sheetkit_core::rich_text::rich_text_to_plain(&runs))
        }
    })
}

pub(crate) fn parse_style_color(s: &str) -> Option<StyleColor> {
    if s.starts_with('#') && s.len() == 7 {
        // #RRGGBB format (stored as-is)
        Some(StyleColor::Rgb(s.to_string()))
    } else if s.len() == 8 && s.chars().all(|c| c.is_ascii_hexdigit()) {
        // AARRGGBB format (e.g. "FFFFFF00")
        Some(StyleColor::Rgb(s.to_string()))
    } else if s.len() == 6 && s.chars().all(|c| c.is_ascii_hexdigit()) {
        // RRGGBB format (e.g. "FF0000")
        Some(StyleColor::Rgb(s.to_string()))
    } else if let Some(theme_str) = s.strip_prefix("theme:") {
        theme_str.parse::<u32>().ok().map(StyleColor::Theme)
    } else if let Some(indexed_str) = s.strip_prefix("indexed:") {
        indexed_str.parse::<u32>().ok().map(StyleColor::Indexed)
    } else {
        None
    }
}

pub(crate) fn parse_pattern_type(s: &str) -> PatternType {
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

pub(crate) fn parse_border_line_style(s: &str) -> BorderLineStyle {
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

pub(crate) fn parse_horizontal_align(s: &str) -> HorizontalAlign {
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

pub(crate) fn parse_vertical_align(s: &str) -> VerticalAlign {
    match s.to_lowercase().as_str() {
        "top" => VerticalAlign::Top,
        "center" => VerticalAlign::Center,
        "bottom" => VerticalAlign::Bottom,
        "justify" => VerticalAlign::Justify,
        "distributed" => VerticalAlign::Distributed,
        _ => VerticalAlign::Bottom,
    }
}

pub(crate) fn js_style_to_core(js: &JsStyle) -> Style {
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
            gradient: None,
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

pub(crate) fn parse_chart_type(s: &str) -> Result<ChartType> {
    let chart_type = match s.to_lowercase().as_str() {
        "col" => ChartType::Col,
        "colstacked" => ChartType::ColStacked,
        "colpercentstacked" => ChartType::ColPercentStacked,
        "col3d" => ChartType::Col3D,
        "col3dstacked" => ChartType::Col3DStacked,
        "col3dpercentstacked" => ChartType::Col3DPercentStacked,
        "bar" => ChartType::Bar,
        "barstacked" => ChartType::BarStacked,
        "barpercentstacked" => ChartType::BarPercentStacked,
        "bar3d" => ChartType::Bar3D,
        "bar3dstacked" => ChartType::Bar3DStacked,
        "bar3dpercentstacked" => ChartType::Bar3DPercentStacked,
        "line" => ChartType::Line,
        "linestacked" => ChartType::LineStacked,
        "linepercentstacked" => ChartType::LinePercentStacked,
        "line3d" => ChartType::Line3D,
        "pie" => ChartType::Pie,
        "pie3d" => ChartType::Pie3D,
        "doughnut" => ChartType::Doughnut,
        "area" => ChartType::Area,
        "areastacked" => ChartType::AreaStacked,
        "areapercentstacked" => ChartType::AreaPercentStacked,
        "area3d" => ChartType::Area3D,
        "area3dstacked" => ChartType::Area3DStacked,
        "area3dpercentstacked" => ChartType::Area3DPercentStacked,
        "scatter" => ChartType::Scatter,
        "scattersmooth" => ChartType::ScatterSmooth,
        "scatterstraight" | "scatterline" => ChartType::ScatterLine,
        "radar" => ChartType::Radar,
        "radarfilled" => ChartType::RadarFilled,
        "radarmarker" => ChartType::RadarMarker,
        "stockhlc" => ChartType::StockHLC,
        "stockohlc" => ChartType::StockOHLC,
        "stockvhlc" => ChartType::StockVHLC,
        "stockvohlc" => ChartType::StockVOHLC,
        "bubble" => ChartType::Bubble,
        "surface" => ChartType::Surface,
        "surfacetop" | "surface3d" => ChartType::Surface3D,
        "surfacewireframe" => ChartType::SurfaceWireframe,
        "surfacetopwireframe" | "surfacewireframe3d" => ChartType::SurfaceWireframe3D,
        "colline" => ChartType::ColLine,
        "collinestacked" => ChartType::ColLineStacked,
        "collinepercentstacked" => ChartType::ColLinePercentStacked,
        "pieofpie" => ChartType::PieOfPie,
        "barofpie" => ChartType::BarOfPie,
        "col3dcone" => ChartType::Col3DCone,
        "col3dconestacked" => ChartType::Col3DConeStacked,
        "col3dconepercentstacked" => ChartType::Col3DConePercentStacked,
        "col3dpyramid" => ChartType::Col3DPyramid,
        "col3dpyramidstacked" => ChartType::Col3DPyramidStacked,
        "col3dpyramidpercentstacked" => ChartType::Col3DPyramidPercentStacked,
        "col3dcylinder" => ChartType::Col3DCylinder,
        "col3dcylinderstacked" => ChartType::Col3DCylinderStacked,
        "col3dcylinderpercentstacked" => ChartType::Col3DCylinderPercentStacked,
        "contour" => ChartType::Contour,
        "wireframecontour" => ChartType::WireframeContour,
        "bubble3d" => ChartType::Bubble3D,
        _ => {
            return Err(Error::from_reason(format!("unknown chart type: {s}")));
        }
    };
    Ok(chart_type)
}

pub(crate) fn parse_shape_type(s: &str) -> Result<sheetkit_core::shape::ShapeType> {
    sheetkit_core::shape::ShapeType::parse(s).map_err(|e| Error::from_reason(e.to_string()))
}

pub(crate) fn parse_image_format(s: &str) -> Result<ImageFormat> {
    ImageFormat::from_extension(s).map_err(|e| Error::from_reason(e.to_string()))
}

pub(crate) fn parse_validation_type(s: &str) -> Result<ValidationType> {
    let validation_type = match s.to_lowercase().as_str() {
        "none" => ValidationType::None,
        "whole" => ValidationType::Whole,
        "decimal" => ValidationType::Decimal,
        "list" => ValidationType::List,
        "date" => ValidationType::Date,
        "time" => ValidationType::Time,
        "textlength" => ValidationType::TextLength,
        "custom" => ValidationType::Custom,
        _ => {
            return Err(Error::from_reason(format!("unknown validation type: {s}")));
        }
    };
    Ok(validation_type)
}

pub(crate) fn parse_validation_operator(s: &str) -> Option<ValidationOperator> {
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

pub(crate) fn parse_error_style(s: &str) -> Option<ErrorStyle> {
    match s.to_lowercase().as_str() {
        "stop" => Some(ErrorStyle::Stop),
        "warning" => Some(ErrorStyle::Warning),
        "information" => Some(ErrorStyle::Information),
        _ => None,
    }
}

pub(crate) fn validation_type_to_string(vt: &ValidationType) -> String {
    match vt {
        ValidationType::None => "none".to_string(),
        ValidationType::Whole => "whole".to_string(),
        ValidationType::Decimal => "decimal".to_string(),
        ValidationType::List => "list".to_string(),
        ValidationType::Date => "date".to_string(),
        ValidationType::Time => "time".to_string(),
        ValidationType::TextLength => "textLength".to_string(),
        ValidationType::Custom => "custom".to_string(),
    }
}

pub(crate) fn validation_operator_to_string(vo: &ValidationOperator) -> String {
    match vo {
        ValidationOperator::Between => "between".to_string(),
        ValidationOperator::NotBetween => "notBetween".to_string(),
        ValidationOperator::Equal => "equal".to_string(),
        ValidationOperator::NotEqual => "notEqual".to_string(),
        ValidationOperator::LessThan => "lessThan".to_string(),
        ValidationOperator::LessThanOrEqual => "lessThanOrEqual".to_string(),
        ValidationOperator::GreaterThan => "greaterThan".to_string(),
        ValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual".to_string(),
    }
}

pub(crate) fn error_style_to_string(es: &ErrorStyle) -> String {
    match es {
        ErrorStyle::Stop => "stop".to_string(),
        ErrorStyle::Warning => "warning".to_string(),
        ErrorStyle::Information => "information".to_string(),
    }
}

pub(crate) fn core_validation_to_js(v: &DataValidationConfig) -> JsDataValidationConfig {
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

pub(crate) fn js_doc_props_to_core(js: &JsDocProperties) -> DocProperties {
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

pub(crate) fn core_doc_props_to_js(props: &DocProperties) -> JsDocProperties {
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

pub(crate) fn js_app_props_to_core(js: &JsAppProperties) -> AppProperties {
    AppProperties {
        application: js.application.clone(),
        doc_security: js.doc_security,
        company: js.company.clone(),
        app_version: js.app_version.clone(),
        manager: js.manager.clone(),
        template: js.template.clone(),
    }
}

pub(crate) fn core_app_props_to_js(props: &AppProperties) -> JsAppProperties {
    JsAppProperties {
        application: props.application.clone(),
        doc_security: props.doc_security,
        company: props.company.clone(),
        app_version: props.app_version.clone(),
        manager: props.manager.clone(),
        template: props.template.clone(),
    }
}

pub(crate) fn parse_cf_operator(s: &str) -> Option<CfOperator> {
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

pub(crate) fn cf_operator_to_string(op: &CfOperator) -> String {
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

pub(crate) fn parse_cf_value_type(s: &str) -> CfValueType {
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

pub(crate) fn cf_value_type_to_string(vt: &CfValueType) -> String {
    match vt {
        CfValueType::Num => "num".to_string(),
        CfValueType::Percent => "percent".to_string(),
        CfValueType::Min => "min".to_string(),
        CfValueType::Max => "max".to_string(),
        CfValueType::Percentile => "percentile".to_string(),
        CfValueType::Formula => "formula".to_string(),
    }
}

pub(crate) fn js_conditional_style_to_core(js: &JsConditionalStyle) -> ConditionalStyle {
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
            gradient: None,
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

pub(crate) fn js_cf_rule_to_core(js: &JsConditionalFormatRule) -> Result<ConditionalFormatRule> {
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

pub(crate) fn core_cf_rule_to_js(rule: &ConditionalFormatRule) -> JsConditionalFormatRule {
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

    let format = rule.format.as_ref().map(|s| JsConditionalStyle {
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
        border: s.border.as_ref().map(|b| {
            let side_to_js = |side: &BorderSideStyle| JsBorderSideStyle {
                style: Some(match side.style {
                    BorderLineStyle::Thin => "thin".to_string(),
                    BorderLineStyle::Medium => "medium".to_string(),
                    BorderLineStyle::Thick => "thick".to_string(),
                    BorderLineStyle::Dashed => "dashed".to_string(),
                    BorderLineStyle::Dotted => "dotted".to_string(),
                    BorderLineStyle::Double => "double".to_string(),
                    BorderLineStyle::Hair => "hair".to_string(),
                    BorderLineStyle::MediumDashed => "mediumDashed".to_string(),
                    BorderLineStyle::DashDot => "dashDot".to_string(),
                    BorderLineStyle::MediumDashDot => "mediumDashDot".to_string(),
                    BorderLineStyle::DashDotDot => "dashDotDot".to_string(),
                    BorderLineStyle::MediumDashDotDot => "mediumDashDotDot".to_string(),
                    BorderLineStyle::SlantDashDot => "slantDashDot".to_string(),
                }),
                color: side.color.as_ref().map(|c| match c {
                    StyleColor::Rgb(rgb) => rgb.clone(),
                    StyleColor::Theme(t) => format!("theme:{t}"),
                    StyleColor::Indexed(i) => format!("indexed:{i}"),
                }),
            };
            JsBorderStyle {
                left: b.left.as_ref().map(&side_to_js),
                right: b.right.as_ref().map(&side_to_js),
                top: b.top.as_ref().map(&side_to_js),
                bottom: b.bottom.as_ref().map(&side_to_js),
                diagonal: b.diagonal.as_ref().map(&side_to_js),
            }
        }),
        custom_num_fmt: s.num_fmt.as_ref().and_then(|nf| match nf {
            NumFmtStyle::Custom(code) => Some(code.clone()),
            _ => None,
        }),
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

pub(crate) fn parse_aggregate_function(s: &str) -> Result<AggregateFunction> {
    let func = match s.to_lowercase().as_str() {
        "sum" => AggregateFunction::Sum,
        "count" => AggregateFunction::Count,
        "average" => AggregateFunction::Average,
        "max" => AggregateFunction::Max,
        "min" => AggregateFunction::Min,
        "product" => AggregateFunction::Product,
        "countnums" => AggregateFunction::CountNums,
        "stddev" => AggregateFunction::StdDev,
        "stddevp" => AggregateFunction::StdDevP,
        "var" => AggregateFunction::Var,
        "varp" => AggregateFunction::VarP,
        _ => {
            return Err(Error::from_reason(format!(
                "unknown aggregate function: {s}"
            )));
        }
    };
    Ok(func)
}

pub(crate) fn js_sheet_protection_to_core(
    js: &crate::types::JsSheetProtectionConfig,
) -> sheetkit_core::sheet::SheetProtectionConfig {
    sheetkit_core::sheet::SheetProtectionConfig {
        password: js.password.clone(),
        select_locked_cells: js.select_locked_cells.unwrap_or(false),
        select_unlocked_cells: js.select_unlocked_cells.unwrap_or(false),
        format_cells: js.format_cells.unwrap_or(false),
        format_columns: js.format_columns.unwrap_or(false),
        format_rows: js.format_rows.unwrap_or(false),
        insert_columns: js.insert_columns.unwrap_or(false),
        insert_rows: js.insert_rows.unwrap_or(false),
        insert_hyperlinks: js.insert_hyperlinks.unwrap_or(false),
        delete_columns: js.delete_columns.unwrap_or(false),
        delete_rows: js.delete_rows.unwrap_or(false),
        sort: js.sort.unwrap_or(false),
        auto_filter: js.auto_filter.unwrap_or(false),
        pivot_tables: js.pivot_tables.unwrap_or(false),
    }
}

pub(crate) fn js_open_options_to_core(
    js: Option<&crate::types::JsOpenOptions>,
) -> sheetkit_core::workbook::OpenOptions {
    let Some(js) = js else {
        return sheetkit_core::workbook::OpenOptions::default();
    };
    sheetkit_core::workbook::OpenOptions {
        sheet_rows: js.sheet_rows,
        sheets: js.sheets.clone(),
        max_unzip_size: js.max_unzip_size.map(|v| v as u64),
        max_zip_entries: js.max_zip_entries.map(|v| v as usize),
    }
}

pub(crate) fn js_sparkline_to_core(
    js: &crate::types::JsSparklineConfig,
) -> sheetkit_core::sparkline::SparklineConfig {
    use sheetkit_core::sparkline::{SparklineConfig, SparklineType};
    SparklineConfig {
        data_range: js.data_range.clone(),
        location: js.location.clone(),
        sparkline_type: js
            .sparkline_type
            .as_deref()
            .and_then(SparklineType::parse)
            .unwrap_or_default(),
        markers: js.markers.unwrap_or(false),
        high_point: js.high_point.unwrap_or(false),
        low_point: js.low_point.unwrap_or(false),
        first_point: js.first_point.unwrap_or(false),
        last_point: js.last_point.unwrap_or(false),
        negative_points: js.negative_points.unwrap_or(false),
        show_axis: js.show_axis.unwrap_or(false),
        line_weight: js.line_weight,
        style: js.style,
    }
}

pub(crate) fn core_sparkline_to_js(
    config: &sheetkit_core::sparkline::SparklineConfig,
) -> crate::types::JsSparklineConfig {
    use sheetkit_core::sparkline::SparklineType;
    crate::types::JsSparklineConfig {
        data_range: config.data_range.clone(),
        location: config.location.clone(),
        sparkline_type: match config.sparkline_type {
            SparklineType::Line => Some("line".to_string()),
            SparklineType::Column => Some("column".to_string()),
            SparklineType::WinLoss => Some("stacked".to_string()),
        },
        markers: if config.markers { Some(true) } else { None },
        high_point: if config.high_point { Some(true) } else { None },
        low_point: if config.low_point { Some(true) } else { None },
        first_point: if config.first_point { Some(true) } else { None },
        last_point: if config.last_point { Some(true) } else { None },
        negative_points: if config.negative_points {
            Some(true)
        } else {
            None
        },
        show_axis: if config.show_axis { Some(true) } else { None },
        line_weight: config.line_weight,
        style: config.style,
    }
}
