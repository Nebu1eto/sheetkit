use std::borrow::Cow;

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

pub(crate) type JsCellInputValue = Either6<String, f64, bool, DateValue, FormulaValue, Null>;
pub(crate) type JsCellOutputValue = Either5<Null, bool, f64, String, DateValue>;

pub(crate) fn js_value_to_cell_value(v: JsCellInputValue) -> Result<CellValue> {
    Ok(match v {
        Either6::A(s) => CellValue::String(s),
        Either6::B(n) => CellValue::Number(n),
        Either6::C(b) => CellValue::Bool(b),
        Either6::D(d) => CellValue::Date(d.serial),
        Either6::E(formula) => CellValue::Formula {
            expr: formula.formula,
            result: formula
                .result
                .map(formula_result_to_cell_value)
                .transpose()?
                .map(Box::new),
        },
        Either6::F(_) => CellValue::Empty,
    })
}

fn formula_result_to_cell_value(value: FormulaResultValue) -> Result<CellValue> {
    match value.value_type.as_str() {
        "empty" => Ok(CellValue::Empty),
        "boolean" => value
            .bool_value
            .map(CellValue::Bool)
            .ok_or_else(|| Error::from_reason("formula boolean result requires boolValue")),
        "number" => value
            .number_value
            .map(CellValue::Number)
            .ok_or_else(|| Error::from_reason("formula number result requires numberValue")),
        "date" => value
            .date
            .map(|date| CellValue::Date(date.serial))
            .ok_or_else(|| Error::from_reason("formula date result requires date")),
        "error" => value
            .value
            .map(CellValue::Error)
            .ok_or_else(|| Error::from_reason("formula error result requires value")),
        "string" => value
            .value
            .map(CellValue::String)
            .ok_or_else(|| Error::from_reason("formula string result requires value")),
        "formula" => Err(Error::from_reason("formula results cannot be nested")),
        other => Err(Error::from_reason(format!(
            "unknown formula result type: {other}"
        ))),
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
        formula: cell_value_to_formula_value(&value),
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
        formula: cell_value_to_formula_value(&value),
    }
}

pub(crate) fn cell_value_to_parts(
    value: &CellValue,
) -> (String, Option<String>, Option<f64>, Option<bool>) {
    let (value_type, str_val, num_val, bool_val) = cell_value_parts_ref(value);
    (
        value_type.to_string(),
        str_val.map(Cow::into_owned),
        num_val,
        bool_val,
    )
}

/// Borrowed form of [`cell_value_to_parts`]: value type, string value,
/// numeric value, and boolean value of one cell without cloning text.
pub(crate) fn cell_value_parts_ref(
    value: &CellValue,
) -> (
    &'static str,
    Option<Cow<'_, str>>,
    Option<f64>,
    Option<bool>,
) {
    match value {
        CellValue::Empty => ("empty", None, None, None),
        CellValue::String(s) => ("string", Some(Cow::Borrowed(s)), None, None),
        CellValue::Number(n) => ("number", None, Some(*n), None),
        CellValue::Bool(b) => ("boolean", None, None, Some(*b)),
        CellValue::Error(e) => ("error", Some(Cow::Borrowed(e)), None, None),
        CellValue::Formula { expr, .. } => ("formula", Some(Cow::Borrowed(expr)), None, None),
        CellValue::Date(serial) => (
            "date",
            serial_to_iso(*serial).map(Cow::Owned),
            Some(*serial),
            None,
        ),
        CellValue::RichString(runs) => {
            let plain = sheetkit_core::rich_text::rich_text_to_plain(runs);
            ("string", Some(Cow::Owned(plain)), None, None)
        }
    }
}

/// Tag strings shared with `decodeStreamBatch` in `index.ts`. Index order is
/// the wire code; any other tag is written literally after
/// [`STREAM_TAG_LITERAL`].
const STREAM_TAGS: [&str; 7] = [
    "empty", "string", "number", "boolean", "error", "formula", "date",
];
const STREAM_TAG_LITERAL: u8 = 0xFF;
const STREAM_HAS_VALUE: u8 = 0x01;
const STREAM_HAS_NUMBER: u8 = 0x02;
const STREAM_HAS_BOOL: u8 = 0x04;
const STREAM_HAS_NESTED: u8 = 0x08;

/// Encode one streaming-reader batch for `decodeStreamBatch` in `index.ts`.
///
/// The decoded objects are field-for-field identical to the `JsRowData`
/// values produced by [`cell_value_to_row_cell`]; only the transfer form
/// differs. Layout (little-endian):
///
/// ```text
/// u32 row_count
/// row:  u32 row_number, u32 cell_count, cell*
/// cell: u16 column (0 = no name), tag value_type, u8 flags,
///       [str value] [f64 number] [u8 bool] [formula]
/// formula: str expression, u8 has_result, [result]
/// result:  tag value_type, u8 flags, [str value] [f64 number] [u8 bool] [date]
/// date:    tag kind, f64 serial, u8 has_iso, [str iso]
/// tag:     u8 index into STREAM_TAGS, or 0xFF followed by str
/// str:     u32 byte length, UTF-8 bytes
/// ```
pub(crate) fn encode_stream_batch(
    rows: &[sheetkit_core::stream_reader::StreamRow],
) -> std::result::Result<Vec<u8>, String> {
    let mut out = StreamBatchWriter {
        bytes: Vec::with_capacity(4096),
    };
    out.u32_len(rows.len())?;
    for row in rows {
        out.bytes.extend_from_slice(&row.row_number.to_le_bytes());
        out.u32_len(row.cells.len())?;
        for (col, value) in &row.cells {
            let col = if (1..=sheetkit_core::utils::constants::MAX_COLUMNS).contains(col) {
                *col as u16
            } else {
                0
            };
            out.bytes.extend_from_slice(&col.to_le_bytes());
            let (value_type, str_val, num_val, bool_val) = cell_value_parts_ref(value);
            let formula = cell_value_to_formula_value(value);
            out.tag(value_type)?;
            out.scalar_parts(str_val.as_deref(), num_val, bool_val, formula.is_some())?;
            if let Some(formula) = formula {
                out.tag(&formula.kind)?;
                out.str(&formula.formula)?;
                match formula.result {
                    None => out.bytes.push(0),
                    Some(result) => {
                        out.bytes.push(1);
                        out.tag(&result.value_type)?;
                        out.scalar_parts(
                            result.value.as_deref(),
                            result.number_value,
                            result.bool_value,
                            result.date.is_some(),
                        )?;
                        if let Some(date) = result.date {
                            out.tag(&date.kind)?;
                            out.bytes.extend_from_slice(&date.serial.to_le_bytes());
                            match date.iso {
                                None => out.bytes.push(0),
                                Some(iso) => {
                                    out.bytes.push(1);
                                    out.str(&iso)?;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    Ok(out.bytes)
}

struct StreamBatchWriter {
    bytes: Vec<u8>,
}

impl StreamBatchWriter {
    fn u32_len(&mut self, len: usize) -> std::result::Result<(), String> {
        let len = u32::try_from(len).map_err(|_| "stream batch field is too large".to_string())?;
        self.bytes.extend_from_slice(&len.to_le_bytes());
        Ok(())
    }

    fn str(&mut self, value: &str) -> std::result::Result<(), String> {
        self.u32_len(value.len())?;
        self.bytes.extend_from_slice(value.as_bytes());
        Ok(())
    }

    fn tag(&mut self, value: &str) -> std::result::Result<(), String> {
        match STREAM_TAGS.iter().position(|tag| *tag == value) {
            Some(index) => {
                self.bytes.push(index as u8);
                Ok(())
            }
            None => {
                self.bytes.push(STREAM_TAG_LITERAL);
                self.str(value)
            }
        }
    }

    fn scalar_parts(
        &mut self,
        value: Option<&str>,
        number: Option<f64>,
        boolean: Option<bool>,
        nested: bool,
    ) -> std::result::Result<(), String> {
        let mut flags = 0;
        if value.is_some() {
            flags |= STREAM_HAS_VALUE;
        }
        if number.is_some() {
            flags |= STREAM_HAS_NUMBER;
        }
        if boolean.is_some() {
            flags |= STREAM_HAS_BOOL;
        }
        if nested {
            flags |= STREAM_HAS_NESTED;
        }
        self.bytes.push(flags);
        if let Some(value) = value {
            self.str(value)?;
        }
        if let Some(number) = number {
            self.bytes.extend_from_slice(&number.to_le_bytes());
        }
        if let Some(boolean) = boolean {
            self.bytes.push(u8::from(boolean));
        }
        Ok(())
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

pub(crate) fn cell_value_to_either(value: CellValue) -> Result<JsCellOutputValue> {
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

pub(crate) fn cell_value_to_formula_value(value: &CellValue) -> Option<FormulaValue> {
    let CellValue::Formula { expr, result } = value else {
        return None;
    };
    Some(FormulaValue {
        kind: "formula".to_string(),
        formula: expr.clone(),
        result: result.as_deref().map(cell_value_to_formula_result),
    })
}

fn cell_value_to_formula_result(value: &CellValue) -> FormulaResultValue {
    match value {
        CellValue::Empty => FormulaResultValue {
            value_type: "empty".to_string(),
            value: None,
            number_value: None,
            bool_value: None,
            date: None,
        },
        CellValue::Bool(value) => FormulaResultValue {
            value_type: "boolean".to_string(),
            value: None,
            number_value: None,
            bool_value: Some(*value),
            date: None,
        },
        CellValue::Number(value) => FormulaResultValue {
            value_type: "number".to_string(),
            value: None,
            number_value: Some(*value),
            bool_value: None,
            date: None,
        },
        CellValue::String(value) => FormulaResultValue {
            value_type: "string".to_string(),
            value: Some(value.clone()),
            number_value: None,
            bool_value: None,
            date: None,
        },
        CellValue::Date(serial) => FormulaResultValue {
            value_type: "date".to_string(),
            value: None,
            number_value: None,
            bool_value: None,
            date: Some(DateValue {
                kind: "date".to_string(),
                serial: *serial,
                iso: serial_to_iso(*serial),
            }),
        },
        CellValue::Error(value) => FormulaResultValue {
            value_type: "error".to_string(),
            value: Some(value.clone()),
            number_value: None,
            bool_value: None,
            date: None,
        },
        CellValue::Formula { expr, .. } => FormulaResultValue {
            value_type: "formula".to_string(),
            value: Some(expr.clone()),
            number_value: None,
            bool_value: None,
            date: None,
        },
        CellValue::RichString(runs) => FormulaResultValue {
            value_type: "string".to_string(),
            value: Some(sheetkit_core::rich_text::rich_text_to_plain(runs)),
            number_value: None,
            bool_value: None,
            date: None,
        },
    }
}

fn serial_to_iso(serial: f64) -> Option<String> {
    sheetkit_core::cell::serial_to_datetime(serial).map(|dt| {
        if serial.fract() == 0.0 {
            dt.format("%Y-%m-%d").to_string()
        } else {
            dt.format("%Y-%m-%dT%H:%M:%S").to_string()
        }
    })
}

pub(crate) fn parse_style_color(s: &str) -> Option<StyleColor> {
    if let Some(rgb) = s.strip_prefix('#') {
        (rgb.len() == 6 && rgb.chars().all(|c| c.is_ascii_hexdigit()))
            .then(|| StyleColor::Rgb(format!("FF{}", rgb.to_ascii_uppercase())))
    } else if s.len() == 8 && s.chars().all(|c| c.is_ascii_hexdigit()) {
        Some(StyleColor::Rgb(s.to_ascii_uppercase()))
    } else if s.len() == 6 && s.chars().all(|c| c.is_ascii_hexdigit()) {
        Some(StyleColor::Rgb(format!("FF{}", s.to_ascii_uppercase())))
    } else if let Some(theme_str) = s.strip_prefix("theme:") {
        theme_str.parse::<u32>().ok().map(StyleColor::Theme)
    } else if let Some(indexed_str) = s.strip_prefix("indexed:") {
        indexed_str.parse::<u32>().ok().map(StyleColor::Indexed)
    } else {
        None
    }
}

fn unknown_enum_error(kind: &str, value: &str, accepted: &[&str]) -> Error {
    invalid_arg(sheetkit_core::error::Error::InvalidArgument(format!(
        "unknown {kind} '{value}'; expected one of: {}",
        accepted.join(", ")
    )))
}

const PATTERN_TYPES: &[&str] = &[
    "none",
    "solid",
    "gray125",
    "darkGray",
    "mediumGray",
    "lightGray",
];

/// Parse a fill pattern name (case-insensitive). Unknown names are rejected.
pub(crate) fn parse_pattern_type(s: &str) -> Result<PatternType> {
    Ok(match s.to_lowercase().as_str() {
        "none" => PatternType::None,
        "solid" => PatternType::Solid,
        "gray125" => PatternType::Gray125,
        "darkgray" => PatternType::DarkGray,
        "mediumgray" => PatternType::MediumGray,
        "lightgray" => PatternType::LightGray,
        _ => return Err(unknown_enum_error("fill pattern", s, PATTERN_TYPES)),
    })
}

const BORDER_LINE_STYLES: &[&str] = &[
    "none",
    "thin",
    "medium",
    "thick",
    "dashed",
    "dotted",
    "double",
    "hair",
    "mediumDashed",
    "dashDot",
    "mediumDashDot",
    "dashDotDot",
    "mediumDashDotDot",
    "slantDashDot",
];

/// Parse a border line style (case-insensitive). `none` yields `None`;
/// unknown names are rejected.
pub(crate) fn parse_border_line_style(s: &str) -> Result<Option<BorderLineStyle>> {
    Ok(match s.to_lowercase().as_str() {
        "none" => None,
        "thin" => Some(BorderLineStyle::Thin),
        "medium" => Some(BorderLineStyle::Medium),
        "thick" => Some(BorderLineStyle::Thick),
        "dashed" => Some(BorderLineStyle::Dashed),
        "dotted" => Some(BorderLineStyle::Dotted),
        "double" => Some(BorderLineStyle::Double),
        "hair" => Some(BorderLineStyle::Hair),
        "mediumdashed" => Some(BorderLineStyle::MediumDashed),
        "dashdot" => Some(BorderLineStyle::DashDot),
        "mediumdashdot" => Some(BorderLineStyle::MediumDashDot),
        "dashdotdot" => Some(BorderLineStyle::DashDotDot),
        "mediumdashdotdot" => Some(BorderLineStyle::MediumDashDotDot),
        "slantdashdot" => Some(BorderLineStyle::SlantDashDot),
        _ => return Err(unknown_enum_error("border style", s, BORDER_LINE_STYLES)),
    })
}

const HORIZONTAL_ALIGNS: &[&str] = &[
    "general",
    "left",
    "center",
    "right",
    "fill",
    "justify",
    "centerContinuous",
    "distributed",
];

/// Parse a horizontal alignment (case-insensitive). Unknown names are rejected.
pub(crate) fn parse_horizontal_align(s: &str) -> Result<HorizontalAlign> {
    Ok(match s.to_lowercase().as_str() {
        "general" => HorizontalAlign::General,
        "left" => HorizontalAlign::Left,
        "center" => HorizontalAlign::Center,
        "right" => HorizontalAlign::Right,
        "fill" => HorizontalAlign::Fill,
        "justify" => HorizontalAlign::Justify,
        "centercontinuous" => HorizontalAlign::CenterContinuous,
        "distributed" => HorizontalAlign::Distributed,
        _ => {
            return Err(unknown_enum_error(
                "horizontal alignment",
                s,
                HORIZONTAL_ALIGNS,
            ))
        }
    })
}

const VERTICAL_ALIGNS: &[&str] = &["top", "center", "bottom", "justify", "distributed"];

/// Parse a vertical alignment (case-insensitive). Unknown names are rejected.
pub(crate) fn parse_vertical_align(s: &str) -> Result<VerticalAlign> {
    Ok(match s.to_lowercase().as_str() {
        "top" => VerticalAlign::Top,
        "center" => VerticalAlign::Center,
        "bottom" => VerticalAlign::Bottom,
        "justify" => VerticalAlign::Justify,
        "distributed" => VerticalAlign::Distributed,
        _ => return Err(unknown_enum_error("vertical alignment", s, VERTICAL_ALIGNS)),
    })
}

fn js_border_side_to_core(side: &JsBorderSideStyle) -> Result<Option<BorderSideStyle>> {
    let style = match side.style.as_deref() {
        Some(style) => parse_border_line_style(style)?,
        None => Some(BorderLineStyle::Thin),
    };
    Ok(style.map(|style| BorderSideStyle {
        style,
        color: side
            .color
            .as_ref()
            .and_then(|color| parse_style_color(color)),
    }))
}

fn js_border_to_core(border: &JsBorderStyle) -> Result<BorderStyle> {
    let side = |side: &Option<JsBorderSideStyle>| -> Result<Option<BorderSideStyle>> {
        match side {
            Some(side) => js_border_side_to_core(side),
            None => Ok(None),
        }
    };
    Ok(BorderStyle {
        left: side(&border.left)?,
        right: side(&border.right)?,
        top: side(&border.top)?,
        bottom: side(&border.bottom)?,
        diagonal: side(&border.diagonal)?,
    })
}

fn js_fill_to_core(f: &JsFillStyle) -> Result<FillStyle> {
    Ok(FillStyle {
        pattern: f
            .pattern
            .as_deref()
            .map(parse_pattern_type)
            .transpose()?
            .unwrap_or(PatternType::None),
        fg_color: f.fg_color.as_ref().and_then(|s| parse_style_color(s)),
        bg_color: f.bg_color.as_ref().and_then(|s| parse_style_color(s)),
        gradient: None,
    })
}

fn js_alignment_to_core(a: &JsAlignmentStyle) -> Result<AlignmentStyle> {
    Ok(AlignmentStyle {
        horizontal: a
            .horizontal
            .as_deref()
            .map(parse_horizontal_align)
            .transpose()?,
        vertical: a
            .vertical
            .as_deref()
            .map(parse_vertical_align)
            .transpose()?,
        wrap_text: a.wrap_text.unwrap_or(false),
        text_rotation: checked_opt_u32("alignment.textRotation", a.text_rotation)
            .map_err(invalid_arg)?,
        indent: checked_opt_u32("alignment.indent", a.indent).map_err(invalid_arg)?,
        shrink_to_fit: a.shrink_to_fit.unwrap_or(false),
    })
}

/// Convert a JS style to a core style, rejecting unknown enum strings and
/// non-integer numeric fields.
pub(crate) fn js_style_to_core(js: &JsStyle) -> Result<Style> {
    Ok(Style {
        font: js.font.as_ref().map(|f| FontStyle {
            name: f.name.clone(),
            size: f.size,
            bold: f.bold.unwrap_or(false),
            italic: f.italic.unwrap_or(false),
            underline: f.underline.unwrap_or(false),
            strikethrough: f.strikethrough.unwrap_or(false),
            color: f.color.as_ref().and_then(|s| parse_style_color(s)),
        }),
        fill: js.fill.as_ref().map(js_fill_to_core).transpose()?,
        border: js.border.as_ref().map(js_border_to_core).transpose()?,
        alignment: js
            .alignment
            .as_ref()
            .map(js_alignment_to_core)
            .transpose()?,
        num_fmt: if let Some(custom) = &js.custom_num_fmt {
            Some(NumFmtStyle::Custom(custom.clone()))
        } else {
            checked_opt_u32("numFmtId", js.num_fmt_id)
                .map_err(invalid_arg)?
                .map(NumFmtStyle::Builtin)
        },
        protection: js.protection.as_ref().map(|p| ProtectionStyle {
            locked: p.locked.unwrap_or(true),
            hidden: p.hidden.unwrap_or(false),
        }),
    })
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

pub(crate) fn js_app_props_to_core(js: &JsAppProperties) -> Result<AppProperties> {
    Ok(AppProperties {
        application: js.application.clone(),
        doc_security: checked_opt_u32("docSecurity", js.doc_security).map_err(invalid_arg)?,
        company: js.company.clone(),
        app_version: js.app_version.clone(),
        manager: js.manager.clone(),
        template: js.template.clone(),
    })
}

pub(crate) fn core_app_props_to_js(props: &AppProperties) -> JsAppProperties {
    JsAppProperties {
        application: props.application.clone(),
        doc_security: props.doc_security.map(f64::from),
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

/// Parse a conditional-format value type (case-insensitive). Unknown names
/// are rejected.
pub(crate) fn parse_cf_value_type(s: &str) -> Result<CfValueType> {
    Ok(match s.to_lowercase().as_str() {
        "num" => CfValueType::Num,
        "percent" => CfValueType::Percent,
        "min" => CfValueType::Min,
        "max" => CfValueType::Max,
        "percentile" => CfValueType::Percentile,
        "formula" => CfValueType::Formula,
        _ => {
            return Err(unknown_enum_error(
                "conditional format value type",
                s,
                &["num", "percent", "min", "max", "percentile", "formula"],
            ))
        }
    })
}

fn parse_cf_value_type_or(value: Option<&str>, default: CfValueType) -> Result<CfValueType> {
    Ok(value
        .map(parse_cf_value_type)
        .transpose()?
        .unwrap_or(default))
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

pub(crate) fn js_conditional_style_to_core(js: &JsConditionalStyle) -> Result<ConditionalStyle> {
    Ok(ConditionalStyle {
        font: js.font.as_ref().map(|f| FontStyle {
            name: f.name.clone(),
            size: f.size,
            bold: f.bold.unwrap_or(false),
            italic: f.italic.unwrap_or(false),
            underline: f.underline.unwrap_or(false),
            strikethrough: f.strikethrough.unwrap_or(false),
            color: f.color.as_ref().and_then(|s| parse_style_color(s)),
        }),
        fill: js.fill.as_ref().map(js_fill_to_core).transpose()?,
        border: js.border.as_ref().map(js_border_to_core).transpose()?,
        num_fmt: js
            .custom_num_fmt
            .as_ref()
            .map(|s| NumFmtStyle::Custom(s.clone())),
    })
}

pub(crate) fn js_cf_rule_to_core(js: &JsConditionalFormatRule) -> Result<ConditionalFormatRule> {
    let rule_type = match js.rule_type.as_str() {
        "cellIs" => {
            let operator = match js.operator.as_deref() {
                None => CfOperator::Equal,
                Some(s) => parse_cf_operator(s).ok_or_else(|| {
                    unknown_enum_error(
                        "conditional format operator",
                        s,
                        &[
                            "lessThan",
                            "lessThanOrEqual",
                            "equal",
                            "notEqual",
                            "greaterThanOrEqual",
                            "greaterThan",
                            "between",
                            "notBetween",
                        ],
                    )
                })?,
            };
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
            min_type: parse_cf_value_type_or(js.min_type.as_deref(), CfValueType::Min)?,
            min_value: js.min_value.clone(),
            min_color: js.min_color.clone().unwrap_or_default(),
            mid_type: js
                .mid_type
                .as_deref()
                .map(parse_cf_value_type)
                .transpose()?,
            mid_value: js.mid_value.clone(),
            mid_color: js.mid_color.clone(),
            max_type: parse_cf_value_type_or(js.max_type.as_deref(), CfValueType::Max)?,
            max_value: js.max_value.clone(),
            max_color: js.max_color.clone().unwrap_or_default(),
        },
        "dataBar" => ConditionalFormatType::DataBar {
            min_type: parse_cf_value_type_or(js.min_type.as_deref(), CfValueType::Min)?,
            min_value: js.min_value.clone(),
            max_type: parse_cf_value_type_or(js.max_type.as_deref(), CfValueType::Max)?,
            max_value: js.max_value.clone(),
            color: js.bar_color.clone().unwrap_or_default(),
            show_value: js.show_value.unwrap_or(true),
        },
        "duplicateValues" => ConditionalFormatType::DuplicateValues,
        "uniqueValues" => ConditionalFormatType::UniqueValues,
        "top10" => ConditionalFormatType::Top10 {
            rank: checked_opt_u32("rank", js.rank)
                .map_err(invalid_arg)?
                .unwrap_or(10),
            percent: js.percent.unwrap_or(false),
        },
        "bottom10" => ConditionalFormatType::Bottom10 {
            rank: checked_opt_u32("rank", js.rank)
                .map_err(invalid_arg)?
                .unwrap_or(10),
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

    let format = js
        .format
        .as_ref()
        .map(js_conditional_style_to_core)
        .transpose()?;

    Ok(ConditionalFormatRule {
        rule_type,
        format,
        priority: checked_opt_u32("priority", js.priority).map_err(invalid_arg)?,
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
                Some(f64::from(*rank)),
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
                Some(f64::from(*rank)),
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
                left: b.left.as_ref().map(side_to_js),
                right: b.right.as_ref().map(side_to_js),
                top: b.top.as_ref().map(side_to_js),
                bottom: b.bottom.as_ref().map(side_to_js),
                diagonal: b.diagonal.as_ref().map(side_to_js),
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
        priority: rule.priority.map(f64::from),
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
) -> std::result::Result<sheetkit_core::workbook::OpenOptions, sheetkit_core::error::Error> {
    let Some(js) = js else {
        return Ok(sheetkit_core::workbook::OpenOptions::default());
    };
    let defaults = sheetkit_core::workbook::OpenOptions::default();
    // Prefer read_mode; fall back to parse_mode for backward compatibility.
    let read_mode = if let Some(rm) = js.read_mode.as_deref() {
        match rm {
            "eager" => sheetkit_core::workbook::ReadMode::Eager,
            "lazy" => sheetkit_core::workbook::ReadMode::Lazy,
            "stream" => sheetkit_core::workbook::ReadMode::Stream,
            _ => {
                return Err(sheetkit_core::error::Error::InvalidArgument(
                    "readMode must be lazy, eager, or stream".to_string(),
                ));
            }
        }
    } else {
        match js.parse_mode.as_deref() {
            Some("readfast") => sheetkit_core::workbook::ReadMode::Lazy,
            Some("full") => sheetkit_core::workbook::ReadMode::Eager,
            None => defaults.read_mode,
            Some(_) => {
                return Err(sheetkit_core::error::Error::InvalidArgument(
                    "parseMode must be readfast or full".to_string(),
                ));
            }
        }
    };
    let aux_parts = match js.aux_parts.as_deref() {
        Some("deferred") => sheetkit_core::workbook::AuxParts::Deferred,
        Some("eager") => sheetkit_core::workbook::AuxParts::EagerLoad,
        Some(_) => {
            return Err(sheetkit_core::error::Error::InvalidArgument(
                "auxParts must be deferred or eager".to_string(),
            ));
        }
        None => defaults.aux_parts,
    };
    let date_interpretation = match js.date_interpretation.as_deref() {
        Some("cellType") => sheetkit_core::workbook::DateInterpretation::CellType,
        Some("numFmt") | None => sheetkit_core::workbook::DateInterpretation::NumFmt,
        Some(_) => {
            return Err(sheetkit_core::error::Error::InvalidArgument(
                "dateInterpretation must be cellType or numFmt".to_string(),
            ));
        }
    };
    let crypt_limits = js_crypt_limits_to_core(js.crypt_limits.as_ref())?;
    let vba_limits = js_vba_limits_to_core(js.vba_limits.as_ref())?;
    let mut limits = defaults.limits;
    limits.outer_input_bytes = optional_finite_limit(
        js.max_outer_input_bytes,
        defaults.limits.outer_input_bytes,
        "maxOuterInputBytes",
    )?;
    limits.xml_bytes = optional_finite_limit(
        js.max_xml_part_bytes,
        defaults.limits.xml_bytes,
        "maxXmlPartBytes",
    )?;
    limits.compression_ratio_floor_bytes = optional_finite_limit(
        js.compression_ratio_floor_bytes,
        defaults.limits.compression_ratio_floor_bytes,
        "compressionRatioFloorBytes",
    )?;
    Ok(sheetkit_core::workbook::OpenOptions {
        sheet_rows: checked_opt_u32("sheetRows", js.sheet_rows)?,
        sheets: js.sheets.clone(),
        max_unzip_size: js
            .max_unzip_size
            .map(|value| finite_limit(value, "maxUnzipSize"))
            .transpose()?,
        max_zip_entries: js
            .max_zip_entries
            .map(|value| {
                finite_limit(value, "maxZipEntries").and_then(|value| {
                    usize::try_from(value).map_err(|_| {
                        sheetkit_core::error::Error::InvalidArgument(
                            "maxZipEntries is not representable on this platform".to_string(),
                        )
                    })
                })
            })
            .transpose()?,
        limits,
        crypt_limits,
        vba_limits,
        read_mode,
        aux_parts,
        date_interpretation,
    })
}

fn optional_finite_limit(
    value: Option<f64>,
    default: u64,
    name: &str,
) -> std::result::Result<u64, sheetkit_core::error::Error> {
    match value {
        None => Ok(default),
        Some(value) => finite_limit(value, name),
    }
}

pub(crate) fn js_crypt_limits_to_core(
    js: Option<&crate::types::JsCryptLimits>,
) -> std::result::Result<sheetkit_core::crypt::CryptLimits, sheetkit_core::error::Error> {
    let defaults = sheetkit_core::crypt::CryptLimits::default();
    let Some(js) = js else {
        return Ok(defaults);
    };
    Ok(sheetkit_core::crypt::CryptLimits {
        outer_cfb_bytes: optional_finite_limit(
            js.outer_cfb_bytes,
            defaults.outer_cfb_bytes,
            "outerCfbBytes",
        )?,
        named_stream_count: optional_finite_limit(
            js.named_stream_count,
            defaults.named_stream_count,
            "namedStreamCount",
        )?,
        named_stream_declared_bytes: optional_finite_limit(
            js.named_stream_declared_bytes,
            defaults.named_stream_declared_bytes,
            "namedStreamDeclaredBytes",
        )?,
        named_stream_read_bytes: optional_finite_limit(
            js.named_stream_read_bytes,
            defaults.named_stream_read_bytes,
            "namedStreamReadBytes",
        )?,
        encrypted_package_bytes: optional_finite_limit(
            js.encrypted_package_bytes,
            defaults.encrypted_package_bytes,
            "encryptedPackageBytes",
        )?,
        kdf_iterations: optional_finite_limit(
            js.kdf_iterations,
            defaults.kdf_iterations,
            "kdfIterations",
        )?,
        kdf_work: optional_finite_limit(js.kdf_work, defaults.kdf_work, "kdfWork")?,
    })
}

pub(crate) fn js_save_limits_to_core(
    js: Option<&crate::types::JsSaveLimits>,
) -> std::result::Result<sheetkit_core::save::SaveLimits, sheetkit_core::error::Error> {
    let defaults = sheetkit_core::save::SaveLimits::default();
    let Some(js) = js else {
        return Ok(defaults);
    };
    Ok(sheetkit_core::save::SaveLimits {
        output_bytes: optional_finite_limit(js.output_bytes, defaults.output_bytes, "outputBytes")?,
        xml_part_bytes: optional_finite_limit(
            js.xml_part_bytes,
            defaults.xml_part_bytes,
            "xmlPartBytes",
        )?,
        row_count: optional_finite_limit(js.row_count, defaults.row_count, "rowCount")?,
        cell_count: optional_finite_limit(js.cell_count, defaults.cell_count, "cellCount")?,
    })
}

pub(crate) fn js_stream_limits_to_core(
    js: Option<&crate::types::JsStreamLimits>,
) -> std::result::Result<sheetkit_core::save::StreamLimits, sheetkit_core::error::Error> {
    let defaults = sheetkit_core::save::StreamLimits::default();
    let Some(js) = js else {
        return Ok(defaults);
    };
    Ok(sheetkit_core::save::StreamLimits {
        row_count: optional_finite_limit(js.row_count, defaults.row_count, "rowCount")?,
        cell_count: optional_finite_limit(js.cell_count, defaults.cell_count, "cellCount")?,
        row_xml_bytes: optional_finite_limit(
            js.row_xml_bytes,
            defaults.row_xml_bytes,
            "rowXmlBytes",
        )?,
        output_bytes: optional_finite_limit(js.output_bytes, defaults.output_bytes, "outputBytes")?,
        config_count: optional_finite_limit(js.config_count, defaults.config_count, "configCount")?,
        config_bytes: optional_finite_limit(js.config_bytes, defaults.config_bytes, "configBytes")?,
    })
}

pub(crate) fn js_vba_limits_to_core(
    js: Option<&crate::types::JsVbaLimits>,
) -> std::result::Result<sheetkit_core::vba::VbaLimits, sheetkit_core::error::Error> {
    let defaults = sheetkit_core::vba::VbaLimits::default();
    let Some(js) = js else {
        return Ok(defaults);
    };
    Ok(sheetkit_core::vba::VbaLimits {
        input_bytes: optional_finite_limit(js.input_bytes, defaults.input_bytes, "inputBytes")?,
        cfb_stream_count: optional_finite_limit(
            js.cfb_stream_count,
            defaults.cfb_stream_count,
            "cfbStreamCount",
        )?,
        cfb_stream_declared_bytes: optional_finite_limit(
            js.cfb_stream_declared_bytes,
            defaults.cfb_stream_declared_bytes,
            "cfbStreamDeclaredBytes",
        )?,
        cfb_aggregate_declared_bytes: optional_finite_limit(
            js.cfb_aggregate_declared_bytes,
            defaults.cfb_aggregate_declared_bytes,
            "cfbAggregateDeclaredBytes",
        )?,
        cfb_stream_read_bytes: optional_finite_limit(
            js.cfb_stream_read_bytes,
            defaults.cfb_stream_read_bytes,
            "cfbStreamReadBytes",
        )?,
        cfb_aggregate_read_bytes: optional_finite_limit(
            js.cfb_aggregate_read_bytes,
            defaults.cfb_aggregate_read_bytes,
            "cfbAggregateReadBytes",
        )?,
        compressed_bytes: optional_finite_limit(
            js.compressed_bytes,
            defaults.compressed_bytes,
            "compressedBytes",
        )?,
        module_count: optional_finite_limit(js.module_count, defaults.module_count, "moduleCount")?,
        chunk_count: optional_finite_limit(js.chunk_count, defaults.chunk_count, "chunkCount")?,
        module_decoded_bytes: optional_finite_limit(
            js.module_decoded_bytes,
            defaults.module_decoded_bytes,
            "moduleDecodedBytes",
        )?,
        aggregate_decoded_bytes: optional_finite_limit(
            js.aggregate_decoded_bytes,
            defaults.aggregate_decoded_bytes,
            "aggregateDecodedBytes",
        )?,
        retained_bytes: optional_finite_limit(
            js.retained_bytes,
            defaults.retained_bytes,
            "retainedBytes",
        )?,
        output_bytes: optional_finite_limit(js.output_bytes, defaults.output_bytes, "outputBytes")?,
    })
}

pub(crate) fn js_formula_limits_to_core(
    js: Option<&crate::types::JsFormulaLimits>,
) -> std::result::Result<sheetkit_core::formula::FormulaLimits, sheetkit_core::error::Error> {
    let defaults = sheetkit_core::formula::FormulaLimits::secure_default();
    let Some(js) = js else {
        return Ok(defaults);
    };
    Ok(sheetkit_core::formula::FormulaLimits {
        per_formula_input_bytes: optional_finite_limit(
            js.per_formula_input_bytes,
            defaults.per_formula_input_bytes,
            "perFormulaInputBytes",
        )?,
        per_formula_nesting_depth: optional_finite_limit(
            js.per_formula_nesting_depth,
            defaults.per_formula_nesting_depth,
            "perFormulaNestingDepth",
        )?,
        per_formula_binary_chain_depth: optional_finite_limit(
            js.per_formula_binary_chain_depth,
            defaults.per_formula_binary_chain_depth,
            "perFormulaBinaryChainDepth",
        )?,
        per_formula_tokens: optional_finite_limit(
            js.per_formula_tokens,
            defaults.per_formula_tokens,
            "perFormulaTokens",
        )?,
        per_formula_ast_nodes: optional_finite_limit(
            js.per_formula_ast_nodes,
            defaults.per_formula_ast_nodes,
            "perFormulaAstNodes",
        )?,
        operation_input_bytes: optional_finite_limit(
            js.operation_input_bytes,
            defaults.operation_input_bytes,
            "operationInputBytes",
        )?,
        operation_tokens: optional_finite_limit(
            js.operation_tokens,
            defaults.operation_tokens,
            "operationTokens",
        )?,
        operation_ast_nodes: optional_finite_limit(
            js.operation_ast_nodes,
            defaults.operation_ast_nodes,
            "operationAstNodes",
        )?,
        operation_range_cells: optional_finite_limit(
            js.operation_range_cells,
            defaults.operation_range_cells,
            "operationRangeCells",
        )?,
        operation_dependency_edges: optional_finite_limit(
            js.operation_dependency_edges,
            defaults.operation_dependency_edges,
            "operationDependencyEdges",
        )?,
        operation_ready_queue_ops: optional_finite_limit(
            js.operation_ready_queue_ops,
            defaults.operation_ready_queue_ops,
            "operationReadyQueueOps",
        )?,
        operation_evaluation_steps: optional_finite_limit(
            js.operation_evaluation_steps,
            defaults.operation_evaluation_steps,
            "operationEvaluationSteps",
        )?,
        operation_match_work: optional_finite_limit(
            js.operation_match_work,
            defaults.operation_match_work,
            "operationMatchWork",
        )?,
        operation_output_bytes: optional_finite_limit(
            js.operation_output_bytes,
            defaults.operation_output_bytes,
            "operationOutputBytes",
        )?,
    })
}

/// Validate a JavaScript number as an integer in `min..=max`.
///
/// Rejects non-finite, fractional, unsafe (beyond `Number.MAX_SAFE_INTEGER`),
/// and out-of-range values with `InvalidArgument` instead of letting napi
/// coerce them (truncation, wrap-around, or NaN to zero).
fn checked_integer(
    name: &str,
    value: f64,
    min: f64,
    max: f64,
) -> std::result::Result<f64, sheetkit_core::error::Error> {
    const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
    if !value.is_finite()
        || value.fract() != 0.0
        || value.abs() > MAX_SAFE_INTEGER
        || value < min
        || value > max
    {
        return Err(sheetkit_core::error::Error::InvalidArgument(format!(
            "{name} must be an integer from {min} to {max}, got {value}"
        )));
    }
    Ok(value)
}

/// Validate a JavaScript number as a `u32` (0 to 4294967295).
pub(crate) fn checked_u32(
    name: &str,
    value: f64,
) -> std::result::Result<u32, sheetkit_core::error::Error> {
    checked_integer(name, value, 0.0, u32::MAX as f64).map(|value| value as u32)
}

/// Validate a JavaScript number as a `u8` (0 to 255).
pub(crate) fn checked_u8(
    name: &str,
    value: f64,
) -> std::result::Result<u8, sheetkit_core::error::Error> {
    checked_integer(name, value, 0.0, u8::MAX as f64).map(|value| value as u8)
}

/// Validate a JavaScript number as an `i32`.
pub(crate) fn checked_i32(
    name: &str,
    value: f64,
) -> std::result::Result<i32, sheetkit_core::error::Error> {
    checked_integer(name, value, i32::MIN as f64, i32::MAX as f64).map(|value| value as i32)
}

/// Validate an optional JavaScript number as a `u32`.
pub(crate) fn checked_opt_u32(
    name: &str,
    value: Option<f64>,
) -> std::result::Result<Option<u32>, sheetkit_core::error::Error> {
    value.map(|value| checked_u32(name, value)).transpose()
}

/// Validate an optional JavaScript number as an `i32`.
pub(crate) fn checked_opt_i32(
    name: &str,
    value: Option<f64>,
) -> std::result::Result<Option<i32>, sheetkit_core::error::Error> {
    value.map(|value| checked_i32(name, value)).transpose()
}

/// Convert a core error from argument validation into a napi error with the
/// same message and status that `map_core_error` uses for non-resource errors.
pub(crate) fn invalid_arg(error: sheetkit_core::error::Error) -> napi::Error {
    napi::Error::from_reason(error.to_string())
}

fn finite_limit(value: f64, name: &str) -> std::result::Result<u64, sheetkit_core::error::Error> {
    const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
    if !value.is_finite() || value < 0.0 || value.fract() != 0.0 || value > MAX_SAFE_INTEGER {
        return Err(sheetkit_core::error::Error::InvalidArgument(format!(
            "{name} must be a finite non-negative safe integer"
        )));
    }
    Ok(value as u64)
}

pub(crate) fn js_traversal_options_to_core(
    js: Option<&crate::types::JsTraversalOptions>,
) -> std::result::Result<sheetkit_core::traversal::TraversalOptions, sheetkit_core::error::Error> {
    use sheetkit_core::traversal::{TraversalMode, TraversalOptions};

    let Some(js) = js else {
        return Ok(TraversalOptions::default());
    };
    let mode = match js.mode.as_deref() {
        None | Some("sparse") => TraversalMode::Sparse,
        Some("dense") => TraversalMode::Dense,
        Some(_) => {
            return Err(sheetkit_core::error::Error::InvalidArgument(
                "traversal mode must be sparse or dense".to_string(),
            ));
        }
    };
    let limits = js_traversal_limits_to_core(js.limits.as_ref())?;
    Ok(TraversalOptions::new(mode).with_limits(limits))
}

pub(crate) fn js_dense_traversal_limits_to_core(
    js: Option<&crate::types::JsTraversalOptions>,
) -> std::result::Result<sheetkit_core::traversal::TraversalLimits, sheetkit_core::error::Error> {
    if let Some(js) = js {
        match js.mode.as_deref() {
            None | Some("dense") => {}
            Some("sparse") => {
                return Err(sheetkit_core::error::Error::InvalidArgument(
                    "this operation requires dense traversal".to_string(),
                ));
            }
            Some(_) => {
                return Err(sheetkit_core::error::Error::InvalidArgument(
                    "traversal mode must be dense".to_string(),
                ));
            }
        }
    }
    js_traversal_limits_to_core(js.and_then(|options| options.limits.as_ref()))
}

fn js_traversal_limits_to_core(
    js: Option<&crate::types::JsTraversalLimits>,
) -> std::result::Result<sheetkit_core::traversal::TraversalLimits, sheetkit_core::error::Error> {
    use sheetkit_core::traversal::TraversalLimits;

    let defaults = TraversalLimits::default();
    match js {
        Some(limits) => TraversalLimits::new(
            limits
                .max_rows
                .map(|value| finite_limit(value, "maxRows"))
                .transpose()?
                .unwrap_or(defaults.max_rows()),
            limits
                .max_columns
                .map(|value| finite_limit(value, "maxColumns"))
                .transpose()?
                .unwrap_or(defaults.max_columns()),
            limits
                .max_cells
                .map(|value| finite_limit(value, "maxCells"))
                .transpose()?
                .unwrap_or(defaults.max_cells()),
            limits
                .max_output_bytes
                .map(|value| finite_limit(value, "maxOutputBytes"))
                .transpose()?
                .unwrap_or(defaults.max_output_bytes()),
        ),
        None => Ok(defaults),
    }
}

pub(crate) fn parse_form_control_type(s: &str) -> Result<sheetkit_core::control::FormControlType> {
    sheetkit_core::control::FormControlType::parse(s).map_err(|e| Error::from_reason(e.to_string()))
}

pub(crate) fn js_form_control_to_core(
    js: &crate::types::JsFormControlConfig,
) -> Result<sheetkit_core::control::FormControlConfig> {
    let control_type = parse_form_control_type(&js.control_type)?;
    Ok(sheetkit_core::control::FormControlConfig {
        control_type,
        cell: js.cell.clone(),
        width: js.width,
        height: js.height,
        text: js.text.clone(),
        macro_name: js.macro_name.clone(),
        cell_link: js.cell_link.clone(),
        checked: js.checked,
        min_value: checked_opt_u32("minValue", js.min_value).map_err(invalid_arg)?,
        max_value: checked_opt_u32("maxValue", js.max_value).map_err(invalid_arg)?,
        increment: checked_opt_u32("increment", js.increment).map_err(invalid_arg)?,
        page_increment: checked_opt_u32("pageIncrement", js.page_increment).map_err(invalid_arg)?,
        current_value: checked_opt_u32("currentValue", js.current_value).map_err(invalid_arg)?,
        three_d: js.three_d,
    })
}

pub(crate) fn core_form_control_info_to_js(
    info: &sheetkit_core::control::FormControlInfo,
) -> crate::types::JsFormControlInfo {
    use sheetkit_core::control::FormControlType;
    let control_type = match info.control_type {
        FormControlType::Button => "button",
        FormControlType::CheckBox => "checkbox",
        FormControlType::OptionButton => "optionButton",
        FormControlType::SpinButton => "spinButton",
        FormControlType::ScrollBar => "scrollBar",
        FormControlType::GroupBox => "groupBox",
        FormControlType::Label => "label",
    }
    .to_string();
    crate::types::JsFormControlInfo {
        control_type,
        cell: info.cell.clone(),
        text: info.text.clone(),
        macro_name: info.macro_name.clone(),
        cell_link: info.cell_link.clone(),
        checked: info.checked,
        current_value: info.current_value,
        min_value: info.min_value,
        max_value: info.max_value,
        increment: info.increment,
        page_increment: info.page_increment,
    }
}

pub(crate) fn js_sparkline_to_core(
    js: &crate::types::JsSparklineConfig,
) -> Result<sheetkit_core::sparkline::SparklineConfig> {
    use sheetkit_core::sparkline::{SparklineConfig, SparklineType};
    Ok(SparklineConfig {
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
        style: checked_opt_u32("style", js.style).map_err(invalid_arg)?,
    })
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
        style: config.style.map(f64::from),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalizes_rgb_style_colors_to_uppercase_argb() {
        assert!(
            matches!(parse_style_color("#a1b2c3"), Some(StyleColor::Rgb(value)) if value == "FFA1B2C3")
        );
        assert!(
            matches!(parse_style_color("a1b2c3"), Some(StyleColor::Rgb(value)) if value == "FFA1B2C3")
        );
        assert!(
            matches!(parse_style_color("80a1b2c3"), Some(StyleColor::Rgb(value)) if value == "80A1B2C3")
        );
        assert!(parse_style_color("#12345").is_none());
        assert!(parse_style_color("GG0000").is_none());
    }

    #[test]
    fn no_border_side_is_omitted() {
        let style = JsStyle {
            font: None,
            fill: None,
            border: Some(JsBorderStyle {
                left: Some(JsBorderSideStyle {
                    style: Some("none".to_string()),
                    color: Some("FF0000".to_string()),
                }),
                right: None,
                top: None,
                bottom: None,
                diagonal: None,
            }),
            alignment: None,
            num_fmt_id: None,
            custom_num_fmt: None,
            protection: None,
        };
        assert!(js_style_to_core(&style)
            .expect("style")
            .border
            .expect("border")
            .left
            .is_none());
    }

    #[test]
    fn checked_integers_reject_coercible_values() {
        for value in [
            f64::NAN,
            f64::INFINITY,
            f64::NEG_INFINITY,
            -1.0,
            1.9,
            4_294_967_296.0,
            9_007_199_254_740_992.0,
        ] {
            assert!(checked_u32("row", value).is_err(), "{value}");
        }
        assert_eq!(checked_u32("row", 4_294_967_295.0).unwrap(), u32::MAX);
        assert_eq!(checked_u32("row", 0.0).unwrap(), 0);
        assert!(checked_u8("level", 256.0).is_err());
        assert_eq!(checked_u8("level", 255.0).unwrap(), 255);
        assert!(checked_i32("rot", -2_147_483_649.0).is_err());
        assert_eq!(checked_i32("rot", -2_147_483_648.0).unwrap(), i32::MIN);
    }

    #[test]
    fn style_enum_strings_reject_unknown_values() {
        assert!(parse_pattern_type("polkadot").is_err());
        assert!(parse_border_line_style("wavy").is_err());
        assert!(parse_horizontal_align("middle").is_err());
        assert!(parse_vertical_align("middle").is_err());
        assert!(parse_cf_value_type("median").is_err());
        assert!(matches!(
            parse_pattern_type("DarkGray"),
            Ok(PatternType::DarkGray)
        ));
    }

    #[test]
    fn formula_conversion_preserves_expression_and_cached_result() {
        let value = CellValue::Formula {
            expr: "A1+1".to_string(),
            result: Some(Box::new(CellValue::Number(42.0))),
        };
        let formula = cell_value_to_formula_value(&value).expect("formula conversion");
        assert_eq!(formula.kind, "formula");
        assert_eq!(formula.formula, "A1+1");
        assert_eq!(
            formula.result.expect("cached result").number_value,
            Some(42.0)
        );
    }

    #[test]
    fn napi_formula_output_returns_expression_string() {
        let value = cell_value_to_either(CellValue::Formula {
            expr: "TODAY()".to_string(),
            result: Some(Box::new(CellValue::Date(45292.5))),
        })
        .expect("formula output");

        let Either5::D(formula) = value else {
            panic!("expected formula expression");
        };
        assert_eq!(formula, "TODAY()");
    }

    #[test]
    fn formula_result_requires_matching_payload() {
        let value = FormulaResultValue {
            value_type: "number".to_string(),
            value: None,
            number_value: None,
            bool_value: None,
            date: None,
        };

        assert!(formula_result_to_cell_value(value).is_err());
    }

    #[test]
    fn formula_result_rejects_unknown_and_nested_types() {
        for value_type in ["unknown", "formula"] {
            let value = FormulaResultValue {
                value_type: value_type.to_string(),
                value: Some("A1+1".to_string()),
                number_value: None,
                bool_value: None,
                date: None,
            };
            assert!(formula_result_to_cell_value(value).is_err());
        }
    }

    #[test]
    fn row_cells_preserve_formula_and_cached_result() {
        let cell = cell_value_to_row_cell(
            "A".to_string(),
            CellValue::Formula {
                expr: "A1+1".to_string(),
                result: Some(Box::new(CellValue::Number(42.0))),
            },
        );
        assert_eq!(cell.value_type, "formula");
        let formula = cell.formula.expect("formula");
        assert_eq!(formula.formula, "A1+1");
        assert_eq!(
            formula.result.expect("cached result").number_value,
            Some(42.0)
        );
    }

    #[test]
    fn open_options_keep_unspecified_defaults_and_parse_mode_fallback() {
        let options = |parse_mode| JsOpenOptions {
            sheet_rows: None,
            sheets: None,
            max_unzip_size: None,
            max_zip_entries: None,
            max_outer_input_bytes: None,
            max_xml_part_bytes: None,
            compression_ratio_floor_bytes: None,
            crypt_limits: None,
            vba_limits: None,
            read_mode: None,
            aux_parts: None,
            parse_mode,
            date_interpretation: None,
        };
        let defaults = js_open_options_to_core(Some(&options(None))).unwrap();
        assert_eq!(
            defaults.limits,
            sheetkit_core::workbook::OpenLimits::default()
        );
        assert!(matches!(
            defaults.read_mode,
            sheetkit_core::workbook::ReadMode::Lazy
        ));
        assert!(matches!(
            defaults.aux_parts,
            sheetkit_core::workbook::AuxParts::Deferred
        ));
        let legacy = js_open_options_to_core(Some(&options(Some("readfast".to_string())))).unwrap();
        assert!(matches!(
            legacy.read_mode,
            sheetkit_core::workbook::ReadMode::Lazy
        ));

        let mut eager = options(None);
        eager.read_mode = Some("eager".to_string());
        assert!(matches!(
            js_open_options_to_core(Some(&eager)).unwrap().read_mode,
            sheetkit_core::workbook::ReadMode::Eager
        ));
    }

    #[test]
    fn numeric_limits_reject_non_finite_fractional_and_unsafe_values() {
        for value in [f64::NAN, f64::INFINITY, -1.0, 1.5, 9_007_199_254_740_992.0] {
            assert!(finite_limit(value, "limit").is_err());
        }
        assert_eq!(
            finite_limit(9_007_199_254_740_991.0, "limit").unwrap(),
            9_007_199_254_740_991
        );
    }
}
