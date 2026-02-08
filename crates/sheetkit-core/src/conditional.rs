//! Conditional formatting builder and utilities.
//!
//! Provides a high-level API for adding, querying, and removing conditional
//! formatting rules on worksheet cells. Supports cell-is comparisons,
//! formula-based rules, color scales, data bars, duplicate/unique values,
//! top/bottom N, above/below average, and text-based rules.

use crate::error::Result;
use crate::style::{
    BorderLineStyle, BorderSideStyle, BorderStyle, FillStyle, FontStyle, NumFmtStyle, PatternType,
    StyleColor,
};
use sheetkit_xml::styles::{Dxf, Dxfs, NumFmt, StyleSheet};
use sheetkit_xml::worksheet::{
    CfColor, CfColorScale, CfDataBar, CfRule, CfVo, ConditionalFormatting, WorksheetXml,
};

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// Comparison operator for CellIs conditional formatting rules.
#[derive(Debug, Clone, PartialEq)]
pub enum CfOperator {
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
    GreaterThanOrEqual,
    GreaterThan,
    Between,
    NotBetween,
}

impl CfOperator {
    /// Convert to the XML attribute string.
    pub fn as_str(&self) -> &str {
        match self {
            CfOperator::LessThan => "lessThan",
            CfOperator::LessThanOrEqual => "lessThanOrEqual",
            CfOperator::Equal => "equal",
            CfOperator::NotEqual => "notEqual",
            CfOperator::GreaterThanOrEqual => "greaterThanOrEqual",
            CfOperator::GreaterThan => "greaterThan",
            CfOperator::Between => "between",
            CfOperator::NotBetween => "notBetween",
        }
    }

    /// Parse from the XML attribute string.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "lessThan" => Some(CfOperator::LessThan),
            "lessThanOrEqual" => Some(CfOperator::LessThanOrEqual),
            "equal" => Some(CfOperator::Equal),
            "notEqual" => Some(CfOperator::NotEqual),
            "greaterThanOrEqual" => Some(CfOperator::GreaterThanOrEqual),
            "greaterThan" => Some(CfOperator::GreaterThan),
            "between" => Some(CfOperator::Between),
            "notBetween" => Some(CfOperator::NotBetween),
            _ => None,
        }
    }
}

/// Value type for color scale and data bar thresholds.
#[derive(Debug, Clone, PartialEq)]
pub enum CfValueType {
    Num,
    Percent,
    Min,
    Max,
    Percentile,
    Formula,
}

impl CfValueType {
    /// Convert to the XML attribute string.
    pub fn as_str(&self) -> &str {
        match self {
            CfValueType::Num => "num",
            CfValueType::Percent => "percent",
            CfValueType::Min => "min",
            CfValueType::Max => "max",
            CfValueType::Percentile => "percentile",
            CfValueType::Formula => "formula",
        }
    }

    /// Parse from the XML attribute string.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "num" => Some(CfValueType::Num),
            "percent" => Some(CfValueType::Percent),
            "min" => Some(CfValueType::Min),
            "max" => Some(CfValueType::Max),
            "percentile" => Some(CfValueType::Percentile),
            "formula" => Some(CfValueType::Formula),
            _ => None,
        }
    }
}

/// The type of conditional formatting rule.
#[derive(Debug, Clone, PartialEq)]
pub enum ConditionalFormatType {
    /// Cell value comparison (e.g., greater than, between).
    CellIs {
        operator: CfOperator,
        formula: String,
        /// Optional second formula for Between/NotBetween.
        formula2: Option<String>,
    },
    /// Formula-based rule.
    Expression { formula: String },
    /// Color scale (2 or 3 color gradient).
    ColorScale {
        min_type: CfValueType,
        min_value: Option<String>,
        min_color: String,
        mid_type: Option<CfValueType>,
        mid_value: Option<String>,
        mid_color: Option<String>,
        max_type: CfValueType,
        max_value: Option<String>,
        max_color: String,
    },
    /// Data bar visualization.
    DataBar {
        min_type: CfValueType,
        min_value: Option<String>,
        max_type: CfValueType,
        max_value: Option<String>,
        color: String,
        show_value: bool,
    },
    /// Duplicate values.
    DuplicateValues,
    /// Unique values.
    UniqueValues,
    /// Top N values.
    Top10 { rank: u32, percent: bool },
    /// Bottom N values.
    Bottom10 { rank: u32, percent: bool },
    /// Above or below average.
    AboveAverage { above: bool, equal_average: bool },
    /// Cells that contain blanks.
    ContainsBlanks,
    /// Cells that do not contain blanks.
    NotContainsBlanks,
    /// Cells that contain errors.
    ContainsErrors,
    /// Cells that do not contain errors.
    NotContainsErrors,
    /// Cells containing specific text.
    ContainsText { text: String },
    /// Cells not containing specific text.
    NotContainsText { text: String },
    /// Cells beginning with specific text.
    BeginsWith { text: String },
    /// Cells ending with specific text.
    EndsWith { text: String },
}

/// Style applied by a conditional formatting rule.
///
/// Uses differential formatting (DXF): only the properties that differ from
/// the cell's base style are specified.
#[derive(Debug, Clone, Default)]
pub struct ConditionalStyle {
    pub font: Option<FontStyle>,
    pub fill: Option<FillStyle>,
    pub border: Option<BorderStyle>,
    pub num_fmt: Option<NumFmtStyle>,
}

/// A single conditional formatting rule with its style.
#[derive(Debug, Clone)]
pub struct ConditionalFormatRule {
    /// The type of conditional formatting rule.
    pub rule_type: ConditionalFormatType,
    /// The differential style to apply when the rule matches.
    pub format: Option<ConditionalStyle>,
    /// Rule priority (lower numbers are evaluated first).
    pub priority: Option<u32>,
    /// If true, no rules with lower priority are applied when this rule matches.
    pub stop_if_true: bool,
}

// ---------------------------------------------------------------------------
// DXF handling
// ---------------------------------------------------------------------------

/// Convert a `ConditionalStyle` to an XML `Dxf` and add it to the stylesheet.
/// Returns the DXF index.
fn add_dxf(stylesheet: &mut StyleSheet, style: &ConditionalStyle) -> u32 {
    let dxf = conditional_style_to_dxf(style);

    let dxfs = stylesheet.dxfs.get_or_insert_with(|| Dxfs {
        count: Some(0),
        dxfs: Vec::new(),
    });

    let id = dxfs.dxfs.len() as u32;
    dxfs.dxfs.push(dxf);
    dxfs.count = Some(dxfs.dxfs.len() as u32);
    id
}

/// Convert a `ConditionalStyle` to the XML `Dxf` struct.
fn conditional_style_to_dxf(style: &ConditionalStyle) -> Dxf {
    use sheetkit_xml::styles::{
        BoolVal, Border, BorderSide, Fill, Font, FontName, FontSize, PatternFill, Underline,
    };

    let font = style.font.as_ref().map(|f| Font {
        b: if f.bold {
            Some(BoolVal { val: None })
        } else {
            None
        },
        i: if f.italic {
            Some(BoolVal { val: None })
        } else {
            None
        },
        strike: if f.strikethrough {
            Some(BoolVal { val: None })
        } else {
            None
        },
        u: if f.underline {
            Some(Underline { val: None })
        } else {
            None
        },
        sz: f.size.map(|val| FontSize { val }),
        color: f.color.as_ref().map(style_color_to_xml_color),
        name: f.name.as_ref().map(|val| FontName { val: val.clone() }),
        family: None,
        scheme: None,
    });

    let fill = style.fill.as_ref().map(|f| Fill {
        pattern_fill: Some(PatternFill {
            pattern_type: Some(pattern_type_str(&f.pattern).to_string()),
            fg_color: f.fg_color.as_ref().map(style_color_to_xml_color),
            bg_color: f.bg_color.as_ref().map(style_color_to_xml_color),
        }),
    });

    let border = style.border.as_ref().map(|b| {
        let side_to_xml = |s: &BorderSideStyle| BorderSide {
            style: Some(border_line_str(&s.style).to_string()),
            color: s.color.as_ref().map(style_color_to_xml_color),
        };
        Border {
            diagonal_up: None,
            diagonal_down: None,
            left: b.left.as_ref().map(side_to_xml),
            right: b.right.as_ref().map(side_to_xml),
            top: b.top.as_ref().map(side_to_xml),
            bottom: b.bottom.as_ref().map(side_to_xml),
            diagonal: b.diagonal.as_ref().map(side_to_xml),
        }
    });

    let num_fmt = style.num_fmt.as_ref().map(|nf| match nf {
        NumFmtStyle::Builtin(id) => NumFmt {
            num_fmt_id: *id,
            format_code: String::new(),
        },
        NumFmtStyle::Custom(code) => NumFmt {
            num_fmt_id: 164,
            format_code: code.clone(),
        },
    });

    Dxf {
        font,
        num_fmt,
        fill,
        border,
    }
}

/// Convert a `StyleColor` to an XML `Color`.
fn style_color_to_xml_color(color: &StyleColor) -> sheetkit_xml::styles::Color {
    sheetkit_xml::styles::Color {
        auto: None,
        indexed: match color {
            StyleColor::Indexed(i) => Some(*i),
            _ => None,
        },
        rgb: match color {
            StyleColor::Rgb(rgb) => Some(rgb.clone()),
            _ => None,
        },
        theme: match color {
            StyleColor::Theme(t) => Some(*t),
            _ => None,
        },
        tint: None,
    }
}

/// Get the string for a PatternType.
fn pattern_type_str(pattern: &PatternType) -> &str {
    match pattern {
        PatternType::None => "none",
        PatternType::Solid => "solid",
        PatternType::Gray125 => "gray125",
        PatternType::DarkGray => "darkGray",
        PatternType::MediumGray => "mediumGray",
        PatternType::LightGray => "lightGray",
    }
}

/// Get the string for a BorderLineStyle.
fn border_line_str(style: &BorderLineStyle) -> &str {
    match style {
        BorderLineStyle::Thin => "thin",
        BorderLineStyle::Medium => "medium",
        BorderLineStyle::Thick => "thick",
        BorderLineStyle::Dashed => "dashed",
        BorderLineStyle::Dotted => "dotted",
        BorderLineStyle::Double => "double",
        BorderLineStyle::Hair => "hair",
        BorderLineStyle::MediumDashed => "mediumDashed",
        BorderLineStyle::DashDot => "dashDot",
        BorderLineStyle::MediumDashDot => "mediumDashDot",
        BorderLineStyle::DashDotDot => "dashDotDot",
        BorderLineStyle::MediumDashDotDot => "mediumDashDotDot",
        BorderLineStyle::SlantDashDot => "slantDashDot",
    }
}

/// Convert an XML `Dxf` back to a `ConditionalStyle`.
fn dxf_to_conditional_style(dxf: &Dxf) -> ConditionalStyle {
    let font = dxf.font.as_ref().map(|f| FontStyle {
        name: f.name.as_ref().map(|n| n.val.clone()),
        size: f.sz.as_ref().map(|s| s.val),
        bold: f.b.is_some(),
        italic: f.i.is_some(),
        underline: f.u.is_some(),
        strikethrough: f.strike.is_some(),
        color: f.color.as_ref().and_then(xml_color_to_style_color),
    });

    let fill = dxf.fill.as_ref().and_then(|f| {
        let pf = f.pattern_fill.as_ref()?;
        Some(FillStyle {
            pattern: pf
                .pattern_type
                .as_ref()
                .map(|s| parse_pattern_type(s))
                .unwrap_or(PatternType::None),
            fg_color: pf.fg_color.as_ref().and_then(xml_color_to_style_color),
            bg_color: pf.bg_color.as_ref().and_then(xml_color_to_style_color),
        })
    });

    let border = dxf.border.as_ref().map(|b| {
        let side = |s: &sheetkit_xml::styles::BorderSide| -> Option<BorderSideStyle> {
            let style_str = s.style.as_ref()?;
            let line_style = parse_border_line_style(style_str)?;
            Some(BorderSideStyle {
                style: line_style,
                color: s.color.as_ref().and_then(xml_color_to_style_color),
            })
        };
        BorderStyle {
            left: b.left.as_ref().and_then(side),
            right: b.right.as_ref().and_then(side),
            top: b.top.as_ref().and_then(side),
            bottom: b.bottom.as_ref().and_then(side),
            diagonal: b.diagonal.as_ref().and_then(side),
        }
    });

    let num_fmt = dxf.num_fmt.as_ref().map(|nf| {
        if nf.format_code.is_empty() {
            NumFmtStyle::Builtin(nf.num_fmt_id)
        } else {
            NumFmtStyle::Custom(nf.format_code.clone())
        }
    });

    ConditionalStyle {
        font,
        fill,
        border,
        num_fmt,
    }
}

/// Convert an XML Color to a StyleColor.
fn xml_color_to_style_color(color: &sheetkit_xml::styles::Color) -> Option<StyleColor> {
    if let Some(ref rgb) = color.rgb {
        Some(StyleColor::Rgb(rgb.clone()))
    } else if let Some(theme) = color.theme {
        Some(StyleColor::Theme(theme))
    } else {
        color.indexed.map(StyleColor::Indexed)
    }
}

/// Parse a pattern type string.
fn parse_pattern_type(s: &str) -> PatternType {
    match s {
        "none" => PatternType::None,
        "solid" => PatternType::Solid,
        "gray125" => PatternType::Gray125,
        "darkGray" => PatternType::DarkGray,
        "mediumGray" => PatternType::MediumGray,
        "lightGray" => PatternType::LightGray,
        _ => PatternType::None,
    }
}

/// Parse a border line style string.
fn parse_border_line_style(s: &str) -> Option<BorderLineStyle> {
    match s {
        "thin" => Some(BorderLineStyle::Thin),
        "medium" => Some(BorderLineStyle::Medium),
        "thick" => Some(BorderLineStyle::Thick),
        "dashed" => Some(BorderLineStyle::Dashed),
        "dotted" => Some(BorderLineStyle::Dotted),
        "double" => Some(BorderLineStyle::Double),
        "hair" => Some(BorderLineStyle::Hair),
        "mediumDashed" => Some(BorderLineStyle::MediumDashed),
        "dashDot" => Some(BorderLineStyle::DashDot),
        "mediumDashDot" => Some(BorderLineStyle::MediumDashDot),
        "dashDotDot" => Some(BorderLineStyle::DashDotDot),
        "mediumDashDotDot" => Some(BorderLineStyle::MediumDashDotDot),
        "slantDashDot" => Some(BorderLineStyle::SlantDashDot),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Conversion: rule -> XML CfRule
// ---------------------------------------------------------------------------

/// Compute the next priority across all existing conditional formatting rules.
fn next_priority(ws: &WorksheetXml) -> u32 {
    let max = ws
        .conditional_formatting
        .iter()
        .flat_map(|cf| cf.cf_rules.iter())
        .map(|r| r.priority)
        .max()
        .unwrap_or(0);
    max + 1
}

/// Convert a `ConditionalFormatRule` to an XML `CfRule`, adding a DXF to the
/// stylesheet if needed. Returns the CfRule ready for insertion.
fn rule_to_xml(rule: &ConditionalFormatRule, stylesheet: &mut StyleSheet, priority: u32) -> CfRule {
    let dxf_id = rule.format.as_ref().map(|style| add_dxf(stylesheet, style));

    let stop_if_true = if rule.stop_if_true { Some(true) } else { None };

    match &rule.rule_type {
        ConditionalFormatType::CellIs {
            operator,
            formula,
            formula2,
        } => {
            let mut formulas = vec![formula.clone()];
            if let Some(f2) = formula2 {
                formulas.push(f2.clone());
            }
            CfRule {
                rule_type: "cellIs".to_string(),
                dxf_id,
                priority,
                operator: Some(operator.as_str().to_string()),
                text: None,
                stop_if_true,
                above_average: None,
                equal_average: None,
                percent: None,
                rank: None,
                bottom: None,
                formulas,
                color_scale: None,
                data_bar: None,
                icon_set: None,
            }
        }
        ConditionalFormatType::Expression { formula } => CfRule {
            rule_type: "expression".to_string(),
            dxf_id,
            priority,
            operator: None,
            text: None,
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![formula.clone()],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
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
        } => {
            let mut cfvos = vec![CfVo {
                value_type: min_type.as_str().to_string(),
                val: min_value.clone(),
            }];
            let mut colors = vec![CfColor {
                rgb: Some(min_color.clone()),
                theme: None,
                tint: None,
            }];

            if let Some(mt) = mid_type {
                cfvos.push(CfVo {
                    value_type: mt.as_str().to_string(),
                    val: mid_value.clone(),
                });
                colors.push(CfColor {
                    rgb: mid_color.clone(),
                    theme: None,
                    tint: None,
                });
            }

            cfvos.push(CfVo {
                value_type: max_type.as_str().to_string(),
                val: max_value.clone(),
            });
            colors.push(CfColor {
                rgb: Some(max_color.clone()),
                theme: None,
                tint: None,
            });

            CfRule {
                rule_type: "colorScale".to_string(),
                dxf_id: None, // color scales do not use DXF
                priority,
                operator: None,
                text: None,
                stop_if_true,
                above_average: None,
                equal_average: None,
                percent: None,
                rank: None,
                bottom: None,
                formulas: vec![],
                color_scale: Some(CfColorScale { cfvos, colors }),
                data_bar: None,
                icon_set: None,
            }
        }
        ConditionalFormatType::DataBar {
            min_type,
            min_value,
            max_type,
            max_value,
            color,
            show_value,
        } => {
            let cfvos = vec![
                CfVo {
                    value_type: min_type.as_str().to_string(),
                    val: min_value.clone(),
                },
                CfVo {
                    value_type: max_type.as_str().to_string(),
                    val: max_value.clone(),
                },
            ];
            CfRule {
                rule_type: "dataBar".to_string(),
                dxf_id: None, // data bars do not use DXF
                priority,
                operator: None,
                text: None,
                stop_if_true,
                above_average: None,
                equal_average: None,
                percent: None,
                rank: None,
                bottom: None,
                formulas: vec![],
                color_scale: None,
                data_bar: Some(CfDataBar {
                    show_value: if *show_value { None } else { Some(false) },
                    cfvos,
                    color: Some(CfColor {
                        rgb: Some(color.clone()),
                        theme: None,
                        tint: None,
                    }),
                }),
                icon_set: None,
            }
        }
        ConditionalFormatType::DuplicateValues => CfRule {
            rule_type: "duplicateValues".to_string(),
            dxf_id,
            priority,
            operator: None,
            text: None,
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::UniqueValues => CfRule {
            rule_type: "uniqueValues".to_string(),
            dxf_id,
            priority,
            operator: None,
            text: None,
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::Top10 { rank, percent } => CfRule {
            rule_type: "top10".to_string(),
            dxf_id,
            priority,
            operator: None,
            text: None,
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: if *percent { Some(true) } else { None },
            rank: Some(*rank),
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::Bottom10 { rank, percent } => CfRule {
            rule_type: "top10".to_string(),
            dxf_id,
            priority,
            operator: None,
            text: None,
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: if *percent { Some(true) } else { None },
            rank: Some(*rank),
            bottom: Some(true),
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::AboveAverage {
            above,
            equal_average,
        } => CfRule {
            rule_type: "aboveAverage".to_string(),
            dxf_id,
            priority,
            operator: None,
            text: None,
            stop_if_true,
            above_average: if *above { None } else { Some(false) },
            equal_average: if *equal_average { Some(true) } else { None },
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::ContainsBlanks => CfRule {
            rule_type: "containsBlanks".to_string(),
            dxf_id,
            priority,
            operator: None,
            text: None,
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::NotContainsBlanks => CfRule {
            rule_type: "notContainsBlanks".to_string(),
            dxf_id,
            priority,
            operator: None,
            text: None,
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::ContainsErrors => CfRule {
            rule_type: "containsErrors".to_string(),
            dxf_id,
            priority,
            operator: None,
            text: None,
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::NotContainsErrors => CfRule {
            rule_type: "notContainsErrors".to_string(),
            dxf_id,
            priority,
            operator: None,
            text: None,
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::ContainsText { text } => CfRule {
            rule_type: "containsText".to_string(),
            dxf_id,
            priority,
            operator: Some("containsText".to_string()),
            text: Some(text.clone()),
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::NotContainsText { text } => CfRule {
            rule_type: "notContainsText".to_string(),
            dxf_id,
            priority,
            operator: Some("notContains".to_string()),
            text: Some(text.clone()),
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::BeginsWith { text } => CfRule {
            rule_type: "beginsWith".to_string(),
            dxf_id,
            priority,
            operator: Some("beginsWith".to_string()),
            text: Some(text.clone()),
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
        ConditionalFormatType::EndsWith { text } => CfRule {
            rule_type: "endsWith".to_string(),
            dxf_id,
            priority,
            operator: Some("endsWith".to_string()),
            text: Some(text.clone()),
            stop_if_true,
            above_average: None,
            equal_average: None,
            percent: None,
            rank: None,
            bottom: None,
            formulas: vec![],
            color_scale: None,
            data_bar: None,
            icon_set: None,
        },
    }
}

// ---------------------------------------------------------------------------
// Conversion: XML CfRule -> rule
// ---------------------------------------------------------------------------

/// Convert an XML `CfRule` to a `ConditionalFormatRule`, looking up the DXF
/// style from the stylesheet.
fn xml_to_rule(cf_rule: &CfRule, stylesheet: &StyleSheet) -> ConditionalFormatRule {
    let format = cf_rule
        .dxf_id
        .and_then(|id| {
            stylesheet
                .dxfs
                .as_ref()
                .and_then(|dxfs| dxfs.dxfs.get(id as usize))
        })
        .map(dxf_to_conditional_style);

    let rule_type = match cf_rule.rule_type.as_str() {
        "cellIs" => {
            let operator = cf_rule
                .operator
                .as_deref()
                .and_then(CfOperator::parse)
                .unwrap_or(CfOperator::Equal);
            let formula = cf_rule.formulas.first().cloned().unwrap_or_default();
            let formula2 = cf_rule.formulas.get(1).cloned();
            ConditionalFormatType::CellIs {
                operator,
                formula,
                formula2,
            }
        }
        "expression" => {
            let formula = cf_rule.formulas.first().cloned().unwrap_or_default();
            ConditionalFormatType::Expression { formula }
        }
        "colorScale" => {
            if let Some(cs) = &cf_rule.color_scale {
                let get_cfvo = |idx: usize| -> (CfValueType, Option<String>) {
                    cs.cfvos
                        .get(idx)
                        .map(|v| {
                            (
                                CfValueType::parse(&v.value_type).unwrap_or(CfValueType::Min),
                                v.val.clone(),
                            )
                        })
                        .unwrap_or((CfValueType::Min, None))
                };
                let get_color = |idx: usize| -> Option<String> {
                    cs.colors.get(idx).and_then(|c| c.rgb.clone())
                };

                let (min_type, min_value) = get_cfvo(0);
                let min_color = get_color(0).unwrap_or_default();

                let is_three_color = cs.cfvos.len() >= 3;
                let (mid_type, mid_value, mid_color) = if is_three_color {
                    let (mt, mv) = get_cfvo(1);
                    (Some(mt), mv, get_color(1))
                } else {
                    (None, None, None)
                };

                let max_idx = if is_three_color { 2 } else { 1 };
                let (max_type, max_value) = get_cfvo(max_idx);
                let max_color = get_color(max_idx).unwrap_or_default();

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
                }
            } else {
                // Fallback for malformed data
                ConditionalFormatType::ColorScale {
                    min_type: CfValueType::Min,
                    min_value: None,
                    min_color: String::new(),
                    mid_type: None,
                    mid_value: None,
                    mid_color: None,
                    max_type: CfValueType::Max,
                    max_value: None,
                    max_color: String::new(),
                }
            }
        }
        "dataBar" => {
            if let Some(db) = &cf_rule.data_bar {
                let get_cfvo = |idx: usize| -> (CfValueType, Option<String>) {
                    db.cfvos
                        .get(idx)
                        .map(|v| {
                            (
                                CfValueType::parse(&v.value_type).unwrap_or(CfValueType::Min),
                                v.val.clone(),
                            )
                        })
                        .unwrap_or((CfValueType::Min, None))
                };
                let (min_type, min_value) = get_cfvo(0);
                let (max_type, max_value) = get_cfvo(1);
                let color = db
                    .color
                    .as_ref()
                    .and_then(|c| c.rgb.clone())
                    .unwrap_or_default();
                let show_value = db.show_value.unwrap_or(true);

                ConditionalFormatType::DataBar {
                    min_type,
                    min_value,
                    max_type,
                    max_value,
                    color,
                    show_value,
                }
            } else {
                ConditionalFormatType::DataBar {
                    min_type: CfValueType::Min,
                    min_value: None,
                    max_type: CfValueType::Max,
                    max_value: None,
                    color: String::new(),
                    show_value: true,
                }
            }
        }
        "duplicateValues" => ConditionalFormatType::DuplicateValues,
        "uniqueValues" => ConditionalFormatType::UniqueValues,
        "top10" => {
            let rank = cf_rule.rank.unwrap_or(10);
            let percent = cf_rule.percent.unwrap_or(false);
            if cf_rule.bottom == Some(true) {
                ConditionalFormatType::Bottom10 { rank, percent }
            } else {
                ConditionalFormatType::Top10 { rank, percent }
            }
        }
        "aboveAverage" => {
            let above = cf_rule.above_average.unwrap_or(true);
            let equal_average = cf_rule.equal_average.unwrap_or(false);
            ConditionalFormatType::AboveAverage {
                above,
                equal_average,
            }
        }
        "containsBlanks" => ConditionalFormatType::ContainsBlanks,
        "notContainsBlanks" => ConditionalFormatType::NotContainsBlanks,
        "containsErrors" => ConditionalFormatType::ContainsErrors,
        "notContainsErrors" => ConditionalFormatType::NotContainsErrors,
        "containsText" => ConditionalFormatType::ContainsText {
            text: cf_rule.text.clone().unwrap_or_default(),
        },
        "notContainsText" => ConditionalFormatType::NotContainsText {
            text: cf_rule.text.clone().unwrap_or_default(),
        },
        "beginsWith" => ConditionalFormatType::BeginsWith {
            text: cf_rule.text.clone().unwrap_or_default(),
        },
        "endsWith" => ConditionalFormatType::EndsWith {
            text: cf_rule.text.clone().unwrap_or_default(),
        },
        // Unknown rule types default to expression
        _ => ConditionalFormatType::Expression {
            formula: cf_rule.formulas.first().cloned().unwrap_or_default(),
        },
    };

    ConditionalFormatRule {
        rule_type,
        format,
        priority: Some(cf_rule.priority),
        stop_if_true: cf_rule.stop_if_true.unwrap_or(false),
    }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Set conditional formatting rules on a cell range. Each call adds a new
/// `conditionalFormatting` element with one or more `cfRule` children.
pub fn set_conditional_format(
    ws: &mut WorksheetXml,
    stylesheet: &mut StyleSheet,
    sqref: &str,
    rules: &[ConditionalFormatRule],
) -> Result<()> {
    let mut xml_rules = Vec::with_capacity(rules.len());
    for rule in rules {
        let priority = rule.priority.unwrap_or_else(|| next_priority(ws));
        let cf_rule = rule_to_xml(rule, stylesheet, priority);
        xml_rules.push(cf_rule);
    }

    ws.conditional_formatting.push(ConditionalFormatting {
        sqref: sqref.to_string(),
        cf_rules: xml_rules,
    });

    Ok(())
}

/// Get all conditional formatting rules from a worksheet.
///
/// Returns a list of `(sqref, rules)` pairs.
pub fn get_conditional_formats(
    ws: &WorksheetXml,
    stylesheet: &StyleSheet,
) -> Vec<(String, Vec<ConditionalFormatRule>)> {
    ws.conditional_formatting
        .iter()
        .map(|cf| {
            let rules = cf
                .cf_rules
                .iter()
                .map(|r| xml_to_rule(r, stylesheet))
                .collect();
            (cf.sqref.clone(), rules)
        })
        .collect()
}

/// Delete all conditional formatting rules for a specific cell range.
pub fn delete_conditional_format(ws: &mut WorksheetXml, sqref: &str) -> Result<()> {
    ws.conditional_formatting.retain(|cf| cf.sqref != sqref);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn default_stylesheet() -> StyleSheet {
        StyleSheet::default()
    }

    // -----------------------------------------------------------------------
    // CellIs tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_cell_is_greater_than() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::CellIs {
                operator: CfOperator::GreaterThan,
                formula: "100".to_string(),
                formula2: None,
            },
            format: Some(ConditionalStyle {
                font: Some(FontStyle {
                    bold: true,
                    color: Some(StyleColor::Rgb("FFFF0000".to_string())),
                    ..FontStyle::default()
                }),
                ..ConditionalStyle::default()
            }),
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        assert_eq!(ws.conditional_formatting.len(), 1);
        assert_eq!(ws.conditional_formatting[0].sqref, "A1:A100");
        assert_eq!(ws.conditional_formatting[0].cf_rules.len(), 1);

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "cellIs");
        assert_eq!(rule.operator, Some("greaterThan".to_string()));
        assert_eq!(rule.formulas, vec!["100".to_string()]);
        assert!(rule.dxf_id.is_some());
        assert_eq!(rule.priority, 1);

        // Verify DXF was added
        let dxfs = ss.dxfs.as_ref().unwrap();
        assert_eq!(dxfs.dxfs.len(), 1);
        assert!(dxfs.dxfs[0].font.is_some());
    }

    #[test]
    fn test_cell_is_between() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::CellIs {
                operator: CfOperator::Between,
                formula: "10".to_string(),
                formula2: Some("20".to_string()),
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "B1:B50", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.operator, Some("between".to_string()));
        assert_eq!(rule.formulas, vec!["10".to_string(), "20".to_string()]);
        assert!(rule.dxf_id.is_none());
    }

    #[test]
    fn test_cell_is_equal() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::CellIs {
                operator: CfOperator::Equal,
                formula: "\"Yes\"".to_string(),
                formula2: None,
            },
            format: Some(ConditionalStyle {
                fill: Some(FillStyle {
                    pattern: PatternType::Solid,
                    fg_color: Some(StyleColor::Rgb("FF00FF00".to_string())),
                    bg_color: None,
                }),
                ..ConditionalStyle::default()
            }),
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "C1:C10", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.operator, Some("equal".to_string()));
        assert_eq!(rule.formulas, vec!["\"Yes\"".to_string()]);
    }

    #[test]
    fn test_cell_is_less_than() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::CellIs {
                operator: CfOperator::LessThan,
                formula: "0".to_string(),
                formula2: None,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "D1:D10", &rules).unwrap();
        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.operator, Some("lessThan".to_string()));
    }

    #[test]
    fn test_cell_is_not_between() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::CellIs {
                operator: CfOperator::NotBetween,
                formula: "1".to_string(),
                formula2: Some("10".to_string()),
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "E1:E10", &rules).unwrap();
        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.operator, Some("notBetween".to_string()));
        assert_eq!(rule.formulas.len(), 2);
    }

    // -----------------------------------------------------------------------
    // Expression tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_expression_rule() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::Expression {
                formula: "MOD(ROW(),2)=0".to_string(),
            },
            format: Some(ConditionalStyle {
                fill: Some(FillStyle {
                    pattern: PatternType::Solid,
                    fg_color: Some(StyleColor::Rgb("FFEEEEEE".to_string())),
                    bg_color: None,
                }),
                ..ConditionalStyle::default()
            }),
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:Z100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "expression");
        assert_eq!(rule.formulas, vec!["MOD(ROW(),2)=0".to_string()]);
        assert!(rule.dxf_id.is_some());
    }

    // -----------------------------------------------------------------------
    // ColorScale tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_two_color_scale() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::ColorScale {
                min_type: CfValueType::Min,
                min_value: None,
                min_color: "FFF8696B".to_string(),
                mid_type: None,
                mid_value: None,
                mid_color: None,
                max_type: CfValueType::Max,
                max_value: None,
                max_color: "FF63BE7B".to_string(),
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "colorScale");
        assert!(rule.dxf_id.is_none());
        let cs = rule.color_scale.as_ref().unwrap();
        assert_eq!(cs.cfvos.len(), 2);
        assert_eq!(cs.colors.len(), 2);
        assert_eq!(cs.cfvos[0].value_type, "min");
        assert_eq!(cs.cfvos[1].value_type, "max");
        assert_eq!(cs.colors[0].rgb, Some("FFF8696B".to_string()));
        assert_eq!(cs.colors[1].rgb, Some("FF63BE7B".to_string()));
    }

    #[test]
    fn test_three_color_scale() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::ColorScale {
                min_type: CfValueType::Min,
                min_value: None,
                min_color: "FFF8696B".to_string(),
                mid_type: Some(CfValueType::Percentile),
                mid_value: Some("50".to_string()),
                mid_color: Some("FFFFEB84".to_string()),
                max_type: CfValueType::Max,
                max_value: None,
                max_color: "FF63BE7B".to_string(),
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "B1:B100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        let cs = rule.color_scale.as_ref().unwrap();
        assert_eq!(cs.cfvos.len(), 3);
        assert_eq!(cs.colors.len(), 3);
        assert_eq!(cs.cfvos[1].value_type, "percentile");
        assert_eq!(cs.cfvos[1].val, Some("50".to_string()));
        assert_eq!(cs.colors[1].rgb, Some("FFFFEB84".to_string()));
    }

    // -----------------------------------------------------------------------
    // DataBar tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_data_bar() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::DataBar {
                min_type: CfValueType::Min,
                min_value: None,
                max_type: CfValueType::Max,
                max_value: None,
                color: "FF638EC6".to_string(),
                show_value: true,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "C1:C50", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "dataBar");
        let db = rule.data_bar.as_ref().unwrap();
        assert!(db.show_value.is_none()); // true is default, not serialized
        assert_eq!(db.cfvos.len(), 2);
        assert_eq!(db.color.as_ref().unwrap().rgb, Some("FF638EC6".to_string()));
    }

    #[test]
    fn test_data_bar_hidden_value() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::DataBar {
                min_type: CfValueType::Num,
                min_value: Some("0".to_string()),
                max_type: CfValueType::Num,
                max_value: Some("100".to_string()),
                color: "FF638EC6".to_string(),
                show_value: false,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "D1:D50", &rules).unwrap();

        let db = ws.conditional_formatting[0].cf_rules[0]
            .data_bar
            .as_ref()
            .unwrap();
        assert_eq!(db.show_value, Some(false));
        assert_eq!(db.cfvos[0].value_type, "num");
        assert_eq!(db.cfvos[0].val, Some("0".to_string()));
    }

    // -----------------------------------------------------------------------
    // DuplicateValues / UniqueValues tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_duplicate_values() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::DuplicateValues,
            format: Some(ConditionalStyle {
                fill: Some(FillStyle {
                    pattern: PatternType::Solid,
                    fg_color: Some(StyleColor::Rgb("FFFF0000".to_string())),
                    bg_color: None,
                }),
                ..ConditionalStyle::default()
            }),
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "duplicateValues");
        assert!(rule.dxf_id.is_some());
    }

    #[test]
    fn test_unique_values() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::UniqueValues,
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "B1:B50", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "uniqueValues");
    }

    // -----------------------------------------------------------------------
    // Top10 / Bottom10 tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_top_10() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::Top10 {
                rank: 5,
                percent: false,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "top10");
        assert_eq!(rule.rank, Some(5));
        assert!(rule.percent.is_none());
        assert!(rule.bottom.is_none());
    }

    #[test]
    fn test_top_10_percent() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::Top10 {
                rank: 10,
                percent: true,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.percent, Some(true));
        assert!(rule.bottom.is_none());
    }

    #[test]
    fn test_bottom_10() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::Bottom10 {
                rank: 3,
                percent: false,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "top10"); // Bottom uses top10 type with bottom=true
        assert_eq!(rule.rank, Some(3));
        assert_eq!(rule.bottom, Some(true));
    }

    // -----------------------------------------------------------------------
    // AboveAverage tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_above_average() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::AboveAverage {
                above: true,
                equal_average: false,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "aboveAverage");
        assert!(rule.above_average.is_none()); // true is default
        assert!(rule.equal_average.is_none()); // false is default
    }

    #[test]
    fn test_below_average() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::AboveAverage {
                above: false,
                equal_average: true,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.above_average, Some(false));
        assert_eq!(rule.equal_average, Some(true));
    }

    // -----------------------------------------------------------------------
    // Text-based tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_contains_text() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::ContainsText {
                text: "error".to_string(),
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "containsText");
        assert_eq!(rule.text, Some("error".to_string()));
        assert_eq!(rule.operator, Some("containsText".to_string()));
    }

    #[test]
    fn test_not_contains_text() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::NotContainsText {
                text: "done".to_string(),
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "notContainsText");
        assert_eq!(rule.operator, Some("notContains".to_string()));
    }

    #[test]
    fn test_begins_with() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::BeginsWith {
                text: "Total".to_string(),
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "beginsWith");
        assert_eq!(rule.text, Some("Total".to_string()));
    }

    #[test]
    fn test_ends_with() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::EndsWith {
                text: "Inc.".to_string(),
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let rule = &ws.conditional_formatting[0].cf_rules[0];
        assert_eq!(rule.rule_type, "endsWith");
        assert_eq!(rule.text, Some("Inc.".to_string()));
    }

    // -----------------------------------------------------------------------
    // Blanks / Errors tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_contains_blanks() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::ContainsBlanks,
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();
        assert_eq!(
            ws.conditional_formatting[0].cf_rules[0].rule_type,
            "containsBlanks"
        );
    }

    #[test]
    fn test_not_contains_blanks() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::NotContainsBlanks,
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();
        assert_eq!(
            ws.conditional_formatting[0].cf_rules[0].rule_type,
            "notContainsBlanks"
        );
    }

    #[test]
    fn test_contains_errors() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::ContainsErrors,
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();
        assert_eq!(
            ws.conditional_formatting[0].cf_rules[0].rule_type,
            "containsErrors"
        );
    }

    #[test]
    fn test_not_contains_errors() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();
        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::NotContainsErrors,
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();
        assert_eq!(
            ws.conditional_formatting[0].cf_rules[0].rule_type,
            "notContainsErrors"
        );
    }

    // -----------------------------------------------------------------------
    // Delete tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_delete_conditional_format() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules1 = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::DuplicateValues,
            format: None,
            priority: None,
            stop_if_true: false,
        }];
        let rules2 = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::UniqueValues,
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules1).unwrap();
        set_conditional_format(&mut ws, &mut ss, "B1:B100", &rules2).unwrap();
        assert_eq!(ws.conditional_formatting.len(), 2);

        delete_conditional_format(&mut ws, "A1:A100").unwrap();
        assert_eq!(ws.conditional_formatting.len(), 1);
        assert_eq!(ws.conditional_formatting[0].sqref, "B1:B100");
    }

    #[test]
    fn test_delete_nonexistent_conditional_format() {
        let mut ws = WorksheetXml::default();
        delete_conditional_format(&mut ws, "Z1:Z99").unwrap();
        assert!(ws.conditional_formatting.is_empty());
    }

    // -----------------------------------------------------------------------
    // Multiple rules on same range
    // -----------------------------------------------------------------------

    #[test]
    fn test_multiple_rules_same_range() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![
            ConditionalFormatRule {
                rule_type: ConditionalFormatType::CellIs {
                    operator: CfOperator::GreaterThan,
                    formula: "90".to_string(),
                    formula2: None,
                },
                format: Some(ConditionalStyle {
                    fill: Some(FillStyle {
                        pattern: PatternType::Solid,
                        fg_color: Some(StyleColor::Rgb("FF00FF00".to_string())),
                        bg_color: None,
                    }),
                    ..ConditionalStyle::default()
                }),
                priority: None,
                stop_if_true: false,
            },
            ConditionalFormatRule {
                rule_type: ConditionalFormatType::CellIs {
                    operator: CfOperator::LessThan,
                    formula: "50".to_string(),
                    formula2: None,
                },
                format: Some(ConditionalStyle {
                    fill: Some(FillStyle {
                        pattern: PatternType::Solid,
                        fg_color: Some(StyleColor::Rgb("FFFF0000".to_string())),
                        bg_color: None,
                    }),
                    ..ConditionalStyle::default()
                }),
                priority: None,
                stop_if_true: false,
            },
        ];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        assert_eq!(ws.conditional_formatting.len(), 1);
        assert_eq!(ws.conditional_formatting[0].cf_rules.len(), 2);

        let dxfs = ss.dxfs.as_ref().unwrap();
        assert_eq!(dxfs.dxfs.len(), 2);
    }

    // -----------------------------------------------------------------------
    // Get (roundtrip) tests
    // -----------------------------------------------------------------------

    #[test]
    fn test_get_conditional_formats_cell_is() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::CellIs {
                operator: CfOperator::GreaterThanOrEqual,
                formula: "50".to_string(),
                formula2: None,
            },
            format: Some(ConditionalStyle {
                font: Some(FontStyle {
                    bold: true,
                    ..FontStyle::default()
                }),
                ..ConditionalStyle::default()
            }),
            priority: None,
            stop_if_true: true,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let formats = get_conditional_formats(&ws, &ss);
        assert_eq!(formats.len(), 1);
        assert_eq!(formats[0].0, "A1:A100");
        assert_eq!(formats[0].1.len(), 1);

        let rule = &formats[0].1[0];
        match &rule.rule_type {
            ConditionalFormatType::CellIs {
                operator,
                formula,
                formula2,
            } => {
                assert_eq!(*operator, CfOperator::GreaterThanOrEqual);
                assert_eq!(formula, "50");
                assert!(formula2.is_none());
            }
            _ => panic!("expected CellIs rule type"),
        }
        assert!(rule.stop_if_true);
        assert!(rule.format.is_some());
        assert!(rule.format.as_ref().unwrap().font.as_ref().unwrap().bold);
    }

    #[test]
    fn test_get_conditional_formats_color_scale() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::ColorScale {
                min_type: CfValueType::Min,
                min_value: None,
                min_color: "FFF8696B".to_string(),
                mid_type: Some(CfValueType::Percentile),
                mid_value: Some("50".to_string()),
                mid_color: Some("FFFFEB84".to_string()),
                max_type: CfValueType::Max,
                max_value: None,
                max_color: "FF63BE7B".to_string(),
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let formats = get_conditional_formats(&ws, &ss);
        let rule = &formats[0].1[0];
        match &rule.rule_type {
            ConditionalFormatType::ColorScale {
                min_type,
                min_color,
                mid_type,
                mid_color,
                max_type,
                max_color,
                ..
            } => {
                assert_eq!(*min_type, CfValueType::Min);
                assert_eq!(min_color, "FFF8696B");
                assert_eq!(*mid_type, Some(CfValueType::Percentile));
                assert_eq!(*mid_color, Some("FFFFEB84".to_string()));
                assert_eq!(*max_type, CfValueType::Max);
                assert_eq!(max_color, "FF63BE7B");
            }
            _ => panic!("expected ColorScale rule type"),
        }
    }

    #[test]
    fn test_get_conditional_formats_data_bar() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::DataBar {
                min_type: CfValueType::Min,
                min_value: None,
                max_type: CfValueType::Max,
                max_value: None,
                color: "FF638EC6".to_string(),
                show_value: true,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let formats = get_conditional_formats(&ws, &ss);
        let rule = &formats[0].1[0];
        match &rule.rule_type {
            ConditionalFormatType::DataBar {
                min_type,
                max_type,
                color,
                show_value,
                ..
            } => {
                assert_eq!(*min_type, CfValueType::Min);
                assert_eq!(*max_type, CfValueType::Max);
                assert_eq!(color, "FF638EC6");
                assert!(*show_value);
            }
            _ => panic!("expected DataBar rule type"),
        }
    }

    #[test]
    fn test_get_conditional_formats_duplicate_values() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::DuplicateValues,
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let formats = get_conditional_formats(&ws, &ss);
        let rule = &formats[0].1[0];
        assert_eq!(rule.rule_type, ConditionalFormatType::DuplicateValues);
    }

    #[test]
    fn test_get_conditional_formats_top10() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::Top10 {
                rank: 5,
                percent: true,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let formats = get_conditional_formats(&ws, &ss);
        let rule = &formats[0].1[0];
        match &rule.rule_type {
            ConditionalFormatType::Top10 { rank, percent } => {
                assert_eq!(*rank, 5);
                assert!(*percent);
            }
            _ => panic!("expected Top10 rule type"),
        }
    }

    #[test]
    fn test_get_conditional_formats_bottom10() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::Bottom10 {
                rank: 3,
                percent: false,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let formats = get_conditional_formats(&ws, &ss);
        let rule = &formats[0].1[0];
        match &rule.rule_type {
            ConditionalFormatType::Bottom10 { rank, percent } => {
                assert_eq!(*rank, 3);
                assert!(!(*percent));
            }
            _ => panic!("expected Bottom10 rule type"),
        }
    }

    #[test]
    fn test_get_conditional_formats_above_average() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::AboveAverage {
                above: false,
                equal_average: true,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let formats = get_conditional_formats(&ws, &ss);
        let rule = &formats[0].1[0];
        match &rule.rule_type {
            ConditionalFormatType::AboveAverage {
                above,
                equal_average,
            } => {
                assert!(!above);
                assert!(*equal_average);
            }
            _ => panic!("expected AboveAverage rule type"),
        }
    }

    #[test]
    fn test_get_conditional_formats_empty() {
        let ws = WorksheetXml::default();
        let ss = default_stylesheet();
        let formats = get_conditional_formats(&ws, &ss);
        assert!(formats.is_empty());
    }

    // -----------------------------------------------------------------------
    // Priority auto-increment
    // -----------------------------------------------------------------------

    #[test]
    fn test_priority_auto_increment() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules1 = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::DuplicateValues,
            format: None,
            priority: None,
            stop_if_true: false,
        }];
        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules1).unwrap();

        let rules2 = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::UniqueValues,
            format: None,
            priority: None,
            stop_if_true: false,
        }];
        set_conditional_format(&mut ws, &mut ss, "B1:B100", &rules2).unwrap();

        assert_eq!(ws.conditional_formatting[0].cf_rules[0].priority, 1);
        assert_eq!(ws.conditional_formatting[1].cf_rules[0].priority, 2);
    }

    #[test]
    fn test_explicit_priority() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::DuplicateValues,
            format: None,
            priority: Some(42),
            stop_if_true: false,
        }];
        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        assert_eq!(ws.conditional_formatting[0].cf_rules[0].priority, 42);
    }

    // -----------------------------------------------------------------------
    // Stop if true
    // -----------------------------------------------------------------------

    #[test]
    fn test_stop_if_true() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::DuplicateValues,
            format: None,
            priority: None,
            stop_if_true: true,
        }];
        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        assert_eq!(
            ws.conditional_formatting[0].cf_rules[0].stop_if_true,
            Some(true)
        );
    }

    // -----------------------------------------------------------------------
    // DXF style roundtrip
    // -----------------------------------------------------------------------

    #[test]
    fn test_dxf_style_roundtrip_font() {
        let style = ConditionalStyle {
            font: Some(FontStyle {
                bold: true,
                italic: true,
                color: Some(StyleColor::Rgb("FFFF0000".to_string())),
                size: Some(14.0),
                name: Some("Arial".to_string()),
                ..FontStyle::default()
            }),
            ..ConditionalStyle::default()
        };

        let dxf = conditional_style_to_dxf(&style);
        let roundtrip = dxf_to_conditional_style(&dxf);

        let font = roundtrip.font.unwrap();
        assert!(font.bold);
        assert!(font.italic);
        assert_eq!(font.color, Some(StyleColor::Rgb("FFFF0000".to_string())));
        assert_eq!(font.size, Some(14.0));
        assert_eq!(font.name, Some("Arial".to_string()));
    }

    #[test]
    fn test_dxf_style_roundtrip_fill() {
        let style = ConditionalStyle {
            fill: Some(FillStyle {
                pattern: PatternType::Solid,
                fg_color: Some(StyleColor::Rgb("FFFFFF00".to_string())),
                bg_color: None,
            }),
            ..ConditionalStyle::default()
        };

        let dxf = conditional_style_to_dxf(&style);
        let roundtrip = dxf_to_conditional_style(&dxf);

        let fill = roundtrip.fill.unwrap();
        assert_eq!(fill.pattern, PatternType::Solid);
        assert_eq!(fill.fg_color, Some(StyleColor::Rgb("FFFFFF00".to_string())));
    }

    #[test]
    fn test_dxf_style_roundtrip_border() {
        let style = ConditionalStyle {
            border: Some(BorderStyle {
                left: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: Some(StyleColor::Rgb("FF000000".to_string())),
                }),
                ..BorderStyle::default()
            }),
            ..ConditionalStyle::default()
        };

        let dxf = conditional_style_to_dxf(&style);
        let roundtrip = dxf_to_conditional_style(&dxf);

        let border = roundtrip.border.unwrap();
        let left = border.left.unwrap();
        assert_eq!(left.style, BorderLineStyle::Thin);
        assert_eq!(left.color, Some(StyleColor::Rgb("FF000000".to_string())));
    }

    // -----------------------------------------------------------------------
    // XML serialization roundtrip
    // -----------------------------------------------------------------------

    #[test]
    fn test_xml_serialization_roundtrip() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::CellIs {
                operator: CfOperator::GreaterThan,
                formula: "100".to_string(),
                formula2: None,
            },
            format: Some(ConditionalStyle {
                font: Some(FontStyle {
                    bold: true,
                    ..FontStyle::default()
                }),
                ..ConditionalStyle::default()
            }),
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let xml = quick_xml::se::to_string(&ws).unwrap();
        assert!(xml.contains("conditionalFormatting"));
        assert!(xml.contains("cfRule"));
        assert!(xml.contains("cellIs"));

        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.conditional_formatting.len(), 1);
        assert_eq!(parsed.conditional_formatting[0].sqref, "A1:A100");
        assert_eq!(parsed.conditional_formatting[0].cf_rules.len(), 1);
        assert_eq!(
            parsed.conditional_formatting[0].cf_rules[0].rule_type,
            "cellIs"
        );
    }

    #[test]
    fn test_color_scale_xml_roundtrip() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::ColorScale {
                min_type: CfValueType::Min,
                min_value: None,
                min_color: "FFF8696B".to_string(),
                mid_type: None,
                mid_value: None,
                mid_color: None,
                max_type: CfValueType::Max,
                max_value: None,
                max_color: "FF63BE7B".to_string(),
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let xml = quick_xml::se::to_string(&ws).unwrap();
        assert!(xml.contains("colorScale"));

        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        let cs = parsed.conditional_formatting[0].cf_rules[0]
            .color_scale
            .as_ref()
            .unwrap();
        assert_eq!(cs.cfvos.len(), 2);
        assert_eq!(cs.colors.len(), 2);
    }

    #[test]
    fn test_data_bar_xml_roundtrip() {
        let mut ws = WorksheetXml::default();
        let mut ss = default_stylesheet();

        let rules = vec![ConditionalFormatRule {
            rule_type: ConditionalFormatType::DataBar {
                min_type: CfValueType::Min,
                min_value: None,
                max_type: CfValueType::Max,
                max_value: None,
                color: "FF638EC6".to_string(),
                show_value: true,
            },
            format: None,
            priority: None,
            stop_if_true: false,
        }];

        set_conditional_format(&mut ws, &mut ss, "A1:A100", &rules).unwrap();

        let xml = quick_xml::se::to_string(&ws).unwrap();
        assert!(xml.contains("dataBar"));

        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        let db = parsed.conditional_formatting[0].cf_rules[0]
            .data_bar
            .as_ref()
            .unwrap();
        assert_eq!(db.cfvos.len(), 2);
    }

    // -----------------------------------------------------------------------
    // CfOperator / CfValueType parse roundtrip
    // -----------------------------------------------------------------------

    #[test]
    fn test_cf_operator_roundtrip() {
        let operators = [
            CfOperator::LessThan,
            CfOperator::LessThanOrEqual,
            CfOperator::Equal,
            CfOperator::NotEqual,
            CfOperator::GreaterThanOrEqual,
            CfOperator::GreaterThan,
            CfOperator::Between,
            CfOperator::NotBetween,
        ];
        for op in &operators {
            let s = op.as_str();
            let parsed = CfOperator::parse(s).unwrap();
            assert_eq!(*op, parsed);
        }
    }

    #[test]
    fn test_cf_value_type_roundtrip() {
        let types = [
            CfValueType::Num,
            CfValueType::Percent,
            CfValueType::Min,
            CfValueType::Max,
            CfValueType::Percentile,
            CfValueType::Formula,
        ];
        for vt in &types {
            let s = vt.as_str();
            let parsed = CfValueType::parse(s).unwrap();
            assert_eq!(*vt, parsed);
        }
    }
}
