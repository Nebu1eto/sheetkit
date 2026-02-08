//! Style builder and runtime management.
//!
//! Provides high-level, ergonomic style types that map to the low-level XML
//! stylesheet structures in `sheetkit-xml`. Styles are registered in the
//! stylesheet with deduplication: identical style components share the same
//! index.

use sheetkit_xml::styles::{
    Alignment, Border, BorderSide, Borders, Color, Fill, Fills, Font, Fonts, NumFmt, NumFmts,
    PatternFill, Protection, StyleSheet, Xf,
};

use crate::error::{Error, Result};

/// Maximum number of cell XFs Excel supports.
const MAX_CELL_XFS: usize = 65430;

/// First ID available for custom number formats (0-163 are reserved).
const CUSTOM_NUM_FMT_BASE: u32 = 164;

/// Built-in number format IDs.
pub mod builtin_num_fmts {
    /// General
    pub const GENERAL: u32 = 0;
    /// 0
    pub const INTEGER: u32 = 1;
    /// 0.00
    pub const DECIMAL_2: u32 = 2;
    /// #,##0
    pub const THOUSANDS: u32 = 3;
    /// #,##0.00
    pub const THOUSANDS_DECIMAL: u32 = 4;
    /// 0%
    pub const PERCENT: u32 = 9;
    /// 0.00%
    pub const PERCENT_DECIMAL: u32 = 10;
    /// 0.00E+00
    pub const SCIENTIFIC: u32 = 11;
    /// m/d/yyyy
    pub const DATE_MDY: u32 = 14;
    /// d-mmm-yy
    pub const DATE_DMY: u32 = 15;
    /// d-mmm
    pub const DATE_DM: u32 = 16;
    /// mmm-yy
    pub const DATE_MY: u32 = 17;
    /// h:mm AM/PM
    pub const TIME_HM_AP: u32 = 18;
    /// h:mm:ss AM/PM
    pub const TIME_HMS_AP: u32 = 19;
    /// h:mm
    pub const TIME_HM: u32 = 20;
    /// h:mm:ss
    pub const TIME_HMS: u32 = 21;
    /// m/d/yyyy h:mm
    pub const DATETIME: u32 = 22;
    /// @
    pub const TEXT: u32 = 49;
}

/// User-facing style definition.
#[derive(Debug, Clone, Default)]
pub struct Style {
    pub font: Option<FontStyle>,
    pub fill: Option<FillStyle>,
    pub border: Option<BorderStyle>,
    pub alignment: Option<AlignmentStyle>,
    pub num_fmt: Option<NumFmtStyle>,
    pub protection: Option<ProtectionStyle>,
}

/// Font style definition.
#[derive(Debug, Clone, Default)]
pub struct FontStyle {
    /// Font name, e.g. "Calibri", "Arial".
    pub name: Option<String>,
    /// Font size, e.g. 11.0.
    pub size: Option<f64>,
    /// Bold.
    pub bold: bool,
    /// Italic.
    pub italic: bool,
    /// Underline.
    pub underline: bool,
    /// Strikethrough.
    pub strikethrough: bool,
    /// Font color.
    pub color: Option<StyleColor>,
}

/// Color specification.
#[derive(Debug, Clone, PartialEq)]
pub enum StyleColor {
    /// ARGB hex color, e.g. "FF0000FF".
    Rgb(String),
    /// Theme color index.
    Theme(u32),
    /// Indexed color.
    Indexed(u32),
}

/// Fill style definition.
#[derive(Debug, Clone)]
pub struct FillStyle {
    /// Pattern type.
    pub pattern: PatternType,
    /// Foreground color.
    pub fg_color: Option<StyleColor>,
    /// Background color.
    pub bg_color: Option<StyleColor>,
    /// Gradient fill (mutually exclusive with pattern fill when gradient is set).
    pub gradient: Option<GradientFillStyle>,
}

/// Gradient fill style definition.
#[derive(Debug, Clone)]
pub struct GradientFillStyle {
    /// Gradient type: linear or path.
    pub gradient_type: GradientType,
    /// Rotation angle in degrees for linear gradients.
    pub degree: Option<f64>,
    /// Left coordinate for path gradients (0.0-1.0).
    pub left: Option<f64>,
    /// Right coordinate for path gradients (0.0-1.0).
    pub right: Option<f64>,
    /// Top coordinate for path gradients (0.0-1.0).
    pub top: Option<f64>,
    /// Bottom coordinate for path gradients (0.0-1.0).
    pub bottom: Option<f64>,
    /// Gradient stops defining the color transitions.
    pub stops: Vec<GradientStop>,
}

/// Gradient type.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum GradientType {
    Linear,
    Path,
}

impl GradientType {
    fn from_str(s: &str) -> Self {
        match s {
            "path" => GradientType::Path,
            _ => GradientType::Linear,
        }
    }
}

/// A single stop in a gradient fill.
#[derive(Debug, Clone)]
pub struct GradientStop {
    /// Position of this stop (0.0-1.0).
    pub position: f64,
    /// Color at this stop.
    pub color: StyleColor,
}

/// Pattern fill type.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum PatternType {
    None,
    Solid,
    Gray125,
    DarkGray,
    MediumGray,
    LightGray,
}

impl PatternType {
    fn as_str(&self) -> &str {
        match self {
            PatternType::None => "none",
            PatternType::Solid => "solid",
            PatternType::Gray125 => "gray125",
            PatternType::DarkGray => "darkGray",
            PatternType::MediumGray => "mediumGray",
            PatternType::LightGray => "lightGray",
        }
    }

    fn from_str(s: &str) -> Self {
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
}

/// Border style definition.
#[derive(Debug, Clone, Default)]
pub struct BorderStyle {
    pub left: Option<BorderSideStyle>,
    pub right: Option<BorderSideStyle>,
    pub top: Option<BorderSideStyle>,
    pub bottom: Option<BorderSideStyle>,
    pub diagonal: Option<BorderSideStyle>,
}

/// Border side style definition.
#[derive(Debug, Clone)]
pub struct BorderSideStyle {
    pub style: BorderLineStyle,
    pub color: Option<StyleColor>,
}

/// Border line style.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BorderLineStyle {
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

impl BorderLineStyle {
    fn as_str(&self) -> &str {
        match self {
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

    fn from_str(s: &str) -> Option<Self> {
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
}

/// Alignment style definition.
#[derive(Debug, Clone, Default)]
pub struct AlignmentStyle {
    pub horizontal: Option<HorizontalAlign>,
    pub vertical: Option<VerticalAlign>,
    pub wrap_text: bool,
    pub text_rotation: Option<u32>,
    pub indent: Option<u32>,
    pub shrink_to_fit: bool,
}

/// Horizontal alignment.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum HorizontalAlign {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

impl HorizontalAlign {
    fn as_str(&self) -> &str {
        match self {
            HorizontalAlign::General => "general",
            HorizontalAlign::Left => "left",
            HorizontalAlign::Center => "center",
            HorizontalAlign::Right => "right",
            HorizontalAlign::Fill => "fill",
            HorizontalAlign::Justify => "justify",
            HorizontalAlign::CenterContinuous => "centerContinuous",
            HorizontalAlign::Distributed => "distributed",
        }
    }

    fn from_str(s: &str) -> Option<Self> {
        match s {
            "general" => Some(HorizontalAlign::General),
            "left" => Some(HorizontalAlign::Left),
            "center" => Some(HorizontalAlign::Center),
            "right" => Some(HorizontalAlign::Right),
            "fill" => Some(HorizontalAlign::Fill),
            "justify" => Some(HorizontalAlign::Justify),
            "centerContinuous" => Some(HorizontalAlign::CenterContinuous),
            "distributed" => Some(HorizontalAlign::Distributed),
            _ => None,
        }
    }
}

/// Vertical alignment.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum VerticalAlign {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

impl VerticalAlign {
    fn as_str(&self) -> &str {
        match self {
            VerticalAlign::Top => "top",
            VerticalAlign::Center => "center",
            VerticalAlign::Bottom => "bottom",
            VerticalAlign::Justify => "justify",
            VerticalAlign::Distributed => "distributed",
        }
    }

    fn from_str(s: &str) -> Option<Self> {
        match s {
            "top" => Some(VerticalAlign::Top),
            "center" => Some(VerticalAlign::Center),
            "bottom" => Some(VerticalAlign::Bottom),
            "justify" => Some(VerticalAlign::Justify),
            "distributed" => Some(VerticalAlign::Distributed),
            _ => None,
        }
    }
}

/// Number format style.
#[derive(Debug, Clone)]
pub enum NumFmtStyle {
    /// Built-in format ID (0-49).
    Builtin(u32),
    /// Custom format code string.
    Custom(String),
}

/// Protection style definition.
#[derive(Debug, Clone)]
pub struct ProtectionStyle {
    pub locked: bool,
    pub hidden: bool,
}

/// Builder for creating Style objects with a fluent API.
///
/// Each setter method initializes the relevant sub-struct if it has not been
/// set yet, then applies the value. Call `build()` to obtain the final Style.
pub struct StyleBuilder {
    style: Style,
}

impl StyleBuilder {
    /// Create a new StyleBuilder with all fields set to None.
    pub fn new() -> Self {
        Self {
            style: Style::default(),
        }
    }

    // -- Font methods --

    /// Set the bold flag on the font.
    pub fn bold(mut self, bold: bool) -> Self {
        self.style.font.get_or_insert_with(FontStyle::default).bold = bold;
        self
    }

    /// Set the italic flag on the font.
    pub fn italic(mut self, italic: bool) -> Self {
        self.style
            .font
            .get_or_insert_with(FontStyle::default)
            .italic = italic;
        self
    }

    /// Set the underline flag on the font.
    pub fn underline(mut self, underline: bool) -> Self {
        self.style
            .font
            .get_or_insert_with(FontStyle::default)
            .underline = underline;
        self
    }

    /// Set the strikethrough flag on the font.
    pub fn strikethrough(mut self, strikethrough: bool) -> Self {
        self.style
            .font
            .get_or_insert_with(FontStyle::default)
            .strikethrough = strikethrough;
        self
    }

    /// Set the font name (e.g. "Arial", "Calibri").
    pub fn font_name(mut self, name: &str) -> Self {
        self.style.font.get_or_insert_with(FontStyle::default).name = Some(name.to_string());
        self
    }

    /// Set the font size in points.
    pub fn font_size(mut self, size: f64) -> Self {
        self.style.font.get_or_insert_with(FontStyle::default).size = Some(size);
        self
    }

    /// Set the font color using a StyleColor value.
    pub fn font_color(mut self, color: StyleColor) -> Self {
        self.style.font.get_or_insert_with(FontStyle::default).color = Some(color);
        self
    }

    /// Set the font color using an ARGB hex string (e.g. "FF0000FF").
    pub fn font_color_rgb(self, rgb: &str) -> Self {
        self.font_color(StyleColor::Rgb(rgb.to_string()))
    }

    // -- Fill methods --

    /// Set the fill pattern type.
    pub fn fill_pattern(mut self, pattern: PatternType) -> Self {
        self.style
            .fill
            .get_or_insert(FillStyle {
                pattern: PatternType::None,
                fg_color: None,
                bg_color: None,
                gradient: None,
            })
            .pattern = pattern;
        self
    }

    /// Set the fill foreground color.
    pub fn fill_fg_color(mut self, color: StyleColor) -> Self {
        self.style
            .fill
            .get_or_insert(FillStyle {
                pattern: PatternType::None,
                fg_color: None,
                bg_color: None,
                gradient: None,
            })
            .fg_color = Some(color);
        self
    }

    /// Set the fill foreground color using an ARGB hex string.
    pub fn fill_fg_color_rgb(self, rgb: &str) -> Self {
        self.fill_fg_color(StyleColor::Rgb(rgb.to_string()))
    }

    /// Set the fill background color.
    pub fn fill_bg_color(mut self, color: StyleColor) -> Self {
        self.style
            .fill
            .get_or_insert(FillStyle {
                pattern: PatternType::None,
                fg_color: None,
                bg_color: None,
                gradient: None,
            })
            .bg_color = Some(color);
        self
    }

    /// Convenience method: set a solid fill with the given ARGB foreground color.
    pub fn solid_fill(mut self, rgb: &str) -> Self {
        self.style.fill = Some(FillStyle {
            pattern: PatternType::Solid,
            fg_color: Some(StyleColor::Rgb(rgb.to_string())),
            bg_color: self.style.fill.and_then(|f| f.bg_color),
            gradient: None,
        });
        self
    }

    // -- Border methods --

    /// Set the left border style and color.
    pub fn border_left(mut self, style: BorderLineStyle, color: StyleColor) -> Self {
        self.style
            .border
            .get_or_insert_with(BorderStyle::default)
            .left = Some(BorderSideStyle {
            style,
            color: Some(color),
        });
        self
    }

    /// Set the right border style and color.
    pub fn border_right(mut self, style: BorderLineStyle, color: StyleColor) -> Self {
        self.style
            .border
            .get_or_insert_with(BorderStyle::default)
            .right = Some(BorderSideStyle {
            style,
            color: Some(color),
        });
        self
    }

    /// Set the top border style and color.
    pub fn border_top(mut self, style: BorderLineStyle, color: StyleColor) -> Self {
        self.style
            .border
            .get_or_insert_with(BorderStyle::default)
            .top = Some(BorderSideStyle {
            style,
            color: Some(color),
        });
        self
    }

    /// Set the bottom border style and color.
    pub fn border_bottom(mut self, style: BorderLineStyle, color: StyleColor) -> Self {
        self.style
            .border
            .get_or_insert_with(BorderStyle::default)
            .bottom = Some(BorderSideStyle {
            style,
            color: Some(color),
        });
        self
    }

    /// Set all four border sides (left, right, top, bottom) to the same style and color.
    pub fn border_all(mut self, style: BorderLineStyle, color: StyleColor) -> Self {
        let side = || BorderSideStyle {
            style,
            color: Some(color.clone()),
        };
        let border = self.style.border.get_or_insert_with(BorderStyle::default);
        border.left = Some(side());
        border.right = Some(side());
        border.top = Some(side());
        border.bottom = Some(side());
        self
    }

    // -- Alignment methods --

    /// Set horizontal alignment.
    pub fn horizontal_align(mut self, align: HorizontalAlign) -> Self {
        self.style
            .alignment
            .get_or_insert_with(AlignmentStyle::default)
            .horizontal = Some(align);
        self
    }

    /// Set vertical alignment.
    pub fn vertical_align(mut self, align: VerticalAlign) -> Self {
        self.style
            .alignment
            .get_or_insert_with(AlignmentStyle::default)
            .vertical = Some(align);
        self
    }

    /// Set the wrap text flag.
    pub fn wrap_text(mut self, wrap: bool) -> Self {
        self.style
            .alignment
            .get_or_insert_with(AlignmentStyle::default)
            .wrap_text = wrap;
        self
    }

    /// Set text rotation in degrees.
    pub fn text_rotation(mut self, degrees: u32) -> Self {
        self.style
            .alignment
            .get_or_insert_with(AlignmentStyle::default)
            .text_rotation = Some(degrees);
        self
    }

    /// Set the indent level.
    pub fn indent(mut self, indent: u32) -> Self {
        self.style
            .alignment
            .get_or_insert_with(AlignmentStyle::default)
            .indent = Some(indent);
        self
    }

    /// Set the shrink to fit flag.
    pub fn shrink_to_fit(mut self, shrink: bool) -> Self {
        self.style
            .alignment
            .get_or_insert_with(AlignmentStyle::default)
            .shrink_to_fit = shrink;
        self
    }

    // -- Number format methods --

    /// Set a built-in number format by ID (see `builtin_num_fmts` constants).
    pub fn num_format_builtin(mut self, id: u32) -> Self {
        self.style.num_fmt = Some(NumFmtStyle::Builtin(id));
        self
    }

    /// Set a custom number format string (e.g. "#,##0.00").
    pub fn num_format_custom(mut self, format: &str) -> Self {
        self.style.num_fmt = Some(NumFmtStyle::Custom(format.to_string()));
        self
    }

    // -- Protection methods --

    /// Set the locked flag for cell protection.
    pub fn locked(mut self, locked: bool) -> Self {
        self.style
            .protection
            .get_or_insert(ProtectionStyle {
                locked: true,
                hidden: false,
            })
            .locked = locked;
        self
    }

    /// Set the hidden flag for cell protection.
    pub fn hidden(mut self, hidden: bool) -> Self {
        self.style
            .protection
            .get_or_insert(ProtectionStyle {
                locked: true,
                hidden: false,
            })
            .hidden = hidden;
        self
    }

    // -- Build --

    /// Consume the builder and return the constructed Style.
    pub fn build(self) -> Style {
        self.style
    }
}

impl Default for StyleBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Convert a `StyleColor` to the XML `Color` struct.
fn style_color_to_xml(color: &StyleColor) -> Color {
    match color {
        StyleColor::Rgb(rgb) => Color {
            auto: None,
            indexed: None,
            rgb: Some(rgb.clone()),
            theme: None,
            tint: None,
        },
        StyleColor::Theme(t) => Color {
            auto: None,
            indexed: None,
            rgb: None,
            theme: Some(*t),
            tint: None,
        },
        StyleColor::Indexed(i) => Color {
            auto: None,
            indexed: Some(*i),
            rgb: None,
            theme: None,
            tint: None,
        },
    }
}

/// Convert an XML `Color` back to a `StyleColor`.
fn xml_color_to_style(color: &Color) -> Option<StyleColor> {
    if let Some(ref rgb) = color.rgb {
        Some(StyleColor::Rgb(rgb.clone()))
    } else if let Some(theme) = color.theme {
        Some(StyleColor::Theme(theme))
    } else {
        color.indexed.map(StyleColor::Indexed)
    }
}

/// Convert a `FontStyle` to the XML `Font` struct.
fn font_style_to_xml(font: &FontStyle) -> Font {
    use sheetkit_xml::styles::{BoolVal, FontName, FontSize, Underline};

    Font {
        b: if font.bold {
            Some(BoolVal { val: None })
        } else {
            None
        },
        i: if font.italic {
            Some(BoolVal { val: None })
        } else {
            None
        },
        strike: if font.strikethrough {
            Some(BoolVal { val: None })
        } else {
            None
        },
        u: if font.underline {
            Some(Underline { val: None })
        } else {
            None
        },
        sz: font.size.map(|val| FontSize { val }),
        color: font.color.as_ref().map(style_color_to_xml),
        name: font.name.as_ref().map(|val| FontName { val: val.clone() }),
        family: None,
        scheme: None,
    }
}

/// Convert an XML `Font` to a `FontStyle`.
fn xml_font_to_style(font: &Font) -> FontStyle {
    FontStyle {
        name: font.name.as_ref().map(|n| n.val.clone()),
        size: font.sz.as_ref().map(|s| s.val),
        bold: font.b.is_some(),
        italic: font.i.is_some(),
        underline: font.u.is_some(),
        strikethrough: font.strike.is_some(),
        color: font.color.as_ref().and_then(xml_color_to_style),
    }
}

/// Convert a `FillStyle` to the XML `Fill` struct.
fn fill_style_to_xml(fill: &FillStyle) -> Fill {
    if let Some(ref grad) = fill.gradient {
        return Fill {
            pattern_fill: None,
            gradient_fill: Some(gradient_style_to_xml(grad)),
        };
    }
    Fill {
        pattern_fill: Some(PatternFill {
            pattern_type: Some(fill.pattern.as_str().to_string()),
            fg_color: fill.fg_color.as_ref().map(style_color_to_xml),
            bg_color: fill.bg_color.as_ref().map(style_color_to_xml),
        }),
        gradient_fill: None,
    }
}

/// Convert a `GradientFillStyle` to the XML `GradientFill` struct.
fn gradient_style_to_xml(grad: &GradientFillStyle) -> sheetkit_xml::styles::GradientFill {
    sheetkit_xml::styles::GradientFill {
        gradient_type: match grad.gradient_type {
            GradientType::Linear => None,
            GradientType::Path => Some("path".to_string()),
        },
        degree: grad.degree,
        left: grad.left,
        right: grad.right,
        top: grad.top,
        bottom: grad.bottom,
        stops: grad
            .stops
            .iter()
            .map(|s| sheetkit_xml::styles::GradientStop {
                position: s.position,
                color: style_color_to_xml(&s.color),
            })
            .collect(),
    }
}

/// Convert an XML `Fill` to a `FillStyle`.
fn xml_fill_to_style(fill: &Fill) -> Option<FillStyle> {
    if let Some(ref gf) = fill.gradient_fill {
        let gradient_type = gf
            .gradient_type
            .as_ref()
            .map(|s| GradientType::from_str(s))
            .unwrap_or(GradientType::Linear);
        let stops: Vec<GradientStop> = gf
            .stops
            .iter()
            .filter_map(|s| {
                xml_color_to_style(&s.color).map(|c| GradientStop {
                    position: s.position,
                    color: c,
                })
            })
            .collect();
        return Some(FillStyle {
            pattern: PatternType::None,
            fg_color: None,
            bg_color: None,
            gradient: Some(GradientFillStyle {
                gradient_type,
                degree: gf.degree,
                left: gf.left,
                right: gf.right,
                top: gf.top,
                bottom: gf.bottom,
                stops,
            }),
        });
    }
    let pf = fill.pattern_fill.as_ref()?;
    let pattern = pf
        .pattern_type
        .as_ref()
        .map(|s| PatternType::from_str(s))
        .unwrap_or(PatternType::None);
    Some(FillStyle {
        pattern,
        fg_color: pf.fg_color.as_ref().and_then(xml_color_to_style),
        bg_color: pf.bg_color.as_ref().and_then(xml_color_to_style),
        gradient: None,
    })
}

/// Convert a `BorderSideStyle` to the XML `BorderSide` struct.
fn border_side_to_xml(side: &BorderSideStyle) -> BorderSide {
    BorderSide {
        style: Some(side.style.as_str().to_string()),
        color: side.color.as_ref().map(style_color_to_xml),
    }
}

/// Convert an XML `BorderSide` to a `BorderSideStyle`.
fn xml_border_side_to_style(side: &BorderSide) -> Option<BorderSideStyle> {
    let style_str = side.style.as_ref()?;
    let style = BorderLineStyle::from_str(style_str)?;
    Some(BorderSideStyle {
        style,
        color: side.color.as_ref().and_then(xml_color_to_style),
    })
}

/// Convert a `BorderStyle` to the XML `Border` struct.
fn border_style_to_xml(border: &BorderStyle) -> Border {
    Border {
        diagonal_up: None,
        diagonal_down: None,
        left: border.left.as_ref().map(border_side_to_xml),
        right: border.right.as_ref().map(border_side_to_xml),
        top: border.top.as_ref().map(border_side_to_xml),
        bottom: border.bottom.as_ref().map(border_side_to_xml),
        diagonal: border.diagonal.as_ref().map(border_side_to_xml),
    }
}

/// Convert an XML `Border` to a `BorderStyle`.
fn xml_border_to_style(border: &Border) -> BorderStyle {
    BorderStyle {
        left: border.left.as_ref().and_then(xml_border_side_to_style),
        right: border.right.as_ref().and_then(xml_border_side_to_style),
        top: border.top.as_ref().and_then(xml_border_side_to_style),
        bottom: border.bottom.as_ref().and_then(xml_border_side_to_style),
        diagonal: border.diagonal.as_ref().and_then(xml_border_side_to_style),
    }
}

/// Convert an `AlignmentStyle` to the XML `Alignment` struct.
fn alignment_style_to_xml(align: &AlignmentStyle) -> Alignment {
    Alignment {
        horizontal: align.horizontal.map(|h| h.as_str().to_string()),
        vertical: align.vertical.map(|v| v.as_str().to_string()),
        wrap_text: if align.wrap_text { Some(true) } else { None },
        text_rotation: align.text_rotation,
        indent: align.indent,
        shrink_to_fit: if align.shrink_to_fit {
            Some(true)
        } else {
            None
        },
    }
}

/// Convert an XML `Alignment` to an `AlignmentStyle`.
fn xml_alignment_to_style(align: &Alignment) -> AlignmentStyle {
    AlignmentStyle {
        horizontal: align
            .horizontal
            .as_ref()
            .and_then(|s| HorizontalAlign::from_str(s)),
        vertical: align
            .vertical
            .as_ref()
            .and_then(|s| VerticalAlign::from_str(s)),
        wrap_text: align.wrap_text.unwrap_or(false),
        text_rotation: align.text_rotation,
        indent: align.indent,
        shrink_to_fit: align.shrink_to_fit.unwrap_or(false),
    }
}

/// Convert a `ProtectionStyle` to the XML `Protection` struct.
fn protection_style_to_xml(prot: &ProtectionStyle) -> Protection {
    Protection {
        locked: Some(prot.locked),
        hidden: Some(prot.hidden),
    }
}

/// Convert an XML `Protection` to a `ProtectionStyle`.
fn xml_protection_to_style(prot: &Protection) -> ProtectionStyle {
    ProtectionStyle {
        locked: prot.locked.unwrap_or(true), // Excel default: locked=true
        hidden: prot.hidden.unwrap_or(false),
    }
}

/// Check if two XML `Font` values are equivalent for deduplication purposes.
fn fonts_equal(a: &Font, b: &Font) -> bool {
    a.b.is_some() == b.b.is_some()
        && a.i.is_some() == b.i.is_some()
        && a.strike.is_some() == b.strike.is_some()
        && a.u.is_some() == b.u.is_some()
        && a.sz == b.sz
        && a.color == b.color
        && a.name == b.name
}

/// Check if two XML `Fill` values are equivalent for deduplication purposes.
fn fills_equal(a: &Fill, b: &Fill) -> bool {
    a.pattern_fill == b.pattern_fill && a.gradient_fill == b.gradient_fill
}

/// Check if two XML `Border` values are equivalent for deduplication purposes.
fn borders_equal(a: &Border, b: &Border) -> bool {
    a.left == b.left
        && a.right == b.right
        && a.top == b.top
        && a.bottom == b.bottom
        && a.diagonal == b.diagonal
}

/// Check if two XML `Xf` values are equivalent for deduplication purposes.
fn xfs_equal(a: &Xf, b: &Xf) -> bool {
    a.num_fmt_id == b.num_fmt_id
        && a.font_id == b.font_id
        && a.fill_id == b.fill_id
        && a.border_id == b.border_id
        && a.alignment == b.alignment
        && a.protection == b.protection
}

/// Convert a `FontStyle` to the XML `Font` struct, find or add it in the fonts list.
/// Returns the 0-based font index.
fn add_or_find_font(fonts: &mut Fonts, font: &FontStyle) -> u32 {
    let xml_font = font_style_to_xml(font);

    for (i, existing) in fonts.fonts.iter().enumerate() {
        if fonts_equal(existing, &xml_font) {
            return i as u32;
        }
    }

    let id = fonts.fonts.len() as u32;
    fonts.fonts.push(xml_font);
    fonts.count = Some(fonts.fonts.len() as u32);
    id
}

/// Convert a `FillStyle` to the XML `Fill` struct, find or add it.
/// Returns the 0-based fill index.
fn add_or_find_fill(fills: &mut Fills, fill: &FillStyle) -> u32 {
    let xml_fill = fill_style_to_xml(fill);

    for (i, existing) in fills.fills.iter().enumerate() {
        if fills_equal(existing, &xml_fill) {
            return i as u32;
        }
    }

    let id = fills.fills.len() as u32;
    fills.fills.push(xml_fill);
    fills.count = Some(fills.fills.len() as u32);
    id
}

/// Convert a `BorderStyle` to the XML `Border` struct, find or add it.
/// Returns the 0-based border index.
fn add_or_find_border(borders: &mut Borders, border: &BorderStyle) -> u32 {
    let xml_border = border_style_to_xml(border);

    for (i, existing) in borders.borders.iter().enumerate() {
        if borders_equal(existing, &xml_border) {
            return i as u32;
        }
    }

    let id = borders.borders.len() as u32;
    borders.borders.push(xml_border);
    borders.count = Some(borders.borders.len() as u32);
    id
}

/// Register a custom number format, return its ID (starting from 164).
/// If an identical format code already exists, returns the existing ID.
fn add_or_find_num_fmt(stylesheet: &mut StyleSheet, fmt: &str) -> u32 {
    let num_fmts = stylesheet.num_fmts.get_or_insert_with(|| NumFmts {
        count: Some(0),
        num_fmts: Vec::new(),
    });

    for nf in &num_fmts.num_fmts {
        if nf.format_code == fmt {
            return nf.num_fmt_id;
        }
    }

    let next_id = num_fmts
        .num_fmts
        .iter()
        .map(|nf| nf.num_fmt_id)
        .max()
        .map(|max_id| max_id + 1)
        .unwrap_or(CUSTOM_NUM_FMT_BASE);

    let next_id = next_id.max(CUSTOM_NUM_FMT_BASE);

    num_fmts.num_fmts.push(NumFmt {
        num_fmt_id: next_id,
        format_code: fmt.to_string(),
    });
    num_fmts.count = Some(num_fmts.num_fmts.len() as u32);

    next_id
}

/// Convert a high-level `Style` to XML components and register in the stylesheet.
/// Returns the style ID (index into cellXfs).
pub fn add_style(stylesheet: &mut StyleSheet, style: &Style) -> Result<u32> {
    if stylesheet.cell_xfs.xfs.len() >= MAX_CELL_XFS {
        return Err(Error::CellStylesExceeded { max: MAX_CELL_XFS });
    }

    let font_id = match &style.font {
        Some(font) => add_or_find_font(&mut stylesheet.fonts, font),
        None => 0, // default font
    };

    let fill_id = match &style.fill {
        Some(fill) => add_or_find_fill(&mut stylesheet.fills, fill),
        None => 0, // default fill (none)
    };

    let border_id = match &style.border {
        Some(border) => add_or_find_border(&mut stylesheet.borders, border),
        None => 0, // default border (empty)
    };

    let num_fmt_id = match &style.num_fmt {
        Some(NumFmtStyle::Builtin(id)) => *id,
        Some(NumFmtStyle::Custom(code)) => add_or_find_num_fmt(stylesheet, code),
        None => 0, // General
    };

    let alignment = style.alignment.as_ref().map(alignment_style_to_xml);
    let protection = style.protection.as_ref().map(protection_style_to_xml);

    let xf = Xf {
        num_fmt_id: Some(num_fmt_id),
        font_id: Some(font_id),
        fill_id: Some(fill_id),
        border_id: Some(border_id),
        xf_id: Some(0),
        apply_number_format: if num_fmt_id != 0 { Some(true) } else { None },
        apply_font: if font_id != 0 { Some(true) } else { None },
        apply_fill: if fill_id != 0 { Some(true) } else { None },
        apply_border: if border_id != 0 { Some(true) } else { None },
        apply_alignment: if alignment.is_some() {
            Some(true)
        } else {
            None
        },
        alignment,
        protection,
    };

    for (i, existing) in stylesheet.cell_xfs.xfs.iter().enumerate() {
        if xfs_equal(existing, &xf) {
            return Ok(i as u32);
        }
    }

    let id = stylesheet.cell_xfs.xfs.len() as u32;
    stylesheet.cell_xfs.xfs.push(xf);
    stylesheet.cell_xfs.count = Some(stylesheet.cell_xfs.xfs.len() as u32);

    Ok(id)
}

/// Get the `Style` from a style ID (reverse lookup from XML components).
pub fn get_style(stylesheet: &StyleSheet, style_id: u32) -> Option<Style> {
    let xf = stylesheet.cell_xfs.xfs.get(style_id as usize)?;

    let font = xf
        .font_id
        .and_then(|id| stylesheet.fonts.fonts.get(id as usize))
        .map(xml_font_to_style);

    let fill = xf
        .fill_id
        .and_then(|id| stylesheet.fills.fills.get(id as usize))
        .and_then(xml_fill_to_style);

    let border = xf
        .border_id
        .and_then(|id| stylesheet.borders.borders.get(id as usize))
        .map(xml_border_to_style);

    let alignment = xf.alignment.as_ref().map(xml_alignment_to_style);

    let num_fmt = xf.num_fmt_id.and_then(|id| {
        if id == 0 {
            return None;
        }
        // Check if it's a built-in format (0-163).
        if id < CUSTOM_NUM_FMT_BASE {
            Some(NumFmtStyle::Builtin(id))
        } else {
            // Look up custom format code.
            stylesheet
                .num_fmts
                .as_ref()
                .and_then(|nfs| nfs.num_fmts.iter().find(|nf| nf.num_fmt_id == id))
                .map(|nf| NumFmtStyle::Custom(nf.format_code.clone()))
        }
    });

    let protection = xf.protection.as_ref().map(xml_protection_to_style);

    Some(Style {
        font,
        fill,
        border,
        alignment,
        num_fmt,
        protection,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Helper to create a fresh default stylesheet for tests.
    fn default_stylesheet() -> StyleSheet {
        StyleSheet::default()
    }

    #[test]
    fn test_add_bold_font_style() {
        let mut ss = default_stylesheet();
        let style = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        // The default stylesheet has one Xf at index 0, so the new one should be at 1.
        assert_eq!(id, 1);
        // Font list should have grown.
        assert_eq!(ss.fonts.fonts.len(), 2);
        assert!(ss.fonts.fonts[1].b.is_some());
    }

    #[test]
    fn test_add_same_style_twice_deduplication() {
        let mut ss = default_stylesheet();
        let style = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };

        let id1 = add_style(&mut ss, &style).unwrap();
        let id2 = add_style(&mut ss, &style).unwrap();
        assert_eq!(id1, id2, "same style should return the same ID");
        // Only 2 fonts (default + the bold one).
        assert_eq!(ss.fonts.fonts.len(), 2);
        // Only 2 Xfs (default + the bold one).
        assert_eq!(ss.cell_xfs.xfs.len(), 2);
    }

    #[test]
    fn test_add_different_styles_different_ids() {
        let mut ss = default_stylesheet();

        let bold_style = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };
        let italic_style = Style {
            font: Some(FontStyle {
                italic: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };

        let id1 = add_style(&mut ss, &bold_style).unwrap();
        let id2 = add_style(&mut ss, &italic_style).unwrap();
        assert_ne!(id1, id2);
    }

    #[test]
    fn test_font_italic() {
        let mut ss = default_stylesheet();
        let style = Style {
            font: Some(FontStyle {
                italic: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        assert!(id > 0);
        let font_id = ss.cell_xfs.xfs[id as usize].font_id.unwrap();
        assert!(ss.fonts.fonts[font_id as usize].i.is_some());
    }

    #[test]
    fn test_font_underline() {
        let mut ss = default_stylesheet();
        let style = Style {
            font: Some(FontStyle {
                underline: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        assert!(id > 0);
        let font_id = ss.cell_xfs.xfs[id as usize].font_id.unwrap();
        assert!(ss.fonts.fonts[font_id as usize].u.is_some());
    }

    #[test]
    fn test_font_strikethrough() {
        let mut ss = default_stylesheet();
        let style = Style {
            font: Some(FontStyle {
                strikethrough: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let font_id = ss.cell_xfs.xfs[id as usize].font_id.unwrap();
        assert!(ss.fonts.fonts[font_id as usize].strike.is_some());
    }

    #[test]
    fn test_font_custom_name_and_size() {
        let mut ss = default_stylesheet();
        let style = Style {
            font: Some(FontStyle {
                name: Some("Arial".to_string()),
                size: Some(14.0),
                ..FontStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let font_id = ss.cell_xfs.xfs[id as usize].font_id.unwrap();
        let xml_font = &ss.fonts.fonts[font_id as usize];
        assert_eq!(xml_font.name.as_ref().unwrap().val, "Arial");
        assert_eq!(xml_font.sz.as_ref().unwrap().val, 14.0);
    }

    #[test]
    fn test_font_with_rgb_color() {
        let mut ss = default_stylesheet();
        let style = Style {
            font: Some(FontStyle {
                color: Some(StyleColor::Rgb("FFFF0000".to_string())),
                ..FontStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let font_id = ss.cell_xfs.xfs[id as usize].font_id.unwrap();
        let xml_font = &ss.fonts.fonts[font_id as usize];
        assert_eq!(
            xml_font.color.as_ref().unwrap().rgb,
            Some("FFFF0000".to_string())
        );
    }

    #[test]
    fn test_fill_solid_color() {
        let mut ss = default_stylesheet();
        let style = Style {
            fill: Some(FillStyle {
                pattern: PatternType::Solid,
                fg_color: Some(StyleColor::Rgb("FFFFFF00".to_string())),
                bg_color: None,
                gradient: None,
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let fill_id = ss.cell_xfs.xfs[id as usize].fill_id.unwrap();
        let xml_fill = &ss.fills.fills[fill_id as usize];
        let pf = xml_fill.pattern_fill.as_ref().unwrap();
        assert_eq!(pf.pattern_type, Some("solid".to_string()));
        assert_eq!(
            pf.fg_color.as_ref().unwrap().rgb,
            Some("FFFFFF00".to_string())
        );
    }

    #[test]
    fn test_fill_pattern() {
        let mut ss = default_stylesheet();
        let style = Style {
            fill: Some(FillStyle {
                pattern: PatternType::LightGray,
                fg_color: None,
                bg_color: None,
                gradient: None,
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let fill_id = ss.cell_xfs.xfs[id as usize].fill_id.unwrap();
        let xml_fill = &ss.fills.fills[fill_id as usize];
        let pf = xml_fill.pattern_fill.as_ref().unwrap();
        assert_eq!(pf.pattern_type, Some("lightGray".to_string()));
    }

    #[test]
    fn test_fill_deduplication() {
        let mut ss = default_stylesheet();
        let style = Style {
            fill: Some(FillStyle {
                pattern: PatternType::Solid,
                fg_color: Some(StyleColor::Rgb("FFFF0000".to_string())),
                bg_color: None,
                gradient: None,
            }),
            ..Style::default()
        };

        let id1 = add_style(&mut ss, &style).unwrap();
        let id2 = add_style(&mut ss, &style).unwrap();
        assert_eq!(id1, id2);
        // Default has 2 fills (none + gray125), we added 1 more.
        assert_eq!(ss.fills.fills.len(), 3);
    }

    #[test]
    fn test_border_thin_all_sides() {
        let mut ss = default_stylesheet();
        let style = Style {
            border: Some(BorderStyle {
                left: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                right: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                top: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                bottom: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                diagonal: None,
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let border_id = ss.cell_xfs.xfs[id as usize].border_id.unwrap();
        let xml_border = &ss.borders.borders[border_id as usize];
        assert_eq!(
            xml_border.left.as_ref().unwrap().style,
            Some("thin".to_string())
        );
        assert_eq!(
            xml_border.right.as_ref().unwrap().style,
            Some("thin".to_string())
        );
        assert_eq!(
            xml_border.top.as_ref().unwrap().style,
            Some("thin".to_string())
        );
        assert_eq!(
            xml_border.bottom.as_ref().unwrap().style,
            Some("thin".to_string())
        );
    }

    #[test]
    fn test_border_medium() {
        let mut ss = default_stylesheet();
        let style = Style {
            border: Some(BorderStyle {
                left: Some(BorderSideStyle {
                    style: BorderLineStyle::Medium,
                    color: Some(StyleColor::Rgb("FF000000".to_string())),
                }),
                right: None,
                top: None,
                bottom: None,
                diagonal: None,
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let border_id = ss.cell_xfs.xfs[id as usize].border_id.unwrap();
        let xml_border = &ss.borders.borders[border_id as usize];
        let left = xml_border.left.as_ref().unwrap();
        assert_eq!(left.style, Some("medium".to_string()));
        assert_eq!(
            left.color.as_ref().unwrap().rgb,
            Some("FF000000".to_string())
        );
    }

    #[test]
    fn test_border_thick() {
        let mut ss = default_stylesheet();
        let style = Style {
            border: Some(BorderStyle {
                bottom: Some(BorderSideStyle {
                    style: BorderLineStyle::Thick,
                    color: None,
                }),
                ..BorderStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let border_id = ss.cell_xfs.xfs[id as usize].border_id.unwrap();
        let xml_border = &ss.borders.borders[border_id as usize];
        assert_eq!(
            xml_border.bottom.as_ref().unwrap().style,
            Some("thick".to_string())
        );
    }

    #[test]
    fn test_num_fmt_builtin() {
        let mut ss = default_stylesheet();
        let style = Style {
            num_fmt: Some(NumFmtStyle::Builtin(builtin_num_fmts::PERCENT)),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let xf = &ss.cell_xfs.xfs[id as usize];
        assert_eq!(xf.num_fmt_id, Some(builtin_num_fmts::PERCENT));
        assert_eq!(xf.apply_number_format, Some(true));
    }

    #[test]
    fn test_num_fmt_custom() {
        let mut ss = default_stylesheet();
        let style = Style {
            num_fmt: Some(NumFmtStyle::Custom("#,##0.000".to_string())),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let xf = &ss.cell_xfs.xfs[id as usize];
        let fmt_id = xf.num_fmt_id.unwrap();
        assert!(fmt_id >= CUSTOM_NUM_FMT_BASE);

        // Verify the format code was stored.
        let num_fmts = ss.num_fmts.as_ref().unwrap();
        let nf = num_fmts
            .num_fmts
            .iter()
            .find(|nf| nf.num_fmt_id == fmt_id)
            .unwrap();
        assert_eq!(nf.format_code, "#,##0.000");
    }

    #[test]
    fn test_num_fmt_custom_deduplication() {
        let mut ss = default_stylesheet();
        let style = Style {
            num_fmt: Some(NumFmtStyle::Custom("0.0%".to_string())),
            ..Style::default()
        };

        let id1 = add_style(&mut ss, &style).unwrap();
        let id2 = add_style(&mut ss, &style).unwrap();
        assert_eq!(id1, id2);

        // Only one custom format should exist.
        let num_fmts = ss.num_fmts.as_ref().unwrap();
        assert_eq!(num_fmts.num_fmts.len(), 1);
    }

    #[test]
    fn test_alignment_horizontal_center() {
        let mut ss = default_stylesheet();
        let style = Style {
            alignment: Some(AlignmentStyle {
                horizontal: Some(HorizontalAlign::Center),
                ..AlignmentStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let xf = &ss.cell_xfs.xfs[id as usize];
        assert_eq!(xf.apply_alignment, Some(true));
        let align = xf.alignment.as_ref().unwrap();
        assert_eq!(align.horizontal, Some("center".to_string()));
    }

    #[test]
    fn test_alignment_vertical_top() {
        let mut ss = default_stylesheet();
        let style = Style {
            alignment: Some(AlignmentStyle {
                vertical: Some(VerticalAlign::Top),
                ..AlignmentStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let xf = &ss.cell_xfs.xfs[id as usize];
        let align = xf.alignment.as_ref().unwrap();
        assert_eq!(align.vertical, Some("top".to_string()));
    }

    #[test]
    fn test_alignment_wrap_text() {
        let mut ss = default_stylesheet();
        let style = Style {
            alignment: Some(AlignmentStyle {
                wrap_text: true,
                ..AlignmentStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let xf = &ss.cell_xfs.xfs[id as usize];
        let align = xf.alignment.as_ref().unwrap();
        assert_eq!(align.wrap_text, Some(true));
    }

    #[test]
    fn test_alignment_text_rotation() {
        let mut ss = default_stylesheet();
        let style = Style {
            alignment: Some(AlignmentStyle {
                text_rotation: Some(90),
                ..AlignmentStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let xf = &ss.cell_xfs.xfs[id as usize];
        let align = xf.alignment.as_ref().unwrap();
        assert_eq!(align.text_rotation, Some(90));
    }

    #[test]
    fn test_protection_locked() {
        let mut ss = default_stylesheet();
        let style = Style {
            protection: Some(ProtectionStyle {
                locked: true,
                hidden: false,
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let xf = &ss.cell_xfs.xfs[id as usize];
        let prot = xf.protection.as_ref().unwrap();
        assert_eq!(prot.locked, Some(true));
        assert_eq!(prot.hidden, Some(false));
    }

    #[test]
    fn test_protection_hidden() {
        let mut ss = default_stylesheet();
        let style = Style {
            protection: Some(ProtectionStyle {
                locked: false,
                hidden: true,
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &style).unwrap();
        let xf = &ss.cell_xfs.xfs[id as usize];
        let prot = xf.protection.as_ref().unwrap();
        assert_eq!(prot.locked, Some(false));
        assert_eq!(prot.hidden, Some(true));
    }

    #[test]
    fn test_combined_style_all_components() {
        let mut ss = default_stylesheet();
        let style = Style {
            font: Some(FontStyle {
                name: Some("Arial".to_string()),
                size: Some(12.0),
                bold: true,
                italic: false,
                underline: false,
                strikethrough: false,
                color: Some(StyleColor::Rgb("FF0000FF".to_string())),
            }),
            fill: Some(FillStyle {
                pattern: PatternType::Solid,
                fg_color: Some(StyleColor::Rgb("FFFFFF00".to_string())),
                bg_color: None,
                gradient: None,
            }),
            border: Some(BorderStyle {
                left: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: Some(StyleColor::Rgb("FF000000".to_string())),
                }),
                right: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: Some(StyleColor::Rgb("FF000000".to_string())),
                }),
                top: Some(BorderSideStyle {
                    style: BorderLineStyle::Medium,
                    color: None,
                }),
                bottom: Some(BorderSideStyle {
                    style: BorderLineStyle::Medium,
                    color: None,
                }),
                diagonal: None,
            }),
            alignment: Some(AlignmentStyle {
                horizontal: Some(HorizontalAlign::Center),
                vertical: Some(VerticalAlign::Center),
                wrap_text: true,
                text_rotation: None,
                indent: None,
                shrink_to_fit: false,
            }),
            num_fmt: Some(NumFmtStyle::Custom("#,##0.00".to_string())),
            protection: Some(ProtectionStyle {
                locked: true,
                hidden: false,
            }),
        };

        let id = add_style(&mut ss, &style).unwrap();
        assert!(id > 0);

        let xf = &ss.cell_xfs.xfs[id as usize];
        assert!(xf.font_id.unwrap() > 0);
        assert!(xf.fill_id.unwrap() > 0);
        assert!(xf.border_id.unwrap() > 0);
        assert!(xf.num_fmt_id.unwrap() >= CUSTOM_NUM_FMT_BASE);
        assert!(xf.alignment.is_some());
        assert!(xf.protection.is_some());
        assert_eq!(xf.apply_font, Some(true));
        assert_eq!(xf.apply_fill, Some(true));
        assert_eq!(xf.apply_border, Some(true));
        assert_eq!(xf.apply_number_format, Some(true));
        assert_eq!(xf.apply_alignment, Some(true));
    }

    #[test]
    fn test_get_style_default() {
        let ss = default_stylesheet();
        let style = get_style(&ss, 0);
        assert!(style.is_some());
        let style = style.unwrap();
        // Default style should have the default font.
        assert!(style.font.is_some());
    }

    #[test]
    fn test_get_style_invalid_id() {
        let ss = default_stylesheet();
        let style = get_style(&ss, 999);
        assert!(style.is_none());
    }

    #[test]
    fn test_get_style_roundtrip_bold() {
        let mut ss = default_stylesheet();
        let original = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &original).unwrap();
        let retrieved = get_style(&ss, id).unwrap();
        assert!(retrieved.font.is_some());
        assert!(retrieved.font.as_ref().unwrap().bold);
    }

    #[test]
    fn test_get_style_roundtrip_fill() {
        let mut ss = default_stylesheet();
        let original = Style {
            fill: Some(FillStyle {
                pattern: PatternType::Solid,
                fg_color: Some(StyleColor::Rgb("FFFF0000".to_string())),
                bg_color: None,
                gradient: None,
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &original).unwrap();
        let retrieved = get_style(&ss, id).unwrap();
        assert!(retrieved.fill.is_some());
        let fill = retrieved.fill.unwrap();
        assert_eq!(fill.pattern, PatternType::Solid);
        assert_eq!(fill.fg_color, Some(StyleColor::Rgb("FFFF0000".to_string())));
    }

    #[test]
    fn test_get_style_roundtrip_alignment() {
        let mut ss = default_stylesheet();
        let original = Style {
            alignment: Some(AlignmentStyle {
                horizontal: Some(HorizontalAlign::Right),
                vertical: Some(VerticalAlign::Bottom),
                wrap_text: true,
                text_rotation: Some(45),
                indent: Some(2),
                shrink_to_fit: false,
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &original).unwrap();
        let retrieved = get_style(&ss, id).unwrap();
        assert!(retrieved.alignment.is_some());
        let align = retrieved.alignment.unwrap();
        assert_eq!(align.horizontal, Some(HorizontalAlign::Right));
        assert_eq!(align.vertical, Some(VerticalAlign::Bottom));
        assert!(align.wrap_text);
        assert_eq!(align.text_rotation, Some(45));
        assert_eq!(align.indent, Some(2));
    }

    #[test]
    fn test_get_style_roundtrip_protection() {
        let mut ss = default_stylesheet();
        let original = Style {
            protection: Some(ProtectionStyle {
                locked: false,
                hidden: true,
            }),
            ..Style::default()
        };

        let id = add_style(&mut ss, &original).unwrap();
        let retrieved = get_style(&ss, id).unwrap();
        assert!(retrieved.protection.is_some());
        let prot = retrieved.protection.unwrap();
        assert!(!prot.locked);
        assert!(prot.hidden);
    }

    #[test]
    fn test_get_style_roundtrip_num_fmt_builtin() {
        let mut ss = default_stylesheet();
        let original = Style {
            num_fmt: Some(NumFmtStyle::Builtin(builtin_num_fmts::DATE_MDY)),
            ..Style::default()
        };

        let id = add_style(&mut ss, &original).unwrap();
        let retrieved = get_style(&ss, id).unwrap();
        assert!(retrieved.num_fmt.is_some());
        match retrieved.num_fmt.unwrap() {
            NumFmtStyle::Builtin(fid) => assert_eq!(fid, builtin_num_fmts::DATE_MDY),
            _ => panic!("expected Builtin num fmt"),
        }
    }

    #[test]
    fn test_get_style_roundtrip_num_fmt_custom() {
        let mut ss = default_stylesheet();
        let original = Style {
            num_fmt: Some(NumFmtStyle::Custom("yyyy-mm-dd".to_string())),
            ..Style::default()
        };

        let id = add_style(&mut ss, &original).unwrap();
        let retrieved = get_style(&ss, id).unwrap();
        assert!(retrieved.num_fmt.is_some());
        match retrieved.num_fmt.unwrap() {
            NumFmtStyle::Custom(code) => assert_eq!(code, "yyyy-mm-dd"),
            _ => panic!("expected Custom num fmt"),
        }
    }

    #[test]
    fn test_builtin_num_fmt_constants() {
        assert_eq!(builtin_num_fmts::GENERAL, 0);
        assert_eq!(builtin_num_fmts::INTEGER, 1);
        assert_eq!(builtin_num_fmts::DECIMAL_2, 2);
        assert_eq!(builtin_num_fmts::THOUSANDS, 3);
        assert_eq!(builtin_num_fmts::THOUSANDS_DECIMAL, 4);
        assert_eq!(builtin_num_fmts::PERCENT, 9);
        assert_eq!(builtin_num_fmts::PERCENT_DECIMAL, 10);
        assert_eq!(builtin_num_fmts::SCIENTIFIC, 11);
        assert_eq!(builtin_num_fmts::DATE_MDY, 14);
        assert_eq!(builtin_num_fmts::DATE_DMY, 15);
        assert_eq!(builtin_num_fmts::DATE_DM, 16);
        assert_eq!(builtin_num_fmts::DATE_MY, 17);
        assert_eq!(builtin_num_fmts::TIME_HM_AP, 18);
        assert_eq!(builtin_num_fmts::TIME_HMS_AP, 19);
        assert_eq!(builtin_num_fmts::TIME_HM, 20);
        assert_eq!(builtin_num_fmts::TIME_HMS, 21);
        assert_eq!(builtin_num_fmts::DATETIME, 22);
        assert_eq!(builtin_num_fmts::TEXT, 49);
    }

    #[test]
    fn test_pattern_type_roundtrip() {
        let types = [
            PatternType::None,
            PatternType::Solid,
            PatternType::Gray125,
            PatternType::DarkGray,
            PatternType::MediumGray,
            PatternType::LightGray,
        ];
        for pt in &types {
            let s = pt.as_str();
            let back = PatternType::from_str(s);
            assert_eq!(*pt, back);
        }
    }

    #[test]
    fn test_border_line_style_roundtrip() {
        let styles = [
            BorderLineStyle::Thin,
            BorderLineStyle::Medium,
            BorderLineStyle::Thick,
            BorderLineStyle::Dashed,
            BorderLineStyle::Dotted,
            BorderLineStyle::Double,
            BorderLineStyle::Hair,
            BorderLineStyle::MediumDashed,
            BorderLineStyle::DashDot,
            BorderLineStyle::MediumDashDot,
            BorderLineStyle::DashDotDot,
            BorderLineStyle::MediumDashDotDot,
            BorderLineStyle::SlantDashDot,
        ];
        for bls in &styles {
            let s = bls.as_str();
            let back = BorderLineStyle::from_str(s).unwrap();
            assert_eq!(*bls, back);
        }
    }

    #[test]
    fn test_horizontal_align_roundtrip() {
        let aligns = [
            HorizontalAlign::General,
            HorizontalAlign::Left,
            HorizontalAlign::Center,
            HorizontalAlign::Right,
            HorizontalAlign::Fill,
            HorizontalAlign::Justify,
            HorizontalAlign::CenterContinuous,
            HorizontalAlign::Distributed,
        ];
        for ha in &aligns {
            let s = ha.as_str();
            let back = HorizontalAlign::from_str(s).unwrap();
            assert_eq!(*ha, back);
        }
    }

    #[test]
    fn test_vertical_align_roundtrip() {
        let aligns = [
            VerticalAlign::Top,
            VerticalAlign::Center,
            VerticalAlign::Bottom,
            VerticalAlign::Justify,
            VerticalAlign::Distributed,
        ];
        for va in &aligns {
            let s = va.as_str();
            let back = VerticalAlign::from_str(s).unwrap();
            assert_eq!(*va, back);
        }
    }

    #[test]
    fn test_style_color_rgb_roundtrip() {
        let color = StyleColor::Rgb("FF00FF00".to_string());
        let xml = style_color_to_xml(&color);
        let back = xml_color_to_style(&xml).unwrap();
        assert_eq!(color, back);
    }

    #[test]
    fn test_style_color_theme_roundtrip() {
        let color = StyleColor::Theme(4);
        let xml = style_color_to_xml(&color);
        let back = xml_color_to_style(&xml).unwrap();
        assert_eq!(color, back);
    }

    #[test]
    fn test_style_color_indexed_roundtrip() {
        let color = StyleColor::Indexed(10);
        let xml = style_color_to_xml(&color);
        let back = xml_color_to_style(&xml).unwrap();
        assert_eq!(color, back);
    }

    #[test]
    fn test_font_deduplication() {
        let mut ss = default_stylesheet();
        let font = FontStyle {
            name: Some("Courier".to_string()),
            size: Some(10.0),
            bold: true,
            ..FontStyle::default()
        };

        let id1 = add_or_find_font(&mut ss.fonts, &font);
        let id2 = add_or_find_font(&mut ss.fonts, &font);
        assert_eq!(id1, id2);
        // Default has 1 font, we added 1.
        assert_eq!(ss.fonts.fonts.len(), 2);
    }

    #[test]
    fn test_multiple_custom_num_fmts() {
        let mut ss = default_stylesheet();
        let id1 = add_or_find_num_fmt(&mut ss, "0.0%");
        let id2 = add_or_find_num_fmt(&mut ss, "#,##0");
        assert_eq!(id1, 164);
        assert_eq!(id2, 165);

        // Same format returns same id.
        let id3 = add_or_find_num_fmt(&mut ss, "0.0%");
        assert_eq!(id3, 164);
    }

    #[test]
    fn test_xf_count_maintained() {
        let mut ss = default_stylesheet();
        assert_eq!(ss.cell_xfs.count, Some(1));

        let style = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };
        add_style(&mut ss, &style).unwrap();
        assert_eq!(ss.cell_xfs.count, Some(2));
    }

    // -- StyleBuilder tests --

    #[test]
    fn test_style_builder_empty() {
        let style = StyleBuilder::new().build();
        assert!(style.font.is_none());
        assert!(style.fill.is_none());
        assert!(style.border.is_none());
        assert!(style.alignment.is_none());
        assert!(style.num_fmt.is_none());
        assert!(style.protection.is_none());
    }

    #[test]
    fn test_style_builder_default_equivalent() {
        let style = StyleBuilder::default().build();
        assert!(style.font.is_none());
        assert!(style.fill.is_none());
    }

    #[test]
    fn test_style_builder_font() {
        let style = StyleBuilder::new()
            .bold(true)
            .italic(true)
            .font_size(14.0)
            .font_name("Arial")
            .font_color_rgb("FF0000FF")
            .build();
        let font = style.font.unwrap();
        assert!(font.bold);
        assert!(font.italic);
        assert_eq!(font.size, Some(14.0));
        assert_eq!(font.name, Some("Arial".to_string()));
        assert_eq!(font.color, Some(StyleColor::Rgb("FF0000FF".to_string())));
    }

    #[test]
    fn test_style_builder_font_underline_strikethrough() {
        let style = StyleBuilder::new()
            .underline(true)
            .strikethrough(true)
            .build();
        let font = style.font.unwrap();
        assert!(font.underline);
        assert!(font.strikethrough);
    }

    #[test]
    fn test_style_builder_font_color_typed() {
        let style = StyleBuilder::new().font_color(StyleColor::Theme(4)).build();
        let font = style.font.unwrap();
        assert_eq!(font.color, Some(StyleColor::Theme(4)));
    }

    #[test]
    fn test_style_builder_solid_fill() {
        let style = StyleBuilder::new().solid_fill("FFFF0000").build();
        let fill = style.fill.unwrap();
        assert_eq!(fill.pattern, PatternType::Solid);
        assert_eq!(fill.fg_color, Some(StyleColor::Rgb("FFFF0000".to_string())));
    }

    #[test]
    fn test_style_builder_fill_pattern_and_colors() {
        let style = StyleBuilder::new()
            .fill_pattern(PatternType::Gray125)
            .fill_fg_color_rgb("FFAABBCC")
            .fill_bg_color(StyleColor::Indexed(64))
            .build();
        let fill = style.fill.unwrap();
        assert_eq!(fill.pattern, PatternType::Gray125);
        assert_eq!(fill.fg_color, Some(StyleColor::Rgb("FFAABBCC".to_string())));
        assert_eq!(fill.bg_color, Some(StyleColor::Indexed(64)));
    }

    #[test]
    fn test_style_builder_border_individual_sides() {
        let style = StyleBuilder::new()
            .border_left(
                BorderLineStyle::Thin,
                StyleColor::Rgb("FF000000".to_string()),
            )
            .border_right(
                BorderLineStyle::Medium,
                StyleColor::Rgb("FF111111".to_string()),
            )
            .border_top(
                BorderLineStyle::Thick,
                StyleColor::Rgb("FF222222".to_string()),
            )
            .border_bottom(
                BorderLineStyle::Dashed,
                StyleColor::Rgb("FF333333".to_string()),
            )
            .build();
        let border = style.border.unwrap();

        let left = border.left.unwrap();
        assert_eq!(left.style, BorderLineStyle::Thin);
        assert_eq!(left.color, Some(StyleColor::Rgb("FF000000".to_string())));

        let right = border.right.unwrap();
        assert_eq!(right.style, BorderLineStyle::Medium);

        let top = border.top.unwrap();
        assert_eq!(top.style, BorderLineStyle::Thick);

        let bottom = border.bottom.unwrap();
        assert_eq!(bottom.style, BorderLineStyle::Dashed);
    }

    #[test]
    fn test_style_builder_border_all() {
        let style = StyleBuilder::new()
            .border_all(
                BorderLineStyle::Thin,
                StyleColor::Rgb("FF000000".to_string()),
            )
            .build();
        let border = style.border.unwrap();
        assert!(border.left.is_some());
        assert!(border.right.is_some());
        assert!(border.top.is_some());
        assert!(border.bottom.is_some());
        // diagonal should not be set by border_all
        assert!(border.diagonal.is_none());

        let left = border.left.unwrap();
        assert_eq!(left.style, BorderLineStyle::Thin);
        assert_eq!(left.color, Some(StyleColor::Rgb("FF000000".to_string())));
    }

    #[test]
    fn test_style_builder_alignment() {
        let style = StyleBuilder::new()
            .horizontal_align(HorizontalAlign::Center)
            .vertical_align(VerticalAlign::Center)
            .wrap_text(true)
            .text_rotation(45)
            .indent(2)
            .shrink_to_fit(true)
            .build();
        let align = style.alignment.unwrap();
        assert_eq!(align.horizontal, Some(HorizontalAlign::Center));
        assert_eq!(align.vertical, Some(VerticalAlign::Center));
        assert!(align.wrap_text);
        assert_eq!(align.text_rotation, Some(45));
        assert_eq!(align.indent, Some(2));
        assert!(align.shrink_to_fit);
    }

    #[test]
    fn test_style_builder_num_format_builtin() {
        let style = StyleBuilder::new().num_format_builtin(2).build();
        match style.num_fmt.unwrap() {
            NumFmtStyle::Builtin(id) => assert_eq!(id, 2),
            _ => panic!("expected builtin format"),
        }
    }

    #[test]
    fn test_style_builder_num_format_custom() {
        let style = StyleBuilder::new().num_format_custom("#,##0.00").build();
        match style.num_fmt.unwrap() {
            NumFmtStyle::Custom(fmt) => assert_eq!(fmt, "#,##0.00"),
            _ => panic!("expected custom format"),
        }
    }

    #[test]
    fn test_style_builder_protection() {
        let style = StyleBuilder::new().locked(true).hidden(true).build();
        let prot = style.protection.unwrap();
        assert!(prot.locked);
        assert!(prot.hidden);
    }

    #[test]
    fn test_style_builder_protection_unlock() {
        let style = StyleBuilder::new().locked(false).hidden(false).build();
        let prot = style.protection.unwrap();
        assert!(!prot.locked);
        assert!(!prot.hidden);
    }

    #[test]
    fn test_style_builder_full_style() {
        let style = StyleBuilder::new()
            .bold(true)
            .font_size(12.0)
            .solid_fill("FFFFFF00")
            .border_all(
                BorderLineStyle::Thin,
                StyleColor::Rgb("FF000000".to_string()),
            )
            .horizontal_align(HorizontalAlign::Center)
            .num_format_builtin(2)
            .locked(true)
            .build();
        assert!(style.font.is_some());
        assert!(style.fill.is_some());
        assert!(style.border.is_some());
        assert!(style.alignment.is_some());
        assert!(style.num_fmt.is_some());
        assert!(style.protection.is_some());

        // Verify specific values survived the chaining
        assert!(style.font.as_ref().unwrap().bold);
        assert_eq!(style.font.as_ref().unwrap().size, Some(12.0));
        assert_eq!(style.fill.as_ref().unwrap().pattern, PatternType::Solid);
    }

    #[test]
    fn test_style_builder_integrates_with_add_style() {
        let mut ss = default_stylesheet();
        let style = StyleBuilder::new()
            .bold(true)
            .font_size(11.0)
            .solid_fill("FFFF0000")
            .horizontal_align(HorizontalAlign::Center)
            .build();

        let id = add_style(&mut ss, &style).unwrap();
        assert!(id > 0);

        // Round-trip: get the style back and verify
        let retrieved = get_style(&ss, id).unwrap();
        assert!(retrieved.font.as_ref().unwrap().bold);
        assert_eq!(retrieved.fill.as_ref().unwrap().pattern, PatternType::Solid);
        assert_eq!(
            retrieved.alignment.as_ref().unwrap().horizontal,
            Some(HorizontalAlign::Center)
        );
    }
}
