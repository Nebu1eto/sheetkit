//! Styles XML schema structures.
//!
//! Represents `xl/styles.xml` in the OOXML package.
//! Minimal implementation for Phase 1 with Excel-compatible default styles.

use serde::{Deserialize, Serialize};

use crate::namespaces;

/// Stylesheet root element (`xl/styles.xml`).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "styleSheet")]
pub struct StyleSheet {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "numFmts", skip_serializing_if = "Option::is_none")]
    pub num_fmts: Option<NumFmts>,

    #[serde(rename = "fonts")]
    pub fonts: Fonts,

    #[serde(rename = "fills")]
    pub fills: Fills,

    #[serde(rename = "borders")]
    pub borders: Borders,

    #[serde(rename = "cellStyleXfs", skip_serializing_if = "Option::is_none")]
    pub cell_style_xfs: Option<CellStyleXfs>,

    #[serde(rename = "cellXfs")]
    pub cell_xfs: CellXfs,

    #[serde(rename = "cellStyles", skip_serializing_if = "Option::is_none")]
    pub cell_styles: Option<CellStyles>,

    #[serde(rename = "dxfs", skip_serializing_if = "Option::is_none")]
    pub dxfs: Option<Dxfs>,

    #[serde(rename = "tableStyles", skip_serializing_if = "Option::is_none")]
    pub table_styles: Option<TableStyles>,
}

/// Number formats container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct NumFmts {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "numFmt", default)]
    pub num_fmts: Vec<NumFmt>,
}

/// Individual number format.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct NumFmt {
    #[serde(rename = "@numFmtId")]
    pub num_fmt_id: u32,

    #[serde(rename = "@formatCode")]
    pub format_code: String,
}

/// Fonts container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Fonts {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "font", default)]
    pub fonts: Vec<Font>,
}

/// Individual font definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Font {
    #[serde(rename = "b", skip_serializing_if = "Option::is_none")]
    pub b: Option<BoolVal>,

    #[serde(rename = "i", skip_serializing_if = "Option::is_none")]
    pub i: Option<BoolVal>,

    #[serde(rename = "strike", skip_serializing_if = "Option::is_none")]
    pub strike: Option<BoolVal>,

    #[serde(rename = "u", skip_serializing_if = "Option::is_none")]
    pub u: Option<Underline>,

    #[serde(rename = "sz", skip_serializing_if = "Option::is_none")]
    pub sz: Option<FontSize>,

    #[serde(rename = "color", skip_serializing_if = "Option::is_none")]
    pub color: Option<Color>,

    #[serde(rename = "name", skip_serializing_if = "Option::is_none")]
    pub name: Option<FontName>,

    #[serde(rename = "family", skip_serializing_if = "Option::is_none")]
    pub family: Option<FontFamily>,

    #[serde(rename = "scheme", skip_serializing_if = "Option::is_none")]
    pub scheme: Option<FontScheme>,
}

/// Fills container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Fills {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "fill", default)]
    pub fills: Vec<Fill>,
}

/// Individual fill definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Fill {
    #[serde(rename = "patternFill", skip_serializing_if = "Option::is_none")]
    pub pattern_fill: Option<PatternFill>,

    #[serde(rename = "gradientFill", skip_serializing_if = "Option::is_none")]
    pub gradient_fill: Option<GradientFill>,
}

/// Pattern fill definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PatternFill {
    #[serde(rename = "@patternType", skip_serializing_if = "Option::is_none")]
    pub pattern_type: Option<String>,

    #[serde(rename = "fgColor", skip_serializing_if = "Option::is_none")]
    pub fg_color: Option<Color>,

    #[serde(rename = "bgColor", skip_serializing_if = "Option::is_none")]
    pub bg_color: Option<Color>,
}

/// Gradient fill definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct GradientFill {
    /// Gradient type: "linear" or "path".
    #[serde(rename = "@type", skip_serializing_if = "Option::is_none")]
    pub gradient_type: Option<String>,

    /// Rotation angle in degrees for linear gradients.
    #[serde(rename = "@degree", skip_serializing_if = "Option::is_none")]
    pub degree: Option<f64>,

    /// Left coordinate for path gradients (0.0-1.0).
    #[serde(rename = "@left", skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,

    /// Right coordinate for path gradients (0.0-1.0).
    #[serde(rename = "@right", skip_serializing_if = "Option::is_none")]
    pub right: Option<f64>,

    /// Top coordinate for path gradients (0.0-1.0).
    #[serde(rename = "@top", skip_serializing_if = "Option::is_none")]
    pub top: Option<f64>,

    /// Bottom coordinate for path gradients (0.0-1.0).
    #[serde(rename = "@bottom", skip_serializing_if = "Option::is_none")]
    pub bottom: Option<f64>,

    /// Gradient stops.
    #[serde(rename = "stop", default)]
    pub stops: Vec<GradientStop>,
}

/// A single gradient stop with position and color.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct GradientStop {
    /// Position of this stop (0.0-1.0).
    #[serde(rename = "@position")]
    pub position: f64,

    /// Color at this stop.
    pub color: Color,
}

/// Borders container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Borders {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "border", default)]
    pub borders: Vec<Border>,
}

/// Individual border definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Border {
    #[serde(rename = "@diagonalUp", skip_serializing_if = "Option::is_none")]
    pub diagonal_up: Option<bool>,

    #[serde(rename = "@diagonalDown", skip_serializing_if = "Option::is_none")]
    pub diagonal_down: Option<bool>,

    #[serde(rename = "left", skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderSide>,

    #[serde(rename = "right", skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderSide>,

    #[serde(rename = "top", skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderSide>,

    #[serde(rename = "bottom", skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderSide>,

    #[serde(rename = "diagonal", skip_serializing_if = "Option::is_none")]
    pub diagonal: Option<BorderSide>,
}

/// Border side definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BorderSide {
    #[serde(rename = "@style", skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,

    #[serde(rename = "color", skip_serializing_if = "Option::is_none")]
    pub color: Option<Color>,
}

/// Cell style XFs container (base style formats).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellStyleXfs {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "xf", default)]
    pub xfs: Vec<Xf>,
}

/// Cell XFs container (applied cell formats).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellXfs {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "xf", default)]
    pub xfs: Vec<Xf>,
}

/// Cell format entry.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Xf {
    #[serde(rename = "@numFmtId", skip_serializing_if = "Option::is_none")]
    pub num_fmt_id: Option<u32>,

    #[serde(rename = "@fontId", skip_serializing_if = "Option::is_none")]
    pub font_id: Option<u32>,

    #[serde(rename = "@fillId", skip_serializing_if = "Option::is_none")]
    pub fill_id: Option<u32>,

    #[serde(rename = "@borderId", skip_serializing_if = "Option::is_none")]
    pub border_id: Option<u32>,

    #[serde(rename = "@xfId", skip_serializing_if = "Option::is_none")]
    pub xf_id: Option<u32>,

    #[serde(rename = "@applyNumberFormat", skip_serializing_if = "Option::is_none")]
    pub apply_number_format: Option<bool>,

    #[serde(rename = "@applyFont", skip_serializing_if = "Option::is_none")]
    pub apply_font: Option<bool>,

    #[serde(rename = "@applyFill", skip_serializing_if = "Option::is_none")]
    pub apply_fill: Option<bool>,

    #[serde(rename = "@applyBorder", skip_serializing_if = "Option::is_none")]
    pub apply_border: Option<bool>,

    #[serde(rename = "@applyAlignment", skip_serializing_if = "Option::is_none")]
    pub apply_alignment: Option<bool>,

    #[serde(rename = "alignment", skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Alignment>,

    #[serde(rename = "protection", skip_serializing_if = "Option::is_none")]
    pub protection: Option<Protection>,
}

/// Cell alignment.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Alignment {
    #[serde(rename = "@horizontal", skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<String>,

    #[serde(rename = "@vertical", skip_serializing_if = "Option::is_none")]
    pub vertical: Option<String>,

    #[serde(rename = "@wrapText", skip_serializing_if = "Option::is_none")]
    pub wrap_text: Option<bool>,

    #[serde(rename = "@textRotation", skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<u32>,

    #[serde(rename = "@indent", skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,

    #[serde(rename = "@shrinkToFit", skip_serializing_if = "Option::is_none")]
    pub shrink_to_fit: Option<bool>,
}

/// Cell protection.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Protection {
    #[serde(rename = "@locked", skip_serializing_if = "Option::is_none")]
    pub locked: Option<bool>,

    #[serde(rename = "@hidden", skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
}

/// Cell styles container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellStyles {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "cellStyle", default)]
    pub cell_styles: Vec<CellStyle>,
}

/// Individual named cell style.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellStyle {
    #[serde(rename = "@name")]
    pub name: String,

    #[serde(rename = "@xfId")]
    pub xf_id: u32,

    #[serde(rename = "@builtinId", skip_serializing_if = "Option::is_none")]
    pub builtin_id: Option<u32>,
}

/// Differential formats container (for conditional formatting).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Dxfs {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "dxf", default)]
    pub dxfs: Vec<Dxf>,
}

/// Individual differential format.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Dxf {
    #[serde(rename = "font", skip_serializing_if = "Option::is_none")]
    pub font: Option<Font>,

    #[serde(rename = "numFmt", skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<NumFmt>,

    #[serde(rename = "fill", skip_serializing_if = "Option::is_none")]
    pub fill: Option<Fill>,

    #[serde(rename = "border", skip_serializing_if = "Option::is_none")]
    pub border: Option<Border>,
}

/// Table styles container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TableStyles {
    #[serde(rename = "@count", skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,

    #[serde(rename = "@defaultTableStyle", skip_serializing_if = "Option::is_none")]
    pub default_table_style: Option<String>,

    #[serde(rename = "@defaultPivotStyle", skip_serializing_if = "Option::is_none")]
    pub default_pivot_style: Option<String>,
}

/// Color definition (used across fonts, fills, borders).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Color {
    #[serde(rename = "@auto", skip_serializing_if = "Option::is_none")]
    pub auto: Option<bool>,

    #[serde(rename = "@indexed", skip_serializing_if = "Option::is_none")]
    pub indexed: Option<u32>,

    #[serde(rename = "@rgb", skip_serializing_if = "Option::is_none")]
    pub rgb: Option<String>,

    #[serde(rename = "@theme", skip_serializing_if = "Option::is_none")]
    pub theme: Option<u32>,

    #[serde(rename = "@tint", skip_serializing_if = "Option::is_none")]
    pub tint: Option<f64>,
}

/// Boolean value wrapper (used for bold, italic, etc.).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BoolVal {
    #[serde(rename = "@val", skip_serializing_if = "Option::is_none")]
    pub val: Option<bool>,
}

/// Underline.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Underline {
    #[serde(rename = "@val", skip_serializing_if = "Option::is_none")]
    pub val: Option<String>,
}

/// Font size.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FontSize {
    #[serde(rename = "@val")]
    pub val: f64,
}

/// Font name.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FontName {
    #[serde(rename = "@val")]
    pub val: String,
}

/// Font family.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FontFamily {
    #[serde(rename = "@val")]
    pub val: u32,
}

/// Font scheme (theme-based).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FontScheme {
    #[serde(rename = "@val")]
    pub val: String,
}

impl Default for StyleSheet {
    /// Creates an Excel-compatible minimal default stylesheet.
    ///
    /// This includes:
    /// - 1 default font (Calibri, 11pt)
    /// - 2 required fills (none, gray125)
    /// - 1 empty border
    /// - 1 cellStyleXf
    /// - 1 cellXf (Normal style)
    /// - 1 cellStyle ("Normal")
    fn default() -> Self {
        Self {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            num_fmts: None,
            fonts: Fonts {
                count: Some(1),
                fonts: vec![Font {
                    b: None,
                    i: None,
                    strike: None,
                    u: None,
                    sz: Some(FontSize { val: 11.0 }),
                    color: Some(Color {
                        auto: None,
                        indexed: None,
                        rgb: None,
                        theme: Some(1),
                        tint: None,
                    }),
                    name: Some(FontName {
                        val: "Calibri".to_string(),
                    }),
                    family: Some(FontFamily { val: 2 }),
                    scheme: Some(FontScheme {
                        val: "minor".to_string(),
                    }),
                }],
            },
            fills: Fills {
                count: Some(2),
                fills: vec![
                    Fill {
                        pattern_fill: Some(PatternFill {
                            pattern_type: Some("none".to_string()),
                            fg_color: None,
                            bg_color: None,
                        }),
                        gradient_fill: None,
                    },
                    Fill {
                        pattern_fill: Some(PatternFill {
                            pattern_type: Some("gray125".to_string()),
                            fg_color: None,
                            bg_color: None,
                        }),
                        gradient_fill: None,
                    },
                ],
            },
            borders: Borders {
                count: Some(1),
                borders: vec![Border {
                    diagonal_up: None,
                    diagonal_down: None,
                    left: Some(BorderSide {
                        style: None,
                        color: None,
                    }),
                    right: Some(BorderSide {
                        style: None,
                        color: None,
                    }),
                    top: Some(BorderSide {
                        style: None,
                        color: None,
                    }),
                    bottom: Some(BorderSide {
                        style: None,
                        color: None,
                    }),
                    diagonal: Some(BorderSide {
                        style: None,
                        color: None,
                    }),
                }],
            },
            cell_style_xfs: Some(CellStyleXfs {
                count: Some(1),
                xfs: vec![Xf {
                    num_fmt_id: Some(0),
                    font_id: Some(0),
                    fill_id: Some(0),
                    border_id: Some(0),
                    xf_id: None,
                    apply_number_format: None,
                    apply_font: None,
                    apply_fill: None,
                    apply_border: None,
                    apply_alignment: None,
                    alignment: None,
                    protection: None,
                }],
            }),
            cell_xfs: CellXfs {
                count: Some(1),
                xfs: vec![Xf {
                    num_fmt_id: Some(0),
                    font_id: Some(0),
                    fill_id: Some(0),
                    border_id: Some(0),
                    xf_id: Some(0),
                    apply_number_format: None,
                    apply_font: None,
                    apply_fill: None,
                    apply_border: None,
                    apply_alignment: None,
                    alignment: None,
                    protection: None,
                }],
            },
            cell_styles: Some(CellStyles {
                count: Some(1),
                cell_styles: vec![CellStyle {
                    name: "Normal".to_string(),
                    xf_id: 0,
                    builtin_id: Some(0),
                }],
            }),
            dxfs: None,
            table_styles: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_stylesheet_default() {
        let ss = StyleSheet::default();
        assert_eq!(ss.xmlns, namespaces::SPREADSHEET_ML);

        // 1 default font
        assert_eq!(ss.fonts.fonts.len(), 1);
        assert_eq!(ss.fonts.count, Some(1));
        let font = &ss.fonts.fonts[0];
        assert_eq!(font.sz.as_ref().unwrap().val, 11.0);
        assert_eq!(font.name.as_ref().unwrap().val, "Calibri");
        assert_eq!(font.family.as_ref().unwrap().val, 2);
        assert_eq!(font.scheme.as_ref().unwrap().val, "minor");

        // 2 required fills (none, gray125)
        assert_eq!(ss.fills.fills.len(), 2);
        assert_eq!(ss.fills.count, Some(2));
        assert_eq!(
            ss.fills.fills[0]
                .pattern_fill
                .as_ref()
                .unwrap()
                .pattern_type,
            Some("none".to_string())
        );
        assert_eq!(
            ss.fills.fills[1]
                .pattern_fill
                .as_ref()
                .unwrap()
                .pattern_type,
            Some("gray125".to_string())
        );

        // 1 border
        assert_eq!(ss.borders.borders.len(), 1);
        assert_eq!(ss.borders.count, Some(1));

        // 1 cellStyleXf
        assert!(ss.cell_style_xfs.is_some());
        assert_eq!(ss.cell_style_xfs.as_ref().unwrap().xfs.len(), 1);

        // 1 cellXf
        assert_eq!(ss.cell_xfs.xfs.len(), 1);
        assert_eq!(ss.cell_xfs.count, Some(1));
        let xf = &ss.cell_xfs.xfs[0];
        assert_eq!(xf.num_fmt_id, Some(0));
        assert_eq!(xf.font_id, Some(0));
        assert_eq!(xf.fill_id, Some(0));
        assert_eq!(xf.border_id, Some(0));

        // 1 cellStyle
        assert!(ss.cell_styles.is_some());
        let styles = ss.cell_styles.as_ref().unwrap();
        assert_eq!(styles.cell_styles.len(), 1);
        assert_eq!(styles.cell_styles[0].name, "Normal");
        assert_eq!(styles.cell_styles[0].xf_id, 0);
        assert_eq!(styles.cell_styles[0].builtin_id, Some(0));
    }

    #[test]
    fn test_stylesheet_roundtrip() {
        let ss = StyleSheet::default();
        let xml = quick_xml::se::to_string(&ss).unwrap();
        let parsed: StyleSheet = quick_xml::de::from_str(&xml).unwrap();

        assert_eq!(ss.xmlns, parsed.xmlns);
        assert_eq!(ss.fonts.fonts.len(), parsed.fonts.fonts.len());
        assert_eq!(ss.fills.fills.len(), parsed.fills.fills.len());
        assert_eq!(ss.borders.borders.len(), parsed.borders.borders.len());
        assert_eq!(ss.cell_xfs.xfs.len(), parsed.cell_xfs.xfs.len());
    }

    #[test]
    fn test_stylesheet_serialize_structure() {
        let ss = StyleSheet::default();
        let xml = quick_xml::se::to_string(&ss).unwrap();
        assert!(xml.contains("<styleSheet"));
        assert!(xml.contains("<fonts"));
        assert!(xml.contains("<font>"));
        assert!(xml.contains("<fills"));
        assert!(xml.contains("<fill>"));
        assert!(xml.contains("<borders"));
        assert!(xml.contains("<border>"));
        assert!(xml.contains("<cellXfs"));
        assert!(xml.contains("<xf "));
    }

    #[test]
    fn test_parse_real_excel_styles_minimal() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"#;

        let parsed: StyleSheet = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(parsed.fonts.fonts.len(), 1);
        assert_eq!(parsed.fonts.fonts[0].sz.as_ref().unwrap().val, 11.0);
        assert_eq!(parsed.fonts.fonts[0].name.as_ref().unwrap().val, "Calibri");
        assert_eq!(parsed.fills.fills.len(), 2);
        assert_eq!(parsed.borders.borders.len(), 1);
        assert_eq!(parsed.cell_xfs.xfs.len(), 1);
        assert_eq!(parsed.cell_xfs.xfs[0].num_fmt_id, Some(0));
    }

    #[test]
    fn test_font_with_bold_italic() {
        let font = Font {
            b: Some(BoolVal { val: None }),
            i: Some(BoolVal { val: None }),
            strike: None,
            u: None,
            sz: Some(FontSize { val: 12.0 }),
            color: None,
            name: Some(FontName {
                val: "Arial".to_string(),
            }),
            family: None,
            scheme: None,
        };
        let xml = quick_xml::se::to_string(&font).unwrap();
        assert!(xml.contains("<b"));
        assert!(xml.contains("<i"));
        let parsed: Font = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.b.is_some());
        assert!(parsed.i.is_some());
        assert_eq!(parsed.sz.unwrap().val, 12.0);
    }

    #[test]
    fn test_color_rgb() {
        let color = Color {
            auto: None,
            indexed: None,
            rgb: Some("FF0000FF".to_string()),
            theme: None,
            tint: None,
        };
        let xml = quick_xml::se::to_string(&color).unwrap();
        assert!(xml.contains("rgb=\"FF0000FF\""));
        let parsed: Color = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.rgb, Some("FF0000FF".to_string()));
    }

    #[test]
    fn test_color_theme_with_tint() {
        let color = Color {
            auto: None,
            indexed: None,
            rgb: None,
            theme: Some(4),
            tint: Some(0.399_975_585_192_419_2),
        };
        let xml = quick_xml::se::to_string(&color).unwrap();
        let parsed: Color = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.theme, Some(4));
        assert!(parsed.tint.is_some());
    }

    #[test]
    fn test_border_with_style() {
        let border = Border {
            diagonal_up: None,
            diagonal_down: None,
            left: Some(BorderSide {
                style: Some("thin".to_string()),
                color: Some(Color {
                    auto: Some(true),
                    indexed: None,
                    rgb: None,
                    theme: None,
                    tint: None,
                }),
            }),
            right: None,
            top: None,
            bottom: None,
            diagonal: None,
        };
        let xml = quick_xml::se::to_string(&border).unwrap();
        assert!(xml.contains("style=\"thin\""));
        let parsed: Border = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(
            parsed.left.as_ref().unwrap().style,
            Some("thin".to_string())
        );
        assert_eq!(
            parsed.left.as_ref().unwrap().color.as_ref().unwrap().auto,
            Some(true)
        );
    }

    #[test]
    fn test_xf_with_alignment() {
        let xf = Xf {
            num_fmt_id: Some(0),
            font_id: Some(0),
            fill_id: Some(0),
            border_id: Some(0),
            xf_id: Some(0),
            apply_number_format: None,
            apply_font: None,
            apply_fill: None,
            apply_border: None,
            apply_alignment: Some(true),
            alignment: Some(Alignment {
                horizontal: Some("center".to_string()),
                vertical: Some("center".to_string()),
                wrap_text: Some(true),
                text_rotation: None,
                indent: None,
                shrink_to_fit: None,
            }),
            protection: None,
        };
        let xml = quick_xml::se::to_string(&xf).unwrap();
        assert!(xml.contains("alignment"));
        assert!(xml.contains("horizontal=\"center\""));
        let parsed: Xf = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.alignment.is_some());
        let align = parsed.alignment.unwrap();
        assert_eq!(align.horizontal, Some("center".to_string()));
        assert_eq!(align.vertical, Some("center".to_string()));
        assert_eq!(align.wrap_text, Some(true));
    }

    #[test]
    fn test_num_fmt() {
        let nf = NumFmt {
            num_fmt_id: 164,
            format_code: "#,##0.00_ ".to_string(),
        };
        let xml = quick_xml::se::to_string(&nf).unwrap();
        let parsed: NumFmt = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.num_fmt_id, 164);
        assert_eq!(parsed.format_code, "#,##0.00_ ");
    }

    #[test]
    fn test_optional_fields_not_serialized() {
        let ss = StyleSheet::default();
        let xml = quick_xml::se::to_string(&ss).unwrap();
        // numFmts is None, so it should not appear
        assert!(!xml.contains("numFmts"));
        // dxfs is None
        assert!(!xml.contains("dxfs"));
        // tableStyles is None
        assert!(!xml.contains("tableStyles"));
    }

    #[test]
    fn test_cell_style_roundtrip() {
        let style = CellStyle {
            name: "Heading 1".to_string(),
            xf_id: 1,
            builtin_id: Some(16),
        };
        let xml = quick_xml::se::to_string(&style).unwrap();
        let parsed: CellStyle = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.name, "Heading 1");
        assert_eq!(parsed.xf_id, 1);
        assert_eq!(parsed.builtin_id, Some(16));
    }
}
