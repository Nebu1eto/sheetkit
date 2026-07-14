//! Theme color resolution.

use sheetkit_xml::theme::ThemeColors;

/// Maps spreadsheet `theme` values to the DrawingML color scheme order.
///
/// SpreadsheetML indexes `lt1`, `dk1`, `lt2`, and `dk2` first, whereas
/// DrawingML stores those color scheme entries as `dk1`, `lt1`, `dk2`, `lt2`.
const SPREADSHEET_THEME_TO_SCHEME: [usize; 12] = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];

/// Resolve a theme color index to an ARGB hex string.
/// Applies tint modification if specified.
pub fn resolve_theme_color(theme: &ThemeColors, index: u32, tint: Option<f64>) -> Option<String> {
    let scheme_index = *SPREADSHEET_THEME_TO_SCHEME.get(index as usize)?;
    let base = theme.get(scheme_index)?;
    if base.is_empty() {
        return None;
    }
    match tint {
        Some(t) if t != 0.0 => Some(apply_tint(base, t)),
        _ => Some(base.to_string()),
    }
}

/// Apply a tint value to an ARGB hex color.
/// Tint > 0 lightens toward white, tint < 0 darkens toward black.
///
/// SpreadsheetML tint changes HLS luminance, with HLS values from 0 to 255.
fn apply_tint(argb: &str, tint: f64) -> String {
    if argb.len() != 8 {
        return argb.to_string();
    }
    let Ok(alpha) = u8::from_str_radix(&argb[0..2], 16) else {
        return argb.to_string();
    };
    let Ok(red) = u8::from_str_radix(&argb[2..4], 16) else {
        return argb.to_string();
    };
    let Ok(green) = u8::from_str_radix(&argb[4..6], 16) else {
        return argb.to_string();
    };
    let Ok(blue) = u8::from_str_radix(&argb[6..8], 16) else {
        return argb.to_string();
    };

    let tint = if tint.is_finite() {
        tint.clamp(-1.0, 1.0)
    } else {
        0.0
    };
    let (hue, saturation, luminance) = rgb_to_hsl(red, green, blue);
    let luminance = if tint < 0.0 {
        luminance * (1.0 + tint)
    } else {
        luminance * (1.0 - tint) + (1.0 - (1.0 - tint))
    };
    let (red, green, blue) = hsl_to_rgb(hue, saturation, luminance.clamp(0.0, 1.0));

    format!("{alpha:02X}{red:02X}{green:02X}{blue:02X}")
}

fn rgb_to_hsl(red: u8, green: u8, blue: u8) -> (f64, f64, f64) {
    let red = f64::from(red) / 255.0;
    let green = f64::from(green) / 255.0;
    let blue = f64::from(blue) / 255.0;
    let max = red.max(green).max(blue);
    let min = red.min(green).min(blue);
    let luminance = (max + min) / 2.0;
    let delta = max - min;
    if delta == 0.0 {
        return (0.0, 0.0, luminance);
    }

    let saturation = delta / (1.0 - (2.0 * luminance - 1.0).abs());
    let hue = if max == red {
        ((green - blue) / delta).rem_euclid(6.0)
    } else if max == green {
        (blue - red) / delta + 2.0
    } else {
        (red - green) / delta + 4.0
    } / 6.0;
    (hue, saturation, luminance)
}

fn hsl_to_rgb(hue: f64, saturation: f64, luminance: f64) -> (u8, u8, u8) {
    if saturation == 0.0 {
        let gray = channel(luminance);
        return (gray, gray, gray);
    }
    let upper = if luminance <= 0.5 {
        luminance * (1.0 + saturation)
    } else {
        luminance + saturation - luminance * saturation
    };
    let lower = 2.0 * luminance - upper;
    (
        channel(hue_to_rgb(lower, upper, hue + 1.0 / 3.0)),
        channel(hue_to_rgb(lower, upper, hue)),
        channel(hue_to_rgb(lower, upper, hue - 1.0 / 3.0)),
    )
}

fn hue_to_rgb(lower: f64, upper: f64, hue: f64) -> f64 {
    let hue = hue.rem_euclid(1.0);
    if hue < 1.0 / 6.0 {
        lower + (upper - lower) * 6.0 * hue
    } else if hue < 0.5 {
        upper
    } else if hue < 2.0 / 3.0 {
        lower + (upper - lower) * (2.0 / 3.0 - hue) * 6.0
    } else {
        lower
    }
}

fn channel(value: f64) -> u8 {
    (value * 255.0).round().clamp(0.0, 255.0) as u8
}

/// Get the default Office theme colors.
pub fn default_theme_colors() -> ThemeColors {
    ThemeColors {
        colors: [
            "FF000000".to_string(),
            "FFFFFFFF".to_string(),
            "FF44546A".to_string(),
            "FFE7E6E6".to_string(),
            "FF4472C4".to_string(),
            "FFED7D31".to_string(),
            "FFA5A5A5".to_string(),
            "FFFFC000".to_string(),
            "FF5B9BD5".to_string(),
            "FF70AD47".to_string(),
            "FF0563C1".to_string(),
            "FF954F72".to_string(),
        ],
    }
}

/// Generate default theme1.xml content as raw bytes.
pub fn default_theme_xml() -> Vec<u8> {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
    xml.as_bytes().to_vec()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_theme_color_no_tint() {
        let theme = default_theme_colors();
        let color = resolve_theme_color(&theme, 1, None);
        assert_eq!(color, Some("FF000000".to_string()));
    }

    #[test]
    fn test_resolve_theme_color_with_positive_tint() {
        let theme = default_theme_colors();
        let color = resolve_theme_color(&theme, 1, Some(0.5));
        assert_eq!(color, Some("FF808080".to_string()));
    }

    #[test]
    fn test_resolve_invalid_index() {
        let theme = default_theme_colors();
        assert!(resolve_theme_color(&theme, 99, None).is_none());
    }

    #[test]
    fn test_apply_tint_lighten() {
        let result = apply_tint("FF000000", 0.5);
        assert_eq!(result, "FF808080");
    }

    #[test]
    fn test_apply_tint_darken() {
        let result = apply_tint("FFFFFFFF", -0.5);
        assert_eq!(result, "FF808080");
    }

    #[test]
    fn test_apply_tint_zero() {
        let theme = default_theme_colors();
        let color = resolve_theme_color(&theme, 4, Some(0.0));
        assert_eq!(color, Some("FF4472C4".to_string()));
    }

    #[test]
    fn test_spreadsheet_theme_indices_match_cell_style_order() {
        let theme = default_theme_colors();
        let expected = [
            "FFFFFFFF", "FF000000", "FFE7E6E6", "FF44546A", "FF4472C4", "FFED7D31", "FFA5A5A5",
            "FFFFC000", "FF5B9BD5", "FF70AD47", "FF0563C1", "FF954F72",
        ];
        for (index, color) in expected.into_iter().enumerate() {
            assert_eq!(
                resolve_theme_color(&theme, index as u32, None),
                Some(color.to_string())
            );
        }
    }

    #[test]
    fn test_apply_tint_uses_hls_luminance_for_chromatic_colors() {
        assert_eq!(apply_tint("804472C4", 0.5), "80A1B8E2");
        assert_eq!(apply_tint("804472C4", -0.5), "80203864");
    }

    #[test]
    fn test_default_theme_has_all_colors() {
        let theme = default_theme_colors();
        for i in 0..12 {
            assert!(!theme.colors[i].is_empty());
        }
    }

    #[test]
    fn test_default_theme_xml_parseable() {
        let xml_bytes = default_theme_xml();
        let colors = sheetkit_xml::theme::parse_theme_colors(&xml_bytes);
        assert_eq!(colors.colors[0], "FF000000");
        assert_eq!(colors.colors[1], "FFFFFFFF");
        assert_eq!(colors.colors[4], "FF4472C4");
    }
}
