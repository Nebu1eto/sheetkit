//! Theme color resolution.

use sheetkit_xml::theme::ThemeColors;

/// Resolve a theme color index to an ARGB hex string.
/// Applies tint modification if specified.
pub fn resolve_theme_color(theme: &ThemeColors, index: u32, tint: Option<f64>) -> Option<String> {
    let base = theme.get(index as usize)?;
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
fn apply_tint(argb: &str, tint: f64) -> String {
    if argb.len() < 8 {
        return argb.to_string();
    }
    let r = u8::from_str_radix(&argb[2..4], 16).unwrap_or(0);
    let g = u8::from_str_radix(&argb[4..6], 16).unwrap_or(0);
    let b = u8::from_str_radix(&argb[6..8], 16).unwrap_or(0);

    let (r, g, b) = if tint < 0.0 {
        let factor = 1.0 + tint;
        (
            (r as f64 * factor) as u8,
            (g as f64 * factor) as u8,
            (b as f64 * factor) as u8,
        )
    } else {
        (
            (r as f64 + (255.0 - r as f64) * tint) as u8,
            (g as f64 + (255.0 - g as f64) * tint) as u8,
            (b as f64 + (255.0 - b as f64) * tint) as u8,
        )
    };

    format!("FF{:02X}{:02X}{:02X}", r, g, b)
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
        let color = resolve_theme_color(&theme, 0, None);
        assert_eq!(color, Some("FF000000".to_string()));
    }

    #[test]
    fn test_resolve_theme_color_with_positive_tint() {
        let theme = default_theme_colors();
        let color = resolve_theme_color(&theme, 0, Some(0.5));
        assert!(color.is_some());
        let c = color.unwrap();
        assert_eq!(&c[0..2], "FF");
    }

    #[test]
    fn test_resolve_invalid_index() {
        let theme = default_theme_colors();
        assert!(resolve_theme_color(&theme, 99, None).is_none());
    }

    #[test]
    fn test_apply_tint_lighten() {
        let result = apply_tint("FF000000", 0.5);
        assert_eq!(result, "FF7F7F7F");
    }

    #[test]
    fn test_apply_tint_darken() {
        let result = apply_tint("FFFFFFFF", -0.5);
        assert_eq!(result, "FF7F7F7F");
    }

    #[test]
    fn test_apply_tint_zero() {
        let theme = default_theme_colors();
        let color = resolve_theme_color(&theme, 4, Some(0.0));
        assert_eq!(color, Some("FF4472C4".to_string()));
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
