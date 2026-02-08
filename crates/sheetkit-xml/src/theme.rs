//! Theme XML schema structures.
//!
//! Represents `xl/theme/theme1.xml` in the OOXML package.
//! Only the color scheme is parsed; other theme elements are preserved as raw XML.

/// Simplified theme representation focusing on the color scheme.
#[derive(Debug, Clone, Default)]
pub struct ThemeColors {
    /// 12 theme color slots: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink.
    /// Each stored as ARGB hex string (e.g., "FF000000").
    pub colors: [String; 12],
}

impl ThemeColors {
    /// Standard Excel theme color slot names.
    pub const SLOT_NAMES: [&str; 12] = [
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];

    /// Get color by theme index (0-11).
    pub fn get(&self, index: usize) -> Option<&str> {
        self.colors.get(index).map(|s| s.as_str())
    }
}

/// Parse theme colors from theme1.xml raw bytes.
/// Uses quick-xml Reader API directly since the theme namespace is complex.
pub fn parse_theme_colors(xml_bytes: &[u8]) -> ThemeColors {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_reader(xml_bytes);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut colors = ThemeColors::default();
    let mut current_slot: Option<usize> = None;
    let mut in_color_scheme = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");
                if name == "clrScheme" {
                    in_color_scheme = true;
                }
                if in_color_scheme {
                    if let Some(idx) = ThemeColors::SLOT_NAMES.iter().position(|&s| s == name) {
                        current_slot = Some(idx);
                    }
                    if let Some(slot) = current_slot {
                        extract_color_from_element(e, &mut colors, slot);
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                if in_color_scheme {
                    let local_name = e.local_name();
                    let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");
                    if let Some(idx) = ThemeColors::SLOT_NAMES.iter().position(|&s| s == name) {
                        current_slot = Some(idx);
                    }
                    if let Some(slot) = current_slot {
                        extract_color_from_element(e, &mut colors, slot);
                    }
                    if ThemeColors::SLOT_NAMES.contains(&name) {
                        current_slot = None;
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                let name = std::str::from_utf8(local.as_ref()).unwrap_or("");
                if name == "clrScheme" {
                    in_color_scheme = false;
                }
                if in_color_scheme && ThemeColors::SLOT_NAMES.contains(&name) {
                    current_slot = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    colors
}

fn extract_color_from_element(
    e: &quick_xml::events::BytesStart<'_>,
    colors: &mut ThemeColors,
    slot_idx: usize,
) {
    let local_name = e.local_name();
    let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");
    if name == "srgbClr" {
        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == b"val" {
                if let Ok(val) = std::str::from_utf8(&attr.value) {
                    colors.colors[slot_idx] = format!("FF{}", val);
                }
            }
        }
    } else if name == "sysClr" {
        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == b"lastClr" {
                if let Ok(val) = std::str::from_utf8(&attr.value) {
                    colors.colors[slot_idx] = format!("FF{}", val);
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_theme_colors() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
  </a:themeElements>
</a:theme>"#;
        let colors = parse_theme_colors(xml);
        assert_eq!(colors.colors[0], "FF000000");
        assert_eq!(colors.colors[1], "FFFFFFFF");
        assert_eq!(colors.colors[4], "FF4472C4");
        assert_eq!(colors.colors[11], "FF954F72");
    }

    #[test]
    fn test_empty_theme() {
        let colors = parse_theme_colors(b"<a:theme></a:theme>");
        assert_eq!(colors.colors[0], "");
    }

    #[test]
    fn test_theme_color_get() {
        let mut colors = ThemeColors::default();
        colors.colors[0] = "FF000000".to_string();
        assert_eq!(colors.get(0), Some("FF000000"));
        assert_eq!(colors.get(12), None);
    }
}
