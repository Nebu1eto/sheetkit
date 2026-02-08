//! Rich text run types and conversion utilities.
//!
//! A rich text value is a sequence of [`RichTextRun`] segments, each with its
//! own formatting. These map to the `<r>` (run) elements inside `<si>` items
//! in `xl/sharedStrings.xml`.

use sheetkit_xml::shared_strings::{BoolVal, Color, FontName, FontSize, RPr, Si, R, T};

/// A single formatted text segment within a rich text cell.
#[derive(Debug, Clone, PartialEq)]
pub struct RichTextRun {
    /// The plain text content of this run.
    pub text: String,
    /// Font name (e.g., "Calibri").
    pub font: Option<String>,
    /// Font size in points.
    pub size: Option<f64>,
    /// Whether the text is bold.
    pub bold: bool,
    /// Whether the text is italic.
    pub italic: bool,
    /// Font color as an RGB hex string (e.g., "#FF0000").
    pub color: Option<String>,
}

/// Convert a high-level [`RichTextRun`] into an XML `<r>` element.
pub fn run_to_xml(run: &RichTextRun) -> R {
    let has_formatting =
        run.bold || run.italic || run.font.is_some() || run.size.is_some() || run.color.is_some();

    let r_pr = if has_formatting {
        Some(RPr {
            b: if run.bold {
                Some(BoolVal { val: None })
            } else {
                None
            },
            i: if run.italic {
                Some(BoolVal { val: None })
            } else {
                None
            },
            sz: run.size.map(|val| FontSize { val }),
            color: run.color.as_ref().map(|rgb| Color {
                rgb: Some(rgb.clone()),
                theme: None,
                tint: None,
            }),
            r_font: run.font.as_ref().map(|val| FontName { val: val.clone() }),
            family: None,
            scheme: None,
        })
    } else {
        None
    };

    R {
        r_pr,
        t: T {
            xml_space: if run.text.starts_with(' ')
                || run.text.ends_with(' ')
                || run.text.contains("  ")
                || run.text.contains('\n')
                || run.text.contains('\t')
            {
                Some("preserve".to_string())
            } else {
                None
            },
            value: run.text.clone(),
        },
    }
}

/// Convert an XML `<r>` element into a high-level [`RichTextRun`].
pub fn xml_to_run(r: &R) -> RichTextRun {
    let (font, size, bold, italic, color) = if let Some(ref rpr) = r.r_pr {
        (
            rpr.r_font.as_ref().map(|f| f.val.clone()),
            rpr.sz.as_ref().map(|s| s.val),
            rpr.b.is_some(),
            rpr.i.is_some(),
            rpr.color.as_ref().and_then(|c| c.rgb.clone()),
        )
    } else {
        (None, None, false, false, None)
    };

    RichTextRun {
        text: r.t.value.clone(),
        font,
        size,
        bold,
        italic,
        color,
    }
}

/// Convert a slice of [`RichTextRun`] into an XML `<si>` element (for the SST).
pub fn runs_to_si(runs: &[RichTextRun]) -> Si {
    Si {
        t: None,
        r: runs.iter().map(run_to_xml).collect(),
    }
}

/// Extract plain text from a slice of rich text runs.
pub fn rich_text_to_plain(runs: &[RichTextRun]) -> String {
    runs.iter().map(|r| r.text.as_str()).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_rich_text_to_plain() {
        let runs = vec![
            RichTextRun {
                text: "Hello ".to_string(),
                font: None,
                size: None,
                bold: true,
                italic: false,
                color: None,
            },
            RichTextRun {
                text: "World".to_string(),
                font: None,
                size: None,
                bold: false,
                italic: false,
                color: None,
            },
        ];
        assert_eq!(rich_text_to_plain(&runs), "Hello World");
    }

    #[test]
    fn test_run_to_xml_plain() {
        let run = RichTextRun {
            text: "plain".to_string(),
            font: None,
            size: None,
            bold: false,
            italic: false,
            color: None,
        };
        let xml_r = run_to_xml(&run);
        assert!(xml_r.r_pr.is_none());
        assert_eq!(xml_r.t.value, "plain");
    }

    #[test]
    fn test_run_to_xml_bold() {
        let run = RichTextRun {
            text: "bold".to_string(),
            font: None,
            size: None,
            bold: true,
            italic: false,
            color: None,
        };
        let xml_r = run_to_xml(&run);
        assert!(xml_r.r_pr.is_some());
        assert!(xml_r.r_pr.as_ref().unwrap().b.is_some());
    }

    #[test]
    fn test_xml_to_run_roundtrip() {
        let original = RichTextRun {
            text: "test".to_string(),
            font: Some("Arial".to_string()),
            size: Some(12.0),
            bold: true,
            italic: true,
            color: Some("#FF0000".to_string()),
        };
        let xml_r = run_to_xml(&original);
        let back = xml_to_run(&xml_r);
        assert_eq!(original, back);
    }

    #[test]
    fn test_runs_to_si() {
        let runs = vec![
            RichTextRun {
                text: "A".to_string(),
                font: None,
                size: None,
                bold: true,
                italic: false,
                color: None,
            },
            RichTextRun {
                text: "B".to_string(),
                font: None,
                size: None,
                bold: false,
                italic: false,
                color: None,
            },
        ];
        let si = runs_to_si(&runs);
        assert!(si.t.is_none());
        assert_eq!(si.r.len(), 2);
    }

    #[test]
    fn test_xml_to_run_no_formatting() {
        let r = R {
            r_pr: None,
            t: T {
                xml_space: None,
                value: "text".to_string(),
            },
        };
        let run = xml_to_run(&r);
        assert_eq!(run.text, "text");
        assert!(!run.bold);
        assert!(!run.italic);
        assert!(run.font.is_none());
        assert!(run.size.is_none());
        assert!(run.color.is_none());
    }

    #[test]
    fn test_space_preservation() {
        let run = RichTextRun {
            text: " leading space".to_string(),
            font: None,
            size: None,
            bold: false,
            italic: false,
            color: None,
        };
        let xml_r = run_to_xml(&run);
        assert_eq!(xml_r.t.xml_space, Some("preserve".to_string()));
    }
}
