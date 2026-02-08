//! Page layout settings for worksheets.
//!
//! Provides functions for configuring page margins, page setup (paper size,
//! orientation, scaling), header/footer, print options, and page breaks.

use sheetkit_xml::worksheet::{
    Break, HeaderFooter, PageMargins, PageSetup, PrintOptions, RowBreaks, WorksheetXml,
};

use crate::error::Result;

// -- Default margin values (inches, matching Excel defaults) --

/// Default left margin in inches.
pub const DEFAULT_MARGIN_LEFT: f64 = 0.7;
/// Default right margin in inches.
pub const DEFAULT_MARGIN_RIGHT: f64 = 0.7;
/// Default top margin in inches.
pub const DEFAULT_MARGIN_TOP: f64 = 0.75;
/// Default bottom margin in inches.
pub const DEFAULT_MARGIN_BOTTOM: f64 = 0.75;
/// Default header margin in inches.
pub const DEFAULT_MARGIN_HEADER: f64 = 0.3;
/// Default footer margin in inches.
pub const DEFAULT_MARGIN_FOOTER: f64 = 0.3;

/// Page margin configuration in inches.
#[derive(Debug, Clone)]
pub struct PageMarginsConfig {
    /// Left margin in inches.
    pub left: f64,
    /// Right margin in inches.
    pub right: f64,
    /// Top margin in inches.
    pub top: f64,
    /// Bottom margin in inches.
    pub bottom: f64,
    /// Header margin in inches.
    pub header: f64,
    /// Footer margin in inches.
    pub footer: f64,
}

impl Default for PageMarginsConfig {
    fn default() -> Self {
        Self {
            left: DEFAULT_MARGIN_LEFT,
            right: DEFAULT_MARGIN_RIGHT,
            top: DEFAULT_MARGIN_TOP,
            bottom: DEFAULT_MARGIN_BOTTOM,
            header: DEFAULT_MARGIN_HEADER,
            footer: DEFAULT_MARGIN_FOOTER,
        }
    }
}

/// Page orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Orientation {
    Portrait,
    Landscape,
}

impl Orientation {
    fn as_str(&self) -> &str {
        match self {
            Orientation::Portrait => "portrait",
            Orientation::Landscape => "landscape",
        }
    }

    fn from_str(s: &str) -> Option<Self> {
        match s {
            "portrait" => Some(Orientation::Portrait),
            "landscape" => Some(Orientation::Landscape),
            _ => None,
        }
    }
}

/// Standard paper sizes. The numeric values follow the OOXML specification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PaperSize {
    Letter = 1,
    Tabloid = 3,
    Legal = 5,
    A3 = 8,
    A4 = 9,
    A5 = 11,
    B4 = 12,
    B5 = 13,
}

impl PaperSize {
    fn as_u32(self) -> u32 {
        self as u32
    }

    /// Convert from a numeric paper size value.
    pub fn from_u32(v: u32) -> Option<Self> {
        match v {
            1 => Some(PaperSize::Letter),
            3 => Some(PaperSize::Tabloid),
            5 => Some(PaperSize::Legal),
            8 => Some(PaperSize::A3),
            9 => Some(PaperSize::A4),
            11 => Some(PaperSize::A5),
            12 => Some(PaperSize::B4),
            13 => Some(PaperSize::B5),
            _ => None,
        }
    }
}

/// Set page margins on a worksheet.
pub fn set_page_margins(ws: &mut WorksheetXml, margins: &PageMarginsConfig) -> Result<()> {
    ws.page_margins = Some(PageMargins {
        left: margins.left,
        right: margins.right,
        top: margins.top,
        bottom: margins.bottom,
        header: margins.header,
        footer: margins.footer,
    });
    Ok(())
}

/// Get page margins from a worksheet, returning defaults if not set.
pub fn get_page_margins(ws: &WorksheetXml) -> PageMarginsConfig {
    match &ws.page_margins {
        Some(pm) => PageMarginsConfig {
            left: pm.left,
            right: pm.right,
            top: pm.top,
            bottom: pm.bottom,
            header: pm.header,
            footer: pm.footer,
        },
        None => PageMarginsConfig::default(),
    }
}

/// Set page setup options on a worksheet.
///
/// Only non-`None` parameters are applied; existing values for `None`
/// parameters are preserved.
pub fn set_page_setup(
    ws: &mut WorksheetXml,
    orientation: Option<Orientation>,
    paper_size: Option<PaperSize>,
    scale: Option<u32>,
    fit_to_width: Option<u32>,
    fit_to_height: Option<u32>,
) -> Result<()> {
    let setup = ws.page_setup.get_or_insert(PageSetup {
        paper_size: None,
        orientation: None,
        scale: None,
        fit_to_width: None,
        fit_to_height: None,
        first_page_number: None,
        horizontal_dpi: None,
        vertical_dpi: None,
        r_id: None,
    });

    if let Some(o) = orientation {
        setup.orientation = Some(o.as_str().to_string());
    }
    if let Some(ps) = paper_size {
        setup.paper_size = Some(ps.as_u32());
    }
    if let Some(s) = scale {
        setup.scale = Some(s);
    }
    if let Some(w) = fit_to_width {
        setup.fit_to_width = Some(w);
    }
    if let Some(h) = fit_to_height {
        setup.fit_to_height = Some(h);
    }
    Ok(())
}

/// Get the current orientation, if set.
pub fn get_orientation(ws: &WorksheetXml) -> Option<Orientation> {
    ws.page_setup
        .as_ref()
        .and_then(|ps| ps.orientation.as_deref())
        .and_then(Orientation::from_str)
}

/// Get the current paper size, if set.
pub fn get_paper_size(ws: &WorksheetXml) -> Option<PaperSize> {
    ws.page_setup
        .as_ref()
        .and_then(|ps| ps.paper_size)
        .and_then(PaperSize::from_u32)
}

/// Get the current scale, if set.
pub fn get_scale(ws: &WorksheetXml) -> Option<u32> {
    ws.page_setup.as_ref().and_then(|ps| ps.scale)
}

/// Get the current fit-to-width value, if set.
pub fn get_fit_to_width(ws: &WorksheetXml) -> Option<u32> {
    ws.page_setup.as_ref().and_then(|ps| ps.fit_to_width)
}

/// Get the current fit-to-height value, if set.
pub fn get_fit_to_height(ws: &WorksheetXml) -> Option<u32> {
    ws.page_setup.as_ref().and_then(|ps| ps.fit_to_height)
}

/// Set header and footer text for printing.
///
/// Both parameters are optional. Pass `None` to leave a field unchanged.
/// Excel header/footer text uses special formatting codes (e.g., `&L`, `&C`,
/// `&R` for left/center/right sections).
pub fn set_header_footer(
    ws: &mut WorksheetXml,
    header: Option<&str>,
    footer: Option<&str>,
) -> Result<()> {
    let hf = ws.header_footer.get_or_insert(HeaderFooter {
        odd_header: None,
        odd_footer: None,
    });
    if let Some(h) = header {
        hf.odd_header = Some(h.to_string());
    }
    if let Some(f) = footer {
        hf.odd_footer = Some(f.to_string());
    }
    Ok(())
}

/// Get the header and footer text.
///
/// Returns `(header, footer)` where each may be `None` if not set.
pub fn get_header_footer(ws: &WorksheetXml) -> (Option<String>, Option<String>) {
    match &ws.header_footer {
        Some(hf) => (hf.odd_header.clone(), hf.odd_footer.clone()),
        None => (None, None),
    }
}

/// Set print options on a worksheet.
///
/// Only non-`None` parameters are applied; existing values for `None`
/// parameters are preserved.
pub fn set_print_options(
    ws: &mut WorksheetXml,
    grid_lines: Option<bool>,
    headings: Option<bool>,
    h_centered: Option<bool>,
    v_centered: Option<bool>,
) -> Result<()> {
    let opts = ws.print_options.get_or_insert(PrintOptions {
        grid_lines: None,
        headings: None,
        horizontal_centered: None,
        vertical_centered: None,
    });
    if let Some(gl) = grid_lines {
        opts.grid_lines = Some(gl);
    }
    if let Some(h) = headings {
        opts.headings = Some(h);
    }
    if let Some(hc) = h_centered {
        opts.horizontal_centered = Some(hc);
    }
    if let Some(vc) = v_centered {
        opts.vertical_centered = Some(vc);
    }
    Ok(())
}

/// Get print options from a worksheet.
///
/// Returns `(grid_lines, headings, horizontal_centered, vertical_centered)`.
pub fn get_print_options(
    ws: &WorksheetXml,
) -> (Option<bool>, Option<bool>, Option<bool>, Option<bool>) {
    match &ws.print_options {
        Some(po) => (
            po.grid_lines,
            po.headings,
            po.horizontal_centered,
            po.vertical_centered,
        ),
        None => (None, None, None, None),
    }
}

/// Insert a horizontal page break before the given 1-based row number.
///
/// If a break already exists at this row, the call is a no-op.
pub fn insert_page_break(ws: &mut WorksheetXml, row: u32) -> Result<()> {
    let rb = ws.row_breaks.get_or_insert_with(|| RowBreaks {
        count: None,
        manual_break_count: None,
        brk: vec![],
    });

    // Avoid duplicate breaks.
    if rb.brk.iter().any(|b| b.id == row) {
        return Ok(());
    }

    rb.brk.push(Break {
        id: row,
        max: Some(16383),
        man: Some(true),
    });

    // Sort breaks by row number.
    rb.brk.sort_by_key(|b| b.id);

    let count = rb.brk.len() as u32;
    rb.count = Some(count);
    rb.manual_break_count = Some(count);

    Ok(())
}

/// Remove a horizontal page break at the given 1-based row number.
///
/// If no break exists at this row, the call is a no-op.
pub fn remove_page_break(ws: &mut WorksheetXml, row: u32) -> Result<()> {
    if let Some(rb) = &mut ws.row_breaks {
        rb.brk.retain(|b| b.id != row);
        if rb.brk.is_empty() {
            ws.row_breaks = None;
        } else {
            let count = rb.brk.len() as u32;
            rb.count = Some(count);
            rb.manual_break_count = Some(count);
        }
    }
    Ok(())
}

/// Get all row page break positions (1-based row numbers).
pub fn get_page_breaks(ws: &WorksheetXml) -> Vec<u32> {
    match &ws.row_breaks {
        Some(rb) => rb.brk.iter().map(|b| b.id).collect(),
        None => vec![],
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetkit_xml::worksheet::WorksheetXml;

    // -- Page margins tests --

    #[test]
    fn test_set_get_page_margins() {
        let mut ws = WorksheetXml::default();
        let margins = PageMarginsConfig {
            left: 1.0,
            right: 1.0,
            top: 1.5,
            bottom: 1.5,
            header: 0.5,
            footer: 0.5,
        };
        set_page_margins(&mut ws, &margins).unwrap();

        let result = get_page_margins(&ws);
        assert_eq!(result.left, 1.0);
        assert_eq!(result.right, 1.0);
        assert_eq!(result.top, 1.5);
        assert_eq!(result.bottom, 1.5);
        assert_eq!(result.header, 0.5);
        assert_eq!(result.footer, 0.5);
    }

    #[test]
    fn test_default_page_margins() {
        let ws = WorksheetXml::default();
        let margins = get_page_margins(&ws);
        assert_eq!(margins.left, DEFAULT_MARGIN_LEFT);
        assert_eq!(margins.right, DEFAULT_MARGIN_RIGHT);
        assert_eq!(margins.top, DEFAULT_MARGIN_TOP);
        assert_eq!(margins.bottom, DEFAULT_MARGIN_BOTTOM);
        assert_eq!(margins.header, DEFAULT_MARGIN_HEADER);
        assert_eq!(margins.footer, DEFAULT_MARGIN_FOOTER);
    }

    #[test]
    fn test_page_margins_default_values() {
        let config = PageMarginsConfig::default();
        assert_eq!(config.left, 0.7);
        assert_eq!(config.right, 0.7);
        assert_eq!(config.top, 0.75);
        assert_eq!(config.bottom, 0.75);
        assert_eq!(config.header, 0.3);
        assert_eq!(config.footer, 0.3);
    }

    // -- Orientation tests --

    #[test]
    fn test_set_orientation_portrait() {
        let mut ws = WorksheetXml::default();
        set_page_setup(&mut ws, Some(Orientation::Portrait), None, None, None, None).unwrap();
        assert_eq!(get_orientation(&ws), Some(Orientation::Portrait));
    }

    #[test]
    fn test_set_orientation_landscape() {
        let mut ws = WorksheetXml::default();
        set_page_setup(
            &mut ws,
            Some(Orientation::Landscape),
            None,
            None,
            None,
            None,
        )
        .unwrap();
        assert_eq!(get_orientation(&ws), Some(Orientation::Landscape));
    }

    #[test]
    fn test_orientation_none_when_not_set() {
        let ws = WorksheetXml::default();
        assert_eq!(get_orientation(&ws), None);
    }

    // -- Paper size tests --

    #[test]
    fn test_set_paper_size_a4() {
        let mut ws = WorksheetXml::default();
        set_page_setup(&mut ws, None, Some(PaperSize::A4), None, None, None).unwrap();
        assert_eq!(get_paper_size(&ws), Some(PaperSize::A4));
    }

    #[test]
    fn test_set_paper_size_letter() {
        let mut ws = WorksheetXml::default();
        set_page_setup(&mut ws, None, Some(PaperSize::Letter), None, None, None).unwrap();
        assert_eq!(get_paper_size(&ws), Some(PaperSize::Letter));
    }

    #[test]
    fn test_paper_size_none_when_not_set() {
        let ws = WorksheetXml::default();
        assert_eq!(get_paper_size(&ws), None);
    }

    // -- Scale tests --

    #[test]
    fn test_set_scale() {
        let mut ws = WorksheetXml::default();
        set_page_setup(&mut ws, None, None, Some(75), None, None).unwrap();
        assert_eq!(get_scale(&ws), Some(75));
    }

    #[test]
    fn test_scale_none_when_not_set() {
        let ws = WorksheetXml::default();
        assert_eq!(get_scale(&ws), None);
    }

    // -- Fit to page tests --

    #[test]
    fn test_set_fit_to_page() {
        let mut ws = WorksheetXml::default();
        set_page_setup(&mut ws, None, None, None, Some(1), Some(2)).unwrap();
        assert_eq!(get_fit_to_width(&ws), Some(1));
        assert_eq!(get_fit_to_height(&ws), Some(2));
    }

    #[test]
    fn test_fit_to_page_none_when_not_set() {
        let ws = WorksheetXml::default();
        assert_eq!(get_fit_to_width(&ws), None);
        assert_eq!(get_fit_to_height(&ws), None);
    }

    // -- Multiple page setup options --

    #[test]
    fn test_set_page_setup_combined() {
        let mut ws = WorksheetXml::default();
        set_page_setup(
            &mut ws,
            Some(Orientation::Landscape),
            Some(PaperSize::Legal),
            Some(80),
            None,
            None,
        )
        .unwrap();

        assert_eq!(get_orientation(&ws), Some(Orientation::Landscape));
        assert_eq!(get_paper_size(&ws), Some(PaperSize::Legal));
        assert_eq!(get_scale(&ws), Some(80));
    }

    #[test]
    fn test_set_page_setup_preserves_existing() {
        let mut ws = WorksheetXml::default();
        set_page_setup(
            &mut ws,
            Some(Orientation::Landscape),
            Some(PaperSize::A4),
            None,
            None,
            None,
        )
        .unwrap();

        // Second call should preserve orientation/paper size.
        set_page_setup(&mut ws, None, None, Some(50), None, None).unwrap();

        assert_eq!(get_orientation(&ws), Some(Orientation::Landscape));
        assert_eq!(get_paper_size(&ws), Some(PaperSize::A4));
        assert_eq!(get_scale(&ws), Some(50));
    }

    // -- Header footer tests --

    #[test]
    fn test_set_header_footer() {
        let mut ws = WorksheetXml::default();
        set_header_footer(&mut ws, Some("&CPage &P"), Some("&LFooter Left")).unwrap();

        let (header, footer) = get_header_footer(&ws);
        assert_eq!(header, Some("&CPage &P".to_string()));
        assert_eq!(footer, Some("&LFooter Left".to_string()));
    }

    #[test]
    fn test_set_header_only() {
        let mut ws = WorksheetXml::default();
        set_header_footer(&mut ws, Some("My Header"), None).unwrap();

        let (header, footer) = get_header_footer(&ws);
        assert_eq!(header, Some("My Header".to_string()));
        assert_eq!(footer, None);
    }

    #[test]
    fn test_set_footer_only() {
        let mut ws = WorksheetXml::default();
        set_header_footer(&mut ws, None, Some("My Footer")).unwrap();

        let (header, footer) = get_header_footer(&ws);
        assert_eq!(header, None);
        assert_eq!(footer, Some("My Footer".to_string()));
    }

    #[test]
    fn test_get_header_footer_none() {
        let ws = WorksheetXml::default();
        let (header, footer) = get_header_footer(&ws);
        assert_eq!(header, None);
        assert_eq!(footer, None);
    }

    #[test]
    fn test_set_header_footer_preserves_existing() {
        let mut ws = WorksheetXml::default();
        set_header_footer(&mut ws, Some("Header1"), Some("Footer1")).unwrap();
        // Update only header.
        set_header_footer(&mut ws, Some("Header2"), None).unwrap();

        let (header, footer) = get_header_footer(&ws);
        assert_eq!(header, Some("Header2".to_string()));
        assert_eq!(footer, Some("Footer1".to_string()));
    }

    // -- Print options tests --

    #[test]
    fn test_set_print_options() {
        let mut ws = WorksheetXml::default();
        set_print_options(&mut ws, Some(true), Some(true), Some(true), Some(false)).unwrap();

        let (gl, hd, hc, vc) = get_print_options(&ws);
        assert_eq!(gl, Some(true));
        assert_eq!(hd, Some(true));
        assert_eq!(hc, Some(true));
        assert_eq!(vc, Some(false));
    }

    #[test]
    fn test_get_print_options_none() {
        let ws = WorksheetXml::default();
        let (gl, hd, hc, vc) = get_print_options(&ws);
        assert_eq!(gl, None);
        assert_eq!(hd, None);
        assert_eq!(hc, None);
        assert_eq!(vc, None);
    }

    #[test]
    fn test_set_print_options_partial() {
        let mut ws = WorksheetXml::default();
        set_print_options(&mut ws, Some(true), None, None, None).unwrap();

        let (gl, hd, hc, vc) = get_print_options(&ws);
        assert_eq!(gl, Some(true));
        assert_eq!(hd, None);
        assert_eq!(hc, None);
        assert_eq!(vc, None);
    }

    #[test]
    fn test_set_print_options_preserves_existing() {
        let mut ws = WorksheetXml::default();
        set_print_options(&mut ws, Some(true), Some(false), None, None).unwrap();
        set_print_options(&mut ws, None, None, Some(true), None).unwrap();

        let (gl, hd, hc, vc) = get_print_options(&ws);
        assert_eq!(gl, Some(true));
        assert_eq!(hd, Some(false));
        assert_eq!(hc, Some(true));
        assert_eq!(vc, None);
    }

    // -- Page break tests --

    #[test]
    fn test_insert_page_break() {
        let mut ws = WorksheetXml::default();
        insert_page_break(&mut ws, 10).unwrap();

        let breaks = get_page_breaks(&ws);
        assert_eq!(breaks, vec![10]);
    }

    #[test]
    fn test_insert_multiple_page_breaks() {
        let mut ws = WorksheetXml::default();
        insert_page_break(&mut ws, 20).unwrap();
        insert_page_break(&mut ws, 10).unwrap();
        insert_page_break(&mut ws, 30).unwrap();

        let breaks = get_page_breaks(&ws);
        assert_eq!(breaks, vec![10, 20, 30]);
    }

    #[test]
    fn test_insert_duplicate_page_break_is_noop() {
        let mut ws = WorksheetXml::default();
        insert_page_break(&mut ws, 10).unwrap();
        insert_page_break(&mut ws, 10).unwrap();

        let breaks = get_page_breaks(&ws);
        assert_eq!(breaks, vec![10]);
    }

    #[test]
    fn test_remove_page_break() {
        let mut ws = WorksheetXml::default();
        insert_page_break(&mut ws, 10).unwrap();
        insert_page_break(&mut ws, 20).unwrap();

        remove_page_break(&mut ws, 10).unwrap();

        let breaks = get_page_breaks(&ws);
        assert_eq!(breaks, vec![20]);
    }

    #[test]
    fn test_remove_last_page_break_clears_element() {
        let mut ws = WorksheetXml::default();
        insert_page_break(&mut ws, 10).unwrap();
        remove_page_break(&mut ws, 10).unwrap();

        assert!(ws.row_breaks.is_none());
        assert!(get_page_breaks(&ws).is_empty());
    }

    #[test]
    fn test_remove_nonexistent_page_break_is_noop() {
        let mut ws = WorksheetXml::default();
        insert_page_break(&mut ws, 10).unwrap();
        remove_page_break(&mut ws, 99).unwrap();

        let breaks = get_page_breaks(&ws);
        assert_eq!(breaks, vec![10]);
    }

    #[test]
    fn test_get_page_breaks_empty() {
        let ws = WorksheetXml::default();
        assert!(get_page_breaks(&ws).is_empty());
    }

    #[test]
    fn test_page_break_max_value() {
        let mut ws = WorksheetXml::default();
        insert_page_break(&mut ws, 5).unwrap();

        let rb = ws.row_breaks.as_ref().unwrap();
        assert_eq!(rb.brk[0].max, Some(16383));
        assert_eq!(rb.brk[0].man, Some(true));
    }

    #[test]
    fn test_page_break_count_updated() {
        let mut ws = WorksheetXml::default();
        insert_page_break(&mut ws, 10).unwrap();
        insert_page_break(&mut ws, 20).unwrap();
        insert_page_break(&mut ws, 30).unwrap();

        let rb = ws.row_breaks.as_ref().unwrap();
        assert_eq!(rb.count, Some(3));
        assert_eq!(rb.manual_break_count, Some(3));

        remove_page_break(&mut ws, 20).unwrap();
        let rb = ws.row_breaks.as_ref().unwrap();
        assert_eq!(rb.count, Some(2));
        assert_eq!(rb.manual_break_count, Some(2));
    }

    // -- Paper size enum tests --

    #[test]
    fn test_paper_size_values() {
        assert_eq!(PaperSize::Letter.as_u32(), 1);
        assert_eq!(PaperSize::Tabloid.as_u32(), 3);
        assert_eq!(PaperSize::Legal.as_u32(), 5);
        assert_eq!(PaperSize::A3.as_u32(), 8);
        assert_eq!(PaperSize::A4.as_u32(), 9);
        assert_eq!(PaperSize::A5.as_u32(), 11);
        assert_eq!(PaperSize::B4.as_u32(), 12);
        assert_eq!(PaperSize::B5.as_u32(), 13);
    }

    #[test]
    fn test_paper_size_from_u32() {
        assert_eq!(PaperSize::from_u32(1), Some(PaperSize::Letter));
        assert_eq!(PaperSize::from_u32(9), Some(PaperSize::A4));
        assert_eq!(PaperSize::from_u32(99), None);
    }

    // -- Orientation enum tests --

    #[test]
    fn test_orientation_as_str() {
        assert_eq!(Orientation::Portrait.as_str(), "portrait");
        assert_eq!(Orientation::Landscape.as_str(), "landscape");
    }

    #[test]
    fn test_orientation_from_str() {
        assert_eq!(
            Orientation::from_str("portrait"),
            Some(Orientation::Portrait)
        );
        assert_eq!(
            Orientation::from_str("landscape"),
            Some(Orientation::Landscape)
        );
        assert_eq!(Orientation::from_str("unknown"), None);
    }
}
