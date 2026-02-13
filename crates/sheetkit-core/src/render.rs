//! SVG renderer for worksheet visual output.
//!
//! Generates an SVG string from worksheet data including cell values,
//! column widths, row heights, cell styles (fonts, fills, borders),
//! optional gridlines, and optional row/column headers.

use crate::cell::CellValue;
use crate::col::get_col_width;
use crate::error::{Error, Result};
use crate::row::{get_row_height, get_rows, resolve_cell_value};
use crate::sst::SharedStringTable;
use crate::style::{
    get_style, AlignmentStyle, BorderLineStyle, FontStyle, HorizontalAlign, PatternType,
    StyleColor, VerticalAlign,
};
use crate::utils::cell_ref::{cell_name_to_coordinates, column_number_to_name};
use sheetkit_xml::styles::StyleSheet;
use sheetkit_xml::worksheet::WorksheetXml;

/// Default column width in pixels (approximately 8.43 characters at 7px each).
const DEFAULT_COL_WIDTH_PX: f64 = 64.0;

/// Default row height in pixels (15 points).
const DEFAULT_ROW_HEIGHT_PX: f64 = 20.0;

/// Pixel width of the row/column header gutter area.
const HEADER_WIDTH: f64 = 40.0;
const HEADER_HEIGHT: f64 = 20.0;

const HEADER_BG_COLOR: &str = "#F0F0F0";
const HEADER_TEXT_COLOR: &str = "#666666";
const GRIDLINE_COLOR: &str = "#D0D0D0";

/// Conversion factor from Excel column width units to pixels.
/// Excel column width is in "number of characters" based on the default font.
/// The approximate conversion is: pixels = width * 7 + 5 (for padding).
fn col_width_to_px(width: f64) -> f64 {
    width * 7.0 + 5.0
}

/// Conversion factor from Excel row height (points) to pixels.
/// 1 point = 4/3 pixels at 96 DPI.
fn row_height_to_px(height: f64) -> f64 {
    height * 4.0 / 3.0
}

/// Options for rendering a worksheet to SVG.
pub struct RenderOptions {
    /// Name of the sheet to render. Required.
    pub sheet_name: String,
    /// Optional cell range to render (e.g. "A1:F20"). None renders the used range.
    pub range: Option<String>,
    /// Whether to draw gridlines between cells.
    pub show_gridlines: bool,
    /// Whether to draw row and column headers (A, B, C... and 1, 2, 3...).
    pub show_headers: bool,
    /// Scale factor for the output (1.0 = 100%).
    pub scale: f64,
    /// Default font family for cell text.
    pub default_font_family: String,
    /// Default font size in points for cell text.
    pub default_font_size: f64,
}

impl Default for RenderOptions {
    fn default() -> Self {
        Self {
            sheet_name: String::new(),
            range: None,
            show_gridlines: true,
            show_headers: true,
            scale: 1.0,
            default_font_family: "Arial".to_string(),
            default_font_size: 11.0,
        }
    }
}

/// Computed layout for a single cell during rendering.
struct CellLayout {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
    col: u32,
    row: u32,
}

/// Render a worksheet to an SVG string.
///
/// Uses the worksheet XML, shared string table, and stylesheet to produce
/// a visual representation of the sheet as SVG. The `options` parameter
/// controls which sheet, range, and visual features to include.
pub fn render_to_svg(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    stylesheet: &StyleSheet,
    options: &RenderOptions,
) -> Result<String> {
    if options.scale <= 0.0 {
        return Err(Error::InvalidArgument(format!(
            "render scale must be positive, got {}",
            options.scale
        )));
    }

    let (min_col, min_row, max_col, max_row) = compute_range(ws, sst, options)?;

    let col_widths = compute_col_widths(ws, min_col, max_col);
    let row_heights = compute_row_heights(ws, min_row, max_row);

    let total_width: f64 = col_widths.iter().sum();
    let total_height: f64 = row_heights.iter().sum();

    let header_x_offset = if options.show_headers {
        HEADER_WIDTH
    } else {
        0.0
    };
    let header_y_offset = if options.show_headers {
        HEADER_HEIGHT
    } else {
        0.0
    };

    let svg_width = (total_width + header_x_offset) * options.scale;
    let svg_height = (total_height + header_y_offset) * options.scale;

    let mut svg = String::with_capacity(4096);
    svg.push_str(&format!(
        r#"<svg xmlns="http://www.w3.org/2000/svg" width="{svg_width}" height="{svg_height}" viewBox="0 0 {} {}">"#,
        total_width + header_x_offset,
        total_height + header_y_offset,
    ));

    svg.push_str(&format!(
        r#"<style>text {{ font-family: {}; font-size: {}px; }}</style>"#,
        &options.default_font_family, options.default_font_size
    ));

    // White background
    svg.push_str(&format!(
        r#"<rect width="{}" height="{}" fill="white"/>"#,
        total_width + header_x_offset,
        total_height + header_y_offset,
    ));

    // Render column headers
    if options.show_headers {
        render_column_headers(&mut svg, &col_widths, min_col, header_x_offset, options);
        render_row_headers(&mut svg, &row_heights, min_row, header_y_offset, options);
    }

    // Build cell layouts
    let layouts = build_cell_layouts(
        &col_widths,
        &row_heights,
        min_col,
        min_row,
        max_col,
        max_row,
        header_x_offset,
        header_y_offset,
    );

    // Render cell fills
    render_cell_fills(&mut svg, ws, sst, stylesheet, &layouts, min_col, min_row);

    // Render gridlines
    if options.show_gridlines {
        render_gridlines(
            &mut svg,
            &col_widths,
            &row_heights,
            total_width,
            total_height,
            header_x_offset,
            header_y_offset,
        );
    }

    // Render cell borders
    render_cell_borders(&mut svg, ws, stylesheet, &layouts, min_col, min_row);

    // Render cell text
    render_cell_text(
        &mut svg, ws, sst, stylesheet, &layouts, min_col, min_row, options,
    );

    svg.push_str("</svg>");
    Ok(svg)
}

/// Determine the range of cells to render.
fn compute_range(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    options: &RenderOptions,
) -> Result<(u32, u32, u32, u32)> {
    if let Some(ref range) = options.range {
        let parts: Vec<&str> = range.split(':').collect();
        if parts.len() != 2 {
            return Err(Error::InvalidCellReference(format!(
                "expected range like 'A1:F20', got '{range}'"
            )));
        }
        let (c1, r1) = cell_name_to_coordinates(parts[0])?;
        let (c2, r2) = cell_name_to_coordinates(parts[1])?;
        Ok((c1.min(c2), r1.min(r2), c1.max(c2), r1.max(r2)))
    } else {
        let rows = get_rows(ws, sst)?;
        if rows.is_empty() {
            return Ok((1, 1, 1, 1));
        }
        let mut min_col = u32::MAX;
        let mut max_col = 0u32;
        let min_row = rows.first().map(|(r, _)| *r).unwrap_or(1);
        let max_row = rows.last().map(|(r, _)| *r).unwrap_or(1);
        for (_, cells) in &rows {
            for (col, _) in cells {
                min_col = min_col.min(*col);
                max_col = max_col.max(*col);
            }
        }
        if min_col == u32::MAX {
            min_col = 1;
        }
        if max_col == 0 {
            max_col = 1;
        }
        Ok((min_col, min_row, max_col, max_row))
    }
}

/// Compute pixel widths for each column in the range.
fn compute_col_widths(ws: &WorksheetXml, min_col: u32, max_col: u32) -> Vec<f64> {
    (min_col..=max_col)
        .map(|col_num| {
            let col_name = column_number_to_name(col_num).unwrap_or_default();
            match get_col_width(ws, &col_name) {
                Some(w) => col_width_to_px(w),
                None => DEFAULT_COL_WIDTH_PX,
            }
        })
        .collect()
}

/// Compute pixel heights for each row in the range.
fn compute_row_heights(ws: &WorksheetXml, min_row: u32, max_row: u32) -> Vec<f64> {
    (min_row..=max_row)
        .map(|row_num| match get_row_height(ws, row_num) {
            Some(h) => row_height_to_px(h),
            None => DEFAULT_ROW_HEIGHT_PX,
        })
        .collect()
}

/// Build a grid of CellLayout structs for every cell position in the range.
#[allow(clippy::too_many_arguments)]
fn build_cell_layouts(
    col_widths: &[f64],
    row_heights: &[f64],
    min_col: u32,
    min_row: u32,
    max_col: u32,
    max_row: u32,
    x_offset: f64,
    y_offset: f64,
) -> Vec<CellLayout> {
    let mut layouts = Vec::new();
    let mut y = y_offset;
    for (ri, row_num) in (min_row..=max_row).enumerate() {
        let h = row_heights[ri];
        let mut x = x_offset;
        for (ci, col_num) in (min_col..=max_col).enumerate() {
            let w = col_widths[ci];
            layouts.push(CellLayout {
                x,
                y,
                width: w,
                height: h,
                col: col_num,
                row: row_num,
            });
            x += w;
        }
        y += h;
    }
    layouts
}

/// Render column header labels (A, B, C, ...).
fn render_column_headers(
    svg: &mut String,
    col_widths: &[f64],
    min_col: u32,
    x_offset: f64,
    _options: &RenderOptions,
) {
    let total_w: f64 = col_widths.iter().sum();
    svg.push_str(&format!(
        "<rect x=\"{x_offset}\" y=\"0\" width=\"{total_w}\" height=\"{HEADER_HEIGHT}\" fill=\"{HEADER_BG_COLOR}\"/>",
    ));

    let mut x = x_offset;
    for (i, &w) in col_widths.iter().enumerate() {
        let col_num = min_col + i as u32;
        let col_name = column_number_to_name(col_num).unwrap_or_default();
        let text_x = x + w / 2.0;
        let text_y = HEADER_HEIGHT / 2.0 + 4.0;
        svg.push_str(&format!(
            "<text x=\"{text_x}\" y=\"{text_y}\" text-anchor=\"middle\" fill=\"{HEADER_TEXT_COLOR}\" font-size=\"10\">{col_name}</text>",
        ));
        x += w;
    }
}

/// Render row header labels (1, 2, 3, ...).
fn render_row_headers(
    svg: &mut String,
    row_heights: &[f64],
    min_row: u32,
    y_offset: f64,
    _options: &RenderOptions,
) {
    let total_h: f64 = row_heights.iter().sum();
    svg.push_str(&format!(
        "<rect x=\"0\" y=\"{y_offset}\" width=\"{HEADER_WIDTH}\" height=\"{total_h}\" fill=\"{HEADER_BG_COLOR}\"/>",
    ));

    let mut y = y_offset;
    for (i, &h) in row_heights.iter().enumerate() {
        let row_num = min_row + i as u32;
        let text_x = HEADER_WIDTH / 2.0;
        let text_y = y + h / 2.0 + 4.0;
        svg.push_str(&format!(
            "<text x=\"{text_x}\" y=\"{text_y}\" text-anchor=\"middle\" fill=\"{HEADER_TEXT_COLOR}\" font-size=\"10\">{row_num}</text>",
        ));
        y += h;
    }
}

/// Render cell background fills.
fn render_cell_fills(
    svg: &mut String,
    ws: &WorksheetXml,
    _sst: &SharedStringTable,
    stylesheet: &StyleSheet,
    layouts: &[CellLayout],
    _min_col: u32,
    _min_row: u32,
) {
    for layout in layouts {
        let style_id = find_cell_style(ws, layout.col, layout.row);
        if style_id == 0 {
            continue;
        }
        if let Some(style) = get_style(stylesheet, style_id) {
            if let Some(ref fill) = style.fill {
                if fill.pattern == PatternType::Solid {
                    if let Some(ref color) = fill.fg_color {
                        let hex = style_color_to_hex(color);
                        svg.push_str(&format!(
                            r#"<rect x="{}" y="{}" width="{}" height="{}" fill="{}"/>"#,
                            layout.x, layout.y, layout.width, layout.height, hex
                        ));
                    }
                }
            }
        }
    }
}

/// Render gridlines.
fn render_gridlines(
    svg: &mut String,
    col_widths: &[f64],
    row_heights: &[f64],
    total_width: f64,
    total_height: f64,
    x_offset: f64,
    y_offset: f64,
) {
    let mut y = y_offset;
    for h in row_heights {
        y += h;
        let x2 = x_offset + total_width;
        svg.push_str(&format!(
            "<line x1=\"{x_offset}\" y1=\"{y}\" x2=\"{x2}\" y2=\"{y}\" stroke=\"{GRIDLINE_COLOR}\" stroke-width=\"0.5\"/>",
        ));
    }

    let mut x = x_offset;
    for w in col_widths {
        x += w;
        let y2 = y_offset + total_height;
        svg.push_str(&format!(
            "<line x1=\"{x}\" y1=\"{y_offset}\" x2=\"{x}\" y2=\"{y2}\" stroke=\"{GRIDLINE_COLOR}\" stroke-width=\"0.5\"/>",
        ));
    }
}

/// Render cell borders.
fn render_cell_borders(
    svg: &mut String,
    ws: &WorksheetXml,
    stylesheet: &StyleSheet,
    layouts: &[CellLayout],
    _min_col: u32,
    _min_row: u32,
) {
    for layout in layouts {
        let style_id = find_cell_style(ws, layout.col, layout.row);
        if style_id == 0 {
            continue;
        }
        let style = match get_style(stylesheet, style_id) {
            Some(s) => s,
            None => continue,
        };
        let border = match &style.border {
            Some(b) => b,
            None => continue,
        };

        let x1 = layout.x;
        let y1 = layout.y;
        let x2 = layout.x + layout.width;
        let y2 = layout.y + layout.height;

        if let Some(ref left) = border.left {
            let (sw, color) = border_line_attrs(left.style, left.color.as_ref());
            svg.push_str(&format!(
                r#"<line x1="{x1}" y1="{y1}" x2="{x1}" y2="{y2}" stroke="{color}" stroke-width="{sw}"/>"#,
            ));
        }
        if let Some(ref right) = border.right {
            let (sw, color) = border_line_attrs(right.style, right.color.as_ref());
            svg.push_str(&format!(
                r#"<line x1="{x2}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{color}" stroke-width="{sw}"/>"#,
            ));
        }
        if let Some(ref top) = border.top {
            let (sw, color) = border_line_attrs(top.style, top.color.as_ref());
            svg.push_str(&format!(
                r#"<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y1}" stroke="{color}" stroke-width="{sw}"/>"#,
            ));
        }
        if let Some(ref bottom) = border.bottom {
            let (sw, color) = border_line_attrs(bottom.style, bottom.color.as_ref());
            svg.push_str(&format!(
                r#"<line x1="{x1}" y1="{y2}" x2="{x2}" y2="{y2}" stroke="{color}" stroke-width="{sw}"/>"#,
            ));
        }
    }
}

/// Render cell text values.
#[allow(clippy::too_many_arguments)]
fn render_cell_text(
    svg: &mut String,
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    stylesheet: &StyleSheet,
    layouts: &[CellLayout],
    _min_col: u32,
    _min_row: u32,
    options: &RenderOptions,
) {
    for layout in layouts {
        let cell_value = find_cell_value(ws, sst, layout.col, layout.row);
        if cell_value == CellValue::Empty {
            continue;
        }

        let display_text = cell_value.to_string();
        if display_text.is_empty() {
            continue;
        }

        let style_id = find_cell_style(ws, layout.col, layout.row);
        let style = get_style(stylesheet, style_id);

        let font = style.as_ref().and_then(|s| s.font.as_ref());
        let alignment = style.as_ref().and_then(|s| s.alignment.as_ref());

        let (text_x, anchor) = compute_text_x(layout, alignment);
        let text_y = compute_text_y(layout, alignment, font, options);

        let escaped = xml_escape(&display_text);

        let mut attrs = String::new();
        attrs.push_str(&format!(r#" x="{text_x}" y="{text_y}""#));
        attrs.push_str(&format!(r#" text-anchor="{anchor}""#));

        if let Some(f) = font {
            if f.bold {
                attrs.push_str(r#" font-weight="bold""#);
            }
            if f.italic {
                attrs.push_str(r#" font-style="italic""#);
            }
            if let Some(ref name) = f.name {
                attrs.push_str(&format!(r#" font-family="{name}""#));
            }
            if let Some(size) = f.size {
                attrs.push_str(&format!(r#" font-size="{size}""#));
            }
            if let Some(ref color) = f.color {
                let hex = style_color_to_hex(color);
                attrs.push_str(&format!(r#" fill="{hex}""#));
            }
            let mut decorations = Vec::new();
            if f.underline {
                decorations.push("underline");
            }
            if f.strikethrough {
                decorations.push("line-through");
            }
            if !decorations.is_empty() {
                attrs.push_str(&format!(r#" text-decoration="{}""#, decorations.join(" ")));
            }
        }

        svg.push_str(&format!("<text{attrs}>{escaped}</text>"));
    }
}

/// Compute the x position and text-anchor for a cell's text based on alignment.
fn compute_text_x(layout: &CellLayout, alignment: Option<&AlignmentStyle>) -> (f64, &'static str) {
    let padding = 3.0;
    match alignment.and_then(|a| a.horizontal) {
        Some(HorizontalAlign::Center) | Some(HorizontalAlign::CenterContinuous) => {
            (layout.x + layout.width / 2.0, "middle")
        }
        Some(HorizontalAlign::Right) => (layout.x + layout.width - padding, "end"),
        _ => (layout.x + padding, "start"),
    }
}

/// Compute the y position for a cell's text based on vertical alignment.
fn compute_text_y(
    layout: &CellLayout,
    alignment: Option<&AlignmentStyle>,
    font: Option<&FontStyle>,
    options: &RenderOptions,
) -> f64 {
    let font_size = font
        .and_then(|f| f.size)
        .unwrap_or(options.default_font_size);
    match alignment.and_then(|a| a.vertical) {
        Some(VerticalAlign::Top) => layout.y + font_size + 2.0,
        Some(VerticalAlign::Center) => layout.y + layout.height / 2.0 + font_size / 3.0,
        _ => layout.y + layout.height - 4.0,
    }
}

/// Find the style ID for a cell at the given coordinates.
fn find_cell_style(ws: &WorksheetXml, col: u32, row: u32) -> u32 {
    ws.sheet_data
        .rows
        .binary_search_by_key(&row, |r| r.r)
        .ok()
        .and_then(|idx| {
            let row_data = &ws.sheet_data.rows[idx];
            row_data
                .cells
                .binary_search_by_key(&col, |c| c.col)
                .ok()
                .and_then(|ci| row_data.cells[ci].s)
        })
        .unwrap_or(0)
}

/// Find the CellValue for a cell at the given coordinates.
fn find_cell_value(ws: &WorksheetXml, sst: &SharedStringTable, col: u32, row: u32) -> CellValue {
    ws.sheet_data
        .rows
        .binary_search_by_key(&row, |r| r.r)
        .ok()
        .and_then(|idx| {
            let row_data = &ws.sheet_data.rows[idx];
            row_data
                .cells
                .binary_search_by_key(&col, |c| c.col)
                .ok()
                .map(|ci| resolve_cell_value(&row_data.cells[ci], sst))
        })
        .unwrap_or(CellValue::Empty)
}

/// Convert a StyleColor to a CSS hex color string.
///
/// Handles several input formats: 8-char ARGB (`FF000000`), 6-char RGB
/// (`000000`), and values already prefixed with `#`. Always returns a
/// `#RRGGBB` string suitable for SVG attributes.
fn style_color_to_hex(color: &StyleColor) -> String {
    match color {
        StyleColor::Rgb(rgb) => {
            let stripped = rgb.strip_prefix('#').unwrap_or(rgb);
            if stripped.len() == 8 {
                // ARGB format (e.g. "FF000000") -> "#000000"
                format!("#{}", &stripped[2..])
            } else {
                format!("#{stripped}")
            }
        }
        StyleColor::Theme(_) | StyleColor::Indexed(_) => "#000000".to_string(),
    }
}

/// Convert a border line style to SVG stroke-width and color.
fn border_line_attrs(style: BorderLineStyle, color: Option<&StyleColor>) -> (f64, String) {
    let stroke_width = match style {
        BorderLineStyle::Thin | BorderLineStyle::Hair => 1.0,
        BorderLineStyle::Medium
        | BorderLineStyle::MediumDashed
        | BorderLineStyle::MediumDashDot
        | BorderLineStyle::MediumDashDotDot => 2.0,
        BorderLineStyle::Thick => 3.0,
        _ => 1.0,
    };
    let color_str = color
        .map(style_color_to_hex)
        .unwrap_or_else(|| "#000000".to_string());
    (stroke_width, color_str)
}

/// Escape special XML characters in text content.
fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(c),
        }
    }
    out
}

#[cfg(test)]
#[allow(clippy::field_reassign_with_default)]
mod tests {
    use super::*;
    use crate::sst::SharedStringTable;
    use crate::style::{add_style, StyleBuilder};
    use sheetkit_xml::styles::StyleSheet;
    use sheetkit_xml::worksheet::{Cell, CellTypeTag, Row, SheetData, WorksheetXml};

    fn default_options(sheet: &str) -> RenderOptions {
        RenderOptions {
            sheet_name: sheet.to_string(),
            ..RenderOptions::default()
        }
    }

    fn make_num_cell(r: &str, col: u32, v: &str) -> Cell {
        Cell {
            r: r.into(),
            col,
            s: None,
            t: CellTypeTag::None,
            v: Some(v.to_string()),
            f: None,
            is: None,
        }
    }

    fn make_sst_cell(r: &str, col: u32, sst_idx: u32) -> Cell {
        Cell {
            r: r.into(),
            col,
            s: None,
            t: CellTypeTag::SharedString,
            v: Some(sst_idx.to_string()),
            f: None,
            is: None,
        }
    }

    fn simple_ws_and_sst() -> (WorksheetXml, SharedStringTable) {
        let mut sst = SharedStringTable::new();
        sst.add("Name"); // 0
        sst.add("Score"); // 1
        sst.add("Alice"); // 2

        let mut ws = WorksheetXml::default();
        ws.sheet_data = SheetData {
            rows: vec![
                Row {
                    r: 1,
                    spans: None,
                    s: None,
                    custom_format: None,
                    ht: None,
                    hidden: None,
                    custom_height: None,
                    outline_level: None,
                    cells: vec![make_sst_cell("A1", 1, 0), make_sst_cell("B1", 2, 1)],
                },
                Row {
                    r: 2,
                    spans: None,
                    s: None,
                    custom_format: None,
                    ht: None,
                    hidden: None,
                    custom_height: None,
                    outline_level: None,
                    cells: vec![make_sst_cell("A2", 1, 2), make_num_cell("B2", 2, "95")],
                },
            ],
        };
        (ws, sst)
    }

    #[test]
    fn test_render_produces_valid_svg() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();
        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(svg.starts_with("<svg"));
        assert!(svg.ends_with("</svg>"));
        assert!(svg.contains("xmlns=\"http://www.w3.org/2000/svg\""));
    }

    #[test]
    fn test_render_contains_cell_text() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();
        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains(">Name<"),
            "SVG should contain cell text 'Name'"
        );
        assert!(
            svg.contains(">Score<"),
            "SVG should contain cell text 'Score'"
        );
        assert!(
            svg.contains(">Alice<"),
            "SVG should contain cell text 'Alice'"
        );
        assert!(svg.contains(">95<"), "SVG should contain cell text '95'");
    }

    #[test]
    fn test_render_contains_headers() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();
        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(svg.contains(">A<"), "SVG should contain column header 'A'");
        assert!(svg.contains(">B<"), "SVG should contain column header 'B'");
        assert!(svg.contains(">1<"), "SVG should contain row header '1'");
        assert!(svg.contains(">2<"), "SVG should contain row header '2'");
    }

    #[test]
    fn test_render_no_headers() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();
        let mut opts = default_options("Sheet1");
        opts.show_headers = false;

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        // With headers off, the header background rects should not appear
        assert!(
            !svg.contains("fill=\"#F0F0F0\""),
            "SVG should not contain header backgrounds"
        );
    }

    #[test]
    fn test_render_no_gridlines() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();
        let mut opts = default_options("Sheet1");
        opts.show_gridlines = false;

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            !svg.contains("stroke=\"#D0D0D0\""),
            "SVG should not contain gridlines"
        );
    }

    #[test]
    fn test_render_with_gridlines() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();
        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains("stroke=\"#D0D0D0\""),
            "SVG should contain gridlines"
        );
    }

    #[test]
    fn test_render_custom_col_widths() {
        let (mut ws, sst) = simple_ws_and_sst();
        crate::col::set_col_width(&mut ws, "A", 20.0).unwrap();

        let ss = StyleSheet::default();
        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(svg.starts_with("<svg"));
        assert!(svg.contains(">Name<"));
    }

    #[test]
    fn test_render_custom_row_heights() {
        let (mut ws, sst) = simple_ws_and_sst();
        crate::row::set_row_height(&mut ws, 1, 30.0).unwrap();

        let ss = StyleSheet::default();
        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(svg.starts_with("<svg"));
        assert!(svg.contains(">Name<"));
    }

    #[test]
    fn test_render_with_range() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();
        let mut opts = default_options("Sheet1");
        opts.range = Some("A1:A2".to_string());

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(svg.contains(">Name<"));
        assert!(svg.contains(">Alice<"));
        // B column should not appear if the range is just A1:A2
        assert!(!svg.contains(">Score<"));
    }

    #[test]
    fn test_render_empty_sheet() {
        let ws = WorksheetXml::default();
        let sst = SharedStringTable::new();
        let ss = StyleSheet::default();
        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(svg.starts_with("<svg"));
        assert!(svg.ends_with("</svg>"));
    }

    #[test]
    fn test_render_bold_text() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();

        let bold_style = StyleBuilder::new().bold(true).build();
        let style_id = add_style(&mut ss, &bold_style).unwrap();

        // Apply style to cell A1
        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains("font-weight=\"bold\""),
            "SVG should contain bold font attribute"
        );
    }

    #[test]
    fn test_render_colored_fill() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();

        let fill_style = StyleBuilder::new().solid_fill("FFFFFF00").build();
        let style_id = add_style(&mut ss, &fill_style).unwrap();

        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains("fill=\"#FFFF00\""),
            "SVG should contain yellow fill color"
        );
    }

    #[test]
    fn test_render_font_color() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();

        let style = StyleBuilder::new().font_color_rgb("FFFF0000").build();
        let style_id = add_style(&mut ss, &style).unwrap();

        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains("fill=\"#FF0000\""),
            "SVG should contain red font color"
        );
    }

    #[test]
    fn test_render_with_shared_strings() {
        let mut sst = SharedStringTable::new();
        sst.add("Hello");
        sst.add("World");

        let mut ws = WorksheetXml::default();
        ws.sheet_data = SheetData {
            rows: vec![Row {
                r: 1,
                spans: None,
                s: None,
                custom_format: None,
                ht: None,
                hidden: None,
                custom_height: None,
                outline_level: None,
                cells: vec![
                    Cell {
                        r: "A1".into(),
                        col: 1,
                        s: None,
                        t: CellTypeTag::SharedString,
                        v: Some("0".to_string()),
                        f: None,
                        is: None,
                    },
                    Cell {
                        r: "B1".into(),
                        col: 2,
                        s: None,
                        t: CellTypeTag::SharedString,
                        v: Some("1".to_string()),
                        f: None,
                        is: None,
                    },
                ],
            }],
        };

        let ss = StyleSheet::default();
        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(svg.contains(">Hello<"));
        assert!(svg.contains(">World<"));
    }

    #[test]
    fn test_render_xml_escaping() {
        let mut ws = WorksheetXml::default();
        ws.sheet_data = SheetData {
            rows: vec![Row {
                r: 1,
                spans: None,
                s: None,
                custom_format: None,
                ht: None,
                hidden: None,
                custom_height: None,
                outline_level: None,
                cells: vec![],
            }],
        };

        let sst = SharedStringTable::new();
        let ss = StyleSheet::default();
        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        // Verify valid XML - at minimum it parses as SVG
        assert!(svg.starts_with("<svg"));
        assert!(svg.ends_with("</svg>"));
    }

    #[test]
    fn test_xml_escape_special_chars() {
        assert_eq!(xml_escape("a&b"), "a&amp;b");
        assert_eq!(xml_escape("a<b"), "a&lt;b");
        assert_eq!(xml_escape("a>b"), "a&gt;b");
        assert_eq!(xml_escape("a\"b"), "a&quot;b");
        assert_eq!(xml_escape("a'b"), "a&apos;b");
        assert_eq!(xml_escape("normal"), "normal");
    }

    #[test]
    fn test_style_color_to_hex_argb() {
        let color = StyleColor::Rgb("FFFF0000".to_string());
        assert_eq!(style_color_to_hex(&color), "#FF0000");
    }

    #[test]
    fn test_style_color_to_hex_rgb() {
        let color = StyleColor::Rgb("00FF00".to_string());
        assert_eq!(style_color_to_hex(&color), "#00FF00");
    }

    #[test]
    fn test_style_color_to_hex_theme_defaults_to_black() {
        let color = StyleColor::Theme(4);
        assert_eq!(style_color_to_hex(&color), "#000000");
    }

    #[test]
    fn test_border_line_attrs_thin() {
        let (sw, color) = border_line_attrs(BorderLineStyle::Thin, None);
        assert_eq!(sw, 1.0);
        assert_eq!(color, "#000000");
    }

    #[test]
    fn test_border_line_attrs_thick_with_color() {
        let c = StyleColor::Rgb("FF0000FF".to_string());
        let (sw, color) = border_line_attrs(BorderLineStyle::Thick, Some(&c));
        assert_eq!(sw, 3.0);
        assert_eq!(color, "#0000FF");
    }

    #[test]
    fn test_render_center_aligned_text() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();

        let style = StyleBuilder::new()
            .horizontal_align(HorizontalAlign::Center)
            .build();
        let style_id = add_style(&mut ss, &style).unwrap();

        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains("text-anchor=\"middle\""),
            "SVG should contain centered text"
        );
    }

    #[test]
    fn test_render_right_aligned_text() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();

        let style = StyleBuilder::new()
            .horizontal_align(HorizontalAlign::Right)
            .build();
        let style_id = add_style(&mut ss, &style).unwrap();

        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains("text-anchor=\"end\""),
            "SVG should contain right-aligned text"
        );
    }

    #[test]
    fn test_render_italic_text() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();

        let style = StyleBuilder::new().italic(true).build();
        let style_id = add_style(&mut ss, &style).unwrap();

        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains("font-style=\"italic\""),
            "SVG should contain italic text"
        );
    }

    #[test]
    fn test_render_border_lines() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();

        let style = StyleBuilder::new()
            .border_all(
                BorderLineStyle::Thin,
                StyleColor::Rgb("FF000000".to_string()),
            )
            .build();
        let style_id = add_style(&mut ss, &style).unwrap();

        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains("stroke=\"#000000\""),
            "SVG should contain border lines"
        );
    }

    #[test]
    fn test_render_invalid_range_returns_error() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();
        let mut opts = default_options("Sheet1");
        opts.range = Some("INVALID".to_string());

        let result = render_to_svg(&ws, &sst, &ss, &opts);
        assert!(result.is_err());
    }

    #[test]
    fn test_render_scale_affects_dimensions() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();

        let mut opts1 = default_options("Sheet1");
        opts1.scale = 1.0;
        let svg1 = render_to_svg(&ws, &sst, &ss, &opts1).unwrap();

        let mut opts2 = default_options("Sheet1");
        opts2.scale = 2.0;
        let svg2 = render_to_svg(&ws, &sst, &ss, &opts2).unwrap();

        // Extract width from the SVG tag
        fn extract_width(svg: &str) -> f64 {
            let start = svg.find("width=\"").unwrap() + 7;
            let end = svg[start..].find('"').unwrap() + start;
            svg[start..end].parse().unwrap()
        }

        let w1 = extract_width(&svg1);
        let w2 = extract_width(&svg2);
        assert!(
            (w2 - w1 * 2.0).abs() < 0.01,
            "scale=2.0 should double the width: {w1} vs {w2}"
        );
    }

    #[test]
    fn test_render_underline_text() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();

        let style = StyleBuilder::new().underline(true).build();
        let style_id = add_style(&mut ss, &style).unwrap();

        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains("text-decoration=\"underline\""),
            "SVG should contain underlined text"
        );
    }

    #[test]
    fn test_render_strikethrough_text() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();

        let style = StyleBuilder::new().strikethrough(true).build();
        let style_id = add_style(&mut ss, &style).unwrap();

        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let opts = default_options("Sheet1");

        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains("text-decoration=\"line-through\""),
            "SVG should contain strikethrough text"
        );
    }

    #[test]
    fn test_style_color_to_hex_already_prefixed() {
        let color = StyleColor::Rgb("#FF0000".to_string());
        assert_eq!(style_color_to_hex(&color), "#FF0000");
    }

    #[test]
    fn test_style_color_to_hex_prefixed_argb() {
        let color = StyleColor::Rgb("#FFFF0000".to_string());
        assert_eq!(style_color_to_hex(&color), "#FF0000");
    }

    #[test]
    fn test_style_color_to_hex_no_double_hash() {
        let color = StyleColor::Rgb("#00FF00".to_string());
        let hex = style_color_to_hex(&color);
        assert!(
            !hex.starts_with("##"),
            "should not produce double hash, got: {hex}"
        );
        assert_eq!(hex, "#00FF00");
    }

    #[test]
    fn test_render_underline_and_strikethrough_merged() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();

        let style = StyleBuilder::new()
            .underline(true)
            .strikethrough(true)
            .build();
        let style_id = add_style(&mut ss, &style).unwrap();

        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let opts = default_options("Sheet1");
        let svg = render_to_svg(&ws, &sst, &ss, &opts).unwrap();

        assert!(
            svg.contains(r#"text-decoration="underline line-through""#),
            "both decorations should be merged in one attribute"
        );
        let count = svg.matches("text-decoration=").count();
        // Only one text-decoration attribute for the cell A1 text element
        assert_eq!(
            count, 1,
            "text-decoration should appear exactly once, found {count}"
        );
    }

    #[test]
    fn test_render_scale_zero_returns_error() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();
        let mut opts = default_options("Sheet1");
        opts.scale = 0.0;

        let result = render_to_svg(&ws, &sst, &ss, &opts);
        assert!(result.is_err(), "scale=0 should return an error");
        let err_msg = result.unwrap_err().to_string();
        assert!(
            err_msg.contains("scale must be positive"),
            "error should mention scale: {err_msg}"
        );
    }

    #[test]
    fn test_render_scale_negative_returns_error() {
        let (ws, sst) = simple_ws_and_sst();
        let ss = StyleSheet::default();
        let mut opts = default_options("Sheet1");
        opts.scale = -1.0;

        let result = render_to_svg(&ws, &sst, &ss, &opts);
        assert!(result.is_err(), "negative scale should return an error");
        let err_msg = result.unwrap_err().to_string();
        assert!(
            err_msg.contains("scale must be positive"),
            "error should mention scale: {err_msg}"
        );
    }
}
