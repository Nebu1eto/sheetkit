//! SVG renderer for worksheet visual output.
//!
//! Generates an SVG string from worksheet data including cell values,
//! column widths, row heights, cell styles (fonts, fills, borders),
//! optional gridlines, and optional row/column headers.

use std::collections::HashMap;

use crate::error::{Error, Result};
use crate::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};
use crate::row::get_row_height;
use crate::sst::SharedStringTable;
use crate::style::{
    get_style, AlignmentStyle, BorderLineStyle, FontStyle, HorizontalAlign, PatternType, Style,
    StyleColor, VerticalAlign,
};
use crate::theme::{default_theme_colors, resolve_theme_color};
use crate::traversal::{CheckedRect, TraversalLimits};
use crate::utils::cell_ref::{cell_name_to_coordinates, column_number_to_name};
use sheetkit_xml::styles::StyleSheet;
use sheetkit_xml::theme::ThemeColors;
use sheetkit_xml::worksheet::{Cell, CellTypeTag, InlineString, WorksheetXml};

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

/// Maximum size of one non-streaming SVG result.
pub const MAX_SVG_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;

struct SvgOutput {
    value: String,
    limit: u64,
    base_offset: u64,
}

impl SvgOutput {
    const fn new(limit: u64) -> Self {
        Self {
            value: String::new(),
            limit,
            base_offset: 0,
        }
    }

    fn stage(&self) -> Self {
        Self {
            value: String::new(),
            limit: self.limit,
            base_offset: self.observed_len(),
        }
    }

    /// Like [`Self::stage`], reusing `buffer`'s allocation.
    fn stage_with(&self, mut buffer: String) -> Self {
        buffer.clear();
        Self {
            value: buffer,
            limit: self.limit,
            base_offset: self.observed_len(),
        }
    }

    /// Return `true` when `additional` more bytes stay within the limit.
    fn fits(&self, additional: usize) -> bool {
        self.observed_len()
            .saturating_add(u64::try_from(additional).unwrap_or(u64::MAX))
            <= self.limit
    }

    /// Push formatted text as one piece, formatting into a reusable buffer
    /// instead of allocating a new `String` per call.
    fn push_fmt(&mut self, scratch: &mut String, args: std::fmt::Arguments<'_>) -> Result<()> {
        use std::fmt::Write as _;
        scratch.clear();
        scratch
            .write_fmt(args)
            .map_err(|_| Error::Internal("SVG formatting failed".to_string()))?;
        self.push_str(scratch)
    }

    fn push_str(&mut self, value: &str) -> Result<()> {
        let additional = u64::try_from(value.len()).unwrap_or(u64::MAX);
        self.ensure_additional(additional)?;

        self.value
            .try_reserve(value.len())
            .map_err(|_| Error::Internal("SVG output allocation failed".to_string()))?;
        self.value.push_str(value);
        Ok(())
    }

    /// Push a run of unescaped characters with the same limit outcome as
    /// pushing them one character at a time: a run that fits is pushed
    /// whole; otherwise characters are pushed individually so the error
    /// reports the same observed length.
    fn push_chars(&mut self, run: &str) -> Result<()> {
        let observed = self
            .observed_len()
            .saturating_add(u64::try_from(run.len()).unwrap_or(u64::MAX));
        if observed <= self.limit {
            return self.push_str(run);
        }
        for c in run.chars() {
            let mut encoded = [0; 4];
            self.push_str(c.encode_utf8(&mut encoded))?;
        }
        Ok(())
    }

    /// Push `value`, replacing each character for which `escape` returns a
    /// replacement. Equivalent to pushing every character or replacement
    /// individually, including the limit error.
    fn push_escaped(
        &mut self,
        value: &str,
        escape: impl Fn(char) -> Option<&'static str>,
    ) -> Result<()> {
        let mut start = 0;
        for (index, c) in value.char_indices() {
            if let Some(replacement) = escape(c) {
                if start < index {
                    self.push_chars(&value[start..index])?;
                }
                self.push_str(replacement)?;
                start = index + c.len_utf8();
            }
        }
        if start < value.len() {
            self.push_chars(&value[start..])?;
        }
        Ok(())
    }

    fn ensure_additional(&self, additional: u64) -> Result<()> {
        let observed = self.observed_len().saturating_add(additional);
        if observed > self.limit {
            return Err(svg_output_limit_error(self.limit, observed));
        }
        Ok(())
    }

    fn observed_len(&self) -> u64 {
        self.base_offset
            .saturating_add(u64::try_from(self.value.len()).unwrap_or(u64::MAX))
    }

    fn finish(self) -> String {
        self.value
    }
}

fn svg_output_limit_error(limit: u64, observed: u64) -> Error {
    let observed = observed.max(limit.saturating_add(1));
    match ResourceLimitError::exceeded(
        ResourceOperation::Render,
        ResourceCategory::Output,
        ResourceDimension::OutputBytes,
        limit,
        observed,
        Some(ResourceStage::Render),
    ) {
        Some(metadata) => metadata.into(),
        None => Error::Internal("invalid SVG output limit".to_string()),
    }
}

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
    column_style: u32,
}

struct ColumnMetadataIndex {
    widths: Vec<f64>,
    styles: Vec<u32>,
}

impl ColumnMetadataIndex {
    fn build(ws: &WorksheetXml, min_col: u32, max_col: u32) -> Result<(Self, usize)> {
        let count = usize::try_from(u64::from(max_col - min_col) + 1)
            .map_err(|_| Error::Internal("render column count is not addressable".to_string()))?;

        let mut widths = Vec::new();
        widths
            .try_reserve_exact(count)
            .map_err(|_| Error::Internal("render column layout allocation failed".to_string()))?;
        widths.resize(count, DEFAULT_COL_WIDTH_PX);

        let mut styles = Vec::new();
        styles
            .try_reserve_exact(count)
            .map_err(|_| Error::Internal("render column layout allocation failed".to_string()))?;
        styles.resize(count, 0);

        let parent_count = count.checked_add(1).ok_or_else(|| {
            Error::Internal("render column metadata index is too large".to_string())
        })?;
        let mut next_unresolved = Vec::new();
        next_unresolved
            .try_reserve_exact(parent_count)
            .map_err(|_| {
                Error::Internal("render column metadata index allocation failed".to_string())
            })?;
        next_unresolved.extend(0..parent_count);

        let mut resolution_work = 0usize;
        if let Some(columns) = ws.cols.as_ref() {
            // Resolve in source order to preserve the renderer's first-match semantics.
            for column in &columns.cols {
                resolution_work = resolution_work.saturating_add(1);
                if column.min > column.max || column.max < min_col || column.min > max_col {
                    continue;
                }

                let start = usize::try_from(column.min.max(min_col) - min_col).map_err(|_| {
                    Error::Internal("render column start is not addressable".to_string())
                })?;
                let end = usize::try_from(column.max.min(max_col) - min_col).map_err(|_| {
                    Error::Internal("render column end is not addressable".to_string())
                })?;
                let mut offset = find_next_unresolved(&mut next_unresolved, start);
                while offset <= end {
                    widths[offset] = column
                        .width
                        .map(col_width_to_px)
                        .unwrap_or(DEFAULT_COL_WIDTH_PX);
                    styles[offset] = column.style.unwrap_or(0);
                    resolution_work = resolution_work.saturating_add(1);

                    let next = find_next_unresolved(&mut next_unresolved, offset + 1);
                    next_unresolved[offset] = next;
                    offset = next;
                }
            }
        }

        Ok((Self { widths, styles }, resolution_work))
    }
}

fn find_next_unresolved(parent: &mut [usize], index: usize) -> usize {
    let mut root = index;
    while parent[root] != root {
        root = parent[root];
    }

    let mut current = index;
    while parent[current] != current {
        let next = parent[current];
        parent[current] = root;
        current = next;
    }
    root
}

/// Per-render memo of [`get_style`] results. Styles are pure functions of
/// the stylesheet, so each referenced style id is resolved once.
struct StyleCache<'s> {
    stylesheet: &'s StyleSheet,
    entries: HashMap<u32, Option<Style>>,
}

impl<'s> StyleCache<'s> {
    fn new(stylesheet: &'s StyleSheet) -> Self {
        Self {
            stylesheet,
            entries: HashMap::new(),
        }
    }

    fn get(&mut self, style_id: u32) -> Option<&Style> {
        let stylesheet = self.stylesheet;
        self.entries
            .entry(style_id)
            .or_insert_with(|| get_style(stylesheet, style_id))
            .as_ref()
    }
}

struct LayoutGrid<'a> {
    col_widths: &'a [f64],
    col_styles: &'a [u32],
    row_heights: &'a [f64],
    min_col: u32,
    min_row: u32,
    x_offset: f64,
    y_offset: f64,
}

impl LayoutGrid<'_> {
    fn for_each(&self, mut visitor: impl FnMut(&CellLayout) -> Result<()>) -> Result<()> {
        let mut y = self.y_offset;
        for (row_index, height) in self.row_heights.iter().copied().enumerate() {
            let row = self.min_row + row_index as u32;
            let mut x = self.x_offset;
            for (column_index, width) in self.col_widths.iter().copied().enumerate() {
                let col = self.min_col + column_index as u32;
                visitor(&CellLayout {
                    x,
                    y,
                    width,
                    height,
                    col,
                    row,
                    column_style: self.col_styles[column_index],
                })?;
                x += width;
            }
            y += height;
        }
        Ok(())
    }
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
    render_to_svg_with_limits(ws, sst, stylesheet, options, TraversalLimits::default())
}

/// Renders a worksheet using explicit finite dense traversal limits.
pub fn render_to_svg_with_limits(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    stylesheet: &StyleSheet,
    options: &RenderOptions,
    limits: TraversalLimits,
) -> Result<String> {
    render_to_svg_with_theme_and_limits(
        ws,
        sst,
        stylesheet,
        &default_theme_colors(),
        options,
        limits,
    )
}

/// Render a worksheet to SVG using the supplied theme colors.
pub(crate) fn render_to_svg_with_theme(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    stylesheet: &StyleSheet,
    theme: &ThemeColors,
    options: &RenderOptions,
) -> Result<String> {
    render_to_svg_with_theme_and_limits(
        ws,
        sst,
        stylesheet,
        theme,
        options,
        TraversalLimits::default(),
    )
}

pub(crate) fn render_to_svg_with_theme_and_limits(
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    stylesheet: &StyleSheet,
    theme: &ThemeColors,
    options: &RenderOptions,
    limits: TraversalLimits,
) -> Result<String> {
    validate_positive_finite(options.scale, "render scale must be positive and finite")?;
    validate_positive_finite(
        options.default_font_size,
        "default font size must be positive and finite",
    )?;
    let output_limit = MAX_SVG_OUTPUT_BYTES.min(limits.max_output_bytes());
    let mut svg = SvgOutput::new(output_limit);
    let default_font_family = css_font_family_attribute(&options.default_font_family, &svg)?;

    let (min_col, min_row, max_col, max_row) = compute_range(ws, sst, options, limits)?;
    let _rect = render_checked_rect(min_col, min_row, max_col, max_row, limits)?;

    let (column_metadata, _metadata_work) = ColumnMetadataIndex::build(ws, min_col, max_col)?;
    let col_widths = column_metadata.widths.as_slice();
    let row_heights = compute_row_heights(ws, min_row, max_row)?;

    let total_width: f64 = col_widths.iter().sum();
    let total_height: f64 = row_heights.iter().sum();
    validate_finite_nonnegative(total_width, "render width must be finite and non-negative")?;
    validate_finite_nonnegative(
        total_height,
        "render height must be finite and non-negative",
    )?;

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
    validate_finite_nonnegative(svg_width, "SVG width must be finite and non-negative")?;
    validate_finite_nonnegative(svg_height, "SVG height must be finite and non-negative")?;

    svg.push_str(&format!(
        r#"<svg xmlns="http://www.w3.org/2000/svg" width="{svg_width}" height="{svg_height}" viewBox="0 0 {} {}" font-family=""#,
        total_width + header_x_offset,
        total_height + header_y_offset,
    ))?;
    svg.push_str(&default_font_family)?;
    svg.push_str(&format!(r#"" font-size="{}">"#, options.default_font_size))?;

    // White background
    svg.push_str(&format!(
        r#"<rect width="{}" height="{}" fill="white"/>"#,
        total_width + header_x_offset,
        total_height + header_y_offset,
    ))?;

    // Render column headers
    if options.show_headers {
        render_column_headers(&mut svg, col_widths, min_col, header_x_offset, options)?;
        render_row_headers(&mut svg, &row_heights, min_row, header_y_offset, options)?;
    }

    let layouts = LayoutGrid {
        col_widths,
        col_styles: &column_metadata.styles,
        row_heights: &row_heights,
        min_col,
        min_row,
        x_offset: header_x_offset,
        y_offset: header_y_offset,
    };

    // Render cell fills
    let mut styles = StyleCache::new(stylesheet);
    render_cell_fills(&mut svg, ws, &mut styles, theme, &layouts)?;

    // Render gridlines
    if options.show_gridlines {
        render_gridlines(
            &mut svg,
            col_widths,
            &row_heights,
            total_width,
            total_height,
            header_x_offset,
            header_y_offset,
        )?;
    }

    // Render cell borders
    render_cell_borders(&mut svg, ws, &mut styles, theme, &layouts, min_col, min_row)?;

    // Render cell text
    render_cell_text(
        &mut svg,
        ws,
        sst,
        &mut styles,
        theme,
        &layouts,
        min_col,
        min_row,
        options,
    )?;

    svg.push_str("</svg>")?;
    Ok(svg.finish())
}

fn render_checked_rect(
    min_col: u32,
    min_row: u32,
    max_col: u32,
    max_row: u32,
    limits: TraversalLimits,
) -> Result<CheckedRect> {
    CheckedRect::new(min_row, max_row, min_col, max_col, limits).map_err(|error| match error {
        Error::ResourceLimitExceeded(metadata) => ResourceLimitError::exceeded(
            ResourceOperation::Render,
            ResourceCategory::Traversal,
            metadata.dimension(),
            metadata.limit(),
            metadata.observed(),
            Some(ResourceStage::Render),
        )
        .map(Error::from)
        .unwrap_or_else(|| Error::Internal("invalid render resource limit".to_string())),
        other => other,
    })
}

/// Determine the range of cells to render.
fn compute_range(
    ws: &WorksheetXml,
    _sst: &SharedStringTable,
    options: &RenderOptions,
    limits: TraversalLimits,
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
        // Rendering computes bounds from cell presence only, so date
        // promotion is not relevant here. Skip the lookup.
        if ws.sheet_data.rows.is_empty() {
            return Ok((1, 1, 1, 1));
        }
        let mut min_col = u32::MAX;
        let mut max_col = 0u32;
        let mut min_row = u32::MAX;
        let mut max_row = 0;
        let mut visited_rows = 0u64;
        let mut visited_cells = 0u64;
        for row in &ws.sheet_data.rows {
            visited_rows = visited_rows.saturating_add(1);
            check_render_visit_limit(ResourceDimension::RowCount, limits.max_rows(), visited_rows)?;
            if row.cells.is_empty() {
                continue;
            }
            min_row = min_row.min(row.r);
            max_row = max_row.max(row.r);
            for cell in &row.cells {
                visited_cells = visited_cells.saturating_add(1);
                check_render_visit_limit(
                    ResourceDimension::CellCount,
                    limits.max_cells(),
                    visited_cells,
                )?;
                let col = if cell.col > 0 {
                    cell.col
                } else {
                    cell_name_to_coordinates(cell.r.as_str())?.0
                };
                min_col = min_col.min(col);
                max_col = max_col.max(col);
            }
        }
        if min_row == u32::MAX {
            return Ok((1, 1, 1, 1));
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

fn check_render_visit_limit(dimension: ResourceDimension, limit: u64, observed: u64) -> Result<()> {
    if observed <= limit {
        return Ok(());
    }
    let metadata = ResourceLimitError::exceeded(
        ResourceOperation::Render,
        ResourceCategory::Traversal,
        dimension,
        limit,
        observed,
        Some(ResourceStage::Render),
    )
    .ok_or_else(|| Error::Internal("invalid render traversal limit".to_string()))?;
    Err(metadata.into())
}

/// Compute pixel heights for each row in the range.
fn compute_row_heights(ws: &WorksheetXml, min_row: u32, max_row: u32) -> Result<Vec<f64>> {
    let count = usize::try_from(u64::from(max_row - min_row) + 1)
        .map_err(|_| Error::Internal("render row count is not addressable".to_string()))?;
    let mut heights = Vec::new();
    heights
        .try_reserve_exact(count)
        .map_err(|_| Error::Internal("render row layout allocation failed".to_string()))?;
    for row_num in min_row..=max_row {
        heights.push(match get_row_height(ws, row_num) {
            Some(h) => row_height_to_px(h),
            None => DEFAULT_ROW_HEIGHT_PX,
        });
    }
    Ok(heights)
}

/// Render column header labels (A, B, C, ...).
fn render_column_headers(
    svg: &mut SvgOutput,
    col_widths: &[f64],
    min_col: u32,
    x_offset: f64,
    _options: &RenderOptions,
) -> Result<()> {
    let total_w: f64 = col_widths.iter().sum();
    svg.push_str(&format!(
        "<rect x=\"{x_offset}\" y=\"0\" width=\"{total_w}\" height=\"{HEADER_HEIGHT}\" fill=\"{HEADER_BG_COLOR}\"/>",
    ))?;

    let mut x = x_offset;
    for (i, &w) in col_widths.iter().enumerate() {
        let col_num = min_col + i as u32;
        let col_name = column_number_to_name(col_num).unwrap_or_default();
        let text_x = x + w / 2.0;
        let text_y = HEADER_HEIGHT / 2.0 + 4.0;
        svg.push_str(&format!(
            "<text x=\"{text_x}\" y=\"{text_y}\" text-anchor=\"middle\" fill=\"{HEADER_TEXT_COLOR}\" font-size=\"10\">{col_name}</text>",
        ))?;
        x += w;
    }
    Ok(())
}

/// Render row header labels (1, 2, 3, ...).
fn render_row_headers(
    svg: &mut SvgOutput,
    row_heights: &[f64],
    min_row: u32,
    y_offset: f64,
    _options: &RenderOptions,
) -> Result<()> {
    let total_h: f64 = row_heights.iter().sum();
    svg.push_str(&format!(
        "<rect x=\"0\" y=\"{y_offset}\" width=\"{HEADER_WIDTH}\" height=\"{total_h}\" fill=\"{HEADER_BG_COLOR}\"/>",
    ))?;

    let mut y = y_offset;
    for (i, &h) in row_heights.iter().enumerate() {
        let row_num = min_row + i as u32;
        let text_x = HEADER_WIDTH / 2.0;
        let text_y = y + h / 2.0 + 4.0;
        svg.push_str(&format!(
            "<text x=\"{text_x}\" y=\"{text_y}\" text-anchor=\"middle\" fill=\"{HEADER_TEXT_COLOR}\" font-size=\"10\">{row_num}</text>",
        ))?;
        y += h;
    }
    Ok(())
}

/// Render cell background fills.
fn render_cell_fills(
    svg: &mut SvgOutput,
    ws: &WorksheetXml,
    styles: &mut StyleCache<'_>,
    theme: &ThemeColors,
    layouts: &LayoutGrid<'_>,
) -> Result<()> {
    let mut scratch = String::new();
    layouts.for_each(|layout| {
        let style_id = find_cell_style(ws, layout.col, layout.row, layout.column_style);
        if style_id == 0 {
            return Ok(());
        }
        if let Some(style) = styles.get(style_id) {
            if let Some(ref fill) = style.fill {
                if fill.pattern == PatternType::Solid {
                    if let Some(ref color) = fill.fg_color {
                        let hex = style_color_to_hex(color, theme)?;
                        svg.push_fmt(
                            &mut scratch,
                            format_args!(
                                r#"<rect x="{}" y="{}" width="{}" height="{}" fill="{}"/>"#,
                                layout.x, layout.y, layout.width, layout.height, hex
                            ),
                        )?;
                    }
                }
            }
        }
        Ok(())
    })
}

/// Render gridlines.
fn render_gridlines(
    svg: &mut SvgOutput,
    col_widths: &[f64],
    row_heights: &[f64],
    total_width: f64,
    total_height: f64,
    x_offset: f64,
    y_offset: f64,
) -> Result<()> {
    let mut y = y_offset;
    for h in row_heights {
        y += h;
        let x2 = x_offset + total_width;
        svg.push_str(&format!(
            "<line x1=\"{x_offset}\" y1=\"{y}\" x2=\"{x2}\" y2=\"{y}\" stroke=\"{GRIDLINE_COLOR}\" stroke-width=\"0.5\"/>",
        ))?;
    }

    let mut x = x_offset;
    for w in col_widths {
        x += w;
        let y2 = y_offset + total_height;
        svg.push_str(&format!(
            "<line x1=\"{x}\" y1=\"{y_offset}\" x2=\"{x}\" y2=\"{y2}\" stroke=\"{GRIDLINE_COLOR}\" stroke-width=\"0.5\"/>",
        ))?;
    }
    Ok(())
}

/// Render cell borders.
fn render_cell_borders(
    svg: &mut SvgOutput,
    ws: &WorksheetXml,
    styles: &mut StyleCache<'_>,
    theme: &ThemeColors,
    layouts: &LayoutGrid<'_>,
    _min_col: u32,
    _min_row: u32,
) -> Result<()> {
    let mut scratch = String::new();
    layouts.for_each(|layout| {
        let style_id = find_cell_style(ws, layout.col, layout.row, layout.column_style);
        if style_id == 0 {
            return Ok(());
        }
        let style = match styles.get(style_id) {
            Some(s) => s,
            None => return Ok(()),
        };
        let border = match &style.border {
            Some(b) => b,
            None => return Ok(()),
        };

        let x1 = layout.x;
        let y1 = layout.y;
        let x2 = layout.x + layout.width;
        let y2 = layout.y + layout.height;

        if let Some(ref left) = border.left {
            let (sw, color) = border_line_attrs(left.style, left.color.as_ref(), theme)?;
            svg.push_fmt(&mut scratch, format_args!(
                r#"<line x1="{x1}" y1="{y1}" x2="{x1}" y2="{y2}" stroke="{color}" stroke-width="{sw}"/>"#,
            ))?;
        }
        if let Some(ref right) = border.right {
            let (sw, color) = border_line_attrs(right.style, right.color.as_ref(), theme)?;
            svg.push_fmt(&mut scratch, format_args!(
                r#"<line x1="{x2}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{color}" stroke-width="{sw}"/>"#,
            ))?;
        }
        if let Some(ref top) = border.top {
            let (sw, color) = border_line_attrs(top.style, top.color.as_ref(), theme)?;
            svg.push_fmt(&mut scratch, format_args!(
                r#"<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y1}" stroke="{color}" stroke-width="{sw}"/>"#,
            ))?;
        }
        if let Some(ref bottom) = border.bottom {
            let (sw, color) = border_line_attrs(bottom.style, bottom.color.as_ref(), theme)?;
            svg.push_fmt(&mut scratch, format_args!(
                r#"<line x1="{x1}" y1="{y2}" x2="{x2}" y2="{y2}" stroke="{color}" stroke-width="{sw}"/>"#,
            ))?;
        }
        Ok(())
    })
}

/// Render cell text values.
#[allow(clippy::too_many_arguments)]
fn render_cell_text(
    svg: &mut SvgOutput,
    ws: &WorksheetXml,
    sst: &SharedStringTable,
    styles: &mut StyleCache<'_>,
    theme: &ThemeColors,
    layouts: &LayoutGrid<'_>,
    _min_col: u32,
    _min_row: u32,
    options: &RenderOptions,
) -> Result<()> {
    let mut scratch = String::new();
    let mut staged = String::new();
    // Escaped font-family attribute values keyed by font name.
    let mut font_families: HashMap<String, String> = HashMap::new();
    layouts.for_each(|layout| {
        let Some(cell) = find_cell(ws, layout.col, layout.row) else {
            return Ok(());
        };
        let cell_text = CellText::from_cell(cell, sst);
        if cell_text.is_empty() {
            return Ok(());
        }

        let style_id = find_cell_style(ws, layout.col, layout.row, layout.column_style);
        let style = styles.get(style_id);

        let font = style.and_then(|s| s.font.as_ref());
        let alignment = style.and_then(|s| s.alignment.as_ref());

        let (text_x, anchor) = compute_text_x(layout, alignment);
        let text_y = compute_text_y(layout, alignment, font, options)?;

        let mut attrs = svg.stage_with(std::mem::take(&mut staged));
        attrs.push_fmt(&mut scratch, format_args!(r#" x="{text_x}" y="{text_y}""#))?;
        attrs.push_fmt(&mut scratch, format_args!(r#" text-anchor="{anchor}""#))?;

        if let Some(f) = font {
            if f.bold {
                attrs.push_str(r#" font-weight="bold""#)?;
            }
            if f.italic {
                attrs.push_str(r#" font-style="italic""#)?;
            }
            if let Some(ref name) = f.name {
                attrs.push_str(r#" font-family=""#)?;
                // A cached value that fits cannot hit any intermediate limit
                // in css_font_family_attribute (its staged lengths never
                // exceed the final escaped length); otherwise recompute so
                // the outcome, including any error, is unchanged.
                match font_families.get(name.as_str()) {
                    Some(escaped) if attrs.fits(escaped.len()) => attrs.push_str(escaped)?,
                    _ => {
                        let escaped = css_font_family_attribute(name, &attrs)?;
                        attrs.push_str(&escaped)?;
                        font_families.insert(name.clone(), escaped);
                    }
                }
                attrs.push_str("\"")?;
            }
            if let Some(size) = f.size {
                validate_positive_finite(size, "font size must be positive and finite")?;
                attrs.push_fmt(&mut scratch, format_args!(r#" font-size="{size}""#))?;
            }
            if let Some(ref color) = f.color {
                let hex = style_color_to_hex(color, theme)?;
                attrs.push_fmt(&mut scratch, format_args!(r#" fill="{hex}""#))?;
            }
            let mut decorations = Vec::new();
            if f.underline {
                decorations.push("underline");
            }
            if f.strikethrough {
                decorations.push("line-through");
            }
            if !decorations.is_empty() {
                attrs.push_str(&format!(r#" text-decoration="{}""#, decorations.join(" ")))?;
            }
        }

        svg.push_str("<text")?;
        let attrs = attrs.finish();
        svg.push_str(&attrs)?;
        staged = attrs;
        svg.push_str(">")?;
        cell_text.write_escaped(svg)?;
        svg.push_str("</text>")?;
        Ok(())
    })
}

enum CellText<'a> {
    Empty,
    Borrowed(&'a str),
    Inline(&'a InlineString),
    Formula(&'a str),
    Static(&'static str),
    Number(f64),
}

impl<'a> CellText<'a> {
    fn from_cell(cell: &'a Cell, sst: &'a SharedStringTable) -> Self {
        if let Some(formula) = cell.f.as_deref() {
            let expression = formula.value.as_deref().unwrap_or("");
            return match (cell.t, cell.v.as_deref()) {
                (CellTypeTag::Boolean, Some(value)) => {
                    Self::Static(if value == "1" { "TRUE" } else { "FALSE" })
                }
                (CellTypeTag::Error | CellTypeTag::FormulaString, Some(value)) => {
                    Self::Borrowed(value)
                }
                (CellTypeTag::SharedString, Some(value)) => sst_text(sst, value),
                (CellTypeTag::InlineString, _) => {
                    cell.is.as_deref().map(Self::Inline).unwrap_or(Self::Empty)
                }
                (CellTypeTag::None | CellTypeTag::Number | CellTypeTag::Date, Some(value)) => value
                    .parse::<f64>()
                    .map(Self::Number)
                    .unwrap_or(Self::Formula(expression)),
                _ => Self::Formula(expression),
            };
        }

        match (cell.t, cell.v.as_deref()) {
            (CellTypeTag::SharedString, Some(value)) => sst_text(sst, value),
            (CellTypeTag::Boolean, Some(value)) => {
                Self::Static(if value == "1" { "TRUE" } else { "FALSE" })
            }
            (CellTypeTag::Error | CellTypeTag::FormulaString, Some(value)) => Self::Borrowed(value),
            (CellTypeTag::InlineString, _) => {
                cell.is.as_deref().map(Self::Inline).unwrap_or(Self::Empty)
            }
            (CellTypeTag::None | CellTypeTag::Number, Some(value)) => value
                .parse::<f64>()
                .map(Self::Number)
                .unwrap_or(Self::Empty),
            _ => Self::Empty,
        }
    }

    fn is_empty(&self) -> bool {
        match self {
            Self::Empty => true,
            Self::Borrowed(value) | Self::Static(value) => value.is_empty(),
            Self::Inline(value) => {
                value.t.as_deref().is_none_or(str::is_empty)
                    && value.r.iter().all(|run| run.t.value.is_empty())
            }
            Self::Formula(_) | Self::Number(_) => false,
        }
    }

    fn write_escaped(&self, output: &mut SvgOutput) -> Result<()> {
        match self {
            Self::Empty => Ok(()),
            Self::Borrowed(value) | Self::Static(value) => push_xml_text(output, value),
            Self::Inline(value) => {
                let escaped_len = value
                    .t
                    .as_deref()
                    .map(xml_text_escaped_len)
                    .transpose()?
                    .unwrap_or(0);
                let escaped_len = value.r.iter().try_fold(escaped_len, |total, run| {
                    Ok::<_, Error>(total.saturating_add(xml_text_escaped_len(&run.t.value)?))
                })?;
                output.ensure_additional(escaped_len)?;
                if let Some(text) = value.t.as_deref() {
                    push_xml_text_unchecked(output, text)?;
                }
                for run in &value.r {
                    push_xml_text_unchecked(output, &run.t.value)?;
                }
                Ok(())
            }
            Self::Formula(expression) => {
                let escaped_len = 1u64.saturating_add(xml_text_escaped_len(expression)?);
                output.ensure_additional(escaped_len)?;
                output.push_str("=")?;
                push_xml_text_unchecked(output, expression)
            }
            Self::Number(number) => push_xml_text(output, &display_number(*number)),
        }
    }
}

fn sst_text<'a>(sst: &'a SharedStringTable, index: &str) -> CellText<'a> {
    index
        .parse::<usize>()
        .ok()
        .and_then(|index| sst.get(index))
        .map(CellText::Borrowed)
        .unwrap_or(CellText::Empty)
}

fn display_number(number: f64) -> String {
    if number.fract() == 0.0 && number.is_finite() {
        (number as i64).to_string()
    } else {
        number.to_string()
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
) -> Result<f64> {
    let font_size = font
        .and_then(|f| f.size)
        .unwrap_or(options.default_font_size);
    validate_positive_finite(font_size, "font size must be positive and finite")?;
    Ok(match alignment.and_then(|a| a.vertical) {
        Some(VerticalAlign::Top) => layout.y + font_size + 2.0,
        Some(VerticalAlign::Center) => layout.y + layout.height / 2.0 + font_size / 3.0,
        _ => layout.y + layout.height - 4.0,
    })
}

/// Find the style ID for a cell at the given coordinates.
fn find_cell_style(ws: &WorksheetXml, col: u32, row: u32, column_style: u32) -> u32 {
    let row_data = ws
        .sheet_data
        .rows
        .binary_search_by_key(&row, |r| r.r)
        .ok()
        .map(|idx| &ws.sheet_data.rows[idx]);

    if let Some(row_data) = row_data {
        if let Some(cell) = row_data
            .cells
            .binary_search_by_key(&col, |c| c.col)
            .ok()
            .map(|idx| &row_data.cells[idx])
        {
            if let Some(style) = cell.s {
                return style;
            }
        }
        if let Some(style) = row_data.s {
            return style;
        }
    }

    column_style
}

/// Find a cell at the given coordinates without cloning its value.
fn find_cell(ws: &WorksheetXml, col: u32, row: u32) -> Option<&Cell> {
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
                .map(|ci| &row_data.cells[ci])
        })
}

/// Convert a StyleColor to a validated SVG hex color string.
fn style_color_to_hex(color: &StyleColor, theme: &ThemeColors) -> Result<String> {
    match color {
        StyleColor::Rgb(rgb) => rgb_to_svg_hex(rgb),
        StyleColor::Theme(index) => match resolve_theme_color(theme, *index, None) {
            Some(color) => rgb_to_svg_hex(&color),
            None => Ok("#000000".to_string()),
        },
        StyleColor::Indexed(_) => Ok("#000000".to_string()),
    }
}

fn rgb_to_svg_hex(rgb: &str) -> Result<String> {
    let digits = rgb.strip_prefix('#').unwrap_or(rgb);
    if !matches!(digits.len(), 6 | 8) || !digits.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(Error::InvalidArgument("invalid SVG color".to_string()));
    }

    let digits = if digits.len() == 8 {
        digits
            .get(2..)
            .ok_or_else(|| Error::InvalidArgument("invalid SVG color".to_string()))?
    } else {
        digits
    };
    Ok(format!("#{digits}"))
}

/// Convert a border line style to SVG stroke-width and color.
fn border_line_attrs(
    style: BorderLineStyle,
    color: Option<&StyleColor>,
    theme: &ThemeColors,
) -> Result<(f64, String)> {
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
        .map(|color| style_color_to_hex(color, theme))
        .transpose()?
        .unwrap_or_else(|| "#000000".to_string());
    Ok((stroke_width, color_str))
}

fn validate_positive_finite(value: f64, message: &'static str) -> Result<()> {
    if !value.is_finite() || value <= 0.0 {
        return Err(Error::InvalidArgument(message.to_string()));
    }
    Ok(())
}

fn validate_finite_nonnegative(value: f64, message: &'static str) -> Result<()> {
    if !value.is_finite() || value < 0.0 {
        return Err(Error::InvalidArgument(message.to_string()));
    }
    Ok(())
}

fn is_xml_char(c: char) -> bool {
    matches!(c, '\u{9}' | '\u{A}' | '\u{D}')
        || ('\u{20}'..='\u{D7FF}').contains(&c)
        || ('\u{E000}'..='\u{FFFD}').contains(&c)
        || ('\u{10000}'..='\u{10FFFF}').contains(&c)
}

fn validate_xml_chars(value: &str, message: &'static str) -> Result<()> {
    if value.chars().any(|c| !is_xml_char(c)) {
        return Err(Error::InvalidArgument(message.to_string()));
    }
    Ok(())
}

/// Escape XML text while preserving printable Unicode.
#[cfg(test)]
fn xml_text_escape(s: &str) -> Result<String> {
    let mut out = SvgOutput::new(MAX_SVG_OUTPUT_BYTES);
    push_xml_text(&mut out, s)?;
    Ok(out.finish())
}

fn push_xml_text(output: &mut SvgOutput, value: &str) -> Result<()> {
    output.ensure_additional(xml_text_escaped_len(value)?)?;
    push_xml_text_unchecked(output, value)
}

/// Validate XML characters and return the escaped text length in one pass.
fn xml_text_escaped_len(value: &str) -> Result<u64> {
    let mut length = 0u64;
    for c in value.chars() {
        if !is_xml_char(c) {
            return Err(Error::InvalidArgument(
                "SVG text contains an XML-forbidden character".to_string(),
            ));
        }
        length = length.saturating_add(match c {
            '&' => 5,
            '<' | '>' => 4,
            _ => u64::from(c.len_utf8() as u32),
        });
    }
    Ok(length)
}

fn xml_text_escape_char(c: char) -> Option<&'static str> {
    match c {
        '&' => Some("&amp;"),
        '<' => Some("&lt;"),
        '>' => Some("&gt;"),
        _ => None,
    }
}

fn push_xml_text_unchecked(out: &mut SvgOutput, s: &str) -> Result<()> {
    out.push_escaped(s, xml_text_escape_char)
}

/// Escape one XML attribute after its CSS value has been encoded.
fn xml_attribute_escape(s: &str, parent: &SvgOutput) -> Result<String> {
    validate_xml_chars(s, "SVG attribute contains an XML-forbidden character")?;
    let mut out = parent.stage();
    out.push_escaped(s, |c| match c {
        '&' => Some("&amp;"),
        '<' => Some("&lt;"),
        '>' => Some("&gt;"),
        '"' => Some("&quot;"),
        '\'' => Some("&apos;"),
        _ => None,
    })?;
    Ok(out.finish())
}

/// Encode a font family as one CSS string, then as one XML attribute value.
fn css_font_family_attribute(value: &str, parent: &SvgOutput) -> Result<String> {
    validate_xml_chars(value, "SVG font family contains an XML-forbidden character")?;
    let mut css = parent.stage();
    css.push_str("\"")?;
    css.push_escaped(value, |c| match c {
        '"' => Some("\\\""),
        '\\' => Some("\\\\"),
        '\n' => Some("\\A "),
        '\r' => Some("\\D "),
        '\t' => Some("\\9 "),
        '<' => Some("\\3C "),
        '>' => Some("\\3E "),
        '&' => Some("\\26 "),
        _ => None,
    })?;
    css.push_str("\"")?;
    xml_attribute_escape(&css.finish(), parent)
}

#[cfg(test)]
#[allow(clippy::field_reassign_with_default)]
mod tests {
    use super::*;
    use crate::sst::SharedStringTable;
    use crate::style::{add_style, StyleBuilder};
    use sheetkit_xml::styles::StyleSheet;
    use sheetkit_xml::worksheet::{Cell, CellTypeTag, Col, Cols, Row, SheetData, WorksheetXml};

    fn default_options(sheet: &str) -> RenderOptions {
        RenderOptions {
            sheet_name: sheet.to_string(),
            ..RenderOptions::default()
        }
    }

    fn make_num_cell(r: &str, col: u32, v: &str) -> Cell {
        Cell {
            r: r.try_into().unwrap(),
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
            r: r.try_into().unwrap(),
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

    /// Per-character reference for `SvgOutput::push_escaped`.
    fn push_escaped_per_char(
        out: &mut SvgOutput,
        value: &str,
        escape: impl Fn(char) -> Option<&'static str>,
    ) -> Result<()> {
        for c in value.chars() {
            match escape(c) {
                Some(replacement) => out.push_str(replacement)?,
                None => {
                    let mut encoded = [0; 4];
                    out.push_str(c.encode_utf8(&mut encoded))?;
                }
            }
        }
        Ok(())
    }

    fn outcome(result: Result<()>, out: SvgOutput) -> std::result::Result<String, (u64, u64)> {
        match result {
            Ok(()) => Ok(out.finish()),
            Err(error) => {
                let metadata = error.resource_metadata().unwrap();
                Err((metadata.limit(), metadata.observed()))
            }
        }
    }

    #[test]
    fn push_escaped_matches_per_char_output_and_limit_errors() {
        let value = "ab&c<d>e\"f'g\u{e9}\u{d55c}\u{1F600}hij";
        for limit in 0..48 {
            for prefix in [0usize, 3] {
                let mut fast = SvgOutput::new(limit);
                let mut slow = SvgOutput::new(limit);
                let prefix = "x".repeat(prefix);
                let fast_prefix = fast.push_str(&prefix);
                let slow_prefix = slow.push_str(&prefix);
                assert_eq!(fast_prefix.is_ok(), slow_prefix.is_ok());
                if fast_prefix.is_err() {
                    continue;
                }
                let fast_result = fast.push_escaped(value, xml_text_escape_char);
                let slow_result = push_escaped_per_char(&mut slow, value, xml_text_escape_char);
                assert_eq!(
                    outcome(fast_result, fast),
                    outcome(slow_result, slow),
                    "limit {limit}"
                );
            }
        }
    }

    #[test]
    fn repeated_font_family_matches_first_rendering() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();
        let style = add_style(
            &mut ss,
            &StyleBuilder::new()
                .font_name("My \"Font\" & <Co>")
                .font_size(12.0)
                .build(),
        )
        .unwrap();
        for row in &mut ws.sheet_data.rows {
            for cell in &mut row.cells {
                cell.s = Some(style);
            }
        }
        let svg = render_to_svg(&ws, &sst, &ss, &default_options("Sheet1")).unwrap();
        let expected = r#" font-family="&quot;My \&quot;Font\&quot; \26  \3C Co\3E &quot;""#;
        assert_eq!(svg.matches(expected).count(), 4, "{svg}");
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
    fn render_dense_area_accepts_exact_limit_and_rejects_plus_one() {
        let (ws, sst) = simple_ws_and_sst();
        let stylesheet = StyleSheet::default();
        let mut options = default_options("Sheet1");
        options.range = Some("A1:B2".to_string());
        let exact = TraversalLimits::new(2, 2, 4, 1024 * 1024).unwrap();
        render_to_svg_with_limits(&ws, &sst, &stylesheet, &options, exact).unwrap();

        let plus_one = TraversalLimits::new(2, 2, 3, 1024 * 1024).unwrap();
        let error =
            render_to_svg_with_limits(&ws, &sst, &stylesheet, &options, plus_one).unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.operation(), ResourceOperation::Render);
        assert_eq!(metadata.dimension(), ResourceDimension::CellCount);
        assert_eq!(metadata.limit(), 3);
        assert_eq!(metadata.observed(), 4);
    }

    #[test]
    fn range_discovery_charges_duplicate_cell_visits() {
        let mut ws = WorksheetXml::default();
        ws.sheet_data.rows = vec![Row {
            r: 1,
            spans: None,
            s: None,
            custom_format: None,
            ht: None,
            hidden: None,
            custom_height: None,
            outline_level: None,
            cells: vec![make_num_cell("A1", 1, "1"), make_num_cell("A1", 1, "2")],
        }];
        let options = default_options("Sheet1");
        let exact = TraversalLimits::new(1, 1, 2, 1024 * 1024).unwrap();
        render_to_svg_with_limits(
            &ws,
            &SharedStringTable::new(),
            &StyleSheet::default(),
            &options,
            exact,
        )
        .unwrap();

        let plus_one = TraversalLimits::new(1, 1, 1, 1024 * 1024).unwrap();
        let error = render_to_svg_with_limits(
            &ws,
            &SharedStringTable::new(),
            &StyleSheet::default(),
            &options,
            plus_one,
        )
        .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.dimension(), ResourceDimension::CellCount);
        assert_eq!(metadata.limit(), 1);
        assert_eq!(metadata.observed(), 2);
    }

    #[test]
    fn range_discovery_charges_duplicate_row_visits() {
        let mut ws = WorksheetXml::default();
        let duplicate_row = Row {
            r: 1,
            spans: None,
            s: None,
            custom_format: None,
            ht: None,
            hidden: None,
            custom_height: None,
            outline_level: None,
            cells: vec![make_num_cell("A1", 1, "1")],
        };
        ws.sheet_data.rows = vec![duplicate_row.clone(), duplicate_row];
        let options = default_options("Sheet1");
        let exact = TraversalLimits::new(2, 1, 2, 1024 * 1024).unwrap();
        render_to_svg_with_limits(
            &ws,
            &SharedStringTable::new(),
            &StyleSheet::default(),
            &options,
            exact,
        )
        .unwrap();

        let plus_one = TraversalLimits::new(1, 1, 2, 1024 * 1024).unwrap();
        let error = render_to_svg_with_limits(
            &ws,
            &SharedStringTable::new(),
            &StyleSheet::default(),
            &options,
            plus_one,
        )
        .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.dimension(), ResourceDimension::RowCount);
        assert_eq!(metadata.limit(), 1);
        assert_eq!(metadata.observed(), 2);
    }

    #[test]
    fn bounded_text_escape_rejects_before_growing_output() {
        let text = "<&>".repeat(1_000_000);
        let mut output = SvgOutput::new(16);

        let error = push_xml_text(&mut output, &text).unwrap_err();

        assert!(matches!(error, Error::ResourceLimitExceeded(_)));
        assert_eq!(output.value.capacity(), 0);
    }

    #[test]
    fn font_family_helper_uses_the_caller_output_limit() {
        let font_family = "<&>".repeat(1_000_000);
        let output = SvgOutput::new(16);

        let error = css_font_family_attribute(&font_family, &output).unwrap_err();

        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
        assert_eq!(metadata.limit(), 16);
    }

    #[test]
    fn render_admits_far_rows_before_layout_allocation() {
        let mut ws = WorksheetXml::default();
        ws.sheet_data.rows = vec![
            Row {
                r: 1,
                spans: None,
                s: None,
                custom_format: None,
                ht: None,
                hidden: None,
                custom_height: None,
                outline_level: None,
                cells: vec![make_num_cell("A1", 1, "1")],
            },
            Row {
                r: 1_048_576,
                spans: None,
                s: None,
                custom_format: None,
                ht: None,
                hidden: None,
                custom_height: None,
                outline_level: None,
                cells: vec![make_num_cell("A1048576", 1, "2")],
            },
        ];
        let options = default_options("Sheet1");
        let limits = TraversalLimits::new(1_048_576, 1, 2, 1024).unwrap();
        let error = render_to_svg_with_limits(
            &ws,
            &SharedStringTable::new(),
            &StyleSheet::default(),
            &options,
            limits,
        )
        .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.dimension(), ResourceDimension::CellCount);
        assert_eq!(metadata.observed(), 1_048_576);
    }

    #[test]
    fn explicit_default_limits_preserve_svg_output() {
        let (ws, sst) = simple_ws_and_sst();
        let stylesheet = StyleSheet::default();
        let options = default_options("Sheet1");
        let default_svg = render_to_svg(&ws, &sst, &stylesheet, &options).unwrap();
        let explicit_svg =
            render_to_svg_with_limits(&ws, &sst, &stylesheet, &options, TraversalLimits::default())
                .unwrap();
        assert_eq!(default_svg, explicit_svg);

        let exact = TraversalLimits::new(2, 2, 4, default_svg.len() as u64).unwrap();
        assert_eq!(
            render_to_svg_with_limits(&ws, &sst, &stylesheet, &options, exact).unwrap(),
            default_svg
        );
        let too_small = TraversalLimits::new(2, 2, 4, default_svg.len() as u64 - 1).unwrap();
        let error =
            render_to_svg_with_limits(&ws, &sst, &stylesheet, &options, too_small).unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
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
    fn column_metadata_index_resolves_overlaps_once_with_first_definition_winning() {
        let mut ws = WorksheetXml::default();
        let definition_count = 4_096usize;
        let visible_columns = 64usize;
        let mut definitions = Vec::with_capacity(definition_count);
        for index in 0..definition_count {
            definitions.push(Col {
                min: 1,
                max: visible_columns as u32,
                width: Some(10.0 + index as f64),
                style: Some(index as u32),
                hidden: None,
                custom_width: Some(true),
                outline_level: None,
            });
        }
        ws.cols = Some(Cols { cols: definitions });

        let (metadata, resolution_work) =
            ColumnMetadataIndex::build(&ws, 1, visible_columns as u32).unwrap();

        assert_eq!(resolution_work, definition_count + visible_columns);
        assert_eq!(metadata.widths.len(), visible_columns);
        assert_eq!(metadata.styles.len(), visible_columns);
        assert!(metadata
            .widths
            .iter()
            .all(|width| *width == col_width_to_px(10.0)));
        assert!(metadata.styles.iter().all(|style| *style == 0));

        let mut options = default_options("Sheet1");
        let last_column = column_number_to_name(visible_columns as u32).unwrap();
        options.range = Some(format!("A1:{last_column}64"));
        let svg = render_to_svg(
            &ws,
            &SharedStringTable::new(),
            &StyleSheet::default(),
            &options,
        )
        .unwrap();
        assert!(svg.ends_with("</svg>"));
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
    fn test_render_applies_row_style_when_cell_has_no_style() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();
        let style_id =
            add_style(&mut ss, &StyleBuilder::new().solid_fill("FFFF0000").build()).unwrap();
        ws.sheet_data.rows[0].s = Some(style_id);

        let svg = render_to_svg(&ws, &sst, &ss, &default_options("Sheet1")).unwrap();

        assert!(svg.contains("fill=\"#FF0000\""));
    }

    #[test]
    fn test_render_applies_column_style_when_cell_and_row_have_no_style() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();
        let style_id =
            add_style(&mut ss, &StyleBuilder::new().solid_fill("FF00FF00").build()).unwrap();
        ws.cols = Some(Cols {
            cols: vec![Col {
                min: 1,
                max: 1,
                width: None,
                style: Some(style_id),
                hidden: None,
                custom_width: None,
                outline_level: None,
            }],
        });

        let svg = render_to_svg(&ws, &sst, &ss, &default_options("Sheet1")).unwrap();

        assert!(svg.contains("fill=\"#00FF00\""));
    }

    #[test]
    fn test_cell_style_precedence_includes_explicit_default_style() {
        let (mut ws, _sst) = simple_ws_and_sst();
        ws.sheet_data.rows[0].s = Some(2);
        ws.sheet_data.rows[0].cells[0].s = Some(0);
        ws.cols = Some(Cols {
            cols: vec![Col {
                min: 1,
                max: 2,
                width: None,
                style: Some(3),
                hidden: None,
                custom_width: None,
                outline_level: None,
            }],
        });

        assert_eq!(find_cell_style(&ws, 1, 1, 3), 0);
        assert_eq!(find_cell_style(&ws, 2, 1, 3), 2);
        assert_eq!(find_cell_style(&ws, 1, 3, 3), 3);
    }

    #[test]
    fn test_render_resolves_theme_colors_with_default_theme() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();
        let style_id = add_style(
            &mut ss,
            &StyleBuilder::new()
                .fill_pattern(PatternType::Solid)
                .fill_fg_color(StyleColor::Theme(4))
                .font_color(StyleColor::Theme(5))
                .border_all(BorderLineStyle::Thin, StyleColor::Theme(6))
                .build(),
        )
        .unwrap();
        ws.sheet_data.rows[0].cells[0].s = Some(style_id);

        let svg = render_to_svg(&ws, &sst, &ss, &default_options("Sheet1")).unwrap();

        assert!(svg.contains("fill=\"#4472C4\""));
        assert!(svg.contains("fill=\"#ED7D31\""));
        assert!(svg.contains("stroke=\"#A5A5A5\""));
    }

    #[test]
    fn test_render_resolves_theme_colors_with_supplied_theme() {
        let (mut ws, sst) = simple_ws_and_sst();
        let mut ss = StyleSheet::default();
        let style_id = add_style(
            &mut ss,
            &StyleBuilder::new()
                .fill_pattern(PatternType::Solid)
                .fill_fg_color(StyleColor::Theme(4))
                .font_color(StyleColor::Theme(5))
                .border_all(BorderLineStyle::Thin, StyleColor::Theme(6))
                .build(),
        )
        .unwrap();
        ws.sheet_data.rows[0].cells[0].s = Some(style_id);
        let mut theme = crate::theme::default_theme_colors();
        theme.colors[4] = "FF112233".to_string();
        theme.colors[5] = "FF445566".to_string();
        theme.colors[6] = "FF778899".to_string();

        let svg =
            render_to_svg_with_theme(&ws, &sst, &ss, &theme, &default_options("Sheet1")).unwrap();

        assert!(svg.contains("fill=\"#112233\""));
        assert!(svg.contains("fill=\"#445566\""));
        assert!(svg.contains("stroke=\"#778899\""));
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
                        r: "A1".try_into().unwrap(),
                        col: 1,
                        s: None,
                        t: CellTypeTag::SharedString,
                        v: Some("0".to_string()),
                        f: None,
                        is: None,
                    },
                    Cell {
                        r: "B1".try_into().unwrap(),
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
        assert_eq!(xml_text_escape("a&b").unwrap(), "a&amp;b");
        assert_eq!(xml_text_escape("a<b").unwrap(), "a&lt;b");
        assert_eq!(xml_text_escape("a>b").unwrap(), "a&gt;b");
        assert_eq!(xml_text_escape("a\"b").unwrap(), "a\"b");
        assert_eq!(xml_text_escape("a'b").unwrap(), "a'b");
        assert_eq!(xml_text_escape("normal").unwrap(), "normal");
    }

    #[test]
    fn test_style_color_to_hex_argb() {
        let color = StyleColor::Rgb("FFFF0000".to_string());
        assert_eq!(
            style_color_to_hex(&color, &default_theme_colors()).unwrap(),
            "#FF0000"
        );
    }

    #[test]
    fn test_style_color_to_hex_rgb() {
        let color = StyleColor::Rgb("00FF00".to_string());
        assert_eq!(
            style_color_to_hex(&color, &default_theme_colors()).unwrap(),
            "#00FF00"
        );
    }

    #[test]
    fn test_style_color_to_hex_theme_uses_default_theme() {
        let color = StyleColor::Theme(4);
        assert_eq!(
            style_color_to_hex(&color, &default_theme_colors()).unwrap(),
            "#4472C4"
        );
    }

    #[test]
    fn test_border_line_attrs_thin() {
        let (sw, color) =
            border_line_attrs(BorderLineStyle::Thin, None, &default_theme_colors()).unwrap();
        assert_eq!(sw, 1.0);
        assert_eq!(color, "#000000");
    }

    #[test]
    fn test_border_line_attrs_thick_with_color() {
        let c = StyleColor::Rgb("FF0000FF".to_string());
        let (sw, color) =
            border_line_attrs(BorderLineStyle::Thick, Some(&c), &default_theme_colors()).unwrap();
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
        assert_eq!(
            style_color_to_hex(&color, &default_theme_colors()).unwrap(),
            "#FF0000"
        );
    }

    #[test]
    fn test_style_color_to_hex_prefixed_argb() {
        let color = StyleColor::Rgb("#FFFF0000".to_string());
        assert_eq!(
            style_color_to_hex(&color, &default_theme_colors()).unwrap(),
            "#FF0000"
        );
    }

    #[test]
    fn test_style_color_to_hex_no_double_hash() {
        let color = StyleColor::Rgb("#00FF00".to_string());
        let hex = style_color_to_hex(&color, &default_theme_colors()).unwrap();
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

    #[test]
    fn svg_output_limit_accepts_exact_bytes_and_rejects_plus_one() {
        let mut exact = SvgOutput::new(3);
        exact.push_str("abc").unwrap();
        assert_eq!(exact.finish(), "abc");

        let mut over = SvgOutput::new(3);
        over.push_str("abc").unwrap();
        let error = over.push_str("d").unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.operation(), ResourceOperation::Render);
        assert_eq!(metadata.category(), ResourceCategory::Output);
        assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
        assert_eq!(metadata.limit(), 3);
        assert_eq!(metadata.observed(), 4);
        assert_eq!(metadata.stage(), Some(ResourceStage::Render));
        assert_eq!(over.finish(), "abc");
    }
}
