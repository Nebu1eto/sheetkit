//! VML (Vector Markup Language) support for Excel legacy comment rendering.
//!
//! Excel uses VML drawing parts (`xl/drawings/vmlDrawingN.vml`) to render
//! comment/note pop-up boxes in the UI. This module generates minimal VML
//! markup for new comments and tracks preserved VML bytes for round-tripping.

use crate::utils::cell_ref::cell_name_to_coordinates;

/// Default comment box width in columns (roughly 2 columns).
const DEFAULT_COMMENT_WIDTH_COLS: u32 = 2;
/// Default comment box height in rows (roughly 4 rows).
const DEFAULT_COMMENT_HEIGHT_ROWS: u32 = 4;

/// Build a complete VML drawing document containing shapes for each comment cell.
///
/// `cells` is a list of cell references (e.g. `["A1", "B3"]`).
/// Returns the VML XML string.
pub fn build_vml_drawing(cells: &[&str]) -> String {
    let mut shapes = String::new();
    for (i, cell) in cells.iter().enumerate() {
        let shape_id = 1025 + i;
        if let Ok((col, row)) = cell_name_to_coordinates(cell) {
            let anchor = comment_anchor(col, row);
            let row_0 = row - 1;
            let col_0 = col - 1;
            let z = i + 1;
            write_vml_shape(&mut shapes, shape_id, z, &anchor, row_0, col_0);
        }
    }

    let mut doc = String::with_capacity(1024 + shapes.len());
    doc.push_str("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"");
    doc.push_str(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
    doc.push_str(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n");
    doc.push_str(" <o:shapelayout v:ext=\"edit\">\n");
    doc.push_str("  <o:idmap v:ext=\"edit\" data=\"1\"/>\n");
    doc.push_str(" </o:shapelayout>\n");
    doc.push_str(" <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\"");
    doc.push_str(" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">\n");
    doc.push_str("  <v:stroke joinstyle=\"miter\"/>\n");
    doc.push_str("  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n");
    doc.push_str(" </v:shapetype>\n");
    doc.push_str(&shapes);
    doc.push_str("</xml>\n");
    doc
}

/// Write a single VML shape element for a comment to the output string.
fn write_vml_shape(
    out: &mut String,
    shape_id: usize,
    z_index: usize,
    anchor: &str,
    row_0: u32,
    col_0: u32,
) {
    use std::fmt::Write;
    let _ = write!(out, " <v:shape id=\"_x0000_s{}\"", shape_id);
    out.push_str(" type=\"#_x0000_t202\"");
    let _ = write!(
        out,
        " style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;\
         width:108pt;height:59.25pt;z-index:{};visibility:hidden\"",
        z_index
    );
    out.push_str(" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\n");
    out.push_str("  <v:fill color2=\"#ffffe1\"/>\n");
    out.push_str("  <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>\n");
    out.push_str("  <v:path o:connecttype=\"none\"/>\n");
    out.push_str("  <v:textbox/>\n");
    out.push_str("  <x:ClientData ObjectType=\"Note\">\n");
    out.push_str("   <x:MoveWithCells/>\n");
    out.push_str("   <x:SizeWithCells/>\n");
    let _ = writeln!(out, "   <x:Anchor>{}</x:Anchor>", anchor);
    let _ = writeln!(out, "   <x:Row>{}</x:Row>", row_0);
    let _ = writeln!(out, "   <x:Column>{}</x:Column>", col_0);
    out.push_str("  </x:ClientData>\n");
    out.push_str(" </v:shape>\n");
}

/// Compute the 8-value anchor string for a comment box near a cell.
///
/// Format: "LeftCol, LeftOff, TopRow, TopOff, RightCol, RightOff, BottomRow, BottomOff"
/// Offsets are in Excel internal units (EMU/15). We use zero offsets for simplicity.
fn comment_anchor(col: u32, row: u32) -> String {
    let left_col = col;
    let top_row = if row > 1 { row - 2 } else { 0 };
    let right_col = col + DEFAULT_COMMENT_WIDTH_COLS;
    let bottom_row = top_row + DEFAULT_COMMENT_HEIGHT_ROWS;
    format!("{left_col}, 15, {top_row}, 10, {right_col}, 15, {bottom_row}, 4")
}

/// Extract comment cell references from an existing VML drawing XML string.
///
/// Scans for `<x:Row>` and `<x:Column>` elements in the VML and returns
/// (row_0based, col_0based) pairs.
pub fn extract_vml_comment_cells(vml_xml: &str) -> Vec<(u32, u32)> {
    let mut cells = Vec::new();
    let mut current_row: Option<u32> = None;
    let mut current_col: Option<u32> = None;

    for line in vml_xml.lines() {
        let trimmed = line.trim();
        if let Some(val) = extract_element_value(trimmed, "x:Row") {
            current_row = val.parse().ok();
        }
        if let Some(val) = extract_element_value(trimmed, "x:Column") {
            current_col = val.parse().ok();
        }
        if current_row.is_some() && current_col.is_some() {
            cells.push((current_row.unwrap(), current_col.unwrap()));
            current_row = None;
            current_col = None;
        }
    }
    cells
}

/// Simple extraction of text content from an XML element like `<tag>value</tag>`.
fn extract_element_value<'a>(line: &'a str, tag: &str) -> Option<&'a str> {
    let open = format!("<{}>", tag);
    let close = format!("</{}>", tag);
    if let Some(start) = line.find(&open) {
        let val_start = start + open.len();
        if let Some(end) = line.find(&close) {
            return Some(&line[val_start..end]);
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_build_vml_drawing_single_cell() {
        let vml = build_vml_drawing(&["A1"]);
        assert!(vml.contains("xmlns:v=\"urn:schemas-microsoft-com:vml\""));
        assert!(vml.contains("<x:Row>0</x:Row>"));
        assert!(vml.contains("<x:Column>0</x:Column>"));
        assert!(vml.contains("ObjectType=\"Note\""));
        assert!(vml.contains("_x0000_t202"));
        assert!(vml.contains("fillcolor=\"#ffffe1\""));
    }

    #[test]
    fn test_build_vml_drawing_multiple_cells() {
        let vml = build_vml_drawing(&["A1", "C5"]);
        assert!(vml.contains("<x:Row>0</x:Row>"));
        assert!(vml.contains("<x:Column>0</x:Column>"));
        assert!(vml.contains("<x:Row>4</x:Row>"));
        assert!(vml.contains("<x:Column>2</x:Column>"));
        assert!(vml.contains("_x0000_s1025"));
        assert!(vml.contains("_x0000_s1026"));
    }

    #[test]
    fn test_build_vml_drawing_empty() {
        let vml = build_vml_drawing(&[]);
        assert!(vml.contains("<o:shapelayout"));
        assert!(vml.contains("<v:shapetype"));
        assert!(!vml.contains("<v:shape id="));
    }

    #[test]
    fn test_extract_vml_comment_cells() {
        let vml = build_vml_drawing(&["B3", "D10"]);
        let cells = extract_vml_comment_cells(&vml);
        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0], (2, 1));
        assert_eq!(cells[1], (9, 3));
    }

    #[test]
    fn test_comment_anchor_format() {
        let anchor = comment_anchor(1, 1);
        assert!(anchor.contains(", "));
        let parts: Vec<&str> = anchor.split(", ").collect();
        assert_eq!(parts.len(), 8);
    }

    #[test]
    fn test_extract_element_value() {
        assert_eq!(
            extract_element_value("<x:Row>5</x:Row>", "x:Row"),
            Some("5")
        );
        assert_eq!(
            extract_element_value("<x:Column>3</x:Column>", "x:Column"),
            Some("3")
        );
        assert_eq!(extract_element_value("no match here", "x:Row"), None);
    }
}
