//! VML (Vector Markup Language) support for Excel legacy comment rendering.
//!
//! Excel uses VML drawing parts (`xl/drawings/vmlDrawingN.vml`) to render
//! comment/note pop-up boxes in the UI. This module generates minimal VML
//! markup for new comments and tracks preserved VML bytes for round-tripping.

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::NsReader;

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
    let shapes = build_comment_shapes(cells, 1025);

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

/// Merge regenerated comment shapes into preserved VML content.
pub fn merge_vml_comments(existing_vml: &[u8], cells: &[&str], start_shape_id: usize) -> Vec<u8> {
    let existing = String::from_utf8_lossy(existing_vml);
    let shapes = build_comment_shapes(cells, start_shape_id);
    if let Some(close_pos) = existing.rfind("</xml>") {
        let mut result = String::with_capacity(existing.len() + shapes.len());
        result.push_str(&existing[..close_pos]);
        if !existing.contains("_x0000_t202") {
            result.push_str(
                " <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" \
                 o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">\n\
                  <v:stroke joinstyle=\"miter\"/>\n\
                  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n\
                 </v:shapetype>\n",
            );
        }
        result.push_str(&shapes);
        result.push_str("</xml>\n");
        result.into_bytes()
    } else {
        build_vml_drawing(cells).into_bytes()
    }
}

/// VML namespace URI for `v:` elements.
const VML_NS: &[u8] = b"urn:schemas-microsoft-com:vml";
/// Excel namespace URI for `x:` elements such as `x:ClientData`.
const EXCEL_NS: &[u8] = b"urn:schemas-microsoft-com:office:excel";

/// Remove comment shapes while retaining non-comment VML content.
///
/// A comment shape is a VML `shape` element containing an Excel `ClientData`
/// element whose `ObjectType` is `Note`. Elements are matched by namespace and
/// local name, so any prefix, quote style, or attribute whitespace is accepted.
/// All other content is preserved byte-for-byte. Returns `None` when no
/// non-comment shape remains or the VML cannot be parsed.
pub fn strip_vml_comment_shapes_from_vml(vml_bytes: &[u8]) -> Option<Vec<u8>> {
    filter_vml_shapes(vml_bytes, |is_note| !is_note)
}

fn build_comment_shapes(cells: &[&str], start_shape_id: usize) -> String {
    let mut shapes = String::new();
    for (i, cell) in cells.iter().enumerate() {
        if let Ok((col, row)) = cell_name_to_coordinates(cell) {
            let anchor = comment_anchor(col, row);
            write_vml_shape(
                &mut shapes,
                start_shape_id + i,
                i + 1,
                &anchor,
                row - 1,
                col - 1,
            );
        }
    }
    shapes
}

/// Whether `ns` is the namespace `uri`, also accepting an undeclared
/// conventional prefix for leniency with hand-written VML.
fn in_namespace(ns: &ResolveResult<'_>, uri: &[u8], conventional_prefix: &[u8]) -> bool {
    match ns {
        ResolveResult::Bound(bound) => bound.as_ref() == uri,
        ResolveResult::Unknown(prefix) => prefix.as_slice() == conventional_prefix,
        ResolveResult::Unbound => false,
    }
}

fn is_vml_shape(ns: &ResolveResult<'_>, local_name: &[u8]) -> bool {
    local_name == b"shape" && in_namespace(ns, VML_NS, b"v")
}

fn is_note_client_data(ns: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    element.local_name().as_ref() == b"ClientData"
        && in_namespace(ns, EXCEL_NS, b"x")
        && element
            .attributes()
            .flatten()
            .find(|attr| attr.key.as_ref() == b"ObjectType")
            .and_then(|attr| attr.unescape_value().ok())
            .is_some_and(|value| value.trim() == "Note")
}

/// Remove VML `shape` elements for which `keep(is_note)` returns false.
///
/// Byte ranges of removed shapes (plus one trailing line break) are cut from
/// the original text; everything else is preserved. Returns `None` if no shape
/// is kept or the markup cannot be parsed.
fn filter_vml_shapes<F>(vml_bytes: &[u8], mut keep: F) -> Option<Vec<u8>>
where
    F: FnMut(bool) -> bool,
{
    let vml = String::from_utf8_lossy(vml_bytes);
    let mut reader = NsReader::from_str(&vml);
    // Legacy VML textboxes may contain HTML-style unclosed tags such as `<br>`.
    reader.config_mut().check_end_names = false;

    // Start offset and note flag of the shape currently open.
    let mut open_shape: Option<(usize, bool)> = None;
    let mut removals: Vec<(usize, usize)> = Vec::new();
    let mut retained = false;

    loop {
        let event_start = usize::try_from(reader.buffer_position()).ok()?;
        let (ns, event) = reader.read_resolved_event().ok()?;
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let is_empty = matches!(event, Event::Empty(_));
                if let Some((_, is_note)) = open_shape.as_mut() {
                    if is_note_client_data(&ns, e) {
                        *is_note = true;
                    }
                } else if is_vml_shape(&ns, e.local_name().as_ref()) {
                    if is_empty {
                        if keep(false) {
                            retained = true;
                        } else {
                            let end = usize::try_from(reader.buffer_position()).ok()?;
                            removals.push((event_start, end));
                        }
                    } else {
                        open_shape = Some((event_start, false));
                    }
                }
            }
            Event::End(ref e) => {
                if open_shape.is_some() && is_vml_shape(&ns, e.local_name().as_ref()) {
                    if let Some((start, is_note)) = open_shape.take() {
                        if keep(is_note) {
                            retained = true;
                        } else {
                            let end = usize::try_from(reader.buffer_position()).ok()?;
                            removals.push((start, end));
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    if open_shape.is_some() || !retained {
        return None;
    }

    let mut result = String::with_capacity(vml.len());
    let mut cursor = 0;
    for (start, end) in removals {
        result.push_str(vml.get(cursor..start)?);
        let rest = vml.get(end..)?;
        cursor = end
            + if rest.starts_with("\r\n") {
                2
            } else if rest.starts_with('\n') {
                1
            } else {
                0
            };
    }
    result.push_str(vml.get(cursor..)?);
    Some(result.into_bytes())
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
        if let (Some(row), Some(col)) = (current_row, current_col) {
            cells.push((row, col));
            current_row = None;
            current_col = None;
        }
    }
    cells
}

/// Simple extraction of text content from an XML element like `<tag>value</tag>`.
///
/// The closing tag is searched only after the first opening tag, so malformed
/// input yields `None` instead of an invalid slice.
fn extract_element_value<'a>(line: &'a str, tag: &str) -> Option<&'a str> {
    let open = format!("<{}>", tag);
    let close = format!("</{}>", tag);
    let val_start = line.find(&open)? + open.len();
    let rest = line.get(val_start..)?;
    let len = rest.find(&close)?;
    rest.get(..len)
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

    #[test]
    fn test_extract_vml_comment_cells_never_panics_on_malformed_tags() {
        let cases: &[(&str, Vec<(u32, u32)>)] = &[
            ("</x:Row><x:Row>0</x:Row>", vec![]),
            (
                "</x:Row><x:Row>0</x:Row><x:Column>1</x:Column>",
                vec![(0, 1)],
            ),
            ("</x:Column><x:Column>", vec![]),
            ("<x:Row>1", vec![]),
            ("<x:Row></x:Row>", vec![]),
            (
                "<x:Row>1</x:Row><x:Row>2</x:Row><x:Column>3</x:Column>",
                vec![(1, 3)],
            ),
            (
                "<x:Row><x:Row>4</x:Row></x:Row><x:Column>5</x:Column>",
                vec![],
            ),
            ("</x:Row>", vec![]),
            ("", vec![]),
        ];
        for (input, expected) in cases {
            assert_eq!(&extract_vml_comment_cells(input), expected, "{input}");
        }

        let pieces = [
            "<x:Row>",
            "</x:Row>",
            "<x:Column>",
            "</x:Column>",
            "7",
            "\n",
            "\u{e9}",
        ];
        let mut input = String::new();
        for a in pieces {
            for b in pieces {
                for c in pieces {
                    for d in pieces {
                        input.clear();
                        input.push_str(a);
                        input.push_str(b);
                        input.push_str(c);
                        input.push_str(d);
                        let _ = extract_vml_comment_cells(&input);
                    }
                }
            }
        }
    }

    const VML_NS_HEADER: &str = "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">";

    #[test]
    fn test_strip_comment_shapes_single_quoted_object_type() {
        let vml = format!(
            "{VML_NS_HEADER}\n<v:shape id='c1'><x:ClientData ObjectType='Note'><x:Row>0</x:Row></x:ClientData></v:shape>\n<v:shape id='b1'><x:ClientData ObjectType='Button'/></v:shape>\n</xml>\n"
        );
        let out =
            String::from_utf8(strip_vml_comment_shapes_from_vml(vml.as_bytes()).unwrap()).unwrap();
        assert!(!out.contains("c1"), "{out}");
        assert!(out.contains("<v:shape id='b1'><x:ClientData ObjectType='Button'/></v:shape>"));
        assert!(out.trim_end().ends_with("</xml>"));
    }

    #[test]
    fn test_strip_comment_shapes_whitespace_and_newlines_in_attributes() {
        let vml = format!(
            "{VML_NS_HEADER}\n<v:shape\n id=\"c1\"><x:ClientData\n   ObjectType = \"Note\"\n></x:ClientData></v:shape\n>\n<v:shape\tid=\"b1\"><x:ClientData ObjectType=\"Drop\"/></v:shape>\n</xml>\n"
        );
        let out =
            String::from_utf8(strip_vml_comment_shapes_from_vml(vml.as_bytes()).unwrap()).unwrap();
        assert!(!out.contains("c1"), "{out}");
        assert!(out.contains("<v:shape\tid=\"b1\">"));
    }

    #[test]
    fn test_strip_comment_shapes_alternate_vml_prefix() {
        let vml = "<xml xmlns:s=\"urn:schemas-microsoft-com:vml\" xmlns:e=\"urn:schemas-microsoft-com:office:excel\">\n<s:shape id=\"c1\"><e:ClientData ObjectType=\"Note\"/></s:shape>\n<s:shape id=\"b1\"><e:ClientData ObjectType=\"Button\"/></s:shape>\n</xml>\n";
        let out =
            String::from_utf8(strip_vml_comment_shapes_from_vml(vml.as_bytes()).unwrap()).unwrap();
        assert!(!out.contains("c1"), "{out}");
        assert!(out.contains("<s:shape id=\"b1\"><e:ClientData ObjectType=\"Button\"/></s:shape>"));
        assert!(out.starts_with("<xml xmlns:s="));
    }

    #[test]
    fn test_strip_comment_shapes_preserves_surrounding_content() {
        let vml = format!(
            "{VML_NS_HEADER}\n <o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>\n<v:shapetype id=\"_x0000_t201\"/>\n<v:shape id=\"b1\"><x:ClientData ObjectType=\"Button\"/></v:shape>\n<v:shapetype id=\"_x0000_t202\"/>\n<v:shape id=\"c1\"><x:ClientData ObjectType=\"Note\"/></v:shape>\n<v:rect id=\"r1\"/>\n</xml>\n"
        );
        let out =
            String::from_utf8(strip_vml_comment_shapes_from_vml(vml.as_bytes()).unwrap()).unwrap();
        assert!(!out.contains("c1"), "{out}");
        for kept in [
            "<o:shapelayout",
            "_x0000_t201",
            "_x0000_t202",
            "id=\"b1\"",
            "<v:rect id=\"r1\"/>",
        ] {
            assert!(out.contains(kept), "missing {kept}: {out}");
        }
        assert_eq!(out.matches("</xml>").count(), 1);
    }

    #[test]
    fn test_strip_comment_shapes_only_comments_returns_none() {
        let vml = build_vml_drawing(&["A1", "B2"]);
        assert!(strip_vml_comment_shapes_from_vml(vml.as_bytes()).is_none());
    }

    #[test]
    fn test_strip_comment_shapes_ignores_foreign_namespace_lookalikes() {
        // A `shape` element outside the VML namespace is not a VML shape, and a
        // `ClientData` element outside the Excel namespace does not mark a note.
        let vml = format!(
            "{VML_NS_HEADER}<other:shape xmlns:other=\"urn:other\" id=\"f1\"/><v:shape id=\"k1\"><y:ClientData xmlns:y=\"urn:other\" ObjectType=\"Note\"/></v:shape></xml>"
        );
        let out =
            String::from_utf8(strip_vml_comment_shapes_from_vml(vml.as_bytes()).unwrap()).unwrap();
        assert!(out.contains("id=\"f1\""));
        assert!(out.contains("id=\"k1\""));
    }

    #[test]
    fn test_strip_comment_shapes_then_merge_has_no_stale_notes() {
        let mut vml = build_vml_drawing(&["A1"]);
        let button = " <v:shape id='_x0000_s2000' type='#_x0000_t201'><x:ClientData ObjectType='Button'/></v:shape>\n";
        vml.insert_str(vml.rfind("</xml>").unwrap(), button);
        let vml = vml.replace("ObjectType=\"Note\"", "ObjectType='Note'");

        let stripped = strip_vml_comment_shapes_from_vml(vml.as_bytes()).unwrap();
        let merged = String::from_utf8(merge_vml_comments(&stripped, &["C3"], 3000)).unwrap();
        assert_eq!(extract_vml_comment_cells(&merged), vec![(2, 2)]);
        assert!(merged.contains("_x0000_s2000"));
    }

    #[test]
    fn test_strip_comment_shapes_never_panics() {
        for input in [
            "",
            "<xml>",
            "<v:shape ",
            "<v:shape id='a'>",
            "</v:shape><v:shape id='a'></v:shape>",
            "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape><v:shape></v:shape>",
            "\u{0}\u{ff}<",
        ] {
            let _ = strip_vml_comment_shapes_from_vml(input.as_bytes());
        }
    }
}
