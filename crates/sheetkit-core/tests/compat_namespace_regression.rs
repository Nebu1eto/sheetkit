//! Real-file namespace compatibility: optional `xmlns:r`, extra namespace
//! declarations, and a prefixed SpreadsheetML default namespace.

use sheetkit_core::cell::CellValue;
use sheetkit_core::workbook::{OpenOptions, ReadMode, Workbook};
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;

const ALL_MODES: [ReadMode; 3] = [ReadMode::Eager, ReadMode::Lazy, ReadMode::Stream];
const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn parts(buffer: &[u8]) -> Vec<(String, Vec<u8>)> {
    let mut archive = zip::ZipArchive::new(Cursor::new(buffer)).unwrap();
    let mut output = Vec::new();
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        output.push((entry.name().to_string(), bytes));
    }
    output
}

fn rezip(parts: &[(String, Vec<u8>)]) -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, bytes) in parts {
        writer
            .start_file(name, SimpleFileOptions::default())
            .unwrap();
        writer.write_all(bytes).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

fn rewrite_part(buffer: &[u8], name: &str, rewrite: impl Fn(&str) -> String) -> Vec<u8> {
    let mut parts = parts(buffer);
    let part = parts
        .iter_mut()
        .find(|(part_name, _)| part_name == name)
        .unwrap();
    part.1 = rewrite(std::str::from_utf8(&part.1).unwrap()).into_bytes();
    rezip(&parts)
}

fn part_text(buffer: &[u8], name: &str) -> String {
    let bytes = parts(buffer)
        .into_iter()
        .find(|(part_name, _)| part_name == name)
        .unwrap()
        .1;
    String::from_utf8(bytes).unwrap()
}

fn base_package() -> Vec<u8> {
    let mut workbook = Workbook::new();
    workbook
        .set_cell_value("Sheet1", "A1", CellValue::String("kept".into()))
        .unwrap();
    workbook
        .set_cell_value("Sheet1", "B2", CellValue::Number(42.0))
        .unwrap();
    workbook.save_to_buffer().unwrap()
}

fn drop_root_rel_ns(xml: &str) -> String {
    let declaration = format!(" xmlns:r=\"{REL_NS}\"");
    assert!(xml.contains(&declaration), "fixture must declare xmlns:r");
    xml.replacen(&declaration, "", 1)
}

/// Remove `xmlns:r` from the workbook root and bind it on each `<sheet>` so
/// the `r:id` attributes stay namespace-valid.
fn workbook_without_root_rel_ns(xml: &str) -> String {
    drop_root_rel_ns(xml).replace("<sheet ", &format!("<sheet xmlns:r=\"{REL_NS}\" "))
}

fn add_unrelated_mc(xml: &str, root: &str) -> String {
    let open = format!("<{root} ");
    assert!(xml.contains(&open));
    xml.replacen(&open, &format!("<{root} xmlns:mc=\"{MC_NS}\" "), 1)
}

/// Rewrite every unprefixed element to use the `x:` prefix bound to the
/// SpreadsheetML main namespace, as OpenXML SDK writers do.
fn prefix_main_namespace(xml: &str) -> String {
    let default_ns = format!(" xmlns=\"{MAIN_NS}\"");
    assert!(xml.contains(&default_ns));
    let xml = xml.replacen(&default_ns, &format!(" xmlns:x=\"{MAIN_NS}\""), 1);
    let mut output = String::with_capacity(xml.len() + xml.len() / 4);
    let mut rest = xml.as_str();
    while let Some(position) = rest.find('<') {
        output.push_str(&rest[..=position]);
        rest = &rest[position + 1..];
        let (closing, name_start) = if let Some(stripped) = rest.strip_prefix('/') {
            (true, stripped)
        } else {
            (false, rest)
        };
        let is_element = name_start
            .chars()
            .next()
            .is_some_and(|ch| ch.is_ascii_alphabetic());
        let name_end = name_start
            .find(|ch: char| ch.is_whitespace() || ch == '>' || ch == '/')
            .unwrap_or(name_start.len());
        let unprefixed = !name_start[..name_end].contains(':');
        if closing {
            output.push('/');
            rest = name_start;
        }
        if is_element && unprefixed {
            output.push_str("x:");
        }
    }
    output.push_str(rest);
    output
}

fn assert_round_trip(buffer: &[u8], context: &str) {
    for mode in ALL_MODES {
        let options = OpenOptions::new().read_mode(mode);
        let workbook = Workbook::open_from_buffer_with_options(buffer, &options)
            .unwrap_or_else(|error| panic!("{context} {mode:?}: open failed: {error}"));
        assert_eq!(workbook.sheet_names(), vec!["Sheet1"], "{context} {mode:?}");
        let rows = workbook
            .get_rows("Sheet1")
            .unwrap_or_else(|error| panic!("{context} {mode:?}: get_rows failed: {error}"));
        assert_eq!(rows.len(), 2, "{context} {mode:?}: rows {rows:?}");
        let streamed = workbook
            .open_sheet_reader("Sheet1")
            .and_then(|mut reader| reader.next_batch(16))
            .unwrap_or_else(|error| panic!("{context} {mode:?}: sheet reader failed: {error}"));
        assert_eq!(streamed.len(), 2, "{context} {mode:?}: streamed rows");
        assert_eq!(
            workbook.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("kept".into()),
            "{context} {mode:?}"
        );
        assert_eq!(
            workbook.get_cell_value("Sheet1", "B2").unwrap(),
            CellValue::Number(42.0),
            "{context} {mode:?}"
        );

        let saved = workbook
            .save_to_buffer()
            .unwrap_or_else(|error| panic!("{context} {mode:?}: save failed: {error}"));
        let reopened = Workbook::open_from_buffer_with_options(&saved, &options)
            .unwrap_or_else(|error| panic!("{context} {mode:?}: reopen failed: {error}"));
        assert_eq!(
            reopened.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("kept".into()),
            "{context} {mode:?}: reopened"
        );
        assert_eq!(
            reopened.get_cell_value("Sheet1", "B2").unwrap(),
            CellValue::Number(42.0),
            "{context} {mode:?}: reopened"
        );
    }
}

#[test]
fn missing_relationship_namespace_opens_in_all_modes() {
    let base = base_package();
    let sheet = rewrite_part(&base, "xl/worksheets/sheet1.xml", drop_root_rel_ns);
    assert_round_trip(&sheet, "sheet without xmlns:r");
    let both = rewrite_part(&sheet, "xl/workbook.xml", workbook_without_root_rel_ns);
    assert_round_trip(&both, "sheet and workbook without root xmlns:r");
}

#[test]
fn missing_relationship_namespace_with_unrelated_mc_opens_in_all_modes() {
    let base = base_package();
    let sheet = rewrite_part(&base, "xl/worksheets/sheet1.xml", |xml| {
        add_unrelated_mc(&drop_root_rel_ns(xml), "worksheet")
    });
    let both = rewrite_part(&sheet, "xl/workbook.xml", |xml| {
        add_unrelated_mc(&workbook_without_root_rel_ns(xml), "workbook")
    });
    assert_round_trip(&both, "xmlns:mc without xmlns:r");
}

#[test]
fn prefixed_main_namespace_opens_in_all_modes() {
    let base = base_package();
    let sheet = rewrite_part(&base, "xl/worksheets/sheet1.xml", prefix_main_namespace);
    assert!(part_text(&sheet, "xl/worksheets/sheet1.xml").contains("<x:worksheet xmlns:x="));
    assert_round_trip(&sheet, "prefixed worksheet");
    let both = rewrite_part(&sheet, "xl/workbook.xml", prefix_main_namespace);
    assert!(part_text(&both, "xl/workbook.xml").contains("<x:sheet "));
    assert_round_trip(&both, "prefixed worksheet and workbook");
    let shared = rewrite_part(&both, "xl/sharedStrings.xml", prefix_main_namespace);
    let all = rewrite_part(&shared, "xl/styles.xml", prefix_main_namespace);
    assert!(part_text(&all, "xl/sharedStrings.xml").contains("<x:si>"));
    assert_round_trip(&all, "prefixed worksheet, workbook, strings, and styles");
}

#[test]
fn saved_roots_still_declare_both_namespaces() {
    let saved = base_package();
    let sheet = part_text(&saved, "xl/worksheets/sheet1.xml");
    let workbook = part_text(&saved, "xl/workbook.xml");
    for (xml, root) in [(&sheet, "worksheet"), (&workbook, "workbook")] {
        let start = xml.find(&format!("<{root} ")).unwrap();
        let tag = &xml[start..start + xml[start..].find('>').unwrap()];
        assert!(tag.contains(&format!("xmlns=\"{MAIN_NS}\"")), "{tag}");
        assert!(tag.contains(&format!("xmlns:r=\"{REL_NS}\"")), "{tag}");
    }

    let reopened = Workbook::open_from_buffer(&saved)
        .unwrap()
        .save_to_buffer()
        .unwrap();
    let resaved_sheet = part_text(&reopened, "xl/worksheets/sheet1.xml");
    assert!(resaved_sheet.contains(&format!("xmlns:r=\"{REL_NS}\"")));

    let stripped = rewrite_part(&saved, "xl/worksheets/sheet1.xml", drop_root_rel_ns);
    let mut opened = Workbook::open_from_buffer_with_options(
        &stripped,
        &OpenOptions::new().read_mode(ReadMode::Eager),
    )
    .unwrap();
    opened
        .set_cell_value("Sheet1", "C3", CellValue::Number(1.0))
        .unwrap();
    let dirty = opened.save_to_buffer().unwrap();
    let dirty_sheet = part_text(&dirty, "xl/worksheets/sheet1.xml");
    assert!(dirty_sheet.contains(&format!("xmlns=\"{MAIN_NS}\"")));
    assert!(dirty_sheet.contains(&format!("xmlns:r=\"{REL_NS}\"")));
}
