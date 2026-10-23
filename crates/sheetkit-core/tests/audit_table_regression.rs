use std::io::{Cursor, Read, Write};

use sheetkit_core::cell::CellValue;
use sheetkit_core::error::Error;
use sheetkit_core::table::{TableColumn, TableConfig};
use sheetkit_core::workbook::{OpenOptions, ReadMode, Workbook};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const RICH_TABLE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" id="1" name="Calc" displayName="Calc" ref="A1:B4" totalsRowCount="1"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="Label" totalsRowLabel="Total"/><tableColumn id="2" name="Value" totalsRowFunction="custom"><calculatedColumnFormula>1+1</calculatedColumnFormula><totalsRowFormula>SUBTOTAL(109,Calc[Value])</totalsRowFormula></tableColumn></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/><extLst><ext uri="{504A1905-F514-4f6f-8877-14C23A59335A}"><x14:table altText="Calc table"/></ext></extLst></table>"#;

fn column(name: &str) -> TableColumn {
    TableColumn {
        name: name.to_string(),
        totals_row_function: None,
        totals_row_label: None,
    }
}

fn config(name: &str, range: &str, columns: &[&str]) -> TableConfig {
    TableConfig {
        name: name.to_string(),
        display_name: name.to_string(),
        range: range.to_string(),
        columns: columns.iter().map(|c| column(c)).collect(),
        ..TableConfig::default()
    }
}

fn read_part(package: &[u8], path: &str) -> String {
    let mut archive = ZipArchive::new(Cursor::new(package)).unwrap();
    let mut part = archive.by_name(path).unwrap();
    let mut xml = String::new();
    part.read_to_string(&mut xml).unwrap();
    xml
}

fn replace_part(package: &[u8], path: &str, data: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(package)).unwrap();
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for i in 0..archive.len() {
        let mut entry = archive.by_index(i).unwrap();
        let name = entry.name().to_string();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        writer.start_file(name.as_str(), options).unwrap();
        if name == path {
            writer.write_all(data.as_bytes()).unwrap();
        } else {
            writer.write_all(&bytes).unwrap();
        }
    }
    writer.finish().unwrap().into_inner()
}

/// A package whose `xl/tables/table1.xml` is `RICH_TABLE_XML` on Sheet1.
fn rich_table_package() -> Vec<u8> {
    let mut wb = Workbook::new();
    wb.set_cell_value("Sheet1", "A2", "a").unwrap();
    wb.add_table("Sheet1", &config("Calc", "A1:B4", &["Label", "Value"]))
        .unwrap();
    let package = wb.save_to_buffer().unwrap();
    assert!(read_part(&package, "xl/tables/table1.xml").contains(r#"name="Calc""#));
    replace_part(&package, "xl/tables/table1.xml", RICH_TABLE_XML)
}

fn add_unrelated_table_and_save(options: &OpenOptions) -> String {
    let package = rich_table_package();
    let mut wb = Workbook::open_from_buffer_with_options(&package, options).unwrap();
    wb.add_table("Sheet1", &config("Other", "D1:E3", &["X", "Y"]))
        .unwrap();
    let saved = wb.save_to_buffer().unwrap();

    let reopened = Workbook::open_from_buffer(&saved).unwrap();
    let mut names: Vec<String> = reopened
        .get_tables("Sheet1")
        .unwrap()
        .into_iter()
        .map(|t| t.name)
        .collect();
    names.sort();
    assert_eq!(names, ["Calc", "Other"]);
    read_part(&saved, "xl/tables/table1.xml")
}

#[test]
fn lazy_add_table_preserves_existing_table_content() {
    let xml = add_unrelated_table_and_save(&OpenOptions::new().read_mode(ReadMode::Lazy));
    assert!(
        xml.contains("<calculatedColumnFormula>1+1</calculatedColumnFormula>"),
        "{xml}"
    );
    assert!(
        xml.contains("<totalsRowFormula>SUBTOTAL(109,Calc[Value])</totalsRowFormula>"),
        "{xml}"
    );
    assert!(
        xml.contains(r#"<x14:table altText="Calc table"/>"#),
        "{xml}"
    );
    assert!(
        xml.contains(
            r#"xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main""#
        ),
        "{xml}"
    );
}

#[test]
fn eager_add_table_preserves_existing_column_formulas() {
    let xml = add_unrelated_table_and_save(&OpenOptions::new());
    assert!(
        xml.contains("<calculatedColumnFormula>1+1</calculatedColumnFormula>"),
        "{xml}"
    );
    assert!(
        xml.contains("<totalsRowFormula>SUBTOTAL(109,Calc[Value])</totalsRowFormula>"),
        "{xml}"
    );
}

fn assert_rejected(wb: &mut Workbook, cfg: &TableConfig) {
    let before = wb.get_cell_value("Sheet1", "H1").unwrap();
    let result = wb.add_table("Sheet1", cfg);
    assert!(
        matches!(
            result,
            Err(Error::InvalidArgument(_)) | Err(Error::TableAlreadyExists { .. })
        ),
        "{} / {}: {result:?}",
        cfg.name,
        cfg.display_name
    );
    assert_eq!(wb.get_cell_value("Sheet1", "H1").unwrap(), before);
    assert!(wb
        .get_tables("Sheet1")
        .unwrap()
        .iter()
        .all(|t| t.name != cfg.name));
}

#[test]
fn add_table_rejects_invalid_names_before_mutation() {
    let mut wb = Workbook::new();
    let long = "a".repeat(256);
    for name in [
        "A1",
        "a1",
        "XFD1048576",
        "R1C1",
        "r",
        "C",
        "RC",
        "R10",
        "Bad Name",
        "1Table",
        ".Table",
        "Tab-le",
        "Table\u{3000}X",
        long.as_str(),
    ] {
        assert_rejected(&mut wb, &config(name, "H1:H3", &["Head"]));
    }

    let mut bad_display = config("Good", "H1:H3", &["Head"]);
    bad_display.display_name = "Bad Display".to_string();
    assert_rejected(&mut wb, &bad_display);
    assert_eq!(wb.get_cell_value("Sheet1", "H1").unwrap(), CellValue::Empty);
}

#[test]
fn add_table_rejects_case_insensitive_duplicates() {
    let mut wb = Workbook::new();
    wb.add_table("Sheet1", &config("Sales", "A1:A3", &["Amount"]))
        .unwrap();
    assert_rejected(&mut wb, &config("sales", "H1:H3", &["Head"]));
    assert_rejected(&mut wb, &config("SALES", "H1:H3", &["Head"]));

    let mut display_clash = config("Unique", "H1:H3", &["Head"]);
    display_clash.display_name = "sAlEs".to_string();
    assert_rejected(&mut wb, &display_clash);

    wb.set_defined_name("Budget", "Sheet1!$A$1", None, None)
        .unwrap();
    assert_rejected(&mut wb, &config("budget", "H1:H3", &["Head"]));
    assert_eq!(wb.get_tables("Sheet1").unwrap().len(), 1);
}

#[test]
fn add_table_accepts_boundary_names() {
    let mut wb = Workbook::new();
    let longest = "t".repeat(255);
    let names = [
        "_t1",
        "Table.Name",
        "\\Back",
        "XFE1",
        "R1C1x",
        "Ventes\u{e9}",
        longest.as_str(),
    ];
    for (i, name) in names.iter().enumerate() {
        let col = char::from(b'A' + i as u8);
        let range = format!("{col}1:{col}3");
        wb.add_table("Sheet1", &config(name, &range, &["Head"]))
            .unwrap_or_else(|e| panic!("{name}: {e}"));
    }
    assert_eq!(wb.get_tables("Sheet1").unwrap().len(), names.len());
}

#[test]
fn add_table_empty_display_name_falls_back_to_name() {
    let mut wb = Workbook::new();
    let mut cfg = config("Plain", "A1:A3", &["Head"]);
    cfg.display_name = String::new();
    wb.add_table("Sheet1", &cfg).unwrap();
    let tables = wb.get_tables("Sheet1").unwrap();
    assert_eq!(tables[0].display_name, "Plain");
}

#[test]
fn add_table_allocates_free_id_when_existing_id_is_u32_max() {
    let package = rich_table_package();
    let package = replace_part(
        &package,
        "xl/tables/table1.xml",
        &RICH_TABLE_XML.replace(r#"id="1""#, &format!(r#"id="{}""#, u32::MAX)),
    );
    let mut wb = Workbook::open_from_buffer(&package).unwrap();
    wb.add_table("Sheet1", &config("Other", "D1:E3", &["X", "Y"]))
        .unwrap();
    let saved = wb.save_to_buffer().unwrap();
    let ids: Vec<String> = (1..=2)
        .map(|n| read_part(&saved, &format!("xl/tables/table{n}.xml")))
        .collect();
    assert!(ids[0].contains(&format!(r#"id="{}""#, u32::MAX)));
    assert!(ids[1].contains(r#" id="1""#), "{}", ids[1]);
}
