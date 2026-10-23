use sheetkit_core::cell::CellValue;
use sheetkit_core::table::{TableColumn, TableConfig};
use sheetkit_core::workbook::{AuxParts, OpenOptions, ReadMode, Workbook};
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;

fn parts(buffer: &[u8]) -> Vec<(String, Vec<u8>)> {
    let mut archive = zip::ZipArchive::new(Cursor::new(buffer)).unwrap();
    (0..archive.len())
        .map(|index| {
            let mut entry = archive.by_index(index).unwrap();
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes).unwrap();
            (entry.name().to_string(), bytes)
        })
        .collect()
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

fn table_part(buffer: &[u8]) -> (String, Vec<u8>) {
    parts(buffer)
        .into_iter()
        .find(|(name, _)| name.starts_with("xl/tables/"))
        .unwrap()
}

#[test]
fn eager_open_round_trips_untouched_table_bytes() {
    let mut workbook = Workbook::new();
    workbook
        .set_cell_value("Sheet1", "A1", CellValue::String("Value".into()))
        .unwrap();
    workbook
        .add_table(
            "Sheet1",
            &TableConfig {
                name: "Data".to_string(),
                display_name: "Data".to_string(),
                range: "A1:A5".to_string(),
                columns: vec![TableColumn {
                    name: "Value".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                }],
                ..Default::default()
            },
        )
        .unwrap();
    let base = workbook.save_to_buffer().unwrap();
    let (table_name, table_bytes) = table_part(&base);

    // Add features the typed model does not represent.
    let original = String::from_utf8(table_bytes).unwrap();
    let close = original.rfind("</table>").unwrap();
    let extended = format!(
        "{}<sortState ref=\"A2:A5\"><sortCondition ref=\"A2:A5\"/></sortState><extLst><ext uri=\"{{00000000-0000-0000-0000-000000000001}}\"><custom/></ext></extLst>{}",
        &original[..close],
        &original[close..]
    );
    let mut package = parts(&base);
    for (name, bytes) in &mut package {
        if *name == table_name {
            *bytes = extended.clone().into_bytes();
        }
    }
    let input = rezip(&package);

    let options = OpenOptions::new()
        .read_mode(ReadMode::Eager)
        .aux_parts(AuxParts::EagerLoad);
    let opened = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
    assert_eq!(opened.get_tables("Sheet1").unwrap().len(), 1);
    let saved = opened.save_to_buffer().unwrap();
    let (saved_name, saved_bytes) = table_part(&saved);
    assert_eq!(saved_name, table_name);
    assert_eq!(String::from_utf8(saved_bytes).unwrap(), extended);
}
