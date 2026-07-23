use std::fmt::Write as _;
use std::io::{Cursor, Read as _, Write as _};
use std::sync::{Arc, Barrier};

use sheetkit_core::cell::CellValue;
use sheetkit_core::comment::CommentConfig;
use sheetkit_core::error::Error;
use sheetkit_core::resource::{
    ResourceCategory, ResourceDimension, ResourceOperation, ResourceStage,
};
use sheetkit_core::sst::SharedStringTable;
use sheetkit_core::stream_reader::{SheetStreamReader, MAX_STREAM_BATCH_ROWS};
use sheetkit_core::table::{TableColumn, TableConfig};
use sheetkit_core::workbook::{AuxParts, OpenOptions, ReadMode, Workbook};

fn worksheet_xml(rows: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>{rows}</sheetData>
</worksheet>"#
    )
    .into_bytes()
}

fn worksheet_xml_with_table(rows: &str, relationship_id: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>{rows}</sheetData><tableParts count="1"><tablePart r:id="{relationship_id}"/></tableParts>
</worksheet>"#
    )
    .into_bytes()
}

fn replace_zip_part(input: &[u8], path: &str, replacement: &[u8]) -> Vec<u8> {
    let mut source = zip::ZipArchive::new(Cursor::new(input)).unwrap();
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..source.len() {
        let entry = source.by_index(index).unwrap();
        if entry.name() == path {
            output
                .start_file(
                    path,
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .unwrap();
            output.write_all(replacement).unwrap();
        } else {
            output.raw_copy_file(entry).unwrap();
        }
    }
    output.finish().unwrap().into_inner()
}

fn raw_zip_part(input: &[u8], path: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(input)).unwrap();
    let index = (0..archive.len())
        .find(|index| archive.by_index(*index).unwrap().name() == path)
        .unwrap();
    let mut entry = archive.by_index_raw(index).unwrap();
    let mut raw = Vec::new();
    entry.read_to_end(&mut raw).unwrap();
    raw
}

fn zip_part(input: &[u8], path: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(input)).unwrap();
    let mut entry = archive.by_name(path).unwrap();
    let mut bytes = Vec::new();
    entry.read_to_end(&mut bytes).unwrap();
    bytes
}

fn two_sheet_package() -> Vec<u8> {
    let mut workbook = Workbook::new();
    workbook.new_sheet("Excluded").unwrap();
    workbook.set_cell_value("Sheet1", "A1", "Included").unwrap();
    workbook
        .set_cell_value("Excluded", "A1", "Original excluded")
        .unwrap();
    workbook.save_to_buffer().unwrap()
}

fn two_sheet_table_package() -> Vec<u8> {
    let mut workbook = Workbook::new();
    workbook.new_sheet("Excluded").unwrap();
    workbook.set_cell_value("Sheet1", "A1", "Included").unwrap();
    workbook.set_cell_value("Excluded", "A1", "Header").unwrap();
    workbook.set_cell_value("Excluded", "A2", "Value").unwrap();
    workbook
        .add_table(
            "Excluded",
            &TableConfig {
                name: "ExcludedTable".to_string(),
                display_name: "ExcludedTable".to_string(),
                range: "A1:A2".to_string(),
                columns: vec![TableColumn {
                    name: "Header".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                }],
                ..TableConfig::default()
            },
        )
        .unwrap();
    workbook.save_to_buffer().unwrap()
}

fn assert_batch_limit(error: Error) {
    let metadata = error
        .resource_metadata()
        .expect("batch overflow must preserve typed resource metadata");
    assert_eq!(metadata.operation(), ResourceOperation::StreamRead);
    assert_eq!(metadata.category(), ResourceCategory::Streaming);
    assert_eq!(metadata.dimension(), ResourceDimension::RowCount);
    assert_eq!(metadata.limit(), MAX_STREAM_BATCH_ROWS as u64);
    assert_eq!(metadata.observed(), MAX_STREAM_BATCH_ROWS as u64 + 1);
    assert_eq!(metadata.stage(), Some(ResourceStage::Validate));
}

fn open_dirty_reader_with_limit(input: &[u8], limit: u64, owned: bool) -> Result<(), Error> {
    let mut workbook = Workbook::open_from_buffer_with_options(
        input,
        &OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .max_xml_part_bytes(limit),
    )?;
    for row in 1..=200 {
        workbook.set_cell_value("Sheet1", &format!("A{row}"), f64::from(row))?;
    }
    if owned {
        let _reader = workbook.open_sheet_reader_owned("Sheet1")?;
    } else {
        let _reader = workbook.open_sheet_reader("Sheet1")?;
    }
    Ok(())
}

#[test]
fn borrowed_reader_validates_batch_policy_before_reserve_or_read() {
    let xml = worksheet_xml(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#);
    let sst = SharedStringTable::new();
    let mut reader = SheetStreamReader::new(Cursor::new(xml.as_slice()), &sst);

    assert!(matches!(
        reader.next_batch(0),
        Err(Error::InvalidArgument(_))
    ));
    assert_batch_limit(reader.next_batch(MAX_STREAM_BATCH_ROWS + 1).unwrap_err());
    assert_eq!(reader.next_batch(MAX_STREAM_BATCH_ROWS).unwrap().len(), 1);
}

#[test]
fn owned_reader_validates_batch_policy_before_reserve_or_read() {
    let mut workbook = Workbook::new();
    workbook.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    let bytes = workbook.save_to_buffer().unwrap();
    let opened = Workbook::open_from_buffer_with_options(
        &bytes,
        &OpenOptions::new().read_mode(ReadMode::Stream),
    )
    .unwrap();
    let mut reader = opened.open_sheet_reader_owned("Sheet1").unwrap();

    assert!(matches!(
        reader.next_batch(0),
        Err(Error::InvalidArgument(_))
    ));
    assert_batch_limit(reader.next_batch(MAX_STREAM_BATCH_ROWS + 1).unwrap_err());
    assert_eq!(reader.next_batch(MAX_STREAM_BATCH_ROWS).unwrap().len(), 1);
}

#[test]
fn row_limit_stops_before_materializing_a_skipped_row() {
    let xml = worksheet_xml(
        r#"<row r="1"><c r="A1"><v>1</v></c></row>
<row r="2"><c r="A2" t="s"><v>999999999</v></c></row>"#,
    );
    let sst = SharedStringTable::new();

    let mut borrowed = SheetStreamReader::new(Cursor::new(xml.as_slice()), &sst).row_limit(Some(1));
    let rows = borrowed.next_batch(MAX_STREAM_BATCH_ROWS).unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].row_number, 1);

    let mut workbook = Workbook::new();
    workbook.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    workbook.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    let bytes = workbook.save_to_buffer().unwrap();
    let opened = Workbook::open_from_buffer_with_options(
        &bytes,
        &OpenOptions::new().read_mode(ReadMode::Stream).sheet_rows(1),
    )
    .unwrap();
    let mut owned = opened.open_sheet_reader_owned("Sheet1").unwrap();
    let rows = owned.next_batch(MAX_STREAM_BATCH_ROWS).unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].row_number, 1);
}

#[test]
fn eager_open_applies_row_limit_before_typed_row_construction() {
    let base = Workbook::new().save_to_buffer().unwrap();
    let xml = worksheet_xml(
        r#"<row r="1"><c r="A1"><v>1</v></c></row>
<row r="2"><c r="not-a-cell"><v>2</v></c></row>"#,
    );
    let input = replace_zip_part(&base, "xl/worksheets/sheet1.xml", &xml);

    let limited = Workbook::open_from_buffer_with_options(
        &input,
        &OpenOptions::new().read_mode(ReadMode::Eager).sheet_rows(1),
    )
    .unwrap();
    assert_eq!(limited.get_cell_value("Sheet1", "A1").unwrap(), 1.0.into());

    assert!(Workbook::open_from_buffer_with_options(
        &input,
        &OpenOptions::new().read_mode(ReadMode::Eager),
    )
    .is_err());
}

#[test]
fn filtered_sheet_is_not_decompressed_and_raw_payload_round_trips() {
    let base = two_sheet_package();
    let mut rows = String::with_capacity(2 * 1024 * 1024);
    for row in 1..=45_000 {
        write!(
            rows,
            "<row r=\"{row}\"><c r=\"A{row}\" t=\"inlineStr\"><is><t>excluded-{row:08}-payload</t></is></c></row>"
        )
        .unwrap();
    }
    let large_xml = worksheet_xml(&rows);
    assert!(large_xml.len() > 2 * 1024 * 1024);
    let input = replace_zip_part(&base, "xl/worksheets/sheet2.xml", &large_xml);
    let original_raw = raw_zip_part(&input, "xl/worksheets/sheet2.xml");

    let options = OpenOptions::new()
        .read_mode(ReadMode::Eager)
        .sheets(vec!["Sheet1".to_string()])
        .max_aggregate_actual_bytes(1024 * 1024);
    let workbook = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
    assert!(workbook
        .open_sheet_reader("Excluded")
        .unwrap()
        .next_batch(1)
        .unwrap()
        .is_empty());

    let output = workbook.save_to_buffer().unwrap();
    assert_eq!(
        raw_zip_part(&output, "xl/worksheets/sheet2.xml"),
        original_raw
    );
}

#[test]
fn eager_aux_does_not_decompress_filtered_sheet() {
    let base = two_sheet_package();
    let mut rows = String::with_capacity(2 * 1024 * 1024);
    for row in 1..=45_000 {
        write!(
            rows,
            "<row r=\"{row}\"><c r=\"A{row}\" t=\"inlineStr\"><is><t>excluded-{row:08}-payload</t></is></c></row>"
        )
        .unwrap();
    }
    let input = replace_zip_part(&base, "xl/worksheets/sheet2.xml", &worksheet_xml(&rows));
    let original_raw = raw_zip_part(&input, "xl/worksheets/sheet2.xml");

    let workbook = Workbook::open_from_buffer_with_options(
        &input,
        &OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad)
            .sheets(vec!["Sheet1".to_string()])
            .max_aggregate_actual_bytes(1024 * 1024),
    )
    .unwrap();
    let output = workbook.save_to_buffer().unwrap();
    assert_eq!(
        raw_zip_part(&output, "xl/worksheets/sheet2.xml"),
        original_raw
    );
}

#[test]
fn table_validation_does_not_decompress_filtered_sheet() {
    let base = two_sheet_table_package();
    let mut rows = String::with_capacity(2 * 1024 * 1024);
    for row in 1..=45_000 {
        write!(
            rows,
            "<row r=\"{row}\"><c r=\"A{row}\" t=\"inlineStr\"><is><t>excluded-{row:08}-payload</t></is></c></row>"
        )
        .unwrap();
    }
    let input = replace_zip_part(
        &base,
        "xl/worksheets/sheet2.xml",
        &worksheet_xml_with_table(&rows, "rId1"),
    );
    let original_raw = raw_zip_part(&input, "xl/worksheets/sheet2.xml");

    for aux_parts in [AuxParts::Deferred, AuxParts::EagerLoad] {
        let workbook = Workbook::open_from_buffer_with_options(
            &input,
            &OpenOptions::new()
                .read_mode(ReadMode::Eager)
                .aux_parts(aux_parts)
                .sheets(vec!["Sheet1".to_string()])
                .max_aggregate_actual_bytes(1024 * 1024),
        )
        .unwrap();
        let output = workbook.save_to_buffer().unwrap();
        assert_eq!(
            raw_zip_part(&output, "xl/worksheets/sheet2.xml"),
            original_raw
        );
    }
}

#[test]
fn file_backed_source_identity_survives_rename_and_path_replacement() {
    let directory = tempfile::TempDir::new().unwrap();
    let original_path = directory.path().join("source.xlsx");
    let moved_path = directory.path().join("moved.xlsx");
    let output_path = directory.path().join("output.xlsx");

    let mut original = Workbook::new();
    original
        .set_cell_value("Sheet1", "A1", "Original source")
        .unwrap();
    original.save(&original_path).unwrap();
    let opened = Workbook::open_with_options(
        &original_path,
        &OpenOptions::new().read_mode(ReadMode::Stream),
    )
    .unwrap();

    std::fs::rename(&original_path, &moved_path).unwrap();
    let mut replacement = Workbook::new();
    replacement
        .set_cell_value("Sheet1", "A1", "Replacement source")
        .unwrap();
    replacement.save(&original_path).unwrap();

    let mut reader = opened.open_sheet_reader_owned("Sheet1").unwrap();
    let rows = reader.next_batch(1).unwrap();
    assert_eq!(
        rows[0].cells[0].1,
        CellValue::String("Original source".to_string())
    );

    opened.save(&output_path).unwrap();
    let round_trip = Workbook::open(&output_path).unwrap();
    assert_eq!(
        round_trip.get_cell_value("Sheet1", "A1").unwrap(),
        CellValue::String("Original source".to_string())
    );
}

#[test]
fn dirty_borrowed_and_owned_readers_stream_current_sheet_state() {
    let bytes = two_sheet_package();
    let mut workbook = Workbook::open_from_buffer_with_options(
        &bytes,
        &OpenOptions::new().read_mode(ReadMode::Stream),
    )
    .unwrap();
    workbook
        .set_cell_value("Sheet1", "A1", "Dirty value")
        .unwrap();

    let mut borrowed = workbook.open_sheet_reader("Sheet1").unwrap();
    let borrowed_rows = borrowed.next_batch(1).unwrap();
    assert_eq!(
        borrowed_rows[0].cells[0].1,
        CellValue::String("Dirty value".to_string())
    );
    drop(borrowed);

    let mut owned = workbook.open_sheet_reader_owned("Sheet1").unwrap();
    let owned_rows = owned.next_batch(1).unwrap();
    assert_eq!(owned_rows, borrowed_rows);
}

#[test]
fn file_backed_source_supports_concurrent_reads_and_saves() {
    let directory = tempfile::TempDir::new().unwrap();
    let source_path = directory.path().join("source.xlsx");

    let mut source = Workbook::new();
    for row in 1..=4_000 {
        source
            .set_cell_value("Sheet1", &format!("A{row}"), format!("source-row-{row:08}"))
            .unwrap();
    }
    source.save(&source_path).unwrap();

    let workbook = Arc::new(
        Workbook::open_with_options(
            &source_path,
            &OpenOptions::new().read_mode(ReadMode::Stream),
        )
        .unwrap(),
    );
    let barrier = Arc::new(Barrier::new(3));

    std::thread::scope(|scope| {
        for _ in 0..2 {
            let workbook = Arc::clone(&workbook);
            let barrier = Arc::clone(&barrier);
            scope.spawn(move || {
                barrier.wait();
                for _ in 0..12 {
                    let mut reader = workbook.open_sheet_reader_owned("Sheet1").unwrap();
                    let rows = reader.next_batch(MAX_STREAM_BATCH_ROWS).unwrap();
                    assert_eq!(rows.len(), 4_000);
                    assert_eq!(
                        rows[3_999].cells[0].1,
                        CellValue::String("source-row-00004000".to_string())
                    );
                }
            });
        }

        barrier.wait();
        for _ in 0..12 {
            let output = workbook.save_to_buffer().unwrap();
            let round_trip = Workbook::open_from_buffer(&output).unwrap();
            assert_eq!(
                round_trip.get_cell_value("Sheet1", "A4000").unwrap(),
                CellValue::String("source-row-00004000".to_string())
            );
        }
    });
}

#[test]
fn filtered_aux_only_mutation_preserves_original_cells() {
    let input = two_sheet_package();
    let mut workbook = Workbook::open_from_buffer_with_options(
        &input,
        &OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .sheets(vec!["Sheet1".to_string()]),
    )
    .unwrap();

    workbook
        .add_comment(
            "Excluded",
            &CommentConfig {
                cell: "A1".to_string(),
                author: "Security regression".to_string(),
                text: "Preserve original cells".to_string(),
            },
        )
        .unwrap();

    let output = workbook.save_to_buffer().unwrap();
    let mut round_trip = Workbook::open_from_buffer(&output).unwrap();
    assert_eq!(
        round_trip.get_cell_value("Excluded", "A1").unwrap(),
        CellValue::String("Original excluded".to_string())
    );
    assert_eq!(round_trip.get_comments("Excluded").unwrap().len(), 1);
}

#[test]
fn lazy_read_then_mutation_saves_current_cells() {
    let input = two_sheet_package();
    let mut workbook = Workbook::open_from_buffer_with_options(
        &input,
        &OpenOptions::new().read_mode(ReadMode::Lazy),
    )
    .unwrap();
    assert_eq!(
        workbook.get_cell_value("Sheet1", "A1").unwrap(),
        CellValue::String("Included".to_string())
    );
    workbook
        .set_cell_value("Sheet1", "A1", "Mutated after lazy read")
        .unwrap();

    let output = workbook.save_to_buffer().unwrap();
    let round_trip = Workbook::open_from_buffer(&output).unwrap();
    assert_eq!(
        round_trip.get_cell_value("Sheet1", "A1").unwrap(),
        CellValue::String("Mutated after lazy read".to_string())
    );
}

#[test]
fn dirty_reader_spool_enforces_exact_xml_byte_limit() {
    let mut source = Workbook::new();
    source.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    let input = source.save_to_buffer().unwrap();

    let mut reference = Workbook::open_from_buffer(&input).unwrap();
    for row in 1..=200 {
        reference
            .set_cell_value("Sheet1", &format!("A{row}"), f64::from(row))
            .unwrap();
    }
    let serialized = zip_part(
        &reference.save_to_buffer().unwrap(),
        "xl/worksheets/sheet1.xml",
    );
    let mut low = u64::try_from(zip_part(&input, "xl/worksheets/sheet1.xml").len()).unwrap();
    let mut high = u64::try_from(serialized.len()).unwrap().saturating_mul(2);
    assert!(open_dirty_reader_with_limit(&input, high, false).is_ok());
    while low < high {
        let middle = low + (high - low) / 2;
        if open_dirty_reader_with_limit(&input, middle, false).is_ok() {
            high = middle;
        } else {
            low = middle + 1;
        }
    }
    let exact = low;
    assert!(open_dirty_reader_with_limit(&input, exact, false).is_ok());

    for owned in [false, true] {
        let limit = exact - 1;
        let error = open_dirty_reader_with_limit(&input, limit, owned).unwrap_err();
        let metadata = error
            .resource_metadata()
            .expect("dirty spool overflow must retain typed metadata");
        assert_eq!(metadata.operation(), ResourceOperation::StreamRead);
        assert_eq!(metadata.category(), ResourceCategory::Streaming);
        assert_eq!(metadata.dimension(), ResourceDimension::XmlBytes);
        assert_eq!(metadata.limit(), limit);
        assert_eq!(metadata.observed(), limit + 1);
        assert_eq!(metadata.stage(), Some(ResourceStage::Serialize));
    }
}
