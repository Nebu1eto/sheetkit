use std::io::{Cursor, Read as _};

use sheetkit_core::resource::{
    ResourceCategory, ResourceDimension, ResourceOperation, ResourceStage,
};
use sheetkit_core::save::SaveLimits;
use sheetkit_core::workbook::Workbook;

fn append_xml_part(source: &[u8], name: &str, data: &[u8]) -> Vec<u8> {
    let mut input = zip::ZipArchive::new(Cursor::new(source)).unwrap();
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for index in 0..input.len() {
        let mut entry = input.by_index(index).unwrap();
        let entry_name = entry.name().to_string();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        output.start_file(entry_name, options).unwrap();
        std::io::Write::write_all(&mut output, &bytes).unwrap();
    }
    output.start_file(name, options).unwrap();
    std::io::Write::write_all(&mut output, data).unwrap();
    output.finish().unwrap().into_inner()
}

fn assert_save_limit(
    error: sheetkit_core::error::Error,
    dimension: ResourceDimension,
    limit: u64,
    observed: u64,
) {
    let metadata = error
        .resource_metadata()
        .expect("save limit must preserve typed metadata");
    assert_eq!(metadata.operation(), ResourceOperation::WorkbookSave);
    assert_eq!(metadata.category(), ResourceCategory::Output);
    assert_eq!(metadata.dimension(), dimension);
    assert_eq!(metadata.limit(), limit);
    assert_eq!(metadata.observed(), observed);
}

fn largest_xml_part(bytes: &[u8]) -> u64 {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut largest = 0u64;
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        if !entry.name().ends_with(".xml") && !entry.name().ends_with(".rels") {
            continue;
        }
        let mut content = Vec::new();
        entry.read_to_end(&mut content).unwrap();
        largest = largest.max(u64::try_from(content.len()).unwrap());
    }
    largest
}

#[test]
fn buffer_save_accepts_exact_archive_and_xml_limits_and_rejects_plus_one() {
    let mut workbook = Workbook::new();
    workbook
        .set_cell_value("Sheet1", "A1", "bounded output")
        .unwrap();
    let baseline = workbook.save_to_buffer().unwrap();
    let archive_bytes = u64::try_from(baseline.len()).unwrap();
    let xml_bytes = largest_xml_part(&baseline);

    let exact = SaveLimits {
        output_bytes: archive_bytes,
        xml_part_bytes: xml_bytes,
        ..SaveLimits::default()
    };
    assert_eq!(
        workbook.save_to_buffer_with_limits(exact).unwrap().len(),
        baseline.len()
    );

    let error = workbook
        .save_to_buffer_with_limits(SaveLimits {
            output_bytes: archive_bytes - 1,
            ..exact
        })
        .unwrap_err();
    assert_save_limit(
        error,
        ResourceDimension::OutputBytes,
        archive_bytes - 1,
        archive_bytes,
    );

    let error = workbook
        .save_to_buffer_with_limits(SaveLimits {
            xml_part_bytes: xml_bytes - 1,
            ..exact
        })
        .unwrap_err();
    let metadata = error.resource_metadata().unwrap();
    assert_eq!(metadata.operation(), ResourceOperation::WorkbookSave);
    assert_eq!(metadata.dimension(), ResourceDimension::XmlBytes);
    assert_eq!(metadata.limit(), xml_bytes - 1);
    assert!(metadata.observed() > metadata.limit());
    assert_eq!(metadata.stage(), Some(ResourceStage::Serialize));
}

#[test]
fn save_row_and_cell_limits_are_operation_local_and_destination_atomic() {
    let mut workbook = Workbook::new();
    workbook.set_cell_value("Sheet1", "A1", "one").unwrap();
    workbook.set_cell_value("Sheet1", "B2", "two").unwrap();
    let directory = tempfile::tempdir().unwrap();
    let destination = directory.path().join("limited.xlsx");
    std::fs::write(&destination, b"existing destination").unwrap();

    for (dimension, limits) in [
        (
            ResourceDimension::RowCount,
            SaveLimits {
                row_count: 1,
                ..SaveLimits::default()
            },
        ),
        (
            ResourceDimension::CellCount,
            SaveLimits {
                cell_count: 1,
                ..SaveLimits::default()
            },
        ),
    ] {
        let error = workbook.save_with_limits(&destination, limits).unwrap_err();
        assert_save_limit(error, dimension, 1, 2);
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            b"existing destination"
        );
    }

    let exact = SaveLimits {
        row_count: 2,
        cell_count: 2,
        ..SaveLimits::default()
    };
    workbook.save_with_limits(&destination, exact).unwrap();
    assert_ne!(
        std::fs::read(&destination).unwrap(),
        b"existing destination"
    );
}

#[test]
fn oversized_cell_xml_is_rejected_by_the_streaming_part_serializer() {
    let mut workbook = Workbook::new();
    workbook
        .set_cell_value("Sheet1", "A1", "<&payload>".repeat(512))
        .unwrap();
    let error = workbook
        .save_to_buffer_with_limits(SaveLimits {
            xml_part_bytes: 1024,
            ..SaveLimits::default()
        })
        .unwrap_err();
    let metadata = error.resource_metadata().unwrap();
    assert_eq!(metadata.operation(), ResourceOperation::WorkbookSave);
    assert_eq!(metadata.dimension(), ResourceDimension::XmlBytes);
    assert_eq!(metadata.limit(), 1024);
    assert!(metadata.observed() > 1024);
}

#[test]
fn lazy_package_backed_rows_cells_and_raw_non_sheet_xml_obey_save_limits() {
    let mut source = Workbook::new();
    source.set_cell_value("Sheet1", "A1", "one").unwrap();
    source.set_cell_value("Sheet1", "B2", "two").unwrap();
    let baseline = source.save_to_buffer().unwrap();
    let mut state = 0x1234_5678u32;
    let payload = (0..100_000)
        .map(|_| {
            state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
            format!("{state:08x}")
        })
        .collect::<String>();
    let raw_part = format!("<root>{payload}</root>");
    let package = append_xml_part(&baseline, "customXml/item1.XML", raw_part.as_bytes());
    let opened = Workbook::open_from_buffer_with_options(
        &package,
        &sheetkit_core::workbook::OpenOptions::default(),
    )
    .unwrap();

    for (dimension, limits) in [
        (
            ResourceDimension::RowCount,
            SaveLimits {
                row_count: 1,
                ..SaveLimits::default()
            },
        ),
        (
            ResourceDimension::CellCount,
            SaveLimits {
                cell_count: 1,
                ..SaveLimits::default()
            },
        ),
    ] {
        let error = opened.save_to_buffer_with_limits(limits).unwrap_err();
        assert_save_limit(error, dimension, 1, 2);
    }

    let exact_xml = u64::try_from(raw_part.len()).unwrap();
    opened
        .save_to_buffer_with_limits(SaveLimits {
            xml_part_bytes: exact_xml,
            ..SaveLimits::default()
        })
        .unwrap();
    let error = opened
        .save_to_buffer_with_limits(SaveLimits {
            xml_part_bytes: exact_xml - 1,
            ..SaveLimits::default()
        })
        .unwrap_err();
    assert_save_limit(error, ResourceDimension::XmlBytes, exact_xml - 1, exact_xml);

    let mut excluded = Workbook::open_from_buffer_with_options(
        &baseline,
        &sheetkit_core::workbook::OpenOptions::default().sheets(Vec::new()),
    )
    .unwrap();
    excluded
        .add_comment(
            "Sheet1",
            &sheetkit_core::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Security".to_string(),
                text: "aux-only mutation".to_string(),
            },
        )
        .unwrap();
    let error = excluded
        .save_to_buffer_with_limits(SaveLimits {
            row_count: 1,
            ..SaveLimits::default()
        })
        .unwrap_err();
    assert_save_limit(error, ResourceDimension::RowCount, 1, 2);
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_save_uses_the_same_output_limit_and_preserves_destination() {
    let workbook = Workbook::new();
    let directory = tempfile::tempdir().unwrap();
    let destination = directory.path().join("limited.xlsx");
    std::fs::write(&destination, b"existing destination").unwrap();

    let error = workbook
        .save_with_password_with_output_limits(
            &destination,
            "password",
            sheetkit_core::crypt::CryptLimits::default(),
            SaveLimits {
                output_bytes: 64,
                ..SaveLimits::default()
            },
        )
        .unwrap_err();
    let metadata = error.resource_metadata().unwrap();
    assert_eq!(metadata.operation(), ResourceOperation::WorkbookSave);
    assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
    assert_eq!(
        std::fs::read(&destination).unwrap(),
        b"existing destination"
    );
}
