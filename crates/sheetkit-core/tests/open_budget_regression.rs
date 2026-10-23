use sheetkit_core::cell::CellValue;
use sheetkit_core::error::Error;
use sheetkit_core::workbook::{AuxParts, OpenOptions, ReadMode, Workbook};
use sheetkit_core::ResourceDimension;
use std::io::{Read, Write};
use tempfile::NamedTempFile;
use zip::write::SimpleFileOptions;

fn open_error(result: Result<Workbook, Error>) -> Error {
    match result {
        Ok(_) => panic!("expected workbook open to fail"),
        Err(error) => error,
    }
}

#[test]
fn default_buffer_open_rejects_outer_input_over_limit() {
    let data = vec![0_u8; 134_217_729];
    let error = open_error(Workbook::open_from_buffer(&data));
    let metadata = match error {
        Error::ResourceLimitExceeded(metadata) => metadata,
        other => panic!("expected resource limit error, got {other}"),
    };
    assert_eq!(metadata.dimension(), ResourceDimension::InputBytes);
    assert_eq!(metadata.limit(), 134_217_728);
    assert_eq!(metadata.observed(), 134_217_729);
}

#[test]
fn exact_outer_input_override_is_accepted_before_zip_validation() {
    let data = vec![0_u8; 32];
    let options = OpenOptions::new().max_outer_input_bytes(32);
    let error = open_error(Workbook::open_from_buffer_with_options(&data, &options));
    assert!(!matches!(error, Error::ResourceLimitExceeded(_)));
}

#[test]
fn outer_input_override_plus_one_is_rejected() {
    let data = vec![0_u8; 33];
    let options = OpenOptions::new().max_outer_input_bytes(32);
    let error = open_error(Workbook::open_from_buffer_with_options(&data, &options));
    assert_eq!(
        error
            .resource_metadata()
            .map(|metadata| metadata.dimension()),
        Some(ResourceDimension::InputBytes)
    );
}

fn basic_package() -> Vec<u8> {
    Workbook::new().save_to_buffer().unwrap()
}

fn string_package() -> Vec<u8> {
    let mut workbook = Workbook::new();
    let value: String = (0..4_000).map(|index| format!("{index:05x}")).collect();
    workbook
        .set_cell_value("Sheet1", "A1", CellValue::String(value))
        .unwrap();
    workbook.save_to_buffer().unwrap()
}

fn high_ratio_package() -> Vec<u8> {
    let base = basic_package();
    let cursor = std::io::Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let mut source = zip::ZipArchive::new(std::io::Cursor::new(base)).unwrap();
    for index in 0..source.len() {
        let mut entry = source.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        writer.start_file(entry.name(), options).unwrap();
        writer.write_all(&bytes).unwrap();
    }
    writer
        .start_file("attacker-controlled-name.bin", options)
        .unwrap();
    writer.write_all(&vec![b'A'; 32 * 1024]).unwrap();
    writer.finish().unwrap().into_inner()
}

fn package_metrics(data: &[u8]) -> (usize, u64, u64, u64) {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(data)).unwrap();
    let mut largest = 0;
    let mut aggregate = 0_u64;
    let mut largest_xml = 0;
    for index in 0..archive.len() {
        let entry = archive.by_index(index).unwrap();
        largest = largest.max(entry.size());
        aggregate += entry.size();
        if entry.name().ends_with(".xml") || entry.name().ends_with(".rels") {
            largest_xml = largest_xml.max(entry.size());
        }
    }
    (archive.len(), largest, aggregate, largest_xml)
}

fn assert_dimension(result: Result<Workbook, Error>, dimension: ResourceDimension) {
    let error = open_error(result);
    let metadata = error.resource_metadata().expect("typed resource metadata");
    assert_eq!(metadata.dimension(), dimension);
    assert!(!error.to_string().contains("attacker-controlled"));
}

#[test]
fn default_and_exact_archive_boundaries_accept_normal_workbook() {
    let data = basic_package();
    let (entries, largest, aggregate, largest_xml) = package_metrics(&data);
    let options = OpenOptions::new()
        .max_zip_entries(entries)
        .max_entry_declared_bytes(largest)
        .max_entry_actual_bytes(largest)
        .max_aggregate_declared_bytes(aggregate)
        .max_xml_part_bytes(largest_xml);

    Workbook::open_from_buffer(&data).expect("secure defaults accept a normal workbook");
    Workbook::open_from_buffer_with_options(&data, &options)
        .expect("exact metadata boundaries are inclusive");
}

#[test]
fn each_archive_preflight_dimension_rejects_one_over() {
    let data = basic_package();
    let (entries, largest, aggregate, _) = package_metrics(&data);

    assert_dimension(
        Workbook::open_from_buffer_with_options(
            &data,
            &OpenOptions::new().max_zip_entries(entries - 1),
        ),
        ResourceDimension::EntryCount,
    );
    assert_dimension(
        Workbook::open_from_buffer_with_options(
            &data,
            &OpenOptions::new().max_entry_declared_bytes(largest - 1),
        ),
        ResourceDimension::EntryDeclaredBytes,
    );
    assert_dimension(
        Workbook::open_from_buffer_with_options(
            &data,
            &OpenOptions::new().max_aggregate_declared_bytes(aggregate - 1),
        ),
        ResourceDimension::AggregateDeclaredBytes,
    );
}

#[test]
fn actual_and_xml_streaming_violations_remain_typed() {
    let data = basic_package();
    assert_dimension(
        Workbook::open_from_buffer_with_options(
            &data,
            &OpenOptions::new().max_entry_actual_bytes(1),
        ),
        ResourceDimension::EntryActualBytes,
    );
    assert_dimension(
        Workbook::open_from_buffer_with_options(&data, &OpenOptions::new().max_xml_part_bytes(1)),
        ResourceDimension::XmlBytes,
    );
    assert_dimension(
        Workbook::open_from_buffer_with_options(
            &data,
            &OpenOptions::new().max_aggregate_actual_bytes(1),
        ),
        ResourceDimension::AggregateActualBytes,
    );
}

#[test]
fn path_and_buffer_use_the_same_outer_input_boundary() {
    let data = basic_package();
    let mut file = NamedTempFile::new().unwrap();
    file.write_all(&data).unwrap();
    let below = data.len() as u64 - 1;
    let options = OpenOptions::new().max_outer_input_bytes(below);

    assert_dimension(
        Workbook::open_from_buffer_with_options(&data, &options),
        ResourceDimension::InputBytes,
    );
    assert_dimension(
        Workbook::open_with_options(file.path(), &options),
        ResourceDimension::InputBytes,
    );
}

#[test]
fn hard_ceiling_cannot_be_disabled_by_an_override() {
    let data = basic_package();
    let options = OpenOptions::new().max_compression_ratio(1_001);
    assert_dimension(
        Workbook::open_from_buffer_with_options(&data, &options),
        ResourceDimension::CompressionRatio,
    );
}

#[test]
fn high_ratio_entry_rejects_without_echoing_its_name() {
    let data = high_ratio_package();

    assert_dimension(
        Workbook::open_from_buffer_with_options(
            &data,
            &OpenOptions::new().max_compression_ratio(10),
        ),
        ResourceDimension::CompressionRatio,
    );
}

#[test]
fn truncated_package_is_not_reported_as_resource_exhaustion() {
    let mut data = basic_package();
    data.truncate(data.len() / 2);
    let error = open_error(Workbook::open_from_buffer(&data));
    assert!(error.resource_metadata().is_none());
}

#[test]
fn crc_corruption_is_not_reported_as_resource_exhaustion() {
    let mut data = basic_package();
    let (data_start, compressed_size) = {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&data)).unwrap();
        let entry = archive.by_index(0).unwrap();
        (
            entry.data_start() as usize,
            entry.compressed_size() as usize,
        )
    };
    assert!(compressed_size > 2);
    data[data_start + compressed_size / 2] ^= 0x5a;
    let error = open_error(Workbook::open_from_buffer(&data));
    assert!(error.resource_metadata().is_none());
}

#[test]
fn appended_and_zip64_marked_packages_remain_accepted() {
    let base = basic_package();
    let cursor = std::io::Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    let options = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .large_file(true);
    let mut source = zip::ZipArchive::new(std::io::Cursor::new(base)).unwrap();
    for index in 0..source.len() {
        let mut entry = source.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        writer.start_file(entry.name(), options).unwrap();
        writer.write_all(&bytes).unwrap();
    }
    let zip64_marked = writer.finish().unwrap().into_inner();
    Workbook::open_from_buffer(&zip64_marked).expect("ZIP64-marked package opens");

    let mut appended = b"self-extracting-prefix".to_vec();
    appended.extend_from_slice(&zip64_marked);
    Workbook::open_from_buffer(&appended).expect("appended archive offset is detected");
}

#[test]
fn misleading_declared_size_is_rejected_during_preflight() {
    let mut data = basic_package();
    let central = data
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
        .expect("central directory header");
    let misleading = 268_435_457_u32.to_le_bytes();
    data[central + 24..central + 28].copy_from_slice(&misleading);

    assert_dimension(
        Workbook::open_from_buffer(&data),
        ResourceDimension::EntryDeclaredBytes,
    );
}

#[test]
fn single_pass_shared_string_read_has_an_exact_aggregate_actual_boundary() {
    let data = string_package();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&data)).unwrap();
    let mut aggregate = 0_u64;
    let mut has_shared_strings = false;
    for index in 0..archive.len() {
        let entry = archive.by_index(index).unwrap();
        aggregate += entry.size();
        if entry.name() == "xl/sharedStrings.xml" {
            has_shared_strings = true;
        }
    }
    assert!(has_shared_strings);
    let exact = aggregate;
    let eager = |limit| {
        OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad)
            .max_aggregate_actual_bytes(limit)
    };

    Workbook::open_from_buffer_with_options(&data, &eager(exact))
        .expect("exact cumulative bytes, including one SST read, are accepted");
    assert_dimension(
        Workbook::open_from_buffer_with_options(&data, &eager(exact - 1)),
        ResourceDimension::AggregateActualBytes,
    );
}

fn retained_raw_bytes(data: &[u8]) -> u64 {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(data)).unwrap();
    let mut retained = 0_u64;
    for index in 0..archive.len() {
        let entry = archive.by_index(index).unwrap();
        let name = entry.name();
        if name == "xl/workbook.xml"
            || (name.starts_with("xl/worksheets/") && name.ends_with(".xml"))
            || name == "xl/theme/theme1.xml"
            || name.starts_with("docProps/")
        {
            retained += entry.size();
        }
    }
    retained
}

#[test]
fn transient_sst_reads_do_not_consume_retained_raw_budget() {
    let data = string_package();
    let retained = retained_raw_bytes(&data);
    let mut file = NamedTempFile::new().unwrap();
    file.write_all(&data).unwrap();
    let options = OpenOptions::new()
        .read_mode(ReadMode::Eager)
        .aux_parts(AuxParts::EagerLoad)
        .max_retained_package_bytes(retained);

    Workbook::open_with_options(file.path(), &options)
        .expect("large transient shared-string reads are not retained raw bytes");
}

#[test]
fn stored_buffer_source_consumes_retained_raw_budget() {
    let data = basic_package();
    assert_dimension(
        Workbook::open_from_buffer_with_options(
            &data,
            &OpenOptions::new().max_retained_package_bytes(data.len() as u64 - 1),
        ),
        ResourceDimension::RetainedBytes,
    );
}

#[test]
fn stored_unknown_part_consumes_retained_raw_budget() {
    let base = basic_package();
    let retained_without_unknown = retained_raw_bytes(&base);
    let mut base_file = NamedTempFile::new().unwrap();
    base_file.write_all(&base).unwrap();
    let eager_with_retained = |limit| {
        OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad)
            .max_retained_package_bytes(limit)
    };
    let mut low = 1_u64;
    let mut high = retained_without_unknown;
    while low < high {
        let middle = low + (high - low) / 2;
        match Workbook::open_with_options(base_file.path(), &eager_with_retained(middle)) {
            Ok(_) => high = middle,
            Err(error)
                if error
                    .resource_metadata()
                    .map(|metadata| metadata.dimension())
                    == Some(ResourceDimension::RetainedBytes) =>
            {
                low = middle + 1;
            }
            Err(error) => panic!("unexpected base package error: {error}"),
        }
    }
    Workbook::open_with_options(base_file.path(), &eager_with_retained(low))
        .expect("exact retained raw boundary accepts the base package");

    let cursor = std::io::Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    let mut source = zip::ZipArchive::new(std::io::Cursor::new(base)).unwrap();
    for index in 0..source.len() {
        let mut entry = source.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        writer
            .start_file(
                entry.name(),
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated),
            )
            .unwrap();
        writer.write_all(&bytes).unwrap();
    }
    writer
        .start_file(
            "custom/retained.bin",
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored),
        )
        .unwrap();
    writer.write_all(&[0x5a; 512]).unwrap();
    let data = writer.finish().unwrap().into_inner();
    let mut file = NamedTempFile::new().unwrap();
    file.write_all(&data).unwrap();
    assert_dimension(
        Workbook::open_with_options(file.path(), &eager_with_retained(low)),
        ResourceDimension::RetainedBytes,
    );
}

#[test]
fn trusted_override_does_not_change_the_next_default_open() {
    let data = high_ratio_package();
    let trusted = OpenOptions::new().max_compression_ratio(1_000);
    Workbook::open_from_buffer_with_options(&data, &trusted)
        .expect("finite trusted override applies to this call");
    assert_dimension(
        Workbook::open_from_buffer(&data),
        ResourceDimension::CompressionRatio,
    );
}

#[test]
fn optional_part_cannot_swallow_a_sticky_resource_violation() {
    let base = basic_package();
    let (_, largest_required, _, _) = package_metrics(&base);
    let cursor = std::io::Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    let deflated =
        SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let mut source = zip::ZipArchive::new(std::io::Cursor::new(base)).unwrap();
    for index in 0..source.len() {
        let mut entry = source.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        writer.start_file(entry.name(), deflated).unwrap();
        writer.write_all(&bytes).unwrap();
    }
    let oversized = usize::try_from(largest_required + 1).unwrap();
    writer
        .start_file(
            "custom/optional.bin",
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored),
        )
        .unwrap();
    writer.write_all(&vec![0x5a; oversized]).unwrap();
    let data = writer.finish().unwrap().into_inner();

    assert_dimension(
        Workbook::open_from_buffer_with_options(
            &data,
            &OpenOptions::new().max_entry_actual_bytes(largest_required),
        ),
        ResourceDimension::EntryActualBytes,
    );
}
