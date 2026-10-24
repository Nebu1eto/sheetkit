//! Compression-ratio enforcement must accept ordinary repetitive spreadsheet
//! content while still bounding large high-ratio expansions.

use sheetkit_core::cell::CellValue;
use sheetkit_core::error::Error;
use sheetkit_core::workbook::{OpenLimits, OpenOptions, ReadMode, Workbook};
use sheetkit_core::ResourceDimension;
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;
use zip::CompressionMethod;

const ALL_MODES: [ReadMode; 3] = [ReadMode::Eager, ReadMode::Lazy, ReadMode::Stream];
const MIB: usize = 1024 * 1024;
const PAYLOAD_NAME: &str = "customXml/payload.bin";

/// Workbook whose 2,049 formula-like strings compress far beyond 100:1; the
/// same shape the CLI CSV output-limit test uses.
fn csv_fixture() -> Vec<u8> {
    let mut workbook = Workbook::new();
    let dangerous = format!("={}", "x".repeat(32_766));
    for column in 1..=2_049 {
        let cell = sheetkit_core::utils::cell_ref::coordinates_to_cell_name(column, 1).unwrap();
        workbook
            .set_cell_value("Sheet1", &cell, dangerous.as_str())
            .unwrap();
    }
    workbook.save_to_buffer().unwrap()
}

fn open(buffer: &[u8], options: &OpenOptions) -> Result<Workbook, Error> {
    Workbook::open_from_buffer_with_options(buffer, options)
}

fn expect_dimension(result: Result<Workbook, Error>, dimension: ResourceDimension, context: &str) {
    match result {
        Ok(_) => panic!("{context}: expected {dimension:?} rejection"),
        Err(error) => {
            let metadata = error
                .resource_metadata()
                .unwrap_or_else(|| panic!("{context}: unexpected error {error}"));
            assert_eq!(metadata.dimension(), dimension, "{context}: {error}");
        }
    }
}

fn with_payload(payload: &[u8], method: CompressionMethod) -> Vec<u8> {
    let mut workbook = Workbook::new();
    workbook
        .set_cell_value("Sheet1", "A1", CellValue::String("ok".into()))
        .unwrap();
    let base = workbook.save_to_buffer().unwrap();
    let mut source = zip::ZipArchive::new(Cursor::new(base)).unwrap();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for index in 0..source.len() {
        let mut entry = source.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        writer.start_file(entry.name(), deflated).unwrap();
        writer.write_all(&bytes).unwrap();
    }
    writer
        .start_file(
            PAYLOAD_NAME,
            SimpleFileOptions::default()
                .compression_method(method)
                .large_file(false),
        )
        .unwrap();
    writer.write_all(payload).unwrap();
    writer.finish().unwrap().into_inner()
}

fn entry_sizes(buffer: &[u8], name: &str) -> (u64, u64) {
    let mut archive = zip::ZipArchive::new(Cursor::new(buffer)).unwrap();
    let entry = archive.by_name(name).unwrap();
    (entry.size(), entry.compressed_size())
}

fn header_positions(buffer: &[u8], name: &str) -> (Vec<usize>, Vec<usize>) {
    let needle = name.as_bytes();
    let mut local = Vec::new();
    let mut central = Vec::new();
    for position in 0..buffer.len().saturating_sub(needle.len()) {
        if &buffer[position..position + needle.len()] != needle {
            continue;
        }
        if position >= 30 && &buffer[position - 30..position - 26] == b"PK\x03\x04" {
            local.push(position - 30);
        }
        if position >= 46 && &buffer[position - 46..position - 42] == b"PK\x01\x02" {
            central.push(position - 46);
        }
    }
    (local, central)
}

#[test]
fn library_output_with_repetitive_parts_reopens_at_defaults() {
    let fixture = csv_fixture();
    let saved = Workbook::open_from_buffer(&fixture)
        .unwrap()
        .save_to_buffer()
        .unwrap();
    let max_ratio = {
        let mut archive = zip::ZipArchive::new(Cursor::new(&saved)).unwrap();
        (0..archive.len())
            .map(|index| {
                let entry = archive.by_index(index).unwrap();
                entry.size().div_ceil(entry.compressed_size().max(1))
            })
            .max()
            .unwrap()
    };
    assert!(
        max_ratio > OpenLimits::default().compression_ratio,
        "fixture must exceed the default ratio, observed {max_ratio}"
    );
    for mode in ALL_MODES {
        let options = OpenOptions::new().read_mode(mode);
        let workbook = open(&saved, &options)
            .unwrap_or_else(|error| panic!("{mode:?}: reopen failed: {error}"));
        assert!(!workbook.get_rows("Sheet1").unwrap().is_empty(), "{mode:?}");
    }
}

#[test]
fn high_ratio_entry_at_or_below_floor_is_accepted() {
    let floor = OpenLimits::default().compression_ratio_floor_bytes;
    assert_eq!(floor, 8 * MIB as u64);
    let data = with_payload(&vec![0; 7 * MIB], CompressionMethod::Deflated);
    let (size, compressed) = entry_sizes(&data, PAYLOAD_NAME);
    assert!(size.div_ceil(compressed) > 1_000, "payload ratio too low");
    for mode in ALL_MODES {
        open(&data, &OpenOptions::new().read_mode(mode))
            .unwrap_or_else(|error| panic!("{mode:?}: {error}"));
    }

    let exact = with_payload(&vec![0; 8 * MIB], CompressionMethod::Deflated);
    open(&exact, &OpenOptions::new()).expect("entry exactly at the floor is accepted");
}

#[test]
fn high_ratio_entry_above_floor_is_rejected_unless_overridden() {
    let data = with_payload(&vec![0; 9 * MIB], CompressionMethod::Deflated);
    for mode in ALL_MODES {
        expect_dimension(
            open(&data, &OpenOptions::new().read_mode(mode)),
            ResourceDimension::CompressionRatio,
            &format!("{mode:?} default"),
        );
    }

    let raised_floor = OpenOptions::new().compression_ratio_floor_bytes(16 * MIB as u64);
    open(&data, &raised_floor).expect("raised floor accepts the entry");

    let (size, compressed) = entry_sizes(&data, PAYLOAD_NAME);
    let observed = size.div_ceil(compressed);
    let raised_ratio = OpenOptions::new().max_compression_ratio(observed.min(1_000));
    if observed <= 1_000 {
        open(&data, &raised_ratio).expect("raised ratio accepts the entry");
    } else {
        expect_dimension(
            open(&data, &raised_ratio),
            ResourceDimension::CompressionRatio,
            "ratio ceiling still applies above the floor",
        );
    }

    let strict = OpenOptions::new().compression_ratio_floor_bytes(0);
    let small = with_payload(&vec![0; 64 * 1024], CompressionMethod::Deflated);
    expect_dimension(
        open(&small, &strict),
        ResourceDimension::CompressionRatio,
        "zero floor enforces the ratio for every entry",
    );
}

#[test]
fn large_stored_entry_is_accepted() {
    let data = with_payload(&vec![0; 9 * MIB], CompressionMethod::Stored);
    for mode in ALL_MODES {
        open(&data, &OpenOptions::new().read_mode(mode))
            .unwrap_or_else(|error| panic!("{mode:?}: {error}"));
    }
}

#[test]
fn non_empty_entry_with_zero_compressed_bytes_is_rejected_below_floor() {
    let mut data = with_payload(&[0x5a; 16], CompressionMethod::Stored);
    let (local, central) = header_positions(&data, PAYLOAD_NAME);
    assert!(!local.is_empty() && !central.is_empty());
    for header in local {
        data[header + 18..header + 22].copy_from_slice(&0u32.to_le_bytes());
    }
    for header in central {
        data[header + 20..header + 24].copy_from_slice(&0u32.to_le_bytes());
    }
    for mode in ALL_MODES {
        expect_dimension(
            open(&data, &OpenOptions::new().read_mode(mode)),
            ResourceDimension::CompressionRatio,
            &format!("{mode:?} zero compressed bytes"),
        );
    }
}

#[test]
fn floor_is_validated_against_its_ceiling() {
    let ceiling = OpenLimits::hard_ceiling();
    assert_eq!(
        ceiling.compression_ratio_floor_bytes,
        ceiling.entry_actual_bytes
    );
    let data = with_payload(&[0; 16], CompressionMethod::Deflated);
    open(
        &data,
        &OpenOptions::new().compression_ratio_floor_bytes(ceiling.compression_ratio_floor_bytes),
    )
    .expect("ceiling value is accepted");
    expect_dimension(
        open(
            &data,
            &OpenOptions::new()
                .compression_ratio_floor_bytes(ceiling.compression_ratio_floor_bytes + 1),
        ),
        ResourceDimension::CompressionRatioFloorBytes,
        "floor above ceiling",
    );
}
