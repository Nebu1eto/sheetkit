use std::cell::Cell;
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::rc::Rc;

use zip::read::{ArchiveOffset, Config, MAX_ZIP64_EXTENSIBLE_DATA_SIZE};
use zip::result::ZipError;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

fn archive_with_entries(count: usize, zip64: bool, comment: Option<Box<[u8]>>) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    if zip64 {
        writer.set_raw_zip64_comment(Some(Box::new([])));
    }
    if let Some(comment) = comment {
        writer.set_raw_comment(comment);
    }

    for index in 0..count {
        writer
            .start_file(format!("entry-{index}"), SimpleFileOptions::default())
            .unwrap();
        writer.write_all(b"value").unwrap();
    }

    writer.finish().unwrap().into_inner()
}

fn archive_with_zip64_sector(count: usize, sector_size: usize) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    writer.set_raw_zip64_comment(Some(vec![0; sector_size].into_boxed_slice()));
    for index in 0..count {
        writer
            .start_file(format!("entry-{index}"), SimpleFileOptions::default())
            .unwrap();
    }
    writer.finish().unwrap().into_inner()
}

#[derive(Debug)]
struct CountingCursor {
    inner: Cursor<Vec<u8>>,
    bytes_read: Rc<Cell<u64>>,
    largest_read: Rc<Cell<usize>>,
}

impl CountingCursor {
    fn new(bytes: Vec<u8>) -> (Self, Rc<Cell<u64>>, Rc<Cell<usize>>) {
        let bytes_read = Rc::new(Cell::new(0));
        let largest_read = Rc::new(Cell::new(0));
        (
            Self {
                inner: Cursor::new(bytes),
                bytes_read: Rc::clone(&bytes_read),
                largest_read: Rc::clone(&largest_read),
            },
            bytes_read,
            largest_read,
        )
    }
}

impl Read for CountingCursor {
    fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
        self.largest_read
            .set(self.largest_read.get().max(buffer.len()));
        let count = self.inner.read(buffer)?;
        self.bytes_read
            .set(self.bytes_read.get().saturating_add(count as u64));
        Ok(count)
    }
}

impl Seek for CountingCursor {
    fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
        self.inner.seek(position)
    }
}

fn limited_config(limit: u64) -> Config {
    Config {
        max_entries: Some(limit),
        ..Config::default()
    }
}

fn assert_entry_limit(error: ZipError, count: u64, limit: u64) {
    let error_debug = format!("{error:?}");
    assert!(
        matches!(
            error,
            ZipError::EntryCountLimitExceeded {
                count: actual_count,
                limit: actual_limit,
            } if actual_count == count && actual_limit == limit
        ),
        "unexpected error: {error_debug}"
    );
}

fn signature_position(bytes: &[u8], signature: [u8; 4]) -> usize {
    bytes
        .windows(signature.len())
        .rposition(|window| window == signature)
        .unwrap()
}

fn zip32_eocd_position(bytes: &[u8]) -> usize {
    signature_position(bytes, 0x0605_4b50u32.to_le_bytes())
}

fn zip64_eocd_position(bytes: &[u8]) -> usize {
    signature_position(bytes, 0x0606_4b50u32.to_le_bytes())
}

fn is_entry_limit(error: &ZipError) -> bool {
    matches!(error, ZipError::EntryCountLimitExceeded { .. })
}

#[test]
fn default_config_does_not_limit_entries() {
    let archive = ZipArchive::new(Cursor::new(archive_with_entries(2, false, None))).unwrap();
    assert_eq!(archive.len(), 2);
}

#[test]
fn normal_eocd_accepts_exact_limit_and_rejects_one_more() {
    let exact = ZipArchive::with_config(
        limited_config(2),
        Cursor::new(archive_with_entries(2, false, None)),
    )
    .unwrap();
    assert_eq!(exact.len(), 2);

    let error = ZipArchive::with_config(
        limited_config(1),
        Cursor::new(archive_with_entries(2, false, None)),
    )
    .unwrap_err();
    assert_entry_limit(error, 2, 1);
}

#[test]
fn zip64_eocd_enforces_the_limit() {
    let error = ZipArchive::with_config(
        limited_config(1),
        Cursor::new(archive_with_entries(2, true, None)),
    )
    .unwrap_err();
    assert_entry_limit(error, 2, 1);
}

#[test]
fn zip64_entry_limit_is_checked_before_reading_the_extensible_sector() {
    let (reader, bytes_read, largest_read) =
        CountingCursor::new(archive_with_zip64_sector(2, 1024 * 1024));
    let error = ZipArchive::with_config(limited_config(1), reader).unwrap_err();
    assert_entry_limit(error, 2, 1);
    assert!(
        largest_read.get() < 256 * 1024,
        "requested a {}-byte read after reading {} bytes",
        largest_read.get(),
        bytes_read.get(),
    );
    assert!(
        bytes_read.get() < 256 * 1024,
        "read {} bytes before rejecting the entry count",
        bytes_read.get(),
    );
}

#[test]
fn default_config_limits_the_zip64_extensible_sector() {
    let error =
        ZipArchive::new(Cursor::new(archive_with_zip64_sector(1, 1024 * 1024 + 1))).unwrap_err();
    assert!(matches!(error, ZipError::InvalidArchive(_)));
}

#[test]
fn hard_zip64_sector_limit_accepts_exact_size() {
    let archive = ZipArchive::new(Cursor::new(archive_with_zip64_sector(
        1,
        MAX_ZIP64_EXTENSIBLE_DATA_SIZE as usize,
    )))
    .unwrap();
    assert_eq!(archive.len(), 1);
}

#[test]
fn hard_zip64_sector_limit_rejects_larger_override() {
    let config = Config {
        max_zip64_extensible_data_size: Some(MAX_ZIP64_EXTENSIBLE_DATA_SIZE + 1),
        ..Config::default()
    };
    assert!(
        ZipArchive::with_config(config, Cursor::new(archive_with_entries(1, false, None))).is_err()
    );
}

#[test]
fn hard_zip64_sector_limit_cannot_be_disabled() {
    let config = Config {
        max_zip64_extensible_data_size: None,
        ..Config::default()
    };
    assert!(
        ZipArchive::with_config(config, Cursor::new(archive_with_entries(1, false, None))).is_err()
    );
}

#[test]
fn known_archive_offset_enforces_the_limit() {
    let prefix = b"self-extracting-prefix";
    let mut bytes = prefix.to_vec();
    bytes.extend_from_slice(&archive_with_entries(2, false, None));

    let config = Config {
        archive_offset: ArchiveOffset::Known(prefix.len() as u64),
        max_entries: Some(1),
        ..Config::default()
    };
    let error = ZipArchive::with_config(config, Cursor::new(bytes)).unwrap_err();
    assert_entry_limit(error, 2, 1);
}

#[test]
fn detected_archive_offset_enforces_the_limit() {
    let mut bytes = b"self-extracting-prefix".to_vec();
    bytes.extend_from_slice(&archive_with_entries(2, false, None));

    let error = ZipArchive::with_config(limited_config(1), Cursor::new(bytes)).unwrap_err();
    assert_entry_limit(error, 2, 1);
}

#[test]
fn false_eocd_candidate_does_not_replace_the_real_entry_count() {
    let mut false_eocd = Vec::new();
    false_eocd.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
    false_eocd.extend_from_slice(&0u16.to_le_bytes());
    false_eocd.extend_from_slice(&0u16.to_le_bytes());
    false_eocd.extend_from_slice(&9u16.to_le_bytes());
    false_eocd.extend_from_slice(&9u16.to_le_bytes());
    false_eocd.extend_from_slice(&0u32.to_le_bytes());
    false_eocd.extend_from_slice(&0u32.to_le_bytes());
    false_eocd.extend_from_slice(&0u16.to_le_bytes());

    let error = ZipArchive::with_config(
        limited_config(1),
        Cursor::new(archive_with_entries(
            2,
            false,
            Some(false_eocd.into_boxed_slice()),
        )),
    )
    .unwrap_err();
    assert_entry_limit(error, 2, 1);
}

#[test]
fn structurally_plausible_false_limit_candidate_does_not_mask_the_real_count() {
    let archive = archive_with_entries(2, false, None);
    let eocd = zip32_eocd_position(&archive);
    let directory_size = u32::from_le_bytes(archive[eocd + 12..eocd + 16].try_into().unwrap());
    let directory_offset = u32::from_le_bytes(archive[eocd + 16..eocd + 20].try_into().unwrap());

    let padding = vec![0; 100];
    let false_directory_size = directory_size + 22 + padding.len() as u32;
    let mut false_eocd = padding;
    false_eocd.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
    false_eocd.extend_from_slice(&0u16.to_le_bytes());
    false_eocd.extend_from_slice(&0u16.to_le_bytes());
    false_eocd.extend_from_slice(&3u16.to_le_bytes());
    false_eocd.extend_from_slice(&3u16.to_le_bytes());
    false_eocd.extend_from_slice(&false_directory_size.to_le_bytes());
    false_eocd.extend_from_slice(&directory_offset.to_le_bytes());
    false_eocd.extend_from_slice(&0u16.to_le_bytes());

    let error = ZipArchive::with_config(
        limited_config(1),
        Cursor::new(archive_with_entries(
            2,
            false,
            Some(false_eocd.into_boxed_slice()),
        )),
    )
    .unwrap_err();
    assert_entry_limit(error, 2, 1);
}

#[test]
fn false_low_count_eocd_candidate_cannot_truncate_the_archive() {
    let archive = archive_with_entries(2, false, None);
    let eocd = zip32_eocd_position(&archive);
    let directory_size = u32::from_le_bytes(archive[eocd + 12..eocd + 16].try_into().unwrap());
    let directory_offset = u32::from_le_bytes(archive[eocd + 16..eocd + 20].try_into().unwrap());

    let mut false_eocd = Vec::new();
    false_eocd.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
    false_eocd.extend_from_slice(&0u16.to_le_bytes());
    false_eocd.extend_from_slice(&0u16.to_le_bytes());
    false_eocd.extend_from_slice(&1u16.to_le_bytes());
    false_eocd.extend_from_slice(&1u16.to_le_bytes());
    false_eocd.extend_from_slice(&directory_size.to_le_bytes());
    false_eocd.extend_from_slice(&directory_offset.to_le_bytes());
    false_eocd.extend_from_slice(&0u16.to_le_bytes());

    let parsed = ZipArchive::new(Cursor::new(archive_with_entries(
        2,
        false,
        Some(false_eocd.into_boxed_slice()),
    )))
    .unwrap();
    assert_eq!(parsed.len(), 2);
}

#[test]
fn zip32_uses_total_count_and_rejects_single_disk_count_mismatch() {
    let mut bytes = archive_with_entries(2, false, None);
    let eocd = zip32_eocd_position(&bytes);
    bytes[eocd + 8..eocd + 10].copy_from_slice(&1u16.to_le_bytes());

    let error = ZipArchive::with_config(limited_config(1), Cursor::new(bytes)).unwrap_err();
    assert!(!is_entry_limit(&error));
}

#[test]
fn zip64_rejects_single_disk_count_mismatch() {
    let mut bytes = archive_with_entries(2, true, None);
    let eocd64 = zip64_eocd_position(&bytes);
    bytes[eocd64 + 24..eocd64 + 32].copy_from_slice(&1u64.to_le_bytes());

    let error = ZipArchive::with_config(limited_config(1), Cursor::new(bytes)).unwrap_err();
    assert!(!is_entry_limit(&error));
}

#[test]
fn zip64_huge_count_is_rejected_without_platform_truncation() {
    let mut bytes = archive_with_entries(2, true, None);
    let eocd64 = zip64_eocd_position(&bytes);
    bytes[eocd64 + 24..eocd64 + 32].copy_from_slice(&u64::MAX.to_le_bytes());
    bytes[eocd64 + 32..eocd64 + 40].copy_from_slice(&u64::MAX.to_le_bytes());

    let error = ZipArchive::with_config(limited_config(1), Cursor::new(bytes)).unwrap_err();
    assert!(!is_entry_limit(&error));
}

#[test]
fn unsaturated_zip32_low_count_must_match_zip64() {
    let mut bytes = archive_with_entries(2, true, None);
    let eocd = zip32_eocd_position(&bytes);
    bytes[eocd + 8..eocd + 10].copy_from_slice(&1u16.to_le_bytes());
    bytes[eocd + 10..eocd + 12].copy_from_slice(&1u16.to_le_bytes());

    let error = ZipArchive::with_config(limited_config(1), Cursor::new(bytes)).unwrap_err();
    assert!(!is_entry_limit(&error));
}

#[test]
fn unsaturated_zip32_high_count_must_match_zip64() {
    let mut bytes = archive_with_entries(2, true, None);
    let eocd = zip32_eocd_position(&bytes);
    bytes[eocd + 8..eocd + 10].copy_from_slice(&3u16.to_le_bytes());
    bytes[eocd + 10..eocd + 12].copy_from_slice(&3u16.to_le_bytes());

    let error = ZipArchive::with_config(limited_config(1), Cursor::new(bytes)).unwrap_err();
    assert!(!is_entry_limit(&error));
}

#[test]
fn unsaturated_zip32_size_and_offset_must_match_zip64() {
    for field in [12..16, 16..20] {
        let mut bytes = archive_with_entries(2, true, None);
        let eocd = zip32_eocd_position(&bytes);
        let value = u32::from_le_bytes(
            bytes[eocd + field.start..eocd + field.end]
                .try_into()
                .unwrap(),
        );
        bytes[eocd + field.start..eocd + field.end]
            .copy_from_slice(&value.saturating_sub(1).to_le_bytes());

        let error = ZipArchive::with_config(limited_config(1), Cursor::new(bytes)).unwrap_err();
        assert!(!is_entry_limit(&error));
    }
}

#[test]
fn declared_directory_size_must_end_at_the_eocd() {
    let mut bytes = archive_with_entries(2, false, None);
    let eocd = zip32_eocd_position(&bytes);
    let directory_size = u32::from_le_bytes(bytes[eocd + 12..eocd + 16].try_into().unwrap());
    bytes[eocd + 12..eocd + 16].copy_from_slice(&(directory_size - 1).to_le_bytes());

    assert!(ZipArchive::new(Cursor::new(bytes)).is_err());
}

#[test]
fn zip64_declared_directory_size_must_end_at_the_zip64_eocd() {
    let mut bytes = archive_with_entries(2, true, None);
    let eocd64 = zip64_eocd_position(&bytes);
    let directory_size = u64::from_le_bytes(bytes[eocd64 + 40..eocd64 + 48].try_into().unwrap());
    bytes[eocd64 + 40..eocd64 + 48].copy_from_slice(&(directory_size - 1).to_le_bytes());

    assert!(ZipArchive::new(Cursor::new(bytes)).is_err());
}

#[test]
fn nonzero_disk_metadata_is_rejected_before_the_entry_limit() {
    let mut bytes = archive_with_entries(2, false, None);
    let eocd = zip32_eocd_position(&bytes);
    bytes[eocd + 4..eocd + 6].copy_from_slice(&1u16.to_le_bytes());
    bytes[eocd + 6..eocd + 8].copy_from_slice(&1u16.to_le_bytes());

    let error = ZipArchive::with_config(limited_config(1), Cursor::new(bytes)).unwrap_err();
    assert!(!is_entry_limit(&error));
}

#[test]
fn structurally_bounded_over_limit_directory_is_rejected_without_entry_parsing() {
    const CENTRAL_HEADER_SIZE: usize = 46;
    let mut bytes = vec![0; CENTRAL_HEADER_SIZE * 2];
    for header in bytes.chunks_exact_mut(CENTRAL_HEADER_SIZE) {
        header[..4].copy_from_slice(&0x0201_4b50u32.to_le_bytes());
        header[28..30].copy_from_slice(&u16::MAX.to_le_bytes());
    }

    bytes.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&(CENTRAL_HEADER_SIZE as u32 * 2).to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());
    bytes.extend_from_slice(&0u16.to_le_bytes());

    let error = ZipArchive::with_config(limited_config(1), Cursor::new(bytes)).unwrap_err();
    assert_entry_limit(error, 2, 1);
}
