use std::io::{Cursor, Write};

use zip::result::ZipError;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

/// Build an archive whose second entry is renamed to `renamed` in both the
/// local and central headers after writing. `renamed` must have the same
/// length as `"entry-b"`.
fn archive_with_renamed_second_entry(renamed: &str) -> Vec<u8> {
    assert_eq!(renamed.len(), "entry-b".len());
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, value) in [("entry-a", b"first"), ("entry-b", b"later")] {
        writer
            .start_file(name, SimpleFileOptions::default())
            .unwrap();
        writer.write_all(value).unwrap();
    }
    let mut bytes = writer.finish().unwrap().into_inner();
    let needle = b"entry-b";
    let positions: Vec<usize> = bytes
        .windows(needle.len())
        .enumerate()
        .filter(|(_, window)| window == needle)
        .map(|(position, _)| position)
        .collect();
    assert_eq!(positions.len(), 2);
    for position in positions {
        bytes[position..position + needle.len()].copy_from_slice(renamed.as_bytes());
    }
    bytes
}

#[test]
fn duplicate_entry_names_are_rejected() {
    let error =
        ZipArchive::new(Cursor::new(archive_with_renamed_second_entry("entry-a"))).unwrap_err();
    assert!(
        matches!(error, ZipError::DuplicateEntryName),
        "unexpected error: {error:?}"
    );
}

#[test]
fn distinct_entry_names_are_accepted() {
    let archive =
        ZipArchive::new(Cursor::new(archive_with_renamed_second_entry("entry-c"))).unwrap();
    assert_eq!(archive.len(), 2);
}

#[test]
fn case_variant_names_remain_distinct_zip_entries() {
    let archive =
        ZipArchive::new(Cursor::new(archive_with_renamed_second_entry("ENTRY-A"))).unwrap();
    assert_eq!(archive.len(), 2);
}
