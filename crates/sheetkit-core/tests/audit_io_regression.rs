use sheetkit_core::cell::CellValue;
use sheetkit_core::error::Error;
use sheetkit_core::save::SaveLimits;
use sheetkit_core::workbook::{AuxParts, OpenOptions, ReadMode, Workbook};
use sheetkit_core::ResourceDimension;
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;

const ALL_MODES: [ReadMode; 3] = [ReadMode::Eager, ReadMode::Lazy, ReadMode::Stream];

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

fn rewrite_part(buffer: &[u8], name: &str, rewrite: impl Fn(&[u8]) -> Vec<u8>) -> Vec<u8> {
    let mut parts = parts(buffer);
    let part = parts
        .iter_mut()
        .find(|(part_name, _)| part_name == name)
        .unwrap();
    part.1 = rewrite(&part.1);
    rezip(&parts)
}

fn with_doc_props() -> Vec<u8> {
    let mut parts = parts(&Workbook::new().save_to_buffer().unwrap());
    for (name, bytes) in [
        (
            "docProps/core.xml",
            &b"<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"/>"[..],
        ),
        (
            "docProps/app.xml",
            &b"<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"/>"[..],
        ),
    ] {
        parts.retain(|(part_name, _)| part_name != name);
        parts.push((name.to_string(), bytes.to_vec()));
    }
    rezip(&parts)
}

fn part_bytes(buffer: &[u8], name: &str) -> Vec<u8> {
    parts(buffer)
        .into_iter()
        .find(|(part_name, _)| part_name == name)
        .unwrap()
        .1
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

/// Rewrite the uncompressed size declared for one entry in both headers.
fn declare_size(buffer: &mut [u8], name: &str, declared: u32) {
    let (local, central) = header_positions(buffer, name);
    assert!(!local.is_empty() && !central.is_empty());
    for header in local {
        buffer[header + 22..header + 26].copy_from_slice(&declared.to_le_bytes());
    }
    for header in central {
        buffer[header + 24..header + 28].copy_from_slice(&declared.to_le_bytes());
    }
}

/// Replace the CRC-32 recorded for one entry in both headers.
fn corrupt_crc(buffer: &mut [u8], name: &str) {
    let (local, central) = header_positions(buffer, name);
    assert!(!local.is_empty() && !central.is_empty());
    for header in local {
        buffer[header + 14..header + 18].copy_from_slice(&[1, 2, 3, 4]);
    }
    for header in central {
        buffer[header + 16..header + 20].copy_from_slice(&[1, 2, 3, 4]);
    }
}

fn entry_size(buffer: &[u8], name: &str) -> u64 {
    let mut archive = zip::ZipArchive::new(Cursor::new(buffer)).unwrap();
    let size = archive.by_name(name).unwrap().size();
    size
}

fn large_sheet_workbook(extra_sheet: bool) -> Vec<u8> {
    let mut workbook = Workbook::new();
    for row in 1..=20_000u32 {
        workbook
            .set_cell_value("Sheet1", &format!("A{row}"), CellValue::Number(1.0))
            .unwrap();
    }
    if extra_sheet {
        workbook.new_sheet("Sheet2").unwrap();
    }
    workbook.save_to_buffer().unwrap()
}

fn open(buffer: &[u8], options: &OpenOptions) -> Result<Workbook, Error> {
    Workbook::open_from_buffer_with_options(buffer, options)
}

fn expect_err<T>(result: Result<T, Error>, context: &str) -> Error {
    match result {
        Ok(_) => panic!("{context}: expected an error"),
        Err(error) => error,
    }
}

fn assert_xml_limit(error: Error, limit: u64) {
    let metadata = error
        .resource_metadata()
        .copied()
        .unwrap_or_else(|| panic!("expected resource limit error, got {error}"));
    assert_eq!(metadata.dimension(), ResourceDimension::XmlBytes);
    assert_eq!(metadata.limit(), limit);
    assert!(metadata.observed() > limit);
}

const SHEET1: &str = "xl/worksheets/sheet1.xml";

#[test]
fn io_f1_lazy_clean_save_counts_actual_worksheet_bytes() {
    let mut buffer = large_sheet_workbook(false);
    assert!(entry_size(&buffer, SHEET1) > 200_000);
    declare_size(&mut buffer, SHEET1, 100);

    let workbook = open(&buffer, &OpenOptions::new().read_mode(ReadMode::Lazy)).unwrap();
    let limits = SaveLimits {
        xml_part_bytes: 200_000,
        ..SaveLimits::default()
    };
    let error = expect_err(workbook.save_to_buffer_with_limits(limits), "lazy save");
    assert_xml_limit(error, 200_000);
}

#[test]
fn io_f1_eager_filtered_sheet_save_counts_actual_worksheet_bytes() {
    let mut buffer = large_sheet_workbook(true);
    declare_size(&mut buffer, SHEET1, 100);

    let options = OpenOptions::new()
        .read_mode(ReadMode::Eager)
        .sheets(vec!["Sheet2".to_string()]);
    let workbook = open(&buffer, &options).unwrap();
    let limits = SaveLimits {
        xml_part_bytes: 200_000,
        ..SaveLimits::default()
    };
    let error = expect_err(workbook.save_to_buffer_with_limits(limits), "eager save");
    assert_xml_limit(error, 200_000);
}

#[test]
fn io_f1_declared_size_mismatch_is_rejected_on_every_read_path() {
    let mut buffer = large_sheet_workbook(false);
    declare_size(&mut buffer, SHEET1, 100);

    let error = expect_err(
        open(&buffer, &OpenOptions::new().read_mode(ReadMode::Eager)),
        "eager open",
    );
    assert!(matches!(error, Error::Zip(_)), "unexpected error: {error}");

    let workbook = open(&buffer, &OpenOptions::new().read_mode(ReadMode::Lazy)).unwrap();
    let error = expect_err(workbook.save_to_buffer(), "lazy passthrough save");
    assert!(matches!(error, Error::Zip(_)), "unexpected error: {error}");
    let error = expect_err(workbook.get_cell_value("Sheet1", "A1"), "lazy hydration");
    assert!(matches!(error, Error::Zip(_)), "unexpected error: {error}");
}

#[test]
fn io_f1_lazy_clean_save_accepts_exact_actual_size() {
    let buffer = large_sheet_workbook(false);
    let actual = entry_size(&buffer, SHEET1);
    let workbook = open(&buffer, &OpenOptions::new().read_mode(ReadMode::Lazy)).unwrap();
    let exact = SaveLimits {
        xml_part_bytes: actual,
        ..SaveLimits::default()
    };
    workbook.save_to_buffer_with_limits(exact).unwrap();

    let workbook = open(&buffer, &OpenOptions::new().read_mode(ReadMode::Lazy)).unwrap();
    let below = SaveLimits {
        xml_part_bytes: actual - 1,
        ..SaveLimits::default()
    };
    assert_xml_limit(
        expect_err(workbook.save_to_buffer_with_limits(below), "below"),
        actual - 1,
    );
}

#[cfg(unix)]
#[test]
fn io_f2_save_preserves_existing_mode_and_default_mode() {
    use std::os::unix::fs::PermissionsExt;

    let directory = tempfile::tempdir().unwrap();
    let existing = directory.path().join("existing.xlsx");
    std::fs::write(&existing, b"old").unwrap();
    std::fs::set_permissions(&existing, std::fs::Permissions::from_mode(0o640)).unwrap();
    Workbook::new().save(&existing).unwrap();
    let mode = std::fs::metadata(&existing).unwrap().permissions().mode() & 0o777;
    assert_eq!(mode, 0o640);
    Workbook::open(&existing).unwrap();

    let reference = directory.path().join("reference.bin");
    std::fs::File::create(&reference).unwrap();
    let default_mode = std::fs::metadata(&reference).unwrap().permissions().mode() & 0o777;
    let created = directory.path().join("created.xlsx");
    Workbook::new().save(&created).unwrap();
    let mode = std::fs::metadata(&created).unwrap().permissions().mode() & 0o777;
    assert_eq!(mode, default_mode);
}

#[cfg(unix)]
#[test]
fn io_f2_save_through_symlink_updates_the_link_target() {
    let directory = tempfile::tempdir().unwrap();
    let nested = directory.path().join("nested");
    std::fs::create_dir(&nested).unwrap();
    let target = nested.join("real.xlsx");
    std::fs::write(&target, b"old").unwrap();
    let link = directory.path().join("link.xlsx");
    std::os::unix::fs::symlink("nested/real.xlsx", &link).unwrap();

    let mut workbook = Workbook::new();
    workbook
        .set_cell_value("Sheet1", "A1", CellValue::String("saved".into()))
        .unwrap();
    workbook.save(&link).unwrap();

    assert!(std::fs::symlink_metadata(&link)
        .unwrap()
        .file_type()
        .is_symlink());
    let reopened = Workbook::open(&target).unwrap();
    assert_eq!(
        reopened.get_cell_value("Sheet1", "A1").unwrap(),
        CellValue::String("saved".into())
    );
    let leftovers: Vec<_> = std::fs::read_dir(&nested)
        .unwrap()
        .chain(std::fs::read_dir(directory.path()).unwrap())
        .map(|entry| entry.unwrap().file_name().to_string_lossy().into_owned())
        .filter(|name| name.starts_with(".sheetkit-"))
        .collect();
    assert!(leftovers.is_empty(), "temporary files left: {leftovers:?}");

    let dangling = directory.path().join("dangling.xlsx");
    std::os::unix::fs::symlink("nested/new.xlsx", &dangling).unwrap();
    workbook.save(&dangling).unwrap();
    assert!(std::fs::symlink_metadata(&dangling)
        .unwrap()
        .file_type()
        .is_symlink());
    Workbook::open(nested.join("new.xlsx")).unwrap();
}

fn with_custom_part() -> Vec<u8> {
    let mut parts = parts(&Workbook::new().save_to_buffer().unwrap());
    parts.push((
        "customXml/item1.xml".to_string(),
        b"<item>hello</item>".to_vec(),
    ));
    rezip(&parts)
}

#[test]
fn io_f3_corrupt_unknown_part_fails_open_in_every_mode() {
    let mut buffer = with_custom_part();
    for mode in ALL_MODES {
        open(&buffer, &OpenOptions::new().read_mode(mode)).unwrap();
    }
    corrupt_crc(&mut buffer, "customXml/item1.xml");
    for mode in ALL_MODES {
        let error = expect_err(open(&buffer, &OpenOptions::new().read_mode(mode)), "open");
        assert!(matches!(error, Error::Zip(_)), "{mode:?}: {error}");
    }
}

#[test]
fn io_f3_corrupt_optional_known_parts_fail_open() {
    let base = with_doc_props();
    let eager = OpenOptions::new()
        .read_mode(ReadMode::Eager)
        .aux_parts(AuxParts::EagerLoad);
    open(&base, &eager).unwrap();
    for name in [
        "xl/theme/theme1.xml",
        "docProps/core.xml",
        "docProps/app.xml",
    ] {
        let mut buffer = base.clone();
        corrupt_crc(&mut buffer, name);
        let error = expect_err(open(&buffer, &eager), name);
        assert!(matches!(error, Error::Zip(_)), "{name}: {error}");
    }

    let mut buffer = base.clone();
    corrupt_crc(&mut buffer, "xl/theme/theme1.xml");
    for mode in ALL_MODES {
        let error = expect_err(open(&buffer, &OpenOptions::new().read_mode(mode)), "theme");
        assert!(matches!(error, Error::Zip(_)), "{mode:?}: {error}");
    }
}

#[test]
fn io_f3_corrupt_required_xml_part_fails_open() {
    let base = Workbook::new().save_to_buffer().unwrap();
    let mut buffer = base.clone();
    corrupt_crc(&mut buffer, "xl/styles.xml");
    let error = expect_err(open(&buffer, &OpenOptions::new()), "styles");
    assert!(matches!(error, Error::Zip(_)), "{error}");
}

#[test]
fn io_f3_malformed_but_intact_optional_part_is_still_preserved() {
    let base = with_doc_props();
    let buffer = rewrite_part(&base, "docProps/app.xml", |_| b"<not-closed".to_vec());
    let eager = OpenOptions::new()
        .read_mode(ReadMode::Eager)
        .aux_parts(AuxParts::EagerLoad);
    let workbook = open(&buffer, &eager).unwrap();
    let saved = workbook.save_to_buffer().unwrap();
    assert_eq!(part_bytes(&saved, "docProps/app.xml"), b"<not-closed");
}

fn two_sheet_package() -> Vec<u8> {
    let mut workbook = Workbook::new();
    workbook.new_sheet("Sheet2").unwrap();
    workbook
        .set_cell_value("Sheet1", "A1", CellValue::String("x".into()))
        .unwrap();
    workbook.save_to_buffer().unwrap()
}

#[test]
fn io_f4_worksheet_relationship_aliases_fail_open_in_every_mode() {
    let base = two_sheet_package();
    let rels = "xl/_rels/workbook.xml.rels";
    let cases = [
        ("worksheets/sheet2.xml", "worksheets/sheet1.xml"),
        ("worksheets/sheet2.xml", "/xl/workbook.xml"),
        ("worksheets/sheet2.xml", "styles.xml"),
        ("worksheets/sheet2.xml", "../../../../etc/passwd"),
        ("worksheets/sheet2.xml", "../../xl/worksheets/sheet1.xml"),
        (
            "Target=\"worksheets/sheet2.xml\"",
            "Target=\"http://example.invalid/sheet.xml\" TargetMode=\"External\"",
        ),
        (
            "Target=\"worksheets/sheet2.xml\"",
            "Target=\"worksheets/sheet2.xml\" TargetMode=\"External\"",
        ),
    ];
    for (from, to) in cases {
        let buffer = rewrite_part(&base, rels, |bytes| {
            let text = String::from_utf8(bytes.to_vec()).unwrap();
            assert!(text.contains(from));
            text.replacen(from, to, 1).into_bytes()
        });
        for mode in ALL_MODES {
            let error = expect_err(
                open(&buffer, &OpenOptions::new().read_mode(mode)),
                &format!("{to} {mode:?}"),
            );
            assert!(
                matches!(error, Error::InvalidArgument(_)),
                "{to} {mode:?}: {error}"
            );
        }
    }
}

#[test]
fn io_f4_valid_worksheet_targets_still_open_in_every_mode() {
    let base = two_sheet_package();
    let buffer = rewrite_part(&base, "xl/_rels/workbook.xml.rels", |bytes| {
        String::from_utf8(bytes.to_vec())
            .unwrap()
            .replacen("worksheets/sheet2.xml", "/xl/worksheets/sheet2.xml", 1)
            .into_bytes()
    });
    for mode in ALL_MODES {
        let workbook = open(&buffer, &OpenOptions::new().read_mode(mode)).unwrap();
        workbook.save_to_buffer().unwrap();
    }
}

#[test]
fn io_f5_eager_load_reads_each_worksheet_once() {
    let mut workbook = Workbook::new();
    for row in 1..=20_000u32 {
        workbook
            .set_cell_value("Sheet1", &format!("A{row}"), CellValue::Number(1.0))
            .unwrap();
    }
    workbook
        .add_sparkline(
            "Sheet1",
            &sheetkit_core::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1"),
        )
        .unwrap();
    workbook
        .add_table(
            "Sheet1",
            &sheetkit_core::table::TableConfig {
                name: "Data".to_string(),
                display_name: "Data".to_string(),
                range: "A1:A10".to_string(),
                columns: vec![sheetkit_core::table::TableColumn {
                    name: "Value".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                }],
                ..Default::default()
            },
        )
        .unwrap();
    let buffer = workbook.save_to_buffer().unwrap();

    let total: u64 = {
        let mut archive = zip::ZipArchive::new(Cursor::new(&buffer)).unwrap();
        (0..archive.len())
            .map(|index| archive.by_index(index).unwrap().size())
            .sum()
    };
    let exact = OpenOptions::new()
        .read_mode(ReadMode::Eager)
        .aux_parts(AuxParts::EagerLoad)
        .max_aggregate_actual_bytes(total);
    let opened = open(&buffer, &exact).unwrap();
    assert_eq!(opened.get_sparklines("Sheet1").unwrap().len(), 1);
    assert_eq!(opened.get_tables("Sheet1").unwrap().len(), 1);

    let below = OpenOptions::new()
        .read_mode(ReadMode::Eager)
        .aux_parts(AuxParts::EagerLoad)
        .max_aggregate_actual_bytes(total - 1);
    let error = expect_err(open(&buffer, &below), "below");
    assert_eq!(
        error
            .resource_metadata()
            .map(|metadata| metadata.dimension()),
        Some(ResourceDimension::AggregateActualBytes)
    );
}

fn padded_package(value: f64, padding_entries: usize) -> Vec<u8> {
    let mut workbook = Workbook::new();
    workbook
        .set_cell_value("Sheet1", "A1", CellValue::Number(value))
        .unwrap();
    let mut parts = parts(&workbook.save_to_buffer().unwrap());
    for index in 0..padding_entries {
        parts.insert(0, (format!("customXml/pad{index}.xml"), b"<p/>".to_vec()));
    }
    rezip(&parts)
}

#[test]
fn io_f6_lazy_reads_do_not_reparse_a_rewritten_package() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("book.xlsx");
    std::fs::write(&path, padded_package(1.0, 0)).unwrap();

    let options = OpenOptions::new()
        .read_mode(ReadMode::Lazy)
        .max_zip_entries(20);
    let workbook = Workbook::open_with_options(&path, &options).unwrap();

    // Rewrite the same inode in place with a different, over-limit package.
    let replacement = padded_package(2.0, 64);
    let mut file = std::fs::OpenOptions::new()
        .write(true)
        .truncate(true)
        .open(&path)
        .unwrap();
    file.write_all(&replacement).unwrap();
    file.sync_all().unwrap();
    drop(file);

    let value = workbook.get_cell_value("Sheet1", "A1");
    assert!(
        !matches!(&value, Ok(CellValue::Number(number)) if *number == 2.0),
        "lazy read observed the rewritten package: {value:?}"
    );
}

/// Duplicate an entry name by renaming a same-length sibling in both headers.
fn with_duplicate_entry(name: &str, placeholder: &str, contents: &[u8]) -> Vec<u8> {
    assert_eq!(name.len(), placeholder.len());
    let mut parts = parts(&Workbook::new().save_to_buffer().unwrap());
    parts.push((placeholder.to_string(), contents.to_vec()));
    let mut buffer = rezip(&parts);
    let (local, central) = header_positions(&buffer, placeholder);
    for header in local {
        buffer[header + 30..header + 30 + name.len()].copy_from_slice(name.as_bytes());
    }
    for header in central {
        buffer[header + 46..header + 46 + name.len()].copy_from_slice(name.as_bytes());
    }
    buffer
}

#[test]
fn io_f7_duplicate_entry_names_fail_open() {
    let buffer = with_duplicate_entry("xl/workbook.xml", "xl/workbook.xmz", b"<workbook/>");
    for mode in ALL_MODES {
        let error = expect_err(open(&buffer, &OpenOptions::new().read_mode(mode)), "dup");
        assert!(matches!(error, Error::Zip(_)), "{mode:?}: {error}");
    }
}

#[test]
fn io_f7_case_insensitive_duplicate_part_names_fail_open() {
    let buffer = with_duplicate_entry("XL/STYLES.XML", "xl/styles.xmz", b"<styleSheet/>");
    for mode in ALL_MODES {
        let error = expect_err(open(&buffer, &OpenOptions::new().read_mode(mode)), "case");
        assert!(
            matches!(error, Error::InvalidArgument(_)),
            "{mode:?}: {error}"
        );
    }
}
