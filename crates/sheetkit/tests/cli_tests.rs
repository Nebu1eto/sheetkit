#![cfg(feature = "cli")]

use std::io::{Read, Write};
use std::path::PathBuf;
use std::process::Command;

use zip::write::SimpleFileOptions;

fn cli_bin() -> PathBuf {
    // cargo test builds the test binary in the target directory.
    // The CLI binary is built separately with the "cli" feature.
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.pop(); // crates/
    path.pop(); // project root
    path.push("target");
    path.push("debug");
    path.push("sheetkit");
    path
}

fn fixture_path() -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests");
    path.push("fixtures");
    path
}

fn create_fixture() -> PathBuf {
    static INIT: std::sync::OnceLock<PathBuf> = std::sync::OnceLock::new();
    INIT.get_or_init(|| {
        let dir = fixture_path();
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("cli_test.xlsx");
        if path.exists() {
            return path;
        }
        let mut wb = sheetkit::Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Name").unwrap();
        wb.set_cell_value("Sheet1", "B1", "Value").unwrap();
        wb.set_cell_value("Sheet1", "A2", "Alpha").unwrap();
        wb.set_cell_value("Sheet1", "B2", 100.0).unwrap();
        wb.set_cell_value("Sheet1", "A3", "Beta").unwrap();
        wb.set_cell_value("Sheet1", "B3", 200.5).unwrap();
        wb.set_cell_value("Sheet1", "C1", "Active").unwrap();
        wb.set_cell_value("Sheet1", "C2", true).unwrap();
        wb.set_cell_value("Sheet1", "C3", false).unwrap();
        wb.new_sheet("Summary").unwrap();
        wb.set_cell_value("Summary", "A1", "Total").unwrap();
        wb.set_cell_value("Summary", "B1", 300.5).unwrap();
        wb.save(&path).unwrap();
        path
    })
    .clone()
}

fn create_csv_limit_fixture() -> PathBuf {
    static INIT: std::sync::OnceLock<PathBuf> = std::sync::OnceLock::new();
    INIT.get_or_init(|| {
        let path = fixture_path().join("csv-output-limit.xlsx");
        let mut wb = sheetkit::Workbook::new();
        let dangerous = format!("={}", "x".repeat(32_766));
        for column in 1..=2_049 {
            let cell = sheetkit::utils::coordinates_to_cell_name(column, 1).unwrap();
            wb.set_cell_value("Sheet1", &cell, dangerous.as_str())
                .unwrap();
        }
        wb.save(&path).unwrap();
        path
    })
    .clone()
}

fn create_table_limit_fixture() -> PathBuf {
    static INIT: std::sync::OnceLock<PathBuf> = std::sync::OnceLock::new();
    INIT.get_or_init(|| {
        let path = fixture_path().join("table-output-limit.xlsx");
        let mut wb = sheetkit::Workbook::new();
        for row in 1..=4_097 {
            wb.set_cell_value("Sheet1", &format!("A{row}"), "first")
                .unwrap();
            wb.set_cell_value("Sheet1", &format!("XFD{row}"), "last")
                .unwrap();
        }
        wb.save(&path).unwrap();
        path
    })
    .clone()
}

fn create_terminal_output_fixture() -> (tempfile::TempDir, PathBuf, String, String) {
    let dir = tempfile::TempDir::new().unwrap();
    let path = dir.path().join("terminal-output.xlsx");
    let sheet_name = "보고서\u{009b}31m".to_string();
    let cell_value = "한글 값\u{009b}31m\u{007f}".to_string();
    let mut wb = sheetkit::Workbook::new();
    wb.new_sheet("SheetPlaceholder").unwrap();
    wb.set_cell_value("SheetPlaceholder", "A1", "CellPlaceholder")
        .unwrap();
    wb.set_doc_props(sheetkit::DocProperties {
        title: Some("TitlePlaceholder".to_string()),
        ..Default::default()
    });
    wb.save(&path).unwrap();
    rewrite_xml_controls(&path);
    (dir, path, sheet_name, cell_value)
}

fn rewrite_xml_controls(path: &std::path::Path) {
    let input = std::fs::File::open(path).unwrap();
    let mut archive = zip::ZipArchive::new(input).unwrap();
    let mut entries = Vec::with_capacity(archive.len());
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        let name = entry.name().to_string();
        if matches!(
            name.as_str(),
            "xl/workbook.xml" | "xl/sharedStrings.xml" | "docProps/core.xml"
        ) {
            let xml = String::from_utf8(bytes).unwrap();
            bytes = xml
                .replace("SheetPlaceholder", "보고서&#x9B;31m")
                .replace("CellPlaceholder", "한글 값&#x9B;31m&#x7F;")
                .replace("TitlePlaceholder", "한글 문서&#x85;title")
                .into_bytes();
        }
        entries.push((name, entry.is_dir(), entry.compression(), bytes));
    }
    drop(archive);

    let staged = tempfile::NamedTempFile::new_in(path.parent().unwrap()).unwrap();
    let mut writer = zip::ZipWriter::new(staged);
    for (name, is_dir, compression, bytes) in entries {
        let options = SimpleFileOptions::default().compression_method(compression);
        if is_dir {
            writer.add_directory(name, options).unwrap();
        } else {
            writer.start_file(name, options).unwrap();
            writer.write_all(&bytes).unwrap();
        }
    }
    let staged = writer.finish().unwrap();
    staged.persist(path).unwrap();
}

fn run_cli(args: &[&str]) -> std::process::Output {
    Command::new(cli_bin())
        .args(args)
        .output()
        .expect("failed to execute CLI binary")
}

fn stdout(output: &std::process::Output) -> String {
    String::from_utf8_lossy(&output.stdout).to_string()
}

fn stderr(output: &std::process::Output) -> String {
    String::from_utf8_lossy(&output.stderr).to_string()
}

#[test]
fn test_cli_no_args_shows_help() {
    let output = run_cli(&[]);
    // With no subcommand, clap should show an error or help text
    assert!(!output.status.success() || !stdout(&output).is_empty());
}

#[test]
fn test_cli_help_flag() {
    let output = run_cli(&["--help"]);
    let out = stdout(&output);
    assert!(out.contains("sheetkit") || out.contains("Excel"));
}

#[test]
fn test_cli_parser_errors_escape_control_payloads_and_help_stays_on_stdout() {
    let fixture = create_fixture();
    let payload = "bad\u{0007}\u{001b}[31m\u{009b}\u{007f}\nnext";
    let output = run_cli(&["read", fixture.to_str().unwrap(), "--format", payload]);

    assert!(!output.status.success());
    assert!(output.stdout.is_empty());
    let error = stderr(&output);
    for control in ['\u{0007}', '\u{001b}', '\u{009b}', '\u{007f}'] {
        assert!(!error.contains(control), "raw terminal control: {error:?}");
    }
    assert!(!error.contains("\nnext"));
    assert!(error.contains("\\x07"));
    assert!(error.contains("\\x1B"));
    assert!(error.contains("\\x9B"));
    assert!(error.contains("\\x7F"));
    assert!(error.contains("\\x0A"));

    let help = run_cli(&["--help"]);
    assert!(help.status.success());
    assert!(help.stderr.is_empty());
    assert!(stdout(&help).contains("Usage:"));
}

#[test]
fn test_cli_sheets_command() {
    let fixture = create_fixture();
    let output = run_cli(&["sheets", fixture.to_str().unwrap()]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    let out = stdout(&output);
    assert!(out.contains("Sheet1"), "output: {out}");
    assert!(out.contains("Summary"), "output: {out}");
}

#[test]
fn test_cli_info_command() {
    let fixture = create_fixture();
    let output = run_cli(&["info", fixture.to_str().unwrap()]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    let out = stdout(&output);
    assert!(out.contains("Sheet1"), "output: {out}");
    assert!(out.contains("Summary"), "output: {out}");
    assert!(out.contains("2"), "should show sheet count, output: {out}");
}

#[test]
fn test_cli_read_command_default_sheet() {
    let fixture = create_fixture();
    let output = run_cli(&["read", fixture.to_str().unwrap()]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    let out = stdout(&output);
    assert!(out.contains("Name"), "output: {out}");
    assert!(out.contains("Alpha"), "output: {out}");
    assert!(out.contains("100"), "output: {out}");
}

#[test]
fn test_cli_read_command_specific_sheet() {
    let fixture = create_fixture();
    let output = run_cli(&["read", fixture.to_str().unwrap(), "--sheet", "Summary"]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    let out = stdout(&output);
    assert!(out.contains("Total"), "output: {out}");
    assert!(out.contains("300.5"), "output: {out}");
}

#[test]
fn test_cli_get_command() {
    let fixture = create_fixture();
    let output = run_cli(&["get", fixture.to_str().unwrap(), "Sheet1", "A2"]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    let out = stdout(&output).trim().to_string();
    assert_eq!(out, "Alpha");
}

#[test]
fn test_cli_get_command_number() {
    let fixture = create_fixture();
    let output = run_cli(&["get", fixture.to_str().unwrap(), "Sheet1", "B2"]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    let out = stdout(&output).trim().to_string();
    assert_eq!(out, "100");
}

#[test]
fn test_cli_get_command_bool() {
    let fixture = create_fixture();
    let output = run_cli(&["get", fixture.to_str().unwrap(), "Sheet1", "C2"]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    let out = stdout(&output).trim().to_string();
    assert_eq!(out, "TRUE");
}

#[test]
fn test_cli_get_command_empty_cell() {
    let fixture = create_fixture();
    let output = run_cli(&["get", fixture.to_str().unwrap(), "Sheet1", "Z99"]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    let out = stdout(&output).trim().to_string();
    assert!(out.is_empty(), "expected empty, got: {out}");
}

#[test]
fn test_cli_set_command() {
    let fixture = create_fixture();
    let dir = tempfile::TempDir::new().unwrap();
    let output_path = dir.path().join("set_test.xlsx");
    let output = run_cli(&[
        "set",
        fixture.to_str().unwrap(),
        "Sheet1",
        "D1",
        "NewValue",
        "-o",
        output_path.to_str().unwrap(),
    ]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    assert!(output_path.exists());

    // Verify the value was set
    let verify = run_cli(&["get", output_path.to_str().unwrap(), "Sheet1", "D1"]);
    assert!(verify.status.success());
    assert_eq!(stdout(&verify).trim(), "NewValue");
}

#[test]
fn test_cli_convert_csv() {
    let fixture = create_fixture();
    let dir = tempfile::TempDir::new().unwrap();
    let output_path = dir.path().join("output.csv");
    let output = run_cli(&[
        "convert",
        fixture.to_str().unwrap(),
        "-f",
        "csv",
        "-o",
        output_path.to_str().unwrap(),
    ]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    assert!(output_path.exists());
    let content = std::fs::read_to_string(&output_path).unwrap();
    assert!(content.contains("Name"), "csv: {content}");
    assert!(content.contains("Alpha"), "csv: {content}");
    assert!(content.contains(","), "csv should have commas: {content}");
}

#[test]
fn test_cli_convert_csv_specific_sheet() {
    let fixture = create_fixture();
    let dir = tempfile::TempDir::new().unwrap();
    let output_path = dir.path().join("summary.csv");
    let output = run_cli(&[
        "convert",
        fixture.to_str().unwrap(),
        "-f",
        "csv",
        "-o",
        output_path.to_str().unwrap(),
        "--sheet",
        "Summary",
    ]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    let content = std::fs::read_to_string(&output_path).unwrap();
    assert!(content.contains("Total"), "csv: {content}");
    assert!(content.contains("300.5"), "csv: {content}");
}

#[test]
fn test_cli_missing_file_error() {
    let output = run_cli(&["info", "/nonexistent/path/file.xlsx"]);
    assert!(!output.status.success());
    let err = stderr(&output);
    assert!(!err.is_empty(), "should have error output");
}

#[test]
fn test_cli_invalid_sheet_error() {
    let fixture = create_fixture();
    let output = run_cli(&["read", fixture.to_str().unwrap(), "--sheet", "NoSuchSheet"]);
    assert!(!output.status.success());
}

#[test]
fn test_cli_invalid_cell_ref_error() {
    let fixture = create_fixture();
    let output = run_cli(&["get", fixture.to_str().unwrap(), "Sheet1", "INVALID!!!"]);
    assert!(!output.status.success());
}

#[test]
fn test_cli_read_csv_format() {
    let fixture = create_fixture();
    let output = run_cli(&["read", fixture.to_str().unwrap(), "--format", "csv"]);
    assert!(output.status.success(), "stderr: {}", stderr(&output));
    let out = stdout(&output);
    // CSV format uses commas
    assert!(out.contains(","), "output: {out}");
    assert!(out.contains("Name"), "output: {out}");
}

#[test]
fn test_cli_read_csv_is_spreadsheet_safe_by_default() {
    let dir = tempfile::TempDir::new().unwrap();
    let path = dir.path().join("csv-security.xlsx");
    let mut wb = sheetkit::Workbook::new();
    wb.set_cell_value("Sheet1", "A1", "=command").unwrap();
    wb.set_cell_value("Sheet1", "B1", -42.0).unwrap();
    wb.set_cell_value(
        "Sheet1",
        "C1",
        sheetkit::CellValue::Formula {
            expr: "1+1".to_string(),
            result: None,
        },
    )
    .unwrap();
    wb.set_cell_value(
        "Sheet1",
        "D1",
        sheetkit::CellValue::Formula {
            expr: "2+2".to_string(),
            result: Some(Box::new(sheetkit::CellValue::Number(4.0))),
        },
    )
    .unwrap();
    wb.save(&path).unwrap();

    let output = run_cli(&["read", path.to_str().unwrap(), "--format", "csv"]);

    assert!(output.status.success(), "stderr: {}", stderr(&output));
    assert_eq!(stdout(&output), "'=command,-42,'=1+1,'=2+2");
}

#[test]
fn test_cli_read_csv_raw_exchange_requires_explicit_policy() {
    let dir = tempfile::TempDir::new().unwrap();
    let path = dir.path().join("csv-raw-exchange.xlsx");
    let mut wb = sheetkit::Workbook::new();
    wb.set_cell_value("Sheet1", "A1", "=command").unwrap();
    wb.save(&path).unwrap();

    let output = run_cli(&[
        "read",
        path.to_str().unwrap(),
        "--format",
        "csv",
        "--csv-formula-policy",
        "raw-exchange",
    ]);

    assert!(output.status.success(), "stderr: {}", stderr(&output));
    assert_eq!(stdout(&output), "=command");
}

#[test]
fn test_cli_read_csv_limit_rejection_has_empty_stdout() {
    let path = create_csv_limit_fixture();

    let output = run_cli(&["read", path.to_str().unwrap(), "--format", "csv"]);

    assert!(!output.status.success());
    assert!(output.stdout.is_empty());
    let error = stderr(&output);
    assert!(error.contains("resource limit exceeded"), "stderr: {error}");
    assert!(!error.contains("xxxxxxxx"), "stderr echoed cell content");
}

#[test]
fn test_cli_read_repeated_sst_materialization_is_bounded_and_atomic() {
    let path = create_csv_limit_fixture();

    for format in ["json", "table"] {
        let output = run_cli(&["read", path.to_str().unwrap(), "--format", format]);
        assert!(
            !output.status.success(),
            "format unexpectedly passed: {format}"
        );
        assert!(output.stdout.is_empty());
        let error = stderr(&output);
        assert!(error.contains("resource limit exceeded"), "stderr: {error}");
        assert!(error.contains("retained_bytes"), "stderr: {error}");
        assert!(!error.contains("xxxxxxxx"), "stderr echoed cell content");
    }
}

#[test]
fn test_cli_read_table_limit_rejection_is_atomic() {
    let path = create_table_limit_fixture();

    let output = run_cli(&["read", path.to_str().unwrap(), "--format", "table"]);

    assert!(!output.status.success());
    assert!(output.stdout.is_empty());
    let error = stderr(&output);
    assert!(error.contains("resource limit exceeded"), "stderr: {error}");
    assert!(error.contains("output_bytes"), "stderr: {error}");
    assert!(!error.contains("first"), "stderr echoed cell content");
}

#[test]
fn test_cli_convert_csv_is_safe_by_default_and_raw_is_explicit() {
    let dir = tempfile::TempDir::new().unwrap();
    let workbook_path = dir.path().join("convert-policy.xlsx");
    let safe_path = dir.path().join("safe.csv");
    let raw_path = dir.path().join("raw.csv");
    let mut wb = sheetkit::Workbook::new();
    wb.set_cell_value("Sheet1", "A1", "=command").unwrap();
    wb.save(&workbook_path).unwrap();
    std::fs::write(&safe_path, b"old destination").unwrap();

    let safe = run_cli(&[
        "convert",
        workbook_path.to_str().unwrap(),
        "--format",
        "csv",
        "--output",
        safe_path.to_str().unwrap(),
    ]);
    let raw = run_cli(&[
        "convert",
        workbook_path.to_str().unwrap(),
        "--format",
        "csv",
        "--output",
        raw_path.to_str().unwrap(),
        "--csv-formula-policy",
        "raw-exchange",
    ]);

    assert!(safe.status.success(), "stderr: {}", stderr(&safe));
    assert!(raw.status.success(), "stderr: {}", stderr(&raw));
    assert_eq!(std::fs::read(&safe_path).unwrap(), b"'=command");
    assert_eq!(std::fs::read(&raw_path).unwrap(), b"=command");
}

#[test]
fn test_cli_convert_csv_failure_preserves_destination_bytes() {
    let workbook_path = create_csv_limit_fixture();
    let dir = tempfile::TempDir::new().unwrap();
    let output_path = dir.path().join("existing.csv");
    let original = b"existing destination bytes\n";
    std::fs::write(&output_path, original).unwrap();

    let output = run_cli(&[
        "convert",
        workbook_path.to_str().unwrap(),
        "--format",
        "csv",
        "--output",
        output_path.to_str().unwrap(),
    ]);

    assert!(!output.status.success());
    assert!(output.stdout.is_empty());
    assert_eq!(std::fs::read(output_path).unwrap(), original);
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}

#[test]
fn test_cli_human_outputs_escape_terminal_controls_and_preserve_unicode() {
    let (_dir, path, sheet_name, _) = create_terminal_output_fixture();
    let path = path.to_str().unwrap();

    let info = run_cli(&["info", path]);
    let sheets = run_cli(&["sheets", path]);
    let read = run_cli(&["read", path, "--sheet", &sheet_name]);
    let get = run_cli(&["get", path, &sheet_name, "A1"]);

    for output in [&info, &sheets, &read, &get] {
        assert!(output.status.success(), "stderr: {}", stderr(output));
        let output = stdout(output);
        assert!(!output.contains('\u{009b}'), "raw C1 control: {output:?}");
        assert!(!output.contains('\u{0085}'), "raw C1 control: {output:?}");
        assert!(!output.contains('\u{007f}'), "raw DEL control: {output:?}");
    }

    assert!(stdout(&info).contains("한글 문서\\x85title"));
    assert!(stdout(&sheets).contains("보고서\\x9B31m"));
    assert!(stdout(&read).contains("한글 값\\x9B31m\\x7F"));
    assert!(stdout(&get).contains("한글 값\\x9B31m\\x7F"));
}

#[test]
fn test_cli_json_outputs_parse_and_preserve_original_controls() {
    let (_dir, path, sheet_name, cell_value) = create_terminal_output_fixture();
    let path = path.to_str().unwrap();

    let info = run_cli(&["info", path, "--format", "json"]);
    let sheets = run_cli(&["sheets", path, "--format", "json"]);
    let read = run_cli(&["read", "--format", "json", "--sheet", &sheet_name, path]);
    let get = run_cli(&["get", path, &sheet_name, "A1", "--format", "json"]);

    for output in [&info, &sheets, &read, &get] {
        assert!(output.status.success(), "stderr: {}", stderr(output));
        assert!(!output.stdout.windows(2).any(|bytes| bytes == [0xc2, 0x9b]));
        assert!(!output.stdout.windows(2).any(|bytes| bytes == [0xc2, 0x85]));
        assert!(!output.stdout.contains(&0x7f));
    }

    let info: serde_json::Value = serde_json::from_slice(&info.stdout).unwrap();
    let sheets: serde_json::Value = serde_json::from_slice(&sheets.stdout).unwrap();
    let read: serde_json::Value = serde_json::from_slice(&read.stdout).unwrap();
    let get: serde_json::Value = serde_json::from_slice(&get.stdout).unwrap();

    assert_eq!(info["properties"]["title"], "한글 문서\u{0085}title");
    assert!(info["sheets"]
        .as_array()
        .unwrap()
        .contains(&sheet_name.clone().into()));
    assert!(sheets
        .as_array()
        .unwrap()
        .contains(&sheet_name.clone().into()));
    assert_eq!(read[0]["cells"][0]["value"], cell_value);
    assert_eq!(get, cell_value);
}
