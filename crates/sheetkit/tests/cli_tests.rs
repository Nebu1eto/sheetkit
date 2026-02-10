use std::path::PathBuf;
use std::process::Command;

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
