use sheetkit::Workbook;
use tempfile::TempDir;

#[test]
fn test_create_and_save_empty_workbook() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("empty.xlsx");

    let wb = Workbook::new();
    wb.save(&path).unwrap();

    assert!(path.exists());
    // Verify file is not empty
    let metadata = std::fs::metadata(&path).unwrap();
    assert!(metadata.len() > 0);
}

#[test]
fn test_roundtrip_preserves_sheet_names() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("roundtrip.xlsx");

    let wb1 = Workbook::new();
    assert_eq!(wb1.sheet_names(), vec!["Sheet1"]);
    wb1.save(&path).unwrap();

    let wb2 = Workbook::open(&path).unwrap();
    assert_eq!(wb2.sheet_names(), vec!["Sheet1"]);
}

#[test]
fn test_open_nonexistent_returns_io_error() {
    let result = Workbook::open("/tmp/nonexistent_file_12345.xlsx");
    assert!(result.is_err());
}

#[test]
fn test_workbook_default_trait() {
    let wb = Workbook::default();
    assert_eq!(wb.sheet_names(), vec!["Sheet1"]);
}

#[test]
fn test_public_api_reexports() {
    // Verify all public types are accessible
    let _wb = sheetkit::Workbook::new();
    let _col = sheetkit::utils::column_name_to_number("A").unwrap();
    let _name = sheetkit::utils::column_number_to_name(1).unwrap();
    let _coords = sheetkit::utils::cell_name_to_coordinates("A1").unwrap();
    let _cell = sheetkit::utils::coordinates_to_cell_name(1, 1).unwrap();
}

#[test]
fn test_error_type_accessible() {
    // Verify Error type is accessible through public API
    let err = sheetkit::Error::InvalidCellReference("bad".to_string());
    assert!(err.to_string().contains("bad"));
}

#[test]
fn test_save_and_reopen_multiple_times() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");

    // Save
    let wb = Workbook::new();
    wb.save(&path).unwrap();

    // Open and save again
    let wb2 = Workbook::open(&path).unwrap();
    let path2 = dir.path().join("multi2.xlsx");
    wb2.save(&path2).unwrap();

    // Open the re-saved file
    let wb3 = Workbook::open(&path2).unwrap();
    assert_eq!(wb3.sheet_names(), vec!["Sheet1"]);
}
