//! Regression tests for row/column structural-edit audit findings.

use std::panic::{catch_unwind, AssertUnwindSafe};

use sheetkit_core::cell::CellValue;
use sheetkit_core::comment::CommentConfig;
use sheetkit_core::conditional::{ConditionalFormatRule, ConditionalFormatType};
use sheetkit_core::error::Error;
use sheetkit_core::hyperlink::HyperlinkType;
use sheetkit_core::table::{TableColumn, TableConfig};
use sheetkit_core::threaded_comment::ThreadedCommentInput;
use sheetkit_core::validation::DataValidationConfig;
use sheetkit_core::workbook::Workbook;
use sheetkit_xml::worksheet::WorksheetXml;

fn formula(wb: &Workbook, sheet: &str, cell: &str) -> String {
    match wb.get_cell_value(sheet, cell).unwrap() {
        CellValue::Formula { expr, .. } => expr,
        other => panic!("expected formula at {sheet}!{cell}, got {other:?}"),
    }
}

fn comment(cell: &str, text: &str) -> CommentConfig {
    CommentConfig {
        cell: cell.to_string(),
        author: "Author".to_string(),
        text: text.to_string(),
    }
}

fn comment_refs(wb: &mut Workbook, sheet: &str) -> Vec<(String, String)> {
    let mut comments: Vec<(String, String)> = wb
        .get_comments(sheet)
        .unwrap()
        .into_iter()
        .map(|comment| (comment.cell, comment.text))
        .collect();
    comments.sort();
    comments
}

fn threaded_refs(wb: &Workbook, sheet: &str) -> Vec<(String, String)> {
    let mut comments: Vec<(String, String)> = wb
        .get_threaded_comments(sheet)
        .unwrap()
        .into_iter()
        .map(|comment| (comment.cell_ref, comment.text))
        .collect();
    comments.sort();
    comments
}

fn table(name: &str, range: &str, columns: &[&str]) -> TableConfig {
    TableConfig {
        name: name.into(),
        display_name: name.into(),
        range: range.into(),
        columns: columns
            .iter()
            .map(|column| TableColumn {
                name: (*column).into(),
                totals_row_function: None,
                totals_row_label: None,
            })
            .collect(),
        ..Default::default()
    }
}

#[test]
fn remove_row_invalidates_deleted_cells_and_shrinks_ranges_in_formulas() {
    let mut wb = Workbook::new();
    wb.new_sheet("Other").unwrap();
    wb.set_cell_formula(
        "Sheet1",
        "B1",
        "A5+SUM(A1:A5)+SUM(A5:A9)+SUM(A5:A5)+Sheet1!A5+A6+SUM(A4:A6)",
    )
    .unwrap();
    wb.set_cell_formula("Other", "C1", "Sheet1!A5+SUM(Sheet1!A1:A5)+A5")
        .unwrap();
    wb.set_defined_name("PrintArea", "Sheet1!$A$1:$C$5", None, None)
        .unwrap();
    wb.set_defined_name("Gone", "Sheet1!$A$5", None, None)
        .unwrap();

    wb.remove_row("Sheet1", 5).unwrap();

    assert_eq!(
        formula(&wb, "Sheet1", "B1"),
        "#REF!+SUM(A1:A4)+SUM(A5:A8)+SUM(#REF!)+Sheet1!#REF!+A5+SUM(A4:A5)"
    );
    assert_eq!(
        formula(&wb, "Other", "C1"),
        "Sheet1!#REF!+SUM(Sheet1!A1:A4)+A5"
    );
    assert_eq!(
        wb.get_defined_name("PrintArea", None)
            .unwrap()
            .unwrap()
            .value,
        "Sheet1!$A$1:$C$4"
    );
    assert_eq!(
        wb.get_defined_name("Gone", None).unwrap().unwrap().value,
        "Sheet1!#REF!"
    );
}

#[test]
fn remove_row_drops_or_shrinks_sqref_owners() {
    let mut wb = Workbook::new();
    wb.merge_cells("Sheet1", "A5", "B5").unwrap();
    wb.merge_cells("Sheet1", "A3", "B4").unwrap();
    wb.merge_cells("Sheet1", "D3", "E6").unwrap();
    wb.merge_cells("Sheet1", "C5", "C6").unwrap();
    wb.add_data_validation("Sheet1", &DataValidationConfig::whole_number("A5", 1, 9))
        .unwrap();
    wb.add_data_validation(
        "Sheet1",
        &DataValidationConfig::whole_number("F4:F6 G5", 1, 9),
    )
    .unwrap();
    let rule = ConditionalFormatRule {
        rule_type: ConditionalFormatType::DuplicateValues,
        format: None,
        priority: Some(1),
        stop_if_true: false,
    };
    wb.set_conditional_format("Sheet1", "H5", std::slice::from_ref(&rule))
        .unwrap();
    wb.set_conditional_format("Sheet1", "H1:H9", std::slice::from_ref(&rule))
        .unwrap();
    wb.set_cell_hyperlink(
        "Sheet1",
        "J5",
        HyperlinkType::External("https://example.com/deleted".into()),
        None,
        None,
    )
    .unwrap();
    wb.set_cell_hyperlink(
        "Sheet1",
        "J6",
        HyperlinkType::External("https://example.com/kept".into()),
        None,
        None,
    )
    .unwrap();
    wb.set_auto_filter("Sheet1", "K5:L5").unwrap();

    wb.remove_row("Sheet1", 5).unwrap();

    assert_eq!(
        wb.get_merge_cells("Sheet1").unwrap(),
        vec!["A3:B4".to_string(), "D3:E5".to_string()]
    );
    let validations: Vec<String> = wb
        .get_data_validations("Sheet1")
        .unwrap()
        .into_iter()
        .map(|validation| validation.sqref)
        .collect();
    assert_eq!(validations, vec!["F4:F5".to_string()]);
    let formats: Vec<String> = wb
        .get_conditional_formats("Sheet1")
        .unwrap()
        .into_iter()
        .map(|(sqref, _)| sqref)
        .collect();
    assert_eq!(formats, vec!["H1:H8".to_string()]);
    assert_eq!(
        wb.get_cell_hyperlink("Sheet1", "J5")
            .unwrap()
            .unwrap()
            .link_type,
        HyperlinkType::External("https://example.com/kept".into())
    );
    assert!(wb.get_cell_hyperlink("Sheet1", "J6").unwrap().is_none());

    let reopened = Workbook::open_from_buffer(&wb.save_to_buffer().unwrap()).unwrap();
    assert_eq!(
        reopened.get_merge_cells("Sheet1").unwrap(),
        vec!["A3:B4".to_string(), "D3:E5".to_string()]
    );
    assert_eq!(
        reopened
            .get_cell_hyperlink("Sheet1", "J5")
            .unwrap()
            .unwrap()
            .link_type,
        HyperlinkType::External("https://example.com/kept".into())
    );
}

#[test]
fn remove_row_drops_fully_deleted_auto_filter() {
    let mut wb = Workbook::new();
    wb.set_auto_filter("Sheet1", "K5:L5").unwrap();
    wb.remove_row("Sheet1", 5).unwrap();
    let sheet_xml = sheet_part(&wb.save_to_buffer().unwrap(), "xl/worksheets/sheet1.xml");
    assert!(!sheet_xml.contains("<autoFilter"), "{sheet_xml}");
}

fn sheet_part(bytes: &[u8], name: &str) -> String {
    use std::io::Read;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut part = String::new();
    archive
        .by_name(name)
        .unwrap()
        .read_to_string(&mut part)
        .unwrap();
    part
}

#[test]
fn remove_col_invalidates_deleted_cells_and_shrinks_ranges() {
    let mut wb = Workbook::new();
    wb.set_cell_formula(
        "Sheet1",
        "A3",
        "B1+SUM(A1:B1)+SUM(B:B)+SUM(A:C)+C1+SUM($B$1:$D$1)",
    )
    .unwrap();
    wb.merge_cells("Sheet1", "B5", "B6").unwrap();
    wb.merge_cells("Sheet1", "A8", "C8").unwrap();
    wb.add_data_validation("Sheet1", &DataValidationConfig::whole_number("B10", 1, 9))
        .unwrap();

    wb.remove_col("Sheet1", "B").unwrap();

    assert_eq!(
        formula(&wb, "Sheet1", "A3"),
        "#REF!+SUM(A1:A1)+SUM(#REF!)+SUM(A:B)+B1+SUM($B$1:$C$1)"
    );
    assert_eq!(
        wb.get_merge_cells("Sheet1").unwrap(),
        vec!["A8:B8".to_string()]
    );
    assert!(wb.get_data_validations("Sheet1").unwrap().is_empty());
}

#[test]
fn remove_row_shifts_and_drops_comments() {
    let mut wb = Workbook::new();
    wb.add_comment("Sheet1", &comment("A5", "deleted")).unwrap();
    wb.add_comment("Sheet1", &comment("A6", "t")).unwrap();
    wb.add_comment("Sheet1", &comment("B2", "above")).unwrap();

    wb.remove_row("Sheet1", 5).unwrap();

    assert_eq!(
        comment_refs(&mut wb, "Sheet1"),
        vec![
            ("A5".to_string(), "t".to_string()),
            ("B2".to_string(), "above".to_string()),
        ]
    );

    let mut reopened = Workbook::open_from_buffer(&wb.save_to_buffer().unwrap()).unwrap();
    assert_eq!(
        comment_refs(&mut reopened, "Sheet1"),
        vec![
            ("A5".to_string(), "t".to_string()),
            ("B2".to_string(), "above".to_string()),
        ]
    );
}

#[test]
fn insert_and_column_edits_shift_comments() {
    let mut wb = Workbook::new();
    wb.add_comment("Sheet1", &comment("B2", "b2")).unwrap();
    wb.add_comment("Sheet1", &comment("C3", "c3")).unwrap();

    wb.insert_rows("Sheet1", 3, 2).unwrap();
    assert_eq!(
        comment_refs(&mut wb, "Sheet1"),
        vec![
            ("B2".to_string(), "b2".to_string()),
            ("C5".to_string(), "c3".to_string()),
        ]
    );

    wb.insert_cols("Sheet1", "A", 1).unwrap();
    assert_eq!(
        comment_refs(&mut wb, "Sheet1"),
        vec![
            ("C2".to_string(), "b2".to_string()),
            ("D5".to_string(), "c3".to_string()),
        ]
    );

    wb.remove_col("Sheet1", "C").unwrap();
    assert_eq!(
        comment_refs(&mut wb, "Sheet1"),
        vec![("C5".to_string(), "c3".to_string())]
    );

    wb.set_cell_value("Sheet1", "A1", "source").unwrap();
    wb.duplicate_row("Sheet1", 1).unwrap();
    assert_eq!(
        comment_refs(&mut wb, "Sheet1"),
        vec![("C6".to_string(), "c3".to_string())]
    );
}

#[test]
fn lazily_opened_comments_shift_with_rows() {
    let mut source = Workbook::new();
    source
        .add_comment("Sheet1", &comment("A5", "deleted"))
        .unwrap();
    source.add_comment("Sheet1", &comment("A6", "t")).unwrap();
    let bytes = source.save_to_buffer().unwrap();

    let mut wb = Workbook::open_from_buffer(&bytes).unwrap();
    wb.remove_row("Sheet1", 5).unwrap();
    let mut reopened = Workbook::open_from_buffer(&wb.save_to_buffer().unwrap()).unwrap();
    assert_eq!(
        comment_refs(&mut reopened, "Sheet1"),
        vec![("A5".to_string(), "t".to_string())]
    );
}

#[test]
fn structural_edits_shift_threaded_comments() {
    let mut wb = Workbook::new();
    let input = |text: &str| ThreadedCommentInput {
        author: "Author".into(),
        text: text.into(),
        parent_id: None,
    };
    wb.add_threaded_comment("Sheet1", "A5", &input("deleted"))
        .unwrap();
    wb.add_threaded_comment("Sheet1", "A6", &input("t"))
        .unwrap();

    wb.remove_row("Sheet1", 5).unwrap();
    assert_eq!(
        threaded_refs(&wb, "Sheet1"),
        vec![("A5".to_string(), "t".to_string())]
    );

    wb.insert_cols("Sheet1", "A", 2).unwrap();
    assert_eq!(
        threaded_refs(&wb, "Sheet1"),
        vec![("C5".to_string(), "t".to_string())]
    );

    let reopened = Workbook::open_from_buffer(&wb.save_to_buffer().unwrap()).unwrap();
    assert_eq!(
        threaded_refs(&reopened, "Sheet1"),
        vec![("C5".to_string(), "t".to_string())]
    );
}

#[test]
fn whole_row_and_column_references_shift_on_both_axes() {
    let mut wb = Workbook::new();
    wb.set_cell_formula(
        "Sheet1",
        "Z100",
        "SUM(A:A)+SUM(1:1)+SUM(B2:B3)+SUM($A:$C)+SUM(Sheet1!B:B)+SUM('Sheet1'!$2:3)+LEN(\"A:A\")+LOG10(2)",
    )
    .unwrap();

    wb.insert_cols("Sheet1", "A", 1).unwrap();
    assert_eq!(
        formula(&wb, "Sheet1", "AA100"),
        "SUM(B:B)+SUM(1:1)+SUM(C2:C3)+SUM($B:$D)+SUM(Sheet1!C:C)+SUM('Sheet1'!$2:3)+LEN(\"A:A\")+LOG10(2)"
    );

    wb.insert_rows("Sheet1", 2, 1).unwrap();
    assert_eq!(
        formula(&wb, "Sheet1", "AA101"),
        "SUM(B:B)+SUM(1:1)+SUM(C3:C4)+SUM($B:$D)+SUM(Sheet1!C:C)+SUM('Sheet1'!$3:4)+LEN(\"A:A\")+LOG10(2)"
    );

    wb.remove_row("Sheet1", 1).unwrap();
    assert_eq!(
        formula(&wb, "Sheet1", "AA100"),
        "SUM(B:B)+SUM(#REF!)+SUM(C2:C3)+SUM($B:$D)+SUM(Sheet1!C:C)+SUM('Sheet1'!$2:3)+LEN(\"A:A\")+LOG10(2)"
    );

    wb.remove_col("Sheet1", "B").unwrap();
    assert_eq!(
        formula(&wb, "Sheet1", "Z100"),
        "SUM(#REF!)+SUM(#REF!)+SUM(B2:B3)+SUM($B:$C)+SUM(Sheet1!B:B)+SUM('Sheet1'!$2:3)+LEN(\"A:A\")+LOG10(2)"
    );
}

#[test]
fn whole_sheet_spans_stay_pinned_and_three_dimensional_ranges_are_untouched() {
    let mut wb = Workbook::new();
    wb.new_sheet("Sheet2").unwrap();
    wb.set_cell_formula(
        "Sheet1",
        "B1",
        "SUM(1:1048576)+SUM(A:XFD)+SUM(Sheet1:Sheet2!A1:D4)+SUM(Sheet1:Sheet2!C:C)",
    )
    .unwrap();

    wb.insert_rows("Sheet1", 3, 1).unwrap();
    wb.insert_cols("Sheet1", "C", 1).unwrap();

    assert_eq!(
        formula(&wb, "Sheet1", "B1"),
        "SUM(1:1048576)+SUM(A:XFD)+SUM(Sheet1:Sheet2!A1:D4)+SUM(Sheet1:Sheet2!C:C)"
    );
}

#[test]
fn duplicate_row_shifts_absolute_rows_structurally_and_relative_rows_by_copy() {
    let mut wb = Workbook::new();
    wb.set_cell_formula("Sheet1", "B5", "A$10+A10+Sheet1!A$10+A$2+A2")
        .unwrap();

    wb.duplicate_row("Sheet1", 5).unwrap();

    assert_eq!(formula(&wb, "Sheet1", "B5"), "A$11+A11+Sheet1!A$11+A$2+A2");
    assert_eq!(formula(&wb, "Sheet1", "B6"), "A$11+A11+Sheet1!A$11+A$2+A3");
}

#[test]
fn structural_edits_shift_and_drop_row_page_breaks() {
    let mut wb = Workbook::new();
    wb.insert_page_break("Sheet1", 10).unwrap();
    wb.insert_page_break("Sheet1", 20).unwrap();

    wb.insert_rows("Sheet1", 1, 5).unwrap();
    assert_eq!(wb.get_page_breaks("Sheet1").unwrap(), vec![15, 25]);

    wb.remove_row("Sheet1", 15).unwrap();
    assert_eq!(wb.get_page_breaks("Sheet1").unwrap(), vec![24]);

    wb.remove_row("Sheet1", 1).unwrap();
    assert_eq!(wb.get_page_breaks("Sheet1").unwrap(), vec![23]);

    wb.set_cell_value("Sheet1", "A2", "source").unwrap();
    wb.duplicate_row("Sheet1", 2).unwrap();
    assert_eq!(wb.get_page_breaks("Sheet1").unwrap(), vec![24]);
}

#[test]
fn insert_cols_with_oversized_col_definition_does_not_overflow() {
    let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cols><col min="1" max="4294967295" width="10"/><col min="16384" max="16384" width="5"/></cols><sheetData/></worksheet>"#;
    let mut ws: WorksheetXml = quick_xml::de::from_str(xml).unwrap();

    let result = catch_unwind(AssertUnwindSafe(|| {
        sheetkit_core::col::insert_cols(&mut ws, "A", 1)
    }))
    .expect("insert_cols must not panic");
    result.unwrap();

    let cols = &ws.cols.as_ref().unwrap().cols;
    assert_eq!(cols.len(), 1);
    assert_eq!((cols[0].min, cols[0].max), (2, 16_384));
}

#[test]
fn col_definitions_with_invalid_bounds_are_rejected() {
    for (min, max) in [(0, 3), (5, 2), (16_385, 16_390)] {
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cols><col min="{min}" max="{max}" width="10"/></cols><sheetData/></worksheet>"#
        );
        let mut ws: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        let before = ws.clone();
        assert!(matches!(
            sheetkit_core::col::insert_cols(&mut ws, "A", 1),
            Err(Error::InvalidColumnNumber(_))
        ));
        assert!(matches!(
            sheetkit_core::col::remove_col(&mut ws, "A"),
            Err(Error::InvalidColumnNumber(_))
        ));
        assert_eq!(ws, before);
    }
}

#[test]
fn remove_col_inside_table_drops_matching_table_column() {
    let mut wb = Workbook::new();
    wb.add_table(
        "Sheet1",
        &table("Sales", "A1:C4", &["Name", "Region", "Total"]),
    )
    .unwrap();

    wb.remove_col("Sheet1", "B").unwrap();

    let tables = wb.get_tables("Sheet1").unwrap();
    assert_eq!(tables[0].range, "A1:B4");
    assert_eq!(
        tables[0].columns,
        vec!["Name".to_string(), "Total".to_string()]
    );
    Workbook::open_from_buffer(&wb.save_to_buffer().unwrap()).unwrap();
}

#[test]
fn deleting_table_header_row_is_rejected_without_mutation() {
    let mut wb = Workbook::new();
    wb.add_table("Sheet1", &table("Sales", "A2:B4", &["Name", "Total"]))
        .unwrap();
    wb.set_cell_formula("Sheet1", "D1", "A3").unwrap();

    let error = wb.remove_row("Sheet1", 2).unwrap_err();
    assert!(matches!(error, Error::InvalidArgument(_)), "{error:?}");
    assert_eq!(wb.get_tables("Sheet1").unwrap()[0].range, "A2:B4");
    assert_eq!(formula(&wb, "Sheet1", "D1"), "A3");

    wb.remove_row("Sheet1", 3).unwrap();
    assert_eq!(wb.get_tables("Sheet1").unwrap()[0].range, "A2:B3");
}

#[test]
fn copy_sheet_table_names_avoid_case_insensitive_and_defined_name_collisions() {
    let mut wb = Workbook::new();
    wb.add_table("Sheet1", &table("Sales", "A1:B3", &["Name", "Total"]))
        .unwrap();
    wb.set_defined_name("SALES_COPY", "Sheet1!$A$1", None, None)
        .unwrap();

    wb.copy_sheet("Sheet1", "Copied").unwrap();

    assert_eq!(wb.get_tables("Copied").unwrap()[0].name, "Sales_Copy2");
}

#[test]
fn copy_sheet_rejects_invalid_generated_table_name_without_mutation() {
    let mut wb = Workbook::new();
    let long_name = format!("T{}", "x".repeat(253));
    wb.add_table("Sheet1", &table(&long_name, "A1:B3", &["Name", "Total"]))
        .unwrap();

    let error = wb.copy_sheet("Sheet1", "Copied").unwrap_err();
    assert!(matches!(error, Error::InvalidArgument(_)), "{error:?}");
    assert_eq!(wb.sheet_names(), vec!["Sheet1"]);
}

fn rewrite_part(bytes: &[u8], name: &str, edit: impl Fn(&str) -> String) -> Vec<u8> {
    use std::io::{Read, Write};
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut out = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let entry_name = entry.name().to_string();
        let mut data = Vec::new();
        entry.read_to_end(&mut data).unwrap();
        if entry_name == name {
            data = edit(std::str::from_utf8(&data).unwrap()).into_bytes();
        }
        out.start_file(entry_name, options).unwrap();
        out.write_all(&data).unwrap();
    }
    out.finish().unwrap().into_inner()
}

#[test]
fn copy_sheet_preserves_table_sort_state_from_lazy_part() {
    let mut source = Workbook::new();
    source
        .add_table("Sheet1", &table("Sales", "A1:B3", &["Name", "Total"]))
        .unwrap();
    let bytes = rewrite_part(
        &source.save_to_buffer().unwrap(),
        "xl/tables/table1.xml",
        |xml| {
            xml.replacen(
                "<tableColumns",
                "<sortState ref=\"A2:B3\"><sortCondition ref=\"A2:A3\"/></sortState><tableColumns",
                1,
            )
        },
    );
    assert!(sheet_part(&bytes, "xl/tables/table1.xml").contains("sortState"));

    let mut wb = Workbook::open_from_buffer(&bytes).unwrap();
    wb.copy_sheet("Sheet1", "Copied").unwrap();
    let saved = wb.save_to_buffer().unwrap();

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(saved.as_slice())).unwrap();
    let table_parts: Vec<String> = (0..archive.len())
        .map(|index| archive.by_index(index).unwrap().name().to_string())
        .filter(|name| name.starts_with("xl/tables/table"))
        .collect();
    assert_eq!(table_parts.len(), 2, "{table_parts:?}");
    for part in table_parts {
        assert!(sheet_part(&saved, &part).contains("sortState"), "{part}");
    }
}
