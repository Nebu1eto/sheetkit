use sheetkit_core::cell::CellValue;
use sheetkit_core::output::{serialize_csv, serialize_csv_with_policy, CsvFormulaPolicy};
use sheetkit_core::rich_text::RichTextRun;

fn row(values: Vec<CellValue>) -> Vec<(u32, Vec<(u32, CellValue)>)> {
    vec![(
        1,
        values
            .into_iter()
            .enumerate()
            .map(|(index, value)| (u32::try_from(index + 1).unwrap(), value))
            .collect(),
    )]
}

#[test]
fn spreadsheet_safe_neutralizes_only_textual_formula_candidates() {
    let rows = row(vec![
        CellValue::String("=SUM(A1:A2)".to_string()),
        CellValue::String(" \t@command".to_string()),
        CellValue::Number(-42.0),
        CellValue::Bool(true),
        CellValue::Date(45_292.0),
        CellValue::Error("#N/A".to_string()),
        CellValue::Formula {
            expr: "2+2".to_string(),
            result: None,
        },
        CellValue::Formula {
            expr: "2+2".to_string(),
            result: Some(Box::new(CellValue::Number(4.0))),
        },
        CellValue::Formula {
            expr: "TRUE".to_string(),
            result: Some(Box::new(CellValue::Bool(false))),
        },
        CellValue::Formula {
            expr: "TODAY()".to_string(),
            result: Some(Box::new(CellValue::Date(45_292.0))),
        },
        CellValue::Formula {
            expr: "NA()".to_string(),
            result: Some(Box::new(CellValue::Error("#DIV/0!".to_string()))),
        },
        CellValue::Formula {
            expr: "DANGEROUS()".to_string(),
            result: Some(Box::new(CellValue::String("+command".to_string()))),
        },
        CellValue::RichString(vec![RichTextRun {
            text: "@rich".to_string(),
            font: None,
            size: None,
            bold: false,
            italic: false,
            color: None,
        }]),
    ]);

    let csv = serialize_csv(&rows).unwrap();
    let csv = String::from_utf8(csv).unwrap();

    assert_eq!(
        csv,
        "'=SUM(A1:A2),' \t@command,-42,TRUE,2024-01-01,#N/A,'=2+2,'=2+2,'=TRUE,'=TODAY(),'=NA(),'=DANGEROUS(),'@rich"
    );
}

#[test]
fn spreadsheet_safe_handles_leading_controls_spaces_and_all_formula_starters() {
    let rows = row(vec![
        CellValue::String("=a".to_string()),
        CellValue::String("+a".to_string()),
        CellValue::String("-a".to_string()),
        CellValue::String("@a".to_string()),
        CellValue::String("\t=a".to_string()),
        CellValue::String("\0 +a".to_string()),
        CellValue::String("\u{00a0}-a".to_string()),
        CellValue::String("ordinary".to_string()),
    ]);

    let csv = serialize_csv(&rows).unwrap();
    let csv = String::from_utf8(csv).unwrap();

    assert_eq!(csv, "'=a,'+a,'-a,'@a,'\t=a,'\0 +a,'\u{00a0}-a,ordinary");
}

#[test]
fn raw_exchange_is_explicit_and_preserves_formula_capable_text() {
    let rows = row(vec![
        CellValue::String("=command".to_string()),
        CellValue::Formula {
            expr: "1+1".to_string(),
            result: None,
        },
    ]);

    let csv = serialize_csv_with_policy(&rows, CsvFormulaPolicy::RawExchange).unwrap();

    assert_eq!(csv, b"=command,=1+1");
}

#[test]
fn csv_quotes_fields_and_preserves_unicode_and_embedded_crlf() {
    let rows = row(vec![
        CellValue::String("Korean: \u{c548}\u{b155}".to_string()),
        CellValue::String("comma, quote \" and\r\nnewline".to_string()),
    ]);

    let csv = serialize_csv(&rows).unwrap();

    assert_eq!(
        csv,
        "Korean: \u{c548}\u{b155},\"comma, quote \"\" and\r\nnewline\"".as_bytes()
    );
}

#[test]
fn csv_preserves_canonical_sparse_cells() {
    let rows = vec![(
        1,
        vec![
            (1, CellValue::String("first".to_string())),
            (3, CellValue::String("third".to_string())),
        ],
    )];

    let csv = serialize_csv(&rows).unwrap();

    assert_eq!(csv, b"first,,third");
}

#[test]
fn invalid_direct_and_cached_error_payloads_are_treated_as_text() {
    let rows = row(vec![
        CellValue::Error("=direct".to_string()),
        CellValue::Formula {
            expr: "ERROR.TYPE(A1)".to_string(),
            result: Some(Box::new(CellValue::Error("+cached".to_string()))),
        },
        CellValue::Error("#REF!".to_string()),
    ]);

    let csv = serialize_csv(&rows).unwrap();

    assert_eq!(csv, b"'=direct,'=ERROR.TYPE(A1),#REF!");
}

#[test]
fn csv_rejects_noncanonical_rows_and_duplicate_coordinates() {
    let out_of_order_rows = vec![
        (2, vec![(1, CellValue::String("data".to_string()))]),
        (1, vec![(1, CellValue::String("header".to_string()))]),
    ];
    let duplicate_rows = vec![
        (1, vec![(1, CellValue::String("first".to_string()))]),
        (1, vec![(1, CellValue::String("second".to_string()))]),
    ];
    let duplicate_columns = vec![(
        1,
        vec![
            (1, CellValue::String("first".to_string())),
            (1, CellValue::String("second".to_string())),
        ],
    )];
    let out_of_order_columns = vec![(
        1,
        vec![
            (2, CellValue::String("second".to_string())),
            (1, CellValue::String("first".to_string())),
        ],
    )];

    assert!(serialize_csv(&out_of_order_rows).is_err());
    assert!(serialize_csv(&duplicate_rows).is_err());
    assert!(serialize_csv(&duplicate_columns).is_err());
    assert!(serialize_csv(&out_of_order_columns).is_err());
}

#[test]
fn csv_rejects_out_of_bounds_coordinates() {
    let invalid_row = vec![(0, vec![(1, CellValue::String("value".to_string()))])];
    let invalid_column = vec![(1, vec![(0, CellValue::String("value".to_string()))])];

    assert!(serialize_csv(&invalid_row).is_err());
    assert!(serialize_csv(&invalid_column).is_err());
}
