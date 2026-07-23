use std::panic::{catch_unwind, AssertUnwindSafe};

use quick_xml::events::Event;
use quick_xml::Reader;
use sheetkit_core::error::Error;
use sheetkit_core::render::{render_to_svg, render_to_svg_with_limits, RenderOptions};
use sheetkit_core::resource::{ResourceDimension, ResourceOperation};
use sheetkit_core::sst::SharedStringTable;
use sheetkit_core::style::{add_style, BorderLineStyle, StyleBuilder, StyleColor};
use sheetkit_core::traversal::TraversalLimits;
use sheetkit_xml::shared_strings::{R, T};
use sheetkit_xml::styles::StyleSheet;
use sheetkit_xml::worksheet::{
    Cell, CellFormula, CellTypeTag, CompactCellRef, InlineString, Row, SheetData, WorksheetXml,
};

fn options() -> RenderOptions {
    RenderOptions {
        sheet_name: "Sheet1".to_string(),
        show_headers: false,
        show_gridlines: false,
        ..RenderOptions::default()
    }
}

fn worksheet_with_text(value: &str) -> (WorksheetXml, SharedStringTable) {
    let mut sst = SharedStringTable::new();
    sst.add(value);
    let worksheet = WorksheetXml {
        sheet_data: SheetData {
            rows: vec![Row {
                r: 1,
                spans: None,
                s: None,
                custom_format: None,
                ht: None,
                hidden: None,
                custom_height: None,
                outline_level: None,
                cells: vec![Cell {
                    r: CompactCellRef::new("A1").unwrap(),
                    col: 1,
                    s: None,
                    t: CellTypeTag::SharedString,
                    v: Some("0".to_string()),
                    f: None,
                    is: None,
                }],
            }],
        },
        ..WorksheetXml::default()
    };
    (worksheet, sst)
}

fn assert_well_formed_svg(svg: &str) {
    let mut reader = Reader::from_str(svg);
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .expect("SVG must be well-formed XML")
        {
            Event::Eof => break,
            _ => buffer.clear(),
        }
    }
}

#[test]
fn default_font_family_cannot_escape_its_css_string_or_xml_attribute() {
    let (worksheet, sst) = worksheet_with_text("ハングル 한글");
    let stylesheet = StyleSheet::default();
    let mut render_options = options();
    render_options.default_font_family =
        r#"Noto \";}</style><script href="https://attacker">\url(https://attacker)"#.to_string();

    let svg = render_to_svg(&worksheet, &sst, &stylesheet, &render_options).unwrap();

    assert_well_formed_svg(&svg);
    assert!(!svg.contains("<style>"));
    assert!(!svg.contains("</style>"));
    assert!(!svg.contains("<script"));
    assert!(!svg.contains(" href=\"https://attacker\""));
    assert!(svg.contains("ハングル 한글"));
}

#[test]
fn workbook_font_name_cannot_create_an_svg_attribute() {
    let (mut worksheet, sst) = worksheet_with_text("safe");
    let mut stylesheet = StyleSheet::default();
    let style = StyleBuilder::new()
        .font_name("Noto Sans\" onload=\"alert(1)")
        .build();
    let style_id = add_style(&mut stylesheet, &style).unwrap();
    worksheet.sheet_data.rows[0].cells[0].s = Some(style_id);

    let svg = render_to_svg(&worksheet, &sst, &stylesheet, &options()).unwrap();

    assert_well_formed_svg(&svg);
    assert!(!svg.contains(" onload=\"alert(1)\""));
}

#[test]
fn workbook_font_name_preserves_unicode() {
    let (mut worksheet, sst) = worksheet_with_text("safe");
    let mut stylesheet = StyleSheet::default();
    let style = StyleBuilder::new().font_name("Noto Sans 한글").build();
    let style_id = add_style(&mut stylesheet, &style).unwrap();
    worksheet.sheet_data.rows[0].cells[0].s = Some(style_id);

    let svg = render_to_svg(&worksheet, &sst, &stylesheet, &options()).unwrap();

    assert_well_formed_svg(&svg);
    assert!(svg.contains("Noto Sans 한글"));
}

#[test]
fn xml_forbidden_controls_are_rejected_without_echoing_input() {
    let (worksheet, sst) = worksheet_with_text("secret\0value");
    let stylesheet = StyleSheet::default();

    let error = render_to_svg(&worksheet, &sst, &stylesheet, &options()).unwrap_err();
    assert!(!error.to_string().contains("secret"));

    let (worksheet, sst) = worksheet_with_text("safe");
    let mut render_options = options();
    render_options.default_font_family = "private\u{1f}font".to_string();
    let error = render_to_svg(&worksheet, &sst, &stylesheet, &render_options).unwrap_err();
    assert!(!error.to_string().contains("private"));

    let (mut worksheet, sst) = worksheet_with_text("safe");
    let mut stylesheet = StyleSheet::default();
    let style_id = add_style(
        &mut stylesheet,
        &StyleBuilder::new().font_name("private\0font").build(),
    )
    .unwrap();
    worksheet.sheet_data.rows[0].cells[0].s = Some(style_id);
    let error = render_to_svg(&worksheet, &sst, &stylesheet, &options()).unwrap_err();
    assert!(!error.to_string().contains("private"));
}

#[test]
fn scale_and_font_sizes_must_be_positive_and_finite() {
    let (mut worksheet, sst) = worksheet_with_text("safe");
    let stylesheet = StyleSheet::default();

    for scale in [0.0, -1.0, f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let mut render_options = options();
        render_options.scale = scale;
        assert!(render_to_svg(&worksheet, &sst, &stylesheet, &render_options).is_err());
    }
    for size in [0.0, -1.0, f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let mut render_options = options();
        render_options.default_font_size = size;
        assert!(render_to_svg(&worksheet, &sst, &stylesheet, &render_options).is_err());
    }

    let mut stylesheet = StyleSheet::default();
    let style_id = add_style(
        &mut stylesheet,
        &StyleBuilder::new().font_size(f64::NAN).build(),
    )
    .unwrap();
    worksheet.sheet_data.rows[0].cells[0].s = Some(style_id);
    assert!(render_to_svg(&worksheet, &sst, &stylesheet, &options()).is_err());
}

#[test]
fn rgb_colors_accept_only_six_or_eight_ascii_hex_digits() {
    let invalid = [
        "12345",
        "1234567",
        "GGGGGG",
        "AéééZ",
        "\" onload=\"alert(1)",
        "url(https://attacker)",
    ];

    for color in invalid {
        let (mut worksheet, sst) = worksheet_with_text("safe");
        let mut stylesheet = StyleSheet::default();
        let style_id = add_style(
            &mut stylesheet,
            &StyleBuilder::new()
                .font_color(StyleColor::Rgb(color.to_string()))
                .build(),
        )
        .unwrap();
        worksheet.sheet_data.rows[0].cells[0].s = Some(style_id);

        let outcome = catch_unwind(AssertUnwindSafe(|| {
            render_to_svg(&worksheet, &sst, &stylesheet, &options())
        }));
        let result = outcome.expect("malformed colors must not panic");
        let error = result.expect_err("malformed colors must be rejected");
        assert!(!error.to_string().contains(color));
    }
}

#[test]
fn rgb_and_argb_colors_render_as_rgb_attributes() {
    for (input, expected) in [
        ("a1B2c3", "#a1B2c3"),
        ("80A1B2C3", "#A1B2C3"),
        ("#A1B2C3", "#A1B2C3"),
        ("#80A1B2C3", "#A1B2C3"),
    ] {
        let (mut worksheet, sst) = worksheet_with_text("safe");
        let mut stylesheet = StyleSheet::default();
        let style_id = add_style(
            &mut stylesheet,
            &StyleBuilder::new()
                .font_color(StyleColor::Rgb(input.to_string()))
                .build(),
        )
        .unwrap();
        worksheet.sheet_data.rows[0].cells[0].s = Some(style_id);

        let svg = render_to_svg(&worksheet, &sst, &stylesheet, &options()).unwrap();
        assert!(svg.contains(&format!("fill=\"{expected}\"")));
    }
}

#[test]
fn malformed_fill_and_border_colors_are_rejected() {
    let styles = [
        StyleBuilder::new().solid_fill("not-hex").build(),
        StyleBuilder::new()
            .border_all(BorderLineStyle::Thin, StyleColor::Rgb("AéééZ".to_string()))
            .build(),
    ];

    for style in styles {
        let (mut worksheet, sst) = worksheet_with_text("safe");
        let mut stylesheet = StyleSheet::default();
        let style_id = add_style(&mut stylesheet, &style).unwrap();
        worksheet.sheet_data.rows[0].cells[0].s = Some(style_id);

        let outcome = catch_unwind(AssertUnwindSafe(|| {
            render_to_svg(&worksheet, &sst, &stylesheet, &options())
        }));
        let result = outcome.expect("malformed colors must not panic");
        result.expect_err("malformed colors must be rejected");
    }
}

#[test]
fn caller_output_limit_rejects_escape_heavy_sst_text() {
    let (baseline_worksheet, baseline_sst) = worksheet_with_text("safe");
    let stylesheet = StyleSheet::default();
    let baseline =
        render_to_svg(&baseline_worksheet, &baseline_sst, &stylesheet, &options()).unwrap();

    let (worksheet, sst) = worksheet_with_text(&"<&>".repeat(1_000_000));
    let limit = baseline.len() as u64 + 32;
    let limits = TraversalLimits::new(1, 1, 1, limit).unwrap();
    let error = render_to_svg_with_limits(&worksheet, &sst, &stylesheet, &options(), limits)
        .expect_err("escaped cell text must be charged to the caller's limit");

    let Error::ResourceLimitExceeded(metadata) = error else {
        panic!("expected resource limit error");
    };
    assert_eq!(metadata.operation(), ResourceOperation::Render);
    assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
    assert_eq!(metadata.limit(), limit);
    assert!(!metadata.to_string().contains("<&>"));
}

#[test]
fn formula_sst_and_inline_rich_text_are_escaped_without_flattening_clones() {
    let escape_heavy = "<&>".repeat(4096);
    let mut sst = SharedStringTable::new();
    sst.add(&escape_heavy);
    let worksheet = WorksheetXml {
        sheet_data: SheetData {
            rows: vec![Row {
                r: 1,
                spans: None,
                s: None,
                custom_format: None,
                ht: None,
                hidden: None,
                custom_height: None,
                outline_level: None,
                cells: vec![
                    Cell {
                        r: CompactCellRef::new("A1").unwrap(),
                        col: 1,
                        s: None,
                        t: CellTypeTag::SharedString,
                        v: Some("0".to_string()),
                        f: None,
                        is: None,
                    },
                    Cell {
                        r: CompactCellRef::new("B1").unwrap(),
                        col: 2,
                        s: None,
                        t: CellTypeTag::None,
                        v: None,
                        f: Some(Box::new(CellFormula {
                            t: None,
                            reference: None,
                            si: None,
                            value: Some("A1<&>".to_string()),
                        })),
                        is: None,
                    },
                    Cell {
                        r: CompactCellRef::new("C1").unwrap(),
                        col: 3,
                        s: None,
                        t: CellTypeTag::InlineString,
                        v: None,
                        f: None,
                        is: Some(Box::new(InlineString {
                            t: None,
                            r: vec![
                                R {
                                    r_pr: None,
                                    t: T {
                                        xml_space: None,
                                        value: "rich<&".to_string(),
                                    },
                                },
                                R {
                                    r_pr: None,
                                    t: T {
                                        xml_space: None,
                                        value: ">text".to_string(),
                                    },
                                },
                            ],
                            r_ph: Vec::new(),
                        })),
                    },
                ],
            }],
        },
        ..WorksheetXml::default()
    };

    let svg = render_to_svg(&worksheet, &sst, &StyleSheet::default(), &options()).unwrap();

    assert_well_formed_svg(&svg);
    assert!(svg.contains("&lt;&amp;&gt;&lt;&amp;&gt;"));
    assert!(svg.contains("=A1&lt;&amp;&gt;"));
    assert!(svg.contains("rich&lt;&amp;&gt;text"));
}

#[test]
fn late_font_overflow_reports_root_limit_and_cumulative_observed_bytes() {
    let (mut baseline_worksheet, baseline_sst) = worksheet_with_text("safe");
    let mut baseline_stylesheet = StyleSheet::default();
    let baseline_style = add_style(
        &mut baseline_stylesheet,
        &StyleBuilder::new().font_name("Arial").build(),
    )
    .unwrap();
    baseline_worksheet.sheet_data.rows[0].cells[0].s = Some(baseline_style);
    let baseline = render_to_svg(
        &baseline_worksheet,
        &baseline_sst,
        &baseline_stylesheet,
        &options(),
    )
    .unwrap();

    let configured_limit = baseline.len() as u64 + 32;
    let (mut worksheet, sst) = worksheet_with_text("safe");
    let mut stylesheet = StyleSheet::default();
    let long_font_name = "A".repeat(configured_limit as usize + 32);
    let style_id = add_style(
        &mut stylesheet,
        &StyleBuilder::new().font_name(&long_font_name).build(),
    )
    .unwrap();
    worksheet.sheet_data.rows[0].cells[0].s = Some(style_id);
    let limits = TraversalLimits::new(1, 1, 1, configured_limit).unwrap();

    let error = render_to_svg_with_limits(&worksheet, &sst, &stylesheet, &options(), limits)
        .expect_err("late font staging must retain the root output budget metadata");

    let Error::ResourceLimitExceeded(metadata) = error else {
        panic!("expected resource limit error");
    };
    assert_eq!(metadata.operation(), ResourceOperation::Render);
    assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
    assert_eq!(metadata.limit(), configured_limit);
    assert_eq!(metadata.observed(), configured_limit + 1);
}
