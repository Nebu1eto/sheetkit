use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::Reader;
use sheetkit_core::chart::{ChartConfig, ChartSeries, ChartType};
use sheetkit_core::comment::CommentConfig;
use sheetkit_core::error::Error;
use sheetkit_core::image::{ImageConfig, ImageFormat};
use sheetkit_core::shape::{ShapeConfig, ShapeType};
use sheetkit_core::sparkline::SparklineConfig;
use sheetkit_core::table::{TableColumn, TableConfig};
use sheetkit_core::workbook::Workbook;
use zip::ZipArchive;

fn local_name(name: &[u8]) -> String {
    String::from_utf8_lossy(name)
        .rsplit(':')
        .next()
        .expect("XML names always have a final component")
        .to_string()
}

fn elements(xml: &str, wanted_name: &str) -> Vec<BTreeMap<String, String>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut matches = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer).expect("valid XML") {
            Event::Start(element) | Event::Empty(element)
                if local_name(element.name().as_ref()) == wanted_name =>
            {
                matches.push(
                    element
                        .attributes()
                        .map(|attribute| {
                            let attribute = attribute.expect("valid XML attribute");
                            let value = attribute
                                .decode_and_unescape_value(reader.decoder())
                                .expect("valid XML attribute value")
                                .into_owned();
                            (local_name(attribute.key.as_ref()), value)
                        })
                        .collect(),
                );
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    matches
}

fn xml_part(archive: &mut ZipArchive<Cursor<Vec<u8>>>, name: &str) -> String {
    let mut entry = archive.by_name(name).expect("required ZIP entry");
    let mut xml = String::new();
    entry.read_to_string(&mut xml).expect("UTF-8 XML part");
    xml
}

fn part_names(archive: &mut ZipArchive<Cursor<Vec<u8>>>) -> BTreeSet<String> {
    (0..archive.len())
        .map(|index| {
            archive
                .by_index(index)
                .expect("valid ZIP entry")
                .name()
                .to_string()
        })
        .collect()
}

fn chart_config(chart_type: ChartType, series_count: usize) -> ChartConfig {
    ChartConfig {
        chart_type,
        title: None,
        series: (0..series_count)
            .map(|index| ChartSeries {
                name: format!("Series {index}"),
                categories: "Sheet1!$A$2:$A$4".to_string(),
                values: format!(
                    "Sheet1!${}$2:${}$4",
                    (b'B' + index as u8) as char,
                    (b'B' + index as u8) as char
                ),
                x_values: None,
                bubble_sizes: None,
            })
            .collect(),
        show_legend: false,
        view_3d: None,
    }
}

fn table() -> TableConfig {
    TableConfig {
        name: "Metrics".to_string(),
        display_name: "Metrics".to_string(),
        range: "A1:B4".to_string(),
        columns: ["Name", "Value"]
            .into_iter()
            .map(|name| TableColumn {
                name: name.to_string(),
                totals_row_function: None,
                totals_row_label: None,
            })
            .collect(),
        ..TableConfig::default()
    }
}

fn assert_package(bytes: Vec<u8>) {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open XLSX package");
    let names = part_names(&mut archive);
    assert_eq!(names.len(), archive.len(), "ZIP paths are unique");

    let chart_parts = names
        .iter()
        .filter(|name| name.starts_with("xl/charts/chart") && name.ends_with(".xml"))
        .cloned()
        .collect::<Vec<_>>();
    assert_eq!(chart_parts.len(), 6, "two standard plus four stock charts");
    assert_eq!(
        chart_parts.iter().collect::<BTreeSet<_>>().len(),
        chart_parts.len()
    );

    let drawing = xml_part(&mut archive, "xl/drawings/drawing1.xml");
    let drawing_ids = elements(&drawing, "cNvPr")
        .into_iter()
        .filter_map(|properties| properties.get("id").cloned())
        .collect::<Vec<_>>();
    assert_eq!(drawing_ids.len(), 6, "one drawing object per chart");
    assert_eq!(
        drawing_ids.iter().collect::<BTreeSet<_>>().len(),
        drawing_ids.len()
    );

    let stock_parts = chart_parts
        .iter()
        .filter(|part| xml_part(&mut archive, part).contains("stockChart"))
        .cloned()
        .collect::<Vec<_>>();
    assert_eq!(stock_parts.len(), 4);
    let stock_signatures = stock_parts
        .iter()
        .map(|part| {
            let xml = xml_part(&mut archive, part);
            let series = elements(&xml, "ser").len();
            let has_high_low_lines = xml.contains("hiLowLines");
            let has_up_down_bars = xml.contains("upDownBars");
            let has_volume_chart = xml.contains("<c:barChart>");
            let value_axes = xml.matches("<c:valAx>").count();
            (
                series,
                has_high_low_lines,
                has_up_down_bars,
                has_volume_chart,
                value_axes,
            )
        })
        .collect::<BTreeSet<_>>();
    assert_eq!(
        stock_signatures.len(),
        4,
        "stock variants serialize distinct required structures"
    );
    assert!(
        stock_signatures.contains(&(3, true, false, false, 1)),
        "HLC structure"
    );
    assert!(
        stock_signatures.contains(&(4, true, true, false, 1)),
        "OHLC structure"
    );
    assert!(
        stock_signatures.contains(&(4, true, false, true, 2)),
        "VHLC structure"
    );
    assert!(
        stock_signatures.contains(&(5, true, true, true, 2)),
        "VOHLC structure"
    );

    let worksheet = xml_part(&mut archive, "xl/worksheets/sheet1.xml");
    let mut reader = Reader::from_str(&worksheet);
    let mut buffer = Vec::new();
    loop {
        if reader
            .read_event_into(&mut buffer)
            .expect("well-formed worksheet XML")
            == Event::Eof
        {
            break;
        }
        buffer.clear();
    }
    let legacy = worksheet.find("<legacyDrawing ").expect("legacy drawing");
    let tables = worksheet.find("<tableParts ").expect("table parts");
    let extensions = worksheet
        .find("<extLst>")
        .expect("sparkline extension list");
    assert!(legacy < tables && tables < extensions);
}

#[test]
fn charts_sparklines_and_legacy_drawing_keep_distinct_valid_package_structure() {
    let mut workbook = Workbook::new();
    for (cell, value) in [
        ("A1", "Name"),
        ("B1", "Value"),
        ("A2", "One"),
        ("A3", "Two"),
        ("A4", "Three"),
    ] {
        workbook
            .set_cell_value("Sheet1", cell, value)
            .expect("write table value");
    }
    for column in b'B'..=b'F' {
        for row in 2..=4 {
            workbook
                .set_cell_value("Sheet1", &format!("{}{}", column as char, row), row as f64)
                .expect("write chart value");
        }
    }
    workbook.add_table("Sheet1", &table()).expect("add table");
    workbook
        .add_comment(
            "Sheet1",
            &CommentConfig {
                cell: "A1".to_string(),
                author: "Reviewer".to_string(),
                text: "Preserve VML drawing".to_string(),
            },
        )
        .expect("add comment");

    let sparkline_range = "'Data & <North>'!$A$1:$A$3";
    workbook
        .add_sparkline("Sheet1", &SparklineConfig::new(sparkline_range, "G2"))
        .expect("add escaped sparkline range");

    for (index, (chart_type, series_count)) in [
        (ChartType::Line, 1),
        (ChartType::Col, 1),
        (ChartType::StockHLC, 3),
        (ChartType::StockOHLC, 4),
        (ChartType::StockVHLC, 4),
        (ChartType::StockVOHLC, 5),
    ]
    .into_iter()
    .enumerate()
    {
        let row = 1 + index * 12;
        workbook
            .add_chart(
                "Sheet1",
                &format!("I{row}"),
                &format!("P{}", row + 10),
                &chart_config(chart_type, series_count),
            )
            .expect("add chart");
    }

    let first_save = workbook.save_to_buffer().expect("save workbook");
    assert_package(first_save.clone());
    let reopened = Workbook::open_from_buffer(&first_save).expect("reopen workbook");
    assert_eq!(
        reopened.get_sparklines("Sheet1").expect("read sparklines")[0].data_range,
        sparkline_range
    );
    let second_save = reopened.save_to_buffer().expect("save reopened workbook");
    assert_package(second_save);
}

#[test]
fn mixed_drawing_objects_reuse_neither_sparse_nor_deleted_ids() {
    let mut workbook = Workbook::new();
    let chart = chart_config(ChartType::Line, 1);
    workbook
        .add_chart("Sheet1", "A1", "H10", &chart)
        .expect("add first chart");
    workbook
        .add_chart("Sheet1", "A12", "H21", &chart)
        .expect("add second chart");
    workbook
        .delete_chart("Sheet1", "A1")
        .expect("delete first chart");
    workbook
        .add_shape(
            "Sheet1",
            &ShapeConfig {
                shape_type: ShapeType::Rect,
                from_cell: "J1".to_string(),
                to_cell: "L4".to_string(),
                text: Some("Status".to_string()),
                fill_color: None,
                line_color: None,
                line_width: None,
            },
        )
        .expect("add shape");
    workbook
        .add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x89, 0x50, 0x4e, 0x47],
                format: ImageFormat::Png,
                from_cell: "J6".to_string(),
                width_px: 32,
                height_px: 32,
            },
        )
        .expect("add image");

    let saved = workbook.save_to_buffer().expect("save mixed drawing");
    let mut archive = ZipArchive::new(Cursor::new(saved)).expect("open XLSX package");
    let drawing = xml_part(&mut archive, "xl/drawings/drawing1.xml");
    let object_ids = elements(&drawing, "cNvPr")
        .into_iter()
        .filter_map(|properties| properties.get("id").cloned())
        .collect::<Vec<_>>();
    assert_eq!(object_ids.len(), 3);
    assert_eq!(
        object_ids.iter().collect::<BTreeSet<_>>().len(),
        object_ids.len(),
        "all drawing objects must have unique non-visual IDs"
    );
}

#[test]
fn stock_chart_series_validation_is_atomic() {
    let mut workbook = Workbook::new();
    let invalid = chart_config(ChartType::StockHLC, 2);

    let error = workbook
        .add_chart("Sheet1", "A1", "H10", &invalid)
        .expect_err("HLC requires three series");
    assert!(matches!(error, Error::InvalidArgument(_)));

    let saved = workbook
        .save_to_buffer()
        .expect("save after rejected chart");
    let mut archive = ZipArchive::new(Cursor::new(saved)).expect("open XLSX package");
    let names = part_names(&mut archive);
    assert!(names.iter().all(|name| !name.starts_with("xl/charts/")));
    assert!(names.iter().all(|name| !name.starts_with("xl/drawings/")));
}
