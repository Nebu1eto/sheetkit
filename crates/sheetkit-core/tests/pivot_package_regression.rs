use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::Reader;
use sheetkit_core::pivot::{AggregateFunction, PivotDataField, PivotField, PivotTableConfig};
use sheetkit_core::workbook::Workbook;
use zip::ZipArchive;

const PIVOT_CACHE_DEFINITION_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const PIVOT_CACHE_RECORDS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
const PIVOT_TABLE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
const PIVOT_TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
const PIVOT_CACHE_DEFINITION_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
const PIVOT_CACHE_RECORDS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";

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
                let attributes = element
                    .attributes()
                    .map(|attribute| {
                        let attribute = attribute.expect("valid XML attribute");
                        let value = attribute
                            .decode_and_unescape_value(reader.decoder())
                            .expect("valid XML attribute value")
                            .into_owned();
                        (local_name(attribute.key.as_ref()), value)
                    })
                    .collect();
                matches.push(attributes);
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    matches
}

fn has_element(xml: &str, name: &str, expected: &[(&str, &str)]) -> bool {
    elements(xml, name).iter().any(|attributes| {
        expected
            .iter()
            .all(|(key, value)| attributes.get(*key).is_some_and(|actual| actual == value))
    })
}

fn cache_record_field_order(xml: &str) -> Vec<Vec<String>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut records = Vec::new();
    let mut fields = None;

    loop {
        match reader.read_event_into(&mut buffer).expect("valid XML") {
            Event::Start(element) => {
                let name = local_name(element.name().as_ref());
                if name == "r" {
                    fields = Some(Vec::new());
                } else if let Some(fields) = &mut fields {
                    if matches!(name.as_str(), "x" | "n" | "s" | "b") {
                        fields.push(name);
                    }
                }
            }
            Event::Empty(element) => {
                let name = local_name(element.name().as_ref());
                if let Some(fields) = &mut fields {
                    if matches!(name.as_str(), "x" | "n" | "s" | "b") {
                        fields.push(name);
                    }
                }
            }
            Event::End(element) if local_name(element.name().as_ref()) == "r" => {
                records.push(fields.take().expect("cache record fields"));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    records
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

fn xml_part(archive: &mut ZipArchive<Cursor<Vec<u8>>>, name: &str) -> String {
    let mut entry = archive.by_name(name).expect("required ZIP entry");
    let mut xml = String::new();
    entry.read_to_string(&mut xml).expect("UTF-8 XML part");
    xml
}

fn report_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    workbook.new_sheet("DataTwo").expect("create source sheet");
    workbook.new_sheet("Report").expect("create report sheet");

    for (sheet, rows) in [
        (
            "Sheet1",
            vec![
                ["Name", "Region", "Sales", "Quarter"],
                ["Alice", "North", "100", "Q1"],
                ["Bob", "South", "200", "Q2"],
                ["Carol", "North", "150", "Q1"],
            ],
        ),
        (
            "DataTwo",
            vec![
                ["Product", "Channel", "Units", "Year"],
                ["Paper", "Online", "10", "2025"],
                ["Pen", "Retail", "20", "2026"],
            ],
        ),
    ] {
        for (row_index, row) in rows.iter().enumerate() {
            for (column_index, value) in row.iter().enumerate() {
                let cell = format!("{}{}", (b'A' + column_index as u8) as char, row_index + 1);
                workbook
                    .set_cell_value(sheet, &cell, *value)
                    .expect("write pivot source value");
            }
        }
    }

    for (sheet, cell, value) in [
        ("Sheet1", "C2", 100.0),
        ("Sheet1", "C3", 200.0),
        ("Sheet1", "C4", 150.0),
        ("DataTwo", "C2", 10.0),
        ("DataTwo", "C3", 20.0),
    ] {
        workbook
            .set_cell_value(sheet, cell, value)
            .expect("write numeric pivot source value");
    }

    workbook
}

fn regional_sales_pivot() -> PivotTableConfig {
    PivotTableConfig {
        name: "RegionalSales".to_string(),
        source_sheet: "Sheet1".to_string(),
        source_range: "A1:D4".to_string(),
        target_sheet: "Report".to_string(),
        target_cell: "B3".to_string(),
        rows: vec![PivotField {
            name: "Region".to_string(),
        }],
        columns: vec![PivotField {
            name: "Quarter".to_string(),
        }],
        data: vec![PivotDataField {
            name: "Sales".to_string(),
            function: AggregateFunction::Sum,
            display_name: None,
        }],
    }
}

fn product_units_pivot() -> PivotTableConfig {
    PivotTableConfig {
        name: "ProductUnits".to_string(),
        source_sheet: "DataTwo".to_string(),
        source_range: "A1:D3".to_string(),
        target_sheet: "Report".to_string(),
        target_cell: "J5".to_string(),
        rows: vec![PivotField {
            name: "Product".to_string(),
        }],
        columns: vec![],
        data: vec![PivotDataField {
            name: "Units".to_string(),
            function: AggregateFunction::Count,
            display_name: None,
        }],
    }
}

#[test]
fn saved_pivot_cache_is_populated_and_linked_as_an_ooxml_package() {
    let mut workbook = report_workbook();
    workbook
        .add_pivot_table(&regional_sales_pivot())
        .expect("add pivot table");
    let bytes = workbook.save_to_buffer().expect("save workbook");
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open XLSX package");

    let names = part_names(&mut archive);
    assert!(names.contains("xl/pivotTables/pivotTable1.xml"));
    assert!(names.contains("xl/pivotCache/pivotCacheDefinition1.xml"));
    assert!(names.contains("xl/pivotCache/pivotCacheRecords1.xml"));
    assert!(names.contains("xl/pivotTables/_rels/pivotTable1.xml.rels"));
    assert!(names.contains("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"));

    let cache_definition = xml_part(&mut archive, "xl/pivotCache/pivotCacheDefinition1.xml");
    let cache_definition_roots = elements(&cache_definition, "pivotCacheDefinition");
    let cache_definition_root = &cache_definition_roots[0];
    assert!(matches!(
        cache_definition_root
            .get("refreshOnLoad")
            .map(String::as_str),
        Some("1" | "true")
    ));
    assert_eq!(
        cache_definition_root.get("recordCount").map(String::as_str),
        Some("3")
    );
    assert!(has_element(
        &cache_definition,
        "worksheetSource",
        &[("sheet", "Sheet1"), ("ref", "A1:D4")],
    ));
    assert_eq!(elements(&cache_definition, "sharedItems").len(), 4);
    assert!(elements(&cache_definition, "sharedItems")
        .iter()
        .all(|items| items.get("count").is_some_and(|count| count != "0")));

    let cache_records = xml_part(&mut archive, "xl/pivotCache/pivotCacheRecords1.xml");
    assert!(has_element(
        &cache_records,
        "pivotCacheRecords",
        &[("count", "3")]
    ));
    assert_eq!(elements(&cache_records, "r").len(), 3);
    assert_eq!(elements(&cache_records, "n").len(), 3);
    for fields in cache_record_field_order(&cache_records) {
        assert_eq!(fields.len(), 4);
        assert_eq!(fields[2], "n");
    }

    let cache_rels = xml_part(
        &mut archive,
        "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
    );
    assert!(has_element(
        &cache_rels,
        "Relationship",
        &[
            ("Type", PIVOT_CACHE_RECORDS_REL),
            ("Target", "pivotCacheRecords1.xml"),
        ],
    ));

    let pivot_table_rels = xml_part(&mut archive, "xl/pivotTables/_rels/pivotTable1.xml.rels");
    assert!(has_element(
        &pivot_table_rels,
        "Relationship",
        &[
            ("Type", PIVOT_CACHE_DEFINITION_REL),
            ("Target", "../pivotCache/pivotCacheDefinition1.xml"),
        ],
    ));

    let pivot_table = xml_part(&mut archive, "xl/pivotTables/pivotTable1.xml");
    let location = elements(&pivot_table, "location");
    assert_eq!(location.len(), 1);
    assert_ne!(location[0].get("ref").map(String::as_str), Some("B3"));
    assert!(location[0]
        .get("ref")
        .is_some_and(|reference| reference.starts_with("B3:")));
    assert!(elements(&pivot_table, "items").len() >= 2);
}

#[test]
fn multiple_saved_pivots_keep_cache_links_and_metadata_distinct_after_reopen() {
    let first = regional_sales_pivot();
    let second = product_units_pivot();
    let mut workbook = report_workbook();
    workbook.add_pivot_table(&first).expect("add first pivot");
    workbook.add_pivot_table(&second).expect("add second pivot");
    let bytes = workbook.save_to_buffer().expect("save workbook");
    let reopened = Workbook::open_from_buffer(&bytes).expect("reopen workbook");
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open XLSX package");

    let names = part_names(&mut archive);
    for number in 1..=2 {
        assert!(names.contains(&format!("xl/pivotTables/pivotTable{number}.xml")));
        assert!(names.contains(&format!("xl/pivotCache/pivotCacheDefinition{number}.xml")));
        assert!(names.contains(&format!("xl/pivotCache/pivotCacheRecords{number}.xml")));
        assert!(names.contains(&format!("xl/pivotTables/_rels/pivotTable{number}.xml.rels")));
        assert!(names.contains(&format!(
            "xl/pivotCache/_rels/pivotCacheDefinition{number}.xml.rels"
        )));
    }

    let workbook_xml = xml_part(&mut archive, "xl/workbook.xml");
    let caches = elements(&workbook_xml, "pivotCache");
    assert_eq!(caches.len(), 2);
    assert_eq!(
        caches
            .iter()
            .filter_map(|cache| cache.get("cacheId"))
            .collect::<BTreeSet<_>>()
            .len(),
        2
    );
    assert_eq!(
        caches
            .iter()
            .filter_map(|cache| cache.get("id"))
            .collect::<BTreeSet<_>>()
            .len(),
        2
    );

    let workbook_rels = xml_part(&mut archive, "xl/_rels/workbook.xml.rels");
    for number in 1..=2 {
        assert!(has_element(
            &workbook_rels,
            "Relationship",
            &[
                ("Type", PIVOT_CACHE_DEFINITION_REL),
                (
                    "Target",
                    &format!("pivotCache/pivotCacheDefinition{number}.xml"),
                ),
            ],
        ));

        let cache_rels = xml_part(
            &mut archive,
            &format!("xl/pivotCache/_rels/pivotCacheDefinition{number}.xml.rels"),
        );
        assert!(has_element(
            &cache_rels,
            "Relationship",
            &[
                ("Type", PIVOT_CACHE_RECORDS_REL),
                ("Target", &format!("pivotCacheRecords{number}.xml")),
            ],
        ));

        let pivot_table = xml_part(
            &mut archive,
            &format!("xl/pivotTables/pivotTable{number}.xml"),
        );
        let pivot_definition = elements(&pivot_table, "pivotTableDefinition");
        let cache_id = pivot_definition[0]
            .get("cacheId")
            .expect("pivot table cache ID");
        let workbook_cache = caches
            .iter()
            .find(|cache| cache.get("cacheId") == Some(cache_id))
            .expect("workbook cache entry for pivot table");
        let relationship_id = workbook_cache
            .get("id")
            .expect("workbook cache relationship ID");
        assert!(has_element(
            &workbook_rels,
            "Relationship",
            &[
                ("Id", relationship_id),
                (
                    "Target",
                    &format!("pivotCache/pivotCacheDefinition{number}.xml"),
                ),
            ],
        ));

        let pivot_table_rels = xml_part(
            &mut archive,
            &format!("xl/pivotTables/_rels/pivotTable{number}.xml.rels"),
        );
        assert!(has_element(
            &pivot_table_rels,
            "Relationship",
            &[
                ("Type", PIVOT_CACHE_DEFINITION_REL),
                (
                    "Target",
                    &format!("../pivotCache/pivotCacheDefinition{number}.xml"),
                ),
            ],
        ));
    }

    let report_rels = xml_part(&mut archive, "xl/worksheets/_rels/sheet3.xml.rels");
    for number in 1..=2 {
        assert!(has_element(
            &report_rels,
            "Relationship",
            &[
                ("Type", PIVOT_TABLE_REL),
                ("Target", &format!("../pivotTables/pivotTable{number}.xml")),
            ],
        ));
    }

    let content_types = xml_part(&mut archive, "[Content_Types].xml");
    for number in 1..=2 {
        assert!(has_element(
            &content_types,
            "Override",
            &[
                ("ContentType", PIVOT_TABLE_CONTENT_TYPE),
                (
                    "PartName",
                    &format!("/xl/pivotTables/pivotTable{number}.xml")
                ),
            ],
        ));
        assert!(has_element(
            &content_types,
            "Override",
            &[
                ("ContentType", PIVOT_CACHE_DEFINITION_CONTENT_TYPE),
                (
                    "PartName",
                    &format!("/xl/pivotCache/pivotCacheDefinition{number}.xml"),
                ),
            ],
        ));
        assert!(has_element(
            &content_types,
            "Override",
            &[
                ("ContentType", PIVOT_CACHE_RECORDS_CONTENT_TYPE),
                (
                    "PartName",
                    &format!("/xl/pivotCache/pivotCacheRecords{number}.xml"),
                ),
            ],
        ));
    }

    let pivots = reopened.get_pivot_tables();
    assert_eq!(pivots.len(), 2);
    for (config, expected_name) in [(&first, "RegionalSales"), (&second, "ProductUnits")] {
        let pivot = pivots
            .iter()
            .find(|pivot| pivot.name == expected_name)
            .expect("reopened pivot table");
        assert_eq!(pivot.source_sheet, config.source_sheet);
        assert_eq!(pivot.source_range, config.source_range);
        assert_eq!(pivot.target_sheet, config.target_sheet);
        assert!(pivot
            .location
            .starts_with(&format!("{}:", config.target_cell)));

        let table_location = (1..=2)
            .map(|number| {
                let table = xml_part(
                    &mut archive,
                    &format!("xl/pivotTables/pivotTable{number}.xml"),
                );
                let definitions = elements(&table, "pivotTableDefinition");
                let locations = elements(&table, "location");
                let definition = &definitions[0];
                let location = &locations[0];
                (
                    definition.get("name").cloned(),
                    location.get("ref").cloned(),
                )
            })
            .find(|(name, _)| name.as_deref() == Some(expected_name))
            .and_then(|(_, location)| location)
            .expect("serialized pivot location");
        assert_eq!(pivot.location, table_location);
    }
}
