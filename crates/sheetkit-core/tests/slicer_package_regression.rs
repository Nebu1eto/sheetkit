use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::Reader;
use sheetkit_core::slicer::SlicerConfig;
use sheetkit_core::table::{TableColumn, TableConfig};
use sheetkit_core::workbook::Workbook;
use zip::ZipArchive;

const DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const SLICER_REL: &str = "http://schemas.microsoft.com/office/2007/relationships/slicer";
const SLICER_CACHE_REL: &str = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
const SLICER_CONTENT_TYPE: &str = "application/vnd.ms-excel.slicer+xml";
const SLICER_CACHE_CONTENT_TYPE: &str = "application/vnd.ms-excel.slicerCache+xml";

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

fn has_element(xml: &str, name: &str, expected: &[(&str, &str)]) -> bool {
    elements(xml, name).iter().any(|attributes| {
        expected
            .iter()
            .all(|(key, value)| attributes.get(*key).is_some_and(|actual| actual == value))
    })
}

fn defined_name_values(xml: &str) -> BTreeMap<String, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut values = BTreeMap::new();
    let mut current_name = None;

    loop {
        match reader.read_event_into(&mut buffer).expect("valid XML") {
            Event::Start(element) if local_name(element.name().as_ref()) == "definedName" => {
                current_name = element.attributes().find_map(|attribute| {
                    let attribute = attribute.expect("valid XML attribute");
                    (local_name(attribute.key.as_ref()) == "name").then(|| {
                        attribute
                            .decode_and_unescape_value(reader.decoder())
                            .expect("valid defined name")
                            .into_owned()
                    })
                });
            }
            Event::Text(text) if current_name.is_some() => {
                values.insert(
                    current_name.take().expect("defined name"),
                    String::from_utf8_lossy(text.as_ref()).into_owned(),
                );
            }
            Event::End(element) if local_name(element.name().as_ref()) == "definedName" => {
                current_name = None;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    values
}

fn root_child_names(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut names = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer).expect("valid XML") {
            Event::Start(element) => {
                if depth == 1 {
                    names.push(local_name(element.name().as_ref()));
                }
                depth += 1;
            }
            Event::Empty(element) if depth == 1 => names.push(local_name(element.name().as_ref())),
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    names
}

fn one_cell_anchors(xml: &str) -> Vec<String> {
    const START: &str = "<xdr:oneCellAnchor";
    const END: &str = "</xdr:oneCellAnchor>";

    let mut anchors = Vec::new();
    let mut remaining = xml;
    while let Some(start) = remaining.find(START) {
        let anchor = &remaining[start..];
        let end = anchor.find(END).expect("complete oneCellAnchor") + END.len();
        anchors.push(anchor[..end].to_string());
        remaining = &anchor[end..];
    }
    anchors
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

fn relationship_targets(xml: &str, relationship_type: &str) -> BTreeMap<String, String> {
    elements(xml, "Relationship")
        .into_iter()
        .filter(|relationship| {
            relationship.get("Type").map(String::as_str) == Some(relationship_type)
        })
        .map(|relationship| {
            (
                relationship.get("Id").expect("relationship ID").clone(),
                relationship
                    .get("Target")
                    .expect("relationship target")
                    .clone(),
            )
        })
        .collect()
}

fn table(name: &str, range: &str, columns: &[&str]) -> TableConfig {
    TableConfig {
        name: name.to_string(),
        display_name: name.to_string(),
        range: range.to_string(),
        columns: columns
            .iter()
            .map(|name| TableColumn {
                name: (*name).to_string(),
                totals_row_function: None,
                totals_row_label: None,
            })
            .collect(),
        ..TableConfig::default()
    }
}

fn slicer(name: &str, cell: &str, table_name: &str, column_name: &str) -> SlicerConfig {
    SlicerConfig {
        name: name.to_string(),
        cell: cell.to_string(),
        table_name: table_name.to_string(),
        column_name: column_name.to_string(),
        caption: None,
        style: None,
        width: None,
        height: None,
        show_caption: None,
        column_count: None,
    }
}

fn workbook_with_two_tables() -> Workbook {
    let mut workbook = Workbook::new();
    workbook.new_sheet("DataTwo").expect("create second sheet");

    for (sheet, rows) in [
        (
            "Sheet1",
            [["Status", "Region"], ["Open", "North"], ["Closed", "South"]],
        ),
        (
            "DataTwo",
            [["Category", "Owner"], ["Office", "Alice"], ["Home", "Bob"]],
        ),
    ] {
        for (row, values) in rows.iter().enumerate() {
            for (column, value) in values.iter().enumerate() {
                let cell = format!("{}{}", (b'A' + column as u8) as char, row + 1);
                workbook
                    .set_cell_value(sheet, &cell, *value)
                    .expect("write table cell");
            }
        }
    }

    workbook
        .add_table(
            "Sheet1",
            &table("StatusTable", "A1:B3", &["Status", "Region"]),
        )
        .expect("add first table");
    workbook
        .add_table(
            "DataTwo",
            &table("CategoryTable", "A1:B3", &["Category", "Owner"]),
        )
        .expect("add second table");
    workbook
}

#[test]
fn saved_slicers_have_complete_package_graph_and_deterministic_placement() {
    let mut workbook = workbook_with_two_tables();
    let default_slicer = slicer("StatusFilter", "F1", "StatusTable", "Status");
    let mut explicit_slicer = slicer("OwnerFilter", "H3", "CategoryTable", "Owner");
    explicit_slicer.width = Some(300);
    explicit_slicer.height = Some(250);
    explicit_slicer.caption = Some("Filter owner".to_string());

    workbook
        .add_slicer("Sheet1", &default_slicer)
        .expect("add default slicer");
    workbook
        .add_slicer("DataTwo", &explicit_slicer)
        .expect("add explicit slicer");

    let first_save = workbook.save_to_buffer().expect("save workbook");
    let reopened = Workbook::open_from_buffer(&first_save).expect("reopen workbook");
    assert_eq!(
        reopened
            .get_slicers("Sheet1")
            .expect("get first slicers")
            .len(),
        1
    );
    assert_eq!(
        reopened
            .get_slicers("DataTwo")
            .expect("get second slicers")
            .len(),
        1
    );
    let second_save = reopened.save_to_buffer().expect("save reopened workbook");

    for bytes in [first_save, second_save] {
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open XLSX package");
        let names = part_names(&mut archive);
        let slicer_parts: Vec<_> = names
            .iter()
            .filter(|name| name.starts_with("xl/slicers/") && name.ends_with(".xml"))
            .collect();
        let cache_parts: Vec<_> = names
            .iter()
            .filter(|name| name.starts_with("xl/slicerCaches/") && name.ends_with(".xml"))
            .collect();
        let drawing_parts: Vec<_> = names
            .iter()
            .filter(|name| name.starts_with("xl/drawings/drawing") && name.ends_with(".xml"))
            .collect();
        assert_eq!(slicer_parts.len(), 2, "one definition part per slicer");
        assert_eq!(cache_parts.len(), 2, "one cache part per slicer");
        assert_eq!(drawing_parts.len(), 2, "one drawing placement per slicer");
        assert_eq!(
            names.len(),
            archive.len(),
            "ZIP must not contain duplicate paths"
        );

        let workbook_xml = xml_part(&mut archive, "xl/workbook.xml");
        let workbook_children = root_child_names(&workbook_xml);
        assert_eq!(
            workbook_children
                .iter()
                .filter(|name| name.as_str() == "extLst")
                .count(),
            1
        );
        assert_eq!(workbook_children.last().map(String::as_str), Some("extLst"));
        assert!(
            workbook_xml.contains("slicer"),
            "workbook extLst links slicer caches"
        );
        assert!(has_element(
            &workbook_xml,
            "ext",
            &[("uri", "{BBE1A952-AA13-448E-AADC-164F8A28A991}")],
        ));
        let cache_defined_names = elements(&workbook_xml, "definedName");
        let defined_name_values = defined_name_values(&workbook_xml);
        assert!(
            cache_defined_names
                .iter()
                .filter(|defined_name| !defined_name.contains_key("localSheetId"))
                .count()
                >= 2
        );
        let workbook_rels = xml_part(&mut archive, "xl/_rels/workbook.xml.rels");
        let workbook_cache_rels = relationship_targets(&workbook_rels, SLICER_CACHE_REL);
        assert_eq!(
            workbook_cache_rels.len(),
            2,
            "every cache has one workbook relationship"
        );
        let workbook_cache_ids: BTreeSet<_> = elements(&workbook_xml, "slicerCache")
            .iter()
            .filter_map(|cache| cache.get("id").cloned())
            .collect();
        assert_eq!(
            workbook_cache_ids.len(),
            2,
            "workbook extLst keeps cache IDs unique"
        );
        for relationship_id in workbook_cache_ids {
            assert!(
                workbook_cache_rels
                    .get(&relationship_id)
                    .is_some_and(|target| target.starts_with("slicerCaches/")),
                "workbook ext slicer cache r:id resolves to its cache part"
            );
        }

        let content_types = xml_part(&mut archive, "[Content_Types].xml");
        for part in &slicer_parts {
            assert!(has_element(
                &content_types,
                "Override",
                &[
                    ("PartName", &format!("/{part}")),
                    ("ContentType", SLICER_CONTENT_TYPE)
                ],
            ));
        }
        for part in &cache_parts {
            assert!(has_element(
                &content_types,
                "Override",
                &[
                    ("PartName", &format!("/{part}")),
                    ("ContentType", SLICER_CACHE_CONTENT_TYPE)
                ],
            ));
        }

        let first_sheet = xml_part(&mut archive, "xl/worksheets/sheet1.xml");
        let second_sheet = xml_part(&mut archive, "xl/worksheets/sheet2.xml");
        for (sheet_number, worksheet) in [(1, &first_sheet), (2, &second_sheet)] {
            let worksheet_children = root_child_names(worksheet);
            assert_eq!(
                worksheet_children
                    .iter()
                    .filter(|name| name.as_str() == "extLst")
                    .count(),
                1
            );
            assert_eq!(
                worksheet_children.last().map(String::as_str),
                Some("extLst")
            );
            assert!(
                worksheet.contains("slicer"),
                "worksheet extLst links slicers"
            );
            assert!(has_element(
                worksheet,
                "ext",
                &[("uri", "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}")],
            ));
            let rels = xml_part(
                &mut archive,
                &format!("xl/worksheets/_rels/sheet{sheet_number}.xml.rels"),
            );
            let slicer_rels = relationship_targets(&rels, SLICER_REL);
            assert_eq!(slicer_rels.len(), 1);
            for relationship_id in elements(worksheet, "slicer")
                .iter()
                .filter_map(|slicer| slicer.get("id"))
            {
                assert!(
                    slicer_rels
                        .get(relationship_id)
                        .is_some_and(|target| target.starts_with("../slicers/")),
                    "worksheet x14:slicer r:id resolves to its slicer definition"
                );
            }
            assert!(has_element(&rels, "Relationship", &[("Type", DRAWING_REL)]));
        }

        let all_drawing_xml = drawing_parts
            .iter()
            .map(|part| xml_part(&mut archive, part))
            .collect::<Vec<_>>();
        let drawing_anchors = all_drawing_xml
            .iter()
            .flat_map(|xml| one_cell_anchors(xml))
            .collect::<Vec<_>>();
        assert_eq!(
            drawing_anchors.len(),
            2,
            "each slicer has one drawing anchor"
        );
        for anchor in &drawing_anchors {
            assert!(anchor.contains("<xdr:graphicFrame"));
            assert!(anchor.contains("<a:graphic"));
            assert!(anchor.contains(
                "<a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\""
            ));
            assert!(anchor.contains("sle:slicer"));
            assert!(anchor.contains("<mc:Choice Requires=\"sle\""));
            assert!(
                anchor.contains("<xdr:sp"),
                "AlternateContent includes xdr:sp fallback"
            );
            let from = anchor.find("<xdr:from").expect("anchor from marker");
            let ext = anchor.find("<xdr:ext").expect("anchor extent");
            let alternate = anchor
                .find("<mc:AlternateContent")
                .expect("anchor AlternateContent");
            let client_data = anchor.find("<xdr:clientData").expect("anchor clientData");
            assert!(from < ext && ext < alternate && alternate < client_data);
        }
        assert!(drawing_anchors.iter().any(|anchor| {
            anchor.contains("name=\"StatusFilter\"")
                && anchor.contains(">5<")
                && anchor.contains(">0<")
                && anchor.contains("cx=\"1905000\"")
                && anchor.contains("cy=\"1905000\"")
        }));

        for drawing in &all_drawing_xml {
            let c_nv_pr_ids = elements(drawing, "cNvPr")
                .iter()
                .filter_map(|properties| properties.get("id").cloned())
                .collect::<Vec<_>>();
            assert_eq!(
                c_nv_pr_ids.iter().collect::<BTreeSet<_>>().len(),
                c_nv_pr_ids.len(),
                "drawing cNvPr IDs are unique within the drawing part"
            );
        }

        assert!(names.iter().all(|path| {
            !(path.starts_with("xl/slicers/_rels/") || path.starts_with("xl/slicerCaches/_rels/"))
        }));
        assert!(drawing_anchors.iter().any(|anchor| {
            anchor.contains("name=\"OwnerFilter\"")
                && anchor.contains(">7<")
                && anchor.contains(">2<")
                && anchor.contains("cx=\"2857500\"")
                && anchor.contains("cy=\"2381250\"")
        }));

        for cache_part in &cache_parts {
            let cache = xml_part(&mut archive, cache_part);
            let table_cache = elements(&cache, "tableSlicerCache");
            assert_eq!(table_cache.len(), 1, "cache links exactly one source table");
            let cache_definition = &elements(&cache, "slicerCacheDefinition")[0];
            let cache_name = cache_definition.get("name").expect("slicer cache name");
            assert_eq!(
                defined_name_values.get(cache_name).map(String::as_str),
                Some("#N/A"),
                "each cache has its matching workbook-scoped defined name"
            );
            let expected_column_id = match cache_definition.get("sourceName").map(String::as_str) {
                Some("Status") => "1",
                Some("Owner") => "2",
                other => panic!("unexpected slicer source column: {other:?}"),
            };
            assert_eq!(
                table_cache[0].get("column").map(String::as_str),
                Some(expected_column_id),
                "tableSlicerCache column is the source tableColumn@id"
            );
            assert!(table_cache[0].contains_key("tableId"));
        }
    }
}

#[test]
fn deleted_slicer_leaves_no_stale_package_graph_after_two_saves() {
    let mut workbook = workbook_with_two_tables();
    workbook
        .add_slicer(
            "Sheet1",
            &slicer("StatusFilter", "F1", "StatusTable", "Status"),
        )
        .expect("add slicer");
    let with_slicer = workbook.save_to_buffer().expect("save slicer");
    let mut workbook = Workbook::open_from_buffer(&with_slicer).expect("reopen slicer");
    workbook
        .delete_slicer("Sheet1", "StatusFilter")
        .expect("delete slicer");

    let first_save = workbook.save_to_buffer().expect("save deletion");
    let reopened = Workbook::open_from_buffer(&first_save).expect("reopen deletion");
    assert!(reopened
        .get_slicers("Sheet1")
        .expect("get deleted slicers")
        .is_empty());
    let second_save = reopened.save_to_buffer().expect("save reopened deletion");

    for bytes in [first_save, second_save] {
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open XLSX package");
        let names = part_names(&mut archive);
        assert!(names
            .iter()
            .all(|name| !name.starts_with("xl/slicers/") && !name.starts_with("xl/slicerCaches/")));

        let workbook_xml = xml_part(&mut archive, "xl/workbook.xml");
        assert!(!workbook_xml.contains("slicerCache"));
        assert!(!defined_name_values(&workbook_xml).contains_key("StatusFilter_Cache"));
        let workbook_rels = xml_part(&mut archive, "xl/_rels/workbook.xml.rels");
        assert!(relationship_targets(&workbook_rels, SLICER_CACHE_REL).is_empty());

        let worksheet = xml_part(&mut archive, "xl/worksheets/sheet1.xml");
        assert!(!worksheet.contains("slicerList"));
        let worksheet_rels = xml_part(&mut archive, "xl/worksheets/_rels/sheet1.xml.rels");
        assert!(relationship_targets(&worksheet_rels, SLICER_REL).is_empty());
        for drawing in names
            .iter()
            .filter(|name| name.starts_with("xl/drawings/drawing") && name.ends_with(".xml"))
        {
            assert!(!xml_part(&mut archive, drawing).contains("sle:slicer"));
        }
    }
}
