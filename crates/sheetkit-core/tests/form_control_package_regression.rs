use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::Reader;
use sheetkit_core::comment::CommentConfig;
use sheetkit_core::control::{FormControlConfig, FormControlType};
use sheetkit_core::workbook::Workbook;
use zip::ZipArchive;

const CONTROL_PROPERTIES_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
const CONTROL_PROPERTIES_CONTENT_TYPE: &str = "application/vnd.ms-excel.controlproperties+xml";
const VML_DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const COMMENTS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

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

fn resolved_sheet_target(target: &str) -> String {
    let mut components = vec!["xl".to_string(), "worksheets".to_string()];
    for component in target.split('/') {
        match component {
            "" | "." => {}
            ".." => {
                components.pop().expect("relative target stays in package");
            }
            component => components.push(component.to_string()),
        }
    }
    components.join("/")
}

fn assert_complete_control_package(bytes: Vec<u8>, expected_controls: usize) {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open XLSX package");
    let names = part_names(&mut archive);
    assert_eq!(
        names.len(),
        archive.len(),
        "ZIP must not contain duplicate paths"
    );

    let worksheet = xml_part(&mut archive, "xl/worksheets/sheet1.xml");
    let controls = elements(&worksheet, "control");
    assert_eq!(controls.len(), expected_controls);
    assert_eq!(
        controls
            .iter()
            .filter_map(|control| control.get("shapeId"))
            .collect::<BTreeSet<_>>()
            .len(),
        expected_controls,
        "worksheet control shape IDs are unique"
    );

    let rels = xml_part(&mut archive, "xl/worksheets/_rels/sheet1.xml.rels");
    assert_eq!(relationship_targets(&rels, VML_DRAWING_REL).len(), 1);
    assert_eq!(relationship_targets(&rels, COMMENTS_REL).len(), 1);
    let control_properties = relationship_targets(&rels, CONTROL_PROPERTIES_REL);
    assert_eq!(control_properties.len(), expected_controls);
    assert_eq!(
        names
            .iter()
            .filter(|name| name.starts_with("xl/ctrlProps/ctrlProp"))
            .count(),
        expected_controls,
        "every ctrlProp part is owned by one worksheet control"
    );
    let mut object_types = BTreeSet::new();
    for control in &controls {
        let relationship_id = control.get("id").expect("control relationship ID");
        let target = control_properties
            .get(relationship_id)
            .expect("worksheet control r:id resolves to a ctrlProp relationship");
        assert!(
            names.contains(&resolved_sheet_target(target)),
            "ctrlProp relationship target exists"
        );
        let properties = xml_part(&mut archive, &resolved_sheet_target(target));
        let attributes = elements(&properties, "formControlPr")
            .into_iter()
            .next()
            .expect("ctrlProp root");
        object_types.insert(
            attributes
                .get("objectType")
                .expect("control object type")
                .clone(),
        );
        if attributes.get("objectType").map(String::as_str) == Some("CheckBox") {
            assert_eq!(
                attributes.get("checked").map(String::as_str),
                Some("Checked")
            );
            assert_eq!(attributes.get("fmlaLink").map(String::as_str), Some("$F$2"));
        }
    }
    let expected_object_types = match expected_controls {
        2 => BTreeSet::from(["Button".to_string(), "CheckBox".to_string()]),
        1 => BTreeSet::from(["CheckBox".to_string()]),
        _ => BTreeSet::new(),
    };
    assert_eq!(object_types, expected_object_types);

    let content_types = xml_part(&mut archive, "[Content_Types].xml");
    let control_property_overrides = elements(&content_types, "Override")
        .into_iter()
        .filter(|entry| {
            entry.get("ContentType").map(String::as_str) == Some(CONTROL_PROPERTIES_CONTENT_TYPE)
        })
        .collect::<Vec<_>>();
    assert_eq!(control_property_overrides.len(), expected_controls);
    for target in control_properties.values() {
        assert!(control_property_overrides.iter().any(|override_entry| {
            override_entry.get("PartName").is_some_and(|part_name| {
                part_name == &format!("/{}", resolved_sheet_target(target))
            })
        }));
    }

    let vml = xml_part(&mut archive, "xl/drawings/vmlDrawing1.vml");
    assert!(vml.contains("ObjectType=\"Note\""));
    if expected_controls > 0 {
        assert!(vml.contains("ObjectType=\"Checkbox\""));
        assert!(vml.contains("3, 15, 4, 10, 4, 527, 5, 61"));
    } else {
        assert!(!vml.contains("ObjectType=\"Checkbox\""));
    }
    if expected_controls == 2 {
        assert!(vml.contains("ObjectType=\"Button\""));
        assert!(vml.contains("1, 15, 1, 10, 3, 15, 3, 10"));
    } else {
        assert!(!vml.contains("ObjectType=\"Button\""));
        assert!(!vml.contains("Café"));
    }
    let shape_ids = elements(&vml, "shape")
        .into_iter()
        .filter_map(|shape| shape.get("id").cloned())
        .collect::<Vec<_>>();
    assert_eq!(
        shape_ids.iter().collect::<BTreeSet<_>>().len(),
        shape_ids.len()
    );
    let vml_control_shape_ids = shape_ids
        .iter()
        .filter_map(|shape_id| shape_id.strip_prefix("_x0000_s"))
        .collect::<BTreeSet<_>>();
    let worksheet_shape_ids = controls
        .iter()
        .filter_map(|control| control.get("shapeId"))
        .map(String::as_str)
        .collect::<BTreeSet<_>>();
    assert!(worksheet_shape_ids.is_subset(&vml_control_shape_ids));
}

#[test]
fn form_controls_keep_complete_package_graph_through_roundtrip_and_delete() {
    let label = "Save & <close> \"Café\"";
    let mut workbook = Workbook::new();
    workbook
        .add_comment(
            "Sheet1",
            &CommentConfig {
                cell: "A1".to_string(),
                author: "Reviewer".to_string(),
                text: "Comment survives controls".to_string(),
            },
        )
        .expect("add comment");

    let mut button = FormControlConfig::button("B2", label);
    button.width = Some(96.0);
    button.height = Some(30.0);
    workbook
        .add_form_control("Sheet1", button)
        .expect("add escaped-label button");

    let mut checkbox = FormControlConfig::checkbox("D5", "Enabled");
    checkbox.cell_link = Some("$F$2".to_string());
    checkbox.checked = Some(true);
    workbook
        .add_form_control("Sheet1", checkbox)
        .expect("add checkbox");

    let first_save = workbook.save_to_buffer().expect("save workbook");
    let mut reopened = Workbook::open_from_buffer(&first_save).expect("reopen workbook");
    let controls = reopened
        .get_form_controls("Sheet1")
        .expect("read form controls");
    assert_eq!(controls.len(), 2);
    assert_eq!(controls[0].control_type, FormControlType::Button);
    assert_eq!(controls[0].text.as_deref(), Some(label));
    assert_eq!(controls[0].cell, "B2");
    assert_eq!(controls[1].control_type, FormControlType::CheckBox);
    assert_eq!(controls[1].cell, "D5");
    assert_eq!(controls[1].cell_link.as_deref(), Some("$F$2"));
    assert_eq!(controls[1].checked, Some(true));
    assert_eq!(
        reopened.get_comments("Sheet1").expect("read comments")[0].text,
        "Comment survives controls"
    );

    let second_save = reopened.save_to_buffer().expect("save reopened workbook");
    assert_complete_control_package(first_save, 2);
    assert_complete_control_package(second_save.clone(), 2);

    let mut after_second_save =
        Workbook::open_from_buffer(&second_save).expect("reopen stable workbook");
    after_second_save
        .delete_form_control("Sheet1", 0)
        .expect("delete button");
    let deleted_save = after_second_save.save_to_buffer().expect("save deletion");
    let mut deleted_reopen = Workbook::open_from_buffer(&deleted_save).expect("reopen deletion");
    let remaining = deleted_reopen
        .get_form_controls("Sheet1")
        .expect("read remaining control");
    assert_eq!(remaining.len(), 1);
    assert_eq!(remaining[0].control_type, FormControlType::CheckBox);
    assert_eq!(remaining[0].cell, "D5");
    assert_eq!(remaining[0].checked, Some(true));
    assert_eq!(
        deleted_reopen.get_comments("Sheet1").expect("read comment")[0].text,
        "Comment survives controls"
    );
    let deleted_second_save = deleted_reopen
        .save_to_buffer()
        .expect("save reopened deletion");
    assert_complete_control_package(deleted_save, 1);
    assert_complete_control_package(deleted_second_save.clone(), 1);

    let mut delete_last =
        Workbook::open_from_buffer(&deleted_second_save).expect("reopen final control");
    delete_last
        .delete_form_control("Sheet1", 0)
        .expect("delete final control");
    let without_controls = delete_last.save_to_buffer().expect("save without controls");
    assert_complete_control_package(without_controls, 0);
}

#[test]
fn dirty_control_sheet_preserves_another_sheets_control_properties() {
    let mut workbook = Workbook::new();
    workbook.new_sheet("Other").expect("create second sheet");
    workbook
        .add_form_control("Sheet1", FormControlConfig::button("A1", "First"))
        .expect("add first control");
    workbook
        .add_form_control("Other", FormControlConfig::checkbox("B2", "Second"))
        .expect("add second control");
    let initial = workbook.save_to_buffer().expect("save both controls");
    let mut reopened = Workbook::open_from_buffer(&initial).expect("reopen controls");

    reopened
        .delete_form_control("Sheet1", 0)
        .expect("delete first control");
    let saved = reopened.save_to_buffer().expect("save one dirty sheet");
    let mut archive = ZipArchive::new(Cursor::new(saved)).expect("open XLSX package");
    let names = part_names(&mut archive);
    let first_rels = xml_part(&mut archive, "xl/worksheets/_rels/sheet1.xml.rels");
    let second_rels = xml_part(&mut archive, "xl/worksheets/_rels/sheet2.xml.rels");

    assert!(
        relationship_targets(&first_rels, CONTROL_PROPERTIES_REL).is_empty(),
        "unexpected Sheet1 relationships: {first_rels}"
    );
    assert_eq!(
        relationship_targets(&second_rels, CONTROL_PROPERTIES_REL).len(),
        1
    );
    assert_eq!(
        names
            .iter()
            .filter(|name| name.starts_with("xl/ctrlProps/ctrlProp"))
            .count(),
        1
    );
}
