//! Regressions found by opening SheetKit output in Microsoft Excel.

use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::Reader;
use sheetkit_core::comment::CommentConfig;
use sheetkit_core::slicer::SlicerConfig;
use sheetkit_core::table::{TableColumn, TableConfig};
use sheetkit_core::threaded_comment::ThreadedCommentInput;
use sheetkit_core::workbook::Workbook;
use zip::ZipArchive;

const X14_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const X15_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

fn part(buffer: &[u8], name: &str) -> Option<String> {
    let mut archive = ZipArchive::new(Cursor::new(buffer)).expect("valid zip");
    let mut file = archive.by_name(name).ok()?;
    let mut xml = String::new();
    file.read_to_string(&mut xml).expect("utf-8 part");
    Some(xml)
}

fn required_part(buffer: &[u8], name: &str) -> String {
    part(buffer, name).unwrap_or_else(|| panic!("missing part {name}"))
}

fn column(name: &str) -> TableColumn {
    TableColumn {
        name: name.to_string(),
        totals_row_function: None,
        totals_row_label: None,
    }
}

fn slicer_workbook() -> Workbook {
    let mut wb = Workbook::new();
    for (cell, value) in [("A2", "Apple"), ("A3", "Pear"), ("A4", "Plum")] {
        wb.set_cell_value("Sheet1", cell, value).unwrap();
    }
    for (cell, value) in [("B2", 3.0), ("B3", 2.0), ("B4", 5.0)] {
        wb.set_cell_value("Sheet1", cell, value).unwrap();
    }
    wb.add_table(
        "Sheet1",
        &TableConfig {
            name: "Sales".to_string(),
            display_name: "Sales".to_string(),
            range: "A1:B4".to_string(),
            columns: vec![column("Product"), column("Qty")],
            style_name: Some("TableStyleMedium2".to_string()),
            ..TableConfig::default()
        },
    )
    .unwrap();
    wb.add_slicer(
        "Sheet1",
        &SlicerConfig {
            name: "Product Slicer".to_string(),
            cell: "E2".to_string(),
            table_name: "Sales".to_string(),
            column_name: "Product".to_string(),
            caption: None,
            style: None,
            width: None,
            height: None,
            show_caption: None,
            column_count: None,
        },
    )
    .unwrap();
    wb
}

/// Return (namespace declarations, Requires) for every `mc:Choice` element.
fn mc_choices(xml: &str) -> Vec<(Vec<(String, String)>, String)> {
    let mut reader = Reader::from_str(xml);
    let mut buffer = Vec::new();
    let mut choices = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).expect("valid XML") {
            Event::Start(element) | Event::Empty(element)
                if element.name().as_ref() == b"mc:Choice" =>
            {
                let mut declarations = Vec::new();
                let mut requires = String::new();
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    let value = String::from_utf8_lossy(&attribute.value).into_owned();
                    if let Some(prefix) = key.strip_prefix("xmlns:") {
                        declarations.push((prefix.to_string(), value));
                    } else if key == "Requires" {
                        requires = value;
                    }
                }
                choices.push((declarations, requires));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    choices
}

fn ext_block<'a>(xml: &'a str, uri: &str) -> Option<&'a str> {
    let start = xml.to_ascii_lowercase().find(&uri.to_ascii_lowercase())?;
    let start = xml[..start].rfind("<ext")?;
    let end = xml[start..].find("</ext>")? + start;
    Some(&xml[start..end])
}

fn assert_table_slicer_package(buffer: &[u8]) {
    let drawing = required_part(buffer, "xl/drawings/drawing1.xml");
    let choices = mc_choices(&drawing);
    assert_eq!(choices.len(), 1, "{drawing}");
    for (declarations, requires) in &choices {
        for prefix in requires.split_whitespace() {
            assert!(
                declarations.iter().any(|(declared, _)| declared == prefix),
                "mc:Choice Requires=\"{requires}\" must declare xmlns:{prefix}: {drawing}"
            );
        }
    }

    for name in ["xl/slicers/slicer1.xml", "xl/slicerCaches/slicerCache1.xml"] {
        assert!(part(buffer, name).is_some(), "missing {name}");
    }
    let cache = required_part(buffer, "xl/slicerCaches/slicerCache1.xml");
    assert!(cache.contains("tableSlicerCache tableId=\"1\" column=\"1\""));

    let content_types = required_part(buffer, "[Content_Types].xml");
    assert!(content_types.contains(
        "PartName=\"/xl/slicers/slicer1.xml\" ContentType=\"application/vnd.ms-excel.slicer+xml\""
    ));
    assert!(content_types.contains(
        "PartName=\"/xl/slicerCaches/slicerCache1.xml\" ContentType=\"application/vnd.ms-excel.slicerCache+xml\""
    ));
    assert!(content_types.contains("/xl/drawings/drawing1.xml"));

    let workbook = required_part(buffer, "xl/workbook.xml");
    let workbook_rels = required_part(buffer, "xl/_rels/workbook.xml.rels");
    let cache_rid = workbook_rels
        .split("<Relationship ")
        .find(|rel| rel.contains("slicerCaches/slicerCache1.xml"))
        .and_then(|rel| rel.split("Id=\"").nth(1))
        .and_then(|rest| rest.split('"').next())
        .expect("workbook slicer cache relationship");
    let caches = ext_block(&workbook, "46BE6895-7355-4a93-B00E-2C351335B9C9")
        .expect("table slicer caches extension");
    assert!(
        caches.contains(&format!("xmlns:x15=\"{X15_NS}\"")),
        "{caches}"
    );
    assert!(caches.contains("<x15:slicerCaches"), "{caches}");
    assert!(
        caches.contains(&format!("r:id=\"{cache_rid}\"")),
        "{caches}"
    );
    assert!(
        ext_block(&workbook, "BBE1A952-AA13-448E-AADC-164F8A28A991").is_none(),
        "table caches must not be listed as pivot caches: {workbook}"
    );
    assert_eq!(workbook.matches("<x14:slicerCache ").count(), 1);

    let sheet = required_part(buffer, "xl/worksheets/sheet1.xml");
    let sheet_rels = required_part(buffer, "xl/worksheets/_rels/sheet1.xml.rels");
    assert!(sheet_rels.contains("../drawings/drawing1.xml"));
    let slicer_rid = sheet_rels
        .split("<Relationship ")
        .find(|rel| rel.contains("../slicers/slicer1.xml"))
        .and_then(|rel| rel.split("Id=\"").nth(1))
        .and_then(|rest| rest.split('"').next())
        .expect("worksheet slicer relationship");
    let list = ext_block(&sheet, "3A4CF648-6AED-40f4-86FF-DC5316D8AED3")
        .expect("table slicer list extension");
    assert!(list.contains(&format!("<x14:slicerList xmlns:x14=\"{X14_NS}\"")));
    assert!(list.contains(&format!("r:id=\"{slicer_rid}\"")), "{list}");
    assert!(
        ext_block(&sheet, "A8765BA9-456A-4DAB-B4F3-ACF838C121DE").is_none(),
        "table slicers must not be listed as pivot slicers: {sheet}"
    );
    assert_eq!(sheet.matches("<x14:slicer ").count(), 1);
}

#[test]
fn table_slicer_package_uses_table_slicer_extensions_and_scoped_namespaces() {
    let buffer = slicer_workbook().save_to_buffer().unwrap();
    assert_table_slicer_package(&buffer);
}

#[test]
fn table_slicer_extensions_survive_open_and_save_without_duplicates() {
    let buffer = slicer_workbook().save_to_buffer().unwrap();
    let mut reopened = Workbook::open_from_buffer(&buffer).unwrap();
    reopened.set_cell_value("Sheet1", "D1", "touched").unwrap();
    let saved = reopened.save_to_buffer().unwrap();
    assert_table_slicer_package(&saved);

    let untouched = Workbook::open_from_buffer(&buffer)
        .unwrap()
        .save_to_buffer()
        .unwrap();
    assert_table_slicer_package(&untouched);
}

fn threaded_input(author: &str, text: &str, parent_id: Option<String>) -> ThreadedCommentInput {
    ThreadedCommentInput {
        author: author.to_string(),
        text: text.to_string(),
        parent_id,
    }
}

fn legacy_comments(buffer: &[u8]) -> Vec<(String, String, String)> {
    let Some(xml) = part(buffer, "xl/comments1.xml") else {
        return Vec::new();
    };
    let comments: sheetkit_xml::comments::Comments = quick_xml::de::from_str(&xml).unwrap();
    let mut entries: Vec<(String, String, String)> = comments
        .comment_list
        .comments
        .iter()
        .map(|c| {
            let author = comments.authors.authors[c.author_id as usize].clone();
            let text = c.text.runs.iter().map(|r| r.t.as_str()).collect::<String>();
            (c.r#ref.clone(), author, text)
        })
        .collect();
    entries.sort();
    entries
}

#[test]
fn threaded_comments_are_mirrored_by_legacy_placeholder_comments() {
    let mut wb = Workbook::new();
    let root = wb
        .add_threaded_comment("Sheet1", "A1", &threaded_input("Alice", "root", None))
        .unwrap();
    wb.add_threaded_comment(
        "Sheet1",
        "A1",
        &threaded_input("Bob", "reply", Some(root.clone())),
    )
    .unwrap();
    let other = wb
        .add_threaded_comment("Sheet1", "B2", &threaded_input("Alice", "other", None))
        .unwrap();
    wb.add_comment(
        "Sheet1",
        &CommentConfig {
            cell: "D4".to_string(),
            author: "Carol".to_string(),
            text: "plain note".to_string(),
        },
    )
    .unwrap();
    let buffer = wb.save_to_buffer().unwrap();

    let sheet = required_part(&buffer, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains("<legacyDrawing r:id="), "{sheet}");
    let content_types = required_part(&buffer, "[Content_Types].xml");
    for needle in [
        "/xl/comments1.xml",
        "vmlDrawing",
        "/xl/threadedComments/threadedComment1.xml",
        "/xl/persons/person.xml",
    ] {
        assert!(content_types.contains(needle), "missing {needle}");
    }
    let vml = required_part(&buffer, "xl/drawings/vmlDrawing1.vml");
    assert_eq!(vml.matches("<x:ClientData ObjectType=\"Note\"").count(), 3);

    let notes = Workbook::open_from_buffer(&buffer)
        .unwrap()
        .get_comments("Sheet1")
        .unwrap();
    assert_eq!(notes.len(), 1, "legacy twins are not listed as notes");
    assert_eq!(notes[0].cell, "D4");

    let comments = legacy_comments(&buffer);
    assert_eq!(comments.len(), 3, "{comments:?}");
    let a1 = comments.iter().find(|c| c.0 == "A1").unwrap();
    assert_eq!(a1.1, format!("tc={root}"));
    assert!(a1
        .2
        .starts_with("[Threaded comment]\n\nYour version of Excel"));
    assert!(
        a1.2.ends_with("\n\nComment:\n    root\nReply:\n    reply"),
        "{}",
        a1.2
    );
    let b2 = comments.iter().find(|c| c.0 == "B2").unwrap();
    assert_eq!(b2.1, format!("tc={other}"));

    let mut reopened = Workbook::open_from_buffer(&buffer).unwrap();
    reopened.delete_threaded_comment("Sheet1", &other).unwrap();
    let saved = reopened.save_to_buffer().unwrap();
    let comments = legacy_comments(&saved);
    let cells: Vec<&str> = comments.iter().map(|c| c.0.as_str()).collect();
    assert_eq!(
        cells,
        ["A1", "D4"],
        "plain notes survive, the B2 twin is removed"
    );
}

#[test]
fn reopened_threaded_comments_do_not_duplicate_their_legacy_twin() {
    let mut wb = Workbook::new();
    let root = wb
        .add_threaded_comment("Sheet1", "A1", &threaded_input("Alice", "root", None))
        .unwrap();
    let buffer = wb.save_to_buffer().unwrap();

    let untouched = Workbook::open_from_buffer(&buffer)
        .unwrap()
        .save_to_buffer()
        .unwrap();
    assert_eq!(legacy_comments(&untouched).len(), 1);
    assert!(part(&untouched, "xl/threadedComments/threadedComment1.xml").is_some());

    let mut reopened = Workbook::open_from_buffer(&untouched).unwrap();
    reopened
        .add_threaded_comment(
            "Sheet1",
            "A1",
            &threaded_input("Bob", "late reply", Some(root)),
        )
        .unwrap();
    let saved = reopened.save_to_buffer().unwrap();
    let comments = legacy_comments(&saved);
    assert_eq!(comments.len(), 1, "{comments:?}");
    assert!(
        comments[0].2.ends_with("Reply:\n    late reply"),
        "{}",
        comments[0].2
    );
    let threaded = required_part(&saved, "xl/threadedComments/threadedComment1.xml");
    assert_eq!(threaded.matches("<threadedComment ").count(), 2);
}

#[test]
fn table_totals_row_is_declared_and_populated() {
    let mut wb = Workbook::new();
    for (cell, value) in [("A2", "Apple"), ("A3", "Pear")] {
        wb.set_cell_value("Sheet1", cell, value).unwrap();
    }
    for (cell, value) in [("B2", 3.0), ("B3", 2.0), ("C2", 1.5), ("C3", 2.0)] {
        wb.set_cell_value("Sheet1", cell, value).unwrap();
    }
    let totals = |name: &str, function: Option<&str>, label: Option<&str>| TableColumn {
        name: name.to_string(),
        totals_row_function: function.map(str::to_string),
        totals_row_label: label.map(str::to_string),
    };
    wb.add_table(
        "Sheet1",
        &TableConfig {
            name: "Sales".to_string(),
            display_name: "Sales".to_string(),
            range: "A1:C4".to_string(),
            columns: vec![
                totals("Product", None, Some("Total")),
                totals("Qty", Some("sum"), None),
                totals("Price #1", Some("average"), None),
            ],
            ..TableConfig::default()
        },
    )
    .unwrap();
    let buffer = wb.save_to_buffer().unwrap();
    let table = required_part(&buffer, "xl/tables/table1.xml");
    assert!(table.contains("ref=\"A1:C4\""), "{table}");
    assert!(table.contains("totalsRowCount=\"1\""), "{table}");
    assert!(table.contains("<autoFilter ref=\"A1:C3\""), "{table}");

    let reopened = Workbook::open_from_buffer(&buffer).unwrap();
    assert_eq!(
        reopened.get_cell_value("Sheet1", "A4").unwrap(),
        sheetkit_core::cell::CellValue::String("Total".to_string())
    );
    let formula = |cell: &str| match reopened.get_cell_value("Sheet1", cell).unwrap() {
        sheetkit_core::cell::CellValue::Formula { expr, .. } => expr,
        other => panic!("expected formula in {cell}, got {other:?}"),
    };
    assert_eq!(formula("B4"), "SUBTOTAL(109,[Qty])");
    assert_eq!(formula("C4"), "SUBTOTAL(101,[Price '#1])");
}

#[test]
fn totals_formulas_escape_at_signs_in_column_names() {
    let mut wb = Workbook::new();
    wb.set_cell_value("Sheet1", "B2", 3.0).unwrap();
    wb.add_table(
        "Sheet1",
        &TableConfig {
            name: "T".to_string(),
            display_name: "T".to_string(),
            range: "B1:B3".to_string(),
            columns: vec![TableColumn {
                name: "@Qty".to_string(),
                totals_row_function: Some("sum".to_string()),
                totals_row_label: None,
            }],
            ..TableConfig::default()
        },
    )
    .unwrap();
    let buffer = wb.save_to_buffer().unwrap();
    let reopened = Workbook::open_from_buffer(&buffer).unwrap();
    match reopened.get_cell_value("Sheet1", "B3").unwrap() {
        sheetkit_core::cell::CellValue::Formula { expr, .. } => {
            // [@Qty] would be a current-row reference to a column named Qty.
            assert_eq!(expr, "SUBTOTAL(109,['@Qty])");
        }
        other => panic!("expected formula, got {other:?}"),
    }
}

#[test]
fn deleting_a_thread_keeps_an_ordinary_note_that_mentions_the_placeholder() {
    let mut wb = Workbook::new();
    let id = wb
        .add_threaded_comment("Sheet1", "A1", &threaded_input("Alice", "thread", None))
        .unwrap();
    // A foreign file may carry an ordinary note on a threaded cell whose text
    // starts with Excel's placeholder label. Rewrite the saved legacy part to
    // that state and reopen it.
    let saved = wb.save_to_buffer().unwrap();
    let patched = rewrite_part(&saved, "xl/comments1.xml", |xml| {
        let author_start = xml.find("<author>").unwrap();
        let author_end = xml.find("</author>").unwrap() + "</author>".len();
        let xml = format!(
            "{}<author>Bob</author>{}",
            &xml[..author_start],
            &xml[author_end..]
        );
        let text_start = xml.find("<text>").unwrap();
        let text_end = xml.find("</text>").unwrap() + "</text>".len();
        format!(
            "{}<text><t>[Threaded comment] is a literal label in this ordinary note</t></text>{}",
            &xml[..text_start],
            &xml[text_end..]
        )
    });
    let mut wb = Workbook::open_from_buffer(&patched).unwrap();
    let notes = wb.get_comments("Sheet1").unwrap();
    assert_eq!(notes.len(), 1, "{notes:?}");
    assert_eq!(notes[0].author, "Bob");
    wb.delete_threaded_comment("Sheet1", &id).unwrap();
    let notes = wb.get_comments("Sheet1").unwrap();
    assert_eq!(notes.len(), 1, "{notes:?}");
    assert_eq!(notes[0].author, "Bob");
}

#[test]
fn note_operations_on_a_threaded_cell_are_rejected_so_the_thread_keeps_its_twin() {
    let mut wb = Workbook::new();
    wb.add_threaded_comment("Sheet1", "A1", &threaded_input("Alice", "thread", None))
        .unwrap();
    let add = wb.add_comment(
        "Sheet1",
        &CommentConfig {
            cell: "A1".to_string(),
            author: "Carol".to_string(),
            text: "plain note".to_string(),
        },
    );
    assert!(
        matches!(add, Err(sheetkit_core::error::Error::InvalidArgument(_))),
        "{add:?}"
    );
    let remove = wb.remove_comment("Sheet1", "A1");
    assert!(
        matches!(remove, Err(sheetkit_core::error::Error::InvalidArgument(_))),
        "{remove:?}"
    );
    let twins = legacy_comments(&wb.save_to_buffer().unwrap());
    assert_eq!(twins.len(), 1, "{twins:?}");
    assert!(twins[0].1.starts_with("tc="), "{twins:?}");
}

#[test]
fn deleting_a_thread_root_removes_its_replies_and_twin() {
    let mut wb = Workbook::new();
    let root = wb
        .add_threaded_comment("Sheet1", "A1", &threaded_input("Alice", "root", None))
        .unwrap();
    wb.add_threaded_comment(
        "Sheet1",
        "A1",
        &threaded_input("Bob", "reply", Some(root.clone())),
    )
    .unwrap();
    assert_eq!(wb.get_threaded_comments("Sheet1").unwrap().len(), 2);
    wb.delete_threaded_comment("Sheet1", &root).unwrap();
    assert!(wb.get_threaded_comments("Sheet1").unwrap().is_empty());
    let saved = wb.save_to_buffer().unwrap();
    assert!(part(&saved, "xl/threadedComments/threadedComment1.xml").is_none());
    assert!(legacy_comments(&saved).is_empty());
}

#[test]
fn table_totals_row_config_is_validated() {
    let config = |range: &str, function: &str| TableConfig {
        name: "T".to_string(),
        display_name: "T".to_string(),
        range: range.to_string(),
        columns: vec![TableColumn {
            name: "Qty".to_string(),
            totals_row_function: Some(function.to_string()),
            totals_row_label: None,
        }],
        ..TableConfig::default()
    };
    let mut wb = Workbook::new();
    assert!(wb.add_table("Sheet1", &config("A1:A2", "sum")).is_err());
    assert!(wb.add_table("Sheet1", &config("A1:A3", "median")).is_err());
    assert!(wb.add_table("Sheet1", &config("A1:A3", "custom")).is_err());
    wb.add_table("Sheet1", &config("A1:A3", "sum")).unwrap();
}

#[test]
fn table_without_totals_keeps_autofilter_on_full_range() {
    let mut wb = Workbook::new();
    wb.add_table(
        "Sheet1",
        &TableConfig {
            name: "Plain".to_string(),
            display_name: "Plain".to_string(),
            range: "A1:A3".to_string(),
            columns: vec![column("Qty")],
            ..TableConfig::default()
        },
    )
    .unwrap();
    let table = required_part(&wb.save_to_buffer().unwrap(), "xl/tables/table1.xml");
    assert!(!table.contains("totalsRowCount"), "{table}");
    assert!(table.contains("<autoFilter ref=\"A1:A3\""), "{table}");
}

fn rewrite_part(buffer: &[u8], name: &str, edit: impl Fn(&str) -> String) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(buffer.to_vec())).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        let bytes = if entry.name() == name {
            edit(&String::from_utf8(bytes).unwrap()).into_bytes()
        } else {
            bytes
        };
        out.start_file(entry.name().to_string(), options).unwrap();
        std::io::Write::write_all(&mut out, &bytes).unwrap();
    }
    out.finish().unwrap().into_inner()
}

#[test]
fn unparsable_slicer_cache_keeps_the_source_extension_on_the_sheet() {
    let buffer = slicer_workbook().save_to_buffer().unwrap();
    let broken = rewrite_part(&buffer, "xl/slicerCaches/slicerCache1.xml", |xml| {
        xml.replacen("sourceName=", "srcName=", 1)
    });
    for mode in [
        sheetkit_core::workbook::ReadMode::Eager,
        sheetkit_core::workbook::ReadMode::Lazy,
    ] {
        let options = sheetkit_core::workbook::OpenOptions::new().read_mode(mode);
        let wb = Workbook::open_from_buffer_with_options(&broken, &options).unwrap();
        let saved = wb.save_to_buffer().unwrap();
        let sheet = required_part(&saved, "xl/worksheets/sheet1.xml");
        assert!(
            ext_block(&sheet, "3A4CF648-6AED-40f4-86FF-DC5316D8AED3").is_some(),
            "{mode:?}: table slicer list must survive an unparsable cache: {sheet}"
        );
        assert!(
            ext_block(&sheet, "A8765BA9-456A-4DAB-B4F3-ACF838C121DE").is_none(),
            "{mode:?}: slicer must not move to the pivot list: {sheet}"
        );
        let workbook = required_part(&saved, "xl/workbook.xml");
        assert!(
            workbook.contains("x15:slicerCaches"),
            "{mode:?}: {workbook}"
        );
    }
}

#[test]
fn threaded_comment_on_a_cell_with_a_note_is_rejected_and_the_note_survives() {
    let mut wb = Workbook::new();
    wb.add_comment(
        "Sheet1",
        &CommentConfig {
            cell: "A1".to_string(),
            author: "Reviewer".to_string(),
            text: "plain note".to_string(),
        },
    )
    .unwrap();
    let error = wb
        .add_threaded_comment("Sheet1", "A1", &threaded_input("Alice", "thread", None))
        .unwrap_err();
    assert!(
        matches!(error, sheetkit_core::error::Error::InvalidArgument(_)),
        "{error}"
    );
    let notes = wb.get_comments("Sheet1").unwrap();
    assert_eq!(notes.len(), 1);
    assert_eq!(notes[0].text, "plain note");
    assert!(wb.get_threaded_comments("Sheet1").unwrap().is_empty());
    // A thread on a free cell still works, and its twin is hidden from notes.
    wb.add_threaded_comment("Sheet1", "B1", &threaded_input("Alice", "thread", None))
        .unwrap();
    let notes = wb.get_comments("Sheet1").unwrap();
    assert_eq!(notes.len(), 1, "{notes:?}");
    assert_eq!(notes[0].cell, "A1");
}
