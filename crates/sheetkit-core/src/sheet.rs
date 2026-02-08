//! Sheet management utilities.
//!
//! Contains validation helpers and internal functions used by [`crate::workbook::Workbook`]
//! for creating, deleting, renaming, and copying worksheets.

use sheetkit_xml::content_types::{mime_types, ContentTypeOverride, ContentTypes};
use sheetkit_xml::relationships::{rel_types, Relationship, Relationships};
use sheetkit_xml::workbook::{SheetEntry, WorkbookXml};
use sheetkit_xml::worksheet::WorksheetXml;

use crate::error::{Error, Result};
use crate::utils::constants::{MAX_SHEET_NAME_LENGTH, SHEET_NAME_INVALID_CHARS};

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

/// Validate a sheet name according to Excel rules.
///
/// A valid sheet name must:
/// - Be non-empty
/// - Be at most [`MAX_SHEET_NAME_LENGTH`] (31) characters
/// - Not contain any of the characters `: \ / ? * [ ]`
/// - Not start or end with a single quote (`'`)
pub fn validate_sheet_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(Error::InvalidSheetName("sheet name cannot be empty".into()));
    }
    if name.len() > MAX_SHEET_NAME_LENGTH {
        return Err(Error::InvalidSheetName(format!(
            "sheet name '{}' exceeds {} characters",
            name, MAX_SHEET_NAME_LENGTH
        )));
    }
    for ch in SHEET_NAME_INVALID_CHARS {
        if name.contains(*ch) {
            return Err(Error::InvalidSheetName(format!(
                "sheet name '{}' contains invalid character '{}'",
                name, ch
            )));
        }
    }
    if name.starts_with('\'') || name.ends_with('\'') {
        return Err(Error::InvalidSheetName(format!(
            "sheet name '{}' cannot start or end with a single quote",
            name
        )));
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// ID generation helpers
// ---------------------------------------------------------------------------

/// Generate the next available rId for workbook relationships.
///
/// Scans existing relationship IDs of the form `rIdN` and returns `rId{max+1}`.
pub fn next_rid(existing_rels: &[Relationship]) -> String {
    let max = existing_rels
        .iter()
        .filter_map(|r| r.id.strip_prefix("rId").and_then(|n| n.parse::<u32>().ok()))
        .max()
        .unwrap_or(0);
    format!("rId{}", max + 1)
}

/// Generate the next available sheet ID.
///
/// Sheet IDs in a workbook are unique but not necessarily contiguous. This
/// function returns one greater than the current maximum.
pub fn next_sheet_id(existing_sheets: &[SheetEntry]) -> u32 {
    existing_sheets
        .iter()
        .map(|s| s.sheet_id)
        .max()
        .unwrap_or(0)
        + 1
}

// ---------------------------------------------------------------------------
// Sheet operations (operate on workbook internals)
// ---------------------------------------------------------------------------

/// Find the index (0-based) of a sheet by name.
pub fn find_sheet_index(worksheets: &[(String, WorksheetXml)], name: &str) -> Option<usize> {
    worksheets.iter().position(|(n, _)| n == name)
}

/// Add a new sheet. Returns the 0-based index of the new sheet.
///
/// This function performs all bookkeeping: adds entries to the sheet list,
/// workbook relationships, and content type overrides.
pub fn add_sheet(
    workbook_xml: &mut WorkbookXml,
    workbook_rels: &mut Relationships,
    content_types: &mut ContentTypes,
    worksheets: &mut Vec<(String, WorksheetXml)>,
    name: &str,
    worksheet_data: WorksheetXml,
) -> Result<usize> {
    validate_sheet_name(name)?;

    // Check for duplicate name
    if worksheets.iter().any(|(n, _)| n == name) {
        return Err(Error::SheetAlreadyExists {
            name: name.to_string(),
        });
    }

    let rid = next_rid(&workbook_rels.relationships);
    let sheet_id = next_sheet_id(&workbook_xml.sheets.sheets);
    let sheet_number = worksheets.len() + 1;
    let target = format!("worksheets/sheet{}.xml", sheet_number);

    // Add SheetEntry to workbook XML
    workbook_xml.sheets.sheets.push(SheetEntry {
        name: name.to_string(),
        sheet_id,
        state: None,
        r_id: rid.clone(),
    });

    // Add Relationship
    workbook_rels.relationships.push(Relationship {
        id: rid,
        rel_type: rel_types::WORKSHEET.to_string(),
        target: target.clone(),
        target_mode: None,
    });

    // Add ContentType override
    content_types.overrides.push(ContentTypeOverride {
        part_name: format!("/xl/{}", target),
        content_type: mime_types::WORKSHEET.to_string(),
    });

    // Add to worksheets vec
    worksheets.push((name.to_string(), worksheet_data));

    Ok(worksheets.len() - 1)
}

/// Delete a sheet by name.
///
/// Returns an error if the sheet does not exist or if it is the last remaining sheet.
pub fn delete_sheet(
    workbook_xml: &mut WorkbookXml,
    workbook_rels: &mut Relationships,
    content_types: &mut ContentTypes,
    worksheets: &mut Vec<(String, WorksheetXml)>,
    name: &str,
) -> Result<()> {
    let idx = find_sheet_index(worksheets, name).ok_or_else(|| Error::SheetNotFound {
        name: name.to_string(),
    })?;

    if worksheets.len() <= 1 {
        return Err(Error::InvalidSheetName(
            "cannot delete the last sheet in a workbook".into(),
        ));
    }

    // Get the rId before removing
    let r_id = workbook_xml.sheets.sheets[idx].r_id.clone();

    // Remove from worksheets vec
    worksheets.remove(idx);

    // Remove from workbook XML sheets
    workbook_xml.sheets.sheets.remove(idx);

    // Remove matching relationship
    workbook_rels.relationships.retain(|r| r.id != r_id);

    // Rebuild content type overrides for worksheets to match new indices
    rebuild_content_type_overrides(content_types, worksheets.len());

    // Rebuild relationship targets to match new worksheet indices
    rebuild_worksheet_relationships(workbook_xml, workbook_rels);

    Ok(())
}

/// Rename a sheet.
pub fn rename_sheet(
    workbook_xml: &mut WorkbookXml,
    worksheets: &mut [(String, WorksheetXml)],
    old_name: &str,
    new_name: &str,
) -> Result<()> {
    validate_sheet_name(new_name)?;

    let idx = find_sheet_index(worksheets, old_name).ok_or_else(|| Error::SheetNotFound {
        name: old_name.to_string(),
    })?;

    // Check new name doesn't conflict
    if worksheets.iter().any(|(n, _)| n == new_name) {
        return Err(Error::SheetAlreadyExists {
            name: new_name.to_string(),
        });
    }

    // Update in worksheets vec
    worksheets[idx].0 = new_name.to_string();

    // Update in workbook XML
    workbook_xml.sheets.sheets[idx].name = new_name.to_string();

    Ok(())
}

/// Copy a sheet, returning the 0-based index of the new copy.
pub fn copy_sheet(
    workbook_xml: &mut WorkbookXml,
    workbook_rels: &mut Relationships,
    content_types: &mut ContentTypes,
    worksheets: &mut Vec<(String, WorksheetXml)>,
    source_name: &str,
    target_name: &str,
) -> Result<usize> {
    let source_idx =
        find_sheet_index(worksheets, source_name).ok_or_else(|| Error::SheetNotFound {
            name: source_name.to_string(),
        })?;

    let cloned_data = worksheets[source_idx].1.clone();

    add_sheet(
        workbook_xml,
        workbook_rels,
        content_types,
        worksheets,
        target_name,
        cloned_data,
    )
}

// ---------------------------------------------------------------------------
// Active sheet helpers
// ---------------------------------------------------------------------------

/// Get the active sheet index (0-based) from bookViews, defaulting to 0.
pub fn active_sheet_index(workbook_xml: &WorkbookXml) -> usize {
    workbook_xml
        .book_views
        .as_ref()
        .and_then(|bv| bv.workbook_views.first())
        .and_then(|v| v.active_tab)
        .unwrap_or(0) as usize
}

/// Set the active sheet by index in bookViews.
pub fn set_active_sheet_index(workbook_xml: &mut WorkbookXml, index: u32) {
    use sheetkit_xml::workbook::{BookViews, WorkbookView};

    let book_views = workbook_xml.book_views.get_or_insert_with(|| BookViews {
        workbook_views: vec![WorkbookView {
            x_window: None,
            y_window: None,
            window_width: None,
            window_height: None,
            active_tab: Some(0),
        }],
    });

    if let Some(view) = book_views.workbook_views.first_mut() {
        view.active_tab = Some(index);
    }
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/// Rebuild content type overrides for worksheets so they match the current
/// worksheet indices (sheet1.xml, sheet2.xml, ...).
fn rebuild_content_type_overrides(content_types: &mut ContentTypes, sheet_count: usize) {
    // Remove all existing worksheet overrides
    content_types
        .overrides
        .retain(|o| o.content_type != mime_types::WORKSHEET);

    // Re-add them with correct indices
    for i in 1..=sheet_count {
        content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/xl/worksheets/sheet{}.xml", i),
            content_type: mime_types::WORKSHEET.to_string(),
        });
    }
}

/// Rebuild worksheet relationship targets so they match the current worksheet indices.
fn rebuild_worksheet_relationships(
    workbook_xml: &mut WorkbookXml,
    workbook_rels: &mut Relationships,
) {
    // Collect the rIds from the sheet entries (which are already in correct order)
    let sheet_rids: Vec<String> = workbook_xml
        .sheets
        .sheets
        .iter()
        .map(|s| s.r_id.clone())
        .collect();

    // Update relationship targets for each sheet
    for (i, rid) in sheet_rids.iter().enumerate() {
        if let Some(rel) = workbook_rels
            .relationships
            .iter_mut()
            .find(|r| r.id == *rid)
        {
            rel.target = format!("worksheets/sheet{}.xml", i + 1);
        }
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use sheetkit_xml::content_types::ContentTypes;
    use sheetkit_xml::relationships;
    use sheetkit_xml::workbook::WorkbookXml;
    use sheetkit_xml::worksheet::WorksheetXml;

    // === Validation tests ===

    #[test]
    fn test_validate_empty_name() {
        let result = validate_sheet_name("");
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(
            err_msg.contains("empty"),
            "Error should mention empty: {err_msg}"
        );
    }

    #[test]
    fn test_validate_too_long_name() {
        let long_name = "a".repeat(32);
        let result = validate_sheet_name(&long_name);
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(
            err_msg.contains("exceeds"),
            "Error should mention exceeds: {err_msg}"
        );
    }

    #[test]
    fn test_validate_exactly_max_length_is_ok() {
        let name = "a".repeat(MAX_SHEET_NAME_LENGTH);
        assert!(validate_sheet_name(&name).is_ok());
    }

    #[test]
    fn test_validate_invalid_chars() {
        for ch in SHEET_NAME_INVALID_CHARS {
            let name = format!("Sheet{}", ch);
            let result = validate_sheet_name(&name);
            assert!(result.is_err(), "Name with '{}' should be invalid", ch);
        }
    }

    #[test]
    fn test_validate_single_quote_boundary() {
        assert!(validate_sheet_name("'Sheet").is_err());
        assert!(validate_sheet_name("Sheet'").is_err());
        assert!(validate_sheet_name("'Sheet'").is_err());
        // Single quote in the middle is OK
        assert!(validate_sheet_name("She'et").is_ok());
    }

    #[test]
    fn test_validate_valid_name() {
        assert!(validate_sheet_name("Sheet1").is_ok());
        assert!(validate_sheet_name("My Data").is_ok());
        assert!(validate_sheet_name("Q1-2024").is_ok());
        assert!(validate_sheet_name("Sheet (2)").is_ok());
    }

    // === ID generation tests ===

    #[test]
    fn test_next_rid() {
        let rels = vec![
            Relationship {
                id: "rId1".to_string(),
                rel_type: "".to_string(),
                target: "".to_string(),
                target_mode: None,
            },
            Relationship {
                id: "rId3".to_string(),
                rel_type: "".to_string(),
                target: "".to_string(),
                target_mode: None,
            },
        ];
        assert_eq!(next_rid(&rels), "rId4");
    }

    #[test]
    fn test_next_rid_empty() {
        assert_eq!(next_rid(&[]), "rId1");
    }

    #[test]
    fn test_next_sheet_id() {
        let sheets = vec![
            SheetEntry {
                name: "Sheet1".to_string(),
                sheet_id: 1,
                state: None,
                r_id: "rId1".to_string(),
            },
            SheetEntry {
                name: "Sheet2".to_string(),
                sheet_id: 5,
                state: None,
                r_id: "rId2".to_string(),
            },
        ];
        assert_eq!(next_sheet_id(&sheets), 6);
    }

    #[test]
    fn test_next_sheet_id_empty() {
        assert_eq!(next_sheet_id(&[]), 1);
    }

    // === Sheet operation tests ===

    /// Helper to create default test workbook internals.
    fn test_workbook_parts() -> (
        WorkbookXml,
        Relationships,
        ContentTypes,
        Vec<(String, WorksheetXml)>,
    ) {
        let workbook_xml = WorkbookXml::default();
        let workbook_rels = relationships::workbook_rels();
        let content_types = ContentTypes::default();
        let worksheets = vec![("Sheet1".to_string(), WorksheetXml::default())];
        (workbook_xml, workbook_rels, content_types, worksheets)
    }

    #[test]
    fn test_add_sheet_basic() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        let idx = add_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "Sheet2",
            WorksheetXml::default(),
        )
        .unwrap();

        assert_eq!(idx, 1);
        assert_eq!(ws.len(), 2);
        assert_eq!(ws[1].0, "Sheet2");
        assert_eq!(wb_xml.sheets.sheets.len(), 2);
        assert_eq!(wb_xml.sheets.sheets[1].name, "Sheet2");

        // Verify a new relationship was added
        let ws_rels: Vec<_> = wb_rels
            .relationships
            .iter()
            .filter(|r| r.rel_type == rel_types::WORKSHEET)
            .collect();
        assert_eq!(ws_rels.len(), 2);

        // Verify content type override was added
        let ws_overrides: Vec<_> = ct
            .overrides
            .iter()
            .filter(|o| o.content_type == mime_types::WORKSHEET)
            .collect();
        assert_eq!(ws_overrides.len(), 2);
    }

    #[test]
    fn test_add_sheet_duplicate_returns_error() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        let result = add_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "Sheet1",
            WorksheetXml::default(),
        );

        assert!(result.is_err());
        assert!(
            matches!(result.unwrap_err(), Error::SheetAlreadyExists { name } if name == "Sheet1")
        );
    }

    #[test]
    fn test_add_sheet_invalid_name_returns_error() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        let result = add_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "Bad[Name",
            WorksheetXml::default(),
        );

        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::InvalidSheetName(_)));
    }

    #[test]
    fn test_delete_sheet_basic() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        // Add a second sheet first
        add_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "Sheet2",
            WorksheetXml::default(),
        )
        .unwrap();

        assert_eq!(ws.len(), 2);

        // Delete the first sheet
        delete_sheet(&mut wb_xml, &mut wb_rels, &mut ct, &mut ws, "Sheet1").unwrap();

        assert_eq!(ws.len(), 1);
        assert_eq!(ws[0].0, "Sheet2");
        assert_eq!(wb_xml.sheets.sheets.len(), 1);
        assert_eq!(wb_xml.sheets.sheets[0].name, "Sheet2");

        // Verify worksheet relationships were rebuilt
        let ws_rels: Vec<_> = wb_rels
            .relationships
            .iter()
            .filter(|r| r.rel_type == rel_types::WORKSHEET)
            .collect();
        assert_eq!(ws_rels.len(), 1);

        // Verify content type overrides were rebuilt
        let ws_overrides: Vec<_> = ct
            .overrides
            .iter()
            .filter(|o| o.content_type == mime_types::WORKSHEET)
            .collect();
        assert_eq!(ws_overrides.len(), 1);
    }

    #[test]
    fn test_delete_last_sheet_returns_error() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        let result = delete_sheet(&mut wb_xml, &mut wb_rels, &mut ct, &mut ws, "Sheet1");
        assert!(result.is_err());
    }

    #[test]
    fn test_delete_nonexistent_sheet_returns_error() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        let result = delete_sheet(&mut wb_xml, &mut wb_rels, &mut ct, &mut ws, "Nonexistent");
        assert!(result.is_err());
        assert!(
            matches!(result.unwrap_err(), Error::SheetNotFound { name } if name == "Nonexistent")
        );
    }

    #[test]
    fn test_rename_sheet_basic() {
        let (mut wb_xml, _, _, mut ws) = test_workbook_parts();

        rename_sheet(&mut wb_xml, &mut ws, "Sheet1", "MySheet").unwrap();

        assert_eq!(ws[0].0, "MySheet");
        assert_eq!(wb_xml.sheets.sheets[0].name, "MySheet");
    }

    #[test]
    fn test_rename_sheet_to_existing_returns_error() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        add_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "Sheet2",
            WorksheetXml::default(),
        )
        .unwrap();

        let result = rename_sheet(&mut wb_xml, &mut ws, "Sheet1", "Sheet2");
        assert!(result.is_err());
        assert!(
            matches!(result.unwrap_err(), Error::SheetAlreadyExists { name } if name == "Sheet2")
        );
    }

    #[test]
    fn test_rename_nonexistent_sheet_returns_error() {
        let (mut wb_xml, _, _, mut ws) = test_workbook_parts();

        let result = rename_sheet(&mut wb_xml, &mut ws, "Nope", "NewName");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { name } if name == "Nope"));
    }

    #[test]
    fn test_copy_sheet_basic() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        let idx = copy_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "Sheet1",
            "Sheet1 Copy",
        )
        .unwrap();

        assert_eq!(idx, 1);
        assert_eq!(ws.len(), 2);
        assert_eq!(ws[1].0, "Sheet1 Copy");
        // The copied worksheet data should be a clone of the source
        assert_eq!(ws[1].1, ws[0].1);
    }

    #[test]
    fn test_copy_nonexistent_sheet_returns_error() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        let result = copy_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "Nonexistent",
            "Copy",
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_copy_sheet_to_existing_name_returns_error() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        let result = copy_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "Sheet1",
            "Sheet1",
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_find_sheet_index() {
        let ws: Vec<(String, WorksheetXml)> = vec![
            ("Sheet1".to_string(), WorksheetXml::default()),
            ("Sheet2".to_string(), WorksheetXml::default()),
        ];

        assert_eq!(find_sheet_index(&ws, "Sheet1"), Some(0));
        assert_eq!(find_sheet_index(&ws, "Sheet2"), Some(1));
        assert_eq!(find_sheet_index(&ws, "Sheet3"), None);
    }

    #[test]
    fn test_active_sheet_index_default() {
        let wb_xml = WorkbookXml::default();
        assert_eq!(active_sheet_index(&wb_xml), 0);
    }

    #[test]
    fn test_set_active_sheet_index() {
        let mut wb_xml = WorkbookXml::default();
        set_active_sheet_index(&mut wb_xml, 2);

        assert_eq!(active_sheet_index(&wb_xml), 2);
    }

    #[test]
    fn test_multiple_add_delete_consistency() {
        let (mut wb_xml, mut wb_rels, mut ct, mut ws) = test_workbook_parts();

        // Add 3 more sheets
        add_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "A",
            WorksheetXml::default(),
        )
        .unwrap();
        add_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "B",
            WorksheetXml::default(),
        )
        .unwrap();
        add_sheet(
            &mut wb_xml,
            &mut wb_rels,
            &mut ct,
            &mut ws,
            "C",
            WorksheetXml::default(),
        )
        .unwrap();

        assert_eq!(ws.len(), 4);

        // Delete "B" (middle sheet)
        delete_sheet(&mut wb_xml, &mut wb_rels, &mut ct, &mut ws, "B").unwrap();

        assert_eq!(ws.len(), 3);
        let names: Vec<&str> = ws.iter().map(|(n, _)| n.as_str()).collect();
        assert_eq!(names, vec!["Sheet1", "A", "C"]);

        // Verify internal consistency
        assert_eq!(wb_xml.sheets.sheets.len(), 3);
        let ws_rels: Vec<_> = wb_rels
            .relationships
            .iter()
            .filter(|r| r.rel_type == rel_types::WORKSHEET)
            .collect();
        assert_eq!(ws_rels.len(), 3);
        let ws_overrides: Vec<_> = ct
            .overrides
            .iter()
            .filter(|o| o.content_type == mime_types::WORKSHEET)
            .collect();
        assert_eq!(ws_overrides.len(), 3);
    }
}
