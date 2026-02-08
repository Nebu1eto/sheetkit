//! Sheet management utilities.
//!
//! Contains validation helpers and internal functions used by [`crate::workbook::Workbook`]
//! for creating, deleting, renaming, and copying worksheets.

use sheetkit_xml::content_types::{mime_types, ContentTypeOverride, ContentTypes};
use sheetkit_xml::relationships::{rel_types, Relationship, Relationships};
use sheetkit_xml::workbook::{SheetEntry, WorkbookXml};
use sheetkit_xml::worksheet::{
    Pane, Selection, SheetFormatPr, SheetPr, SheetProtection, SheetView, SheetViews, TabColor,
    WorksheetXml,
};

use crate::error::{Error, Result};
use crate::protection::legacy_password_hash;
use crate::utils::cell_ref::cell_name_to_coordinates;
use crate::utils::constants::{
    DEFAULT_ROW_HEIGHT, MAX_COLUMN_WIDTH, MAX_ROW_HEIGHT, MAX_SHEET_NAME_LENGTH,
    SHEET_NAME_INVALID_CHARS,
};

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

    if worksheets.iter().any(|(n, _)| n == name) {
        return Err(Error::SheetAlreadyExists {
            name: name.to_string(),
        });
    }

    let rid = next_rid(&workbook_rels.relationships);
    let sheet_id = next_sheet_id(&workbook_xml.sheets.sheets);
    let sheet_number = worksheets.len() + 1;
    let target = format!("worksheets/sheet{}.xml", sheet_number);

    workbook_xml.sheets.sheets.push(SheetEntry {
        name: name.to_string(),
        sheet_id,
        state: None,
        r_id: rid.clone(),
    });

    workbook_rels.relationships.push(Relationship {
        id: rid,
        rel_type: rel_types::WORKSHEET.to_string(),
        target: target.clone(),
        target_mode: None,
    });

    content_types.overrides.push(ContentTypeOverride {
        part_name: format!("/xl/{}", target),
        content_type: mime_types::WORKSHEET.to_string(),
    });

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

    let r_id = workbook_xml.sheets.sheets[idx].r_id.clone();

    worksheets.remove(idx);
    workbook_xml.sheets.sheets.remove(idx);
    workbook_rels.relationships.retain(|r| r.id != r_id);

    rebuild_content_type_overrides(content_types, worksheets.len());
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

    if worksheets.iter().any(|(n, _)| n == new_name) {
        return Err(Error::SheetAlreadyExists {
            name: new_name.to_string(),
        });
    }

    worksheets[idx].0 = new_name.to_string();
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

/// Configuration for sheet protection.
///
/// All boolean fields default to `false`, meaning the corresponding action is
/// forbidden when protection is enabled. Set a field to `true` to allow that
/// action even when the sheet is protected.
#[derive(Debug, Clone, Default)]
pub struct SheetProtectionConfig {
    /// Optional password. Hashed with the legacy Excel algorithm.
    pub password: Option<String>,
    /// Allow selecting locked cells.
    pub select_locked_cells: bool,
    /// Allow selecting unlocked cells.
    pub select_unlocked_cells: bool,
    /// Allow formatting cells.
    pub format_cells: bool,
    /// Allow formatting columns.
    pub format_columns: bool,
    /// Allow formatting rows.
    pub format_rows: bool,
    /// Allow inserting columns.
    pub insert_columns: bool,
    /// Allow inserting rows.
    pub insert_rows: bool,
    /// Allow inserting hyperlinks.
    pub insert_hyperlinks: bool,
    /// Allow deleting columns.
    pub delete_columns: bool,
    /// Allow deleting rows.
    pub delete_rows: bool,
    /// Allow sorting.
    pub sort: bool,
    /// Allow using auto-filter.
    pub auto_filter: bool,
    /// Allow using pivot tables.
    pub pivot_tables: bool,
}

/// Protect a sheet with optional password and permission settings.
///
/// When a sheet is protected, users cannot edit cells unless specific
/// permissions are granted via the config. The password is hashed using
/// the legacy Excel algorithm.
pub fn protect_sheet(ws: &mut WorksheetXml, config: &SheetProtectionConfig) -> Result<()> {
    let hashed = config.password.as_ref().map(|p| {
        let h = legacy_password_hash(p);
        format!("{:04X}", h)
    });

    let to_opt = |v: bool| if v { Some(true) } else { None };

    ws.sheet_protection = Some(SheetProtection {
        password: hashed,
        sheet: Some(true),
        objects: Some(true),
        scenarios: Some(true),
        select_locked_cells: to_opt(config.select_locked_cells),
        select_unlocked_cells: to_opt(config.select_unlocked_cells),
        format_cells: to_opt(config.format_cells),
        format_columns: to_opt(config.format_columns),
        format_rows: to_opt(config.format_rows),
        insert_columns: to_opt(config.insert_columns),
        insert_rows: to_opt(config.insert_rows),
        insert_hyperlinks: to_opt(config.insert_hyperlinks),
        delete_columns: to_opt(config.delete_columns),
        delete_rows: to_opt(config.delete_rows),
        sort: to_opt(config.sort),
        auto_filter: to_opt(config.auto_filter),
        pivot_tables: to_opt(config.pivot_tables),
    });

    Ok(())
}

/// Remove sheet protection.
pub fn unprotect_sheet(ws: &mut WorksheetXml) -> Result<()> {
    ws.sheet_protection = None;
    Ok(())
}

/// Check if a sheet is protected.
pub fn is_sheet_protected(ws: &WorksheetXml) -> bool {
    ws.sheet_protection
        .as_ref()
        .and_then(|p| p.sheet)
        .unwrap_or(false)
}

/// Set the tab color of a sheet using an RGB hex string (e.g. "FF0000" for red).
pub fn set_tab_color(ws: &mut WorksheetXml, rgb: &str) -> Result<()> {
    let sheet_pr = ws.sheet_pr.get_or_insert_with(SheetPr::default);
    sheet_pr.tab_color = Some(TabColor {
        rgb: Some(rgb.to_string()),
        theme: None,
        indexed: None,
    });
    Ok(())
}

/// Get the tab color of a sheet as an RGB hex string.
pub fn get_tab_color(ws: &WorksheetXml) -> Option<String> {
    ws.sheet_pr
        .as_ref()
        .and_then(|pr| pr.tab_color.as_ref())
        .and_then(|tc| tc.rgb.clone())
}

/// Set the default row height for a sheet.
///
/// Returns an error if the height exceeds [`MAX_ROW_HEIGHT`] (409).
pub fn set_default_row_height(ws: &mut WorksheetXml, height: f64) -> Result<()> {
    if height > MAX_ROW_HEIGHT {
        return Err(Error::RowHeightExceeded {
            height,
            max: MAX_ROW_HEIGHT,
        });
    }
    let fmt = ws.sheet_format_pr.get_or_insert(SheetFormatPr {
        default_row_height: DEFAULT_ROW_HEIGHT,
        default_col_width: None,
        custom_height: None,
        outline_level_row: None,
        outline_level_col: None,
    });
    fmt.default_row_height = height;
    Ok(())
}

/// Get the default row height for a sheet.
///
/// Returns [`DEFAULT_ROW_HEIGHT`] (15.0) if no sheet format properties are set.
pub fn get_default_row_height(ws: &WorksheetXml) -> f64 {
    ws.sheet_format_pr
        .as_ref()
        .map(|f| f.default_row_height)
        .unwrap_or(DEFAULT_ROW_HEIGHT)
}

/// Set the default column width for a sheet.
///
/// Returns an error if the width exceeds [`MAX_COLUMN_WIDTH`] (255).
pub fn set_default_col_width(ws: &mut WorksheetXml, width: f64) -> Result<()> {
    if width > MAX_COLUMN_WIDTH {
        return Err(Error::ColumnWidthExceeded {
            width,
            max: MAX_COLUMN_WIDTH,
        });
    }
    let fmt = ws.sheet_format_pr.get_or_insert(SheetFormatPr {
        default_row_height: DEFAULT_ROW_HEIGHT,
        default_col_width: None,
        custom_height: None,
        outline_level_row: None,
        outline_level_col: None,
    });
    fmt.default_col_width = Some(width);
    Ok(())
}

/// Get the default column width for a sheet.
///
/// Returns `None` if no default column width has been set.
pub fn get_default_col_width(ws: &WorksheetXml) -> Option<f64> {
    ws.sheet_format_pr
        .as_ref()
        .and_then(|f| f.default_col_width)
}

/// Set freeze panes on a worksheet.
///
/// The cell reference indicates the top-left cell of the scrollable (unfrozen) area.
/// For example, `"A2"` freezes row 1, `"B1"` freezes column A, and `"B2"` freezes
/// both row 1 and column A.
///
/// Returns an error if the cell reference is invalid or is `"A1"` (which would
/// freeze nothing).
pub fn set_panes(ws: &mut WorksheetXml, cell: &str) -> Result<()> {
    let (col, row) = cell_name_to_coordinates(cell)?;

    if col == 1 && row == 1 {
        return Err(Error::InvalidCellReference(
            "freeze pane at A1 has no effect".to_string(),
        ));
    }

    let x_split = col - 1;
    let y_split = row - 1;

    let active_pane = match (x_split > 0, y_split > 0) {
        (true, true) => "bottomRight",
        (true, false) => "topRight",
        (false, true) => "bottomLeft",
        (false, false) => unreachable!(),
    };

    let pane = Pane {
        x_split: if x_split > 0 { Some(x_split) } else { None },
        y_split: if y_split > 0 { Some(y_split) } else { None },
        top_left_cell: Some(cell.to_string()),
        active_pane: Some(active_pane.to_string()),
        state: Some("frozen".to_string()),
    };

    let selection = Selection {
        pane: Some(active_pane.to_string()),
        active_cell: Some(cell.to_string()),
        sqref: Some(cell.to_string()),
    };

    let sheet_views = ws.sheet_views.get_or_insert_with(|| SheetViews {
        sheet_views: vec![SheetView {
            tab_selected: None,
            zoom_scale: None,
            workbook_view_id: 0,
            pane: None,
            selection: vec![],
        }],
    });

    if let Some(view) = sheet_views.sheet_views.first_mut() {
        view.pane = Some(pane);
        view.selection = vec![selection];
    }

    Ok(())
}

/// Remove any freeze or split panes from a worksheet.
pub fn unset_panes(ws: &mut WorksheetXml) {
    if let Some(ref mut sheet_views) = ws.sheet_views {
        for view in &mut sheet_views.sheet_views {
            view.pane = None;
            // Reset selection to default (no pane attribute).
            view.selection = vec![];
        }
    }
}

/// Get the current freeze pane cell reference, if any.
///
/// Returns the top-left cell of the unfrozen area (e.g., `"A2"` if row 1 is
/// frozen), or `None` if no panes are configured.
pub fn get_panes(ws: &WorksheetXml) -> Option<String> {
    ws.sheet_views
        .as_ref()
        .and_then(|sv| sv.sheet_views.first())
        .and_then(|view| view.pane.as_ref())
        .and_then(|pane| pane.top_left_cell.clone())
}

/// Rebuild content type overrides for worksheets so they match the current
/// worksheet indices (sheet1.xml, sheet2.xml, ...).
fn rebuild_content_type_overrides(content_types: &mut ContentTypes, sheet_count: usize) {
    content_types
        .overrides
        .retain(|o| o.content_type != mime_types::WORKSHEET);

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
    let sheet_rids: Vec<String> = workbook_xml
        .sheets
        .sheets
        .iter()
        .map(|s| s.r_id.clone())
        .collect();

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

#[cfg(test)]
mod tests {
    use super::*;
    use sheetkit_xml::content_types::ContentTypes;
    use sheetkit_xml::relationships;
    use sheetkit_xml::workbook::WorkbookXml;
    use sheetkit_xml::worksheet::WorksheetXml;

    // -- Sheet protection tests --

    #[test]
    fn test_protect_sheet_no_password() {
        let mut ws = WorksheetXml::default();
        let config = SheetProtectionConfig::default();
        protect_sheet(&mut ws, &config).unwrap();

        assert!(ws.sheet_protection.is_some());
        let prot = ws.sheet_protection.as_ref().unwrap();
        assert_eq!(prot.sheet, Some(true));
        assert_eq!(prot.objects, Some(true));
        assert_eq!(prot.scenarios, Some(true));
        assert!(prot.password.is_none());
    }

    #[test]
    fn test_protect_sheet_with_password() {
        let mut ws = WorksheetXml::default();
        let config = SheetProtectionConfig {
            password: Some("secret".to_string()),
            ..SheetProtectionConfig::default()
        };
        protect_sheet(&mut ws, &config).unwrap();

        let prot = ws.sheet_protection.as_ref().unwrap();
        assert!(prot.password.is_some());
        let pw = prot.password.as_ref().unwrap();
        // Should be a 4-char uppercase hex string
        assert_eq!(pw.len(), 4);
        assert!(pw.chars().all(|c| c.is_ascii_hexdigit()));
        // Should be deterministic
        let expected = format!("{:04X}", legacy_password_hash("secret"));
        assert_eq!(pw, &expected);
    }

    #[test]
    fn test_unprotect_sheet() {
        let mut ws = WorksheetXml::default();
        let config = SheetProtectionConfig {
            password: Some("test".to_string()),
            ..SheetProtectionConfig::default()
        };
        protect_sheet(&mut ws, &config).unwrap();
        assert!(ws.sheet_protection.is_some());

        unprotect_sheet(&mut ws).unwrap();
        assert!(ws.sheet_protection.is_none());
    }

    #[test]
    fn test_is_sheet_protected() {
        let mut ws = WorksheetXml::default();
        assert!(!is_sheet_protected(&ws));

        let config = SheetProtectionConfig::default();
        protect_sheet(&mut ws, &config).unwrap();
        assert!(is_sheet_protected(&ws));

        unprotect_sheet(&mut ws).unwrap();
        assert!(!is_sheet_protected(&ws));
    }

    #[test]
    fn test_protect_sheet_with_permissions() {
        let mut ws = WorksheetXml::default();
        let config = SheetProtectionConfig {
            password: None,
            format_cells: true,
            insert_rows: true,
            delete_columns: true,
            sort: true,
            ..SheetProtectionConfig::default()
        };
        protect_sheet(&mut ws, &config).unwrap();

        let prot = ws.sheet_protection.as_ref().unwrap();
        assert_eq!(prot.format_cells, Some(true));
        assert_eq!(prot.insert_rows, Some(true));
        assert_eq!(prot.delete_columns, Some(true));
        assert_eq!(prot.sort, Some(true));
        // Fields not set should be None (meaning forbidden)
        assert!(prot.format_columns.is_none());
        assert!(prot.format_rows.is_none());
        assert!(prot.insert_columns.is_none());
        assert!(prot.insert_hyperlinks.is_none());
        assert!(prot.delete_rows.is_none());
        assert!(prot.auto_filter.is_none());
        assert!(prot.pivot_tables.is_none());
        assert!(prot.select_locked_cells.is_none());
        assert!(prot.select_unlocked_cells.is_none());
    }

    // -- Tab color tests --

    #[test]
    fn test_set_tab_color() {
        let mut ws = WorksheetXml::default();
        set_tab_color(&mut ws, "FF0000").unwrap();

        assert!(ws.sheet_pr.is_some());
        let tab_color = ws.sheet_pr.as_ref().unwrap().tab_color.as_ref().unwrap();
        assert_eq!(tab_color.rgb, Some("FF0000".to_string()));
    }

    #[test]
    fn test_get_tab_color() {
        let mut ws = WorksheetXml::default();
        set_tab_color(&mut ws, "00FF00").unwrap();
        assert_eq!(get_tab_color(&ws), Some("00FF00".to_string()));
    }

    #[test]
    fn test_get_tab_color_none() {
        let ws = WorksheetXml::default();
        assert_eq!(get_tab_color(&ws), None);
    }

    // -- Default row height tests --

    #[test]
    fn test_set_default_row_height() {
        let mut ws = WorksheetXml::default();
        set_default_row_height(&mut ws, 20.0).unwrap();

        assert!(ws.sheet_format_pr.is_some());
        assert_eq!(
            ws.sheet_format_pr.as_ref().unwrap().default_row_height,
            20.0
        );
    }

    #[test]
    fn test_get_default_row_height() {
        let ws = WorksheetXml::default();
        assert_eq!(get_default_row_height(&ws), DEFAULT_ROW_HEIGHT);

        let mut ws2 = WorksheetXml::default();
        set_default_row_height(&mut ws2, 25.0).unwrap();
        assert_eq!(get_default_row_height(&ws2), 25.0);
    }

    #[test]
    fn test_set_default_row_height_exceeds_max() {
        let mut ws = WorksheetXml::default();
        let result = set_default_row_height(&mut ws, 500.0);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::RowHeightExceeded { .. }
        ));
    }

    // -- Default column width tests --

    #[test]
    fn test_set_default_col_width() {
        let mut ws = WorksheetXml::default();
        set_default_col_width(&mut ws, 12.0).unwrap();

        assert!(ws.sheet_format_pr.is_some());
        assert_eq!(
            ws.sheet_format_pr.as_ref().unwrap().default_col_width,
            Some(12.0)
        );
    }

    #[test]
    fn test_get_default_col_width() {
        let ws = WorksheetXml::default();
        assert_eq!(get_default_col_width(&ws), None);

        let mut ws2 = WorksheetXml::default();
        set_default_col_width(&mut ws2, 18.5).unwrap();
        assert_eq!(get_default_col_width(&ws2), Some(18.5));
    }

    #[test]
    fn test_set_default_col_width_exceeds_max() {
        let mut ws = WorksheetXml::default();
        let result = set_default_col_width(&mut ws, 300.0);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::ColumnWidthExceeded { .. }
        ));
    }

    // -- Existing tests below --

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

        let ws_rels: Vec<_> = wb_rels
            .relationships
            .iter()
            .filter(|r| r.rel_type == rel_types::WORKSHEET)
            .collect();
        assert_eq!(ws_rels.len(), 2);

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

        delete_sheet(&mut wb_xml, &mut wb_rels, &mut ct, &mut ws, "Sheet1").unwrap();

        assert_eq!(ws.len(), 1);
        assert_eq!(ws[0].0, "Sheet2");
        assert_eq!(wb_xml.sheets.sheets.len(), 1);
        assert_eq!(wb_xml.sheets.sheets[0].name, "Sheet2");

        let ws_rels: Vec<_> = wb_rels
            .relationships
            .iter()
            .filter(|r| r.rel_type == rel_types::WORKSHEET)
            .collect();
        assert_eq!(ws_rels.len(), 1);

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

        delete_sheet(&mut wb_xml, &mut wb_rels, &mut ct, &mut ws, "B").unwrap();

        assert_eq!(ws.len(), 3);
        let names: Vec<&str> = ws.iter().map(|(n, _)| n.as_str()).collect();
        assert_eq!(names, vec!["Sheet1", "A", "C"]);

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

    // -- Freeze pane tests --

    #[test]
    fn test_set_panes_freeze_row() {
        let mut ws = WorksheetXml::default();
        set_panes(&mut ws, "A2").unwrap();

        let pane = ws.sheet_views.as_ref().unwrap().sheet_views[0]
            .pane
            .as_ref()
            .unwrap();
        assert_eq!(pane.y_split, Some(1));
        assert!(pane.x_split.is_none());
        assert_eq!(pane.top_left_cell, Some("A2".to_string()));
        assert_eq!(pane.active_pane, Some("bottomLeft".to_string()));
        assert_eq!(pane.state, Some("frozen".to_string()));
    }

    #[test]
    fn test_set_panes_freeze_col() {
        let mut ws = WorksheetXml::default();
        set_panes(&mut ws, "B1").unwrap();

        let pane = ws.sheet_views.as_ref().unwrap().sheet_views[0]
            .pane
            .as_ref()
            .unwrap();
        assert_eq!(pane.x_split, Some(1));
        assert!(pane.y_split.is_none());
        assert_eq!(pane.top_left_cell, Some("B1".to_string()));
        assert_eq!(pane.active_pane, Some("topRight".to_string()));
        assert_eq!(pane.state, Some("frozen".to_string()));
    }

    #[test]
    fn test_set_panes_freeze_both() {
        let mut ws = WorksheetXml::default();
        set_panes(&mut ws, "B2").unwrap();

        let pane = ws.sheet_views.as_ref().unwrap().sheet_views[0]
            .pane
            .as_ref()
            .unwrap();
        assert_eq!(pane.x_split, Some(1));
        assert_eq!(pane.y_split, Some(1));
        assert_eq!(pane.top_left_cell, Some("B2".to_string()));
        assert_eq!(pane.active_pane, Some("bottomRight".to_string()));
        assert_eq!(pane.state, Some("frozen".to_string()));
    }

    #[test]
    fn test_set_panes_freeze_multiple_rows() {
        let mut ws = WorksheetXml::default();
        set_panes(&mut ws, "A4").unwrap();

        let pane = ws.sheet_views.as_ref().unwrap().sheet_views[0]
            .pane
            .as_ref()
            .unwrap();
        assert_eq!(pane.y_split, Some(3));
        assert!(pane.x_split.is_none());
        assert_eq!(pane.top_left_cell, Some("A4".to_string()));
        assert_eq!(pane.active_pane, Some("bottomLeft".to_string()));
    }

    #[test]
    fn test_set_panes_freeze_multiple_cols() {
        let mut ws = WorksheetXml::default();
        set_panes(&mut ws, "D1").unwrap();

        let pane = ws.sheet_views.as_ref().unwrap().sheet_views[0]
            .pane
            .as_ref()
            .unwrap();
        assert_eq!(pane.x_split, Some(3));
        assert!(pane.y_split.is_none());
        assert_eq!(pane.top_left_cell, Some("D1".to_string()));
        assert_eq!(pane.active_pane, Some("topRight".to_string()));
    }

    #[test]
    fn test_set_panes_a1_error() {
        let mut ws = WorksheetXml::default();
        let result = set_panes(&mut ws, "A1");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::InvalidCellReference(_)
        ));
    }

    #[test]
    fn test_set_panes_invalid_cell_error() {
        let mut ws = WorksheetXml::default();
        let result = set_panes(&mut ws, "ZZZZ1");
        assert!(result.is_err());
    }

    #[test]
    fn test_unset_panes() {
        let mut ws = WorksheetXml::default();
        set_panes(&mut ws, "B2").unwrap();
        assert!(get_panes(&ws).is_some());

        unset_panes(&mut ws);
        assert!(get_panes(&ws).is_none());
        // SheetViews should still exist but without pane.
        let view = &ws.sheet_views.as_ref().unwrap().sheet_views[0];
        assert!(view.pane.is_none());
        assert!(view.selection.is_empty());
    }

    #[test]
    fn test_get_panes_none_when_not_set() {
        let ws = WorksheetXml::default();
        assert!(get_panes(&ws).is_none());
    }

    #[test]
    fn test_get_panes_returns_value_after_set() {
        let mut ws = WorksheetXml::default();
        set_panes(&mut ws, "C5").unwrap();
        assert_eq!(get_panes(&ws), Some("C5".to_string()));
    }

    #[test]
    fn test_set_panes_selection_has_pane_attribute() {
        let mut ws = WorksheetXml::default();
        set_panes(&mut ws, "B2").unwrap();

        let selection = &ws.sheet_views.as_ref().unwrap().sheet_views[0].selection[0];
        assert_eq!(selection.pane, Some("bottomRight".to_string()));
        assert_eq!(selection.active_cell, Some("B2".to_string()));
        assert_eq!(selection.sqref, Some("B2".to_string()));
    }

    #[test]
    fn test_set_panes_overwrites_previous() {
        let mut ws = WorksheetXml::default();
        set_panes(&mut ws, "A2").unwrap();
        assert_eq!(get_panes(&ws), Some("A2".to_string()));

        set_panes(&mut ws, "C3").unwrap();
        assert_eq!(get_panes(&ws), Some("C3".to_string()));

        let pane = ws.sheet_views.as_ref().unwrap().sheet_views[0]
            .pane
            .as_ref()
            .unwrap();
        assert_eq!(pane.x_split, Some(2));
        assert_eq!(pane.y_split, Some(2));
        assert_eq!(pane.active_pane, Some("bottomRight".to_string()));
    }

    #[test]
    fn test_unset_panes_noop_when_no_views() {
        let mut ws = WorksheetXml::default();
        // Should not panic when there are no sheet views.
        unset_panes(&mut ws);
        assert!(get_panes(&ws).is_none());
    }
}
