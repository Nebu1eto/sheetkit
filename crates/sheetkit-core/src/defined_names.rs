//! Defined names (named ranges) management.
//!
//! Provides functions to add, retrieve, update, delete, and list defined names
//! in an Excel workbook. Defined names can be workbook-scoped (visible from all
//! sheets) or sheet-scoped (visible only in the specified sheet).

use sheetkit_xml::workbook::{DefinedName, DefinedNames, WorkbookXml};

use crate::error::{Error, Result};

/// Characters that are not allowed in defined names.
const DEFINED_NAME_INVALID_CHARS: &[char] = &['/', '?', '*', '[', ']'];

/// Scope of a defined name.
#[derive(Debug, Clone, PartialEq)]
pub enum DefinedNameScope {
    /// Workbook-level scope (visible from all sheets).
    Workbook,
    /// Sheet-level scope (only visible in the specified sheet).
    Sheet(u32),
}

/// Information about a defined name.
#[derive(Debug, Clone)]
pub struct DefinedNameInfo {
    pub name: String,
    /// The reference or formula, e.g. "Sheet1!$A$1:$D$10".
    pub value: String,
    pub scope: DefinedNameScope,
    pub comment: Option<String>,
}

/// Validate a defined name.
///
/// A valid defined name must:
/// - Be non-empty
/// - Be at most 255 Unicode scalar characters
/// - Start with a Unicode letter, underscore, or backslash
/// - Contain only Unicode letters or digits, underscores, periods, or backslashes
/// - Not be an A1 or R1C1 reference, or the reserved names R and C
fn validate_defined_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(Error::InvalidDefinedName(
            "defined name cannot be empty".into(),
        ));
    }
    if name.chars().count() > 255 {
        return Err(Error::InvalidDefinedName(
            "defined name cannot exceed 255 characters".into(),
        ));
    }
    if name != name.trim() {
        return Err(Error::InvalidDefinedName(
            "defined name cannot start or end with whitespace".into(),
        ));
    }
    for ch in DEFINED_NAME_INVALID_CHARS {
        if name.contains(*ch) {
            return Err(Error::InvalidDefinedName(format!(
                "defined name '{}' contains invalid character '{}'",
                name, ch
            )));
        }
    }
    let mut chars = name.chars();
    let first = chars.next().expect("name was checked as non-empty");
    if !(first.is_alphabetic() || matches!(first, '_' | '\\')) {
        return Err(Error::InvalidDefinedName(
            "defined name must start with a letter, underscore, or backslash".into(),
        ));
    }
    if !chars.all(|ch| ch.is_alphanumeric() || matches!(ch, '_' | '.' | '\\')) {
        return Err(Error::InvalidDefinedName(
            "defined name contains characters outside Excel's name grammar".into(),
        ));
    }
    if name.eq_ignore_ascii_case("R") || name.eq_ignore_ascii_case("C") {
        return Err(Error::InvalidDefinedName(
            "defined name cannot be the reserved name R or C".into(),
        ));
    }
    if is_a1_reference(name) || is_r1c1_reference(name) {
        return Err(Error::InvalidDefinedName(
            "defined name cannot be an Excel cell reference".into(),
        ));
    }
    Ok(())
}

fn is_a1_reference(name: &str) -> bool {
    name.is_ascii() && crate::utils::cell_ref::cell_name_to_coordinates(name).is_ok()
}

fn is_r1c1_reference(name: &str) -> bool {
    if !name.is_ascii() {
        return false;
    }
    let upper = name.to_ascii_uppercase();
    let Some(rest) = upper.strip_prefix('R') else {
        return false;
    };
    let Some((row, col)) = rest.split_once('C') else {
        return false;
    };
    !row.is_empty()
        && !col.is_empty()
        && r1c1_component_is_valid(row)
        && r1c1_component_is_valid(col)
}

fn defined_names_equal(left: &str, right: &str) -> bool {
    left.to_lowercase() == right.to_lowercase()
}

fn r1c1_component_is_valid(component: &str) -> bool {
    component.chars().all(|ch| ch.is_ascii_digit())
        || (component.starts_with('[')
            && component.ends_with(']')
            && component[1..component.len() - 1]
                .strip_prefix('-')
                .unwrap_or(&component[1..component.len() - 1])
                .chars()
                .all(|ch| ch.is_ascii_digit()))
}

/// Convert a `DefinedNameScope` to the corresponding `local_sheet_id` value.
fn scope_to_local_sheet_id(scope: &DefinedNameScope) -> Option<u32> {
    match scope {
        DefinedNameScope::Workbook => None,
        DefinedNameScope::Sheet(id) => Some(*id),
    }
}

/// Convert an optional `local_sheet_id` to a `DefinedNameScope`.
fn local_sheet_id_to_scope(local_sheet_id: Option<u32>) -> DefinedNameScope {
    match local_sheet_id {
        None => DefinedNameScope::Workbook,
        Some(id) => DefinedNameScope::Sheet(id),
    }
}

/// Add or update a defined name in the workbook.
///
/// If a defined name with the same name and scope already exists, its value
/// and comment are updated. Otherwise, a new entry is created.
pub fn set_defined_name(
    wb: &mut WorkbookXml,
    name: &str,
    value: &str,
    scope: DefinedNameScope,
    comment: Option<&str>,
) -> Result<()> {
    validate_defined_name(name)?;

    let local_sheet_id = scope_to_local_sheet_id(&scope);

    let defined_names = wb.defined_names.get_or_insert_with(|| DefinedNames {
        defined_names: Vec::new(),
    });

    // Check if a name with the same name+scope already exists.
    if let Some(existing) = defined_names
        .defined_names
        .iter_mut()
        .find(|dn| defined_names_equal(&dn.name, name) && dn.local_sheet_id == local_sheet_id)
    {
        existing.value = value.to_string();
        existing.comment = comment.map(|c| c.to_string());
        return Ok(());
    }

    defined_names.defined_names.push(DefinedName {
        name: name.to_string(),
        local_sheet_id,
        comment: comment.map(|c| c.to_string()),
        hidden: None,
        value: value.to_string(),
    });

    Ok(())
}

/// Get a defined name by name and scope.
///
/// Returns `None` if no matching defined name is found.
pub fn get_defined_name(
    wb: &WorkbookXml,
    name: &str,
    scope: DefinedNameScope,
) -> Option<DefinedNameInfo> {
    let defined_names = wb.defined_names.as_ref()?;
    let local_sheet_id = scope_to_local_sheet_id(&scope);

    defined_names
        .defined_names
        .iter()
        .find(|dn| defined_names_equal(&dn.name, name) && dn.local_sheet_id == local_sheet_id)
        .map(|dn| DefinedNameInfo {
            name: dn.name.clone(),
            value: dn.value.clone(),
            scope: local_sheet_id_to_scope(dn.local_sheet_id),
            comment: dn.comment.clone(),
        })
}

/// Delete a defined name by name and scope.
///
/// Returns an error if the defined name does not exist.
pub fn delete_defined_name(
    wb: &mut WorkbookXml,
    name: &str,
    scope: DefinedNameScope,
) -> Result<()> {
    let local_sheet_id = scope_to_local_sheet_id(&scope);

    let defined_names = wb
        .defined_names
        .as_mut()
        .ok_or_else(|| Error::DefinedNameNotFound {
            name: name.to_string(),
        })?;

    let idx = defined_names
        .defined_names
        .iter()
        .position(|dn| defined_names_equal(&dn.name, name) && dn.local_sheet_id == local_sheet_id)
        .ok_or_else(|| Error::DefinedNameNotFound {
            name: name.to_string(),
        })?;

    defined_names.defined_names.remove(idx);

    // Clean up: remove the container if empty.
    if defined_names.defined_names.is_empty() {
        wb.defined_names = None;
    }

    Ok(())
}

/// List all defined names in the workbook.
pub fn get_all_defined_names(wb: &WorkbookXml) -> Vec<DefinedNameInfo> {
    let Some(defined_names) = wb.defined_names.as_ref() else {
        return Vec::new();
    };

    defined_names
        .defined_names
        .iter()
        .map(|dn| DefinedNameInfo {
            name: dn.name.clone(),
            value: dn.value.clone(),
            scope: local_sheet_id_to_scope(dn.local_sheet_id),
            comment: dn.comment.clone(),
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Create a minimal WorkbookXml for testing.
    fn test_workbook() -> WorkbookXml {
        WorkbookXml::default()
    }

    #[test]
    fn test_set_defined_name_workbook_scope() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "SalesData",
            "Sheet1!$A$1:$D$10",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();

        let dn = wb.defined_names.as_ref().unwrap();
        assert_eq!(dn.defined_names.len(), 1);
        assert_eq!(dn.defined_names[0].name, "SalesData");
        assert_eq!(dn.defined_names[0].value, "Sheet1!$A$1:$D$10");
        assert!(dn.defined_names[0].local_sheet_id.is_none());
    }

    #[test]
    fn test_set_defined_name_sheet_scope() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "LocalRange",
            "Sheet1!$B$2:$C$5",
            DefinedNameScope::Sheet(0),
            None,
        )
        .unwrap();

        let dn = wb.defined_names.as_ref().unwrap();
        assert_eq!(dn.defined_names.len(), 1);
        assert_eq!(dn.defined_names[0].name, "LocalRange");
        assert_eq!(dn.defined_names[0].local_sheet_id, Some(0));
    }

    #[test]
    fn test_get_defined_name() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "Revenue",
            "Sheet1!$E$1:$E$100",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();

        let info = get_defined_name(&wb, "Revenue", DefinedNameScope::Workbook).unwrap();
        assert_eq!(info.name, "Revenue");
        assert_eq!(info.value, "Sheet1!$E$1:$E$100");
        assert_eq!(info.scope, DefinedNameScope::Workbook);
        assert!(info.comment.is_none());
    }

    #[test]
    fn test_get_defined_name_not_found() {
        let wb = test_workbook();
        let result = get_defined_name(&wb, "NonExistent", DefinedNameScope::Workbook);
        assert!(result.is_none());
    }

    #[test]
    fn test_update_defined_name() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "DataRange",
            "Sheet1!$A$1:$A$10",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();

        // Update the same name with a new value.
        set_defined_name(
            &mut wb,
            "DataRange",
            "Sheet1!$A$1:$A$50",
            DefinedNameScope::Workbook,
            Some("Updated range"),
        )
        .unwrap();

        let dn = wb.defined_names.as_ref().unwrap();
        assert_eq!(dn.defined_names.len(), 1, "should not duplicate the entry");
        assert_eq!(dn.defined_names[0].value, "Sheet1!$A$1:$A$50");
        assert_eq!(
            dn.defined_names[0].comment,
            Some("Updated range".to_string())
        );
    }

    #[test]
    fn test_delete_defined_name() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "ToDelete",
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();
        assert!(wb.defined_names.is_some());

        delete_defined_name(&mut wb, "ToDelete", DefinedNameScope::Workbook).unwrap();
        // Container should be cleaned up since it is now empty.
        assert!(wb.defined_names.is_none());
    }

    #[test]
    fn test_delete_defined_name_not_found() {
        let mut wb = test_workbook();
        let result = delete_defined_name(&mut wb, "Ghost", DefinedNameScope::Workbook);
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(
            err.to_string().contains("Ghost"),
            "error message should contain the name"
        );
    }

    #[test]
    fn test_get_all_defined_names() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "Alpha",
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();
        set_defined_name(
            &mut wb,
            "Beta",
            "Sheet1!$B$1",
            DefinedNameScope::Sheet(0),
            None,
        )
        .unwrap();
        set_defined_name(
            &mut wb,
            "Gamma",
            "Sheet1!$C$1",
            DefinedNameScope::Sheet(1),
            None,
        )
        .unwrap();

        let all = get_all_defined_names(&wb);
        assert_eq!(all.len(), 3);
        assert_eq!(all[0].name, "Alpha");
        assert_eq!(all[1].name, "Beta");
        assert_eq!(all[2].name, "Gamma");
    }

    #[test]
    fn test_get_all_defined_names_empty() {
        let wb = test_workbook();
        let all = get_all_defined_names(&wb);
        assert!(all.is_empty());
    }

    #[test]
    fn test_same_name_different_scopes() {
        let mut wb = test_workbook();

        // Set workbook-scoped name.
        set_defined_name(
            &mut wb,
            "Total",
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();

        // Set sheet-scoped name with the same name.
        set_defined_name(
            &mut wb,
            "Total",
            "Sheet1!$B$1",
            DefinedNameScope::Sheet(0),
            None,
        )
        .unwrap();

        let dn = wb.defined_names.as_ref().unwrap();
        assert_eq!(dn.defined_names.len(), 2, "both scopes should coexist");

        let wb_info = get_defined_name(&wb, "Total", DefinedNameScope::Workbook).unwrap();
        assert_eq!(wb_info.value, "Sheet1!$A$1");

        let sheet_info = get_defined_name(&wb, "Total", DefinedNameScope::Sheet(0)).unwrap();
        assert_eq!(sheet_info.value, "Sheet1!$B$1");
    }

    #[test]
    fn test_defined_name_with_comment() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "Annotated",
            "Sheet1!$A$1:$Z$100",
            DefinedNameScope::Workbook,
            Some("Main data area"),
        )
        .unwrap();

        let info = get_defined_name(&wb, "Annotated", DefinedNameScope::Workbook).unwrap();
        assert_eq!(info.comment, Some("Main data area".to_string()));
    }

    #[test]
    fn test_invalid_defined_name_empty() {
        let mut wb = test_workbook();
        let result = set_defined_name(&mut wb, "", "Sheet1!$A$1", DefinedNameScope::Workbook, None);
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("empty"));
    }

    #[test]
    fn test_invalid_defined_name_leading_whitespace() {
        let mut wb = test_workbook();
        let result = set_defined_name(
            &mut wb,
            " Leading",
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        );
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("whitespace"));
    }

    #[test]
    fn test_invalid_defined_name_trailing_whitespace() {
        let mut wb = test_workbook();
        let result = set_defined_name(
            &mut wb,
            "Trailing ",
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        );
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("whitespace"));
    }

    #[test]
    fn test_invalid_defined_name_special_chars() {
        let mut wb = test_workbook();
        for ch in DEFINED_NAME_INVALID_CHARS {
            let name = format!("Bad{}Name", ch);
            let result = set_defined_name(
                &mut wb,
                &name,
                "Sheet1!$A$1",
                DefinedNameScope::Workbook,
                None,
            );
            assert!(result.is_err(), "should reject '{}' in name", ch);
            assert!(
                result
                    .unwrap_err()
                    .to_string()
                    .contains("invalid character"),
                "error for '{}' should mention invalid character",
                ch
            );
        }
    }

    #[test]
    fn test_defined_name_grammar_boundaries_and_reference_forms() {
        let mut wb = test_workbook();
        for name in [
            "1Name",
            "Name With Space",
            "A1",
            "XFD1048576",
            "R1C1",
            "R[-1]C[2]",
            "R",
            "C",
        ] {
            assert!(
                set_defined_name(
                    &mut wb,
                    name,
                    "Sheet1!$A$1",
                    DefinedNameScope::Workbook,
                    None
                )
                .is_err(),
                "{name} should be rejected"
            );
        }
        assert!(set_defined_name(
            &mut wb,
            &"N".repeat(255),
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        )
        .is_ok());
        assert!(set_defined_name(
            &mut wb,
            &"N".repeat(256),
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        )
        .is_err());
        assert!(set_defined_name(
            &mut wb,
            &"가".repeat(255),
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        )
        .is_ok());
        assert!(set_defined_name(
            &mut wb,
            &"가".repeat(256),
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        )
        .is_err());

        let korean_name = "매출합계";
        assert!(set_defined_name(
            &mut wb,
            korean_name,
            "Sheet1!$C$1",
            DefinedNameScope::Sheet(0),
            None,
        )
        .is_ok());
        assert_eq!(
            get_defined_name(&wb, korean_name, DefinedNameScope::Sheet(0))
                .unwrap()
                .value,
            "Sheet1!$C$1"
        );
    }

    #[test]
    fn test_defined_names_are_case_insensitive_within_scope() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "Revenue",
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();
        set_defined_name(
            &mut wb,
            "revenue",
            "Sheet1!$B$1",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();

        assert_eq!(wb.defined_names.as_ref().unwrap().defined_names.len(), 1);
        assert_eq!(
            get_defined_name(&wb, "REVENUE", DefinedNameScope::Workbook)
                .unwrap()
                .value,
            "Sheet1!$B$1"
        );
        delete_defined_name(&mut wb, "ReVeNuE", DefinedNameScope::Workbook).unwrap();
        assert!(wb.defined_names.is_none());
    }

    #[test]
    fn test_defined_name_unicode_case_identity() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "Résumé",
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();
        set_defined_name(
            &mut wb,
            "RÉSUMÉ",
            "Sheet1!$B$1",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();

        assert_eq!(wb.defined_names.as_ref().unwrap().defined_names.len(), 1);
        assert_eq!(
            get_defined_name(&wb, "résumé", DefinedNameScope::Workbook)
                .unwrap()
                .value,
            "Sheet1!$B$1"
        );
    }

    #[test]
    fn test_delete_one_keeps_others() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "Keep",
            "Sheet1!$A$1",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();
        set_defined_name(
            &mut wb,
            "Remove",
            "Sheet1!$B$1",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();

        delete_defined_name(&mut wb, "Remove", DefinedNameScope::Workbook).unwrap();

        let dn = wb.defined_names.as_ref().unwrap();
        assert_eq!(dn.defined_names.len(), 1);
        assert_eq!(dn.defined_names[0].name, "Keep");
    }

    #[test]
    fn test_delete_wrong_scope_not_found() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "ScopedName",
            "Sheet1!$A$1",
            DefinedNameScope::Sheet(0),
            None,
        )
        .unwrap();

        // Deleting with workbook scope should fail because it only exists at sheet scope.
        let result = delete_defined_name(&mut wb, "ScopedName", DefinedNameScope::Workbook);
        assert!(result.is_err());
    }

    #[test]
    fn test_xml_roundtrip_with_defined_names() {
        let mut wb = test_workbook();
        set_defined_name(
            &mut wb,
            "RangeA",
            "Sheet1!$A$1:$A$10",
            DefinedNameScope::Workbook,
            Some("First range"),
        )
        .unwrap();
        set_defined_name(
            &mut wb,
            "RangeB",
            "Sheet1!$B$1:$B$5",
            DefinedNameScope::Sheet(0),
            None,
        )
        .unwrap();
        set_defined_name(
            &mut wb,
            "매출합계",
            "Sheet1!$C$1:$C$5",
            DefinedNameScope::Workbook,
            None,
        )
        .unwrap();

        let xml = quick_xml::se::to_string(&wb).unwrap();
        let parsed: WorkbookXml = quick_xml::de::from_str(&xml).unwrap();

        let all = get_all_defined_names(&parsed);
        assert_eq!(all.len(), 3);
        assert_eq!(all[0].name, "RangeA");
        assert_eq!(all[0].value, "Sheet1!$A$1:$A$10");
        assert_eq!(all[0].scope, DefinedNameScope::Workbook);
        assert_eq!(all[0].comment, Some("First range".to_string()));
        assert_eq!(all[1].name, "RangeB");
        assert_eq!(all[1].value, "Sheet1!$B$1:$B$5");
        assert_eq!(all[1].scope, DefinedNameScope::Sheet(0));
        assert!(all[1].comment.is_none());
        assert_eq!(all[2].name, "매출합계");
        assert_eq!(all[2].value, "Sheet1!$C$1:$C$5");
    }
}
