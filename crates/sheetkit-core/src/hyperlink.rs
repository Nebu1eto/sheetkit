//! Hyperlink management for worksheet cells.
//!
//! Provides types and functions for setting, getting, and deleting hyperlinks
//! on individual cells. Supports external URLs, internal sheet references,
//! and email (mailto) links.

use sheetkit_xml::relationships::{rel_types, Relationship, Relationships};
use sheetkit_xml::worksheet::{Hyperlink, Hyperlinks, WorksheetXml};

use crate::error::Result;

/// Type of hyperlink target.
#[derive(Debug, Clone, PartialEq)]
pub enum HyperlinkType {
    /// External URL (e.g., "https://example.com" or "file:///path").
    External(String),
    /// Internal sheet reference (e.g., "Sheet2!A1").
    Internal(String),
    /// Email link (e.g., "mailto:user@example.com").
    Email(String),
}

/// Hyperlink information returned by get operations.
#[derive(Debug, Clone, PartialEq)]
pub struct HyperlinkInfo {
    /// The hyperlink target.
    pub link_type: HyperlinkType,
    /// Optional display text.
    pub display: Option<String>,
    /// Optional tooltip text.
    pub tooltip: Option<String>,
}

/// Set a hyperlink on a cell.
///
/// For external URLs and email links, a relationship entry is created in the
/// worksheet `.rels` file with `TargetMode="External"`. For internal sheet
/// references, only the `location` attribute is set on the hyperlink element
/// (no relationship is needed).
///
/// If a hyperlink already exists on the cell, it is replaced.
pub fn set_cell_hyperlink(
    ws: &mut WorksheetXml,
    rels: &mut Relationships,
    cell: &str,
    link: &HyperlinkType,
    display: Option<&str>,
    tooltip: Option<&str>,
) -> Result<()> {
    // Remove any existing hyperlink on this cell first.
    delete_cell_hyperlink(ws, rels, cell)?;

    let hyperlinks = ws
        .hyperlinks
        .get_or_insert_with(|| Hyperlinks { hyperlinks: vec![] });

    match link {
        HyperlinkType::External(url) | HyperlinkType::Email(url) => {
            let rid = next_rel_id(rels);
            rels.relationships.push(Relationship {
                id: rid.clone(),
                rel_type: rel_types::HYPERLINK.to_string(),
                target: url.clone(),
                target_mode: Some("External".to_string()),
            });
            hyperlinks.hyperlinks.push(Hyperlink {
                reference: cell.to_string(),
                r_id: Some(rid),
                location: None,
                display: display.map(|s| s.to_string()),
                tooltip: tooltip.map(|s| s.to_string()),
            });
        }
        HyperlinkType::Internal(location) => {
            hyperlinks.hyperlinks.push(Hyperlink {
                reference: cell.to_string(),
                r_id: None,
                location: Some(location.clone()),
                display: display.map(|s| s.to_string()),
                tooltip: tooltip.map(|s| s.to_string()),
            });
        }
    }

    Ok(())
}

/// Get hyperlink information for a cell.
///
/// Returns `None` if the cell has no hyperlink.
pub fn get_cell_hyperlink(
    ws: &WorksheetXml,
    rels: &Relationships,
    cell: &str,
) -> Result<Option<HyperlinkInfo>> {
    let hyperlinks = match &ws.hyperlinks {
        Some(h) => h,
        None => return Ok(None),
    };

    let hl = match hyperlinks.hyperlinks.iter().find(|h| h.reference == cell) {
        Some(h) => h,
        None => return Ok(None),
    };

    let link_type = if let Some(ref rid) = hl.r_id {
        // Look up the relationship target.
        let rel = rels.relationships.iter().find(|r| r.id == *rid);
        match rel {
            Some(r) => {
                let target = &r.target;
                if target.starts_with("mailto:") {
                    HyperlinkType::Email(target.clone())
                } else {
                    HyperlinkType::External(target.clone())
                }
            }
            None => {
                // Relationship not found; treat as external with empty target.
                HyperlinkType::External(String::new())
            }
        }
    } else if let Some(ref location) = hl.location {
        HyperlinkType::Internal(location.clone())
    } else {
        // No r:id and no location; should not happen in valid files.
        return Ok(None);
    };

    Ok(Some(HyperlinkInfo {
        link_type,
        display: hl.display.clone(),
        tooltip: hl.tooltip.clone(),
    }))
}

/// Delete a hyperlink from a cell.
///
/// Removes the hyperlink element from the worksheet XML and, if the hyperlink
/// used a relationship, removes the corresponding relationship entry.
pub fn delete_cell_hyperlink(
    ws: &mut WorksheetXml,
    rels: &mut Relationships,
    cell: &str,
) -> Result<()> {
    let hyperlinks = match &mut ws.hyperlinks {
        Some(h) => h,
        None => return Ok(()),
    };

    // Find and remove the hyperlink for this cell.
    let removed: Vec<Hyperlink> = hyperlinks
        .hyperlinks
        .extract_if(.., |h| h.reference == cell)
        .collect();

    // Remove associated relationships.
    for hl in &removed {
        if let Some(ref rid) = hl.r_id {
            rels.relationships.retain(|r| r.id != *rid);
        }
    }

    // If no hyperlinks remain, remove the container.
    if hyperlinks.hyperlinks.is_empty() {
        ws.hyperlinks = None;
    }

    Ok(())
}

/// Generate the next relationship ID for a rels collection.
fn next_rel_id(rels: &Relationships) -> String {
    let max = rels
        .relationships
        .iter()
        .filter_map(|r| r.id.strip_prefix("rId").and_then(|n| n.parse::<u32>().ok()))
        .max()
        .unwrap_or(0);
    format!("rId{}", max + 1)
}

#[cfg(test)]
#[allow(clippy::field_reassign_with_default)]
mod tests {
    use super::*;
    use sheetkit_xml::namespaces;

    fn empty_rels() -> Relationships {
        Relationships {
            xmlns: namespaces::PACKAGE_RELATIONSHIPS.to_string(),
            relationships: vec![],
        }
    }

    #[test]
    fn test_set_external_hyperlink() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::External("https://example.com".to_string()),
            None,
            None,
        )
        .unwrap();

        // Hyperlink element should exist.
        let hls = ws.hyperlinks.as_ref().unwrap();
        assert_eq!(hls.hyperlinks.len(), 1);
        assert_eq!(hls.hyperlinks[0].reference, "A1");
        assert!(hls.hyperlinks[0].r_id.is_some());
        assert!(hls.hyperlinks[0].location.is_none());

        // Relationship should exist.
        assert_eq!(rels.relationships.len(), 1);
        assert_eq!(rels.relationships[0].target, "https://example.com");
        assert_eq!(
            rels.relationships[0].target_mode,
            Some("External".to_string())
        );
        assert_eq!(rels.relationships[0].rel_type, rel_types::HYPERLINK);
    }

    #[test]
    fn test_set_internal_hyperlink() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "B2",
            &HyperlinkType::Internal("Sheet2!A1".to_string()),
            None,
            None,
        )
        .unwrap();

        let hls = ws.hyperlinks.as_ref().unwrap();
        assert_eq!(hls.hyperlinks.len(), 1);
        assert_eq!(hls.hyperlinks[0].reference, "B2");
        assert!(hls.hyperlinks[0].r_id.is_none());
        assert_eq!(hls.hyperlinks[0].location, Some("Sheet2!A1".to_string()));

        // No relationship should be created for internal links.
        assert!(rels.relationships.is_empty());
    }

    #[test]
    fn test_set_email_hyperlink() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "C3",
            &HyperlinkType::Email("mailto:user@example.com".to_string()),
            None,
            None,
        )
        .unwrap();

        let hls = ws.hyperlinks.as_ref().unwrap();
        assert_eq!(hls.hyperlinks.len(), 1);
        assert!(hls.hyperlinks[0].r_id.is_some());

        assert_eq!(rels.relationships.len(), 1);
        assert_eq!(rels.relationships[0].target, "mailto:user@example.com");
        assert_eq!(
            rels.relationships[0].target_mode,
            Some("External".to_string())
        );
    }

    #[test]
    fn test_get_hyperlink_external() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::External("https://rust-lang.org".to_string()),
            Some("Rust"),
            Some("Visit Rust"),
        )
        .unwrap();

        let info = get_cell_hyperlink(&ws, &rels, "A1").unwrap().unwrap();
        assert_eq!(
            info.link_type,
            HyperlinkType::External("https://rust-lang.org".to_string())
        );
        assert_eq!(info.display, Some("Rust".to_string()));
        assert_eq!(info.tooltip, Some("Visit Rust".to_string()));
    }

    #[test]
    fn test_get_hyperlink_internal() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "D4",
            &HyperlinkType::Internal("Summary!B10".to_string()),
            Some("Go to Summary"),
            None,
        )
        .unwrap();

        let info = get_cell_hyperlink(&ws, &rels, "D4").unwrap().unwrap();
        assert_eq!(
            info.link_type,
            HyperlinkType::Internal("Summary!B10".to_string())
        );
        assert_eq!(info.display, Some("Go to Summary".to_string()));
        assert!(info.tooltip.is_none());
    }

    #[test]
    fn test_get_hyperlink_email() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::Email("mailto:test@test.com".to_string()),
            None,
            None,
        )
        .unwrap();

        let info = get_cell_hyperlink(&ws, &rels, "A1").unwrap().unwrap();
        assert_eq!(
            info.link_type,
            HyperlinkType::Email("mailto:test@test.com".to_string())
        );
    }

    #[test]
    fn test_delete_hyperlink() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::External("https://example.com".to_string()),
            None,
            None,
        )
        .unwrap();

        assert!(ws.hyperlinks.is_some());
        assert_eq!(rels.relationships.len(), 1);

        delete_cell_hyperlink(&mut ws, &mut rels, "A1").unwrap();

        // Hyperlinks container should be removed when empty.
        assert!(ws.hyperlinks.is_none());
        // Relationship should be cleaned up.
        assert!(rels.relationships.is_empty());
    }

    #[test]
    fn test_delete_internal_hyperlink() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::Internal("Sheet2!A1".to_string()),
            None,
            None,
        )
        .unwrap();

        delete_cell_hyperlink(&mut ws, &mut rels, "A1").unwrap();

        assert!(ws.hyperlinks.is_none());
        assert!(rels.relationships.is_empty());
    }

    #[test]
    fn test_hyperlink_with_display_and_tooltip() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::External("https://example.com".to_string()),
            Some("Click here"),
            Some("Opens example.com"),
        )
        .unwrap();

        let hls = ws.hyperlinks.as_ref().unwrap();
        assert_eq!(hls.hyperlinks[0].display, Some("Click here".to_string()));
        assert_eq!(
            hls.hyperlinks[0].tooltip,
            Some("Opens example.com".to_string())
        );

        let info = get_cell_hyperlink(&ws, &rels, "A1").unwrap().unwrap();
        assert_eq!(info.display, Some("Click here".to_string()));
        assert_eq!(info.tooltip, Some("Opens example.com".to_string()));
    }

    #[test]
    fn test_overwrite_hyperlink() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        // Set first hyperlink.
        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::External("https://old.com".to_string()),
            None,
            None,
        )
        .unwrap();

        // Overwrite with a new hyperlink.
        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::External("https://new.com".to_string()),
            Some("New Link"),
            None,
        )
        .unwrap();

        // Should have only one hyperlink.
        let hls = ws.hyperlinks.as_ref().unwrap();
        assert_eq!(hls.hyperlinks.len(), 1);

        // Should have only one relationship (old one cleaned up).
        assert_eq!(rels.relationships.len(), 1);
        assert_eq!(rels.relationships[0].target, "https://new.com");

        let info = get_cell_hyperlink(&ws, &rels, "A1").unwrap().unwrap();
        assert_eq!(
            info.link_type,
            HyperlinkType::External("https://new.com".to_string())
        );
        assert_eq!(info.display, Some("New Link".to_string()));
    }

    #[test]
    fn test_overwrite_external_with_internal() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::External("https://example.com".to_string()),
            None,
            None,
        )
        .unwrap();

        assert_eq!(rels.relationships.len(), 1);

        // Overwrite with internal link.
        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::Internal("Sheet2!A1".to_string()),
            None,
            None,
        )
        .unwrap();

        // Old relationship should be cleaned up.
        assert!(rels.relationships.is_empty());

        let hls = ws.hyperlinks.as_ref().unwrap();
        assert_eq!(hls.hyperlinks.len(), 1);
        assert!(hls.hyperlinks[0].r_id.is_none());
        assert_eq!(hls.hyperlinks[0].location, Some("Sheet2!A1".to_string()));
    }

    #[test]
    fn test_get_nonexistent_hyperlink() {
        let ws = WorksheetXml::default();
        let rels = empty_rels();

        let result = get_cell_hyperlink(&ws, &rels, "Z99").unwrap();
        assert!(result.is_none());
    }

    #[test]
    fn test_get_nonexistent_hyperlink_with_empty_container() {
        let mut ws = WorksheetXml::default();
        ws.hyperlinks = Some(Hyperlinks { hyperlinks: vec![] });
        let rels = empty_rels();

        let result = get_cell_hyperlink(&ws, &rels, "A1").unwrap();
        assert!(result.is_none());
    }

    #[test]
    fn test_delete_nonexistent_hyperlink() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        // Deleting a hyperlink that does not exist should succeed silently.
        delete_cell_hyperlink(&mut ws, &mut rels, "A1").unwrap();
        assert!(ws.hyperlinks.is_none());
    }

    #[test]
    fn test_multiple_hyperlinks() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::External("https://example.com".to_string()),
            Some("Example"),
            None,
        )
        .unwrap();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "B1",
            &HyperlinkType::Internal("Sheet2!A1".to_string()),
            Some("Sheet 2"),
            None,
        )
        .unwrap();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "C1",
            &HyperlinkType::Email("mailto:info@example.com".to_string()),
            Some("Email Us"),
            Some("Send email"),
        )
        .unwrap();

        let hls = ws.hyperlinks.as_ref().unwrap();
        assert_eq!(hls.hyperlinks.len(), 3);

        // External and email should have relationships; internal should not.
        assert_eq!(rels.relationships.len(), 2);

        // Verify each hyperlink.
        let a1 = get_cell_hyperlink(&ws, &rels, "A1").unwrap().unwrap();
        assert_eq!(
            a1.link_type,
            HyperlinkType::External("https://example.com".to_string())
        );
        assert_eq!(a1.display, Some("Example".to_string()));

        let b1 = get_cell_hyperlink(&ws, &rels, "B1").unwrap().unwrap();
        assert_eq!(
            b1.link_type,
            HyperlinkType::Internal("Sheet2!A1".to_string())
        );
        assert_eq!(b1.display, Some("Sheet 2".to_string()));

        let c1 = get_cell_hyperlink(&ws, &rels, "C1").unwrap().unwrap();
        assert_eq!(
            c1.link_type,
            HyperlinkType::Email("mailto:info@example.com".to_string())
        );
        assert_eq!(c1.display, Some("Email Us".to_string()));
        assert_eq!(c1.tooltip, Some("Send email".to_string()));
    }

    #[test]
    fn test_delete_one_of_multiple() {
        let mut ws = WorksheetXml::default();
        let mut rels = empty_rels();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "A1",
            &HyperlinkType::External("https://a.com".to_string()),
            None,
            None,
        )
        .unwrap();

        set_cell_hyperlink(
            &mut ws,
            &mut rels,
            "B1",
            &HyperlinkType::External("https://b.com".to_string()),
            None,
            None,
        )
        .unwrap();

        assert_eq!(ws.hyperlinks.as_ref().unwrap().hyperlinks.len(), 2);
        assert_eq!(rels.relationships.len(), 2);

        // Delete only A1.
        delete_cell_hyperlink(&mut ws, &mut rels, "A1").unwrap();

        let hls = ws.hyperlinks.as_ref().unwrap();
        assert_eq!(hls.hyperlinks.len(), 1);
        assert_eq!(hls.hyperlinks[0].reference, "B1");

        // Only B1's relationship should remain.
        assert_eq!(rels.relationships.len(), 1);
        assert_eq!(rels.relationships[0].target, "https://b.com");
    }

    #[test]
    fn test_next_rel_id_empty() {
        let rels = empty_rels();
        assert_eq!(next_rel_id(&rels), "rId1");
    }

    #[test]
    fn test_next_rel_id_with_existing() {
        let rels = Relationships {
            xmlns: namespaces::PACKAGE_RELATIONSHIPS.to_string(),
            relationships: vec![
                Relationship {
                    id: "rId1".to_string(),
                    rel_type: rel_types::HYPERLINK.to_string(),
                    target: "https://a.com".to_string(),
                    target_mode: Some("External".to_string()),
                },
                Relationship {
                    id: "rId3".to_string(),
                    rel_type: rel_types::HYPERLINK.to_string(),
                    target: "https://b.com".to_string(),
                    target_mode: Some("External".to_string()),
                },
            ],
        };
        assert_eq!(next_rel_id(&rels), "rId4");
    }
}
