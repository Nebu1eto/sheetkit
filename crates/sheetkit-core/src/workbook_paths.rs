use sheetkit_xml::relationships::Relationships;

/// Create an empty relationships container.
pub(crate) fn default_relationships() -> Relationships {
    Relationships {
        xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
        relationships: vec![],
    }
}

/// Resolve a relationship target against the source part path.
///
/// Both arguments are OOXML package-internal paths (e.g. `xl/workbook.xml`).
pub(crate) fn resolve_relationship_target(source_part: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }

    let base_dir = source_part
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or_default();
    let mut parts: Vec<&str> = if base_dir.is_empty() {
        vec![]
    } else {
        base_dir.split('/').collect()
    };

    for seg in target.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(seg),
        }
    }

    parts.join("/")
}

/// Get the `.rels` part path for a package part.
pub(crate) fn relationship_part_path(part_path: &str) -> String {
    let normalized = part_path.trim_start_matches('/');
    let (dir, file) = normalized.rsplit_once('/').unwrap_or(("", normalized));
    if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    }
}

/// Build a relative relationship target from `source_part` to `target_part`.
pub(crate) fn relative_relationship_target(source_part: &str, target_part: &str) -> String {
    let source_dir = source_part
        .trim_start_matches('/')
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or_default();
    let source_parts: Vec<&str> = if source_dir.is_empty() {
        vec![]
    } else {
        source_dir.split('/').collect()
    };
    let target_parts: Vec<&str> = target_part.trim_start_matches('/').split('/').collect();

    let mut common = 0usize;
    while common < source_parts.len()
        && common < target_parts.len()
        && source_parts[common] == target_parts[common]
    {
        common += 1;
    }

    let mut rel_parts: Vec<String> = Vec::new();
    for _ in 0..(source_parts.len() - common) {
        rel_parts.push("..".to_string());
    }
    rel_parts.extend(target_parts[common..].iter().map(|s| s.to_string()));

    if rel_parts.is_empty() {
        ".".to_string()
    } else {
        rel_parts.join("/")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_relationship_target() {
        assert_eq!(
            resolve_relationship_target("xl/workbook.xml", "worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_relationship_target("xl/worksheets/sheet1.xml", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
    }

    #[test]
    fn test_relationship_part_path() {
        assert_eq!(
            relationship_part_path("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );
        assert_eq!(
            relationship_part_path("xl/drawings/drawing3.xml"),
            "xl/drawings/_rels/drawing3.xml.rels"
        );
    }

    #[test]
    fn test_relative_relationship_target() {
        assert_eq!(
            relative_relationship_target("xl/worksheets/sheet1.xml", "xl/drawings/drawing1.xml"),
            "../drawings/drawing1.xml"
        );
        assert_eq!(
            relative_relationship_target("xl/drawings/drawing1.xml", "xl/charts/chart1.xml"),
            "../charts/chart1.xml"
        );
    }
}
