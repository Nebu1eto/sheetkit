/// Typed deferred auxiliary parts for on-demand parsing.
///
/// When a workbook is opened in Lazy or Stream mode, auxiliary parts
/// (comments, charts, images, doc properties, pivot tables, etc.) are
/// stored as raw bytes instead of being parsed immediately. This module
/// provides a typed index over those raw bytes and on-demand loaders
/// that parse each category the first time it is accessed.
///
/// Dirty tracking records which categories have been modified since
/// loading so that only changed parts are re-serialized on save.
/// Untouched deferred parts are written back as raw bytes for perfect
/// round-trip fidelity.
use std::collections::HashMap;

/// Classifies a ZIP path into a deferred part category, or `None`
/// if the path does not belong to any recognized auxiliary category.
pub(crate) fn classify_deferred_path(path: &str) -> Option<AuxCategory> {
    if path.starts_with("xl/comments") && path.ends_with(".xml") {
        return Some(AuxCategory::Comments);
    }
    if path.starts_with("xl/drawings/vmlDrawing") && path.ends_with(".vml") {
        return Some(AuxCategory::Vml);
    }
    if path.starts_with("xl/drawings/") && path.ends_with(".xml") {
        return Some(AuxCategory::Drawings);
    }
    if path.starts_with("xl/charts/") && path.ends_with(".xml") {
        return Some(AuxCategory::Charts);
    }
    if path.starts_with("xl/media/") {
        return Some(AuxCategory::Images);
    }
    if path == "docProps/core.xml" || path == "docProps/app.xml" || path == "docProps/custom.xml" {
        return Some(AuxCategory::DocProperties);
    }
    if path.starts_with("xl/pivotTables/") && path.ends_with(".xml") {
        return Some(AuxCategory::PivotTables);
    }
    if path.starts_with("xl/pivotCache/") && path.ends_with(".xml") {
        return Some(AuxCategory::PivotCaches);
    }
    if path.starts_with("xl/tables/") && path.ends_with(".xml") {
        return Some(AuxCategory::Tables);
    }
    if path.starts_with("xl/slicers/") && path.ends_with(".xml") {
        return Some(AuxCategory::Slicers);
    }
    if path.starts_with("xl/slicerCaches/") && path.ends_with(".xml") {
        return Some(AuxCategory::SlicerCaches);
    }
    if path.starts_with("xl/threadedComments/") && path.ends_with(".xml") {
        return Some(AuxCategory::ThreadedComments);
    }
    if path == "xl/persons/person.xml" {
        return Some(AuxCategory::PersonList);
    }
    if path == "xl/vbaProject.bin" {
        return Some(AuxCategory::Vba);
    }
    if path.starts_with("xl/drawings/_rels/") && path.ends_with(".rels") {
        return Some(AuxCategory::DrawingRels);
    }
    None
}

/// Categories of auxiliary parts that can be deferred.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum AuxCategory {
    Comments,
    Vml,
    Drawings,
    DrawingRels,
    Charts,
    Images,
    DocProperties,
    PivotTables,
    PivotCaches,
    Tables,
    Slicers,
    SlicerCaches,
    ThreadedComments,
    PersonList,
    Vba,
}

/// Typed index of deferred auxiliary parts.
///
/// Each entry stores the ZIP path and raw bytes, grouped by category.
/// Parts that have been hydrated (parsed into their typed fields on the
/// Workbook) are removed from this index. On save, remaining entries
/// are written back as raw bytes.
#[derive(Debug, Default)]
pub(crate) struct DeferredAuxParts {
    entries: HashMap<AuxCategory, Vec<(String, Vec<u8>)>>,
    hydrated: u32,
    dirty: u32,
}

impl DeferredAuxParts {
    /// Create a new empty deferred parts index.
    pub(crate) fn new() -> Self {
        Self::default()
    }

    /// Insert a deferred part, classifying it by ZIP path.
    /// Returns true if the path was classified and stored, false if unknown.
    pub(crate) fn insert(&mut self, path: String, data: Vec<u8>) -> bool {
        if let Some(cat) = classify_deferred_path(&path) {
            self.entries.entry(cat).or_default().push((path, data));
            true
        } else {
            false
        }
    }

    /// Returns true when no deferred parts are stored.
    pub(crate) fn is_empty(&self) -> bool {
        self.entries.values().all(|v| v.is_empty())
    }

    /// Check whether any parts exist in the given category.
    pub(crate) fn has_category(&self, cat: AuxCategory) -> bool {
        self.entries.get(&cat).is_some_and(|v| !v.is_empty())
    }

    /// Take all parts for a category, removing them from the index.
    /// Used during hydration to move raw bytes into typed fields.
    pub(crate) fn take(&mut self, cat: AuxCategory) -> Vec<(String, Vec<u8>)> {
        self.hydrated |= category_bit(cat);
        self.entries.remove(&cat).unwrap_or_default()
    }

    /// Remove a specific path from a category. Returns the raw bytes if found.
    pub(crate) fn remove_path(&mut self, cat: AuxCategory, path: &str) -> Option<Vec<u8>> {
        if let Some(entries) = self.entries.get_mut(&cat) {
            if let Some(pos) = entries.iter().position(|(p, _)| p == path) {
                let (_, data) = entries.remove(pos);
                return Some(data);
            }
        }
        None
    }

    /// Mark a category as dirty (modified after hydration).
    pub(crate) fn mark_dirty(&mut self, cat: AuxCategory) {
        self.dirty |= category_bit(cat);
    }

    /// Check if a category has been hydrated.
    #[allow(dead_code)]
    pub(crate) fn is_hydrated(&self, cat: AuxCategory) -> bool {
        self.hydrated & category_bit(cat) != 0
    }

    /// Check if a category is dirty (modified after hydration).
    #[allow(dead_code)]
    pub(crate) fn is_dirty(&self, cat: AuxCategory) -> bool {
        self.dirty & category_bit(cat) != 0
    }

    /// Iterate over all remaining (non-hydrated) parts across all categories.
    /// Yields (zip_path, raw_bytes) pairs.
    pub(crate) fn remaining_parts(&self) -> impl Iterator<Item = (&str, &[u8])> {
        self.entries
            .values()
            .flat_map(|v| v.iter().map(|(p, d)| (p.as_str(), d.as_slice())))
    }

    /// Check whether any deferred parts exist at all (across all categories).
    /// Used by the save path to decide whether deferred-aware logic is needed.
    pub(crate) fn has_any(&self) -> bool {
        !self.is_empty()
    }
}

/// Map a category to a unique bit position for the bitfield flags.
fn category_bit(cat: AuxCategory) -> u32 {
    match cat {
        AuxCategory::Comments => 1 << 0,
        AuxCategory::Vml => 1 << 1,
        AuxCategory::Drawings => 1 << 2,
        AuxCategory::DrawingRels => 1 << 3,
        AuxCategory::Charts => 1 << 4,
        AuxCategory::Images => 1 << 5,
        AuxCategory::DocProperties => 1 << 6,
        AuxCategory::PivotTables => 1 << 7,
        AuxCategory::PivotCaches => 1 << 8,
        AuxCategory::Tables => 1 << 9,
        AuxCategory::Slicers => 1 << 10,
        AuxCategory::SlicerCaches => 1 << 11,
        AuxCategory::ThreadedComments => 1 << 12,
        AuxCategory::PersonList => 1 << 13,
        AuxCategory::Vba => 1 << 14,
    }
}

use super::Workbook;
use sheetkit_xml::relationships::rel_types;

/// Inverse of `relationship_part_path`: given a rels path like
/// `xl/drawings/_rels/drawing1.xml.rels`, return the owner part path
/// `xl/drawings/drawing1.xml`.
fn rels_path_to_owner(rels_path: &str) -> String {
    // Strip the .rels suffix.
    let without_rels_ext = rels_path.strip_suffix(".rels").unwrap_or(rels_path);
    // The path should be like `<dir>/_rels/<file>`. Remove `/_rels/` segment.
    if let Some(pos) = without_rels_ext.rfind("/_rels/") {
        let dir = &without_rels_ext[..pos];
        let file = &without_rels_ext[pos + 7..]; // skip "/_rels/"
        format!("{dir}/{file}")
    } else if let Some(file) = without_rels_ext.strip_prefix("_rels/") {
        file.to_string()
    } else {
        without_rels_ext.to_string()
    }
}

impl Workbook {
    /// Hydrate deferred document properties (core, app, custom) into typed fields.
    ///
    /// Parses raw bytes for `docProps/core.xml`, `docProps/app.xml`, and
    /// `docProps/custom.xml` from the deferred parts index. Called automatically
    /// before any doc prop mutation to preserve existing values.
    pub(crate) fn hydrate_doc_props(&mut self) {
        if !self.deferred_parts.has_category(AuxCategory::DocProperties) {
            return;
        }

        if let Some(bytes) = self
            .deferred_parts
            .remove_path(AuxCategory::DocProperties, "docProps/core.xml")
        {
            let xml_str = String::from_utf8_lossy(&bytes);
            if let Ok(parsed) = sheetkit_xml::doc_props::deserialize_core_properties(&xml_str) {
                if self.core_properties.is_none() {
                    self.core_properties = Some(parsed);
                }
            }
        }

        if let Some(bytes) = self
            .deferred_parts
            .remove_path(AuxCategory::DocProperties, "docProps/app.xml")
        {
            let xml_str = String::from_utf8_lossy(&bytes);
            if let Ok(parsed) =
                quick_xml::de::from_str::<sheetkit_xml::doc_props::ExtendedProperties>(&xml_str)
            {
                if self.app_properties.is_none() {
                    self.app_properties = Some(parsed);
                }
            }
        }

        if let Some(bytes) = self
            .deferred_parts
            .remove_path(AuxCategory::DocProperties, "docProps/custom.xml")
        {
            let xml_str = String::from_utf8_lossy(&bytes);
            if let Ok(parsed) = sheetkit_xml::doc_props::deserialize_custom_properties(&xml_str) {
                if self.custom_properties.is_none() {
                    self.custom_properties = Some(parsed);
                }
            }
        }

        self.deferred_parts.mark_dirty(AuxCategory::DocProperties);
    }

    /// Hydrate deferred table parts into `self.tables`.
    ///
    /// Parses deferred table XML parts and associates them with sheets via
    /// worksheet relationships. Called before table queries or mutations.
    pub(crate) fn hydrate_tables(&mut self) {
        if !self.deferred_parts.has_category(AuxCategory::Tables) {
            return;
        }

        let table_entries = self.deferred_parts.take(AuxCategory::Tables);
        for (path, bytes) in table_entries {
            let xml_str = String::from_utf8_lossy(&bytes);
            if let Ok(table_xml) =
                quick_xml::de::from_str::<sheetkit_xml::table::TableXml>(&xml_str)
            {
                let sheet_idx = self.find_sheet_for_table_path(&path);
                self.tables.push((path, table_xml, sheet_idx));
            }
        }

        self.deferred_parts.mark_dirty(AuxCategory::Tables);
    }

    /// Hydrate deferred pivot table and cache parts into typed fields.
    pub(crate) fn hydrate_pivot_tables(&mut self) {
        if !self.deferred_parts.has_category(AuxCategory::PivotTables)
            && !self.deferred_parts.has_category(AuxCategory::PivotCaches)
        {
            return;
        }

        let pt_entries = self.deferred_parts.take(AuxCategory::PivotTables);
        for (path, bytes) in pt_entries {
            let xml_str = String::from_utf8_lossy(&bytes);
            if let Ok(pt) =
                quick_xml::de::from_str::<sheetkit_xml::pivot_table::PivotTableDefinition>(&xml_str)
            {
                self.pivot_tables.push((path, pt));
            }
        }

        let cache_entries = self.deferred_parts.take(AuxCategory::PivotCaches);
        for (path, bytes) in cache_entries {
            let xml_str = String::from_utf8_lossy(&bytes);
            if path.contains("pivotCacheRecords") {
                if let Ok(pcr) = quick_xml::de::from_str::<
                    sheetkit_xml::pivot_cache::PivotCacheRecords,
                >(&xml_str)
                {
                    self.pivot_cache_records.push((path, pcr));
                }
            } else if let Ok(pcd) =
                quick_xml::de::from_str::<sheetkit_xml::pivot_cache::PivotCacheDefinition>(&xml_str)
            {
                self.pivot_cache_defs.push((path, pcd));
            }
        }

        self.deferred_parts.mark_dirty(AuxCategory::PivotTables);
        self.deferred_parts.mark_dirty(AuxCategory::PivotCaches);
    }

    /// Hydrate deferred slicer and slicer cache parts.
    pub(crate) fn hydrate_slicers(&mut self) {
        if !self.deferred_parts.has_category(AuxCategory::Slicers)
            && !self.deferred_parts.has_category(AuxCategory::SlicerCaches)
        {
            return;
        }

        let slicer_entries = self.deferred_parts.take(AuxCategory::Slicers);
        for (path, bytes) in slicer_entries {
            let xml_str = String::from_utf8_lossy(&bytes);
            if let Ok(sd) =
                quick_xml::de::from_str::<sheetkit_xml::slicer::SlicerDefinitions>(&xml_str)
            {
                self.slicer_defs.push((path, sd));
            }
        }

        let cache_entries = self.deferred_parts.take(AuxCategory::SlicerCaches);
        for (path, bytes) in cache_entries {
            let xml_str = String::from_utf8_lossy(&bytes);
            if let Some(scd) = sheetkit_xml::slicer::parse_slicer_cache(&xml_str) {
                self.slicer_caches.push((path, scd));
            }
        }

        self.deferred_parts.mark_dirty(AuxCategory::Slicers);
        self.deferred_parts.mark_dirty(AuxCategory::SlicerCaches);
    }

    /// Hydrate deferred threaded comments and person list.
    pub(crate) fn hydrate_threaded_comments(&mut self) {
        if !self
            .deferred_parts
            .has_category(AuxCategory::ThreadedComments)
            && !self.deferred_parts.has_category(AuxCategory::PersonList)
        {
            return;
        }

        let tc_entries = self.deferred_parts.take(AuxCategory::ThreadedComments);
        for (path, bytes) in tc_entries {
            let xml_str = String::from_utf8_lossy(&bytes);
            if let Ok(tc) = quick_xml::de::from_str::<
                sheetkit_xml::threaded_comment::ThreadedComments,
            >(&xml_str)
            {
                let sheet_idx = self.find_sheet_for_threaded_comment_path(&path);
                if let Some(idx) = sheet_idx {
                    if idx < self.sheet_threaded_comments.len() {
                        self.sheet_threaded_comments[idx] = Some(tc);
                    }
                }
            }
        }

        let person_entries = self.deferred_parts.take(AuxCategory::PersonList);
        for (_path, bytes) in person_entries {
            let xml_str = String::from_utf8_lossy(&bytes);
            if let Ok(pl) =
                quick_xml::de::from_str::<sheetkit_xml::threaded_comment::PersonList>(&xml_str)
            {
                self.person_list = pl;
            }
        }

        self.deferred_parts
            .mark_dirty(AuxCategory::ThreadedComments);
        self.deferred_parts.mark_dirty(AuxCategory::PersonList);
    }

    /// Hydrate deferred VBA project binary.
    #[allow(dead_code)]
    pub(crate) fn hydrate_vba(&mut self) {
        if !self.deferred_parts.has_category(AuxCategory::Vba) {
            return;
        }

        if let Some(bytes) = self
            .deferred_parts
            .remove_path(AuxCategory::Vba, "xl/vbaProject.bin")
        {
            if self.vba_blob.is_none() {
                self.vba_blob = Some(bytes);
            }
        }

        self.deferred_parts.mark_dirty(AuxCategory::Vba);
    }

    /// Hydrate deferred drawings, charts, images, and their relationships.
    ///
    /// This is the most complex hydration because drawings reference charts and
    /// images through relationship files. The method parses all drawing XML,
    /// drawing relationship files, chart XML, and image binary data.
    pub(crate) fn hydrate_drawings(&mut self) {
        use sheetkit_xml::chart::ChartSpace;
        use sheetkit_xml::drawing::WsDr;
        use sheetkit_xml::relationships::Relationships;

        let needs_drawings = self.deferred_parts.has_category(AuxCategory::Drawings);
        let needs_charts = self.deferred_parts.has_category(AuxCategory::Charts);
        let needs_images = self.deferred_parts.has_category(AuxCategory::Images);
        let needs_drawing_rels = self.deferred_parts.has_category(AuxCategory::DrawingRels);

        if !needs_drawings && !needs_charts && !needs_images && !needs_drawing_rels {
            return;
        }

        if needs_drawings {
            let drawing_entries = self.deferred_parts.take(AuxCategory::Drawings);
            for (path, bytes) in drawing_entries {
                let xml_str = String::from_utf8_lossy(&bytes);
                if let Ok(drawing) = quick_xml::de::from_str::<WsDr>(&xml_str) {
                    let idx = self.drawings.len();
                    self.drawings.push((path.clone(), drawing));

                    // Link to sheets via worksheet_drawings.
                    for (sheet_idx, rels) in &self.worksheet_rels {
                        if rels.relationships.iter().any(|r| {
                            r.rel_type == rel_types::DRAWING
                                && crate::workbook_paths::resolve_relationship_target(
                                    &self.sheet_part_path(*sheet_idx),
                                    &r.target,
                                ) == path
                        }) {
                            self.worksheet_drawings.insert(*sheet_idx, idx);
                        }
                    }
                }
            }
        }

        if needs_drawing_rels {
            let rels_entries = self.deferred_parts.take(AuxCategory::DrawingRels);
            for (path, bytes) in rels_entries {
                let xml_str = String::from_utf8_lossy(&bytes);
                if let Ok(rels) = quick_xml::de::from_str::<Relationships>(&xml_str) {
                    // Find the drawing index this .rels file belongs to.
                    let drawing_path = rels_path_to_owner(&path);
                    if let Some(idx) = self.drawings.iter().position(|(p, _)| *p == drawing_path) {
                        self.drawing_rels.insert(idx, rels);
                    }
                }
            }
        }

        if needs_charts {
            let chart_entries = self.deferred_parts.take(AuxCategory::Charts);
            for (path, bytes) in chart_entries {
                let xml_str = String::from_utf8_lossy(&bytes);
                match quick_xml::de::from_str::<ChartSpace>(&xml_str) {
                    Ok(chart) => {
                        self.charts.push((path, chart));
                    }
                    Err(_) => {
                        self.raw_charts.push((path, bytes));
                    }
                }
            }
        }

        if needs_images {
            let image_entries = self.deferred_parts.take(AuxCategory::Images);
            for (path, bytes) in image_entries {
                self.images.push((path, bytes));
            }
        }

        self.deferred_parts.mark_dirty(AuxCategory::Drawings);
        self.deferred_parts.mark_dirty(AuxCategory::DrawingRels);
        self.deferred_parts.mark_dirty(AuxCategory::Charts);
        self.deferred_parts.mark_dirty(AuxCategory::Images);
    }

    /// Find the sheet index that references a given table path via worksheet rels.
    fn find_sheet_for_table_path(&self, table_path: &str) -> usize {
        for (sheet_idx, rels) in &self.worksheet_rels {
            for rel in &rels.relationships {
                if rel.rel_type == rel_types::TABLE {
                    let resolved = crate::workbook_paths::resolve_relationship_target(
                        &self.sheet_part_path(*sheet_idx),
                        &rel.target,
                    );
                    if resolved == table_path {
                        return *sheet_idx;
                    }
                }
            }
        }
        0
    }

    /// Find the sheet index that references a given threaded comment path.
    fn find_sheet_for_threaded_comment_path(&self, tc_path: &str) -> Option<usize> {
        for (sheet_idx, rels) in &self.worksheet_rels {
            for rel in &rels.relationships {
                if rel.rel_type == sheetkit_xml::threaded_comment::REL_TYPE_THREADED_COMMENT {
                    let resolved = crate::workbook_paths::resolve_relationship_target(
                        &self.sheet_part_path(*sheet_idx),
                        &rel.target,
                    );
                    if resolved == tc_path {
                        return Some(*sheet_idx);
                    }
                }
            }
        }
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn classify_comment_path() {
        assert_eq!(
            classify_deferred_path("xl/comments1.xml"),
            Some(AuxCategory::Comments)
        );
        assert_eq!(
            classify_deferred_path("xl/comments12.xml"),
            Some(AuxCategory::Comments)
        );
    }

    #[test]
    fn classify_vml_path() {
        assert_eq!(
            classify_deferred_path("xl/drawings/vmlDrawing1.vml"),
            Some(AuxCategory::Vml)
        );
    }

    #[test]
    fn classify_drawing_path() {
        assert_eq!(
            classify_deferred_path("xl/drawings/drawing1.xml"),
            Some(AuxCategory::Drawings)
        );
    }

    #[test]
    fn classify_drawing_rels_path() {
        assert_eq!(
            classify_deferred_path("xl/drawings/_rels/drawing1.xml.rels"),
            Some(AuxCategory::DrawingRels)
        );
    }

    #[test]
    fn classify_chart_path() {
        assert_eq!(
            classify_deferred_path("xl/charts/chart1.xml"),
            Some(AuxCategory::Charts)
        );
    }

    #[test]
    fn classify_image_path() {
        assert_eq!(
            classify_deferred_path("xl/media/image1.png"),
            Some(AuxCategory::Images)
        );
    }

    #[test]
    fn classify_doc_props_paths() {
        assert_eq!(
            classify_deferred_path("docProps/core.xml"),
            Some(AuxCategory::DocProperties)
        );
        assert_eq!(
            classify_deferred_path("docProps/app.xml"),
            Some(AuxCategory::DocProperties)
        );
        assert_eq!(
            classify_deferred_path("docProps/custom.xml"),
            Some(AuxCategory::DocProperties)
        );
    }

    #[test]
    fn classify_pivot_table_path() {
        assert_eq!(
            classify_deferred_path("xl/pivotTables/pivotTable1.xml"),
            Some(AuxCategory::PivotTables)
        );
    }

    #[test]
    fn classify_pivot_cache_path() {
        assert_eq!(
            classify_deferred_path("xl/pivotCache/pivotCacheDefinition1.xml"),
            Some(AuxCategory::PivotCaches)
        );
        assert_eq!(
            classify_deferred_path("xl/pivotCache/pivotCacheRecords1.xml"),
            Some(AuxCategory::PivotCaches)
        );
    }

    #[test]
    fn classify_table_path() {
        assert_eq!(
            classify_deferred_path("xl/tables/table1.xml"),
            Some(AuxCategory::Tables)
        );
    }

    #[test]
    fn classify_slicer_paths() {
        assert_eq!(
            classify_deferred_path("xl/slicers/slicer1.xml"),
            Some(AuxCategory::Slicers)
        );
        assert_eq!(
            classify_deferred_path("xl/slicerCaches/slicerCache1.xml"),
            Some(AuxCategory::SlicerCaches)
        );
    }

    #[test]
    fn classify_threaded_comment_path() {
        assert_eq!(
            classify_deferred_path("xl/threadedComments/threadedComment1.xml"),
            Some(AuxCategory::ThreadedComments)
        );
    }

    #[test]
    fn classify_person_list_path() {
        assert_eq!(
            classify_deferred_path("xl/persons/person.xml"),
            Some(AuxCategory::PersonList)
        );
    }

    #[test]
    fn classify_vba_path() {
        assert_eq!(
            classify_deferred_path("xl/vbaProject.bin"),
            Some(AuxCategory::Vba)
        );
    }

    #[test]
    fn classify_unknown_path() {
        assert_eq!(classify_deferred_path("xl/foo/bar.xml"), None);
        assert_eq!(
            classify_deferred_path("xl/printerSettings/printerSettings1.bin"),
            None
        );
    }

    #[test]
    fn insert_and_take() {
        let mut deferred = DeferredAuxParts::new();
        deferred.insert("xl/comments1.xml".to_string(), b"<Comments/>".to_vec());
        deferred.insert(
            "xl/charts/chart1.xml".to_string(),
            b"<c:chartSpace/>".to_vec(),
        );

        assert!(deferred.has_category(AuxCategory::Comments));
        assert!(deferred.has_category(AuxCategory::Charts));
        assert!(!deferred.has_category(AuxCategory::Tables));

        let comments = deferred.take(AuxCategory::Comments);
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].0, "xl/comments1.xml");
        assert!(deferred.is_hydrated(AuxCategory::Comments));
        assert!(!deferred.has_category(AuxCategory::Comments));
    }

    #[test]
    fn remove_path() {
        let mut deferred = DeferredAuxParts::new();
        deferred.insert("xl/comments1.xml".to_string(), b"data1".to_vec());
        deferred.insert("xl/comments2.xml".to_string(), b"data2".to_vec());

        let data = deferred.remove_path(AuxCategory::Comments, "xl/comments1.xml");
        assert_eq!(data.as_deref(), Some(b"data1".as_slice()));
        assert!(deferred.has_category(AuxCategory::Comments));

        let data2 = deferred.remove_path(AuxCategory::Comments, "xl/comments2.xml");
        assert_eq!(data2.as_deref(), Some(b"data2".as_slice()));
        assert!(!deferred.has_category(AuxCategory::Comments));
    }

    #[test]
    fn dirty_tracking() {
        let mut deferred = DeferredAuxParts::new();
        assert!(!deferred.is_dirty(AuxCategory::DocProperties));

        deferred.mark_dirty(AuxCategory::DocProperties);
        assert!(deferred.is_dirty(AuxCategory::DocProperties));
        assert!(!deferred.is_dirty(AuxCategory::Charts));
    }

    #[test]
    fn remaining_parts_iteration() {
        let mut deferred = DeferredAuxParts::new();
        deferred.insert("xl/comments1.xml".to_string(), b"a".to_vec());
        deferred.insert("xl/charts/chart1.xml".to_string(), b"b".to_vec());

        deferred.take(AuxCategory::Comments);

        let remaining: Vec<(&str, &[u8])> = deferred.remaining_parts().collect();
        assert_eq!(remaining.len(), 1);
        assert_eq!(remaining[0].0, "xl/charts/chart1.xml");
    }

    #[test]
    fn is_empty_checks() {
        let mut deferred = DeferredAuxParts::new();
        assert!(deferred.is_empty());

        deferred.insert("xl/comments1.xml".to_string(), b"data".to_vec());
        assert!(!deferred.is_empty());

        deferred.take(AuxCategory::Comments);
        assert!(deferred.is_empty());
    }

    #[test]
    fn rels_path_to_owner_basic() {
        assert_eq!(
            super::rels_path_to_owner("xl/drawings/_rels/drawing1.xml.rels"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            super::rels_path_to_owner("xl/worksheets/_rels/sheet1.xml.rels"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(super::rels_path_to_owner("_rels/.rels"), "");
    }

    use crate::workbook::open_options::{OpenOptions, ReadMode};
    use crate::workbook::Workbook;

    #[test]
    fn lazy_open_doc_props_roundtrip() {
        let mut wb = Workbook::new();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Test Title".to_string()),
            creator: Some("Test Author".to_string()),
            ..Default::default()
        });
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let props = wb3.get_doc_props();
        assert_eq!(props.title.as_deref(), Some("Test Title"));
        assert_eq!(props.creator.as_deref(), Some("Test Author"));
    }

    #[test]
    fn lazy_open_set_doc_props_hydrates_first() {
        let mut wb = Workbook::new();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Original".to_string()),
            ..Default::default()
        });
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Setting new doc props should hydrate first, then overwrite.
        wb2.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Updated".to_string()),
            ..Default::default()
        });
        let saved = wb2.save_to_buffer().unwrap();

        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let props = wb3.get_doc_props();
        assert_eq!(props.title.as_deref(), Some("Updated"));
    }

    #[test]
    fn lazy_open_table_roundtrip() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "A1",
            crate::cell::CellValue::String("Name".to_string()),
        )
        .unwrap();
        wb.set_cell_value(
            "Sheet1",
            "B1",
            crate::cell::CellValue::String("Value".to_string()),
        )
        .unwrap();
        wb.set_cell_value(
            "Sheet1",
            "A2",
            crate::cell::CellValue::String("Alice".to_string()),
        )
        .unwrap();
        wb.set_cell_value("Sheet1", "B2", crate::cell::CellValue::Number(10.0))
            .unwrap();
        wb.add_table(
            "Sheet1",
            &crate::table::TableConfig {
                name: "TestTable".to_string(),
                display_name: "TestTable".to_string(),
                range: "A1:B2".to_string(),
                columns: vec![
                    crate::table::TableColumn {
                        name: "Name".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    crate::table::TableColumn {
                        name: "Value".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                ],
                ..Default::default()
            },
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let tables = wb3.get_tables("Sheet1").unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "TestTable");
    }

    #[test]
    fn lazy_open_add_table_after_deferred() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "A1",
            crate::cell::CellValue::String("Name".to_string()),
        )
        .unwrap();
        wb.set_cell_value(
            "Sheet1",
            "B1",
            crate::cell::CellValue::String("Value".to_string()),
        )
        .unwrap();
        wb.set_cell_value(
            "Sheet1",
            "A2",
            crate::cell::CellValue::String("Alice".to_string()),
        )
        .unwrap();
        wb.set_cell_value("Sheet1", "B2", crate::cell::CellValue::Number(10.0))
            .unwrap();
        wb.add_table(
            "Sheet1",
            &crate::table::TableConfig {
                name: "ExistingTable".to_string(),
                display_name: "ExistingTable".to_string(),
                range: "A1:B2".to_string(),
                columns: vec![
                    crate::table::TableColumn {
                        name: "Name".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    crate::table::TableColumn {
                        name: "Value".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                ],
                ..Default::default()
            },
        )
        .unwrap();

        wb.set_cell_value(
            "Sheet1",
            "D1",
            crate::cell::CellValue::String("Col1".to_string()),
        )
        .unwrap();
        wb.set_cell_value(
            "Sheet1",
            "D2",
            crate::cell::CellValue::String("Row1".to_string()),
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Adding a new table should hydrate existing tables first.
        wb2.add_table(
            "Sheet1",
            &crate::table::TableConfig {
                name: "NewTable".to_string(),
                display_name: "NewTable".to_string(),
                range: "D1:D2".to_string(),
                columns: vec![crate::table::TableColumn {
                    name: "Col1".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                }],
                ..Default::default()
            },
        )
        .unwrap();

        let saved = wb2.save_to_buffer().unwrap();
        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let tables = wb3.get_tables("Sheet1").unwrap();
        assert_eq!(tables.len(), 2);
        let names: Vec<&str> = tables.iter().map(|t| t.name.as_str()).collect();
        assert!(names.contains(&"ExistingTable"));
        assert!(names.contains(&"NewTable"));
    }

    #[test]
    fn typed_deferred_parts_classifies_correctly() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "A1",
            crate::cell::CellValue::String("data".to_string()),
        )
        .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Tester".to_string(),
                text: "comment".to_string(),
            },
        )
        .unwrap();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Test".to_string()),
            ..Default::default()
        });
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // The typed deferred parts should have entries.
        assert!(wb2.deferred_parts.has_any());
        // Comments should be in the Comments category.
        assert!(wb2.deferred_parts.has_category(AuxCategory::Comments));
        // Doc props should be in the DocProperties category.
        assert!(wb2.deferred_parts.has_category(AuxCategory::DocProperties));
    }

    #[test]
    fn deferred_parts_not_duplicated_on_save() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "A1",
            crate::cell::CellValue::String("data".to_string()),
        )
        .unwrap();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Round-trip".to_string()),
            ..Default::default()
        });
        let buf = wb.save_to_buffer().unwrap();

        // Open in lazy mode and save without touching anything.
        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        // Re-open in eager mode and check the doc props are preserved.
        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let props = wb3.get_doc_props();
        assert_eq!(props.title.as_deref(), Some("Round-trip"));
    }

    #[test]
    fn hydrate_doc_props_on_set_custom_property() {
        let mut wb = Workbook::new();
        wb.set_custom_property(
            "OriginalProp",
            crate::doc_props::CustomPropertyValue::String("original".to_string()),
        );
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Setting a new custom property should hydrate existing ones.
        wb2.set_custom_property(
            "NewProp",
            crate::doc_props::CustomPropertyValue::String("new".to_string()),
        );
        let saved = wb2.save_to_buffer().unwrap();

        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let orig = wb3.get_custom_property("OriginalProp");
        assert!(orig.is_some(), "Original custom property must survive");
        let new_val = wb3.get_custom_property("NewProp");
        assert!(new_val.is_some(), "New custom property must be present");
    }
}
