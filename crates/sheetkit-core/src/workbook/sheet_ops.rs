use super::*;

fn sheet_name_matches(sheet_name: &str, target_sheet_name_lowercase: &str) -> bool {
    sheet_name.to_lowercase() == target_sheet_name_lowercase
}

fn validate_anchor_shift<F>(col: u32, row: u32, shift_cell: F) -> Result<()>
where
    F: Fn(u32, u32) -> (u32, u32),
{
    let col = col.checked_add(1).ok_or(Error::InvalidColumnNumber(col))?;
    let row = row.checked_add(1).ok_or(Error::InvalidRowNumber(row))?;
    let (col, row) = shift_cell(col, row);
    if !(1..=crate::utils::constants::MAX_COLUMNS).contains(&col) {
        return Err(Error::InvalidColumnNumber(col));
    }
    if !(1..=crate::utils::constants::MAX_ROWS).contains(&row) {
        return Err(Error::InvalidRowNumber(row));
    }
    Ok(())
}

fn shift_references_for_owner<F>(
    text: &str,
    owner_sheet_idx: usize,
    target_sheet_idx: usize,
    target_sheet_name_lowercase: &str,
    shift_cell: F,
) -> Result<String>
where
    F: Fn(u32, u32) -> (u32, u32) + Copy,
{
    crate::cell_ref_shift::shift_cell_references_with_abs_and_scope(
        text,
        owner_sheet_idx == target_sheet_idx,
        |sheet_name| sheet_name_matches(sheet_name, target_sheet_name_lowercase),
        |col, row, _, _| shift_cell(col, row),
    )
}

fn shift_duplicated_row_formula(formula: &str) -> Result<String> {
    crate::cell_ref_shift::shift_cell_references_with_abs_and_scope(
        formula,
        true,
        |_| true,
        |col, row, _abs_col, abs_row| (col, if abs_row { row } else { row.saturating_add(1) }),
    )
}

fn next_numbered_part_path(
    prefix: &str,
    suffix: &str,
    existing: &std::collections::HashSet<String>,
) -> String {
    let mut number = 1usize;
    loop {
        let candidate = format!("{prefix}{number}{suffix}");
        if !existing.contains(&candidate) {
            return candidate;
        }
        number += 1;
    }
}

fn validate_drawing_relationship_ids(drawing: &WsDr, relationships: &Relationships) -> Result<()> {
    let mut required = Vec::new();
    for anchor in &drawing.two_cell_anchors {
        if let Some(frame) = &anchor.graphic_frame {
            required.push((
                frame.graphic.graphic_data.chart.r_id.as_str(),
                rel_types::CHART,
            ));
        }
        if let Some(picture) = &anchor.pic {
            required.push((picture.blip_fill.blip.r_embed.as_str(), rel_types::IMAGE));
        }
    }
    for anchor in &drawing.one_cell_anchors {
        if let Some(picture) = &anchor.pic {
            required.push((picture.blip_fill.blip.r_embed.as_str(), rel_types::IMAGE));
        }
    }
    for (r_id, rel_type) in required {
        if !relationships
            .relationships
            .iter()
            .any(|relationship| relationship.id == r_id && relationship.rel_type == rel_type)
        {
            return Err(Error::InvalidArgument(format!(
                "drawing relationship '{r_id}' is missing or has the wrong type"
            )));
        }
    }
    Ok(())
}

fn validate_worksheet_relationship_ids(
    worksheet: &WorksheetXml,
    relationships: &Relationships,
) -> Result<()> {
    let relationship_matches = |r_id: &str, rel_type: &str| {
        relationships
            .relationships
            .iter()
            .any(|relationship| relationship.id == r_id && relationship.rel_type == rel_type)
    };
    if let Some(drawing) = &worksheet.drawing {
        if !relationship_matches(&drawing.r_id, rel_types::DRAWING) {
            return Err(Error::InvalidArgument(format!(
                "worksheet drawing relationship '{}' is missing or has the wrong type",
                drawing.r_id
            )));
        }
    }
    if let Some(legacy_drawing) = &worksheet.legacy_drawing {
        if !relationship_matches(&legacy_drawing.r_id, rel_types::VML_DRAWING) {
            return Err(Error::InvalidArgument(format!(
                "worksheet legacy drawing relationship '{}' is missing or has the wrong type",
                legacy_drawing.r_id
            )));
        }
    }
    if let Some(page_setup) = &worksheet.page_setup {
        if let Some(r_id) = &page_setup.r_id {
            if !relationship_matches(r_id, rel_types::PRINTER_SETTINGS) {
                return Err(Error::InvalidArgument(format!(
                    "worksheet page setup relationship '{r_id}' is missing or has the wrong type"
                )));
            }
        }
    }
    if let Some(hyperlinks) = &worksheet.hyperlinks {
        for hyperlink in &hyperlinks.hyperlinks {
            if let Some(r_id) = &hyperlink.r_id {
                if !relationship_matches(r_id, rel_types::HYPERLINK) {
                    return Err(Error::InvalidArgument(format!(
                        "worksheet hyperlink relationship '{r_id}' is missing or has the wrong type"
                    )));
                }
            }
        }
    }
    if let Some(table_parts) = &worksheet.table_parts {
        for table_part in &table_parts.table_parts {
            if !relationship_matches(&table_part.r_id, rel_types::TABLE) {
                return Err(Error::InvalidArgument(format!(
                    "worksheet table relationship '{}' is missing or has the wrong type",
                    table_part.r_id
                )));
            }
        }
    }
    Ok(())
}

fn clone_threaded_comments_with_new_ids(
    source: &sheetkit_xml::threaded_comment::ThreadedComments,
    occupied_ids: &mut std::collections::HashSet<String>,
) -> sheetkit_xml::threaded_comment::ThreadedComments {
    let mut cloned = source.clone();
    let mut remapped = std::collections::HashMap::new();
    for comment in &mut cloned.comments {
        let new_id = loop {
            let candidate = format!("{{{}}}", uuid::Uuid::new_v4().to_string().to_uppercase());
            if occupied_ids.insert(candidate.clone()) {
                break candidate;
            }
        };
        remapped.insert(comment.id.clone(), new_id.clone());
        comment.id = new_id;
    }
    for comment in &mut cloned.comments {
        if let Some(parent_id) = &mut comment.parent_id {
            if let Some(new_parent_id) = remapped.get(parent_id) {
                *parent_id = new_parent_id.clone();
            }
        }
    }
    cloned
}

fn visit_series_references<F>(
    series: &mut sheetkit_xml::chart::Series,
    visitor: &mut F,
) -> Result<()>
where
    F: FnMut(&mut String) -> Result<()>,
{
    if let Some(text) = &mut series.tx {
        if let Some(reference) = &mut text.str_ref {
            visitor(&mut reference.f)?;
        }
    }
    if let Some(category) = &mut series.cat {
        if let Some(reference) = &mut category.str_ref {
            visitor(&mut reference.f)?;
        }
        if let Some(reference) = &mut category.num_ref {
            visitor(&mut reference.f)?;
        }
    }
    if let Some(value) = &mut series.val {
        if let Some(reference) = &mut value.num_ref {
            visitor(&mut reference.f)?;
        }
    }
    Ok(())
}

fn visit_scatter_references<F>(
    series: &mut sheetkit_xml::chart::ScatterSeries,
    visitor: &mut F,
) -> Result<()>
where
    F: FnMut(&mut String) -> Result<()>,
{
    if let Some(text) = &mut series.tx {
        if let Some(reference) = &mut text.str_ref {
            visitor(&mut reference.f)?;
        }
    }
    if let Some(category) = &mut series.x_val {
        if let Some(reference) = &mut category.str_ref {
            visitor(&mut reference.f)?;
        }
        if let Some(reference) = &mut category.num_ref {
            visitor(&mut reference.f)?;
        }
    }
    if let Some(value) = &mut series.y_val {
        if let Some(reference) = &mut value.num_ref {
            visitor(&mut reference.f)?;
        }
    }
    Ok(())
}

fn visit_bubble_references<F>(
    series: &mut sheetkit_xml::chart::BubbleSeries,
    visitor: &mut F,
) -> Result<()>
where
    F: FnMut(&mut String) -> Result<()>,
{
    if let Some(text) = &mut series.tx {
        if let Some(reference) = &mut text.str_ref {
            visitor(&mut reference.f)?;
        }
    }
    if let Some(category) = &mut series.x_val {
        if let Some(reference) = &mut category.str_ref {
            visitor(&mut reference.f)?;
        }
        if let Some(reference) = &mut category.num_ref {
            visitor(&mut reference.f)?;
        }
    }
    for value in [&mut series.y_val, &mut series.bubble_size]
        .into_iter()
        .flatten()
    {
        if let Some(reference) = &mut value.num_ref {
            visitor(&mut reference.f)?;
        }
    }
    Ok(())
}

fn visit_chart_references<F>(chart: &mut ChartSpace, mut visitor: F) -> Result<()>
where
    F: FnMut(&mut String) -> Result<()>,
{
    macro_rules! visit_series_chart {
        ($chart:expr) => {
            if let Some(chart) = &mut $chart {
                for series in &mut chart.series {
                    visit_series_references(series, &mut visitor)?;
                }
            }
        };
    }
    let plot_area = &mut chart.chart.plot_area;
    visit_series_chart!(plot_area.bar_chart);
    visit_series_chart!(plot_area.bar_3d_chart);
    visit_series_chart!(plot_area.line_chart);
    visit_series_chart!(plot_area.line_3d_chart);
    visit_series_chart!(plot_area.pie_chart);
    visit_series_chart!(plot_area.pie_3d_chart);
    visit_series_chart!(plot_area.doughnut_chart);
    visit_series_chart!(plot_area.area_chart);
    visit_series_chart!(plot_area.area_3d_chart);
    visit_series_chart!(plot_area.radar_chart);
    visit_series_chart!(plot_area.stock_chart);
    visit_series_chart!(plot_area.surface_chart);
    visit_series_chart!(plot_area.surface_3d_chart);
    visit_series_chart!(plot_area.of_pie_chart);
    if let Some(chart) = &mut plot_area.scatter_chart {
        for series in &mut chart.series {
            visit_scatter_references(series, &mut visitor)?;
        }
    }
    if let Some(chart) = &mut plot_area.bubble_chart {
        for series in &mut chart.series {
            visit_bubble_references(series, &mut visitor)?;
        }
    }
    Ok(())
}

impl Workbook {
    fn append_empty_sheet_state(&mut self) {
        if self.sheet_comments.len() < self.worksheets.len() {
            self.sheet_comments.push(None);
        }
        if self.sheet_sparklines.len() < self.worksheets.len() {
            self.sheet_sparklines.push(vec![]);
        }
        if self.sheet_vml.len() < self.worksheets.len() {
            self.sheet_vml.push(None);
        }
        if self.raw_sheet_xml.len() < self.worksheets.len() {
            self.raw_sheet_xml.push(None);
        }
        if self.sheet_dirty.len() < self.worksheets.len() {
            self.sheet_dirty.push(true);
        }
        if self.sheet_threaded_comments.len() < self.worksheets.len() {
            self.sheet_threaded_comments.push(None);
        }
        if self.sheet_form_controls.len() < self.worksheets.len() {
            self.sheet_form_controls.push(vec![]);
        }
    }

    /// Return the names of all sheets in workbook order.
    pub fn sheet_names(&self) -> Vec<&str> {
        self.worksheets
            .iter()
            .map(|(name, _)| name.as_str())
            .collect()
    }

    /// Create a new empty sheet with the given name. Returns the 0-based sheet index.
    pub fn new_sheet(&mut self, name: &str) -> Result<usize> {
        let occupied_paths = self.occupied_part_paths();
        let idx = crate::sheet::add_sheet_with_occupied_paths(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            name,
            WorksheetXml::default(),
            occupied_paths,
        )?;
        self.append_empty_sheet_state();
        self.rebuild_sheet_index();
        Ok(idx)
    }

    /// Delete a sheet by name.
    pub fn delete_sheet(&mut self, name: &str) -> Result<()> {
        let idx = self.sheet_index(name)?;
        self.assert_parallel_vecs_in_sync();
        if self.worksheets.len() <= 1 {
            return Err(Error::InvalidSheetName(
                "cannot delete the last sheet in a workbook".into(),
            ));
        }
        if self.worksheet_rels.get(&idx).is_some_and(|rels| {
            rels.relationships.iter().any(|relationship| {
                relationship.rel_type == rel_types::PIVOT_TABLE
                    || relationship.rel_type == rel_types::SLICER
                    || relationship.rel_type == rel_types::SLICER_CACHE
            })
        }) {
            return Err(Error::InvalidArgument(
                "cannot delete a sheet with pivot or slicer relationships safely".into(),
            ));
        }
        self.ensure_drawing_relationships_hydratable(idx)?;
        self.hydrate_lifecycle_parts_for_delete(idx)?;
        let deleted_drawing_idx = self.ensure_owned_drawing_parsed(idx)?;
        let deleted_sheet_path = self.sheet_part_path(idx);
        let deleted_direct_targets: Vec<String> = self
            .worksheet_rels
            .get(&idx)
            .into_iter()
            .flat_map(|rels| &rels.relationships)
            .filter(|relationship| relationship.target_mode.as_deref() != Some("External"))
            .map(|relationship| {
                crate::workbook_paths::resolve_relationship_target(
                    &deleted_sheet_path,
                    &relationship.target,
                )
            })
            .collect();

        crate::sheet::delete_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            name,
        )?;

        if let Some(defined_names) = &mut self.workbook_xml.defined_names {
            defined_names
                .defined_names
                .retain(|defined_name| defined_name.local_sheet_id != Some(idx as u32));
            for defined_name in &mut defined_names.defined_names {
                if let Some(local_sheet_id) = &mut defined_name.local_sheet_id {
                    if *local_sheet_id > idx as u32 {
                        *local_sheet_id -= 1;
                    }
                }
            }
        }
        if let Some(book_views) = &mut self.workbook_xml.book_views {
            let last_sheet_idx = self.worksheets.len().saturating_sub(1) as u32;
            for view in &mut book_views.workbook_views {
                if let Some(active_tab) = &mut view.active_tab {
                    if *active_tab > idx as u32 {
                        *active_tab -= 1;
                    }
                    *active_tab = (*active_tab).min(last_sheet_idx);
                }
            }
        }

        // Remove all per-sheet parallel data at once. After delete_sheet
        // above, worksheets has already been shortened by 1 so these
        // vectors must follow.
        self.sheet_comments.remove(idx);
        self.sheet_sparklines.remove(idx);
        self.sheet_vml.remove(idx);
        self.raw_sheet_xml.remove(idx);
        self.sheet_dirty.remove(idx);
        self.sheet_threaded_comments.remove(idx);
        self.sheet_form_controls.remove(idx);

        // Remove tables belonging to the deleted sheet and re-index remaining.
        self.tables.retain(|(_, _, si)| *si != idx);
        for (_, _, si) in &mut self.tables {
            if *si > idx {
                *si -= 1;
            }
        }

        // Remove and reindex streamed sheet data.
        self.streamed_sheets.remove(&idx);
        self.streamed_sheets = self
            .streamed_sheets
            .drain()
            .map(|(i, data)| if i > idx { (i - 1, data) } else { (i, data) })
            .collect();

        self.reindex_sheet_maps_after_delete(idx);
        if let Some(drawing_idx) = deleted_drawing_idx {
            self.remove_unreferenced_drawing(drawing_idx);
        }
        self.remove_unreferenced_unknown_targets(&deleted_direct_targets);
        self.rebuild_sheet_index();
        Ok(())
    }

    fn remove_unreferenced_unknown_targets(&mut self, deleted_targets: &[String]) {
        for target in deleted_targets {
            let still_referenced =
                self.worksheet_rels.iter().any(|(sheet_idx, rels)| {
                    let sheet_path = self.sheet_part_path(*sheet_idx);
                    rels.relationships.iter().any(|relationship| {
                        relationship.target_mode.as_deref() != Some("External")
                            && crate::workbook_paths::resolve_relationship_target(
                                &sheet_path,
                                &relationship.target,
                            ) == *target
                    })
                }) || self.workbook_rels.relationships.iter().any(|relationship| {
                    relationship.target_mode.as_deref() != Some("External")
                        && crate::workbook_paths::resolve_relationship_target(
                            "xl/workbook.xml",
                            &relationship.target,
                        ) == *target
                }) || self.package_rels.relationships.iter().any(|relationship| {
                    relationship.target_mode.as_deref() != Some("External")
                        && crate::workbook_paths::resolve_relationship_target(
                            "",
                            &relationship.target,
                        ) == *target
                }) || self.drawing_rels.iter().any(|(drawing_idx, rels)| {
                    let drawing_path = &self.drawings[*drawing_idx].0;
                    rels.relationships.iter().any(|relationship| {
                        relationship.target_mode.as_deref() != Some("External")
                            && crate::workbook_paths::resolve_relationship_target(
                                drawing_path,
                                &relationship.target,
                            ) == *target
                    })
                });
            if still_referenced {
                continue;
            }
            self.unknown_parts.retain(|(path, _)| path != target);
            self.content_types
                .overrides
                .retain(|entry| entry.part_name != format!("/{target}"));
        }
    }

    fn hydrate_lifecycle_parts_for_delete(&mut self, deleted_sheet_idx: usize) -> Result<()> {
        use crate::workbook::aux::AuxCategory;

        let parse_entries = |category, validate: &dyn Fn(&str) -> bool| -> Result<()> {
            if let Some(entries) = self.deferred_parts.entries(category) {
                for (path, bytes) in entries {
                    let xml = std::str::from_utf8(bytes)
                        .map_err(|error| Error::XmlParse(format!("{path}: {error}")))?;
                    if !validate(xml) {
                        return Err(Error::XmlDeserialize(format!(
                            "cannot hydrate lifecycle part '{path}'"
                        )));
                    }
                }
            }
            Ok(())
        };
        parse_entries(AuxCategory::Comments, &|xml| {
            quick_xml::de::from_str::<Comments>(xml).is_ok()
        })?;
        parse_entries(AuxCategory::Tables, &|xml| {
            quick_xml::de::from_str::<sheetkit_xml::table::TableXml>(xml).is_ok()
        })?;
        parse_entries(AuxCategory::ThreadedComments, &|xml| {
            quick_xml::de::from_str::<sheetkit_xml::threaded_comment::ThreadedComments>(xml).is_ok()
        })?;
        parse_entries(AuxCategory::PersonList, &|xml| {
            quick_xml::de::from_str::<sheetkit_xml::threaded_comment::PersonList>(xml).is_ok()
        })?;
        let owned_drawing = self.owned_drawing_path(deleted_sheet_idx);
        if let Some(drawing_path) = owned_drawing.as_deref() {
            let drawing_rels_path = relationship_part_path(drawing_path);
            for (path, bytes) in self
                .deferred_parts
                .entries(AuxCategory::Drawings)
                .into_iter()
                .flatten()
                .filter(|(path, _)| path == drawing_path)
            {
                let xml = std::str::from_utf8(bytes)
                    .map_err(|error| Error::XmlParse(format!("{path}: {error}")))?;
                quick_xml::de::from_str::<WsDr>(xml).map_err(|error| {
                    Error::XmlDeserialize(format!(
                        "cannot hydrate lifecycle part '{path}': {error}"
                    ))
                })?;
            }
            for (path, bytes) in self
                .deferred_parts
                .entries(AuxCategory::DrawingRels)
                .into_iter()
                .flatten()
                .filter(|(path, _)| path == &drawing_rels_path)
            {
                let xml = std::str::from_utf8(bytes)
                    .map_err(|error| Error::XmlParse(format!("{path}: {error}")))?;
                quick_xml::de::from_str::<Relationships>(xml).map_err(|error| {
                    Error::XmlDeserialize(format!(
                        "cannot hydrate lifecycle part '{path}': {error}"
                    ))
                })?;
            }
        }

        let comment_entries = self.deferred_parts.take(AuxCategory::Comments);
        for (path, bytes) in comment_entries {
            let xml = std::str::from_utf8(&bytes)
                .map_err(|error| Error::XmlParse(format!("{path}: {error}")))?;
            let comments: Comments = quick_xml::de::from_str(xml)
                .map_err(|error| Error::XmlDeserialize(format!("{path}: {error}")))?;
            if let Some(sheet_idx) = self.find_sheet_for_owned_rel_path(rel_types::COMMENTS, &path)
            {
                self.sheet_comments[sheet_idx] = Some(comments);
            }
        }
        let vml_entries = self.deferred_parts.take(AuxCategory::Vml);
        for (path, bytes) in vml_entries {
            if let Some(sheet_idx) =
                self.find_sheet_for_owned_rel_path(rel_types::VML_DRAWING, &path)
            {
                self.sheet_vml[sheet_idx] = Some(bytes);
            }
        }
        if !self.deferred_parts.is_empty() {
            self.deferred_parts.mark_dirty(AuxCategory::Comments);
            self.deferred_parts.mark_dirty(AuxCategory::Vml);
        }
        self.hydrate_tables();
        self.hydrate_threaded_comments();
        self.hydrate_drawings();
        Ok(())
    }

    fn find_sheet_for_owned_rel_path(&self, rel_type: &str, part_path: &str) -> Option<usize> {
        self.worksheet_rels.iter().find_map(|(sheet_idx, rels)| {
            rels.relationships
                .iter()
                .any(|relationship| {
                    relationship.rel_type == rel_type
                        && crate::workbook_paths::resolve_relationship_target(
                            &self.sheet_part_path(*sheet_idx),
                            &relationship.target,
                        ) == part_path
                })
                .then_some(*sheet_idx)
        })
    }

    fn remove_unreferenced_drawing(&mut self, drawing_idx: usize) {
        if self
            .worksheet_drawings
            .values()
            .any(|existing_idx| *existing_idx == drawing_idx)
        {
            return;
        }
        let Some((drawing_path, _)) = self.drawings.get(drawing_idx) else {
            return;
        };
        let drawing_path = drawing_path.clone();
        let removed_chart_paths: Vec<String> = self
            .drawing_rels
            .get(&drawing_idx)
            .into_iter()
            .flat_map(|rels| &rels.relationships)
            .filter(|relationship| relationship.rel_type == rel_types::CHART)
            .map(|relationship| {
                crate::workbook_paths::resolve_relationship_target(
                    &drawing_path,
                    &relationship.target,
                )
            })
            .collect();
        let removed_image_paths: Vec<String> = self
            .drawing_rels
            .get(&drawing_idx)
            .into_iter()
            .flat_map(|rels| &rels.relationships)
            .filter(|relationship| relationship.rel_type == rel_types::IMAGE)
            .map(|relationship| {
                crate::workbook_paths::resolve_relationship_target(
                    &drawing_path,
                    &relationship.target,
                )
            })
            .collect();

        self.drawings.remove(drawing_idx);
        self.drawing_rels.remove(&drawing_idx);
        self.remove_graph_part(&drawing_path);
        self.remove_graph_part(&relationship_part_path(&drawing_path));
        self.drawing_rels = self
            .drawing_rels
            .drain()
            .map(|(idx, rels)| {
                if idx > drawing_idx {
                    (idx - 1, rels)
                } else {
                    (idx, rels)
                }
            })
            .collect();
        for existing_idx in self.worksheet_drawings.values_mut() {
            if *existing_idx > drawing_idx {
                *existing_idx -= 1;
            }
        }
        self.content_types
            .overrides
            .retain(|entry| entry.part_name != format!("/{drawing_path}"));

        for chart_path in removed_chart_paths {
            let still_referenced = self.drawing_rels.iter().any(|(idx, rels)| {
                let owner_path = &self.drawings[*idx].0;
                rels.relationships.iter().any(|relationship| {
                    relationship.rel_type == rel_types::CHART
                        && crate::workbook_paths::resolve_relationship_target(
                            owner_path,
                            &relationship.target,
                        ) == chart_path
                })
            });
            if !still_referenced {
                self.charts.retain(|(path, _)| path != &chart_path);
                self.raw_charts.retain(|(path, _)| path != &chart_path);
                self.remove_graph_part(&chart_path);
                self.content_types
                    .overrides
                    .retain(|entry| entry.part_name != format!("/{chart_path}"));
            }
        }
        for image_path in removed_image_paths {
            let still_referenced = self.drawing_rels.iter().any(|(idx, rels)| {
                let owner_path = &self.drawings[*idx].0;
                rels.relationships.iter().any(|relationship| {
                    relationship.rel_type == rel_types::IMAGE
                        && crate::workbook_paths::resolve_relationship_target(
                            owner_path,
                            &relationship.target,
                        ) == image_path
                })
            });
            if !still_referenced {
                self.images.retain(|(path, _)| path != &image_path);
            }
        }
    }

    /// Debug assertion that all per-sheet parallel vectors have the same
    /// length as `worksheets`. Catching desync early prevents silent data
    /// corruption from mismatched indices.
    fn assert_parallel_vecs_in_sync(&self) {
        let n = self.worksheets.len();
        debug_assert_eq!(self.sheet_comments.len(), n, "sheet_comments desync");
        debug_assert_eq!(self.sheet_sparklines.len(), n, "sheet_sparklines desync");
        debug_assert_eq!(self.sheet_vml.len(), n, "sheet_vml desync");
        debug_assert_eq!(self.raw_sheet_xml.len(), n, "raw_sheet_xml desync");
        debug_assert_eq!(self.sheet_dirty.len(), n, "sheet_dirty desync");
        debug_assert_eq!(
            self.sheet_threaded_comments.len(),
            n,
            "sheet_threaded_comments desync"
        );
        debug_assert_eq!(
            self.sheet_form_controls.len(),
            n,
            "sheet_form_controls desync"
        );
    }

    /// Rename a sheet.
    pub fn set_sheet_name(&mut self, old_name: &str, new_name: &str) -> Result<()> {
        crate::sheet::rename_sheet(
            &mut self.workbook_xml,
            &mut self.worksheets,
            old_name,
            new_name,
        )?;
        self.rebuild_sheet_index();
        Ok(())
    }

    /// Copy a sheet, returning the 0-based index of the new copy.
    pub fn copy_sheet(&mut self, source: &str, target: &str) -> Result<usize> {
        let src_idx = self.sheet_index(source)?;
        crate::sheet::validate_sheet_name(target)?;
        if self.worksheets.iter().any(|(name, _)| name == target) {
            return Err(Error::SheetAlreadyExists {
                name: target.to_string(),
            });
        }

        let cloned_streamed = self
            .streamed_sheets
            .get(&src_idx)
            .map(crate::stream::StreamedSheetData::try_clone)
            .transpose()?;
        let cloned_worksheet = if let Some(worksheet) = self.worksheets[src_idx].1.get() {
            worksheet.clone()
        } else if let Some(raw) = self.raw_sheet_xml[src_idx].as_ref() {
            let xml =
                std::str::from_utf8(raw).map_err(|error| Error::XmlParse(error.to_string()))?;
            quick_xml::de::from_str(xml)
                .map_err(|error| Error::XmlDeserialize(error.to_string()))?
        } else {
            WorksheetXml::default()
        };

        use crate::workbook::aux::AuxCategory;
        let source_rels = self
            .worksheet_rels
            .get(&src_idx)
            .cloned()
            .unwrap_or_else(crate::workbook_paths::default_relationships);
        validate_worksheet_relationship_ids(&cloned_worksheet, &source_rels)?;
        let has_rel_type = |rel_type: &str| {
            source_rels
                .relationships
                .iter()
                .any(|relationship| relationship.rel_type == rel_type)
        };
        let source_has_drawing = has_rel_type(rel_types::DRAWING);
        let source_has_comments = has_rel_type(rel_types::COMMENTS);
        let source_has_vml = has_rel_type(rel_types::VML_DRAWING);
        let source_has_tables = has_rel_type(rel_types::TABLE);
        let source_has_threaded =
            has_rel_type(sheetkit_xml::threaded_comment::REL_TYPE_THREADED_COMMENT);
        let occupied_aux_paths = self.occupied_part_paths();
        let deferred_source_data = (source_has_drawing
            && [
                AuxCategory::Drawings,
                AuxCategory::DrawingRels,
                AuxCategory::Charts,
                AuxCategory::Images,
            ]
            .into_iter()
            .any(|category| self.deferred_parts.has_category(category)))
            || (source_has_comments && self.deferred_parts.has_category(AuxCategory::Comments))
            || (source_has_vml && self.deferred_parts.has_category(AuxCategory::Vml))
            || (source_has_tables && self.deferred_parts.has_category(AuxCategory::Tables))
            || (source_has_threaded
                && self
                    .deferred_parts
                    .has_category(AuxCategory::ThreadedComments));
        if deferred_source_data {
            return Err(Error::InvalidArgument(
                "cannot copy a sheet with deferred relationship parts; hydrate them first".into(),
            ));
        }

        for relationship in &source_rels.relationships {
            let supported = relationship.rel_type == rel_types::HYPERLINK
                || relationship.rel_type == rel_types::DRAWING
                || relationship.rel_type == rel_types::COMMENTS
                || relationship.rel_type == rel_types::VML_DRAWING
                || relationship.rel_type == rel_types::TABLE
                || relationship.rel_type
                    == sheetkit_xml::threaded_comment::REL_TYPE_THREADED_COMMENT;
            if !supported {
                return Err(Error::InvalidArgument(format!(
                    "cannot copy unsupported worksheet relationship type '{}'",
                    relationship.rel_type
                )));
            }
            if relationship.rel_type == rel_types::HYPERLINK
                && relationship.target_mode.as_deref() != Some("External")
            {
                return Err(Error::InvalidArgument(
                    "worksheet hyperlink relationships must be external".into(),
                ));
            }
        }

        let source_sheet_path = self.sheet_part_path(src_idx);
        let managed_rel_count = |rel_type: &str| {
            source_rels
                .relationships
                .iter()
                .filter(|relationship| relationship.rel_type == rel_type)
                .count()
        };
        let content_type_matches = |path: &str, content_type: &str| {
            self.content_types.overrides.iter().any(|entry| {
                entry.part_name.trim_start_matches('/') == path
                    && entry.content_type == content_type
            })
        };
        if source_has_comments {
            let relationship = source_rels
                .relationships
                .iter()
                .find(|relationship| relationship.rel_type == rel_types::COMMENTS)
                .unwrap();
            let path = crate::workbook_paths::resolve_relationship_target(
                &source_sheet_path,
                &relationship.target,
            );
            if managed_rel_count(rel_types::COMMENTS) != 1
                || self.sheet_comments[src_idx].is_none()
                || !content_type_matches(&path, mime_types::COMMENTS)
            {
                return Err(Error::InvalidArgument(
                    "worksheet comments relationship is unresolved".into(),
                ));
            }
        }
        if source_has_vml {
            let relationship = source_rels
                .relationships
                .iter()
                .find(|relationship| relationship.rel_type == rel_types::VML_DRAWING)
                .unwrap();
            let path = crate::workbook_paths::resolve_relationship_target(
                &source_sheet_path,
                &relationship.target,
            );
            let has_managed_vml = self.sheet_vml[src_idx].is_some()
                || !self.sheet_form_controls[src_idx].is_empty()
                || self.sheet_comments[src_idx].is_some();
            if managed_rel_count(rel_types::VML_DRAWING) != 1
                || !has_managed_vml
                || !occupied_aux_paths.contains(&path)
            {
                return Err(Error::InvalidArgument(
                    "worksheet VML relationship is unresolved".into(),
                ));
            }
        }
        if source_has_tables {
            let table_relationships: Vec<&Relationship> = source_rels
                .relationships
                .iter()
                .filter(|relationship| relationship.rel_type == rel_types::TABLE)
                .collect();
            let source_tables: Vec<&str> = self
                .tables
                .iter()
                .filter(|(_, _, sheet_idx)| *sheet_idx == src_idx)
                .map(|(path, _, _)| path.as_str())
                .collect();
            let all_resolved = table_relationships.iter().all(|relationship| {
                let path = crate::workbook_paths::resolve_relationship_target(
                    &source_sheet_path,
                    &relationship.target,
                );
                source_tables.contains(&path.as_str())
                    && content_type_matches(&path, mime_types::TABLE)
            });
            if !all_resolved || table_relationships.len() != source_tables.len() {
                return Err(Error::InvalidArgument(
                    "worksheet table relationship is unresolved".into(),
                ));
            }
        }
        if source_has_threaded {
            let relationship = source_rels
                .relationships
                .iter()
                .find(|relationship| {
                    relationship.rel_type
                        == sheetkit_xml::threaded_comment::REL_TYPE_THREADED_COMMENT
                })
                .unwrap();
            let path = crate::workbook_paths::resolve_relationship_target(
                &source_sheet_path,
                &relationship.target,
            );
            if managed_rel_count(sheetkit_xml::threaded_comment::REL_TYPE_THREADED_COMMENT) != 1
                || self.sheet_threaded_comments[src_idx].is_none()
                || !content_type_matches(
                    &path,
                    sheetkit_xml::threaded_comment::THREADED_COMMENTS_CONTENT_TYPE,
                )
            {
                return Err(Error::InvalidArgument(
                    "worksheet threaded-comment relationship is unresolved".into(),
                ));
            }
        }
        if source_has_drawing {
            let drawing_relationships: Vec<&Relationship> = source_rels
                .relationships
                .iter()
                .filter(|relationship| relationship.rel_type == rel_types::DRAWING)
                .collect();
            let mapped_drawing_path = self
                .worksheet_drawings
                .get(&src_idx)
                .and_then(|drawing_idx| self.drawings.get(*drawing_idx))
                .map(|(path, _)| path.as_str());
            let resolved_path = drawing_relationships.first().map(|relationship| {
                crate::workbook_paths::resolve_relationship_target(
                    &source_sheet_path,
                    &relationship.target,
                )
            });
            if drawing_relationships.len() != 1
                || resolved_path.as_deref() != mapped_drawing_path
                || !resolved_path
                    .as_deref()
                    .is_some_and(|path| content_type_matches(path, mime_types::DRAWING))
            {
                return Err(Error::InvalidArgument(
                    "worksheet drawing relationship is unresolved".into(),
                ));
            }
        }

        let mut cloned_drawing = None;
        if source_has_drawing {
            let drawing_idx = *self.worksheet_drawings.get(&src_idx).ok_or_else(|| {
                Error::InvalidArgument("worksheet drawing relationship is unresolved".into())
            })?;
            let (source_drawing_path, source_drawing) =
                self.drawings.get(drawing_idx).ok_or_else(|| {
                    Error::InvalidArgument("worksheet drawing part is missing".into())
                })?;
            let new_drawing_path =
                next_numbered_part_path("xl/drawings/drawing", ".xml", &occupied_aux_paths);
            let mut new_drawing_rels =
                self.drawing_rels
                    .get(&drawing_idx)
                    .cloned()
                    .ok_or_else(|| {
                        Error::InvalidArgument("worksheet drawing relationships are missing".into())
                    })?;
            validate_drawing_relationship_ids(source_drawing, &new_drawing_rels)?;
            let mut cloned_charts = Vec::new();
            let mut cloned_images = Vec::new();
            let mut reserved_chart_paths = occupied_aux_paths.clone();
            let mut reserved_image_paths = occupied_aux_paths.clone();
            for relationship in &mut new_drawing_rels.relationships {
                if relationship.rel_type == rel_types::IMAGE {
                    let source_image_path = crate::workbook_paths::resolve_relationship_target(
                        source_drawing_path,
                        &relationship.target,
                    );
                    let (_, image_bytes) = self
                        .images
                        .iter()
                        .find(|(path, _)| *path == source_image_path)
                        .ok_or_else(|| {
                            Error::InvalidArgument(format!(
                                "drawing image target '{source_image_path}' is missing"
                            ))
                        })?;
                    let extension = source_image_path
                        .rsplit_once('.')
                        .map(|(_, extension)| extension)
                        .ok_or_else(|| {
                            Error::InvalidArgument(format!(
                                "drawing image target '{source_image_path}' has no extension"
                            ))
                        })?;
                    let new_image_path = next_numbered_part_path(
                        "xl/media/image",
                        &format!(".{extension}"),
                        &reserved_image_paths,
                    );
                    reserved_image_paths.insert(new_image_path.clone());
                    cloned_images.push((new_image_path.clone(), image_bytes.clone()));
                    relationship.target = crate::workbook_paths::relative_relationship_target(
                        &new_drawing_path,
                        &new_image_path,
                    );
                    continue;
                }
                if relationship.rel_type != rel_types::CHART {
                    return Err(Error::InvalidArgument(format!(
                        "cannot copy unsupported drawing relationship type '{}'",
                        relationship.rel_type
                    )));
                }
                let source_chart_path = crate::workbook_paths::resolve_relationship_target(
                    source_drawing_path,
                    &relationship.target,
                );
                let new_chart_path =
                    next_numbered_part_path("xl/charts/chart", ".xml", &reserved_chart_paths);
                if let Some((_, chart)) = self
                    .charts
                    .iter()
                    .find(|(path, _)| *path == source_chart_path)
                {
                    cloned_charts.push((new_chart_path.clone(), Some(chart.clone()), None));
                } else if let Some((_, raw)) = self
                    .raw_charts
                    .iter()
                    .find(|(path, _)| *path == source_chart_path)
                {
                    cloned_charts.push((new_chart_path.clone(), None, Some(raw.clone())));
                } else {
                    return Err(Error::InvalidArgument(format!(
                        "drawing chart target '{source_chart_path}' is missing"
                    )));
                }
                reserved_chart_paths.insert(new_chart_path.clone());
                relationship.target = crate::workbook_paths::relative_relationship_target(
                    &new_drawing_path,
                    &new_chart_path,
                );
            }
            cloned_drawing = Some((
                new_drawing_path,
                source_drawing.clone(),
                new_drawing_rels,
                cloned_charts,
                cloned_images,
            ));
        }

        let max_table_id = self
            .tables
            .iter()
            .map(|(_, table, _)| table.id)
            .max()
            .unwrap_or(0);
        let mut reserved_table_paths = occupied_aux_paths.clone();
        let mut cloned_tables = Vec::new();
        for (offset, (_, source_table, _)) in self
            .tables
            .iter()
            .filter(|(_, _, sheet_idx)| *sheet_idx == src_idx)
            .enumerate()
        {
            let new_path =
                next_numbered_part_path("xl/tables/table", ".xml", &reserved_table_paths);
            reserved_table_paths.insert(new_path.clone());
            let mut table = source_table.clone();
            let id_increment = u32::try_from(offset)
                .ok()
                .and_then(|offset| offset.checked_add(1))
                .ok_or_else(|| Error::InvalidArgument("table ID allocation overflow".into()))?;
            table.id = max_table_id
                .checked_add(id_increment)
                .ok_or_else(|| Error::InvalidArgument("table ID allocation overflow".into()))?;
            let base_name = format!("{}_Copy", table.name);
            let mut name = base_name.clone();
            let mut suffix = 2usize;
            while self
                .tables
                .iter()
                .any(|(_, existing, _)| existing.name == name || existing.display_name == name)
                || cloned_tables.iter().any(
                    |(_, existing): &(String, sheetkit_xml::table::TableXml)| {
                        existing.name == name || existing.display_name == name
                    },
                )
            {
                name = format!("{base_name}{suffix}");
                suffix += 1;
            }
            table.name = name.clone();
            table.display_name = name;
            cloned_tables.push((new_path, table));
        }

        let idx = crate::sheet::add_sheet_with_occupied_paths(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            target,
            cloned_worksheet,
            occupied_aux_paths,
        )?;
        self.sheet_comments
            .push(self.sheet_comments[src_idx].clone());
        self.sheet_sparklines
            .push(self.sheet_sparklines[src_idx].clone());
        self.sheet_vml.push(self.sheet_vml[src_idx].clone());
        self.raw_sheet_xml.push(None);
        self.sheet_dirty.push(true);
        let mut occupied_thread_ids: std::collections::HashSet<String> = self
            .sheet_threaded_comments
            .iter()
            .flatten()
            .flat_map(|threaded| threaded.comments.iter().map(|comment| comment.id.clone()))
            .collect();
        self.sheet_threaded_comments
            .push(
                self.sheet_threaded_comments[src_idx]
                    .as_ref()
                    .map(|threaded| {
                        clone_threaded_comments_with_new_ids(threaded, &mut occupied_thread_ids)
                    }),
            );
        self.sheet_form_controls
            .push(self.sheet_form_controls[src_idx].clone());
        if let Some(cloned) = cloned_streamed {
            self.streamed_sheets.insert(idx, cloned);
        }

        let mut copied_rels = source_rels;
        copied_rels.relationships.retain(|relationship| {
            relationship.rel_type == rel_types::HYPERLINK
                || relationship.rel_type == rel_types::DRAWING
        });
        if let Some((drawing_path, drawing, drawing_rels, cloned_charts, cloned_images)) =
            cloned_drawing
        {
            let drawing_idx = self.drawings.len();
            for relationship in &mut copied_rels.relationships {
                if relationship.rel_type == rel_types::DRAWING {
                    relationship.target = crate::workbook_paths::relative_relationship_target(
                        &self.sheet_part_path(idx),
                        &drawing_path,
                    );
                }
            }
            self.drawings.push((drawing_path.clone(), drawing));
            self.mark_graph_part_dirty(&drawing_path);
            self.worksheet_drawings.insert(idx, drawing_idx);
            self.drawing_rels.insert(drawing_idx, drawing_rels);
            self.mark_drawing_relationships_dirty(drawing_idx);
            self.content_types.overrides.push(ContentTypeOverride {
                part_name: format!("/{drawing_path}"),
                content_type: mime_types::DRAWING.to_string(),
            });
            for (path, typed, raw) in cloned_charts {
                self.content_types.overrides.push(ContentTypeOverride {
                    part_name: format!("/{path}"),
                    content_type: mime_types::CHART.to_string(),
                });
                if let Some(chart) = typed {
                    self.mark_graph_part_dirty(&path);
                    self.charts.push((path, chart));
                } else if let Some(bytes) = raw {
                    self.mark_graph_part_dirty(&path);
                    self.raw_charts.push((path, bytes));
                }
            }
            self.images.extend(cloned_images);
        }
        if !copied_rels.relationships.is_empty() {
            self.worksheet_rels.insert(idx, copied_rels);
        }
        for (path, table) in cloned_tables {
            self.tables.push((path, table, idx));
        }
        self.rebuild_sheet_index();
        Ok(idx)
    }

    /// Get a sheet's 0-based index by name. Returns `None` if not found.
    pub fn get_sheet_index(&self, name: &str) -> Option<usize> {
        crate::sheet::find_sheet_index(&self.worksheets, name)
    }

    /// Get the name of the active sheet.
    pub fn get_active_sheet(&self) -> &str {
        let idx = crate::sheet::active_sheet_index(&self.workbook_xml);
        self.worksheets
            .get(idx)
            .map(|(n, _)| n.as_str())
            .or_else(|| self.worksheets.first().map(|(n, _)| n.as_str()))
            .unwrap_or("")
    }

    /// Set the active sheet by name.
    pub fn set_active_sheet(&mut self, name: &str) -> Result<()> {
        let idx = crate::sheet::find_sheet_index(&self.worksheets, name).ok_or_else(|| {
            Error::SheetNotFound {
                name: name.to_string(),
            }
        })?;
        crate::sheet::set_active_sheet_index(&mut self.workbook_xml, idx as u32);
        Ok(())
    }

    /// Create a [`StreamWriter`](crate::stream::StreamWriter) for a new sheet.
    ///
    /// The sheet will be added to the workbook when the StreamWriter is applied
    /// via [`apply_stream_writer`](Self::apply_stream_writer).
    pub fn new_stream_writer(&self, sheet_name: &str) -> Result<crate::stream::StreamWriter> {
        crate::sheet::validate_sheet_name(sheet_name)?;
        if self.worksheets.iter().any(|(n, _)| n == sheet_name) {
            return Err(Error::SheetAlreadyExists {
                name: sheet_name.to_string(),
            });
        }
        Ok(crate::stream::StreamWriter::new(sheet_name))
    }

    /// Apply a completed [`StreamWriter`](crate::stream::StreamWriter) to the
    /// workbook, adding it as a new sheet.
    ///
    /// The streamed row data stays on disk (in a temp file) and is written
    /// directly to the ZIP archive during save, keeping memory usage constant
    /// regardless of the number of rows.
    ///
    /// **Note:** Cell values in streamed sheets cannot be read back via
    /// [`get_cell_value`](Self::get_cell_value) before saving. Save the
    /// workbook and reopen it to read the data.
    ///
    /// Returns the 0-based index of the new sheet.
    pub fn apply_stream_writer(&mut self, writer: crate::stream::StreamWriter) -> Result<usize> {
        let (sheet_name, streamed_data) = writer.into_streamed_data()?;

        // Add an empty WorksheetXml placeholder for sheet management
        // (sheet names, indices, metadata). The actual data lives in the
        // temp file and is streamed to the ZIP during save.
        let occupied_paths = self.occupied_part_paths();
        let idx = crate::sheet::add_sheet_with_occupied_paths(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            &sheet_name,
            WorksheetXml::default(),
            occupied_paths,
        )?;
        self.append_empty_sheet_state();

        // Store the streamed data for use during save.
        self.streamed_sheets.insert(idx, streamed_data);

        self.rebuild_sheet_index();
        Ok(idx)
    }

    /// Insert `count` empty rows starting at `start_row` in the named sheet.
    pub fn insert_rows(&mut self, sheet: &str, start_row: u32, count: u32) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        if count == 0 {
            return Ok(());
        }
        self.preflight_structural_target(sheet_idx, |ws| {
            crate::row::insert_rows(ws, start_row, count)
        })?;
        self.prepare_reference_shift(sheet_idx, |col, row| {
            if row >= start_row {
                (col, row.saturating_add(count))
            } else {
                (col, row)
            }
        })?;
        {
            let ws = self.worksheet_mut_by_index(sheet_idx)?;
            crate::row::insert_rows(ws, start_row, count)?;
        }
        self.apply_reference_shift_for_sheet(sheet_idx, |col, row| {
            if row >= start_row {
                (col, row.saturating_add(count))
            } else {
                (col, row)
            }
        })
    }

    /// Remove a single row from the named sheet, shifting rows below it up.
    pub fn remove_row(&mut self, sheet: &str, row: u32) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        self.preflight_structural_target(sheet_idx, |ws| crate::row::remove_row(ws, row))?;
        self.prepare_reference_shift(sheet_idx, |col, referenced_row| {
            if referenced_row > row {
                (col, referenced_row - 1)
            } else {
                (col, referenced_row)
            }
        })?;
        {
            let ws = self.worksheet_mut_by_index(sheet_idx)?;
            crate::row::remove_row(ws, row)?;
        }
        self.apply_reference_shift_for_sheet(sheet_idx, |col, r| {
            if r > row {
                (col, r - 1)
            } else {
                (col, r)
            }
        })
    }

    /// Duplicate a row, inserting the copy directly below.
    pub fn duplicate_row(&mut self, sheet: &str, row: u32) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let target_row = row.checked_add(1).ok_or(Error::InvalidRowNumber(row))?;
        self.preflight_structural_target(sheet_idx, |ws| crate::row::duplicate_row(ws, row))?;
        self.preflight_duplicate_formula_copy(sheet_idx, row)?;
        self.prepare_reference_shift(sheet_idx, |col, referenced_row| {
            if referenced_row >= target_row {
                (col, referenced_row.saturating_add(1))
            } else {
                (col, referenced_row)
            }
        })?;
        {
            let ws = self.worksheet_mut_by_index(sheet_idx)?;
            crate::row::duplicate_row(ws, row)?;
        }
        let duplicated_formulas = self.take_duplicated_row_formulas(sheet_idx, target_row)?;
        self.apply_reference_shift_for_sheet(sheet_idx, |col, r| {
            if r >= target_row {
                (col, r.saturating_add(1))
            } else {
                (col, r)
            }
        })?;
        self.restore_duplicated_row_formulas(sheet_idx, target_row, duplicated_formulas)
    }

    /// Set the height of a row in points.
    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::row::set_row_height(ws, row, height)
    }

    /// Get the height of a row.
    pub fn get_row_height(&self, sheet: &str, row: u32) -> Result<Option<f64>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::row::get_row_height(ws, row))
    }

    /// Set the visibility of a row.
    pub fn set_row_visible(&mut self, sheet: &str, row: u32, visible: bool) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::row::set_row_visible(ws, row, visible)
    }

    /// Get the visibility of a row. Returns true if visible (not hidden).
    pub fn get_row_visible(&self, sheet: &str, row: u32) -> Result<bool> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::row::get_row_visible(ws, row))
    }

    /// Set the outline level of a row.
    pub fn set_row_outline_level(&mut self, sheet: &str, row: u32, level: u8) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::row::set_row_outline_level(ws, row, level)
    }

    /// Get the outline level of a row. Returns 0 if not set.
    pub fn get_row_outline_level(&self, sheet: &str, row: u32) -> Result<u8> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::row::get_row_outline_level(ws, row))
    }

    /// Set the style for an entire row.
    ///
    /// The `style_id` must be a valid index in cellXfs (returned by `add_style`).
    pub fn set_row_style(&mut self, sheet: &str, row: u32, style_id: u32) -> Result<()> {
        if style_id as usize >= self.stylesheet.cell_xfs.xfs.len() {
            return Err(Error::StyleNotFound { id: style_id });
        }
        let ws = self.worksheet_mut(sheet)?;
        crate::row::set_row_style(ws, row, style_id)
    }

    /// Get the style ID for a row. Returns 0 (default) if not set.
    pub fn get_row_style(&self, sheet: &str, row: u32) -> Result<u32> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::row::get_row_style(ws, row))
    }

    /// Get all rows with their data from a sheet.
    ///
    /// Returns a Vec of `(row_number, Vec<(column_number, CellValue)>)` tuples.
    /// Column numbers are 1-based (A=1, B=2, ...). Only rows that contain at
    /// least one cell are included (sparse).
    #[allow(clippy::type_complexity)]
    pub fn get_rows(&self, sheet: &str) -> Result<Vec<(u32, Vec<(u32, CellValue)>)>> {
        let ws = self.worksheet_ref(sheet)?;
        let style_is_date = self.computed_style_is_date();
        crate::row::get_rows(ws, &self.sst_runtime, &style_is_date)
    }

    /// Get all columns with their data from a sheet.
    ///
    /// Returns a Vec of `(column_name, Vec<(row_number, CellValue)>)` tuples.
    /// Only columns that have data are included (sparse).
    #[allow(clippy::type_complexity)]
    pub fn get_cols(&self, sheet: &str) -> Result<Vec<(String, Vec<(u32, CellValue)>)>> {
        let ws = self.worksheet_ref(sheet)?;
        let style_is_date = self.computed_style_is_date();
        crate::col::get_cols(ws, &self.sst_runtime, &style_is_date)
    }

    /// Set the width of a column.
    pub fn set_col_width(&mut self, sheet: &str, col: &str, width: f64) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::col::set_col_width(ws, col, width)
    }

    /// Get the width of a column.
    pub fn get_col_width(&self, sheet: &str, col: &str) -> Result<Option<f64>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::col::get_col_width(ws, col))
    }

    /// Set the visibility of a column.
    pub fn set_col_visible(&mut self, sheet: &str, col: &str, visible: bool) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::col::set_col_visible(ws, col, visible)
    }

    /// Get the visibility of a column. Returns true if visible (not hidden).
    pub fn get_col_visible(&self, sheet: &str, col: &str) -> Result<bool> {
        let ws = self.worksheet_ref(sheet)?;
        crate::col::get_col_visible(ws, col)
    }

    /// Set the outline level of a column.
    pub fn set_col_outline_level(&mut self, sheet: &str, col: &str, level: u8) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::col::set_col_outline_level(ws, col, level)
    }

    /// Get the outline level of a column. Returns 0 if not set.
    pub fn get_col_outline_level(&self, sheet: &str, col: &str) -> Result<u8> {
        let ws = self.worksheet_ref(sheet)?;
        crate::col::get_col_outline_level(ws, col)
    }

    /// Set the style for an entire column.
    ///
    /// The `style_id` must be a valid index in cellXfs (returned by `add_style`).
    pub fn set_col_style(&mut self, sheet: &str, col: &str, style_id: u32) -> Result<()> {
        if style_id as usize >= self.stylesheet.cell_xfs.xfs.len() {
            return Err(Error::StyleNotFound { id: style_id });
        }
        let ws = self.worksheet_mut(sheet)?;
        crate::col::set_col_style(ws, col, style_id)
    }

    /// Get the style ID for a column. Returns 0 (default) if not set.
    pub fn get_col_style(&self, sheet: &str, col: &str) -> Result<u32> {
        let ws = self.worksheet_ref(sheet)?;
        crate::col::get_col_style(ws, col)
    }

    /// Insert `count` columns starting at `col` in the named sheet.
    pub fn insert_cols(&mut self, sheet: &str, col: &str, count: u32) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let start_col = column_name_to_number(col)?;
        if count == 0 {
            return Ok(());
        }
        self.preflight_structural_target(sheet_idx, |ws| crate::col::insert_cols(ws, col, count))?;
        self.prepare_reference_shift(sheet_idx, |c, row| {
            if c >= start_col {
                (c.saturating_add(count), row)
            } else {
                (c, row)
            }
        })?;
        {
            let ws = self.worksheet_mut_by_index(sheet_idx)?;
            crate::col::insert_cols(ws, col, count)?;
        }
        self.apply_reference_shift_for_sheet(sheet_idx, |c, row| {
            if c >= start_col {
                (c.saturating_add(count), row)
            } else {
                (c, row)
            }
        })
    }

    /// Remove a single column from the named sheet.
    pub fn remove_col(&mut self, sheet: &str, col: &str) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let col_num = column_name_to_number(col)?;
        self.preflight_structural_target(sheet_idx, |ws| crate::col::remove_col(ws, col))?;
        self.prepare_reference_shift(
            sheet_idx,
            |c, row| {
                if c > col_num {
                    (c - 1, row)
                } else {
                    (c, row)
                }
            },
        )?;
        {
            let ws = self.worksheet_mut_by_index(sheet_idx)?;
            crate::col::remove_col(ws, col)?;
        }
        self.apply_reference_shift_for_sheet(sheet_idx, |c, row| {
            if c > col_num {
                (c - 1, row)
            } else {
                (c, row)
            }
        })
    }

    /// Reindex per-sheet maps after deleting a sheet.
    pub(crate) fn reindex_sheet_maps_after_delete(&mut self, removed_idx: usize) {
        self.worksheet_rels = self
            .worksheet_rels
            .iter()
            .filter_map(|(idx, rels)| {
                if *idx == removed_idx {
                    None
                } else if *idx > removed_idx {
                    Some((idx - 1, rels.clone()))
                } else {
                    Some((*idx, rels.clone()))
                }
            })
            .collect();

        self.worksheet_drawings = self
            .worksheet_drawings
            .iter()
            .filter_map(|(idx, drawing_idx)| {
                if *idx == removed_idx {
                    None
                } else if *idx > removed_idx {
                    Some((idx - 1, *drawing_idx))
                } else {
                    Some((*idx, *drawing_idx))
                }
            })
            .collect();
    }

    /// Prepare every reference-bearing worksheet part before a structural edit.
    ///
    /// A row-limited or streamed workbook cannot safely update references
    /// outside the materialized window, so reject it before any mutation.
    fn prepare_reference_shift<F>(&mut self, target_sheet_idx: usize, shift_cell: F) -> Result<()>
    where
        F: Fn(u32, u32) -> (u32, u32) + Copy,
    {
        if self.sheet_rows_limit.is_some() {
            return Err(Error::InvalidArgument(
                "cannot structurally edit a workbook opened with sheet_rows".into(),
            ));
        }
        if !self.streamed_sheets.is_empty() {
            return Err(Error::InvalidArgument(
                "cannot structurally edit a workbook containing streamed sheets".into(),
            ));
        }
        if !self.raw_charts.is_empty() {
            return Err(Error::InvalidArgument(
                "cannot structurally edit a workbook with raw charts".into(),
            ));
        }
        self.ensure_drawing_relationships_hydratable(target_sheet_idx)?;

        // Validate deferred XML without consuming passthrough bytes. A failed
        // edit must not turn an untouched lazy workbook into a dirty one.
        self.validate_deferred_drawing_and_chart_parts(target_sheet_idx, shift_cell)?;
        self.validate_reference_shift(target_sheet_idx, shift_cell)?;
        for idx in 0..self.worksheets.len() {
            self.ensure_hydrated(idx)?;
        }
        self.hydrate_drawings();
        self.hydrate_tables();
        self.validate_typed_chart_ownership()?;
        Ok(())
    }

    fn validate_deferred_drawing_and_chart_parts<F>(
        &self,
        target_sheet_idx: usize,
        shift_cell: F,
    ) -> Result<()>
    where
        F: Fn(u32, u32) -> (u32, u32) + Copy,
    {
        use crate::workbook::aux::AuxCategory;

        let target_sheet_name_lowercase = self.worksheets[target_sheet_idx].0.to_lowercase();
        let validate = |text: &str, owner_sheet_idx: usize| {
            shift_references_for_owner(
                text,
                owner_sheet_idx,
                target_sheet_idx,
                &target_sheet_name_lowercase,
                shift_cell,
            )
            .map(|_| ())
        };
        if let Some(entries) = self.deferred_parts.entries(AuxCategory::DrawingRels) {
            for (_, bytes) in entries {
                quick_xml::de::from_str::<Relationships>(&String::from_utf8_lossy(bytes))
                    .map_err(|error| Error::XmlDeserialize(error.to_string()))?;
            }
        }
        if let Some(entries) = self.deferred_parts.entries(AuxCategory::Drawings) {
            for (path, bytes) in entries {
                let drawing = quick_xml::de::from_str::<WsDr>(&String::from_utf8_lossy(bytes))
                    .map_err(|error| Error::XmlDeserialize(error.to_string()))?;
                let owner = self
                    .worksheet_rels
                    .iter()
                    .find_map(|(sheet_idx, relationships)| {
                        relationships
                            .relationships
                            .iter()
                            .any(|relationship| {
                                relationship.rel_type == rel_types::DRAWING
                                    && resolve_relationship_target(
                                        &self.sheet_part_path(*sheet_idx),
                                        &relationship.target,
                                    ) == *path
                            })
                            .then_some(*sheet_idx)
                    });
                if owner == Some(target_sheet_idx) {
                    for anchor in &drawing.one_cell_anchors {
                        validate_anchor_shift(anchor.from.col, anchor.from.row, shift_cell)?;
                    }
                    for anchor in &drawing.two_cell_anchors {
                        validate_anchor_shift(anchor.from.col, anchor.from.row, shift_cell)?;
                        validate_anchor_shift(anchor.to.col, anchor.to.row, shift_cell)?;
                    }
                }
            }
        }
        if let Some(entries) = self.deferred_parts.entries(AuxCategory::Charts) {
            for (path, bytes) in entries {
                let owner_sheet_idx = self.deferred_chart_owner_sheet_idx(path)?;
                let mut chart =
                    quick_xml::de::from_str::<ChartSpace>(&String::from_utf8_lossy(bytes))
                        .map_err(|error| Error::XmlDeserialize(error.to_string()))?;
                visit_chart_references(&mut chart, |formula| validate(formula, owner_sheet_idx))?;
            }
        }
        Ok(())
    }

    fn deferred_chart_owner_sheet_idx(&self, chart_path: &str) -> Result<usize> {
        self.worksheet_rels
            .iter()
            .find_map(|(sheet_idx, worksheet_rels)| {
                worksheet_rels.relationships.iter().find_map(|relationship| {
                    (relationship.rel_type == rel_types::DRAWING).then(|| {
                        let drawing_path = resolve_relationship_target(
                            &self.sheet_part_path(*sheet_idx),
                            &relationship.target,
                        );
                        let has_drawing = self
                            .drawings
                            .iter()
                            .any(|(path, _)| path == &drawing_path)
                            || self
                                .deferred_parts
                                .get_path(
                                    crate::workbook::aux::AuxCategory::Drawings,
                                    &drawing_path,
                                )
                                .is_some();
                        if !has_drawing {
                            return None;
                        }
                        let drawing_rels_path = relationship_part_path(&drawing_path);
                        let rels = self
                            .drawing_rels
                            .iter()
                            .find_map(|(drawing_idx, rels)| {
                                (self.drawings.get(*drawing_idx).is_some_and(|(path, _)| {
                                    path == &drawing_path
                                }))
                                .then_some(rels)
                            })
                            .cloned()
                            .or_else(|| {
                                self.deferred_parts
                                    .get_path(
                                        crate::workbook::aux::AuxCategory::DrawingRels,
                                        &drawing_rels_path,
                                    )
                                    .and_then(|bytes| {
                                        quick_xml::de::from_reader::<_, Relationships>(bytes).ok()
                                    })
                            });
                        rels.and_then(|rels| {
                            rels.relationships
                                .iter()
                                .any(|relationship| {
                                    relationship.rel_type == rel_types::CHART
                                        && resolve_relationship_target(
                                            &drawing_path,
                                            &relationship.target,
                                        ) == chart_path
                                })
                                .then_some(*sheet_idx)
                        })
                    })
                    .flatten()
                })
            })
            .ok_or_else(|| {
                Error::InvalidArgument(format!(
                    "cannot structurally edit a workbook with an unresolved chart owner '{chart_path}'"
                ))
            })
    }

    fn chart_owner_sheet_idx(&self, chart_path: &str) -> Result<usize> {
        self.worksheet_drawings
            .iter()
            .find_map(|(sheet_idx, drawing_idx)| {
                self.drawing_rels
                    .get(drawing_idx)
                    .and_then(|relationships| {
                        relationships.relationships.iter().find_map(|relationship| {
                            (relationship.rel_type == rel_types::CHART
                                && resolve_relationship_target(
                                    &self.drawings[*drawing_idx].0,
                                    &relationship.target,
                                ) == chart_path)
                                .then_some(*sheet_idx)
                        })
                    })
            })
            .ok_or_else(|| {
                Error::InvalidArgument(format!(
                    "cannot structurally edit a workbook with an unresolved chart owner '{chart_path}'"
                ))
            })
    }

    fn validate_typed_chart_ownership(&self) -> Result<()> {
        for (path, _) in &self.charts {
            self.chart_owner_sheet_idx(path)?;
        }
        Ok(())
    }

    fn validate_reference_shift<F>(&self, target_sheet_idx: usize, shift_cell: F) -> Result<()>
    where
        F: Fn(u32, u32) -> (u32, u32) + Copy,
    {
        let target_sheet_name_lowercase = self.worksheets[target_sheet_idx].0.to_lowercase();
        let validate = |text: &str, owner_sheet_idx: usize| {
            shift_references_for_owner(
                text,
                owner_sheet_idx,
                target_sheet_idx,
                &target_sheet_name_lowercase,
                shift_cell,
            )
            .map(|_| ())
        };

        for owner_sheet_idx in 0..self.worksheets.len() {
            let ws = self.preflight_worksheet(owner_sheet_idx)?;
            for row in &ws.sheet_data.rows {
                for cell in &row.cells {
                    if let Some(formula) = &cell.f {
                        if let Some(expr) = &formula.value {
                            validate(expr, owner_sheet_idx)?;
                        }
                        if let Some(reference) = &formula.reference {
                            validate(reference, owner_sheet_idx)?;
                        }
                    }
                }
            }
            if let Some(merges) = &ws.merge_cells {
                for merge in &merges.merge_cells {
                    validate(&merge.reference, owner_sheet_idx)?;
                }
            }
            if let Some(auto_filter) = &ws.auto_filter {
                validate(&auto_filter.reference, owner_sheet_idx)?;
            }
            if let Some(validations) = &ws.data_validations {
                for validation in &validations.data_validations {
                    validate(&validation.sqref, owner_sheet_idx)?;
                    if let Some(formula) = &validation.formula1 {
                        validate(formula, owner_sheet_idx)?;
                    }
                    if let Some(formula) = &validation.formula2 {
                        validate(formula, owner_sheet_idx)?;
                    }
                }
            }
            for formatting in &ws.conditional_formatting {
                validate(&formatting.sqref, owner_sheet_idx)?;
                for rule in &formatting.cf_rules {
                    for formula in &rule.formulas {
                        validate(formula, owner_sheet_idx)?;
                    }
                }
            }
            if let Some(hyperlinks) = &ws.hyperlinks {
                for hyperlink in &hyperlinks.hyperlinks {
                    validate(&hyperlink.reference, owner_sheet_idx)?;
                    if let Some(location) = &hyperlink.location {
                        validate(location, owner_sheet_idx)?;
                    }
                }
            }
            if let Some(views) = &ws.sheet_views {
                for view in &views.sheet_views {
                    if let Some(pane) = &view.pane {
                        if let Some(top_left) = &pane.top_left_cell {
                            validate(top_left, owner_sheet_idx)?;
                        }
                    }
                    for selection in &view.selection {
                        if let Some(active_cell) = &selection.active_cell {
                            validate(active_cell, owner_sheet_idx)?;
                        }
                        if let Some(sqref) = &selection.sqref {
                            validate(sqref, owner_sheet_idx)?;
                        }
                    }
                }
            }
        }
        if let Some(defined_names) = &self.workbook_xml.defined_names {
            for defined_name in &defined_names.defined_names {
                crate::cell_ref_shift::shift_cell_references_with_abs_and_scope(
                    &defined_name.value,
                    defined_name.local_sheet_id == Some(target_sheet_idx as u32),
                    |name| sheet_name_matches(name, &target_sheet_name_lowercase),
                    |col, row, _, _| shift_cell(col, row),
                )?;
            }
        }
        for (owner_sheet_idx, sparklines) in self.sheet_sparklines.iter().enumerate() {
            for sparkline in sparklines {
                validate(&sparkline.data_range, owner_sheet_idx)?;
                if owner_sheet_idx == target_sheet_idx {
                    validate(&sparkline.location, owner_sheet_idx)?;
                }
            }
        }
        for (_, table, owner_sheet_idx) in &self.tables {
            if *owner_sheet_idx == target_sheet_idx {
                validate(&table.reference, target_sheet_idx)?;
                if let Some(auto_filter) = &table.auto_filter {
                    validate(&auto_filter.reference, target_sheet_idx)?;
                }
            }
        }
        for (path, chart) in &self.charts {
            let owner_sheet_idx = self.chart_owner_sheet_idx(path)?;
            let mut chart = chart.clone();
            visit_chart_references(&mut chart, |formula| validate(formula, owner_sheet_idx))?;
        }
        if let Some(&drawing_idx) = self.worksheet_drawings.get(&target_sheet_idx) {
            if let Some((_, drawing)) = self.drawings.get(drawing_idx) {
                for anchor in &drawing.one_cell_anchors {
                    validate_anchor_shift(anchor.from.col, anchor.from.row, shift_cell)?;
                }
                for anchor in &drawing.two_cell_anchors {
                    validate_anchor_shift(anchor.from.col, anchor.from.row, shift_cell)?;
                    validate_anchor_shift(anchor.to.col, anchor.to.row, shift_cell)?;
                }
            }
        }
        if let Some(entries) = self
            .deferred_parts
            .entries(crate::workbook::aux::AuxCategory::Tables)
        {
            for (path, bytes) in entries {
                let table = quick_xml::de::from_str::<sheetkit_xml::table::TableXml>(
                    &String::from_utf8_lossy(bytes),
                )
                .map_err(|error| Error::XmlDeserialize(error.to_string()))?;
                let owner_sheet_idx = self
                    .worksheet_rels
                    .iter()
                    .find_map(|(sheet_idx, rels)| {
                        rels.relationships.iter().find_map(|rel| {
                            (rel.rel_type == rel_types::TABLE
                                && resolve_relationship_target(
                                    &self.sheet_part_path(*sheet_idx),
                                    &rel.target,
                                ) == *path)
                                .then_some(*sheet_idx)
                        })
                    })
                    .unwrap_or(0);
                if owner_sheet_idx != target_sheet_idx {
                    continue;
                }
                validate(&table.reference, target_sheet_idx)?;
                if let Some(auto_filter) = &table.auto_filter {
                    validate(&auto_filter.reference, target_sheet_idx)?;
                }
            }
        }
        Ok(())
    }

    fn preflight_worksheet(&self, sheet_idx: usize) -> Result<WorksheetXml> {
        if !self.sheet_dirty.get(sheet_idx).copied().unwrap_or(true) {
            if let Some(bytes) = self.raw_sheet_xml.get(sheet_idx).and_then(Option::as_deref) {
                return io::deserialize_worksheet_xml(bytes);
            }
        }
        Ok(self.worksheet_ref_by_index(sheet_idx)?.clone())
    }

    fn preflight_structural_target<F>(&self, sheet_idx: usize, edit: F) -> Result<()>
    where
        F: FnOnce(&mut WorksheetXml) -> Result<()>,
    {
        if self.sheet_rows_limit.is_some() {
            return Err(Error::InvalidArgument(
                "cannot structurally edit a workbook opened with sheet_rows".into(),
            ));
        }
        if !self.streamed_sheets.is_empty() {
            return Err(Error::InvalidArgument(
                "cannot structurally edit a workbook containing streamed sheets".into(),
            ));
        }
        let mut worksheet = self.preflight_worksheet(sheet_idx)?;
        edit(&mut worksheet)
    }

    fn preflight_duplicate_formula_copy(&self, sheet_idx: usize, row: u32) -> Result<()> {
        let worksheet = self.preflight_worksheet(sheet_idx)?;
        let source = worksheet
            .sheet_data
            .rows
            .iter()
            .find(|candidate| candidate.r == row)
            .ok_or(Error::InvalidRowNumber(row))?;
        for cell in &source.cells {
            let Some(formula) = cell.f.as_ref().and_then(|formula| formula.value.as_deref()) else {
                continue;
            };
            shift_duplicated_row_formula(formula)?;
        }
        Ok(())
    }

    fn take_duplicated_row_formulas(
        &mut self,
        sheet_idx: usize,
        row: u32,
    ) -> Result<Vec<Option<String>>> {
        let ws = self.worksheet_mut_by_index(sheet_idx)?;
        let Some(duplicated) = ws
            .sheet_data
            .rows
            .iter_mut()
            .find(|candidate| candidate.r == row)
        else {
            return Ok(Vec::new());
        };
        Ok(duplicated
            .cells
            .iter_mut()
            .map(|cell| cell.f.as_mut().and_then(|formula| formula.value.take()))
            .collect())
    }

    fn restore_duplicated_row_formulas(
        &mut self,
        sheet_idx: usize,
        row: u32,
        formulas: Vec<Option<String>>,
    ) -> Result<()> {
        let ws = self.worksheet_mut_by_index(sheet_idx)?;
        let Some(duplicated) = ws
            .sheet_data
            .rows
            .iter_mut()
            .find(|candidate| candidate.r == row)
        else {
            return Ok(());
        };
        for (cell, formula) in duplicated.cells.iter_mut().zip(formulas) {
            if let (Some(cell_formula), Some(formula)) = (&mut cell.f, formula) {
                cell_formula.value = Some(shift_duplicated_row_formula(&formula)?);
            }
        }
        Ok(())
    }

    /// Apply a cell-reference shift transformation to sheet-scoped structures.
    pub(crate) fn apply_reference_shift_for_sheet<F>(
        &mut self,
        sheet_idx: usize,
        shift_cell: F,
    ) -> Result<()>
    where
        F: Fn(u32, u32) -> (u32, u32) + Copy,
    {
        let target_sheet_name_lowercase = self.worksheets[sheet_idx].0.to_lowercase();
        for owner_sheet_idx in 0..self.worksheets.len() {
            let changed = {
                let ws = self.worksheets[owner_sheet_idx]
                    .1
                    .get_mut()
                    .ok_or_else(|| {
                        Error::Internal(
                            "worksheet must be hydrated before shifting references".into(),
                        )
                    })?;
                let before = ws.clone();

                for row in &mut ws.sheet_data.rows {
                    for cell in &mut row.cells {
                        if let Some(formula) = &mut cell.f {
                            if let Some(expr) = &mut formula.value {
                                *expr = shift_references_for_owner(
                                    expr,
                                    owner_sheet_idx,
                                    sheet_idx,
                                    &target_sheet_name_lowercase,
                                    shift_cell,
                                )?;
                            }
                            if let Some(reference) = &mut formula.reference {
                                *reference = shift_references_for_owner(
                                    reference,
                                    owner_sheet_idx,
                                    sheet_idx,
                                    &target_sheet_name_lowercase,
                                    shift_cell,
                                )?;
                            }
                        }
                    }
                }

                if let Some(merges) = &mut ws.merge_cells {
                    for merge in &mut merges.merge_cells {
                        merge.reference = shift_references_for_owner(
                            &merge.reference,
                            owner_sheet_idx,
                            sheet_idx,
                            &target_sheet_name_lowercase,
                            shift_cell,
                        )?;
                    }
                    merges.cached_coords.clear();
                }

                if let Some(auto_filter) = &mut ws.auto_filter {
                    auto_filter.reference = shift_references_for_owner(
                        &auto_filter.reference,
                        owner_sheet_idx,
                        sheet_idx,
                        &target_sheet_name_lowercase,
                        shift_cell,
                    )?;
                }

                if let Some(validations) = &mut ws.data_validations {
                    for validation in &mut validations.data_validations {
                        validation.sqref = shift_references_for_owner(
                            &validation.sqref,
                            owner_sheet_idx,
                            sheet_idx,
                            &target_sheet_name_lowercase,
                            shift_cell,
                        )?;
                        for formula in [&mut validation.formula1, &mut validation.formula2]
                            .into_iter()
                            .flatten()
                        {
                            *formula = shift_references_for_owner(
                                formula,
                                owner_sheet_idx,
                                sheet_idx,
                                &target_sheet_name_lowercase,
                                shift_cell,
                            )?;
                        }
                    }
                }

                for formatting in &mut ws.conditional_formatting {
                    formatting.sqref = shift_references_for_owner(
                        &formatting.sqref,
                        owner_sheet_idx,
                        sheet_idx,
                        &target_sheet_name_lowercase,
                        shift_cell,
                    )?;
                    for rule in &mut formatting.cf_rules {
                        for formula in &mut rule.formulas {
                            *formula = shift_references_for_owner(
                                formula,
                                owner_sheet_idx,
                                sheet_idx,
                                &target_sheet_name_lowercase,
                                shift_cell,
                            )?;
                        }
                    }
                }

                if let Some(hyperlinks) = &mut ws.hyperlinks {
                    for hyperlink in &mut hyperlinks.hyperlinks {
                        hyperlink.reference = shift_references_for_owner(
                            &hyperlink.reference,
                            owner_sheet_idx,
                            sheet_idx,
                            &target_sheet_name_lowercase,
                            shift_cell,
                        )?;
                        if let Some(location) = &mut hyperlink.location {
                            *location = shift_references_for_owner(
                                location,
                                owner_sheet_idx,
                                sheet_idx,
                                &target_sheet_name_lowercase,
                                shift_cell,
                            )?;
                        }
                    }
                }

                if let Some(views) = &mut ws.sheet_views {
                    for view in &mut views.sheet_views {
                        if let Some(pane) = &mut view.pane {
                            if let Some(top_left) = &mut pane.top_left_cell {
                                *top_left = shift_references_for_owner(
                                    top_left,
                                    owner_sheet_idx,
                                    sheet_idx,
                                    &target_sheet_name_lowercase,
                                    shift_cell,
                                )?;
                            }
                        }
                        for selection in &mut view.selection {
                            for reference in [&mut selection.active_cell, &mut selection.sqref]
                                .into_iter()
                                .flatten()
                            {
                                *reference = shift_references_for_owner(
                                    reference,
                                    owner_sheet_idx,
                                    sheet_idx,
                                    &target_sheet_name_lowercase,
                                    shift_cell,
                                )?;
                            }
                        }
                    }
                }

                *ws != before
            };
            if changed {
                self.mark_sheet_dirty(owner_sheet_idx);
            }
        }

        if let Some(defined_names) = &mut self.workbook_xml.defined_names {
            for defined_name in &mut defined_names.defined_names {
                let owner = defined_name.local_sheet_id.map(|id| id as usize);
                defined_name.value =
                    crate::cell_ref_shift::shift_cell_references_with_abs_and_scope(
                        &defined_name.value,
                        owner == Some(sheet_idx),
                        |name| sheet_name_matches(name, &target_sheet_name_lowercase),
                        |col, row, _, _| shift_cell(col, row),
                    )?;
            }
        }

        for (owner_sheet_idx, sparklines) in self.sheet_sparklines.iter_mut().enumerate() {
            for sparkline in sparklines {
                sparkline.data_range = shift_references_for_owner(
                    &sparkline.data_range,
                    owner_sheet_idx,
                    sheet_idx,
                    &target_sheet_name_lowercase,
                    shift_cell,
                )?;
                if owner_sheet_idx == sheet_idx {
                    sparkline.location = shift_references_for_owner(
                        &sparkline.location,
                        owner_sheet_idx,
                        sheet_idx,
                        &target_sheet_name_lowercase,
                        shift_cell,
                    )?;
                }
            }
        }

        let mut shifted_table = false;
        for (_, table, owner_sheet_idx) in &mut self.tables {
            if *owner_sheet_idx != sheet_idx {
                continue;
            }
            shifted_table = true;
            table.reference = shift_references_for_owner(
                &table.reference,
                sheet_idx,
                sheet_idx,
                &target_sheet_name_lowercase,
                shift_cell,
            )?;
            if let Some(auto_filter) = &mut table.auto_filter {
                auto_filter.reference = shift_references_for_owner(
                    &auto_filter.reference,
                    sheet_idx,
                    sheet_idx,
                    &target_sheet_name_lowercase,
                    shift_cell,
                )?;
            }
        }
        if shifted_table {
            self.deferred_parts
                .mark_dirty(crate::workbook::aux::AuxCategory::Tables);
        }

        let mut dirty_chart_paths = Vec::new();
        for chart_idx in 0..self.charts.len() {
            let mut changed = false;
            let owner_sheet_idx = self.chart_owner_sheet_idx(&self.charts[chart_idx].0)?;
            {
                let chart = &mut self.charts[chart_idx].1;
                visit_chart_references(chart, |formula| {
                    let shifted = shift_references_for_owner(
                        formula,
                        owner_sheet_idx,
                        sheet_idx,
                        &target_sheet_name_lowercase,
                        shift_cell,
                    )?;
                    if shifted != *formula {
                        changed = true;
                        *formula = shifted;
                    }
                    Ok(())
                })?;
            }
            if changed {
                dirty_chart_paths.push(self.charts[chart_idx].0.clone());
            }
        }
        for path in dirty_chart_paths {
            self.mark_graph_part_dirty(&path);
        }

        // Drawing anchors attached to this sheet.
        if let Some(&drawing_idx) = self.worksheet_drawings.get(&sheet_idx) {
            let mut changed = false;
            if let Some((_, drawing)) = self.drawings.get_mut(drawing_idx) {
                for anchor in &mut drawing.one_cell_anchors {
                    let (new_col, new_row) = shift_cell(anchor.from.col + 1, anchor.from.row + 1);
                    changed |= new_col != anchor.from.col + 1 || new_row != anchor.from.row + 1;
                    anchor.from.col = new_col - 1;
                    anchor.from.row = new_row - 1;
                }
                for anchor in &mut drawing.two_cell_anchors {
                    let (from_col, from_row) = shift_cell(anchor.from.col + 1, anchor.from.row + 1);
                    changed |= from_col != anchor.from.col + 1 || from_row != anchor.from.row + 1;
                    anchor.from.col = from_col - 1;
                    anchor.from.row = from_row - 1;
                    let (to_col, to_row) = shift_cell(anchor.to.col + 1, anchor.to.row + 1);
                    changed |= to_col != anchor.to.col + 1 || to_row != anchor.to.row + 1;
                    anchor.to.col = to_col - 1;
                    anchor.to.row = to_row - 1;
                }
            }
            if changed {
                self.mark_drawing_dirty(drawing_idx);
            }
        }

        Ok(())
    }

    /// Ensure a drawing exists for the given sheet index, creating one if needed.
    /// Returns the drawing index.
    pub(crate) fn ensure_drawing_for_sheet(&mut self, sheet_idx: usize) -> usize {
        if let Some(&idx) = self.worksheet_drawings.get(&sheet_idx) {
            return idx;
        }

        let idx = self.drawings.len();
        let drawing_path = self.next_available_part_path("xl/drawings/drawing", ".xml");
        self.drawings.push((drawing_path.clone(), WsDr::default()));
        self.mark_graph_part_dirty(&drawing_path);
        self.worksheet_drawings.insert(sheet_idx, idx);

        // Add drawing reference to the worksheet.
        let ws_rid = self.next_worksheet_rid(sheet_idx);
        // ensure_hydrated can only fail if the sheet was never loaded, which
        // should not happen for a sheet we're actively attaching a drawing to.
        // Use expect instead of `?` because this method returns `usize`.
        self.ensure_hydrated(sheet_idx)
            .expect("sheet must be hydrated before attaching a drawing");
        self.mark_sheet_dirty(sheet_idx);
        self.worksheets[sheet_idx].1.get_mut().unwrap().drawing = Some(DrawingRef {
            r_id: ws_rid.clone(),
        });

        // Add worksheet->drawing relationship.
        let drawing_rel_target = crate::workbook_paths::relative_relationship_target(
            &self.sheet_part_path(sheet_idx),
            &drawing_path,
        );
        let ws_rels = self
            .worksheet_rels
            .entry(sheet_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        ws_rels.relationships.push(Relationship {
            id: ws_rid,
            rel_type: rel_types::DRAWING.to_string(),
            target: drawing_rel_target,
            target_mode: None,
        });

        // Add content type for the drawing.
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/{drawing_path}"),
            content_type: mime_types::DRAWING.to_string(),
        });

        idx
    }

    pub(crate) fn occupied_part_paths(&self) -> std::collections::HashSet<String> {
        let mut paths: std::collections::HashSet<String> = self
            .content_types
            .overrides
            .iter()
            .map(|entry| entry.part_name.trim_start_matches('/').to_string())
            .collect();
        paths.extend((0..self.worksheets.len()).map(|idx| self.sheet_part_path(idx)));
        paths.extend(self.charts.iter().map(|(path, _)| path.clone()));
        paths.extend(self.raw_charts.iter().map(|(path, _)| path.clone()));
        paths.extend(self.drawings.iter().map(|(path, _)| path.clone()));
        paths.extend(self.raw_graph_parts.keys().cloned());
        paths.extend(self.images.iter().map(|(path, _)| path.clone()));
        paths.extend(self.tables.iter().map(|(path, _, _)| path.clone()));
        paths.extend(self.pivot_tables.iter().map(|(path, _)| path.clone()));
        paths.extend(self.pivot_cache_defs.iter().map(|(path, _)| path.clone()));
        paths.extend(
            self.pivot_cache_records
                .iter()
                .map(|(path, _)| path.clone()),
        );
        paths.extend(self.slicer_defs.iter().map(|(path, _)| path.clone()));
        paths.extend(self.slicer_caches.iter().map(|(path, _)| path.clone()));
        paths.extend(self.unknown_parts.iter().map(|(path, _)| path.clone()));
        paths.extend(
            self.deferred_parts
                .remaining_parts()
                .map(|(path, _)| path.to_string()),
        );
        paths
    }

    pub(crate) fn next_available_part_path(&self, prefix: &str, suffix: &str) -> String {
        let occupied = self.occupied_part_paths();
        next_numbered_part_path(prefix, suffix, &occupied)
    }

    /// Generate the next relationship ID for a worksheet's rels.
    pub(crate) fn next_worksheet_rid(&self, sheet_idx: usize) -> String {
        let existing = self
            .worksheet_rels
            .get(&sheet_idx)
            .map(|r| r.relationships.as_slice())
            .unwrap_or(&[]);
        crate::sheet::next_rid(existing)
    }

    /// Generate the next relationship ID for a drawing's rels.
    pub(crate) fn next_drawing_rid(&self, drawing_idx: usize) -> String {
        let existing = self
            .drawing_rels
            .get(&drawing_idx)
            .map(|r| r.relationships.as_slice())
            .unwrap_or(&[]);
        crate::sheet::next_rid(existing)
    }
}

#[cfg(test)]
#[allow(clippy::approx_constant)]
mod tests {
    use super::*;
    use crate::workbook::aux::AuxCategory;
    use std::io::Read;
    use tempfile::TempDir;

    fn zip_part(buffer: &[u8], name: &str) -> Vec<u8> {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer)).unwrap();
        let mut bytes = Vec::new();
        archive
            .by_name(name)
            .unwrap()
            .read_to_end(&mut bytes)
            .unwrap();
        bytes
    }

    fn zip_entry_count(buffer: &[u8], name: &str) -> usize {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer)).unwrap();
        (0..archive.len())
            .filter(|index| archive.by_index(*index).unwrap().name() == name)
            .count()
    }

    fn assert_package_integrity(bytes: &[u8]) {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let names: Vec<String> = (0..archive.len())
            .map(|index| archive.by_index(index).unwrap().name().to_string())
            .collect();
        let unique_names: std::collections::HashSet<&str> =
            names.iter().map(String::as_str).collect();
        assert_eq!(names.len(), unique_names.len(), "duplicate ZIP entry");

        let mut content_types_xml = String::new();
        archive
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_string(&mut content_types_xml)
            .unwrap();
        let content_types: ContentTypes = quick_xml::de::from_str(&content_types_xml).unwrap();
        let mut part_names = std::collections::HashSet::new();
        for entry in content_types.overrides {
            assert!(part_names.insert(entry.part_name.clone()));
            let is_unset_document_property = matches!(
                entry.part_name.as_str(),
                "/docProps/core.xml" | "/docProps/app.xml" | "/docProps/custom.xml"
            ) && !unique_names
                .contains(entry.part_name.trim_start_matches('/'));
            if is_unset_document_property {
                continue;
            }
            assert!(
                unique_names.contains(entry.part_name.trim_start_matches('/')),
                "content type target missing: {}",
                entry.part_name
            );
        }

        let rel_paths: Vec<String> = names
            .iter()
            .filter(|name| name.ends_with(".rels"))
            .cloned()
            .collect();
        for rel_path in rel_paths {
            let mut xml = String::new();
            archive
                .by_name(&rel_path)
                .unwrap()
                .read_to_string(&mut xml)
                .unwrap();
            let relationships: Relationships = quick_xml::de::from_str(&xml).unwrap();
            let owner = if rel_path == "_rels/.rels" {
                String::new()
            } else {
                let (dir, file) = rel_path.rsplit_once("/_rels/").unwrap();
                format!("{dir}/{}", file.trim_end_matches(".rels"))
            };
            for relationship in relationships.relationships {
                if relationship.target_mode.as_deref() == Some("External") {
                    continue;
                }
                let target = crate::workbook_paths::resolve_relationship_target(
                    &owner,
                    &relationship.target,
                );
                let is_unset_document_property = rel_path == "_rels/.rels"
                    && matches!(
                        relationship.rel_type.as_str(),
                        rel_types::CORE_PROPERTIES
                            | rel_types::EXTENDED_PROPERTIES
                            | rel_types::CUSTOM_PROPERTIES
                    )
                    && !unique_names.contains(target.as_str());
                if is_unset_document_property {
                    continue;
                }
                assert!(
                    unique_names.contains(target.as_str()),
                    "relationship target missing: {rel_path} -> {target}"
                );
            }
        }
    }

    #[test]
    fn test_new_sheet_basic() {
        let mut wb = Workbook::new();
        let idx = wb.new_sheet("Sheet2").unwrap();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_names(), vec!["Sheet1", "Sheet2"]);
    }

    #[test]
    fn test_sparse_worksheet_paths_save_with_package_integrity() {
        let mut wb = Workbook::new();
        let first_rid = wb.workbook_xml.sheets.sheets[0].r_id.clone();
        wb.workbook_rels
            .relationships
            .iter_mut()
            .find(|relationship| relationship.id == first_rid)
            .unwrap()
            .target = "worksheets/sheet2.xml".to_string();
        let worksheet_override = wb
            .content_types
            .overrides
            .iter_mut()
            .find(|entry| entry.content_type == mime_types::WORKSHEET)
            .unwrap();
        worksheet_override.part_name = "/xl/worksheets/sheet2.xml".to_string();

        wb.new_sheet("Added").unwrap();
        assert_eq!(wb.sheet_part_path(0), "xl/worksheets/sheet2.xml");
        assert_eq!(wb.sheet_part_path(1), "xl/worksheets/sheet1.xml");
        let bytes = wb.save_to_buffer().unwrap();
        assert_package_integrity(&bytes);
    }

    #[test]
    fn test_new_sheet_skips_unknown_worksheet_path_collision() {
        let mut wb = Workbook::new();
        let first_rid = wb.workbook_xml.sheets.sheets[0].r_id.clone();
        wb.workbook_rels
            .relationships
            .iter_mut()
            .find(|relationship| relationship.id == first_rid)
            .unwrap()
            .target = "worksheets/sheet9.xml".to_string();
        wb.content_types
            .overrides
            .iter_mut()
            .find(|entry| entry.content_type == mime_types::WORKSHEET)
            .unwrap()
            .part_name = "/xl/worksheets/sheet9.xml".to_string();
        wb.unknown_parts
            .push(("xl/worksheets/sheet1.xml".to_string(), b"reserved".to_vec()));

        wb.new_sheet("Added").unwrap();

        assert_eq!(wb.sheet_part_path(1), "xl/worksheets/sheet2.xml");
    }

    #[test]
    fn test_new_sheet_duplicate_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.new_sheet("Sheet1");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::SheetAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_new_sheet_invalid_name_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.new_sheet("Bad/Name");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::InvalidSheetName(_)));
    }

    #[test]
    fn test_delete_sheet_basic() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.delete_sheet("Sheet1").unwrap();
        assert_eq!(wb.sheet_names(), vec!["Sheet2"]);
    }

    #[test]
    fn test_delete_sheet_keeps_parallel_vecs_in_sync() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.new_sheet("Sheet3").unwrap();

        // Add comments to Sheet2 (middle sheet).
        wb.add_comment(
            "Sheet2",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Test".to_string(),
                text: "note".to_string(),
            },
        )
        .unwrap();

        // Delete the middle sheet and verify no panic.
        wb.delete_sheet("Sheet2").unwrap();
        assert_eq!(wb.sheet_names(), vec!["Sheet1", "Sheet3"]);

        // After deletion, adding a comment to Sheet3 (now index 1)
        // should work without index mismatch.
        wb.add_comment(
            "Sheet3",
            &crate::comment::CommentConfig {
                cell: "B2".to_string(),
                author: "Test".to_string(),
                text: "note2".to_string(),
            },
        )
        .unwrap();
    }

    #[test]
    fn test_delete_sheet_reindexes_local_names_and_all_active_tabs() {
        use sheetkit_xml::workbook::{DefinedName, DefinedNames};

        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.new_sheet("Sheet3").unwrap();
        wb.set_defined_name("OnFirst", "Sheet1!$A$1", Some("Sheet1"), None)
            .unwrap();
        wb.set_defined_name("OnSecond", "Sheet2!$A$1", Some("Sheet2"), None)
            .unwrap();
        wb.set_defined_name("OnThird", "Sheet3!$A$1", Some("Sheet3"), None)
            .unwrap();
        wb.workbook_xml
            .defined_names
            .get_or_insert_with(|| DefinedNames {
                defined_names: vec![],
            })
            .defined_names
            .push(DefinedName {
                name: "Global".to_string(),
                local_sheet_id: None,
                comment: None,
                hidden: None,
                value: "Sheet3!$B$2".to_string(),
            });
        let views = wb.workbook_xml.book_views.as_mut().unwrap();
        let mut extra_view = views.workbook_views[0].clone();
        views.workbook_views[0].active_tab = Some(1);
        extra_view.active_tab = Some(99);
        views.workbook_views.push(extra_view);

        wb.delete_sheet("Sheet2").unwrap();

        let names = wb.workbook_xml.defined_names.as_ref().unwrap();
        assert!(names
            .defined_names
            .iter()
            .all(|name| name.name != "OnSecond"));
        assert_eq!(
            names
                .defined_names
                .iter()
                .find(|name| name.name == "OnThird")
                .unwrap()
                .local_sheet_id,
            Some(1)
        );
        assert_eq!(
            wb.workbook_xml
                .book_views
                .as_ref()
                .unwrap()
                .workbook_views
                .iter()
                .map(|view| view.active_tab)
                .collect::<Vec<_>>(),
            vec![Some(1), Some(1)]
        );
    }

    #[test]
    fn test_delete_last_sheet_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.delete_sheet("Sheet1");
        assert!(result.is_err());
    }

    #[test]
    fn test_delete_nonexistent_sheet_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.delete_sheet("NoSuchSheet");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_set_sheet_name_basic() {
        let mut wb = Workbook::new();
        wb.set_sheet_name("Sheet1", "Renamed").unwrap();
        assert_eq!(wb.sheet_names(), vec!["Renamed"]);
    }

    #[test]
    fn test_set_sheet_name_to_existing_returns_error() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        let result = wb.set_sheet_name("Sheet1", "Sheet2");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::SheetAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_copy_sheet_basic() {
        let mut wb = Workbook::new();
        let idx = wb.copy_sheet("Sheet1", "Sheet1 Copy").unwrap();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_names(), vec!["Sheet1", "Sheet1 Copy"]);
    }

    #[test]
    fn test_copy_sheet_clones_mixed_relationship_graph_and_roundtrips() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        use crate::hyperlink::HyperlinkType;
        use crate::table::{TableColumn, TableConfig};

        let mut wb = Workbook::new();
        wb.set_cell_hyperlink(
            "Sheet1",
            "A1",
            HyperlinkType::External("https://example.com/source".to_string()),
            Some("Source"),
            None,
        )
        .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "B2".to_string(),
                author: "Author".to_string(),
                text: "Copied note".to_string(),
            },
        )
        .unwrap();
        wb.add_threaded_comment(
            "Sheet1",
            "C3",
            &crate::threaded_comment::ThreadedCommentInput {
                author: "Author".to_string(),
                text: "Copied thread".to_string(),
                parent_id: None,
            },
        )
        .unwrap();
        wb.add_table(
            "Sheet1",
            &TableConfig {
                name: "SourceTable".to_string(),
                display_name: "SourceTable".to_string(),
                range: "A1:B3".to_string(),
                columns: vec![
                    TableColumn {
                        name: "Name".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumn {
                        name: "Value".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                ],
                ..TableConfig::default()
            },
        )
        .unwrap();
        wb.add_chart(
            "Sheet1",
            "E1",
            "L10",
            &ChartConfig {
                chart_type: ChartType::Line,
                title: None,
                series: vec![ChartSeries {
                    name: "Series".to_string(),
                    categories: "Sheet1!$A$1:$A$3".to_string(),
                    values: "Sheet1!$B$1:$B$3".to_string(),
                    x_values: None,
                    bubble_sizes: None,
                }],
                show_legend: false,
                view_3d: None,
            },
        )
        .unwrap();

        wb.copy_sheet("Sheet1", "Copied").unwrap();

        let copied_link = wb.get_cell_hyperlink("Copied", "A1").unwrap().unwrap();
        assert_eq!(
            copied_link.link_type,
            HyperlinkType::External("https://example.com/source".to_string())
        );
        assert_eq!(wb.get_comments("Copied").unwrap().len(), 1);
        assert_eq!(wb.get_threaded_comments("Copied").unwrap().len(), 1);
        let copied_tables = wb.get_tables("Copied").unwrap();
        assert_eq!(copied_tables.len(), 1);
        assert_ne!(copied_tables[0].name, "SourceTable");
        assert_eq!(wb.drawings.len(), 2);
        assert_eq!(wb.charts.len(), 2);

        wb.delete_sheet("Sheet1").unwrap();
        let bytes = wb.save_to_buffer().unwrap();
        assert_package_integrity(&bytes);
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes.clone())).unwrap();
        let names: Vec<String> = (0..archive.len())
            .map(|index| archive.by_index(index).unwrap().name().to_string())
            .collect();
        let unique_names: std::collections::HashSet<&str> =
            names.iter().map(String::as_str).collect();
        assert_eq!(names.len(), unique_names.len());
        assert_eq!(
            names
                .iter()
                .filter(|name| name.starts_with("xl/charts/chart"))
                .count(),
            1
        );
        let mut reopened = Workbook::open_from_buffer(&bytes).unwrap();
        assert_eq!(reopened.sheet_names(), vec!["Copied"]);
        assert!(reopened
            .get_cell_hyperlink("Copied", "A1")
            .unwrap()
            .is_some());
        assert_eq!(reopened.get_comments("Copied").unwrap().len(), 1);
        assert_eq!(reopened.get_threaded_comments("Copied").unwrap().len(), 1);
        assert_eq!(reopened.get_tables("Copied").unwrap().len(), 1);
    }

    #[test]
    fn test_delete_sheet_retargets_threaded_comments_by_surviving_sheet() {
        let mut wb = Workbook::new();
        wb.new_sheet("Second").unwrap();
        wb.new_sheet("Third").unwrap();
        for (sheet, text) in [("Second", "Second thread"), ("Third", "Third thread")] {
            wb.add_threaded_comment(
                sheet,
                "A1",
                &crate::threaded_comment::ThreadedCommentInput {
                    author: "Author".to_string(),
                    text: text.to_string(),
                    parent_id: None,
                },
            )
            .unwrap();
        }

        wb.delete_sheet("Sheet1").unwrap();
        let bytes = wb.save_to_buffer().unwrap();
        let reopened = Workbook::open_from_buffer(&bytes).unwrap();

        assert_eq!(
            reopened.get_threaded_comments("Second").unwrap()[0].text,
            "Second thread"
        );
        assert_eq!(
            reopened.get_threaded_comments("Third").unwrap()[0].text,
            "Third thread"
        );
    }

    #[test]
    fn test_delete_last_threaded_sheet_removes_person_and_comment_parts() {
        let mut wb = Workbook::new();
        wb.new_sheet("Survivor").unwrap();
        wb.add_threaded_comment(
            "Sheet1",
            "A1",
            &crate::threaded_comment::ThreadedCommentInput {
                author: "Author".to_string(),
                text: "Removed thread".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        wb.delete_sheet("Sheet1").unwrap();
        let bytes = wb.save_to_buffer().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes.clone())).unwrap();
        let names: Vec<String> = (0..archive.len())
            .map(|index| archive.by_index(index).unwrap().name().to_string())
            .collect();
        assert!(!names
            .iter()
            .any(|name| name.starts_with("xl/threadedComments/")));
        assert!(!names.iter().any(|name| name == "xl/persons/person.xml"));

        let reopened = Workbook::open_from_buffer(&bytes).unwrap();
        assert!(!reopened
            .workbook_rels
            .relationships
            .iter()
            .any(|relationship| {
                relationship.rel_type == sheetkit_xml::threaded_comment::REL_TYPE_PERSON
            }));
        assert!(!reopened.content_types.overrides.iter().any(|entry| {
            entry.content_type == sheetkit_xml::threaded_comment::THREADED_COMMENTS_CONTENT_TYPE
                || entry.content_type == sheetkit_xml::threaded_comment::PERSON_LIST_CONTENT_TYPE
        }));
    }

    #[test]
    fn test_untouched_lazy_threaded_comments_survive_save() {
        let mut wb = Workbook::new();
        wb.add_threaded_comment(
            "Sheet1",
            "A1",
            &crate::threaded_comment::ThreadedCommentInput {
                author: "Author".to_string(),
                text: "Deferred thread".to_string(),
                parent_id: None,
            },
        )
        .unwrap();
        let original = wb.save_to_buffer().unwrap();
        let options = OpenOptions::new().read_mode(ReadMode::Lazy);
        let lazy = Workbook::open_from_buffer_with_options(&original, &options).unwrap();

        let saved = lazy.save_to_buffer().unwrap();
        let reopened = Workbook::open_from_buffer(&saved).unwrap();

        let comments = reopened.get_threaded_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].text, "Deferred thread");
        assert_eq!(reopened.get_persons().len(), 1);
    }

    #[test]
    fn test_copy_sheet_rejects_unsupported_relationship_atomically() {
        let mut wb = Workbook::new();
        wb.worksheet_rels.insert(
            0,
            Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![Relationship {
                    id: "rId1".to_string(),
                    rel_type: rel_types::PRINTER_SETTINGS.to_string(),
                    target: "../printerSettings/printerSettings1.bin".to_string(),
                    target_mode: None,
                }],
            },
        );
        let before_names: Vec<String> = wb.sheet_names().into_iter().map(str::to_string).collect();
        let before_sheet_count = wb.workbook_xml.sheets.sheets.len();

        let error = wb.copy_sheet("Sheet1", "Copied").unwrap_err();

        assert!(matches!(error, Error::InvalidArgument(_)));
        assert_eq!(
            wb.sheet_names(),
            before_names.iter().map(String::as_str).collect::<Vec<_>>()
        );
        assert_eq!(wb.workbook_xml.sheets.sheets.len(), before_sheet_count);
    }

    fn assert_copy_rejected_without_mutation(mut wb: Workbook) {
        let before_names: Vec<String> = wb.sheet_names().into_iter().map(str::to_string).collect();
        assert!(matches!(
            wb.copy_sheet("Sheet1", "Copied"),
            Err(Error::InvalidArgument(_))
        ));
        assert_eq!(
            wb.sheet_names(),
            before_names.iter().map(String::as_str).collect::<Vec<_>>()
        );
    }

    #[test]
    fn test_copy_sheet_rejects_unresolved_supported_relationships() {
        let relationship = |rel_type: &str, target: &str| Relationship {
            id: "rId99".to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: None,
        };

        let mut comments = Workbook::new();
        comments
            .add_comment(
                "Sheet1",
                &crate::comment::CommentConfig {
                    cell: "A1".to_string(),
                    author: "Author".to_string(),
                    text: "Note".to_string(),
                },
            )
            .unwrap();
        comments
            .worksheet_rels
            .entry(0)
            .or_insert_with(crate::workbook_paths::default_relationships)
            .relationships
            .push(relationship(rel_types::COMMENTS, "../comments99.xml"));
        assert_copy_rejected_without_mutation(comments);

        let mut vml = Workbook::new();
        vml.sheet_vml[0] = Some(b"<xml/>".to_vec());
        vml.worksheet_rels
            .entry(0)
            .or_insert_with(crate::workbook_paths::default_relationships)
            .relationships
            .push(relationship(
                rel_types::VML_DRAWING,
                "../drawings/vmlDrawing99.vml",
            ));
        assert_copy_rejected_without_mutation(vml);

        let mut table = Workbook::new();
        table
            .worksheet_rels
            .entry(0)
            .or_insert_with(crate::workbook_paths::default_relationships)
            .relationships
            .push(relationship(rel_types::TABLE, "../tables/table99.xml"));
        assert_copy_rejected_without_mutation(table);

        let mut threaded = Workbook::new();
        threaded
            .add_threaded_comment(
                "Sheet1",
                "A1",
                &crate::threaded_comment::ThreadedCommentInput {
                    author: "Author".to_string(),
                    text: "Thread".to_string(),
                    parent_id: None,
                },
            )
            .unwrap();
        threaded
            .worksheet_rels
            .entry(0)
            .or_insert_with(crate::workbook_paths::default_relationships)
            .relationships
            .push(relationship(
                sheetkit_xml::threaded_comment::REL_TYPE_THREADED_COMMENT,
                "../threadedComments/threadedComment99.xml",
            ));
        assert_copy_rejected_without_mutation(threaded);
    }

    #[test]
    fn test_copy_sheet_rejects_missing_drawing_relationships_and_rids() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let config = ChartConfig {
            chart_type: ChartType::Line,
            title: None,
            series: vec![ChartSeries {
                name: "Series".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        };

        let mut missing_rels = Workbook::new();
        missing_rels
            .add_chart("Sheet1", "E1", "L10", &config)
            .unwrap();
        missing_rels.drawing_rels.clear();
        assert_copy_rejected_without_mutation(missing_rels);

        let mut missing_rid = Workbook::new();
        missing_rid
            .add_chart("Sheet1", "E1", "L10", &config)
            .unwrap();
        missing_rid.drawings[0].1.two_cell_anchors[0]
            .graphic_frame
            .as_mut()
            .unwrap()
            .graphic
            .graphic_data
            .chart
            .r_id = "missing".to_string();
        assert_copy_rejected_without_mutation(missing_rid);
    }

    #[test]
    fn test_copy_sheet_rejects_unresolved_worksheet_rids_atomically() {
        let mut drawing = Workbook::new();
        drawing.worksheets[0].1.get_mut().unwrap().drawing =
            Some(sheetkit_xml::worksheet::DrawingRef {
                r_id: "missingDrawing".to_string(),
            });
        assert_copy_rejected_without_mutation(drawing);

        let mut legacy = Workbook::new();
        legacy.worksheets[0].1.get_mut().unwrap().legacy_drawing =
            Some(sheetkit_xml::worksheet::LegacyDrawingRef {
                r_id: "missingLegacy".to_string(),
            });
        assert_copy_rejected_without_mutation(legacy);

        let mut hyperlink = Workbook::new();
        hyperlink.worksheets[0].1.get_mut().unwrap().hyperlinks =
            Some(sheetkit_xml::worksheet::Hyperlinks {
                hyperlinks: vec![sheetkit_xml::worksheet::Hyperlink {
                    reference: "A1".to_string(),
                    r_id: Some("missingHyperlink".to_string()),
                    location: None,
                    display: None,
                    tooltip: None,
                }],
            });
        assert_copy_rejected_without_mutation(hyperlink);

        let mut table = Workbook::new();
        table.worksheets[0].1.get_mut().unwrap().table_parts =
            Some(sheetkit_xml::worksheet::TableParts {
                count: Some(1),
                table_parts: vec![sheetkit_xml::worksheet::TablePart {
                    r_id: "missingTable".to_string(),
                }],
            });
        assert_copy_rejected_without_mutation(table);

        let mut page_setup = Workbook::new();
        page_setup.worksheets[0].1.get_mut().unwrap().page_setup =
            Some(sheetkit_xml::worksheet::PageSetup {
                paper_size: None,
                orientation: None,
                scale: None,
                fit_to_width: None,
                fit_to_height: None,
                first_page_number: None,
                horizontal_dpi: None,
                vertical_dpi: None,
                r_id: Some("missingPrinter".to_string()),
            });
        assert_copy_rejected_without_mutation(page_setup);
    }

    #[test]
    fn test_copy_sheet_clones_image_and_source_delete_keeps_copy() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x89, 0x50, 0x4e, 0x47],
                format: ImageFormat::Png,
                from_cell: "B2".to_string(),
                width_px: 32,
                height_px: 32,
            },
        )
        .unwrap();

        wb.copy_sheet("Sheet1", "Copied").unwrap();
        assert_eq!(wb.images.len(), 2);
        assert_ne!(wb.images[0].0, wb.images[1].0);
        assert_eq!(wb.images[0].1, wb.images[1].1);

        wb.delete_sheet("Sheet1").unwrap();
        assert_eq!(wb.images.len(), 1);
        let bytes = wb.save_to_buffer().unwrap();
        assert_package_integrity(&bytes);
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes.clone())).unwrap();
        let media_count = (0..archive.len())
            .filter(|index| {
                archive
                    .by_index(*index)
                    .unwrap()
                    .name()
                    .starts_with("xl/media/")
            })
            .count();
        assert_eq!(media_count, 1);
        Workbook::open_from_buffer(&bytes).unwrap();
    }

    #[test]
    fn test_delete_sheet_retains_image_referenced_by_surviving_drawing() {
        use crate::image::{ImageConfig, ImageFormat};
        let config = ImageConfig {
            data: vec![0x89, 0x50, 0x4e, 0x47],
            format: ImageFormat::Png,
            from_cell: "B2".to_string(),
            width_px: 32,
            height_px: 32,
        };
        let mut wb = Workbook::new();
        wb.add_image("Sheet1", &config).unwrap();
        wb.new_sheet("Survivor").unwrap();
        wb.add_image("Survivor", &config).unwrap();
        let shared_image_path = wb.images[0].0.clone();
        let survivor_drawing_idx = wb.worksheet_drawings[&1];
        let survivor_drawing_path = wb.drawings[survivor_drawing_idx].0.clone();
        wb.drawing_rels
            .get_mut(&survivor_drawing_idx)
            .unwrap()
            .relationships[0]
            .target = crate::workbook_paths::relative_relationship_target(
            &survivor_drawing_path,
            &shared_image_path,
        );
        wb.images.remove(1);

        wb.delete_sheet("Sheet1").unwrap();

        assert_eq!(wb.images.len(), 1);
        assert_eq!(wb.images[0].0, shared_image_path);
        let bytes = wb.save_to_buffer().unwrap();
        assert_package_integrity(&bytes);
    }

    #[test]
    fn test_copy_sheet_remaps_threaded_reply_ids() {
        let mut wb = Workbook::new();
        let root_id = wb
            .add_threaded_comment(
                "Sheet1",
                "A1",
                &crate::threaded_comment::ThreadedCommentInput {
                    author: "Author".to_string(),
                    text: "Root".to_string(),
                    parent_id: None,
                },
            )
            .unwrap();
        wb.add_threaded_comment(
            "Sheet1",
            "A1",
            &crate::threaded_comment::ThreadedCommentInput {
                author: "Author".to_string(),
                text: "Reply".to_string(),
                parent_id: Some(root_id),
            },
        )
        .unwrap();

        wb.copy_sheet("Sheet1", "Copied").unwrap();
        let source = wb.sheet_threaded_comments[0].as_ref().unwrap();
        let copied = wb.sheet_threaded_comments[1].as_ref().unwrap();
        let source_ids: std::collections::HashSet<&str> = source
            .comments
            .iter()
            .map(|comment| comment.id.as_str())
            .collect();
        assert!(copied
            .comments
            .iter()
            .all(|comment| !source_ids.contains(comment.id.as_str())));
        assert_eq!(
            copied.comments[1].parent_id.as_deref(),
            Some(copied.comments[0].id.as_str())
        );
        assert_eq!(source.comments[0].person_id, copied.comments[0].person_id);
    }

    #[test]
    fn test_copy_sheet_table_id_overflow_is_atomic() {
        use crate::table::{TableColumn, TableConfig};
        let mut wb = Workbook::new();
        wb.add_table(
            "Sheet1",
            &TableConfig {
                name: "AtLimit".to_string(),
                display_name: "AtLimit".to_string(),
                range: "A1:A2".to_string(),
                columns: vec![TableColumn {
                    name: "Value".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                }],
                ..TableConfig::default()
            },
        )
        .unwrap();
        wb.tables[0].1.id = u32::MAX;
        assert_copy_rejected_without_mutation(wb);
    }

    #[test]
    fn test_delete_sheet_rejects_pivot_and_slicer_closure_atomically() {
        use crate::workbook::aux::AuxCategory;
        for rel_type in [rel_types::PIVOT_TABLE, rel_types::SLICER] {
            let mut wb = Workbook::new();
            wb.new_sheet("Survivor").unwrap();
            wb.worksheet_rels.insert(
                0,
                Relationships {
                    xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                    relationships: vec![Relationship {
                        id: "rId1".to_string(),
                        rel_type: rel_type.to_string(),
                        target: "../unsupported/part1.xml".to_string(),
                        target_mode: None,
                    }],
                },
            );
            wb.deferred_parts.insert(
                "xl/pivotTables/pivotTable1.xml".to_string(),
                b"deferred".to_vec(),
            );
            assert!(matches!(
                wb.delete_sheet("Sheet1"),
                Err(Error::InvalidArgument(_))
            ));
            assert_eq!(wb.sheet_names(), vec!["Sheet1", "Survivor"]);
            assert!(wb.deferred_parts.has_category(AuxCategory::PivotTables));
        }
    }

    #[test]
    fn test_copy_sheet_skips_unknown_part_path_collisions() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};

        let mut wb = Workbook::new();
        wb.add_chart(
            "Sheet1",
            "E1",
            "L10",
            &ChartConfig {
                chart_type: ChartType::Line,
                title: None,
                series: vec![ChartSeries {
                    name: "Series".to_string(),
                    categories: "Sheet1!$A$1:$A$3".to_string(),
                    values: "Sheet1!$B$1:$B$3".to_string(),
                    x_values: None,
                    bubble_sizes: None,
                }],
                show_legend: false,
                view_3d: None,
            },
        )
        .unwrap();
        wb.unknown_parts.extend([
            ("xl/drawings/drawing2.xml".to_string(), b"opaque".to_vec()),
            ("xl/charts/chart2.xml".to_string(), b"opaque".to_vec()),
        ]);

        wb.copy_sheet("Sheet1", "Copied").unwrap();

        assert!(wb
            .drawings
            .iter()
            .any(|(path, _)| path == "xl/drawings/drawing3.xml"));
        assert!(wb
            .charts
            .iter()
            .any(|(path, _)| path == "xl/charts/chart3.xml"));
    }

    #[test]
    fn test_normal_drawing_chart_and_image_allocation_skips_reserved_paths() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        wb.unknown_parts.extend([
            ("xl/drawings/drawing1.xml".to_string(), b"opaque".to_vec()),
            ("xl/charts/chart1.xml".to_string(), b"opaque".to_vec()),
            ("xl/media/image1.png".to_string(), b"opaque".to_vec()),
        ]);
        wb.content_types.overrides.extend([
            ContentTypeOverride {
                part_name: "/xl/drawings/drawing2.xml".to_string(),
                content_type: mime_types::DRAWING.to_string(),
            },
            ContentTypeOverride {
                part_name: "/xl/charts/chart2.xml".to_string(),
                content_type: mime_types::CHART.to_string(),
            },
        ]);
        wb.add_chart(
            "Sheet1",
            "E1",
            "L10",
            &ChartConfig {
                chart_type: ChartType::Line,
                title: None,
                series: vec![ChartSeries {
                    name: "Series".to_string(),
                    categories: "Sheet1!$A$1:$A$3".to_string(),
                    values: "Sheet1!$B$1:$B$3".to_string(),
                    x_values: None,
                    bubble_sizes: None,
                }],
                show_legend: false,
                view_3d: None,
            },
        )
        .unwrap();
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x89, 0x50, 0x4e, 0x47],
                format: ImageFormat::Png,
                from_cell: "B2".to_string(),
                width_px: 32,
                height_px: 32,
            },
        )
        .unwrap();

        assert_eq!(wb.drawings[0].0, "xl/drawings/drawing3.xml");
        assert_eq!(wb.charts[0].0, "xl/charts/chart3.xml");
        assert_eq!(wb.images[0].0, "xl/media/image2.png");
    }

    #[test]
    fn test_save_rejects_raw_collision_without_dropping_unknown_part() {
        let mut wb = Workbook::new();
        wb.unknown_parts
            .push(("xl/comments1.xml".to_string(), b"opaque".to_vec()));
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "Note".to_string(),
            },
        )
        .unwrap();

        assert!(matches!(
            wb.save_to_buffer(),
            Err(Error::InvalidArgument(_))
        ));
        assert_eq!(wb.unknown_parts[0].0, "xl/comments1.xml");
        assert_eq!(wb.unknown_parts[0].1, b"opaque");
    }

    #[test]
    fn test_save_rejects_relationship_part_collision_before_writing() {
        use crate::hyperlink::HyperlinkType;
        let mut wb = Workbook::new();
        wb.set_cell_hyperlink(
            "Sheet1",
            "A1",
            HyperlinkType::External("https://example.com".to_string()),
            None,
            None,
        )
        .unwrap();
        let rel_path = crate::workbook_paths::relationship_part_path(&wb.sheet_part_path(0));
        wb.unknown_parts
            .push((rel_path.clone(), b"preserved relationships".to_vec()));

        assert!(matches!(
            wb.save_to_buffer(),
            Err(Error::InvalidArgument(_))
        ));
        assert_eq!(wb.unknown_parts[0].0, rel_path);
        assert_eq!(wb.unknown_parts[0].1, b"preserved relationships");

        wb.unknown_parts.clear();
        let bytes = wb.save_to_buffer().unwrap();
        assert_package_integrity(&bytes);
    }

    #[test]
    fn test_delete_sheet_rejects_malformed_deferred_lifecycle_part_atomically() {
        use crate::workbook::aux::AuxCategory;

        let mut wb = Workbook::new();
        wb.new_sheet("Survivor").unwrap();
        wb.worksheet_rels.insert(
            0,
            Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![Relationship {
                    id: "rId1".to_string(),
                    rel_type: rel_types::COMMENTS.to_string(),
                    target: "../comments1.xml".to_string(),
                    target_mode: None,
                }],
            },
        );
        assert!(wb.deferred_parts.insert(
            "xl/comments1.xml".to_string(),
            b"<comments><broken>".to_vec(),
        ));
        assert!(wb.deferred_parts.has_category(AuxCategory::Comments));
        let before_names: Vec<String> = wb.sheet_names().into_iter().map(str::to_string).collect();

        let error = wb.delete_sheet("Sheet1").unwrap_err();

        assert!(matches!(error, Error::XmlDeserialize(_)));
        assert_eq!(
            wb.sheet_names(),
            before_names.iter().map(String::as_str).collect::<Vec<_>>()
        );
        assert!(wb.deferred_parts.has_category(AuxCategory::Comments));
    }

    #[test]
    fn test_get_sheet_index() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        assert_eq!(wb.get_sheet_index("Sheet1"), Some(0));
        assert_eq!(wb.get_sheet_index("Sheet2"), Some(1));
        assert_eq!(wb.get_sheet_index("Nonexistent"), None);
    }

    #[test]
    fn test_get_active_sheet_default() {
        let wb = Workbook::new();
        assert_eq!(wb.get_active_sheet(), "Sheet1");
    }

    #[test]
    fn test_get_active_sheet_is_safe_for_empty_workbook() {
        let mut wb = Workbook::new();
        wb.worksheets.clear();

        assert_eq!(wb.get_active_sheet(), "");
    }

    #[test]
    fn test_set_active_sheet() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.set_active_sheet("Sheet2").unwrap();
        assert_eq!(wb.get_active_sheet(), "Sheet2");
    }

    #[test]
    fn test_set_active_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.set_active_sheet("NoSuchSheet");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_sheet_management_roundtrip_save_open() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("sheet_mgmt.xlsx");

        let mut wb = Workbook::new();
        wb.new_sheet("Data").unwrap();
        wb.new_sheet("Summary").unwrap();
        wb.set_sheet_name("Sheet1", "Overview").unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Overview", "Data", "Summary"]);
    }

    #[test]
    fn test_workbook_insert_rows() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "stay").unwrap();
        wb.set_cell_value("Sheet1", "A2", "shift").unwrap();
        wb.insert_rows("Sheet1", 2, 1).unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("stay".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "A3").unwrap(),
            CellValue::String("shift".to_string())
        );
        assert_eq!(wb.get_cell_value("Sheet1", "A2").unwrap(), CellValue::Empty);
    }

    #[test]
    fn test_zero_count_structural_edits_leave_lazy_sheet_clean() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("zero_count_structural_edit.xlsx");
        Workbook::new().save(&path).unwrap();

        let mut wb = Workbook::open(&path).unwrap();
        assert!(!wb.is_sheet_dirty(0));

        wb.insert_rows("Sheet1", 1, 0).unwrap();
        wb.insert_cols("Sheet1", "A", 0).unwrap();

        assert!(!wb.is_sheet_dirty(0));
    }

    #[test]
    fn test_workbook_insert_rows_updates_formula_and_ranges() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "C1",
            CellValue::Formula {
                expr: "SUM(A2:B2)".to_string(),
                result: None,
            },
        )
        .unwrap();
        wb.add_data_validation(
            "Sheet1",
            &crate::validation::DataValidationConfig::whole_number("A2:A5", 1, 9),
        )
        .unwrap();
        wb.set_auto_filter("Sheet1", "A2:B10").unwrap();
        wb.merge_cells("Sheet1", "A2", "B3").unwrap();

        wb.insert_rows("Sheet1", 2, 1).unwrap();

        match wb.get_cell_value("Sheet1", "C1").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "SUM(A3:B3)"),
            other => panic!("expected formula, got {other:?}"),
        }

        let validations = wb.get_data_validations("Sheet1").unwrap();
        assert_eq!(validations.len(), 1);
        assert_eq!(validations[0].sqref, "A3:A6");

        let merges = wb.get_merge_cells("Sheet1").unwrap();
        assert_eq!(merges, vec!["A3:B4".to_string()]);

        let ws = wb.worksheet_ref("Sheet1").unwrap();
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A3:B11");
    }

    #[test]
    fn test_structural_edit_updates_target_references_workbook_wide() {
        let mut wb = Workbook::new();
        wb.new_sheet("Other").unwrap();
        wb.new_sheet("다른 시트").unwrap();
        wb.set_cell_formula(
            "Other",
            "C1",
            "sheet1!A2+'다른 시트'!A2+A2+\"A2\"+LOG10(A2)",
        )
        .unwrap();
        wb.set_defined_name("BookRange", "Sheet1!$A$2:$A$3", None, None)
            .unwrap();
        wb.set_defined_name("LocalRange", "$A$2:$A$3", Some("Sheet1"), None)
            .unwrap();
        wb.set_defined_name("OtherRange", "$A$2:$A$3+Sheet1!A2", Some("Other"), None)
            .unwrap();
        wb.add_sparkline(
            "Other",
            &crate::sparkline::SparklineConfig::new("Sheet1!A2:A3", "B2"),
        )
        .unwrap();
        wb.add_table(
            "Sheet1",
            &crate::table::TableConfig {
                name: "ShiftedTable".into(),
                display_name: "ShiftedTable".into(),
                range: "A2:B4".into(),
                columns: vec![
                    crate::table::TableColumn {
                        name: "First".into(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    crate::table::TableColumn {
                        name: "Second".into(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                ],
                ..Default::default()
            },
        )
        .unwrap();
        wb.set_cell_formula("Sheet1", "A2", "A2").unwrap();

        {
            let ws = wb.worksheet_mut("Sheet1").unwrap();
            let formula = ws.sheet_data.rows[0].cells[0].f.as_mut().unwrap();
            formula.t = Some("shared".into());
            formula.reference = Some("A2:B2".into());
            formula.si = Some(0);
        }

        wb.insert_rows("Sheet1", 2, 1).unwrap();

        match wb.get_cell_value("Other", "C1").unwrap() {
            CellValue::Formula { expr, .. } => {
                assert_eq!(expr, "sheet1!A3+'다른 시트'!A2+A2+\"A2\"+LOG10(A2)");
            }
            value => panic!("expected formula, got {value:?}"),
        }
        assert_eq!(
            wb.get_defined_name("BookRange", None)
                .unwrap()
                .unwrap()
                .value,
            "Sheet1!$A$3:$A$4"
        );
        assert_eq!(
            wb.get_defined_name("LocalRange", Some("Sheet1"))
                .unwrap()
                .unwrap()
                .value,
            "$A$3:$A$4"
        );
        assert_eq!(
            wb.get_defined_name("OtherRange", Some("Other"))
                .unwrap()
                .unwrap()
                .value,
            "$A$2:$A$3+Sheet1!A3"
        );
        assert_eq!(
            wb.get_sparklines("Other").unwrap()[0].data_range,
            "Sheet1!A3:A4"
        );
        assert_eq!(wb.get_sparklines("Other").unwrap()[0].location, "B2");
        let shared = wb.worksheet_ref("Sheet1").unwrap().sheet_data.rows[0].cells[0]
            .f
            .as_ref()
            .unwrap();
        assert_eq!(shared.reference.as_deref(), Some("A3:B3"));
        assert_eq!(wb.get_tables("Sheet1").unwrap()[0].range, "A3:B5");
        assert_eq!(
            wb.tables[0].1.auto_filter.as_ref().unwrap().reference,
            "A3:B5"
        );
    }

    #[test]
    fn test_structural_edit_rejects_streamed_workbook_before_mutation() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "unchanged").unwrap();
        let writer = wb.new_stream_writer("Streamed").unwrap();
        wb.apply_stream_writer(writer).unwrap();

        let error = wb.insert_rows("Sheet1", 1, 1).unwrap_err();
        assert!(matches!(error, Error::InvalidArgument(_)));
        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("unchanged".into())
        );
    }

    #[test]
    fn test_structural_edit_rejects_sheet_rows_limit_before_mutation() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "unchanged").unwrap();
        wb.sheet_rows_limit = Some(1);

        let error = wb.insert_rows("Sheet1", 1, 1).unwrap_err();
        assert!(matches!(error, Error::InvalidArgument(_)));
        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("unchanged".into())
        );
    }

    #[test]
    fn test_structural_edit_updates_typed_chart_series_references() {
        let mut wb = Workbook::new();
        wb.add_chart(
            "Sheet1",
            "D1",
            "K10",
            &ChartConfig {
                chart_type: crate::chart::ChartType::Col,
                title: None,
                series: vec![crate::chart::ChartSeries {
                    name: "sheet1!$A$2".into(),
                    categories: "Sheet1!$A$2:$A$4".into(),
                    values: "Sheet1!$B$2:$B$4".into(),
                    x_values: None,
                    bubble_sizes: None,
                }],
                show_legend: false,
                view_3d: None,
            },
        )
        .unwrap();

        wb.insert_rows("Sheet1", 2, 1).unwrap();

        let series = &wb.charts[0]
            .1
            .chart
            .plot_area
            .bar_chart
            .as_ref()
            .unwrap()
            .series[0];
        assert_eq!(
            series.tx.as_ref().unwrap().str_ref.as_ref().unwrap().f,
            "sheet1!$A$3"
        );
        assert_eq!(
            series.cat.as_ref().unwrap().str_ref.as_ref().unwrap().f,
            "Sheet1!$A$3:$A$5"
        );
        assert_eq!(
            series.val.as_ref().unwrap().num_ref.as_ref().unwrap().f,
            "Sheet1!$B$3:$B$5"
        );
    }

    #[test]
    fn test_structural_edit_rejects_unparseable_deferred_chart_without_mutation() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("lazy_chart.xlsx");
        let mut wb = Workbook::new();
        wb.add_chart(
            "Sheet1",
            "D1",
            "K10",
            &ChartConfig {
                chart_type: crate::chart::ChartType::Col,
                title: None,
                series: vec![crate::chart::ChartSeries {
                    name: "Sheet1!$A$2".into(),
                    categories: "Sheet1!$A$2:$A$4".into(),
                    values: "Sheet1!$B$2:$B$4".into(),
                    x_values: None,
                    bubble_sizes: None,
                }],
                show_legend: false,
                view_3d: None,
            },
        )
        .unwrap();
        wb.set_cell_value("Sheet1", "A1", "unchanged").unwrap();
        wb.save(&path).unwrap();

        let mut lazy = Workbook::open(&path).unwrap();
        let error = lazy.insert_rows("Sheet1", 2, 1).unwrap_err();
        assert!(matches!(error, Error::XmlDeserialize(_)));
        assert!(lazy.raw_charts.is_empty());
        assert!(lazy.charts.is_empty());
        assert!(lazy
            .deferred_parts
            .has_category(crate::workbook::aux::AuxCategory::Charts));
        assert_eq!(
            lazy.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("unchanged".into())
        );
        assert!(!lazy.is_sheet_dirty(0));
        lazy.save(&path).unwrap();
        let reopened = Workbook::open(&path).unwrap();
        assert_eq!(
            reopened.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("unchanged".into())
        );
    }

    #[test]
    fn test_structural_edit_rejects_lazy_orphan_chart_before_mutation() {
        let mut workbook = Workbook::new();
        workbook.new_sheet("Sheet2").unwrap();
        workbook
            .add_chart(
                "Sheet1",
                "D1",
                "K10",
                &ChartConfig {
                    chart_type: crate::chart::ChartType::Col,
                    title: None,
                    series: vec![crate::chart::ChartSeries {
                        name: "Sheet1!$A$1".into(),
                        categories: "Sheet1!$A$1:$A$2".into(),
                        values: "Sheet1!$B$1:$B$2".into(),
                        x_values: None,
                        bubble_sizes: None,
                    }],
                    show_legend: false,
                    view_3d: None,
                },
            )
            .unwrap();
        let chart_path = workbook.charts[0].0.clone();
        let parseable_chart = concat!(
            r#"<chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            "<chart><plotArea/></chart></chartSpace>"
        )
        .as_bytes()
        .to_vec();
        let original = workbook.save_to_buffer().unwrap();
        let options = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut lazy = Workbook::open_from_buffer_with_options(&original, &options).unwrap();
        lazy.deferred_parts
            .remove_path(AuxCategory::Charts, &chart_path);
        assert!(lazy
            .deferred_parts
            .insert(chart_path.clone(), parseable_chart.clone()));
        let drawing_path = lazy
            .worksheet_rels
            .get(&0)
            .unwrap()
            .relationships
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::DRAWING)
            .map(|relationship| {
                resolve_relationship_target(&lazy.sheet_part_path(0), &relationship.target)
            })
            .unwrap();
        let drawing_rels_path = relationship_part_path(&drawing_path);
        let raw_drawing_rels = lazy
            .deferred_parts
            .get_path(AuxCategory::DrawingRels, &drawing_rels_path)
            .unwrap()
            .to_vec();
        assert!(lazy
            .deferred_parts
            .remove_path(AuxCategory::Drawings, &drawing_path)
            .is_some());
        let before = lazy.save_to_buffer().unwrap();

        let error = lazy.insert_rows("Sheet2", 1, 1).unwrap_err();

        assert!(
            matches!(&error, Error::InvalidArgument(message) if message.contains("unresolved chart owner")),
            "{error:?}"
        );
        assert!(lazy.charts.is_empty());
        assert!(lazy.raw_graph_parts.is_empty());
        assert_eq!(
            lazy.deferred_parts
                .get_path(AuxCategory::Charts, &chart_path),
            Some(parseable_chart.as_slice())
        );
        assert_eq!(
            lazy.deferred_parts
                .get_path(AuxCategory::DrawingRels, &drawing_rels_path),
            Some(raw_drawing_rels.as_slice())
        );
        assert!(lazy
            .deferred_parts
            .get_path(AuxCategory::Drawings, &drawing_path)
            .is_none());
        assert!(!lazy.is_sheet_dirty(0));
        assert!(!lazy.is_sheet_dirty(1));
        assert_eq!(lazy.save_to_buffer().unwrap(), before);
    }

    #[test]
    fn test_duplicate_row_hydrates_filtered_target_from_raw_xml() {
        let mut workbook = Workbook::new();
        workbook.new_sheet("Selected").unwrap();
        workbook.set_cell_value("Sheet1", "A1", "source").unwrap();
        let original = workbook.save_to_buffer().unwrap();
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .sheets(vec!["Selected".to_string()]);
        let mut opened = Workbook::open_from_buffer_with_options(&original, &options).unwrap();
        let selected_raw = zip_part(&original, "xl/worksheets/sheet2.xml");

        opened.duplicate_row("Sheet1", 1).unwrap();
        let saved = opened.save_to_buffer().unwrap();

        let duplicated = opened
            .worksheet_ref("Sheet1")
            .unwrap()
            .sheet_data
            .rows
            .iter()
            .find(|row| row.r == 2)
            .unwrap();
        assert_eq!(duplicated.cells[0].r, "A2");
        assert!(!opened.is_sheet_dirty(1));
        assert_eq!(zip_part(&saved, "xl/worksheets/sheet2.xml"), selected_raw);
    }

    #[test]
    fn test_structural_edit_preserves_untouched_sheet_raw_bytes() {
        let mut workbook = Workbook::new();
        workbook.new_sheet("Untouched").unwrap();
        let base = workbook.save_to_buffer().unwrap();
        let mut workbook = Workbook::open_from_buffer(&base).unwrap();
        let sheet_path = "xl/worksheets/sheet2.xml";
        let raw_sheet = String::from_utf8(zip_part(&base, sheet_path))
            .unwrap()
            .replacen(
                "</worksheet>",
                "<extLst><ext uri=\"{opaque}\"><opaque:payload xmlns:opaque=\"urn:test\"/></ext></extLst></worksheet>",
                1,
            )
            .into_bytes();
        workbook.raw_sheet_xml[1] = Some(raw_sheet.clone());

        workbook.insert_rows("Sheet1", 1, 1).unwrap();
        let saved = workbook.save_to_buffer().unwrap();

        assert!(!workbook.is_sheet_dirty(1));
        assert_eq!(zip_part(&saved, sheet_path), raw_sheet);
    }

    #[test]
    fn test_structural_edit_rejects_unparsed_eager_drawing_relationships_atomically() {
        let mut workbook = Workbook::new();
        workbook
            .add_image(
                "Sheet1",
                &crate::image::ImageConfig {
                    data: vec![0x89, 0x50, 0x4e, 0x47],
                    format: crate::image::ImageFormat::Png,
                    from_cell: "A1".to_string(),
                    width_px: 1,
                    height_px: 1,
                },
            )
            .unwrap();
        let base = workbook.save_to_buffer().unwrap();
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut opened = Workbook::open_from_buffer_with_options(&base, &options).unwrap();
        let drawing_idx = opened.worksheet_drawings[&0];
        let drawing_path = opened.drawings[drawing_idx].0.clone();
        let rels_path = relationship_part_path(&drawing_path);
        opened.drawing_rels.remove(&drawing_idx);
        opened
            .raw_graph_parts
            .insert(rels_path.clone(), b"<Relationships><Relationship".to_vec());
        let before = opened.save_to_buffer().unwrap();

        let error = opened.insert_rows("Sheet1", 1, 1).unwrap_err();

        assert!(matches!(error, Error::InvalidArgument(_)));
        assert_eq!(opened.save_to_buffer().unwrap(), before);
        assert_eq!(
            zip_part(&before, &rels_path),
            b"<Relationships><Relationship"
        );
        assert!(!opened.is_sheet_dirty(0));
    }

    #[test]
    fn test_structural_edit_rejects_unresolved_chart_owner_on_other_sheet() {
        let mut workbook = Workbook::new();
        workbook.new_sheet("Other").unwrap();
        workbook
            .add_chart(
                "Other",
                "D1",
                "K10",
                &ChartConfig {
                    chart_type: crate::chart::ChartType::Col,
                    title: None,
                    series: vec![crate::chart::ChartSeries {
                        name: "Other!$A$1".into(),
                        categories: "Other!$A$1:$A$2".into(),
                        values: "Other!$B$1:$B$2".into(),
                        x_values: None,
                        bubble_sizes: None,
                    }],
                    show_legend: false,
                    view_3d: None,
                },
            )
            .unwrap();
        let base = workbook.save_to_buffer().unwrap();
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut opened = Workbook::open_from_buffer_with_options(&base, &options).unwrap();
        let drawing_idx = opened.worksheet_drawings[&1];
        let drawing_path = opened.drawings[drawing_idx].0.clone();
        let rels_path = relationship_part_path(&drawing_path);
        opened.drawing_rels.remove(&drawing_idx);
        opened
            .raw_graph_parts
            .insert(rels_path.clone(), b"<Relationships><Relationship".to_vec());
        let before = opened.save_to_buffer().unwrap();

        let error = opened.insert_rows("Sheet1", 1, 1).unwrap_err();

        assert!(matches!(error, Error::InvalidArgument(_)));
        assert_eq!(opened.save_to_buffer().unwrap(), before);
        assert_eq!(
            zip_part(&before, &rels_path),
            b"<Relationships><Relationship"
        );
        assert!(!opened.is_sheet_dirty(0));
        assert!(!opened.is_sheet_dirty(1));
    }

    #[test]
    fn test_structural_edit_rejects_filtered_sheet_reference_overflow_atomically() {
        let mut workbook = Workbook::new();
        workbook.new_sheet("Filtered").unwrap();
        workbook
            .set_cell_formula("Filtered", "A1", "Sheet1!XFD1")
            .unwrap();
        let original = workbook.save_to_buffer().unwrap();
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .sheets(vec!["Sheet1".to_string()]);
        let mut opened = Workbook::open_from_buffer_with_options(&original, &options).unwrap();
        let before = opened.save_to_buffer().unwrap();

        let error = opened.insert_cols("Sheet1", "A", 1).unwrap_err();

        assert!(matches!(error, Error::InvalidColumnNumber(_)));
        assert_eq!(opened.save_to_buffer().unwrap(), before);
        assert!(!opened.is_sheet_dirty(0));
        assert!(!opened.is_sheet_dirty(1));
    }

    #[test]
    fn test_delete_sheet_rejects_unparsed_eager_drawing_relationships_atomically() {
        let mut workbook = Workbook::new();
        workbook.new_sheet("Survivor").unwrap();
        workbook
            .add_image(
                "Sheet1",
                &crate::image::ImageConfig {
                    data: vec![0x89, 0x50, 0x4e, 0x47],
                    format: crate::image::ImageFormat::Png,
                    from_cell: "A1".to_string(),
                    width_px: 1,
                    height_px: 1,
                },
            )
            .unwrap();
        let base = workbook.save_to_buffer().unwrap();
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut opened = Workbook::open_from_buffer_with_options(&base, &options).unwrap();
        let drawing_idx = opened.worksheet_drawings[&0];
        let drawing_path = opened.drawings[drawing_idx].0.clone();
        let rels_path = relationship_part_path(&drawing_path);
        opened.drawing_rels.remove(&drawing_idx);
        opened
            .raw_graph_parts
            .insert(rels_path.clone(), b"<Relationships><Relationship".to_vec());
        let before = opened.save_to_buffer().unwrap();

        let error = opened.delete_sheet("Sheet1").unwrap_err();

        assert!(matches!(error, Error::InvalidArgument(_)));
        assert_eq!(opened.sheet_names(), vec!["Sheet1", "Survivor"]);
        assert_eq!(opened.save_to_buffer().unwrap(), before);
        assert_eq!(
            zip_part(&before, &rels_path),
            b"<Relationships><Relationship"
        );
    }

    #[test]
    fn test_delete_graph_free_sheet_ignores_unrelated_deferred_drawing() {
        let mut workbook = Workbook::new();
        workbook.new_sheet("Survivor").unwrap();
        let original = workbook.save_to_buffer().unwrap();
        let options = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut lazy = Workbook::open_from_buffer_with_options(&original, &options).unwrap();
        let drawing_path = "xl/drawings/drawing99.xml";
        let raw_drawing = b"<xdr:wsDr><opaque></xdr:wsDr>".to_vec();
        assert!(lazy
            .deferred_parts
            .insert(drawing_path.to_string(), raw_drawing.clone()));

        lazy.delete_sheet("Sheet1").unwrap();
        let saved = lazy.save_to_buffer().unwrap();

        assert_eq!(lazy.sheet_names(), vec!["Survivor"]);
        assert_eq!(zip_part(&saved, drawing_path), raw_drawing);
    }

    #[test]
    fn test_delete_graph_free_sheet_preserves_opaque_survivor_graph_closure() {
        let mut workbook = Workbook::new();
        workbook.new_sheet("Survivor").unwrap();
        workbook
            .add_image(
                "Survivor",
                &crate::image::ImageConfig {
                    data: vec![0x89, 0x50, 0x4e, 0x47],
                    format: crate::image::ImageFormat::Png,
                    from_cell: "A1".to_string(),
                    width_px: 1,
                    height_px: 1,
                },
            )
            .unwrap();
        let original = workbook.save_to_buffer().unwrap();
        let options = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut lazy = Workbook::open_from_buffer_with_options(&original, &options).unwrap();
        let sheet_path = lazy.sheet_part_path(1);
        let drawing_path = lazy
            .worksheet_rels
            .get(&1)
            .unwrap()
            .relationships
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::DRAWING)
            .map(|relationship| resolve_relationship_target(&sheet_path, &relationship.target))
            .unwrap();
        let drawing_rels_path = relationship_part_path(&drawing_path);
        let drawing_rels = quick_xml::de::from_reader::<_, Relationships>(
            lazy.deferred_parts
                .get_path(AuxCategory::DrawingRels, &drawing_rels_path)
                .unwrap(),
        )
        .unwrap();
        let image_path = drawing_rels
            .relationships
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::IMAGE)
            .map(|relationship| resolve_relationship_target(&drawing_path, &relationship.target))
            .unwrap();
        let raw_drawing = b"<xdr:wsDr><opaque></xdr:wsDr>".to_vec();
        lazy.deferred_parts
            .remove_path(AuxCategory::Drawings, &drawing_path);
        assert!(lazy
            .deferred_parts
            .insert(drawing_path.clone(), raw_drawing.clone()));
        let raw_sheet_rels = zip_part(&original, &relationship_part_path(&sheet_path));
        let raw_drawing_rels = zip_part(&original, &drawing_rels_path);
        let raw_image = zip_part(&original, &image_path);

        lazy.delete_sheet("Sheet1").unwrap();
        let saved = lazy.save_to_buffer().unwrap();

        assert_eq!(lazy.sheet_names(), vec!["Survivor"]);
        assert_eq!(zip_part(&saved, &drawing_path), raw_drawing);
        assert_eq!(zip_part(&saved, &drawing_rels_path), raw_drawing_rels);
        assert_eq!(zip_part(&saved, &image_path), raw_image);
        assert_eq!(
            zip_part(&saved, "xl/worksheets/_rels/sheet2.xml.rels"),
            raw_sheet_rels
        );
        assert_eq!(zip_entry_count(&saved, &drawing_path), 1);
        assert_package_integrity(&saved);
    }

    #[test]
    fn test_structural_edit_reference_overflow_leaves_workbook_unchanged() {
        let mut wb = Workbook::new();
        wb.set_cell_formula("Sheet1", "A1", "XFD1").unwrap();
        wb.set_auto_filter("Sheet1", "A1:XFD1").unwrap();

        let error = wb.insert_cols("Sheet1", "A", 1).unwrap_err();
        assert!(matches!(error, Error::InvalidColumnNumber(_)));
        match wb.get_cell_value("Sheet1", "A1").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "XFD1"),
            value => panic!("expected formula, got {value:?}"),
        }
        assert_eq!(
            wb.worksheet_ref("Sheet1")
                .unwrap()
                .auto_filter
                .as_ref()
                .unwrap()
                .reference,
            "A1:XFD1"
        );
    }

    #[test]
    fn test_workbook_insert_rows_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.insert_rows("NoSheet", 1, 1);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_remove_row() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "first").unwrap();
        wb.set_cell_value("Sheet1", "A2", "second").unwrap();
        wb.set_cell_value("Sheet1", "A3", "third").unwrap();
        wb.remove_row("Sheet1", 2).unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("first".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "A2").unwrap(),
            CellValue::String("third".to_string())
        );
    }

    #[test]
    fn test_workbook_remove_row_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.remove_row("NoSheet", 1);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_duplicate_row() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "original").unwrap();
        wb.duplicate_row("Sheet1", 1).unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("original".to_string())
        );
        // The duplicated row at row 2 has the same SST index.
        assert_eq!(
            wb.get_cell_value("Sheet1", "A2").unwrap(),
            CellValue::String("original".to_string())
        );
    }

    #[test]
    fn test_duplicate_row_adjusts_only_relative_formula_rows() {
        let mut wb = Workbook::new();
        wb.new_sheet("Other").unwrap();
        wb.set_cell_formula("Sheet1", "B1", "A1+$A$1+Other!A1+Other!A2+Other!$A$1")
            .unwrap();

        wb.duplicate_row("Sheet1", 1).unwrap();

        match wb.get_cell_value("Sheet1", "B2").unwrap() {
            CellValue::Formula { expr, .. } => {
                assert_eq!(expr, "A2+$A$1+Other!A2+Other!A3+Other!$A$1");
            }
            value => panic!("expected formula, got {value:?}"),
        }
    }

    #[test]
    fn test_duplicate_row_shifts_existing_merge_and_hyperlink_once() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "source").unwrap();
        wb.merge_cells("Sheet1", "C3", "D3").unwrap();
        wb.set_cell_hyperlink(
            "Sheet1",
            "A3",
            crate::hyperlink::HyperlinkType::Internal("Sheet1!B3".into()),
            None,
            None,
        )
        .unwrap();

        wb.duplicate_row("Sheet1", 1).unwrap();

        assert_eq!(wb.get_merge_cells("Sheet1").unwrap(), vec!["C4:D4"]);
        let hyperlink = wb.get_cell_hyperlink("Sheet1", "A4").unwrap().unwrap();
        assert_eq!(
            hyperlink.link_type,
            crate::hyperlink::HyperlinkType::Internal("Sheet1!B4".into())
        );
    }

    #[test]
    fn test_duplicate_row_formula_overflow_leaves_workbook_unchanged() {
        let mut wb = Workbook::new();
        wb.new_sheet("Other").unwrap();
        wb.set_cell_formula("Sheet1", "A1", "Other!A1048576")
            .unwrap();

        let error = wb.duplicate_row("Sheet1", 1).unwrap_err();
        assert!(matches!(error, Error::InvalidRowNumber(1_048_577)));
        match wb.get_cell_value("Sheet1", "A1").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "Other!A1048576"),
            value => panic!("expected formula, got {value:?}"),
        }
        assert_eq!(wb.get_cell_value("Sheet1", "A2").unwrap(), CellValue::Empty);
    }

    #[test]
    fn test_workbook_set_and_get_row_height() {
        let mut wb = Workbook::new();
        wb.set_row_height("Sheet1", 3, 25.0).unwrap();
        assert_eq!(wb.get_row_height("Sheet1", 3).unwrap(), Some(25.0));
    }

    #[test]
    fn test_workbook_get_row_height_sheet_not_found() {
        let wb = Workbook::new();
        let result = wb.get_row_height("NoSheet", 1);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_set_row_visible() {
        let mut wb = Workbook::new();
        wb.set_row_visible("Sheet1", 1, false).unwrap();
    }

    #[test]
    fn test_workbook_set_row_visible_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.set_row_visible("NoSheet", 1, false);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_set_and_get_col_width() {
        let mut wb = Workbook::new();
        wb.set_col_width("Sheet1", "A", 18.0).unwrap();
        assert_eq!(wb.get_col_width("Sheet1", "A").unwrap(), Some(18.0));
    }

    #[test]
    fn test_workbook_get_col_width_sheet_not_found() {
        let wb = Workbook::new();
        let result = wb.get_col_width("NoSheet", "A");
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_set_col_visible() {
        let mut wb = Workbook::new();
        wb.set_col_visible("Sheet1", "B", false).unwrap();
    }

    #[test]
    fn test_workbook_set_col_visible_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.set_col_visible("NoSheet", "A", false);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_insert_cols() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "a").unwrap();
        wb.set_cell_value("Sheet1", "B1", "b").unwrap();
        wb.insert_cols("Sheet1", "B", 1).unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("a".to_string())
        );
        assert_eq!(wb.get_cell_value("Sheet1", "B1").unwrap(), CellValue::Empty);
        assert_eq!(
            wb.get_cell_value("Sheet1", "C1").unwrap(),
            CellValue::String("b".to_string())
        );
    }

    #[test]
    fn test_workbook_insert_cols_updates_formula_and_ranges() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "D1",
            CellValue::Formula {
                expr: "SUM(A1:B1)".to_string(),
                result: None,
            },
        )
        .unwrap();
        wb.add_data_validation(
            "Sheet1",
            &crate::validation::DataValidationConfig::whole_number("B2:C3", 1, 9),
        )
        .unwrap();
        wb.set_auto_filter("Sheet1", "A1:C10").unwrap();
        wb.merge_cells("Sheet1", "B3", "C4").unwrap();

        wb.insert_cols("Sheet1", "B", 2).unwrap();

        match wb.get_cell_value("Sheet1", "F1").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "SUM(A1:D1)"),
            other => panic!("expected formula, got {other:?}"),
        }

        let validations = wb.get_data_validations("Sheet1").unwrap();
        assert_eq!(validations.len(), 1);
        assert_eq!(validations[0].sqref, "D2:E3");

        let merges = wb.get_merge_cells("Sheet1").unwrap();
        assert_eq!(merges, vec!["D3:E4".to_string()]);

        let ws = wb.worksheet_ref("Sheet1").unwrap();
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:E10");
    }

    #[test]
    fn test_workbook_insert_cols_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.insert_cols("NoSheet", "A", 1);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_remove_col() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "a").unwrap();
        wb.set_cell_value("Sheet1", "B1", "b").unwrap();
        wb.set_cell_value("Sheet1", "C1", "c").unwrap();
        wb.remove_col("Sheet1", "B").unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("a".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "B1").unwrap(),
            CellValue::String("c".to_string())
        );
    }

    #[test]
    fn test_workbook_remove_col_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.remove_col("NoSheet", "A");
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_new_stream_writer_validates_name() {
        let wb = Workbook::new();
        let result = wb.new_stream_writer("Bad[Name");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::InvalidSheetName(_)));
    }

    #[test]
    fn test_new_stream_writer_rejects_duplicate() {
        let wb = Workbook::new();
        let result = wb.new_stream_writer("Sheet1");
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::SheetAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_new_stream_writer_valid_name() {
        let wb = Workbook::new();
        let sw = wb.new_stream_writer("StreamSheet").unwrap();
        assert_eq!(sw.sheet_name(), "StreamSheet");
    }

    #[test]
    fn test_apply_stream_writer_adds_sheet() {
        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("StreamSheet").unwrap();
        sw.write_row(1, &[CellValue::from("Hello"), CellValue::from(42)])
            .unwrap();
        let idx = wb.apply_stream_writer(sw).unwrap();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_names(), vec!["Sheet1", "StreamSheet"]);
    }

    #[test]
    fn test_apply_stream_writer_uses_inline_strings() {
        // Streamed sheets use inline strings, not the shared string table.
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Existing").unwrap();
        let sst_before = wb.sst_runtime.len();

        let mut sw = wb.new_stream_writer("StreamSheet").unwrap();
        sw.write_row(1, &[CellValue::from("New"), CellValue::from("Existing")])
            .unwrap();
        wb.apply_stream_writer(sw).unwrap();

        // SST should not grow because streamed sheets use inline strings.
        assert_eq!(wb.sst_runtime.len(), sst_before);
    }

    #[test]
    fn test_stream_writer_save_and_reopen() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_test.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Normal").unwrap();

        let mut sw = wb.new_stream_writer("Streamed").unwrap();
        sw.write_row(1, &[CellValue::from("Name"), CellValue::from("Value")])
            .unwrap();
        sw.write_row(2, &[CellValue::from("Alice"), CellValue::from(100)])
            .unwrap();
        sw.write_row(3, &[CellValue::from("Bob"), CellValue::from(200)])
            .unwrap();
        wb.apply_stream_writer(sw).unwrap();

        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1", "Streamed"]);
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Normal".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Streamed", "A1").unwrap(),
            CellValue::String("Name".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Streamed", "B2").unwrap(),
            CellValue::Number(100.0)
        );
        assert_eq!(
            wb2.get_cell_value("Streamed", "A3").unwrap(),
            CellValue::String("Bob".to_string())
        );
    }

    #[test]
    fn test_workbook_get_rows_empty_sheet() {
        let wb = Workbook::new();
        let rows = wb.get_rows("Sheet1").unwrap();
        assert!(rows.is_empty());
    }

    #[test]
    fn test_workbook_get_rows_with_data() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Name").unwrap();
        wb.set_cell_value("Sheet1", "B1", 42.0).unwrap();
        wb.set_cell_value("Sheet1", "A2", "Alice").unwrap();
        wb.set_cell_value("Sheet1", "B2", true).unwrap();

        let rows = wb.get_rows("Sheet1").unwrap();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].0, 1);
        assert_eq!(rows[0].1.len(), 2);
        assert_eq!(rows[0].1[0].0, 1);
        assert_eq!(rows[0].1[0].1, CellValue::String("Name".to_string()));
        assert_eq!(rows[0].1[1].0, 2);
        assert_eq!(rows[0].1[1].1, CellValue::Number(42.0));
        assert_eq!(rows[1].0, 2);
        assert_eq!(rows[1].1[0].1, CellValue::String("Alice".to_string()));
        assert_eq!(rows[1].1[1].1, CellValue::Bool(true));
    }

    #[test]
    fn test_workbook_get_rows_sheet_not_found() {
        let wb = Workbook::new();
        assert!(wb.get_rows("NoSheet").is_err());
    }

    #[test]
    fn test_workbook_get_cols_empty_sheet() {
        let wb = Workbook::new();
        let cols = wb.get_cols("Sheet1").unwrap();
        assert!(cols.is_empty());
    }

    #[test]
    fn test_workbook_get_cols_with_data() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Name").unwrap();
        wb.set_cell_value("Sheet1", "B1", 42.0).unwrap();
        wb.set_cell_value("Sheet1", "A2", "Alice").unwrap();
        wb.set_cell_value("Sheet1", "B2", 30.0).unwrap();

        let cols = wb.get_cols("Sheet1").unwrap();
        assert_eq!(cols.len(), 2);
        assert_eq!(cols[0].0, "A");
        assert_eq!(cols[0].1.len(), 2);
        assert_eq!(cols[0].1[0], (1, CellValue::String("Name".to_string())));
        assert_eq!(cols[0].1[1], (2, CellValue::String("Alice".to_string())));
        assert_eq!(cols[1].0, "B");
        assert_eq!(cols[1].1[0], (1, CellValue::Number(42.0)));
        assert_eq!(cols[1].1[1], (2, CellValue::Number(30.0)));
    }

    #[test]
    fn test_workbook_get_cols_sheet_not_found() {
        let wb = Workbook::new();
        assert!(wb.get_cols("NoSheet").is_err());
    }

    #[test]
    fn test_streamed_sheet_cells_empty_before_save() {
        // Streamed sheet data lives in a temp file, not in the WorksheetXml.
        // Reading cells before save returns Empty.
        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("Streamed").unwrap();
        sw.write_row(1, &[CellValue::from("Name"), CellValue::from("Age")])
            .unwrap();
        wb.apply_stream_writer(sw).unwrap();

        assert_eq!(
            wb.get_cell_value("Streamed", "A1").unwrap(),
            CellValue::Empty
        );
        assert_eq!(
            wb.get_cell_value("Streamed", "B1").unwrap(),
            CellValue::Empty
        );
    }

    #[test]
    fn test_streamed_sheet_readable_after_save_reopen() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_reopen.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("Streamed").unwrap();
        sw.write_row(1, &[CellValue::from("Name"), CellValue::from("Age")])
            .unwrap();
        sw.write_row(2, &[CellValue::from("Alice"), CellValue::from(30)])
            .unwrap();
        wb.apply_stream_writer(sw).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(
            wb2.get_cell_value("Streamed", "A1").unwrap(),
            CellValue::String("Name".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Streamed", "B1").unwrap(),
            CellValue::String("Age".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Streamed", "A2").unwrap(),
            CellValue::String("Alice".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Streamed", "B2").unwrap(),
            CellValue::Number(30.0)
        );
    }

    #[test]
    fn test_workbook_get_rows_roundtrip_save_open() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "hello").unwrap();
        wb.set_cell_value("Sheet1", "B1", 99.0).unwrap();
        wb.set_cell_value("Sheet1", "A2", true).unwrap();

        let tmp = std::env::temp_dir().join("test_get_rows_roundtrip.xlsx");
        wb.save(&tmp).unwrap();

        let wb2 = Workbook::open(&tmp).unwrap();
        let rows = wb2.get_rows("Sheet1").unwrap();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].1[0].1, CellValue::String("hello".to_string()));
        assert_eq!(rows[0].1[1].1, CellValue::Number(99.0));
        assert_eq!(rows[1].1[0].1, CellValue::Bool(true));

        let _ = std::fs::remove_file(&tmp);
    }

    #[test]
    fn test_stream_save_reopen_basic() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_basic.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("Optimized").unwrap();
        sw.write_row(1, &[CellValue::from("Hello"), CellValue::from(42)])
            .unwrap();
        sw.write_row(2, &[CellValue::from("World"), CellValue::from(99)])
            .unwrap();
        let idx = wb.apply_stream_writer(sw).unwrap();
        assert_eq!(idx, 1);

        wb.save(&path).unwrap();
        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(
            wb2.get_cell_value("Optimized", "A1").unwrap(),
            CellValue::String("Hello".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Optimized", "B1").unwrap(),
            CellValue::Number(42.0)
        );
        assert_eq!(
            wb2.get_cell_value("Optimized", "A2").unwrap(),
            CellValue::String("World".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Optimized", "B2").unwrap(),
            CellValue::Number(99.0)
        );
    }

    #[test]
    fn test_stream_save_reopen_all_types() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_types.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("Types").unwrap();
        sw.write_row(
            1,
            &[
                CellValue::from("text"),
                CellValue::from(42),
                CellValue::from(3.14),
                CellValue::from(true),
                CellValue::Formula {
                    expr: "SUM(B1:C1)".to_string(),
                    result: None,
                },
                CellValue::Error("#N/A".to_string()),
                CellValue::Empty,
            ],
        )
        .unwrap();
        wb.apply_stream_writer(sw).unwrap();

        wb.save(&path).unwrap();
        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(
            wb2.get_cell_value("Types", "A1").unwrap(),
            CellValue::String("text".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Types", "B1").unwrap(),
            CellValue::Number(42.0)
        );
        assert_eq!(
            wb2.get_cell_value("Types", "D1").unwrap(),
            CellValue::Bool(true)
        );
        match wb2.get_cell_value("Types", "E1").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "SUM(B1:C1)"),
            other => panic!("expected formula, got {other:?}"),
        }
        assert_eq!(
            wb2.get_cell_value("Types", "F1").unwrap(),
            CellValue::Error("#N/A".to_string())
        );
        assert_eq!(wb2.get_cell_value("Types", "G1").unwrap(), CellValue::Empty);
    }

    #[test]
    fn test_apply_stream_optimized_save_reopen() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_optimized.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Normal").unwrap();

        let mut sw = wb.new_stream_writer("Fast").unwrap();
        sw.write_row(1, &[CellValue::from("Name"), CellValue::from("Value")])
            .unwrap();
        sw.write_row(2, &[CellValue::from("Alice"), CellValue::from(100)])
            .unwrap();
        sw.write_row(3, &[CellValue::from("Bob"), CellValue::from(200)])
            .unwrap();
        wb.apply_stream_writer(sw).unwrap();

        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1", "Fast"]);
        assert_eq!(
            wb2.get_cell_value("Fast", "A1").unwrap(),
            CellValue::String("Name".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Fast", "B2").unwrap(),
            CellValue::Number(100.0)
        );
        assert_eq!(
            wb2.get_cell_value("Fast", "A3").unwrap(),
            CellValue::String("Bob".to_string())
        );
    }

    #[test]
    fn test_stream_freeze_panes_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_freeze.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("FreezeSheet").unwrap();
        sw.set_freeze_panes("B3").unwrap();
        sw.write_row(1, &[CellValue::from("A"), CellValue::from("B")])
            .unwrap();
        sw.write_row(2, &[CellValue::from("C"), CellValue::from("D")])
            .unwrap();
        wb.apply_stream_writer(sw).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(
            wb2.get_panes("FreezeSheet").unwrap(),
            Some("B3".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("FreezeSheet", "A1").unwrap(),
            CellValue::String("A".to_string())
        );
    }

    #[test]
    fn test_stream_merge_cells_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_merge.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("MergeSheet").unwrap();
        sw.add_merge_cell("A1:C1").unwrap();
        sw.add_merge_cell("A3:B4").unwrap();
        sw.write_row(1, &[CellValue::from("Header")]).unwrap();
        sw.write_row(2, &[CellValue::from("Data")]).unwrap();
        wb.apply_stream_writer(sw).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let merges = wb2.get_merge_cells("MergeSheet").unwrap();
        assert!(merges.contains(&"A1:C1".to_string()));
        assert!(merges.contains(&"A3:B4".to_string()));
        assert_eq!(
            wb2.get_cell_value("MergeSheet", "A1").unwrap(),
            CellValue::String("Header".to_string())
        );
    }

    #[test]
    fn test_stream_col_widths_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_colw.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("ColSheet").unwrap();
        sw.set_col_width(1, 25.0).unwrap();
        sw.set_col_width(2, 12.5).unwrap();
        sw.write_row(1, &[CellValue::from("Wide"), CellValue::from("Narrow")])
            .unwrap();
        wb.apply_stream_writer(sw).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let w1 = wb2.get_col_width("ColSheet", "A").unwrap().unwrap();
        let w2 = wb2.get_col_width("ColSheet", "B").unwrap().unwrap();
        assert!((w1 - 25.0).abs() < 0.01);
        assert!((w2 - 12.5).abs() < 0.01);
    }

    #[test]
    fn test_stream_multiple_sheets() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_multi.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Normal").unwrap();

        let mut sw1 = wb.new_stream_writer("Stream1").unwrap();
        sw1.write_row(1, &[CellValue::from("S1R1")]).unwrap();
        sw1.write_row(2, &[CellValue::from("S1R2")]).unwrap();
        wb.apply_stream_writer(sw1).unwrap();

        let mut sw2 = wb.new_stream_writer("Stream2").unwrap();
        sw2.write_row(1, &[CellValue::from("S2R1")]).unwrap();
        wb.apply_stream_writer(sw2).unwrap();

        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1", "Stream1", "Stream2"]);
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Normal".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Stream1", "A1").unwrap(),
            CellValue::String("S1R1".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Stream1", "A2").unwrap(),
            CellValue::String("S1R2".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Stream2", "A1").unwrap(),
            CellValue::String("S2R1".to_string())
        );
    }

    #[test]
    fn test_stream_delete_sheet() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_delete.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("ToDelete").unwrap();
        sw.write_row(1, &[CellValue::from("Gone")]).unwrap();
        wb.apply_stream_writer(sw).unwrap();

        let mut sw2 = wb.new_stream_writer("Kept").unwrap();
        sw2.write_row(1, &[CellValue::from("Stays")]).unwrap();
        wb.apply_stream_writer(sw2).unwrap();

        wb.delete_sheet("ToDelete").unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1", "Kept"]);
        assert_eq!(
            wb2.get_cell_value("Kept", "A1").unwrap(),
            CellValue::String("Stays".to_string())
        );
    }

    #[test]
    fn test_stream_combined_features_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_combined.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("Combined").unwrap();
        sw.set_freeze_panes("A2").unwrap();
        sw.set_col_width(1, 30.0).unwrap();
        sw.set_col_width_range(2, 3, 15.0).unwrap();
        sw.add_merge_cell("B1:C1").unwrap();
        sw.write_row(
            1,
            &[
                CellValue::from("Name"),
                CellValue::from("Merged Header"),
                CellValue::Empty,
            ],
        )
        .unwrap();
        sw.write_row(
            2,
            &[
                CellValue::from("Alice"),
                CellValue::from(100),
                CellValue::from(true),
            ],
        )
        .unwrap();
        wb.apply_stream_writer(sw).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.get_panes("Combined").unwrap(), Some("A2".to_string()));
        let merges = wb2.get_merge_cells("Combined").unwrap();
        assert!(merges.contains(&"B1:C1".to_string()));
        let w1 = wb2.get_col_width("Combined", "A").unwrap().unwrap();
        assert!((w1 - 30.0).abs() < 0.01);
        assert_eq!(
            wb2.get_cell_value("Combined", "A1").unwrap(),
            CellValue::String("Name".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Combined", "B2").unwrap(),
            CellValue::Number(100.0)
        );
        assert_eq!(
            wb2.get_cell_value("Combined", "C2").unwrap(),
            CellValue::Bool(true)
        );
    }

    // --- Regression tests for P1 bugs ---

    #[test]
    fn test_stream_formula_result_types_roundtrip() {
        // Regression: formula cached results must preserve their type via the
        // cell t attribute (t="str", t="b", t="e"). Without it, string results
        // are dropped and bool results are decoded as Number(1.0).
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_formula_types.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("Formulas").unwrap();
        sw.write_row(
            1,
            &[
                CellValue::Formula {
                    expr: "A2&B2".to_string(),
                    result: Some(Box::new(CellValue::String("hello".to_string()))),
                },
                CellValue::Formula {
                    expr: "A2>0".to_string(),
                    result: Some(Box::new(CellValue::Bool(true))),
                },
                CellValue::Formula {
                    expr: "1/0".to_string(),
                    result: Some(Box::new(CellValue::Error("#DIV/0!".to_string()))),
                },
                CellValue::Formula {
                    expr: "SUM(A2:A10)".to_string(),
                    result: Some(Box::new(CellValue::Number(55.0))),
                },
            ],
        )
        .unwrap();
        wb.apply_stream_writer(sw).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        // String result
        assert_eq!(
            wb2.get_cell_value("Formulas", "A1").unwrap(),
            CellValue::Formula {
                expr: "A2&B2".to_string(),
                result: Some(Box::new(CellValue::String("hello".to_string()))),
            }
        );
        // Bool result
        assert_eq!(
            wb2.get_cell_value("Formulas", "B1").unwrap(),
            CellValue::Formula {
                expr: "A2>0".to_string(),
                result: Some(Box::new(CellValue::Bool(true))),
            }
        );
        // Error result
        assert_eq!(
            wb2.get_cell_value("Formulas", "C1").unwrap(),
            CellValue::Formula {
                expr: "1/0".to_string(),
                result: Some(Box::new(CellValue::Error("#DIV/0!".to_string()))),
            }
        );
        // Numeric result
        assert_eq!(
            wb2.get_cell_value("Formulas", "D1").unwrap(),
            CellValue::Formula {
                expr: "SUM(A2:A10)".to_string(),
                result: Some(Box::new(CellValue::Number(55.0))),
            }
        );
    }

    #[test]
    fn test_stream_edit_after_apply_takes_effect() {
        // Regression: edits via set_cell_value after apply_stream_writer must
        // not be silently ignored. The edit invalidates the streamed data so
        // the normal WorksheetXml serialization path is used on save.
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_edit_after.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("S").unwrap();
        sw.write_row(1, &[CellValue::from("old")]).unwrap();
        wb.apply_stream_writer(sw).unwrap();

        // Edit the streamed sheet: this should invalidate streamed data.
        wb.set_cell_value("S", "A1", "new").unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(
            wb2.get_cell_value("S", "A1").unwrap(),
            CellValue::String("new".to_string())
        );
    }

    #[test]
    fn test_stream_copy_sheet_preserves_data() {
        // Regression: copy_sheet must clone the streamed payload so both
        // source and target sheets have the streamed data on save.
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("stream_copy.xlsx");

        let mut wb = Workbook::new();
        let mut sw = wb.new_stream_writer("Src").unwrap();
        sw.write_row(1, &[CellValue::from("x")]).unwrap();
        sw.write_row(2, &[CellValue::from("y")]).unwrap();
        wb.apply_stream_writer(sw).unwrap();

        wb.copy_sheet("Src", "Dst").unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(
            wb2.get_cell_value("Src", "A1").unwrap(),
            CellValue::String("x".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Src", "A2").unwrap(),
            CellValue::String("y".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Dst", "A1").unwrap(),
            CellValue::String("x".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Dst", "A2").unwrap(),
            CellValue::String("y".to_string())
        );
    }
}
