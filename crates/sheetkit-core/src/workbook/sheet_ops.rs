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
    /// Return the names of all sheets in workbook order.
    pub fn sheet_names(&self) -> Vec<&str> {
        self.worksheets
            .iter()
            .map(|(name, _)| name.as_str())
            .collect()
    }

    /// Create a new empty sheet with the given name. Returns the 0-based sheet index.
    pub fn new_sheet(&mut self, name: &str) -> Result<usize> {
        let idx = crate::sheet::add_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            name,
            WorksheetXml::default(),
        )?;
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
        self.rebuild_sheet_index();
        Ok(idx)
    }

    /// Delete a sheet by name.
    pub fn delete_sheet(&mut self, name: &str) -> Result<()> {
        let idx = self.sheet_index(name)?;
        self.assert_parallel_vecs_in_sync();

        crate::sheet::delete_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            name,
        )?;

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
        self.rebuild_sheet_index();
        Ok(())
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
        // Resolve the source index before copy_sheet changes the array.
        let src_idx = self.sheet_index(source)?;
        // Hydrate the source sheet so copy_sheet clones the real data,
        // not an empty default.
        self.ensure_hydrated(src_idx)?;
        let idx = crate::sheet::copy_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            source,
            target,
        )?;
        if self.sheet_comments.len() < self.worksheets.len() {
            self.sheet_comments.push(None);
        }
        let source_sparklines = self
            .sheet_sparklines
            .get(src_idx)
            .cloned()
            .unwrap_or_default();
        if self.sheet_sparklines.len() < self.worksheets.len() {
            self.sheet_sparklines.push(source_sparklines);
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
        // Copy streamed data if the source sheet was streamed.
        if let Some(src_streamed) = self.streamed_sheets.get(&src_idx) {
            let cloned = src_streamed.try_clone()?;
            self.streamed_sheets.insert(idx, cloned);
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
        let idx = crate::sheet::add_sheet(
            &mut self.workbook_xml,
            &mut self.workbook_rels,
            &mut self.content_types,
            &mut self.worksheets,
            &sheet_name,
            WorksheetXml::default(),
        )?;
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

        // Validate deferred XML without consuming passthrough bytes. A failed
        // edit must not turn an untouched lazy workbook into a dirty one.
        self.validate_deferred_drawing_and_chart_parts(target_sheet_idx, shift_cell)?;
        self.validate_reference_shift(target_sheet_idx, shift_cell)?;
        for idx in 0..self.worksheets.len() {
            self.ensure_hydrated(idx)?;
        }
        self.hydrate_drawings();
        self.hydrate_tables();
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
            for (_, bytes) in entries {
                let mut chart =
                    quick_xml::de::from_str::<ChartSpace>(&String::from_utf8_lossy(bytes))
                        .map_err(|error| Error::XmlDeserialize(error.to_string()))?;
                // A deferred chart's drawing relationship is not yet hydrated.
                // Treat unqualified formulas as target-owned during dry-run so
                // overflow cannot escape validation; qualified formulas retain
                // their exact sheet selection.
                visit_chart_references(&mut chart, |formula| validate(formula, target_sheet_idx))?;
            }
        }
        Ok(())
    }

    fn chart_owner_sheet_idx(&self, chart_path: &str) -> usize {
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
            .unwrap_or(0)
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

        for (owner_sheet_idx, (_, worksheet)) in self.worksheets.iter().enumerate() {
            let deferred;
            let ws = if let Some(ws) = worksheet.get() {
                ws
            } else {
                let bytes = self.raw_sheet_xml[owner_sheet_idx]
                    .as_deref()
                    .ok_or_else(|| {
                        Error::Internal("worksheet has no materialized or deferred data".into())
                    })?;
                deferred = io::deserialize_worksheet_xml(bytes)?;
                &deferred
            };
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
            let owner_sheet_idx = self.chart_owner_sheet_idx(path);
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
        let mut worksheet = self.worksheet_ref_by_index(sheet_idx)?.clone();
        edit(&mut worksheet)
    }

    fn preflight_duplicate_formula_copy(&self, sheet_idx: usize, row: u32) -> Result<()> {
        let worksheet = self.worksheet_ref_by_index(sheet_idx)?;
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
            let ws = self.worksheet_mut_by_index(owner_sheet_idx)?;

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

        for chart_idx in 0..self.charts.len() {
            let owner_sheet_idx = self.chart_owner_sheet_idx(&self.charts[chart_idx].0);
            let chart = &mut self.charts[chart_idx].1;
            visit_chart_references(chart, |formula| {
                *formula = shift_references_for_owner(
                    formula,
                    owner_sheet_idx,
                    sheet_idx,
                    &target_sheet_name_lowercase,
                    shift_cell,
                )?;
                Ok(())
            })?;
        }

        // Drawing anchors attached to this sheet.
        if let Some(&drawing_idx) = self.worksheet_drawings.get(&sheet_idx) {
            if let Some((_, drawing)) = self.drawings.get_mut(drawing_idx) {
                for anchor in &mut drawing.one_cell_anchors {
                    let (new_col, new_row) = shift_cell(anchor.from.col + 1, anchor.from.row + 1);
                    anchor.from.col = new_col - 1;
                    anchor.from.row = new_row - 1;
                }
                for anchor in &mut drawing.two_cell_anchors {
                    let (from_col, from_row) = shift_cell(anchor.from.col + 1, anchor.from.row + 1);
                    anchor.from.col = from_col - 1;
                    anchor.from.row = from_row - 1;
                    let (to_col, to_row) = shift_cell(anchor.to.col + 1, anchor.to.row + 1);
                    anchor.to.col = to_col - 1;
                    anchor.to.row = to_row - 1;
                }
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
        let drawing_path = format!("xl/drawings/drawing{}.xml", idx + 1);
        self.drawings.push((drawing_path, WsDr::default()));
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
        let drawing_rel_target = format!("../drawings/drawing{}.xml", idx + 1);
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
            part_name: format!("/xl/drawings/drawing{}.xml", idx + 1),
            content_type: mime_types::DRAWING.to_string(),
        });

        idx
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
    use tempfile::TempDir;

    #[test]
    fn test_new_sheet_basic() {
        let mut wb = Workbook::new();
        let idx = wb.new_sheet("Sheet2").unwrap();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_names(), vec!["Sheet1", "Sheet2"]);
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
