use super::*;

impl Workbook {
    /// Add a pivot table to the workbook.
    ///
    /// The pivot table summarizes data from `config.source_sheet` /
    /// `config.source_range` and places its output on `config.target_sheet`
    /// starting at `config.target_cell`.
    pub fn add_pivot_table(&mut self, config: &PivotTableConfig) -> Result<()> {
        // Validate source sheet exists.
        let _src_idx = self.sheet_index(&config.source_sheet)?;

        // Validate target sheet exists.
        let target_idx = self.sheet_index(&config.target_sheet)?;

        // Check for duplicate name.
        if self
            .pivot_tables
            .iter()
            .any(|(_, pt)| pt.name == config.name)
        {
            return Err(Error::PivotTableAlreadyExists {
                name: config.name.clone(),
            });
        }

        // Read header row from the source data.
        let field_names = self.read_header_row(&config.source_sheet, &config.source_range)?;
        if field_names.is_empty() {
            return Err(Error::InvalidSourceRange(
                "source range header row is empty".to_string(),
            ));
        }

        // Assign a cache ID (next available).
        let cache_id = self
            .pivot_tables
            .iter()
            .map(|(_, pt)| pt.cache_id)
            .max()
            .map(|m| m + 1)
            .unwrap_or(0);

        // Build XML structures.
        let pt_def = crate::pivot::build_pivot_table_xml(config, cache_id, &field_names)?;
        let pcd = crate::pivot::build_pivot_cache_definition(
            &config.source_sheet,
            &config.source_range,
            &field_names,
        );
        let pcr = sheetkit_xml::pivot_cache::PivotCacheRecords {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            xmlns_r: sheetkit_xml::namespaces::RELATIONSHIPS.to_string(),
            count: Some(0),
            records: vec![],
        };

        // Determine part numbers.
        let pt_num = self.pivot_tables.len() + 1;
        let cache_num = self.pivot_cache_defs.len() + 1;

        let pt_path = format!("xl/pivotTables/pivotTable{}.xml", pt_num);
        let pcd_path = format!("xl/pivotCache/pivotCacheDefinition{}.xml", cache_num);
        let pcr_path = format!("xl/pivotCache/pivotCacheRecords{}.xml", cache_num);

        // Store parts.
        self.pivot_tables.push((pt_path.clone(), pt_def));
        self.pivot_cache_defs.push((pcd_path.clone(), pcd));
        self.pivot_cache_records.push((pcr_path.clone(), pcr));

        // Add content type overrides.
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/{}", pt_path),
            content_type: mime_types::PIVOT_TABLE.to_string(),
        });
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/{}", pcd_path),
            content_type: mime_types::PIVOT_CACHE_DEFINITION.to_string(),
        });
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/{}", pcr_path),
            content_type: mime_types::PIVOT_CACHE_RECORDS.to_string(),
        });

        // Add workbook relationship for pivot cache definition.
        let wb_rid = crate::sheet::next_rid(&self.workbook_rels.relationships);
        self.workbook_rels.relationships.push(Relationship {
            id: wb_rid.clone(),
            rel_type: rel_types::PIVOT_CACHE_DEF.to_string(),
            target: format!("pivotCache/pivotCacheDefinition{}.xml", cache_num),
            target_mode: None,
        });

        // Update workbook_xml.pivot_caches.
        let pivot_caches = self
            .workbook_xml
            .pivot_caches
            .get_or_insert_with(|| sheetkit_xml::workbook::PivotCaches { caches: vec![] });
        pivot_caches
            .caches
            .push(sheetkit_xml::workbook::PivotCacheEntry {
                cache_id,
                r_id: wb_rid,
            });

        // Add worksheet relationship for pivot table on the target sheet.
        let ws_rid = self.next_worksheet_rid(target_idx);
        let ws_rels = self
            .worksheet_rels
            .entry(target_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        ws_rels.relationships.push(Relationship {
            id: ws_rid,
            rel_type: rel_types::PIVOT_TABLE.to_string(),
            target: format!("../pivotTables/pivotTable{}.xml", pt_num),
            target_mode: None,
        });

        Ok(())
    }

    /// Get information about all pivot tables in the workbook.
    pub fn get_pivot_tables(&self) -> Vec<PivotTableInfo> {
        self.pivot_tables
            .iter()
            .map(|(_path, pt)| {
                // Find the matching cache definition by cache_id.
                let (source_sheet, source_range) = self
                    .pivot_cache_defs
                    .iter()
                    .enumerate()
                    .find(|(i, _)| {
                        self.workbook_xml
                            .pivot_caches
                            .as_ref()
                            .and_then(|pc| pc.caches.iter().find(|e| e.cache_id == pt.cache_id))
                            .is_some()
                            || *i == pt.cache_id as usize
                    })
                    .and_then(|(_, (_, pcd))| {
                        pcd.cache_source
                            .worksheet_source
                            .as_ref()
                            .map(|ws| (ws.sheet.clone(), ws.reference.clone()))
                    })
                    .unwrap_or_default();

                // Determine target sheet from the pivot table path.
                let target_sheet = self.find_pivot_table_target_sheet(pt).unwrap_or_default();

                PivotTableInfo {
                    name: pt.name.clone(),
                    source_sheet,
                    source_range,
                    target_sheet,
                    location: pt.location.reference.clone(),
                }
            })
            .collect()
    }

    /// Delete a pivot table by name.
    pub fn delete_pivot_table(&mut self, name: &str) -> Result<()> {
        // Find the pivot table.
        let pt_idx = self
            .pivot_tables
            .iter()
            .position(|(_, pt)| pt.name == name)
            .ok_or_else(|| Error::PivotTableNotFound {
                name: name.to_string(),
            })?;

        let (pt_path, pt_def) = self.pivot_tables.remove(pt_idx);
        let cache_id = pt_def.cache_id;

        // Remove the matching pivot cache definition and records.
        // Find the workbook_xml pivot cache entry for this cache_id.
        let mut wb_cache_rid = None;
        if let Some(ref mut pivot_caches) = self.workbook_xml.pivot_caches {
            if let Some(pos) = pivot_caches
                .caches
                .iter()
                .position(|e| e.cache_id == cache_id)
            {
                wb_cache_rid = Some(pivot_caches.caches[pos].r_id.clone());
                pivot_caches.caches.remove(pos);
            }
            if pivot_caches.caches.is_empty() {
                self.workbook_xml.pivot_caches = None;
            }
        }

        // Remove the workbook relationship for this cache.
        if let Some(ref rid) = wb_cache_rid {
            // Find the target to determine which cache def to remove.
            if let Some(rel) = self
                .workbook_rels
                .relationships
                .iter()
                .find(|r| r.id == *rid)
            {
                let target_path = format!("xl/{}", rel.target);
                self.pivot_cache_defs.retain(|(p, _)| *p != target_path);

                // Remove matching cache records (same numbering).
                let records_path = target_path.replace("pivotCacheDefinition", "pivotCacheRecords");
                self.pivot_cache_records.retain(|(p, _)| *p != records_path);
            }
            self.workbook_rels.relationships.retain(|r| r.id != *rid);
        }

        // Remove content type overrides for the removed parts.
        let pt_part = format!("/{}", pt_path);
        self.content_types
            .overrides
            .retain(|o| o.part_name != pt_part);

        // Also remove cache def and records content types if the paths were removed.
        self.content_types.overrides.retain(|o| {
            let p = o.part_name.trim_start_matches('/');
            // Keep if it is still in our live lists.
            if o.content_type == mime_types::PIVOT_CACHE_DEFINITION {
                return self.pivot_cache_defs.iter().any(|(path, _)| path == p);
            }
            if o.content_type == mime_types::PIVOT_CACHE_RECORDS {
                return self.pivot_cache_records.iter().any(|(path, _)| path == p);
            }
            if o.content_type == mime_types::PIVOT_TABLE {
                return self.pivot_tables.iter().any(|(path, _)| path == p);
            }
            true
        });

        // Remove worksheet relationship for this pivot table.
        for (_idx, rels) in self.worksheet_rels.iter_mut() {
            rels.relationships.retain(|r| {
                if r.rel_type != rel_types::PIVOT_TABLE {
                    return true;
                }
                // Check if the target matches the removed pivot table.
                let full_target = format!(
                    "xl/pivotTables/{}",
                    r.target.trim_start_matches("../pivotTables/")
                );
                full_target != pt_path
            });
        }

        Ok(())
    }

    /// Add a sparkline to a worksheet.
    pub fn add_sparkline(
        &mut self,
        sheet: &str,
        config: &crate::sparkline::SparklineConfig,
    ) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        crate::sparkline::validate_sparkline_config(config)?;
        while self.sheet_sparklines.len() <= idx {
            self.sheet_sparklines.push(vec![]);
        }
        self.sheet_sparklines[idx].push(config.clone());
        Ok(())
    }

    /// Get all sparklines for a worksheet.
    pub fn get_sparklines(&self, sheet: &str) -> Result<Vec<crate::sparkline::SparklineConfig>> {
        let idx = self.sheet_index(sheet)?;
        Ok(self.sheet_sparklines.get(idx).cloned().unwrap_or_default())
    }

    /// Remove a sparkline by its location cell reference.
    pub fn remove_sparkline(&mut self, sheet: &str, location: &str) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        let sparklines = self
            .sheet_sparklines
            .get_mut(idx)
            .ok_or_else(|| Error::Internal(format!("no sparkline data for sheet '{sheet}'")))?;
        let pos = sparklines
            .iter()
            .position(|s| s.location == location)
            .ok_or_else(|| {
                Error::Internal(format!(
                    "sparkline at location '{location}' not found on sheet '{sheet}'"
                ))
            })?;
        sparklines.remove(pos);
        Ok(())
    }

    /// Evaluate a single formula string in the context of `sheet`.
    ///
    /// A [`CellSnapshot`] is built from the current workbook state so
    /// that cell references within the formula can be resolved.
    pub fn evaluate_formula(&self, sheet: &str, formula: &str) -> Result<CellValue> {
        // Validate the sheet exists.
        let _ = self.sheet_index(sheet)?;
        let parsed = crate::formula::parser::parse_formula(formula)?;
        let snapshot = self.build_cell_snapshot(sheet)?;
        crate::formula::eval::evaluate(&parsed, &snapshot)
    }

    /// Recalculate every formula cell across all sheets and store the
    /// computed result back into each cell. Uses a dependency graph and
    /// topological sort so formulas are evaluated after their dependencies.
    pub fn calculate_all(&mut self) -> Result<()> {
        use crate::formula::eval::{build_dependency_graph, topological_sort, CellCoord};

        let sheet_names: Vec<String> = self.sheet_names().iter().map(|s| s.to_string()).collect();

        // Collect all formula cells with their coordinates and formula strings.
        let mut formula_cells: Vec<(CellCoord, String)> = Vec::new();
        for (idx, sn) in sheet_names.iter().enumerate() {
            self.ensure_hydrated(idx)?;
            let ws = self.worksheets[idx].1.get().unwrap();
            for row in &ws.sheet_data.rows {
                for cell in &row.cells {
                    if let Some(ref f) = cell.f {
                        let formula_str = f.value.clone().unwrap_or_default();
                        if !formula_str.is_empty() {
                            if let Ok((c, r)) = cell_name_to_coordinates(cell.r.as_str()) {
                                formula_cells.push((
                                    CellCoord {
                                        sheet: sn.clone(),
                                        col: c,
                                        row: r,
                                    },
                                    formula_str,
                                ));
                            }
                        }
                    }
                }
            }
        }

        if formula_cells.is_empty() {
            return Ok(());
        }

        // Build dependency graph and determine evaluation order.
        let deps = build_dependency_graph(&formula_cells)?;
        let coords: Vec<CellCoord> = formula_cells.iter().map(|(c, _)| c.clone()).collect();
        let eval_order = topological_sort(&coords, &deps)?;

        // Build a lookup from coord to formula string.
        let formula_map: HashMap<CellCoord, String> = formula_cells.into_iter().collect();

        // Build a snapshot of all cell data.
        let first_sheet = sheet_names.first().cloned().unwrap_or_default();
        let mut snapshot = self.build_cell_snapshot(&first_sheet)?;

        // Evaluate in dependency order, updating the snapshot progressively
        // so later formulas see already-computed results.
        let mut results: Vec<(CellCoord, String, CellValue)> = Vec::new();
        for coord in &eval_order {
            if let Some(formula_str) = formula_map.get(coord) {
                snapshot.set_current_sheet(&coord.sheet);
                let parsed = crate::formula::parser::parse_formula(formula_str)?;
                let mut evaluator = crate::formula::eval::Evaluator::new(&snapshot);
                let result = evaluator.eval_expr(&parsed)?;
                snapshot.set_cell(&coord.sheet, coord.col, coord.row, result.clone());
                results.push((coord.clone(), formula_str.clone(), result));
            }
        }

        // Write results back directly to the XML cells, preserving the
        // formula element and storing the computed value in the v/t fields.
        for (coord, _formula_str, result) in results {
            let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(coord.col, coord.row)?;
            if let Some((_, ws_lock)) = self.worksheets.iter_mut().find(|(n, _)| *n == coord.sheet)
            {
                let Some(ws) = ws_lock.get_mut() else {
                    continue;
                };
                if let Some(row) = ws.sheet_data.rows.iter_mut().find(|r| r.r == coord.row) {
                    if let Some(cell) = row.cells.iter_mut().find(|c| c.r == *cell_ref) {
                        match &result {
                            CellValue::Number(n) => {
                                cell.v = Some(n.to_string());
                                cell.t = CellTypeTag::None;
                            }
                            CellValue::String(s) => {
                                cell.v = Some(s.clone());
                                cell.t = CellTypeTag::FormulaString;
                            }
                            CellValue::Bool(b) => {
                                cell.v = Some(if *b { "1".to_string() } else { "0".to_string() });
                                cell.t = CellTypeTag::Boolean;
                            }
                            CellValue::Error(e) => {
                                cell.v = Some(e.clone());
                                cell.t = CellTypeTag::Error;
                            }
                            CellValue::Date(n) => {
                                cell.v = Some(n.to_string());
                                cell.t = CellTypeTag::None;
                            }
                            _ => {}
                        }
                    }
                }
            }
        }

        Ok(())
    }

    /// Build a [`CellSnapshot`] for formula evaluation, with the given
    /// sheet as the current-sheet context.
    fn build_cell_snapshot(
        &self,
        current_sheet: &str,
    ) -> Result<crate::formula::eval::CellSnapshot> {
        let mut snapshot = crate::formula::eval::CellSnapshot::new(current_sheet.to_string());
        for (idx, (sn, _)) in self.worksheets.iter().enumerate() {
            let ws = self.worksheet_ref_by_index(idx)?;
            for row in &ws.sheet_data.rows {
                for cell in &row.cells {
                    if let Ok((c, r)) = cell_name_to_coordinates(cell.r.as_str()) {
                        let cv = self.xml_cell_to_value(cell)?;
                        snapshot.set_cell(sn, c, r, cv);
                    }
                }
            }
        }
        Ok(snapshot)
    }

    /// Return `(col, row)` pairs for all occupied cells on the named sheet.
    pub fn get_occupied_cells(&self, sheet: &str) -> Result<Vec<(u32, u32)>> {
        let ws = self.worksheet_ref(sheet)?;
        let mut cells = Vec::new();
        for row in &ws.sheet_data.rows {
            for cell in &row.cells {
                if let Ok((c, r)) = cell_name_to_coordinates(cell.r.as_str()) {
                    cells.push((c, r));
                }
            }
        }
        Ok(cells)
    }

    /// Read the header row (first row) of a range from a sheet, returning cell
    /// values as strings.
    fn read_header_row(&self, sheet: &str, range: &str) -> Result<Vec<String>> {
        let parts: Vec<&str> = range.split(':').collect();
        if parts.len() != 2 {
            return Err(Error::InvalidSourceRange(range.to_string()));
        }
        let (start_col, start_row) = cell_name_to_coordinates(parts[0])
            .map_err(|_| Error::InvalidSourceRange(range.to_string()))?;
        let (end_col, _end_row) = cell_name_to_coordinates(parts[1])
            .map_err(|_| Error::InvalidSourceRange(range.to_string()))?;

        let mut headers = Vec::new();
        for col in start_col..=end_col {
            let cell_name = crate::utils::cell_ref::coordinates_to_cell_name(col, start_row)?;
            let val = self.get_cell_value(sheet, &cell_name)?;
            let s = match val {
                CellValue::String(s) => s,
                CellValue::Number(n) => n.to_string(),
                CellValue::Bool(b) => b.to_string(),
                CellValue::RichString(runs) => crate::rich_text::rich_text_to_plain(&runs),
                _ => String::new(),
            };
            headers.push(s);
        }
        Ok(headers)
    }

    /// Find the target sheet name for a pivot table by looking at worksheet
    /// relationships that reference its path.
    fn find_pivot_table_target_sheet(
        &self,
        pt: &sheetkit_xml::pivot_table::PivotTableDefinition,
    ) -> Option<String> {
        // Find the pivot table path.
        let pt_path = self
            .pivot_tables
            .iter()
            .find(|(_, p)| p.name == pt.name)
            .map(|(path, _)| path.as_str())?;

        // Find which worksheet has a relationship pointing to this pivot table.
        for (sheet_idx, rels) in &self.worksheet_rels {
            for r in &rels.relationships {
                if r.rel_type == rel_types::PIVOT_TABLE {
                    let full_target = format!(
                        "xl/pivotTables/{}",
                        r.target.trim_start_matches("../pivotTables/")
                    );
                    if full_target == pt_path {
                        return self
                            .worksheets
                            .get(*sheet_idx)
                            .map(|(name, _)| name.clone());
                    }
                }
            }
        }
        None
    }

    /// Set the core document properties (title, author, etc.).
    pub fn set_doc_props(&mut self, props: crate::doc_props::DocProperties) {
        self.core_properties = Some(props.to_core_properties());
        self.ensure_doc_props_content_types();
    }

    /// Get the core document properties.
    pub fn get_doc_props(&self) -> crate::doc_props::DocProperties {
        self.core_properties
            .as_ref()
            .map(crate::doc_props::DocProperties::from)
            .unwrap_or_default()
    }

    /// Set the application properties (company, app version, etc.).
    pub fn set_app_props(&mut self, props: crate::doc_props::AppProperties) {
        self.app_properties = Some(props.to_extended_properties());
        self.ensure_doc_props_content_types();
    }

    /// Get the application properties.
    pub fn get_app_props(&self) -> crate::doc_props::AppProperties {
        self.app_properties
            .as_ref()
            .map(crate::doc_props::AppProperties::from)
            .unwrap_or_default()
    }

    /// Set a custom property by name. If a property with the same name already
    /// exists, its value is replaced.
    pub fn set_custom_property(
        &mut self,
        name: &str,
        value: crate::doc_props::CustomPropertyValue,
    ) {
        let props = self
            .custom_properties
            .get_or_insert_with(sheetkit_xml::doc_props::CustomProperties::default);
        crate::doc_props::set_custom_property(props, name, value);
        self.ensure_custom_props_content_types();
    }

    /// Get a custom property value by name, or `None` if it does not exist.
    pub fn get_custom_property(&self, name: &str) -> Option<crate::doc_props::CustomPropertyValue> {
        self.custom_properties
            .as_ref()
            .and_then(|p| crate::doc_props::find_custom_property(p, name))
    }

    /// Remove a custom property by name. Returns `true` if a property was
    /// found and removed.
    pub fn delete_custom_property(&mut self, name: &str) -> bool {
        if let Some(ref mut props) = self.custom_properties {
            crate::doc_props::delete_custom_property(props, name)
        } else {
            false
        }
    }

    /// Ensure content types contains entries for core and extended properties.
    fn ensure_doc_props_content_types(&mut self) {
        let core_part = "/docProps/core.xml";
        let app_part = "/docProps/app.xml";

        let has_core = self
            .content_types
            .overrides
            .iter()
            .any(|o| o.part_name == core_part);
        if !has_core {
            self.content_types.overrides.push(ContentTypeOverride {
                part_name: core_part.to_string(),
                content_type: mime_types::CORE_PROPERTIES.to_string(),
            });
        }

        let has_app = self
            .content_types
            .overrides
            .iter()
            .any(|o| o.part_name == app_part);
        if !has_app {
            self.content_types.overrides.push(ContentTypeOverride {
                part_name: app_part.to_string(),
                content_type: mime_types::EXTENDED_PROPERTIES.to_string(),
            });
        }
    }

    /// Ensure content types and package rels contain entries for custom properties.
    fn ensure_custom_props_content_types(&mut self) {
        self.ensure_doc_props_content_types();

        let custom_part = "/docProps/custom.xml";
        let has_custom = self
            .content_types
            .overrides
            .iter()
            .any(|o| o.part_name == custom_part);
        if !has_custom {
            self.content_types.overrides.push(ContentTypeOverride {
                part_name: custom_part.to_string(),
                content_type: mime_types::CUSTOM_PROPERTIES.to_string(),
            });
        }

        let has_custom_rel = self
            .package_rels
            .relationships
            .iter()
            .any(|r| r.rel_type == rel_types::CUSTOM_PROPERTIES);
        if !has_custom_rel {
            let next_id = self.package_rels.relationships.len() + 1;
            self.package_rels.relationships.push(Relationship {
                id: format!("rId{next_id}"),
                rel_type: rel_types::CUSTOM_PROPERTIES.to_string(),
                target: "docProps/custom.xml".to_string(),
                target_mode: None,
            });
        }
    }

    /// Look up a table by name across all sheets. Returns a reference to the
    /// table XML, path, and sheet index from the main table storage.
    fn find_table_by_name(
        &self,
        name: &str,
    ) -> Option<(&String, &sheetkit_xml::table::TableXml, usize)> {
        self.tables
            .iter()
            .find(|(_, t, _)| t.name == name)
            .map(|(path, t, idx)| (path, t, *idx))
    }

    /// Look up the table name for a given table ID.
    fn table_name_by_id(&self, table_id: u32) -> Option<&str> {
        self.tables
            .iter()
            .find(|(_, t, _)| t.id == table_id)
            .map(|(_, t, _)| t.name.as_str())
    }

    /// Look up the column name by 1-based column index and table ID.
    fn table_column_name_by_index(&self, table_id: u32, column_index: u32) -> Option<&str> {
        self.tables
            .iter()
            .find(|(_, t, _)| t.id == table_id)
            .and_then(|(_, t, _)| {
                t.table_columns
                    .columns
                    .get((column_index - 1) as usize)
                    .map(|c| c.name.as_str())
            })
    }

    /// Add a slicer to a sheet targeting a table column.
    ///
    /// Creates the slicer definition, slicer cache, content type overrides,
    /// and worksheet relationships needed for Excel to render the slicer.
    pub fn add_slicer(&mut self, sheet: &str, config: &crate::slicer::SlicerConfig) -> Result<()> {
        use sheetkit_xml::content_types::ContentTypeOverride;
        use sheetkit_xml::slicer::{
            SlicerCacheDefinition, SlicerDefinition, SlicerDefinitions, TableSlicerCache,
        };

        crate::slicer::validate_slicer_config(config)?;

        let sheet_idx = self.sheet_index(sheet)?;

        // Check for duplicate name across all slicer definitions.
        for (_, sd) in &self.slicer_defs {
            for s in &sd.slicers {
                if s.name == config.name {
                    return Err(Error::SlicerAlreadyExists {
                        name: config.name.clone(),
                    });
                }
            }
        }

        let cache_name = crate::slicer::slicer_cache_name(&config.name);
        let caption = config
            .caption
            .clone()
            .unwrap_or_else(|| config.column_name.clone());

        // Determine part numbers.
        let slicer_num = self.slicer_defs.len() + 1;
        let cache_num = self.slicer_caches.len() + 1;

        let slicer_path = format!("xl/slicers/slicer{}.xml", slicer_num);
        let cache_path = format!("xl/slicerCaches/slicerCache{}.xml", cache_num);

        // Build the slicer definition.
        let slicer_def = SlicerDefinition {
            name: config.name.clone(),
            cache: cache_name.clone(),
            caption: Some(caption),
            start_item: None,
            column_count: config.column_count,
            show_caption: config.show_caption,
            style: config.style.clone(),
            locked_position: None,
            row_height: crate::slicer::DEFAULT_ROW_HEIGHT_EMU,
        };

        let slicer_defs = SlicerDefinitions {
            xmlns: sheetkit_xml::namespaces::SLICER_2009.to_string(),
            xmlns_mc: Some(sheetkit_xml::namespaces::MC.to_string()),
            slicers: vec![slicer_def],
        };

        // Look up the actual table by name and validate it exists on this sheet.
        let (_path, table_xml, table_sheet_idx) = self
            .find_table_by_name(&config.table_name)
            .ok_or_else(|| Error::TableNotFound {
                name: config.table_name.clone(),
            })?;

        if table_sheet_idx != sheet_idx {
            return Err(Error::TableNotFound {
                name: config.table_name.clone(),
            });
        }

        // Validate the column exists in the table and get its 1-based index.
        let column_index = table_xml
            .table_columns
            .columns
            .iter()
            .position(|c| c.name == config.column_name)
            .ok_or_else(|| Error::TableColumnNotFound {
                table: config.table_name.clone(),
                column: config.column_name.clone(),
            })?;

        let real_table_id = table_xml.id;
        let real_column = (column_index + 1) as u32;

        // Build the slicer cache definition with real table metadata.
        let slicer_cache = SlicerCacheDefinition {
            name: cache_name.clone(),
            source_name: config.column_name.clone(),
            table_slicer_cache: Some(TableSlicerCache {
                table_id: real_table_id,
                column: real_column,
            }),
        };

        // Store parts.
        self.slicer_defs.push((slicer_path.clone(), slicer_defs));
        self.slicer_caches.push((cache_path.clone(), slicer_cache));

        // Add content type overrides.
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/{}", slicer_path),
            content_type: mime_types::SLICER.to_string(),
        });
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/{}", cache_path),
            content_type: mime_types::SLICER_CACHE.to_string(),
        });

        // Add workbook relationship for slicer cache.
        let wb_rid = crate::sheet::next_rid(&self.workbook_rels.relationships);
        self.workbook_rels.relationships.push(Relationship {
            id: wb_rid,
            rel_type: rel_types::SLICER_CACHE.to_string(),
            target: format!("slicerCaches/slicerCache{}.xml", cache_num),
            target_mode: None,
        });

        // Add worksheet relationship for slicer part.
        let ws_rid = self.next_worksheet_rid(sheet_idx);
        let ws_rels = self
            .worksheet_rels
            .entry(sheet_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        ws_rels.relationships.push(Relationship {
            id: ws_rid,
            rel_type: rel_types::SLICER.to_string(),
            target: format!("../slicers/slicer{}.xml", slicer_num),
            target_mode: None,
        });

        Ok(())
    }

    /// Get information about all slicers on a sheet.
    pub fn get_slicers(&self, sheet: &str) -> Result<Vec<crate::slicer::SlicerInfo>> {
        let sheet_idx = self.sheet_index(sheet)?;
        let mut result = Vec::new();

        // Find slicer parts referenced by this sheet's relationships.
        let empty_rels = Relationships {
            xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
            relationships: vec![],
        };
        let rels = self.worksheet_rels.get(&sheet_idx).unwrap_or(&empty_rels);

        let slicer_targets: Vec<String> = rels
            .relationships
            .iter()
            .filter(|r| r.rel_type == rel_types::SLICER)
            .map(|r| {
                let sheet_path = self.sheet_part_path(sheet_idx);
                crate::workbook_paths::resolve_relationship_target(&sheet_path, &r.target)
            })
            .collect();

        for (path, sd) in &self.slicer_defs {
            if !slicer_targets.contains(path) {
                continue;
            }
            for slicer in &sd.slicers {
                // Find the matching cache to get source info.
                let cache = self
                    .slicer_caches
                    .iter()
                    .find(|(_, sc)| sc.name == slicer.cache);

                let (table_name, column_name) = if let Some((_, sc)) = cache {
                    let tname = sc
                        .table_slicer_cache
                        .as_ref()
                        .and_then(|tsc| self.table_name_by_id(tsc.table_id))
                        .unwrap_or("")
                        .to_string();
                    let cname = sc
                        .table_slicer_cache
                        .as_ref()
                        .and_then(|tsc| self.table_column_name_by_index(tsc.table_id, tsc.column))
                        .unwrap_or(&sc.source_name)
                        .to_string();
                    (tname, cname)
                } else {
                    (String::new(), String::new())
                };

                result.push(crate::slicer::SlicerInfo {
                    name: slicer.name.clone(),
                    caption: slicer
                        .caption
                        .clone()
                        .unwrap_or_else(|| slicer.name.clone()),
                    table_name,
                    column_name,
                    style: slicer.style.clone(),
                });
            }
        }

        Ok(result)
    }

    /// Delete a slicer by name from a sheet.
    ///
    /// Removes the slicer definition, cache, content types, and relationships.
    pub fn delete_slicer(&mut self, sheet: &str, name: &str) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;

        // Find the slicer definition containing this slicer name.
        let sd_idx = self
            .slicer_defs
            .iter()
            .position(|(_, sd)| sd.slicers.iter().any(|s| s.name == name))
            .ok_or_else(|| Error::SlicerNotFound {
                name: name.to_string(),
            })?;

        let (sd_path, sd) = &self.slicer_defs[sd_idx];

        // Find the cache name linked to this slicer.
        let cache_name = sd
            .slicers
            .iter()
            .find(|s| s.name == name)
            .map(|s| s.cache.clone())
            .unwrap_or_default();

        // If this is the only slicer in this definitions part, remove the whole part.
        let remove_whole_part = sd.slicers.len() == 1;

        if remove_whole_part {
            let sd_path_clone = sd_path.clone();
            self.slicer_defs.remove(sd_idx);

            // Remove content type override.
            let sd_part = format!("/{}", sd_path_clone);
            self.content_types
                .overrides
                .retain(|o| o.part_name != sd_part);

            // Remove worksheet relationship pointing to this slicer part.
            let ws_path = self.sheet_part_path(sheet_idx);
            if let Some(rels) = self.worksheet_rels.get_mut(&sheet_idx) {
                rels.relationships.retain(|r| {
                    if r.rel_type != rel_types::SLICER {
                        return true;
                    }
                    let target =
                        crate::workbook_paths::resolve_relationship_target(&ws_path, &r.target);
                    target != sd_path_clone
                });
            }
        } else {
            // Remove just this slicer from the definitions.
            self.slicer_defs[sd_idx]
                .1
                .slicers
                .retain(|s| s.name != name);
        }

        // Remove the matching slicer cache.
        if !cache_name.is_empty() {
            if let Some(sc_idx) = self
                .slicer_caches
                .iter()
                .position(|(_, sc)| sc.name == cache_name)
            {
                let (sc_path, _) = self.slicer_caches.remove(sc_idx);
                let sc_part = format!("/{}", sc_path);
                self.content_types
                    .overrides
                    .retain(|o| o.part_name != sc_part);

                // Remove workbook relationship for this cache.
                self.workbook_rels.relationships.retain(|r| {
                    if r.rel_type != rel_types::SLICER_CACHE {
                        return true;
                    }
                    let full_target = format!("xl/{}", r.target);
                    full_target != sc_path
                });
            }
        }

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    fn make_pivot_workbook() -> Workbook {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Name").unwrap();
        wb.set_cell_value("Sheet1", "B1", "Region").unwrap();
        wb.set_cell_value("Sheet1", "C1", "Sales").unwrap();
        wb.set_cell_value("Sheet1", "A2", "Alice").unwrap();
        wb.set_cell_value("Sheet1", "B2", "North").unwrap();
        wb.set_cell_value("Sheet1", "C2", 100.0).unwrap();
        wb.set_cell_value("Sheet1", "A3", "Bob").unwrap();
        wb.set_cell_value("Sheet1", "B3", "South").unwrap();
        wb.set_cell_value("Sheet1", "C3", 200.0).unwrap();
        wb.set_cell_value("Sheet1", "A4", "Carol").unwrap();
        wb.set_cell_value("Sheet1", "B4", "North").unwrap();
        wb.set_cell_value("Sheet1", "C4", 150.0).unwrap();
        wb
    }

    fn basic_pivot_config() -> PivotTableConfig {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        PivotTableConfig {
            name: "PivotTable1".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "E1".to_string(),
            rows: vec![PivotField {
                name: "Name".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        }
    }

    #[test]
    fn test_add_pivot_table_basic() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        assert_eq!(wb.pivot_tables.len(), 1);
        assert_eq!(wb.pivot_cache_defs.len(), 1);
        assert_eq!(wb.pivot_cache_records.len(), 1);
        assert_eq!(wb.pivot_tables[0].1.name, "PivotTable1");
        assert_eq!(wb.pivot_tables[0].1.cache_id, 0);
    }

    #[test]
    fn test_add_pivot_table_with_columns() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();
        let config = PivotTableConfig {
            name: "PT2".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "E1".to_string(),
            rows: vec![PivotField {
                name: "Name".to_string(),
            }],
            columns: vec![PivotField {
                name: "Region".to_string(),
            }],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Average,
                display_name: Some("Avg Sales".to_string()),
            }],
        };
        wb.add_pivot_table(&config).unwrap();

        let pt = &wb.pivot_tables[0].1;
        assert!(pt.row_fields.is_some());
        assert!(pt.col_fields.is_some());
        assert!(pt.data_fields.is_some());
    }

    #[test]
    fn test_add_pivot_table_source_sheet_not_found() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = Workbook::new();
        let config = PivotTableConfig {
            name: "PT".to_string(),
            source_sheet: "NonExistent".to_string(),
            source_range: "A1:B2".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![PivotField {
                name: "Col1".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Col2".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        let result = wb.add_pivot_table(&config);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_add_pivot_table_target_sheet_not_found() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();
        let config = PivotTableConfig {
            name: "PT".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Report".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![PivotField {
                name: "Name".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        let result = wb.add_pivot_table(&config);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_add_pivot_table_duplicate_name() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        let result = wb.add_pivot_table(&config);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::PivotTableAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_get_pivot_tables_empty() {
        let wb = Workbook::new();
        let pts = wb.get_pivot_tables();
        assert!(pts.is_empty());
    }

    #[test]
    fn test_get_pivot_tables_after_add() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        let pts = wb.get_pivot_tables();
        assert_eq!(pts.len(), 1);
        assert_eq!(pts[0].name, "PivotTable1");
        assert_eq!(pts[0].source_sheet, "Sheet1");
        assert_eq!(pts[0].source_range, "A1:C4");
        assert_eq!(pts[0].target_sheet, "Sheet1");
        assert_eq!(pts[0].location, "E1");
    }

    #[test]
    fn test_delete_pivot_table() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();
        assert_eq!(wb.pivot_tables.len(), 1);

        wb.delete_pivot_table("PivotTable1").unwrap();
        assert!(wb.pivot_tables.is_empty());
        assert!(wb.pivot_cache_defs.is_empty());
        assert!(wb.pivot_cache_records.is_empty());
        assert!(wb.workbook_xml.pivot_caches.is_none());

        // Content type overrides for pivot parts should be gone.
        let pivot_overrides: Vec<_> = wb
            .content_types
            .overrides
            .iter()
            .filter(|o| {
                o.content_type == mime_types::PIVOT_TABLE
                    || o.content_type == mime_types::PIVOT_CACHE_DEFINITION
                    || o.content_type == mime_types::PIVOT_CACHE_RECORDS
            })
            .collect();
        assert!(pivot_overrides.is_empty());
    }

    #[test]
    fn test_delete_pivot_table_not_found() {
        let wb_result = Workbook::new().delete_pivot_table("NonExistent");
        assert!(wb_result.is_err());
        assert!(matches!(
            wb_result.unwrap_err(),
            Error::PivotTableNotFound { .. }
        ));
    }

    #[test]
    fn test_pivot_table_save_open_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("pivot_roundtrip.xlsx");

        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        wb.save(&path).unwrap();

        // Verify the ZIP contains pivot parts.
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/pivotTables/pivotTable1.xml").is_ok());
        assert!(archive
            .by_name("xl/pivotCache/pivotCacheDefinition1.xml")
            .is_ok());
        assert!(archive
            .by_name("xl/pivotCache/pivotCacheRecords1.xml")
            .is_ok());

        // Re-open and verify pivot table is parsed.
        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.pivot_tables.len(), 1);
        assert_eq!(wb2.pivot_tables[0].1.name, "PivotTable1");
        assert_eq!(wb2.pivot_cache_defs.len(), 1);
        assert_eq!(wb2.pivot_cache_records.len(), 1);
    }

    #[test]
    fn test_add_multiple_pivot_tables() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();

        let config1 = basic_pivot_config();
        wb.add_pivot_table(&config1).unwrap();

        let config2 = PivotTableConfig {
            name: "PivotTable2".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "H1".to_string(),
            rows: vec![PivotField {
                name: "Region".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Count,
                display_name: None,
            }],
        };
        wb.add_pivot_table(&config2).unwrap();

        assert_eq!(wb.pivot_tables.len(), 2);
        assert_eq!(wb.pivot_cache_defs.len(), 2);
        assert_eq!(wb.pivot_tables[0].1.cache_id, 0);
        assert_eq!(wb.pivot_tables[1].1.cache_id, 1);

        let pts = wb.get_pivot_tables();
        assert_eq!(pts.len(), 2);
        assert_eq!(pts[0].name, "PivotTable1");
        assert_eq!(pts[1].name, "PivotTable2");
    }

    #[test]
    fn test_add_pivot_table_content_types_added() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        let has_pt_ct = wb.content_types.overrides.iter().any(|o| {
            o.content_type == mime_types::PIVOT_TABLE
                && o.part_name == "/xl/pivotTables/pivotTable1.xml"
        });
        assert!(has_pt_ct);

        let has_pcd_ct = wb.content_types.overrides.iter().any(|o| {
            o.content_type == mime_types::PIVOT_CACHE_DEFINITION
                && o.part_name == "/xl/pivotCache/pivotCacheDefinition1.xml"
        });
        assert!(has_pcd_ct);

        let has_pcr_ct = wb.content_types.overrides.iter().any(|o| {
            o.content_type == mime_types::PIVOT_CACHE_RECORDS
                && o.part_name == "/xl/pivotCache/pivotCacheRecords1.xml"
        });
        assert!(has_pcr_ct);
    }

    #[test]
    fn test_add_pivot_table_workbook_rels_and_pivot_caches() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        // Workbook rels should have a pivot cache definition relationship.
        let cache_rel = wb
            .workbook_rels
            .relationships
            .iter()
            .find(|r| r.rel_type == rel_types::PIVOT_CACHE_DEF);
        assert!(cache_rel.is_some());
        let cache_rel = cache_rel.unwrap();
        assert_eq!(cache_rel.target, "pivotCache/pivotCacheDefinition1.xml");

        // Workbook XML should have pivot caches.
        let pivot_caches = wb.workbook_xml.pivot_caches.as_ref().unwrap();
        assert_eq!(pivot_caches.caches.len(), 1);
        assert_eq!(pivot_caches.caches[0].cache_id, 0);
        assert_eq!(pivot_caches.caches[0].r_id, cache_rel.id);
    }

    #[test]
    fn test_add_pivot_table_worksheet_rels_added() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        // Sheet1 is index 0; its rels should have a pivot table relationship.
        let ws_rels = wb.worksheet_rels.get(&0).unwrap();
        let pt_rel = ws_rels
            .relationships
            .iter()
            .find(|r| r.rel_type == rel_types::PIVOT_TABLE);
        assert!(pt_rel.is_some());
        assert_eq!(pt_rel.unwrap().target, "../pivotTables/pivotTable1.xml");
    }

    #[test]
    fn test_add_pivot_table_on_separate_target_sheet() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();
        wb.new_sheet("Report").unwrap();

        let config = PivotTableConfig {
            name: "CrossSheet".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Report".to_string(),
            target_cell: "A1".to_string(),
            rows: vec![PivotField {
                name: "Name".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        wb.add_pivot_table(&config).unwrap();

        let pts = wb.get_pivot_tables();
        assert_eq!(pts.len(), 1);
        assert_eq!(pts[0].target_sheet, "Report");
        assert_eq!(pts[0].source_sheet, "Sheet1");

        // Worksheet rels should be on the Report sheet (index 1).
        let ws_rels = wb.worksheet_rels.get(&1).unwrap();
        let pt_rel = ws_rels
            .relationships
            .iter()
            .find(|r| r.rel_type == rel_types::PIVOT_TABLE);
        assert!(pt_rel.is_some());
    }

    #[test]
    fn test_pivot_table_invalid_source_range() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();
        let config = PivotTableConfig {
            name: "BadRange".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "INVALID".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "E1".to_string(),
            rows: vec![PivotField {
                name: "Name".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        let result = wb.add_pivot_table(&config);
        assert!(result.is_err());
    }

    #[test]
    fn test_delete_pivot_table_then_add_another() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = make_pivot_workbook();
        let config1 = basic_pivot_config();
        wb.add_pivot_table(&config1).unwrap();
        wb.delete_pivot_table("PivotTable1").unwrap();

        let config2 = PivotTableConfig {
            name: "PivotTable2".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "E1".to_string(),
            rows: vec![PivotField {
                name: "Region".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Max,
                display_name: None,
            }],
        };
        wb.add_pivot_table(&config2).unwrap();

        assert_eq!(wb.pivot_tables.len(), 1);
        assert_eq!(wb.pivot_tables[0].1.name, "PivotTable2");
    }

    #[test]
    fn test_pivot_table_cache_definition_stores_source_info() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        let pcd = &wb.pivot_cache_defs[0].1;
        let ws_source = pcd.cache_source.worksheet_source.as_ref().unwrap();
        assert_eq!(ws_source.sheet, "Sheet1");
        assert_eq!(ws_source.reference, "A1:C4");
        assert_eq!(pcd.cache_fields.fields.len(), 3);
        assert_eq!(pcd.cache_fields.fields[0].name, "Name");
        assert_eq!(pcd.cache_fields.fields[1].name, "Region");
        assert_eq!(pcd.cache_fields.fields[2].name, "Sales");
    }

    #[test]
    fn test_pivot_table_field_names_from_data() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();

        let pt = &wb.pivot_tables[0].1;
        assert_eq!(pt.pivot_fields.fields.len(), 3);
        // Name is a row field.
        assert_eq!(pt.pivot_fields.fields[0].axis, Some("axisRow".to_string()));
        // Region is not used.
        assert_eq!(pt.pivot_fields.fields[1].axis, None);
        // Sales is a data field.
        assert_eq!(pt.pivot_fields.fields[2].data_field, Some(true));
    }

    #[test]
    fn test_pivot_table_empty_header_row_error() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let mut wb = Workbook::new();
        // No data set in the sheet.
        let config = PivotTableConfig {
            name: "Empty".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:B1".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "D1".to_string(),
            rows: vec![PivotField {
                name: "X".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Y".to_string(),
                function: AggregateFunction::Sum,
                display_name: None,
            }],
        };
        let result = wb.add_pivot_table(&config);
        assert!(result.is_err());
    }

    #[test]
    fn test_pivot_table_multiple_save_roundtrip() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("multi_pivot.xlsx");

        let mut wb = make_pivot_workbook();
        let config1 = basic_pivot_config();
        wb.add_pivot_table(&config1).unwrap();

        let config2 = PivotTableConfig {
            name: "PT2".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:C4".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "H1".to_string(),
            rows: vec![PivotField {
                name: "Region".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Sales".to_string(),
                function: AggregateFunction::Min,
                display_name: None,
            }],
        };
        wb.add_pivot_table(&config2).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.pivot_tables.len(), 2);
        let names: Vec<&str> = wb2
            .pivot_tables
            .iter()
            .map(|(_, pt)| pt.name.as_str())
            .collect();
        assert!(names.contains(&"PivotTable1"));
        assert!(names.contains(&"PT2"));
    }

    #[test]
    fn test_calculate_all_with_dependency_order() {
        let mut wb = Workbook::new();
        // A1 = 10 (value)
        wb.set_cell_value("Sheet1", "A1", 10.0).unwrap();
        // A2 = A1 * 2 (formula depends on A1)
        wb.set_cell_value(
            "Sheet1",
            "A2",
            CellValue::Formula {
                expr: "A1*2".to_string(),
                result: None,
            },
        )
        .unwrap();
        // A3 = A2 + A1 (formula depends on A2 and A1)
        wb.set_cell_value(
            "Sheet1",
            "A3",
            CellValue::Formula {
                expr: "A2+A1".to_string(),
                result: None,
            },
        )
        .unwrap();

        wb.calculate_all().unwrap();

        // A2 should be 20 (10 * 2)
        let a2 = wb.get_cell_value("Sheet1", "A2").unwrap();
        match a2 {
            CellValue::Formula { result, .. } => {
                assert_eq!(*result.unwrap(), CellValue::Number(20.0));
            }
            _ => panic!("A2 should be a formula cell"),
        }

        // A3 should be 30 (20 + 10)
        let a3 = wb.get_cell_value("Sheet1", "A3").unwrap();
        match a3 {
            CellValue::Formula { result, .. } => {
                assert_eq!(*result.unwrap(), CellValue::Number(30.0));
            }
            _ => panic!("A3 should be a formula cell"),
        }
    }

    #[test]
    fn test_calculate_all_no_formulas() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", 10.0).unwrap();
        wb.set_cell_value("Sheet1", "B1", 20.0).unwrap();
        // Should succeed without error when there are no formulas.
        wb.calculate_all().unwrap();
    }

    #[test]
    fn test_calculate_all_cycle_detection() {
        let mut wb = Workbook::new();
        // A1 = B1, B1 = A1
        wb.set_cell_value(
            "Sheet1",
            "A1",
            CellValue::Formula {
                expr: "B1".to_string(),
                result: None,
            },
        )
        .unwrap();
        wb.set_cell_value(
            "Sheet1",
            "B1",
            CellValue::Formula {
                expr: "A1".to_string(),
                result: None,
            },
        )
        .unwrap();

        let result = wb.calculate_all();
        assert!(result.is_err());
        let err_str = result.unwrap_err().to_string();
        assert!(
            err_str.contains("circular reference"),
            "expected circular reference error, got: {err_str}"
        );
    }

    #[test]
    fn test_set_get_doc_props() {
        let mut wb = Workbook::new();
        let props = crate::doc_props::DocProperties {
            title: Some("My Title".to_string()),
            subject: Some("My Subject".to_string()),
            creator: Some("Author".to_string()),
            keywords: Some("rust, excel".to_string()),
            description: Some("A test workbook".to_string()),
            last_modified_by: Some("Editor".to_string()),
            revision: Some("2".to_string()),
            created: Some("2024-01-01T00:00:00Z".to_string()),
            modified: Some("2024-06-01T12:00:00Z".to_string()),
            category: Some("Testing".to_string()),
            content_status: Some("Draft".to_string()),
        };
        wb.set_doc_props(props);

        let got = wb.get_doc_props();
        assert_eq!(got.title.as_deref(), Some("My Title"));
        assert_eq!(got.subject.as_deref(), Some("My Subject"));
        assert_eq!(got.creator.as_deref(), Some("Author"));
        assert_eq!(got.keywords.as_deref(), Some("rust, excel"));
        assert_eq!(got.description.as_deref(), Some("A test workbook"));
        assert_eq!(got.last_modified_by.as_deref(), Some("Editor"));
        assert_eq!(got.revision.as_deref(), Some("2"));
        assert_eq!(got.created.as_deref(), Some("2024-01-01T00:00:00Z"));
        assert_eq!(got.modified.as_deref(), Some("2024-06-01T12:00:00Z"));
        assert_eq!(got.category.as_deref(), Some("Testing"));
        assert_eq!(got.content_status.as_deref(), Some("Draft"));
    }

    #[test]
    fn test_set_get_app_props() {
        let mut wb = Workbook::new();
        let props = crate::doc_props::AppProperties {
            application: Some("SheetKit".to_string()),
            doc_security: Some(0),
            company: Some("Acme Corp".to_string()),
            app_version: Some("1.0.0".to_string()),
            manager: Some("Boss".to_string()),
            template: Some("default.xltx".to_string()),
        };
        wb.set_app_props(props);

        let got = wb.get_app_props();
        assert_eq!(got.application.as_deref(), Some("SheetKit"));
        assert_eq!(got.doc_security, Some(0));
        assert_eq!(got.company.as_deref(), Some("Acme Corp"));
        assert_eq!(got.app_version.as_deref(), Some("1.0.0"));
        assert_eq!(got.manager.as_deref(), Some("Boss"));
        assert_eq!(got.template.as_deref(), Some("default.xltx"));
    }

    #[test]
    fn test_custom_property_crud() {
        let mut wb = Workbook::new();

        // Set
        wb.set_custom_property(
            "Project",
            crate::doc_props::CustomPropertyValue::String("SheetKit".to_string()),
        );

        // Get
        let val = wb.get_custom_property("Project");
        assert_eq!(
            val,
            Some(crate::doc_props::CustomPropertyValue::String(
                "SheetKit".to_string()
            ))
        );

        // Update
        wb.set_custom_property(
            "Project",
            crate::doc_props::CustomPropertyValue::String("Updated".to_string()),
        );
        let val = wb.get_custom_property("Project");
        assert_eq!(
            val,
            Some(crate::doc_props::CustomPropertyValue::String(
                "Updated".to_string()
            ))
        );

        // Delete
        assert!(wb.delete_custom_property("Project"));
        assert!(wb.get_custom_property("Project").is_none());
        assert!(!wb.delete_custom_property("Project")); // already gone
    }

    #[test]
    fn test_doc_props_save_open_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("doc_props.xlsx");

        let mut wb = Workbook::new();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Test Title".to_string()),
            creator: Some("Test Author".to_string()),
            created: Some("2024-01-01T00:00:00Z".to_string()),
            ..Default::default()
        });
        wb.set_app_props(crate::doc_props::AppProperties {
            application: Some("SheetKit".to_string()),
            company: Some("TestCorp".to_string()),
            ..Default::default()
        });
        wb.set_custom_property("Version", crate::doc_props::CustomPropertyValue::Int(42));
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let doc = wb2.get_doc_props();
        assert_eq!(doc.title.as_deref(), Some("Test Title"));
        assert_eq!(doc.creator.as_deref(), Some("Test Author"));
        assert_eq!(doc.created.as_deref(), Some("2024-01-01T00:00:00Z"));

        let app = wb2.get_app_props();
        assert_eq!(app.application.as_deref(), Some("SheetKit"));
        assert_eq!(app.company.as_deref(), Some("TestCorp"));

        let custom = wb2.get_custom_property("Version");
        assert_eq!(custom, Some(crate::doc_props::CustomPropertyValue::Int(42)));
    }

    #[test]
    fn test_open_without_doc_props() {
        // A newly created workbook saved without setting doc props should
        // still open gracefully (core/app/custom properties are all None).
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("no_props.xlsx");

        let wb = Workbook::new();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let doc = wb2.get_doc_props();
        assert!(doc.title.is_none());
        assert!(doc.creator.is_none());

        let app = wb2.get_app_props();
        assert!(app.application.is_none());

        assert!(wb2.get_custom_property("anything").is_none());
    }

    #[test]
    fn test_custom_property_multiple_types() {
        let mut wb = Workbook::new();

        wb.set_custom_property(
            "StringProp",
            crate::doc_props::CustomPropertyValue::String("hello".to_string()),
        );
        wb.set_custom_property("IntProp", crate::doc_props::CustomPropertyValue::Int(-7));
        wb.set_custom_property(
            "FloatProp",
            crate::doc_props::CustomPropertyValue::Float(3.15),
        );
        wb.set_custom_property(
            "BoolProp",
            crate::doc_props::CustomPropertyValue::Bool(true),
        );
        wb.set_custom_property(
            "DateProp",
            crate::doc_props::CustomPropertyValue::DateTime("2024-01-01T00:00:00Z".to_string()),
        );

        assert_eq!(
            wb.get_custom_property("StringProp"),
            Some(crate::doc_props::CustomPropertyValue::String(
                "hello".to_string()
            ))
        );
        assert_eq!(
            wb.get_custom_property("IntProp"),
            Some(crate::doc_props::CustomPropertyValue::Int(-7))
        );
        assert_eq!(
            wb.get_custom_property("FloatProp"),
            Some(crate::doc_props::CustomPropertyValue::Float(3.15))
        );
        assert_eq!(
            wb.get_custom_property("BoolProp"),
            Some(crate::doc_props::CustomPropertyValue::Bool(true))
        );
        assert_eq!(
            wb.get_custom_property("DateProp"),
            Some(crate::doc_props::CustomPropertyValue::DateTime(
                "2024-01-01T00:00:00Z".to_string()
            ))
        );
    }

    #[test]
    fn test_doc_props_default_values() {
        let wb = Workbook::new();
        let doc = wb.get_doc_props();
        assert!(doc.title.is_none());
        assert!(doc.subject.is_none());
        assert!(doc.creator.is_none());
        assert!(doc.keywords.is_none());
        assert!(doc.description.is_none());
        assert!(doc.last_modified_by.is_none());
        assert!(doc.revision.is_none());
        assert!(doc.created.is_none());
        assert!(doc.modified.is_none());
        assert!(doc.category.is_none());
        assert!(doc.content_status.is_none());

        let app = wb.get_app_props();
        assert!(app.application.is_none());
        assert!(app.doc_security.is_none());
        assert!(app.company.is_none());
        assert!(app.app_version.is_none());
        assert!(app.manager.is_none());
        assert!(app.template.is_none());
    }

    #[test]
    fn test_add_sparkline_and_get_sparklines() {
        let mut wb = Workbook::new();
        let config = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1");
        wb.add_sparkline("Sheet1", &config).unwrap();

        let sparklines = wb.get_sparklines("Sheet1").unwrap();
        assert_eq!(sparklines.len(), 1);
        assert_eq!(sparklines[0].data_range, "Sheet1!A1:A10");
        assert_eq!(sparklines[0].location, "B1");
    }

    #[test]
    fn test_add_multiple_sparklines_to_same_sheet() {
        let mut wb = Workbook::new();
        let config1 = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1");
        let config2 = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B2");
        let mut config3 = crate::sparkline::SparklineConfig::new("Sheet1!C1:C10", "D1");
        config3.sparkline_type = crate::sparkline::SparklineType::Column;

        wb.add_sparkline("Sheet1", &config1).unwrap();
        wb.add_sparkline("Sheet1", &config2).unwrap();
        wb.add_sparkline("Sheet1", &config3).unwrap();

        let sparklines = wb.get_sparklines("Sheet1").unwrap();
        assert_eq!(sparklines.len(), 3);
        assert_eq!(
            sparklines[2].sparkline_type,
            crate::sparkline::SparklineType::Column
        );
    }

    #[test]
    fn test_remove_sparkline_by_location() {
        let mut wb = Workbook::new();
        let config1 = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1");
        let config2 = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B2");
        wb.add_sparkline("Sheet1", &config1).unwrap();
        wb.add_sparkline("Sheet1", &config2).unwrap();

        wb.remove_sparkline("Sheet1", "B1").unwrap();

        let sparklines = wb.get_sparklines("Sheet1").unwrap();
        assert_eq!(sparklines.len(), 1);
        assert_eq!(sparklines[0].location, "B2");
    }

    #[test]
    fn test_remove_nonexistent_sparkline_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.remove_sparkline("Sheet1", "Z99");
        assert!(result.is_err());
    }

    #[test]
    fn test_sparkline_on_nonexistent_sheet_returns_error() {
        let mut wb = Workbook::new();
        let config = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1");
        let result = wb.add_sparkline("NoSuchSheet", &config);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));

        let result = wb.get_sparklines("NoSuchSheet");
        assert!(result.is_err());
    }

    #[test]
    fn test_sparkline_save_open_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("sparkline_roundtrip.xlsx");

        let mut wb = Workbook::new();
        for i in 1..=10 {
            wb.set_cell_value(
                "Sheet1",
                &format!("A{i}"),
                CellValue::Number(i as f64 * 10.0),
            )
            .unwrap();
        }

        let mut config = crate::sparkline::SparklineConfig::new("Sheet1!A1:A10", "B1");
        config.sparkline_type = crate::sparkline::SparklineType::Column;
        config.markers = true;
        config.high_point = true;
        config.line_weight = Some(1.5);

        wb.add_sparkline("Sheet1", &config).unwrap();

        let config2 = crate::sparkline::SparklineConfig::new("Sheet1!A1:A5", "C1");
        wb.add_sparkline("Sheet1", &config2).unwrap();

        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let sparklines = wb2.get_sparklines("Sheet1").unwrap();
        assert_eq!(sparklines.len(), 2);
        assert_eq!(sparklines[0].data_range, "Sheet1!A1:A10");
        assert_eq!(sparklines[0].location, "B1");
        assert_eq!(
            sparklines[0].sparkline_type,
            crate::sparkline::SparklineType::Column
        );
        assert!(sparklines[0].markers);
        assert!(sparklines[0].high_point);
        assert_eq!(sparklines[0].line_weight, Some(1.5));
        assert_eq!(sparklines[1].data_range, "Sheet1!A1:A5");
        assert_eq!(sparklines[1].location, "C1");
    }

    #[test]
    fn test_sparkline_empty_sheet_returns_empty_vec() {
        let wb = Workbook::new();
        let sparklines = wb.get_sparklines("Sheet1").unwrap();
        assert!(sparklines.is_empty());
    }

    fn make_table_config(cols: &[&str]) -> crate::table::TableConfig {
        crate::table::TableConfig {
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            range: "A1:D10".to_string(),
            columns: cols
                .iter()
                .map(|c| crate::table::TableColumn {
                    name: c.to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                })
                .collect(),
            ..crate::table::TableConfig::default()
        }
    }

    fn make_slicer_workbook() -> Workbook {
        let mut wb = Workbook::new();
        let table = make_table_config(&["Status", "Region", "Category", "Col1", "Col2"]);
        wb.add_table("Sheet1", &table).unwrap();
        wb
    }

    fn make_slicer_config(name: &str, col: &str) -> crate::slicer::SlicerConfig {
        crate::slicer::SlicerConfig {
            name: name.to_string(),
            cell: "F1".to_string(),
            table_name: "Table1".to_string(),
            column_name: col.to_string(),
            caption: None,
            style: None,
            width: None,
            height: None,
            show_caption: None,
            column_count: None,
        }
    }

    #[test]
    fn test_add_slicer_basic() {
        let mut wb = make_slicer_workbook();
        let config = make_slicer_config("StatusFilter", "Status");
        wb.add_slicer("Sheet1", &config).unwrap();

        let slicers = wb.get_slicers("Sheet1").unwrap();
        assert_eq!(slicers.len(), 1);
        assert_eq!(slicers[0].name, "StatusFilter");
        assert_eq!(slicers[0].column_name, "Status");
        assert_eq!(slicers[0].table_name, "Table1");
    }

    #[test]
    fn test_add_slicer_with_options() {
        let mut wb = make_slicer_workbook();
        let config = crate::slicer::SlicerConfig {
            name: "RegionSlicer".to_string(),
            cell: "G2".to_string(),
            table_name: "Table1".to_string(),
            column_name: "Region".to_string(),
            caption: Some("Filter by Region".to_string()),
            style: Some("SlicerStyleLight1".to_string()),
            width: Some(300),
            height: Some(250),
            show_caption: Some(true),
            column_count: Some(2),
        };
        wb.add_slicer("Sheet1", &config).unwrap();

        let slicers = wb.get_slicers("Sheet1").unwrap();
        assert_eq!(slicers.len(), 1);
        assert_eq!(slicers[0].caption, "Filter by Region");
        assert_eq!(slicers[0].style, Some("SlicerStyleLight1".to_string()));
    }

    #[test]
    fn test_add_slicer_duplicate_name() {
        let mut wb = make_slicer_workbook();
        let config = make_slicer_config("MySlicer", "Status");
        wb.add_slicer("Sheet1", &config).unwrap();

        let result = wb.add_slicer("Sheet1", &config);
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("already exists"));
    }

    #[test]
    fn test_add_slicer_invalid_sheet() {
        let mut wb = make_slicer_workbook();
        let config = make_slicer_config("S1", "Status");
        let result = wb.add_slicer("NoSuchSheet", &config);
        assert!(result.is_err());
    }

    #[test]
    fn test_add_slicer_table_not_found() {
        let mut wb = Workbook::new();
        let config = crate::slicer::SlicerConfig {
            name: "S1".to_string(),
            cell: "F1".to_string(),
            table_name: "NonExistent".to_string(),
            column_name: "Col".to_string(),
            caption: None,
            style: None,
            width: None,
            height: None,
            show_caption: None,
            column_count: None,
        };
        let result = wb.add_slicer("Sheet1", &config);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::TableNotFound { .. }));
    }

    #[test]
    fn test_add_slicer_column_not_found() {
        let mut wb = make_slicer_workbook();
        let config = crate::slicer::SlicerConfig {
            name: "S1".to_string(),
            cell: "F1".to_string(),
            table_name: "Table1".to_string(),
            column_name: "NonExistentColumn".to_string(),
            caption: None,
            style: None,
            width: None,
            height: None,
            show_caption: None,
            column_count: None,
        };
        let result = wb.add_slicer("Sheet1", &config);
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::TableColumnNotFound { .. }
        ));
    }

    #[test]
    fn test_add_slicer_correct_table_id_and_column() {
        let mut wb = make_slicer_workbook();
        let config = make_slicer_config("RegFilter", "Region");
        wb.add_slicer("Sheet1", &config).unwrap();

        // Region is at index 1 (0-based), so column should be 2 (1-based).
        let cache = &wb.slicer_caches[0].1;
        let tsc = cache.table_slicer_cache.as_ref().unwrap();
        assert_eq!(tsc.table_id, 1);
        assert_eq!(tsc.column, 2);
    }

    #[test]
    fn test_get_slicers_resolves_table_name() {
        let mut wb = make_slicer_workbook();
        wb.add_slicer("Sheet1", &make_slicer_config("S1", "Category"))
            .unwrap();

        let slicers = wb.get_slicers("Sheet1").unwrap();
        assert_eq!(slicers.len(), 1);
        assert_eq!(slicers[0].table_name, "Table1");
        assert_eq!(slicers[0].column_name, "Category");
    }

    #[test]
    fn test_get_slicers_empty() {
        let wb = Workbook::new();
        let slicers = wb.get_slicers("Sheet1").unwrap();
        assert!(slicers.is_empty());
    }

    #[test]
    fn test_delete_slicer() {
        let mut wb = make_slicer_workbook();
        let config = make_slicer_config("S1", "Status");
        wb.add_slicer("Sheet1", &config).unwrap();

        assert_eq!(wb.get_slicers("Sheet1").unwrap().len(), 1);

        wb.delete_slicer("Sheet1", "S1").unwrap();
        assert_eq!(wb.get_slicers("Sheet1").unwrap().len(), 0);
    }

    #[test]
    fn test_delete_slicer_not_found() {
        let mut wb = Workbook::new();
        let result = wb.delete_slicer("Sheet1", "NonExistent");
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("not found"));
    }

    #[test]
    fn test_delete_slicer_cleans_content_types() {
        let mut wb = make_slicer_workbook();
        let config = make_slicer_config("S1", "Status");
        wb.add_slicer("Sheet1", &config).unwrap();

        let ct_before = wb.content_types.overrides.len();
        wb.delete_slicer("Sheet1", "S1").unwrap();
        let ct_after = wb.content_types.overrides.len();

        // Two content type overrides (slicer + cache) should be removed.
        assert_eq!(ct_before - ct_after, 2);
    }

    #[test]
    fn test_delete_slicer_cleans_workbook_rels() {
        let mut wb = make_slicer_workbook();
        let config = make_slicer_config("S1", "Status");
        wb.add_slicer("Sheet1", &config).unwrap();

        let has_cache_rel = wb
            .workbook_rels
            .relationships
            .iter()
            .any(|r| r.rel_type == rel_types::SLICER_CACHE);
        assert!(has_cache_rel);

        wb.delete_slicer("Sheet1", "S1").unwrap();

        let has_cache_rel = wb
            .workbook_rels
            .relationships
            .iter()
            .any(|r| r.rel_type == rel_types::SLICER_CACHE);
        assert!(!has_cache_rel);
    }

    #[test]
    fn test_multiple_slicers_on_same_sheet() {
        let mut wb = make_slicer_workbook();
        wb.add_slicer("Sheet1", &make_slicer_config("S1", "Col1"))
            .unwrap();
        wb.add_slicer("Sheet1", &make_slicer_config("S2", "Col2"))
            .unwrap();

        let slicers = wb.get_slicers("Sheet1").unwrap();
        assert_eq!(slicers.len(), 2);
    }

    #[test]
    fn test_slicer_roundtrip() {
        let tmp = TempDir::new().unwrap();
        let path = tmp.path().join("slicer_rt.xlsx");

        let mut wb = make_slicer_workbook();
        wb.add_slicer("Sheet1", &make_slicer_config("MySlicer", "Category"))
            .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let slicers = wb2.get_slicers("Sheet1").unwrap();
        assert_eq!(slicers.len(), 1);
        assert_eq!(slicers[0].name, "MySlicer");
        assert_eq!(slicers[0].column_name, "Category");
        assert_eq!(slicers[0].table_name, "Table1");
    }

    #[test]
    fn test_slicer_content_types_added() {
        let mut wb = make_slicer_workbook();
        wb.add_slicer("Sheet1", &make_slicer_config("S1", "Status"))
            .unwrap();

        let has_slicer_ct = wb
            .content_types
            .overrides
            .iter()
            .any(|o| o.content_type == mime_types::SLICER);
        let has_cache_ct = wb
            .content_types
            .overrides
            .iter()
            .any(|o| o.content_type == mime_types::SLICER_CACHE);

        assert!(has_slicer_ct);
        assert!(has_cache_ct);
    }

    #[test]
    fn test_slicer_worksheet_rels_added() {
        let mut wb = make_slicer_workbook();
        wb.add_slicer("Sheet1", &make_slicer_config("S1", "Status"))
            .unwrap();

        let rels = wb.worksheet_rels.get(&0).unwrap();
        let has_slicer_rel = rels
            .relationships
            .iter()
            .any(|r| r.rel_type == rel_types::SLICER);
        assert!(has_slicer_rel);
    }

    #[test]
    fn test_slicer_error_display() {
        let err = Error::SlicerNotFound {
            name: "Missing".to_string(),
        };
        assert_eq!(err.to_string(), "slicer 'Missing' not found");

        let err = Error::SlicerAlreadyExists {
            name: "Dup".to_string(),
        };
        assert_eq!(err.to_string(), "slicer 'Dup' already exists");
    }

    #[test]
    fn test_add_table_and_get_tables() {
        let mut wb = Workbook::new();
        let table = make_table_config(&["Name", "Age", "City"]);
        wb.add_table("Sheet1", &table).unwrap();

        let tables = wb.get_tables("Sheet1").unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "Table1");
        assert_eq!(tables[0].columns, vec!["Name", "Age", "City"]);
    }

    #[test]
    fn test_add_table_duplicate_name() {
        let mut wb = Workbook::new();
        let table = make_table_config(&["Col"]);
        wb.add_table("Sheet1", &table).unwrap();

        let result = wb.add_table("Sheet1", &table);
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("already exists"));
    }

    #[test]
    fn test_slicer_table_on_wrong_sheet() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        let table = make_table_config(&["Status"]);
        wb.add_table("Sheet2", &table).unwrap();

        let config = make_slicer_config("S1", "Status");
        let result = wb.add_slicer("Sheet1", &config);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::TableNotFound { .. }));
    }
}
