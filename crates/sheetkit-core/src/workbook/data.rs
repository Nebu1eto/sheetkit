use super::io::serialize_xml;
use super::*;
use crate::traversal::{checked_output_add, CheckedRect, TraversalLimits};

struct PivotOutputBudget {
    limits: TraversalLimits,
    used: u64,
}

impl PivotOutputBudget {
    fn new(limits: TraversalLimits) -> Self {
        Self { limits, used: 0 }
    }

    fn charge(&mut self, bytes: u64) -> Result<()> {
        let remaining = self.limits.max_output_bytes().saturating_sub(self.used);
        let admitted = bytes.min(remaining.saturating_add(1));
        self.used = checked_output_add(self.used, admitted, self.limits)?;
        Ok(())
    }

    fn charge_usize(&mut self, bytes: usize) -> Result<()> {
        self.charge(u64::try_from(bytes).unwrap_or(u64::MAX))
    }

    fn charge_array<T>(&mut self, count: usize) -> Result<()> {
        let bytes = count.saturating_mul(std::mem::size_of::<T>());
        self.charge_usize(bytes)
    }
}

struct PivotXmlCounter<'a> {
    budget: &'a mut PivotOutputBudget,
    budget_error: Option<Error>,
}

impl std::fmt::Write for PivotXmlCounter<'_> {
    fn write_str(&mut self, value: &str) -> std::fmt::Result {
        if let Err(error) = self.budget.charge_usize(value.len()) {
            self.budget_error = Some(error);
            return Err(std::fmt::Error);
        }
        Ok(())
    }
}

fn charge_pivot_xml<T: serde::Serialize>(budget: &mut PivotOutputBudget, value: &T) -> Result<()> {
    budget.charge_usize(XML_DECLARATION.len() + 1)?;
    let mut counter = PivotXmlCounter {
        budget,
        budget_error: None,
    };
    let result = value.serialize(quick_xml::se::Serializer::new(&mut counter));
    if let Some(error) = counter.budget_error {
        return Err(error);
    }
    result.map_err(|error| Error::XmlParse(error.to_string()))
}

fn pivot_cache_string_bytes(value: &CellValue) -> usize {
    match value {
        CellValue::Empty => 0,
        CellValue::Bool(value) => {
            if *value {
                4
            } else {
                5
            }
        }
        CellValue::Number(value) | CellValue::Date(value) if value.is_finite() => 0,
        CellValue::String(value) | CellValue::Error(value) => value.len(),
        CellValue::RichString(runs) => runs.iter().map(|run| run.text.len()).sum(),
        CellValue::Formula {
            result: Some(result),
            ..
        } => pivot_cache_string_bytes(result),
        CellValue::Formula { expr, result: None } => expr.len().saturating_add(1),
        CellValue::Number(_) | CellValue::Date(_) => 32,
    }
}

impl Workbook {
    /// Add a pivot table to the workbook.
    ///
    /// The pivot table summarizes data from `config.source_sheet` /
    /// `config.source_range` and places its output on `config.target_sheet`
    /// starting at `config.target_cell`.
    pub fn add_pivot_table(&mut self, config: &PivotTableConfig) -> Result<()> {
        self.add_pivot_table_with_limits(config, TraversalLimits::default())
    }

    /// Add a pivot table with explicit finite source traversal limits.
    pub fn add_pivot_table_with_limits(
        &mut self,
        config: &PivotTableConfig,
        limits: TraversalLimits,
    ) -> Result<()> {
        let mut output_budget = PivotOutputBudget::new(limits);
        Self::charge_pivot_config(config, &mut output_budget)?;
        self.sheet_index(&config.source_sheet)?;
        let target_idx = self.sheet_index(&config.target_sheet)?;

        if self
            .get_pivot_tables()
            .iter()
            .any(|pivot| pivot.name == config.name)
        {
            return Err(Error::PivotTableAlreadyExists {
                name: config.name.clone(),
            });
        }

        // Admit the complete dense source before allocating rows or mutating
        // deferred/typed pivot state.
        let (field_names, source_rows) = self.read_pivot_source_with_limits(
            &config.source_sheet,
            &config.source_range,
            limits,
            &mut output_budget,
        )?;
        if field_names.is_empty() || field_names.iter().any(String::is_empty) {
            return Err(Error::InvalidSourceRange(
                "source range contains an empty header".to_string(),
            ));
        }

        // Lowest unused cache ID across pivot tables and workbook caches;
        // gaps are reused and parsed IDs near u32::MAX cannot overflow.
        let used_cache_ids: std::collections::HashSet<u32> = self
            .pivot_tables
            .iter()
            .map(|(_, pt)| pt.cache_id)
            .chain(
                self.workbook_xml
                    .pivot_caches
                    .iter()
                    .flat_map(|caches| caches.caches.iter().map(|cache| cache.cache_id)),
            )
            .collect();
        let cache_id = (0..=u32::MAX)
            .find(|id| !used_cache_ids.contains(id))
            .ok_or_else(|| Error::InvalidArgument("no unused pivot cache ID available".into()))?;

        let mut occupied_paths = self.occupied_part_paths();
        let allocate_part =
            |prefix: &str,
             has_relationship_part: bool,
             occupied: &mut std::collections::HashSet<String>| {
                let mut number = 1usize;
                loop {
                    let path = format!("{prefix}{number}.xml");
                    let rels_path = relationship_part_path(&path);
                    if !occupied.contains(&path)
                        && (!has_relationship_part || !occupied.contains(&rels_path))
                    {
                        occupied.insert(path.clone());
                        if has_relationship_part {
                            occupied.insert(rels_path);
                        }
                        break path;
                    }
                    number += 1;
                }
            };
        let pt_path = allocate_part("xl/pivotTables/pivotTable", true, &mut occupied_paths);
        let pcd_path = allocate_part(
            "xl/pivotCache/pivotCacheDefinition",
            true,
            &mut occupied_paths,
        );
        let pcr_path = allocate_part(
            "xl/pivotCache/pivotCacheRecords",
            false,
            &mut occupied_paths,
        );

        let cache_records_rid = "rId1".to_string();
        Self::charge_pivot_table_construction(
            config,
            &field_names,
            source_rows.len(),
            &mut output_budget,
        )?;
        self.charge_pivot_cache_construction(&field_names, &source_rows, &mut output_budget)?;
        let (pcd, pcr, item_counts) = crate::pivot::build_pivot_cache_parts(
            &config.source_sheet,
            &config.source_range,
            &field_names,
            &source_rows,
            cache_records_rid.clone(),
        );
        let pt_def = crate::pivot::build_pivot_table_xml_with_items(
            config,
            cache_id,
            &field_names,
            &item_counts,
        )?;

        let cache_rels = Relationships {
            xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
            relationships: vec![Relationship {
                id: cache_records_rid,
                rel_type: rel_types::PIVOT_CACHE_RECORDS.to_string(),
                target: relative_relationship_target(&pcd_path, &pcr_path),
                target_mode: None,
            }],
        };
        let table_rels = Relationships {
            xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
            relationships: vec![Relationship {
                id: "rId1".to_string(),
                rel_type: rel_types::PIVOT_CACHE_DEF.to_string(),
                target: relative_relationship_target(&pt_path, &pcd_path),
                target_mode: None,
            }],
        };
        let cache_rels_bytes = serialize_xml(&cache_rels)?.into_bytes();
        let table_rels_bytes = serialize_xml(&table_rels)?.into_bytes();

        charge_pivot_xml(&mut output_budget, &pcd)?;
        charge_pivot_xml(&mut output_budget, &pcr)?;
        charge_pivot_xml(&mut output_budget, &pt_def)?;
        output_budget.charge_usize(cache_rels_bytes.len())?;
        output_budget.charge_usize(table_rels_bytes.len())?;

        let wb_rid = crate::sheet::next_rid(&self.workbook_rels.relationships);
        let ws_rid = self.next_worksheet_rid(target_idx);
        let ws_target = relative_relationship_target(&self.sheet_part_path(target_idx), &pt_path);

        // All parsing, allocation, construction, and serialization has
        // succeeded. Hydrate and commit the staged graph in one final phase.
        self.hydrate_pivot_tables();
        self.pivot_tables.push((pt_path.clone(), pt_def));
        self.pivot_cache_defs.push((pcd_path.clone(), pcd));
        self.pivot_cache_records.push((pcr_path.clone(), pcr));
        self.raw_graph_parts
            .insert(relationship_part_path(&pt_path), table_rels_bytes);
        self.raw_graph_parts
            .insert(relationship_part_path(&pcd_path), cache_rels_bytes);

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

        self.workbook_rels.relationships.push(Relationship {
            id: wb_rid.clone(),
            rel_type: rel_types::PIVOT_CACHE_DEF.to_string(),
            target: relative_relationship_target("xl/workbook.xml", &pcd_path),
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
            target: ws_target,
            target_mode: None,
        });

        Ok(())
    }

    /// Get information about all pivot tables in the workbook.
    pub fn get_pivot_tables(&self) -> Vec<PivotTableInfo> {
        use crate::workbook::aux::AuxCategory;

        let mut pivot_tables = self.pivot_tables.clone();
        if let Some(entries) = self.deferred_parts.entries(AuxCategory::PivotTables) {
            for (path, bytes) in entries {
                if pivot_tables.iter().any(|(existing, _)| existing == path) {
                    continue;
                }
                let xml_str = String::from_utf8_lossy(bytes);
                if let Ok(pt) = quick_xml::de::from_str::<
                    sheetkit_xml::pivot_table::PivotTableDefinition,
                >(&xml_str)
                {
                    pivot_tables.push((path.clone(), pt));
                }
            }
        }

        let mut pivot_cache_defs = self.pivot_cache_defs.clone();
        if let Some(entries) = self.deferred_parts.entries(AuxCategory::PivotCaches) {
            for (path, bytes) in entries {
                if !path.contains("pivotCacheDefinition")
                    || pivot_cache_defs
                        .iter()
                        .any(|(existing, _)| existing == path)
                {
                    continue;
                }
                let xml_str = String::from_utf8_lossy(bytes);
                if let Ok(pcd) = quick_xml::de::from_str::<
                    sheetkit_xml::pivot_cache::PivotCacheDefinition,
                >(&xml_str)
                {
                    pivot_cache_defs.push((path.clone(), pcd));
                }
            }
        }

        pivot_tables
            .iter()
            .map(|(path, pt)| {
                let cache_path = self.workbook_xml.pivot_caches.as_ref().and_then(|caches| {
                    let entry = caches
                        .caches
                        .iter()
                        .find(|entry| entry.cache_id == pt.cache_id)?;
                    let relationship = self
                        .workbook_rels
                        .relationships
                        .iter()
                        .find(|relationship| relationship.id == entry.r_id)?;
                    Some(resolve_relationship_target(
                        "xl/workbook.xml",
                        &relationship.target,
                    ))
                });
                let (source_sheet, source_range) = cache_path
                    .as_ref()
                    .and_then(|cache_path| {
                        pivot_cache_defs.iter().find(|(path, _)| path == cache_path)
                    })
                    .and_then(|(_, pcd)| {
                        pcd.cache_source
                            .worksheet_source
                            .as_ref()
                            .map(|ws| (ws.sheet.clone(), ws.reference.clone()))
                    })
                    .unwrap_or_default();

                // Determine target sheet from the pivot table path.
                let target_sheet = self
                    .find_pivot_table_target_sheet_by_path(path)
                    .unwrap_or_default();

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
        self.hydrate_pivot_tables();
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
        let pt_rels_path = relationship_part_path(&pt_path);
        self.raw_graph_parts.remove(&pt_rels_path);
        self.dirty_graph_parts.remove(&pt_rels_path);
        self.unknown_parts.retain(|(path, _)| path != &pt_rels_path);

        let cache_still_used = self
            .pivot_tables
            .iter()
            .any(|(_, table)| table.cache_id == cache_id);

        let mut wb_cache_rid = None;
        if !cache_still_used {
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
        }

        if let Some(ref rid) = wb_cache_rid {
            let cache_path = self
                .workbook_rels
                .relationships
                .iter()
                .find(|r| r.id == *rid)
                .map(|relationship| {
                    resolve_relationship_target("xl/workbook.xml", &relationship.target)
                });
            if let Some(cache_path) = cache_path {
                let cache_rels_path = relationship_part_path(&cache_path);
                let relationship_bytes = self
                    .raw_graph_parts
                    .get(&cache_rels_path)
                    .map(Vec::as_slice)
                    .or_else(|| {
                        self.unknown_parts
                            .iter()
                            .find(|(path, _)| path == &cache_rels_path)
                            .map(|(_, bytes)| bytes.as_slice())
                    });
                let records_path = relationship_bytes
                    .and_then(|bytes| quick_xml::de::from_reader::<_, Relationships>(bytes).ok())
                    .and_then(|relationships| {
                        relationships
                            .relationships
                            .into_iter()
                            .find(|relationship| {
                                relationship.rel_type == rel_types::PIVOT_CACHE_RECORDS
                            })
                    })
                    .map(|relationship| {
                        resolve_relationship_target(&cache_path, &relationship.target)
                    });

                self.pivot_cache_defs
                    .retain(|(path, _)| path != &cache_path);
                if let Some(records_path) = records_path {
                    self.pivot_cache_records
                        .retain(|(path, _)| path != &records_path);
                }
                self.raw_graph_parts.remove(&cache_rels_path);
                self.dirty_graph_parts.remove(&cache_rels_path);
                self.unknown_parts
                    .retain(|(path, _)| path != &cache_rels_path);
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
        for rels in self.worksheet_rels.values_mut() {
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
        self.hydrate_sparklines_for_sheet(idx);
        while self.sheet_sparklines.len() <= idx {
            self.sheet_sparklines.push(vec![]);
            self.sheet_sparklines_hydrated.push(true);
        }
        self.sheet_sparklines[idx].push(config.clone());
        self.mark_sheet_dirty(idx);
        Ok(())
    }

    /// Get all sparklines for a worksheet.
    pub fn get_sparklines(&self, sheet: &str) -> Result<Vec<crate::sparkline::SparklineConfig>> {
        let idx = self.sheet_index(sheet)?;
        Ok(self.sparklines_for_sheet(idx))
    }

    /// Remove a sparkline by its location cell reference.
    pub fn remove_sparkline(&mut self, sheet: &str, location: &str) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        self.hydrate_sparklines_for_sheet(idx);
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
        self.mark_sheet_dirty(idx);
        Ok(())
    }

    /// Return sparklines from typed data or a read-only parse of lazy sheet XML.
    fn sparklines_for_sheet(&self, sheet_idx: usize) -> Vec<crate::sparkline::SparklineConfig> {
        if self
            .sheet_sparklines_hydrated
            .get(sheet_idx)
            .copied()
            .unwrap_or(false)
        {
            return self
                .sheet_sparklines
                .get(sheet_idx)
                .cloned()
                .unwrap_or_default();
        }
        if let Some(bytes) = self.raw_sheet_xml.get(sheet_idx).and_then(Option::as_deref) {
            return crate::workbook::io::parse_sparklines_from_xml(&String::from_utf8_lossy(bytes));
        }
        if let Some(bytes) = self
            .raw_sheet_extensions
            .get(sheet_idx)
            .and_then(Option::as_deref)
        {
            return crate::workbook::io::parse_sparklines_from_xml(&String::from_utf8_lossy(bytes));
        }
        if self
            .sheet_package_backed
            .get(sheet_idx)
            .copied()
            .unwrap_or(false)
        {
            self.package_source
                .as_ref()
                .and_then(|source| {
                    source
                        .worksheet_extension_fragment(&self.sheet_part_path(sheet_idx))
                        .ok()
                        .flatten()
                })
                .map(|bytes| {
                    crate::workbook::io::parse_sparklines_from_xml(&String::from_utf8_lossy(&bytes))
                })
                .unwrap_or_default()
        } else {
            Vec::new()
        }
    }

    /// Materialize lazy sparkline data before mutating it.
    pub(super) fn hydrate_sparklines_for_sheet(&mut self, sheet_idx: usize) {
        if self
            .sheet_sparklines_hydrated
            .get(sheet_idx)
            .copied()
            .unwrap_or(false)
        {
            return;
        }
        while self.raw_sheet_extensions.len() <= sheet_idx {
            self.raw_sheet_extensions.push(None);
        }
        if self
            .raw_sheet_xml
            .get(sheet_idx)
            .and_then(Option::as_ref)
            .is_none()
            && self.raw_sheet_extensions[sheet_idx].is_none()
        {
            let extension = if self
                .sheet_package_backed
                .get(sheet_idx)
                .copied()
                .unwrap_or(false)
            {
                self.package_source.as_ref().and_then(|source| {
                    source
                        .worksheet_extension_fragment(&self.sheet_part_path(sheet_idx))
                        .ok()
                        .flatten()
                })
            } else {
                None
            };
            self.raw_sheet_extensions[sheet_idx] = extension;
        }
        let sparklines = self.sparklines_for_sheet(sheet_idx);
        while self.sheet_sparklines.len() <= sheet_idx {
            self.sheet_sparklines.push(vec![]);
            self.sheet_sparklines_hydrated.push(false);
        }
        self.sheet_sparklines[sheet_idx] = sparklines;
        self.sheet_sparklines_hydrated[sheet_idx] = true;
    }

    /// Evaluate a single formula string in the context of `sheet`.
    ///
    /// A [`CellSnapshot`](crate::formula::eval::CellSnapshot) is built from the current workbook state so
    /// that cell references within the formula can be resolved.
    pub fn evaluate_formula(&self, sheet: &str, formula: &str) -> Result<CellValue> {
        self.evaluate_formula_with_limits(
            sheet,
            formula,
            crate::formula::budget::FormulaLimits::secure_default(),
        )
    }

    /// Evaluates one formula with finite, validated resource limits.
    pub fn evaluate_formula_with_limits(
        &self,
        sheet: &str,
        formula: &str,
        limits: crate::formula::budget::FormulaLimits,
    ) -> Result<CellValue> {
        // Validate the sheet exists.
        let _ = self.sheet_index(sheet)?;
        let budget = crate::formula::budget::FormulaBudget::for_calculation(limits)?;
        let parsed = crate::formula::parser::parse_formula_with_budget(formula, &budget)?;
        let snapshot = self.build_cell_snapshot(sheet)?;
        let mut evaluator = crate::formula::eval::Evaluator::with_budget(&snapshot, budget);
        evaluator.eval_expr(&parsed)
    }

    /// Recalculate every formula cell across all sheets and store the
    /// computed result back into each cell. Uses a dependency graph and
    /// topological sort so formulas are evaluated after their dependencies.
    pub fn calculate_all(&mut self) -> Result<()> {
        self.calculate_all_with_limits(crate::formula::budget::FormulaLimits::secure_default())
    }

    /// Recalculate all formulas with finite, validated resource limits.
    pub fn calculate_all_with_limits(
        &mut self,
        limits: crate::formula::budget::FormulaLimits,
    ) -> Result<()> {
        use crate::formula::eval::{
            build_dependency_graph_from_parsed_with_budget, topological_sort_with_budget, CellCoord,
        };

        let budget = crate::formula::budget::FormulaBudget::for_calculation(limits)?;

        let sheet_names: Vec<String> = self.sheet_names().iter().map(|s| s.to_string()).collect();

        // Parse as formulas are discovered so cumulative limits are enforced
        // before retaining attacker-controlled formula collections.
        let mut parsed_cells = Vec::new();
        let mut results = Vec::new();
        for (idx, sn) in sheet_names.iter().enumerate() {
            self.ensure_hydrated(idx)?;
            let ws = self.worksheets[idx].1.get().ok_or_else(|| {
                Error::Internal("formula source sheet is not hydrated".to_string())
            })?;
            for row in &ws.sheet_data.rows {
                for cell in &row.cells {
                    if let Some(ref f) = cell.f {
                        let formula_str = f.value.as_deref().unwrap_or_default();
                        if !formula_str.is_empty() {
                            let (c, r) = cell_name_to_coordinates(cell.r.as_str())?;
                            let coord = CellCoord {
                                sheet: sn.clone(),
                                col: c,
                                row: r,
                            };
                            match crate::formula::parser::parse_formula_with_budget(
                                formula_str,
                                &budget,
                            ) {
                                Ok(expr) => parsed_cells.push((coord, expr)),
                                Err(error @ Error::ResourceLimitExceeded(_)) => return Err(error),
                                Err(_) => {
                                    budget.charge_output_bytes(7)?;
                                    results.push((coord, CellValue::Error("#VALUE!".to_string())))
                                }
                            }
                        }
                    }
                }
            }
        }

        if parsed_cells.is_empty() && results.is_empty() {
            return Ok(());
        }

        // Build the graph only from valid expressions. Invalid formulas are
        // written to the progressive snapshot before dependent expressions run.
        let mut parsed_coords = Vec::new();
        parsed_coords
            .try_reserve_exact(parsed_cells.len())
            .map_err(|_| Error::Internal("formula coordinate allocation failed".to_string()))?;
        parsed_coords.extend(parsed_cells.iter().map(|(coord, _)| coord.clone()));
        let deps = build_dependency_graph_from_parsed_with_budget(&parsed_cells, &budget)?;
        let eval_order = topological_sort_with_budget(&parsed_coords, &deps, &budget)?;

        // Build a lookup from coord to its already-parsed formula expression.
        let mut formula_map: HashMap<CellCoord, crate::formula::ast::Expr> = HashMap::new();
        formula_map
            .try_reserve(parsed_cells.len())
            .map_err(|_| Error::Internal("formula lookup allocation failed".to_string()))?;
        formula_map.extend(parsed_cells);

        // Build a snapshot of all cell data.
        let first_sheet = sheet_names.first().cloned().unwrap_or_default();
        let mut snapshot = self.build_cell_snapshot(&first_sheet)?;

        // Evaluate in dependency order, updating the snapshot progressively
        // so later formulas see already-computed results.
        for (coord, result) in &results {
            snapshot.set_cell(&coord.sheet, coord.col, coord.row, result.clone());
        }

        for coord in &eval_order {
            if let Some(parsed) = formula_map.get(coord) {
                snapshot.set_current_sheet(&coord.sheet);
                let mut evaluator =
                    crate::formula::eval::Evaluator::with_budget(&snapshot, budget.clone());
                evaluator.set_position(coord.col, coord.row);
                let result = match evaluator.eval_expr(parsed) {
                    Ok(result) => result,
                    Err(error) => {
                        let cached = formula_error_to_cached_value(&error).ok_or(error)?;
                        let output_bytes = match &cached {
                            CellValue::String(value) | CellValue::Error(value) => {
                                value.len() as u64
                            }
                            _ => 0,
                        };
                        budget.charge_output_bytes(output_bytes)?;
                        cached
                    }
                };
                snapshot.set_cell(&coord.sheet, coord.col, coord.row, result.clone());
                results.push((coord.clone(), result));
            }
        }

        // Write results back directly to the XML cells, preserving the
        // formula element and storing the computed value in the v/t fields.
        let sheet_name_to_idx: HashMap<String, usize> = sheet_names
            .iter()
            .enumerate()
            .map(|(idx, name)| (name.clone(), idx))
            .collect();

        for (coord, result) in results {
            let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(coord.col, coord.row)?;
            let sheet_idx = *sheet_name_to_idx.get(&coord.sheet).ok_or_else(|| {
                Error::Internal(format!("formula result sheet '{}' is missing", coord.sheet))
            })?;
            let ws = self.worksheets[sheet_idx].1.get_mut().ok_or_else(|| {
                Error::Internal(format!(
                    "formula result sheet '{}' is not hydrated",
                    coord.sheet
                ))
            })?;
            let row = ws
                .sheet_data
                .rows
                .iter_mut()
                .find(|r| r.r == coord.row)
                .ok_or_else(|| {
                    Error::Internal(format!("formula result row {} is missing", coord.row))
                })?;
            let cell = row
                .cells
                .iter_mut()
                .find(|c| c.r == *cell_ref)
                .ok_or_else(|| {
                    Error::Internal(format!("formula result cell {cell_ref} is missing"))
                })?;
            let (new_v, new_t) = match &result {
                CellValue::Number(n) => (Some(n.to_string()), CellTypeTag::None),
                CellValue::String(s) => (Some(s.clone()), CellTypeTag::FormulaString),
                CellValue::Bool(b) => (
                    Some(if *b { "1".to_string() } else { "0".to_string() }),
                    CellTypeTag::Boolean,
                ),
                CellValue::Error(e) => (Some(e.clone()), CellTypeTag::Error),
                CellValue::Date(n) => (Some(n.to_string()), CellTypeTag::None),
                _ => continue,
            };

            let changed = cell.v != new_v || cell.t != new_t;
            if changed {
                cell.v = new_v;
                cell.t = new_t;
                self.mark_sheet_dirty(sheet_idx);
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

    fn charge_pivot_config(
        config: &PivotTableConfig,
        budget: &mut PivotOutputBudget,
    ) -> Result<()> {
        use sheetkit_xml::pivot_table::{DataFieldDef, FieldRef};

        for value in [
            config.name.as_str(),
            config.source_sheet.as_str(),
            config.source_range.as_str(),
            config.target_sheet.as_str(),
            config.target_cell.as_str(),
        ] {
            budget.charge_usize(value.len())?;
        }
        for field in config.rows.iter().chain(&config.columns) {
            budget.charge_usize(field.name.len())?;
        }
        for field in &config.data {
            budget.charge_usize(field.name.len())?;
            if let Some(display_name) = field.display_name.as_ref() {
                budget.charge_usize(display_name.len())?;
            } else {
                budget.charge_usize(
                    field
                        .function
                        .to_xml_str()
                        .len()
                        .saturating_add(" of ".len())
                        .saturating_add(field.name.len()),
                )?;
            }
            budget.charge_usize(field.function.to_xml_str().len())?;
        }

        budget.charge_array::<FieldRef>(config.rows.len().saturating_add(config.columns.len()))?;
        budget.charge_array::<DataFieldDef>(config.data.len())?;
        // The target is uppercased and then copied into the final location.
        budget.charge_usize(
            config
                .target_cell
                .len()
                .saturating_mul(2)
                .saturating_add(11),
        )?;
        Ok(())
    }

    fn charge_pivot_table_construction(
        config: &PivotTableConfig,
        field_names: &[String],
        source_row_count: usize,
        budget: &mut PivotOutputBudget,
    ) -> Result<()> {
        use sheetkit_xml::pivot_table::{FieldItem, PivotFieldDef};

        budget.charge_array::<PivotFieldDef>(field_names.len())?;
        for field_name in field_names {
            let row_field = config.rows.iter().any(|field| field.name == *field_name);
            let column_field = config.columns.iter().any(|field| field.name == *field_name);
            if row_field || column_field {
                budget.charge_array::<FieldItem>(source_row_count.saturating_add(1))?;
                budget.charge_usize("default".len())?;
                budget.charge_usize(if row_field {
                    "axisRow".len()
                } else {
                    "axisCol".len()
                })?;
            }
        }
        Ok(())
    }

    fn pivot_cell_retained_bytes(&self, cell: &sheetkit_xml::worksheet::Cell) -> Result<usize> {
        use sheetkit_xml::worksheet::CellTypeTag;

        if let Some(formula) = cell.f.as_ref() {
            let expression = formula.value.as_ref().map_or(0, String::len);
            let cached = match cell.t {
                CellTypeTag::Error | CellTypeTag::FormulaString => {
                    cell.v.as_ref().map_or(0, String::len)
                }
                _ => 0,
            };
            let cached_value = if cell.v.is_some() {
                std::mem::size_of::<CellValue>()
            } else {
                0
            };
            return Ok(expression
                .saturating_add(cached)
                .saturating_add(cached_value));
        }

        match cell.t {
            CellTypeTag::SharedString => {
                let Some(index) = cell.v.as_deref() else {
                    return Ok(0);
                };
                let index = index
                    .parse::<usize>()
                    .map_err(|_| Error::Internal(format!("invalid SST index: {index}")))?;
                self.sst_runtime
                    .get(index)
                    .map(str::len)
                    .ok_or_else(|| Error::Internal(format!("SST index {index} out of bounds")))
            }
            CellTypeTag::InlineString => Ok(cell
                .is
                .as_ref()
                .and_then(|inline| inline.t.as_ref())
                .map_or(0, String::len)),
            CellTypeTag::Error | CellTypeTag::FormulaString => {
                Ok(cell.v.as_ref().map_or(0, String::len))
            }
            _ => Ok(0),
        }
    }

    fn charge_pivot_cache_construction(
        &self,
        field_names: &[String],
        rows: &[Vec<CellValue>],
        budget: &mut PivotOutputBudget,
    ) -> Result<()> {
        use sheetkit_xml::pivot_cache::{CacheField, CacheRecord, IndexField, NumberField};

        let cell_count = rows.iter().try_fold(0usize, |total, row| {
            total.checked_add(row.len()).ok_or_else(|| {
                Error::Internal("pivot source cell count overflowed usize".to_string())
            })
        })?;
        budget.charge_array::<Vec<CellValue>>(rows.len())?;
        budget.charge_array::<CellValue>(cell_count)?;
        budget.charge_array::<CacheField>(field_names.len())?;
        budget.charge_array::<Vec<u32>>(field_names.len())?;
        budget.charge_array::<u32>(field_names.len())?;
        budget.charge_array::<bool>(field_names.len())?;
        budget.charge_array::<u32>(cell_count)?;
        budget.charge_array::<(&str, u32)>(cell_count)?;
        budget.charge_array::<(u64, u32)>(cell_count)?;
        budget.charge_array::<CacheRecord>(rows.len())?;
        budget.charge_array::<IndexField>(cell_count)?;
        budget.charge_array::<NumberField>(cell_count)?;

        for name in field_names {
            budget.charge_usize(name.len())?;
        }
        for value in rows.iter().flatten() {
            let string_bytes = pivot_cache_string_bytes(value);
            // CacheValue retains one copy. SharedItems may retain another;
            // charging every value is a conservative preflight that avoids
            // allocating an attacker-controlled deduplication set here.
            budget.charge_usize(string_bytes)?;
            budget.charge_usize(string_bytes)?;
        }
        Ok(())
    }

    fn read_pivot_row(
        &self,
        worksheet: &sheetkit_xml::worksheet::WorksheetXml,
        row_number: u32,
        start_col: u32,
        end_col: u32,
        width: usize,
        output_budget: &mut PivotOutputBudget,
    ) -> Result<Vec<CellValue>> {
        let xml_row = worksheet
            .sheet_data
            .rows
            .binary_search_by_key(&row_number, |row| row.r)
            .ok()
            .map(|index| &worksheet.sheet_data.rows[index]);
        output_budget.charge_array::<CellValue>(width)?;
        let mut values = Vec::new();
        values
            .try_reserve_exact(width)
            .map_err(|_| Error::Internal("failed to allocate admitted pivot row".to_string()))?;
        for col in start_col..=end_col {
            let xml_cell = xml_row.and_then(|row| {
                row.cells
                    .binary_search_by_key(&col, |cell| cell.col)
                    .ok()
                    .map(|index| &row.cells[index])
            });
            let value = match xml_cell {
                Some(cell) => {
                    output_budget.charge_usize(self.pivot_cell_retained_bytes(cell)?)?;
                    self.xml_cell_to_value(cell)?
                }
                None => CellValue::Empty,
            };
            values.push(value);
        }
        Ok(values)
    }

    fn read_pivot_source_with_limits(
        &self,
        sheet: &str,
        range: &str,
        limits: TraversalLimits,
        output_budget: &mut PivotOutputBudget,
    ) -> Result<(Vec<String>, Vec<Vec<CellValue>>)> {
        let (start, end) = range
            .split_once(':')
            .ok_or_else(|| Error::InvalidSourceRange(range.to_string()))?;
        let (start_col, start_row) = cell_name_to_coordinates(start)
            .map_err(|_| Error::InvalidSourceRange(range.to_string()))?;
        let (end_col, end_row) = cell_name_to_coordinates(end)
            .map_err(|_| Error::InvalidSourceRange(range.to_string()))?;
        if start_col > end_col || start_row > end_row {
            return Err(Error::InvalidSourceRange(range.to_string()));
        }

        let rect = CheckedRect::new(start_row, end_row, start_col, end_col, limits)?;
        let width = usize::try_from(rect.column_count())
            .map_err(|_| Error::InvalidSourceRange(range.to_string()))?;
        let data_height = usize::try_from(rect.row_count().saturating_sub(1))
            .map_err(|_| Error::InvalidSourceRange(range.to_string()))?;
        let ws = self.worksheet_ref(sheet)?;

        output_budget.charge_array::<String>(width)?;
        output_budget.charge_array::<Vec<CellValue>>(data_height)?;

        let header_values =
            self.read_pivot_row(ws, start_row, start_col, end_col, width, output_budget)?;
        for value in &header_values {
            let allocated = match value {
                CellValue::Number(_) => 32,
                CellValue::Bool(value) => {
                    if *value {
                        4
                    } else {
                        5
                    }
                }
                CellValue::RichString(runs) => runs.iter().map(|run| run.text.len()).sum(),
                _ => 0,
            };
            output_budget.charge_usize(allocated)?;
        }
        let field_names = header_values
            .into_iter()
            .map(|value| match value {
                CellValue::String(value) => value,
                CellValue::Number(value) => value.to_string(),
                CellValue::Bool(value) => value.to_string(),
                CellValue::RichString(runs) => crate::rich_text::rich_text_to_plain(&runs),
                _ => String::new(),
            })
            .collect();

        let mut rows = Vec::new();
        rows.try_reserve_exact(data_height)
            .map_err(|_| Error::Internal("failed to allocate admitted pivot rows".to_string()))?;
        for row in start_row.saturating_add(1)..=end_row {
            rows.push(self.read_pivot_row(ws, row, start_col, end_col, width, output_budget)?);
        }
        Ok((field_names, rows))
    }

    /// Find the target sheet name for a pivot table by its part path.
    fn find_pivot_table_target_sheet_by_path(&self, pt_path: &str) -> Option<String> {
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
        self.hydrate_doc_props();
        self.core_properties = Some(props.to_core_properties());
        self.mark_doc_prop_dirty("docProps/core.xml");
        self.ensure_doc_props_content_types();
    }

    /// Get the core document properties.
    pub fn get_doc_props(&self) -> crate::doc_props::DocProperties {
        use crate::workbook::aux::AuxCategory;

        if let Some(props) = self.core_properties.as_ref() {
            return crate::doc_props::DocProperties::from(props);
        }

        self.deferred_parts
            .get_path(AuxCategory::DocProperties, "docProps/core.xml")
            .and_then(|bytes| {
                let xml_str = String::from_utf8_lossy(bytes);
                sheetkit_xml::doc_props::deserialize_core_properties(&xml_str).ok()
            })
            .as_ref()
            .map(crate::doc_props::DocProperties::from)
            .unwrap_or_default()
    }

    /// Set the application properties (company, app version, etc.).
    pub fn set_app_props(&mut self, props: crate::doc_props::AppProperties) {
        self.hydrate_doc_props();
        self.app_properties = Some(props.to_extended_properties());
        self.mark_doc_prop_dirty("docProps/app.xml");
        self.ensure_doc_props_content_types();
    }

    /// Get the application properties.
    pub fn get_app_props(&self) -> crate::doc_props::AppProperties {
        use crate::workbook::aux::AuxCategory;

        if let Some(props) = self.app_properties.as_ref() {
            return crate::doc_props::AppProperties::from(props);
        }

        self.deferred_parts
            .get_path(AuxCategory::DocProperties, "docProps/app.xml")
            .and_then(|bytes| {
                let xml_str = String::from_utf8_lossy(bytes);
                quick_xml::de::from_str::<sheetkit_xml::doc_props::ExtendedProperties>(&xml_str)
                    .ok()
            })
            .as_ref()
            .map(crate::doc_props::AppProperties::from)
            .unwrap_or_default()
    }

    /// Set a custom property by name. If a property with the same name already
    /// exists, its value is replaced.
    ///
    /// A new property gets the lowest unused pid (>= 2). In the unreachable
    /// case that every pid is taken, the workbook is left unchanged.
    pub fn set_custom_property(
        &mut self,
        name: &str,
        value: crate::doc_props::CustomPropertyValue,
    ) {
        self.hydrate_doc_props();
        let props = self
            .custom_properties
            .get_or_insert_with(sheetkit_xml::doc_props::CustomProperties::default);
        // Exhaustion needs more than four billion properties; the signature
        // has no error channel, so skip the write rather than panic.
        if crate::doc_props::set_custom_property(props, name, value).is_err() {
            return;
        }
        self.mark_doc_prop_dirty("docProps/custom.xml");
        self.ensure_custom_props_content_types();
    }

    /// Get a custom property value by name, or `None` if it does not exist.
    pub fn get_custom_property(&self, name: &str) -> Option<crate::doc_props::CustomPropertyValue> {
        use crate::workbook::aux::AuxCategory;

        if let Some(props) = self.custom_properties.as_ref() {
            return crate::doc_props::find_custom_property(props, name);
        }

        self.deferred_parts
            .get_path(AuxCategory::DocProperties, "docProps/custom.xml")
            .and_then(|bytes| {
                let xml_str = String::from_utf8_lossy(bytes);
                sheetkit_xml::doc_props::deserialize_custom_properties(&xml_str).ok()
            })
            .as_ref()
            .and_then(|props| crate::doc_props::find_custom_property(props, name))
    }

    /// Remove a custom property by name. Returns `true` if a property was
    /// found and removed.
    pub fn delete_custom_property(&mut self, name: &str) -> bool {
        self.hydrate_doc_props();
        if let Some(ref mut props) = self.custom_properties {
            let deleted = crate::doc_props::delete_custom_property(props, name);
            if deleted {
                self.mark_doc_prop_dirty("docProps/custom.xml");
            }
            deleted
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
            let rid = crate::sheet::next_rid(&self.package_rels.relationships);
            self.package_rels.relationships.push(Relationship {
                id: rid,
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

    /// Add a slicer to a sheet targeting a table column.
    ///
    /// Creates the slicer definition, slicer cache, content type overrides,
    /// and worksheet relationships needed for Excel to render the slicer.
    pub fn add_slicer(&mut self, sheet: &str, config: &crate::slicer::SlicerConfig) -> Result<()> {
        use sheetkit_xml::content_types::ContentTypeOverride;
        use sheetkit_xml::slicer::{
            SlicerCacheDefinition, SlicerDefinition, SlicerDefinitions, TableSlicerCache,
        };

        self.hydrate_slicers();
        self.hydrate_tables();
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
        if self
            .workbook_xml
            .defined_names
            .as_ref()
            .is_some_and(|names| {
                names
                    .defined_names
                    .iter()
                    .any(|defined| defined.name.eq_ignore_ascii_case(&cache_name))
            })
        {
            return Err(Error::InvalidArgument(format!(
                "slicer cache name already exists: {cache_name}"
            )));
        }
        let caption = config
            .caption
            .clone()
            .unwrap_or_else(|| config.column_name.clone());

        // Allocate package paths from the complete occupied graph, not vector
        // lengths; reopened workbooks commonly have sparse part numbers.
        let slicer_path = self.next_available_part_path("xl/slicers/slicer", ".xml");
        let cache_path = self.next_available_part_path("xl/slicerCaches/slicerCache", ".xml");

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

        // The placement sheet may differ from the table's owner sheet.
        let _ = table_sheet_idx;

        // CT_TableSlicerCache uses tableColumn@id, not its vector position.
        let column = table_xml
            .table_columns
            .columns
            .iter()
            .find(|c| c.name == config.column_name)
            .ok_or_else(|| Error::TableColumnNotFound {
                table: config.table_name.clone(),
                column: config.column_name.clone(),
            })?;

        let real_table_id = table_xml.id;
        let real_column = column.id;

        // Build the slicer cache definition with real table metadata.
        let slicer_cache = SlicerCacheDefinition {
            name: cache_name.clone(),
            source_name: config.column_name.clone(),
            table_slicer_cache: Some(TableSlicerCache {
                table_id: real_table_id,
                column: real_column,
            }),
        };

        // Prove the anchor can be allocated before committing package state.
        self.validate_slicer_anchor(sheet_idx, &config.cell)?;
        crate::defined_names::set_defined_name(
            &mut self.workbook_xml,
            &cache_name,
            "#N/A",
            crate::defined_names::DefinedNameScope::Workbook,
            None,
        )?;

        // Store parts after all fallible validation has completed.
        self.slicer_defs.push((slicer_path.clone(), slicer_defs));
        self.slicer_caches.push((cache_path.clone(), slicer_cache));

        // Add content type overrides.
        for (path, content_type) in [
            (&slicer_path, mime_types::SLICER),
            (&cache_path, mime_types::SLICER_CACHE),
        ] {
            if !self
                .content_types
                .overrides
                .iter()
                .any(|entry| entry.part_name == format!("/{path}"))
            {
                self.content_types.overrides.push(ContentTypeOverride {
                    part_name: format!("/{path}"),
                    content_type: content_type.to_string(),
                });
            }
        }

        // Add workbook relationship for slicer cache.
        let wb_rid = crate::sheet::next_rid(&self.workbook_rels.relationships);
        self.workbook_rels.relationships.push(Relationship {
            id: wb_rid,
            rel_type: rel_types::SLICER_CACHE.to_string(),
            target: crate::workbook_paths::relative_relationship_target(
                "xl/workbook.xml",
                &cache_path,
            ),
            target_mode: None,
        });

        // Add worksheet relationship for slicer part.
        let ws_rid = self.next_worksheet_rid(sheet_idx);
        let sheet_path = self.sheet_part_path(sheet_idx);
        let slicer_target =
            crate::workbook_paths::relative_relationship_target(&sheet_path, &slicer_path);
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
            target: slicer_target,
            target_mode: None,
        });

        self.add_slicer_anchor(
            sheet_idx,
            &config.cell,
            config.width.unwrap_or(crate::slicer::DEFAULT_WIDTH_PX),
            config.height.unwrap_or(crate::slicer::DEFAULT_HEIGHT_PX),
            &config.name,
            &config
                .caption
                .clone()
                .unwrap_or_else(|| config.column_name.clone()),
        )?;

        Ok(())
    }

    /// Get information about all slicers on a sheet.
    pub fn get_slicers(&self, sheet: &str) -> Result<Vec<crate::slicer::SlicerInfo>> {
        use crate::workbook::aux::AuxCategory;

        let sheet_idx = self.sheet_index(sheet)?;
        let mut result = Vec::new();
        let mut slicer_defs = self.slicer_defs.clone();
        let mut slicer_caches = self.slicer_caches.clone();
        let mut tables = self.tables.clone();

        if let Some(entries) = self.deferred_parts.entries(AuxCategory::Slicers) {
            for (path, bytes) in entries {
                if slicer_defs.iter().any(|(existing, _)| existing == path) {
                    continue;
                }
                let xml_str = String::from_utf8_lossy(bytes);
                if let Ok(sd) =
                    quick_xml::de::from_str::<sheetkit_xml::slicer::SlicerDefinitions>(&xml_str)
                {
                    slicer_defs.push((path.clone(), sd));
                }
            }
        }

        if let Some(entries) = self.deferred_parts.entries(AuxCategory::SlicerCaches) {
            for (path, bytes) in entries {
                if slicer_caches.iter().any(|(existing, _)| existing == path) {
                    continue;
                }
                let xml_str = String::from_utf8_lossy(bytes);
                if let Some(cache) = sheetkit_xml::slicer::parse_slicer_cache(&xml_str) {
                    slicer_caches.push((path.clone(), cache));
                }
            }
        }

        let resolve_table_sheet_idx = |table_path: &str| -> usize {
            for (idx, rels) in &self.worksheet_rels {
                for rel in &rels.relationships {
                    if rel.rel_type != rel_types::TABLE {
                        continue;
                    }
                    let resolved = crate::workbook_paths::resolve_relationship_target(
                        &self.sheet_part_path(*idx),
                        &rel.target,
                    );
                    if resolved == table_path {
                        return *idx;
                    }
                }
            }
            0
        };

        if let Some(entries) = self.deferred_parts.entries(AuxCategory::Tables) {
            for (path, bytes) in entries {
                if tables.iter().any(|(existing, _, _)| existing == path) {
                    continue;
                }
                let xml_str = String::from_utf8_lossy(bytes);
                if let Ok(table_xml) =
                    quick_xml::de::from_str::<sheetkit_xml::table::TableXml>(&xml_str)
                {
                    let table_sheet_idx = resolve_table_sheet_idx(path);
                    tables.push((path.clone(), table_xml, table_sheet_idx));
                }
            }
        }

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

        for (path, sd) in &slicer_defs {
            if !slicer_targets.contains(path) {
                continue;
            }
            for slicer in &sd.slicers {
                // Find the matching cache to get source info.
                let cache = slicer_caches.iter().find(|(_, sc)| sc.name == slicer.cache);

                let (table_name, column_name) = if let Some((_, sc)) = cache {
                    let tname = sc
                        .table_slicer_cache
                        .as_ref()
                        .and_then(|tsc| {
                            tables
                                .iter()
                                .find(|(_, t, _)| t.id == tsc.table_id)
                                .map(|(_, t, _)| t.name.clone())
                        })
                        .unwrap_or_default();
                    let cname = sc
                        .table_slicer_cache
                        .as_ref()
                        .and_then(|tsc| {
                            tables
                                .iter()
                                .find(|(_, t, _)| t.id == tsc.table_id)
                                .and_then(|(_, t, _)| {
                                    t.table_columns
                                        .columns
                                        .iter()
                                        .find(|column| column.id == tsc.column)
                                })
                                .map(|c| c.name.clone())
                        })
                        .unwrap_or_else(|| sc.source_name.clone());
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
        self.hydrate_slicers();
        let sheet_idx = self.sheet_index(sheet)?;
        let sheet_path = self.sheet_part_path(sheet_idx);
        let slicer_targets: Vec<String> = self
            .worksheet_rels
            .get(&sheet_idx)
            .into_iter()
            .flat_map(|relationships| &relationships.relationships)
            .filter(|relationship| relationship.rel_type == rel_types::SLICER)
            .map(|relationship| {
                crate::workbook_paths::resolve_relationship_target(
                    &sheet_path,
                    &relationship.target,
                )
            })
            .collect();

        let sd_idx = self
            .slicer_defs
            .iter()
            .position(|(path, definitions)| {
                slicer_targets.contains(path)
                    && definitions.slicers.iter().any(|slicer| slicer.name == name)
            })
            .ok_or_else(|| Error::SlicerNotFound {
                name: name.to_string(),
            })?;

        self.ensure_owned_drawing_hydratable(sheet_idx)?;
        self.hydrate_drawings();
        self.ensure_owned_drawing_parsed(sheet_idx)?;
        self.ensure_hydrated(sheet_idx)?;

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
            if let Some(rels) = self.worksheet_rels.get_mut(&sheet_idx) {
                rels.relationships.retain(|r| {
                    if r.rel_type != rel_types::SLICER {
                        return true;
                    }
                    let target =
                        crate::workbook_paths::resolve_relationship_target(&sheet_path, &r.target);
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

        let cache_is_still_used = self.slicer_defs.iter().any(|(_, definitions)| {
            definitions
                .slicers
                .iter()
                .any(|slicer| slicer.cache == cache_name)
        });

        // Remove the matching slicer cache after its final view is removed.
        if !cache_name.is_empty() && !cache_is_still_used {
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
                    crate::workbook_paths::resolve_relationship_target("xl/workbook.xml", &r.target)
                        != sc_path
                });
            }
            if let Some(defined_names) = self.workbook_xml.defined_names.as_mut() {
                defined_names
                    .defined_names
                    .retain(|defined| !defined.name.eq_ignore_ascii_case(&cache_name));
                if defined_names.defined_names.is_empty() {
                    self.workbook_xml.defined_names = None;
                }
            }
        }

        // Remove the visible one-cell slicer anchor from the owning drawing.
        let mut empty_drawing = None;
        if let Some(&drawing_idx) = self.worksheet_drawings.get(&sheet_idx) {
            let mut anchor_removed = false;
            let mut anchors_empty = false;
            if let Some((_, drawing)) = self.drawings.get_mut(drawing_idx) {
                let before = drawing.one_cell_anchors.len();
                drawing.one_cell_anchors.retain(|anchor| {
                    anchor.alternate_content.as_ref().is_none_or(|content| {
                        content
                            .choice
                            .graphic_frame
                            .graphic
                            .graphic_data
                            .slicer
                            .name
                            != name
                    })
                });
                anchor_removed = drawing.one_cell_anchors.len() != before;
                anchors_empty =
                    drawing.one_cell_anchors.is_empty() && drawing.two_cell_anchors.is_empty();
            }
            if anchor_removed {
                self.mark_drawing_dirty(drawing_idx);
            }
            if anchors_empty
                && self
                    .drawing_rels
                    .get(&drawing_idx)
                    .is_none_or(|relationships| relationships.relationships.is_empty())
            {
                empty_drawing = Some(drawing_idx);
            }
        }
        if let Some(drawing_idx) = empty_drawing {
            self.remove_empty_drawing(sheet_idx, drawing_idx);
        }
        self.mark_sheet_dirty(sheet_idx);

        Ok(())
    }

    fn remove_empty_drawing(&mut self, sheet_idx: usize, drawing_idx: usize) {
        let last_idx = self.drawings.len() - 1;
        let (drawing_path, _) = self.drawings.swap_remove(drawing_idx);
        let relationships_path = relationship_part_path(&drawing_path);
        self.drawing_rels.remove(&drawing_idx);
        self.remove_graph_part(&drawing_path);
        self.remove_graph_part(&relationships_path);
        self.worksheet_drawings.remove(&sheet_idx);

        if drawing_idx != last_idx {
            if let Some(relationships) = self.drawing_rels.remove(&last_idx) {
                self.drawing_rels.insert(drawing_idx, relationships);
            }
            for existing_idx in self.worksheet_drawings.values_mut() {
                if *existing_idx == last_idx {
                    *existing_idx = drawing_idx;
                }
            }
        }

        self.content_types
            .overrides
            .retain(|override_| override_.part_name != format!("/{drawing_path}"));
        let sheet_path = self.sheet_part_path(sheet_idx);
        if let Some(relationships) = self.worksheet_rels.get_mut(&sheet_idx) {
            relationships.relationships.retain(|relationship| {
                if relationship.rel_type != rel_types::DRAWING {
                    return true;
                }
                crate::workbook_paths::resolve_relationship_target(
                    &sheet_path,
                    &relationship.target,
                ) != drawing_path
            });
        }
        if let Some(worksheet) = self.worksheets[sheet_idx].1.get_mut() {
            worksheet.drawing = None;
        }
    }
}

fn formula_error_to_cached_value(error: &Error) -> Option<CellValue> {
    match error {
        Error::UnknownFunction { .. } => Some(CellValue::Error("#NAME?".to_string())),
        Error::WrongArgCount { .. } => Some(CellValue::Error("#VALUE!".to_string())),
        Error::FormulaError(message) => {
            let code = match message.as_str() {
                "#DIV/0!" | "#N/A" | "#NAME?" | "#NULL!" | "#NUM!" | "#REF!" | "#VALUE!" => {
                    message.as_str()
                }
                _ => "#VALUE!",
            };
            Some(CellValue::Error(code.to_string()))
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::workbook::open_options::{AuxParts, OpenOptions, ReadMode};
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
        assert_eq!(pts[0].location, "E1:F5");
    }

    #[test]
    fn test_lazy_get_pivot_tables_without_mutation() {
        let mut wb = make_pivot_workbook();
        wb.add_pivot_table(&basic_pivot_config()).unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let pts = wb2.get_pivot_tables();
        assert_eq!(pts.len(), 1);
        assert_eq!(pts[0].name, "PivotTable1");
        assert_eq!(pts[0].source_sheet, "Sheet1");
        assert_eq!(pts[0].source_range, "A1:C4");
        assert_eq!(pts[0].target_sheet, "Sheet1");
    }

    #[test]
    fn test_delete_pivot_table() {
        let mut wb = make_pivot_workbook();
        let config = basic_pivot_config();
        wb.add_pivot_table(&config).unwrap();
        let table_rels_path = relationship_part_path(&wb.pivot_tables[0].0);
        let cache_rels_path = relationship_part_path(&wb.pivot_cache_defs[0].0);
        assert_eq!(wb.pivot_tables.len(), 1);

        wb.delete_pivot_table("PivotTable1").unwrap();
        assert!(wb.pivot_tables.is_empty());
        assert!(wb.pivot_cache_defs.is_empty());
        assert!(wb.pivot_cache_records.is_empty());
        assert!(wb.workbook_xml.pivot_caches.is_none());
        assert!(!wb.raw_graph_parts.contains_key(&table_rels_path));
        assert!(!wb.raw_graph_parts.contains_key(&cache_rels_path));
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
        let opts = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let wb2 = Workbook::open_with_options(&path, &opts).unwrap();
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
    fn test_add_pivot_table_avoids_existing_part_and_relationship_paths() {
        let mut wb = make_pivot_workbook();
        wb.unknown_parts.push((
            "xl/pivotTables/_rels/pivotTable1.xml.rels".to_string(),
            b"preserved".to_vec(),
        ));
        wb.unknown_parts.push((
            "xl/pivotCache/pivotCacheDefinition1.xml".to_string(),
            b"preserved".to_vec(),
        ));
        wb.unknown_parts.push((
            "xl/pivotCache/pivotCacheRecords1.xml".to_string(),
            b"preserved".to_vec(),
        ));

        wb.add_pivot_table(&basic_pivot_config()).unwrap();

        assert_eq!(wb.pivot_tables[0].0, "xl/pivotTables/pivotTable2.xml");
        assert_eq!(
            wb.pivot_cache_defs[0].0,
            "xl/pivotCache/pivotCacheDefinition2.xml"
        );
        assert_eq!(
            wb.pivot_cache_records[0].0,
            "xl/pivotCache/pivotCacheRecords2.xml"
        );
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
        assert_eq!(wb.pivot_cache_defs.len(), 1);
        assert_eq!(wb.pivot_cache_records.len(), 1);
        assert!(wb
            .raw_graph_parts
            .contains_key(&relationship_part_path(&wb.pivot_tables[0].0)));
        assert!(wb
            .raw_graph_parts
            .contains_key(&relationship_part_path(&wb.pivot_cache_defs[0].0)));
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
    fn test_pivot_table_populates_cache_records_and_shared_items() {
        let mut wb = make_pivot_workbook();
        wb.add_pivot_table(&basic_pivot_config()).unwrap();

        let pcd = &wb.pivot_cache_defs[0].1;
        assert_eq!(pcd.record_count, Some(3));
        assert_eq!(pcd.refresh_on_load, Some(true));
        assert_eq!(pcd.r_id.as_deref(), Some("rId1"));
        assert_eq!(
            pcd.cache_fields.fields[0]
                .shared_items
                .as_ref()
                .unwrap()
                .count,
            Some(3)
        );
        assert_eq!(
            pcd.cache_fields.fields[1]
                .shared_items
                .as_ref()
                .unwrap()
                .count,
            Some(2)
        );
        assert_eq!(
            pcd.cache_fields.fields[2]
                .shared_items
                .as_ref()
                .unwrap()
                .number_items
                .len(),
            3
        );

        let records = &wb.pivot_cache_records[0].1;
        assert_eq!(records.count, Some(3));
        assert_eq!(records.records.len(), 3);
        assert!(records
            .records
            .iter()
            .all(|record| { record.index_fields.len() == 2 && record.number_fields.len() == 1 }));
    }

    #[test]
    fn test_pivot_table_builds_complete_part_relationship_graph() {
        let mut wb = make_pivot_workbook();
        wb.add_pivot_table(&basic_pivot_config()).unwrap();

        let pt_path = &wb.pivot_tables[0].0;
        let cache_path = &wb.pivot_cache_defs[0].0;
        let pt_rels_path = relationship_part_path(pt_path);
        let cache_rels_path = relationship_part_path(cache_path);
        let pt_rels: Relationships =
            quick_xml::de::from_reader(wb.raw_graph_parts.get(&pt_rels_path).unwrap().as_slice())
                .unwrap();
        let cache_rels: Relationships = quick_xml::de::from_reader(
            wb.raw_graph_parts.get(&cache_rels_path).unwrap().as_slice(),
        )
        .unwrap();

        assert_eq!(pt_rels.relationships.len(), 1);
        assert_eq!(
            pt_rels.relationships[0].rel_type,
            rel_types::PIVOT_CACHE_DEF
        );
        assert_eq!(cache_rels.relationships.len(), 1);
        assert_eq!(
            cache_rels.relationships[0].rel_type,
            rel_types::PIVOT_CACHE_RECORDS
        );
        assert_eq!(
            wb.pivot_cache_defs[0].1.r_id.as_deref(),
            Some(cache_rels.relationships[0].id.as_str())
        );
    }

    #[test]
    fn test_pivot_table_axis_items_and_output_location_range() {
        let mut wb = make_pivot_workbook();
        wb.add_pivot_table(&basic_pivot_config()).unwrap();

        let pt = &wb.pivot_tables[0].1;
        let row_items = pt.pivot_fields.fields[0].items.as_ref().unwrap();
        assert_eq!(row_items.items.len(), 4);
        assert_eq!(
            row_items.items.last().unwrap().item_type.as_deref(),
            Some("default")
        );
        assert_eq!(pt.location.reference, "E1:F5");
    }

    #[test]
    fn test_get_pivot_tables_resolves_cache_through_workbook_relationship() {
        let mut wb = make_pivot_workbook();
        let first = basic_pivot_config();
        wb.add_pivot_table(&first).unwrap();

        let mut second = basic_pivot_config();
        second.name = "PivotTable2".to_string();
        second.source_range = "A1:B4".to_string();
        second.target_cell = "H1".to_string();
        second.data[0].name = "Region".to_string();
        wb.add_pivot_table(&second).unwrap();

        wb.pivot_cache_defs.swap(0, 1);
        let infos = wb.get_pivot_tables();
        assert_eq!(infos[0].source_range, "A1:C4");
        assert_eq!(infos[1].source_range, "A1:B4");
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
    fn pivot_source_limits_accept_exact_area_and_reject_plus_one_atomically() {
        let config = basic_pivot_config();
        let mut exact = make_pivot_workbook();
        let limits = TraversalLimits::new(4, 3, 12, 64 * 1024).unwrap();
        exact.add_pivot_table_with_limits(&config, limits).unwrap();
        assert_eq!(exact.pivot_tables.len(), 1);

        let mut over = make_pivot_workbook();
        let before_graph_parts = over.raw_graph_parts.len();
        let before_overrides = over.content_types.overrides.len();
        let before_workbook_rels = over.workbook_rels.relationships.len();
        let before_worksheet_rels = over.worksheet_rels.len();
        let before_pivot_caches = over
            .workbook_xml
            .pivot_caches
            .as_ref()
            .map_or(0, |caches| caches.caches.len());
        let limits = TraversalLimits::new(4, 3, 11, 64 * 1024).unwrap();
        let error = over
            .add_pivot_table_with_limits(&config, limits)
            .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(
            metadata.dimension(),
            crate::resource::ResourceDimension::CellCount
        );
        assert_eq!(metadata.limit(), 11);
        assert_eq!(metadata.observed(), 12);
        assert!(over.pivot_tables.is_empty());
        assert!(over.pivot_cache_defs.is_empty());
        assert!(over.pivot_cache_records.is_empty());
        assert_eq!(over.raw_graph_parts.len(), before_graph_parts);
        assert_eq!(over.content_types.overrides.len(), before_overrides);
        assert_eq!(over.workbook_rels.relationships.len(), before_workbook_rels);
        assert_eq!(over.worksheet_rels.len(), before_worksheet_rels);
        assert_eq!(
            over.workbook_xml
                .pivot_caches
                .as_ref()
                .map_or(0, |caches| caches.caches.len()),
            before_pivot_caches
        );
    }

    #[test]
    fn pivot_output_budget_accepts_exact_and_rejects_plus_one() {
        let config = basic_pivot_config();
        let accepts = |max_output_bytes| {
            let mut workbook = make_pivot_workbook();
            let limits = TraversalLimits::new(4, 3, 12, max_output_bytes).unwrap();
            workbook.add_pivot_table_with_limits(&config, limits)
        };

        let mut low = 1_u64;
        let mut high = 64 * 1024_u64;
        accepts(high).unwrap();
        while low < high {
            let midpoint = low + (high - low) / 2;
            if accepts(midpoint).is_ok() {
                high = midpoint;
            } else {
                low = midpoint + 1;
            }
        }
        let exact = low;
        accepts(exact).unwrap();

        let error = accepts(exact - 1).unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected output resource limit error");
        };
        assert_eq!(
            metadata.dimension(),
            crate::resource::ResourceDimension::OutputBytes
        );
        assert_eq!(metadata.limit(), exact - 1);
        assert_eq!(metadata.observed(), exact);
    }

    #[test]
    fn pivot_config_strings_are_admitted_before_source_parsing() {
        let mut workbook = make_pivot_workbook();
        let before = workbook.save_to_buffer().unwrap();
        let mut config = basic_pivot_config();
        config.name = "x".repeat(32_000);
        config.source_range = "not-a-range".to_string();

        let limits = TraversalLimits::new(4, 3, 12, 1024).unwrap();
        let error = workbook
            .add_pivot_table_with_limits(&config, limits)
            .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected config output limit error before source parsing");
        };
        assert_eq!(
            metadata.dimension(),
            crate::resource::ResourceDimension::OutputBytes
        );
        assert_eq!(metadata.limit(), 1024);
        assert_eq!(metadata.observed(), 1025);
        assert_eq!(workbook.save_to_buffer().unwrap(), before);
    }

    #[test]
    fn pivot_config_collections_are_admitted_before_definition_allocation() {
        let mut workbook = make_pivot_workbook();
        let before = workbook.save_to_buffer().unwrap();
        let mut config = basic_pivot_config();
        config.source_range = "not-a-range".to_string();
        config.rows = (0..50_000)
            .map(|_| crate::pivot::PivotField {
                name: "Name".to_string(),
            })
            .collect();

        let limits = TraversalLimits::new(4, 3, 12, 64 * 1024).unwrap();
        let error = workbook
            .add_pivot_table_with_limits(&config, limits)
            .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected definition output limit error before source parsing");
        };
        assert_eq!(
            metadata.dimension(),
            crate::resource::ResourceDimension::OutputBytes
        );
        assert_eq!(metadata.limit(), 64 * 1024);
        assert_eq!(metadata.observed(), 64 * 1024 + 1);
        assert_eq!(workbook.save_to_buffer().unwrap(), before);
    }

    #[test]
    fn pivot_repeated_sst_amplification_is_bounded_and_atomic() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};

        let mut workbook = Workbook::new();
        workbook.set_cell_value("Sheet1", "A1", "Category").unwrap();
        let repeated = "x".repeat(32_000);
        for row in 2..=33 {
            workbook
                .set_cell_value("Sheet1", &format!("A{row}"), repeated.as_str())
                .unwrap();
        }
        assert_eq!(workbook.sst_runtime.len(), 2);

        let config = PivotTableConfig {
            name: "RepeatedSst".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1:A33".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "C1".to_string(),
            rows: vec![PivotField {
                name: "Category".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Category".to_string(),
                function: AggregateFunction::Count,
                display_name: None,
            }],
        };
        let before = workbook.save_to_buffer().unwrap();
        let limits = TraversalLimits::new(33, 1, 33, 256 * 1024).unwrap();
        let error = workbook
            .add_pivot_table_with_limits(&config, limits)
            .unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected output resource limit error");
        };
        assert_eq!(
            metadata.dimension(),
            crate::resource::ResourceDimension::OutputBytes
        );
        assert_eq!(metadata.limit(), 256 * 1024);
        assert_eq!(metadata.observed(), 256 * 1024 + 1);
        assert!(workbook.pivot_tables.is_empty());
        assert!(workbook.pivot_cache_defs.is_empty());
        assert!(workbook.pivot_cache_records.is_empty());
        assert_eq!(workbook.save_to_buffer().unwrap(), before);
    }

    #[test]
    fn pivot_source_preserves_far_rows_and_save_reopen_roundtrip() {
        use crate::pivot::{AggregateFunction, PivotDataField, PivotField};

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1048575", "Category").unwrap();
        wb.set_cell_value("Sheet1", "A1048576", "North").unwrap();
        let config = PivotTableConfig {
            name: "FarRows".to_string(),
            source_sheet: "Sheet1".to_string(),
            source_range: "A1048575:A1048576".to_string(),
            target_sheet: "Sheet1".to_string(),
            target_cell: "C1".to_string(),
            rows: vec![PivotField {
                name: "Category".to_string(),
            }],
            columns: vec![],
            data: vec![PivotDataField {
                name: "Category".to_string(),
                function: AggregateFunction::Count,
                display_name: None,
            }],
        };
        let limits = TraversalLimits::new(2, 1, 2, 64 * 1024).unwrap();
        wb.add_pivot_table_with_limits(&config, limits).unwrap();
        assert_eq!(wb.pivot_cache_records[0].1.count, Some(1));

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("far-row-pivot.xlsx");
        wb.save(&path).unwrap();
        let reopened = Workbook::open(&path).unwrap();
        let pivots = reopened.get_pivot_tables();
        assert_eq!(pivots.len(), 1);
        assert_eq!(pivots[0].source_range, "A1048575:A1048576");
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

        let opts = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let wb2 = Workbook::open_with_options(&path, &opts).unwrap();
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
    fn test_add_slicer_drawing_id_overflow_is_atomic() {
        let mut wb = make_slicer_workbook();
        wb.add_shape(
            "Sheet1",
            &crate::shape::ShapeConfig {
                shape_type: crate::shape::ShapeType::Rect,
                from_cell: "J1".to_string(),
                to_cell: "K2".to_string(),
                text: None,
                fill_color: None,
                line_color: None,
                line_width: None,
            },
        )
        .unwrap();
        wb.drawings[0].1.two_cell_anchors[0]
            .shape
            .as_mut()
            .unwrap()
            .nv_sp_pr
            .c_nv_pr
            .id = u32::MAX;
        let before = wb.save_to_buffer().unwrap();

        let result = wb.add_slicer("Sheet1", &make_slicer_config("S1", "Status"));

        assert!(
            matches!(result, Err(Error::InvalidArgument(message)) if message == "drawing object ID overflow")
        );
        assert!(wb.slicer_defs.is_empty());
        assert!(wb.slicer_caches.is_empty());
        assert!(wb.workbook_xml.defined_names.is_none());
        assert_eq!(wb.save_to_buffer().unwrap(), before);
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
    fn test_calculate_all_isolates_formula_errors() {
        let mut wb = Workbook::new();
        for (cell, expr) in [
            ("A1", "1+1"),
            ("B1", "1+\"abc\""),
            ("C1", "NONEXISTENT(1)"),
            ("D1", "SUM(A:A)"),
            ("E1", "A1+3"),
            ("F1", "B1+1"),
            ("G1", "#N/A+1"),
        ] {
            wb.set_cell_value(
                "Sheet1",
                cell,
                CellValue::Formula {
                    expr: expr.to_string(),
                    result: None,
                },
            )
            .unwrap();
        }

        wb.calculate_all().unwrap();

        let cached = |cell: &str| match wb.get_cell_value("Sheet1", cell).unwrap() {
            CellValue::Formula {
                result: Some(result),
                ..
            } => *result,
            value => panic!("{cell} should have a cached formula result, got {value:?}"),
        };

        assert_eq!(cached("A1"), CellValue::Number(2.0));
        assert_eq!(cached("B1"), CellValue::Error("#VALUE!".to_string()));
        assert_eq!(cached("C1"), CellValue::Error("#NAME?".to_string()));
        assert_eq!(cached("D1"), CellValue::Error("#VALUE!".to_string()));
        assert_eq!(cached("E1"), CellValue::Number(5.0));
        assert_eq!(cached("F1"), CellValue::Error("#VALUE!".to_string()));
        assert_eq!(cached("G1"), CellValue::Error("#N/A".to_string()));
    }

    #[test]
    fn test_calculate_all_caches_range_errors_and_continues() {
        let mut wb = Workbook::new();
        for (cell, expr) in [
            ("A1", "1/0"),
            ("B1", "SUM(A1:A2)"),
            ("C1", "2+2"),
            ("C5", "ROW()"),
            ("D5", "COLUMN()"),
        ] {
            wb.set_cell_value(
                "Sheet1",
                cell,
                CellValue::Formula {
                    expr: expr.to_string(),
                    result: None,
                },
            )
            .unwrap();
        }

        wb.calculate_all().unwrap();

        let cached = |cell: &str| match wb.get_cell_value("Sheet1", cell).unwrap() {
            CellValue::Formula {
                result: Some(result),
                ..
            } => *result,
            value => panic!("{cell} should have a cached formula result, got {value:?}"),
        };

        assert_eq!(cached("A1"), CellValue::Error("#DIV/0!".to_string()));
        assert_eq!(cached("B1"), CellValue::Error("#DIV/0!".to_string()));
        assert_eq!(cached("C1"), CellValue::Number(4.0));
        assert_eq!(cached("C5"), CellValue::Number(5.0));
        assert_eq!(cached("D5"), CellValue::Number(4.0));
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
    fn test_calculate_all_no_formulas_keeps_lazy_sheet_clean() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", 10.0).unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert!(!wb2.is_sheet_dirty(0));

        wb2.calculate_all().unwrap();
        assert!(
            !wb2.is_sheet_dirty(0),
            "calculate_all without formulas must not dirty the sheet"
        );
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
    fn test_calculate_all_resource_failure_is_atomic() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "A1",
            CellValue::Formula {
                expr: "1+1".to_string(),
                result: Some(Box::new(CellValue::Number(7.0))),
            },
        )
        .unwrap();
        wb.set_cell_value(
            "Sheet1",
            "B1",
            CellValue::Formula {
                expr: "A1+1".to_string(),
                result: Some(Box::new(CellValue::Number(8.0))),
            },
        )
        .unwrap();
        let buffer = wb.save_to_buffer().unwrap();
        let options = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb = Workbook::open_from_buffer_with_options(&buffer, &options).unwrap();
        let before_a1 = wb.get_cell_value("Sheet1", "A1").unwrap();
        let before_b1 = wb.get_cell_value("Sheet1", "B1").unwrap();
        assert!(!wb.is_sheet_dirty(0));
        let limits = crate::formula::budget::FormulaLimits {
            operation_input_bytes: 3,
            ..crate::formula::budget::FormulaLimits::secure_default()
        };

        let error = wb.calculate_all_with_limits(limits).unwrap_err();

        assert_eq!(
            error.resource_metadata().unwrap().dimension_code(),
            "input_bytes"
        );
        assert_eq!(wb.get_cell_value("Sheet1", "A1").unwrap(), before_a1);
        assert_eq!(wb.get_cell_value("Sheet1", "B1").unwrap(), before_b1);
        assert!(!wb.is_sheet_dirty(0));
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
    fn test_add_pivot_table_after_parsed_u32_max_cache_id() {
        let mut wb = make_pivot_workbook();
        wb.add_pivot_table(&basic_pivot_config()).unwrap();
        wb.pivot_tables[0].1.cache_id = u32::MAX;
        wb.workbook_xml.pivot_caches.as_mut().unwrap().caches[0].cache_id = u32::MAX;

        let mut config = basic_pivot_config();
        config.name = "PivotTable2".to_string();
        config.target_cell = "I1".to_string();
        wb.add_pivot_table(&config).unwrap();

        assert_eq!(wb.pivot_tables[1].1.cache_id, 0);
        let caches = &wb.workbook_xml.pivot_caches.as_ref().unwrap().caches;
        let ids: Vec<u32> = caches.iter().map(|cache| cache.cache_id).collect();
        assert_eq!(ids, vec![u32::MAX, 0]);
    }

    #[test]
    fn test_add_pivot_table_reuses_lowest_free_cache_id() {
        let mut wb = make_pivot_workbook();
        wb.add_pivot_table(&basic_pivot_config()).unwrap();
        wb.pivot_tables[0].1.cache_id = 5;
        wb.workbook_xml.pivot_caches.as_mut().unwrap().caches[0].cache_id = 5;

        let mut config = basic_pivot_config();
        config.name = "PivotTable2".to_string();
        config.target_cell = "I1".to_string();
        wb.add_pivot_table(&config).unwrap();

        assert_eq!(wb.pivot_tables[1].1.cache_id, 0);
    }

    #[test]
    fn test_custom_property_package_rel_id_does_not_collide() {
        let mut wb = Workbook::new();
        let count = wb.package_rels.relationships.len();
        wb.package_rels.relationships[0].id = format!("rId{}", count + 1);

        wb.set_custom_property("Key", crate::doc_props::CustomPropertyValue::Int(1));

        let ids: std::collections::HashSet<&str> = wb
            .package_rels
            .relationships
            .iter()
            .map(|rel| rel.id.as_str())
            .collect();
        assert_eq!(ids.len(), wb.package_rels.relationships.len());
        assert!(wb
            .package_rels
            .relationships
            .iter()
            .any(|rel| rel.rel_type == rel_types::CUSTOM_PROPERTIES));
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
    fn test_lazy_get_doc_properties_without_mutation() {
        let mut wb = Workbook::new();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Deferred Title".to_string()),
            creator: Some("Deferred Author".to_string()),
            ..Default::default()
        });
        wb.set_app_props(crate::doc_props::AppProperties {
            company: Some("Deferred Corp".to_string()),
            application: Some("SheetKit".to_string()),
            ..Default::default()
        });
        wb.set_custom_property(
            "DeferredVersion",
            crate::doc_props::CustomPropertyValue::Int(7),
        );
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        let doc = wb2.get_doc_props();
        assert_eq!(doc.title.as_deref(), Some("Deferred Title"));
        assert_eq!(doc.creator.as_deref(), Some("Deferred Author"));

        let app = wb2.get_app_props();
        assert_eq!(app.company.as_deref(), Some("Deferred Corp"));
        assert_eq!(app.application.as_deref(), Some("SheetKit"));

        assert_eq!(
            wb2.get_custom_property("DeferredVersion"),
            Some(crate::doc_props::CustomPropertyValue::Int(7))
        );
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

        let opts = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let wb2 = Workbook::open_with_options(&path, &opts).unwrap();
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
    fn lazy_sparklines_match_eager_and_preserve_existing_configs_on_mutation() {
        let mut source = Workbook::new();
        let mut existing = crate::sparkline::SparklineConfig::new("Sheet1!A1:A3", "B1");
        existing.sparkline_type = crate::sparkline::SparklineType::Column;
        existing.markers = true;
        existing.line_weight = Some(1.5);
        source.add_sparkline("Sheet1", &existing).unwrap();
        let bytes = source.save_to_buffer().unwrap();

        let eager_options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let eager = Workbook::open_from_buffer_with_options(&bytes, &eager_options).unwrap();
        let mut lazy = Workbook::open_from_buffer(&bytes).unwrap();
        let lazy_configs = lazy.get_sparklines("Sheet1").unwrap();
        let eager_configs = eager.get_sparklines("Sheet1").unwrap();
        assert_eq!(lazy_configs.len(), 1);
        assert_eq!(lazy_configs[0].data_range, eager_configs[0].data_range);
        assert_eq!(lazy_configs[0].location, eager_configs[0].location);
        assert_eq!(
            lazy_configs[0].sparkline_type,
            eager_configs[0].sparkline_type
        );
        assert_eq!(lazy_configs[0].markers, eager_configs[0].markers);
        assert_eq!(lazy_configs[0].line_weight, eager_configs[0].line_weight);

        let added = crate::sparkline::SparklineConfig::new("Sheet1!C1:C3", "D1");
        lazy.add_sparkline("Sheet1", &added).unwrap();
        let reopened = Workbook::open_from_buffer(&lazy.save_to_buffer().unwrap()).unwrap();
        let configs = reopened.get_sparklines("Sheet1").unwrap();
        assert_eq!(configs.len(), 2);
        assert_eq!(configs[0].data_range, existing.data_range);
        assert_eq!(configs[0].location, existing.location);
        assert_eq!(configs[0].sparkline_type, existing.sparkline_type);
        assert_eq!(configs[0].markers, existing.markers);
        assert_eq!(configs[0].line_weight, existing.line_weight);
        assert_eq!(configs[1].data_range, added.data_range);
        assert_eq!(configs[1].location, added.location);

        let mut lazy = Workbook::open_from_buffer(&bytes).unwrap();
        lazy.remove_sparkline("Sheet1", "B1").unwrap();
        assert!(lazy.get_sparklines("Sheet1").unwrap().is_empty());
        let reopened = Workbook::open_from_buffer(&lazy.save_to_buffer().unwrap()).unwrap();
        assert!(reopened.get_sparklines("Sheet1").unwrap().is_empty());
    }

    #[test]
    fn test_sparkline_empty_sheet_returns_empty_vec() {
        let wb = Workbook::new();
        let sparklines = wb.get_sparklines("Sheet1").unwrap();
        assert!(sparklines.is_empty());
    }

    fn make_table_config(cols: &[&str]) -> crate::table::TableConfig {
        let end = crate::utils::cell_ref::coordinates_to_cell_name(cols.len() as u32, 10)
            .expect("test table columns must fit within worksheet bounds");
        crate::table::TableConfig {
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            range: format!("A1:{end}"),
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
    fn test_get_slicers_resolves_nonsequential_table_column_id() {
        let mut wb = make_slicer_workbook();
        wb.tables[0].1.table_columns.columns[2].id = 17;
        wb.add_slicer("Sheet1", &make_slicer_config("S1", "Category"))
            .unwrap();

        let slicers = wb.get_slicers("Sheet1").unwrap();
        assert_eq!(slicers[0].column_name, "Category");
        assert_eq!(
            wb.slicer_caches[0]
                .1
                .table_slicer_cache
                .as_ref()
                .unwrap()
                .column,
            17
        );
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
    fn test_delete_slicer_rejects_name_from_another_sheet_without_mutation() {
        let mut wb = make_slicer_workbook();
        wb.new_sheet("Other").unwrap();
        wb.add_slicer("Sheet1", &make_slicer_config("S1", "Status"))
            .unwrap();
        let before = wb.save_to_buffer().unwrap();

        let result = wb.delete_slicer("Other", "S1");

        assert!(matches!(result, Err(Error::SlicerNotFound { .. })));
        assert_eq!(wb.get_slicers("Sheet1").unwrap().len(), 1);
        assert_eq!(wb.save_to_buffer().unwrap(), before);
    }

    #[test]
    fn test_delete_slicer_cleans_content_types() {
        let mut wb = make_slicer_workbook();
        let config = make_slicer_config("S1", "Status");
        wb.add_slicer("Sheet1", &config).unwrap();

        let ct_before = wb.content_types.overrides.len();
        wb.delete_slicer("Sheet1", "S1").unwrap();
        let ct_after = wb.content_types.overrides.len();

        // The slicer, cache, and now-empty drawing overrides are removed.
        assert_eq!(ct_before - ct_after, 3);
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
    fn test_delete_slicer_preserves_a_cache_used_by_another_view() {
        let mut wb = make_slicer_workbook();
        wb.add_slicer("Sheet1", &make_slicer_config("S1", "Status"))
            .unwrap();
        wb.add_slicer("Sheet1", &make_slicer_config("S2", "Region"))
            .unwrap();
        let shared_cache = wb.slicer_defs[0].1.slicers[0].cache.clone();
        wb.slicer_defs[1].1.slicers[0].cache = shared_cache.clone();

        wb.delete_slicer("Sheet1", "S1").unwrap();

        assert!(wb
            .slicer_caches
            .iter()
            .any(|(_, cache)| cache.name == shared_cache));
        assert!(wb
            .workbook_rels
            .relationships
            .iter()
            .any(|relationship| relationship.rel_type == rel_types::SLICER_CACHE));
        assert!(wb.workbook_xml.defined_names.as_ref().is_some_and(|names| {
            names
                .defined_names
                .iter()
                .any(|defined| defined.name == shared_cache)
        }));
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
    fn test_lazy_get_slicers_without_mutation() {
        let mut wb = make_slicer_workbook();
        wb.add_slicer("Sheet1", &make_slicer_config("LazySlicer", "Status"))
            .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let slicers = wb2.get_slicers("Sheet1").unwrap();
        assert_eq!(slicers.len(), 1);
        assert_eq!(slicers[0].name, "LazySlicer");
        assert_eq!(slicers[0].column_name, "Status");
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
    fn test_slicer_can_be_placed_away_from_its_source_table() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        let table = make_table_config(&["Status"]);
        wb.add_table("Sheet2", &table).unwrap();

        let config = make_slicer_config("S1", "Status");
        wb.add_slicer("Sheet1", &config).unwrap();
        assert_eq!(wb.get_slicers("Sheet1").unwrap().len(), 1);
    }
}
