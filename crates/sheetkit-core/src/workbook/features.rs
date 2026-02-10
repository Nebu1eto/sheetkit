use super::*;

impl Workbook {
    /// Add a data validation rule to a sheet.
    pub fn add_data_validation(
        &mut self,
        sheet: &str,
        config: &DataValidationConfig,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::validation::add_validation(ws, config)
    }

    /// Get all data validation rules for a sheet.
    pub fn get_data_validations(&self, sheet: &str) -> Result<Vec<DataValidationConfig>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::validation::get_validations(ws))
    }

    /// Remove a data validation rule matching the given cell range from a sheet.
    pub fn remove_data_validation(&mut self, sheet: &str, sqref: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::validation::remove_validation(ws, sqref)
    }

    /// Set conditional formatting rules on a cell range of a sheet.
    pub fn set_conditional_format(
        &mut self,
        sheet: &str,
        sqref: &str,
        rules: &[ConditionalFormatRule],
    ) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        let ws = &mut self.worksheets[idx].1;
        crate::conditional::set_conditional_format(ws, &mut self.stylesheet, sqref, rules)
    }

    /// Get all conditional formatting rules for a sheet.
    ///
    /// Returns a list of `(sqref, rules)` pairs.
    pub fn get_conditional_formats(
        &self,
        sheet: &str,
    ) -> Result<Vec<(String, Vec<ConditionalFormatRule>)>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::conditional::get_conditional_formats(
            ws,
            &self.stylesheet,
        ))
    }

    /// Delete conditional formatting rules for a specific cell range on a sheet.
    pub fn delete_conditional_format(&mut self, sheet: &str, sqref: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::conditional::delete_conditional_format(ws, sqref)
    }

    /// Add a comment to a cell on the given sheet.
    ///
    /// A VML drawing part is generated automatically when saving so that
    /// the comment renders correctly in Excel.
    pub fn add_comment(&mut self, sheet: &str, config: &CommentConfig) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        crate::comment::add_comment(&mut self.sheet_comments[idx], config);
        // Invalidate cached VML so save() regenerates it from current comments.
        if idx < self.sheet_vml.len() {
            self.sheet_vml[idx] = None;
        }
        Ok(())
    }

    /// Get all comments for a sheet.
    pub fn get_comments(&self, sheet: &str) -> Result<Vec<CommentConfig>> {
        let idx = self.sheet_index(sheet)?;
        Ok(crate::comment::get_all_comments(&self.sheet_comments[idx]))
    }

    /// Remove a comment from a cell on the given sheet.
    ///
    /// When the last comment on a sheet is removed, the VML drawing part is
    /// cleaned up automatically during save.
    pub fn remove_comment(&mut self, sheet: &str, cell: &str) -> Result<()> {
        let idx = self.sheet_index(sheet)?;
        crate::comment::remove_comment(&mut self.sheet_comments[idx], cell);
        // Invalidate cached VML so save() regenerates or omits it.
        if idx < self.sheet_vml.len() {
            self.sheet_vml[idx] = None;
        }
        Ok(())
    }

    /// Set an auto-filter on a sheet for the given cell range.
    pub fn set_auto_filter(&mut self, sheet: &str, range: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::table::set_auto_filter(ws, range)
    }

    /// Remove the auto-filter from a sheet.
    pub fn remove_auto_filter(&mut self, sheet: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::table::remove_auto_filter(ws);
        Ok(())
    }

    /// Add a table to a sheet.
    ///
    /// Creates the table XML part, adds the appropriate relationship and
    /// content type entries. The table name must be unique within the workbook.
    pub fn add_table(&mut self, sheet: &str, config: &crate::table::TableConfig) -> Result<()> {
        crate::table::validate_table_config(config)?;
        let sheet_idx = self.sheet_index(sheet)?;

        // Check for duplicate table name across the entire workbook.
        if self.tables.iter().any(|(_, t, _)| t.name == config.name) {
            return Err(Error::TableAlreadyExists {
                name: config.name.clone(),
            });
        }

        // Assign a unique table ID (max existing + 1).
        let table_id = self.tables.iter().map(|(_, t, _)| t.id).max().unwrap_or(0) + 1;

        let table_num = self.tables.len() + 1;
        let table_path = format!("xl/tables/table{}.xml", table_num);
        let table_xml = crate::table::build_table_xml(config, table_id);

        self.tables.push((table_path, table_xml, sheet_idx));
        Ok(())
    }

    /// List all tables on a sheet.
    ///
    /// Returns metadata for each table associated with the given sheet.
    pub fn get_tables(&self, sheet: &str) -> Result<Vec<crate::table::TableInfo>> {
        let sheet_idx = self.sheet_index(sheet)?;
        let infos = self
            .tables
            .iter()
            .filter(|(_, _, idx)| *idx == sheet_idx)
            .map(|(_, table_xml, _)| crate::table::table_xml_to_info(table_xml))
            .collect();
        Ok(infos)
    }

    /// Delete a table from a sheet by name.
    ///
    /// Removes the table part, relationship, and content type entries.
    pub fn delete_table(&mut self, sheet: &str, table_name: &str) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;

        let pos = self
            .tables
            .iter()
            .position(|(_, t, idx)| t.name == table_name && *idx == sheet_idx);
        match pos {
            Some(i) => {
                self.tables.remove(i);
                Ok(())
            }
            None => Err(Error::TableNotFound {
                name: table_name.to_string(),
            }),
        }
    }

    /// Set freeze panes on a sheet.
    ///
    /// The cell reference indicates the top-left cell of the scrollable area.
    /// For example, `"A2"` freezes row 1, `"B1"` freezes column A, and `"B2"`
    /// freezes both row 1 and column A.
    pub fn set_panes(&mut self, sheet: &str, cell: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::sheet::set_panes(ws, cell)
    }

    /// Remove any freeze or split panes from a sheet.
    pub fn unset_panes(&mut self, sheet: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::sheet::unset_panes(ws);
        Ok(())
    }

    /// Get the current freeze pane cell reference for a sheet, if any.
    ///
    /// Returns the top-left cell of the unfrozen area (e.g., `"A2"` if row 1
    /// is frozen), or `None` if no panes are configured.
    pub fn get_panes(&self, sheet: &str) -> Result<Option<String>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::sheet::get_panes(ws))
    }

    /// Set page margins on a sheet.
    pub fn set_page_margins(
        &mut self,
        sheet: &str,
        margins: &crate::page_layout::PageMarginsConfig,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::set_page_margins(ws, margins)
    }

    /// Get page margins for a sheet, returning Excel defaults if not set.
    pub fn get_page_margins(&self, sheet: &str) -> Result<crate::page_layout::PageMarginsConfig> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_page_margins(ws))
    }

    /// Set page setup options (orientation, paper size, scale, fit-to-page).
    ///
    /// Only non-`None` parameters are applied; existing values for `None`
    /// parameters are preserved.
    pub fn set_page_setup(
        &mut self,
        sheet: &str,
        orientation: Option<crate::page_layout::Orientation>,
        paper_size: Option<crate::page_layout::PaperSize>,
        scale: Option<u32>,
        fit_to_width: Option<u32>,
        fit_to_height: Option<u32>,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::set_page_setup(
            ws,
            orientation,
            paper_size,
            scale,
            fit_to_width,
            fit_to_height,
        )
    }

    /// Get the page orientation for a sheet.
    pub fn get_orientation(&self, sheet: &str) -> Result<Option<crate::page_layout::Orientation>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_orientation(ws))
    }

    /// Get the paper size for a sheet.
    pub fn get_paper_size(&self, sheet: &str) -> Result<Option<crate::page_layout::PaperSize>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_paper_size(ws))
    }

    /// Get scale, fit-to-width, and fit-to-height values for a sheet.
    ///
    /// Returns `(scale, fit_to_width, fit_to_height)`, each `None` if not set.
    pub fn get_page_setup_details(
        &self,
        sheet: &str,
    ) -> Result<(Option<u32>, Option<u32>, Option<u32>)> {
        let ws = self.worksheet_ref(sheet)?;
        Ok((
            crate::page_layout::get_scale(ws),
            crate::page_layout::get_fit_to_width(ws),
            crate::page_layout::get_fit_to_height(ws),
        ))
    }

    /// Set header and footer text for printing.
    pub fn set_header_footer(
        &mut self,
        sheet: &str,
        header: Option<&str>,
        footer: Option<&str>,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::set_header_footer(ws, header, footer)
    }

    /// Get the header and footer text for a sheet.
    pub fn get_header_footer(&self, sheet: &str) -> Result<(Option<String>, Option<String>)> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_header_footer(ws))
    }

    /// Set print options on a sheet.
    pub fn set_print_options(
        &mut self,
        sheet: &str,
        grid_lines: Option<bool>,
        headings: Option<bool>,
        h_centered: Option<bool>,
        v_centered: Option<bool>,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::set_print_options(ws, grid_lines, headings, h_centered, v_centered)
    }

    /// Get print options for a sheet.
    ///
    /// Returns `(grid_lines, headings, horizontal_centered, vertical_centered)`.
    #[allow(clippy::type_complexity)]
    pub fn get_print_options(
        &self,
        sheet: &str,
    ) -> Result<(Option<bool>, Option<bool>, Option<bool>, Option<bool>)> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_print_options(ws))
    }

    /// Insert a horizontal page break before the given 1-based row.
    pub fn insert_page_break(&mut self, sheet: &str, row: u32) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::insert_page_break(ws, row)
    }

    /// Remove a horizontal page break at the given 1-based row.
    pub fn remove_page_break(&mut self, sheet: &str, row: u32) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::page_layout::remove_page_break(ws, row)
    }

    /// Get all row page break positions (1-based row numbers).
    pub fn get_page_breaks(&self, sheet: &str) -> Result<Vec<u32>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::page_layout::get_page_breaks(ws))
    }

    /// Set a hyperlink on a cell.
    ///
    /// For external URLs and email links, a relationship entry is created in
    /// the worksheet's `.rels` file. Internal sheet references use only the
    /// `location` attribute without a relationship.
    pub fn set_cell_hyperlink(
        &mut self,
        sheet: &str,
        cell: &str,
        link: crate::hyperlink::HyperlinkType,
        display: Option<&str>,
        tooltip: Option<&str>,
    ) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let ws = &mut self.worksheets[sheet_idx].1;
        let rels = self
            .worksheet_rels
            .entry(sheet_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        crate::hyperlink::set_cell_hyperlink(ws, rels, cell, &link, display, tooltip)
    }

    /// Get hyperlink information for a cell.
    ///
    /// Returns `None` if the cell has no hyperlink.
    pub fn get_cell_hyperlink(
        &self,
        sheet: &str,
        cell: &str,
    ) -> Result<Option<crate::hyperlink::HyperlinkInfo>> {
        let sheet_idx = self.sheet_index(sheet)?;
        let ws = &self.worksheets[sheet_idx].1;
        let empty_rels = Relationships {
            xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
            relationships: vec![],
        };
        let rels = self.worksheet_rels.get(&sheet_idx).unwrap_or(&empty_rels);
        crate::hyperlink::get_cell_hyperlink(ws, rels, cell)
    }

    /// Delete a hyperlink from a cell.
    ///
    /// Removes both the hyperlink element from the worksheet XML and any
    /// associated relationship entry.
    pub fn delete_cell_hyperlink(&mut self, sheet: &str, cell: &str) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let ws = &mut self.worksheets[sheet_idx].1;
        let rels = self
            .worksheet_rels
            .entry(sheet_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        crate::hyperlink::delete_cell_hyperlink(ws, rels, cell)
    }

    /// Protect the workbook structure and/or windows.
    pub fn protect_workbook(&mut self, config: WorkbookProtectionConfig) {
        let password_hash = config.password.as_ref().map(|p| {
            let hash = crate::protection::legacy_password_hash(p);
            format!("{:04X}", hash)
        });
        self.workbook_xml.workbook_protection = Some(WorkbookProtection {
            workbook_password: password_hash,
            lock_structure: if config.lock_structure {
                Some(true)
            } else {
                None
            },
            lock_windows: if config.lock_windows {
                Some(true)
            } else {
                None
            },
            revisions_password: None,
            lock_revision: if config.lock_revision {
                Some(true)
            } else {
                None
            },
        });
    }

    /// Remove workbook protection.
    pub fn unprotect_workbook(&mut self) {
        self.workbook_xml.workbook_protection = None;
    }

    /// Check if the workbook is protected.
    pub fn is_workbook_protected(&self) -> bool {
        self.workbook_xml.workbook_protection.is_some()
    }

    /// Resolve a theme color by index (0-11) with optional tint.
    /// Returns the ARGB hex string (e.g. "FF4472C4") or None if the index is out of range.
    pub fn get_theme_color(&self, index: u32, tint: Option<f64>) -> Option<String> {
        crate::theme::resolve_theme_color(&self.theme_colors, index, tint)
    }

    /// Add or update a defined name in the workbook.
    ///
    /// If `scope` is `None`, the name is workbook-scoped (visible from all sheets).
    /// If `scope` is `Some(sheet_name)`, it is sheet-scoped using the sheet's 0-based index.
    /// If a name with the same name and scope already exists, its value and comment are updated.
    pub fn set_defined_name(
        &mut self,
        name: &str,
        value: &str,
        scope: Option<&str>,
        comment: Option<&str>,
    ) -> Result<()> {
        let dn_scope = self.resolve_defined_name_scope(scope)?;
        crate::defined_names::set_defined_name(
            &mut self.workbook_xml,
            name,
            value,
            dn_scope,
            comment,
        )
    }

    /// Get a defined name by name and scope.
    ///
    /// If `scope` is `None`, looks for a workbook-scoped name.
    /// If `scope` is `Some(sheet_name)`, looks for a sheet-scoped name.
    /// Returns `None` if no matching defined name is found.
    pub fn get_defined_name(
        &self,
        name: &str,
        scope: Option<&str>,
    ) -> Result<Option<crate::defined_names::DefinedNameInfo>> {
        let dn_scope = self.resolve_defined_name_scope(scope)?;
        Ok(crate::defined_names::get_defined_name(
            &self.workbook_xml,
            name,
            dn_scope,
        ))
    }

    /// List all defined names in the workbook.
    pub fn get_all_defined_names(&self) -> Vec<crate::defined_names::DefinedNameInfo> {
        crate::defined_names::get_all_defined_names(&self.workbook_xml)
    }

    /// Delete a defined name by name and scope.
    ///
    /// Returns an error if the name does not exist for the given scope.
    pub fn delete_defined_name(&mut self, name: &str, scope: Option<&str>) -> Result<()> {
        let dn_scope = self.resolve_defined_name_scope(scope)?;
        crate::defined_names::delete_defined_name(&mut self.workbook_xml, name, dn_scope)
    }

    /// Protect a sheet with optional password and permission settings.
    ///
    /// Delegates to [`crate::sheet::protect_sheet`] after looking up the sheet.
    pub fn protect_sheet(
        &mut self,
        sheet: &str,
        config: &crate::sheet::SheetProtectionConfig,
    ) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::sheet::protect_sheet(ws, config)
    }

    /// Remove sheet protection.
    pub fn unprotect_sheet(&mut self, sheet: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::sheet::unprotect_sheet(ws)
    }

    /// Check if a sheet is protected.
    pub fn is_sheet_protected(&self, sheet: &str) -> Result<bool> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::sheet::is_sheet_protected(ws))
    }

    /// Resolve an optional sheet name to a [`DefinedNameScope`](crate::defined_names::DefinedNameScope).
    fn resolve_defined_name_scope(
        &self,
        scope: Option<&str>,
    ) -> Result<crate::defined_names::DefinedNameScope> {
        match scope {
            None => Ok(crate::defined_names::DefinedNameScope::Workbook),
            Some(sheet_name) => {
                let idx = self.sheet_index(sheet_name)?;
                Ok(crate::defined_names::DefinedNameScope::Sheet(idx as u32))
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn test_workbook_add_data_validation() {
        let mut wb = Workbook::new();
        let config =
            crate::validation::DataValidationConfig::dropdown("A1:A100", &["Yes", "No", "Maybe"]);
        wb.add_data_validation("Sheet1", &config).unwrap();

        let validations = wb.get_data_validations("Sheet1").unwrap();
        assert_eq!(validations.len(), 1);
        assert_eq!(validations[0].sqref, "A1:A100");
    }

    #[test]
    fn test_workbook_remove_data_validation() {
        let mut wb = Workbook::new();
        let config1 = crate::validation::DataValidationConfig::dropdown("A1:A100", &["Yes", "No"]);
        let config2 = crate::validation::DataValidationConfig::whole_number("B1:B100", 1, 100);
        wb.add_data_validation("Sheet1", &config1).unwrap();
        wb.add_data_validation("Sheet1", &config2).unwrap();

        wb.remove_data_validation("Sheet1", "A1:A100").unwrap();

        let validations = wb.get_data_validations("Sheet1").unwrap();
        assert_eq!(validations.len(), 1);
        assert_eq!(validations[0].sqref, "B1:B100");
    }

    #[test]
    fn test_workbook_data_validation_sheet_not_found() {
        let mut wb = Workbook::new();
        let config = crate::validation::DataValidationConfig::dropdown("A1:A100", &["Yes", "No"]);
        let result = wb.add_data_validation("NoSheet", &config);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_data_validation_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("validation_roundtrip.xlsx");

        let mut wb = Workbook::new();
        let config =
            crate::validation::DataValidationConfig::dropdown("A1:A50", &["Red", "Blue", "Green"]);
        wb.add_data_validation("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let validations = wb2.get_data_validations("Sheet1").unwrap();
        assert_eq!(validations.len(), 1);
        assert_eq!(validations[0].sqref, "A1:A50");
        assert_eq!(
            validations[0].validation_type,
            crate::validation::ValidationType::List
        );
    }

    #[test]
    fn test_workbook_add_comment() {
        let mut wb = Workbook::new();
        let config = crate::comment::CommentConfig {
            cell: "A1".to_string(),
            author: "Alice".to_string(),
            text: "Test comment".to_string(),
        };
        wb.add_comment("Sheet1", &config).unwrap();

        let comments = wb.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell, "A1");
        assert_eq!(comments[0].author, "Alice");
        assert_eq!(comments[0].text, "Test comment");
    }

    #[test]
    fn test_workbook_remove_comment() {
        let mut wb = Workbook::new();
        let config = crate::comment::CommentConfig {
            cell: "A1".to_string(),
            author: "Alice".to_string(),
            text: "Test comment".to_string(),
        };
        wb.add_comment("Sheet1", &config).unwrap();
        wb.remove_comment("Sheet1", "A1").unwrap();

        let comments = wb.get_comments("Sheet1").unwrap();
        assert!(comments.is_empty());
    }

    #[test]
    fn test_workbook_multiple_comments() {
        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Alice".to_string(),
                text: "First".to_string(),
            },
        )
        .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "B2".to_string(),
                author: "Bob".to_string(),
                text: "Second".to_string(),
            },
        )
        .unwrap();

        let comments = wb.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 2);
    }

    #[test]
    fn test_workbook_comment_sheet_not_found() {
        let mut wb = Workbook::new();
        let config = crate::comment::CommentConfig {
            cell: "A1".to_string(),
            author: "Alice".to_string(),
            text: "Test".to_string(),
        };
        let result = wb.add_comment("NoSheet", &config);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_comment_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("comment_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "A saved comment".to_string(),
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        // Verify the comments XML was written to the ZIP.
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(
            archive.by_name("xl/comments1.xml").is_ok(),
            "comments1.xml should be present in the ZIP"
        );
    }

    #[test]
    fn test_workbook_comment_roundtrip_open() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("comment_roundtrip_open.xlsx");

        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "Persist me".to_string(),
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let comments = wb2.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell, "A1");
        assert_eq!(comments[0].author, "Author");
        assert_eq!(comments[0].text, "Persist me");
    }

    #[test]
    fn test_workbook_comment_produces_vml_part() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("comment_vml.xlsx");

        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "B3".to_string(),
                author: "Tester".to_string(),
                text: "VML check".to_string(),
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(
            archive.by_name("xl/drawings/vmlDrawing1.vml").is_ok(),
            "vmlDrawing1.vml should be present in the ZIP"
        );

        // Verify the VML content references the correct cell.
        let mut vml_data = Vec::new();
        archive
            .by_name("xl/drawings/vmlDrawing1.vml")
            .unwrap()
            .read_to_end(&mut vml_data)
            .unwrap();
        let vml_str = String::from_utf8(vml_data).unwrap();
        assert!(vml_str.contains("<x:Row>2</x:Row>"));
        assert!(vml_str.contains("<x:Column>1</x:Column>"));
        assert!(vml_str.contains("ObjectType=\"Note\""));
    }

    #[test]
    fn test_workbook_comment_vml_roundtrip_open() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("comment_vml_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "Roundtrip VML".to_string(),
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        // Reopen and re-save.
        let wb2 = Workbook::open(&path).unwrap();
        let path2 = dir.path().join("comment_vml_roundtrip2.xlsx");
        wb2.save(&path2).unwrap();

        // Verify VML part is preserved through the round-trip.
        let file = std::fs::File::open(&path2).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/drawings/vmlDrawing1.vml").is_ok());

        // Comments should still be readable.
        let wb3 = Workbook::open(&path2).unwrap();
        let comments = wb3.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].text, "Roundtrip VML");
    }

    #[test]
    fn test_workbook_comment_vml_legacy_drawing_ref() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("comment_vml_legacy_ref.xlsx");

        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "C5".to_string(),
                author: "Author".to_string(),
                text: "Legacy drawing test".to_string(),
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        // Verify the worksheet XML contains a legacyDrawing element.
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        let mut ws_data = Vec::new();
        archive
            .by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_end(&mut ws_data)
            .unwrap();
        let ws_str = String::from_utf8(ws_data).unwrap();
        assert!(
            ws_str.contains("legacyDrawing"),
            "worksheet should contain legacyDrawing element"
        );
    }

    #[test]
    fn test_workbook_comment_vml_cleanup_on_last_remove() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("comment_vml_cleanup.xlsx");

        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "Will be removed".to_string(),
            },
        )
        .unwrap();
        wb.remove_comment("Sheet1", "A1").unwrap();
        wb.save(&path).unwrap();

        // Verify no VML part when all comments are removed.
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(
            archive.by_name("xl/drawings/vmlDrawing1.vml").is_err(),
            "vmlDrawing1.vml should not be present when there are no comments"
        );
    }

    #[test]
    fn test_workbook_multiple_comments_vml() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("multi_comment_vml.xlsx");

        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Alice".to_string(),
                text: "First".to_string(),
            },
        )
        .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "D10".to_string(),
                author: "Bob".to_string(),
                text: "Second".to_string(),
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        let mut vml_data = Vec::new();
        archive
            .by_name("xl/drawings/vmlDrawing1.vml")
            .unwrap()
            .read_to_end(&mut vml_data)
            .unwrap();
        let vml_str = String::from_utf8(vml_data).unwrap();
        // Should have two shapes.
        assert!(vml_str.contains("_x0000_s1025"));
        assert!(vml_str.contains("_x0000_s1026"));
    }

    #[test]
    fn test_workbook_set_auto_filter() {
        let mut wb = Workbook::new();
        wb.set_auto_filter("Sheet1", "A1:D10").unwrap();

        let ws = wb.worksheet_ref("Sheet1").unwrap();
        assert!(ws.auto_filter.is_some());
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:D10");
    }

    #[test]
    fn test_workbook_remove_auto_filter() {
        let mut wb = Workbook::new();
        wb.set_auto_filter("Sheet1", "A1:D10").unwrap();
        wb.remove_auto_filter("Sheet1").unwrap();

        let ws = wb.worksheet_ref("Sheet1").unwrap();
        assert!(ws.auto_filter.is_none());
    }

    #[test]
    fn test_workbook_auto_filter_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.set_auto_filter("NoSheet", "A1:D10");
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_workbook_auto_filter_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("autofilter_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.set_auto_filter("Sheet1", "A1:C50").unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let ws = wb2.worksheet_ref("Sheet1").unwrap();
        assert!(ws.auto_filter.is_some());
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:C50");
    }

    #[test]
    fn test_protect_unprotect_workbook() {
        let mut wb = Workbook::new();
        assert!(!wb.is_workbook_protected());

        wb.protect_workbook(crate::protection::WorkbookProtectionConfig {
            password: None,
            lock_structure: true,
            lock_windows: false,
            lock_revision: false,
        });
        assert!(wb.is_workbook_protected());

        wb.unprotect_workbook();
        assert!(!wb.is_workbook_protected());
    }

    #[test]
    fn test_protect_workbook_with_password() {
        let mut wb = Workbook::new();
        wb.protect_workbook(crate::protection::WorkbookProtectionConfig {
            password: Some("secret".to_string()),
            lock_structure: true,
            lock_windows: false,
            lock_revision: false,
        });

        let prot = wb.workbook_xml.workbook_protection.as_ref().unwrap();
        assert!(prot.workbook_password.is_some());
        let hash_str = prot.workbook_password.as_ref().unwrap();
        // Should be a 4-character uppercase hex string
        assert_eq!(hash_str.len(), 4);
        assert!(hash_str.chars().all(|c| c.is_ascii_hexdigit()));
        assert_eq!(prot.lock_structure, Some(true));
    }

    #[test]
    fn test_protect_workbook_structure_only() {
        let mut wb = Workbook::new();
        wb.protect_workbook(crate::protection::WorkbookProtectionConfig {
            password: None,
            lock_structure: true,
            lock_windows: false,
            lock_revision: false,
        });

        let prot = wb.workbook_xml.workbook_protection.as_ref().unwrap();
        assert!(prot.workbook_password.is_none());
        assert_eq!(prot.lock_structure, Some(true));
        assert!(prot.lock_windows.is_none());
        assert!(prot.lock_revision.is_none());
    }

    #[test]
    fn test_protect_workbook_save_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("protected.xlsx");

        let mut wb = Workbook::new();
        wb.protect_workbook(crate::protection::WorkbookProtectionConfig {
            password: Some("hello".to_string()),
            lock_structure: true,
            lock_windows: true,
            lock_revision: false,
        });
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert!(wb2.is_workbook_protected());
        let prot = wb2.workbook_xml.workbook_protection.as_ref().unwrap();
        assert!(prot.workbook_password.is_some());
        assert_eq!(prot.lock_structure, Some(true));
        assert_eq!(prot.lock_windows, Some(true));
    }

    #[test]
    fn test_is_workbook_protected() {
        let wb = Workbook::new();
        assert!(!wb.is_workbook_protected());

        let mut wb2 = Workbook::new();
        wb2.protect_workbook(crate::protection::WorkbookProtectionConfig {
            password: None,
            lock_structure: false,
            lock_windows: false,
            lock_revision: false,
        });
        // Even with no locks, the protection element is present
        assert!(wb2.is_workbook_protected());
    }

    #[test]
    fn test_unprotect_already_unprotected() {
        let mut wb = Workbook::new();
        assert!(!wb.is_workbook_protected());
        // Should be a no-op, not panic
        wb.unprotect_workbook();
        assert!(!wb.is_workbook_protected());
    }

    #[test]
    fn test_set_and_get_external_hyperlink() {
        use crate::hyperlink::HyperlinkType;

        let mut wb = Workbook::new();
        wb.set_cell_hyperlink(
            "Sheet1",
            "A1",
            HyperlinkType::External("https://example.com".to_string()),
            Some("Example"),
            Some("Visit Example"),
        )
        .unwrap();

        let info = wb.get_cell_hyperlink("Sheet1", "A1").unwrap().unwrap();
        assert_eq!(
            info.link_type,
            HyperlinkType::External("https://example.com".to_string())
        );
        assert_eq!(info.display, Some("Example".to_string()));
        assert_eq!(info.tooltip, Some("Visit Example".to_string()));
    }

    #[test]
    fn test_set_and_get_internal_hyperlink() {
        use crate::hyperlink::HyperlinkType;

        let mut wb = Workbook::new();
        wb.new_sheet("Data").unwrap();
        wb.set_cell_hyperlink(
            "Sheet1",
            "B2",
            HyperlinkType::Internal("Data!A1".to_string()),
            Some("Go to Data"),
            None,
        )
        .unwrap();

        let info = wb.get_cell_hyperlink("Sheet1", "B2").unwrap().unwrap();
        assert_eq!(
            info.link_type,
            HyperlinkType::Internal("Data!A1".to_string())
        );
        assert_eq!(info.display, Some("Go to Data".to_string()));
    }

    #[test]
    fn test_set_and_get_email_hyperlink() {
        use crate::hyperlink::HyperlinkType;

        let mut wb = Workbook::new();
        wb.set_cell_hyperlink(
            "Sheet1",
            "C3",
            HyperlinkType::Email("mailto:user@example.com".to_string()),
            None,
            None,
        )
        .unwrap();

        let info = wb.get_cell_hyperlink("Sheet1", "C3").unwrap().unwrap();
        assert_eq!(
            info.link_type,
            HyperlinkType::Email("mailto:user@example.com".to_string())
        );
    }

    #[test]
    fn test_delete_hyperlink_via_workbook() {
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

        wb.delete_cell_hyperlink("Sheet1", "A1").unwrap();

        let info = wb.get_cell_hyperlink("Sheet1", "A1").unwrap();
        assert!(info.is_none());
    }

    #[test]
    fn test_hyperlink_roundtrip_save_open() {
        use crate::hyperlink::HyperlinkType;

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("hyperlink.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_hyperlink(
            "Sheet1",
            "A1",
            HyperlinkType::External("https://rust-lang.org".to_string()),
            Some("Rust"),
            Some("Rust Homepage"),
        )
        .unwrap();
        wb.set_cell_hyperlink(
            "Sheet1",
            "B1",
            HyperlinkType::Internal("Sheet1!C1".to_string()),
            Some("Go to C1"),
            None,
        )
        .unwrap();
        wb.set_cell_hyperlink(
            "Sheet1",
            "C1",
            HyperlinkType::Email("mailto:hello@example.com".to_string()),
            Some("Email"),
            None,
        )
        .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();

        // External link roundtrip.
        let a1 = wb2.get_cell_hyperlink("Sheet1", "A1").unwrap().unwrap();
        assert_eq!(
            a1.link_type,
            HyperlinkType::External("https://rust-lang.org".to_string())
        );
        assert_eq!(a1.display, Some("Rust".to_string()));
        assert_eq!(a1.tooltip, Some("Rust Homepage".to_string()));

        // Internal link roundtrip.
        let b1 = wb2.get_cell_hyperlink("Sheet1", "B1").unwrap().unwrap();
        assert_eq!(
            b1.link_type,
            HyperlinkType::Internal("Sheet1!C1".to_string())
        );
        assert_eq!(b1.display, Some("Go to C1".to_string()));

        // Email link roundtrip.
        let c1 = wb2.get_cell_hyperlink("Sheet1", "C1").unwrap().unwrap();
        assert_eq!(
            c1.link_type,
            HyperlinkType::Email("mailto:hello@example.com".to_string())
        );
        assert_eq!(c1.display, Some("Email".to_string()));
    }

    #[test]
    fn test_hyperlink_on_nonexistent_sheet() {
        use crate::hyperlink::HyperlinkType;

        let mut wb = Workbook::new();
        let result = wb.set_cell_hyperlink(
            "NoSheet",
            "A1",
            HyperlinkType::External("https://example.com".to_string()),
            None,
            None,
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_set_defined_name_workbook_scope() {
        let mut wb = Workbook::new();
        wb.set_defined_name("SalesData", "Sheet1!$A$1:$D$10", None, None)
            .unwrap();

        let info = wb.get_defined_name("SalesData", None).unwrap().unwrap();
        assert_eq!(info.name, "SalesData");
        assert_eq!(info.value, "Sheet1!$A$1:$D$10");
        assert_eq!(info.scope, crate::defined_names::DefinedNameScope::Workbook);
        assert!(info.comment.is_none());
    }

    #[test]
    fn test_set_defined_name_sheet_scope() {
        let mut wb = Workbook::new();
        wb.set_defined_name("LocalRange", "Sheet1!$B$2:$C$5", Some("Sheet1"), None)
            .unwrap();

        let info = wb
            .get_defined_name("LocalRange", Some("Sheet1"))
            .unwrap()
            .unwrap();
        assert_eq!(info.name, "LocalRange");
        assert_eq!(info.value, "Sheet1!$B$2:$C$5");
        assert_eq!(info.scope, crate::defined_names::DefinedNameScope::Sheet(0));
    }

    #[test]
    fn test_update_existing_defined_name() {
        let mut wb = Workbook::new();
        wb.set_defined_name("DataRange", "Sheet1!$A$1:$A$10", None, None)
            .unwrap();

        wb.set_defined_name("DataRange", "Sheet1!$A$1:$A$50", None, Some("Updated"))
            .unwrap();

        let all = wb.get_all_defined_names();
        assert_eq!(all.len(), 1, "should not duplicate the entry");
        assert_eq!(all[0].value, "Sheet1!$A$1:$A$50");
        assert_eq!(all[0].comment, Some("Updated".to_string()));
    }

    #[test]
    fn test_get_all_defined_names() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();

        wb.set_defined_name("Alpha", "Sheet1!$A$1", None, None)
            .unwrap();
        wb.set_defined_name("Beta", "Sheet1!$B$1", Some("Sheet1"), None)
            .unwrap();
        wb.set_defined_name("Gamma", "Sheet2!$C$1", Some("Sheet2"), None)
            .unwrap();

        let all = wb.get_all_defined_names();
        assert_eq!(all.len(), 3);
        assert_eq!(all[0].name, "Alpha");
        assert_eq!(all[1].name, "Beta");
        assert_eq!(all[2].name, "Gamma");
    }

    #[test]
    fn test_delete_defined_name() {
        let mut wb = Workbook::new();
        wb.set_defined_name("ToDelete", "Sheet1!$A$1", None, None)
            .unwrap();
        assert!(wb.get_defined_name("ToDelete", None).unwrap().is_some());

        wb.delete_defined_name("ToDelete", None).unwrap();
        assert!(wb.get_defined_name("ToDelete", None).unwrap().is_none());
    }

    #[test]
    fn test_delete_nonexistent_defined_name_returns_error() {
        let mut wb = Workbook::new();
        let result = wb.delete_defined_name("Ghost", None);
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("Ghost"));
    }

    #[test]
    fn test_defined_name_sheet_scope_requires_existing_sheet() {
        let mut wb = Workbook::new();
        let result = wb.set_defined_name("TestName", "Sheet1!$A$1", Some("NonExistent"), None);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_defined_name_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("defined_names.xlsx");

        let mut wb = Workbook::new();
        wb.set_defined_name("Revenue", "Sheet1!$E$1:$E$100", None, Some("Total revenue"))
            .unwrap();
        wb.set_defined_name("LocalName", "Sheet1!$A$1", Some("Sheet1"), None)
            .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let all = wb2.get_all_defined_names();
        assert_eq!(all.len(), 2);
        assert_eq!(all[0].name, "Revenue");
        assert_eq!(all[0].value, "Sheet1!$E$1:$E$100");
        assert_eq!(all[0].comment, Some("Total revenue".to_string()));
        assert_eq!(all[1].name, "LocalName");
        assert_eq!(all[1].value, "Sheet1!$A$1");
        assert_eq!(
            all[1].scope,
            crate::defined_names::DefinedNameScope::Sheet(0)
        );
    }

    #[test]
    fn test_protect_sheet_via_workbook() {
        let mut wb = Workbook::new();
        let config = crate::sheet::SheetProtectionConfig::default();
        wb.protect_sheet("Sheet1", &config).unwrap();

        assert!(wb.is_sheet_protected("Sheet1").unwrap());
    }

    #[test]
    fn test_unprotect_sheet_via_workbook() {
        let mut wb = Workbook::new();
        let config = crate::sheet::SheetProtectionConfig::default();
        wb.protect_sheet("Sheet1", &config).unwrap();
        assert!(wb.is_sheet_protected("Sheet1").unwrap());

        wb.unprotect_sheet("Sheet1").unwrap();
        assert!(!wb.is_sheet_protected("Sheet1").unwrap());
    }

    #[test]
    fn test_protect_sheet_nonexistent_returns_error() {
        let mut wb = Workbook::new();
        let config = crate::sheet::SheetProtectionConfig::default();
        let result = wb.protect_sheet("NoSuchSheet", &config);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_is_sheet_protected_nonexistent_returns_error() {
        let wb = Workbook::new();
        let result = wb.is_sheet_protected("NoSuchSheet");
        assert!(result.is_err());
    }

    #[test]
    fn test_protect_sheet_with_password_and_permissions() {
        let mut wb = Workbook::new();
        let config = crate::sheet::SheetProtectionConfig {
            password: Some("secret".to_string()),
            format_cells: true,
            insert_rows: true,
            sort: true,
            ..crate::sheet::SheetProtectionConfig::default()
        };
        wb.protect_sheet("Sheet1", &config).unwrap();
        assert!(wb.is_sheet_protected("Sheet1").unwrap());
    }

    #[test]
    fn test_sheet_protection_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("sheet_protection.xlsx");

        let mut wb = Workbook::new();
        let config = crate::sheet::SheetProtectionConfig {
            password: Some("pass".to_string()),
            format_cells: true,
            ..crate::sheet::SheetProtectionConfig::default()
        };
        wb.protect_sheet("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert!(wb2.is_sheet_protected("Sheet1").unwrap());
    }

    #[test]
    fn test_add_table() {
        use crate::table::{TableColumn, TableConfig};

        let mut wb = Workbook::new();
        let config = TableConfig {
            name: "Sales".to_string(),
            display_name: "Sales".to_string(),
            range: "A1:C5".to_string(),
            columns: vec![
                TableColumn {
                    name: "Product".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Quantity".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Price".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
            ],
            show_header_row: true,
            style_name: Some("TableStyleMedium2".to_string()),
            auto_filter: true,
            ..TableConfig::default()
        };
        wb.add_table("Sheet1", &config).unwrap();

        let tables = wb.get_tables("Sheet1").unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "Sales");
        assert_eq!(tables[0].display_name, "Sales");
        assert_eq!(tables[0].range, "A1:C5");
        assert_eq!(tables[0].columns, vec!["Product", "Quantity", "Price"]);
        assert!(tables[0].auto_filter);
        assert!(tables[0].show_header_row);
        assert_eq!(tables[0].style_name, Some("TableStyleMedium2".to_string()));
    }

    #[test]
    fn test_add_table_duplicate_name_error() {
        use crate::table::{TableColumn, TableConfig};

        let mut wb = Workbook::new();
        let config = TableConfig {
            name: "T1".to_string(),
            display_name: "T1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        wb.add_table("Sheet1", &config).unwrap();
        let result = wb.add_table("Sheet1", &config);
        assert!(matches!(
            result.unwrap_err(),
            Error::TableAlreadyExists { .. }
        ));
    }

    #[test]
    fn test_add_table_invalid_config() {
        use crate::table::TableConfig;

        let mut wb = Workbook::new();
        let config = TableConfig {
            name: String::new(),
            range: "A1:B5".to_string(),
            ..TableConfig::default()
        };
        assert!(wb.add_table("Sheet1", &config).is_err());
    }

    #[test]
    fn test_add_table_sheet_not_found() {
        use crate::table::{TableColumn, TableConfig};

        let mut wb = Workbook::new();
        let config = TableConfig {
            name: "T1".to_string(),
            display_name: "T1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        let result = wb.add_table("NoSheet", &config);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_get_tables_empty() {
        let wb = Workbook::new();
        let tables = wb.get_tables("Sheet1").unwrap();
        assert!(tables.is_empty());
    }

    #[test]
    fn test_get_tables_sheet_not_found() {
        let wb = Workbook::new();
        let result = wb.get_tables("NoSheet");
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_delete_table() {
        use crate::table::{TableColumn, TableConfig};

        let mut wb = Workbook::new();
        let config = TableConfig {
            name: "T1".to_string(),
            display_name: "T1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        wb.add_table("Sheet1", &config).unwrap();
        assert_eq!(wb.get_tables("Sheet1").unwrap().len(), 1);

        wb.delete_table("Sheet1", "T1").unwrap();
        assert!(wb.get_tables("Sheet1").unwrap().is_empty());
    }

    #[test]
    fn test_delete_table_not_found() {
        let mut wb = Workbook::new();
        let result = wb.delete_table("Sheet1", "NoTable");
        assert!(matches!(result.unwrap_err(), Error::TableNotFound { .. }));
    }

    #[test]
    fn test_delete_table_wrong_sheet() {
        use crate::table::{TableColumn, TableConfig};

        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        let config = TableConfig {
            name: "T1".to_string(),
            display_name: "T1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        wb.add_table("Sheet1", &config).unwrap();

        let result = wb.delete_table("Sheet2", "T1");
        assert!(matches!(result.unwrap_err(), Error::TableNotFound { .. }));
        // Table should still exist on Sheet1.
        assert_eq!(wb.get_tables("Sheet1").unwrap().len(), 1);
    }

    #[test]
    fn test_multiple_tables_on_sheet() {
        use crate::table::{TableColumn, TableConfig};

        let mut wb = Workbook::new();
        let config1 = TableConfig {
            name: "T1".to_string(),
            display_name: "T1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![
                TableColumn {
                    name: "Name".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Score".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
            ],
            ..TableConfig::default()
        };
        let config2 = TableConfig {
            name: "T2".to_string(),
            display_name: "T2".to_string(),
            range: "D1:E5".to_string(),
            columns: vec![
                TableColumn {
                    name: "City".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Population".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
            ],
            ..TableConfig::default()
        };
        wb.add_table("Sheet1", &config1).unwrap();
        wb.add_table("Sheet1", &config2).unwrap();

        let tables = wb.get_tables("Sheet1").unwrap();
        assert_eq!(tables.len(), 2);
        assert_eq!(tables[0].name, "T1");
        assert_eq!(tables[1].name, "T2");
    }

    #[test]
    fn test_tables_on_different_sheets() {
        use crate::table::{TableColumn, TableConfig};

        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        let config1 = TableConfig {
            name: "T1".to_string(),
            display_name: "T1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col1".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        let config2 = TableConfig {
            name: "T2".to_string(),
            display_name: "T2".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col2".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        wb.add_table("Sheet1", &config1).unwrap();
        wb.add_table("Sheet2", &config2).unwrap();

        assert_eq!(wb.get_tables("Sheet1").unwrap().len(), 1);
        assert_eq!(wb.get_tables("Sheet2").unwrap().len(), 1);
        assert_eq!(wb.get_tables("Sheet1").unwrap()[0].name, "T1");
        assert_eq!(wb.get_tables("Sheet2").unwrap()[0].name, "T2");
    }

    #[test]
    fn test_table_save_produces_zip_parts() {
        use crate::table::{TableColumn, TableConfig};

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("table_parts.xlsx");

        let mut wb = Workbook::new();
        let config = TableConfig {
            name: "Sales".to_string(),
            display_name: "Sales".to_string(),
            range: "A1:C5".to_string(),
            columns: vec![
                TableColumn {
                    name: "Product".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Qty".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Price".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
            ],
            style_name: Some("TableStyleMedium2".to_string()),
            ..TableConfig::default()
        };
        wb.add_table("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        assert!(
            archive.by_name("xl/tables/table1.xml").is_ok(),
            "table1.xml should be present in the ZIP"
        );
        assert!(
            archive
                .by_name("xl/worksheets/_rels/sheet1.xml.rels")
                .is_ok(),
            "worksheet rels should be present"
        );

        // Verify table XML content.
        let mut table_data = Vec::new();
        archive
            .by_name("xl/tables/table1.xml")
            .unwrap()
            .read_to_end(&mut table_data)
            .unwrap();
        let table_str = String::from_utf8(table_data).unwrap();
        assert!(table_str.contains("Sales"));
        assert!(table_str.contains("A1:C5"));
        assert!(table_str.contains("TableStyleMedium2"));
        assert!(table_str.contains("autoFilter"));
        assert!(table_str.contains("tableColumn"));

        // Verify worksheet XML has tableParts element.
        let mut ws_data = Vec::new();
        archive
            .by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_end(&mut ws_data)
            .unwrap();
        let ws_str = String::from_utf8(ws_data).unwrap();
        assert!(
            ws_str.contains("tableParts"),
            "worksheet should contain tableParts element"
        );
        assert!(
            ws_str.contains("tablePart"),
            "worksheet should contain tablePart reference"
        );

        // Verify content types include the table.
        let mut ct_data = Vec::new();
        archive
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_end(&mut ct_data)
            .unwrap();
        let ct_str = String::from_utf8(ct_data).unwrap();
        assert!(
            ct_str.contains("table+xml"),
            "content types should reference the table"
        );

        // Verify worksheet rels include a table relationship.
        let mut rels_data = Vec::new();
        archive
            .by_name("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap()
            .read_to_end(&mut rels_data)
            .unwrap();
        let rels_str = String::from_utf8(rels_data).unwrap();
        assert!(
            rels_str.contains("relationships/table"),
            "worksheet rels should reference the table"
        );
    }

    #[test]
    fn test_table_roundtrip_save_open() {
        use crate::table::{TableColumn, TableConfig};

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("table_roundtrip.xlsx");

        let mut wb = Workbook::new();
        let config = TableConfig {
            name: "Inventory".to_string(),
            display_name: "Inventory".to_string(),
            range: "A1:D10".to_string(),
            columns: vec![
                TableColumn {
                    name: "Item".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Stock".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Price".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Supplier".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
            ],
            show_header_row: true,
            style_name: Some("TableStyleLight1".to_string()),
            auto_filter: true,
            ..TableConfig::default()
        };
        wb.add_table("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let tables = wb2.get_tables("Sheet1").unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "Inventory");
        assert_eq!(tables[0].display_name, "Inventory");
        assert_eq!(tables[0].range, "A1:D10");
        assert_eq!(
            tables[0].columns,
            vec!["Item", "Stock", "Price", "Supplier"]
        );
        assert!(tables[0].auto_filter);
        assert!(tables[0].show_header_row);
        assert_eq!(tables[0].style_name, Some("TableStyleLight1".to_string()));
    }

    #[test]
    fn test_table_roundtrip_multiple_tables() {
        use crate::table::{TableColumn, TableConfig};

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("multi_table_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.add_table(
            "Sheet1",
            &TableConfig {
                name: "T1".to_string(),
                display_name: "T1".to_string(),
                range: "A1:B5".to_string(),
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
        wb.add_table(
            "Sheet2",
            &TableConfig {
                name: "T2".to_string(),
                display_name: "T2".to_string(),
                range: "C1:D8".to_string(),
                columns: vec![
                    TableColumn {
                        name: "Category".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumn {
                        name: "Count".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                ],
                auto_filter: false,
                ..TableConfig::default()
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let t1 = wb2.get_tables("Sheet1").unwrap();
        assert_eq!(t1.len(), 1);
        assert_eq!(t1[0].name, "T1");
        assert_eq!(t1[0].range, "A1:B5");

        let t2 = wb2.get_tables("Sheet2").unwrap();
        assert_eq!(t2.len(), 1);
        assert_eq!(t2[0].name, "T2");
        assert_eq!(t2[0].range, "C1:D8");
        assert!(!t2[0].auto_filter);
    }

    #[test]
    fn test_table_roundtrip_resave() {
        use crate::table::{TableColumn, TableConfig};

        let dir = TempDir::new().unwrap();
        let path1 = dir.path().join("table_resave1.xlsx");
        let path2 = dir.path().join("table_resave2.xlsx");

        let mut wb = Workbook::new();
        wb.add_table(
            "Sheet1",
            &TableConfig {
                name: "T1".to_string(),
                display_name: "T1".to_string(),
                range: "A1:B3".to_string(),
                columns: vec![
                    TableColumn {
                        name: "X".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumn {
                        name: "Y".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                ],
                ..TableConfig::default()
            },
        )
        .unwrap();
        wb.save(&path1).unwrap();

        let wb2 = Workbook::open(&path1).unwrap();
        wb2.save(&path2).unwrap();

        let wb3 = Workbook::open(&path2).unwrap();
        let tables = wb3.get_tables("Sheet1").unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "T1");
        assert_eq!(tables[0].columns, vec!["X", "Y"]);
    }

    #[test]
    fn test_auto_filter_not_regressed_by_tables() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("autofilter_with_table.xlsx");

        let mut wb = Workbook::new();
        wb.set_auto_filter("Sheet1", "A1:C50").unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let ws = wb2.worksheet_ref("Sheet1").unwrap();
        assert!(ws.auto_filter.is_some());
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:C50");
    }

    #[test]
    fn test_delete_sheet_removes_tables() {
        use crate::table::{TableColumn, TableConfig};

        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.add_table(
            "Sheet1",
            &TableConfig {
                name: "T1".to_string(),
                display_name: "T1".to_string(),
                range: "A1:B5".to_string(),
                columns: vec![TableColumn {
                    name: "Col".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                }],
                ..TableConfig::default()
            },
        )
        .unwrap();
        wb.add_table(
            "Sheet2",
            &TableConfig {
                name: "T2".to_string(),
                display_name: "T2".to_string(),
                range: "A1:B5".to_string(),
                columns: vec![TableColumn {
                    name: "Col".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                }],
                ..TableConfig::default()
            },
        )
        .unwrap();

        wb.delete_sheet("Sheet1").unwrap();
        // T1 should be gone, T2 should still exist on Sheet2.
        let tables = wb.get_tables("Sheet2").unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "T2");
    }

    #[test]
    fn test_table_with_no_auto_filter() {
        use crate::table::{TableColumn, TableConfig};

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("table_no_filter.xlsx");

        let mut wb = Workbook::new();
        wb.add_table(
            "Sheet1",
            &TableConfig {
                name: "Plain".to_string(),
                display_name: "Plain".to_string(),
                range: "A1:B3".to_string(),
                columns: vec![
                    TableColumn {
                        name: "A".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                    TableColumn {
                        name: "B".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                    },
                ],
                auto_filter: false,
                ..TableConfig::default()
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let tables = wb2.get_tables("Sheet1").unwrap();
        assert_eq!(tables.len(), 1);
        assert!(!tables[0].auto_filter);
    }
}
