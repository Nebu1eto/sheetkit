use super::*;
use crate::workbook::open_options::OpenOptions;

/// VBA project relationship type URI.
const VBA_PROJECT_REL_TYPE: &str =
    "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

/// VBA project content type.
const VBA_PROJECT_CONTENT_TYPE: &str = "application/vnd.ms-office.vbaProject";

impl Workbook {
    /// Create a new empty workbook containing a single empty sheet named "Sheet1".
    pub fn new() -> Self {
        let sst_runtime = SharedStringTable::new();
        let mut sheet_name_index = HashMap::new();
        sheet_name_index.insert("Sheet1".to_string(), 0);
        Self {
            format: WorkbookFormat::default(),
            content_types: ContentTypes::default(),
            package_rels: relationships::package_rels(),
            workbook_xml: WorkbookXml::default(),
            workbook_rels: relationships::workbook_rels(),
            worksheets: vec![("Sheet1".to_string(), WorksheetXml::default())],
            stylesheet: StyleSheet::default(),
            sst_runtime,
            sheet_comments: vec![None],
            charts: vec![],
            raw_charts: vec![],
            drawings: vec![],
            images: vec![],
            worksheet_drawings: HashMap::new(),
            worksheet_rels: HashMap::new(),
            drawing_rels: HashMap::new(),
            core_properties: None,
            app_properties: None,
            custom_properties: None,
            pivot_tables: vec![],
            pivot_cache_defs: vec![],
            pivot_cache_records: vec![],
            theme_xml: None,
            theme_colors: crate::theme::default_theme_colors(),
            sheet_name_index,
            sheet_sparklines: vec![vec![]],
            sheet_vml: vec![None],
            unknown_parts: vec![],
            deferred_parts: HashMap::new(),
            vba_blob: None,
            tables: vec![],
            raw_sheet_xml: vec![None],
            slicer_defs: vec![],
            slicer_caches: vec![],
            sheet_threaded_comments: vec![None],
            person_list: sheetkit_xml::threaded_comment::PersonList::default(),
            sheet_form_controls: vec![vec![]],
            streamed_sheets: HashMap::new(),
        }
    }

    /// Open an existing `.xlsx` file from disk.
    ///
    /// If the file is encrypted (CFB container), returns
    /// [`Error::FileEncrypted`]. Use [`Workbook::open_with_password`] instead.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_options(path, &OpenOptions::default())
    }

    /// Open an existing `.xlsx` file with custom parsing options.
    ///
    /// See [`OpenOptions`] for available options including row limits,
    /// sheet filtering, and ZIP safety limits.
    pub fn open_with_options<P: AsRef<Path>>(path: P, options: &OpenOptions) -> Result<Self> {
        let data = std::fs::read(path.as_ref())?;

        // Detect encrypted files (CFB container)
        #[cfg(feature = "encryption")]
        if data.len() >= 8 {
            if let Ok(crate::crypt::ContainerFormat::Cfb) =
                crate::crypt::detect_container_format(&data)
            {
                return Err(Error::FileEncrypted);
            }
        }

        let cursor = std::io::Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor).map_err(|e| Error::Zip(e.to_string()))?;
        Self::from_archive(&mut archive, options)
    }

    /// Build a Workbook from an already-opened ZIP archive.
    fn from_archive<R: std::io::Read + std::io::Seek>(
        archive: &mut zip::ZipArchive<R>,
        options: &OpenOptions,
    ) -> Result<Self> {
        // ZIP safety checks: entry count and total decompressed size.
        if let Some(max_entries) = options.max_zip_entries {
            let count = archive.len();
            if count > max_entries {
                return Err(Error::ZipEntryCountExceeded {
                    count,
                    limit: max_entries,
                });
            }
        }
        if let Some(max_size) = options.max_unzip_size {
            let mut total_size: u64 = 0;
            for i in 0..archive.len() {
                let entry = archive.by_index(i).map_err(|e| Error::Zip(e.to_string()))?;
                total_size = total_size.saturating_add(entry.size());
                if total_size > max_size {
                    return Err(Error::ZipSizeExceeded {
                        size: total_size,
                        limit: max_size,
                    });
                }
            }
        }

        // Track all ZIP entry paths that are explicitly handled so that the
        // remaining entries can be preserved as unknown parts.
        let mut known_paths: HashSet<String> = HashSet::new();

        // Parse [Content_Types].xml
        let content_types: ContentTypes = read_xml_part(archive, "[Content_Types].xml")?;
        known_paths.insert("[Content_Types].xml".to_string());

        // Infer the workbook format from the content type of xl/workbook.xml.
        let format = content_types
            .overrides
            .iter()
            .find(|o| o.part_name == "/xl/workbook.xml")
            .and_then(|o| WorkbookFormat::from_content_type(&o.content_type))
            .unwrap_or_default();

        // Parse _rels/.rels
        let package_rels: Relationships = read_xml_part(archive, "_rels/.rels")?;
        known_paths.insert("_rels/.rels".to_string());

        // Parse xl/workbook.xml
        let workbook_xml: WorkbookXml = read_xml_part(archive, "xl/workbook.xml")?;
        known_paths.insert("xl/workbook.xml".to_string());

        // Parse xl/_rels/workbook.xml.rels
        let workbook_rels: Relationships = read_xml_part(archive, "xl/_rels/workbook.xml.rels")?;
        known_paths.insert("xl/_rels/workbook.xml.rels".to_string());

        // Parse each worksheet referenced in the workbook.
        let sheet_count = workbook_xml.sheets.sheets.len();
        let mut worksheets = Vec::with_capacity(sheet_count);
        let mut worksheet_paths = Vec::with_capacity(sheet_count);
        let mut raw_sheet_xml: Vec<Option<Vec<u8>>> = Vec::with_capacity(sheet_count);
        for sheet_entry in &workbook_xml.sheets.sheets {
            // Find the relationship target for this sheet's rId.
            let rel = workbook_rels
                .relationships
                .iter()
                .find(|r| r.id == sheet_entry.r_id && r.rel_type == rel_types::WORKSHEET);

            let rel = rel.ok_or_else(|| {
                Error::Internal(format!(
                    "missing worksheet relationship for sheet '{}'",
                    sheet_entry.name
                ))
            })?;

            let sheet_path = resolve_relationship_target("xl/workbook.xml", &rel.target);

            if options.should_parse_sheet(&sheet_entry.name) {
                let mut ws: WorksheetXml = read_xml_part(archive, &sheet_path)?;
                for row in &mut ws.sheet_data.rows {
                    row.cells.shrink_to_fit();
                }
                ws.sheet_data.rows.shrink_to_fit();
                worksheets.push((sheet_entry.name.clone(), ws));
                raw_sheet_xml.push(None);
            } else {
                // Store the raw XML bytes so the sheet round-trips unchanged on save.
                let raw_bytes = read_bytes_part(archive, &sheet_path)?;
                worksheets.push((sheet_entry.name.clone(), WorksheetXml::default()));
                raw_sheet_xml.push(Some(raw_bytes));
            };
            known_paths.insert(sheet_path.clone());
            worksheet_paths.push(sheet_path);
        }

        // Parse xl/styles.xml
        let stylesheet: StyleSheet = read_xml_part(archive, "xl/styles.xml")?;
        known_paths.insert("xl/styles.xml".to_string());

        // Parse xl/sharedStrings.xml (optional -- may not exist for workbooks with no strings)
        let shared_strings: Sst =
            read_xml_part(archive, "xl/sharedStrings.xml").unwrap_or_default();
        known_paths.insert("xl/sharedStrings.xml".to_string());

        let sst_runtime = SharedStringTable::from_sst(shared_strings);

        // Parse xl/theme/theme1.xml (optional -- preserved as raw bytes for round-trip).
        let (theme_xml, theme_colors) = match read_bytes_part(archive, "xl/theme/theme1.xml") {
            Ok(bytes) => {
                let colors = sheetkit_xml::theme::parse_theme_colors(&bytes);
                (Some(bytes), colors)
            }
            Err(_) => (None, crate::theme::default_theme_colors()),
        };
        known_paths.insert("xl/theme/theme1.xml".to_string());

        // Parse per-sheet worksheet relationship files (optional).
        // Always loaded: needed for hyperlinks, on-demand comment loading, etc.
        let mut worksheet_rels: HashMap<usize, Relationships> = HashMap::with_capacity(sheet_count);
        for (i, sheet_path) in worksheet_paths.iter().enumerate() {
            let rels_path = relationship_part_path(sheet_path);
            if let Ok(rels) = read_xml_part::<Relationships, _>(archive, &rels_path) {
                worksheet_rels.insert(i, rels);
                known_paths.insert(rels_path);
            }
        }

        let read_fast = options.is_read_fast();

        // Auxiliary part parsing: skipped in ReadFast mode.
        let mut sheet_comments: Vec<Option<Comments>> = vec![None; worksheets.len()];
        let mut sheet_vml: Vec<Option<Vec<u8>>> = vec![None; worksheets.len()];
        let mut drawings: Vec<(String, WsDr)> = Vec::new();
        let mut worksheet_drawings: HashMap<usize, usize> = HashMap::new();
        let mut drawing_rels: HashMap<usize, Relationships> = HashMap::new();
        let mut charts: Vec<(String, ChartSpace)> = Vec::new();
        let mut raw_charts: Vec<(String, Vec<u8>)> = Vec::new();
        let mut images: Vec<(String, Vec<u8>)> = Vec::new();
        let mut core_properties: Option<sheetkit_xml::doc_props::CoreProperties> = None;
        let mut app_properties: Option<sheetkit_xml::doc_props::ExtendedProperties> = None;
        let mut custom_properties: Option<sheetkit_xml::doc_props::CustomProperties> = None;
        let mut pivot_cache_defs = Vec::new();
        let mut pivot_tables = Vec::new();
        let mut pivot_cache_records = Vec::new();
        let mut slicer_defs = Vec::new();
        let mut slicer_caches = Vec::new();
        let mut sheet_threaded_comments: Vec<
            Option<sheetkit_xml::threaded_comment::ThreadedComments>,
        > = vec![None; worksheets.len()];
        let mut person_list = sheetkit_xml::threaded_comment::PersonList::default();
        let mut sheet_sparklines: Vec<Vec<crate::sparkline::SparklineConfig>> =
            vec![vec![]; worksheets.len()];
        let mut vba_blob: Option<Vec<u8>> = None;
        let mut tables: Vec<(String, sheetkit_xml::table::TableXml, usize)> = Vec::new();

        if !read_fast {
            let mut drawing_path_to_idx: HashMap<String, usize> = HashMap::new();

            for (sheet_idx, sheet_path) in worksheet_paths.iter().enumerate() {
                let Some(rels) = worksheet_rels.get(&sheet_idx) else {
                    continue;
                };

                if let Some(comment_rel) = rels
                    .relationships
                    .iter()
                    .find(|r| r.rel_type == rel_types::COMMENTS)
                {
                    let comment_path = resolve_relationship_target(sheet_path, &comment_rel.target);
                    if let Ok(comments) = read_xml_part::<Comments, _>(archive, &comment_path) {
                        sheet_comments[sheet_idx] = Some(comments);
                        known_paths.insert(comment_path);
                    }
                }

                if let Some(vml_rel) = rels
                    .relationships
                    .iter()
                    .find(|r| r.rel_type == rel_types::VML_DRAWING)
                {
                    let vml_path = resolve_relationship_target(sheet_path, &vml_rel.target);
                    if let Ok(bytes) = read_bytes_part(archive, &vml_path) {
                        sheet_vml[sheet_idx] = Some(bytes);
                        known_paths.insert(vml_path);
                    }
                }

                if let Some(drawing_rel) = rels
                    .relationships
                    .iter()
                    .find(|r| r.rel_type == rel_types::DRAWING)
                {
                    let drawing_path = resolve_relationship_target(sheet_path, &drawing_rel.target);
                    let drawing_idx = if let Some(idx) = drawing_path_to_idx.get(&drawing_path) {
                        *idx
                    } else if let Ok(drawing) = read_xml_part::<WsDr, _>(archive, &drawing_path) {
                        let idx = drawings.len();
                        drawings.push((drawing_path.clone(), drawing));
                        drawing_path_to_idx.insert(drawing_path.clone(), idx);
                        known_paths.insert(drawing_path);
                        idx
                    } else {
                        continue;
                    };
                    worksheet_drawings.insert(sheet_idx, drawing_idx);
                }
            }

            // Fallback: load drawing parts listed in content types even when they
            // are not discoverable via worksheet rel parsing.
            for ovr in &content_types.overrides {
                if ovr.content_type != mime_types::DRAWING {
                    continue;
                }
                let drawing_path = ovr.part_name.trim_start_matches('/').to_string();
                if drawing_path_to_idx.contains_key(&drawing_path) {
                    continue;
                }
                if let Ok(drawing) = read_xml_part::<WsDr, _>(archive, &drawing_path) {
                    let idx = drawings.len();
                    drawings.push((drawing_path.clone(), drawing));
                    known_paths.insert(drawing_path.clone());
                    drawing_path_to_idx.insert(drawing_path, idx);
                }
            }

            let mut seen_chart_paths: HashSet<String> = HashSet::new();
            let mut seen_image_paths: HashSet<String> = HashSet::new();

            for (drawing_idx, (drawing_path, _)) in drawings.iter().enumerate() {
                let drawing_rels_path = relationship_part_path(drawing_path);
                let Ok(rels) = read_xml_part::<Relationships, _>(archive, &drawing_rels_path)
                else {
                    continue;
                };
                known_paths.insert(drawing_rels_path);

                for rel in &rels.relationships {
                    if rel.rel_type == rel_types::CHART {
                        let chart_path = resolve_relationship_target(drawing_path, &rel.target);
                        if seen_chart_paths.insert(chart_path.clone()) {
                            match read_xml_part::<ChartSpace, _>(archive, &chart_path) {
                                Ok(chart) => {
                                    known_paths.insert(chart_path.clone());
                                    charts.push((chart_path, chart));
                                }
                                Err(_) => {
                                    if let Ok(bytes) = read_bytes_part(archive, &chart_path) {
                                        known_paths.insert(chart_path.clone());
                                        raw_charts.push((chart_path, bytes));
                                    }
                                }
                            }
                        }
                    } else if rel.rel_type == rel_types::IMAGE {
                        let image_path = resolve_relationship_target(drawing_path, &rel.target);
                        if seen_image_paths.insert(image_path.clone()) {
                            if let Ok(bytes) = read_bytes_part(archive, &image_path) {
                                known_paths.insert(image_path.clone());
                                images.push((image_path, bytes));
                            }
                        }
                    }
                }

                drawing_rels.insert(drawing_idx, rels);
            }

            // Fallback: load chart parts listed in content types even when no
            // drawing relationship was read.
            for ovr in &content_types.overrides {
                if ovr.content_type != mime_types::CHART {
                    continue;
                }
                let chart_path = ovr.part_name.trim_start_matches('/').to_string();
                if seen_chart_paths.insert(chart_path.clone()) {
                    match read_xml_part::<ChartSpace, _>(archive, &chart_path) {
                        Ok(chart) => {
                            known_paths.insert(chart_path.clone());
                            charts.push((chart_path, chart));
                        }
                        Err(_) => {
                            if let Ok(bytes) = read_bytes_part(archive, &chart_path) {
                                known_paths.insert(chart_path.clone());
                                raw_charts.push((chart_path, bytes));
                            }
                        }
                    }
                }
            }

            // Parse docProps/core.xml (optional - uses manual XML parsing)
            core_properties = read_string_part(archive, "docProps/core.xml")
                .ok()
                .and_then(|xml_str| {
                    sheetkit_xml::doc_props::deserialize_core_properties(&xml_str).ok()
                });
            known_paths.insert("docProps/core.xml".to_string());

            // Parse docProps/app.xml (optional - uses serde)
            app_properties = read_xml_part(archive, "docProps/app.xml").ok();
            known_paths.insert("docProps/app.xml".to_string());

            // Parse docProps/custom.xml (optional - uses manual XML parsing)
            custom_properties = read_string_part(archive, "docProps/custom.xml")
                .ok()
                .and_then(|xml_str| {
                    sheetkit_xml::doc_props::deserialize_custom_properties(&xml_str).ok()
                });
            known_paths.insert("docProps/custom.xml".to_string());

            // Parse pivot cache definitions, pivot tables, and pivot cache records.
            for ovr in &content_types.overrides {
                let path = ovr.part_name.trim_start_matches('/');
                if ovr.content_type == mime_types::PIVOT_CACHE_DEFINITION {
                    if let Ok(pcd) = read_xml_part::<
                        sheetkit_xml::pivot_cache::PivotCacheDefinition,
                        _,
                    >(archive, path)
                    {
                        known_paths.insert(path.to_string());
                        pivot_cache_defs.push((path.to_string(), pcd));
                    }
                } else if ovr.content_type == mime_types::PIVOT_TABLE {
                    if let Ok(pt) = read_xml_part::<
                        sheetkit_xml::pivot_table::PivotTableDefinition,
                        _,
                    >(archive, path)
                    {
                        known_paths.insert(path.to_string());
                        pivot_tables.push((path.to_string(), pt));
                    }
                } else if ovr.content_type == mime_types::PIVOT_CACHE_RECORDS {
                    if let Ok(pcr) = read_xml_part::<sheetkit_xml::pivot_cache::PivotCacheRecords, _>(
                        archive, path,
                    ) {
                        known_paths.insert(path.to_string());
                        pivot_cache_records.push((path.to_string(), pcr));
                    }
                }
            }

            // Parse slicer definitions and slicer cache definitions.
            for ovr in &content_types.overrides {
                let path = ovr.part_name.trim_start_matches('/');
                if ovr.content_type == mime_types::SLICER {
                    if let Ok(sd) =
                        read_xml_part::<sheetkit_xml::slicer::SlicerDefinitions, _>(archive, path)
                    {
                        slicer_defs.push((path.to_string(), sd));
                    }
                } else if ovr.content_type == mime_types::SLICER_CACHE {
                    if let Ok(raw) = read_string_part(archive, path) {
                        if let Some(scd) = sheetkit_xml::slicer::parse_slicer_cache(&raw) {
                            slicer_caches.push((path.to_string(), scd));
                        }
                    }
                }
            }

            // Parse threaded comments per-sheet and the workbook-level person list.
            for (sheet_idx, sheet_path) in worksheet_paths.iter().enumerate() {
                let Some(rels) = worksheet_rels.get(&sheet_idx) else {
                    continue;
                };
                if let Some(tc_rel) = rels.relationships.iter().find(|r| {
                    r.rel_type == sheetkit_xml::threaded_comment::REL_TYPE_THREADED_COMMENT
                }) {
                    let tc_path = resolve_relationship_target(sheet_path, &tc_rel.target);
                    if let Ok(tc) = read_xml_part::<
                        sheetkit_xml::threaded_comment::ThreadedComments,
                        _,
                    >(archive, &tc_path)
                    {
                        sheet_threaded_comments[sheet_idx] = Some(tc);
                        known_paths.insert(tc_path);
                    }
                }
            }

            // Parse person list (workbook-level).
            person_list = {
                let mut found = None;
                if let Some(person_rel) = workbook_rels
                    .relationships
                    .iter()
                    .find(|r| r.rel_type == sheetkit_xml::threaded_comment::REL_TYPE_PERSON)
                {
                    let person_path =
                        resolve_relationship_target("xl/workbook.xml", &person_rel.target);
                    if let Ok(pl) = read_xml_part::<sheetkit_xml::threaded_comment::PersonList, _>(
                        archive,
                        &person_path,
                    ) {
                        known_paths.insert(person_path);
                        found = Some(pl);
                    }
                }
                if found.is_none() {
                    if let Ok(pl) = read_xml_part::<sheetkit_xml::threaded_comment::PersonList, _>(
                        archive,
                        "xl/persons/person.xml",
                    ) {
                        known_paths.insert("xl/persons/person.xml".to_string());
                        found = Some(pl);
                    }
                }
                found.unwrap_or_default()
            };

            // Parse sparklines from worksheet extension lists.
            for (i, ws_path) in worksheet_paths.iter().enumerate() {
                if let Ok(raw) = read_string_part(archive, ws_path) {
                    let parsed = parse_sparklines_from_xml(&raw);
                    if !parsed.is_empty() {
                        sheet_sparklines[i] = parsed;
                    }
                }
            }

            // Load VBA project binary blob if present (macro-enabled files).
            vba_blob = read_bytes_part(archive, "xl/vbaProject.bin").ok();
            if vba_blob.is_some() {
                known_paths.insert("xl/vbaProject.bin".to_string());
            }

            // Parse table parts referenced from worksheet relationships.
            for (sheet_idx, sheet_path) in worksheet_paths.iter().enumerate() {
                let Some(rels) = worksheet_rels.get(&sheet_idx) else {
                    continue;
                };
                for rel in &rels.relationships {
                    if rel.rel_type != rel_types::TABLE {
                        continue;
                    }
                    let table_path = resolve_relationship_target(sheet_path, &rel.target);
                    if let Ok(table_xml) =
                        read_xml_part::<sheetkit_xml::table::TableXml, _>(archive, &table_path)
                    {
                        known_paths.insert(table_path.clone());
                        tables.push((table_path, table_xml, sheet_idx));
                    }
                }
            }
            // Fallback: load table parts from content type overrides if not found via rels.
            for ovr in &content_types.overrides {
                if ovr.content_type != mime_types::TABLE {
                    continue;
                }
                let table_path = ovr.part_name.trim_start_matches('/').to_string();
                if tables.iter().any(|(p, _, _)| p == &table_path) {
                    continue;
                }
                if let Ok(table_xml) =
                    read_xml_part::<sheetkit_xml::table::TableXml, _>(archive, &table_path)
                {
                    known_paths.insert(table_path.clone());
                    tables.push((table_path, table_xml, 0));
                }
            }
        }

        let sheet_form_controls: Vec<Vec<crate::control::FormControlConfig>> =
            vec![vec![]; worksheets.len()];

        // Build sheet name -> index lookup.
        let mut sheet_name_index = HashMap::with_capacity(worksheets.len());
        for (i, (name, _)) in worksheets.iter().enumerate() {
            sheet_name_index.insert(name.clone(), i);
        }

        // Collect remaining ZIP entries. In ReadFast mode, unhandled entries
        // go into deferred_parts; in Full mode, they go into unknown_parts.
        let mut unknown_parts: Vec<(String, Vec<u8>)> = Vec::new();
        let mut deferred_parts: HashMap<String, Vec<u8>> = HashMap::new();
        for i in 0..archive.len() {
            let Ok(entry) = archive.by_index(i) else {
                continue;
            };
            let name = entry.name().to_string();
            drop(entry);
            if !known_paths.contains(&name) {
                if let Ok(bytes) = read_bytes_part(archive, &name) {
                    if read_fast {
                        deferred_parts.insert(name, bytes);
                    } else {
                        unknown_parts.push((name, bytes));
                    }
                }
            }
        }

        // Populate cached column numbers on all cells, apply row limit, and
        // ensure sorted order for binary search correctness.
        for (_name, ws) in &mut worksheets {
            // Ensure rows are sorted by row number (some writers output unsorted data).
            ws.sheet_data.rows.sort_unstable_by_key(|r| r.r);

            // Apply sheet_rows limit: keep only the first N rows.
            if let Some(max_rows) = options.sheet_rows {
                ws.sheet_data.rows.truncate(max_rows as usize);
            }

            for row in &mut ws.sheet_data.rows {
                for cell in &mut row.cells {
                    cell.col = fast_col_number(cell.r.as_str());
                }
                // Ensure cells within a row are sorted by column number.
                row.cells.sort_unstable_by_key(|c| c.col);
            }
        }

        Ok(Self {
            format,
            content_types,
            package_rels,
            workbook_xml,
            workbook_rels,
            worksheets,
            stylesheet,
            sst_runtime,
            sheet_comments,
            charts,
            raw_charts,
            drawings,
            images,
            worksheet_drawings,
            worksheet_rels,
            drawing_rels,
            core_properties,
            app_properties,
            custom_properties,
            pivot_tables,
            pivot_cache_defs,
            pivot_cache_records,
            theme_xml,
            theme_colors,
            sheet_name_index,
            sheet_sparklines,
            sheet_vml,
            unknown_parts,
            deferred_parts,
            vba_blob,
            tables,
            raw_sheet_xml,
            slicer_defs,
            slicer_caches,
            sheet_threaded_comments,
            person_list,
            sheet_form_controls,
            streamed_sheets: HashMap::new(),
        })
    }

    /// Save the workbook to a file at the given path.
    ///
    /// The target format is inferred from the file extension. Supported
    /// extensions are `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, and `.xlam`.
    /// An unsupported extension returns [`Error::UnsupportedFileExtension`].
    ///
    /// The inferred format overrides the workbook's stored format so that
    /// the content type in the output always matches the extension.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let path = path.as_ref();
        let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");
        let target_format = WorkbookFormat::from_extension(ext)
            .ok_or_else(|| Error::UnsupportedFileExtension(ext.to_string()))?;

        let file = std::fs::File::create(path)?;
        let mut zip = zip::ZipWriter::new(file);
        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated)
            .compression_level(Some(1));
        self.write_zip_contents(&mut zip, options, Some(target_format))?;
        zip.finish().map_err(|e| Error::Zip(e.to_string()))?;
        Ok(())
    }

    /// Serialize the workbook to an in-memory buffer using the stored format.
    pub fn save_to_buffer(&self) -> Result<Vec<u8>> {
        // Estimate compressed output size to reduce reallocations.
        let estimated = self.worksheets.len() * 4000
            + self.sst_runtime.len() * 60
            + self.images.iter().map(|(_, d)| d.len()).sum::<usize>()
            + 32_000;
        let mut buf = Vec::with_capacity(estimated);
        {
            let cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            self.write_zip_contents(&mut zip, options, None)?;
            zip.finish().map_err(|e| Error::Zip(e.to_string()))?;
        }
        Ok(buf)
    }

    /// Open a workbook from an in-memory `.xlsx` buffer.
    pub fn open_from_buffer(data: &[u8]) -> Result<Self> {
        Self::open_from_buffer_with_options(data, &OpenOptions::default())
    }

    /// Open a workbook from an in-memory buffer with custom parsing options.
    pub fn open_from_buffer_with_options(data: &[u8], options: &OpenOptions) -> Result<Self> {
        // Detect encrypted files (CFB container)
        #[cfg(feature = "encryption")]
        if data.len() >= 8 {
            if let Ok(crate::crypt::ContainerFormat::Cfb) =
                crate::crypt::detect_container_format(data)
            {
                return Err(Error::FileEncrypted);
            }
        }

        let cursor = std::io::Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor).map_err(|e| Error::Zip(e.to_string()))?;
        Self::from_archive(&mut archive, options)
    }

    /// Open an encrypted `.xlsx` file using a password.
    ///
    /// The file must be in OLE/CFB container format. Supports both Standard
    /// Encryption (Office 2007, AES-128-ECB) and Agile Encryption (Office
    /// 2010+, AES-256-CBC).
    #[cfg(feature = "encryption")]
    pub fn open_with_password<P: AsRef<Path>>(path: P, password: &str) -> Result<Self> {
        let data = std::fs::read(path.as_ref())?;
        let decrypted_zip = crate::crypt::decrypt_xlsx(&data, password)?;
        let cursor = std::io::Cursor::new(decrypted_zip);
        let mut archive = zip::ZipArchive::new(cursor).map_err(|e| Error::Zip(e.to_string()))?;
        Self::from_archive(&mut archive, &OpenOptions::default())
    }

    /// Save the workbook as an encrypted `.xlsx` file using Agile Encryption
    /// (AES-256-CBC + SHA-512, 100K iterations).
    #[cfg(feature = "encryption")]
    pub fn save_with_password<P: AsRef<Path>>(&self, path: P, password: &str) -> Result<()> {
        // First, serialize to an in-memory ZIP buffer
        let mut zip_buf = Vec::new();
        {
            let cursor = std::io::Cursor::new(&mut zip_buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            self.write_zip_contents(&mut zip, options, None)?;
            zip.finish().map_err(|e| Error::Zip(e.to_string()))?;
        }

        // Encrypt and write to CFB container
        let cfb_data = crate::crypt::encrypt_xlsx(&zip_buf, password)?;
        std::fs::write(path.as_ref(), &cfb_data)?;
        Ok(())
    }

    /// Write all workbook parts into the given ZIP writer.
    ///
    /// When `format_override` is `Some`, that format is used for the workbook
    /// content type instead of the stored `self.format`. This allows `save()`
    /// to infer the format from the file extension without mutating `self`.
    fn write_zip_contents<W: std::io::Write + std::io::Seek>(
        &self,
        zip: &mut zip::ZipWriter<W>,
        options: SimpleFileOptions,
        format_override: Option<WorkbookFormat>,
    ) -> Result<()> {
        let effective_format = format_override.unwrap_or(self.format);
        let mut content_types = self.content_types.clone();

        // Ensure the workbook override content type matches the effective format.
        if let Some(wb_override) = content_types
            .overrides
            .iter_mut()
            .find(|o| o.part_name == "/xl/workbook.xml")
        {
            wb_override.content_type = effective_format.content_type().to_string();
        }

        // Ensure VBA project content type override and workbook relationship are
        // present when a VBA blob exists, and absent when it does not.
        // Skip when deferred_parts is non-empty: relationships are already correct.
        let has_deferred = !self.deferred_parts.is_empty();
        let mut workbook_rels = self.workbook_rels.clone();
        if self.vba_blob.is_some() {
            let vba_part_name = "/xl/vbaProject.bin";
            if !content_types
                .overrides
                .iter()
                .any(|o| o.part_name == vba_part_name)
            {
                content_types.overrides.push(ContentTypeOverride {
                    part_name: vba_part_name.to_string(),
                    content_type: VBA_PROJECT_CONTENT_TYPE.to_string(),
                });
            }
            if !content_types.defaults.iter().any(|d| d.extension == "bin") {
                content_types.defaults.push(ContentTypeDefault {
                    extension: "bin".to_string(),
                    content_type: VBA_PROJECT_CONTENT_TYPE.to_string(),
                });
            }
            if !workbook_rels
                .relationships
                .iter()
                .any(|r| r.rel_type == VBA_PROJECT_REL_TYPE)
            {
                let rid = crate::sheet::next_rid(&workbook_rels.relationships);
                workbook_rels.relationships.push(Relationship {
                    id: rid,
                    rel_type: VBA_PROJECT_REL_TYPE.to_string(),
                    target: "vbaProject.bin".to_string(),
                    target_mode: None,
                });
            }
        } else if !has_deferred {
            content_types
                .overrides
                .retain(|o| o.content_type != VBA_PROJECT_CONTENT_TYPE);
            workbook_rels
                .relationships
                .retain(|r| r.rel_type != VBA_PROJECT_REL_TYPE);
        }

        let mut worksheet_rels = self.worksheet_rels.clone();

        // Synchronize comment/form-control VML parts with worksheet relationships/content types.
        // Per-sheet VML bytes to write: (sheet_idx, zip_path, bytes).
        let mut vml_parts_to_write: Vec<(usize, String, Vec<u8>)> = Vec::new();
        // Per-sheet legacy drawing relationship IDs for worksheet XML serialization.
        let mut legacy_drawing_rids: HashMap<usize, String> = HashMap::new();

        // Ensure the vml extension default content type is present if any VML exists.
        let mut has_any_vml = false;

        // When deferred_parts is non-empty (ReadFast open), skip comment/VML
        // synchronization. The original relationships and content types are already
        // correct, and deferred_parts will supply the raw bytes on save.
        for sheet_idx in 0..self.worksheets.len() {
            let has_comments = self
                .sheet_comments
                .get(sheet_idx)
                .and_then(|c| c.as_ref())
                .is_some();
            let has_form_controls = self
                .sheet_form_controls
                .get(sheet_idx)
                .map(|v| !v.is_empty())
                .unwrap_or(false);
            let has_preserved_vml = self
                .sheet_vml
                .get(sheet_idx)
                .and_then(|v| v.as_ref())
                .is_some();

            if has_deferred {
                continue;
            }

            if let Some(rels) = worksheet_rels.get_mut(&sheet_idx) {
                rels.relationships
                    .retain(|r| r.rel_type != rel_types::COMMENTS);
                rels.relationships
                    .retain(|r| r.rel_type != rel_types::VML_DRAWING);
            }

            let needs_vml = has_comments || has_form_controls || has_preserved_vml;
            if !needs_vml && !has_comments {
                continue;
            }

            if has_comments {
                let comment_path = format!("xl/comments{}.xml", sheet_idx + 1);
                let part_name = format!("/{}", comment_path);
                if !content_types
                    .overrides
                    .iter()
                    .any(|o| o.part_name == part_name && o.content_type == mime_types::COMMENTS)
                {
                    content_types.overrides.push(ContentTypeOverride {
                        part_name,
                        content_type: mime_types::COMMENTS.to_string(),
                    });
                }

                let sheet_path = self.sheet_part_path(sheet_idx);
                let target = relative_relationship_target(&sheet_path, &comment_path);
                let rels = worksheet_rels
                    .entry(sheet_idx)
                    .or_insert_with(default_relationships);
                let rid = crate::sheet::next_rid(&rels.relationships);
                rels.relationships.push(Relationship {
                    id: rid,
                    rel_type: rel_types::COMMENTS.to_string(),
                    target,
                    target_mode: None,
                });
            }

            if !needs_vml {
                continue;
            }

            // Build VML bytes combining comments and form controls.
            let vml_path = format!("xl/drawings/vmlDrawing{}.vml", sheet_idx + 1);
            let vml_bytes = if has_comments && has_form_controls {
                // Both comments and form controls: start with comment VML, then append controls.
                let comment_vml =
                    if let Some(bytes) = self.sheet_vml.get(sheet_idx).and_then(|v| v.as_ref()) {
                        bytes.clone()
                    } else if let Some(Some(comments)) = self.sheet_comments.get(sheet_idx) {
                        let cells: Vec<&str> = comments
                            .comment_list
                            .comments
                            .iter()
                            .map(|c| c.r#ref.as_str())
                            .collect();
                        crate::vml::build_vml_drawing(&cells).into_bytes()
                    } else {
                        continue;
                    };
                let shape_count = crate::control::count_vml_shapes(&comment_vml);
                let start_id = 1025 + shape_count;
                let form_controls = &self.sheet_form_controls[sheet_idx];
                crate::control::merge_vml_controls(&comment_vml, form_controls, start_id)
            } else if has_comments {
                if let Some(bytes) = self.sheet_vml.get(sheet_idx).and_then(|v| v.as_ref()) {
                    bytes.clone()
                } else if let Some(Some(comments)) = self.sheet_comments.get(sheet_idx) {
                    let cells: Vec<&str> = comments
                        .comment_list
                        .comments
                        .iter()
                        .map(|c| c.r#ref.as_str())
                        .collect();
                    crate::vml::build_vml_drawing(&cells).into_bytes()
                } else {
                    continue;
                }
            } else if has_form_controls {
                // Hydrated form controls only (no comments).
                let form_controls = &self.sheet_form_controls[sheet_idx];
                crate::control::build_form_control_vml(form_controls, 1025).into_bytes()
            } else if let Some(Some(vml)) = self.sheet_vml.get(sheet_idx) {
                // Preserved VML bytes only (controls not hydrated, no comments).
                vml.clone()
            } else {
                continue;
            };

            let vml_part_name = format!("/{}", vml_path);
            if !content_types
                .overrides
                .iter()
                .any(|o| o.part_name == vml_part_name && o.content_type == mime_types::VML_DRAWING)
            {
                content_types.overrides.push(ContentTypeOverride {
                    part_name: vml_part_name,
                    content_type: mime_types::VML_DRAWING.to_string(),
                });
            }

            let sheet_path = self.sheet_part_path(sheet_idx);
            let rels = worksheet_rels
                .entry(sheet_idx)
                .or_insert_with(default_relationships);
            let vml_target = relative_relationship_target(&sheet_path, &vml_path);
            let vml_rid = crate::sheet::next_rid(&rels.relationships);
            rels.relationships.push(Relationship {
                id: vml_rid.clone(),
                rel_type: rel_types::VML_DRAWING.to_string(),
                target: vml_target,
                target_mode: None,
            });

            legacy_drawing_rids.insert(sheet_idx, vml_rid);
            vml_parts_to_write.push((sheet_idx, vml_path, vml_bytes));
            has_any_vml = true;
        }

        // Add vml extension default content type if needed.
        if has_any_vml && !content_types.defaults.iter().any(|d| d.extension == "vml") {
            content_types.defaults.push(ContentTypeDefault {
                extension: "vml".to_string(),
                content_type: mime_types::VML_DRAWING.to_string(),
            });
        }

        // Synchronize table parts with worksheet relationships and content types.
        // Also build tableParts references for each worksheet.
        // Skip when deferred_parts is non-empty: relationships are already correct.
        let mut table_parts_by_sheet: HashMap<usize, Vec<String>> = HashMap::new();
        if !has_deferred {
            for (sheet_idx, _) in self.worksheets.iter().enumerate() {
                if let Some(rels) = worksheet_rels.get_mut(&sheet_idx) {
                    rels.relationships
                        .retain(|r| r.rel_type != rel_types::TABLE);
                }
            }
            content_types
                .overrides
                .retain(|o| o.content_type != mime_types::TABLE);
        }
        for (table_path, _table_xml, sheet_idx) in &self.tables {
            let part_name = format!("/{table_path}");
            content_types.overrides.push(ContentTypeOverride {
                part_name,
                content_type: mime_types::TABLE.to_string(),
            });

            let sheet_path = self.sheet_part_path(*sheet_idx);
            let target = relative_relationship_target(&sheet_path, table_path);
            let rels = worksheet_rels
                .entry(*sheet_idx)
                .or_insert_with(default_relationships);
            let rid = crate::sheet::next_rid(&rels.relationships);
            rels.relationships.push(Relationship {
                id: rid.clone(),
                rel_type: rel_types::TABLE.to_string(),
                target,
                target_mode: None,
            });
            table_parts_by_sheet
                .entry(*sheet_idx)
                .or_default()
                .push(rid);
        }

        // Register threaded comment content types and relationships before writing.
        let has_any_threaded = self.sheet_threaded_comments.iter().any(|tc| tc.is_some());
        if has_any_threaded {
            for (i, tc) in self.sheet_threaded_comments.iter().enumerate() {
                if tc.is_some() {
                    let tc_path = format!("xl/threadedComments/threadedComment{}.xml", i + 1);
                    let tc_part_name = format!("/{tc_path}");
                    if !content_types.overrides.iter().any(|o| {
                        o.part_name == tc_part_name
                            && o.content_type
                                == sheetkit_xml::threaded_comment::THREADED_COMMENTS_CONTENT_TYPE
                    }) {
                        content_types.overrides.push(ContentTypeOverride {
                            part_name: tc_part_name,
                            content_type:
                                sheetkit_xml::threaded_comment::THREADED_COMMENTS_CONTENT_TYPE
                                    .to_string(),
                        });
                    }

                    let sheet_path = self.sheet_part_path(i);
                    let target = relative_relationship_target(&sheet_path, &tc_path);
                    let rels = worksheet_rels
                        .entry(i)
                        .or_insert_with(default_relationships);
                    if !rels.relationships.iter().any(|r| {
                        r.rel_type == sheetkit_xml::threaded_comment::REL_TYPE_THREADED_COMMENT
                    }) {
                        let rid = crate::sheet::next_rid(&rels.relationships);
                        rels.relationships.push(Relationship {
                            id: rid,
                            rel_type: sheetkit_xml::threaded_comment::REL_TYPE_THREADED_COMMENT
                                .to_string(),
                            target,
                            target_mode: None,
                        });
                    }
                }
            }

            let person_part_name = "/xl/persons/person.xml";
            if !content_types.overrides.iter().any(|o| {
                o.part_name == person_part_name
                    && o.content_type == sheetkit_xml::threaded_comment::PERSON_LIST_CONTENT_TYPE
            }) {
                content_types.overrides.push(ContentTypeOverride {
                    part_name: person_part_name.to_string(),
                    content_type: sheetkit_xml::threaded_comment::PERSON_LIST_CONTENT_TYPE
                        .to_string(),
                });
            }

            // Add person relationship to workbook_rels so Excel can discover the person list.
            if !workbook_rels
                .relationships
                .iter()
                .any(|r| r.rel_type == sheetkit_xml::threaded_comment::REL_TYPE_PERSON)
            {
                let rid = crate::sheet::next_rid(&workbook_rels.relationships);
                workbook_rels.relationships.push(Relationship {
                    id: rid,
                    rel_type: sheetkit_xml::threaded_comment::REL_TYPE_PERSON.to_string(),
                    target: "persons/person.xml".to_string(),
                    target_mode: None,
                });
            }
        }

        // [Content_Types].xml
        write_xml_part(zip, "[Content_Types].xml", &content_types, options)?;

        // _rels/.rels
        write_xml_part(zip, "_rels/.rels", &self.package_rels, options)?;

        // xl/workbook.xml
        write_xml_part(zip, "xl/workbook.xml", &self.workbook_xml, options)?;

        // xl/_rels/workbook.xml.rels
        write_xml_part(zip, "xl/_rels/workbook.xml.rels", &workbook_rels, options)?;

        // xl/worksheets/sheet{N}.xml
        for (i, (_name, ws)) in self.worksheets.iter().enumerate() {
            let entry_name = self.sheet_part_path(i);

            // If the sheet has streamed data, write it directly from the temp file.
            if let Some(streamed) = self.streamed_sheets.get(&i) {
                crate::stream::write_streamed_sheet(zip, &entry_name, streamed, options)?;
                continue;
            }

            // If the sheet was not parsed (selective open), write raw bytes directly.
            if let Some(Some(raw_bytes)) = self.raw_sheet_xml.get(i) {
                zip.start_file(&entry_name, options)
                    .map_err(|e| Error::Zip(e.to_string()))?;
                zip.write_all(raw_bytes)?;
                continue;
            }

            let empty_sparklines: Vec<crate::sparkline::SparklineConfig> = vec![];
            let sparklines = self.sheet_sparklines.get(i).unwrap_or(&empty_sparklines);
            let legacy_rid = legacy_drawing_rids.get(&i).map(|s| s.as_str());
            let sheet_table_rids = table_parts_by_sheet.get(&i);
            let stale_table_parts = sheet_table_rids.is_none() && ws.table_parts.is_some();
            let has_extras = legacy_rid.is_some()
                || !sparklines.is_empty()
                || sheet_table_rids.is_some()
                || stale_table_parts;

            if !has_extras {
                write_xml_part(zip, &entry_name, ws, options)?;
            } else {
                let ws_to_serialize;
                let ws_ref = if let Some(rids) = sheet_table_rids {
                    ws_to_serialize = {
                        let mut cloned = ws.clone();
                        use sheetkit_xml::worksheet::{TablePart, TableParts};
                        cloned.table_parts = Some(TableParts {
                            count: Some(rids.len() as u32),
                            table_parts: rids
                                .iter()
                                .map(|rid| TablePart { r_id: rid.clone() })
                                .collect(),
                        });
                        cloned
                    };
                    &ws_to_serialize
                } else if stale_table_parts {
                    ws_to_serialize = {
                        let mut cloned = ws.clone();
                        cloned.table_parts = None;
                        cloned
                    };
                    &ws_to_serialize
                } else {
                    ws
                };
                let xml = serialize_worksheet_with_extras(ws_ref, sparklines, legacy_rid)?;
                zip.start_file(&entry_name, options)
                    .map_err(|e| Error::Zip(e.to_string()))?;
                zip.write_all(xml.as_bytes())?;
            }
        }

        // xl/styles.xml
        write_xml_part(zip, "xl/styles.xml", &self.stylesheet, options)?;

        // xl/sharedStrings.xml -- write from the runtime SST
        let sst_xml = self.sst_runtime.to_sst();
        write_xml_part(zip, "xl/sharedStrings.xml", &sst_xml, options)?;

        // xl/comments{N}.xml -- write per-sheet comments
        for (i, comments) in self.sheet_comments.iter().enumerate() {
            if let Some(ref c) = comments {
                let entry_name = format!("xl/comments{}.xml", i + 1);
                write_xml_part(zip, &entry_name, c, options)?;
            }
        }

        // xl/drawings/vmlDrawing{N}.vml -- write VML drawing parts
        for (_sheet_idx, vml_path, vml_bytes) in &vml_parts_to_write {
            zip.start_file(vml_path, options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(vml_bytes)?;
        }

        // xl/drawings/drawing{N}.xml -- write drawing parts
        for (path, drawing) in &self.drawings {
            write_xml_part(zip, path, drawing, options)?;
        }

        // xl/charts/chart{N}.xml -- write chart parts
        for (path, chart) in &self.charts {
            write_xml_part(zip, path, chart, options)?;
        }
        for (path, data) in &self.raw_charts {
            if self.charts.iter().any(|(p, _)| p == path) {
                continue;
            }
            zip.start_file(path, options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(data)?;
        }

        // xl/media/image{N}.{ext} -- write image data
        for (path, data) in &self.images {
            zip.start_file(path, options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(data)?;
        }

        // xl/worksheets/_rels/sheet{N}.xml.rels -- write worksheet relationships
        for (sheet_idx, rels) in &worksheet_rels {
            let sheet_path = self.sheet_part_path(*sheet_idx);
            let path = relationship_part_path(&sheet_path);
            write_xml_part(zip, &path, rels, options)?;
        }

        // xl/drawings/_rels/drawing{N}.xml.rels -- write drawing relationships
        for (drawing_idx, rels) in &self.drawing_rels {
            if let Some((drawing_path, _)) = self.drawings.get(*drawing_idx) {
                let path = relationship_part_path(drawing_path);
                write_xml_part(zip, &path, rels, options)?;
            }
        }

        // xl/pivotTables/pivotTable{N}.xml
        for (path, pt) in &self.pivot_tables {
            write_xml_part(zip, path, pt, options)?;
        }

        // xl/pivotCache/pivotCacheDefinition{N}.xml
        for (path, pcd) in &self.pivot_cache_defs {
            write_xml_part(zip, path, pcd, options)?;
        }

        // xl/pivotCache/pivotCacheRecords{N}.xml
        for (path, pcr) in &self.pivot_cache_records {
            write_xml_part(zip, path, pcr, options)?;
        }

        // xl/tables/table{N}.xml
        for (path, table_xml, _sheet_idx) in &self.tables {
            write_xml_part(zip, path, table_xml, options)?;
        }

        // xl/slicers/slicer{N}.xml
        for (path, sd) in &self.slicer_defs {
            write_xml_part(zip, path, sd, options)?;
        }

        // xl/slicerCaches/slicerCache{N}.xml (manual serialization)
        for (path, scd) in &self.slicer_caches {
            let xml_str = format!(
                "{}\n{}",
                XML_DECLARATION,
                sheetkit_xml::slicer::serialize_slicer_cache(scd),
            );
            zip.start_file(path, options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(xml_str.as_bytes())?;
        }

        // xl/theme/theme1.xml
        {
            let default_theme = crate::theme::default_theme_xml();
            let theme_bytes = self.theme_xml.as_deref().unwrap_or(&default_theme);
            zip.start_file("xl/theme/theme1.xml", options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(theme_bytes)?;
        }

        // xl/vbaProject.bin -- write VBA blob if present
        if let Some(ref blob) = self.vba_blob {
            zip.start_file("xl/vbaProject.bin", options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(blob)?;
        }

        // docProps/core.xml
        if let Some(ref props) = self.core_properties {
            let xml_str = sheetkit_xml::doc_props::serialize_core_properties(props);
            zip.start_file("docProps/core.xml", options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(xml_str.as_bytes())?;
        }

        // docProps/app.xml
        if let Some(ref props) = self.app_properties {
            write_xml_part(zip, "docProps/app.xml", props, options)?;
        }

        // docProps/custom.xml
        if let Some(ref props) = self.custom_properties {
            let xml_str = sheetkit_xml::doc_props::serialize_custom_properties(props);
            zip.start_file("docProps/custom.xml", options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(xml_str.as_bytes())?;
        }

        // xl/threadedComments/threadedComment{N}.xml
        if has_any_threaded {
            for (i, tc) in self.sheet_threaded_comments.iter().enumerate() {
                if let Some(ref tc_data) = tc {
                    let tc_path = format!("xl/threadedComments/threadedComment{}.xml", i + 1);
                    write_xml_part(zip, &tc_path, tc_data, options)?;
                }
            }
            write_xml_part(zip, "xl/persons/person.xml", &self.person_list, options)?;
        }

        // Write back unknown parts preserved from the original file.
        for (path, data) in &self.unknown_parts {
            zip.start_file(path, options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(data)?;
        }

        // Write back deferred parts from ReadFast open (raw bytes, unparsed).
        for (path, data) in &self.deferred_parts {
            zip.start_file(path, options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(data)?;
        }

        Ok(())
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

/// Serialize a value to XML with the standard XML declaration prepended.
pub(crate) fn serialize_xml<T: Serialize>(value: &T) -> Result<String> {
    let body = quick_xml::se::to_string(value).map_err(|e| Error::XmlParse(e.to_string()))?;
    let mut result = String::with_capacity(XML_DECLARATION.len() + 1 + body.len());
    result.push_str(XML_DECLARATION);
    result.push('\n');
    result.push_str(&body);
    Ok(result)
}

/// Read a ZIP entry and deserialize it from XML.
pub(crate) fn read_xml_part<T: serde::de::DeserializeOwned, R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    name: &str,
) -> Result<T> {
    let mut entry = archive
        .by_name(name)
        .map_err(|e| Error::Zip(e.to_string()))?;
    let size_hint = entry.size() as usize;
    let mut content = String::with_capacity(size_hint);
    entry
        .read_to_string(&mut content)
        .map_err(|e| Error::Zip(e.to_string()))?;
    quick_xml::de::from_str(&content).map_err(|e| Error::XmlDeserialize(e.to_string()))
}

/// Read a ZIP entry as a raw string (no serde deserialization).
pub(crate) fn read_string_part<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    name: &str,
) -> Result<String> {
    let mut entry = archive
        .by_name(name)
        .map_err(|e| Error::Zip(e.to_string()))?;
    let size_hint = entry.size() as usize;
    let mut content = String::with_capacity(size_hint);
    entry
        .read_to_string(&mut content)
        .map_err(|e| Error::Zip(e.to_string()))?;
    Ok(content)
}

/// Read a ZIP entry as raw bytes.
pub(crate) fn read_bytes_part<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    name: &str,
) -> Result<Vec<u8>> {
    let mut entry = archive
        .by_name(name)
        .map_err(|e| Error::Zip(e.to_string()))?;
    let size_hint = entry.size() as usize;
    let mut content = Vec::with_capacity(size_hint);
    entry
        .read_to_end(&mut content)
        .map_err(|e| Error::Zip(e.to_string()))?;
    Ok(content)
}

/// Serialize a worksheet with optional sparklines and legacy drawing injected
/// via string manipulation, avoiding a full WorksheetXml clone.
pub(crate) fn serialize_worksheet_with_extras(
    ws: &WorksheetXml,
    sparklines: &[crate::sparkline::SparklineConfig],
    legacy_drawing_rid: Option<&str>,
) -> Result<String> {
    let body = quick_xml::se::to_string(ws).map_err(|e| Error::XmlParse(e.to_string()))?;

    let closing = "</worksheet>";
    let ext_xml = if sparklines.is_empty() {
        String::new()
    } else {
        build_sparkline_ext_xml(sparklines)
    };
    let legacy_xml = if let Some(rid) = legacy_drawing_rid {
        format!("<legacyDrawing r:id=\"{rid}\"/>")
    } else {
        String::new()
    };

    if let Some(pos) = body.rfind(closing) {
        // If injecting a legacy drawing, strip any existing one from the serde output
        // to avoid duplicates (the original ws.legacy_drawing may already be set).
        let body_prefix = &body[..pos];
        let stripped;
        let prefix = if !legacy_xml.is_empty() {
            if let Some(ld_start) = body_prefix.find("<legacyDrawing ") {
                // Find the end of the self-closing element.
                let ld_end = body_prefix[ld_start..]
                    .find("/>")
                    .map(|e| ld_start + e + 2)
                    .unwrap_or(ld_start);
                stripped = format!("{}{}", &body_prefix[..ld_start], &body_prefix[ld_end..]);
                stripped.as_str()
            } else {
                body_prefix
            }
        } else {
            body_prefix
        };

        let extra_len = ext_xml.len() + legacy_xml.len();
        let mut result = String::with_capacity(XML_DECLARATION.len() + 1 + body.len() + extra_len);
        result.push_str(XML_DECLARATION);
        result.push('\n');
        result.push_str(prefix);
        result.push_str(&legacy_xml);
        result.push_str(&ext_xml);
        result.push_str(closing);
        Ok(result)
    } else {
        Ok(format!("{XML_DECLARATION}\n{body}"))
    }
}

/// Build the extLst XML block for sparklines using manual string construction.
pub(crate) fn build_sparkline_ext_xml(sparklines: &[crate::sparkline::SparklineConfig]) -> String {
    use std::fmt::Write;
    let mut xml = String::new();
    let _ = write!(
        xml,
        "<extLst>\
         <ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" \
         uri=\"{{05C60535-1F16-4fd2-B633-F4F36F0B64E0}}\">\
         <x14:sparklineGroups \
         xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
    );
    for config in sparklines {
        let group = crate::sparkline::config_to_xml_group(config);
        let _ = write!(xml, "<x14:sparklineGroup");
        if let Some(ref t) = group.sparkline_type {
            let _ = write!(xml, " type=\"{t}\"");
        }
        if group.markers == Some(true) {
            let _ = write!(xml, " markers=\"1\"");
        }
        if group.high == Some(true) {
            let _ = write!(xml, " high=\"1\"");
        }
        if group.low == Some(true) {
            let _ = write!(xml, " low=\"1\"");
        }
        if group.first == Some(true) {
            let _ = write!(xml, " first=\"1\"");
        }
        if group.last == Some(true) {
            let _ = write!(xml, " last=\"1\"");
        }
        if group.negative == Some(true) {
            let _ = write!(xml, " negative=\"1\"");
        }
        if group.display_x_axis == Some(true) {
            let _ = write!(xml, " displayXAxis=\"1\"");
        }
        if let Some(w) = group.line_weight {
            let _ = write!(xml, " lineWeight=\"{w}\"");
        }
        let _ = write!(xml, "><x14:sparklines>");
        for sp in &group.sparklines.items {
            let _ = write!(
                xml,
                "<x14:sparkline><xm:f>{}</xm:f><xm:sqref>{}</xm:sqref></x14:sparkline>",
                sp.formula, sp.sqref
            );
        }
        let _ = write!(xml, "</x14:sparklines></x14:sparklineGroup>");
    }
    let _ = write!(xml, "</x14:sparklineGroups></ext></extLst>");
    xml
}

/// Parse sparkline configurations from raw worksheet XML content.
pub(crate) fn parse_sparklines_from_xml(xml: &str) -> Vec<crate::sparkline::SparklineConfig> {
    use crate::sparkline::{SparklineConfig, SparklineType};

    let mut sparklines = Vec::new();

    // Find all sparklineGroup elements and parse their attributes and children.
    let mut search_from = 0;
    while let Some(group_start) = xml[search_from..].find("<x14:sparklineGroup") {
        let abs_start = search_from + group_start;
        let group_end_tag = "</x14:sparklineGroup>";
        let abs_end = match xml[abs_start..].find(group_end_tag) {
            Some(pos) => abs_start + pos + group_end_tag.len(),
            None => break,
        };
        let group_xml = &xml[abs_start..abs_end];

        // Parse group-level attributes.
        let sparkline_type = extract_xml_attr(group_xml, "type")
            .and_then(|s| SparklineType::parse(&s))
            .unwrap_or_default();
        let markers = extract_xml_bool_attr(group_xml, "markers");
        let high_point = extract_xml_bool_attr(group_xml, "high");
        let low_point = extract_xml_bool_attr(group_xml, "low");
        let first_point = extract_xml_bool_attr(group_xml, "first");
        let last_point = extract_xml_bool_attr(group_xml, "last");
        let negative_points = extract_xml_bool_attr(group_xml, "negative");
        let show_axis = extract_xml_bool_attr(group_xml, "displayXAxis");
        let line_weight =
            extract_xml_attr(group_xml, "lineWeight").and_then(|s| s.parse::<f64>().ok());

        // Parse individual sparkline entries within this group.
        let mut sp_from = 0;
        while let Some(sp_start) = group_xml[sp_from..].find("<x14:sparkline>") {
            let sp_abs = sp_from + sp_start;
            let sp_end_tag = "</x14:sparkline>";
            let sp_abs_end = match group_xml[sp_abs..].find(sp_end_tag) {
                Some(pos) => sp_abs + pos + sp_end_tag.len(),
                None => break,
            };
            let sp_xml = &group_xml[sp_abs..sp_abs_end];

            let formula = extract_xml_element(sp_xml, "xm:f").unwrap_or_default();
            let sqref = extract_xml_element(sp_xml, "xm:sqref").unwrap_or_default();

            if !formula.is_empty() && !sqref.is_empty() {
                sparklines.push(SparklineConfig {
                    data_range: formula,
                    location: sqref,
                    sparkline_type: sparkline_type.clone(),
                    markers,
                    high_point,
                    low_point,
                    first_point,
                    last_point,
                    negative_points,
                    show_axis,
                    line_weight,
                    style: None,
                });
            }
            sp_from = sp_abs_end;
        }
        search_from = abs_end;
    }
    sparklines
}

/// Extract an XML attribute value from an element's opening tag.
///
/// Uses manual search to avoid allocating format strings for patterns.
pub(crate) fn extract_xml_attr(xml: &str, attr: &str) -> Option<String> {
    // Search for ` attr="` or ` attr='` without allocating pattern strings.
    for quote in ['"', '\''] {
        // Build the search target: " attr=" (space + attr name + = + quote)
        let haystack = xml.as_bytes();
        let attr_bytes = attr.as_bytes();
        let mut pos = 0;
        while pos + 1 + attr_bytes.len() + 2 <= haystack.len() {
            if haystack[pos] == b' '
                && haystack[pos + 1..pos + 1 + attr_bytes.len()] == *attr_bytes
                && haystack[pos + 1 + attr_bytes.len()] == b'='
                && haystack[pos + 1 + attr_bytes.len() + 1] == quote as u8
            {
                let val_start = pos + 1 + attr_bytes.len() + 2;
                if let Some(end) = xml[val_start..].find(quote) {
                    return Some(xml[val_start..val_start + end].to_string());
                }
            }
            pos += 1;
        }
    }
    None
}

/// Extract a boolean attribute from an XML element (true for "1" or "true").
pub(crate) fn extract_xml_bool_attr(xml: &str, attr: &str) -> bool {
    extract_xml_attr(xml, attr)
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        .unwrap_or(false)
}

/// Extract the text content of an XML element like `<tag>content</tag>`.
pub(crate) fn extract_xml_element(xml: &str, tag: &str) -> Option<String> {
    let open = format!("<{tag}>");
    let close = format!("</{tag}>");
    let start = xml.find(&open)?;
    let content_start = start + open.len();
    let end = xml[content_start..].find(&close)?;
    Some(xml[content_start..content_start + end].to_string())
}

/// Serialize a value to XML and write it as a ZIP entry.
pub(crate) fn write_xml_part<T: Serialize, W: std::io::Write + std::io::Seek>(
    zip: &mut zip::ZipWriter<W>,
    name: &str,
    value: &T,
    options: SimpleFileOptions,
) -> Result<()> {
    let xml = serialize_xml(value)?;
    zip.start_file(name, options)
        .map_err(|e| Error::Zip(e.to_string()))?;
    zip.write_all(xml.as_bytes())?;
    Ok(())
}

/// Fast column number extraction from a cell reference string like "A1", "BC42".
///
/// Parses only the alphabetic prefix (column letters) and converts to a
/// 1-based column number. Much faster than [`cell_name_to_coordinates`] because
/// it skips row parsing and avoids error handling overhead.
fn fast_col_number(cell_ref: &str) -> u32 {
    let mut col: u32 = 0;
    for b in cell_ref.bytes() {
        if b.is_ascii_alphabetic() {
            col = col * 26 + (b.to_ascii_uppercase() - b'A') as u32 + 1;
        } else {
            break;
        }
    }
    col
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn test_fast_col_number() {
        assert_eq!(fast_col_number("A1"), 1);
        assert_eq!(fast_col_number("B1"), 2);
        assert_eq!(fast_col_number("Z1"), 26);
        assert_eq!(fast_col_number("AA1"), 27);
        assert_eq!(fast_col_number("AZ1"), 52);
        assert_eq!(fast_col_number("BA1"), 53);
        assert_eq!(fast_col_number("XFD1"), 16384);
    }

    #[test]
    fn test_extract_xml_attr() {
        let xml = r#"<tag type="column" markers="1" weight="2.5">"#;
        assert_eq!(extract_xml_attr(xml, "type"), Some("column".to_string()));
        assert_eq!(extract_xml_attr(xml, "markers"), Some("1".to_string()));
        assert_eq!(extract_xml_attr(xml, "weight"), Some("2.5".to_string()));
        assert_eq!(extract_xml_attr(xml, "missing"), None);
        // Single-quoted attributes
        let xml2 = "<tag name='hello'>";
        assert_eq!(extract_xml_attr(xml2, "name"), Some("hello".to_string()));
    }

    #[test]
    fn test_extract_xml_bool_attr() {
        let xml = r#"<tag markers="1" hidden="0" visible="true">"#;
        assert!(extract_xml_bool_attr(xml, "markers"));
        assert!(!extract_xml_bool_attr(xml, "hidden"));
        assert!(extract_xml_bool_attr(xml, "visible"));
        assert!(!extract_xml_bool_attr(xml, "missing"));
    }

    #[test]
    fn test_new_workbook_has_sheet1() {
        let wb = Workbook::new();
        assert_eq!(wb.sheet_names(), vec!["Sheet1"]);
    }

    #[test]
    fn test_new_workbook_save_creates_file() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("test.xlsx");
        let wb = Workbook::new();
        wb.save(&path).unwrap();
        assert!(path.exists());
    }

    #[test]
    fn test_save_and_open_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("roundtrip.xlsx");

        let wb = Workbook::new();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1"]);
    }

    #[test]
    fn test_saved_file_is_valid_zip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("valid.xlsx");
        let wb = Workbook::new();
        wb.save(&path).unwrap();

        // Verify it's a valid ZIP with expected entries
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        let expected_files = [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/worksheets/sheet1.xml",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
        ];

        for name in &expected_files {
            assert!(archive.by_name(name).is_ok(), "Missing ZIP entry: {}", name);
        }
    }

    #[test]
    fn test_open_nonexistent_file_returns_error() {
        let result = Workbook::open("/nonexistent/path.xlsx");
        assert!(result.is_err());
    }

    #[test]
    fn test_saved_xml_has_declarations() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("decl.xlsx");
        let wb = Workbook::new();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        let mut content = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("[Content_Types].xml").unwrap(),
            &mut content,
        )
        .unwrap();
        assert!(content.starts_with("<?xml"));
    }

    #[test]
    fn test_default_trait() {
        let wb = Workbook::default();
        assert_eq!(wb.sheet_names(), vec!["Sheet1"]);
    }

    #[test]
    fn test_serialize_xml_helper() {
        let ct = ContentTypes::default();
        let xml = serialize_xml(&ct).unwrap();
        assert!(xml.starts_with("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"));
        assert!(xml.contains("<Types"));
    }

    #[test]
    fn test_save_to_buffer_and_open_from_buffer_roundtrip() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet1", "B2", CellValue::Number(42.0))
            .unwrap();

        let buf = wb.save_to_buffer().unwrap();
        assert!(!buf.is_empty());

        let wb2 = Workbook::open_from_buffer(&buf).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1"]);
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Hello".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Sheet1", "B2").unwrap(),
            CellValue::Number(42.0)
        );
    }

    #[test]
    fn test_save_to_buffer_produces_valid_zip() {
        let wb = Workbook::new();
        let buf = wb.save_to_buffer().unwrap();

        let cursor = std::io::Cursor::new(buf);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let expected_files = [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/worksheets/sheet1.xml",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
        ];

        for name in &expected_files {
            assert!(archive.by_name(name).is_ok(), "Missing ZIP entry: {}", name);
        }
    }

    #[test]
    fn test_open_from_buffer_invalid_data() {
        let result = Workbook::open_from_buffer(b"not a zip file");
        assert!(result.is_err());
    }

    #[cfg(feature = "encryption")]
    #[test]
    fn test_save_and_open_with_password_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("encrypted.xlsx");

        // Create a workbook with some data
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet1", "B2", CellValue::Number(42.0))
            .unwrap();

        // Save with password
        wb.save_with_password(&path, "test123").unwrap();

        // Verify it's a CFB file, not a ZIP
        let data = std::fs::read(&path).unwrap();
        assert_eq!(
            &data[..8],
            &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
        );

        // Open without password should fail
        let result = Workbook::open(&path);
        assert!(matches!(result, Err(Error::FileEncrypted)));

        // Open with wrong password should fail
        let result = Workbook::open_with_password(&path, "wrong");
        assert!(matches!(result, Err(Error::IncorrectPassword)));

        // Open with correct password should succeed
        let wb2 = Workbook::open_with_password(&path, "test123").unwrap();
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Hello".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Sheet1", "B2").unwrap(),
            CellValue::Number(42.0)
        );
    }

    /// Create a test xlsx buffer with extra custom ZIP entries that sheetkit
    /// does not natively handle.
    fn create_xlsx_with_custom_entries() -> Vec<u8> {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("hello".to_string()))
            .unwrap();
        let base_buf = wb.save_to_buffer().unwrap();

        // Re-open the ZIP and inject custom entries.
        let cursor = std::io::Cursor::new(&base_buf);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut out = Vec::new();
        {
            let out_cursor = std::io::Cursor::new(&mut out);
            let mut zip_writer = zip::ZipWriter::new(out_cursor);
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

            // Copy all existing entries.
            for i in 0..archive.len() {
                let mut entry = archive.by_index(i).unwrap();
                let name = entry.name().to_string();
                let mut data = Vec::new();
                std::io::Read::read_to_end(&mut entry, &mut data).unwrap();
                zip_writer.start_file(&name, options).unwrap();
                std::io::Write::write_all(&mut zip_writer, &data).unwrap();
            }

            // Add custom entries that sheetkit does not handle.
            zip_writer
                .start_file("customXml/item1.xml", options)
                .unwrap();
            std::io::Write::write_all(&mut zip_writer, b"<custom>data1</custom>").unwrap();

            zip_writer
                .start_file("customXml/itemProps1.xml", options)
                .unwrap();
            std::io::Write::write_all(
                &mut zip_writer,
                b"<ds:datastoreItem xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\"/>",
            )
            .unwrap();

            zip_writer
                .start_file("xl/printerSettings/printerSettings1.bin", options)
                .unwrap();
            std::io::Write::write_all(&mut zip_writer, b"\x00\x01\x02\x03PRINTER").unwrap();

            zip_writer.finish().unwrap();
        }
        out
    }

    #[test]
    fn test_unknown_zip_entries_preserved_on_roundtrip() {
        let buf = create_xlsx_with_custom_entries();

        // Open, verify the data is still accessible.
        let wb = Workbook::open_from_buffer(&buf).unwrap();
        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("hello".to_string())
        );

        // Save and re-open.
        let saved = wb.save_to_buffer().unwrap();
        let cursor = std::io::Cursor::new(&saved);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        // Verify custom entries are present in the output.
        let mut custom_xml = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("customXml/item1.xml").unwrap(),
            &mut custom_xml,
        )
        .unwrap();
        assert_eq!(custom_xml, "<custom>data1</custom>");

        let mut props_xml = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("customXml/itemProps1.xml").unwrap(),
            &mut props_xml,
        )
        .unwrap();
        assert!(props_xml.contains("datastoreItem"));

        let mut printer = Vec::new();
        std::io::Read::read_to_end(
            &mut archive
                .by_name("xl/printerSettings/printerSettings1.bin")
                .unwrap(),
            &mut printer,
        )
        .unwrap();
        assert_eq!(printer, b"\x00\x01\x02\x03PRINTER");
    }

    #[test]
    fn test_unknown_entries_survive_multiple_roundtrips() {
        let buf = create_xlsx_with_custom_entries();
        let wb1 = Workbook::open_from_buffer(&buf).unwrap();
        let buf2 = wb1.save_to_buffer().unwrap();
        let wb2 = Workbook::open_from_buffer(&buf2).unwrap();
        let buf3 = wb2.save_to_buffer().unwrap();

        let cursor = std::io::Cursor::new(&buf3);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let mut custom_xml = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("customXml/item1.xml").unwrap(),
            &mut custom_xml,
        )
        .unwrap();
        assert_eq!(custom_xml, "<custom>data1</custom>");

        let mut printer = Vec::new();
        std::io::Read::read_to_end(
            &mut archive
                .by_name("xl/printerSettings/printerSettings1.bin")
                .unwrap(),
            &mut printer,
        )
        .unwrap();
        assert_eq!(printer, b"\x00\x01\x02\x03PRINTER");
    }

    #[test]
    fn test_new_workbook_has_no_unknown_parts() {
        let wb = Workbook::new();
        let buf = wb.save_to_buffer().unwrap();
        let wb2 = Workbook::open_from_buffer(&buf).unwrap();
        assert!(wb2.unknown_parts.is_empty());
    }

    #[test]
    fn test_known_entries_not_duplicated_as_unknown() {
        let wb = Workbook::new();
        let buf = wb.save_to_buffer().unwrap();
        let wb2 = Workbook::open_from_buffer(&buf).unwrap();

        // None of the standard entries should appear in unknown_parts.
        let unknown_paths: Vec<&str> = wb2.unknown_parts.iter().map(|(p, _)| p.as_str()).collect();
        assert!(
            !unknown_paths.contains(&"[Content_Types].xml"),
            "Content_Types should not be in unknown_parts"
        );
        assert!(
            !unknown_paths.contains(&"xl/workbook.xml"),
            "workbook.xml should not be in unknown_parts"
        );
        assert!(
            !unknown_paths.contains(&"xl/styles.xml"),
            "styles.xml should not be in unknown_parts"
        );
    }

    #[test]
    fn test_modifications_preserved_alongside_unknown_parts() {
        let buf = create_xlsx_with_custom_entries();
        let mut wb = Workbook::open_from_buffer(&buf).unwrap();

        // Modify data in the workbook.
        wb.set_cell_value("Sheet1", "B1", CellValue::Number(42.0))
            .unwrap();

        let saved = wb.save_to_buffer().unwrap();
        let wb2 = Workbook::open_from_buffer(&saved).unwrap();

        // Original data preserved.
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("hello".to_string())
        );
        // New data present.
        assert_eq!(
            wb2.get_cell_value("Sheet1", "B1").unwrap(),
            CellValue::Number(42.0)
        );
        // Unknown parts still present.
        let cursor = std::io::Cursor::new(&saved);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        assert!(archive.by_name("customXml/item1.xml").is_ok());
    }

    #[test]
    fn test_threaded_comment_person_rel_in_workbook_rels() {
        let mut wb = Workbook::new();
        wb.add_threaded_comment(
            "Sheet1",
            "A1",
            &crate::threaded_comment::ThreadedCommentInput {
                author: "Alice".to_string(),
                text: "Test comment".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        let buf = wb.save_to_buffer().unwrap();
        let wb2 = Workbook::open_from_buffer(&buf).unwrap();

        // Verify workbook_rels contains a REL_TYPE_PERSON relationship.
        let has_person_rel = wb2.workbook_rels.relationships.iter().any(|r| {
            r.rel_type == sheetkit_xml::threaded_comment::REL_TYPE_PERSON
                && r.target == "persons/person.xml"
        });
        assert!(
            has_person_rel,
            "workbook_rels must contain a person relationship for threaded comments"
        );
    }

    #[test]
    fn test_no_person_rel_without_threaded_comments() {
        let wb = Workbook::new();
        let buf = wb.save_to_buffer().unwrap();
        let wb2 = Workbook::open_from_buffer(&buf).unwrap();

        let has_person_rel = wb2
            .workbook_rels
            .relationships
            .iter()
            .any(|r| r.rel_type == sheetkit_xml::threaded_comment::REL_TYPE_PERSON);
        assert!(
            !has_person_rel,
            "workbook_rels must not contain a person relationship when there are no threaded comments"
        );
    }

    #[cfg(feature = "encryption")]
    #[test]
    fn test_open_encrypted_file_without_password_returns_file_encrypted() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("encrypted2.xlsx");

        let wb = Workbook::new();
        wb.save_with_password(&path, "secret").unwrap();

        let result = Workbook::open(&path);
        assert!(matches!(result, Err(Error::FileEncrypted)))
    }

    #[test]
    fn test_workbook_format_from_content_type() {
        use sheetkit_xml::content_types::mime_types;
        assert_eq!(
            WorkbookFormat::from_content_type(mime_types::WORKBOOK),
            Some(WorkbookFormat::Xlsx)
        );
        assert_eq!(
            WorkbookFormat::from_content_type(mime_types::WORKBOOK_MACRO),
            Some(WorkbookFormat::Xlsm)
        );
        assert_eq!(
            WorkbookFormat::from_content_type(mime_types::WORKBOOK_TEMPLATE),
            Some(WorkbookFormat::Xltx)
        );
        assert_eq!(
            WorkbookFormat::from_content_type(mime_types::WORKBOOK_TEMPLATE_MACRO),
            Some(WorkbookFormat::Xltm)
        );
        assert_eq!(
            WorkbookFormat::from_content_type(mime_types::WORKBOOK_ADDIN_MACRO),
            Some(WorkbookFormat::Xlam)
        );
        assert_eq!(
            WorkbookFormat::from_content_type("application/unknown"),
            None
        );
    }

    #[test]
    fn test_workbook_format_content_type_roundtrip() {
        for fmt in [
            WorkbookFormat::Xlsx,
            WorkbookFormat::Xlsm,
            WorkbookFormat::Xltx,
            WorkbookFormat::Xltm,
            WorkbookFormat::Xlam,
        ] {
            let ct = fmt.content_type();
            assert_eq!(WorkbookFormat::from_content_type(ct), Some(fmt));
        }
    }

    #[test]
    fn test_new_workbook_defaults_to_xlsx_format() {
        let wb = Workbook::new();
        assert_eq!(wb.format(), WorkbookFormat::Xlsx);
    }

    #[test]
    fn test_xlsx_roundtrip_preserves_format() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("roundtrip_format.xlsx");

        let wb = Workbook::new();
        assert_eq!(wb.format(), WorkbookFormat::Xlsx);
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.format(), WorkbookFormat::Xlsx);
    }

    #[test]
    fn test_save_writes_correct_content_type_for_each_extension() {
        let dir = TempDir::new().unwrap();

        let cases = [
            (WorkbookFormat::Xlsx, "test.xlsx"),
            (WorkbookFormat::Xlsm, "test.xlsm"),
            (WorkbookFormat::Xltx, "test.xltx"),
            (WorkbookFormat::Xltm, "test.xltm"),
            (WorkbookFormat::Xlam, "test.xlam"),
        ];

        for (expected_fmt, filename) in cases {
            let path = dir.path().join(filename);
            let wb = Workbook::new();
            wb.save(&path).unwrap();

            let file = std::fs::File::open(&path).unwrap();
            let mut archive = zip::ZipArchive::new(file).unwrap();

            let ct: ContentTypes = read_xml_part(&mut archive, "[Content_Types].xml").unwrap();
            let wb_override = ct
                .overrides
                .iter()
                .find(|o| o.part_name == "/xl/workbook.xml")
                .expect("workbook override must exist");
            assert_eq!(
                wb_override.content_type,
                expected_fmt.content_type(),
                "content type mismatch for {}",
                filename
            );
        }
    }

    #[test]
    fn test_set_format_changes_workbook_format() {
        let mut wb = Workbook::new();
        assert_eq!(wb.format(), WorkbookFormat::Xlsx);

        wb.set_format(WorkbookFormat::Xlsm);
        assert_eq!(wb.format(), WorkbookFormat::Xlsm);
    }

    #[test]
    fn test_save_buffer_roundtrip_with_xlsm_format() {
        let mut wb = Workbook::new();
        wb.set_format(WorkbookFormat::Xlsm);
        wb.set_cell_value("Sheet1", "A1", CellValue::String("test".to_string()))
            .unwrap();

        let buf = wb.save_to_buffer().unwrap();
        let wb2 = Workbook::open_from_buffer(&buf).unwrap();
        assert_eq!(wb2.format(), WorkbookFormat::Xlsm);
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("test".to_string())
        );
    }

    #[test]
    fn test_open_with_default_options_is_equivalent_to_open() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("default_opts.xlsx");
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("test".to_string()))
            .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open_with_options(&path, &OpenOptions::default()).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1"]);
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("test".to_string())
        );
    }

    #[test]
    fn test_format_inference_from_content_types_overrides() {
        use sheetkit_xml::content_types::mime_types;

        // Simulate a content_types with xlsm workbook type.
        let ct = ContentTypes {
            xmlns: "http://schemas.openxmlformats.org/package/2006/content-types".to_string(),
            defaults: vec![],
            overrides: vec![ContentTypeOverride {
                part_name: "/xl/workbook.xml".to_string(),
                content_type: mime_types::WORKBOOK_MACRO.to_string(),
            }],
        };

        let detected = ct
            .overrides
            .iter()
            .find(|o| o.part_name == "/xl/workbook.xml")
            .and_then(|o| WorkbookFormat::from_content_type(&o.content_type))
            .unwrap_or_default();
        assert_eq!(detected, WorkbookFormat::Xlsm);
    }

    #[test]
    fn test_workbook_format_default_is_xlsx() {
        assert_eq!(WorkbookFormat::default(), WorkbookFormat::Xlsx);
    }

    fn build_xlsm_with_vba(vba_bytes: &[u8]) -> Vec<u8> {
        use std::io::Write;
        let mut buf = Vec::new();
        {
            let cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

            let ct_xml = format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="{wb_ct}"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="{ws_ct}"/>
  <Override PartName="/xl/styles.xml" ContentType="{st_ct}"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="{sst_ct}"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</Types>"#,
                wb_ct = mime_types::WORKBOOK_MACRO,
                ws_ct = mime_types::WORKSHEET,
                st_ct = mime_types::STYLES,
                sst_ct = mime_types::SHARED_STRINGS,
            );
            zip.start_file("[Content_Types].xml", opts).unwrap();
            zip.write_all(ct_xml.as_bytes()).unwrap();

            let pkg_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;
            zip.start_file("_rels/.rels", opts).unwrap();
            zip.write_all(pkg_rels.as_bytes()).unwrap();

            let wb_rels = format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{ws_rel}" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{st_rel}" Target="styles.xml"/>
  <Relationship Id="rId3" Type="{sst_rel}" Target="sharedStrings.xml"/>
  <Relationship Id="rId4" Type="{vba_rel}" Target="vbaProject.bin"/>
</Relationships>"#,
                ws_rel = rel_types::WORKSHEET,
                st_rel = rel_types::STYLES,
                sst_rel = rel_types::SHARED_STRINGS,
                vba_rel = VBA_PROJECT_REL_TYPE,
            );
            zip.start_file("xl/_rels/workbook.xml.rels", opts).unwrap();
            zip.write_all(wb_rels.as_bytes()).unwrap();

            let wb_xml = concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#,
                r#" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
                r#"<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>"#,
                r#"</workbook>"#,
            );
            zip.start_file("xl/workbook.xml", opts).unwrap();
            zip.write_all(wb_xml.as_bytes()).unwrap();

            let ws_xml = concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#,
                r#" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
                r#"<sheetData/>"#,
                r#"</worksheet>"#,
            );
            zip.start_file("xl/worksheets/sheet1.xml", opts).unwrap();
            zip.write_all(ws_xml.as_bytes()).unwrap();

            let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>"#;
            zip.start_file("xl/styles.xml", opts).unwrap();
            zip.write_all(styles_xml.as_bytes()).unwrap();

            let sst_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>"#;
            zip.start_file("xl/sharedStrings.xml", opts).unwrap();
            zip.write_all(sst_xml.as_bytes()).unwrap();

            zip.start_file("xl/vbaProject.bin", opts).unwrap();
            zip.write_all(vba_bytes).unwrap();

            zip.finish().unwrap();
        }
        buf
    }

    #[test]
    fn test_vba_blob_loaded_when_present() {
        let vba_data = b"FAKE_VBA_PROJECT_BINARY_DATA_1234567890";
        let xlsm = build_xlsm_with_vba(vba_data);
        let wb = Workbook::open_from_buffer(&xlsm).unwrap();
        assert!(wb.vba_blob.is_some());
        assert_eq!(wb.vba_blob.as_deref().unwrap(), vba_data);
    }

    #[test]
    fn test_vba_blob_none_for_plain_xlsx() {
        let wb = Workbook::new();
        assert!(wb.vba_blob.is_none());

        let buf = wb.save_to_buffer().unwrap();
        let wb2 = Workbook::open_from_buffer(&buf).unwrap();
        assert!(wb2.vba_blob.is_none());
    }

    #[test]
    fn test_vba_blob_survives_roundtrip_with_identical_bytes() {
        let vba_data: Vec<u8> = (0..=255).cycle().take(1024).collect();
        let xlsm = build_xlsm_with_vba(&vba_data);

        let wb = Workbook::open_from_buffer(&xlsm).unwrap();
        assert_eq!(wb.vba_blob.as_deref().unwrap(), &vba_data[..]);

        let saved = wb.save_to_buffer().unwrap();
        let cursor = std::io::Cursor::new(&saved);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let mut roundtripped = Vec::new();
        std::io::Read::read_to_end(
            &mut archive.by_name("xl/vbaProject.bin").unwrap(),
            &mut roundtripped,
        )
        .unwrap();
        assert_eq!(roundtripped, vba_data);
    }

    #[test]
    fn test_vba_relationship_preserved_on_roundtrip() {
        let vba_data = b"VBA_BLOB";
        let xlsm = build_xlsm_with_vba(vba_data);

        let wb = Workbook::open_from_buffer(&xlsm).unwrap();
        let saved = wb.save_to_buffer().unwrap();

        let cursor = std::io::Cursor::new(&saved);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let rels: Relationships =
            read_xml_part(&mut archive, "xl/_rels/workbook.xml.rels").unwrap();
        let vba_rel = rels
            .relationships
            .iter()
            .find(|r| r.rel_type == VBA_PROJECT_REL_TYPE);
        assert!(vba_rel.is_some(), "VBA relationship must be preserved");
        assert_eq!(vba_rel.unwrap().target, "vbaProject.bin");
    }

    #[test]
    fn test_vba_content_type_preserved_on_roundtrip() {
        let vba_data = b"VBA_BLOB";
        let xlsm = build_xlsm_with_vba(vba_data);

        let wb = Workbook::open_from_buffer(&xlsm).unwrap();
        let saved = wb.save_to_buffer().unwrap();

        let cursor = std::io::Cursor::new(&saved);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let ct: ContentTypes = read_xml_part(&mut archive, "[Content_Types].xml").unwrap();
        let vba_override = ct
            .overrides
            .iter()
            .find(|o| o.part_name == "/xl/vbaProject.bin");
        assert!(
            vba_override.is_some(),
            "VBA content type override must be preserved"
        );
        assert_eq!(vba_override.unwrap().content_type, VBA_PROJECT_CONTENT_TYPE);
    }

    #[test]
    fn test_non_vba_save_has_no_vba_entries() {
        let wb = Workbook::new();
        let buf = wb.save_to_buffer().unwrap();

        let cursor = std::io::Cursor::new(&buf);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        assert!(
            archive.by_name("xl/vbaProject.bin").is_err(),
            "plain xlsx must not contain vbaProject.bin"
        );

        let rels: Relationships =
            read_xml_part(&mut archive, "xl/_rels/workbook.xml.rels").unwrap();
        assert!(
            !rels
                .relationships
                .iter()
                .any(|r| r.rel_type == VBA_PROJECT_REL_TYPE),
            "plain xlsx must not have VBA relationship"
        );

        let ct: ContentTypes = read_xml_part(&mut archive, "[Content_Types].xml").unwrap();
        assert!(
            !ct.overrides
                .iter()
                .any(|o| o.content_type == VBA_PROJECT_CONTENT_TYPE),
            "plain xlsx must not have VBA content type override"
        );
    }

    #[test]
    fn test_xlsm_format_detected_with_vba() {
        let vba_data = b"VBA_BLOB";
        let xlsm = build_xlsm_with_vba(vba_data);
        let wb = Workbook::open_from_buffer(&xlsm).unwrap();
        assert_eq!(wb.format(), WorkbookFormat::Xlsm);
    }

    #[test]
    fn test_from_extension_recognized() {
        assert_eq!(
            WorkbookFormat::from_extension("xlsx"),
            Some(WorkbookFormat::Xlsx)
        );
        assert_eq!(
            WorkbookFormat::from_extension("xlsm"),
            Some(WorkbookFormat::Xlsm)
        );
        assert_eq!(
            WorkbookFormat::from_extension("xltx"),
            Some(WorkbookFormat::Xltx)
        );
        assert_eq!(
            WorkbookFormat::from_extension("xltm"),
            Some(WorkbookFormat::Xltm)
        );
        assert_eq!(
            WorkbookFormat::from_extension("xlam"),
            Some(WorkbookFormat::Xlam)
        );
    }

    #[test]
    fn test_from_extension_case_insensitive() {
        assert_eq!(
            WorkbookFormat::from_extension("XLSX"),
            Some(WorkbookFormat::Xlsx)
        );
        assert_eq!(
            WorkbookFormat::from_extension("Xlsm"),
            Some(WorkbookFormat::Xlsm)
        );
        assert_eq!(
            WorkbookFormat::from_extension("XLTX"),
            Some(WorkbookFormat::Xltx)
        );
    }

    #[test]
    fn test_from_extension_unrecognized() {
        assert_eq!(WorkbookFormat::from_extension("csv"), None);
        assert_eq!(WorkbookFormat::from_extension("xls"), None);
        assert_eq!(WorkbookFormat::from_extension("txt"), None);
        assert_eq!(WorkbookFormat::from_extension("pdf"), None);
        assert_eq!(WorkbookFormat::from_extension(""), None);
    }

    #[test]
    fn test_save_unsupported_extension_csv() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("output.csv");
        let wb = Workbook::new();
        let result = wb.save(&path);
        assert!(matches!(result, Err(Error::UnsupportedFileExtension(ext)) if ext == "csv"));
    }

    #[test]
    fn test_save_unsupported_extension_xls() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("output.xls");
        let wb = Workbook::new();
        let result = wb.save(&path);
        assert!(matches!(result, Err(Error::UnsupportedFileExtension(ext)) if ext == "xls"));
    }

    #[test]
    fn test_save_unsupported_extension_unknown() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("output.foo");
        let wb = Workbook::new();
        let result = wb.save(&path);
        assert!(matches!(result, Err(Error::UnsupportedFileExtension(ext)) if ext == "foo"));
    }

    #[test]
    fn test_save_no_extension_fails() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("noext");
        let wb = Workbook::new();
        let result = wb.save(&path);
        assert!(matches!(
            result,
            Err(Error::UnsupportedFileExtension(ext)) if ext.is_empty()
        ));
    }

    #[test]
    fn test_save_as_xlsm_writes_xlsm_content_type() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("output.xlsm");
        let wb = Workbook::new();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        let ct: ContentTypes = read_xml_part(&mut archive, "[Content_Types].xml").unwrap();
        let wb_ct = ct
            .overrides
            .iter()
            .find(|o| o.part_name == "/xl/workbook.xml")
            .expect("workbook override must exist");
        assert_eq!(wb_ct.content_type, WorkbookFormat::Xlsm.content_type());
    }

    #[test]
    fn test_save_as_xltx_writes_template_content_type() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("output.xltx");
        let wb = Workbook::new();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        let ct: ContentTypes = read_xml_part(&mut archive, "[Content_Types].xml").unwrap();
        let wb_ct = ct
            .overrides
            .iter()
            .find(|o| o.part_name == "/xl/workbook.xml")
            .expect("workbook override must exist");
        assert_eq!(wb_ct.content_type, WorkbookFormat::Xltx.content_type());
    }

    #[test]
    fn test_save_as_xltm_writes_template_macro_content_type() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("output.xltm");
        let wb = Workbook::new();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        let ct: ContentTypes = read_xml_part(&mut archive, "[Content_Types].xml").unwrap();
        let wb_ct = ct
            .overrides
            .iter()
            .find(|o| o.part_name == "/xl/workbook.xml")
            .expect("workbook override must exist");
        assert_eq!(wb_ct.content_type, WorkbookFormat::Xltm.content_type());
    }

    #[test]
    fn test_save_as_xlam_writes_addin_content_type() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("output.xlam");
        let wb = Workbook::new();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        let ct: ContentTypes = read_xml_part(&mut archive, "[Content_Types].xml").unwrap();
        let wb_ct = ct
            .overrides
            .iter()
            .find(|o| o.part_name == "/xl/workbook.xml")
            .expect("workbook override must exist");
        assert_eq!(wb_ct.content_type, WorkbookFormat::Xlam.content_type());
    }

    #[test]
    fn test_save_extension_overrides_stored_format() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("output.xlsm");

        // Workbook has Xlsx format stored, but saved as .xlsm
        let wb = Workbook::new();
        assert_eq!(wb.format(), WorkbookFormat::Xlsx);
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        let ct: ContentTypes = read_xml_part(&mut archive, "[Content_Types].xml").unwrap();
        let wb_ct = ct
            .overrides
            .iter()
            .find(|o| o.part_name == "/xl/workbook.xml")
            .expect("workbook override must exist");
        assert_eq!(
            wb_ct.content_type,
            WorkbookFormat::Xlsm.content_type(),
            "extension .xlsm must override stored Xlsx format"
        );
    }

    #[test]
    fn test_save_to_buffer_preserves_stored_format() {
        let mut wb = Workbook::new();
        wb.set_format(WorkbookFormat::Xltx);

        let buf = wb.save_to_buffer().unwrap();
        let wb2 = Workbook::open_from_buffer(&buf).unwrap();
        assert_eq!(
            wb2.format(),
            WorkbookFormat::Xltx,
            "save_to_buffer must use the stored format, not infer from extension"
        );
    }

    #[test]
    fn test_sheet_rows_limits_rows_read() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("sheet_rows.xlsx");

        let mut wb = Workbook::new();
        for i in 1..=20 {
            let cell = format!("A{}", i);
            wb.set_cell_value("Sheet1", &cell, CellValue::Number(i as f64))
                .unwrap();
        }
        wb.save(&path).unwrap();

        let opts = OpenOptions::new().sheet_rows(5);
        let wb2 = Workbook::open_with_options(&path, &opts).unwrap();

        // First 5 rows should be present
        for i in 1..=5 {
            let cell = format!("A{}", i);
            assert_eq!(
                wb2.get_cell_value("Sheet1", &cell).unwrap(),
                CellValue::Number(i as f64)
            );
        }

        // Rows 6+ should return Empty
        for i in 6..=20 {
            let cell = format!("A{}", i);
            assert_eq!(
                wb2.get_cell_value("Sheet1", &cell).unwrap(),
                CellValue::Empty
            );
        }
    }

    #[test]
    fn test_sheet_rows_with_buffer() {
        let mut wb = Workbook::new();
        for i in 1..=10 {
            let cell = format!("A{}", i);
            wb.set_cell_value("Sheet1", &cell, CellValue::Number(i as f64))
                .unwrap();
        }
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().sheet_rows(3);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        assert_eq!(
            wb2.get_cell_value("Sheet1", "A3").unwrap(),
            CellValue::Number(3.0)
        );
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A4").unwrap(),
            CellValue::Empty
        );
    }

    #[test]
    fn test_save_xlsx_preserves_existing_behavior() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("preserved.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("hello".to_string()))
            .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(wb2.format(), WorkbookFormat::Xlsx);
        assert_eq!(wb2.sheet_names(), vec!["Sheet1"]);
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("hello".to_string())
        );
    }

    #[test]
    fn test_selective_sheet_parsing() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("selective.xlsx");

        let mut wb = Workbook::new();
        wb.new_sheet("Sales").unwrap();
        wb.new_sheet("Data").unwrap();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("Sheet1 data".to_string()))
            .unwrap();
        wb.set_cell_value("Sales", "A1", CellValue::String("Sales data".to_string()))
            .unwrap();
        wb.set_cell_value("Data", "A1", CellValue::String("Data data".to_string()))
            .unwrap();
        wb.save(&path).unwrap();

        let opts = OpenOptions::new().sheets(vec!["Sales".to_string()]);
        let wb2 = Workbook::open_with_options(&path, &opts).unwrap();

        // All sheets exist in the workbook
        assert_eq!(wb2.sheet_names(), vec!["Sheet1", "Sales", "Data"]);

        // Only Sales should have data
        assert_eq!(
            wb2.get_cell_value("Sales", "A1").unwrap(),
            CellValue::String("Sales data".to_string())
        );

        // Sheet1 and Data were not parsed, so they should be empty
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::Empty
        );
        assert_eq!(wb2.get_cell_value("Data", "A1").unwrap(), CellValue::Empty);
    }

    #[test]
    fn test_selective_sheets_multiple() {
        let mut wb = Workbook::new();
        wb.new_sheet("Alpha").unwrap();
        wb.new_sheet("Beta").unwrap();
        wb.set_cell_value("Sheet1", "A1", CellValue::Number(1.0))
            .unwrap();
        wb.set_cell_value("Alpha", "A1", CellValue::Number(2.0))
            .unwrap();
        wb.set_cell_value("Beta", "A1", CellValue::Number(3.0))
            .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().sheets(vec!["Sheet1".to_string(), "Beta".to_string()]);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::Number(1.0)
        );
        assert_eq!(wb2.get_cell_value("Alpha", "A1").unwrap(), CellValue::Empty);
        assert_eq!(
            wb2.get_cell_value("Beta", "A1").unwrap(),
            CellValue::Number(3.0)
        );
    }

    #[test]
    fn test_save_does_not_mutate_stored_format() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("test.xlsm");
        let wb = Workbook::new();
        assert_eq!(wb.format(), WorkbookFormat::Xlsx);
        wb.save(&path).unwrap();
        // The save call takes &self, so the stored format is unchanged.
        assert_eq!(wb.format(), WorkbookFormat::Xlsx);
    }

    #[test]
    fn test_max_zip_entries_exceeded() {
        let wb = Workbook::new();
        let buf = wb.save_to_buffer().unwrap();

        // A basic workbook has at least 8 ZIP entries -- set limit to 2
        let opts = OpenOptions::new().max_zip_entries(2);
        let result = Workbook::open_from_buffer_with_options(&buf, &opts);
        assert!(matches!(result, Err(Error::ZipEntryCountExceeded { .. })));
    }

    #[test]
    fn test_max_zip_entries_within_limit() {
        let wb = Workbook::new();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().max_zip_entries(1000);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1"]);
    }

    #[test]
    fn test_max_unzip_size_exceeded() {
        let mut wb = Workbook::new();
        // Write enough data so the decompressed size is non-trivial
        for i in 1..=100 {
            let cell = format!("A{}", i);
            wb.set_cell_value(
                "Sheet1",
                &cell,
                CellValue::String("long_value_for_size_check".repeat(10)),
            )
            .unwrap();
        }
        let buf = wb.save_to_buffer().unwrap();

        // Set a very small decompressed size limit
        let opts = OpenOptions::new().max_unzip_size(100);
        let result = Workbook::open_from_buffer_with_options(&buf, &opts);
        assert!(matches!(result, Err(Error::ZipSizeExceeded { .. })));
    }

    #[test]
    fn test_max_unzip_size_within_limit() {
        let wb = Workbook::new();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().max_unzip_size(1_000_000_000);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1"]);
    }

    #[test]
    fn test_combined_options() {
        let mut wb = Workbook::new();
        wb.new_sheet("Parsed").unwrap();
        wb.new_sheet("Skipped").unwrap();
        for i in 1..=10 {
            let cell = format!("A{}", i);
            wb.set_cell_value("Parsed", &cell, CellValue::Number(i as f64))
                .unwrap();
            wb.set_cell_value("Skipped", &cell, CellValue::Number(i as f64))
                .unwrap();
        }
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new()
            .sheets(vec!["Parsed".to_string()])
            .sheet_rows(3);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Parsed sheet has only 3 rows
        assert_eq!(
            wb2.get_cell_value("Parsed", "A3").unwrap(),
            CellValue::Number(3.0)
        );
        assert_eq!(
            wb2.get_cell_value("Parsed", "A4").unwrap(),
            CellValue::Empty
        );

        // Skipped sheet is empty
        assert_eq!(
            wb2.get_cell_value("Skipped", "A1").unwrap(),
            CellValue::Empty
        );
    }

    #[test]
    fn test_sheet_rows_zero_means_no_rows() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::Number(1.0))
            .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().sheet_rows(0);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::Empty
        );
    }

    #[test]
    fn test_selective_sheet_parsing_preserves_unparsed_sheets_on_save() {
        let dir = TempDir::new().unwrap();
        let path1 = dir.path().join("original.xlsx");
        let path2 = dir.path().join("resaved.xlsx");

        // Create a workbook with 3 sheets, each with distinct data.
        let mut wb = Workbook::new();
        wb.new_sheet("Sales").unwrap();
        wb.new_sheet("Data").unwrap();
        wb.set_cell_value(
            "Sheet1",
            "A1",
            CellValue::String("Sheet1 value".to_string()),
        )
        .unwrap();
        wb.set_cell_value("Sheet1", "B2", CellValue::Number(100.0))
            .unwrap();
        wb.set_cell_value("Sales", "A1", CellValue::String("Sales value".to_string()))
            .unwrap();
        wb.set_cell_value("Sales", "C3", CellValue::Number(200.0))
            .unwrap();
        wb.set_cell_value("Data", "A1", CellValue::String("Data value".to_string()))
            .unwrap();
        wb.set_cell_value("Data", "D4", CellValue::Bool(true))
            .unwrap();
        wb.save(&path1).unwrap();

        // Reopen with only Sheet1 parsed.
        let opts = OpenOptions::new().sheets(vec!["Sheet1".to_string()]);
        let wb2 = Workbook::open_with_options(&path1, &opts).unwrap();

        // Verify Sheet1 was parsed.
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Sheet1 value".to_string())
        );

        // Save to a new file.
        wb2.save(&path2).unwrap();

        // Reopen the resaved file with all sheets parsed.
        let wb3 = Workbook::open(&path2).unwrap();
        assert_eq!(wb3.sheet_names(), vec!["Sheet1", "Sales", "Data"]);

        // Sheet1 data should be intact.
        assert_eq!(
            wb3.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Sheet1 value".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Sheet1", "B2").unwrap(),
            CellValue::Number(100.0)
        );

        // Sales data should be preserved from raw XML.
        assert_eq!(
            wb3.get_cell_value("Sales", "A1").unwrap(),
            CellValue::String("Sales value".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Sales", "C3").unwrap(),
            CellValue::Number(200.0)
        );

        // Data sheet should be preserved from raw XML.
        assert_eq!(
            wb3.get_cell_value("Data", "A1").unwrap(),
            CellValue::String("Data value".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Data", "D4").unwrap(),
            CellValue::Bool(true)
        );
    }

    #[test]
    fn test_open_from_buffer_with_options_backwards_compatible() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".to_string()))
            .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let wb2 = Workbook::open_from_buffer(&buf).unwrap();
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Hello".to_string())
        );
    }

    use crate::workbook::open_options::ParseMode;

    #[test]
    fn test_readfast_open_reads_cell_data() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("Hello".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet1", "B2", CellValue::Number(42.0))
            .unwrap();
        wb.set_cell_value("Sheet1", "C3", CellValue::Bool(true))
            .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().parse_mode(ParseMode::ReadFast);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1"]);
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Hello".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Sheet1", "B2").unwrap(),
            CellValue::Number(42.0)
        );
        assert_eq!(
            wb2.get_cell_value("Sheet1", "C3").unwrap(),
            CellValue::Bool(true)
        );
    }

    #[test]
    fn test_readfast_open_multi_sheet() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("S1".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet2", "A1", CellValue::String("S2".to_string()))
            .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().parse_mode(ParseMode::ReadFast);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1", "Sheet2"]);
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("S1".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Sheet2", "A1").unwrap(),
            CellValue::String("S2".to_string())
        );
    }

    #[test]
    fn test_readfast_skips_comments() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("data".to_string()))
            .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Tester".to_string(),
                text: "A test comment".to_string(),
            },
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().parse_mode(ParseMode::ReadFast);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Cell data is readable.
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("data".to_string())
        );
        // Comments are not loaded in ReadFast mode.
        let comments = wb2.get_comments("Sheet1").unwrap();
        assert!(comments.is_empty());
    }

    #[test]
    fn test_readfast_skips_doc_properties() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::Number(1.0))
            .unwrap();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Test Title".to_string()),
            ..Default::default()
        });
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().parse_mode(ParseMode::ReadFast);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Cell data is readable.
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::Number(1.0)
        );
        // Doc properties are not loaded.
        let props = wb2.get_doc_props();
        assert!(props.title.is_none());
    }

    #[test]
    fn test_readfast_save_roundtrip_preserves_all_parts() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("data".to_string()))
            .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Tester".to_string(),
                text: "A comment".to_string(),
            },
        )
        .unwrap();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Title".to_string()),
            ..Default::default()
        });
        let buf = wb.save_to_buffer().unwrap();

        // Open in ReadFast mode.
        let opts = OpenOptions::new().parse_mode(ParseMode::ReadFast);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        // Re-open in Full mode and verify all parts were preserved.
        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        assert_eq!(
            wb3.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("data".to_string())
        );
        let comments = wb3.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].text, "A comment");
        let props = wb3.get_doc_props();
        assert_eq!(props.title, Some("Title".to_string()));
    }

    #[test]
    fn test_readfast_with_sheet_rows_limit() {
        let mut wb = Workbook::new();
        for i in 1..=100 {
            wb.set_cell_value("Sheet1", &format!("A{}", i), CellValue::Number(i as f64))
                .unwrap();
        }
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new()
            .parse_mode(ParseMode::ReadFast)
            .sheet_rows(10);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let rows = wb2.get_rows("Sheet1").unwrap();
        assert_eq!(rows.len(), 10);
    }

    #[test]
    fn test_readfast_with_sheets_filter() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("S1".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet2", "A1", CellValue::String("S2".to_string()))
            .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new()
            .parse_mode(ParseMode::ReadFast)
            .sheets(vec!["Sheet2".to_string()]);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert_eq!(wb2.sheet_names(), vec!["Sheet1", "Sheet2"]);
        assert_eq!(
            wb2.get_cell_value("Sheet2", "A1").unwrap(),
            CellValue::String("S2".to_string())
        );
        // Sheet1 was not parsed, should return empty.
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::Empty
        );
    }

    #[test]
    fn test_readfast_preserves_styles() {
        let mut wb = Workbook::new();
        let style_id = wb
            .add_style(&crate::style::Style {
                font: Some(crate::style::FontStyle {
                    bold: true,
                    ..Default::default()
                }),
                ..Default::default()
            })
            .unwrap();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("bold".to_string()))
            .unwrap();
        wb.set_cell_style("Sheet1", "A1", style_id).unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().parse_mode(ParseMode::ReadFast);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let sid = wb2.get_cell_style("Sheet1", "A1").unwrap();
        assert!(sid.is_some());
        let style = crate::style::get_style(&wb2.stylesheet, sid.unwrap());
        assert!(style.is_some());
        assert!(style.unwrap().font.map_or(false, |f| f.bold));
    }

    #[test]
    fn test_readfast_full_mode_unchanged() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("test".to_string()))
            .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "comment text".to_string(),
            },
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        // Full mode: everything should be parsed.
        let opts = OpenOptions::new().parse_mode(ParseMode::Full);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let comments = wb2.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
    }

    #[test]
    fn test_readfast_open_from_file() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("readfast_test.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("file test".to_string()))
            .unwrap();
        wb.save(&path).unwrap();

        let opts = OpenOptions::new().parse_mode(ParseMode::ReadFast);
        let wb2 = Workbook::open_with_options(&path, &opts).unwrap();
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("file test".to_string())
        );
    }

    #[test]
    fn test_readfast_roundtrip_with_custom_zip_entries() {
        let buf = create_xlsx_with_custom_entries();

        let opts = OpenOptions::new().parse_mode(ParseMode::ReadFast);
        let wb = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("hello".to_string())
        );

        let saved = wb.save_to_buffer().unwrap();
        let cursor = std::io::Cursor::new(&saved);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        // Custom entries should be preserved through ReadFast open/save.
        let mut custom_xml = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("customXml/item1.xml").unwrap(),
            &mut custom_xml,
        )
        .unwrap();
        assert_eq!(custom_xml, "<custom>data1</custom>");

        let mut printer = Vec::new();
        std::io::Read::read_to_end(
            &mut archive
                .by_name("xl/printerSettings/printerSettings1.bin")
                .unwrap(),
            &mut printer,
        )
        .unwrap();
        assert_eq!(printer, b"\x00\x01\x02\x03PRINTER");
    }

    #[test]
    fn test_readfast_deferred_parts_not_empty_when_auxiliary_exist() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("data".to_string()))
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
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().parse_mode(ParseMode::ReadFast);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        // When auxiliary parts exist, they should be captured in deferred_parts.
        assert!(
            !wb2.deferred_parts.is_empty(),
            "deferred_parts should contain skipped auxiliary parts"
        );
    }

    #[test]
    fn test_readfast_default_mode_has_no_deferred_parts() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("data".to_string()))
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
        let buf = wb.save_to_buffer().unwrap();

        // Full mode: deferred_parts should be empty.
        let wb2 = Workbook::open_from_buffer(&buf).unwrap();
        assert!(
            wb2.deferred_parts.is_empty(),
            "Full mode should not have deferred parts"
        );
    }
}
