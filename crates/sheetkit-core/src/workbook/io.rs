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
        let workbook_xml = WorkbookXml {
            // Excel-like workbook.xml defaults for interoperability.
            file_version: Some(sheetkit_xml::workbook::FileVersion {
                app_name: Some("xl".to_string()),
                last_edited: Some("7".to_string()),
                lowest_edited: Some("7".to_string()),
                rup_build: Some("27425".to_string()),
            }),
            // Excel-like workbookPr default.
            workbook_pr: Some(sheetkit_xml::workbook::WorkbookPr {
                date1904: None,
                date_compatibility: None,
                filter_privacy: None,
                default_theme_version: Some(166925),
                show_objects: None,
                backup_file: None,
                code_name: None,
                check_compatibility: None,
                auto_compress_pictures: None,
                refresh_all_connections: None,
                save_external_link_values: None,
                update_links: None,
                hide_pivot_field_list: None,
                show_pivot_chart_filter: None,
                allow_refresh_query: None,
                publish_items: None,
                show_border_unselected_tables: None,
                prompted_solutions: None,
                show_ink_annotation: None,
            }),
            // Minimal book view default (active tab only).
            book_views: Some(sheetkit_xml::workbook::BookViews {
                workbook_views: vec![sheetkit_xml::workbook::WorkbookView {
                    x_window: None,
                    y_window: None,
                    window_width: None,
                    window_height: None,
                    active_tab: Some(0),
                }],
            }),
            ..WorkbookXml::default()
        };

        let sst_runtime = SharedStringTable::new();
        let mut sheet_name_index = HashMap::new();
        sheet_name_index.insert("Sheet1".to_string(), 0);
        Self {
            format: WorkbookFormat::default(),
            content_types: ContentTypes::default(),
            package_rels: relationships::package_rels(),
            workbook_xml,
            raw_workbook_xml: None,
            workbook_xml_baseline: None,
            workbook_rels: relationships::workbook_rels(),
            worksheets: vec![(
                "Sheet1".to_string(),
                initialized_lock(WorksheetXml::default()),
            )],
            stylesheet: StyleSheet::default(),
            sst_runtime,
            sheet_comments: vec![None],
            charts: vec![],
            raw_charts: vec![],
            drawings: vec![],
            raw_graph_parts: HashMap::new(),
            dirty_graph_parts: HashSet::new(),
            images: vec![],
            worksheet_drawings: HashMap::new(),
            worksheet_rels: HashMap::new(),
            drawing_rels: HashMap::new(),
            core_properties: None,
            app_properties: None,
            custom_properties: None,
            raw_doc_props: HashMap::new(),
            dirty_doc_props: HashSet::new(),
            pivot_tables: vec![],
            pivot_cache_defs: vec![],
            pivot_cache_records: vec![],
            theme_xml: None,
            theme_colors: crate::theme::default_theme_colors(),
            sheet_name_index,
            sheet_sparklines: vec![vec![]],
            sheet_vml: vec![None],
            unknown_parts: vec![],
            deferred_parts: crate::workbook::aux::DeferredAuxParts::new(),
            vba_blob: None,
            tables: vec![],
            raw_sheet_xml: vec![None],
            sheet_dirty: vec![true],
            slicer_defs: vec![],
            slicer_caches: vec![],
            sheet_threaded_comments: vec![None],
            person_list: sheetkit_xml::threaded_comment::PersonList::default(),
            sheet_form_controls: vec![vec![]],
            streamed_sheets: HashMap::new(),
            package_source: None,
            read_mode: ReadMode::default(),
            sheet_rows_limit: None,
            date_interpretation: super::DateInterpretation::default(),
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
    ///
    /// The file is opened directly via `std::fs::File` and the ZIP archive
    /// is read from the file handle, avoiding a full `std::fs::read` copy.
    pub fn open_with_options<P: AsRef<Path>>(path: P, options: &OpenOptions) -> Result<Self> {
        let file_path = path.as_ref();

        // Detect encrypted files (CFB container) by reading the magic bytes.
        #[cfg(feature = "encryption")]
        {
            let mut header = [0u8; 8];
            if let Ok(mut f) = std::fs::File::open(file_path) {
                use std::io::Read as _;
                if f.read_exact(&mut header).is_ok() {
                    if let Ok(crate::crypt::ContainerFormat::Cfb) =
                        crate::crypt::detect_container_format(&header)
                    {
                        return Err(Error::FileEncrypted);
                    }
                }
            }
        }

        let file = std::fs::File::open(file_path)?;
        let mut archive = zip::ZipArchive::new(file).map_err(|e| Error::Zip(e.to_string()))?;
        let mut wb = Self::from_archive(&mut archive, options)?;
        wb.package_source = Some(PackageSource::Path(file_path.to_path_buf()));
        wb.read_mode = options.read_mode;
        Ok(wb)
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

        // Parse xl/workbook.xml while retaining exact source bytes for
        // passthrough when its typed owner remains unchanged.
        let raw_workbook_xml = read_bytes_part(archive, "xl/workbook.xml")?;
        let workbook_xml: WorkbookXml = deserialize_xml_bytes(&raw_workbook_xml)?;
        let workbook_xml_baseline = workbook_xml.clone();
        known_paths.insert("xl/workbook.xml".to_string());

        // Parse xl/_rels/workbook.xml.rels
        let workbook_rels: Relationships = read_xml_part(archive, "xl/_rels/workbook.xml.rels")?;
        known_paths.insert("xl/_rels/workbook.xml.rels".to_string());

        // Parse each worksheet referenced in the workbook.
        let sheet_count = workbook_xml.sheets.sheets.len();
        let mut worksheets: Vec<(String, OnceLock<WorksheetXml>)> = Vec::with_capacity(sheet_count);
        let mut worksheet_paths = Vec::with_capacity(sheet_count);
        let mut raw_sheet_xml: Vec<Option<Vec<u8>>> = Vec::with_capacity(sheet_count);

        let defer_sheets = matches!(options.read_mode, ReadMode::Lazy | ReadMode::Stream);

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

            let should_parse = options.should_parse_sheet(&sheet_entry.name);

            if should_parse && !defer_sheets {
                // Eager mode + selected: parse immediately while retaining
                // exact source bytes for clean-sheet passthrough.
                let raw_bytes = read_bytes_part(archive, &sheet_path)?;
                let ws = deserialize_worksheet_xml(&raw_bytes)?;
                worksheets.push((sheet_entry.name.clone(), initialized_lock(ws)));
                raw_sheet_xml.push(Some(raw_bytes));
            } else if !should_parse {
                // Filtered out (any mode): store raw bytes for round-trip save
                // but initialize the OnceLock with an empty worksheet so that
                // cell queries return Empty instead of hydrating real data.
                let raw_bytes = read_bytes_part(archive, &sheet_path)?;
                worksheets.push((
                    sheet_entry.name.clone(),
                    initialized_lock(WorksheetXml::default()),
                ));
                raw_sheet_xml.push(Some(raw_bytes));
            } else {
                // Lazy/Stream mode + selected: store raw bytes for on-demand
                // hydration. OnceLock is left empty; `worksheet_ref` will
                // parse from `raw_sheet_xml` on first access.
                let raw_bytes = read_bytes_part(archive, &sheet_path)?;
                worksheets.push((sheet_entry.name.clone(), OnceLock::new()));
                raw_sheet_xml.push(Some(raw_bytes));
            }
            known_paths.insert(sheet_path.clone());
            worksheet_paths.push(sheet_path);
        }

        // Parse xl/styles.xml
        let stylesheet: StyleSheet = read_xml_part(archive, "xl/styles.xml")?;
        known_paths.insert("xl/styles.xml".to_string());

        // Parse xl/sharedStrings.xml (optional -- may not exist for workbooks with no strings).
        let shared_strings = match read_shared_strings_part(archive, "xl/sharedStrings.xml")? {
            Some(sst) => {
                known_paths.insert("xl/sharedStrings.xml".to_string());
                sst
            }
            None => Sst::default(),
        };

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

        let skip_aux = options.skip_aux_parts();

        // Auxiliary part parsing: skipped in Lazy/Stream mode.
        let mut sheet_comments: Vec<Option<Comments>> = vec![None; worksheets.len()];
        let mut sheet_vml: Vec<Option<Vec<u8>>> = vec![None; worksheets.len()];
        let mut drawings: Vec<(String, WsDr)> = Vec::new();
        let mut worksheet_drawings: HashMap<usize, usize> = HashMap::new();
        let mut drawing_rels: HashMap<usize, Relationships> = HashMap::new();
        let mut charts: Vec<(String, ChartSpace)> = Vec::new();
        let mut raw_charts: Vec<(String, Vec<u8>)> = Vec::new();
        let mut raw_graph_parts: HashMap<String, Vec<u8>> = HashMap::new();
        let mut images: Vec<(String, Vec<u8>)> = Vec::new();
        let mut core_properties: Option<sheetkit_xml::doc_props::CoreProperties> = None;
        let mut app_properties: Option<sheetkit_xml::doc_props::ExtendedProperties> = None;
        let mut custom_properties: Option<sheetkit_xml::doc_props::CustomProperties> = None;
        let mut raw_doc_props: HashMap<String, Vec<u8>> = HashMap::new();
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

        if !skip_aux {
            let mut drawing_path_to_idx: HashMap<String, usize> = HashMap::new();
            let mut drawing_paths: HashSet<String> = HashSet::new();

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
                    drawing_paths.insert(drawing_path.clone());
                    let drawing_idx = if let Some(idx) = drawing_path_to_idx.get(&drawing_path) {
                        *idx
                    } else {
                        let Ok(bytes) = read_bytes_part(archive, &drawing_path) else {
                            continue;
                        };
                        known_paths.insert(drawing_path.clone());
                        raw_graph_parts.insert(drawing_path.clone(), bytes.clone());
                        let Ok(drawing) = deserialize_xml_bytes::<WsDr>(&bytes) else {
                            continue;
                        };
                        let idx = drawings.len();
                        drawings.push((drawing_path.clone(), drawing));
                        drawing_path_to_idx.insert(drawing_path, idx);
                        idx
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
                if !drawing_paths.insert(drawing_path.clone()) {
                    continue;
                }
                if let Ok(bytes) = read_bytes_part(archive, &drawing_path) {
                    known_paths.insert(drawing_path.clone());
                    raw_graph_parts.insert(drawing_path.clone(), bytes.clone());
                    let Ok(drawing) = deserialize_xml_bytes::<WsDr>(&bytes) else {
                        continue;
                    };
                    let idx = drawings.len();
                    drawings.push((drawing_path.clone(), drawing));
                    drawing_path_to_idx.insert(drawing_path, idx);
                }
            }

            let mut seen_chart_paths: HashSet<String> = HashSet::new();
            let mut seen_image_paths: HashSet<String> = HashSet::new();

            for drawing_path in drawing_paths {
                let drawing_rels_path = relationship_part_path(&drawing_path);
                let Ok(rels_bytes) = read_bytes_part(archive, &drawing_rels_path) else {
                    continue;
                };
                known_paths.insert(drawing_rels_path.clone());
                raw_graph_parts.insert(drawing_rels_path, rels_bytes.clone());
                let Ok(rels) = deserialize_xml_bytes::<Relationships>(&rels_bytes) else {
                    continue;
                };
                let Some(&drawing_idx) = drawing_path_to_idx.get(&drawing_path) else {
                    continue;
                };

                for rel in &rels.relationships {
                    if rel.rel_type == rel_types::CHART {
                        let chart_path = resolve_relationship_target(&drawing_path, &rel.target);
                        if seen_chart_paths.insert(chart_path.clone()) {
                            let Ok(bytes) = read_bytes_part(archive, &chart_path) else {
                                continue;
                            };
                            known_paths.insert(chart_path.clone());
                            match deserialize_xml_bytes::<ChartSpace>(&bytes) {
                                Ok(chart) => {
                                    raw_graph_parts.insert(chart_path.clone(), bytes);
                                    charts.push((chart_path, chart));
                                }
                                Err(_) => {
                                    raw_charts.push((chart_path, bytes));
                                }
                            }
                        }
                    } else if rel.rel_type == rel_types::IMAGE {
                        let image_path = resolve_relationship_target(&drawing_path, &rel.target);
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
                    let Ok(bytes) = read_bytes_part(archive, &chart_path) else {
                        continue;
                    };
                    known_paths.insert(chart_path.clone());
                    match deserialize_xml_bytes::<ChartSpace>(&bytes) {
                        Ok(chart) => {
                            raw_graph_parts.insert(chart_path.clone(), bytes);
                            charts.push((chart_path, chart));
                        }
                        Err(_) => {
                            raw_charts.push((chart_path, bytes));
                        }
                    }
                }
            }

            // Parse each document-property part independently while retaining
            // its original bytes, including malformed or unsupported content.
            if let Ok(bytes) = read_bytes_part(archive, "docProps/core.xml") {
                let xml = String::from_utf8_lossy(&bytes);
                core_properties = sheetkit_xml::doc_props::deserialize_core_properties(&xml).ok();
                raw_doc_props.insert("docProps/core.xml".to_string(), bytes);
                known_paths.insert("docProps/core.xml".to_string());
            }
            if let Ok(bytes) = read_bytes_part(archive, "docProps/app.xml") {
                app_properties = deserialize_xml_bytes(&bytes).ok();
                raw_doc_props.insert("docProps/app.xml".to_string(), bytes);
                known_paths.insert("docProps/app.xml".to_string());
            }
            if let Ok(bytes) = read_bytes_part(archive, "docProps/custom.xml") {
                let xml = String::from_utf8_lossy(&bytes);
                custom_properties =
                    sheetkit_xml::doc_props::deserialize_custom_properties(&xml).ok();
                raw_doc_props.insert("docProps/custom.xml".to_string(), bytes);
                known_paths.insert("docProps/custom.xml".to_string());
            }

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
                        known_paths.insert(path.to_string());
                        slicer_defs.push((path.to_string(), sd));
                    }
                } else if ovr.content_type == mime_types::SLICER_CACHE {
                    if let Ok(raw) = read_string_part(archive, path) {
                        if let Some(scd) = sheetkit_xml::slicer::parse_slicer_cache(&raw) {
                            known_paths.insert(path.to_string());
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

        // Collect remaining ZIP entries. In Lazy/Stream mode, unhandled entries
        // go into deferred_parts (typed index); in Eager mode, they go into
        // unknown_parts for round-trip preservation.
        let mut unknown_parts: Vec<(String, Vec<u8>)> = Vec::new();
        let mut deferred_parts = crate::workbook::aux::DeferredAuxParts::new();
        for i in 0..archive.len() {
            let Ok(entry) = archive.by_index(i) else {
                continue;
            };
            let name = entry.name().to_string();
            drop(entry);
            if !known_paths.contains(&name) {
                if let Ok(bytes) = read_bytes_part(archive, &name) {
                    if skip_aux && crate::workbook::aux::classify_deferred_path(&name).is_some() {
                        deferred_parts.insert(name, bytes);
                    } else {
                        unknown_parts.push((name, bytes));
                    }
                }
            }
        }

        // Populate cached column numbers on all eagerly-parsed cells, apply
        // row limit, and ensure sorted order for binary search correctness.
        // Deferred sheets (empty OnceLock) are skipped here; they are
        // post-processed on demand in `deserialize_worksheet_xml`.
        for (_name, ws_lock) in &mut worksheets {
            let Some(ws) = ws_lock.get_mut() else {
                continue;
            };
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
            raw_workbook_xml: Some(raw_workbook_xml),
            workbook_xml_baseline: Some(workbook_xml_baseline),
            workbook_rels,
            worksheets,
            stylesheet,
            sst_runtime,
            sheet_comments,
            charts,
            raw_charts,
            drawings,
            raw_graph_parts,
            dirty_graph_parts: HashSet::new(),
            images,
            worksheet_drawings,
            worksheet_rels,
            drawing_rels,
            core_properties,
            app_properties,
            custom_properties,
            raw_doc_props,
            dirty_doc_props: HashSet::new(),
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
            // Every opened sheet retains raw bytes and starts clean. A later
            // owner mutation marks only its sheet dirty.
            sheet_dirty: raw_sheet_xml.iter().map(|raw| raw.is_none()).collect(),
            raw_sheet_xml,
            slicer_defs,
            slicer_caches,
            sheet_threaded_comments,
            person_list,
            sheet_form_controls,
            streamed_sheets: HashMap::new(),
            package_source: None,
            read_mode: options.read_mode,
            sheet_rows_limit: options.sheet_rows,
            date_interpretation: options.date_interpretation,
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
        let mut wb = Self::from_archive(&mut archive, options)?;
        wb.package_source = Some(PackageSource::Buffer(data.into()));
        wb.read_mode = options.read_mode;
        Ok(wb)
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
        let has_deferred = self.deferred_parts.has_any();
        let mut workbook_rels = self.workbook_rels.clone();

        // The SST is always emitted, so normalize its package metadata in the
        // save-local clones. This also makes repeated open/save cycles idempotent.
        content_types
            .overrides
            .retain(|entry| entry.part_name != "/xl/sharedStrings.xml");
        content_types.overrides.push(ContentTypeOverride {
            part_name: "/xl/sharedStrings.xml".to_string(),
            content_type: mime_types::SHARED_STRINGS.to_string(),
        });
        workbook_rels.relationships.retain(|relationship| {
            relationship.rel_type != rel_types::SHARED_STRINGS
                && relationship.target != "sharedStrings.xml"
        });
        workbook_rels.relationships.push(Relationship {
            id: crate::sheet::next_rid(&workbook_rels.relationships),
            rel_type: rel_types::SHARED_STRINGS.to_string(),
            target: "sharedStrings.xml".to_string(),
            target_mode: None,
        });
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

        if !self
            .deferred_parts
            .has_category(crate::workbook::aux::AuxCategory::Comments)
        {
            content_types
                .overrides
                .retain(|entry| entry.content_type != mime_types::COMMENTS);
        }
        if !self
            .deferred_parts
            .has_category(crate::workbook::aux::AuxCategory::Vml)
        {
            content_types
                .overrides
                .retain(|entry| entry.content_type != mime_types::VML_DRAWING);
        }

        // When deferred_parts is non-empty (Lazy open), skip comment/VML
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

            // When deferred_parts is non-empty (Lazy open), skip comment/VML
            // synchronization only for sheets whose comment data is still deferred
            // (not yet hydrated). Hydrated sheets need normal relationship sync.
            if has_deferred && !has_comments && !has_form_controls && !has_preserved_vml {
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

            // Regenerate only owned shapes while preserving unrelated VML.
            let vml_path = format!("xl/drawings/vmlDrawing{}.vml", sheet_idx + 1);
            let mut vml_bytes = self
                .sheet_vml
                .get(sheet_idx)
                .and_then(|v| v.as_ref())
                .cloned();

            if has_comments
                && !vml_bytes
                    .as_deref()
                    .is_some_and(|vml| String::from_utf8_lossy(vml).contains("ObjectType=\"Note\""))
            {
                let comments = self.sheet_comments[sheet_idx].as_ref().unwrap();
                let cells: Vec<&str> = comments
                    .comment_list
                    .comments
                    .iter()
                    .map(|comment| comment.r#ref.as_str())
                    .collect();
                vml_bytes = Some(match vml_bytes {
                    Some(existing) => {
                        let start_id = 1025 + crate::control::count_vml_shapes(&existing);
                        crate::vml::merge_vml_comments(&existing, &cells, start_id)
                    }
                    None => crate::vml::build_vml_drawing(&cells).into_bytes(),
                });
            }

            if has_form_controls {
                let controls = &self.sheet_form_controls[sheet_idx];
                vml_bytes = Some(match vml_bytes {
                    Some(existing) => {
                        let start_id = 1025 + crate::control::count_vml_shapes(&existing);
                        crate::control::merge_vml_controls(&existing, controls, start_id)
                    }
                    None => crate::control::build_form_control_vml(controls, 1025).into_bytes(),
                });
            }

            let Some(vml_bytes) = vml_bytes else {
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
        // In Lazy mode, untouched deferred tables should remain pass-through.
        // Once table data is mutated (or new live tables exist), we fully
        // resynchronize worksheet rels/content types/tableParts.
        use crate::workbook::aux::AuxCategory;
        let mut table_parts_by_sheet: HashMap<usize, Vec<String>> = HashMap::new();
        let should_sync_tables = !has_deferred
            || self.deferred_parts.is_dirty(AuxCategory::Tables)
            || !self.tables.is_empty();
        if should_sync_tables {
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

        let has_deferred_threaded = self
            .deferred_parts
            .has_category(crate::workbook::aux::AuxCategory::ThreadedComments)
            || self
                .deferred_parts
                .has_category(crate::workbook::aux::AuxCategory::PersonList);
        if !has_deferred_threaded {
            // Threaded-comment targets are sheet-index based, so rebuild the entire
            // relationship and override set after sheet lifecycle changes.
            content_types.overrides.retain(|override_entry| {
                override_entry.content_type
                    != sheetkit_xml::threaded_comment::THREADED_COMMENTS_CONTENT_TYPE
                    && override_entry.content_type
                        != sheetkit_xml::threaded_comment::PERSON_LIST_CONTENT_TYPE
            });
            for rels in worksheet_rels.values_mut() {
                rels.relationships.retain(|relationship| {
                    relationship.rel_type
                        != sheetkit_xml::threaded_comment::REL_TYPE_THREADED_COMMENT
                });
            }
            workbook_rels.relationships.retain(|relationship| {
                relationship.rel_type != sheetkit_xml::threaded_comment::REL_TYPE_PERSON
            });
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
            let rid = crate::sheet::next_rid(&workbook_rels.relationships);
            workbook_rels.relationships.push(Relationship {
                id: rid,
                rel_type: sheetkit_xml::threaded_comment::REL_TYPE_PERSON.to_string(),
                target: "persons/person.xml".to_string(),
                target_mode: None,
            });
        }

        let raw_preserved_paths: HashSet<String> = self
            .unknown_parts
            .iter()
            .map(|(path, _)| path.clone())
            .chain(
                self.deferred_parts
                    .remaining_parts()
                    .map(|(path, _)| path.to_string()),
            )
            .collect();
        let mut generated_lifecycle_paths: HashSet<String> = HashSet::from([
            "[Content_Types].xml".to_string(),
            "_rels/.rels".to_string(),
            "xl/workbook.xml".to_string(),
            "xl/_rels/workbook.xml.rels".to_string(),
            "xl/styles.xml".to_string(),
            "xl/sharedStrings.xml".to_string(),
        ]);
        generated_lifecycle_paths
            .extend((0..self.worksheets.len()).map(|i| self.sheet_part_path(i)));
        generated_lifecycle_paths.extend(
            self.sheet_comments
                .iter()
                .enumerate()
                .filter(|(_, comments)| comments.is_some())
                .map(|(i, _)| format!("xl/comments{}.xml", i + 1)),
        );
        generated_lifecycle_paths
            .extend(vml_parts_to_write.iter().map(|(_, path, _)| path.clone()));
        generated_lifecycle_paths.extend(self.drawings.iter().map(|(path, _)| path.clone()));
        generated_lifecycle_paths.extend(self.charts.iter().map(|(path, _)| path.clone()));
        generated_lifecycle_paths.extend(self.raw_charts.iter().map(|(path, _)| path.clone()));
        generated_lifecycle_paths.extend(self.raw_graph_parts.keys().cloned());
        generated_lifecycle_paths.extend(self.images.iter().map(|(path, _)| path.clone()));
        generated_lifecycle_paths.extend(self.tables.iter().map(|(path, _, _)| path.clone()));
        generated_lifecycle_paths.extend(self.pivot_tables.iter().map(|(path, _)| path.clone()));
        generated_lifecycle_paths
            .extend(self.pivot_cache_defs.iter().map(|(path, _)| path.clone()));
        generated_lifecycle_paths.extend(
            self.pivot_cache_records
                .iter()
                .map(|(path, _)| path.clone()),
        );
        generated_lifecycle_paths.extend(self.slicer_defs.iter().map(|(path, _)| path.clone()));
        generated_lifecycle_paths.extend(self.slicer_caches.iter().map(|(path, _)| path.clone()));
        generated_lifecycle_paths.extend(
            worksheet_rels
                .keys()
                .map(|sheet_idx| relationship_part_path(&self.sheet_part_path(*sheet_idx))),
        );
        generated_lifecycle_paths.extend(self.drawing_rels.keys().filter_map(|drawing_idx| {
            self.drawings
                .get(*drawing_idx)
                .map(|(path, _)| relationship_part_path(path))
        }));
        if self.theme_xml.is_some() {
            generated_lifecycle_paths.insert("xl/theme/theme1.xml".to_string());
        }
        if self.vba_blob.is_some() {
            generated_lifecycle_paths.insert("xl/vbaProject.bin".to_string());
        }
        if self.core_properties.is_some() {
            generated_lifecycle_paths.insert("docProps/core.xml".to_string());
        }
        if self.app_properties.is_some() {
            generated_lifecycle_paths.insert("docProps/app.xml".to_string());
        }
        if self.custom_properties.is_some() {
            generated_lifecycle_paths.insert("docProps/custom.xml".to_string());
        }
        generated_lifecycle_paths.extend(self.raw_doc_props.keys().cloned());
        if has_any_threaded {
            generated_lifecycle_paths.extend(
                self.sheet_threaded_comments
                    .iter()
                    .enumerate()
                    .filter(|(_, threaded)| threaded.is_some())
                    .map(|(i, _)| format!("xl/threadedComments/threadedComment{}.xml", i + 1)),
            );
            generated_lifecycle_paths.insert("xl/persons/person.xml".to_string());
        }
        if let Some(conflict) = generated_lifecycle_paths
            .intersection(&raw_preserved_paths)
            .next()
        {
            return Err(Error::InvalidArgument(format!(
                "cannot save because owned part path '{conflict}' conflicts with preserved raw data"
            )));
        }

        // [Content_Types].xml
        write_xml_part(zip, "[Content_Types].xml", &content_types, options)?;

        // _rels/.rels
        write_xml_part(zip, "_rels/.rels", &self.package_rels, options)?;

        // xl/workbook.xml
        let slicer_cache_rids: Vec<&str> = workbook_rels
            .relationships
            .iter()
            .filter(|rel| rel.rel_type == rel_types::SLICER_CACHE)
            .map(|rel| rel.id.as_str())
            .collect();
        let workbook_needs_slicer_extensions = !slicer_cache_rids.is_empty();
        if workbook_needs_slicer_extensions {
            let raw = self.raw_workbook_xml.as_deref();
            let xml = serialize_workbook_with_slicer_extensions(
                &self.workbook_xml,
                raw,
                &slicer_cache_rids,
            )?;
            write_bytes_part(zip, "xl/workbook.xml", xml.as_bytes(), options)?;
        } else if let (Some(raw), Some(baseline)) =
            (&self.raw_workbook_xml, &self.workbook_xml_baseline)
        {
            if self.workbook_xml == *baseline {
                write_bytes_part(zip, "xl/workbook.xml", raw, options)?;
            } else {
                write_xml_part(zip, "xl/workbook.xml", &self.workbook_xml, options)?;
            }
        } else {
            write_xml_part(zip, "xl/workbook.xml", &self.workbook_xml, options)?;
        }

        // xl/_rels/workbook.xml.rels
        write_xml_part(zip, "xl/_rels/workbook.xml.rels", &workbook_rels, options)?;

        // xl/worksheets/sheet{N}.xml
        for (i, (_name, ws_lock)) in self.worksheets.iter().enumerate() {
            let entry_name = self.sheet_part_path(i);
            let dirty = self.sheet_dirty.get(i).copied().unwrap_or(true);

            // If the sheet has streamed data, write it directly from the temp file.
            if let Some(streamed) = self.streamed_sheets.get(&i) {
                crate::stream::write_streamed_sheet(zip, &entry_name, streamed, options)?;
                continue;
            }

            // Copy-on-write passthrough: if the sheet has not been modified
            // (not dirty) and raw XML bytes are available, write them directly.
            // This avoids deserialize-then-serialize overhead for untouched
            // sheets. Dirty sheets always take the serialization path even if
            // raw bytes happen to still be present.
            //
            // The passthrough is also disabled when auxiliary parts (comments,
            // tables, sparklines) require XML injection into the worksheet,
            // since the raw bytes would lack those references.
            let slicer_rids: Vec<&str> = worksheet_rels
                .get(&i)
                .map(|rels| {
                    rels.relationships
                        .iter()
                        .filter(|rel| rel.rel_type == rel_types::SLICER)
                        .map(|rel| rel.id.as_str())
                        .collect()
                })
                .unwrap_or_default();
            let needs_aux_injection = legacy_drawing_rids.contains_key(&i)
                || table_parts_by_sheet.contains_key(&i)
                || !slicer_rids.is_empty();
            if !dirty && !needs_aux_injection {
                if let Some(Some(raw_bytes)) = self.raw_sheet_xml.get(i) {
                    write_bytes_part(zip, &entry_name, raw_bytes, options)?;
                    continue;
                }
            }

            // For non-dirty sheets that need aux injection (comments/tables),
            // or lazy/deferred sheets whose OnceLock is uninitialized, hydrate
            // from raw bytes. We intentionally avoid worksheet_ref_by_index
            // here because it applies sheet_rows truncation, which would cause
            // data loss on save for sheets that were never read by the user.
            //
            // Filtered-out sheets (via OpenOptions::sheets) have their OnceLock
            // initialized with an empty placeholder while raw_sheet_xml holds
            // the real data. We must prefer raw bytes over the placeholder.
            let hydrated_for_save: WorksheetXml;
            let ws = if !dirty {
                if let Some(Some(raw_bytes)) = self.raw_sheet_xml.get(i) {
                    hydrated_for_save = deserialize_worksheet_xml(raw_bytes)?;
                    &hydrated_for_save
                } else {
                    match ws_lock.get() {
                        Some(ws) => ws,
                        None => continue,
                    }
                }
            } else {
                match ws_lock.get() {
                    Some(ws) => ws,
                    None => {
                        if let Some(Some(raw_bytes)) = self.raw_sheet_xml.get(i) {
                            hydrated_for_save = deserialize_worksheet_xml(raw_bytes)?;
                            &hydrated_for_save
                        } else {
                            continue;
                        }
                    }
                }
            };

            let empty_sparklines: Vec<crate::sparkline::SparklineConfig> = vec![];
            let sparklines = self.sheet_sparklines.get(i).unwrap_or(&empty_sparklines);
            let legacy_rid = legacy_drawing_rids.get(&i).map(|s| s.as_str());
            let sheet_table_rids = table_parts_by_sheet.get(&i);
            let stale_table_parts =
                should_sync_tables && sheet_table_rids.is_none() && ws.table_parts.is_some();
            let has_extras = legacy_rid.is_some()
                || !sparklines.is_empty()
                || sheet_table_rids.is_some()
                || stale_table_parts
                || !slicer_rids.is_empty();

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
                let xml = serialize_worksheet_with_slicer_extras(
                    ws_ref,
                    sparklines,
                    legacy_rid,
                    self.raw_sheet_xml.get(i).and_then(|raw| raw.as_deref()),
                    &slicer_rids,
                )?;
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

        let mut emitted_graph_paths: HashSet<String> = HashSet::new();

        // xl/drawings/drawing{N}.xml -- write drawing parts
        for (path, drawing) in &self.drawings {
            if let Some(bytes) = self.clean_raw_graph_part(path) {
                write_bytes_part(zip, path, bytes, options)?;
            } else {
                write_xml_part(zip, path, drawing, options)?;
            }
            emitted_graph_paths.insert(path.clone());
        }

        // xl/charts/chart{N}.xml -- write chart parts
        for (path, chart) in &self.charts {
            if let Some(bytes) = self.clean_raw_graph_part(path) {
                write_bytes_part(zip, path, bytes, options)?;
            } else {
                write_xml_part(zip, path, chart, options)?;
            }
            emitted_graph_paths.insert(path.clone());
        }
        for (path, data) in &self.raw_charts {
            if self.charts.iter().any(|(p, _)| p == path) {
                continue;
            }
            zip.start_file(path, options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(self.clean_raw_graph_part(path).unwrap_or(data))?;
            emitted_graph_paths.insert(path.clone());
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
                if let Some(bytes) = self.clean_raw_graph_part(&path) {
                    write_bytes_part(zip, &path, bytes, options)?;
                } else {
                    write_xml_part(zip, &path, rels, options)?;
                }
                emitted_graph_paths.insert(path);
            }
        }

        // Raw drawing or relationship parts that could not be parsed eagerly.
        for (path, data) in &self.raw_graph_parts {
            if self.dirty_graph_parts.contains(path) || emitted_graph_paths.contains(path) {
                continue;
            }
            write_bytes_part(zip, path, data, options)?;
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
            let cache_definition =
                self.pivot_cache_defs
                    .iter()
                    .find_map(|(definition_path, definition)| {
                        let rels_path = relationship_part_path(definition_path);
                        let bytes = self
                            .raw_graph_parts
                            .get(&rels_path)
                            .map(Vec::as_slice)
                            .or_else(|| {
                                self.unknown_parts
                                    .iter()
                                    .find(|(unknown_path, _)| unknown_path == &rels_path)
                                    .map(|(_, bytes)| bytes.as_slice())
                            })?;
                        let relationships =
                            quick_xml::de::from_reader::<_, Relationships>(bytes).ok()?;
                        relationships
                            .relationships
                            .iter()
                            .any(|relationship| {
                                relationship.rel_type == rel_types::PIVOT_CACHE_RECORDS
                                    && resolve_relationship_target(
                                        definition_path,
                                        &relationship.target,
                                    ) == *path
                            })
                            .then_some(definition)
                    });
            let xml = if let Some(definition) = cache_definition {
                serialize_pivot_cache_records(pcr, definition)?
            } else {
                serialize_xml(pcr)?
            };
            zip.start_file(path, options)
                .map_err(|error| Error::Zip(error.to_string()))?;
            zip.write_all(xml.as_bytes())?;
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
        if let Some(raw) = self.clean_raw_doc_prop("docProps/core.xml") {
            write_bytes_part(zip, "docProps/core.xml", raw, options)?;
        } else if let Some(ref props) = self.core_properties {
            let xml_str = sheetkit_xml::doc_props::serialize_core_properties(props);
            zip.start_file("docProps/core.xml", options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(xml_str.as_bytes())?;
        }

        // docProps/app.xml
        if let Some(raw) = self.clean_raw_doc_prop("docProps/app.xml") {
            write_bytes_part(zip, "docProps/app.xml", raw, options)?;
        } else if let Some(ref props) = self.app_properties {
            write_xml_part(zip, "docProps/app.xml", props, options)?;
        }

        // docProps/custom.xml
        if let Some(raw) = self.clean_raw_doc_prop("docProps/custom.xml") {
            write_bytes_part(zip, "docProps/custom.xml", raw, options)?;
        } else if let Some(ref props) = self.custom_properties {
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

        // Write back deferred parts from Lazy open (raw bytes, unparsed).
        // Skip any path that was already written by the normal code above. This
        // prevents duplicate ZIP entries when an auxiliary part (comments, doc
        // properties, etc.) is mutated after a Lazy open.
        if self.deferred_parts.has_any() {
            let mut emitted_owned: HashSet<String> = HashSet::new();
            // Essential parts always written.
            emitted_owned.insert("[Content_Types].xml".to_string());
            emitted_owned.insert("_rels/.rels".to_string());
            emitted_owned.insert("xl/workbook.xml".to_string());
            emitted_owned.insert("xl/_rels/workbook.xml.rels".to_string());
            emitted_owned.insert("xl/styles.xml".to_string());
            emitted_owned.insert("xl/sharedStrings.xml".to_string());
            emitted_owned.insert("xl/theme/theme1.xml".to_string());
            // Per-sheet worksheet paths.
            for i in 0..self.worksheets.len() {
                emitted_owned.insert(self.sheet_part_path(i));
            }
            for (i, comments) in self.sheet_comments.iter().enumerate() {
                if comments.is_some() {
                    emitted_owned.insert(format!("xl/comments{}.xml", i + 1));
                }
            }
            for (_sheet_idx, vml_path, _) in &vml_parts_to_write {
                emitted_owned.insert(vml_path.clone());
            }
            for (path, _) in &self.drawings {
                emitted_owned.insert(path.clone());
            }
            for (path, _) in &self.charts {
                emitted_owned.insert(path.clone());
            }
            for (path, _) in &self.raw_charts {
                emitted_owned.insert(path.clone());
            }
            emitted_owned.extend(self.raw_graph_parts.keys().cloned());
            for (path, _) in &self.images {
                emitted_owned.insert(path.clone());
            }
            for sheet_idx in worksheet_rels.keys() {
                let sheet_path = self.sheet_part_path(*sheet_idx);
                emitted_owned.insert(relationship_part_path(&sheet_path));
            }
            for drawing_idx in self.drawing_rels.keys() {
                if let Some((drawing_path, _)) = self.drawings.get(*drawing_idx) {
                    emitted_owned.insert(relationship_part_path(drawing_path));
                }
            }
            for (path, _) in &self.pivot_tables {
                emitted_owned.insert(path.clone());
            }
            for (path, _) in &self.pivot_cache_defs {
                emitted_owned.insert(path.clone());
            }
            for (path, _) in &self.pivot_cache_records {
                emitted_owned.insert(path.clone());
            }
            for (path, _, _) in &self.tables {
                emitted_owned.insert(path.clone());
            }
            for (path, _) in &self.slicer_defs {
                emitted_owned.insert(path.clone());
            }
            for (path, _) in &self.slicer_caches {
                emitted_owned.insert(path.clone());
            }
            if self.vba_blob.is_some() {
                emitted_owned.insert("xl/vbaProject.bin".to_string());
            }
            if self.core_properties.is_some() {
                emitted_owned.insert("docProps/core.xml".to_string());
            }
            if self.app_properties.is_some() {
                emitted_owned.insert("docProps/app.xml".to_string());
            }
            if self.custom_properties.is_some() {
                emitted_owned.insert("docProps/custom.xml".to_string());
            }
            emitted_owned.extend(self.raw_doc_props.keys().cloned());
            if has_any_threaded {
                for (i, tc) in self.sheet_threaded_comments.iter().enumerate() {
                    if tc.is_some() {
                        emitted_owned
                            .insert(format!("xl/threadedComments/threadedComment{}.xml", i + 1));
                    }
                }
                emitted_owned.insert("xl/persons/person.xml".to_string());
            }
            for (path, _) in &self.unknown_parts {
                emitted_owned.insert(path.clone());
            }

            for (path, data) in self.deferred_parts.remaining_parts() {
                if emitted_owned.contains(path) {
                    continue;
                }
                zip.start_file(path, options)
                    .map_err(|e| Error::Zip(e.to_string()))?;
                zip.write_all(data)?;
            }
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

fn serialize_pivot_cache_records(
    records: &sheetkit_xml::pivot_cache::PivotCacheRecords,
    definition: &sheetkit_xml::pivot_cache::PivotCacheDefinition,
) -> Result<String> {
    use std::fmt::Write as _;

    let field_kinds = definition
        .cache_fields
        .fields
        .iter()
        .map(|field| {
            let items = field.shared_items.as_ref();
            items.is_some_and(|items| {
                items.contains_number == Some(true) && items.contains_string != Some(true)
            })
        })
        .collect::<Vec<_>>();
    let expected_numbers = field_kinds.iter().filter(|is_number| **is_number).count();
    let expected_indexes = field_kinds.len().saturating_sub(expected_numbers);
    if records.records.iter().any(|record| {
        record.number_fields.len() != expected_numbers
            || record.index_fields.len() != expected_indexes
            || !record.string_fields.is_empty()
            || !record.bool_fields.is_empty()
    }) {
        return serialize_xml(records);
    }

    let mut xml = String::new();
    xml.push_str(XML_DECLARATION);
    xml.push('\n');
    write!(
        xml,
        "<pivotCacheRecords xmlns=\"{}\" xmlns:r=\"{}\" count=\"{}\">",
        quick_xml::escape::escape(&records.xmlns),
        quick_xml::escape::escape(&records.xmlns_r),
        records.count.unwrap_or(records.records.len() as u32),
    )
    .map_err(|error| Error::Internal(error.to_string()))?;
    for record in &records.records {
        xml.push_str("<r>");
        let mut indexes = record.index_fields.iter();
        let mut numbers = record.number_fields.iter();
        for is_number in &field_kinds {
            if *is_number {
                let value = numbers
                    .next()
                    .ok_or_else(|| Error::Internal("missing pivot cache number".to_string()))?;
                write!(xml, "<n v=\"{}\"/>", value.v)
                    .map_err(|error| Error::Internal(error.to_string()))?;
            } else {
                let value = indexes
                    .next()
                    .ok_or_else(|| Error::Internal("missing pivot cache index".to_string()))?;
                write!(xml, "<x v=\"{}\"/>", value.v)
                    .map_err(|error| Error::Internal(error.to_string()))?;
            }
        }
        xml.push_str("</r>");
    }
    xml.push_str("</pivotCacheRecords>");
    Ok(xml)
}

/// Deserialize a `WorksheetXml` from raw XML bytes.
///
/// This is the on-demand counterpart of `read_xml_part` for worksheet data
/// that was stored as raw bytes during open (lazy mode or filtered-out sheets).
/// After deserialization, cell column numbers are populated and rows/cells
/// are sorted for binary-search correctness.
pub(super) fn deserialize_worksheet_xml(bytes: &[u8]) -> Result<WorksheetXml> {
    let buf_cap = bytes.len().clamp(8192, LARGE_BUF_CAPACITY);
    let reader = std::io::BufReader::with_capacity(buf_cap, bytes);
    let mut ws: WorksheetXml =
        quick_xml::de::from_reader(reader).map_err(|e| Error::XmlDeserialize(e.to_string()))?;
    // Post-process: populate cached column numbers, sort rows and cells.
    ws.sheet_data.rows.sort_unstable_by_key(|r| r.r);
    for row in &mut ws.sheet_data.rows {
        for cell in &mut row.cells {
            cell.col = fast_col_number(cell.r.as_str());
        }
        row.cells.sort_unstable_by_key(|c| c.col);
        row.cells.shrink_to_fit();
    }
    ws.sheet_data.rows.shrink_to_fit();
    Ok(ws)
}

/// BufReader capacity for large XML parts (worksheets, sharedStrings).
/// 64 KB reduces read-syscall overhead compared to the 8 KB default.
const LARGE_BUF_CAPACITY: usize = 64 * 1024;

/// Read a ZIP entry and deserialize it from XML.
///
/// Uses `quick_xml::de::from_reader` to deserialize directly from the
/// decompressed ZIP stream, avoiding the intermediate `String` allocation
/// that `read_to_string` + `from_str` would require. The BufReader
/// capacity is scaled based on the uncompressed entry size, up to 64 KB,
/// to reduce read-syscall overhead on large parts.
pub(crate) fn read_xml_part<T: serde::de::DeserializeOwned, R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    name: &str,
) -> Result<T> {
    let entry = archive
        .by_name(name)
        .map_err(|e| Error::Zip(e.to_string()))?;
    let size = entry.size() as usize;
    let buf_cap = size.clamp(8192, LARGE_BUF_CAPACITY);
    let reader = std::io::BufReader::with_capacity(buf_cap, entry);
    quick_xml::de::from_reader(reader).map_err(|e| Error::XmlDeserialize(e.to_string()))
}

pub(crate) fn deserialize_xml_bytes<T: serde::de::DeserializeOwned>(bytes: &[u8]) -> Result<T> {
    quick_xml::de::from_reader(bytes).map_err(|e| Error::XmlDeserialize(e.to_string()))
}

#[derive(Default)]
struct RawSharedStringText {
    plain: Option<String>,
    runs: Vec<String>,
}

enum SharedStringTextTarget {
    Plain,
    Run,
}

/// Deserialize an SST while retaining whitespace that quick-xml's serde
/// deserializer trims from text nodes.
fn read_shared_strings_part<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    name: &str,
) -> Result<Option<Sst>> {
    let mut entry = match archive.by_name(name) {
        Ok(entry) => entry,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(error) => return Err(Error::Zip(error.to_string())),
    };
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes)?;
    deserialize_shared_strings(&bytes).map(Some)
}

fn deserialize_shared_strings(bytes: &[u8]) -> Result<Sst> {
    let reader =
        std::io::BufReader::with_capacity(bytes.len().clamp(8192, LARGE_BUF_CAPACITY), bytes);
    let mut sst: Sst =
        quick_xml::de::from_reader(reader).map_err(|e| Error::XmlDeserialize(e.to_string()))?;
    let raw_items = scan_shared_string_text(bytes)?;

    if raw_items.len() != sst.items.len() {
        return Err(Error::XmlDeserialize(format!(
            "shared string item count mismatch: parsed {}, scanned {}",
            sst.items.len(),
            raw_items.len()
        )));
    }

    for (index, (item, raw)) in sst.items.iter_mut().zip(raw_items).enumerate() {
        match (&mut item.t, item.r.as_mut_slice(), raw.plain, raw.runs) {
            (Some(text), runs, Some(raw_text), raw_runs) if runs.len() == raw_runs.len() => {
                text.value = raw_text;
                for (run, raw_text) in runs.iter_mut().zip(raw_runs) {
                    run.t.value = raw_text;
                }
            }
            (None, runs, None, raw_runs) if runs.len() == raw_runs.len() => {
                for (run, raw_text) in runs.iter_mut().zip(raw_runs) {
                    run.t.value = raw_text;
                }
            }
            (None, [], None, raw_runs) if raw_runs.is_empty() => {}
            _ => {
                return Err(Error::XmlDeserialize(format!(
                    "shared string item {index} text shape mismatch"
                )));
            }
        }
    }

    Ok(sst)
}

fn scan_shared_string_text(bytes: &[u8]) -> Result<Vec<RawSharedStringText>> {
    use quick_xml::events::Event;

    let mut reader = quick_xml::Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut items = Vec::new();
    let mut current_item: Option<RawSharedStringText> = None;
    let mut text_target: Option<SharedStringTextTarget> = None;
    let mut text_value = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                let name = start.local_name().as_ref().to_vec();
                let parent = stack.last().map(Vec::as_slice);
                if name.as_slice() == b"si" {
                    if current_item.is_some() {
                        return Err(Error::XmlDeserialize(
                            "nested shared string item".to_string(),
                        ));
                    }
                    current_item = Some(RawSharedStringText::default());
                } else if name.as_slice() == b"t" && current_item.is_some() {
                    text_target = match parent {
                        Some(b"si") => Some(SharedStringTextTarget::Plain),
                        Some(b"r") => Some(SharedStringTextTarget::Run),
                        _ => None,
                    };
                    text_value.clear();
                }
                stack.push(name);
            }
            Ok(Event::Empty(empty)) => {
                let name = empty.local_name();
                let parent = stack.last().map(Vec::as_slice);
                if name.as_ref() == b"si" {
                    items.push(RawSharedStringText::default());
                } else if name.as_ref() == b"t" {
                    if let Some(item) = current_item.as_mut() {
                        match parent {
                            Some(b"si") => item.plain = Some(String::new()),
                            Some(b"r") => item.runs.push(String::new()),
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Text(text)) => {
                if text_target.is_some() {
                    text_value.push_str(
                        &text
                            .unescape()
                            .map_err(|error| Error::XmlParse(error.to_string()))?,
                    );
                }
            }
            Ok(Event::CData(cdata)) => {
                if text_target.is_some() {
                    text_value.push_str(
                        &reader
                            .decoder()
                            .decode(cdata.as_ref())
                            .map_err(|error| Error::XmlParse(error.to_string()))?,
                    );
                }
            }
            Ok(Event::End(end)) => {
                let name = end.local_name();
                if name.as_ref() == b"t" {
                    if let (Some(item), Some(target)) = (current_item.as_mut(), text_target.take())
                    {
                        let value = std::mem::take(&mut text_value);
                        match target {
                            SharedStringTextTarget::Plain => item.plain = Some(value),
                            SharedStringTextTarget::Run => item.runs.push(value),
                        }
                    }
                } else if name.as_ref() == b"si" {
                    let item = current_item.take().ok_or_else(|| {
                        Error::XmlDeserialize("shared string end without start".to_string())
                    })?;
                    items.push(item);
                }
                stack.pop();
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::XmlParse(error.to_string())),
        }
    }

    if current_item.is_some() || text_target.is_some() {
        return Err(Error::XmlDeserialize(
            "unterminated shared string item".to_string(),
        ));
    }
    Ok(items)
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
        // Clean owners bypass this helper through raw passthrough. Dirty
        // owners are rebuilt in worksheet schema order without raw merging.
        if !legacy_xml.is_empty() {
            if let Some(table_parts_start) = prefix.find("<tableParts") {
                result.push_str(&prefix[..table_parts_start]);
                result.push_str(&legacy_xml);
                result.push_str(&prefix[table_parts_start..]);
            } else {
                result.push_str(prefix);
                result.push_str(&legacy_xml);
            }
        } else {
            result.push_str(prefix);
        }
        result.push_str(&ext_xml);
        result.push_str(closing);
        Ok(result)
    } else {
        Ok(format!("{XML_DECLARATION}\n{body}"))
    }
}

fn serialize_worksheet_with_slicer_extras(
    ws: &WorksheetXml,
    sparklines: &[crate::sparkline::SparklineConfig],
    legacy_drawing_rid: Option<&str>,
    raw_sheet: Option<&[u8]>,
    slicer_rids: &[&str],
) -> Result<String> {
    let mut xml = serialize_worksheet_with_extras(ws, sparklines, legacy_drawing_rid)?;
    if slicer_rids.is_empty() {
        return Ok(xml);
    }
    let generated_entries = ext_list_contents(&xml).unwrap_or_default();
    xml = remove_extension_list(xml);
    let raw_entries = raw_sheet
        .and_then(|bytes| std::str::from_utf8(bytes).ok())
        .and_then(ext_list_contents)
        .map(|contents| {
            let contents =
                remove_known_slicer_extensions(contents, "A8765BA9-456A-4DAB-B4F3-ACF838C121DE");
            remove_known_slicer_extensions(contents, "05C60535-1F16-4fd2-B633-F4F36F0B64E0")
        })
        .unwrap_or_default();
    let entries = format!("{raw_entries}{generated_entries}");
    let slicers = slicer_rids
        .iter()
        .map(|rid| {
            format!(
                "<x14:slicer r:id=\"{}\" xmlns:r=\"{}\"/>",
                rid,
                sheetkit_xml::namespaces::RELATIONSHIPS,
            )
        })
        .collect::<String>();
    let extension = format!("<extLst>{entries}<ext uri=\"{{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}}\"><x14:slicerList xmlns:x14=\"{}\">{slicers}</x14:slicerList></ext></extLst>", sheetkit_xml::namespaces::SLICER_2009);
    let closing = "</worksheet>";
    let position = xml
        .rfind(closing)
        .ok_or_else(|| Error::XmlParse("worksheet closing tag missing".to_string()))?;
    xml.insert_str(position, &extension);
    Ok(xml)
}

fn serialize_workbook_with_slicer_extensions(
    workbook: &WorkbookXml,
    raw_workbook: Option<&[u8]>,
    slicer_cache_rids: &[&str],
) -> Result<String> {
    let mut body =
        quick_xml::se::to_string(workbook).map_err(|error| Error::XmlParse(error.to_string()))?;
    let entries = raw_workbook
        .and_then(|bytes| std::str::from_utf8(bytes).ok())
        .and_then(ext_list_contents)
        .map(|contents| {
            remove_known_slicer_extensions(contents, "BBE1A952-AA13-448E-AADC-164F8A28A991")
        })
        .unwrap_or_default();
    let caches = slicer_cache_rids
        .iter()
        .map(|rid| {
            format!(
                "<x14:slicerCache r:id=\"{}\" xmlns:r=\"{}\"/>",
                rid,
                sheetkit_xml::namespaces::RELATIONSHIPS,
            )
        })
        .collect::<String>();
    let closing = "</workbook>";
    let position = body
        .rfind(closing)
        .ok_or_else(|| Error::XmlParse("workbook closing tag missing".to_string()))?;
    body.insert_str(position, &format!("<extLst>{entries}<ext uri=\"{{BBE1A952-AA13-448E-AADC-164F8A28A991}}\"><x14:slicerCaches xmlns:x14=\"{}\">{caches}</x14:slicerCaches></ext></extLst>", sheetkit_xml::namespaces::SLICER_2009));
    Ok(format!("{XML_DECLARATION}\n{body}"))
}

fn ext_list_contents(xml: &str) -> Option<String> {
    let start = xml.find("<extLst")?;
    let content_start = xml[start..].find('>')? + start + 1;
    let end = xml[content_start..].find("</extLst>")? + content_start;
    Some(xml[content_start..end].to_string())
}

fn remove_extension_list(mut xml: String) -> String {
    if let Some(start) = xml.find("<extLst") {
        if let Some(end) = xml[start..].find("</extLst>") {
            xml.replace_range(start..start + end + "</extLst>".len(), "");
        }
    }
    xml
}

fn remove_known_slicer_extensions(mut contents: String, uri: &str) -> String {
    while let Some(offset) = contents
        .to_ascii_lowercase()
        .find(&uri.to_ascii_lowercase())
    {
        let start = contents[..offset].rfind("<ext").unwrap_or(offset);
        let Some(end) = contents[offset..].find("</ext>") else {
            break;
        };
        contents.replace_range(start..offset + end + "</ext>".len(), "");
    }
    contents
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
    write_bytes_part(zip, name, xml.as_bytes(), options)
}

fn write_bytes_part<W: std::io::Write + std::io::Seek>(
    zip: &mut zip::ZipWriter<W>,
    name: &str,
    bytes: &[u8],
    options: SimpleFileOptions,
) -> Result<()> {
    zip.start_file(name, options)
        .map_err(|e| Error::Zip(e.to_string()))?;
    zip.write_all(bytes)?;
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
#[allow(clippy::unnecessary_map_or)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn known_slicer_extensions_are_removed_case_insensitively() {
        let contents = concat!(
            "<ext uri=\"{bbe1a952-aa13-448e-aadc-164f8a28a991}\">",
            "<x14:slicerCaches/>",
            "</ext>",
            "<ext uri=\"{opaque}\"><opaque:payload/></ext>"
        );

        let remaining = remove_known_slicer_extensions(
            contents.to_string(),
            "BBE1A952-AA13-448E-AADC-164F8A28A991",
        );

        assert!(!remaining.contains("slicerCaches"));
        assert!(remaining.contains("{opaque}"));
    }

    #[test]
    fn adding_a_slicer_merges_existing_workbook_and_sheet_extensions() {
        let mut workbook = Workbook::new();
        workbook
            .add_table(
                "Sheet1",
                &crate::table::TableConfig {
                    name: "Table1".to_string(),
                    display_name: "Table1".to_string(),
                    range: "A1:B3".to_string(),
                    columns: vec![
                        crate::table::TableColumn {
                            name: "Status".to_string(),
                            totals_row_function: None,
                            totals_row_label: None,
                        },
                        crate::table::TableColumn {
                            name: "Value".to_string(),
                            totals_row_function: None,
                            totals_row_label: None,
                        },
                    ],
                    ..crate::table::TableConfig::default()
                },
            )
            .unwrap();
        let base = workbook.save_to_buffer().unwrap();
        let workbook_xml = String::from_utf8(zip_part(&base, "xl/workbook.xml"))
            .unwrap()
            .replacen(
                "</workbook>",
                concat!(
                    "<extLst>",
                    "<ext uri=\"{opaque-workbook}\">",
                    "<opaque:payload xmlns:opaque=\"urn:sheetkit:workbook\"/>",
                    "</ext>",
                    "<ext uri=\"{bbe1a952-aa13-448e-aadc-164f8a28a991}\">",
                    "<x14:slicerCaches xmlns:x14=\"",
                    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                    "\"/>",
                    "</ext>",
                    "</extLst></workbook>"
                ),
                1,
            );
        let sheet_xml = String::from_utf8(zip_part(&base, "xl/worksheets/sheet1.xml"))
            .unwrap()
            .replacen(
                "</worksheet>",
                concat!(
                    "<extLst>",
                    "<ext uri=\"{opaque-sheet}\">",
                    "<opaque:payload xmlns:opaque=\"urn:sheetkit:sheet\"/>",
                    "</ext>",
                    "<ext uri=\"{a8765ba9-456a-4dab-b4f3-acf838c121de}\">",
                    "<x14:slicerList xmlns:x14=\"",
                    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                    "\"/>",
                    "</ext>",
                    "</extLst></worksheet>"
                ),
                1,
            );
        let input = rewrite_zip_parts(
            &base,
            &[
                ("xl/workbook.xml", Some(workbook_xml.into_bytes())),
                ("xl/worksheets/sheet1.xml", Some(sheet_xml.into_bytes())),
            ],
        );
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut opened = Workbook::open_from_buffer_with_options(&input, &options).unwrap();

        opened
            .add_slicer(
                "Sheet1",
                &crate::slicer::SlicerConfig {
                    name: "StatusFilter".to_string(),
                    cell: "D1".to_string(),
                    table_name: "Table1".to_string(),
                    column_name: "Status".to_string(),
                    caption: None,
                    style: None,
                    width: None,
                    height: None,
                    show_caption: None,
                    column_count: None,
                },
            )
            .unwrap();
        let saved = opened.save_to_buffer().unwrap();
        let saved_workbook = String::from_utf8(zip_part(&saved, "xl/workbook.xml")).unwrap();
        let saved_sheet = String::from_utf8(zip_part(&saved, "xl/worksheets/sheet1.xml")).unwrap();

        assert!(saved_workbook.contains("urn:sheetkit:workbook"));
        assert!(saved_sheet.contains("urn:sheetkit:sheet"));
        assert_eq!(saved_workbook.matches("<extLst>").count(), 1);
        assert_eq!(saved_sheet.matches("<extLst>").count(), 1);
        assert_eq!(saved_workbook.matches("<x14:slicerCaches").count(), 1);
        assert_eq!(saved_sheet.matches("<x14:slicerList").count(), 1);
    }

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

    fn rewrite_zip_parts(buffer: &[u8], replacements: &[(&str, Option<Vec<u8>>)]) -> Vec<u8> {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer)).unwrap();
        let mut output = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(&mut output));
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            for index in 0..archive.len() {
                let mut entry = archive.by_index(index).unwrap();
                let name = entry.name().to_string();
                let replacement = replacements
                    .iter()
                    .find(|(path, _)| *path == name)
                    .map(|(_, value)| value);
                if matches!(replacement, Some(None)) {
                    continue;
                }
                let mut bytes = Vec::new();
                entry.read_to_end(&mut bytes).unwrap();
                let bytes = replacement.and_then(Option::as_ref).unwrap_or(&bytes);
                writer.start_file(name, options).unwrap();
                writer.write_all(bytes).unwrap();
            }
            writer.finish().unwrap();
        }
        output
    }

    fn workbook_with_chart_buffer() -> Vec<u8> {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};

        let mut workbook = Workbook::new();
        workbook.new_sheet("Sheet2").unwrap();
        workbook
            .add_chart(
                "Sheet1",
                "E2",
                "L15",
                &ChartConfig {
                    chart_type: ChartType::Col,
                    title: Some("Preserved chart".to_string()),
                    series: vec![ChartSeries {
                        name: "Values".to_string(),
                        categories: "Sheet1!$A$1:$A$3".to_string(),
                        values: "Sheet1!$B$1:$B$3".to_string(),
                        x_values: None,
                        bubble_sizes: None,
                    }],
                    show_legend: true,
                    view_3d: None,
                },
            )
            .unwrap();
        workbook.save_to_buffer().unwrap()
    }

    fn opaque_graph_buffer() -> (Vec<u8>, Vec<u8>, Vec<u8>, Vec<u8>) {
        let base = workbook_with_chart_buffer();

        let chart = String::from_utf8(zip_part(&base, "xl/charts/chart1.xml")).unwrap();
        let chart = chart.replacen(
            "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"",
            "xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart'",
            1,
        );
        assert!(chart.contains("xmlns:c='"));

        let drawing = String::from_utf8(zip_part(&base, "xl/drawings/drawing1.xml")).unwrap();
        let drawing = drawing.replacen(
            "xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"",
            "xmlns:xdr='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'",
            1,
        );
        assert!(drawing.contains("xmlns:xdr='"));
        if let Err(error) = deserialize_xml_bytes::<WsDr>(drawing.as_bytes()) {
            panic!("opaque drawing must remain partially parseable: {error}");
        }

        let drawing_rels =
            String::from_utf8(zip_part(&base, "xl/drawings/_rels/drawing1.xml.rels"))
                .unwrap()
                .replacen(
                    "</Relationships>",
                    "<!--preserve-this-comment--></Relationships>",
                    1,
                );

        let chart = chart.into_bytes();
        let drawing = drawing.into_bytes();
        let drawing_rels = drawing_rels.into_bytes();
        let rewritten = rewrite_zip_parts(
            &base,
            &[
                ("xl/charts/chart1.xml", Some(chart.clone())),
                ("xl/drawings/drawing1.xml", Some(drawing.clone())),
                (
                    "xl/drawings/_rels/drawing1.xml.rels",
                    Some(drawing_rels.clone()),
                ),
            ],
        );
        (rewritten, chart, drawing, drawing_rels)
    }

    fn unsupported_graph_buffer() -> (Vec<u8>, Vec<u8>, Vec<u8>) {
        let base = workbook_with_chart_buffer();
        let chart = String::from_utf8(zip_part(&base, "xl/charts/chart1.xml")).unwrap();
        let chart = chart.replacen(
            "</c:barChart>",
            "<c:dLbls><c:showVal val=\"1\"/></c:dLbls></c:barChart>",
            1,
        );
        let chart = chart.replacen(
            "</c:plotArea>",
            concat!(
                "<c:dateAx><c:axId val=\"9001\"/></c:dateAx>",
                "<c:valAx><c:axId val=\"9002\"/></c:valAx>",
                "<c:spPr><a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/",
                "drawingml/2006/main\"><a:srgbClr val=\"123456\"/></a:solidFill></c:spPr>",
                "</c:plotArea>"
            ),
            1,
        );

        let drawing = String::from_utf8(zip_part(&base, "xl/drawings/drawing1.xml")).unwrap();
        let drawing = drawing.replacen(
            "<xdr:twoCellAnchor>",
            "<xdr:twoCellAnchor editAs=\"oneCell\">",
            1,
        );
        let drawing = drawing.replacen(
            "</xdr:twoCellAnchor>",
            concat!(
                "<xdr:cxnSp/><xdr:grpSp/>",
                "</xdr:twoCellAnchor>",
                "<xdr:absoluteAnchor><xdr:pos x=\"0\" y=\"0\"/>",
                "<xdr:ext cx=\"1\" cy=\"1\"/><xdr:sp/></xdr:absoluteAnchor>"
            ),
            1,
        );

        let chart = chart.into_bytes();
        let drawing = drawing.into_bytes();
        let rewritten = rewrite_zip_parts(
            &base,
            &[
                ("xl/charts/chart1.xml", Some(chart.clone())),
                ("xl/drawings/drawing1.xml", Some(drawing.clone())),
            ],
        );
        (rewritten, chart, drawing)
    }

    fn assert_canonical_sst_metadata(buffer: &[u8]) {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer)).unwrap();
        let content_types: ContentTypes =
            read_xml_part(&mut archive, "[Content_Types].xml").unwrap();
        let relationships: Relationships =
            read_xml_part(&mut archive, "xl/_rels/workbook.xml.rels").unwrap();
        assert_eq!(
            content_types
                .overrides
                .iter()
                .filter(|entry| {
                    entry.part_name == "/xl/sharedStrings.xml"
                        && entry.content_type == mime_types::SHARED_STRINGS
                })
                .count(),
            1
        );
        assert_eq!(
            relationships
                .relationships
                .iter()
                .filter(|relationship| {
                    relationship.rel_type == rel_types::SHARED_STRINGS
                        && relationship.target == "sharedStrings.xml"
                })
                .count(),
            1
        );
    }

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
    fn test_new_workbook_writes_interop_workbook_defaults() {
        let wb = Workbook::new();
        let buf = wb.save_to_buffer().unwrap();

        let cursor = std::io::Cursor::new(buf);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut workbook_xml = String::new();
        std::io::Read::read_to_string(
            &mut archive.by_name("xl/workbook.xml").unwrap(),
            &mut workbook_xml,
        )
        .unwrap();

        assert!(workbook_xml.contains("<fileVersion"));
        assert!(workbook_xml.contains("appName=\"xl\""));
        assert!(workbook_xml.contains("lastEdited=\"7\""));
        assert!(workbook_xml.contains("lowestEdited=\"7\""));
        assert!(workbook_xml.contains("rupBuild=\"27425\""));

        assert!(workbook_xml.contains("<workbookPr"));
        assert!(workbook_xml.contains("defaultThemeVersion=\"166925\""));

        assert!(workbook_xml.contains("<bookViews>"));
        assert!(workbook_xml.contains("<workbookView"));
        assert!(workbook_xml.contains("activeTab=\"0\""));
        assert!(!workbook_xml.contains("xWindow="));
        assert!(!workbook_xml.contains("yWindow="));
        assert!(!workbook_xml.contains("windowWidth="));
        assert!(!workbook_xml.contains("windowHeight="));
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
    fn test_save_adds_canonical_sst_metadata_to_inline_string_package() {
        let base = Workbook::new().save_to_buffer().unwrap();
        let mut content_types: ContentTypes = {
            let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&base)).unwrap();
            read_xml_part(&mut archive, "[Content_Types].xml").unwrap()
        };
        content_types
            .overrides
            .retain(|entry| entry.part_name != "/xl/sharedStrings.xml");
        let mut workbook_rels: Relationships = {
            let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&base)).unwrap();
            read_xml_part(&mut archive, "xl/_rels/workbook.xml.rels").unwrap()
        };
        workbook_rels
            .relationships
            .retain(|relationship| relationship.rel_type != rel_types::SHARED_STRINGS);
        let worksheet = String::from_utf8(zip_part(&base, "xl/worksheets/sheet1.xml"))
            .unwrap()
            .replace(
                "<sheetData/>",
                "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Inline</t></is></c></row></sheetData>",
            );
        let input = rewrite_zip_parts(
            &base,
            &[
                ("xl/sharedStrings.xml", None),
                (
                    "[Content_Types].xml",
                    Some(serialize_xml(&content_types).unwrap().into_bytes()),
                ),
                (
                    "xl/_rels/workbook.xml.rels",
                    Some(serialize_xml(&workbook_rels).unwrap().into_bytes()),
                ),
                ("xl/worksheets/sheet1.xml", Some(worksheet.into_bytes())),
            ],
        );

        let mut workbook = Workbook::open_from_buffer(&input).unwrap();
        workbook
            .set_cell_value("Sheet1", "B1", CellValue::String("Shared".to_string()))
            .unwrap();
        let first = workbook.save_to_buffer().unwrap();
        assert_canonical_sst_metadata(&first);
        let reopened = Workbook::open_from_buffer(&first).unwrap();
        let second = reopened.save_to_buffer().unwrap();
        assert_canonical_sst_metadata(&second);
        assert_eq!(zip_entry_count(&second, "xl/sharedStrings.xml"), 1);
        let reopened_second = Workbook::open_from_buffer(&second).unwrap();
        assert_eq!(
            reopened_second.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Inline".to_string())
        );
        assert_eq!(
            reopened_second.get_cell_value("Sheet1", "B1").unwrap(),
            CellValue::String("Shared".to_string())
        );
    }

    #[test]
    fn test_deserialize_shared_strings_preserves_text_exactly() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t xml:space="preserve">  plain&#9;line&#10;&amp;  </t></si>
  <si><t xml:space="preserve"> direct </t><r><rPr><b/></rPr><t xml:space="preserve"> rich </t></r><r><t><![CDATA[tail  ]]></t></r></si>
  <si><t>entity &lt;value&gt;</t></si>
</sst>"#;
        let sst = deserialize_shared_strings(xml).unwrap();
        assert_eq!(sst.items[0].t.as_ref().unwrap().value, "  plain\tline\n&  ");
        assert_eq!(sst.items[1].t.as_ref().unwrap().value, " direct ");
        assert_eq!(sst.items[1].r[0].t.value, " rich ");
        assert_eq!(sst.items[1].r[1].t.value, "tail  ");
        assert!(sst.items[1].r[0].r_pr.as_ref().unwrap().b.is_some());
        assert_eq!(sst.items[2].t.as_ref().unwrap().value, "entity <value>");
    }

    #[test]
    fn test_shared_string_whitespace_survives_open_save_reopen() {
        let mut workbook = Workbook::new();
        workbook
            .set_cell_value("Sheet1", "A1", CellValue::String("placeholder".to_string()))
            .unwrap();
        let base = workbook.save_to_buffer().unwrap();
        let exact = "  leading\tand\nmultiple  spaces & entities <ok>  rich tail  ";
        let sst = format!(
            "{XML_DECLARATION}\n<sst xmlns=\"{}\" count=\"1\" uniqueCount=\"1\"><si><t xml:space=\"preserve\">  leading&#9;and&#10;multiple  spaces &amp; entities &lt;ok&gt;  </t><r><rPr><b/></rPr><t xml:space=\"preserve\">rich </t></r><r><t><![CDATA[tail  ]]></t></r></si></sst>",
            sheetkit_xml::namespaces::SPREADSHEET_ML
        );
        let input = rewrite_zip_parts(&base, &[("xl/sharedStrings.xml", Some(sst.into_bytes()))]);
        let opened = Workbook::open_from_buffer(&input).unwrap();
        assert_eq!(
            opened.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String(exact.to_string())
        );
        let saved = opened.save_to_buffer().unwrap();
        let reopened = Workbook::open_from_buffer(&saved).unwrap();
        assert_eq!(
            reopened.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String(exact.to_string())
        );
        let roundtripped = reopened.sst_runtime.to_sst();
        let item = &roundtripped.items[0];
        assert_eq!(
            item.t.as_ref().unwrap().value,
            "  leading\tand\nmultiple  spaces & entities <ok>  "
        );
        assert_eq!(item.r[0].t.value, "rich ");
        assert!(item.r[0].r_pr.as_ref().unwrap().b.is_some());
        assert_eq!(item.r[1].t.value, "tail  ");
    }

    fn workbook_with_slicer() -> Workbook {
        let mut workbook = Workbook::new();
        workbook
            .add_table(
                "Sheet1",
                &crate::table::TableConfig {
                    name: "Table1".to_string(),
                    display_name: "Table1".to_string(),
                    range: "A1:B3".to_string(),
                    columns: vec![
                        crate::table::TableColumn {
                            name: "Status".to_string(),
                            totals_row_function: None,
                            totals_row_label: None,
                        },
                        crate::table::TableColumn {
                            name: "Value".to_string(),
                            totals_row_function: None,
                            totals_row_label: None,
                        },
                    ],
                    ..crate::table::TableConfig::default()
                },
            )
            .unwrap();
        workbook
            .add_slicer(
                "Sheet1",
                &crate::slicer::SlicerConfig {
                    name: "StatusFilter".to_string(),
                    cell: "D1".to_string(),
                    table_name: "Table1".to_string(),
                    column_name: "Status".to_string(),
                    caption: None,
                    style: None,
                    width: None,
                    height: None,
                    show_caption: None,
                    column_count: None,
                },
            )
            .unwrap();
        workbook
    }

    #[test]
    fn test_eager_slicer_parts_are_emitted_once_across_two_saves() {
        use crate::workbook::open_options::{AuxParts, OpenOptions, ReadMode};

        let initial = workbook_with_slicer().save_to_buffer().unwrap();
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let opened = Workbook::open_from_buffer_with_options(&initial, &options).unwrap();
        let first = opened.save_to_buffer().unwrap();
        let reopened = Workbook::open_from_buffer_with_options(&first, &options).unwrap();
        let second = reopened.save_to_buffer().unwrap();

        for buffer in [&first, &second] {
            assert_eq!(zip_entry_count(buffer, "xl/slicers/slicer1.xml"), 1);
            assert_eq!(
                zip_entry_count(buffer, "xl/slicerCaches/slicerCache1.xml"),
                1
            );
        }
    }

    #[test]
    fn test_lazy_hydration_preserves_malformed_slicer_parts_as_raw() {
        use crate::workbook::open_options::{AuxParts, OpenOptions, ReadMode};

        let initial = workbook_with_slicer().save_to_buffer().unwrap();
        let malformed_slicer = b"<x14:slicers><broken>".to_vec();
        let malformed_cache = b"<x14:slicerCacheDefinition><broken>".to_vec();
        let input = rewrite_zip_parts(
            &initial,
            &[
                ("xl/slicers/slicer1.xml", Some(malformed_slicer.clone())),
                (
                    "xl/slicerCaches/slicerCache1.xml",
                    Some(malformed_cache.clone()),
                ),
            ],
        );
        let options = OpenOptions::new()
            .read_mode(ReadMode::Lazy)
            .aux_parts(AuxParts::Deferred);
        let mut opened = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
        assert!(opened.delete_slicer("Sheet1", "Missing").is_err());
        let saved = opened.save_to_buffer().unwrap();

        assert_eq!(zip_entry_count(&saved, "xl/slicers/slicer1.xml"), 1);
        assert_eq!(zip_part(&saved, "xl/slicers/slicer1.xml"), malformed_slicer);
        assert_eq!(
            zip_entry_count(&saved, "xl/slicerCaches/slicerCache1.xml"),
            1
        );
        assert_eq!(
            zip_part(&saved, "xl/slicerCaches/slicerCache1.xml"),
            malformed_cache
        );
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
        use crate::workbook::open_options::{AuxParts, OpenOptions, ReadMode};

        let vba_data = b"FAKE_VBA_PROJECT_BINARY_DATA_1234567890";
        let xlsm = build_xlsm_with_vba(vba_data);
        let opts = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let wb = Workbook::open_from_buffer_with_options(&xlsm, &opts).unwrap();
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
        use crate::workbook::open_options::{AuxParts, OpenOptions, ReadMode};

        let vba_data: Vec<u8> = (0..=255).cycle().take(1024).collect();
        let xlsm = build_xlsm_with_vba(&vba_data);

        let opts = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let wb = Workbook::open_from_buffer_with_options(&xlsm, &opts).unwrap();
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

    use crate::workbook::open_options::ReadMode;

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

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
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

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
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

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Cell data is readable.
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("data".to_string())
        );
        // Comments are hydrated on demand from deferred parts.
        let comments = wb2.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].text, "A test comment");
    }

    #[test]
    fn test_readfast_get_doc_properties_without_mutation() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::Number(1.0))
            .unwrap();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Test Title".to_string()),
            ..Default::default()
        });
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Cell data is readable.
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::Number(1.0)
        );
        // Doc properties should be readable directly from deferred parts.
        let props = wb2.get_doc_props();
        assert_eq!(props.title.as_deref(), Some("Test Title"));
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

        // Open in Lazy mode.
        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        // Re-open in Eager mode and verify all parts were preserved.
        let mut wb3 = Workbook::open_from_buffer(&saved).unwrap();
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

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy).sheet_rows(10);
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
            .read_mode(ReadMode::Lazy)
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

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
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

        // Eager mode: everything should be parsed.
        let opts = OpenOptions::new().read_mode(ReadMode::Eager);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
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

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_with_options(&path, &opts).unwrap();
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("file test".to_string())
        );
    }

    #[test]
    fn test_readfast_roundtrip_with_custom_zip_entries() {
        let buf = create_xlsx_with_custom_entries();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("hello".to_string())
        );

        let saved = wb.save_to_buffer().unwrap();
        let cursor = std::io::Cursor::new(&saved);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        // Custom entries should be preserved through Lazy open/save.
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

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        // When auxiliary parts exist, they should be captured in deferred_parts.
        assert!(
            wb2.deferred_parts.has_any(),
            "deferred_parts should contain skipped auxiliary parts"
        );
    }

    #[test]
    fn test_readfast_eager_mode_has_no_deferred_parts() {
        use crate::workbook::open_options::{AuxParts, OpenOptions, ReadMode};

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

        // Eager mode: deferred_parts should be empty.
        let opts = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert!(
            !wb2.deferred_parts.has_any(),
            "Eager mode should not have deferred parts"
        );
    }

    #[test]
    fn test_readfast_table_parts_preserved_on_roundtrip() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("Name".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet1", "B1", CellValue::String("Value".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet1", "A2", CellValue::String("Alice".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet1", "B2", CellValue::Number(10.0))
            .unwrap();
        wb.add_table(
            "Sheet1",
            &crate::table::TableConfig {
                name: "Table1".to_string(),
                display_name: "Table1".to_string(),
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

        // Open in Lazy mode and save.
        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        // Re-open in Eager mode and verify the table survived the round-trip.
        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let tables = wb3.get_tables("Sheet1").unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "Table1");
    }

    #[test]
    fn test_readfast_delete_table_with_other_deferred_cleans_references() {
        use std::io::Read as _;

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("Name".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet1", "B1", CellValue::String("Value".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet1", "A2", CellValue::String("Alice".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet1", "B2", CellValue::Number(10.0))
            .unwrap();
        wb.add_table(
            "Sheet1",
            &crate::table::TableConfig {
                name: "Table1".to_string(),
                display_name: "Table1".to_string(),
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
        // Keep another deferred category so has_deferred remains true in Lazy mode.
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Keep deferred".to_string()),
            ..Default::default()
        });
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        wb2.delete_table("Sheet1", "Table1").unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        assert!(wb3.get_tables("Sheet1").unwrap().is_empty());

        let cursor = std::io::Cursor::new(saved);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();

        let mut ct_xml = String::new();
        archive
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_string(&mut ct_xml)
            .unwrap();
        assert!(
            !ct_xml.contains("/xl/tables/table1.xml"),
            "content types must not reference the deleted table part"
        );
        assert!(
            !ct_xml.contains(mime_types::TABLE),
            "content types must not keep table override after deletion"
        );

        let mut rels_xml = String::new();
        archive
            .by_name("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap()
            .read_to_string(&mut rels_xml)
            .unwrap();
        assert!(
            !rels_xml.contains(rel_types::TABLE),
            "worksheet rels must not contain table relationship after deletion"
        );

        let mut sheet_xml = String::new();
        archive
            .by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();
        assert!(
            !sheet_xml.contains("tableParts"),
            "worksheet XML must not contain tableParts after deletion"
        );
    }

    #[test]
    fn test_readfast_add_comment_then_save_no_duplicate() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("data".to_string()))
            .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Tester".to_string(),
                text: "Original comment".to_string(),
            },
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        // Open in Lazy mode, add a new comment, and save.
        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        wb2.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "B1".to_string(),
                author: "Tester".to_string(),
                text: "New comment".to_string(),
            },
        )
        .unwrap();
        // This must not fail with a duplicate ZIP entry error.
        let saved = wb2.save_to_buffer().unwrap();

        // Re-open and verify both old and new comments are present.
        let mut wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let comments = wb3.get_comments("Sheet1").unwrap();
        assert!(
            comments.iter().any(|c| c.text == "New comment"),
            "New comment should be present after Lazy + add_comment round-trip"
        );
        assert!(
            comments.iter().any(|c| c.text == "Original comment"),
            "Original comment must be preserved after Lazy + add_comment round-trip"
        );
        assert_eq!(
            comments.len(),
            2,
            "Both original and new comments must survive"
        );
    }

    #[test]
    fn test_readfast_add_comment_preserves_existing_comments() {
        // Regression test: opening with Lazy mode, adding a comment, and saving
        // must not drop pre-existing comments.
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("data".to_string()))
            .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Alice".to_string(),
                text: "First comment".to_string(),
            },
        )
        .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "B2".to_string(),
                author: "Bob".to_string(),
                text: "Second comment".to_string(),
            },
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Add a third comment.
        wb2.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "C3".to_string(),
                author: "Charlie".to_string(),
                text: "Third comment".to_string(),
            },
        )
        .unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        let mut wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let comments = wb3.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 3, "All three comments must be present");
        assert!(comments
            .iter()
            .any(|c| c.cell == "A1" && c.text == "First comment"));
        assert!(comments
            .iter()
            .any(|c| c.cell == "B2" && c.text == "Second comment"));
        assert!(comments
            .iter()
            .any(|c| c.cell == "C3" && c.text == "Third comment"));
    }

    #[test]
    fn test_readfast_get_comments_hydrates_deferred() {
        // get_comments should return deferred comments even if no mutation occurred.
        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "Deferred comment".to_string(),
            },
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // get_comments should hydrate and return the deferred comment.
        let comments = wb2.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell, "A1");
        assert_eq!(comments[0].text, "Deferred comment");
    }

    #[test]
    fn test_readfast_remove_comment_hydrates_first() {
        // remove_comment on a Lazy workbook must hydrate deferred comments,
        // then remove only the target comment, preserving others.
        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Alice".to_string(),
                text: "Keep me".to_string(),
            },
        )
        .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "B2".to_string(),
                author: "Bob".to_string(),
                text: "Remove me".to_string(),
            },
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        wb2.remove_comment("Sheet1", "B2").unwrap();

        let saved = wb2.save_to_buffer().unwrap();
        let mut wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let comments = wb3.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell, "A1");
        assert_eq!(comments[0].text, "Keep me");
    }

    #[test]
    fn test_readfast_add_comment_no_preexisting_comments() {
        // Adding a comment to a sheet that had no comments when opened in Lazy mode
        // must create proper relationships and content types on save, even when
        // deferred_parts is non-empty due to other auxiliary parts (e.g. doc props).
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("data".to_string()))
            .unwrap();
        // Add doc props so that Lazy mode will have non-empty deferred_parts.
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Trigger deferred".to_string()),
            ..Default::default()
        });
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        wb2.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Newcomer".to_string(),
                text: "Brand new comment".to_string(),
            },
        )
        .unwrap();

        let saved = wb2.save_to_buffer().unwrap();

        // Verify the comment is readable after re-open.
        let mut wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let comments = wb3.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].text, "Brand new comment");

        // Verify the ZIP contains the comment XML and VML parts.
        let reader = std::io::Cursor::new(&saved);
        let mut archive = zip::ZipArchive::new(reader).unwrap();
        assert!(
            archive.by_name("xl/comments1.xml").is_ok(),
            "comments1.xml must be present"
        );
        assert!(
            archive.by_name("xl/drawings/vmlDrawing1.vml").is_ok(),
            "vmlDrawing1.vml must be present for the comment"
        );
    }

    #[test]
    fn test_readfast_add_comment_vml_roundtrip() {
        // Verify that VML parts are correct after Lazy hydration + add comment.
        let mut wb = Workbook::new();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Original".to_string(),
                text: "Has VML".to_string(),
            },
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        wb2.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "B2".to_string(),
                author: "New".to_string(),
                text: "Also has VML".to_string(),
            },
        )
        .unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        // Verify VML part is present and references both cells.
        let reader = std::io::Cursor::new(&saved);
        let mut archive = zip::ZipArchive::new(reader).unwrap();
        assert!(archive.by_name("xl/drawings/vmlDrawing1.vml").is_ok());

        // Verify both comments survive a full open.
        let mut wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let comments = wb3.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 2);
    }

    #[test]
    fn test_readfast_set_doc_props_then_save_no_duplicate() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::Number(1.0))
            .unwrap();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Original Title".to_string()),
            ..Default::default()
        });
        let buf = wb.save_to_buffer().unwrap();

        // Open in Lazy mode, update doc props, and save.
        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        wb2.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Updated Title".to_string()),
            ..Default::default()
        });
        // This must not fail with a duplicate ZIP entry error.
        let saved = wb2.save_to_buffer().unwrap();

        // Re-open and verify the updated doc props.
        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        let props = wb3.get_doc_props();
        assert_eq!(props.title, Some("Updated Title".to_string()));
    }

    #[test]
    fn test_read_xml_part_from_reader_worksheet() {
        use sheetkit_xml::worksheet::WorksheetXml;
        let ws_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2"><v>42</v></c></row>
  </sheetData>
</worksheet>"#;
        let mut buf = Vec::new();
        {
            let cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            zip.start_file("test.xml", opts).unwrap();
            use std::io::Write;
            zip.write_all(ws_xml.as_bytes()).unwrap();
            zip.finish().unwrap();
        }
        let cursor = std::io::Cursor::new(&buf);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let ws: WorksheetXml = read_xml_part(&mut archive, "test.xml").unwrap();
        assert_eq!(ws.sheet_data.rows.len(), 2);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[0].cells[0].r, "A1");
        assert_eq!(ws.sheet_data.rows[1].r, 2);
        assert_eq!(ws.sheet_data.rows[1].cells[0].v, Some("42".to_string()));
    }

    #[test]
    fn test_read_xml_part_from_reader_sst() {
        use sheetkit_xml::shared_strings::Sst;
        let sst_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
</sst>"#;
        let mut buf = Vec::new();
        {
            let cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            zip.start_file("sst.xml", opts).unwrap();
            use std::io::Write;
            zip.write_all(sst_xml.as_bytes()).unwrap();
            zip.finish().unwrap();
        }
        let cursor = std::io::Cursor::new(&buf);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let sst: Sst = read_xml_part(&mut archive, "sst.xml").unwrap();
        assert_eq!(sst.count, Some(2));
        assert_eq!(sst.unique_count, Some(2));
        assert_eq!(sst.items.len(), 2);
        assert_eq!(sst.items[0].t.as_ref().unwrap().value, "Hello");
        assert_eq!(sst.items[1].t.as_ref().unwrap().value, "World");
    }

    #[test]
    fn test_read_xml_part_from_reader_large_worksheet() {
        use sheetkit_xml::worksheet::WorksheetXml;
        let mut ws_xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>"#,
        );
        for i in 1..=500 {
            ws_xml.push_str(&format!(
                "<row r=\"{i}\"><c r=\"A{i}\"><v>{}</v></c><c r=\"B{i}\"><v>{}</v></c></row>",
                i * 10,
                i * 20,
            ));
        }
        ws_xml.push_str("</sheetData></worksheet>");

        let mut buf = Vec::new();
        {
            let cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            zip.start_file("sheet.xml", opts).unwrap();
            use std::io::Write;
            zip.write_all(ws_xml.as_bytes()).unwrap();
            zip.finish().unwrap();
        }
        let cursor = std::io::Cursor::new(&buf);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let ws: WorksheetXml = read_xml_part(&mut archive, "sheet.xml").unwrap();
        assert_eq!(ws.sheet_data.rows.len(), 500);
        assert_eq!(ws.sheet_data.rows[0].r, 1);
        assert_eq!(ws.sheet_data.rows[0].cells[0].v, Some("10".to_string()));
        assert_eq!(ws.sheet_data.rows[499].r, 500);
        assert_eq!(
            ws.sheet_data.rows[499].cells[1].v,
            Some("10000".to_string())
        );
    }

    // -- Copy-on-write passthrough tests --

    #[test]
    fn test_lazy_open_save_without_modification_roundtrips() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Hello").unwrap();
        wb.set_cell_value("Sheet1", "B1", 42.0f64).unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = crate::workbook::open_options::OpenOptions::new()
            .read_mode(crate::workbook::open_options::ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // No modifications: save should use passthrough for the worksheet.
        let saved = wb2.save_to_buffer().unwrap();

        // Re-open in Eager mode and verify data integrity.
        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        assert_eq!(
            wb3.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Hello".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Sheet1", "B1").unwrap(),
            CellValue::Number(42.0)
        );
    }

    #[test]
    fn test_lazy_open_modify_one_sheet_passthroughs_others() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "First sheet").unwrap();
        wb.new_sheet("Sheet2").unwrap();
        wb.set_cell_value("Sheet2", "A1", "Second sheet").unwrap();
        wb.new_sheet("Sheet3").unwrap();
        wb.set_cell_value("Sheet3", "A1", "Third sheet").unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = crate::workbook::open_options::OpenOptions::new()
            .read_mode(crate::workbook::open_options::ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Only modify Sheet2; Sheet1 and Sheet3 should use passthrough.
        wb2.set_cell_value("Sheet2", "B1", "Modified").unwrap();

        // Verify dirty tracking.
        assert!(!wb2.is_sheet_dirty(0), "Sheet1 should not be dirty");
        assert!(wb2.is_sheet_dirty(1), "Sheet2 should be dirty");
        assert!(!wb2.is_sheet_dirty(2), "Sheet3 should not be dirty");

        let saved = wb2.save_to_buffer().unwrap();

        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        assert_eq!(
            wb3.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("First sheet".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Sheet2", "A1").unwrap(),
            CellValue::String("Second sheet".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Sheet2", "B1").unwrap(),
            CellValue::String("Modified".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Sheet3", "A1").unwrap(),
            CellValue::String("Third sheet".to_string())
        );
    }

    #[test]
    fn test_lazy_open_deferred_aux_parts_preserved() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "data").unwrap();
        wb.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Test Title".to_string()),
            creator: Some("Test Author".to_string()),
            ..Default::default()
        });
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Tester".to_string(),
                text: "A comment".to_string(),
            },
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = crate::workbook::open_options::OpenOptions::new()
            .read_mode(crate::workbook::open_options::ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Save without touching anything; deferred aux parts should be preserved.
        let saved = wb2.save_to_buffer().unwrap();

        let mut wb3 = Workbook::open_from_buffer(&saved).unwrap();
        assert_eq!(
            wb3.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("data".to_string())
        );
        let props = wb3.get_doc_props();
        assert_eq!(props.title.as_deref(), Some("Test Title"));
        assert_eq!(props.creator.as_deref(), Some("Test Author"));
        let comments = wb3.get_comments("Sheet1").unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].text, "A comment");
    }

    #[test]
    fn test_eager_open_save_preserves_all_data() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "data").unwrap();
        wb.set_cell_value("Sheet1", "B1", 42.0f64).unwrap();
        wb.new_sheet("Sheet2").unwrap();
        wb.set_cell_value("Sheet2", "A1", "sheet2").unwrap();
        let buf = wb.save_to_buffer().unwrap();

        // Eager mode (default): all sheets parsed at open time.
        let wb2 = Workbook::open_from_buffer(&buf).unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        assert_eq!(
            wb3.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("data".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Sheet1", "B1").unwrap(),
            CellValue::Number(42.0)
        );
        assert_eq!(
            wb3.get_cell_value("Sheet2", "A1").unwrap(),
            CellValue::String("sheet2".to_string())
        );
    }

    #[test]
    fn test_lazy_read_then_save_passthrough() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "value").unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = crate::workbook::open_options::OpenOptions::new()
            .read_mode(crate::workbook::open_options::ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Read the value (triggers hydration via worksheet_ref, not mutation).
        let val = wb2.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::String("value".to_string()));

        // Sheet was read but not modified, so it should NOT be dirty.
        assert!(!wb2.is_sheet_dirty(0));

        // Save should still use passthrough for the untouched sheet.
        let saved = wb2.save_to_buffer().unwrap();
        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        assert_eq!(
            wb3.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("value".to_string())
        );
    }

    #[test]
    fn test_cow_passthrough_with_styles_and_formulas() {
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
        wb.set_cell_value("Sheet1", "A1", "styled").unwrap();
        wb.set_cell_style("Sheet1", "A1", style_id).unwrap();
        wb.set_cell_formula("Sheet1", "B1", "LEN(A1)").unwrap();
        wb.new_sheet("Sheet2").unwrap();
        wb.set_cell_value("Sheet2", "A1", "other").unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = crate::workbook::open_options::OpenOptions::new()
            .read_mode(crate::workbook::open_options::ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let saved = wb2.save_to_buffer().unwrap();

        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        assert_eq!(
            wb3.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("styled".to_string())
        );
        assert_eq!(wb3.get_cell_style("Sheet1", "A1").unwrap(), Some(style_id));
        match wb3.get_cell_value("Sheet1", "B1").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "LEN(A1)"),
            other => panic!("expected formula, got {:?}", other),
        }
        assert_eq!(
            wb3.get_cell_value("Sheet2", "A1").unwrap(),
            CellValue::String("other".to_string())
        );
    }

    #[test]
    fn test_new_workbook_sheets_are_dirty() {
        let wb = Workbook::new();
        assert!(wb.is_sheet_dirty(0), "new workbook sheet should be dirty");
    }

    #[test]
    fn test_eager_open_sheets_start_clean() {
        use crate::workbook::open_options::{AuxParts, OpenOptions, ReadMode};

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "test").unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert!(
            !wb2.is_sheet_dirty(0),
            "eagerly parsed sheet should retain raw bytes and start clean"
        );
    }

    #[test]
    fn test_lazy_open_sheets_start_clean() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "test").unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = crate::workbook::open_options::OpenOptions::new()
            .read_mode(crate::workbook::open_options::ReadMode::Lazy);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert!(
            !wb2.is_sheet_dirty(0),
            "lazily deferred sheet should start clean"
        );
    }

    #[test]
    fn test_lazy_mutation_marks_dirty() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "test").unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = crate::workbook::open_options::OpenOptions::new()
            .read_mode(crate::workbook::open_options::ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert!(!wb2.is_sheet_dirty(0));

        wb2.set_cell_value("Sheet1", "B1", "new").unwrap();
        assert!(
            wb2.is_sheet_dirty(0),
            "sheet should be dirty after mutation"
        );
    }

    #[test]
    fn test_lazy_open_multi_sheet_selective_dirty() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "s1").unwrap();
        wb.new_sheet("Sheet2").unwrap();
        wb.set_cell_value("Sheet2", "A1", "s2").unwrap();
        wb.new_sheet("Sheet3").unwrap();
        wb.set_cell_value("Sheet3", "A1", "s3").unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = crate::workbook::open_options::OpenOptions::new()
            .read_mode(crate::workbook::open_options::ReadMode::Lazy);
        let mut wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // All sheets start clean.
        assert!(!wb2.is_sheet_dirty(0));
        assert!(!wb2.is_sheet_dirty(1));
        assert!(!wb2.is_sheet_dirty(2));

        // Read Sheet1 (no mutation).
        let _ = wb2.get_cell_value("Sheet1", "A1").unwrap();
        assert!(!wb2.is_sheet_dirty(0), "reading should not dirty a sheet");

        // Mutate Sheet3.
        wb2.set_cell_value("Sheet3", "B1", "modified").unwrap();
        assert!(!wb2.is_sheet_dirty(0));
        assert!(!wb2.is_sheet_dirty(1));
        assert!(wb2.is_sheet_dirty(2));

        // Save and verify all data.
        let saved = wb2.save_to_buffer().unwrap();
        let wb3 = Workbook::open_from_buffer(&saved).unwrap();
        assert_eq!(
            wb3.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("s1".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Sheet2", "A1").unwrap(),
            CellValue::String("s2".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Sheet3", "A1").unwrap(),
            CellValue::String("s3".to_string())
        );
        assert_eq!(
            wb3.get_cell_value("Sheet3", "B1").unwrap(),
            CellValue::String("modified".to_string())
        );
    }

    #[test]
    fn test_sheets_filter_preserves_filtered_sheet_with_comments_on_save() {
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("keep_me".to_string()))
            .unwrap();
        wb.set_cell_value("Sheet2", "A1", CellValue::String("s2".to_string()))
            .unwrap();
        wb.add_comment(
            "Sheet1",
            &crate::comment::CommentConfig {
                cell: "A1".to_string(),
                author: "Test".to_string(),
                text: "a comment".to_string(),
            },
        )
        .unwrap();
        let buf = wb.save_to_buffer().unwrap();

        let opts = OpenOptions::new().sheets(vec!["Sheet2".to_string()]);
        let wb2 = Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::Empty
        );

        let buf2 = wb2.save_to_buffer().unwrap();
        let opts_all = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let wb3 = Workbook::open_from_buffer_with_options(&buf2, &opts_all).unwrap();
        assert_eq!(
            wb3.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("keep_me".to_string()),
        );
        assert_eq!(
            wb3.get_cell_value("Sheet2", "A1").unwrap(),
            CellValue::String("s2".to_string()),
        );
    }

    #[test]
    fn eager_open_preserves_clean_graph_xml() {
        let (input, chart, drawing, drawing_rels) = opaque_graph_buffer();
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut workbook = Workbook::open_from_buffer_with_options(&input, &options).unwrap();

        assert_eq!(workbook.raw_charts.len(), 1);
        assert!(!workbook
            .raw_graph_parts
            .contains_key("xl/charts/chart1.xml"));
        assert_eq!(
            workbook.drawings.len(),
            1,
            "drawing must parse successfully"
        );
        workbook
            .set_cell_value("Sheet1", "A20", "unrelated edit")
            .unwrap();
        assert!(workbook.dirty_graph_parts.is_empty());

        let saved = workbook.save_to_buffer().unwrap();
        assert_eq!(zip_part(&saved, "xl/charts/chart1.xml"), chart);
        assert_eq!(zip_part(&saved, "xl/drawings/drawing1.xml"), drawing);
        assert_eq!(
            zip_part(&saved, "xl/drawings/_rels/drawing1.xml.rels"),
            drawing_rels
        );
    }

    #[test]
    fn clean_typed_chart_owner_prefers_its_raw_original() {
        let base = workbook_with_chart_buffer();
        let raw_chart = zip_part(&base, "xl/charts/chart1.xml")
            .into_iter()
            .chain(b"\n".iter().copied())
            .collect::<Vec<_>>();
        let mut workbook = Workbook::new();
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        workbook
            .add_chart(
                "Sheet1",
                "E2",
                "L15",
                &ChartConfig {
                    chart_type: ChartType::Col,
                    title: None,
                    series: vec![ChartSeries {
                        name: "Values".to_string(),
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
        workbook.remember_raw_graph_part("xl/charts/chart1.xml".to_string(), raw_chart.clone());
        workbook.dirty_graph_parts.remove("xl/charts/chart1.xml");

        let saved = workbook.save_to_buffer().unwrap();
        assert_eq!(zip_part(&saved, "xl/charts/chart1.xml"), raw_chart);
    }

    #[test]
    fn raw_graph_paths_are_reserved_for_part_allocation() {
        let mut workbook = Workbook::new();
        workbook.remember_raw_graph_part("xl/charts/chart1.xml".to_string(), b"chart".to_vec());
        workbook
            .remember_raw_graph_part("xl/drawings/drawing1.xml".to_string(), b"drawing".to_vec());

        assert_eq!(
            workbook.next_available_part_path("xl/charts/chart", ".xml"),
            "xl/charts/chart2.xml"
        );
        assert_eq!(
            workbook.next_available_part_path("xl/drawings/drawing", ".xml"),
            "xl/drawings/drawing2.xml"
        );
    }

    #[test]
    fn lazy_hydration_is_read_only_for_graph_parts() {
        let (input, chart, drawing, drawing_rels) = opaque_graph_buffer();
        let options = OpenOptions::new().read_mode(ReadMode::Lazy);
        let mut workbook = Workbook::open_from_buffer_with_options(&input, &options).unwrap();

        workbook.hydrate_drawings();
        assert_eq!(workbook.raw_charts.len(), 1);
        assert!(!workbook
            .raw_graph_parts
            .contains_key("xl/charts/chart1.xml"));
        assert_eq!(
            workbook.drawings.len(),
            1,
            "drawing must hydrate successfully"
        );
        assert!(workbook.dirty_graph_parts.is_empty());

        let saved = workbook.save_to_buffer().unwrap();
        assert_eq!(zip_part(&saved, "xl/charts/chart1.xml"), chart);
        assert_eq!(zip_part(&saved, "xl/drawings/drawing1.xml"), drawing);
        assert_eq!(
            zip_part(&saved, "xl/drawings/_rels/drawing1.xml.rels"),
            drawing_rels
        );
    }

    #[test]
    fn eager_open_preserves_parse_failed_drawing_and_relationships() {
        let base = workbook_with_chart_buffer();
        let drawing = b"<xdr:wsDr><broken></xdr:wsDr>".to_vec();
        let drawing_rels = b"<Relationships><broken></Relationships>".to_vec();
        let input = rewrite_zip_parts(
            &base,
            &[
                ("xl/drawings/drawing1.xml", Some(drawing.clone())),
                (
                    "xl/drawings/_rels/drawing1.xml.rels",
                    Some(drawing_rels.clone()),
                ),
            ],
        );

        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut workbook = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
        workbook
            .set_cell_value("Sheet1", "A20", "unrelated edit")
            .unwrap();
        let saved = workbook.save_to_buffer().unwrap();

        assert_eq!(zip_part(&saved, "xl/drawings/drawing1.xml"), drawing);
        assert_eq!(
            zip_part(&saved, "xl/drawings/_rels/drawing1.xml.rels"),
            drawing_rels
        );
    }

    #[test]
    fn eager_open_preserves_unsupported_chart_and_drawing_nodes() {
        let (input, chart, drawing) = unsupported_graph_buffer();
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut workbook = Workbook::open_from_buffer_with_options(&input, &options).unwrap();

        workbook
            .set_cell_value("Sheet1", "A20", "unrelated edit")
            .unwrap();
        let saved = workbook.save_to_buffer().unwrap();

        assert_eq!(zip_part(&saved, "xl/charts/chart1.xml"), chart);
        assert_eq!(zip_part(&saved, "xl/drawings/drawing1.xml"), drawing);
    }

    #[test]
    fn mutating_drawing_keeps_existing_chart_raw_xml_clean() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};

        let (input, chart, drawing, _) = opaque_graph_buffer();
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut workbook = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
        workbook
            .add_chart(
                "Sheet1",
                "E20",
                "L30",
                &ChartConfig {
                    chart_type: ChartType::Line,
                    title: None,
                    series: vec![ChartSeries {
                        name: "New values".to_string(),
                        categories: "Sheet1!$A$1:$A$3".to_string(),
                        values: "Sheet1!$C$1:$C$3".to_string(),
                        x_values: None,
                        bubble_sizes: None,
                    }],
                    show_legend: false,
                    view_3d: None,
                },
            )
            .unwrap();
        assert!(!workbook.dirty_graph_parts.contains("xl/charts/chart1.xml"));
        assert!(workbook.dirty_graph_parts.contains("xl/charts/chart2.xml"));
        assert!(workbook
            .dirty_graph_parts
            .contains("xl/drawings/drawing1.xml"));
        assert!(workbook
            .dirty_graph_parts
            .contains("xl/drawings/_rels/drawing1.xml.rels"));

        let saved = workbook.save_to_buffer().unwrap();
        assert_eq!(zip_part(&saved, "xl/charts/chart1.xml"), chart);
        assert_ne!(zip_part(&saved, "xl/drawings/drawing1.xml"), drawing);
        assert!(zip_entry_count(&saved, "xl/charts/chart2.xml") == 1);
    }

    #[test]
    fn graph_relationship_mutations_reject_unparseable_raw_rels_atomically() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        use crate::image::{ImageConfig, ImageFormat};

        let base = workbook_with_chart_buffer();
        let input = rewrite_zip_parts(
            &base,
            &[(
                "xl/drawings/_rels/drawing1.xml.rels",
                Some(b"<Relationships><opaque></Relationships>".to_vec()),
            )],
        );
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut workbook = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
        assert_eq!(workbook.drawings.len(), 1);
        assert!(!workbook.drawing_rels.contains_key(&0));
        let before = workbook.save_to_buffer().unwrap();

        let chart_error = workbook
            .add_chart(
                "Sheet1",
                "E20",
                "L30",
                &ChartConfig {
                    chart_type: ChartType::Line,
                    title: None,
                    series: vec![ChartSeries {
                        name: "New values".to_string(),
                        categories: "Sheet1!$A$1:$A$3".to_string(),
                        values: "Sheet1!$C$1:$C$3".to_string(),
                        x_values: None,
                        bubble_sizes: None,
                    }],
                    show_legend: false,
                    view_3d: None,
                },
            )
            .unwrap_err();
        assert!(matches!(chart_error, Error::InvalidArgument(_)));
        assert_eq!(workbook.save_to_buffer().unwrap(), before);

        let image_error = workbook
            .add_image(
                "Sheet1",
                &ImageConfig {
                    data: vec![0x89, 0x50, 0x4e, 0x47],
                    format: ImageFormat::Png,
                    from_cell: "A20".to_string(),
                    width_px: 10,
                    height_px: 10,
                },
            )
            .unwrap_err();
        assert!(matches!(image_error, Error::InvalidArgument(_)));
        assert_eq!(workbook.save_to_buffer().unwrap(), before);
    }

    #[test]
    fn delete_sheet_rejects_owned_unparseable_raw_drawing_atomically() {
        let base = workbook_with_chart_buffer();
        let opaque_drawing = b"<xdr:wsDr><opaque></xdr:wsDr>".to_vec();
        let input = rewrite_zip_parts(
            &base,
            &[("xl/drawings/drawing1.xml", Some(opaque_drawing.clone()))],
        );
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut workbook = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
        assert!(!workbook.worksheet_drawings.contains_key(&0));
        let before = workbook.save_to_buffer().unwrap();

        let error = workbook.delete_sheet("Sheet1").unwrap_err();
        assert!(matches!(error, Error::InvalidArgument(_)));
        assert_eq!(workbook.sheet_names(), vec!["Sheet1", "Sheet2"]);
        let after = workbook.save_to_buffer().unwrap();
        assert_eq!(after, before);
        assert_eq!(zip_part(&after, "xl/drawings/drawing1.xml"), opaque_drawing);
        assert_eq!(zip_entry_count(&after, "xl/drawings/drawing1.xml"), 1);
    }

    #[test]
    fn eager_edit_preserves_untouched_worksheet_bytes_with_extensions() {
        let mut workbook = Workbook::new();
        workbook.new_sheet("Untouched").unwrap();
        let base = workbook.save_to_buffer().unwrap();
        let sheet_path = "xl/worksheets/sheet2.xml";
        let raw_sheet = String::from_utf8(zip_part(&base, sheet_path)).unwrap();
        let raw_sheet = raw_sheet.replacen(
            "</worksheet>",
            concat!(
                "<colBreaks count=\"1\" manualBreakCount=\"1\">",
                "<brk id=\"3\" min=\"0\" max=\"1048575\" man=\"true\"/>",
                "</colBreaks>",
                "<extLst><ext uri=\"{opaque-sheet-extension}\">",
                "<opaque:payload xmlns:opaque=\"urn:sheetkit:test\" value=\"keep\"/>",
                "</ext></extLst></worksheet>"
            ),
            1,
        );
        let raw_sheet = raw_sheet.into_bytes();
        let input = rewrite_zip_parts(&base, &[(sheet_path, Some(raw_sheet.clone()))]);
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut opened = Workbook::open_from_buffer_with_options(&input, &options).unwrap();

        assert!(!opened.is_sheet_dirty(0));
        assert!(!opened.is_sheet_dirty(1));
        opened.set_cell_value("Sheet1", "A1", "changed").unwrap();
        let saved = opened.save_to_buffer().unwrap();

        assert_eq!(zip_part(&saved, sheet_path), raw_sheet);
        assert_eq!(zip_entry_count(&saved, sheet_path), 1);
        let saved_sheet = String::from_utf8(zip_part(&saved, sheet_path)).unwrap();
        assert_eq!(saved_sheet.matches("<extLst>").count(), 1);
    }

    #[test]
    fn eager_sparkline_removal_does_not_reuse_stale_raw_sheet_xml() {
        let mut workbook = Workbook::new();
        workbook
            .add_sparkline(
                "Sheet1",
                &crate::sparkline::SparklineConfig::new("Sheet1!A1:A3", "B1"),
            )
            .unwrap();
        let base = workbook.save_to_buffer().unwrap();
        assert!(
            String::from_utf8(zip_part(&base, "xl/worksheets/sheet1.xml"))
                .unwrap()
                .contains("sparklineGroups")
        );

        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let mut opened = Workbook::open_from_buffer_with_options(&base, &options).unwrap();
        opened.remove_sparkline("Sheet1", "B1").unwrap();
        let saved = opened.save_to_buffer().unwrap();
        let saved_sheet = String::from_utf8(zip_part(&saved, "xl/worksheets/sheet1.xml")).unwrap();

        assert!(!saved_sheet.contains("sparklineGroups"));
        assert!(!saved_sheet.contains("<extLst>"));
        assert_eq!(zip_entry_count(&saved, "xl/worksheets/sheet1.xml"), 1);
    }

    #[test]
    fn dirty_sheet_extras_follow_worksheet_schema_order() {
        let mut workbook = Workbook::new();
        workbook
            .add_table(
                "Sheet1",
                &crate::table::TableConfig {
                    name: "Table1".to_string(),
                    display_name: "Table1".to_string(),
                    range: "A1:B3".to_string(),
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
                    ..crate::table::TableConfig::default()
                },
            )
            .unwrap();
        workbook
            .add_comment(
                "Sheet1",
                &crate::comment::CommentConfig {
                    cell: "A1".to_string(),
                    author: "Author".to_string(),
                    text: "Comment".to_string(),
                },
            )
            .unwrap();
        workbook
            .add_sparkline(
                "Sheet1",
                &crate::sparkline::SparklineConfig::new("Sheet1!B1:B3", "C1"),
            )
            .unwrap();

        let saved = workbook.save_to_buffer().unwrap();
        let sheet = String::from_utf8(zip_part(&saved, "xl/worksheets/sheet1.xml")).unwrap();
        let legacy = sheet.find("<legacyDrawing ").unwrap();
        let tables = sheet.find("<tableParts ").unwrap();
        let extensions = sheet.find("<extLst>").unwrap();
        let closing = sheet.rfind("</worksheet>").unwrap();

        assert!(legacy < tables, "legacyDrawing must precede tableParts");
        assert!(tables < extensions, "tableParts must precede extLst");
        assert!(
            extensions < closing,
            "extLst must be the final worksheet child"
        );
    }

    #[test]
    fn workbook_xml_uses_raw_bytes_until_its_typed_owner_changes() {
        let base = Workbook::new().save_to_buffer().unwrap();
        let raw_workbook = String::from_utf8(zip_part(&base, "xl/workbook.xml")).unwrap();
        let raw_workbook = raw_workbook
            .replacen(
                "<workbookPr",
                concat!(
                    "<fileSharing readOnlyRecommended=\"true\" userName=\"owner\"/>",
                    "<workbookPr"
                ),
                1,
            )
            .replacen(
                "</sheets>",
                concat!(
                    "</sheets><externalReferences>",
                    "<externalReference r:id=\"rIdExternal\"/>",
                    "</externalReferences>"
                ),
                1,
            )
            .replacen(
                "</workbook>",
                concat!(
                    "<oleSize ref=\"A1:C3\"/>",
                    "<extLst><ext uri=\"{opaque-workbook-extension}\">",
                    "<opaque:payload xmlns:opaque=\"urn:sheetkit:test\" value=\"keep\"/>",
                    "</ext></extLst></workbook>"
                ),
                1,
            );
        assert!(raw_workbook.contains("fileSharing"));
        let raw_workbook = raw_workbook.into_bytes();
        let input = rewrite_zip_parts(&base, &[("xl/workbook.xml", Some(raw_workbook.clone()))]);
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);

        let mut cell_edit = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
        assert_eq!(
            cell_edit
                .workbook_xml
                .file_sharing
                .as_ref()
                .and_then(|sharing| sharing.user_name.as_deref()),
            Some("owner")
        );
        assert_eq!(
            cell_edit
                .workbook_xml
                .ole_size
                .as_ref()
                .map(|size| size.reference.as_str()),
            Some("A1:C3")
        );
        assert_eq!(
            cell_edit
                .workbook_xml
                .external_references
                .as_ref()
                .and_then(|references| references.references.first())
                .map(|reference| reference.r_id.as_str()),
            Some("rIdExternal")
        );
        cell_edit.set_cell_value("Sheet1", "A1", "changed").unwrap();
        let cell_saved = cell_edit.save_to_buffer().unwrap();
        assert_eq!(zip_part(&cell_saved, "xl/workbook.xml"), raw_workbook);
        assert_eq!(zip_entry_count(&cell_saved, "xl/workbook.xml"), 1);

        let mut owner_edit = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
        owner_edit.new_sheet("Added").unwrap();
        let owner_saved = owner_edit.save_to_buffer().unwrap();
        let owner_xml = zip_part(&owner_saved, "xl/workbook.xml");
        assert_ne!(owner_xml, raw_workbook);
        let owner_xml = String::from_utf8(owner_xml).unwrap();
        assert!(owner_xml.contains("name=\"Added\""));
        assert!(owner_xml.contains("r:id=\"rIdExternal\""));
        assert_eq!(zip_entry_count(&owner_saved, "xl/workbook.xml"), 1);
    }

    fn opaque_doc_props_buffer() -> (Vec<u8>, Vec<u8>, Vec<u8>, Vec<u8>) {
        let mut workbook = Workbook::new();
        workbook.set_doc_props(crate::doc_props::DocProperties {
            title: Some("Original".to_string()),
            ..Default::default()
        });
        workbook.set_app_props(crate::doc_props::AppProperties {
            application: Some("Original app".to_string()),
            ..Default::default()
        });
        workbook.set_custom_property(
            "Original",
            crate::doc_props::CustomPropertyValue::String("value".to_string()),
        );
        let base = workbook.save_to_buffer().unwrap();
        let core = b"<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"><broken></cp:coreProperties>".to_vec();
        let app = b"<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><broken></Properties>".to_vec();
        let custom = concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" ",
            "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">",
            "<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"Empty\">",
            "<vt:empty/></property></Properties>"
        )
        .as_bytes()
        .to_vec();
        let input = rewrite_zip_parts(
            &base,
            &[
                ("docProps/core.xml", Some(core.clone())),
                ("docProps/app.xml", Some(app.clone())),
                ("docProps/custom.xml", Some(custom.clone())),
            ],
        );
        (input, core, app, custom)
    }

    fn doc_props_options(read_mode: ReadMode) -> OpenOptions {
        let options = OpenOptions::new().read_mode(read_mode);
        if read_mode == ReadMode::Eager {
            options.aux_parts(AuxParts::EagerLoad)
        } else {
            options
        }
    }

    #[test]
    fn eager_and_lazy_noop_preserve_each_opaque_doc_props_part() {
        let (input, core, app, custom) = opaque_doc_props_buffer();
        for read_mode in [ReadMode::Eager, ReadMode::Lazy] {
            let opened =
                Workbook::open_from_buffer_with_options(&input, &doc_props_options(read_mode))
                    .unwrap();
            let saved = opened.save_to_buffer().unwrap();

            for (path, expected) in [
                ("docProps/core.xml", core.as_slice()),
                ("docProps/app.xml", app.as_slice()),
                ("docProps/custom.xml", custom.as_slice()),
            ] {
                assert_eq!(zip_part(&saved, path), expected, "{read_mode:?}: {path}");
                assert_eq!(zip_entry_count(&saved, path), 1, "{read_mode:?}: {path}");
            }
        }
    }

    #[test]
    fn eager_and_lazy_doc_props_mutations_only_rewrite_the_target_part() {
        let (input, core, app, custom) = opaque_doc_props_buffer();
        for read_mode in [ReadMode::Eager, ReadMode::Lazy] {
            let options = doc_props_options(read_mode);

            let mut core_edit = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
            core_edit.set_doc_props(crate::doc_props::DocProperties {
                title: Some("Updated".to_string()),
                ..Default::default()
            });
            let saved = core_edit.save_to_buffer().unwrap();
            assert_ne!(zip_part(&saved, "docProps/core.xml"), core);
            assert_eq!(zip_part(&saved, "docProps/app.xml"), app);
            assert_eq!(zip_part(&saved, "docProps/custom.xml"), custom);

            let mut app_edit = Workbook::open_from_buffer_with_options(&input, &options).unwrap();
            app_edit.set_app_props(crate::doc_props::AppProperties {
                application: Some("Updated app".to_string()),
                ..Default::default()
            });
            let saved = app_edit.save_to_buffer().unwrap();
            assert_eq!(zip_part(&saved, "docProps/core.xml"), core);
            assert_ne!(zip_part(&saved, "docProps/app.xml"), app);
            assert_eq!(zip_part(&saved, "docProps/custom.xml"), custom);

            let mut custom_edit =
                Workbook::open_from_buffer_with_options(&input, &options).unwrap();
            custom_edit
                .set_custom_property("Updated", crate::doc_props::CustomPropertyValue::Bool(true));
            let saved = custom_edit.save_to_buffer().unwrap();
            assert_eq!(zip_part(&saved, "docProps/core.xml"), core);
            assert_eq!(zip_part(&saved, "docProps/app.xml"), app);
            assert_ne!(zip_part(&saved, "docProps/custom.xml"), custom);
            for path in [
                "docProps/core.xml",
                "docProps/app.xml",
                "docProps/custom.xml",
            ] {
                assert_eq!(zip_entry_count(&saved, path), 1, "{read_mode:?}: {path}");
            }
        }
    }
}
