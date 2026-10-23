use super::*;
use crate::workbook::budget::OpenBudget;
use crate::workbook::open_options::OpenOptions;
use crate::workbook::package_reader::{ArchiveAccess, PackageReader};
use crate::workbook::source::{open_package_archive, PackageInput};

/// VBA project relationship type URI.
const VBA_PROJECT_REL_TYPE: &str =
    "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

/// VBA project content type.
const VBA_PROJECT_CONTENT_TYPE: &str = "application/vnd.ms-office.vbaProject";

fn is_canonical_vba_package_path(path: &str) -> bool {
    path.trim_start_matches('/')
        .eq_ignore_ascii_case("xl/vbaProject.bin")
}

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
            sheet_sparklines_hydrated: vec![true],
            sheet_vml: vec![None],
            unknown_parts: vec![],
            deferred_parts: crate::workbook::aux::DeferredAuxParts::new(),
            vba_blob: None,
            tables: vec![],
            raw_sheet_xml: vec![None],
            raw_sheet_extensions: vec![None],
            sheet_dirty: vec![true],
            sheet_included: vec![true],
            sheet_materialized: vec![true],
            sheet_package_backed: vec![false],
            slicer_defs: vec![],
            slicer_caches: vec![],
            sheet_threaded_comments: vec![None],
            person_list: sheetkit_xml::threaded_comment::PersonList::default(),
            sheet_form_controls: vec![vec![]],
            sheet_controls_dirty: vec![false],
            streamed_sheets: HashMap::new(),
            package_source: None,
            read_mode: ReadMode::default(),
            sheet_rows_limit: None,
            worksheet_spool_limit: OpenLimits::default()
                .xml_bytes
                .min(OpenLimits::default().entry_actual_bytes),
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
        let outer_size = file.metadata()?.len();
        let limits = options.resolved_limits()?;
        let mut budget = OpenBudget::new(limits)?;
        budget.check_outer_input(outer_size)?;
        // One central-directory parse serves both open and later lazy reads.
        let package = open_package_archive(PackageInput::file(file), limits)?;
        let mut archive = PackageReader::from_archive(package.clone(), &mut budget)?;
        let mut wb = Self::from_archive(&mut archive, options)?;
        wb.package_source = Some(PackageSource::new(package, limits));
        wb.read_mode = options.read_mode;
        Ok(wb)
    }

    /// Build a Workbook from an already-opened ZIP archive.
    fn from_archive<A: ArchiveAccess>(archive: &mut A, options: &OpenOptions) -> Result<Self> {
        let resolved_limits = options.resolved_limits()?;
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

        let package_paths: HashSet<String> = archive.entry_names().into_iter().collect();
        validate_package_part_names(&package_paths)?;

        // Every worksheet must own a distinct internal package part in every
        // read mode, so lazy saves cannot alias or escape the package.
        let sheet_count = workbook_xml.sheets.sheets.len();
        let worksheet_paths = resolve_worksheet_paths(
            &workbook_xml,
            &workbook_rels,
            &content_types,
            &package_paths,
        )?;

        // Parse per-sheet worksheet relationship files (optional).
        // Always loaded: needed for hyperlinks, on-demand comment loading, etc.
        let mut worksheet_rels: HashMap<usize, Relationships> = HashMap::with_capacity(sheet_count);
        for (i, sheet_path) in worksheet_paths.iter().enumerate() {
            let rels_path = relationship_part_path(sheet_path);
            if let Some(rels) = read_optional_xml_part::<Relationships, _>(archive, &rels_path)? {
                worksheet_rels.insert(i, rels);
                known_paths.insert(rels_path);
            }
        }
        validate_table_relationship_ownership(
            &worksheet_paths,
            &worksheet_rels,
            &content_types,
            &package_paths,
        )?;

        let skip_aux = options.skip_aux_parts();
        let mut worksheets: Vec<(String, OnceLock<WorksheetXml>)> = Vec::with_capacity(sheet_count);
        let mut raw_sheet_xml: Vec<Option<Vec<u8>>> = Vec::with_capacity(sheet_count);
        let mut sheet_included: Vec<bool> = Vec::with_capacity(sheet_count);
        let mut sheet_materialized: Vec<bool> = Vec::with_capacity(sheet_count);
        let mut eager_sparklines: Vec<Vec<crate::sparkline::SparklineConfig>> =
            vec![vec![]; sheet_count];

        let defer_sheets = !options.is_eager();

        for (sheet_idx, sheet_entry) in workbook_xml.sheets.sheets.iter().enumerate() {
            let sheet_path = &worksheet_paths[sheet_idx];
            let should_parse = options.should_parse_sheet(&sheet_entry.name);
            sheet_included.push(should_parse);

            if should_parse && !defer_sheets {
                // Eager mode + selected: parse immediately. Clean save
                // passthrough uses the pinned package source, so no second
                // sheet-sized raw allocation is retained. Table and sparkline
                // checks reuse a spool captured by the same decompression.
                let table_ids = table_relationship_ids(worksheet_rels.get(&sheet_idx));
                let capture = !skip_aux || !table_ids.is_empty();
                let (ws, spool) =
                    read_eager_worksheet(archive, sheet_path, options.sheet_rows, capture)?;
                if let Some(mut spool) = spool {
                    if !table_ids.is_empty() {
                        validate_worksheet_table_part_stream(
                            std::io::BufReader::with_capacity(LARGE_BUF_CAPACITY, &mut spool),
                            &table_ids,
                            sheet_path,
                        )?;
                    }
                    if !skip_aux {
                        use std::io::Seek as _;
                        spool.seek(std::io::SeekFrom::Start(0))?;
                        if let Some(fragment) =
                            crate::workbook::source::extract_worksheet_extension_fragment(
                                std::io::BufReader::new(spool),
                            )?
                        {
                            eager_sparklines[sheet_idx] =
                                parse_sparklines_from_xml(&String::from_utf8_lossy(&fragment));
                        }
                    }
                }
                worksheets.push((sheet_entry.name.clone(), initialized_lock(ws)));
                sheet_materialized.push(true);
                raw_sheet_xml.push(None);
            } else if !should_parse {
                // Filtered out (any mode): do not decompress or retain the
                // worksheet. Keep an empty typed placeholder for queries; a
                // clean save raw-copies the original compressed entry.
                worksheets.push((
                    sheet_entry.name.clone(),
                    initialized_lock(WorksheetXml::default()),
                ));
                sheet_materialized.push(false);
                raw_sheet_xml.push(None);
            } else {
                // Lazy/Stream mode + selected: keep the sheet package-backed
                // until first access. OnceLock is left empty.
                worksheets.push((sheet_entry.name.clone(), OnceLock::new()));
                sheet_materialized.push(false);
                raw_sheet_xml.push(None);
            }
            known_paths.insert(sheet_path.clone());
        }

        // Deferred sheets are validated directly from the archive; eagerly
        // parsed sheets were validated from their capture spool above.
        let archive_table_checks: Vec<bool> = sheet_included
            .iter()
            .map(|included| *included && defer_sheets)
            .collect();
        validate_worksheet_table_parts(
            archive,
            &worksheet_paths,
            &worksheet_rels,
            &archive_table_checks,
        )?;

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
        let (theme_xml, theme_colors) =
            match read_optional_bytes_part(archive, "xl/theme/theme1.xml")? {
                Some(bytes) => {
                    let colors = sheetkit_xml::theme::parse_theme_colors(&bytes);
                    (Some(bytes), colors)
                }
                None => (None, crate::theme::default_theme_colors()),
            };
        known_paths.insert("xl/theme/theme1.xml".to_string());

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
        let sheet_sparklines = eager_sparklines;
        let sheet_sparklines_hydrated = vec![!skip_aux; worksheets.len()];
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
                    if let Some(comments) =
                        read_optional_xml_part::<Comments, _>(archive, &comment_path)?
                    {
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
                    if let Some(bytes) = read_optional_bytes_part(archive, &vml_path)? {
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
                        let Some(bytes) = read_optional_bytes_part(archive, &drawing_path)? else {
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
                if let Some(bytes) = read_optional_bytes_part(archive, &drawing_path)? {
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
                let Some(rels_bytes) = read_optional_bytes_part(archive, &drawing_rels_path)?
                else {
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
                            let Some(bytes) = read_optional_bytes_part(archive, &chart_path)?
                            else {
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
                            if let Some(bytes) = read_optional_bytes_part(archive, &image_path)? {
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
                    let Some(bytes) = read_optional_bytes_part(archive, &chart_path)? else {
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
            if let Some(bytes) = read_optional_bytes_part(archive, "docProps/core.xml")? {
                let xml = String::from_utf8_lossy(&bytes);
                core_properties = sheetkit_xml::doc_props::deserialize_core_properties(&xml).ok();
                raw_doc_props.insert("docProps/core.xml".to_string(), bytes);
                known_paths.insert("docProps/core.xml".to_string());
            }
            if let Some(bytes) = read_optional_bytes_part(archive, "docProps/app.xml")? {
                app_properties = deserialize_xml_bytes(&bytes).ok();
                raw_doc_props.insert("docProps/app.xml".to_string(), bytes);
                known_paths.insert("docProps/app.xml".to_string());
            }
            if let Some(bytes) = read_optional_bytes_part(archive, "docProps/custom.xml")? {
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
                    if let Some(pcd) = read_optional_xml_part::<
                        sheetkit_xml::pivot_cache::PivotCacheDefinition,
                        _,
                    >(archive, path)?
                    {
                        known_paths.insert(path.to_string());
                        pivot_cache_defs.push((path.to_string(), pcd));
                    }
                } else if ovr.content_type == mime_types::PIVOT_TABLE {
                    if let Some(pt) = read_optional_xml_part::<
                        sheetkit_xml::pivot_table::PivotTableDefinition,
                        _,
                    >(archive, path)?
                    {
                        known_paths.insert(path.to_string());
                        pivot_tables.push((path.to_string(), pt));
                    }
                } else if ovr.content_type == mime_types::PIVOT_CACHE_RECORDS {
                    if let Some(pcr) = read_optional_xml_part::<
                        sheetkit_xml::pivot_cache::PivotCacheRecords,
                        _,
                    >(archive, path)?
                    {
                        known_paths.insert(path.to_string());
                        pivot_cache_records.push((path.to_string(), pcr));
                    }
                }
            }

            // Parse slicer definitions and slicer cache definitions.
            for ovr in &content_types.overrides {
                let path = ovr.part_name.trim_start_matches('/');
                if ovr.content_type == mime_types::SLICER {
                    if let Some(sd) = read_optional_xml_part::<
                        sheetkit_xml::slicer::SlicerDefinitions,
                        _,
                    >(archive, path)?
                    {
                        known_paths.insert(path.to_string());
                        slicer_defs.push((path.to_string(), sd));
                    }
                } else if ovr.content_type == mime_types::SLICER_CACHE {
                    if let Some(raw) = read_optional_part(archive, path, true, false)? {
                        if let Some(scd) = std::str::from_utf8(&raw)
                            .ok()
                            .and_then(sheetkit_xml::slicer::parse_slicer_cache)
                        {
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
                    if let Some(tc) = read_optional_xml_part::<
                        sheetkit_xml::threaded_comment::ThreadedComments,
                        _,
                    >(archive, &tc_path)?
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
                    if let Some(pl) = read_optional_xml_part::<
                        sheetkit_xml::threaded_comment::PersonList,
                        _,
                    >(archive, &person_path)?
                    {
                        known_paths.insert(person_path);
                        found = Some(pl);
                    }
                }
                if found.is_none() {
                    if let Some(pl) = read_optional_xml_part::<
                        sheetkit_xml::threaded_comment::PersonList,
                        _,
                    >(archive, "xl/persons/person.xml")?
                    {
                        known_paths.insert("xl/persons/person.xml".to_string());
                        found = Some(pl);
                    }
                }
                found.unwrap_or_default()
            };

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
                    // parse_table_xml keeps extLst/sortState and the source
                    // snapshot used for byte-identical passthrough.
                    let parsed = read_optional_part(archive, &table_path, true, false)?
                        .and_then(|bytes| String::from_utf8(bytes).ok())
                        .and_then(|xml| sheetkit_xml::table::parse_table_xml(&xml).ok());
                    if let Some(table_xml) = parsed {
                        known_paths.insert(table_path.clone());
                        tables.push((table_path, table_xml, sheet_idx));
                    }
                }
            }
        }

        // VBA access and preservation must not depend on eager/lazy auxiliary
        // parsing. Apply the independent VBA input limit before retaining it.
        let mut vba_candidates: Vec<String> = workbook_rels
            .relationships
            .iter()
            .filter(|relationship| relationship.rel_type == VBA_PROJECT_REL_TYPE)
            .map(|relationship| {
                resolve_relationship_target("xl/workbook.xml", &relationship.target)
            })
            .collect();
        if !vba_candidates
            .iter()
            .any(|path| is_canonical_vba_package_path(path))
        {
            vba_candidates.push("xl/vbaProject.bin".to_string());
        }
        for path in vba_candidates {
            if let Some(bytes) = read_optional_part(archive, &path, false, true)? {
                crate::vba::validate_vba_input(&bytes, options.vba_limits)?;
                known_paths.insert(path);
                vba_blob = Some(bytes);
                break;
            }
        }

        let sheet_form_controls: Vec<Vec<crate::control::FormControlConfig>> =
            vec![vec![]; worksheets.len()];
        let sheet_controls_dirty = vec![false; worksheets.len()];

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
        for name in archive.entry_names() {
            if !known_paths.contains(&name) {
                // Every remaining entry is preserved, so a corrupt one must
                // fail the open instead of disappearing on the next save.
                let bytes = read_bytes_part(archive, &name)?;
                if skip_aux && crate::workbook::aux::classify_deferred_path(&name).is_some() {
                    deferred_parts.insert(name, bytes);
                } else {
                    unknown_parts.push((name, bytes));
                }
            }
        }

        // Validate and index all eagerly-parsed cells before applying row
        // limits or using the cached coordinates for binary search.
        // Deferred sheets (empty OnceLock) are skipped here; they are
        // post-processed on demand in `deserialize_worksheet_xml`.
        for (_name, ws_lock) in &mut worksheets {
            let Some(ws) = ws_lock.get_mut() else {
                continue;
            };
            crate::utils::cell_ref::validate_and_index_worksheet(ws)?;

            // Apply sheet_rows limit: keep only the first N rows.
            if let Some(max_rows) = options.sheet_rows {
                ws.sheet_data.rows.truncate(max_rows as usize);
            }
        }

        if let Some(error) = archive.pending_error() {
            return Err(error);
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
            sheet_sparklines_hydrated,
            sheet_vml,
            unknown_parts,
            deferred_parts,
            vba_blob,
            tables,
            // Every opened sheet retains raw bytes and starts clean. A later
            // owner mutation marks only its sheet dirty.
            sheet_dirty: vec![false; raw_sheet_xml.len()],
            sheet_included,
            sheet_materialized,
            sheet_package_backed: vec![true; sheet_count],
            raw_sheet_xml,
            raw_sheet_extensions: vec![None; sheet_count],
            slicer_defs,
            slicer_caches,
            sheet_threaded_comments,
            person_list,
            sheet_form_controls,
            sheet_controls_dirty,
            streamed_sheets: HashMap::new(),
            package_source: None,
            read_mode: options.read_mode,
            sheet_rows_limit: options.sheet_rows,
            worksheet_spool_limit: resolved_limits
                .xml_bytes
                .min(resolved_limits.entry_actual_bytes),
            date_interpretation: options.date_interpretation,
        })
    }

    fn validate_save_model(&self, limits: crate::save::SaveLimits) -> Result<()> {
        let limits = limits.validate()?;
        let mut rows = 0u64;
        let mut cells = 0u64;

        for (index, (_, worksheet)) in self.worksheets.iter().enumerate() {
            if let Some(streamed) = self.streamed_sheets.get(&index) {
                rows = rows.saturating_add(streamed.row_count);
                cells = cells.saturating_add(streamed.cell_count);
                let xml_bytes = streamed_worksheet_xml_bytes(streamed);
                reject_save_metric(
                    crate::resource::ResourceDimension::XmlBytes,
                    xml_bytes,
                    limits.xml_part_bytes,
                )?;
            } else if self
                .sheet_package_backed
                .get(index)
                .copied()
                .unwrap_or(false)
                && (!self.sheet_dirty.get(index).copied().unwrap_or(true)
                    || (!self.sheet_included.get(index).copied().unwrap_or(true)
                        && !self.sheet_materialized.get(index).copied().unwrap_or(true)))
            {
                let source = self.package_source.as_ref().ok_or_else(|| {
                    Error::Internal("package-backed sheet has no source".to_string())
                })?;
                let (_, raw_rows, raw_cells) = source
                    .worksheet_save_metrics(&self.sheet_part_path(index), limits.xml_part_bytes)?;
                rows = rows.saturating_add(raw_rows);
                cells = cells.saturating_add(raw_cells);
            } else if !self.sheet_dirty.get(index).copied().unwrap_or(true) {
                if let Some(Some(raw)) = self.raw_sheet_xml.get(index) {
                    let (raw_rows, raw_cells) = worksheet_xml_metrics(raw, limits.xml_part_bytes)?;
                    rows = rows.saturating_add(raw_rows);
                    cells = cells.saturating_add(raw_cells);
                } else if let Some(worksheet) = worksheet.get() {
                    rows = rows.saturating_add(
                        u64::try_from(worksheet.sheet_data.rows.len()).unwrap_or(u64::MAX),
                    );
                    for row in &worksheet.sheet_data.rows {
                        cells = cells
                            .saturating_add(u64::try_from(row.cells.len()).unwrap_or(u64::MAX));
                    }
                }
            } else if let Some(worksheet) = worksheet.get() {
                rows = rows.saturating_add(
                    u64::try_from(worksheet.sheet_data.rows.len()).unwrap_or(u64::MAX),
                );
                for row in &worksheet.sheet_data.rows {
                    cells =
                        cells.saturating_add(u64::try_from(row.cells.len()).unwrap_or(u64::MAX));
                }
            }

            reject_save_metric(
                crate::resource::ResourceDimension::RowCount,
                rows,
                limits.row_count,
            )?;
            reject_save_metric(
                crate::resource::ResourceDimension::CellCount,
                cells,
                limits.cell_count,
            )?;
        }
        Ok(())
    }

    fn write_zip_archive_with_limits<W: std::io::Write + std::io::Seek>(
        &self,
        writer: W,
        options: SimpleFileOptions,
        format_override: Option<WorkbookFormat>,
        limits: crate::save::SaveLimits,
    ) -> Result<W> {
        let limits = limits.validate()?;
        self.validate_save_model(limits)?;
        let mut output = SaveOutputWriter::new(writer, limits.output_bytes);
        let result = {
            let mut zip = zip::ZipWriter::new(&mut output);
            let contents = self.write_zip_contents(&mut zip, options, format_override, limits);
            let finish = zip
                .finish()
                .map(|_| ())
                .map_err(|error| Error::Zip(error.to_string()));
            contents.and(finish)
        };
        if let Some(metadata) = output.resource_error() {
            return Err(metadata.into());
        }
        result?;
        Ok(output.into_inner())
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
        self.save_with_limits(path, crate::save::SaveLimits::default())
    }

    /// Save the workbook to a file with explicit finite output limits.
    pub fn save_with_limits<P: AsRef<Path>>(
        &self,
        path: P,
        limits: crate::save::SaveLimits,
    ) -> Result<()> {
        let path = path.as_ref();
        let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");
        let target_format = WorkbookFormat::from_extension(ext)
            .ok_or_else(|| Error::UnsupportedFileExtension(ext.to_string()))?;

        atomic_write_file(path, |file| {
            let options = SimpleFileOptions::default()
                .compression_method(CompressionMethod::Deflated)
                .compression_level(Some(1));
            self.write_zip_archive_with_limits(file, options, Some(target_format), limits)?;
            Ok(())
        })
    }

    /// Serialize the workbook to an in-memory buffer using the stored format.
    pub fn save_to_buffer(&self) -> Result<Vec<u8>> {
        self.save_to_buffer_with_limits(crate::save::SaveLimits::default())
    }

    /// Serialize the workbook to a buffer with explicit finite output limits.
    pub fn save_to_buffer_with_limits(&self, limits: crate::save::SaveLimits) -> Result<Vec<u8>> {
        let cursor = std::io::Cursor::new(Vec::new());
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        let cursor = self.write_zip_archive_with_limits(cursor, options, None, limits)?;
        Ok(cursor.into_inner())
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

        let outer_size = u64::try_from(data.len()).unwrap_or(u64::MAX);
        let limits = options.resolved_limits()?;
        let mut budget = OpenBudget::new(limits)?;
        budget.check_outer_input(outer_size)?;
        budget.charge_retained(outer_size)?;
        let package = open_package_archive(PackageInput::buffer(data.into()), limits)?;
        let mut archive = PackageReader::from_archive(package.clone(), &mut budget)?;
        let mut wb = Self::from_archive(&mut archive, options)?;
        wb.package_source = Some(PackageSource::new(package, limits));
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
        Self::open_with_password_with_options(path, password, &OpenOptions::default())
    }

    /// Open an encrypted workbook from a path with explicit finite limits.
    #[cfg(feature = "encryption")]
    pub fn open_with_password_with_options<P: AsRef<Path>>(
        path: P,
        password: &str,
        options: &OpenOptions,
    ) -> Result<Self> {
        let path = path.as_ref();
        let budget = OpenBudget::new(options.resolved_limits()?)?;
        let data = read_bounded_file(path, &budget, options.crypt_limits)?;
        Self::open_encrypted_bytes(data, password, options, budget)
    }

    /// Open an encrypted workbook from an in-memory CFB container.
    #[cfg(feature = "encryption")]
    pub fn open_from_buffer_with_password(data: &[u8], password: &str) -> Result<Self> {
        Self::open_from_buffer_with_password_with_options(data, password, &OpenOptions::default())
    }

    /// Open an encrypted in-memory CFB container with explicit finite limits.
    #[cfg(feature = "encryption")]
    pub fn open_from_buffer_with_password_with_options(
        data: &[u8],
        password: &str,
        options: &OpenOptions,
    ) -> Result<Self> {
        let budget = OpenBudget::new(options.resolved_limits()?)?;
        let observed = u64::try_from(data.len()).unwrap_or(u64::MAX);
        budget.check_outer_input(observed)?;
        crate::crypt::validate_decrypt_outer_cfb(options.crypt_limits, observed)?;
        Self::open_encrypted_bytes(data.to_vec(), password, options, budget)
    }

    #[cfg(feature = "encryption")]
    fn open_encrypted_bytes(
        encrypted: Vec<u8>,
        password: &str,
        options: &OpenOptions,
        mut budget: OpenBudget,
    ) -> Result<Self> {
        let decrypted_zip =
            crate::crypt::decrypt_xlsx_with_limits(&encrypted, password, options.crypt_limits)?;
        let decrypted_size = u64::try_from(decrypted_zip.len()).unwrap_or(u64::MAX);
        budget.check_outer_input(decrypted_size)?;
        budget.charge_retained(decrypted_size)?;

        let limits = budget.limits();
        let package = open_package_archive(PackageInput::buffer(decrypted_zip.into()), limits)?;
        let mut workbook = {
            let mut archive = PackageReader::from_archive(package.clone(), &mut budget)?;
            Self::from_archive(&mut archive, options)?
        };
        workbook.package_source = Some(PackageSource::new(package, limits));
        workbook.read_mode = options.read_mode;
        Ok(workbook)
    }

    /// Save the workbook as an encrypted `.xlsx` file using Agile Encryption
    /// (AES-256-CBC + SHA-512, 100K iterations).
    #[cfg(feature = "encryption")]
    pub fn save_with_password<P: AsRef<Path>>(&self, path: P, password: &str) -> Result<()> {
        self.save_with_password_with_output_limits(
            path,
            password,
            crate::crypt::CryptLimits::default(),
            crate::save::SaveLimits::default(),
        )
    }

    /// Save an encrypted workbook with explicit finite encryption limits.
    #[cfg(feature = "encryption")]
    pub fn save_with_password_with_limits<P: AsRef<Path>>(
        &self,
        path: P,
        password: &str,
        limits: crate::crypt::CryptLimits,
    ) -> Result<()> {
        self.save_with_password_with_output_limits(
            path,
            password,
            limits,
            crate::save::SaveLimits::default(),
        )
    }

    /// Save an encrypted workbook with explicit encryption and output limits.
    #[cfg(feature = "encryption")]
    pub fn save_with_password_with_output_limits<P: AsRef<Path>>(
        &self,
        path: P,
        password: &str,
        crypt_limits: crate::crypt::CryptLimits,
        save_limits: crate::save::SaveLimits,
    ) -> Result<()> {
        let path = path.as_ref();
        let ext = path
            .extension()
            .and_then(|extension| extension.to_str())
            .unwrap_or("");
        let target_format = WorkbookFormat::from_extension(ext)
            .ok_or_else(|| Error::UnsupportedFileExtension(ext.to_string()))?;

        // First, serialize to an in-memory ZIP buffer
        let cursor = std::io::Cursor::new(Vec::new());
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        let zip_buf = self
            .write_zip_archive_with_limits(cursor, options, Some(target_format), save_limits)?
            .into_inner();

        // Encrypt and write to CFB container
        let cfb_data = crate::crypt::encrypt_xlsx_with_limits(&zip_buf, password, crypt_limits)?;
        let observed = u64::try_from(cfb_data.len()).unwrap_or(u64::MAX);
        if observed > save_limits.output_bytes {
            return Err(crate::save::resource_limit(
                crate::resource::ResourceOperation::WorkbookSave,
                crate::resource::ResourceCategory::Output,
                crate::resource::ResourceDimension::OutputBytes,
                save_limits.output_bytes,
                observed,
                crate::resource::ResourceStage::Write,
            ));
        }
        atomic_write_file(path, |file| {
            file.write_all(&cfb_data)?;
            Ok(())
        })
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
        limits: crate::save::SaveLimits,
    ) -> Result<()> {
        let effective_format = format_override.unwrap_or(self.format);
        let preserves_vba = matches!(
            effective_format,
            WorkbookFormat::Xlsm | WorkbookFormat::Xltm | WorkbookFormat::Xlam
        );
        let emits_vba = preserves_vba && self.vba_blob.is_some();
        // OPC part names compare case-insensitively, so every VBA path match
        // uses the same ASCII case folding.
        let mut vba_part_paths = HashSet::from(["xl/vbaproject.bin".to_string()]);
        vba_part_paths.extend(
            self.workbook_rels
                .relationships
                .iter()
                .filter(|relationship| relationship.rel_type == VBA_PROJECT_REL_TYPE)
                .map(|relationship| {
                    resolve_relationship_target("xl/workbook.xml", &relationship.target)
                        .to_ascii_lowercase()
                }),
        );
        let is_vba_path = |path: &str| {
            vba_part_paths.contains(&path.trim_start_matches('/').to_ascii_lowercase())
        };
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
        if emits_vba {
            let vba_part_name = "/xl/vbaProject.bin";
            content_types.overrides.retain(|entry| {
                entry.content_type != VBA_PROJECT_CONTENT_TYPE && !is_vba_path(&entry.part_name)
            });
            content_types.overrides.push(ContentTypeOverride {
                part_name: vba_part_name.to_string(),
                content_type: VBA_PROJECT_CONTENT_TYPE.to_string(),
            });
            if !content_types.defaults.iter().any(|d| d.extension == "bin") {
                content_types.defaults.push(ContentTypeDefault {
                    extension: "bin".to_string(),
                    content_type: VBA_PROJECT_CONTENT_TYPE.to_string(),
                });
            }
            workbook_rels
                .relationships
                .retain(|relationship| relationship.rel_type != VBA_PROJECT_REL_TYPE);
            let rid = crate::sheet::next_rid(&workbook_rels.relationships);
            workbook_rels.relationships.push(Relationship {
                id: rid,
                rel_type: VBA_PROJECT_REL_TYPE.to_string(),
                target: "vbaProject.bin".to_string(),
                target_mode: None,
            });
        } else if !preserves_vba || !has_deferred {
            content_types.overrides.retain(|entry| {
                entry.content_type != VBA_PROJECT_CONTENT_TYPE && !is_vba_path(&entry.part_name)
            });
            content_types
                .defaults
                .retain(|entry| entry.content_type != VBA_PROJECT_CONTENT_TYPE);
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
        // Per-sheet controls paired with their VML shape IDs and ctrlProp rIds.
        let mut worksheet_controls: HashMap<usize, Vec<(usize, String, String)>> = HashMap::new();
        let mut comment_parts_to_write: Vec<(usize, String)> = Vec::new();
        let mut control_parts_to_write: Vec<(String, String)> = Vec::new();
        let mut allocated_control_paths = self.occupied_part_paths();
        let mut allocated_comment_paths = self.occupied_part_paths();
        let mut allocated_vml_paths = self.occupied_part_paths();
        let mut removed_control_paths: HashSet<String> = HashSet::new();

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
            let sheet_path = self.sheet_part_path(sheet_idx);
            let original_comment_paths: Vec<String> = worksheet_rels
                .get(&sheet_idx)
                .into_iter()
                .flat_map(|relationships| &relationships.relationships)
                .filter(|relationship| relationship.rel_type == rel_types::COMMENTS)
                .map(|relationship| resolve_relationship_target(&sheet_path, &relationship.target))
                .collect();
            let original_vml_paths: Vec<String> = worksheet_rels
                .get(&sheet_idx)
                .into_iter()
                .flat_map(|relationships| &relationships.relationships)
                .filter(|relationship| relationship.rel_type == rel_types::VML_DRAWING)
                .map(|relationship| resolve_relationship_target(&sheet_path, &relationship.target))
                .collect();
            let controls_dirty = self
                .sheet_controls_dirty
                .get(sheet_idx)
                .copied()
                .unwrap_or(false);
            let original_control_entries = if controls_dirty {
                if let Some(entries) = self
                    .raw_sheet_xml
                    .get(sheet_idx)
                    .and_then(|raw| raw.as_deref())
                    .and_then(|raw| std::str::from_utf8(raw).ok())
                    .map(parse_worksheet_control_entries)
                {
                    entries
                } else if self
                    .sheet_package_backed
                    .get(sheet_idx)
                    .copied()
                    .unwrap_or(false)
                {
                    let source = self.package_source.as_ref().ok_or_else(|| {
                        Error::Internal("package-backed sheet has no source".to_string())
                    })?;
                    let spool = source.spool_xml_part(&sheet_path)?;
                    parse_worksheet_control_entries_reader(std::io::BufReader::new(spool))?
                } else {
                    Vec::new()
                }
            } else {
                Vec::new()
            };
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
            if has_deferred
                && !has_comments
                && !has_form_controls
                && !has_preserved_vml
                && !controls_dirty
            {
                continue;
            }

            content_types.overrides.retain(|entry| {
                let path = entry.part_name.trim_start_matches('/');
                !original_comment_paths
                    .iter()
                    .any(|original| original == path)
                    && !original_vml_paths.iter().any(|original| original == path)
            });

            let raw_control_shape_ids = self
                .sheet_vml
                .get(sheet_idx)
                .and_then(|vml| vml.as_deref())
                .map(crate::control::form_control_shape_ids)
                .unwrap_or_default();
            let mut retained_control_entries = if controls_dirty {
                original_control_entries
                    .iter()
                    .filter(|(shape_id, _, _)| raw_control_shape_ids.contains(shape_id))
                    .cloned()
                    .collect::<Vec<_>>()
            } else {
                Vec::new()
            };
            let removed_control_rids: HashSet<&str> = if controls_dirty {
                original_control_entries
                    .iter()
                    .filter(|(shape_id, _, _)| !raw_control_shape_ids.contains(shape_id))
                    .map(|(_, rid, _)| rid.as_str())
                    .collect()
            } else {
                HashSet::new()
            };

            if let Some(rels) = worksheet_rels.get_mut(&sheet_idx) {
                rels.relationships
                    .retain(|r| r.rel_type != rel_types::COMMENTS);
                rels.relationships
                    .retain(|r| r.rel_type != rel_types::VML_DRAWING);
                if !removed_control_rids.is_empty() {
                    for relationship in &rels.relationships {
                        if removed_control_rids.contains(relationship.id.as_str()) {
                            removed_control_paths.insert(resolve_relationship_target(
                                &sheet_path,
                                &relationship.target,
                            ));
                        }
                    }
                    rels.relationships.retain(|relationship| {
                        !removed_control_rids.contains(relationship.id.as_str())
                    });
                }
            }

            let needs_vml = has_comments || has_form_controls || has_preserved_vml;
            if !needs_vml && !has_comments {
                if controls_dirty {
                    worksheet_controls.insert(sheet_idx, Vec::new());
                }
                continue;
            }

            if has_comments {
                let comment_path = original_comment_paths.first().cloned().unwrap_or_else(|| {
                    next_numbered_package_path(&allocated_comment_paths, "xl/comments", ".xml")
                });
                allocated_comment_paths.insert(comment_path.clone());
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
                comment_parts_to_write.push((sheet_idx, comment_path));
            }

            if !needs_vml {
                continue;
            }

            // Regenerate only owned shapes while preserving unrelated VML.
            let vml_path = original_vml_paths.first().cloned().unwrap_or_else(|| {
                next_numbered_package_path(&allocated_vml_paths, "xl/drawings/vmlDrawing", ".vml")
            });
            allocated_vml_paths.insert(vml_path.clone());
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
                        let start_id = crate::control::next_vml_shape_id(&existing);
                        crate::vml::merge_vml_comments(&existing, &cells, start_id)
                    }
                    None => crate::vml::build_vml_drawing(&cells).into_bytes(),
                });
            }

            if controls_dirty && has_form_controls {
                let controls = &self.sheet_form_controls[sheet_idx];
                if raw_control_shape_ids.len() > controls.len() {
                    return Err(Error::InvalidArgument(
                        "cannot safely preserve existing form controls".to_string(),
                    ));
                }
                let new_controls = &controls[raw_control_shape_ids.len()..];
                let start_id = vml_bytes
                    .as_deref()
                    .map(crate::control::next_vml_shape_id)
                    .unwrap_or(1025);
                let last_shape_id = start_id
                    .checked_add(new_controls.len().saturating_sub(1))
                    .ok_or_else(|| {
                        Error::InvalidArgument("form control shape ID overflow".to_string())
                    })?;
                if (!new_controls.is_empty() && last_shape_id > 67_098_623)
                    || raw_control_shape_ids
                        .iter()
                        .any(|shape_id| *shape_id > 67_098_623)
                {
                    return Err(Error::InvalidArgument(
                        "form control shape IDs exceed Excel's supported range".to_string(),
                    ));
                }
                if !new_controls.is_empty() {
                    vml_bytes = Some(match vml_bytes {
                        Some(existing) => {
                            crate::control::merge_vml_controls(&existing, new_controls, start_id)
                        }
                        None => {
                            crate::control::build_form_control_vml(new_controls, 1025).into_bytes()
                        }
                    });
                }

                let mut shape_ids = raw_control_shape_ids.clone();
                shape_ids.extend((0..new_controls.len()).map(|offset| start_id + offset));

                let rels = worksheet_rels
                    .entry(sheet_idx)
                    .or_insert_with(default_relationships);
                let preserved_by_shape: HashMap<usize, (String, String)> = retained_control_entries
                    .drain(..)
                    .map(|(shape_id, rid, name)| (shape_id, (rid, name)))
                    .collect();
                let mut control_entries = Vec::with_capacity(controls.len());
                for (offset, (shape_id, control)) in shape_ids.into_iter().zip(controls).enumerate()
                {
                    if let Some((rid, name)) = preserved_by_shape.get(&shape_id) {
                        if rels.relationships.iter().any(|relationship| {
                            relationship.id == *rid
                                && relationship.rel_type == rel_types::CONTROL_PROPERTIES
                        }) {
                            control_entries.push((shape_id, rid.clone(), name.clone()));
                            continue;
                        }
                    }
                    let control_path = next_control_property_path(&allocated_control_paths);
                    allocated_control_paths.insert(control_path.clone());
                    let rid = crate::sheet::next_rid(&rels.relationships);
                    rels.relationships.push(Relationship {
                        id: rid.clone(),
                        rel_type: rel_types::CONTROL_PROPERTIES.to_string(),
                        target: relative_relationship_target(&sheet_path, &control_path),
                        target_mode: None,
                    });
                    content_types.overrides.push(ContentTypeOverride {
                        part_name: format!("/{control_path}"),
                        content_type: mime_types::CONTROL_PROPERTIES.to_string(),
                    });
                    control_parts_to_write.push((
                        control_path,
                        crate::control::build_control_property_xml(control),
                    ));
                    control_entries.push((
                        shape_id,
                        rid,
                        format!("{} {}", control.control_type.object_type(), offset + 1),
                    ));
                }
                worksheet_controls.insert(sheet_idx, control_entries);
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

        if !removed_control_paths.is_empty() {
            content_types.overrides.retain(|entry| {
                let path = entry.part_name.trim_start_matches('/');
                !removed_control_paths.contains(path)
            });
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
            .filter(|(path, _)| !is_vba_path(path))
            .map(|(path, _)| path.clone())
            .chain(
                self.deferred_parts
                    .remaining_parts()
                    .filter(|(path, _)| !is_vba_path(path))
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
        generated_lifecycle_paths
            .extend(control_parts_to_write.iter().map(|(path, _)| path.clone()));
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
        if emits_vba {
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
        write_xml_part(
            zip,
            "[Content_Types].xml",
            &content_types,
            options,
            limits.xml_part_bytes,
        )?;

        // _rels/.rels
        write_xml_part(
            zip,
            "_rels/.rels",
            &self.package_rels,
            options,
            limits.xml_part_bytes,
        )?;

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
            write_bytes_part(
                zip,
                "xl/workbook.xml",
                xml.as_bytes(),
                options,
                limits.xml_part_bytes,
            )?;
        } else if let (Some(raw), Some(baseline)) =
            (&self.raw_workbook_xml, &self.workbook_xml_baseline)
        {
            if self.workbook_xml == *baseline {
                write_bytes_part(zip, "xl/workbook.xml", raw, options, limits.xml_part_bytes)?;
            } else {
                write_xml_part(
                    zip,
                    "xl/workbook.xml",
                    &self.workbook_xml,
                    options,
                    limits.xml_part_bytes,
                )?;
            }
        } else {
            write_xml_part(
                zip,
                "xl/workbook.xml",
                &self.workbook_xml,
                options,
                limits.xml_part_bytes,
            )?;
        }

        // xl/_rels/workbook.xml.rels
        write_xml_part(
            zip,
            "xl/_rels/workbook.xml.rels",
            &workbook_rels,
            options,
            limits.xml_part_bytes,
        )?;

        // xl/worksheets/sheet{N}.xml
        for (i, (_name, ws_lock)) in self.worksheets.iter().enumerate() {
            let entry_name = self.sheet_part_path(i);
            let dirty = self.sheet_dirty.get(i).copied().unwrap_or(true);

            // If the sheet has streamed data, write it directly from the temp file.
            if let Some(streamed) = self.streamed_sheets.get(&i) {
                crate::stream::write_streamed_sheet(
                    zip,
                    &entry_name,
                    streamed,
                    options,
                    limits.xml_part_bytes,
                )?;
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
                || !slicer_rids.is_empty()
                || self
                    .sheet_sparklines
                    .get(i)
                    .is_some_and(|sparklines| !sparklines.is_empty());
            let included = self.sheet_included.get(i).copied().unwrap_or(true);
            let package_backed = self.sheet_package_backed.get(i).copied().unwrap_or(false);
            if !dirty && (!included || !needs_aux_injection) {
                if let Some(Some(raw_bytes)) = self.raw_sheet_xml.get(i) {
                    write_bytes_part(zip, &entry_name, raw_bytes, options, limits.xml_part_bytes)?;
                    continue;
                }
                if package_backed {
                    let source = self.package_source.as_ref().ok_or_else(|| {
                        Error::Internal("package-backed sheet has no source".to_string())
                    })?;
                    source.raw_copy_part(&entry_name, zip)?;
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
                } else if package_backed && self.package_source.is_some() && included {
                    hydrated_for_save = self.deserialize_package_sheet_for_save(i)?;
                    &hydrated_for_save
                } else {
                    match ws_lock.get() {
                        Some(ws) => ws,
                        None => continue,
                    }
                }
            } else if !included
                && !self.sheet_materialized.get(i).copied().unwrap_or(true)
                && package_backed
                && self.package_source.is_some()
            {
                // Auxiliary-only mutations can dirty an excluded sheet while
                // its typed worksheet is still the empty filter placeholder.
                // Merge those auxiliary changes into the complete pinned
                // package worksheet rather than serializing the placeholder.
                hydrated_for_save = self.deserialize_package_sheet_for_save(i)?;
                &hydrated_for_save
            } else {
                match ws_lock.get() {
                    Some(ws) => ws,
                    None => {
                        if let Some(Some(raw_bytes)) = self.raw_sheet_xml.get(i) {
                            hydrated_for_save = deserialize_worksheet_xml(raw_bytes)?;
                            &hydrated_for_save
                        } else if package_backed && self.package_source.is_some() && included {
                            hydrated_for_save = self.deserialize_package_sheet_for_save(i)?;
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
                || !slicer_rids.is_empty()
                || self.sheet_controls_dirty.get(i).copied().unwrap_or(false);

            if !has_extras {
                write_xml_part(zip, &entry_name, ws, options, limits.xml_part_bytes)?;
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
                let retained_extension = self
                    .raw_sheet_xml
                    .get(i)
                    .and_then(|raw| raw.as_deref())
                    .or_else(|| {
                        self.raw_sheet_extensions
                            .get(i)
                            .and_then(|raw| raw.as_deref())
                    });
                let package_extension = if retained_extension.is_none() && package_backed {
                    self.package_source
                        .as_ref()
                        .map(|source| source.worksheet_extension_fragment(&entry_name))
                        .transpose()?
                        .flatten()
                } else {
                    None
                };
                let raw_sheet = retained_extension.or(package_extension.as_deref());
                let xml = serialize_worksheet_with_slicer_extras(
                    ws_ref,
                    sparklines,
                    legacy_rid,
                    raw_sheet,
                    &slicer_rids,
                    worksheet_controls.get(&i).map(Vec::as_slice).unwrap_or(&[]),
                    self.sheet_controls_dirty.get(i).copied().unwrap_or(false),
                )?;
                write_bytes_part(
                    zip,
                    &entry_name,
                    xml.as_bytes(),
                    options,
                    limits.xml_part_bytes,
                )?;
            }
        }

        // xl/styles.xml
        write_xml_part(
            zip,
            "xl/styles.xml",
            &self.stylesheet,
            options,
            limits.xml_part_bytes,
        )?;

        // xl/sharedStrings.xml -- write from the runtime SST
        let sst_xml = self.sst_runtime.to_sst();
        write_xml_part(
            zip,
            "xl/sharedStrings.xml",
            &sst_xml,
            options,
            limits.xml_part_bytes,
        )?;

        // Write hydrated comments at their retained or newly allocated paths.
        for (sheet_idx, path) in &comment_parts_to_write {
            if let Some(comments) = self.sheet_comments[*sheet_idx].as_ref() {
                write_xml_part(zip, path, comments, options, limits.xml_part_bytes)?;
            }
        }

        // xl/drawings/vmlDrawing{N}.vml -- write VML drawing parts
        for (_sheet_idx, vml_path, vml_bytes) in &vml_parts_to_write {
            write_bytes_part(zip, vml_path, vml_bytes, options, limits.xml_part_bytes)?;
        }

        for (path, xml) in &control_parts_to_write {
            write_bytes_part(zip, path, xml.as_bytes(), options, limits.xml_part_bytes)?;
        }

        let mut emitted_graph_paths: HashSet<String> = HashSet::new();

        // xl/drawings/drawing{N}.xml -- write drawing parts
        for (path, drawing) in &self.drawings {
            if let Some(bytes) = self.clean_raw_graph_part(path) {
                write_bytes_part(zip, path, bytes, options, limits.xml_part_bytes)?;
            } else {
                write_xml_part(zip, path, drawing, options, limits.xml_part_bytes)?;
            }
            emitted_graph_paths.insert(path.clone());
        }

        // xl/charts/chart{N}.xml -- write chart parts
        for (path, chart) in &self.charts {
            if let Some(bytes) = self.clean_raw_graph_part(path) {
                write_bytes_part(zip, path, bytes, options, limits.xml_part_bytes)?;
            } else {
                write_xml_part(zip, path, chart, options, limits.xml_part_bytes)?;
            }
            emitted_graph_paths.insert(path.clone());
        }
        for (path, data) in &self.raw_charts {
            if self.charts.iter().any(|(p, _)| p == path) {
                continue;
            }
            write_bytes_part(
                zip,
                path,
                self.clean_raw_graph_part(path).unwrap_or(data),
                options,
                limits.xml_part_bytes,
            )?;
            emitted_graph_paths.insert(path.clone());
        }

        // xl/media/image{N}.{ext} -- write image data
        for (path, data) in &self.images {
            write_package_bytes_part(zip, path, data, options, limits.xml_part_bytes)?;
        }

        // xl/worksheets/_rels/sheet{N}.xml.rels -- write worksheet relationships
        for (sheet_idx, rels) in &worksheet_rels {
            let sheet_path = self.sheet_part_path(*sheet_idx);
            let path = relationship_part_path(&sheet_path);
            write_xml_part(zip, &path, rels, options, limits.xml_part_bytes)?;
        }

        // xl/drawings/_rels/drawing{N}.xml.rels -- write drawing relationships
        for (drawing_idx, rels) in &self.drawing_rels {
            if let Some((drawing_path, _)) = self.drawings.get(*drawing_idx) {
                let path = relationship_part_path(drawing_path);
                if let Some(bytes) = self.clean_raw_graph_part(&path) {
                    write_bytes_part(zip, &path, bytes, options, limits.xml_part_bytes)?;
                } else {
                    write_xml_part(zip, &path, rels, options, limits.xml_part_bytes)?;
                }
                emitted_graph_paths.insert(path);
            }
        }

        // Raw drawing or relationship parts that could not be parsed eagerly.
        for (path, data) in &self.raw_graph_parts {
            if self.dirty_graph_parts.contains(path) || emitted_graph_paths.contains(path) {
                continue;
            }
            write_bytes_part(zip, path, data, options, limits.xml_part_bytes)?;
        }

        // xl/pivotTables/pivotTable{N}.xml
        for (path, pt) in &self.pivot_tables {
            write_xml_part(zip, path, pt, options, limits.xml_part_bytes)?;
        }

        // xl/pivotCache/pivotCacheDefinition{N}.xml
        for (path, pcd) in &self.pivot_cache_defs {
            write_xml_part(zip, path, pcd, options, limits.xml_part_bytes)?;
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
            write_bytes_part(zip, path, xml.as_bytes(), options, limits.xml_part_bytes)?;
        }

        // xl/tables/table{N}.xml
        for (path, table_xml, _sheet_idx) in &self.tables {
            if let Some(raw) = table_xml.unchanged_source() {
                write_bytes_part(zip, path, raw.as_bytes(), options, limits.xml_part_bytes)?;
            } else {
                write_xml_part(zip, path, table_xml, options, limits.xml_part_bytes)?;
            }
        }

        // xl/slicers/slicer{N}.xml
        for (path, sd) in &self.slicer_defs {
            write_xml_part(zip, path, sd, options, limits.xml_part_bytes)?;
        }

        // xl/slicerCaches/slicerCache{N}.xml (manual serialization)
        for (path, scd) in &self.slicer_caches {
            let xml_str = format!(
                "{}\n{}",
                XML_DECLARATION,
                sheetkit_xml::slicer::serialize_slicer_cache(scd),
            );
            write_bytes_part(
                zip,
                path,
                xml_str.as_bytes(),
                options,
                limits.xml_part_bytes,
            )?;
        }

        // xl/theme/theme1.xml
        {
            let default_theme = crate::theme::default_theme_xml();
            let theme_bytes = self.theme_xml.as_deref().unwrap_or(&default_theme);
            write_bytes_part(
                zip,
                "xl/theme/theme1.xml",
                theme_bytes,
                options,
                limits.xml_part_bytes,
            )?;
        }

        // xl/vbaProject.bin -- write VBA blob if present
        if let Some(blob) = self.vba_blob.as_ref().filter(|_| emits_vba) {
            write_package_bytes_part(
                zip,
                "xl/vbaProject.bin",
                blob,
                options,
                limits.xml_part_bytes,
            )?;
        }

        // docProps/core.xml
        if let Some(raw) = self.clean_raw_doc_prop("docProps/core.xml") {
            write_bytes_part(
                zip,
                "docProps/core.xml",
                raw,
                options,
                limits.xml_part_bytes,
            )?;
        } else if let Some(ref props) = self.core_properties {
            let xml_str = sheetkit_xml::doc_props::serialize_core_properties(props);
            write_bytes_part(
                zip,
                "docProps/core.xml",
                xml_str.as_bytes(),
                options,
                limits.xml_part_bytes,
            )?;
        }

        // docProps/app.xml
        if let Some(raw) = self.clean_raw_doc_prop("docProps/app.xml") {
            write_bytes_part(zip, "docProps/app.xml", raw, options, limits.xml_part_bytes)?;
        } else if let Some(ref props) = self.app_properties {
            write_xml_part(
                zip,
                "docProps/app.xml",
                props,
                options,
                limits.xml_part_bytes,
            )?;
        }

        // docProps/custom.xml
        if let Some(raw) = self.clean_raw_doc_prop("docProps/custom.xml") {
            write_bytes_part(
                zip,
                "docProps/custom.xml",
                raw,
                options,
                limits.xml_part_bytes,
            )?;
        } else if let Some(ref props) = self.custom_properties {
            let xml_str = sheetkit_xml::doc_props::serialize_custom_properties(props);
            write_bytes_part(
                zip,
                "docProps/custom.xml",
                xml_str.as_bytes(),
                options,
                limits.xml_part_bytes,
            )?;
        }

        // xl/threadedComments/threadedComment{N}.xml
        if has_any_threaded {
            for (i, tc) in self.sheet_threaded_comments.iter().enumerate() {
                if let Some(ref tc_data) = tc {
                    let tc_path = format!("xl/threadedComments/threadedComment{}.xml", i + 1);
                    write_xml_part(zip, &tc_path, tc_data, options, limits.xml_part_bytes)?;
                }
            }
            write_xml_part(
                zip,
                "xl/persons/person.xml",
                &self.person_list,
                options,
                limits.xml_part_bytes,
            )?;
        }

        // Write back unknown parts preserved from the original file.
        for (path, data) in &self.unknown_parts {
            if removed_control_paths.contains(path) || is_vba_path(path) {
                continue;
            }
            write_package_bytes_part(zip, path, data, options, limits.xml_part_bytes)?;
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
            for (path, _) in &control_parts_to_write {
                emitted_owned.insert(path.clone());
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
            if emits_vba || !preserves_vba {
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
                if emitted_owned.contains(path) || is_vba_path(path) {
                    continue;
                }
                write_package_bytes_part(zip, path, data, options, limits.xml_part_bytes)?;
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

#[cfg(feature = "encryption")]
fn read_bounded_file(
    path: &Path,
    budget: &OpenBudget,
    crypt_limits: crate::crypt::CryptLimits,
) -> Result<Vec<u8>> {
    let file = std::fs::File::open(path)?;
    let declared = file.metadata()?.len();
    budget.check_outer_input(declared)?;
    let crypt_limit = crate::crypt::validate_decrypt_outer_cfb(crypt_limits, declared)?;

    let limit = budget.limits().outer_input_bytes.min(crypt_limit);
    let mut reader = file.take(limit.saturating_add(1));
    let initial_capacity =
        usize::try_from(declared.min(LARGE_BUF_CAPACITY as u64)).unwrap_or(LARGE_BUF_CAPACITY);
    let mut data = Vec::new();
    data.try_reserve(initial_capacity).map_err(|_| {
        Error::Io(std::io::Error::new(
            std::io::ErrorKind::OutOfMemory,
            "unable to reserve bounded encrypted workbook buffer",
        ))
    })?;
    reader.read_to_end(&mut data)?;
    let observed = u64::try_from(data.len()).unwrap_or(u64::MAX);
    budget.check_outer_input(observed)?;
    crate::crypt::validate_decrypt_outer_cfb(crypt_limits, observed)?;
    Ok(data)
}

fn atomic_write_file(
    path: &Path,
    write: impl FnOnce(&mut std::fs::File) -> Result<()>,
) -> Result<()> {
    atomic_write_file_with_parent_sync(path, write, sync_parent_directory)
}

fn atomic_write_file_with_parent_sync(
    path: &Path,
    write: impl FnOnce(&mut std::fs::File) -> Result<()>,
    sync_parent: impl FnOnce(&Path) -> std::io::Result<()>,
) -> Result<()> {
    // Replace the file a symlink points to rather than the link itself.
    let (destination, existing_permissions) = resolve_save_destination(path)?;
    let parent = destination
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut builder = tempfile::Builder::new();
    builder.prefix(".sheetkit-");
    #[cfg(unix)]
    if existing_permissions.is_none() {
        // New files get the conventional mode, still narrowed by the umask.
        use std::os::unix::fs::PermissionsExt as _;
        builder.permissions(std::fs::Permissions::from_mode(0o666));
    }
    let mut temporary = builder.tempfile_in(parent)?;
    write(temporary.as_file_mut())?;
    temporary.as_file_mut().flush()?;
    #[cfg(unix)]
    if let Some(permissions) = existing_permissions {
        temporary.as_file().set_permissions(permissions)?;
    }
    #[cfg(not(unix))]
    let _ = existing_permissions;
    temporary.as_file().sync_all()?;
    temporary
        .persist(&destination)
        .map_err(|error| Error::Io(error.error))?;

    // The rename is the commit point. A failed directory sync cannot undo it.
    let _ = sync_parent(parent);
    Ok(())
}

/// Follow symlinks from `path` to the file a save must replace. Returns the
/// final path and the permissions of an existing destination file.
fn resolve_save_destination(
    path: &Path,
) -> std::io::Result<(std::path::PathBuf, Option<std::fs::Permissions>)> {
    const MAX_SYMLINK_HOPS: usize = 40;
    let mut current = path.to_path_buf();
    for _ in 0..MAX_SYMLINK_HOPS {
        match std::fs::symlink_metadata(&current) {
            Ok(metadata) if metadata.file_type().is_symlink() => {
                let target = std::fs::read_link(&current)?;
                current = if target.is_absolute() {
                    target
                } else {
                    current
                        .parent()
                        .map(|parent| parent.join(&target))
                        .unwrap_or(target)
                };
            }
            Ok(metadata) => return Ok((current, Some(metadata.permissions()))),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
                return Ok((current, None));
            }
            Err(error) => return Err(error),
        }
    }
    Err(std::io::Error::new(
        std::io::ErrorKind::InvalidInput,
        "too many levels of symbolic links in save destination",
    ))
}

fn sync_parent_directory(parent: &Path) -> std::io::Result<()> {
    #[cfg(unix)]
    {
        std::fs::File::open(parent)?.sync_all()
    }

    #[cfg(not(unix))]
    {
        let _ = parent;
        Ok(())
    }
}

fn reject_save_metric(
    dimension: crate::resource::ResourceDimension,
    observed: u64,
    limit: u64,
) -> Result<()> {
    if observed > limit {
        return Err(crate::save::resource_limit(
            crate::resource::ResourceOperation::WorkbookSave,
            crate::resource::ResourceCategory::Output,
            dimension,
            limit,
            observed,
            crate::resource::ResourceStage::Serialize,
        ));
    }
    Ok(())
}

fn worksheet_xml_metrics(bytes: &[u8], limit: u64) -> Result<(u64, u64)> {
    reject_save_metric(
        crate::resource::ResourceDimension::XmlBytes,
        u64::try_from(bytes.len()).unwrap_or(u64::MAX),
        limit,
    )?;
    let mut reader = quick_xml::Reader::from_reader(std::io::Cursor::new(bytes));
    let mut buffer = Vec::with_capacity(4096);
    let mut rows = 0u64;
    let mut cells = 0u64;
    loop {
        buffer.clear();
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::XmlParse(error.to_string()))?
        {
            quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element) => {
                match element.local_name().as_ref() {
                    b"row" => rows = rows.saturating_add(1),
                    b"c" => cells = cells.saturating_add(1),
                    _ => {}
                }
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    Ok((rows, cells))
}

fn streamed_worksheet_xml_bytes(streamed: &crate::stream::StreamedSheetData) -> u64 {
    u64::try_from(crate::stream::build_worksheet_header(streamed).len())
        .unwrap_or(u64::MAX)
        .saturating_add(streamed.data_len)
        .saturating_add(
            u64::try_from(crate::stream::build_worksheet_footer(streamed).len())
                .unwrap_or(u64::MAX),
        )
}

struct SaveOutputWriter<W> {
    inner: W,
    limit: u64,
    position: u64,
    length: u64,
    resource_error: Option<crate::resource::ResourceLimitError>,
}

impl<W> SaveOutputWriter<W> {
    fn new(inner: W, limit: u64) -> Self {
        Self {
            inner,
            limit,
            position: 0,
            length: 0,
            resource_error: None,
        }
    }

    fn resource_error(&self) -> Option<crate::resource::ResourceLimitError> {
        self.resource_error
    }

    fn into_inner(self) -> W {
        self.inner
    }

    fn mark_output_exceeded(&mut self) {
        if self.resource_error.is_none() {
            self.resource_error = crate::resource::ResourceLimitError::exceeded(
                crate::resource::ResourceOperation::WorkbookSave,
                crate::resource::ResourceCategory::Output,
                crate::resource::ResourceDimension::OutputBytes,
                self.limit,
                self.limit.saturating_add(1),
                Some(crate::resource::ResourceStage::Write),
            );
        }
    }

    fn virtual_seek(&self, position: std::io::SeekFrom) -> std::io::Result<u64> {
        let next = match position {
            std::io::SeekFrom::Start(offset) => i128::from(offset),
            std::io::SeekFrom::Current(offset) => {
                i128::from(self.position).saturating_add(i128::from(offset))
            }
            std::io::SeekFrom::End(offset) => {
                i128::from(self.length).saturating_add(i128::from(offset))
            }
        };
        u64::try_from(next).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "invalid output seek")
        })
    }
}

impl<W: std::io::Write> std::io::Write for SaveOutputWriter<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let amount = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
        let observed = self.position.saturating_add(amount);
        if self.resource_error.is_some() || observed > self.limit {
            self.mark_output_exceeded();
            self.position = observed;
            self.length = self.length.max(observed);
            return Ok(bytes.len());
        }
        let written = self.inner.write(bytes)?;
        self.position = self
            .position
            .saturating_add(u64::try_from(written).unwrap_or(u64::MAX));
        self.length = self.length.max(self.position);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        if self.resource_error.is_some() {
            Ok(())
        } else {
            self.inner.flush()
        }
    }
}

impl<W: std::io::Seek> std::io::Seek for SaveOutputWriter<W> {
    fn seek(&mut self, position: std::io::SeekFrom) -> std::io::Result<u64> {
        let virtual_next = self.virtual_seek(position)?;
        if self.resource_error.is_some() || virtual_next > self.limit {
            self.mark_output_exceeded();
            self.position = virtual_next;
            self.length = self.length.max(virtual_next);
            return Ok(virtual_next);
        }
        let next = self.inner.seek(position)?;
        self.position = next;
        self.length = self.length.max(next);
        Ok(next)
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
    deserialize_worksheet_xml_with_row_limit(reader, None)
}

/// Incrementally filter worksheet rows before typed deserialization.
///
/// Rows beyond `row_limit` are scanned but never copied into the bounded
/// intermediate spool and therefore cannot create typed row/cell allocations.
pub(super) fn deserialize_worksheet_xml_with_row_limit<R: std::io::BufRead>(
    source: R,
    row_limit: Option<u32>,
) -> Result<WorksheetXml> {
    let mut ws: WorksheetXml = if let Some(row_limit) = row_limit {
        use std::io::{Seek as _, Write as _};

        let mut reader = quick_xml::Reader::from_reader(source);
        reader.config_mut().trim_text(false);
        let spool = tempfile::tempfile()?;
        let mut writer = quick_xml::Writer::new(std::io::BufWriter::new(spool));
        let mut buffer = Vec::with_capacity(4096);
        let mut in_sheet_data = false;
        let mut rows_kept = 0u32;
        let mut skipped_depth: Option<usize> = None;

        loop {
            buffer.clear();
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(|error| Error::XmlParse(error.to_string()))?;

            if let Some(depth) = &mut skipped_depth {
                match &event {
                    quick_xml::events::Event::Start(_) => *depth += 1,
                    quick_xml::events::Event::End(_) => {
                        *depth -= 1;
                        if *depth == 0 {
                            skipped_depth = None;
                        }
                    }
                    quick_xml::events::Event::Eof => {
                        return Err(Error::XmlParse(
                            "unexpected EOF while skipping worksheet row".to_string(),
                        ));
                    }
                    _ => {}
                }
                continue;
            }

            match &event {
                quick_xml::events::Event::Start(element)
                    if element.local_name().as_ref() == b"sheetData" =>
                {
                    in_sheet_data = true;
                }
                quick_xml::events::Event::End(element)
                    if element.local_name().as_ref() == b"sheetData" =>
                {
                    in_sheet_data = false;
                }
                quick_xml::events::Event::Start(element)
                    if in_sheet_data && element.local_name().as_ref() == b"row" =>
                {
                    if rows_kept >= row_limit {
                        skipped_depth = Some(1);
                        continue;
                    }
                    rows_kept += 1;
                }
                quick_xml::events::Event::Empty(element)
                    if in_sheet_data && element.local_name().as_ref() == b"row" =>
                {
                    if rows_kept >= row_limit {
                        continue;
                    }
                    rows_kept += 1;
                }
                quick_xml::events::Event::Eof => break,
                _ => {}
            }
            writer
                .write_event(event.into_owned())
                .map_err(|error| Error::XmlParse(error.to_string()))?;
        }

        let mut spool = writer
            .into_inner()
            .into_inner()
            .map_err(|error| Error::Io(error.into_error()))?;
        spool.flush()?;
        spool.seek(std::io::SeekFrom::Start(0))?;
        quick_xml::de::from_reader(std::io::BufReader::new(spool))
            .map_err(|error| Error::XmlDeserialize(error.to_string()))?
    } else {
        quick_xml::de::from_reader(source)
            .map_err(|error| Error::XmlDeserialize(error.to_string()))?
    };
    crate::utils::cell_ref::validate_and_index_worksheet(&mut ws)?;
    for row in &mut ws.sheet_data.rows {
        row.cells.shrink_to_fit();
    }
    ws.sheet_data.rows.shrink_to_fit();
    Ok(ws)
}

fn resolve_checked_table_target(source_part: &str, target: &str) -> Result<String> {
    if target.is_empty()
        || target.contains('\\')
        || target
            .chars()
            .any(|character| matches!(character, '?' | '#' | '\0'))
        || target.split('/').any(|segment| segment == ".")
    {
        return Err(Error::InvalidArgument(
            "invalid table relationship target".to_string(),
        ));
    }

    let mut segments: Vec<&str> = if target.starts_with('/') {
        Vec::new()
    } else {
        source_part
            .rsplit_once('/')
            .map(|(directory, _)| directory.split('/').collect())
            .unwrap_or_default()
    };

    for segment in target.trim_start_matches('/').split('/') {
        match segment {
            "" => {
                return Err(Error::InvalidArgument(
                    "invalid table relationship target".to_string(),
                ));
            }
            ".." => {
                if segments.pop().is_none() {
                    return Err(Error::InvalidArgument(
                        "table relationship target escapes package root".to_string(),
                    ));
                }
            }
            value if value.contains(':') => {
                return Err(Error::InvalidArgument(
                    "external table relationship target is not allowed".to_string(),
                ));
            }
            value => segments.push(value),
        }
    }

    let resolved = segments.join("/");
    if !resolved.starts_with("xl/tables/") || !resolved.ends_with(".xml") {
        return Err(Error::InvalidArgument(format!(
            "table relationship target is outside xl/tables: {resolved}"
        )));
    }
    Ok(resolved)
}

fn validate_table_relationship_ownership(
    worksheet_paths: &[String],
    worksheet_rels: &HashMap<usize, Relationships>,
    content_types: &ContentTypes,
    package_paths: &HashSet<String>,
) -> Result<HashMap<String, usize>> {
    let mut declared_tables = HashSet::new();
    for entry in &content_types.overrides {
        if entry.content_type != mime_types::TABLE {
            continue;
        }
        if !entry.part_name.starts_with('/') {
            return Err(Error::InvalidArgument(
                "table content type path must be package-absolute".to_string(),
            ));
        }
        let path = resolve_checked_table_target("", &entry.part_name)?;
        if !declared_tables.insert(path.clone()) {
            return Err(Error::InvalidArgument(format!(
                "duplicate table content type declaration: {path}"
            )));
        }
        if !package_paths.contains(&path) {
            return Err(Error::InvalidArgument(format!(
                "table relationship target is missing from package: {path}"
            )));
        }
    }

    let mut owners = HashMap::new();
    for (sheet_index, relationships) in worksheet_rels {
        let sheet_path = worksheet_paths.get(*sheet_index).ok_or_else(|| {
            Error::InvalidArgument(format!(
                "table relationship references unknown sheet index {sheet_index}"
            ))
        })?;
        let mut id_counts: HashMap<&str, usize> = HashMap::new();
        for relationship in &relationships.relationships {
            *id_counts.entry(&relationship.id).or_default() += 1;
        }

        for relationship in &relationships.relationships {
            if relationship.rel_type != rel_types::TABLE {
                continue;
            }
            if relationship.target_mode.is_some() {
                return Err(Error::InvalidArgument(format!(
                    "external table relationship is not allowed: {}",
                    relationship.id
                )));
            }
            if relationship.id.is_empty() || id_counts[relationship.id.as_str()] != 1 {
                return Err(Error::InvalidArgument(format!(
                    "duplicate or empty table relationship id: {}",
                    relationship.id
                )));
            }

            let path = resolve_checked_table_target(sheet_path, &relationship.target)?;
            if !declared_tables.contains(&path) || !package_paths.contains(&path) {
                return Err(Error::InvalidArgument(format!(
                    "table relationship target is missing or undeclared: {path}"
                )));
            }
            if let Some(previous_owner) = owners.insert(path.clone(), *sheet_index) {
                return Err(Error::InvalidArgument(format!(
                    "table part has multiple owners ({previous_owner}, {sheet_index}): {path}"
                )));
            }
        }
    }

    for path in declared_tables {
        if !owners.contains_key(&path) {
            return Err(Error::InvalidArgument(format!(
                "table part has no worksheet relationship owner: {path}"
            )));
        }
    }
    Ok(owners)
}

fn resolved_namespace_is(result: &quick_xml::name::ResolveResult<'_>, expected: &str) -> bool {
    matches!(
        result,
        quick_xml::name::ResolveResult::Bound(namespace)
            if namespace.as_ref() == expected.as_bytes()
    )
}

fn table_part_relationship_id<R: std::io::BufRead>(
    reader: &quick_xml::NsReader<R>,
    element: &quick_xml::events::BytesStart<'_>,
    sheet_path: &str,
) -> Result<String> {
    let mut relationship_id = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::XmlParse(error.to_string()))?;
        let (namespace, local_name) = reader.resolve_attribute(attribute.key);
        if local_name.as_ref() == b"id"
            && resolved_namespace_is(&namespace, sheetkit_xml::namespaces::RELATIONSHIPS)
        {
            if relationship_id.is_some() {
                return Err(Error::InvalidArgument(format!(
                    "worksheet tablePart has duplicate r:id attributes: {sheet_path}"
                )));
            }
            relationship_id = Some(
                attribute
                    .decode_and_unescape_value(reader.decoder())
                    .map_err(|error| Error::XmlParse(error.to_string()))?
                    .into_owned(),
            );
        }
    }
    relationship_id.ok_or_else(|| {
        Error::InvalidArgument(format!("worksheet tablePart is missing r:id: {sheet_path}"))
    })
}

fn validate_worksheet_table_parts<A: ArchiveAccess>(
    archive: &mut A,
    worksheet_paths: &[String],
    worksheet_rels: &HashMap<usize, Relationships>,
    sheet_checked: &[bool],
) -> Result<()> {
    for (sheet_index, relationships) in worksheet_rels {
        if !sheet_checked.get(*sheet_index).copied().unwrap_or(false) {
            continue;
        }
        let expected_ids = table_relationship_ids(Some(relationships));
        if expected_ids.is_empty() {
            continue;
        }

        let sheet_path = worksheet_paths.get(*sheet_index).ok_or_else(|| {
            Error::InvalidArgument(format!(
                "table relationship references unknown sheet index {sheet_index}"
            ))
        })?;
        let (result, read_error) = {
            let entry = archive.entry(sheet_path, true)?;
            let mut source = PartReader::new(entry, None);
            let result = validate_worksheet_table_part_stream(
                std::io::BufReader::with_capacity(LARGE_BUF_CAPACITY, &mut source),
                &expected_ids,
                sheet_path,
            );
            (result, source.read_error)
        };
        if let Some(error) = archive.pending_error() {
            return Err(error);
        }
        if let Some(message) = read_error {
            return Err(Error::Zip(message));
        }
        result?;
    }
    Ok(())
}

/// Table relationship ids declared by one worksheet relationship part.
fn table_relationship_ids(relationships: Option<&Relationships>) -> HashSet<String> {
    relationships
        .into_iter()
        .flat_map(|relationships| &relationships.relationships)
        .filter(|relationship| relationship.rel_type == rel_types::TABLE)
        .map(|relationship| relationship.id.clone())
        .collect()
}

/// Check that a worksheet's `tableParts` list matches its table relationships.
fn validate_worksheet_table_part_stream<R: std::io::BufRead>(
    source: R,
    expected_ids: &HashSet<String>,
    sheet_path: &str,
) -> Result<()> {
    let mut reader = quick_xml::NsReader::from_reader(source);
    let mut buffer = Vec::with_capacity(4096);
    let mut declared_count = None;
    let mut seen_container = false;
    let mut in_container = false;
    let mut open_table_part = false;
    let mut actual_ids = HashSet::with_capacity(expected_ids.len());

    loop {
        buffer.clear();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::XmlParse(error.to_string()))?;
        let is_spreadsheet_ml =
            resolved_namespace_is(&namespace, sheetkit_xml::namespaces::SPREADSHEET_ML);
        if open_table_part {
            match event {
                quick_xml::events::Event::End(element)
                    if is_spreadsheet_ml && element.local_name().as_ref() == b"tablePart" =>
                {
                    open_table_part = false;
                    continue;
                }
                quick_xml::events::Event::Text(text)
                    if text.as_ref().iter().all(u8::is_ascii_whitespace) =>
                {
                    continue;
                }
                quick_xml::events::Event::Comment(_) => continue,
                _ => {
                    return Err(Error::InvalidArgument(format!(
                        "worksheet tablePart must be empty: {sheet_path}"
                    )));
                }
            }
        }
        match event {
            quick_xml::events::Event::Start(element)
                if is_spreadsheet_ml && element.local_name().as_ref() == b"tableParts" =>
            {
                if seen_container {
                    return Err(Error::InvalidArgument(format!(
                        "worksheet has duplicate tableParts containers: {sheet_path}"
                    )));
                }
                seen_container = true;
                in_container = true;
                for attribute in element.attributes() {
                    let attribute =
                        attribute.map_err(|error| Error::XmlParse(error.to_string()))?;
                    let (attribute_namespace, local_name) = reader.resolve_attribute(attribute.key);
                    if matches!(attribute_namespace, quick_xml::name::ResolveResult::Unbound)
                        && local_name.as_ref() == b"count"
                    {
                        if declared_count.is_some() {
                            return Err(Error::InvalidArgument(format!(
                                "worksheet tableParts has duplicate count attributes: {sheet_path}"
                            )));
                        }
                        declared_count = Some(
                            attribute
                                .decode_and_unescape_value(reader.decoder())
                                .map_err(|error| Error::XmlParse(error.to_string()))?
                                .parse::<u32>()
                                .map_err(|error| Error::XmlParse(error.to_string()))?,
                        );
                    }
                }
            }
            quick_xml::events::Event::End(element)
                if is_spreadsheet_ml && element.local_name().as_ref() == b"tableParts" =>
            {
                in_container = false;
            }
            quick_xml::events::Event::Empty(element)
                if in_container
                    && is_spreadsheet_ml
                    && element.local_name().as_ref() == b"tablePart" =>
            {
                let relationship_id = table_part_relationship_id(&reader, &element, sheet_path)?;
                if !actual_ids.insert(relationship_id) || actual_ids.len() > expected_ids.len() {
                    return Err(Error::InvalidArgument(format!(
                        "worksheet tableParts do not match table relationships: {sheet_path}"
                    )));
                }
            }
            quick_xml::events::Event::Start(element)
                if in_container
                    && is_spreadsheet_ml
                    && element.local_name().as_ref() == b"tablePart" =>
            {
                let relationship_id = table_part_relationship_id(&reader, &element, sheet_path)?;
                if !actual_ids.insert(relationship_id) || actual_ids.len() > expected_ids.len() {
                    return Err(Error::InvalidArgument(format!(
                        "worksheet tableParts do not match table relationships: {sheet_path}"
                    )));
                }
                open_table_part = true;
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }

    if open_table_part || declared_count.is_some_and(|count| count as usize != actual_ids.len()) {
        return Err(Error::InvalidArgument(format!(
            "worksheet tableParts count does not match entries: {sheet_path}"
        )));
    }
    if actual_ids.len() != expected_ids.len()
        || !actual_ids
            .iter()
            .all(|relationship_id| expected_ids.contains(relationship_id.as_str()))
    {
        return Err(Error::InvalidArgument(format!(
            "worksheet tableParts do not match table relationships: {sheet_path}"
        )));
    }
    Ok(())
}

pub(crate) fn validate_materialized_worksheet_table_parts(
    worksheet: &WorksheetXml,
    relationships: Option<&Relationships>,
    sheet_path: &str,
) -> Result<()> {
    let expected_ids: HashSet<&str> = relationships
        .into_iter()
        .flat_map(|relationships| &relationships.relationships)
        .filter(|relationship| relationship.rel_type == rel_types::TABLE)
        .map(|relationship| relationship.id.as_str())
        .collect();
    if expected_ids.is_empty() {
        return Ok(());
    }
    let Some(table_parts) = &worksheet.table_parts else {
        return Err(Error::InvalidArgument(format!(
            "worksheet tableParts do not match table relationships: {sheet_path}"
        )));
    };
    if table_parts
        .count
        .is_some_and(|count| usize::try_from(count).ok() != Some(table_parts.table_parts.len()))
    {
        return Err(Error::InvalidArgument(format!(
            "worksheet tableParts count does not match entries: {sheet_path}"
        )));
    }
    let actual_ids: HashSet<&str> = table_parts
        .table_parts
        .iter()
        .map(|table_part| table_part.r_id.as_str())
        .collect();
    if actual_ids.len() != table_parts.table_parts.len()
        || actual_ids.len() != expected_ids.len()
        || !actual_ids.iter().all(|id| expected_ids.contains(id))
    {
        return Err(Error::InvalidArgument(format!(
            "worksheet tableParts do not match table relationships: {sheet_path}"
        )));
    }
    Ok(())
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
pub(crate) fn read_xml_part<T: serde::de::DeserializeOwned, A: ArchiveAccess>(
    archive: &mut A,
    name: &str,
) -> Result<T> {
    read_xml_entry(archive, name, PartPresence::Required)?
        .ok_or_else(|| Error::Internal(format!("required part was not read: {name}")))?
}

/// Read an optional XML part. Returns `None` when the entry is absent or its
/// intact bytes do not deserialize; such parts remain preserved as unknown
/// parts. Archive read, checksum, and size failures are returned as errors.
fn read_optional_xml_part<T: serde::de::DeserializeOwned, A: ArchiveAccess>(
    archive: &mut A,
    name: &str,
) -> Result<Option<T>> {
    Ok(read_xml_entry(archive, name, PartPresence::Optional)?.and_then(Result::ok))
}

/// Whether a missing archive entry is an error or an absent optional part.
#[derive(Clone, Copy, PartialEq, Eq)]
enum PartPresence {
    Required,
    Optional,
}

/// Stream-deserialize one XML entry. The outer result carries archive
/// failures; the inner result carries XML deserialization failures.
fn read_xml_entry<T: serde::de::DeserializeOwned, A: ArchiveAccess>(
    archive: &mut A,
    name: &str,
    presence: PartPresence,
) -> Result<Option<Result<T>>> {
    let (parsed, read_error) = {
        let entry = match presence {
            PartPresence::Required => archive.entry(name, true)?,
            PartPresence::Optional => match archive.entry_optional(name, true)? {
                Some(entry) => entry,
                None => return Ok(None),
            },
        };
        let mut source = PartReader::new(entry, None);
        let parsed = {
            let mut reader = std::io::BufReader::with_capacity(LARGE_BUF_CAPACITY, &mut source);
            let parsed: Result<T> = quick_xml::de::from_reader(&mut reader)
                .map_err(|error| Error::XmlDeserialize(error.to_string()));
            if parsed.is_ok() {
                // Read to EOF so the checksum and declared size are verified.
                let _ = std::io::copy(&mut reader, &mut std::io::sink());
            }
            parsed
        };
        (parsed, source.read_error)
    };
    if let Some(error) = archive.pending_error() {
        return Err(error);
    }
    if let Some(message) = read_error {
        return Err(Error::Zip(message));
    }
    Ok(Some(parsed))
}

/// Entry reader that records archive read failures separately from parser
/// failures and can mirror consumed bytes into a spool.
struct PartReader<R> {
    inner: R,
    copy: Option<std::io::BufWriter<std::fs::File>>,
    read_error: Option<String>,
    copy_error: Option<std::io::Error>,
}

impl<R> PartReader<R> {
    fn new(inner: R, copy: Option<std::fs::File>) -> Self {
        Self {
            inner,
            copy: copy.map(std::io::BufWriter::new),
            read_error: None,
            copy_error: None,
        }
    }
}

impl<R: std::io::Read> std::io::Read for PartReader<R> {
    fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
        let count = match self.inner.read(buffer) {
            Ok(count) => count,
            Err(error) => {
                if self.read_error.is_none() {
                    self.read_error = Some(error.to_string());
                }
                return Err(error);
            }
        };
        if let Some(copy) = &mut self.copy {
            if let Err(error) = copy.write_all(&buffer[..count]) {
                let message = error.to_string();
                self.copy_error.get_or_insert(error);
                return Err(std::io::Error::other(message));
            }
        }
        Ok(count)
    }
}

/// Decompress and parse one eagerly opened worksheet exactly once. When
/// `capture` is set, the decompressed bytes are also spooled to disk for
/// table and extension checks that must not decompress the entry again.
fn read_eager_worksheet<A: ArchiveAccess>(
    archive: &mut A,
    path: &str,
    row_limit: Option<u32>,
    capture: bool,
) -> Result<(WorksheetXml, Option<std::fs::File>)> {
    let spool = if capture {
        Some(tempfile::tempfile()?)
    } else {
        None
    };
    let (parsed, source) = {
        let entry = archive.entry(path, true)?;
        let mut source = PartReader::new(entry, spool);
        let parsed = {
            let mut reader = std::io::BufReader::with_capacity(LARGE_BUF_CAPACITY, &mut source);
            let parsed = deserialize_worksheet_xml_with_row_limit(&mut reader, row_limit);
            if parsed.is_ok() {
                let _ = std::io::copy(&mut reader, &mut std::io::sink());
            }
            parsed
        };
        (parsed, (source.read_error, source.copy_error, source.copy))
    };
    let (read_error, copy_error, copy) = source;
    if let Some(error) = archive.pending_error() {
        return Err(error);
    }
    if let Some(message) = read_error {
        return Err(Error::Zip(message));
    }
    if let Some(error) = copy_error {
        return Err(Error::Io(error));
    }
    let worksheet = parsed?;
    let spool = match copy {
        Some(writer) => {
            use std::io::Seek as _;
            let mut file = writer
                .into_inner()
                .map_err(|error| Error::Io(error.into_error()))?;
            file.seek(std::io::SeekFrom::Start(0))?;
            Some(file)
        }
        None => None,
    };
    Ok((worksheet, spool))
}

/// Reject package part names that differ only by ASCII case. OPC part names
/// are case-insensitive, so such entries alias one logical part.
fn validate_package_part_names(package_paths: &HashSet<String>) -> Result<()> {
    let mut folded = HashSet::with_capacity(package_paths.len());
    for path in package_paths {
        if !folded.insert(path.to_ascii_lowercase()) {
            return Err(Error::InvalidArgument(format!(
                "package part names differ only by case: {path}"
            )));
        }
    }
    Ok(())
}

/// Resolve each sheet's worksheet part and require a distinct, internal,
/// present worksheet part for every sheet.
fn resolve_worksheet_paths(
    workbook_xml: &WorkbookXml,
    workbook_rels: &Relationships,
    content_types: &ContentTypes,
    package_paths: &HashSet<String>,
) -> Result<Vec<String>> {
    const RESERVED_PARTS: [&str; 6] = [
        "[content_types].xml",
        "xl/workbook.xml",
        "xl/styles.xml",
        "xl/sharedstrings.xml",
        "xl/theme/theme1.xml",
        "xl/vbaproject.bin",
    ];
    let mut seen = HashSet::with_capacity(workbook_xml.sheets.sheets.len());
    let mut paths = Vec::with_capacity(workbook_xml.sheets.sheets.len());
    for sheet in &workbook_xml.sheets.sheets {
        let relationship = workbook_rels
            .relationships
            .iter()
            .find(|r| r.id == sheet.r_id && r.rel_type == rel_types::WORKSHEET)
            .ok_or_else(|| {
                Error::Internal(format!(
                    "missing worksheet relationship for sheet '{}'",
                    sheet.name
                ))
            })?;
        if relationship.target_mode.is_some() {
            return Err(Error::InvalidArgument(format!(
                "external worksheet relationship is not allowed: {}",
                relationship.id
            )));
        }
        if relationship_target_escapes("xl/workbook.xml", &relationship.target) {
            return Err(Error::InvalidArgument(format!(
                "worksheet relationship target escapes package root: {}",
                relationship.id
            )));
        }
        let path = resolve_relationship_target("xl/workbook.xml", &relationship.target);
        let folded = path.to_ascii_lowercase();
        if !package_paths.contains(&path) {
            return Err(Error::InvalidArgument(format!(
                "worksheet relationship target is missing from package: {path}"
            )));
        }
        if RESERVED_PARTS.contains(&folded.as_str())
            || folded.ends_with(".rels")
            || content_types.overrides.iter().any(|entry| {
                entry
                    .part_name
                    .trim_start_matches('/')
                    .eq_ignore_ascii_case(&path)
                    && entry.content_type != mime_types::WORKSHEET
            })
        {
            return Err(Error::InvalidArgument(format!(
                "worksheet relationship target is not a worksheet part: {path}"
            )));
        }
        if !seen.insert(folded) {
            return Err(Error::InvalidArgument(format!(
                "worksheet part is shared by multiple sheets: {path}"
            )));
        }
        paths.push(path);
    }
    Ok(paths)
}

/// Whether a relative or absolute target leaves the package root or uses a
/// URI scheme instead of a package part name.
fn relationship_target_escapes(source_part: &str, target: &str) -> bool {
    if target.is_empty() || target.contains('\\') || target.contains(':') {
        return true;
    }
    let mut depth = if target.starts_with('/') {
        0usize
    } else {
        source_part.matches('/').count()
    };
    for segment in target.trim_start_matches('/').split('/') {
        match segment {
            "" | "." => {}
            ".." => match depth.checked_sub(1) {
                Some(next) => depth = next,
                None => return true,
            },
            _ => depth += 1,
        }
    }
    false
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
fn read_shared_strings_part<A: ArchiveAccess>(archive: &mut A, name: &str) -> Result<Option<Sst>> {
    let Some(bytes) = read_optional_part(archive, name, true, false)? else {
        return Ok(None);
    };
    deserialize_shared_strings_from_bytes(&bytes).map(Some)
}

fn deserialize_shared_strings<R: std::io::BufRead>(reader: R) -> Result<Sst> {
    quick_xml::de::from_reader(reader).map_err(|e| Error::XmlDeserialize(e.to_string()))
}

fn restore_shared_string_text(mut sst: Sst, raw_items: Vec<RawSharedStringText>) -> Result<Sst> {
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

fn deserialize_shared_strings_from_bytes(bytes: &[u8]) -> Result<Sst> {
    let sst = deserialize_shared_strings(std::io::BufReader::new(bytes))?;
    if !shared_strings_need_raw_scan(bytes) {
        return Ok(sst);
    }
    let raw_items = scan_shared_string_text(std::io::BufReader::new(bytes))?;
    restore_shared_string_text(sst, raw_items)
}

fn shared_strings_need_raw_scan(bytes: &[u8]) -> bool {
    [b":space".as_slice(), b"<![CDATA[".as_slice()]
        .iter()
        .any(|marker| bytes.windows(marker.len()).any(|window| window == *marker))
}

fn scan_shared_string_text<R: std::io::BufRead>(reader: R) -> Result<Vec<RawSharedStringText>> {
    use quick_xml::events::Event;

    let mut reader = quick_xml::Reader::from_reader(reader);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut items = Vec::new();
    let mut current_item: Option<RawSharedStringText> = None;
    let mut text_target: Option<SharedStringTextTarget> = None;
    let mut text_value = String::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
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
        buf.clear();
    }

    if current_item.is_some() || text_target.is_some() {
        return Err(Error::XmlDeserialize(
            "unterminated shared string item".to_string(),
        ));
    }
    Ok(items)
}

/// Read a required ZIP entry as raw bytes, charging them as retained.
pub(crate) fn read_bytes_part<A: ArchiveAccess>(archive: &mut A, name: &str) -> Result<Vec<u8>> {
    read_part(
        archive,
        name,
        PartPresence::Required,
        is_xml_part(name),
        true,
    )?
    .ok_or_else(|| Error::Internal(format!("required part was not read: {name}")))
}

/// Read an optional ZIP entry as retained raw bytes. Only a missing entry
/// yields `None`; read, checksum, and size failures are returned as errors.
fn read_optional_bytes_part<A: ArchiveAccess>(
    archive: &mut A,
    name: &str,
) -> Result<Option<Vec<u8>>> {
    read_part(
        archive,
        name,
        PartPresence::Optional,
        is_xml_part(name),
        true,
    )
}

/// Read an optional ZIP entry with explicit XML and retention accounting.
fn read_optional_part<A: ArchiveAccess>(
    archive: &mut A,
    name: &str,
    xml: bool,
    retained: bool,
) -> Result<Option<Vec<u8>>> {
    read_part(archive, name, PartPresence::Optional, xml, retained)
}

fn read_part<A: ArchiveAccess>(
    archive: &mut A,
    name: &str,
    presence: PartPresence,
    xml: bool,
    retained: bool,
) -> Result<Option<Vec<u8>>> {
    let mut content = Vec::new();
    let read_result = {
        let mut entry = match presence {
            PartPresence::Required => archive.entry(name, xml)?,
            PartPresence::Optional => match archive.entry_optional(name, xml)? {
                Some(entry) => entry,
                None => return Ok(None),
            },
        };
        let mut chunk = [0_u8; LARGE_BUF_CAPACITY];
        loop {
            let count = match entry.read(&mut chunk) {
                Ok(0) => break Ok(()),
                Ok(count) => count,
                Err(error) => break Err(error),
            };
            content.try_reserve(count).map_err(|_| {
                Error::Io(std::io::Error::new(
                    std::io::ErrorKind::OutOfMemory,
                    "unable to reserve bounded workbook part buffer",
                ))
            })?;
            content.extend_from_slice(&chunk[..count]);
        }
    };
    if let Some(error) = archive.pending_error() {
        return Err(error);
    }
    read_result.map_err(|error| Error::Zip(error.to_string()))?;
    if retained {
        archive.charge_retained(content.len() as u64)?;
    }
    Ok(Some(content))
}

fn is_xml_part(name: &str) -> bool {
    name.ends_with(".xml") || name.ends_with(".rels") || name.ends_with(".vml")
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
    controls: &[(usize, String, String)],
    controls_dirty: bool,
) -> Result<String> {
    let mut xml = serialize_worksheet_with_extras(ws, sparklines, legacy_drawing_rid)?;
    if controls_dirty {
        xml = replace_worksheet_controls(xml, controls)?;
    }
    if slicer_rids.is_empty() {
        // Sparkline groups carry a broad set of optional style, color, and
        // axis children which the public configuration does not model. When
        // a structural edit has patched the retained worksheet XML, prefer
        // that original extension list over rebuilding a reduced group.
        if let Some(raw_xml) = raw_sheet.and_then(|bytes| std::str::from_utf8(bytes).ok()) {
            let raw_sparklines = parse_sparklines_from_xml(raw_xml);
            if raw_sparklines.as_slice() == sparklines {
                if let Some(raw_ext) = ext_list_contents(raw_xml)
                    .filter(|contents| contents.contains("sparklineGroups"))
                {
                    xml = remove_extension_list(xml);
                    let closing = "</worksheet>";
                    let position = xml.rfind(closing).ok_or_else(|| {
                        Error::XmlParse("worksheet closing tag missing".to_string())
                    })?;
                    xml.insert_str(position, &format!("<extLst>{raw_ext}</extLst>"));
                }
            }
        }
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

fn replace_worksheet_controls(
    mut xml: String,
    controls: &[(usize, String, String)],
) -> Result<String> {
    if let Some(start) = xml.find("<controls") {
        let end = xml[start..]
            .find("</controls>")
            .map(|offset| start + offset + "</controls>".len())
            .ok_or_else(|| Error::XmlParse("worksheet controls closing tag missing".to_string()))?;
        xml.replace_range(start..end, "");
    }
    if controls.is_empty() {
        return Ok(xml);
    }

    let entries = controls
        .iter()
        .map(|(shape_id, rid, name)| {
            format!(
                "<control shapeId=\"{shape_id}\" r:id=\"{}\" name=\"{}\"/>",
                quick_xml::escape::escape(rid),
                quick_xml::escape::escape(name),
            )
        })
        .collect::<String>();
    let controls_xml = format!("<controls>{entries}</controls>");
    let position = xml
        .find("<tableParts")
        .or_else(|| xml.find("<extLst"))
        .or_else(|| xml.rfind("</worksheet>"))
        .ok_or_else(|| Error::XmlParse("worksheet closing tag missing".to_string()))?;
    xml.insert_str(position, &controls_xml);
    Ok(xml)
}

fn parse_worksheet_control_entries(xml: &str) -> Vec<(usize, String, String)> {
    let mut entries = Vec::new();
    let mut search_from = 0;
    while let Some(relative_start) = xml[search_from..].find("<control ") {
        let start = search_from + relative_start;
        let Some(relative_end) = xml[start..].find("/>") else {
            break;
        };
        let end = start + relative_end + 2;
        let element = &xml[start..end];
        let attribute = |name: &str| {
            let marker = format!("{name}=\"");
            let value_start = element.find(&marker)? + marker.len();
            let value_end = element[value_start..].find('"')? + value_start;
            Some(element[value_start..value_end].to_string())
        };
        if let (Some(shape_id), Some(rid), Some(name)) =
            (attribute("shapeId"), attribute("r:id"), attribute("name"))
        {
            if let Ok(shape_id) = shape_id.parse() {
                entries.push((shape_id, rid, name));
            }
        }
        search_from = end;
    }
    entries
}

fn parse_worksheet_control_entries_reader<R: std::io::BufRead>(
    reader: R,
) -> Result<Vec<(usize, String, String)>> {
    let mut reader = quick_xml::Reader::from_reader(reader);
    let mut buffer = Vec::with_capacity(4096);
    let mut entries = Vec::new();
    loop {
        buffer.clear();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::XmlParse(error.to_string()))?;
        match event {
            quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element)
                if element.local_name().as_ref() == b"control" =>
            {
                let mut shape_id = None;
                let mut relationship_id = None;
                let mut name = None;
                for attribute in element.attributes() {
                    let attribute =
                        attribute.map_err(|error| Error::XmlParse(error.to_string()))?;
                    let value = attribute
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(|error| Error::XmlParse(error.to_string()))?
                        .into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"shapeId" => shape_id = value.parse::<usize>().ok(),
                        b"id" => relationship_id = Some(value),
                        b"name" => name = Some(value),
                        _ => {}
                    }
                }
                if let (Some(shape_id), Some(relationship_id), Some(name)) =
                    (shape_id, relationship_id, name)
                {
                    entries.push((shape_id, relationship_id, name));
                }
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    Ok(entries)
}

fn next_control_property_path(existing: &HashSet<String>) -> String {
    let mut number = 1usize;
    loop {
        let candidate = format!("xl/ctrlProps/ctrlProp{number}.xml");
        if !existing.contains(&candidate) {
            return candidate;
        }
        number += 1;
    }
}

fn next_numbered_package_path(existing: &HashSet<String>, prefix: &str, suffix: &str) -> String {
    let mut number = 1usize;
    loop {
        let candidate = format!("{prefix}{number}{suffix}");
        if !existing.contains(&candidate) {
            return candidate;
        }
        number += 1;
    }
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
                quick_xml::escape::escape(&sp.formula),
                quick_xml::escape::escape(&sp.sqref)
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
    let value = &xml[content_start..content_start + end];
    quick_xml::escape::unescape(value)
        .map(|value| value.into_owned())
        .ok()
}

/// Serialize a value to XML and write it as a ZIP entry.
pub(crate) fn write_xml_part<T: Serialize, W: std::io::Write + std::io::Seek>(
    zip: &mut zip::ZipWriter<W>,
    name: &str,
    value: &T,
    options: SimpleFileOptions,
    limit: u64,
) -> Result<()> {
    use std::fmt::Write as _;

    struct XmlPartWriter<'a, W: std::io::Write + std::io::Seek> {
        zip: &'a mut zip::ZipWriter<W>,
        limit: u64,
        written: u64,
        error: Option<Error>,
    }

    impl<W: std::io::Write + std::io::Seek> std::fmt::Write for XmlPartWriter<'_, W> {
        fn write_str(&mut self, value: &str) -> std::fmt::Result {
            if self.error.is_some() {
                return Err(std::fmt::Error);
            }
            let amount = u64::try_from(value.len()).unwrap_or(u64::MAX);
            let observed = self.written.saturating_add(amount);
            if observed > self.limit {
                self.error = Some(crate::save::resource_limit(
                    crate::resource::ResourceOperation::WorkbookSave,
                    crate::resource::ResourceCategory::Output,
                    crate::resource::ResourceDimension::XmlBytes,
                    self.limit,
                    observed,
                    crate::resource::ResourceStage::Serialize,
                ));
                return Err(std::fmt::Error);
            }
            if let Err(error) = self.zip.write_all(value.as_bytes()) {
                self.error = Some(Error::Io(error));
                return Err(std::fmt::Error);
            }
            self.written = observed;
            Ok(())
        }
    }

    zip.start_file(name, options)
        .map_err(|error| Error::Zip(error.to_string()))?;
    let mut output = XmlPartWriter {
        zip,
        limit,
        written: 0,
        error: None,
    };
    let declaration = writeln!(output, "{XML_DECLARATION}");
    if let Some(error) = output.error.take() {
        return Err(error);
    }
    declaration.map_err(|_| Error::Internal("XML declaration serialization failed".to_string()))?;
    let result = value.serialize(quick_xml::se::Serializer::new(&mut output));
    if let Some(error) = output.error {
        return Err(error);
    }
    result.map_err(|error| Error::XmlParse(error.to_string()))
}

fn write_bytes_part<W: std::io::Write + std::io::Seek>(
    zip: &mut zip::ZipWriter<W>,
    name: &str,
    bytes: &[u8],
    options: SimpleFileOptions,
    limit: u64,
) -> Result<()> {
    reject_save_metric(
        crate::resource::ResourceDimension::XmlBytes,
        u64::try_from(bytes.len()).unwrap_or(u64::MAX),
        limit,
    )?;
    zip.start_file(name, options)
        .map_err(|e| Error::Zip(e.to_string()))?;
    zip.write_all(bytes)?;
    Ok(())
}

fn write_package_bytes_part<W: std::io::Write + std::io::Seek>(
    zip: &mut zip::ZipWriter<W>,
    name: &str,
    bytes: &[u8],
    options: SimpleFileOptions,
    xml_limit: u64,
) -> Result<()> {
    let extension = name.rsplit_once('.').map(|(_, extension)| extension);
    if extension.is_some_and(|extension| {
        extension.eq_ignore_ascii_case("xml")
            || extension.eq_ignore_ascii_case("rels")
            || extension.eq_ignore_ascii_case("vml")
    }) {
        return write_bytes_part(zip, name, bytes, options, xml_limit);
    }
    zip.start_file(name, options)
        .map_err(|error| Error::Zip(error.to_string()))?;
    zip.write_all(bytes)?;
    Ok(())
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

    fn assert_all_package_targets_exist(buffer: &[u8]) {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer)).unwrap();
        let names: HashSet<String> = (0..archive.len())
            .map(|index| archive.by_index(index).unwrap().name().to_string())
            .collect();
        let content_types: ContentTypes = {
            let mut xml = String::new();
            archive
                .by_name("[Content_Types].xml")
                .unwrap()
                .read_to_string(&mut xml)
                .unwrap();
            quick_xml::de::from_str(&xml).unwrap()
        };
        for entry in content_types.overrides {
            let path = entry.part_name.trim_start_matches('/');
            if matches!(
                path,
                "docProps/core.xml" | "docProps/app.xml" | "docProps/custom.xml"
            ) && !names.contains(path)
            {
                continue;
            }
            assert!(
                names.contains(path),
                "content type target missing: {}",
                entry.part_name
            );
        }

        let relationship_paths: Vec<String> = names
            .iter()
            .filter(|name| name.ends_with(".rels"))
            .cloned()
            .collect();
        for relationship_path in relationship_paths {
            let relationships: Relationships = {
                let mut xml = String::new();
                archive
                    .by_name(&relationship_path)
                    .unwrap()
                    .read_to_string(&mut xml)
                    .unwrap();
                quick_xml::de::from_str(&xml).unwrap()
            };
            let owner = if relationship_path == "_rels/.rels" {
                String::new()
            } else {
                let (directory, filename) = relationship_path.rsplit_once("/_rels/").unwrap();
                format!("{directory}/{}", filename.trim_end_matches(".rels"))
            };
            for relationship in relationships.relationships {
                if relationship.target_mode.as_deref() == Some("External") {
                    continue;
                }
                let target = resolve_relationship_target(&owner, &relationship.target);
                if relationship_path == "_rels/.rels"
                    && matches!(
                        relationship.rel_type.as_str(),
                        rel_types::CORE_PROPERTIES
                            | rel_types::EXTENDED_PROPERTIES
                            | rel_types::CUSTOM_PROPERTIES
                    )
                    && !names.contains(&target)
                {
                    continue;
                }
                assert!(
                    names.contains(&target),
                    "relationship target missing: {relationship_path} -> {target}"
                );
            }
        }
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
        let sst = deserialize_shared_strings_from_bytes(xml).unwrap();
        assert_eq!(sst.items[0].t.as_ref().unwrap().value, "  plain\tline\n&  ");
        assert_eq!(sst.items[1].t.as_ref().unwrap().value, " direct ");
        assert_eq!(sst.items[1].r[0].t.value, " rich ");
        assert_eq!(sst.items[1].r[1].t.value, "tail  ");
        assert!(sst.items[1].r[0].r_pr.as_ref().unwrap().b.is_some());
        assert_eq!(sst.items[2].t.as_ref().unwrap().value, "entity <value>");
    }

    #[test]
    fn shared_string_raw_scan_fast_path_preserves_signaled_text() {
        let ordinary = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>ordinary text</t></si><si><r><t>rich text</t></r></si></sst>"#;
        assert!(!shared_strings_need_raw_scan(ordinary));
        let ordinary_sst = deserialize_shared_strings_from_bytes(ordinary).unwrap();
        assert_eq!(
            ordinary_sst.items[0].t.as_ref().unwrap().value,
            "ordinary text"
        );
        assert_eq!(ordinary_sst.items[1].r[0].t.value, "rich text");

        let single_quoted = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><x:t xml:space='preserve' xmlns:x="urn:test">  kept  </x:t></si></sst>"#;
        assert!(shared_strings_need_raw_scan(single_quoted));
        let single_quoted_sst = deserialize_shared_strings_from_bytes(single_quoted).unwrap();
        assert_eq!(
            single_quoted_sst.items[0].t.as_ref().unwrap().value,
            "  kept  "
        );

        let cdata = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t><![CDATA[  cdata  ]]></t></si></sst>"#;
        assert!(shared_strings_need_raw_scan(cdata));
        let cdata_sst = deserialize_shared_strings_from_bytes(cdata).unwrap();
        assert_eq!(cdata_sst.items[0].t.as_ref().unwrap().value, "  cdata  ");
    }

    #[test]
    fn scan_shared_string_text_reads_from_buffered_stream() {
        struct ChunkedReader {
            bytes: Vec<u8>,
            offset: usize,
        }

        impl std::io::Read for ChunkedReader {
            fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
                if self.offset == self.bytes.len() {
                    return Ok(0);
                }
                let len = (self.bytes.len() - self.offset).min(buf.len()).min(3);
                buf[..len].copy_from_slice(&self.bytes[self.offset..self.offset + len]);
                self.offset += len;
                Ok(len)
            }
        }

        let reader = std::io::BufReader::with_capacity(
            5,
            ChunkedReader {
                bytes: b"<sst><si><t>  spaced  </t></si></sst>".to_vec(),
                offset: 0,
            },
        );
        let items = scan_shared_string_text(reader).unwrap();
        assert_eq!(items[0].plain.as_deref(), Some("  spaced  "));
    }

    #[test]
    fn read_shared_strings_part_preserves_highly_compressible_whitespace() {
        use std::io::Write;

        let padding = " ".repeat(1024 * 1024);
        let xml = format!(
            "<sst xmlns=\"{}\"><si><t xml:space=\"preserve\">{padding}value{padding}</t></si></sst>",
            sheetkit_xml::namespaces::SPREADSHEET_ML
        );
        let mut zip_bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(&mut zip_bytes));
            writer
                .start_file(
                    "xl/sharedStrings.xml",
                    SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
                )
                .unwrap();
            writer.write_all(xml.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        assert!(zip_bytes.len() < xml.len() / 100);

        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(zip_bytes)).unwrap();
        let sst = read_shared_strings_part(&mut archive, "xl/sharedStrings.xml")
            .unwrap()
            .unwrap();
        assert_eq!(
            sst.items[0].t.as_ref().unwrap().value,
            format!("{padding}value{padding}")
        );
    }

    #[test]
    fn read_shared_strings_part_opens_and_reads_entry_once() {
        struct CountingEntry<'archive> {
            reader: std::io::Cursor<&'archive [u8]>,
            actual_bytes: &'archive mut usize,
        }

        impl std::io::Read for CountingEntry<'_> {
            fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
                let count = self.reader.read(buffer)?;
                *self.actual_bytes += count;
                Ok(count)
            }
        }

        struct CountingArchive {
            bytes: Vec<u8>,
            opens: usize,
            actual_bytes: usize,
        }

        impl ArchiveAccess for CountingArchive {
            type Entry<'entry>
                = CountingEntry<'entry>
            where
                Self: 'entry;

            fn entry_names(&self) -> Vec<String> {
                vec!["xl/sharedStrings.xml".to_string()]
            }

            fn entry<'entry>(
                &'entry mut self,
                name: &str,
                _xml: bool,
            ) -> Result<Self::Entry<'entry>> {
                assert_eq!(name, "xl/sharedStrings.xml");
                self.opens += 1;
                Ok(CountingEntry {
                    reader: std::io::Cursor::new(&self.bytes),
                    actual_bytes: &mut self.actual_bytes,
                })
            }

            fn entry_optional<'entry>(
                &'entry mut self,
                name: &str,
                xml: bool,
            ) -> Result<Option<Self::Entry<'entry>>> {
                self.entry(name, xml).map(Some)
            }

            fn pending_error(&mut self) -> Option<Error> {
                None
            }

            fn charge_retained(&mut self, _amount: u64) -> Result<()> {
                Ok(())
            }
        }

        let bytes = format!(
            "<sst xmlns=\"{}\"><si><t xml:space=\"preserve\"> value </t></si></sst>",
            sheetkit_xml::namespaces::SPREADSHEET_ML
        )
        .into_bytes();
        let mut archive = CountingArchive {
            bytes,
            opens: 0,
            actual_bytes: 0,
        };

        let sst = read_shared_strings_part(&mut archive, "xl/sharedStrings.xml")
            .unwrap()
            .unwrap();

        assert_eq!(archive.opens, 1);
        assert_eq!(archive.actual_bytes, archive.bytes.len());
        assert_eq!(sst.items[0].t.as_ref().unwrap().value, " value ");
    }

    #[test]
    fn clean_control_metadata_round_trips_without_rebuild() {
        let mut workbook = Workbook::new();
        workbook
            .add_form_control(
                "Sheet1",
                crate::control::FormControlConfig::button("B2", "Run"),
            )
            .unwrap();
        let first = workbook.save_to_buffer().unwrap();
        let options = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        let reopened = Workbook::open_from_buffer_with_options(&first, &options).unwrap();

        assert!(!reopened.sheet_controls_dirty[0]);
        let second = reopened.save_to_buffer().unwrap();

        let first_sheet = String::from_utf8(zip_part(&first, "xl/worksheets/sheet1.xml")).unwrap();
        let second_sheet =
            String::from_utf8(zip_part(&second, "xl/worksheets/sheet1.xml")).unwrap();
        assert_eq!(
            parse_worksheet_control_entries(&second_sheet),
            parse_worksheet_control_entries(&first_sheet)
        );
        assert!(second_sheet.contains("<legacyDrawing r:id=\"rId2\"/>"));
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
    fn lazy_vba_open_applies_exact_and_plus_one_input_limits() {
        let vba_data = b"FINITE_VBA_PROJECT";
        let xlsm = build_xlsm_with_vba(vba_data);
        let exact = crate::vba::VbaLimits {
            input_bytes: vba_data.len() as u64,
            ..crate::vba::VbaLimits::default()
        };
        let workbook =
            Workbook::open_from_buffer_with_options(&xlsm, &OpenOptions::new().vba_limits(exact))
                .unwrap();
        assert_eq!(workbook.get_vba_project(), Some(vba_data.as_slice()));

        let rejected = crate::vba::VbaLimits {
            input_bytes: vba_data.len() as u64 - 1,
            ..crate::vba::VbaLimits::default()
        };
        let error = match Workbook::open_from_buffer_with_options(
            &xlsm,
            &OpenOptions::new().vba_limits(rejected),
        ) {
            Ok(_) => panic!("expected VBA input limit"),
            Err(error) => error,
        };
        assert_eq!(
            error.resource_metadata().unwrap().dimension(),
            crate::resource::ResourceDimension::InputBytes
        );
    }

    #[test]
    fn macro_free_target_strips_vba_without_mutating_source() {
        let vba_data = b"VALIDATED_VBA_PROJECT";
        let source = build_xlsm_with_vba(vba_data);
        let workbook = Workbook::open_from_buffer(&source).unwrap();
        let directory = TempDir::new().unwrap();
        let xlsx_path = directory.path().join("converted.xlsx");
        let xlsm_path = directory.path().join("preserved.xlsm");

        workbook.save(&xlsx_path).unwrap();
        let mut xlsx = zip::ZipArchive::new(std::fs::File::open(&xlsx_path).unwrap()).unwrap();
        assert!(xlsx.by_name("xl/vbaProject.bin").is_err());
        let content_types: ContentTypes = read_xml_part(&mut xlsx, "[Content_Types].xml").unwrap();
        assert!(content_types
            .overrides
            .iter()
            .all(|entry| entry.content_type != VBA_PROJECT_CONTENT_TYPE));
        let relationships: Relationships =
            read_xml_part(&mut xlsx, "xl/_rels/workbook.xml.rels").unwrap();
        assert!(relationships
            .relationships
            .iter()
            .all(|entry| entry.rel_type != VBA_PROJECT_REL_TYPE));
        drop(xlsx);

        workbook.save(&xlsm_path).unwrap();
        let mut xlsm = zip::ZipArchive::new(std::fs::File::open(&xlsm_path).unwrap()).unwrap();
        let mut preserved = Vec::new();
        xlsm.by_name("xl/vbaProject.bin")
            .unwrap()
            .read_to_end(&mut preserved)
            .unwrap();
        assert_eq!(preserved, vba_data);
    }

    #[cfg(feature = "encryption")]
    #[test]
    fn encrypted_save_infers_target_format_for_vba_filtering() {
        let password = "bounded-password";
        let vba_data = b"ENCRYPTED_VBA_PROJECT";
        let workbook = Workbook::open_from_buffer(&build_xlsm_with_vba(vba_data)).unwrap();
        let directory = TempDir::new().unwrap();

        for (extension, preserves_vba) in [("xlsx", false), ("xlsm", true)] {
            let path = directory.path().join(format!("encrypted.{extension}"));
            workbook.save_with_password(&path, password).unwrap();
            let encrypted = std::fs::read(path).unwrap();
            let decrypted = crate::crypt::decrypt_xlsx(&encrypted, password).unwrap();
            let mut archive = zip::ZipArchive::new(std::io::Cursor::new(decrypted)).unwrap();
            assert_eq!(archive.by_name("xl/vbaProject.bin").is_ok(), preserves_vba);
        }
    }

    #[cfg(feature = "encryption")]
    #[test]
    fn encrypted_save_limits_accept_exact_package_and_reject_plus_one() {
        let workbook = Workbook::new();
        let directory = TempDir::new().unwrap();
        let baseline_path = directory.path().join("baseline.xlsx");
        let exact_path = directory.path().join("exact.xlsx");
        let rejected_path = directory.path().join("rejected.xlsx");
        workbook
            .save_with_password(&baseline_path, "password")
            .unwrap();
        let baseline = std::fs::read(&baseline_path).unwrap();
        let compound = cfb::CompoundFile::open(std::io::Cursor::new(baseline)).unwrap();
        let package_limit = compound.entry("/EncryptedPackage").unwrap().len();

        workbook
            .save_with_password_with_limits(
                &exact_path,
                "password",
                crate::crypt::CryptLimits {
                    encrypted_package_bytes: package_limit,
                    ..crate::crypt::CryptLimits::default()
                },
            )
            .unwrap();
        let error = workbook
            .save_with_password_with_limits(
                &rejected_path,
                "password",
                crate::crypt::CryptLimits {
                    encrypted_package_bytes: package_limit - 1,
                    ..crate::crypt::CryptLimits::default()
                },
            )
            .unwrap_err();
        let metadata = error.resource_metadata().expect("resource metadata");
        assert_eq!(metadata.operation(), crate::ResourceOperation::Encrypt);
        assert_eq!(metadata.dimension(), crate::ResourceDimension::InputBytes);
        assert_eq!(metadata.limit(), package_limit - 1);
        assert_eq!(metadata.observed(), package_limit);
        assert!(!rejected_path.exists());
    }

    #[test]
    fn failed_atomic_xlsx_stage_preserves_destination_and_workbook() {
        let vba_data = b"REUSABLE_VBA_PROJECT";
        let workbook = Workbook::open_from_buffer(&build_xlsm_with_vba(vba_data)).unwrap();
        let directory = TempDir::new().unwrap();
        let destination = directory.path().join("existing.xlsx");
        let later = directory.path().join("later.xlsm");
        let original = b"existing destination bytes";
        std::fs::write(&destination, original).unwrap();

        let error = atomic_write_file(&destination, |file| {
            let mut zip = zip::ZipWriter::new(file);
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            workbook.write_zip_contents(
                &mut zip,
                options,
                Some(WorkbookFormat::Xlsx),
                crate::save::SaveLimits::default(),
            )?;
            zip.finish()
                .map_err(|error| Error::Zip(error.to_string()))?;
            Err(Error::Internal(
                "injected before atomic replace".to_string(),
            ))
        })
        .unwrap_err();
        assert!(matches!(error, Error::Internal(_)));
        assert_eq!(std::fs::read(&destination).unwrap(), original);
        assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 1);

        workbook.save(&later).unwrap();
        let mut archive = zip::ZipArchive::new(std::fs::File::open(later).unwrap()).unwrap();
        let mut preserved = Vec::new();
        archive
            .by_name("xl/vbaProject.bin")
            .unwrap()
            .read_to_end(&mut preserved)
            .unwrap();
        assert_eq!(preserved, vba_data);
    }

    #[test]
    fn post_rename_directory_sync_failure_does_not_report_save_failure() {
        let directory = TempDir::new().unwrap();
        let destination = directory.path().join("existing.xlsx");
        std::fs::write(&destination, b"old contents").unwrap();

        atomic_write_file_with_parent_sync(
            &destination,
            |file| {
                file.write_all(b"new contents")?;
                Ok(())
            },
            |_| Err(std::io::Error::other("injected directory sync failure")),
        )
        .unwrap();

        assert_eq!(std::fs::read(&destination).unwrap(), b"new contents");
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
        let Err(Error::ResourceLimitExceeded(metadata)) = result else {
            panic!("expected an entry count resource error");
        };
        assert_eq!(metadata.dimension(), crate::ResourceDimension::EntryCount);
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
        let Err(Error::ResourceLimitExceeded(metadata)) = result else {
            panic!("expected an aggregate declared bytes resource error");
        };
        assert_eq!(
            metadata.dimension(),
            crate::ResourceDimension::AggregateDeclaredBytes
        );
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
    fn malformed_cell_structure_fails_in_eager_lazy_and_stream_modes() {
        let mut workbook = Workbook::new();
        workbook.set_cell_value("Sheet1", "A1", 1.0).unwrap();
        let base = workbook.save_to_buffer().unwrap();
        let sheet_path = "xl/worksheets/sheet1.xml";
        let malformed = String::from_utf8(zip_part(&base, sheet_path))
            .unwrap()
            .replacen("r=\"A1\"", "r=\"A2\"", 1);
        let input = rewrite_zip_parts(&base, &[(sheet_path, Some(malformed.into_bytes()))]);
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("malformed-cell.xlsx");
        std::fs::write(&path, &input).unwrap();

        let assert_invalid = |result: Result<Workbook>| match result {
            Err(error) => assert!(matches!(error, Error::InvalidCellReference(_))),
            Ok(workbook) => {
                let error = workbook.get_cell_value("Sheet1", "A1").unwrap_err();
                assert!(matches!(error, Error::InvalidCellReference(_)));
            }
        };

        for mode in [ReadMode::Eager, ReadMode::Lazy, ReadMode::Stream] {
            let options = OpenOptions::new().read_mode(mode);
            assert_invalid(Workbook::open_from_buffer_with_options(&input, &options));
            assert_invalid(Workbook::open_with_options(&path, &options));
        }
    }

    #[test]
    fn table_ownership_fails_closed_in_every_read_mode() {
        let mut workbook = Workbook::new();
        workbook
            .add_table(
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
        let base = workbook.save_to_buffer().unwrap();
        let relationships_path = "xl/worksheets/_rels/sheet1.xml.rels";
        let relationships_xml = zip_part(&base, relationships_path);
        let relationships: Relationships = deserialize_xml_bytes(&relationships_xml).unwrap();
        let table_relationship_id = relationships
            .relationships
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::TABLE)
            .unwrap()
            .id
            .clone();
        let sheet_path = "xl/worksheets/sheet1.xml";
        let sheet_xml = String::from_utf8(zip_part(&base, sheet_path)).unwrap();
        let table_parts = format!(
            "<tableParts count=\"1\"><tablePart r:id=\"{table_relationship_id}\"/></tableParts>"
        );
        assert!(sheet_xml.contains(&table_parts));

        let mut orphan = relationships.clone();
        orphan
            .relationships
            .retain(|relationship| relationship.rel_type != rel_types::TABLE);
        let orphan = rewrite_zip_parts(
            &base,
            &[(
                relationships_path,
                Some(serialize_xml(&orphan).unwrap().into_bytes()),
            )],
        );

        let mut external = relationships.clone();
        external
            .relationships
            .iter_mut()
            .find(|relationship| relationship.rel_type == rel_types::TABLE)
            .unwrap()
            .target_mode = Some("External".to_string());
        let external = rewrite_zip_parts(
            &base,
            &[(
                relationships_path,
                Some(serialize_xml(&external).unwrap().into_bytes()),
            )],
        );

        let mut traversal = relationships.clone();
        traversal
            .relationships
            .iter_mut()
            .find(|relationship| relationship.rel_type == rel_types::TABLE)
            .unwrap()
            .target = "../../../tables/table1.xml".to_string();
        let traversal = rewrite_zip_parts(
            &base,
            &[(
                relationships_path,
                Some(serialize_xml(&traversal).unwrap().into_bytes()),
            )],
        );

        let mut duplicate = relationships;
        let mut duplicate_relationship = duplicate
            .relationships
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::TABLE)
            .unwrap()
            .clone();
        duplicate_relationship.id = "duplicateTableOwner".to_string();
        duplicate.relationships.push(duplicate_relationship);
        let duplicate = rewrite_zip_parts(
            &base,
            &[(
                relationships_path,
                Some(serialize_xml(&duplicate).unwrap().into_bytes()),
            )],
        );

        let missing = rewrite_zip_parts(&base, &[("xl/tables/table1.xml", None)]);

        let missing_table_parts = rewrite_zip_parts(
            &base,
            &[(
                sheet_path,
                Some(sheet_xml.replacen(&table_parts, "", 1).into_bytes()),
            )],
        );
        let unknown_table_part = rewrite_zip_parts(
            &base,
            &[(
                sheet_path,
                Some(
                    sheet_xml
                        .replacen(
                            &format!("r:id=\"{table_relationship_id}\""),
                            "r:id=\"unknownTableRelationship\"",
                            1,
                        )
                        .into_bytes(),
                ),
            )],
        );
        let duplicate_table_part = rewrite_zip_parts(
            &base,
            &[(
                sheet_path,
                Some(
                    sheet_xml
                        .replacen(
                            &format!("<tablePart r:id=\"{table_relationship_id}\"/>"),
                            &format!(
                                "<tablePart r:id=\"{table_relationship_id}\"/><tablePart r:id=\"{table_relationship_id}\"/>"
                            ),
                            1,
                        )
                        .into_bytes(),
                ),
            )],
        );
        let mismatched_table_part_count = rewrite_zip_parts(
            &base,
            &[(
                sheet_path,
                Some(
                    sheet_xml
                        .replacen("<tableParts count=\"1\">", "<tableParts count=\"2\">", 1)
                        .into_bytes(),
                ),
            )],
        );
        let foreign_table_parts_xml = sheet_xml
            .replacen("<worksheet ", "<worksheet xmlns:f=\"urn:foreign\" ", 1)
            .replacen(
                &table_parts,
                &format!(
                    "<f:tableParts count=\"1\"><f:tablePart f:id=\"{table_relationship_id}\"/></f:tableParts>"
                ),
                1,
            );
        let foreign_table_parts = rewrite_zip_parts(
            &base,
            &[(sheet_path, Some(foreign_table_parts_xml.into_bytes()))],
        );
        let expanded_table_part = rewrite_zip_parts(
            &base,
            &[(
                sheet_path,
                Some(
                    sheet_xml
                        .replacen(
                            &format!("<tablePart r:id=\"{table_relationship_id}\"/>"),
                            &format!("<tablePart r:id=\"{table_relationship_id}\"></tablePart>"),
                            1,
                        )
                        .into_bytes(),
                ),
            )],
        );
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("invalid-table-owner.xlsx");

        for input in [
            &orphan,
            &external,
            &traversal,
            &duplicate,
            &missing,
            &missing_table_parts,
            &unknown_table_part,
            &duplicate_table_part,
            &mismatched_table_part_count,
            &foreign_table_parts,
        ] {
            std::fs::write(&path, input).unwrap();
            for mode in [ReadMode::Eager, ReadMode::Lazy, ReadMode::Stream] {
                let options = OpenOptions::new().read_mode(mode);
                assert!(Workbook::open_from_buffer_with_options(input, &options).is_err());
                assert!(Workbook::open_with_options(&path, &options).is_err());
            }
        }

        std::fs::write(&path, &expanded_table_part).unwrap();
        for mode in [ReadMode::Eager, ReadMode::Lazy, ReadMode::Stream] {
            let options = OpenOptions::new().read_mode(mode);
            Workbook::open_from_buffer_with_options(&expanded_table_part, &options).unwrap();
            Workbook::open_with_options(&path, &options).unwrap();
        }
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
    fn test_lazy_sparse_comment_paths_remain_package_consistent_after_mutation() {
        use crate::workbook::aux::AuxCategory;

        let mut workbook = Workbook::new();
        workbook.new_sheet("Sheet2").unwrap();
        for (sheet, cell) in [("Sheet1", "A1"), ("Sheet2", "B2")] {
            workbook
                .add_comment(
                    sheet,
                    &crate::comment::CommentConfig {
                        cell: cell.to_string(),
                        author: "Author".to_string(),
                        text: format!("Comment on {sheet}"),
                    },
                )
                .unwrap();
        }
        let options = OpenOptions::new().read_mode(ReadMode::Lazy);
        let base = workbook.save_to_buffer().unwrap();
        let mut sparse = Workbook::open_from_buffer_with_options(&base, &options).unwrap();

        for (sheet_idx, old_number, new_number) in [(0, 1, 7), (1, 2, 9)] {
            let sheet_path = sparse.sheet_part_path(sheet_idx);
            for (category, old_path, new_path) in [
                (
                    AuxCategory::Comments,
                    format!("xl/comments{old_number}.xml"),
                    format!("xl/comments{new_number}.xml"),
                ),
                (
                    AuxCategory::Vml,
                    format!("xl/drawings/vmlDrawing{old_number}.vml"),
                    format!("xl/drawings/vmlDrawing{new_number}.vml"),
                ),
            ] {
                let bytes = sparse
                    .deferred_parts
                    .remove_path(category, &old_path)
                    .unwrap();
                assert!(sparse.deferred_parts.insert(new_path.clone(), bytes));
                for entry in &mut sparse.content_types.overrides {
                    if entry.part_name == format!("/{old_path}") {
                        entry.part_name = format!("/{new_path}");
                    }
                }
                for relationship in &mut sparse
                    .worksheet_rels
                    .get_mut(&sheet_idx)
                    .unwrap()
                    .relationships
                {
                    if resolve_relationship_target(&sheet_path, &relationship.target) == old_path {
                        relationship.target = relative_relationship_target(&sheet_path, &new_path);
                    }
                }
            }
        }
        let sparse_package = sparse.save_to_buffer().unwrap();
        let mut opened =
            Workbook::open_from_buffer_with_options(&sparse_package, &options).unwrap();
        opened
            .add_comment(
                "Sheet1",
                &crate::comment::CommentConfig {
                    cell: "C3".to_string(),
                    author: "Author".to_string(),
                    text: "Added comment".to_string(),
                },
            )
            .unwrap();

        let saved = opened.save_to_buffer().unwrap();
        assert_all_package_targets_exist(&saved);
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&saved)).unwrap();
        for path in [
            "xl/comments7.xml",
            "xl/comments9.xml",
            "xl/drawings/vmlDrawing7.vml",
            "xl/drawings/vmlDrawing9.vml",
        ] {
            assert!(
                archive.by_name(path).is_ok(),
                "missing retained part: {path}"
            );
        }
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
    fn lazy_sparkline_addition_bypasses_clean_package_raw_copy() {
        let base = Workbook::new().save_to_buffer().unwrap();
        let mut opened = Workbook::open_from_buffer_with_options(
            &base,
            &OpenOptions::new().read_mode(ReadMode::Lazy),
        )
        .unwrap();
        opened
            .add_sparkline(
                "Sheet1",
                &crate::sparkline::SparklineConfig::new("Sheet1!A1:A3", "B1"),
            )
            .unwrap();
        opened
            .add_sparkline(
                "Sheet1",
                &crate::sparkline::SparklineConfig::new("Sheet1!C1:C3", "D1"),
            )
            .unwrap();

        let saved = opened.save_to_buffer().unwrap();
        let reopened = Workbook::open_from_buffer_with_options(
            &saved,
            &OpenOptions::new()
                .read_mode(ReadMode::Eager)
                .aux_parts(AuxParts::EagerLoad),
        )
        .unwrap();
        assert_eq!(reopened.get_sparklines("Sheet1").unwrap().len(), 2);
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
