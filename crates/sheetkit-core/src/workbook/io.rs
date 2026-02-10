use super::*;

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
            sheet_sparklines: vec![vec![]],
            sheet_vml: vec![None],
            sheet_name_index,
        }
    }

    /// Open an existing `.xlsx` file from disk.
    ///
    /// If the file is encrypted (CFB container), returns
    /// [`Error::FileEncrypted`]. Use [`Workbook::open_with_password`] instead.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
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
        Self::from_archive(&mut archive)
    }

    /// Build a Workbook from an already-opened ZIP archive.
    fn from_archive<R: std::io::Read + std::io::Seek>(
        archive: &mut zip::ZipArchive<R>,
    ) -> Result<Self> {
        // Parse [Content_Types].xml
        let content_types: ContentTypes = read_xml_part(archive, "[Content_Types].xml")?;

        // Infer the workbook format from the content type of xl/workbook.xml.
        let format = content_types
            .overrides
            .iter()
            .find(|o| o.part_name == "/xl/workbook.xml")
            .and_then(|o| WorkbookFormat::from_content_type(&o.content_type))
            .unwrap_or_default();

        // Parse _rels/.rels
        let package_rels: Relationships = read_xml_part(archive, "_rels/.rels")?;

        // Parse xl/workbook.xml
        let workbook_xml: WorkbookXml = read_xml_part(archive, "xl/workbook.xml")?;

        // Parse xl/_rels/workbook.xml.rels
        let workbook_rels: Relationships = read_xml_part(archive, "xl/_rels/workbook.xml.rels")?;

        // Parse each worksheet referenced in the workbook.
        let sheet_count = workbook_xml.sheets.sheets.len();
        let mut worksheets = Vec::with_capacity(sheet_count);
        let mut worksheet_paths = Vec::with_capacity(sheet_count);
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
            let ws: WorksheetXml = read_xml_part(archive, &sheet_path)?;
            worksheets.push((sheet_entry.name.clone(), ws));
            worksheet_paths.push(sheet_path);
        }

        // Parse xl/styles.xml
        let stylesheet: StyleSheet = read_xml_part(archive, "xl/styles.xml")?;

        // Parse xl/sharedStrings.xml (optional -- may not exist for workbooks with no strings)
        let shared_strings: Sst =
            read_xml_part(archive, "xl/sharedStrings.xml").unwrap_or_default();

        let sst_runtime = SharedStringTable::from_sst(shared_strings);

        // Parse xl/theme/theme1.xml (optional -- preserved as raw bytes for round-trip).
        let (theme_xml, theme_colors) = match read_bytes_part(archive, "xl/theme/theme1.xml") {
            Ok(bytes) => {
                let colors = sheetkit_xml::theme::parse_theme_colors(&bytes);
                (Some(bytes), colors)
            }
            Err(_) => (None, crate::theme::default_theme_colors()),
        };

        // Parse per-sheet worksheet relationship files (optional).
        let mut worksheet_rels: HashMap<usize, Relationships> = HashMap::with_capacity(sheet_count);
        for (i, sheet_path) in worksheet_paths.iter().enumerate() {
            let rels_path = relationship_part_path(sheet_path);
            if let Ok(rels) = read_xml_part::<Relationships, _>(archive, &rels_path) {
                worksheet_rels.insert(i, rels);
            }
        }

        // Parse comments, VML drawings, drawings, drawing rels, charts, and images.
        let mut sheet_comments: Vec<Option<Comments>> = vec![None; worksheets.len()];
        let mut sheet_vml: Vec<Option<Vec<u8>>> = vec![None; worksheets.len()];
        let mut drawings: Vec<(String, WsDr)> = Vec::new();
        let mut worksheet_drawings: HashMap<usize, usize> = HashMap::new();
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
                drawing_path_to_idx.insert(drawing_path, idx);
            }
        }

        let mut drawing_rels: HashMap<usize, Relationships> = HashMap::new();
        let mut charts: Vec<(String, ChartSpace)> = Vec::new();
        let mut raw_charts: Vec<(String, Vec<u8>)> = Vec::new();
        let mut images: Vec<(String, Vec<u8>)> = Vec::new();
        let mut seen_chart_paths: HashSet<String> = HashSet::new();
        let mut seen_image_paths: HashSet<String> = HashSet::new();

        for (drawing_idx, (drawing_path, _)) in drawings.iter().enumerate() {
            let drawing_rels_path = relationship_part_path(drawing_path);
            let Ok(rels) = read_xml_part::<Relationships, _>(archive, &drawing_rels_path) else {
                continue;
            };

            for rel in &rels.relationships {
                if rel.rel_type == rel_types::CHART {
                    let chart_path = resolve_relationship_target(drawing_path, &rel.target);
                    if seen_chart_paths.insert(chart_path.clone()) {
                        match read_xml_part::<ChartSpace, _>(archive, &chart_path) {
                            Ok(chart) => charts.push((chart_path, chart)),
                            Err(_) => {
                                if let Ok(bytes) = read_bytes_part(archive, &chart_path) {
                                    raw_charts.push((chart_path, bytes));
                                }
                            }
                        }
                    }
                } else if rel.rel_type == rel_types::IMAGE {
                    let image_path = resolve_relationship_target(drawing_path, &rel.target);
                    if seen_image_paths.insert(image_path.clone()) {
                        if let Ok(bytes) = read_bytes_part(archive, &image_path) {
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
                    Ok(chart) => charts.push((chart_path, chart)),
                    Err(_) => {
                        if let Ok(bytes) = read_bytes_part(archive, &chart_path) {
                            raw_charts.push((chart_path, bytes));
                        }
                    }
                }
            }
        }

        // Parse docProps/core.xml (optional - uses manual XML parsing)
        let core_properties = read_string_part(archive, "docProps/core.xml")
            .ok()
            .and_then(|xml_str| {
                sheetkit_xml::doc_props::deserialize_core_properties(&xml_str).ok()
            });

        // Parse docProps/app.xml (optional - uses serde)
        let app_properties: Option<sheetkit_xml::doc_props::ExtendedProperties> =
            read_xml_part(archive, "docProps/app.xml").ok();

        // Parse docProps/custom.xml (optional - uses manual XML parsing)
        let custom_properties = read_string_part(archive, "docProps/custom.xml")
            .ok()
            .and_then(|xml_str| {
                sheetkit_xml::doc_props::deserialize_custom_properties(&xml_str).ok()
            });

        // Parse pivot cache definitions, pivot tables, and pivot cache records.
        let mut pivot_cache_defs = Vec::new();
        let mut pivot_tables = Vec::new();
        let mut pivot_cache_records = Vec::new();
        for ovr in &content_types.overrides {
            let path = ovr.part_name.trim_start_matches('/');
            if ovr.content_type == mime_types::PIVOT_CACHE_DEFINITION {
                if let Ok(pcd) = read_xml_part::<sheetkit_xml::pivot_cache::PivotCacheDefinition, _>(
                    archive, path,
                ) {
                    pivot_cache_defs.push((path.to_string(), pcd));
                }
            } else if ovr.content_type == mime_types::PIVOT_TABLE {
                if let Ok(pt) = read_xml_part::<sheetkit_xml::pivot_table::PivotTableDefinition, _>(
                    archive, path,
                ) {
                    pivot_tables.push((path.to_string(), pt));
                }
            } else if ovr.content_type == mime_types::PIVOT_CACHE_RECORDS {
                if let Ok(pcr) =
                    read_xml_part::<sheetkit_xml::pivot_cache::PivotCacheRecords, _>(archive, path)
                {
                    pivot_cache_records.push((path.to_string(), pcr));
                }
            }
        }

        // Parse sparklines from worksheet extension lists.
        let mut sheet_sparklines: Vec<Vec<crate::sparkline::SparklineConfig>> =
            vec![vec![]; worksheets.len()];
        for (i, ws_path) in worksheet_paths.iter().enumerate() {
            if let Ok(raw) = read_string_part(archive, ws_path) {
                let parsed = parse_sparklines_from_xml(&raw);
                if !parsed.is_empty() {
                    sheet_sparklines[i] = parsed;
                }
            }
        }

        // Build sheet name -> index lookup.
        let mut sheet_name_index = HashMap::with_capacity(worksheets.len());
        for (i, (name, _)) in worksheets.iter().enumerate() {
            sheet_name_index.insert(name.clone(), i);
        }

        // Populate cached column numbers on all cells and ensure sorted order
        // for binary search correctness.
        for (_name, ws) in &mut worksheets {
            // Ensure rows are sorted by row number (some writers output unsorted data).
            ws.sheet_data.rows.sort_unstable_by_key(|r| r.r);
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
            sheet_sparklines,
            sheet_vml,
            sheet_name_index,
        })
    }

    /// Save the workbook to a `.xlsx` file at the given path.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let file = std::fs::File::create(path)?;
        let mut zip = zip::ZipWriter::new(file);
        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated)
            .compression_level(Some(1));
        self.write_zip_contents(&mut zip, options)?;
        zip.finish().map_err(|e| Error::Zip(e.to_string()))?;
        Ok(())
    }

    /// Serialize the workbook to an in-memory `.xlsx` buffer.
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
            self.write_zip_contents(&mut zip, options)?;
            zip.finish().map_err(|e| Error::Zip(e.to_string()))?;
        }
        Ok(buf)
    }

    /// Open a workbook from an in-memory `.xlsx` buffer.
    pub fn open_from_buffer(data: &[u8]) -> Result<Self> {
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
        Self::from_archive(&mut archive)
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
        Self::from_archive(&mut archive)
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
            self.write_zip_contents(&mut zip, options)?;
            zip.finish().map_err(|e| Error::Zip(e.to_string()))?;
        }

        // Encrypt and write to CFB container
        let cfb_data = crate::crypt::encrypt_xlsx(&zip_buf, password)?;
        std::fs::write(path.as_ref(), &cfb_data)?;
        Ok(())
    }

    /// Write all workbook parts into the given ZIP writer.
    fn write_zip_contents<W: std::io::Write + std::io::Seek>(
        &self,
        zip: &mut zip::ZipWriter<W>,
        options: SimpleFileOptions,
    ) -> Result<()> {
        let mut content_types = self.content_types.clone();

        // Ensure the workbook override content type matches the stored format.
        if let Some(wb_override) = content_types
            .overrides
            .iter_mut()
            .find(|o| o.part_name == "/xl/workbook.xml")
        {
            wb_override.content_type = self.format.content_type().to_string();
        }

        let mut worksheet_rels = self.worksheet_rels.clone();

        // Synchronize comment and VML parts with worksheet relationships/content types.
        let comment_count = self.sheet_comments.iter().filter(|c| c.is_some()).count();
        // Per-sheet VML bytes to write: (sheet_idx, zip_path, bytes).
        let mut vml_parts_to_write: Vec<(usize, String, Vec<u8>)> =
            Vec::with_capacity(comment_count);
        // Per-sheet legacy drawing relationship IDs for worksheet XML serialization.
        let mut legacy_drawing_rids: HashMap<usize, String> = HashMap::with_capacity(comment_count);

        // Ensure the vml extension default content type is present if any VML exists.
        let mut has_any_vml = false;

        for sheet_idx in 0..self.worksheets.len() {
            let has_comments = self
                .sheet_comments
                .get(sheet_idx)
                .and_then(|c| c.as_ref())
                .is_some();
            if let Some(rels) = worksheet_rels.get_mut(&sheet_idx) {
                rels.relationships
                    .retain(|r| r.rel_type != rel_types::COMMENTS);
                rels.relationships
                    .retain(|r| r.rel_type != rel_types::VML_DRAWING);
            }
            if !has_comments {
                continue;
            }

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

            // Determine VML bytes: use preserved bytes if available, otherwise generate.
            let vml_path = format!("xl/drawings/vmlDrawing{}.vml", sheet_idx + 1);
            let vml_bytes =
                if let Some(bytes) = self.sheet_vml.get(sheet_idx).and_then(|v| v.as_ref()) {
                    bytes.clone()
                } else {
                    // Generate VML from comment cell references.
                    let comments = self.sheet_comments[sheet_idx].as_ref().unwrap();
                    let cells: Vec<&str> = comments
                        .comment_list
                        .comments
                        .iter()
                        .map(|c| c.r#ref.as_str())
                        .collect();
                    crate::vml::build_vml_drawing(&cells).into_bytes()
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

        // [Content_Types].xml
        write_xml_part(zip, "[Content_Types].xml", &content_types, options)?;

        // _rels/.rels
        write_xml_part(zip, "_rels/.rels", &self.package_rels, options)?;

        // xl/workbook.xml
        write_xml_part(zip, "xl/workbook.xml", &self.workbook_xml, options)?;

        // xl/_rels/workbook.xml.rels
        write_xml_part(
            zip,
            "xl/_rels/workbook.xml.rels",
            &self.workbook_rels,
            options,
        )?;

        // xl/worksheets/sheet{N}.xml
        for (i, (_name, ws)) in self.worksheets.iter().enumerate() {
            let entry_name = self.sheet_part_path(i);
            let empty_sparklines: Vec<crate::sparkline::SparklineConfig> = vec![];
            let sparklines = self.sheet_sparklines.get(i).unwrap_or(&empty_sparklines);
            let legacy_rid = legacy_drawing_rids.get(&i).map(|s| s.as_str());

            if legacy_rid.is_none() && sparklines.is_empty() {
                write_xml_part(zip, &entry_name, ws, options)?;
            } else {
                // Serialize worksheet to XML and inject extras via string manipulation,
                // avoiding a full WorksheetXml clone.
                let xml = serialize_worksheet_with_extras(ws, sparklines, legacy_rid)?;
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

        // xl/theme/theme1.xml
        {
            let default_theme = crate::theme::default_theme_xml();
            let theme_bytes = self.theme_xml.as_deref().unwrap_or(&default_theme);
            zip.start_file("xl/theme/theme1.xml", options)
                .map_err(|e| Error::Zip(e.to_string()))?;
            zip.write_all(theme_bytes)?;
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
    fn test_save_writes_correct_content_type_for_format() {
        let dir = TempDir::new().unwrap();

        for fmt in [
            WorkbookFormat::Xlsx,
            WorkbookFormat::Xlsm,
            WorkbookFormat::Xltx,
            WorkbookFormat::Xltm,
            WorkbookFormat::Xlam,
        ] {
            let path = dir.path().join(format!("test_{:?}.xlsx", fmt));
            let mut wb = Workbook::new();
            wb.set_format(fmt);
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
                fmt.content_type(),
                "content type mismatch for {:?}",
                fmt
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
}
