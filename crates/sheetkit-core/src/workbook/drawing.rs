use super::*;

impl Workbook {
    /// Add a chart to a sheet, anchored between two cells.
    ///
    /// The chart spans from `from_cell` (e.g., `"B2"`) to `to_cell`
    /// (e.g., `"J15"`). The `config` specifies the chart type, series data,
    /// title, and legend visibility.
    pub fn add_chart(
        &mut self,
        sheet: &str,
        from_cell: &str,
        to_cell: &str,
        config: &ChartConfig,
    ) -> Result<()> {
        let sheet_idx =
            crate::sheet::find_sheet_index(&self.worksheets, sheet).ok_or_else(|| {
                Error::SheetNotFound {
                    name: sheet.to_string(),
                }
            })?;

        // Parse cell references to marker coordinates (0-based).
        let (from_col, from_row) = cell_name_to_coordinates(from_cell)?;
        let (to_col, to_row) = cell_name_to_coordinates(to_cell)?;

        let from_marker = MarkerType {
            col: from_col - 1,
            col_off: 0,
            row: from_row - 1,
            row_off: 0,
        };
        let to_marker = MarkerType {
            col: to_col - 1,
            col_off: 0,
            row: to_row - 1,
            row_off: 0,
        };

        // Allocate chart part.
        let chart_num = self.charts.len() + 1;
        let chart_path = format!("xl/charts/chart{}.xml", chart_num);
        let chart_space = crate::chart::build_chart_xml(config);
        self.charts.push((chart_path, chart_space));

        // Get or create drawing for this sheet.
        let drawing_idx = self.ensure_drawing_for_sheet(sheet_idx);

        // Add chart reference to the drawing's relationships.
        let chart_rid = self.next_drawing_rid(drawing_idx);
        let chart_rel_target = format!("../charts/chart{}.xml", chart_num);

        let dr_rels = self
            .drawing_rels
            .entry(drawing_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        dr_rels.relationships.push(Relationship {
            id: chart_rid.clone(),
            rel_type: rel_types::CHART.to_string(),
            target: chart_rel_target,
            target_mode: None,
        });

        // Build the chart anchor and add it to the drawing.
        let drawing = &mut self.drawings[drawing_idx].1;
        let anchor = crate::chart::build_drawing_with_chart(&chart_rid, from_marker, to_marker);
        drawing.two_cell_anchors.extend(anchor.two_cell_anchors);

        // Add content type for the chart.
        self.content_types.overrides.push(ContentTypeOverride {
            part_name: format!("/xl/charts/chart{}.xml", chart_num),
            content_type: mime_types::CHART.to_string(),
        });

        Ok(())
    }

    /// Add a shape to a sheet, anchored between two cells.
    ///
    /// The shape spans from `config.from_cell` to `config.to_cell`. Unlike
    /// charts and images, shapes do not reference external parts and therefore
    /// do not need a relationship entry.
    pub fn add_shape(&mut self, sheet: &str, config: &crate::shape::ShapeConfig) -> Result<()> {
        let sheet_idx =
            crate::sheet::find_sheet_index(&self.worksheets, sheet).ok_or_else(|| {
                Error::SheetNotFound {
                    name: sheet.to_string(),
                }
            })?;

        let drawing_idx = self.ensure_drawing_for_sheet(sheet_idx);

        let drawing = &mut self.drawings[drawing_idx].1;
        let shape_id = (drawing.one_cell_anchors.len() + drawing.two_cell_anchors.len() + 2) as u32;

        let anchor = crate::shape::build_shape_anchor(config, shape_id)?;
        drawing.two_cell_anchors.push(anchor);

        Ok(())
    }

    /// Add an image to a sheet from bytes.
    ///
    /// The image is anchored to the cell specified in `config.from_cell`.
    /// Dimensions are specified in pixels via `config.width_px` and
    /// `config.height_px`.
    pub fn add_image(&mut self, sheet: &str, config: &ImageConfig) -> Result<()> {
        crate::image::validate_image_config(config)?;

        let sheet_idx =
            crate::sheet::find_sheet_index(&self.worksheets, sheet).ok_or_else(|| {
                Error::SheetNotFound {
                    name: sheet.to_string(),
                }
            })?;

        // Allocate image media part.
        let image_num = self.images.len() + 1;
        let image_path = format!("xl/media/image{}.{}", image_num, config.format.extension());
        self.images.push((image_path, config.data.clone()));

        // Ensure the image extension has a default content type.
        let ext = config.format.extension().to_string();
        if !self
            .content_types
            .defaults
            .iter()
            .any(|d| d.extension == ext)
        {
            self.content_types.defaults.push(ContentTypeDefault {
                extension: ext,
                content_type: config.format.content_type().to_string(),
            });
        }

        // Get or create drawing for this sheet.
        let drawing_idx = self.ensure_drawing_for_sheet(sheet_idx);

        // Add image reference to the drawing's relationships.
        let image_rid = self.next_drawing_rid(drawing_idx);
        let image_rel_target = format!("../media/image{}.{}", image_num, config.format.extension());

        let dr_rels = self
            .drawing_rels
            .entry(drawing_idx)
            .or_insert_with(|| Relationships {
                xmlns: sheetkit_xml::namespaces::PACKAGE_RELATIONSHIPS.to_string(),
                relationships: vec![],
            });
        dr_rels.relationships.push(Relationship {
            id: image_rid.clone(),
            rel_type: rel_types::IMAGE.to_string(),
            target: image_rel_target,
            target_mode: None,
        });

        // Count existing objects in the drawing to assign a unique ID.
        let drawing = &mut self.drawings[drawing_idx].1;
        let pic_id = (drawing.one_cell_anchors.len() + drawing.two_cell_anchors.len() + 2) as u32;

        // Add image anchor to the drawing.
        crate::image::add_image_to_drawing(drawing, &image_rid, config, pic_id)?;

        Ok(())
    }

    /// Delete a chart anchored at the given cell.
    ///
    /// Removes the drawing anchor, chart data, relationship entry, and content
    /// type override for the chart at `cell` on `sheet`.
    pub fn delete_chart(&mut self, sheet: &str, cell: &str) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let (col, row) = cell_name_to_coordinates(cell)?;
        let target_col = col - 1;
        let target_row = row - 1;

        let &drawing_idx =
            self.worksheet_drawings
                .get(&sheet_idx)
                .ok_or_else(|| Error::ChartNotFound {
                    sheet: sheet.to_string(),
                    cell: cell.to_string(),
                })?;

        let drawing = &self.drawings[drawing_idx].1;
        let anchor_pos = drawing
            .two_cell_anchors
            .iter()
            .position(|a| {
                a.from.col == target_col && a.from.row == target_row && a.graphic_frame.is_some()
            })
            .ok_or_else(|| Error::ChartNotFound {
                sheet: sheet.to_string(),
                cell: cell.to_string(),
            })?;

        let anchor = &drawing.two_cell_anchors[anchor_pos];
        let chart_rid = anchor
            .graphic_frame
            .as_ref()
            .unwrap()
            .graphic
            .graphic_data
            .chart
            .r_id
            .clone();

        let chart_path = self
            .drawing_rels
            .get(&drawing_idx)
            .and_then(|rels| {
                rels.relationships
                    .iter()
                    .find(|r| r.id == chart_rid)
                    .map(|r| {
                        let drawing_path = &self.drawings[drawing_idx].0;
                        let base_dir = drawing_path
                            .rfind('/')
                            .map(|i| &drawing_path[..i])
                            .unwrap_or("");
                        if r.target.starts_with("../") {
                            let rel_target = r.target.trim_start_matches("../");
                            let parent = base_dir.rfind('/').map(|i| &base_dir[..i]).unwrap_or("");
                            if parent.is_empty() {
                                rel_target.to_string()
                            } else {
                                format!("{}/{}", parent, rel_target)
                            }
                        } else {
                            format!("{}/{}", base_dir, r.target)
                        }
                    })
            })
            .ok_or_else(|| Error::ChartNotFound {
                sheet: sheet.to_string(),
                cell: cell.to_string(),
            })?;

        self.charts.retain(|(path, _)| path != &chart_path);
        self.raw_charts.retain(|(path, _)| path != &chart_path);

        if let Some(rels) = self.drawing_rels.get_mut(&drawing_idx) {
            rels.relationships.retain(|r| r.id != chart_rid);
        }

        self.drawings[drawing_idx]
            .1
            .two_cell_anchors
            .remove(anchor_pos);

        let ct_part_name = format!("/{}", chart_path);
        self.content_types
            .overrides
            .retain(|o| o.part_name != ct_part_name);

        Ok(())
    }

    /// Delete a picture anchored at the given cell.
    ///
    /// Removes the drawing anchor, image data, relationship entry, and content
    /// type for the picture at `cell` on `sheet`. Searches both one-cell and
    /// two-cell anchors.
    pub fn delete_picture(&mut self, sheet: &str, cell: &str) -> Result<()> {
        let sheet_idx = self.sheet_index(sheet)?;
        let (col, row) = cell_name_to_coordinates(cell)?;
        let target_col = col - 1;
        let target_row = row - 1;

        let &drawing_idx =
            self.worksheet_drawings
                .get(&sheet_idx)
                .ok_or_else(|| Error::PictureNotFound {
                    sheet: sheet.to_string(),
                    cell: cell.to_string(),
                })?;

        let drawing = &self.drawings[drawing_idx].1;

        // Check one-cell anchors first
        if let Some(pos) = drawing
            .one_cell_anchors
            .iter()
            .position(|a| a.from.col == target_col && a.from.row == target_row && a.pic.is_some())
        {
            let image_rid = drawing.one_cell_anchors[pos]
                .pic
                .as_ref()
                .unwrap()
                .blip_fill
                .blip
                .r_embed
                .clone();

            self.remove_picture_data(drawing_idx, &image_rid);
            self.drawings[drawing_idx].1.one_cell_anchors.remove(pos);
            return Ok(());
        }

        // Check two-cell anchors
        if let Some(pos) = drawing
            .two_cell_anchors
            .iter()
            .position(|a| a.from.col == target_col && a.from.row == target_row && a.pic.is_some())
        {
            let image_rid = drawing.two_cell_anchors[pos]
                .pic
                .as_ref()
                .unwrap()
                .blip_fill
                .blip
                .r_embed
                .clone();

            self.remove_picture_data(drawing_idx, &image_rid);
            self.drawings[drawing_idx].1.two_cell_anchors.remove(pos);
            return Ok(());
        }

        Err(Error::PictureNotFound {
            sheet: sheet.to_string(),
            cell: cell.to_string(),
        })
    }

    /// Remove image data, relationship, and content type for a picture reference.
    fn remove_picture_data(&mut self, drawing_idx: usize, image_rid: &str) {
        if let Some(image_path) = self.resolve_drawing_rel_target(drawing_idx, image_rid) {
            self.images.retain(|(path, _)| path != &image_path);
        }

        if let Some(rels) = self.drawing_rels.get_mut(&drawing_idx) {
            rels.relationships.retain(|r| r.id != image_rid);
        }
    }

    /// Resolve a relationship target to a full zip path.
    fn resolve_drawing_rel_target(&self, drawing_idx: usize, rid: &str) -> Option<String> {
        self.drawing_rels.get(&drawing_idx).and_then(|rels| {
            rels.relationships.iter().find(|r| r.id == rid).map(|r| {
                let drawing_path = &self.drawings[drawing_idx].0;
                let base_dir = drawing_path
                    .rfind('/')
                    .map(|i| &drawing_path[..i])
                    .unwrap_or("");
                if r.target.starts_with("../") {
                    let rel_target = r.target.trim_start_matches("../");
                    let parent = base_dir.rfind('/').map(|i| &base_dir[..i]).unwrap_or("");
                    if parent.is_empty() {
                        rel_target.to_string()
                    } else {
                        format!("{}/{}", parent, rel_target)
                    }
                } else {
                    format!("{}/{}", base_dir, r.target)
                }
            })
        })
    }

    /// Get all pictures anchored at the given cell.
    ///
    /// Returns picture data, format, anchor cell, and dimensions for each
    /// picture found at the specified cell.
    pub fn get_pictures(&self, sheet: &str, cell: &str) -> Result<Vec<crate::image::PictureInfo>> {
        let sheet_idx = self.sheet_index(sheet)?;
        let (col, row) = cell_name_to_coordinates(cell)?;
        let target_col = col - 1;
        let target_row = row - 1;

        let drawing_idx = match self.worksheet_drawings.get(&sheet_idx) {
            Some(&idx) => idx,
            None => return Ok(vec![]),
        };

        let drawing = &self.drawings[drawing_idx].1;
        let mut results = Vec::new();

        for anchor in &drawing.one_cell_anchors {
            if anchor.from.col == target_col && anchor.from.row == target_row {
                if let Some(pic) = &anchor.pic {
                    if let Some(info) = self.extract_picture_info(drawing_idx, pic, cell) {
                        results.push(info);
                    }
                }
            }
        }

        for anchor in &drawing.two_cell_anchors {
            if anchor.from.col == target_col && anchor.from.row == target_row {
                if let Some(pic) = &anchor.pic {
                    if let Some(info) = self.extract_picture_info(drawing_idx, pic, cell) {
                        results.push(info);
                    }
                }
            }
        }

        Ok(results)
    }

    /// Extract picture info from a Picture element by resolving its relationship.
    fn extract_picture_info(
        &self,
        drawing_idx: usize,
        pic: &sheetkit_xml::drawing::Picture,
        cell: &str,
    ) -> Option<crate::image::PictureInfo> {
        let rid = &pic.blip_fill.blip.r_embed;
        let image_path = self.resolve_drawing_rel_target(drawing_idx, rid)?;
        let (data, format) = self.find_image_with_format(&image_path)?;

        let cx = pic.sp_pr.xfrm.ext.cx;
        let cy = pic.sp_pr.xfrm.ext.cy;
        let width_px = (cx / crate::image::EMU_PER_PIXEL) as u32;
        let height_px = (cy / crate::image::EMU_PER_PIXEL) as u32;

        Some(crate::image::PictureInfo {
            data: data.clone(),
            format,
            cell: cell.to_string(),
            width_px,
            height_px,
        })
    }

    /// Find image data and determine format from the zip path extension.
    fn find_image_with_format(
        &self,
        image_path: &str,
    ) -> Option<(&Vec<u8>, crate::image::ImageFormat)> {
        self.images
            .iter()
            .find(|(p, _)| p == image_path)
            .and_then(|(path, data)| {
                let ext = path.rsplit('.').next()?;
                let format = crate::image::ImageFormat::from_extension(ext).ok()?;
                Some((data, format))
            })
    }

    /// Get all cells that have pictures anchored to them on the given sheet.
    pub fn get_picture_cells(&self, sheet: &str) -> Result<Vec<String>> {
        let sheet_idx = self.sheet_index(sheet)?;

        let drawing_idx = match self.worksheet_drawings.get(&sheet_idx) {
            Some(&idx) => idx,
            None => return Ok(vec![]),
        };

        let drawing = &self.drawings[drawing_idx].1;
        let mut cells = Vec::new();

        for anchor in &drawing.one_cell_anchors {
            if anchor.pic.is_some() {
                if let Ok(name) = crate::utils::cell_ref::coordinates_to_cell_name(
                    anchor.from.col + 1,
                    anchor.from.row + 1,
                ) {
                    cells.push(name);
                }
            }
        }

        for anchor in &drawing.two_cell_anchors {
            if anchor.pic.is_some() {
                if let Ok(name) = crate::utils::cell_ref::coordinates_to_cell_name(
                    anchor.from.col + 1,
                    anchor.from.row + 1,
                ) {
                    cells.push(name);
                }
            }
        }

        Ok(cells)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn test_add_chart_basic() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: Some("Test Chart".to_string()),
            series: vec![ChartSeries {
                name: "Sales".to_string(),
                categories: "Sheet1!$A$1:$A$5".to_string(),
                values: "Sheet1!$B$1:$B$5".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "E1", "L15", &config).unwrap();

        assert_eq!(wb.charts.len(), 1);
        assert_eq!(wb.drawings.len(), 1);
        assert!(wb.worksheet_drawings.contains_key(&0));
        assert!(wb.drawing_rels.contains_key(&0));
        assert!(wb.worksheets[0].1.drawing.is_some());
    }

    #[test]
    fn test_add_chart_sheet_not_found() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Line,
            title: None,
            series: vec![ChartSeries {
                name: String::new(),
                categories: "Sheet1!$A$1:$A$5".to_string(),
                values: "Sheet1!$B$1:$B$5".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        };
        let result = wb.add_chart("NoSheet", "A1", "H10", &config);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_add_multiple_charts_same_sheet() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        let config1 = ChartConfig {
            chart_type: ChartType::Col,
            title: Some("Chart 1".to_string()),
            series: vec![ChartSeries {
                name: "S1".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
        };
        let config2 = ChartConfig {
            chart_type: ChartType::Line,
            title: Some("Chart 2".to_string()),
            series: vec![ChartSeries {
                name: "S2".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$C$1:$C$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "A1", "F10", &config1).unwrap();
        wb.add_chart("Sheet1", "A12", "F22", &config2).unwrap();

        assert_eq!(wb.charts.len(), 2);
        assert_eq!(wb.drawings.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 2);
    }

    #[test]
    fn test_add_charts_different_sheets() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        wb.new_sheet("Sheet2").unwrap();

        let config = ChartConfig {
            chart_type: ChartType::Pie,
            title: None,
            series: vec![ChartSeries {
                name: String::new(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "A1", "F10", &config).unwrap();
        wb.add_chart("Sheet2", "A1", "F10", &config).unwrap();

        assert_eq!(wb.charts.len(), 2);
        assert_eq!(wb.drawings.len(), 2);
    }

    #[test]
    fn test_save_with_chart() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("with_chart.xlsx");

        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Bar,
            title: Some("Bar Chart".to_string()),
            series: vec![ChartSeries {
                name: "Data".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "E2", "L15", &config).unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        assert!(archive.by_name("xl/charts/chart1.xml").is_ok());
        assert!(archive.by_name("xl/drawings/drawing1.xml").is_ok());
        assert!(archive
            .by_name("xl/worksheets/_rels/sheet1.xml.rels")
            .is_ok());
        assert!(archive
            .by_name("xl/drawings/_rels/drawing1.xml.rels")
            .is_ok());
    }

    #[test]
    fn test_add_image_basic() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x89, 0x50, 0x4E, 0x47],
            format: ImageFormat::Png,
            from_cell: "B2".to_string(),
            width_px: 400,
            height_px: 300,
        };
        wb.add_image("Sheet1", &config).unwrap();

        assert_eq!(wb.images.len(), 1);
        assert_eq!(wb.drawings.len(), 1);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 1);
        assert!(wb.worksheet_drawings.contains_key(&0));
    }

    #[test]
    fn test_add_image_sheet_not_found() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x89],
            format: ImageFormat::Png,
            from_cell: "A1".to_string(),
            width_px: 100,
            height_px: 100,
        };
        let result = wb.add_image("NoSheet", &config);
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_add_image_invalid_config() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![],
            format: ImageFormat::Png,
            from_cell: "A1".to_string(),
            width_px: 100,
            height_px: 100,
        };
        assert!(wb.add_image("Sheet1", &config).is_err());

        let config = ImageConfig {
            data: vec![1],
            format: ImageFormat::Jpeg,
            from_cell: "A1".to_string(),
            width_px: 0,
            height_px: 100,
        };
        assert!(wb.add_image("Sheet1", &config).is_err());
    }

    #[test]
    fn test_save_with_image() {
        use crate::image::{ImageConfig, ImageFormat};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("with_image.xlsx");

        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A],
            format: ImageFormat::Png,
            from_cell: "C3".to_string(),
            width_px: 200,
            height_px: 150,
        };
        wb.add_image("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        assert!(archive.by_name("xl/media/image1.png").is_ok());
        assert!(archive.by_name("xl/drawings/drawing1.xml").is_ok());
        assert!(archive
            .by_name("xl/worksheets/_rels/sheet1.xml.rels")
            .is_ok());
        assert!(archive
            .by_name("xl/drawings/_rels/drawing1.xml.rels")
            .is_ok());
    }

    #[test]
    fn test_save_with_jpeg_image() {
        use crate::image::{ImageConfig, ImageFormat};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("with_jpeg.xlsx");

        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0xFF, 0xD8, 0xFF, 0xE0],
            format: ImageFormat::Jpeg,
            from_cell: "A1".to_string(),
            width_px: 640,
            height_px: 480,
        };
        wb.add_image("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/media/image1.jpeg").is_ok());
    }

    #[test]
    fn test_save_with_new_image_formats() {
        use crate::image::{ImageConfig, ImageFormat};

        let formats = [
            (ImageFormat::Bmp, "bmp"),
            (ImageFormat::Ico, "ico"),
            (ImageFormat::Tiff, "tiff"),
            (ImageFormat::Svg, "svg"),
            (ImageFormat::Emf, "emf"),
            (ImageFormat::Emz, "emz"),
            (ImageFormat::Wmf, "wmf"),
            (ImageFormat::Wmz, "wmz"),
        ];

        for (format, ext) in &formats {
            let dir = TempDir::new().unwrap();
            let path = dir.path().join(format!("with_{ext}.xlsx"));

            let mut wb = Workbook::new();
            let config = ImageConfig {
                data: vec![0x00, 0x01, 0x02, 0x03],
                format: format.clone(),
                from_cell: "A1".to_string(),
                width_px: 100,
                height_px: 100,
            };
            wb.add_image("Sheet1", &config).unwrap();
            wb.save(&path).unwrap();

            let file = std::fs::File::open(&path).unwrap();
            let mut archive = zip::ZipArchive::new(file).unwrap();
            let media_path = format!("xl/media/image1.{ext}");
            assert!(
                archive.by_name(&media_path).is_ok(),
                "expected {media_path} in archive for format {ext}"
            );
        }
    }

    #[test]
    fn test_add_image_new_format_content_type_default() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x42, 0x4D],
            format: ImageFormat::Bmp,
            from_cell: "A1".to_string(),
            width_px: 100,
            height_px: 100,
        };
        wb.add_image("Sheet1", &config).unwrap();

        let has_bmp_default = wb
            .content_types
            .defaults
            .iter()
            .any(|d| d.extension == "bmp" && d.content_type == "image/bmp");
        assert!(has_bmp_default, "content types should have bmp default");
    }

    #[test]
    fn test_add_image_svg_content_type_default() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x3C, 0x73, 0x76, 0x67],
            format: ImageFormat::Svg,
            from_cell: "B3".to_string(),
            width_px: 200,
            height_px: 200,
        };
        wb.add_image("Sheet1", &config).unwrap();

        let has_svg_default = wb
            .content_types
            .defaults
            .iter()
            .any(|d| d.extension == "svg" && d.content_type == "image/svg+xml");
        assert!(has_svg_default, "content types should have svg default");
    }

    #[test]
    fn test_add_image_emf_content_type_and_path() {
        use crate::image::{ImageConfig, ImageFormat};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("with_emf.xlsx");

        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x01, 0x00, 0x00, 0x00],
            format: ImageFormat::Emf,
            from_cell: "A1".to_string(),
            width_px: 150,
            height_px: 150,
        };
        wb.add_image("Sheet1", &config).unwrap();

        let has_emf_default = wb
            .content_types
            .defaults
            .iter()
            .any(|d| d.extension == "emf" && d.content_type == "image/x-emf");
        assert!(has_emf_default);

        wb.save(&path).unwrap();
        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/media/image1.emf").is_ok());
    }

    #[test]
    fn test_add_multiple_new_format_images() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();

        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x42, 0x4D],
                format: ImageFormat::Bmp,
                from_cell: "A1".to_string(),
                width_px: 100,
                height_px: 100,
            },
        )
        .unwrap();

        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x3C, 0x73],
                format: ImageFormat::Svg,
                from_cell: "C1".to_string(),
                width_px: 100,
                height_px: 100,
            },
        )
        .unwrap();

        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x01, 0x00],
                format: ImageFormat::Wmf,
                from_cell: "E1".to_string(),
                width_px: 100,
                height_px: 100,
            },
        )
        .unwrap();

        assert_eq!(wb.images.len(), 3);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 3);

        let ext_defaults: Vec<&str> = wb
            .content_types
            .defaults
            .iter()
            .map(|d| d.extension.as_str())
            .collect();
        assert!(ext_defaults.contains(&"bmp"));
        assert!(ext_defaults.contains(&"svg"));
        assert!(ext_defaults.contains(&"wmf"));
    }

    #[test]
    fn test_add_shape_basic() {
        use crate::shape::{ShapeConfig, ShapeType};
        let mut wb = Workbook::new();
        let config = ShapeConfig {
            shape_type: ShapeType::Rect,
            from_cell: "B2".to_string(),
            to_cell: "F10".to_string(),
            text: Some("Test Shape".to_string()),
            fill_color: Some("FF0000".to_string()),
            line_color: None,
            line_width: None,
        };
        wb.add_shape("Sheet1", &config).unwrap();

        assert_eq!(wb.drawings.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 1);
        let anchor = &wb.drawings[0].1.two_cell_anchors[0];
        assert!(anchor.shape.is_some());
        assert!(anchor.graphic_frame.is_none());
        assert!(anchor.pic.is_none());
    }

    #[test]
    fn test_add_multiple_shapes_same_sheet() {
        use crate::shape::{ShapeConfig, ShapeType};
        let mut wb = Workbook::new();
        wb.add_shape(
            "Sheet1",
            &ShapeConfig {
                shape_type: ShapeType::Rect,
                from_cell: "A1".to_string(),
                to_cell: "C3".to_string(),
                text: None,
                fill_color: None,
                line_color: None,
                line_width: None,
            },
        )
        .unwrap();
        wb.add_shape(
            "Sheet1",
            &ShapeConfig {
                shape_type: ShapeType::Ellipse,
                from_cell: "E1".to_string(),
                to_cell: "H5".to_string(),
                text: Some("Circle".to_string()),
                fill_color: Some("00FF00".to_string()),
                line_color: None,
                line_width: None,
            },
        )
        .unwrap();

        assert_eq!(wb.drawings.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 2);
    }

    #[test]
    fn test_add_shape_and_chart_same_sheet() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        use crate::shape::{ShapeConfig, ShapeType};

        let mut wb = Workbook::new();
        wb.add_chart(
            "Sheet1",
            "E1",
            "L10",
            &ChartConfig {
                chart_type: ChartType::Col,
                title: Some("Chart".to_string()),
                series: vec![ChartSeries {
                    name: "S1".to_string(),
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
        wb.add_shape(
            "Sheet1",
            &ShapeConfig {
                shape_type: ShapeType::Rect,
                from_cell: "A12".to_string(),
                to_cell: "D18".to_string(),
                text: Some("Label".to_string()),
                fill_color: None,
                line_color: None,
                line_width: None,
            },
        )
        .unwrap();

        assert_eq!(wb.drawings.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 2);
    }

    #[test]
    fn test_add_chart_and_image_same_sheet() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        use crate::image::{ImageConfig, ImageFormat};

        let mut wb = Workbook::new();

        let chart_config = ChartConfig {
            chart_type: ChartType::Col,
            title: Some("My Chart".to_string()),
            series: vec![ChartSeries {
                name: "Series 1".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "E1", "L10", &chart_config).unwrap();

        let image_config = ImageConfig {
            data: vec![0x89, 0x50, 0x4E, 0x47],
            format: ImageFormat::Png,
            from_cell: "E12".to_string(),
            width_px: 300,
            height_px: 200,
        };
        wb.add_image("Sheet1", &image_config).unwrap();

        assert_eq!(wb.drawings.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 1);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 1);
        assert_eq!(wb.charts.len(), 1);
        assert_eq!(wb.images.len(), 1);
    }

    #[test]
    fn test_save_with_chart_roundtrip_drawing_ref() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("chart_drawref.xlsx");

        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: None,
            series: vec![ChartSeries {
                name: "Series 1".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "A1", "F10", &config).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let ws = wb2.worksheet_ref("Sheet1").unwrap();
        assert!(ws.drawing.is_some());
    }

    #[test]
    fn test_open_save_preserves_existing_drawing_chart_and_image_parts() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        use crate::image::{ImageConfig, ImageFormat};
        let dir = TempDir::new().unwrap();
        let path1 = dir.path().join("source_with_parts.xlsx");
        let path2 = dir.path().join("resaved_with_parts.xlsx");

        let mut wb = Workbook::new();
        wb.add_chart(
            "Sheet1",
            "E1",
            "L10",
            &ChartConfig {
                chart_type: ChartType::Col,
                title: Some("Chart".to_string()),
                series: vec![ChartSeries {
                    name: "Series 1".to_string(),
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
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x89, 0x50, 0x4E, 0x47],
                format: ImageFormat::Png,
                from_cell: "E12".to_string(),
                width_px: 120,
                height_px: 80,
            },
        )
        .unwrap();
        wb.save(&path1).unwrap();

        let wb2 = Workbook::open(&path1).unwrap();
        assert_eq!(wb2.charts.len() + wb2.raw_charts.len(), 1);
        assert_eq!(wb2.drawings.len(), 1);
        assert_eq!(wb2.images.len(), 1);
        assert_eq!(wb2.drawing_rels.len(), 1);
        assert_eq!(wb2.worksheet_drawings.len(), 1);

        wb2.save(&path2).unwrap();

        let file = std::fs::File::open(&path2).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/charts/chart1.xml").is_ok());
        assert!(archive.by_name("xl/drawings/drawing1.xml").is_ok());
        assert!(archive.by_name("xl/media/image1.png").is_ok());
        assert!(archive
            .by_name("xl/worksheets/_rels/sheet1.xml.rels")
            .is_ok());
        assert!(archive
            .by_name("xl/drawings/_rels/drawing1.xml.rels")
            .is_ok());
    }

    #[test]
    fn test_delete_chart_basic() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: Some("Chart".to_string()),
            series: vec![ChartSeries {
                name: "S1".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "E1", "L10", &config).unwrap();
        assert_eq!(wb.charts.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 1);

        wb.delete_chart("Sheet1", "E1").unwrap();
        assert_eq!(wb.charts.len(), 0);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 0);
    }

    #[test]
    fn test_delete_chart_not_found() {
        let mut wb = Workbook::new();
        let result = wb.delete_chart("Sheet1", "A1");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::ChartNotFound { .. }));
    }

    #[test]
    fn test_delete_chart_wrong_cell() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: None,
            series: vec![ChartSeries {
                name: "S".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "E1", "L10", &config).unwrap();

        let result = wb.delete_chart("Sheet1", "A1");
        assert!(result.is_err());
        assert_eq!(wb.charts.len(), 1);
    }

    #[test]
    fn test_delete_chart_removes_content_type() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: None,
            series: vec![ChartSeries {
                name: "S".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "E1", "L10", &config).unwrap();
        let has_chart_ct = wb
            .content_types
            .overrides
            .iter()
            .any(|o| o.part_name.contains("chart"));
        assert!(has_chart_ct);

        wb.delete_chart("Sheet1", "E1").unwrap();
        let has_chart_ct = wb
            .content_types
            .overrides
            .iter()
            .any(|o| o.part_name.contains("chart"));
        assert!(!has_chart_ct);
    }

    #[test]
    fn test_delete_one_chart_keeps_others() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let mut wb = Workbook::new();
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: None,
            series: vec![ChartSeries {
                name: "S".to_string(),
                categories: "Sheet1!$A$1:$A$3".to_string(),
                values: "Sheet1!$B$1:$B$3".to_string(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        };
        wb.add_chart("Sheet1", "A1", "F10", &config).unwrap();
        wb.add_chart("Sheet1", "A12", "F22", &config).unwrap();
        assert_eq!(wb.charts.len(), 2);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 2);

        wb.delete_chart("Sheet1", "A1").unwrap();
        assert_eq!(wb.charts.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors[0].from.row, 11);
    }

    #[test]
    fn test_delete_picture_basic() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x89, 0x50, 0x4E, 0x47],
            format: ImageFormat::Png,
            from_cell: "B2".to_string(),
            width_px: 200,
            height_px: 150,
        };
        wb.add_image("Sheet1", &config).unwrap();
        assert_eq!(wb.images.len(), 1);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 1);

        wb.delete_picture("Sheet1", "B2").unwrap();
        assert_eq!(wb.images.len(), 0);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 0);
    }

    #[test]
    fn test_delete_picture_not_found() {
        let mut wb = Workbook::new();
        let result = wb.delete_picture("Sheet1", "A1");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::PictureNotFound { .. }));
    }

    #[test]
    fn test_delete_picture_wrong_cell() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let config = ImageConfig {
            data: vec![0x89, 0x50, 0x4E, 0x47],
            format: ImageFormat::Png,
            from_cell: "C3".to_string(),
            width_px: 100,
            height_px: 100,
        };
        wb.add_image("Sheet1", &config).unwrap();

        let result = wb.delete_picture("Sheet1", "A1");
        assert!(result.is_err());
        assert_eq!(wb.images.len(), 1);
    }

    #[test]
    fn test_delete_one_picture_keeps_others() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x89, 0x50, 0x4E, 0x47],
                format: ImageFormat::Png,
                from_cell: "A1".to_string(),
                width_px: 100,
                height_px: 100,
            },
        )
        .unwrap();
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0xFF, 0xD8, 0xFF, 0xE0],
                format: ImageFormat::Jpeg,
                from_cell: "C3".to_string(),
                width_px: 200,
                height_px: 200,
            },
        )
        .unwrap();
        assert_eq!(wb.images.len(), 2);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 2);

        wb.delete_picture("Sheet1", "A1").unwrap();
        assert_eq!(wb.images.len(), 1);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 1);
        assert_eq!(wb.drawings[0].1.one_cell_anchors[0].from.col, 2);
    }

    #[test]
    fn test_get_picture_cells_empty() {
        let wb = Workbook::new();
        let cells = wb.get_picture_cells("Sheet1").unwrap();
        assert!(cells.is_empty());
    }

    #[test]
    fn test_get_picture_cells_returns_cells() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x89, 0x50],
                format: ImageFormat::Png,
                from_cell: "B2".to_string(),
                width_px: 100,
                height_px: 100,
            },
        )
        .unwrap();
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0xFF, 0xD8],
                format: ImageFormat::Jpeg,
                from_cell: "D5".to_string(),
                width_px: 200,
                height_px: 150,
            },
        )
        .unwrap();

        let cells = wb.get_picture_cells("Sheet1").unwrap();
        assert_eq!(cells.len(), 2);
        assert!(cells.contains(&"B2".to_string()));
        assert!(cells.contains(&"D5".to_string()));
    }

    #[test]
    fn test_get_pictures_returns_data() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        let image_data = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: image_data.clone(),
                format: ImageFormat::Png,
                from_cell: "B2".to_string(),
                width_px: 400,
                height_px: 300,
            },
        )
        .unwrap();

        let pics = wb.get_pictures("Sheet1", "B2").unwrap();
        assert_eq!(pics.len(), 1);
        assert_eq!(pics[0].data, image_data);
        assert_eq!(pics[0].format, ImageFormat::Png);
        assert_eq!(pics[0].cell, "B2");
        assert_eq!(pics[0].width_px, 400);
        assert_eq!(pics[0].height_px, 300);
    }

    #[test]
    fn test_get_pictures_empty_cell() {
        let wb = Workbook::new();
        let pics = wb.get_pictures("Sheet1", "A1").unwrap();
        assert!(pics.is_empty());
    }

    #[test]
    fn test_get_pictures_wrong_cell() {
        use crate::image::{ImageConfig, ImageFormat};
        let mut wb = Workbook::new();
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x89, 0x50],
                format: ImageFormat::Png,
                from_cell: "B2".to_string(),
                width_px: 100,
                height_px: 100,
            },
        )
        .unwrap();

        let pics = wb.get_pictures("Sheet1", "A1").unwrap();
        assert!(pics.is_empty());
    }

    #[test]
    fn test_delete_chart_roundtrip() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let dir = TempDir::new().unwrap();
        let path1 = dir.path().join("chart_delete_rt1.xlsx");
        let path2 = dir.path().join("chart_delete_rt2.xlsx");

        let mut wb = Workbook::new();
        wb.add_chart(
            "Sheet1",
            "E1",
            "L10",
            &ChartConfig {
                chart_type: ChartType::Col,
                title: Some("Chart".to_string()),
                series: vec![ChartSeries {
                    name: "S1".to_string(),
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
        wb.save(&path1).unwrap();

        // Delete the chart from the in-memory workbook before save
        wb.delete_chart("Sheet1", "E1").unwrap();
        wb.save(&path2).unwrap();

        let file = std::fs::File::open(&path2).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/charts/chart1.xml").is_err());
    }

    #[test]
    fn test_delete_picture_roundtrip() {
        use crate::image::{ImageConfig, ImageFormat};
        let dir = TempDir::new().unwrap();
        let path1 = dir.path().join("pic_delete_rt1.xlsx");
        let path2 = dir.path().join("pic_delete_rt2.xlsx");

        let mut wb = Workbook::new();
        wb.add_image(
            "Sheet1",
            &ImageConfig {
                data: vec![0x89, 0x50, 0x4E, 0x47],
                format: ImageFormat::Png,
                from_cell: "B2".to_string(),
                width_px: 200,
                height_px: 150,
            },
        )
        .unwrap();
        wb.save(&path1).unwrap();

        // Delete the picture from the in-memory workbook before save
        wb.delete_picture("Sheet1", "B2").unwrap();
        wb.save(&path2).unwrap();

        let file = std::fs::File::open(&path2).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/media/image1.png").is_err());
    }

    #[test]
    fn test_delete_chart_preserves_image() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        use crate::image::{ImageConfig, ImageFormat};

        let mut wb = Workbook::new();
        wb.add_chart(
            "Sheet1",
            "E1",
            "L10",
            &ChartConfig {
                chart_type: ChartType::Col,
                title: None,
                series: vec![ChartSeries {
                    name: "S1".to_string(),
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
                data: vec![0x89, 0x50, 0x4E, 0x47],
                format: ImageFormat::Png,
                from_cell: "E12".to_string(),
                width_px: 200,
                height_px: 150,
            },
        )
        .unwrap();

        wb.delete_chart("Sheet1", "E1").unwrap();
        assert_eq!(wb.charts.len(), 0);
        assert_eq!(wb.images.len(), 1);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 0);
    }

    #[test]
    fn test_delete_picture_preserves_chart() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        use crate::image::{ImageConfig, ImageFormat};

        let mut wb = Workbook::new();
        wb.add_chart(
            "Sheet1",
            "E1",
            "L10",
            &ChartConfig {
                chart_type: ChartType::Col,
                title: None,
                series: vec![ChartSeries {
                    name: "S1".to_string(),
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
                data: vec![0x89, 0x50, 0x4E, 0x47],
                format: ImageFormat::Png,
                from_cell: "E12".to_string(),
                width_px: 200,
                height_px: 150,
            },
        )
        .unwrap();

        wb.delete_picture("Sheet1", "E12").unwrap();
        assert_eq!(wb.images.len(), 0);
        assert_eq!(wb.charts.len(), 1);
        assert_eq!(wb.drawings[0].1.two_cell_anchors.len(), 1);
        assert_eq!(wb.drawings[0].1.one_cell_anchors.len(), 0);
    }
}
