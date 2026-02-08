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
}
