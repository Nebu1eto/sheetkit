//! Shape insertion and management.
//!
//! Provides types for configuring preset geometry shapes in worksheets
//! and helpers for building the corresponding drawing XML structures.

use sheetkit_xml::drawing::{
    AExt, BodyPr, CNvPr, CNvSpPr, ClientData, Ln, LstStyle, MarkerType, NvSpPr, Offset, Paragraph,
    PrstGeom, RunProperties, Shape, ShapeSpPr, SolidFill, SrgbClr, TextRun, TwoCellAnchor, TxBody,
    Xfrm,
};

use crate::error::{Error, Result};
use crate::utils::cell_ref::cell_name_to_coordinates;

/// Preset geometry shape types supported by OOXML.
#[derive(Debug, Clone, PartialEq)]
pub enum ShapeType {
    Rect,
    RoundRect,
    Ellipse,
    Triangle,
    Diamond,
    Pentagon,
    Hexagon,
    Octagon,
    RightArrow,
    LeftArrow,
    UpArrow,
    DownArrow,
    LeftRightArrow,
    UpDownArrow,
    Star4,
    Star5,
    Star6,
    FlowchartProcess,
    FlowchartDecision,
    FlowchartTerminator,
    FlowchartData,
    Heart,
    Lightning,
    Plus,
    Minus,
    Cloud,
    Callout1,
    Callout2,
}

impl ShapeType {
    /// Return the OOXML preset geometry string for this shape type.
    pub fn preset_name(&self) -> &str {
        match self {
            ShapeType::Rect => "rect",
            ShapeType::RoundRect => "roundRect",
            ShapeType::Ellipse => "ellipse",
            ShapeType::Triangle => "triangle",
            ShapeType::Diamond => "diamond",
            ShapeType::Pentagon => "pentagon",
            ShapeType::Hexagon => "hexagon",
            ShapeType::Octagon => "octagon",
            ShapeType::RightArrow => "rightArrow",
            ShapeType::LeftArrow => "leftArrow",
            ShapeType::UpArrow => "upArrow",
            ShapeType::DownArrow => "downArrow",
            ShapeType::LeftRightArrow => "leftRightArrow",
            ShapeType::UpDownArrow => "upDownArrow",
            ShapeType::Star4 => "star4",
            ShapeType::Star5 => "star5",
            ShapeType::Star6 => "star6",
            ShapeType::FlowchartProcess => "flowChartProcess",
            ShapeType::FlowchartDecision => "flowChartDecision",
            ShapeType::FlowchartTerminator => "flowChartTerminator",
            ShapeType::FlowchartData => "flowChartInputOutput",
            ShapeType::Heart => "heart",
            ShapeType::Lightning => "lightningBolt",
            ShapeType::Plus => "mathPlus",
            ShapeType::Minus => "mathMinus",
            ShapeType::Cloud => "cloud",
            ShapeType::Callout1 => "wedgeRectCallout",
            ShapeType::Callout2 => "wedgeRoundRectCallout",
        }
    }

    /// Parse a string into a `ShapeType`.
    ///
    /// Accepts both the camelCase OOXML preset names and simplified lowercase
    /// identifiers (e.g., `"rect"`, `"roundRect"`, `"ellipse"`).
    pub fn parse(s: &str) -> Result<Self> {
        match s.to_lowercase().as_str() {
            "rect" | "rectangle" => Ok(ShapeType::Rect),
            "roundrect" | "roundedrectangle" => Ok(ShapeType::RoundRect),
            "ellipse" | "circle" | "oval" => Ok(ShapeType::Ellipse),
            "triangle" => Ok(ShapeType::Triangle),
            "diamond" => Ok(ShapeType::Diamond),
            "pentagon" => Ok(ShapeType::Pentagon),
            "hexagon" => Ok(ShapeType::Hexagon),
            "octagon" => Ok(ShapeType::Octagon),
            "rightarrow" => Ok(ShapeType::RightArrow),
            "leftarrow" => Ok(ShapeType::LeftArrow),
            "uparrow" => Ok(ShapeType::UpArrow),
            "downarrow" => Ok(ShapeType::DownArrow),
            "leftrightarrow" => Ok(ShapeType::LeftRightArrow),
            "updownarrow" => Ok(ShapeType::UpDownArrow),
            "star4" => Ok(ShapeType::Star4),
            "star5" => Ok(ShapeType::Star5),
            "star6" => Ok(ShapeType::Star6),
            "flowchartprocess" => Ok(ShapeType::FlowchartProcess),
            "flowchartdecision" => Ok(ShapeType::FlowchartDecision),
            "flowchartterminator" => Ok(ShapeType::FlowchartTerminator),
            "flowchartdata" => Ok(ShapeType::FlowchartData),
            "heart" => Ok(ShapeType::Heart),
            "lightning" | "lightningbolt" => Ok(ShapeType::Lightning),
            "plus" | "mathplus" => Ok(ShapeType::Plus),
            "minus" | "mathminus" => Ok(ShapeType::Minus),
            "cloud" => Ok(ShapeType::Cloud),
            "callout1" | "wedgerectcallout" => Ok(ShapeType::Callout1),
            "callout2" | "wedgeroundrectcallout" => Ok(ShapeType::Callout2),
            _ => Err(Error::Internal(format!("unknown shape type: {s}"))),
        }
    }
}

/// Configuration for inserting a shape into a worksheet.
#[derive(Debug, Clone)]
pub struct ShapeConfig {
    /// Preset geometry shape type.
    pub shape_type: ShapeType,
    /// Top-left anchor cell (e.g., `"B2"`).
    pub from_cell: String,
    /// Bottom-right anchor cell (e.g., `"F10"`).
    pub to_cell: String,
    /// Optional text content displayed inside the shape.
    pub text: Option<String>,
    /// Optional fill color as a hex string (e.g., `"4472C4"`).
    pub fill_color: Option<String>,
    /// Optional line/border color as a hex string (e.g., `"2F528F"`).
    pub line_color: Option<String>,
    /// Optional line width in points. Converted to EMU internally.
    pub line_width: Option<f64>,
}

/// Points-to-EMU conversion factor. 1 point = 12700 EMU.
const EMU_PER_POINT: f64 = 12700.0;

/// Build a `TwoCellAnchor` containing a shape.
///
/// The shape spans from `config.from_cell` to `config.to_cell`. An
/// incrementing `shape_id` is used for the non-visual properties.
pub fn build_shape_anchor(config: &ShapeConfig, shape_id: u32) -> Result<TwoCellAnchor> {
    let (from_col, from_row) = cell_name_to_coordinates(&config.from_cell)?;
    let (to_col, to_row) = cell_name_to_coordinates(&config.to_cell)?;

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

    let solid_fill = config.fill_color.as_ref().map(|color| SolidFill {
        srgb_clr: SrgbClr { val: color.clone() },
    });

    let ln = if config.line_color.is_some() || config.line_width.is_some() {
        Some(Ln {
            w: config.line_width.map(|pts| (pts * EMU_PER_POINT) as u64),
            solid_fill: config.line_color.as_ref().map(|color| SolidFill {
                srgb_clr: SrgbClr { val: color.clone() },
            }),
        })
    } else {
        None
    };

    let tx_body = config.text.as_ref().map(|text| TxBody {
        body_pr: BodyPr {},
        lst_style: LstStyle {},
        paragraphs: vec![Paragraph {
            runs: vec![TextRun {
                r_pr: Some(RunProperties {
                    lang: Some("en-US".to_string()),
                    sz: None,
                }),
                t: text.clone(),
            }],
        }],
    });

    let shape = Shape {
        nv_sp_pr: NvSpPr {
            c_nv_pr: CNvPr {
                id: shape_id,
                name: format!("Shape {}", shape_id),
            },
            c_nv_sp_pr: CNvSpPr {},
        },
        sp_pr: ShapeSpPr {
            xfrm: Xfrm {
                off: Offset { x: 0, y: 0 },
                ext: AExt { cx: 0, cy: 0 },
            },
            prst_geom: PrstGeom {
                prst: config.shape_type.preset_name().to_string(),
            },
            solid_fill,
            ln,
        },
        tx_body,
    };

    Ok(TwoCellAnchor {
        from: from_marker,
        to: to_marker,
        graphic_frame: None,
        pic: None,
        shape: Some(shape),
        client_data: ClientData {},
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::workbook::Workbook;
    use tempfile::TempDir;

    #[test]
    fn test_shape_type_preset_names() {
        assert_eq!(ShapeType::Rect.preset_name(), "rect");
        assert_eq!(ShapeType::RoundRect.preset_name(), "roundRect");
        assert_eq!(ShapeType::Ellipse.preset_name(), "ellipse");
        assert_eq!(ShapeType::Triangle.preset_name(), "triangle");
        assert_eq!(ShapeType::Diamond.preset_name(), "diamond");
        assert_eq!(ShapeType::Heart.preset_name(), "heart");
        assert_eq!(ShapeType::Lightning.preset_name(), "lightningBolt");
        assert_eq!(ShapeType::Plus.preset_name(), "mathPlus");
        assert_eq!(ShapeType::Minus.preset_name(), "mathMinus");
        assert_eq!(ShapeType::Cloud.preset_name(), "cloud");
        assert_eq!(ShapeType::Callout1.preset_name(), "wedgeRectCallout");
        assert_eq!(ShapeType::Callout2.preset_name(), "wedgeRoundRectCallout");
        assert_eq!(
            ShapeType::FlowchartProcess.preset_name(),
            "flowChartProcess"
        );
        assert_eq!(
            ShapeType::FlowchartDecision.preset_name(),
            "flowChartDecision"
        );
    }

    #[test]
    fn test_shape_type_from_str() {
        assert_eq!(ShapeType::parse("rect").unwrap(), ShapeType::Rect);
        assert_eq!(ShapeType::parse("rectangle").unwrap(), ShapeType::Rect);
        assert_eq!(ShapeType::parse("roundRect").unwrap(), ShapeType::RoundRect);
        assert_eq!(ShapeType::parse("ellipse").unwrap(), ShapeType::Ellipse);
        assert_eq!(ShapeType::parse("circle").unwrap(), ShapeType::Ellipse);
        assert_eq!(ShapeType::parse("heart").unwrap(), ShapeType::Heart);
        assert_eq!(ShapeType::parse("lightning").unwrap(), ShapeType::Lightning);
        assert_eq!(ShapeType::parse("callout1").unwrap(), ShapeType::Callout1);
        assert!(ShapeType::parse("nonexistent").is_err());
    }

    #[test]
    fn test_build_shape_anchor_basic() {
        let config = ShapeConfig {
            shape_type: ShapeType::Rect,
            from_cell: "B2".to_string(),
            to_cell: "F10".to_string(),
            text: None,
            fill_color: None,
            line_color: None,
            line_width: None,
        };

        let anchor = build_shape_anchor(&config, 2).unwrap();
        assert_eq!(anchor.from.col, 1);
        assert_eq!(anchor.from.row, 1);
        assert_eq!(anchor.to.col, 5);
        assert_eq!(anchor.to.row, 9);
        assert!(anchor.graphic_frame.is_none());
        assert!(anchor.pic.is_none());

        let shape = anchor.shape.as_ref().unwrap();
        assert_eq!(shape.nv_sp_pr.c_nv_pr.id, 2);
        assert_eq!(shape.nv_sp_pr.c_nv_pr.name, "Shape 2");
        assert_eq!(shape.sp_pr.prst_geom.prst, "rect");
        assert!(shape.sp_pr.solid_fill.is_none());
        assert!(shape.sp_pr.ln.is_none());
        assert!(shape.tx_body.is_none());
    }

    #[test]
    fn test_build_shape_anchor_with_text() {
        let config = ShapeConfig {
            shape_type: ShapeType::Ellipse,
            from_cell: "A1".to_string(),
            to_cell: "D5".to_string(),
            text: Some("Hello World".to_string()),
            fill_color: None,
            line_color: None,
            line_width: None,
        };

        let anchor = build_shape_anchor(&config, 3).unwrap();
        let shape = anchor.shape.as_ref().unwrap();
        assert_eq!(shape.sp_pr.prst_geom.prst, "ellipse");

        let tx_body = shape.tx_body.as_ref().unwrap();
        assert_eq!(tx_body.paragraphs.len(), 1);
        assert_eq!(tx_body.paragraphs[0].runs.len(), 1);
        assert_eq!(tx_body.paragraphs[0].runs[0].t, "Hello World");
    }

    #[test]
    fn test_build_shape_anchor_with_fill_and_line() {
        let config = ShapeConfig {
            shape_type: ShapeType::Diamond,
            from_cell: "C3".to_string(),
            to_cell: "H8".to_string(),
            text: None,
            fill_color: Some("4472C4".to_string()),
            line_color: Some("2F528F".to_string()),
            line_width: Some(2.0),
        };

        let anchor = build_shape_anchor(&config, 4).unwrap();
        let shape = anchor.shape.as_ref().unwrap();
        assert_eq!(shape.sp_pr.prst_geom.prst, "diamond");

        let fill = shape.sp_pr.solid_fill.as_ref().unwrap();
        assert_eq!(fill.srgb_clr.val, "4472C4");

        let ln = shape.sp_pr.ln.as_ref().unwrap();
        assert_eq!(ln.w, Some(25400));
        let ln_fill = ln.solid_fill.as_ref().unwrap();
        assert_eq!(ln_fill.srgb_clr.val, "2F528F");
    }

    #[test]
    fn test_build_shape_anchor_invalid_cell() {
        let config = ShapeConfig {
            shape_type: ShapeType::Rect,
            from_cell: "INVALID".to_string(),
            to_cell: "F10".to_string(),
            text: None,
            fill_color: None,
            line_color: None,
            line_width: None,
        };

        assert!(build_shape_anchor(&config, 1).is_err());
    }

    #[test]
    fn test_build_shape_various_types() {
        let types = vec![
            (ShapeType::Star4, "star4"),
            (ShapeType::Star5, "star5"),
            (ShapeType::Star6, "star6"),
            (ShapeType::RightArrow, "rightArrow"),
            (ShapeType::LeftArrow, "leftArrow"),
            (ShapeType::UpArrow, "upArrow"),
            (ShapeType::DownArrow, "downArrow"),
            (ShapeType::Pentagon, "pentagon"),
            (ShapeType::Hexagon, "hexagon"),
            (ShapeType::Octagon, "octagon"),
        ];

        for (shape_type, expected_preset) in types {
            let config = ShapeConfig {
                shape_type,
                from_cell: "A1".to_string(),
                to_cell: "C3".to_string(),
                text: None,
                fill_color: None,
                line_color: None,
                line_width: None,
            };
            let anchor = build_shape_anchor(&config, 1).unwrap();
            let shape = anchor.shape.as_ref().unwrap();
            assert_eq!(shape.sp_pr.prst_geom.prst, expected_preset);
        }
    }

    #[test]
    fn test_add_shape_to_workbook_and_save() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("add_shape_basic.xlsx");

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
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/drawings/drawing1.xml").is_ok());

        let drawing_xml = {
            use std::io::Read;
            let mut buf = String::new();
            archive
                .by_name("xl/drawings/drawing1.xml")
                .unwrap()
                .read_to_string(&mut buf)
                .unwrap();
            buf
        };
        assert!(drawing_xml.contains("rect"));
        assert!(drawing_xml.contains("FF0000"));
        assert!(drawing_xml.contains("Test Shape"));
    }

    #[test]
    fn test_add_shape_sheet_not_found() {
        let mut wb = Workbook::new();
        let config = ShapeConfig {
            shape_type: ShapeType::Rect,
            from_cell: "A1".to_string(),
            to_cell: "C3".to_string(),
            text: None,
            fill_color: None,
            line_color: None,
            line_width: None,
        };
        let result = wb.add_shape("NoSheet", &config);
        assert!(matches!(
            result.unwrap_err(),
            crate::error::Error::SheetNotFound { .. }
        ));
    }

    #[test]
    fn test_add_multiple_shapes_same_sheet_save() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("multi_shape.xlsx");

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
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        let drawing_xml = {
            use std::io::Read;
            let mut buf = String::new();
            archive
                .by_name("xl/drawings/drawing1.xml")
                .unwrap()
                .read_to_string(&mut buf)
                .unwrap();
            buf
        };
        assert!(drawing_xml.contains("rect"));
        assert!(drawing_xml.contains("ellipse"));
        assert!(drawing_xml.contains("Circle"));
        assert!(drawing_xml.contains("00FF00"));
    }

    #[test]
    fn test_save_with_shape() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("with_shape.xlsx");

        let mut wb = Workbook::new();
        let config = ShapeConfig {
            shape_type: ShapeType::RoundRect,
            from_cell: "B2".to_string(),
            to_cell: "F10".to_string(),
            text: Some("Hello".to_string()),
            fill_color: Some("4472C4".to_string()),
            line_color: Some("2F528F".to_string()),
            line_width: Some(1.5),
        };
        wb.add_shape("Sheet1", &config).unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/drawings/drawing1.xml").is_ok());

        let drawing_xml = {
            use std::io::Read;
            let mut buf = String::new();
            archive
                .by_name("xl/drawings/drawing1.xml")
                .unwrap()
                .read_to_string(&mut buf)
                .unwrap();
            buf
        };
        assert!(drawing_xml.contains("roundRect"));
        assert!(drawing_xml.contains("4472C4"));
        assert!(drawing_xml.contains("Hello"));
    }

    #[test]
    fn test_save_shape_roundtrip() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("shape_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.add_shape(
            "Sheet1",
            &ShapeConfig {
                shape_type: ShapeType::Heart,
                from_cell: "A1".to_string(),
                to_cell: "D5".to_string(),
                text: None,
                fill_color: Some("FF0000".to_string()),
                line_color: None,
                line_width: None,
            },
        )
        .unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let ws = wb2.worksheet_ref("Sheet1").unwrap();
        assert!(ws.drawing.is_some());
    }

    #[test]
    fn test_add_shape_with_chart_same_sheet_save() {
        use crate::chart::{ChartConfig, ChartSeries, ChartType};
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("shape_and_chart.xlsx");

        let mut wb = Workbook::new();
        let chart_config = ChartConfig {
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
        wb.add_chart("Sheet1", "E1", "L10", &chart_config).unwrap();

        let shape_config = ShapeConfig {
            shape_type: ShapeType::Rect,
            from_cell: "A12".to_string(),
            to_cell: "D18".to_string(),
            text: Some("Label".to_string()),
            fill_color: None,
            line_color: None,
            line_width: None,
        };
        wb.add_shape("Sheet1", &shape_config).unwrap();
        wb.save(&path).unwrap();

        let file = std::fs::File::open(&path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();
        assert!(archive.by_name("xl/drawings/drawing1.xml").is_ok());
        assert!(archive.by_name("xl/charts/chart1.xml").is_ok());

        let drawing_xml = {
            use std::io::Read;
            let mut buf = String::new();
            archive
                .by_name("xl/drawings/drawing1.xml")
                .unwrap()
                .read_to_string(&mut buf)
                .unwrap();
            buf
        };
        assert!(drawing_xml.contains("rect"));
        assert!(drawing_xml.contains("Label"));
    }

    #[test]
    fn test_shape_line_width_only() {
        let config = ShapeConfig {
            shape_type: ShapeType::Rect,
            from_cell: "A1".to_string(),
            to_cell: "C3".to_string(),
            text: None,
            fill_color: None,
            line_color: None,
            line_width: Some(3.0),
        };
        let anchor = build_shape_anchor(&config, 1).unwrap();
        let shape = anchor.shape.as_ref().unwrap();
        let ln = shape.sp_pr.ln.as_ref().unwrap();
        assert_eq!(ln.w, Some(38100));
        assert!(ln.solid_fill.is_none());
    }
}
