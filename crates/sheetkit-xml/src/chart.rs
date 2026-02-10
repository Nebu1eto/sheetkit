//! Chart XML schema structures.
//!
//! Represents `xl/charts/chart{N}.xml` in the OOXML package.

use serde::{Deserialize, Serialize};

use crate::namespaces;

/// Root element for a chart part.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "chartSpace")]
pub struct ChartSpace {
    #[serde(rename = "@xmlns:c")]
    pub xmlns_c: String,

    #[serde(rename = "@xmlns:a")]
    pub xmlns_a: String,

    #[serde(rename = "@xmlns:r")]
    pub xmlns_r: String,

    #[serde(rename = "c:chart")]
    pub chart: Chart,
}

impl Default for ChartSpace {
    fn default() -> Self {
        Self {
            xmlns_c: namespaces::DRAWING_ML_CHART.to_string(),
            xmlns_a: namespaces::DRAWING_ML.to_string(),
            xmlns_r: namespaces::RELATIONSHIPS.to_string(),
            chart: Chart::default(),
        }
    }
}

/// The chart element containing plot area, legend, and title.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct Chart {
    #[serde(rename = "c:title", skip_serializing_if = "Option::is_none")]
    pub title: Option<ChartTitle>,

    #[serde(rename = "c:view3D", skip_serializing_if = "Option::is_none")]
    pub view_3d: Option<View3D>,

    #[serde(rename = "c:plotArea")]
    pub plot_area: PlotArea,

    #[serde(rename = "c:legend", skip_serializing_if = "Option::is_none")]
    pub legend: Option<Legend>,

    #[serde(rename = "c:plotVisOnly", skip_serializing_if = "Option::is_none")]
    pub plot_vis_only: Option<BoolVal>,
}

/// Chart title.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ChartTitle {
    #[serde(rename = "c:tx")]
    pub tx: TitleTx,
}

/// Title text body.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TitleTx {
    #[serde(rename = "c:rich")]
    pub rich: RichText,
}

/// Rich text body for chart titles.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct RichText {
    #[serde(rename = "a:bodyPr")]
    pub body_pr: BodyPr,

    #[serde(rename = "a:p")]
    pub paragraphs: Vec<Paragraph>,
}

/// Body properties (empty marker for chart titles).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BodyPr {}

/// A paragraph in rich text.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Paragraph {
    #[serde(rename = "a:r", default)]
    pub runs: Vec<Run>,
}

/// A text run within a paragraph.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Run {
    #[serde(rename = "a:t")]
    pub t: String,
}

/// Plot area containing chart type definitions and axes.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct PlotArea {
    #[serde(rename = "c:layout", skip_serializing_if = "Option::is_none")]
    pub layout: Option<Layout>,

    #[serde(rename = "c:barChart", skip_serializing_if = "Option::is_none")]
    pub bar_chart: Option<BarChart>,

    #[serde(rename = "c:bar3DChart", skip_serializing_if = "Option::is_none")]
    pub bar_3d_chart: Option<Bar3DChart>,

    #[serde(rename = "c:lineChart", skip_serializing_if = "Option::is_none")]
    pub line_chart: Option<LineChart>,

    #[serde(rename = "c:line3DChart", skip_serializing_if = "Option::is_none")]
    pub line_3d_chart: Option<Line3DChart>,

    #[serde(rename = "c:pieChart", skip_serializing_if = "Option::is_none")]
    pub pie_chart: Option<PieChart>,

    #[serde(rename = "c:pie3DChart", skip_serializing_if = "Option::is_none")]
    pub pie_3d_chart: Option<Pie3DChart>,

    #[serde(rename = "c:doughnutChart", skip_serializing_if = "Option::is_none")]
    pub doughnut_chart: Option<DoughnutChart>,

    #[serde(rename = "c:areaChart", skip_serializing_if = "Option::is_none")]
    pub area_chart: Option<AreaChart>,

    #[serde(rename = "c:area3DChart", skip_serializing_if = "Option::is_none")]
    pub area_3d_chart: Option<Area3DChart>,

    #[serde(rename = "c:scatterChart", skip_serializing_if = "Option::is_none")]
    pub scatter_chart: Option<ScatterChart>,

    #[serde(rename = "c:bubbleChart", skip_serializing_if = "Option::is_none")]
    pub bubble_chart: Option<BubbleChart>,

    #[serde(rename = "c:radarChart", skip_serializing_if = "Option::is_none")]
    pub radar_chart: Option<RadarChart>,

    #[serde(rename = "c:stockChart", skip_serializing_if = "Option::is_none")]
    pub stock_chart: Option<StockChart>,

    #[serde(rename = "c:surfaceChart", skip_serializing_if = "Option::is_none")]
    pub surface_chart: Option<SurfaceChart>,

    #[serde(rename = "c:surface3DChart", skip_serializing_if = "Option::is_none")]
    pub surface_3d_chart: Option<Surface3DChart>,

    #[serde(rename = "c:ofPieChart", skip_serializing_if = "Option::is_none")]
    pub of_pie_chart: Option<OfPieChart>,

    #[serde(rename = "c:catAx", skip_serializing_if = "Option::is_none")]
    pub cat_ax: Option<CatAx>,

    #[serde(rename = "c:valAx", skip_serializing_if = "Option::is_none")]
    pub val_ax: Option<ValAx>,

    #[serde(rename = "c:serAx", skip_serializing_if = "Option::is_none")]
    pub ser_ax: Option<SerAx>,
}

/// Layout (empty, uses automatic layout).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Layout {}

/// Bar chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BarChart {
    #[serde(rename = "c:barDir")]
    pub bar_dir: StringVal,

    #[serde(rename = "c:grouping")]
    pub grouping: StringVal,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// 3D bar chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Bar3DChart {
    #[serde(rename = "c:barDir")]
    pub bar_dir: StringVal,

    #[serde(rename = "c:grouping")]
    pub grouping: StringVal,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:shape", skip_serializing_if = "Option::is_none")]
    pub shape: Option<StringVal>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// Line chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct LineChart {
    #[serde(rename = "c:grouping")]
    pub grouping: StringVal,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// 3D line chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Line3DChart {
    #[serde(rename = "c:grouping")]
    pub grouping: StringVal,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// Pie chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PieChart {
    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,
}

/// 3D pie chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Pie3DChart {
    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,
}

/// Doughnut chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DoughnutChart {
    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:holeSize", skip_serializing_if = "Option::is_none")]
    pub hole_size: Option<UintVal>,
}

/// Area chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct AreaChart {
    #[serde(rename = "c:grouping")]
    pub grouping: StringVal,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// 3D area chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Area3DChart {
    #[serde(rename = "c:grouping")]
    pub grouping: StringVal,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// Scatter chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ScatterChart {
    #[serde(rename = "c:scatterStyle")]
    pub scatter_style: StringVal,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<ScatterSeries>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// Scatter series (uses xVal/yVal instead of cat/val).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ScatterSeries {
    #[serde(rename = "c:idx")]
    pub idx: UintVal,

    #[serde(rename = "c:order")]
    pub order: UintVal,

    #[serde(rename = "c:tx", skip_serializing_if = "Option::is_none")]
    pub tx: Option<SeriesText>,

    #[serde(rename = "c:xVal", skip_serializing_if = "Option::is_none")]
    pub x_val: Option<CategoryRef>,

    #[serde(rename = "c:yVal", skip_serializing_if = "Option::is_none")]
    pub y_val: Option<ValueRef>,
}

/// Bubble chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BubbleChart {
    #[serde(rename = "c:ser", default)]
    pub series: Vec<BubbleSeries>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// Bubble series (uses xVal/yVal/bubbleSize).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BubbleSeries {
    #[serde(rename = "c:idx")]
    pub idx: UintVal,

    #[serde(rename = "c:order")]
    pub order: UintVal,

    #[serde(rename = "c:tx", skip_serializing_if = "Option::is_none")]
    pub tx: Option<SeriesText>,

    #[serde(rename = "c:xVal", skip_serializing_if = "Option::is_none")]
    pub x_val: Option<CategoryRef>,

    #[serde(rename = "c:yVal", skip_serializing_if = "Option::is_none")]
    pub y_val: Option<ValueRef>,

    #[serde(rename = "c:bubbleSize", skip_serializing_if = "Option::is_none")]
    pub bubble_size: Option<ValueRef>,

    #[serde(rename = "c:bubble3D", skip_serializing_if = "Option::is_none")]
    pub bubble_3d: Option<BoolVal>,
}

/// Radar chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct RadarChart {
    #[serde(rename = "c:radarStyle")]
    pub radar_style: StringVal,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// Stock chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct StockChart {
    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// Surface chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SurfaceChart {
    #[serde(rename = "c:wireframe", skip_serializing_if = "Option::is_none")]
    pub wireframe: Option<BoolVal>,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// 3D surface chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Surface3DChart {
    #[serde(rename = "c:wireframe", skip_serializing_if = "Option::is_none")]
    pub wireframe: Option<BoolVal>,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:axId", default)]
    pub ax_ids: Vec<UintVal>,
}

/// Of-pie chart definition (pie-of-pie or bar-of-pie).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct OfPieChart {
    #[serde(rename = "c:ofPieType")]
    pub of_pie_type: StringVal,

    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,

    #[serde(rename = "c:serLines", skip_serializing_if = "Option::is_none")]
    pub ser_lines: Option<SerLines>,
}

/// Series lines marker element.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SerLines {}

/// A data series within a chart.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Series {
    #[serde(rename = "c:idx")]
    pub idx: UintVal,

    #[serde(rename = "c:order")]
    pub order: UintVal,

    #[serde(rename = "c:tx", skip_serializing_if = "Option::is_none")]
    pub tx: Option<SeriesText>,

    #[serde(rename = "c:cat", skip_serializing_if = "Option::is_none")]
    pub cat: Option<CategoryRef>,

    #[serde(rename = "c:val", skip_serializing_if = "Option::is_none")]
    pub val: Option<ValueRef>,
}

/// Series text (name) reference.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SeriesText {
    #[serde(rename = "c:strRef", skip_serializing_if = "Option::is_none")]
    pub str_ref: Option<StrRef>,

    #[serde(rename = "c:v", skip_serializing_if = "Option::is_none")]
    pub v: Option<String>,
}

/// String reference (a formula to a cell range).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct StrRef {
    #[serde(rename = "c:f")]
    pub f: String,
}

/// Category axis data reference.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CategoryRef {
    #[serde(rename = "c:strRef", skip_serializing_if = "Option::is_none")]
    pub str_ref: Option<StrRef>,

    #[serde(rename = "c:numRef", skip_serializing_if = "Option::is_none")]
    pub num_ref: Option<NumRef>,
}

/// Value axis data reference.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ValueRef {
    #[serde(rename = "c:numRef", skip_serializing_if = "Option::is_none")]
    pub num_ref: Option<NumRef>,
}

/// Numeric reference (a formula to a numeric cell range).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct NumRef {
    #[serde(rename = "c:f")]
    pub f: String,
}

/// Chart legend.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Legend {
    #[serde(rename = "c:legendPos")]
    pub legend_pos: StringVal,
}

/// Category axis.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CatAx {
    #[serde(rename = "c:axId")]
    pub ax_id: UintVal,

    #[serde(rename = "c:scaling")]
    pub scaling: Scaling,

    #[serde(rename = "c:delete")]
    pub delete: BoolVal,

    #[serde(rename = "c:axPos")]
    pub ax_pos: StringVal,

    #[serde(rename = "c:crossAx")]
    pub cross_ax: UintVal,
}

/// Value axis.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ValAx {
    #[serde(rename = "c:axId")]
    pub ax_id: UintVal,

    #[serde(rename = "c:scaling")]
    pub scaling: Scaling,

    #[serde(rename = "c:delete")]
    pub delete: BoolVal,

    #[serde(rename = "c:axPos")]
    pub ax_pos: StringVal,

    #[serde(rename = "c:crossAx")]
    pub cross_ax: UintVal,
}

/// Series axis (used by surface and some 3D charts).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SerAx {
    #[serde(rename = "c:axId")]
    pub ax_id: UintVal,

    #[serde(rename = "c:scaling")]
    pub scaling: Scaling,

    #[serde(rename = "c:delete")]
    pub delete: BoolVal,

    #[serde(rename = "c:axPos")]
    pub ax_pos: StringVal,

    #[serde(rename = "c:crossAx")]
    pub cross_ax: UintVal,
}

/// 3D view settings.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct View3D {
    #[serde(rename = "c:rotX", skip_serializing_if = "Option::is_none")]
    pub rot_x: Option<IntVal>,

    #[serde(rename = "c:rotY", skip_serializing_if = "Option::is_none")]
    pub rot_y: Option<IntVal>,

    #[serde(rename = "c:depthPercent", skip_serializing_if = "Option::is_none")]
    pub depth_percent: Option<UintVal>,

    #[serde(rename = "c:rAngAx", skip_serializing_if = "Option::is_none")]
    pub r_ang_ax: Option<BoolVal>,

    #[serde(rename = "c:perspective", skip_serializing_if = "Option::is_none")]
    pub perspective: Option<UintVal>,
}

/// Axis scaling (orientation).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Scaling {
    #[serde(rename = "c:orientation")]
    pub orientation: StringVal,
}

/// A wrapper for a string `val` attribute.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct StringVal {
    #[serde(rename = "@val")]
    pub val: String,
}

/// A wrapper for an unsigned integer `val` attribute.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct UintVal {
    #[serde(rename = "@val")]
    pub val: u32,
}

/// A wrapper for a signed integer `val` attribute.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct IntVal {
    #[serde(rename = "@val")]
    pub val: i32,
}

/// A wrapper for a boolean `val` attribute.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BoolVal {
    #[serde(rename = "@val")]
    pub val: bool,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_chart_space_default() {
        let cs = ChartSpace::default();
        assert_eq!(cs.xmlns_c, namespaces::DRAWING_ML_CHART);
        assert_eq!(cs.xmlns_a, namespaces::DRAWING_ML);
        assert_eq!(cs.xmlns_r, namespaces::RELATIONSHIPS);
    }

    #[test]
    fn test_string_val_serialize() {
        let sv = StringVal {
            val: "col".to_string(),
        };
        let xml = quick_xml::se::to_string(&sv).unwrap();
        assert!(xml.contains("val=\"col\""));
    }

    #[test]
    fn test_uint_val_serialize() {
        let uv = UintVal { val: 42 };
        let xml = quick_xml::se::to_string(&uv).unwrap();
        assert!(xml.contains("val=\"42\""));
    }

    #[test]
    fn test_int_val_serialize() {
        let iv = IntVal { val: -15 };
        let xml = quick_xml::se::to_string(&iv).unwrap();
        assert!(xml.contains("val=\"-15\""));
    }

    #[test]
    fn test_bool_val_serialize() {
        let bv = BoolVal { val: true };
        let xml = quick_xml::se::to_string(&bv).unwrap();
        assert!(xml.contains("val=\"true\""));
    }

    #[test]
    fn test_series_serialize() {
        let series = Series {
            idx: UintVal { val: 0 },
            order: UintVal { val: 0 },
            tx: Some(SeriesText {
                str_ref: None,
                v: Some("Sales".to_string()),
            }),
            cat: Some(CategoryRef {
                str_ref: Some(StrRef {
                    f: "Sheet1!$A$2:$A$6".to_string(),
                }),
                num_ref: None,
            }),
            val: Some(ValueRef {
                num_ref: Some(NumRef {
                    f: "Sheet1!$B$2:$B$6".to_string(),
                }),
            }),
        };
        let xml = quick_xml::se::to_string(&series).unwrap();
        assert!(xml.contains("Sheet1!$A$2:$A$6"));
        assert!(xml.contains("Sheet1!$B$2:$B$6"));
    }

    #[test]
    fn test_bar_chart_serialize() {
        let bar = BarChart {
            bar_dir: StringVal {
                val: "col".to_string(),
            },
            grouping: StringVal {
                val: "clustered".to_string(),
            },
            series: vec![],
            ax_ids: vec![UintVal { val: 1 }, UintVal { val: 2 }],
        };
        let xml = quick_xml::se::to_string(&bar).unwrap();
        assert!(xml.contains("col"));
        assert!(xml.contains("clustered"));
    }

    #[test]
    fn test_bar_3d_chart_serialize() {
        let bar = Bar3DChart {
            bar_dir: StringVal {
                val: "col".to_string(),
            },
            grouping: StringVal {
                val: "clustered".to_string(),
            },
            series: vec![],
            shape: None,
            ax_ids: vec![UintVal { val: 1 }, UintVal { val: 2 }],
        };
        let xml = quick_xml::se::to_string(&bar).unwrap();
        assert!(xml.contains("col"));
        assert!(xml.contains("clustered"));
    }

    #[test]
    fn test_area_chart_serialize() {
        let area = AreaChart {
            grouping: StringVal {
                val: "standard".to_string(),
            },
            series: vec![],
            ax_ids: vec![UintVal { val: 1 }, UintVal { val: 2 }],
        };
        let xml = quick_xml::se::to_string(&area).unwrap();
        assert!(xml.contains("standard"));
    }

    #[test]
    fn test_scatter_chart_serialize() {
        let scatter = ScatterChart {
            scatter_style: StringVal {
                val: "lineMarker".to_string(),
            },
            series: vec![ScatterSeries {
                idx: UintVal { val: 0 },
                order: UintVal { val: 0 },
                tx: None,
                x_val: Some(CategoryRef {
                    str_ref: None,
                    num_ref: Some(NumRef {
                        f: "Sheet1!$A$2:$A$6".to_string(),
                    }),
                }),
                y_val: Some(ValueRef {
                    num_ref: Some(NumRef {
                        f: "Sheet1!$B$2:$B$6".to_string(),
                    }),
                }),
            }],
            ax_ids: vec![UintVal { val: 1 }, UintVal { val: 2 }],
        };
        let xml = quick_xml::se::to_string(&scatter).unwrap();
        assert!(xml.contains("lineMarker"));
        assert!(xml.contains("Sheet1!$A$2:$A$6"));
        assert!(xml.contains("Sheet1!$B$2:$B$6"));
    }

    #[test]
    fn test_bubble_chart_serialize() {
        let bubble = BubbleChart {
            series: vec![BubbleSeries {
                idx: UintVal { val: 0 },
                order: UintVal { val: 0 },
                tx: None,
                x_val: None,
                y_val: None,
                bubble_size: Some(ValueRef {
                    num_ref: Some(NumRef {
                        f: "Sheet1!$C$2:$C$6".to_string(),
                    }),
                }),
                bubble_3d: None,
            }],
            ax_ids: vec![UintVal { val: 1 }, UintVal { val: 2 }],
        };
        let xml = quick_xml::se::to_string(&bubble).unwrap();
        assert!(xml.contains("Sheet1!$C$2:$C$6"));
    }

    #[test]
    fn test_radar_chart_serialize() {
        let radar = RadarChart {
            radar_style: StringVal {
                val: "marker".to_string(),
            },
            series: vec![],
            ax_ids: vec![UintVal { val: 1 }, UintVal { val: 2 }],
        };
        let xml = quick_xml::se::to_string(&radar).unwrap();
        assert!(xml.contains("marker"));
    }

    #[test]
    fn test_surface_chart_serialize() {
        let surface = SurfaceChart {
            wireframe: Some(BoolVal { val: true }),
            series: vec![],
            ax_ids: vec![UintVal { val: 1 }, UintVal { val: 2 }, UintVal { val: 3 }],
        };
        let xml = quick_xml::se::to_string(&surface).unwrap();
        assert!(xml.contains("val=\"true\""));
    }

    #[test]
    fn test_view_3d_serialize() {
        let view = View3D {
            rot_x: Some(IntVal { val: 15 }),
            rot_y: Some(IntVal { val: 20 }),
            depth_percent: Some(UintVal { val: 150 }),
            r_ang_ax: Some(BoolVal { val: true }),
            perspective: Some(UintVal { val: 30 }),
        };
        let xml = quick_xml::se::to_string(&view).unwrap();
        assert!(xml.contains("val=\"15\""));
        assert!(xml.contains("val=\"20\""));
        assert!(xml.contains("val=\"150\""));
    }

    #[test]
    fn test_ser_ax_serialize() {
        let ser_ax = SerAx {
            ax_id: UintVal { val: 3 },
            scaling: Scaling {
                orientation: StringVal {
                    val: "minMax".to_string(),
                },
            },
            delete: BoolVal { val: false },
            ax_pos: StringVal {
                val: "b".to_string(),
            },
            cross_ax: UintVal { val: 1 },
        };
        let xml = quick_xml::se::to_string(&ser_ax).unwrap();
        assert!(xml.contains("val=\"3\""));
        assert!(xml.contains("minMax"));
    }

    #[test]
    fn test_legend_serialize() {
        let legend = Legend {
            legend_pos: StringVal {
                val: "b".to_string(),
            },
        };
        let xml = quick_xml::se::to_string(&legend).unwrap();
        assert!(xml.contains("val=\"b\""));
    }

    #[test]
    fn test_chart_title_serialize() {
        let title = ChartTitle {
            tx: TitleTx {
                rich: RichText {
                    body_pr: BodyPr {},
                    paragraphs: vec![Paragraph {
                        runs: vec![Run {
                            t: "My Chart".to_string(),
                        }],
                    }],
                },
            },
        };
        let xml = quick_xml::se::to_string(&title).unwrap();
        assert!(xml.contains("My Chart"));
    }

    #[test]
    fn test_num_ref_serialize() {
        let num_ref = NumRef {
            f: "Sheet1!$B$1:$B$5".to_string(),
        };
        let xml = quick_xml::se::to_string(&num_ref).unwrap();
        assert!(xml.contains("Sheet1!$B$1:$B$5"));
    }

    #[test]
    fn test_str_ref_serialize() {
        let str_ref = StrRef {
            f: "Sheet1!$A$1".to_string(),
        };
        let xml = quick_xml::se::to_string(&str_ref).unwrap();
        assert!(xml.contains("Sheet1!$A$1"));
    }

    #[test]
    fn test_plot_area_default_all_none() {
        let pa = PlotArea::default();
        assert!(pa.layout.is_none());
        assert!(pa.bar_chart.is_none());
        assert!(pa.bar_3d_chart.is_none());
        assert!(pa.line_chart.is_none());
        assert!(pa.line_3d_chart.is_none());
        assert!(pa.pie_chart.is_none());
        assert!(pa.pie_3d_chart.is_none());
        assert!(pa.doughnut_chart.is_none());
        assert!(pa.area_chart.is_none());
        assert!(pa.area_3d_chart.is_none());
        assert!(pa.scatter_chart.is_none());
        assert!(pa.bubble_chart.is_none());
        assert!(pa.radar_chart.is_none());
        assert!(pa.stock_chart.is_none());
        assert!(pa.surface_chart.is_none());
        assert!(pa.surface_3d_chart.is_none());
        assert!(pa.of_pie_chart.is_none());
        assert!(pa.cat_ax.is_none());
        assert!(pa.val_ax.is_none());
        assert!(pa.ser_ax.is_none());
    }

    #[test]
    fn test_chart_with_view_3d() {
        let chart = Chart {
            title: None,
            view_3d: Some(View3D {
                rot_x: Some(IntVal { val: 15 }),
                rot_y: Some(IntVal { val: 20 }),
                depth_percent: None,
                r_ang_ax: Some(BoolVal { val: true }),
                perspective: None,
            }),
            plot_area: PlotArea::default(),
            legend: None,
            plot_vis_only: None,
        };
        let xml = quick_xml::se::to_string(&chart).unwrap();
        assert!(xml.contains("val=\"15\""));
        assert!(xml.contains("val=\"20\""));
    }

    #[test]
    fn test_doughnut_chart_serialize() {
        let doughnut = DoughnutChart {
            series: vec![],
            hole_size: Some(UintVal { val: 50 }),
        };
        let xml = quick_xml::se::to_string(&doughnut).unwrap();
        assert!(xml.contains("val=\"50\""));
    }

    #[test]
    fn test_bar_3d_chart_with_shape() {
        let bar = Bar3DChart {
            bar_dir: StringVal {
                val: "col".to_string(),
            },
            grouping: StringVal {
                val: "clustered".to_string(),
            },
            series: vec![],
            shape: Some(StringVal {
                val: "cone".to_string(),
            }),
            ax_ids: vec![UintVal { val: 1 }, UintVal { val: 2 }],
        };
        let xml = quick_xml::se::to_string(&bar).unwrap();
        assert!(xml.contains("cone"));
    }

    #[test]
    fn test_of_pie_chart_serialize() {
        let of_pie = OfPieChart {
            of_pie_type: StringVal {
                val: "pie".to_string(),
            },
            series: vec![Series {
                idx: UintVal { val: 0 },
                order: UintVal { val: 0 },
                tx: None,
                cat: None,
                val: None,
            }],
            ser_lines: Some(SerLines {}),
        };
        let xml = quick_xml::se::to_string(&of_pie).unwrap();
        assert!(xml.contains("val=\"pie\""));
        assert!(xml.contains("serLines"));
    }

    #[test]
    fn test_bubble_series_with_bubble_3d() {
        let bs = BubbleSeries {
            idx: UintVal { val: 0 },
            order: UintVal { val: 0 },
            tx: None,
            x_val: None,
            y_val: None,
            bubble_size: None,
            bubble_3d: Some(BoolVal { val: true }),
        };
        let xml = quick_xml::se::to_string(&bs).unwrap();
        assert!(xml.contains("bubble3D"));
        assert!(xml.contains("val=\"true\""));
    }
}
