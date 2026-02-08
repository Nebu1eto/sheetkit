//! Chart builder and management.
//!
//! Provides types for configuring charts and functions for building the
//! corresponding XML structures (ChartSpace and drawing anchors).

use sheetkit_xml::chart::{
    Area3DChart, AreaChart, Bar3DChart, BarChart, BodyPr, BoolVal, BubbleChart, BubbleSeries,
    CatAx, CategoryRef, Chart, ChartSpace, ChartTitle, DoughnutChart, IntVal, Layout, Legend,
    Line3DChart, LineChart, NumRef, Paragraph, Pie3DChart, PieChart, PlotArea, RadarChart,
    RichText, Run, Scaling, ScatterChart, ScatterSeries, SerAx, Series, SeriesText, StockChart,
    StrRef, StringVal, Surface3DChart, SurfaceChart, TitleTx, UintVal, ValAx, ValueRef, View3D,
};
use sheetkit_xml::drawing::{
    AExt, CNvGraphicFramePr, CNvPr, ChartRef, ClientData, Graphic, GraphicData, GraphicFrame,
    MarkerType, NvGraphicFramePr, Offset, TwoCellAnchor, WsDr, Xfrm,
};
use sheetkit_xml::namespaces;

/// The chart type to render.
#[derive(Debug, Clone, PartialEq)]
pub enum ChartType {
    /// Vertical bar chart (column).
    Col,
    /// Vertical bar chart, stacked.
    ColStacked,
    /// Vertical bar chart, percent stacked.
    ColPercentStacked,
    /// Horizontal bar chart.
    Bar,
    /// Horizontal bar chart, stacked.
    BarStacked,
    /// Horizontal bar chart, percent stacked.
    BarPercentStacked,
    /// Line chart.
    Line,
    /// Pie chart.
    Pie,
    /// Area chart.
    Area,
    /// Area chart, stacked.
    AreaStacked,
    /// Area chart, percent stacked.
    AreaPercentStacked,
    /// 3D area chart.
    Area3D,
    /// 3D area chart, stacked.
    Area3DStacked,
    /// 3D area chart, percent stacked.
    Area3DPercentStacked,
    /// 3D column chart.
    Col3D,
    /// 3D column chart, stacked.
    Col3DStacked,
    /// 3D column chart, percent stacked.
    Col3DPercentStacked,
    /// 3D horizontal bar chart.
    Bar3D,
    /// 3D horizontal bar chart, stacked.
    Bar3DStacked,
    /// 3D horizontal bar chart, percent stacked.
    Bar3DPercentStacked,
    /// Line chart, stacked.
    LineStacked,
    /// Line chart, percent stacked.
    LinePercentStacked,
    /// 3D line chart.
    Line3D,
    /// 3D pie chart.
    Pie3D,
    /// Doughnut chart.
    Doughnut,
    /// Scatter chart (markers only).
    Scatter,
    /// Scatter chart with straight lines.
    ScatterLine,
    /// Scatter chart with smooth lines.
    ScatterSmooth,
    /// Radar chart (standard).
    Radar,
    /// Radar chart with filled area.
    RadarFilled,
    /// Radar chart with markers.
    RadarMarker,
    /// Stock chart: High-Low-Close.
    StockHLC,
    /// Stock chart: Open-High-Low-Close.
    StockOHLC,
    /// Stock chart: Volume-High-Low-Close.
    StockVHLC,
    /// Stock chart: Volume-Open-High-Low-Close.
    StockVOHLC,
    /// Bubble chart.
    Bubble,
    /// Surface chart.
    Surface,
    /// 3D surface chart.
    Surface3D,
    /// Wireframe surface chart.
    SurfaceWireframe,
    /// Wireframe 3D surface chart.
    SurfaceWireframe3D,
    /// Combo chart: column + line.
    ColLine,
    /// Combo chart: column stacked + line.
    ColLineStacked,
    /// Combo chart: column percent stacked + line.
    ColLinePercentStacked,
}

/// 3D view configuration for charts.
#[derive(Debug, Clone, Default)]
pub struct View3DConfig {
    /// X-axis rotation angle in degrees.
    pub rot_x: Option<i32>,
    /// Y-axis rotation angle in degrees.
    pub rot_y: Option<i32>,
    /// Depth as a percentage (100 = normal depth).
    pub depth_percent: Option<u32>,
    /// Whether to use right angle axes.
    pub right_angle_axes: Option<bool>,
    /// Perspective angle in degrees.
    pub perspective: Option<u32>,
}

/// Configuration for a chart.
#[derive(Debug, Clone)]
pub struct ChartConfig {
    /// The type of chart.
    pub chart_type: ChartType,
    /// Optional chart title.
    pub title: Option<String>,
    /// Data series for the chart.
    pub series: Vec<ChartSeries>,
    /// Whether to show the legend.
    pub show_legend: bool,
    /// Optional 3D view settings (auto-populated for 3D chart types if not set).
    pub view_3d: Option<View3DConfig>,
}

/// A single data series within a chart.
#[derive(Debug, Clone)]
pub struct ChartSeries {
    /// Series name (a literal string or cell reference like `"Sheet1!$A$1"`).
    pub name: String,
    /// Category axis data range (e.g., `"Sheet1!$A$2:$A$6"`).
    pub categories: String,
    /// Value axis data range (e.g., `"Sheet1!$B$2:$B$6"`).
    pub values: String,
    /// X-axis values for scatter/bubble charts (e.g., `"Sheet1!$A$2:$A$6"`).
    pub x_values: Option<String>,
    /// Bubble sizes for bubble charts (e.g., `"Sheet1!$C$2:$C$6"`).
    pub bubble_sizes: Option<String>,
}

/// Build a `ChartSpace` XML structure from a chart configuration.
pub fn build_chart_xml(config: &ChartConfig) -> ChartSpace {
    let title = config.title.as_ref().map(|t| build_chart_title(t));
    let legend = if config.show_legend {
        Some(Legend {
            legend_pos: StringVal {
                val: "b".to_string(),
            },
        })
    } else {
        None
    };
    let view_3d = build_view_3d(config);
    let plot_area = build_plot_area(config);

    ChartSpace {
        chart: Chart {
            title,
            view_3d,
            plot_area,
            legend,
            plot_vis_only: Some(BoolVal { val: true }),
        },
        ..ChartSpace::default()
    }
}

/// Build a drawing XML structure containing a chart reference.
pub fn build_drawing_with_chart(chart_ref_id: &str, from: MarkerType, to: MarkerType) -> WsDr {
    let graphic_frame = GraphicFrame {
        nv_graphic_frame_pr: NvGraphicFramePr {
            c_nv_pr: CNvPr {
                id: 2,
                name: "Chart 1".to_string(),
            },
            c_nv_graphic_frame_pr: CNvGraphicFramePr {},
        },
        xfrm: Xfrm {
            off: Offset { x: 0, y: 0 },
            ext: AExt { cx: 0, cy: 0 },
        },
        graphic: Graphic {
            graphic_data: GraphicData {
                uri: namespaces::DRAWING_ML_CHART.to_string(),
                chart: ChartRef {
                    xmlns_c: namespaces::DRAWING_ML_CHART.to_string(),
                    r_id: chart_ref_id.to_string(),
                },
            },
        },
    };
    let anchor = TwoCellAnchor {
        from,
        to,
        graphic_frame: Some(graphic_frame),
        pic: None,
        client_data: ClientData {},
    };
    WsDr {
        two_cell_anchors: vec![anchor],
        ..WsDr::default()
    }
}

fn is_no_axis_chart(ct: &ChartType) -> bool {
    matches!(ct, ChartType::Pie | ChartType::Pie3D | ChartType::Doughnut)
}

fn is_3d_chart(ct: &ChartType) -> bool {
    matches!(
        ct,
        ChartType::Area3D
            | ChartType::Area3DStacked
            | ChartType::Area3DPercentStacked
            | ChartType::Col3D
            | ChartType::Col3DStacked
            | ChartType::Col3DPercentStacked
            | ChartType::Bar3D
            | ChartType::Bar3DStacked
            | ChartType::Bar3DPercentStacked
            | ChartType::Line3D
            | ChartType::Pie3D
            | ChartType::Surface3D
            | ChartType::SurfaceWireframe3D
    )
}

fn needs_ser_ax(ct: &ChartType) -> bool {
    matches!(
        ct,
        ChartType::Surface
            | ChartType::Surface3D
            | ChartType::SurfaceWireframe
            | ChartType::SurfaceWireframe3D
    )
}

fn build_view_3d(config: &ChartConfig) -> Option<View3D> {
    if let Some(v) = &config.view_3d {
        return Some(View3D {
            rot_x: v.rot_x.map(|val| IntVal { val }),
            rot_y: v.rot_y.map(|val| IntVal { val }),
            depth_percent: v.depth_percent.map(|val| UintVal { val }),
            r_ang_ax: v.right_angle_axes.map(|val| BoolVal { val }),
            perspective: v.perspective.map(|val| UintVal { val }),
        });
    }
    if is_3d_chart(&config.chart_type) {
        Some(View3D {
            rot_x: Some(IntVal { val: 15 }),
            rot_y: Some(IntVal { val: 20 }),
            depth_percent: None,
            r_ang_ax: Some(BoolVal { val: true }),
            perspective: Some(UintVal { val: 30 }),
        })
    } else {
        None
    }
}

fn build_series_text(series: &ChartSeries) -> Option<SeriesText> {
    if series.name.is_empty() {
        None
    } else if series.name.contains('!') {
        Some(SeriesText {
            str_ref: Some(StrRef {
                f: series.name.clone(),
            }),
            v: None,
        })
    } else {
        Some(SeriesText {
            str_ref: None,
            v: Some(series.name.clone()),
        })
    }
}

fn build_series(index: u32, series: &ChartSeries) -> Series {
    let tx = build_series_text(series);
    let cat = if series.categories.is_empty() {
        None
    } else {
        Some(CategoryRef {
            str_ref: Some(StrRef {
                f: series.categories.clone(),
            }),
            num_ref: None,
        })
    };
    let val = if series.values.is_empty() {
        None
    } else {
        Some(ValueRef {
            num_ref: Some(NumRef {
                f: series.values.clone(),
            }),
        })
    };
    Series {
        idx: UintVal { val: index },
        order: UintVal { val: index },
        tx,
        cat,
        val,
    }
}

fn build_scatter_series(index: u32, series: &ChartSeries) -> ScatterSeries {
    let tx = build_series_text(series);
    let x_val = series
        .x_values
        .as_ref()
        .or(if series.categories.is_empty() {
            None
        } else {
            Some(&series.categories)
        })
        .map(|ref_str| CategoryRef {
            str_ref: None,
            num_ref: Some(NumRef { f: ref_str.clone() }),
        });
    let y_val = if series.values.is_empty() {
        None
    } else {
        Some(ValueRef {
            num_ref: Some(NumRef {
                f: series.values.clone(),
            }),
        })
    };
    ScatterSeries {
        idx: UintVal { val: index },
        order: UintVal { val: index },
        tx,
        x_val,
        y_val,
    }
}

fn build_bubble_series(index: u32, series: &ChartSeries) -> BubbleSeries {
    let tx = build_series_text(series);
    let x_val = series
        .x_values
        .as_ref()
        .or(if series.categories.is_empty() {
            None
        } else {
            Some(&series.categories)
        })
        .map(|ref_str| CategoryRef {
            str_ref: None,
            num_ref: Some(NumRef { f: ref_str.clone() }),
        });
    let y_val = if series.values.is_empty() {
        None
    } else {
        Some(ValueRef {
            num_ref: Some(NumRef {
                f: series.values.clone(),
            }),
        })
    };
    let bubble_size = series.bubble_sizes.as_ref().map(|ref_str| ValueRef {
        num_ref: Some(NumRef { f: ref_str.clone() }),
    });
    BubbleSeries {
        idx: UintVal { val: index },
        order: UintVal { val: index },
        tx,
        x_val,
        y_val,
        bubble_size,
    }
}

fn build_chart_title(text: &str) -> ChartTitle {
    ChartTitle {
        tx: TitleTx {
            rich: RichText {
                body_pr: BodyPr {},
                paragraphs: vec![Paragraph {
                    runs: vec![Run {
                        t: text.to_string(),
                    }],
                }],
            },
        },
    }
}

fn build_standard_axes() -> (Option<CatAx>, Option<ValAx>) {
    (
        Some(CatAx {
            ax_id: UintVal { val: 1 },
            scaling: Scaling {
                orientation: StringVal {
                    val: "minMax".to_string(),
                },
            },
            delete: BoolVal { val: false },
            ax_pos: StringVal {
                val: "b".to_string(),
            },
            cross_ax: UintVal { val: 2 },
        }),
        Some(ValAx {
            ax_id: UintVal { val: 2 },
            scaling: Scaling {
                orientation: StringVal {
                    val: "minMax".to_string(),
                },
            },
            delete: BoolVal { val: false },
            ax_pos: StringVal {
                val: "l".to_string(),
            },
            cross_ax: UintVal { val: 1 },
        }),
    )
}

fn build_ser_ax() -> SerAx {
    SerAx {
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
    }
}

fn standard_ax_ids() -> Vec<UintVal> {
    vec![UintVal { val: 1 }, UintVal { val: 2 }]
}

fn surface_ax_ids() -> Vec<UintVal> {
    vec![UintVal { val: 1 }, UintVal { val: 2 }, UintVal { val: 3 }]
}

fn build_plot_area(config: &ChartConfig) -> PlotArea {
    let ct = &config.chart_type;
    let no_axes = is_no_axis_chart(ct);
    let (cat_ax, val_ax) = if no_axes {
        (None, None)
    } else {
        build_standard_axes()
    };
    let ser_ax = if needs_ser_ax(ct) {
        Some(build_ser_ax())
    } else {
        None
    };
    let ax_ids = if no_axes {
        vec![]
    } else if needs_ser_ax(ct) {
        surface_ax_ids()
    } else {
        standard_ax_ids()
    };

    let xml_series: Vec<Series> = config
        .series
        .iter()
        .enumerate()
        .map(|(i, s)| build_series(i as u32, s))
        .collect();

    let mut plot_area = PlotArea {
        layout: Some(Layout {}),
        bar_chart: None,
        bar_3d_chart: None,
        line_chart: None,
        line_3d_chart: None,
        pie_chart: None,
        pie_3d_chart: None,
        doughnut_chart: None,
        area_chart: None,
        area_3d_chart: None,
        scatter_chart: None,
        bubble_chart: None,
        radar_chart: None,
        stock_chart: None,
        surface_chart: None,
        surface_3d_chart: None,
        cat_ax,
        val_ax,
        ser_ax,
    };

    match ct {
        ChartType::Col => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal { val: "col".into() },
                grouping: StringVal {
                    val: "clustered".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::ColStacked => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal { val: "col".into() },
                grouping: StringVal {
                    val: "stacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::ColPercentStacked => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal { val: "col".into() },
                grouping: StringVal {
                    val: "percentStacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Bar => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal { val: "bar".into() },
                grouping: StringVal {
                    val: "clustered".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::BarStacked => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal { val: "bar".into() },
                grouping: StringVal {
                    val: "stacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::BarPercentStacked => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal { val: "bar".into() },
                grouping: StringVal {
                    val: "percentStacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Line => {
            plot_area.line_chart = Some(LineChart {
                grouping: StringVal {
                    val: "standard".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::LineStacked => {
            plot_area.line_chart = Some(LineChart {
                grouping: StringVal {
                    val: "stacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::LinePercentStacked => {
            plot_area.line_chart = Some(LineChart {
                grouping: StringVal {
                    val: "percentStacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Line3D => {
            plot_area.line_3d_chart = Some(Line3DChart {
                grouping: StringVal {
                    val: "standard".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Pie => {
            plot_area.pie_chart = Some(PieChart { series: xml_series });
        }
        ChartType::Pie3D => {
            plot_area.pie_3d_chart = Some(Pie3DChart { series: xml_series });
        }
        ChartType::Doughnut => {
            plot_area.doughnut_chart = Some(DoughnutChart {
                series: xml_series,
                hole_size: Some(UintVal { val: 50 }),
            });
        }
        ChartType::Area => {
            plot_area.area_chart = Some(AreaChart {
                grouping: StringVal {
                    val: "standard".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::AreaStacked => {
            plot_area.area_chart = Some(AreaChart {
                grouping: StringVal {
                    val: "stacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::AreaPercentStacked => {
            plot_area.area_chart = Some(AreaChart {
                grouping: StringVal {
                    val: "percentStacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Area3D => {
            plot_area.area_3d_chart = Some(Area3DChart {
                grouping: StringVal {
                    val: "standard".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Area3DStacked => {
            plot_area.area_3d_chart = Some(Area3DChart {
                grouping: StringVal {
                    val: "stacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Area3DPercentStacked => {
            plot_area.area_3d_chart = Some(Area3DChart {
                grouping: StringVal {
                    val: "percentStacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Col3D => {
            plot_area.bar_3d_chart = Some(Bar3DChart {
                bar_dir: StringVal { val: "col".into() },
                grouping: StringVal {
                    val: "clustered".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Col3DStacked => {
            plot_area.bar_3d_chart = Some(Bar3DChart {
                bar_dir: StringVal { val: "col".into() },
                grouping: StringVal {
                    val: "stacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Col3DPercentStacked => {
            plot_area.bar_3d_chart = Some(Bar3DChart {
                bar_dir: StringVal { val: "col".into() },
                grouping: StringVal {
                    val: "percentStacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Bar3D => {
            plot_area.bar_3d_chart = Some(Bar3DChart {
                bar_dir: StringVal { val: "bar".into() },
                grouping: StringVal {
                    val: "clustered".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Bar3DStacked => {
            plot_area.bar_3d_chart = Some(Bar3DChart {
                bar_dir: StringVal { val: "bar".into() },
                grouping: StringVal {
                    val: "stacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Bar3DPercentStacked => {
            plot_area.bar_3d_chart = Some(Bar3DChart {
                bar_dir: StringVal { val: "bar".into() },
                grouping: StringVal {
                    val: "percentStacked".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Scatter => {
            let ss: Vec<ScatterSeries> = config
                .series
                .iter()
                .enumerate()
                .map(|(i, s)| build_scatter_series(i as u32, s))
                .collect();
            plot_area.scatter_chart = Some(ScatterChart {
                scatter_style: StringVal {
                    val: "lineMarker".into(),
                },
                series: ss,
                ax_ids,
            });
        }
        ChartType::ScatterLine => {
            let ss: Vec<ScatterSeries> = config
                .series
                .iter()
                .enumerate()
                .map(|(i, s)| build_scatter_series(i as u32, s))
                .collect();
            plot_area.scatter_chart = Some(ScatterChart {
                scatter_style: StringVal { val: "line".into() },
                series: ss,
                ax_ids,
            });
        }
        ChartType::ScatterSmooth => {
            let ss: Vec<ScatterSeries> = config
                .series
                .iter()
                .enumerate()
                .map(|(i, s)| build_scatter_series(i as u32, s))
                .collect();
            plot_area.scatter_chart = Some(ScatterChart {
                scatter_style: StringVal {
                    val: "smoothMarker".into(),
                },
                series: ss,
                ax_ids,
            });
        }
        ChartType::Bubble => {
            let bs: Vec<BubbleSeries> = config
                .series
                .iter()
                .enumerate()
                .map(|(i, s)| build_bubble_series(i as u32, s))
                .collect();
            plot_area.bubble_chart = Some(BubbleChart { series: bs, ax_ids });
        }
        ChartType::Radar => {
            plot_area.radar_chart = Some(RadarChart {
                radar_style: StringVal {
                    val: "standard".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::RadarFilled => {
            plot_area.radar_chart = Some(RadarChart {
                radar_style: StringVal {
                    val: "filled".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::RadarMarker => {
            plot_area.radar_chart = Some(RadarChart {
                radar_style: StringVal {
                    val: "marker".into(),
                },
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::StockHLC
        | ChartType::StockOHLC
        | ChartType::StockVHLC
        | ChartType::StockVOHLC => {
            plot_area.stock_chart = Some(StockChart {
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Surface => {
            plot_area.surface_chart = Some(SurfaceChart {
                wireframe: None,
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::SurfaceWireframe => {
            plot_area.surface_chart = Some(SurfaceChart {
                wireframe: Some(BoolVal { val: true }),
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::Surface3D => {
            plot_area.surface_3d_chart = Some(Surface3DChart {
                wireframe: None,
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::SurfaceWireframe3D => {
            plot_area.surface_3d_chart = Some(Surface3DChart {
                wireframe: Some(BoolVal { val: true }),
                series: xml_series,
                ax_ids,
            });
        }
        ChartType::ColLine | ChartType::ColLineStacked | ChartType::ColLinePercentStacked => {
            let grouping = match ct {
                ChartType::ColLineStacked => "stacked",
                ChartType::ColLinePercentStacked => "percentStacked",
                _ => "clustered",
            };
            let total = xml_series.len();
            let bar_count = total.div_ceil(2);
            let bar_series: Vec<Series> = xml_series.iter().take(bar_count).cloned().collect();
            let line_series: Vec<Series> = xml_series.iter().skip(bar_count).cloned().collect();
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal { val: "col".into() },
                grouping: StringVal {
                    val: grouping.into(),
                },
                series: bar_series,
                ax_ids: ax_ids.clone(),
            });
            plot_area.line_chart = Some(LineChart {
                grouping: StringVal {
                    val: "standard".into(),
                },
                series: line_series,
                ax_ids,
            });
        }
    }

    plot_area
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ss() -> Vec<ChartSeries> {
        vec![ChartSeries {
            name: "Revenue".into(),
            categories: "Sheet1!$A$2:$A$6".into(),
            values: "Sheet1!$B$2:$B$6".into(),
            x_values: None,
            bubble_sizes: None,
        }]
    }

    fn mc(chart_type: ChartType) -> ChartConfig {
        ChartConfig {
            chart_type,
            title: None,
            series: ss(),
            show_legend: false,
            view_3d: None,
        }
    }

    #[test]
    fn test_build_chart_xml_col() {
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: Some("Sales Chart".into()),
            series: ss(),
            show_legend: true,
            view_3d: None,
        };
        let cs = build_chart_xml(&config);
        assert!(cs.chart.title.is_some());
        assert!(cs.chart.legend.is_some());
        assert!(cs.chart.plot_area.bar_chart.is_some());
        assert!(cs.chart.plot_area.line_chart.is_none());
        assert!(cs.chart.plot_area.pie_chart.is_none());
        assert!(cs.chart.view_3d.is_none());
        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "col");
        assert_eq!(bar.grouping.val, "clustered");
        assert_eq!(bar.series.len(), 1);
        assert_eq!(bar.ax_ids.len(), 2);
    }

    #[test]
    fn test_build_chart_xml_bar() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::Bar,
            title: None,
            series: vec![],
            show_legend: false,
            view_3d: None,
        });
        assert!(cs.chart.title.is_none());
        assert!(cs.chart.legend.is_none());
        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "bar");
        assert_eq!(bar.grouping.val, "clustered");
    }

    #[test]
    fn test_bar_stacked() {
        let cs = build_chart_xml(&mc(ChartType::BarStacked));
        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "bar");
        assert_eq!(bar.grouping.val, "stacked");
    }

    #[test]
    fn test_col_stacked() {
        let cs = build_chart_xml(&mc(ChartType::ColStacked));
        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "col");
        assert_eq!(bar.grouping.val, "stacked");
    }

    #[test]
    fn test_col_percent_stacked() {
        let cs = build_chart_xml(&mc(ChartType::ColPercentStacked));
        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.grouping.val, "percentStacked");
    }

    #[test]
    fn test_bar_percent_stacked() {
        let cs = build_chart_xml(&mc(ChartType::BarPercentStacked));
        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "bar");
        assert_eq!(bar.grouping.val, "percentStacked");
    }

    #[test]
    fn test_line() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::Line,
            title: Some("Trend".into()),
            series: vec![ChartSeries {
                name: "Sheet1!$A$1".into(),
                categories: "Sheet1!$A$2:$A$6".into(),
                values: "Sheet1!$B$2:$B$6".into(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
        });
        assert!(cs.chart.plot_area.line_chart.is_some());
        let line = cs.chart.plot_area.line_chart.unwrap();
        assert_eq!(line.grouping.val, "standard");
        assert_eq!(line.series.len(), 1);
        let tx = line.series[0].tx.as_ref().unwrap();
        assert!(tx.str_ref.is_some());
        assert!(tx.v.is_none());
    }

    #[test]
    fn test_pie() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::Pie,
            title: Some("Distribution".into()),
            series: vec![ChartSeries {
                name: "Data".into(),
                categories: "Sheet1!$A$2:$A$6".into(),
                values: "Sheet1!$B$2:$B$6".into(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: true,
            view_3d: None,
        });
        assert!(cs.chart.plot_area.pie_chart.is_some());
        assert!(cs.chart.plot_area.cat_ax.is_none());
        assert!(cs.chart.plot_area.val_ax.is_none());
        let pie = cs.chart.plot_area.pie_chart.unwrap();
        let tx = pie.series[0].tx.as_ref().unwrap();
        assert!(tx.str_ref.is_none());
        assert_eq!(tx.v.as_deref(), Some("Data"));
    }

    #[test]
    fn test_no_legend() {
        let cs = build_chart_xml(&mc(ChartType::Col));
        assert!(cs.chart.legend.is_none());
    }

    #[test]
    fn test_axes_present_for_non_pie() {
        let cs = build_chart_xml(&mc(ChartType::Line));
        assert!(cs.chart.plot_area.cat_ax.is_some());
        assert!(cs.chart.plot_area.val_ax.is_some());
    }

    #[test]
    fn test_drawing_with_chart() {
        let from = MarkerType {
            col: 1,
            col_off: 0,
            row: 1,
            row_off: 0,
        };
        let to = MarkerType {
            col: 10,
            col_off: 0,
            row: 15,
            row_off: 0,
        };
        let dr = build_drawing_with_chart("rId1", from, to);
        assert_eq!(dr.two_cell_anchors.len(), 1);
        let anchor = &dr.two_cell_anchors[0];
        assert!(anchor.graphic_frame.is_some());
        assert_eq!(anchor.from.col, 1);
        assert_eq!(anchor.to.col, 10);
        let gf = anchor.graphic_frame.as_ref().unwrap();
        assert_eq!(gf.graphic.graphic_data.chart.r_id, "rId1");
    }

    #[test]
    fn test_series_literal_name() {
        let s = ChartSeries {
            name: "MyName".into(),
            categories: "Sheet1!$A$2:$A$6".into(),
            values: "Sheet1!$B$2:$B$6".into(),
            x_values: None,
            bubble_sizes: None,
        };
        let xs = build_series(0, &s);
        let tx = xs.tx.as_ref().unwrap();
        assert!(tx.str_ref.is_none());
        assert_eq!(tx.v.as_deref(), Some("MyName"));
    }

    #[test]
    fn test_series_cell_ref_name() {
        let s = ChartSeries {
            name: "Sheet1!$C$1".into(),
            categories: "".into(),
            values: "Sheet1!$B$2:$B$6".into(),
            x_values: None,
            bubble_sizes: None,
        };
        let xs = build_series(0, &s);
        let tx = xs.tx.as_ref().unwrap();
        assert!(tx.str_ref.is_some());
        assert!(tx.v.is_none());
        assert!(xs.cat.is_none());
    }

    #[test]
    fn test_series_empty_name() {
        let s = ChartSeries {
            name: "".into(),
            categories: "Sheet1!$A$2:$A$6".into(),
            values: "Sheet1!$B$2:$B$6".into(),
            x_values: None,
            bubble_sizes: None,
        };
        let xs = build_series(0, &s);
        assert!(xs.tx.is_none());
    }

    #[test]
    fn test_multiple_series() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::Col,
            title: None,
            series: vec![
                ChartSeries {
                    name: "A".into(),
                    categories: "Sheet1!$A$2:$A$6".into(),
                    values: "Sheet1!$B$2:$B$6".into(),
                    x_values: None,
                    bubble_sizes: None,
                },
                ChartSeries {
                    name: "B".into(),
                    categories: "Sheet1!$A$2:$A$6".into(),
                    values: "Sheet1!$C$2:$C$6".into(),
                    x_values: None,
                    bubble_sizes: None,
                },
            ],
            show_legend: true,
            view_3d: None,
        });
        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.series.len(), 2);
        assert_eq!(bar.series[0].idx.val, 0);
        assert_eq!(bar.series[1].idx.val, 1);
    }

    #[test]
    fn test_area_chart() {
        let cs = build_chart_xml(&mc(ChartType::Area));
        assert!(cs.chart.plot_area.area_chart.is_some());
        let a = cs.chart.plot_area.area_chart.unwrap();
        assert_eq!(a.grouping.val, "standard");
        assert_eq!(a.ax_ids.len(), 2);
        assert!(cs.chart.view_3d.is_none());
    }

    #[test]
    fn test_area_stacked() {
        let cs = build_chart_xml(&mc(ChartType::AreaStacked));
        assert_eq!(
            cs.chart.plot_area.area_chart.unwrap().grouping.val,
            "stacked"
        );
    }

    #[test]
    fn test_area_percent_stacked() {
        let cs = build_chart_xml(&mc(ChartType::AreaPercentStacked));
        assert_eq!(
            cs.chart.plot_area.area_chart.unwrap().grouping.val,
            "percentStacked"
        );
    }

    #[test]
    fn test_area_3d() {
        let cs = build_chart_xml(&mc(ChartType::Area3D));
        assert!(cs.chart.view_3d.is_some());
        assert_eq!(
            cs.chart.plot_area.area_3d_chart.unwrap().grouping.val,
            "standard"
        );
    }

    #[test]
    fn test_area_3d_stacked() {
        let cs = build_chart_xml(&mc(ChartType::Area3DStacked));
        assert!(cs.chart.view_3d.is_some());
        assert_eq!(
            cs.chart.plot_area.area_3d_chart.unwrap().grouping.val,
            "stacked"
        );
    }

    #[test]
    fn test_area_3d_percent_stacked() {
        let cs = build_chart_xml(&mc(ChartType::Area3DPercentStacked));
        assert!(cs.chart.view_3d.is_some());
        assert_eq!(
            cs.chart.plot_area.area_3d_chart.unwrap().grouping.val,
            "percentStacked"
        );
    }

    #[test]
    fn test_col_3d() {
        let cs = build_chart_xml(&mc(ChartType::Col3D));
        assert!(cs.chart.view_3d.is_some());
        let b = cs.chart.plot_area.bar_3d_chart.unwrap();
        assert_eq!(b.bar_dir.val, "col");
        assert_eq!(b.grouping.val, "clustered");
    }

    #[test]
    fn test_col_3d_stacked() {
        let cs = build_chart_xml(&mc(ChartType::Col3DStacked));
        assert_eq!(
            cs.chart.plot_area.bar_3d_chart.unwrap().grouping.val,
            "stacked"
        );
    }

    #[test]
    fn test_col_3d_percent_stacked() {
        let cs = build_chart_xml(&mc(ChartType::Col3DPercentStacked));
        assert_eq!(
            cs.chart.plot_area.bar_3d_chart.unwrap().grouping.val,
            "percentStacked"
        );
    }

    #[test]
    fn test_bar_3d() {
        let cs = build_chart_xml(&mc(ChartType::Bar3D));
        assert!(cs.chart.view_3d.is_some());
        let b = cs.chart.plot_area.bar_3d_chart.unwrap();
        assert_eq!(b.bar_dir.val, "bar");
        assert_eq!(b.grouping.val, "clustered");
    }

    #[test]
    fn test_bar_3d_stacked() {
        let cs = build_chart_xml(&mc(ChartType::Bar3DStacked));
        assert_eq!(
            cs.chart.plot_area.bar_3d_chart.unwrap().grouping.val,
            "stacked"
        );
    }

    #[test]
    fn test_bar_3d_percent_stacked() {
        let cs = build_chart_xml(&mc(ChartType::Bar3DPercentStacked));
        assert_eq!(
            cs.chart.plot_area.bar_3d_chart.unwrap().grouping.val,
            "percentStacked"
        );
    }

    #[test]
    fn test_line_stacked() {
        let cs = build_chart_xml(&mc(ChartType::LineStacked));
        assert_eq!(
            cs.chart.plot_area.line_chart.unwrap().grouping.val,
            "stacked"
        );
    }

    #[test]
    fn test_line_percent_stacked() {
        let cs = build_chart_xml(&mc(ChartType::LinePercentStacked));
        assert_eq!(
            cs.chart.plot_area.line_chart.unwrap().grouping.val,
            "percentStacked"
        );
    }

    #[test]
    fn test_line_3d() {
        let cs = build_chart_xml(&mc(ChartType::Line3D));
        assert!(cs.chart.view_3d.is_some());
        assert!(cs.chart.plot_area.line_3d_chart.is_some());
    }

    #[test]
    fn test_pie_3d() {
        let cs = build_chart_xml(&mc(ChartType::Pie3D));
        assert!(cs.chart.view_3d.is_some());
        assert!(cs.chart.plot_area.pie_3d_chart.is_some());
        assert!(cs.chart.plot_area.cat_ax.is_none());
    }

    #[test]
    fn test_doughnut() {
        let cs = build_chart_xml(&mc(ChartType::Doughnut));
        assert!(cs.chart.plot_area.doughnut_chart.is_some());
        assert!(cs.chart.plot_area.cat_ax.is_none());
        let d = cs.chart.plot_area.doughnut_chart.unwrap();
        assert_eq!(d.hole_size.as_ref().unwrap().val, 50);
    }

    #[test]
    fn test_scatter() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::Scatter,
            title: None,
            series: vec![ChartSeries {
                name: "XY".into(),
                categories: "Sheet1!$A$2:$A$6".into(),
                values: "Sheet1!$B$2:$B$6".into(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        });
        let sc = cs.chart.plot_area.scatter_chart.unwrap();
        assert_eq!(sc.scatter_style.val, "lineMarker");
        assert_eq!(sc.series.len(), 1);
        let s = &sc.series[0];
        assert_eq!(
            s.x_val.as_ref().unwrap().num_ref.as_ref().unwrap().f,
            "Sheet1!$A$2:$A$6"
        );
        assert_eq!(
            s.y_val.as_ref().unwrap().num_ref.as_ref().unwrap().f,
            "Sheet1!$B$2:$B$6"
        );
    }

    #[test]
    fn test_scatter_explicit_x() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::Scatter,
            title: None,
            series: vec![ChartSeries {
                name: "".into(),
                categories: "Sheet1!$A$2:$A$6".into(),
                values: "Sheet1!$B$2:$B$6".into(),
                x_values: Some("Sheet1!$D$2:$D$6".into()),
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        });
        let s = &cs.chart.plot_area.scatter_chart.unwrap().series[0];
        assert_eq!(
            s.x_val.as_ref().unwrap().num_ref.as_ref().unwrap().f,
            "Sheet1!$D$2:$D$6"
        );
    }

    #[test]
    fn test_scatter_line() {
        let cs = build_chart_xml(&mc(ChartType::ScatterLine));
        assert_eq!(
            cs.chart.plot_area.scatter_chart.unwrap().scatter_style.val,
            "line"
        );
    }

    #[test]
    fn test_scatter_smooth() {
        let cs = build_chart_xml(&mc(ChartType::ScatterSmooth));
        assert_eq!(
            cs.chart.plot_area.scatter_chart.unwrap().scatter_style.val,
            "smoothMarker"
        );
    }

    #[test]
    fn test_bubble() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::Bubble,
            title: None,
            series: vec![ChartSeries {
                name: "B".into(),
                categories: "Sheet1!$A$2:$A$6".into(),
                values: "Sheet1!$B$2:$B$6".into(),
                x_values: None,
                bubble_sizes: Some("Sheet1!$C$2:$C$6".into()),
            }],
            show_legend: false,
            view_3d: None,
        });
        let b = cs.chart.plot_area.bubble_chart.unwrap();
        assert_eq!(b.series.len(), 1);
        assert_eq!(
            b.series[0]
                .bubble_size
                .as_ref()
                .unwrap()
                .num_ref
                .as_ref()
                .unwrap()
                .f,
            "Sheet1!$C$2:$C$6"
        );
    }

    #[test]
    fn test_radar() {
        let cs = build_chart_xml(&mc(ChartType::Radar));
        assert_eq!(
            cs.chart.plot_area.radar_chart.unwrap().radar_style.val,
            "standard"
        );
    }

    #[test]
    fn test_radar_filled() {
        let cs = build_chart_xml(&mc(ChartType::RadarFilled));
        assert_eq!(
            cs.chart.plot_area.radar_chart.unwrap().radar_style.val,
            "filled"
        );
    }

    #[test]
    fn test_radar_marker() {
        let cs = build_chart_xml(&mc(ChartType::RadarMarker));
        assert_eq!(
            cs.chart.plot_area.radar_chart.unwrap().radar_style.val,
            "marker"
        );
    }

    #[test]
    fn test_stock_hlc() {
        let cs = build_chart_xml(&mc(ChartType::StockHLC));
        assert!(cs.chart.plot_area.stock_chart.is_some());
    }

    #[test]
    fn test_stock_ohlc() {
        let cs = build_chart_xml(&mc(ChartType::StockOHLC));
        assert!(cs.chart.plot_area.stock_chart.is_some());
    }

    #[test]
    fn test_stock_vhlc() {
        let cs = build_chart_xml(&mc(ChartType::StockVHLC));
        assert!(cs.chart.plot_area.stock_chart.is_some());
    }

    #[test]
    fn test_stock_vohlc() {
        let cs = build_chart_xml(&mc(ChartType::StockVOHLC));
        assert!(cs.chart.plot_area.stock_chart.is_some());
    }

    #[test]
    fn test_surface() {
        let cs = build_chart_xml(&mc(ChartType::Surface));
        assert!(cs.chart.plot_area.surface_chart.is_some());
        assert!(cs.chart.plot_area.ser_ax.is_some());
        let sf = cs.chart.plot_area.surface_chart.unwrap();
        assert!(sf.wireframe.is_none());
        assert_eq!(sf.ax_ids.len(), 3);
    }

    #[test]
    fn test_surface_wireframe() {
        let cs = build_chart_xml(&mc(ChartType::SurfaceWireframe));
        let sf = cs.chart.plot_area.surface_chart.unwrap();
        assert!(sf.wireframe.as_ref().unwrap().val);
        assert_eq!(sf.ax_ids.len(), 3);
    }

    #[test]
    fn test_surface_3d() {
        let cs = build_chart_xml(&mc(ChartType::Surface3D));
        assert!(cs.chart.view_3d.is_some());
        assert!(cs.chart.plot_area.surface_3d_chart.is_some());
        assert!(cs.chart.plot_area.ser_ax.is_some());
    }

    #[test]
    fn test_surface_wireframe_3d() {
        let cs = build_chart_xml(&mc(ChartType::SurfaceWireframe3D));
        assert!(cs.chart.view_3d.is_some());
        let sf = cs.chart.plot_area.surface_3d_chart.unwrap();
        assert!(sf.wireframe.as_ref().unwrap().val);
    }

    #[test]
    fn test_col_line_combo() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::ColLine,
            title: None,
            series: vec![
                ChartSeries {
                    name: "A".into(),
                    categories: "Sheet1!$A$2:$A$6".into(),
                    values: "Sheet1!$B$2:$B$6".into(),
                    x_values: None,
                    bubble_sizes: None,
                },
                ChartSeries {
                    name: "B".into(),
                    categories: "Sheet1!$A$2:$A$6".into(),
                    values: "Sheet1!$C$2:$C$6".into(),
                    x_values: None,
                    bubble_sizes: None,
                },
            ],
            show_legend: true,
            view_3d: None,
        });
        assert!(cs.chart.plot_area.bar_chart.is_some());
        assert!(cs.chart.plot_area.line_chart.is_some());
        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.grouping.val, "clustered");
        assert_eq!(bar.series.len(), 1);
        assert_eq!(cs.chart.plot_area.line_chart.unwrap().series.len(), 1);
    }

    #[test]
    fn test_col_line_stacked_combo() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::ColLineStacked,
            title: None,
            series: vec![
                ChartSeries {
                    name: "A".into(),
                    categories: "".into(),
                    values: "Sheet1!$B$2:$B$6".into(),
                    x_values: None,
                    bubble_sizes: None,
                },
                ChartSeries {
                    name: "B".into(),
                    categories: "".into(),
                    values: "Sheet1!$C$2:$C$6".into(),
                    x_values: None,
                    bubble_sizes: None,
                },
                ChartSeries {
                    name: "C".into(),
                    categories: "".into(),
                    values: "Sheet1!$D$2:$D$6".into(),
                    x_values: None,
                    bubble_sizes: None,
                },
            ],
            show_legend: false,
            view_3d: None,
        });
        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.grouping.val, "stacked");
        assert_eq!(bar.series.len(), 2);
        assert_eq!(cs.chart.plot_area.line_chart.unwrap().series.len(), 1);
    }

    #[test]
    fn test_col_line_percent_stacked_combo() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::ColLinePercentStacked,
            title: None,
            series: vec![ChartSeries {
                name: "A".into(),
                categories: "".into(),
                values: "Sheet1!$B$2:$B$6".into(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        });
        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.grouping.val, "percentStacked");
        assert_eq!(bar.series.len(), 1);
        assert_eq!(cs.chart.plot_area.line_chart.unwrap().series.len(), 0);
    }

    #[test]
    fn test_view_3d_explicit() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::Col3D,
            title: None,
            series: vec![],
            show_legend: false,
            view_3d: Some(View3DConfig {
                rot_x: Some(30),
                rot_y: Some(40),
                depth_percent: Some(200),
                right_angle_axes: Some(false),
                perspective: Some(10),
            }),
        });
        let v = cs.chart.view_3d.unwrap();
        assert_eq!(v.rot_x.unwrap().val, 30);
        assert_eq!(v.rot_y.unwrap().val, 40);
        assert_eq!(v.depth_percent.unwrap().val, 200);
        assert!(!v.r_ang_ax.unwrap().val);
        assert_eq!(v.perspective.unwrap().val, 10);
    }

    #[test]
    fn test_view_3d_auto_defaults() {
        let cs = build_chart_xml(&mc(ChartType::Col3D));
        let v = cs.chart.view_3d.unwrap();
        assert_eq!(v.rot_x.unwrap().val, 15);
        assert_eq!(v.rot_y.unwrap().val, 20);
        assert!(v.r_ang_ax.unwrap().val);
        assert_eq!(v.perspective.unwrap().val, 30);
    }

    #[test]
    fn test_non_3d_no_view() {
        let cs = build_chart_xml(&mc(ChartType::Col));
        assert!(cs.chart.view_3d.is_none());
    }

    #[test]
    fn test_chart_type_enum_coverage() {
        let types = [
            ChartType::Col,
            ChartType::ColStacked,
            ChartType::ColPercentStacked,
            ChartType::Bar,
            ChartType::BarStacked,
            ChartType::BarPercentStacked,
            ChartType::Line,
            ChartType::Pie,
            ChartType::Area,
            ChartType::AreaStacked,
            ChartType::AreaPercentStacked,
            ChartType::Area3D,
            ChartType::Area3DStacked,
            ChartType::Area3DPercentStacked,
            ChartType::Col3D,
            ChartType::Col3DStacked,
            ChartType::Col3DPercentStacked,
            ChartType::Bar3D,
            ChartType::Bar3DStacked,
            ChartType::Bar3DPercentStacked,
            ChartType::LineStacked,
            ChartType::LinePercentStacked,
            ChartType::Line3D,
            ChartType::Pie3D,
            ChartType::Doughnut,
            ChartType::Scatter,
            ChartType::ScatterLine,
            ChartType::ScatterSmooth,
            ChartType::Radar,
            ChartType::RadarFilled,
            ChartType::RadarMarker,
            ChartType::StockHLC,
            ChartType::StockOHLC,
            ChartType::StockVHLC,
            ChartType::StockVOHLC,
            ChartType::Bubble,
            ChartType::Surface,
            ChartType::Surface3D,
            ChartType::SurfaceWireframe,
            ChartType::SurfaceWireframe3D,
            ChartType::ColLine,
            ChartType::ColLineStacked,
            ChartType::ColLinePercentStacked,
        ];
        for ct in &types {
            let _ = build_chart_xml(&ChartConfig {
                chart_type: ct.clone(),
                title: None,
                series: vec![],
                show_legend: false,
                view_3d: None,
            });
        }
    }

    #[test]
    fn test_scatter_empty_categories() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::Scatter,
            title: None,
            series: vec![ChartSeries {
                name: "".into(),
                categories: "".into(),
                values: "Sheet1!$B$2:$B$6".into(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        });
        let s = &cs.chart.plot_area.scatter_chart.unwrap().series[0];
        assert!(s.x_val.is_none());
        assert!(s.y_val.is_some());
    }

    #[test]
    fn test_bubble_no_sizes() {
        let cs = build_chart_xml(&ChartConfig {
            chart_type: ChartType::Bubble,
            title: None,
            series: vec![ChartSeries {
                name: "".into(),
                categories: "Sheet1!$A$2:$A$6".into(),
                values: "Sheet1!$B$2:$B$6".into(),
                x_values: None,
                bubble_sizes: None,
            }],
            show_legend: false,
            view_3d: None,
        });
        assert!(cs.chart.plot_area.bubble_chart.unwrap().series[0]
            .bubble_size
            .is_none());
    }
}
