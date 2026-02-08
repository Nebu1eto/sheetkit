//! Chart builder and management.
//!
//! Provides types for configuring charts and functions for building the
//! corresponding XML structures (ChartSpace and drawing anchors).

use sheetkit_xml::chart::{
    BarChart, BodyPr, BoolVal, CatAx, CategoryRef, Chart, ChartSpace, ChartTitle, Layout, Legend,
    LineChart, NumRef, Paragraph, PieChart, PlotArea, RichText, Run, Scaling, Series, SeriesText,
    StrRef, StringVal, TitleTx, UintVal, ValAx, ValueRef,
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
}

/// Build a `ChartSpace` XML structure from a chart configuration.
pub fn build_chart_xml(config: &ChartConfig) -> ChartSpace {
    let xml_series: Vec<Series> = config
        .series
        .iter()
        .enumerate()
        .map(|(i, s)| build_series(i as u32, s))
        .collect();

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

    let is_pie = matches!(config.chart_type, ChartType::Pie);

    let plot_area = build_plot_area(&config.chart_type, xml_series, is_pie);

    ChartSpace {
        chart: Chart {
            title,
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

/// Build a single XML Series from a ChartSeries.
fn build_series(index: u32, series: &ChartSeries) -> Series {
    let tx = if series.name.is_empty() {
        None
    } else if series.name.contains('!') {
        // Looks like a cell reference
        Some(SeriesText {
            str_ref: Some(StrRef {
                f: series.name.clone(),
            }),
            v: None,
        })
    } else {
        // Literal string value
        Some(SeriesText {
            str_ref: None,
            v: Some(series.name.clone()),
        })
    };

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

/// Build a ChartTitle from a string.
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

/// Build a PlotArea from a chart type and series.
fn build_plot_area(chart_type: &ChartType, series: Vec<Series>, is_pie: bool) -> PlotArea {
    let (cat_ax, val_ax) = if is_pie {
        (None, None)
    } else {
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
    };

    let ax_ids = if is_pie {
        vec![]
    } else {
        vec![UintVal { val: 1 }, UintVal { val: 2 }]
    };

    let mut plot_area = PlotArea {
        layout: Some(Layout {}),
        bar_chart: None,
        line_chart: None,
        pie_chart: None,
        cat_ax,
        val_ax,
    };

    match chart_type {
        ChartType::Col => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal {
                    val: "col".to_string(),
                },
                grouping: StringVal {
                    val: "clustered".to_string(),
                },
                series,
                ax_ids,
            });
        }
        ChartType::ColStacked => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal {
                    val: "col".to_string(),
                },
                grouping: StringVal {
                    val: "stacked".to_string(),
                },
                series,
                ax_ids,
            });
        }
        ChartType::ColPercentStacked => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal {
                    val: "col".to_string(),
                },
                grouping: StringVal {
                    val: "percentStacked".to_string(),
                },
                series,
                ax_ids,
            });
        }
        ChartType::Bar => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal {
                    val: "bar".to_string(),
                },
                grouping: StringVal {
                    val: "clustered".to_string(),
                },
                series,
                ax_ids,
            });
        }
        ChartType::BarStacked => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal {
                    val: "bar".to_string(),
                },
                grouping: StringVal {
                    val: "stacked".to_string(),
                },
                series,
                ax_ids,
            });
        }
        ChartType::BarPercentStacked => {
            plot_area.bar_chart = Some(BarChart {
                bar_dir: StringVal {
                    val: "bar".to_string(),
                },
                grouping: StringVal {
                    val: "percentStacked".to_string(),
                },
                series,
                ax_ids,
            });
        }
        ChartType::Line => {
            plot_area.line_chart = Some(LineChart {
                grouping: StringVal {
                    val: "standard".to_string(),
                },
                series,
                ax_ids,
            });
        }
        ChartType::Pie => {
            plot_area.pie_chart = Some(PieChart { series });
        }
    }

    plot_area
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_config() -> ChartConfig {
        ChartConfig {
            chart_type: ChartType::Col,
            title: Some("Sales Chart".to_string()),
            series: vec![ChartSeries {
                name: "Revenue".to_string(),
                categories: "Sheet1!$A$2:$A$6".to_string(),
                values: "Sheet1!$B$2:$B$6".to_string(),
            }],
            show_legend: true,
        }
    }

    #[test]
    fn test_build_chart_xml_col() {
        let config = sample_config();
        let cs = build_chart_xml(&config);

        assert!(cs.chart.title.is_some());
        assert!(cs.chart.legend.is_some());
        assert!(cs.chart.plot_area.bar_chart.is_some());
        assert!(cs.chart.plot_area.line_chart.is_none());
        assert!(cs.chart.plot_area.pie_chart.is_none());

        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "col");
        assert_eq!(bar.grouping.val, "clustered");
        assert_eq!(bar.series.len(), 1);
        assert_eq!(bar.ax_ids.len(), 2);
    }

    #[test]
    fn test_build_chart_xml_bar() {
        let config = ChartConfig {
            chart_type: ChartType::Bar,
            title: None,
            series: vec![],
            show_legend: false,
        };
        let cs = build_chart_xml(&config);

        assert!(cs.chart.title.is_none());
        assert!(cs.chart.legend.is_none());

        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "bar");
        assert_eq!(bar.grouping.val, "clustered");
    }

    #[test]
    fn test_build_chart_xml_bar_stacked() {
        let config = ChartConfig {
            chart_type: ChartType::BarStacked,
            title: None,
            series: vec![],
            show_legend: false,
        };
        let cs = build_chart_xml(&config);

        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "bar");
        assert_eq!(bar.grouping.val, "stacked");
    }

    #[test]
    fn test_build_chart_xml_col_stacked() {
        let config = ChartConfig {
            chart_type: ChartType::ColStacked,
            title: None,
            series: vec![],
            show_legend: false,
        };
        let cs = build_chart_xml(&config);

        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "col");
        assert_eq!(bar.grouping.val, "stacked");
    }

    #[test]
    fn test_build_chart_xml_col_percent_stacked() {
        let config = ChartConfig {
            chart_type: ChartType::ColPercentStacked,
            title: None,
            series: vec![],
            show_legend: false,
        };
        let cs = build_chart_xml(&config);

        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "col");
        assert_eq!(bar.grouping.val, "percentStacked");
    }

    #[test]
    fn test_build_chart_xml_bar_percent_stacked() {
        let config = ChartConfig {
            chart_type: ChartType::BarPercentStacked,
            title: None,
            series: vec![],
            show_legend: false,
        };
        let cs = build_chart_xml(&config);

        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.bar_dir.val, "bar");
        assert_eq!(bar.grouping.val, "percentStacked");
    }

    #[test]
    fn test_build_chart_xml_line() {
        let config = ChartConfig {
            chart_type: ChartType::Line,
            title: Some("Trend".to_string()),
            series: vec![ChartSeries {
                name: "Sheet1!$A$1".to_string(),
                categories: "Sheet1!$A$2:$A$6".to_string(),
                values: "Sheet1!$B$2:$B$6".to_string(),
            }],
            show_legend: true,
        };
        let cs = build_chart_xml(&config);

        assert!(cs.chart.plot_area.line_chart.is_some());
        assert!(cs.chart.plot_area.bar_chart.is_none());
        assert!(cs.chart.plot_area.pie_chart.is_none());

        let line = cs.chart.plot_area.line_chart.unwrap();
        assert_eq!(line.grouping.val, "standard");
        assert_eq!(line.series.len(), 1);

        // The series name is a cell reference, so it should use str_ref
        let tx = line.series[0].tx.as_ref().unwrap();
        assert!(tx.str_ref.is_some());
        assert!(tx.v.is_none());
    }

    #[test]
    fn test_build_chart_xml_pie() {
        let config = ChartConfig {
            chart_type: ChartType::Pie,
            title: Some("Distribution".to_string()),
            series: vec![ChartSeries {
                name: "Data".to_string(),
                categories: "Sheet1!$A$2:$A$6".to_string(),
                values: "Sheet1!$B$2:$B$6".to_string(),
            }],
            show_legend: true,
        };
        let cs = build_chart_xml(&config);

        assert!(cs.chart.plot_area.pie_chart.is_some());
        assert!(cs.chart.plot_area.bar_chart.is_none());
        assert!(cs.chart.plot_area.line_chart.is_none());
        // Pie charts have no axes
        assert!(cs.chart.plot_area.cat_ax.is_none());
        assert!(cs.chart.plot_area.val_ax.is_none());

        let pie = cs.chart.plot_area.pie_chart.unwrap();
        assert_eq!(pie.series.len(), 1);

        // Literal name should use v, not str_ref
        let tx = pie.series[0].tx.as_ref().unwrap();
        assert!(tx.str_ref.is_none());
        assert_eq!(tx.v.as_deref(), Some("Data"));
    }

    #[test]
    fn test_build_chart_xml_no_legend() {
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: None,
            series: vec![],
            show_legend: false,
        };
        let cs = build_chart_xml(&config);

        assert!(cs.chart.legend.is_none());
    }

    #[test]
    fn test_build_chart_xml_axes_present_for_non_pie() {
        let config = ChartConfig {
            chart_type: ChartType::Line,
            title: None,
            series: vec![],
            show_legend: false,
        };
        let cs = build_chart_xml(&config);

        assert!(cs.chart.plot_area.cat_ax.is_some());
        assert!(cs.chart.plot_area.val_ax.is_some());
    }

    #[test]
    fn test_build_drawing_with_chart() {
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
        assert!(anchor.pic.is_none());
        assert_eq!(anchor.from.col, 1);
        assert_eq!(anchor.to.col, 10);

        let gf = anchor.graphic_frame.as_ref().unwrap();
        assert_eq!(gf.graphic.graphic_data.chart.r_id, "rId1");
    }

    #[test]
    fn test_build_series_literal_name() {
        let series = ChartSeries {
            name: "MyName".to_string(),
            categories: "Sheet1!$A$2:$A$6".to_string(),
            values: "Sheet1!$B$2:$B$6".to_string(),
        };
        let xml_series = build_series(0, &series);

        let tx = xml_series.tx.as_ref().unwrap();
        assert!(tx.str_ref.is_none());
        assert_eq!(tx.v.as_deref(), Some("MyName"));
    }

    #[test]
    fn test_build_series_cell_ref_name() {
        let series = ChartSeries {
            name: "Sheet1!$C$1".to_string(),
            categories: "".to_string(),
            values: "Sheet1!$B$2:$B$6".to_string(),
        };
        let xml_series = build_series(0, &series);

        let tx = xml_series.tx.as_ref().unwrap();
        assert!(tx.str_ref.is_some());
        assert!(tx.v.is_none());
        assert_eq!(tx.str_ref.as_ref().unwrap().f, "Sheet1!$C$1");

        // Empty categories should be None
        assert!(xml_series.cat.is_none());
    }

    #[test]
    fn test_build_series_empty_name() {
        let series = ChartSeries {
            name: "".to_string(),
            categories: "Sheet1!$A$2:$A$6".to_string(),
            values: "Sheet1!$B$2:$B$6".to_string(),
        };
        let xml_series = build_series(0, &series);

        // Empty name should result in no tx
        assert!(xml_series.tx.is_none());
    }

    #[test]
    fn test_build_chart_xml_multiple_series() {
        let config = ChartConfig {
            chart_type: ChartType::Col,
            title: None,
            series: vec![
                ChartSeries {
                    name: "Series A".to_string(),
                    categories: "Sheet1!$A$2:$A$6".to_string(),
                    values: "Sheet1!$B$2:$B$6".to_string(),
                },
                ChartSeries {
                    name: "Series B".to_string(),
                    categories: "Sheet1!$A$2:$A$6".to_string(),
                    values: "Sheet1!$C$2:$C$6".to_string(),
                },
            ],
            show_legend: true,
        };
        let cs = build_chart_xml(&config);

        let bar = cs.chart.plot_area.bar_chart.unwrap();
        assert_eq!(bar.series.len(), 2);
        assert_eq!(bar.series[0].idx.val, 0);
        assert_eq!(bar.series[0].order.val, 0);
        assert_eq!(bar.series[1].idx.val, 1);
        assert_eq!(bar.series[1].order.val, 1);
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
        ];

        for ct in &types {
            let config = ChartConfig {
                chart_type: ct.clone(),
                title: None,
                series: vec![],
                show_legend: false,
            };
            // Should not panic
            let _cs = build_chart_xml(&config);
        }
    }
}
