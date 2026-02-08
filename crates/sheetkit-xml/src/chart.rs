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

    #[serde(rename = "c:lineChart", skip_serializing_if = "Option::is_none")]
    pub line_chart: Option<LineChart>,

    #[serde(rename = "c:pieChart", skip_serializing_if = "Option::is_none")]
    pub pie_chart: Option<PieChart>,

    #[serde(rename = "c:catAx", skip_serializing_if = "Option::is_none")]
    pub cat_ax: Option<CatAx>,

    #[serde(rename = "c:valAx", skip_serializing_if = "Option::is_none")]
    pub val_ax: Option<ValAx>,
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

/// Pie chart definition.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PieChart {
    #[serde(rename = "c:ser", default)]
    pub series: Vec<Series>,
}

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
}
