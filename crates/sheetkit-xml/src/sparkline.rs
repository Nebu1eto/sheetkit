//! Sparkline XML schema structures.
//!
//! Sparklines are stored in worksheet extension lists (extLst) using the
//! x14 namespace (http://schemas.microsoft.com/office/spreadsheetml/2009/9/main).

use serde::{Deserialize, Serialize};

/// A group of sparklines with shared settings.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "x14:sparklineGroup")]
pub struct SparklineGroup {
    /// Sparkline type: "line", "column", or "stacked" (win/loss).
    #[serde(rename = "@type", skip_serializing_if = "Option::is_none")]
    pub sparkline_type: Option<String>,

    /// Whether to display markers on data points.
    #[serde(rename = "@markers", skip_serializing_if = "Option::is_none")]
    pub markers: Option<bool>,

    /// Whether to highlight the high point.
    #[serde(rename = "@high", skip_serializing_if = "Option::is_none")]
    pub high: Option<bool>,

    /// Whether to highlight the low point.
    #[serde(rename = "@low", skip_serializing_if = "Option::is_none")]
    pub low: Option<bool>,

    /// Whether to highlight the first point.
    #[serde(rename = "@first", skip_serializing_if = "Option::is_none")]
    pub first: Option<bool>,

    /// Whether to highlight the last point.
    #[serde(rename = "@last", skip_serializing_if = "Option::is_none")]
    pub last: Option<bool>,

    /// Whether to highlight negative points.
    #[serde(rename = "@negative", skip_serializing_if = "Option::is_none")]
    pub negative: Option<bool>,

    /// Whether to display an axis.
    #[serde(rename = "@displayXAxis", skip_serializing_if = "Option::is_none")]
    pub display_x_axis: Option<bool>,

    /// Line weight in points.
    #[serde(rename = "@lineWeight", skip_serializing_if = "Option::is_none")]
    pub line_weight: Option<f64>,

    /// Minimum axis type: "individual" or "group".
    #[serde(rename = "@minAxisType", skip_serializing_if = "Option::is_none")]
    pub min_axis_type: Option<String>,

    /// Maximum axis type: "individual" or "group".
    #[serde(rename = "@maxAxisType", skip_serializing_if = "Option::is_none")]
    pub max_axis_type: Option<String>,

    /// Individual sparklines in this group.
    #[serde(rename = "x14:sparklines")]
    pub sparklines: SparklineList,
}

/// Container for individual sparklines.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SparklineList {
    #[serde(rename = "x14:sparkline", default)]
    pub items: Vec<Sparkline>,
}

/// A single sparkline mapping data to a cell location.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Sparkline {
    /// Data source range (e.g., "Sheet1!A1:A10").
    #[serde(rename = "xm:f")]
    pub formula: String,

    /// Cell location where the sparkline is rendered (e.g., "B1").
    #[serde(rename = "xm:sqref")]
    pub sqref: String,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sparkline_group_default() {
        let group = SparklineGroup {
            sparkline_type: Some("line".to_string()),
            markers: None,
            high: None,
            low: None,
            first: None,
            last: None,
            negative: None,
            display_x_axis: None,
            line_weight: None,
            min_axis_type: None,
            max_axis_type: None,
            sparklines: SparklineList {
                items: vec![Sparkline {
                    formula: "Sheet1!A1:A10".to_string(),
                    sqref: "B1".to_string(),
                }],
            },
        };
        assert_eq!(group.sparklines.items.len(), 1);
    }

    #[test]
    fn test_sparkline_with_options() {
        let group = SparklineGroup {
            sparkline_type: Some("column".to_string()),
            markers: Some(true),
            high: Some(true),
            low: Some(true),
            first: None,
            last: None,
            negative: Some(true),
            display_x_axis: Some(true),
            line_weight: Some(0.75),
            min_axis_type: Some("group".to_string()),
            max_axis_type: Some("group".to_string()),
            sparklines: SparklineList { items: vec![] },
        };
        assert_eq!(group.sparkline_type, Some("column".to_string()));
        assert_eq!(group.markers, Some(true));
    }
}
