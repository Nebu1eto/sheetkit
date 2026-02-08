//! Sparkline creation and management.

use crate::error::{Error, Result};

/// Sparkline type.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum SparklineType {
    /// Line sparkline (default).
    #[default]
    Line,
    /// Column (bar) sparkline.
    Column,
    /// Win/Loss sparkline.
    WinLoss,
}

impl SparklineType {
    /// Return the OOXML string representation.
    pub fn as_str(&self) -> &str {
        match self {
            SparklineType::Line => "line",
            SparklineType::Column => "column",
            SparklineType::WinLoss => "stacked",
        }
    }

    /// Parse from OOXML string representation.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "line" => Some(SparklineType::Line),
            "column" => Some(SparklineType::Column),
            "stacked" | "winloss" => Some(SparklineType::WinLoss),
            _ => None,
        }
    }
}

/// Configuration for creating a sparkline group.
#[derive(Debug, Clone)]
pub struct SparklineConfig {
    /// Data source range (e.g., "Sheet1!A1:A10").
    pub data_range: String,
    /// Cell where the sparkline is rendered (e.g., "B1").
    pub location: String,
    /// Sparkline type.
    pub sparkline_type: SparklineType,
    /// Show data markers.
    pub markers: bool,
    /// Highlight high point.
    pub high_point: bool,
    /// Highlight low point.
    pub low_point: bool,
    /// Highlight first point.
    pub first_point: bool,
    /// Highlight last point.
    pub last_point: bool,
    /// Highlight negative values.
    pub negative_points: bool,
    /// Show horizontal axis.
    pub show_axis: bool,
    /// Line weight in points (for line sparklines).
    pub line_weight: Option<f64>,
    /// Style preset index (0-35).
    pub style: Option<u32>,
}

impl SparklineConfig {
    /// Create a simple sparkline config.
    pub fn new(data_range: &str, location: &str) -> Self {
        Self {
            data_range: data_range.to_string(),
            location: location.to_string(),
            sparkline_type: SparklineType::Line,
            markers: false,
            high_point: false,
            low_point: false,
            first_point: false,
            last_point: false,
            negative_points: false,
            show_axis: false,
            line_weight: None,
            style: None,
        }
    }
}

/// A sparkline group configuration for adding to worksheets.
#[derive(Debug, Clone)]
pub struct SparklineGroupConfig {
    /// Sparklines in this group (all share the same settings).
    pub sparklines: Vec<SparklineConfig>,
    /// Sparkline type for the group.
    pub sparkline_type: SparklineType,
    /// Show data markers.
    pub markers: bool,
    /// Highlight high point.
    pub high_point: bool,
    /// Highlight low point.
    pub low_point: bool,
    /// Line weight in points.
    pub line_weight: Option<f64>,
}

/// Convert a SparklineConfig to the XML representation.
pub(crate) fn config_to_xml_group(
    config: &SparklineConfig,
) -> sheetkit_xml::sparkline::SparklineGroup {
    use sheetkit_xml::sparkline::*;

    SparklineGroup {
        sparkline_type: match config.sparkline_type {
            SparklineType::Line => None,
            _ => Some(config.sparkline_type.as_str().to_string()),
        },
        markers: if config.markers { Some(true) } else { None },
        high: if config.high_point { Some(true) } else { None },
        low: if config.low_point { Some(true) } else { None },
        first: if config.first_point { Some(true) } else { None },
        last: if config.last_point { Some(true) } else { None },
        negative: if config.negative_points {
            Some(true)
        } else {
            None
        },
        display_x_axis: if config.show_axis { Some(true) } else { None },
        line_weight: config.line_weight,
        min_axis_type: None,
        max_axis_type: None,
        sparklines: SparklineList {
            items: vec![Sparkline {
                formula: config.data_range.clone(),
                sqref: config.location.clone(),
            }],
        },
    }
}

/// Validate a sparkline configuration.
pub fn validate_sparkline_config(config: &SparklineConfig) -> Result<()> {
    if config.data_range.is_empty() {
        return Err(Error::Internal("sparkline data range is empty".to_string()));
    }
    if config.location.is_empty() {
        return Err(Error::Internal("sparkline location is empty".to_string()));
    }
    if let Some(weight) = config.line_weight {
        if weight <= 0.0 {
            return Err(Error::Internal(
                "sparkline line weight must be positive".to_string(),
            ));
        }
    }
    if let Some(style) = config.style {
        if style > 35 {
            return Err(Error::Internal(format!(
                "sparkline style {} out of range (0-35)",
                style
            )));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sparkline_type_default() {
        assert_eq!(SparklineType::default(), SparklineType::Line);
    }

    #[test]
    fn test_sparkline_type_as_str() {
        assert_eq!(SparklineType::Line.as_str(), "line");
        assert_eq!(SparklineType::Column.as_str(), "column");
        assert_eq!(SparklineType::WinLoss.as_str(), "stacked");
    }

    #[test]
    fn test_sparkline_type_parse() {
        assert_eq!(SparklineType::parse("line"), Some(SparklineType::Line));
        assert_eq!(SparklineType::parse("column"), Some(SparklineType::Column));
        assert_eq!(
            SparklineType::parse("stacked"),
            Some(SparklineType::WinLoss)
        );
        assert_eq!(
            SparklineType::parse("winloss"),
            Some(SparklineType::WinLoss)
        );
        assert_eq!(SparklineType::parse("invalid"), None);
    }

    #[test]
    fn test_sparkline_config_new() {
        let config = SparklineConfig::new("Sheet1!A1:A10", "B1");
        assert_eq!(config.data_range, "Sheet1!A1:A10");
        assert_eq!(config.location, "B1");
        assert_eq!(config.sparkline_type, SparklineType::Line);
        assert!(!config.markers);
    }

    #[test]
    fn test_validate_sparkline_config_ok() {
        let config = SparklineConfig::new("Sheet1!A1:A10", "B1");
        assert!(validate_sparkline_config(&config).is_ok());
    }

    #[test]
    fn test_validate_sparkline_empty_range() {
        let config = SparklineConfig {
            data_range: String::new(),
            ..SparklineConfig::new("", "B1")
        };
        assert!(validate_sparkline_config(&config).is_err());
    }

    #[test]
    fn test_validate_sparkline_empty_location() {
        let config = SparklineConfig {
            location: String::new(),
            ..SparklineConfig::new("Sheet1!A1:A10", "")
        };
        assert!(validate_sparkline_config(&config).is_err());
    }

    #[test]
    fn test_validate_sparkline_invalid_weight() {
        let mut config = SparklineConfig::new("Sheet1!A1:A10", "B1");
        config.line_weight = Some(-1.0);
        assert!(validate_sparkline_config(&config).is_err());
    }

    #[test]
    fn test_validate_sparkline_invalid_style() {
        let mut config = SparklineConfig::new("Sheet1!A1:A10", "B1");
        config.style = Some(36);
        assert!(validate_sparkline_config(&config).is_err());
    }

    #[test]
    fn test_config_to_xml_group_line() {
        let config = SparklineConfig::new("Sheet1!A1:A10", "B1");
        let group = config_to_xml_group(&config);
        assert!(group.sparkline_type.is_none());
        assert_eq!(group.sparklines.items.len(), 1);
    }

    #[test]
    fn test_config_to_xml_group_column() {
        let mut config = SparklineConfig::new("Sheet1!A1:A10", "B1");
        config.sparkline_type = SparklineType::Column;
        config.markers = true;
        config.high_point = true;
        let group = config_to_xml_group(&config);
        assert_eq!(group.sparkline_type, Some("column".to_string()));
        assert_eq!(group.markers, Some(true));
        assert_eq!(group.high, Some(true));
    }

    #[test]
    fn test_config_to_xml_group_winloss() {
        let mut config = SparklineConfig::new("Sheet1!A1:A10", "B1");
        config.sparkline_type = SparklineType::WinLoss;
        let group = config_to_xml_group(&config);
        assert_eq!(group.sparkline_type, Some("stacked".to_string()));
    }
}
