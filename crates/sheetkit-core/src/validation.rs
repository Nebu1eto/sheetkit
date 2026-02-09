//! Data validation builder and utilities.
//!
//! Provides a high-level API for adding, querying, and removing data validation
//! rules on worksheet cells.

use crate::error::{Error, Result};
use sheetkit_xml::worksheet::{DataValidation, DataValidations, WorksheetXml};

/// The type of data validation to apply.
#[derive(Debug, Clone, PartialEq)]
pub enum ValidationType {
    /// No value restriction (prompt/message only).
    None,
    Whole,
    Decimal,
    List,
    Date,
    Time,
    TextLength,
    Custom,
}

impl ValidationType {
    /// Convert to the XML attribute string.
    pub fn as_str(&self) -> &str {
        match self {
            ValidationType::None => "none",
            ValidationType::Whole => "whole",
            ValidationType::Decimal => "decimal",
            ValidationType::List => "list",
            ValidationType::Date => "date",
            ValidationType::Time => "time",
            ValidationType::TextLength => "textLength",
            ValidationType::Custom => "custom",
        }
    }

    /// Parse from the XML attribute string.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "none" => Some(ValidationType::None),
            "whole" => Some(ValidationType::Whole),
            "decimal" => Some(ValidationType::Decimal),
            "list" => Some(ValidationType::List),
            "date" => Some(ValidationType::Date),
            "time" => Some(ValidationType::Time),
            "textLength" => Some(ValidationType::TextLength),
            "custom" => Some(ValidationType::Custom),
            _ => None,
        }
    }

    /// Whether this type requires an operator.
    pub fn uses_operator(&self) -> bool {
        matches!(
            self,
            ValidationType::Whole
                | ValidationType::Decimal
                | ValidationType::Date
                | ValidationType::Time
                | ValidationType::TextLength
        )
    }
}

/// The comparison operator for data validation.
#[derive(Debug, Clone, PartialEq)]
pub enum ValidationOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
}

impl ValidationOperator {
    /// Convert to the XML attribute string.
    pub fn as_str(&self) -> &str {
        match self {
            ValidationOperator::Between => "between",
            ValidationOperator::NotBetween => "notBetween",
            ValidationOperator::Equal => "equal",
            ValidationOperator::NotEqual => "notEqual",
            ValidationOperator::LessThan => "lessThan",
            ValidationOperator::LessThanOrEqual => "lessThanOrEqual",
            ValidationOperator::GreaterThan => "greaterThan",
            ValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        }
    }

    /// Parse from the XML attribute string.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "between" => Some(ValidationOperator::Between),
            "notBetween" => Some(ValidationOperator::NotBetween),
            "equal" => Some(ValidationOperator::Equal),
            "notEqual" => Some(ValidationOperator::NotEqual),
            "lessThan" => Some(ValidationOperator::LessThan),
            "lessThanOrEqual" => Some(ValidationOperator::LessThanOrEqual),
            "greaterThan" => Some(ValidationOperator::GreaterThan),
            "greaterThanOrEqual" => Some(ValidationOperator::GreaterThanOrEqual),
            _ => None,
        }
    }

    /// Whether this operator requires two formulas.
    pub fn needs_formula2(&self) -> bool {
        matches!(
            self,
            ValidationOperator::Between | ValidationOperator::NotBetween
        )
    }
}

/// The error display style for validation failures.
#[derive(Debug, Clone, PartialEq)]
pub enum ErrorStyle {
    Stop,
    Warning,
    Information,
}

impl ErrorStyle {
    /// Convert to the XML attribute string.
    pub fn as_str(&self) -> &str {
        match self {
            ErrorStyle::Stop => "stop",
            ErrorStyle::Warning => "warning",
            ErrorStyle::Information => "information",
        }
    }

    /// Parse from the XML attribute string.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "stop" => Some(ErrorStyle::Stop),
            "warning" => Some(ErrorStyle::Warning),
            "information" => Some(ErrorStyle::Information),
            _ => None,
        }
    }
}

/// Configuration for a data validation rule.
#[derive(Debug, Clone)]
pub struct DataValidationConfig {
    /// The cell range to apply validation to (e.g. "A1:A100").
    pub sqref: String,
    /// The type of validation.
    pub validation_type: ValidationType,
    /// The comparison operator (not used for list validations).
    pub operator: Option<ValidationOperator>,
    /// The first formula/value for the validation constraint.
    pub formula1: Option<String>,
    /// The second formula/value (used with Between/NotBetween operators).
    pub formula2: Option<String>,
    /// Whether blank cells are allowed.
    pub allow_blank: bool,
    /// The error display style.
    pub error_style: Option<ErrorStyle>,
    /// The title for the error dialog.
    pub error_title: Option<String>,
    /// The message for the error dialog.
    pub error_message: Option<String>,
    /// The title for the input prompt.
    pub prompt_title: Option<String>,
    /// The message for the input prompt.
    pub prompt_message: Option<String>,
    /// Whether to show the input message when the cell is selected.
    pub show_input_message: bool,
    /// Whether to show the error message on invalid input.
    pub show_error_message: bool,
}

impl DataValidationConfig {
    /// Create a dropdown list validation.
    ///
    /// The items are joined with commas and quoted for the formula.
    /// Individual items must not contain commas (Excel limitation).
    pub fn dropdown(sqref: &str, items: &[&str]) -> Self {
        let formula = format!("\"{}\"", items.join(","));
        Self {
            sqref: sqref.to_string(),
            validation_type: ValidationType::List,
            operator: None,
            formula1: Some(formula),
            formula2: None,
            allow_blank: true,
            error_style: Some(ErrorStyle::Stop),
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
            show_input_message: true,
            show_error_message: true,
        }
    }

    /// Create a whole number range validation (between min and max).
    pub fn whole_number(sqref: &str, min: i64, max: i64) -> Self {
        Self {
            sqref: sqref.to_string(),
            validation_type: ValidationType::Whole,
            operator: Some(ValidationOperator::Between),
            formula1: Some(min.to_string()),
            formula2: Some(max.to_string()),
            allow_blank: true,
            error_style: Some(ErrorStyle::Stop),
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
            show_input_message: true,
            show_error_message: true,
        }
    }

    /// Create a decimal range validation (between min and max).
    pub fn decimal(sqref: &str, min: f64, max: f64) -> Self {
        Self {
            sqref: sqref.to_string(),
            validation_type: ValidationType::Decimal,
            operator: Some(ValidationOperator::Between),
            formula1: Some(min.to_string()),
            formula2: Some(max.to_string()),
            allow_blank: true,
            error_style: Some(ErrorStyle::Stop),
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
            show_input_message: true,
            show_error_message: true,
        }
    }

    /// Create a text length validation.
    pub fn text_length(sqref: &str, operator: ValidationOperator, length: u32) -> Self {
        Self {
            sqref: sqref.to_string(),
            validation_type: ValidationType::TextLength,
            operator: Some(operator),
            formula1: Some(length.to_string()),
            formula2: None,
            allow_blank: true,
            error_style: Some(ErrorStyle::Stop),
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
            show_input_message: true,
            show_error_message: true,
        }
    }
}

/// Validate that `sqref` looks like a valid cell range reference.
///
/// Accepts single refs ("A1"), ranges ("A1:B10"), and space-separated
/// multi-area refs ("A1:B10 D1:E10"). This is not exhaustive but catches
/// obvious mistakes like empty strings.
fn validate_sqref(sqref: &str) -> Result<()> {
    if sqref.is_empty() {
        return Err(Error::InvalidReference {
            reference: sqref.to_string(),
        });
    }
    // Each part (split by space) must match a cell or range pattern.
    for part in sqref.split(' ') {
        if part.is_empty() {
            return Err(Error::InvalidReference {
                reference: sqref.to_string(),
            });
        }
        // Allow "A1" or "A1:B10" shapes.  Each side must start with a letter
        // and contain at least one digit.
        for side in part.split(':') {
            let has_alpha = side.chars().any(|c| c.is_ascii_alphabetic());
            let has_digit = side.chars().any(|c| c.is_ascii_digit());
            if !has_alpha || !has_digit {
                return Err(Error::InvalidReference {
                    reference: sqref.to_string(),
                });
            }
        }
    }
    Ok(())
}

/// Validate formula constraints for the given validation type and operator.
fn validate_formulas(config: &DataValidationConfig) -> Result<()> {
    match &config.validation_type {
        ValidationType::None => {}
        ValidationType::List | ValidationType::Custom => {
            if config.formula1.as_ref().is_none_or(|f| f.is_empty()) {
                return Err(Error::InvalidArgument(format!(
                    "formula1 is required for {:?} validation",
                    config.validation_type
                )));
            }
        }
        _ => {
            // Types that use an operator need formula1 at minimum.
            if config.formula1.as_ref().is_none_or(|f| f.is_empty()) {
                return Err(Error::InvalidArgument(format!(
                    "formula1 is required for {:?} validation",
                    config.validation_type
                )));
            }
            if let Some(op) = &config.operator {
                if op.needs_formula2() && config.formula2.as_ref().is_none_or(|f| f.is_empty()) {
                    return Err(Error::InvalidArgument(format!(
                        "formula2 is required for {:?} operator",
                        op
                    )));
                }
            }
        }
    }
    Ok(())
}

/// Convert a `DataValidationConfig` to the XML `DataValidation` struct.
pub fn config_to_xml(config: &DataValidationConfig) -> DataValidation {
    DataValidation {
        validation_type: Some(config.validation_type.as_str().to_string()),
        operator: config.operator.as_ref().map(|o| o.as_str().to_string()),
        allow_blank: if config.allow_blank { Some(true) } else { None },
        show_drop_down: None,
        show_input_message: if config.show_input_message {
            Some(true)
        } else {
            None
        },
        show_error_message: if config.show_error_message {
            Some(true)
        } else {
            None
        },
        error_style: config.error_style.as_ref().map(|e| e.as_str().to_string()),
        ime_mode: None,
        error_title: config.error_title.clone(),
        error: config.error_message.clone(),
        prompt_title: config.prompt_title.clone(),
        prompt: config.prompt_message.clone(),
        sqref: config.sqref.clone(),
        formula1: config.formula1.clone(),
        formula2: config.formula2.clone(),
    }
}

/// Convert an XML `DataValidation` to a `DataValidationConfig`.
fn xml_to_config(dv: &DataValidation) -> DataValidationConfig {
    DataValidationConfig {
        sqref: dv.sqref.clone(),
        validation_type: dv
            .validation_type
            .as_deref()
            .and_then(ValidationType::parse)
            .unwrap_or(ValidationType::None),
        operator: dv.operator.as_deref().and_then(ValidationOperator::parse),
        formula1: dv.formula1.clone(),
        formula2: dv.formula2.clone(),
        allow_blank: dv.allow_blank.unwrap_or(false),
        error_style: dv.error_style.as_deref().and_then(ErrorStyle::parse),
        error_title: dv.error_title.clone(),
        error_message: dv.error.clone(),
        prompt_title: dv.prompt_title.clone(),
        prompt_message: dv.prompt.clone(),
        show_input_message: dv.show_input_message.unwrap_or(false),
        show_error_message: dv.show_error_message.unwrap_or(false),
    }
}

/// Add a data validation to a worksheet.
pub fn add_validation(ws: &mut WorksheetXml, config: &DataValidationConfig) -> Result<()> {
    validate_sqref(&config.sqref)?;
    validate_formulas(config)?;
    let dv = config_to_xml(config);
    let dvs = ws.data_validations.get_or_insert_with(|| DataValidations {
        count: Some(0),
        disable_prompts: None,
        x_window: None,
        y_window: None,
        data_validations: Vec::new(),
    });
    dvs.data_validations.push(dv);
    dvs.count = Some(dvs.data_validations.len() as u32);
    Ok(())
}

/// Get all data validations from a worksheet.
pub fn get_validations(ws: &WorksheetXml) -> Vec<DataValidationConfig> {
    match &ws.data_validations {
        Some(dvs) => dvs.data_validations.iter().map(xml_to_config).collect(),
        None => Vec::new(),
    }
}

/// Remove validations matching a specific cell range from a worksheet.
///
/// Returns `Ok(())` regardless of whether any validations were actually removed.
pub fn remove_validation(ws: &mut WorksheetXml, sqref: &str) -> Result<()> {
    if let Some(ref mut dvs) = ws.data_validations {
        dvs.data_validations.retain(|dv| dv.sqref != sqref);
        dvs.count = Some(dvs.data_validations.len() as u32);
        if dvs.data_validations.is_empty() {
            ws.data_validations = None;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_dropdown_validation() {
        let config = DataValidationConfig::dropdown("A1:A100", &["Yes", "No", "Maybe"]);
        assert_eq!(config.sqref, "A1:A100");
        assert_eq!(config.validation_type, ValidationType::List);
        assert_eq!(config.formula1, Some("\"Yes,No,Maybe\"".to_string()));
        assert!(config.allow_blank);
        assert!(config.show_input_message);
        assert!(config.show_error_message);
    }

    #[test]
    fn test_whole_number_validation() {
        let config = DataValidationConfig::whole_number("B1:B50", 1, 100);
        assert_eq!(config.sqref, "B1:B50");
        assert_eq!(config.validation_type, ValidationType::Whole);
        assert_eq!(config.operator, Some(ValidationOperator::Between));
        assert_eq!(config.formula1, Some("1".to_string()));
        assert_eq!(config.formula2, Some("100".to_string()));
    }

    #[test]
    fn test_decimal_validation() {
        let config = DataValidationConfig::decimal("C1:C10", 0.0, 99.99);
        assert_eq!(config.sqref, "C1:C10");
        assert_eq!(config.validation_type, ValidationType::Decimal);
        assert_eq!(config.operator, Some(ValidationOperator::Between));
        assert_eq!(config.formula1, Some("0".to_string()));
        assert_eq!(config.formula2, Some("99.99".to_string()));
    }

    #[test]
    fn test_text_length_validation() {
        let config =
            DataValidationConfig::text_length("D1:D10", ValidationOperator::LessThanOrEqual, 255);
        assert_eq!(config.sqref, "D1:D10");
        assert_eq!(config.validation_type, ValidationType::TextLength);
        assert_eq!(config.operator, Some(ValidationOperator::LessThanOrEqual));
        assert_eq!(config.formula1, Some("255".to_string()));
    }

    #[test]
    fn test_config_to_xml_roundtrip() {
        let config = DataValidationConfig::dropdown("A1:A10", &["Red", "Blue"]);
        let xml = config_to_xml(&config);
        assert_eq!(xml.validation_type, Some("list".to_string()));
        assert_eq!(xml.sqref, "A1:A10");
        assert_eq!(xml.formula1, Some("\"Red,Blue\"".to_string()));
        assert_eq!(xml.allow_blank, Some(true));
        assert_eq!(xml.show_input_message, Some(true));
        assert_eq!(xml.show_error_message, Some(true));
    }

    #[test]
    fn test_add_validation_to_worksheet() {
        let mut ws = WorksheetXml::default();
        let config = DataValidationConfig::dropdown("A1:A100", &["Yes", "No"]);
        add_validation(&mut ws, &config).unwrap();

        assert!(ws.data_validations.is_some());
        let dvs = ws.data_validations.as_ref().unwrap();
        assert_eq!(dvs.count, Some(1));
        assert_eq!(dvs.data_validations.len(), 1);
        assert_eq!(dvs.data_validations[0].sqref, "A1:A100");
    }

    #[test]
    fn test_add_multiple_validations() {
        let mut ws = WorksheetXml::default();
        let config1 = DataValidationConfig::dropdown("A1:A100", &["Yes", "No"]);
        let config2 = DataValidationConfig::whole_number("B1:B100", 1, 100);
        add_validation(&mut ws, &config1).unwrap();
        add_validation(&mut ws, &config2).unwrap();

        let dvs = ws.data_validations.as_ref().unwrap();
        assert_eq!(dvs.count, Some(2));
        assert_eq!(dvs.data_validations.len(), 2);
    }

    #[test]
    fn test_get_validations() {
        let mut ws = WorksheetXml::default();
        let config = DataValidationConfig::dropdown("A1:A100", &["Yes", "No"]);
        add_validation(&mut ws, &config).unwrap();

        let configs = get_validations(&ws);
        assert_eq!(configs.len(), 1);
        assert_eq!(configs[0].sqref, "A1:A100");
        assert_eq!(configs[0].validation_type, ValidationType::List);
    }

    #[test]
    fn test_get_validations_empty() {
        let ws = WorksheetXml::default();
        let configs = get_validations(&ws);
        assert!(configs.is_empty());
    }

    #[test]
    fn test_remove_validation() {
        let mut ws = WorksheetXml::default();
        let config1 = DataValidationConfig::dropdown("A1:A100", &["Yes", "No"]);
        let config2 = DataValidationConfig::whole_number("B1:B100", 1, 100);
        add_validation(&mut ws, &config1).unwrap();
        add_validation(&mut ws, &config2).unwrap();

        remove_validation(&mut ws, "A1:A100").unwrap();

        let dvs = ws.data_validations.as_ref().unwrap();
        assert_eq!(dvs.count, Some(1));
        assert_eq!(dvs.data_validations.len(), 1);
        assert_eq!(dvs.data_validations[0].sqref, "B1:B100");
    }

    #[test]
    fn test_remove_last_validation_clears_container() {
        let mut ws = WorksheetXml::default();
        let config = DataValidationConfig::dropdown("A1:A100", &["Yes", "No"]);
        add_validation(&mut ws, &config).unwrap();
        remove_validation(&mut ws, "A1:A100").unwrap();

        assert!(ws.data_validations.is_none());
    }

    #[test]
    fn test_remove_nonexistent_validation() {
        let mut ws = WorksheetXml::default();
        // Should not error when removing from empty worksheet
        remove_validation(&mut ws, "Z1:Z99").unwrap();
        assert!(ws.data_validations.is_none());
    }

    #[test]
    fn test_validation_xml_serialization_roundtrip() {
        let mut ws = WorksheetXml::default();
        let config = DataValidationConfig::dropdown("A1:A10", &["Apple", "Banana"]);
        add_validation(&mut ws, &config).unwrap();

        let xml = quick_xml::se::to_string(&ws).unwrap();
        assert!(xml.contains("dataValidations"));
        assert!(xml.contains("A1:A10"));

        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.data_validations.is_some());
        let dvs = parsed.data_validations.as_ref().unwrap();
        assert_eq!(dvs.data_validations.len(), 1);
        assert_eq!(dvs.data_validations[0].sqref, "A1:A10");
        assert_eq!(
            dvs.data_validations[0].validation_type,
            Some("list".to_string())
        );
    }

    #[test]
    fn test_whole_number_validation_xml_roundtrip() {
        let mut ws = WorksheetXml::default();
        let config = DataValidationConfig::whole_number("B1:B50", 10, 200);
        add_validation(&mut ws, &config).unwrap();

        let xml = quick_xml::se::to_string(&ws).unwrap();
        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();

        let configs = get_validations(&parsed);
        assert_eq!(configs.len(), 1);
        assert_eq!(configs[0].sqref, "B1:B50");
        assert_eq!(configs[0].validation_type, ValidationType::Whole);
        assert_eq!(configs[0].operator, Some(ValidationOperator::Between));
        assert_eq!(configs[0].formula1, Some("10".to_string()));
        assert_eq!(configs[0].formula2, Some("200".to_string()));
    }

    #[test]
    fn test_decimal_validation_xml_roundtrip() {
        let mut ws = WorksheetXml::default();
        let config = DataValidationConfig::decimal("C1:C10", 1.5, 99.9);
        add_validation(&mut ws, &config).unwrap();

        let xml = quick_xml::se::to_string(&ws).unwrap();
        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();

        let configs = get_validations(&parsed);
        assert_eq!(configs.len(), 1);
        assert_eq!(configs[0].validation_type, ValidationType::Decimal);
    }

    #[test]
    fn test_validation_type_as_str() {
        assert_eq!(ValidationType::None.as_str(), "none");
        assert_eq!(ValidationType::Whole.as_str(), "whole");
        assert_eq!(ValidationType::Decimal.as_str(), "decimal");
        assert_eq!(ValidationType::List.as_str(), "list");
        assert_eq!(ValidationType::Date.as_str(), "date");
        assert_eq!(ValidationType::Time.as_str(), "time");
        assert_eq!(ValidationType::TextLength.as_str(), "textLength");
        assert_eq!(ValidationType::Custom.as_str(), "custom");
    }

    #[test]
    fn test_validation_operator_as_str() {
        assert_eq!(ValidationOperator::Between.as_str(), "between");
        assert_eq!(ValidationOperator::NotBetween.as_str(), "notBetween");
        assert_eq!(ValidationOperator::Equal.as_str(), "equal");
        assert_eq!(ValidationOperator::NotEqual.as_str(), "notEqual");
        assert_eq!(ValidationOperator::LessThan.as_str(), "lessThan");
        assert_eq!(
            ValidationOperator::LessThanOrEqual.as_str(),
            "lessThanOrEqual"
        );
        assert_eq!(ValidationOperator::GreaterThan.as_str(), "greaterThan");
        assert_eq!(
            ValidationOperator::GreaterThanOrEqual.as_str(),
            "greaterThanOrEqual"
        );
    }

    #[test]
    fn test_error_style_as_str() {
        assert_eq!(ErrorStyle::Stop.as_str(), "stop");
        assert_eq!(ErrorStyle::Warning.as_str(), "warning");
        assert_eq!(ErrorStyle::Information.as_str(), "information");
    }

    #[test]
    fn test_none_type_roundtrip() {
        assert_eq!(ValidationType::parse("none"), Some(ValidationType::None));
        assert_eq!(ValidationType::None.as_str(), "none");
    }

    #[test]
    fn test_unknown_type_defaults_to_none() {
        let dv = DataValidation {
            validation_type: Some("unknownFuture".to_string()),
            operator: None,
            allow_blank: None,
            show_drop_down: None,
            show_input_message: None,
            show_error_message: None,
            error_style: None,
            ime_mode: None,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            sqref: "A1".to_string(),
            formula1: None,
            formula2: None,
        };
        let config = xml_to_config(&dv);
        assert_eq!(config.validation_type, ValidationType::None);
    }

    #[test]
    fn test_validate_sqref_valid() {
        assert!(validate_sqref("A1").is_ok());
        assert!(validate_sqref("A1:B10").is_ok());
        assert!(validate_sqref("A1:B10 D1:E10").is_ok());
        assert!(validate_sqref("AA100:ZZ999").is_ok());
    }

    #[test]
    fn test_validate_sqref_invalid() {
        assert!(validate_sqref("").is_err());
        assert!(validate_sqref("hello").is_err());
        assert!(validate_sqref("123").is_err());
        assert!(validate_sqref("A1: B10").is_err()); // space inside range
    }

    #[test]
    fn test_add_validation_rejects_empty_sqref() {
        let mut ws = WorksheetXml::default();
        let config = DataValidationConfig {
            sqref: "".to_string(),
            validation_type: ValidationType::List,
            operator: None,
            formula1: Some("\"A,B\"".to_string()),
            formula2: None,
            allow_blank: false,
            error_style: None,
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
            show_input_message: false,
            show_error_message: false,
        };
        assert!(add_validation(&mut ws, &config).is_err());
    }

    #[test]
    fn test_add_validation_rejects_missing_formula1_for_list() {
        let mut ws = WorksheetXml::default();
        let config = DataValidationConfig {
            sqref: "A1:A10".to_string(),
            validation_type: ValidationType::List,
            operator: None,
            formula1: None,
            formula2: None,
            allow_blank: false,
            error_style: None,
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
            show_input_message: false,
            show_error_message: false,
        };
        assert!(add_validation(&mut ws, &config).is_err());
    }

    #[test]
    fn test_add_validation_rejects_missing_formula2_for_between() {
        let mut ws = WorksheetXml::default();
        let config = DataValidationConfig {
            sqref: "A1:A10".to_string(),
            validation_type: ValidationType::Whole,
            operator: Some(ValidationOperator::Between),
            formula1: Some("1".to_string()),
            formula2: None,
            allow_blank: false,
            error_style: None,
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
            show_input_message: false,
            show_error_message: false,
        };
        assert!(add_validation(&mut ws, &config).is_err());
    }

    #[test]
    fn test_none_type_no_formula_required() {
        let mut ws = WorksheetXml::default();
        let config = DataValidationConfig {
            sqref: "A1:A10".to_string(),
            validation_type: ValidationType::None,
            operator: None,
            formula1: None,
            formula2: None,
            allow_blank: false,
            error_style: None,
            error_title: None,
            error_message: None,
            prompt_title: Some("Hint".to_string()),
            prompt_message: Some("Enter a value".to_string()),
            show_input_message: true,
            show_error_message: false,
        };
        assert!(add_validation(&mut ws, &config).is_ok());
        let configs = get_validations(&ws);
        assert_eq!(configs[0].validation_type, ValidationType::None);
    }

    #[test]
    fn test_uses_operator() {
        assert!(!ValidationType::None.uses_operator());
        assert!(ValidationType::Whole.uses_operator());
        assert!(ValidationType::Decimal.uses_operator());
        assert!(!ValidationType::List.uses_operator());
        assert!(ValidationType::Date.uses_operator());
        assert!(ValidationType::Time.uses_operator());
        assert!(ValidationType::TextLength.uses_operator());
        assert!(!ValidationType::Custom.uses_operator());
    }

    #[test]
    fn test_needs_formula2() {
        assert!(ValidationOperator::Between.needs_formula2());
        assert!(ValidationOperator::NotBetween.needs_formula2());
        assert!(!ValidationOperator::Equal.needs_formula2());
        assert!(!ValidationOperator::GreaterThan.needs_formula2());
    }

    #[test]
    fn test_show_drop_down_preserved_in_xml() {
        let dv = DataValidation {
            validation_type: Some("list".to_string()),
            operator: None,
            allow_blank: None,
            show_drop_down: Some(true),
            show_input_message: None,
            show_error_message: None,
            error_style: None,
            ime_mode: None,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            sqref: "A1".to_string(),
            formula1: Some("\"A,B\"".to_string()),
            formula2: None,
        };
        let xml = quick_xml::se::to_string(&dv).unwrap();
        assert!(xml.contains("showDropDown"));

        let parsed: DataValidation = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.show_drop_down, Some(true));
    }

    #[test]
    fn test_ime_mode_preserved_in_xml() {
        let dv = DataValidation {
            validation_type: Some("whole".to_string()),
            operator: None,
            allow_blank: None,
            show_drop_down: None,
            show_input_message: None,
            show_error_message: None,
            error_style: None,
            ime_mode: Some("hiragana".to_string()),
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            sqref: "A1".to_string(),
            formula1: Some("1".to_string()),
            formula2: None,
        };
        let xml = quick_xml::se::to_string(&dv).unwrap();
        assert!(xml.contains("imeMode"));

        let parsed: DataValidation = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.ime_mode, Some("hiragana".to_string()));
    }

    #[test]
    fn test_container_attrs_preserved_in_xml() {
        let dvs = DataValidations {
            count: Some(0),
            disable_prompts: Some(true),
            x_window: Some(100),
            y_window: Some(200),
            data_validations: Vec::new(),
        };
        let xml = quick_xml::se::to_string(&dvs).unwrap();
        assert!(xml.contains("disablePrompts"));
        assert!(xml.contains("xWindow"));
        assert!(xml.contains("yWindow"));

        let parsed: DataValidations = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.disable_prompts, Some(true));
        assert_eq!(parsed.x_window, Some(100));
        assert_eq!(parsed.y_window, Some(200));
    }
}
