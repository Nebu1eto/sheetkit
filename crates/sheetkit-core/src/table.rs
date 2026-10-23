//! Table and auto-filter management.
//!
//! Provides functions for creating, listing, and deleting OOXML table parts on
//! worksheets, as well as the existing auto-filter helpers.

use crate::error::{Error, Result};
use crate::utils::cell_ref::{cell_name_to_coordinates, coordinates_to_cell_name};
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS};
use sheetkit_xml::table::{
    TableAutoFilter, TableColumnXml, TableColumnsXml, TableStyleInfoXml, TableXml,
};
use sheetkit_xml::worksheet::{AutoFilter, WorksheetXml};

/// Configuration for creating a table.
///
/// When any column sets `totals_row_function` or `totals_row_label`, the
/// table gets a totals row: `range` must then include it as its last row
/// (header, at least one data row, totals row), the auto-filter covers the
/// range without the totals row, and `add_table` writes each label or
/// `SUBTOTAL` formula into the totals row cells.
#[derive(Debug, Clone)]
pub struct TableConfig {
    /// The table name (used internally, must be unique within the workbook).
    pub name: String,
    /// The display name shown in the UI.
    pub display_name: String,
    /// The cell range (e.g. "A1:D10"), including the header row and, when
    /// present, the totals row.
    pub range: String,
    /// Column definitions.
    pub columns: Vec<TableColumn>,
    /// Whether to show the header row. Defaults to true.
    pub show_header_row: bool,
    /// The table style name (e.g. "TableStyleMedium2").
    pub style_name: Option<String>,
    /// Whether to enable auto-filter on the table.
    pub auto_filter: bool,
    /// Whether to show first column formatting.
    pub show_first_column: bool,
    /// Whether to show last column formatting.
    pub show_last_column: bool,
    /// Whether to show row stripes.
    pub show_row_stripes: bool,
    /// Whether to show column stripes.
    pub show_column_stripes: bool,
}

impl Default for TableConfig {
    fn default() -> Self {
        Self {
            name: String::new(),
            display_name: String::new(),
            range: String::new(),
            columns: Vec::new(),
            show_header_row: true,
            style_name: None,
            auto_filter: true,
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        }
    }
}

/// A column within a table.
#[derive(Debug, Clone)]
pub struct TableColumn {
    /// The column header name.
    pub name: String,
    /// Optional totals row function: "sum", "min", "max", "average",
    /// "count", "countNums", "stdDev", "var", or "none".
    pub totals_row_function: Option<String>,
    /// Optional totals row label (used for the first column in totals row).
    /// Takes precedence over `totals_row_function` for the cell content.
    pub totals_row_label: Option<String>,
}

/// Metadata about an existing table, returned by list/get operations.
#[derive(Debug, Clone, PartialEq)]
pub struct TableInfo {
    /// The table name.
    pub name: String,
    /// The display name.
    pub display_name: String,
    /// The cell range (e.g. "A1:D10").
    pub range: String,
    /// Whether the table has a header row.
    pub show_header_row: bool,
    /// Whether auto-filter is enabled.
    pub auto_filter: bool,
    /// Column names.
    pub columns: Vec<String>,
    /// The style name, if any.
    pub style_name: Option<String>,
}

/// Maximum length, in characters, of a table name or display name.
pub const MAX_TABLE_NAME_LEN: usize = 255;

/// Return the display name written for `config`: the configured display
/// name, or the table name when the display name is empty.
pub(crate) fn effective_display_name(config: &TableConfig) -> &str {
    if config.display_name.is_empty() {
        &config.name
    } else {
        &config.display_name
    }
}

/// Returns true when `name` would be read by Excel as an A1 cell reference
/// within sheet bounds (e.g. `A1`, `XFD1048576`).
fn is_a1_cell_reference(name: &str) -> bool {
    let letters = name.chars().take_while(char::is_ascii_alphabetic).count();
    if !(1..=3).contains(&letters) || letters == name.len() {
        return false;
    }
    let (col_part, row_part) = name.split_at(letters);
    if !row_part.chars().all(|c| c.is_ascii_digit()) {
        return false;
    }
    let col = col_part.chars().fold(0u32, |acc, c| {
        acc * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
    });
    let row = row_part.parse::<u64>().unwrap_or(u64::MAX);
    col <= MAX_COLUMNS && (1..=u64::from(MAX_ROWS)).contains(&row)
}

/// Returns true when `name` would be read by Excel as an R1C1 reference
/// (`R`, `C`, `RC`, `R1`, `C1`, `R1C1`, ...).
fn is_r1c1_reference(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    let digits = |s: &str| s.chars().all(|c| c.is_ascii_digit());
    if let Some(rest) = upper.strip_prefix('R') {
        match rest.split_once('C') {
            Some((row, col)) => digits(row) && digits(col),
            None => digits(rest),
        }
    } else if let Some(rest) = upper.strip_prefix('C') {
        digits(rest)
    } else {
        false
    }
}

/// Validate a table name or display name against Excel naming rules.
///
/// The name must be 1 to 255 characters, start with a letter, underscore,
/// or backslash, continue with letters, digits, periods, or underscores,
/// and must not look like an A1 or R1C1 cell reference.
pub(crate) fn validate_table_name(kind: &str, name: &str) -> Result<()> {
    let len = name.chars().count();
    if len == 0 {
        return Err(Error::InvalidArgument(format!("{kind} cannot be empty")));
    }
    if len > MAX_TABLE_NAME_LEN {
        return Err(Error::InvalidArgument(format!(
            "{kind} exceeds {MAX_TABLE_NAME_LEN} characters"
        )));
    }
    let mut chars = name.chars();
    let first = chars.next().unwrap_or_default();
    if !(first.is_alphabetic() || first == '_' || first == '\\') {
        return Err(Error::InvalidArgument(format!(
            "{kind} must start with a letter, underscore, or backslash: {name}"
        )));
    }
    if let Some(bad) = chars.find(|c| !(c.is_alphanumeric() || *c == '.' || *c == '_')) {
        return Err(Error::InvalidArgument(format!(
            "{kind} contains invalid character {bad:?}: {name}"
        )));
    }
    if is_a1_cell_reference(name) || is_r1c1_reference(name) {
        return Err(Error::InvalidArgument(format!(
            "{kind} must not be a cell reference: {name}"
        )));
    }
    Ok(())
}

/// Build a `TableXml` from a `TableConfig` and a unique table ID.
pub(crate) fn build_table_xml(config: &TableConfig, table_id: u32) -> TableXml {
    let columns: Vec<TableColumnXml> = config
        .columns
        .iter()
        .enumerate()
        .map(|(i, col)| TableColumnXml {
            id: (i + 1) as u32,
            name: col.name.clone(),
            totals_row_function: col.totals_row_function.clone(),
            totals_row_label: col.totals_row_label.clone(),
            ..Default::default()
        })
        .collect();

    let totals_row = has_totals_row(config);
    let auto_filter = if config.auto_filter {
        Some(TableAutoFilter {
            reference: if totals_row {
                range_without_last_row(&config.range).unwrap_or_else(|| config.range.clone())
            } else {
                config.range.clone()
            },
            ..Default::default()
        })
    } else {
        None
    };

    let style_info = if config.style_name.is_some()
        || config.show_first_column
        || config.show_last_column
        || config.show_row_stripes
        || config.show_column_stripes
    {
        Some(TableStyleInfoXml {
            name: config.style_name.clone(),
            show_first_column: Some(config.show_first_column),
            show_last_column: Some(config.show_last_column),
            show_row_stripes: Some(config.show_row_stripes),
            show_column_stripes: Some(config.show_column_stripes),
        })
    } else {
        None
    };

    let header_row_count = if !config.show_header_row {
        Some(0)
    } else {
        None
    };

    TableXml {
        xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
        id: table_id,
        name: config.name.clone(),
        display_name: effective_display_name(config).to_string(),
        reference: config.range.clone(),
        totals_row_count: totals_row.then_some(1),
        totals_row_shown: None,
        header_row_count,
        auto_filter,
        table_columns: TableColumnsXml {
            count: columns.len() as u32,
            columns,
        },
        table_style_info: style_info,
        ..Default::default()
    }
}

/// Convert a `TableXml` into a `TableInfo` for external consumption.
pub(crate) fn table_xml_to_info(table_xml: &TableXml) -> TableInfo {
    let columns = table_xml
        .table_columns
        .columns
        .iter()
        .map(|c| c.name.clone())
        .collect();
    let show_header_row = table_xml.header_row_count != Some(0);
    let auto_filter = table_xml.auto_filter.is_some();
    let style_name = table_xml
        .table_style_info
        .as_ref()
        .and_then(|s| s.name.clone());

    TableInfo {
        name: table_xml.name.clone(),
        display_name: table_xml.display_name.clone(),
        range: table_xml.reference.clone(),
        show_header_row,
        auto_filter,
        columns,
        style_name,
    }
}

fn table_range_bounds(range: &str) -> Result<(u32, u32, u32, u32)> {
    let mut cells = range.split(':');
    let start = cells.next().unwrap_or_default();
    let end = cells.next().unwrap_or(start);
    if cells.next().is_some() {
        return Err(Error::InvalidArgument(format!(
            "table range must contain one cell or a single cell range: {range}"
        )));
    }

    let (start_col, start_row) = cell_name_to_coordinates(start)?;
    let (end_col, end_row) = cell_name_to_coordinates(end)?;
    Ok((
        start_col.min(end_col),
        start_row.min(end_row),
        start_col.max(end_col),
        start_row.max(end_row),
    ))
}

/// Totals row functions accepted by `TableColumn::totals_row_function` and
/// the `SUBTOTAL` function number written for each.
const TOTALS_ROW_FUNCTIONS: [(&str, Option<u32>); 9] = [
    ("none", None),
    ("sum", Some(109)),
    ("min", Some(105)),
    ("max", Some(104)),
    ("average", Some(101)),
    ("count", Some(103)),
    ("countNums", Some(102)),
    ("stdDev", Some(107)),
    ("var", Some(110)),
];

/// Returns true when any column defines a totals row label or function.
pub(crate) fn has_totals_row(config: &TableConfig) -> bool {
    config
        .columns
        .iter()
        .any(|c| c.totals_row_function.is_some() || c.totals_row_label.is_some())
}

fn range_without_last_row(range: &str) -> Option<String> {
    let (start_col, start_row, end_col, end_row) = table_range_bounds(range).ok()?;
    let start = coordinates_to_cell_name(start_col, start_row).ok()?;
    let end = coordinates_to_cell_name(end_col, end_row.checked_sub(1)?).ok()?;
    Some(format!("{start}:{end}"))
}

/// Quote a column name for use inside a structured reference.
fn structured_reference_column(name: &str) -> String {
    let mut escaped = String::with_capacity(name.len());
    for c in name.chars() {
        if matches!(c, '\'' | '#' | '[' | ']' | '@') {
            escaped.push('\'');
        }
        escaped.push(c);
    }
    escaped
}

/// Content written into a totals row cell.
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum TotalsCell {
    Label(String),
    Formula(String),
}

/// Return the totals row cell reference and content for each column that
/// defines a label or a non-`none` function. Empty without a totals row.
pub(crate) fn table_totals_cells(config: &TableConfig) -> Result<Vec<(String, TotalsCell)>> {
    if !has_totals_row(config) {
        return Ok(Vec::new());
    }
    let (start_col, _, _, end_row) = table_range_bounds(&config.range)?;
    let mut cells = Vec::new();
    for (offset, column) in config.columns.iter().enumerate() {
        let cell = coordinates_to_cell_name(start_col + offset as u32, end_row)?;
        if let Some(label) = &column.totals_row_label {
            cells.push((cell, TotalsCell::Label(label.clone())));
        } else if let Some(function) = &column.totals_row_function {
            let code = TOTALS_ROW_FUNCTIONS
                .iter()
                .find(|(name, _)| name == function)
                .and_then(|(_, code)| *code);
            if let Some(code) = code {
                cells.push((
                    cell,
                    TotalsCell::Formula(format!(
                        "SUBTOTAL({code},[{}])",
                        structured_reference_column(&column.name)
                    )),
                ));
            }
        }
    }
    Ok(cells)
}

/// Return the header cell reference and name for each configured table column.
pub(crate) fn table_header_cells(config: &TableConfig) -> Result<Vec<(String, String)>> {
    if !config.show_header_row {
        return Ok(Vec::new());
    }

    let (start_col, start_row, _, _) = table_range_bounds(&config.range)?;
    config
        .columns
        .iter()
        .enumerate()
        .map(|(offset, column)| {
            Ok((
                coordinates_to_cell_name(start_col + offset as u32, start_row)?,
                column.name.clone(),
            ))
        })
        .collect()
}

/// Validate that a table name, range, and column definitions are consistent.
pub(crate) fn validate_table_config(config: &TableConfig) -> Result<()> {
    validate_table_name("table name", &config.name)?;
    validate_table_name("table display name", effective_display_name(config))?;
    if config.range.is_empty() {
        return Err(Error::InvalidArgument("table range cannot be empty".into()));
    }
    if config.columns.is_empty() {
        return Err(Error::InvalidArgument(
            "table must have at least one column".into(),
        ));
    }

    let (start_col, start_row, end_col, end_row) = table_range_bounds(&config.range)?;
    let range_width = end_col - start_col + 1;
    if range_width != config.columns.len() as u32 {
        return Err(Error::InvalidArgument(format!(
            "table range width ({range_width}) must match column count ({})",
            config.columns.len()
        )));
    }
    for column in &config.columns {
        if let Some(function) = &column.totals_row_function {
            if !TOTALS_ROW_FUNCTIONS
                .iter()
                .any(|(name, _)| name == function)
            {
                return Err(Error::InvalidArgument(format!(
                    "unsupported totals row function {function:?} for column {}",
                    column.name
                )));
            }
        }
    }
    if has_totals_row(config) {
        let header_rows = u32::from(config.show_header_row);
        let rows = end_row - start_row + 1;
        if rows < header_rows + 2 {
            return Err(Error::InvalidArgument(format!(
                "table range {} must include at least one data row and the totals row",
                config.range
            )));
        }
    }
    Ok(())
}

/// Set an auto-filter on a worksheet for the given cell range.
pub fn set_auto_filter(ws: &mut WorksheetXml, range: &str) -> Result<()> {
    ws.auto_filter = Some(AutoFilter {
        reference: range.to_string(),
    });
    Ok(())
}

/// Remove any auto-filter from a worksheet.
pub fn remove_auto_filter(ws: &mut WorksheetXml) {
    ws.auto_filter = None;
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_set_auto_filter() {
        let mut ws = WorksheetXml::default();
        set_auto_filter(&mut ws, "A1:D10").unwrap();

        assert!(ws.auto_filter.is_some());
        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:D10");
    }

    #[test]
    fn test_remove_auto_filter() {
        let mut ws = WorksheetXml::default();
        set_auto_filter(&mut ws, "A1:D10").unwrap();
        remove_auto_filter(&mut ws);

        assert!(ws.auto_filter.is_none());
    }

    #[test]
    fn test_auto_filter_xml_roundtrip() {
        let mut ws = WorksheetXml::default();
        set_auto_filter(&mut ws, "A1:C100").unwrap();

        let xml = quick_xml::se::to_string(&ws).unwrap();
        assert!(xml.contains("autoFilter"));
        assert!(xml.contains("A1:C100"));

        let parsed: WorksheetXml = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.auto_filter.is_some());
        assert_eq!(parsed.auto_filter.as_ref().unwrap().reference, "A1:C100");
    }

    #[test]
    fn test_remove_auto_filter_when_none() {
        let mut ws = WorksheetXml::default();
        remove_auto_filter(&mut ws);
        assert!(ws.auto_filter.is_none());
    }

    #[test]
    fn test_overwrite_auto_filter() {
        let mut ws = WorksheetXml::default();
        set_auto_filter(&mut ws, "A1:B10").unwrap();
        set_auto_filter(&mut ws, "A1:D20").unwrap();

        assert_eq!(ws.auto_filter.as_ref().unwrap().reference, "A1:D20");
    }

    #[test]
    fn test_table_config_creation() {
        let config = TableConfig {
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            range: "A1:D10".to_string(),
            columns: vec![
                TableColumn {
                    name: "Name".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Age".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "City".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Score".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
            ],
            show_header_row: true,
            style_name: Some("TableStyleMedium2".to_string()),
            auto_filter: true,
            ..TableConfig::default()
        };

        assert_eq!(config.name, "Table1");
        assert_eq!(config.columns.len(), 4);
        assert!(config.auto_filter);
    }

    #[test]
    fn test_build_table_xml() {
        let config = TableConfig {
            name: "Sales".to_string(),
            display_name: "Sales".to_string(),
            range: "A1:C5".to_string(),
            columns: vec![
                TableColumn {
                    name: "Product".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Quantity".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
                TableColumn {
                    name: "Price".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                },
            ],
            show_header_row: true,
            style_name: Some("TableStyleMedium2".to_string()),
            auto_filter: true,
            show_row_stripes: true,
            ..TableConfig::default()
        };

        let table_xml = build_table_xml(&config, 1);
        assert_eq!(table_xml.id, 1);
        assert_eq!(table_xml.name, "Sales");
        assert_eq!(table_xml.reference, "A1:C5");
        assert_eq!(table_xml.table_columns.count, 3);
        assert!(table_xml.auto_filter.is_some());
        assert!(table_xml.table_style_info.is_some());
        assert!(table_xml.header_row_count.is_none());
    }

    #[test]
    fn test_build_table_xml_no_header() {
        let config = TableConfig {
            name: "TestTable1".to_string(),
            display_name: "TestTable1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col1".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            show_header_row: false,
            ..TableConfig::default()
        };

        let table_xml = build_table_xml(&config, 2);
        assert_eq!(table_xml.header_row_count, Some(0));
    }

    #[test]
    fn test_table_xml_to_info() {
        let table_xml = TableXml {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            id: 1,
            name: "Inventory".to_string(),
            display_name: "Inventory".to_string(),
            reference: "A1:D20".to_string(),
            totals_row_count: None,
            totals_row_shown: None,
            header_row_count: None,
            auto_filter: Some(TableAutoFilter {
                reference: "A1:D20".to_string(),
                ..Default::default()
            }),
            table_columns: TableColumnsXml {
                count: 4,
                columns: vec![
                    TableColumnXml {
                        id: 1,
                        name: "Item".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                        ..Default::default()
                    },
                    TableColumnXml {
                        id: 2,
                        name: "Stock".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                        ..Default::default()
                    },
                    TableColumnXml {
                        id: 3,
                        name: "Price".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                        ..Default::default()
                    },
                    TableColumnXml {
                        id: 4,
                        name: "Supplier".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                        ..Default::default()
                    },
                ],
            },
            table_style_info: Some(TableStyleInfoXml {
                name: Some("TableStyleLight1".to_string()),
                show_first_column: Some(false),
                show_last_column: Some(false),
                show_row_stripes: Some(true),
                show_column_stripes: Some(false),
            }),
            ..Default::default()
        };

        let info = table_xml_to_info(&table_xml);
        assert_eq!(info.name, "Inventory");
        assert_eq!(info.display_name, "Inventory");
        assert_eq!(info.range, "A1:D20");
        assert!(info.show_header_row);
        assert!(info.auto_filter);
        assert_eq!(info.columns, vec!["Item", "Stock", "Price", "Supplier"]);
        assert_eq!(info.style_name, Some("TableStyleLight1".to_string()));
    }

    #[test]
    fn test_validate_table_config_empty_name() {
        let config = TableConfig {
            name: String::new(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Col".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        assert!(validate_table_config(&config).is_err());
    }

    #[test]
    fn test_validate_table_config_empty_range() {
        let config = TableConfig {
            name: "TestTable1".to_string(),
            range: String::new(),
            columns: vec![TableColumn {
                name: "Col".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        assert!(validate_table_config(&config).is_err());
    }

    #[test]
    fn test_validate_table_config_no_columns() {
        let config = TableConfig {
            name: "TestTable1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![],
            ..TableConfig::default()
        };
        assert!(validate_table_config(&config).is_err());
    }

    #[test]
    fn test_validate_table_config_valid() {
        let config = TableConfig {
            name: "TestTable1".to_string(),
            display_name: "TestTable1".to_string(),
            range: "A1:A5".to_string(),
            columns: vec![TableColumn {
                name: "Col".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };
        assert!(validate_table_config(&config).is_ok());
    }

    #[test]
    fn test_validate_table_config_rejects_range_width_mismatch() {
        let config = TableConfig {
            name: "TestTable1".to_string(),
            display_name: "TestTable1".to_string(),
            range: "A1:B5".to_string(),
            columns: vec![TableColumn {
                name: "Only column".to_string(),
                totals_row_function: None,
                totals_row_label: None,
            }],
            ..TableConfig::default()
        };

        assert!(matches!(
            validate_table_config(&config),
            Err(Error::InvalidArgument(_))
        ));
    }

    #[test]
    fn test_table_config_default() {
        let config = TableConfig::default();
        assert!(config.show_header_row);
        assert!(config.auto_filter);
        assert!(config.show_row_stripes);
        assert!(!config.show_first_column);
        assert!(!config.show_last_column);
        assert!(!config.show_column_stripes);
    }
}
