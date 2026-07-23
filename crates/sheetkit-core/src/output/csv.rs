use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::output::budget::OutputBudget;
use crate::resource::ResourceOperation;
use crate::rich_text::RichTextRun;
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS};

/// Maximum bytes returned by one CSV serialization.
///
/// This ceiling is deliberately finite and cannot be raised by callers.
pub const MAX_CSV_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;

/// Policy for text that spreadsheet applications may interpret as a formula.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CsvFormulaPolicy {
    /// Prefix formula-capable textual fields with an apostrophe.
    #[default]
    SpreadsheetSafe,
    /// Preserve textual fields byte-for-byte for non-spreadsheet exchange.
    RawExchange,
}

/// Serializes sparse typed rows as spreadsheet-safe CSV.
///
/// The result is staged completely before it is returned. Records are
/// separated by LF without a trailing terminator. Formula cells export their
/// expressions regardless of cached results, and formula-capable textual
/// fields are neutralized before quoting.
/// Row and column coordinates must be in-grid, strictly increasing, and unique.
pub fn serialize_csv(rows: &[(u32, Vec<(u32, CellValue)>)]) -> Result<Vec<u8>> {
    serialize_csv_with_policy(rows, CsvFormulaPolicy::SpreadsheetSafe)
}

/// Serializes sparse typed rows as RFC 4180 CSV with an explicit policy.
///
/// Select [`CsvFormulaPolicy::RawExchange`] only for consumers that will not
/// interpret fields as spreadsheet formulas.
/// Row and column coordinates must be in-grid, strictly increasing, and unique.
pub fn serialize_csv_with_policy(
    rows: &[(u32, Vec<(u32, CellValue)>)],
    formula_policy: CsvFormulaPolicy,
) -> Result<Vec<u8>> {
    serialize_csv_with_limit(rows, formula_policy, MAX_CSV_OUTPUT_BYTES)
}

fn serialize_csv_with_limit(
    rows: &[(u32, Vec<(u32, CellValue)>)],
    formula_policy: CsvFormulaPolicy,
    limit: u64,
) -> Result<Vec<u8>> {
    let mut budget = OutputBudget::new(ResourceOperation::Export, limit);
    serialize_csv_with_reserve(rows, formula_policy, |output, bytes| {
        budget.reserve(output, bytes)
    })
}

fn serialize_csv_with_reserve(
    rows: &[(u32, Vec<(u32, CellValue)>)],
    formula_policy: CsvFormulaPolicy,
    mut reserve: impl FnMut(&mut Vec<u8>, u64) -> Result<()>,
) -> Result<Vec<u8>> {
    let Some(max_column) = validate_csv_coordinates(rows)? else {
        return Ok(Vec::new());
    };

    let output_bytes = measure_csv_output(rows, max_column, formula_policy);
    let mut output = Vec::new();
    reserve(&mut output, output_bytes)?;

    for (row_index, (_, cells)) in rows.iter().enumerate() {
        if row_index > 0 {
            output.push(b'\n');
        }
        let mut cell_index = 0usize;
        for column in 1..=max_column {
            if column > 1 {
                output.push(b',');
            }

            while cells
                .get(cell_index)
                .is_some_and(|(cell_column, _)| *cell_column < column)
            {
                cell_index += 1;
            }
            if let Some((cell_column, value)) = cells.get(cell_index) {
                if *cell_column == column {
                    append_field(&mut output, value, formula_policy);
                    cell_index += 1;
                }
            }
        }
    }

    debug_assert_eq!(u64::try_from(output.len()).ok(), Some(output_bytes));
    Ok(output)
}

fn measure_csv_output(
    rows: &[(u32, Vec<(u32, CellValue)>)],
    max_column: u32,
    formula_policy: CsvFormulaPolicy,
) -> u64 {
    let separators = u64::try_from(rows.len().saturating_sub(1)).unwrap_or(u64::MAX);
    let structural_per_row = u64::from(max_column.saturating_sub(1));
    let mut output_bytes = structural_per_row
        .saturating_mul(u64::try_from(rows.len()).unwrap_or(u64::MAX))
        .saturating_add(separators);
    for (_, cells) in rows {
        for (_, value) in cells {
            let field = CsvField::from_cell(value);
            let mut stats = FieldStats::default();
            field.for_each_segment(|segment| stats.observe(segment));
            output_bytes = output_bytes.saturating_add(
                stats.output_bytes(
                    formula_policy == CsvFormulaPolicy::SpreadsheetSafe
                        && field.is_textual()
                        && stats
                            .first_significant
                            .is_some_and(is_formula_trigger_character),
                ),
            );
        }
    }
    output_bytes
}

fn validate_csv_coordinates(rows: &[(u32, Vec<(u32, CellValue)>)]) -> Result<Option<u32>> {
    let mut previous_row = 0u32;
    let mut max_column = None;

    for (row, cells) in rows {
        if *row == 0 || *row > MAX_ROWS {
            return Err(Error::InvalidRowNumber(*row));
        }
        if *row <= previous_row {
            return Err(Error::Internal(
                "CSV row coordinates must be strictly increasing and unique".to_string(),
            ));
        }
        previous_row = *row;

        let mut previous_column = 0u32;
        for (column, _) in cells {
            if *column == 0 || *column > MAX_COLUMNS {
                return Err(Error::InvalidColumnNumber(*column));
            }
            if *column <= previous_column {
                return Err(Error::Internal(
                    "CSV column coordinates must be strictly increasing and unique".to_string(),
                ));
            }
            previous_column = *column;
            max_column = Some(max_column.unwrap_or(0).max(*column));
        }
    }

    Ok(max_column)
}

enum CsvField<'a> {
    Borrowed { text: &'a str, textual: bool },
    Owned { text: String, textual: bool },
    Rich(&'a [RichTextRun]),
    Formula(&'a str),
}

impl<'a> CsvField<'a> {
    fn from_cell(value: &'a CellValue) -> Self {
        match value {
            CellValue::Formula { expr, .. } => Self::Formula(expr),
            CellValue::Empty => Self::Borrowed {
                text: "",
                textual: false,
            },
            CellValue::Bool(value) => Self::Borrowed {
                text: if *value { "TRUE" } else { "FALSE" },
                textual: false,
            },
            CellValue::Number(_) | CellValue::Date(_) => Self::Owned {
                text: value.to_string(),
                textual: false,
            },
            CellValue::String(text) => Self::Borrowed {
                text,
                textual: true,
            },
            CellValue::Error(text) => Self::Borrowed {
                text,
                textual: !is_excel_error_token(text),
            },
            CellValue::RichString(runs) => Self::Rich(runs),
        }
    }

    fn for_each_segment(&self, mut visit: impl FnMut(&str)) {
        match self {
            Self::Borrowed { text, .. } => visit(text),
            Self::Owned { text, .. } => visit(text),
            Self::Rich(runs) => {
                for run in *runs {
                    visit(&run.text);
                }
            }
            Self::Formula(expression) => {
                visit("=");
                visit(expression);
            }
        }
    }

    fn is_textual(&self) -> bool {
        match self {
            Self::Borrowed { textual, .. } | Self::Owned { textual, .. } => *textual,
            Self::Rich(_) | Self::Formula(_) => true,
        }
    }
}

#[derive(Default)]
struct FieldStats {
    content_bytes: u64,
    quote_count: u64,
    needs_quotes: bool,
    first_significant: Option<char>,
}

impl FieldStats {
    fn observe(&mut self, segment: &str) {
        self.content_bytes = self
            .content_bytes
            .saturating_add(u64::try_from(segment.len()).unwrap_or(u64::MAX));
        self.quote_count = self.quote_count.saturating_add(
            u64::try_from(segment.bytes().filter(|byte| *byte == b'"').count()).unwrap_or(u64::MAX),
        );
        self.needs_quotes |= segment
            .bytes()
            .any(|byte| matches!(byte, b',' | b'"' | b'\r' | b'\n'));
        if self.first_significant.is_none() {
            self.first_significant = segment
                .chars()
                .find(|character| !is_formula_leading_character(*character));
        }
    }

    fn output_bytes(&self, neutralized: bool) -> u64 {
        let neutralizer = u64::from(neutralized);
        let structural = if self.needs_quotes { 2 } else { 0 };
        self.content_bytes
            .checked_add(self.quote_count)
            .and_then(|value| value.checked_add(neutralizer))
            .and_then(|value| value.checked_add(structural))
            .unwrap_or(u64::MAX)
    }
}

fn append_field(output: &mut Vec<u8>, value: &CellValue, formula_policy: CsvFormulaPolicy) {
    let field = CsvField::from_cell(value);
    let mut stats = FieldStats::default();
    field.for_each_segment(|segment| stats.observe(segment));
    let neutralized = formula_policy == CsvFormulaPolicy::SpreadsheetSafe
        && field.is_textual()
        && stats
            .first_significant
            .is_some_and(is_formula_trigger_character);
    if stats.needs_quotes {
        output.push(b'"');
    }
    if neutralized {
        output.push(b'\'');
    }
    field.for_each_segment(|segment| append_segment(output, segment, stats.needs_quotes));
    if stats.needs_quotes {
        output.push(b'"');
    }
}

fn append_segment(output: &mut Vec<u8>, segment: &str, quoted: bool) {
    if !quoted {
        output.extend_from_slice(segment.as_bytes());
        return;
    }

    let mut start = 0usize;
    for (index, byte) in segment.bytes().enumerate() {
        if byte == b'"' {
            output.extend_from_slice(&segment.as_bytes()[start..index]);
            output.extend_from_slice(b"\"\"");
            start = index + 1;
        }
    }
    output.extend_from_slice(&segment.as_bytes()[start..]);
}

fn is_formula_leading_character(character: char) -> bool {
    character.is_whitespace() || character.is_control() || character == '\u{feff}'
}

fn is_formula_trigger_character(character: char) -> bool {
    matches!(character, '=' | '+' | '-' | '@')
}

pub(crate) fn is_excel_error_token(value: &str) -> bool {
    matches!(
        value,
        "#NULL!"
            | "#DIV/0!"
            | "#VALUE!"
            | "#REF!"
            | "#NAME?"
            | "#NUM!"
            | "#N/A"
            | "#GETTING_DATA"
            | "#SPILL!"
            | "#CALC!"
            | "#FIELD!"
            | "#BLOCKED!"
            | "#UNKNOWN!"
            | "#CONNECT!"
            | "#BUSY!"
            | "#PYTHON!"
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::resource::ResourceDimension;

    fn one_cell(value: &str) -> Vec<(u32, Vec<(u32, CellValue)>)> {
        vec![(1, vec![(1, CellValue::String(value.to_string()))])]
    }

    fn metadata(error: Error) -> crate::ResourceLimitError {
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        metadata
    }

    #[test]
    fn exact_output_limit_succeeds_and_plus_one_is_rejected() {
        let rows = one_cell("a");
        assert_eq!(
            serialize_csv_with_limit(&rows, CsvFormulaPolicy::SpreadsheetSafe, 1).unwrap(),
            b"a"
        );

        let error =
            serialize_csv_with_limit(&rows, CsvFormulaPolicy::SpreadsheetSafe, 0).unwrap_err();
        let metadata = metadata(error);
        assert_eq!(metadata.limit(), 0);
        assert_eq!(metadata.observed(), 1);
        assert_eq!(metadata.operation(), ResourceOperation::Export);
        assert_eq!(metadata.category(), crate::ResourceCategory::Output);
        assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
        assert_eq!(metadata.stage(), Some(crate::ResourceStage::Serialize));
    }

    #[test]
    fn formula_cells_export_expressions_regardless_of_cached_results() {
        let rows = vec![
            (
                1,
                vec![(
                    1,
                    CellValue::Formula {
                        expr: "1+1".to_string(),
                        result: Some(Box::new(CellValue::Number(2.0))),
                    },
                )],
            ),
            (
                2,
                vec![(
                    1,
                    CellValue::Formula {
                        expr: "2+2".to_string(),
                        result: None,
                    },
                )],
            ),
        ];

        assert_eq!(
            serialize_csv(&rows).unwrap(),
            b"'=1+1\n'=2+2",
            "default CSV bytes must match the Node and CLI contract"
        );
        assert_eq!(
            serialize_csv_with_policy(&rows, CsvFormulaPolicy::RawExchange).unwrap(),
            b"=1+1\n=2+2"
        );
    }

    #[test]
    fn rejection_returns_no_partial_output_value() {
        let rows = one_cell("=dangerous");
        assert!(serialize_csv_with_limit(&rows, CsvFormulaPolicy::SpreadsheetSafe, 1).is_err());
    }

    #[test]
    fn serialization_reserves_the_exact_output_once_before_emission() {
        let rows = vec![(
            1,
            (1..=4_096)
                .map(|column| (column, CellValue::String("value".to_string())))
                .collect(),
        )];
        let mut reserve_calls = Vec::new();

        let output = serialize_csv_with_reserve(
            &rows,
            CsvFormulaPolicy::SpreadsheetSafe,
            |output, bytes| {
                reserve_calls.push(bytes);
                output
                    .try_reserve_exact(usize::try_from(bytes).unwrap())
                    .map_err(|_| Error::Internal("test allocation failed".to_string()))
            },
        )
        .unwrap();

        assert_eq!(reserve_calls, [u64::try_from(output.len()).unwrap()]);
    }
}
