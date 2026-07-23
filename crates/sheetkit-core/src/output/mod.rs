//! Bounded, context-specific output serializers.

mod budget;
mod csv;

pub(crate) use csv::is_excel_error_token;
pub use csv::{serialize_csv, serialize_csv_with_policy, CsvFormulaPolicy, MAX_CSV_OUTPUT_BYTES};
