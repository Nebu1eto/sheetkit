use std::ffi::OsString;
use std::fmt::Write as FmtWrite;
use std::io::Write as IoWrite;
use std::path::{Path, PathBuf};
use std::process;

use clap::error::{ContextKind, ContextValue, ErrorKind};
use clap::{Parser, Subcommand, ValueEnum};
use sheetkit::Workbook;

mod output;

use output::{
    append_terminal_safe_cell, cell_value_json, info_json, is_terminal_control, json_bytes,
    rows_json, terminal_safe, terminal_safe_cell_len, TerminalOutput, MAX_CLI_OUTPUT_BYTES,
};

#[derive(Parser)]
#[command(
    name = "sheetkit",
    version,
    about = "Excel (.xlsx) file toolkit",
    long_about = "A command-line tool for reading, writing, and converting Excel (.xlsx) files."
)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Show workbook information (sheets, dimensions, properties).
    Info {
        /// Path to the .xlsx file.
        file: PathBuf,
        /// Output format.
        #[arg(short, long, default_value = "text")]
        format: HumanOutputFormat,
    },
    /// List all sheet names in the workbook.
    Sheets {
        /// Path to the .xlsx file.
        file: PathBuf,
        /// Output format.
        #[arg(short, long, default_value = "text")]
        format: HumanOutputFormat,
    },
    /// Read and display sheet data.
    Read {
        /// Path to the .xlsx file.
        file: PathBuf,
        /// Sheet name to read. Defaults to the active sheet.
        #[arg(short, long)]
        sheet: Option<String>,
        /// Output format.
        #[arg(short, long, default_value = "table")]
        format: OutputFormat,
        /// Formula handling for CSV fields.
        #[arg(long, default_value = "spreadsheet-safe")]
        csv_formula_policy: CliCsvFormulaPolicy,
    },
    /// Get a single cell value.
    Get {
        /// Path to the .xlsx file.
        file: PathBuf,
        /// Sheet name.
        sheet: String,
        /// Cell reference (e.g. A1, B2).
        cell: String,
        /// Output format.
        #[arg(short, long, default_value = "text")]
        format: HumanOutputFormat,
    },
    /// Set a cell value and write to a new file.
    Set {
        /// Path to the input .xlsx file.
        file: PathBuf,
        /// Sheet name.
        sheet: String,
        /// Cell reference (e.g. A1, B2).
        cell: String,
        /// Value to set.
        value: String,
        /// Output file path.
        #[arg(short, long)]
        output: PathBuf,
    },
    /// Convert a sheet to another format.
    Convert {
        /// Path to the .xlsx file.
        file: PathBuf,
        /// Target format.
        #[arg(short, long)]
        format: ConvertFormat,
        /// Output file path.
        #[arg(short, long)]
        output: PathBuf,
        /// Sheet name. Defaults to the active sheet.
        #[arg(short, long)]
        sheet: Option<String>,
        /// Formula handling for CSV fields.
        #[arg(long, default_value = "spreadsheet-safe")]
        csv_formula_policy: CliCsvFormulaPolicy,
    },
}

#[derive(Clone, ValueEnum)]
enum OutputFormat {
    /// Tab-separated table output.
    Table,
    /// Comma-separated values.
    Csv,
    /// Structured JSON.
    Json,
}

#[derive(Clone, ValueEnum)]
enum HumanOutputFormat {
    /// Terminal-safe human-readable text.
    Text,
    /// Structured JSON.
    Json,
}

#[derive(Clone, ValueEnum)]
enum ConvertFormat {
    /// Comma-separated values.
    Csv,
}

#[derive(Clone, Copy, ValueEnum)]
enum CliCsvFormulaPolicy {
    /// Neutralize formula-capable textual fields for spreadsheet use.
    SpreadsheetSafe,
    /// Preserve formula-capable text for non-spreadsheet exchange.
    RawExchange,
}

impl From<CliCsvFormulaPolicy> for sheetkit::CsvFormulaPolicy {
    fn from(value: CliCsvFormulaPolicy) -> Self {
        match value {
            CliCsvFormulaPolicy::SpreadsheetSafe => Self::SpreadsheetSafe,
            CliCsvFormulaPolicy::RawExchange => Self::RawExchange,
        }
    }
}

fn main() {
    let args: Vec<OsString> = std::env::args_os().collect();
    let cli = match Cli::try_parse_from(&args) {
        Ok(cli) => cli,
        Err(mut error) => {
            let is_display = matches!(
                error.kind(),
                ErrorKind::DisplayHelp | ErrorKind::DisplayVersion
            );
            if !is_display {
                sanitize_clap_error_context(&mut error);
            }
            let message = terminal_safe_clap_message(&error.to_string());
            if is_display {
                let _ = std::io::stdout().lock().write_all(message.as_bytes());
            } else {
                let _ = std::io::stderr().lock().write_all(message.as_bytes());
            }
            process::exit(error.exit_code());
        }
    };
    if let Err(e) = run(cli) {
        let message = format!("Error: {}\n", terminal_safe(&e.to_string()));
        let _ = std::io::stderr().lock().write_all(message.as_bytes());
        process::exit(1);
    }
}

fn sanitize_clap_error_context(error: &mut clap::Error) {
    const UNTRUSTED_CONTEXTS: &[ContextKind] = &[
        ContextKind::InvalidSubcommand,
        ContextKind::InvalidArg,
        ContextKind::PriorArg,
        ContextKind::InvalidValue,
        ContextKind::SuggestedCommand,
        ContextKind::SuggestedSubcommand,
        ContextKind::SuggestedArg,
        ContextKind::SuggestedValue,
        ContextKind::TrailingArg,
        ContextKind::Suggested,
        ContextKind::Custom,
    ];

    for kind in UNTRUSTED_CONTEXTS {
        let replacement = match error.get(*kind) {
            Some(ContextValue::String(value)) => Some(ContextValue::String(terminal_safe(value))),
            Some(ContextValue::Strings(values)) => Some(ContextValue::Strings(
                values.iter().map(|value| terminal_safe(value)).collect(),
            )),
            Some(ContextValue::StyledStr(value)) => {
                Some(ContextValue::String(terminal_safe(&value.to_string())))
            }
            Some(ContextValue::StyledStrs(values)) => Some(ContextValue::Strings(
                values
                    .iter()
                    .map(|value| terminal_safe(&value.to_string()))
                    .collect(),
            )),
            _ => None,
        };
        if let Some(replacement) = replacement {
            error.insert(*kind, replacement);
        }
    }
}

fn terminal_safe_clap_message(message: &str) -> String {
    let mut output = String::with_capacity(message.len());
    for character in message.chars() {
        if character == '\n' {
            output.push(character);
        } else if is_terminal_control(character) {
            let code = u32::from(character);
            write!(output, "\\x{code:02X}").expect("writing to a String cannot fail");
        } else {
            output.push(character);
        }
    }
    output
}

fn run(cli: Cli) -> Result<(), Box<dyn std::error::Error>> {
    match cli.command {
        Commands::Info { file, format } => cmd_info(&file, &format),
        Commands::Sheets { file, format } => cmd_sheets(&file, &format),
        Commands::Read {
            file,
            sheet,
            format,
            csv_formula_policy,
        } => cmd_read(&file, sheet.as_deref(), &format, csv_formula_policy.into()),
        Commands::Get {
            file,
            sheet,
            cell,
            format,
        } => cmd_get(&file, &sheet, &cell, &format),
        Commands::Set {
            file,
            sheet,
            cell,
            value,
            output,
        } => cmd_set(&file, &sheet, &cell, &value, &output),
        Commands::Convert {
            file,
            format,
            output,
            sheet,
            csv_formula_policy,
        } => cmd_convert(
            &file,
            sheet.as_deref(),
            &format,
            &output,
            csv_formula_policy.into(),
        ),
    }
}

fn cmd_info(file: &PathBuf, format: &HumanOutputFormat) -> Result<(), Box<dyn std::error::Error>> {
    let wb = Workbook::open(file)?;
    let sheets = wb.sheet_names();
    let active = wb.get_active_sheet();
    let props = wb.get_doc_props();

    if matches!(format, HumanOutputFormat::Json) {
        write_json(&info_json(file, &sheets, active, &props))?;
        return Ok(());
    }

    let mut output = TerminalOutput::new(MAX_CLI_OUTPUT_BYTES);
    output.push_literal("File: ")?;
    output.push_safe(&file.to_string_lossy())?;
    output.push_literal("\nSheets: ")?;
    output.push_literal(&sheets.len().to_string())?;
    output.push_literal("\n")?;
    for (i, name) in sheets.iter().enumerate() {
        let marker = if *name == active { " (active)" } else { "" };
        output.push_literal("  ")?;
        output.push_literal(&(i + 1).to_string())?;
        output.push_literal(": ")?;
        output.push_safe(name)?;
        output.push_literal(marker)?;
        output.push_literal("\n")?;
    }

    if let Some(title) = &props.title {
        if !title.is_empty() {
            output.push_literal("Title: ")?;
            output.push_safe(title)?;
            output.push_literal("\n")?;
        }
    }
    if let Some(creator) = &props.creator {
        if !creator.is_empty() {
            output.push_literal("Creator: ")?;
            output.push_safe(creator)?;
            output.push_literal("\n")?;
        }
    }
    if let Some(modified) = &props.modified {
        if !modified.is_empty() {
            output.push_literal("Modified: ")?;
            output.push_safe(modified)?;
            output.push_literal("\n")?;
        }
    }
    write_stdout(&output.finish())
}

fn cmd_sheets(
    file: &PathBuf,
    format: &HumanOutputFormat,
) -> Result<(), Box<dyn std::error::Error>> {
    let wb = Workbook::open(file)?;
    let sheets = wb.sheet_names();
    if matches!(format, HumanOutputFormat::Json) {
        write_json(&sheets)?;
        return Ok(());
    }
    let mut output = TerminalOutput::new(MAX_CLI_OUTPUT_BYTES);
    for name in sheets {
        output.push_safe(name)?;
        output.push_literal("\n")?;
    }
    write_stdout(&output.finish())
}

fn cmd_read(
    file: &PathBuf,
    sheet: Option<&str>,
    format: &OutputFormat,
    csv_formula_policy: sheetkit::CsvFormulaPolicy,
) -> Result<(), Box<dyn std::error::Error>> {
    let wb = Workbook::open(file)?;
    let sheet_name = sheet.unwrap_or_else(|| wb.get_active_sheet());
    let rows = wb.get_rows_with_limits(sheet_name, cli_row_limits()?)?;

    if matches!(format, OutputFormat::Csv) {
        let csv = sheetkit::serialize_csv_with_policy(&rows, csv_formula_policy)?;
        return write_stdout(&csv);
    }
    if matches!(format, OutputFormat::Json) {
        write_json(&rows_json(&rows))?;
        return Ok(());
    }
    if rows.is_empty() {
        return Ok(());
    }

    // Determine the maximum column index to produce a dense grid.
    let max_col = rows
        .iter()
        .flat_map(|(_, cells)| cells.iter().map(|(c, _)| *c))
        .max()
        .unwrap_or(0);

    let output = render_table(&rows, max_col, MAX_CLI_OUTPUT_BYTES)?.0;
    write_stdout(output.as_bytes())
}

fn render_table(
    rows: &[(u32, Vec<(u32, sheetkit::CellValue)>)],
    max_column: u32,
    limit: u64,
) -> sheetkit::Result<(String, usize)> {
    let output_bytes = measure_table_output(rows, max_column, limit)?;
    let capacity = usize::try_from(output_bytes)
        .map_err(|_| sheetkit::Error::Internal("table output is not addressable".to_string()))?;
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_| sheetkit::Error::Internal("table output allocation failed".to_string()))?;
    let mut work = 0usize;
    for (_, cells) in rows {
        let mut previous_column = 0u32;
        for (column, _) in cells {
            if *column <= previous_column {
                return Err(sheetkit::Error::Internal(
                    "table cell coordinates must be strictly increasing and unique".to_string(),
                ));
            }
            previous_column = *column;
        }

        let mut cell_index = 0usize;
        for column in 1..=max_column {
            work = work.saturating_add(1);
            if column > 1 {
                output.push('\t');
            }
            if let Some((cell_column, value)) = cells.get(cell_index) {
                if *cell_column == column {
                    work = work.saturating_add(1);
                    append_terminal_safe_cell(&mut output, value);
                    cell_index += 1;
                }
            }
        }
        output.push('\n');
    }
    Ok((output, work))
}

fn measure_table_output(
    rows: &[(u32, Vec<(u32, sheetkit::CellValue)>)],
    max_column: u32,
    limit: u64,
) -> sheetkit::Result<u64> {
    let structural_per_row = u64::from(max_column.saturating_sub(1)).saturating_add(1);
    let mut observed =
        structural_per_row.saturating_mul(u64::try_from(rows.len()).unwrap_or(u64::MAX));
    for (_, cells) in rows {
        let mut previous_column = 0u32;
        for (column, value) in cells {
            if *column <= previous_column {
                return Err(sheetkit::Error::Internal(
                    "table cell coordinates must be strictly increasing and unique".to_string(),
                ));
            }
            previous_column = *column;
            observed = observed.saturating_add(terminal_safe_cell_len(value));
        }
    }
    if observed > limit {
        return Err(sheetkit::ResourceLimitError::exceeded(
            sheetkit::ResourceOperation::Export,
            sheetkit::ResourceCategory::Output,
            sheetkit::ResourceDimension::OutputBytes,
            limit,
            observed,
            Some(sheetkit::ResourceStage::Serialize),
        )
        .map(sheetkit::Error::from)
        .unwrap_or_else(|| sheetkit::Error::Internal("invalid table output limit".to_string())));
    }
    Ok(observed)
}

fn cmd_get(
    file: &PathBuf,
    sheet: &str,
    cell: &str,
    format: &HumanOutputFormat,
) -> Result<(), Box<dyn std::error::Error>> {
    let wb = Workbook::open(file)?;
    let value = wb.get_cell_value_with_limits(sheet, cell, cli_row_limits()?)?;
    if matches!(format, HumanOutputFormat::Json) {
        write_json(&cell_value_json(&value))?;
        return Ok(());
    }
    if !matches!(value, sheetkit::CellValue::Empty) {
        let mut output = TerminalOutput::new(MAX_CLI_OUTPUT_BYTES);
        output.push_cell_value(&value)?;
        output.push_literal("\n")?;
        return write_stdout(&output.finish());
    }
    Ok(())
}

fn cmd_set(
    file: &PathBuf,
    sheet: &str,
    cell: &str,
    value: &str,
    output: &PathBuf,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut wb = Workbook::open(file)?;
    // Try to parse as number or bool; fall back to string.
    let cell_value = parse_cell_input(value);
    wb.set_cell_value(sheet, cell, cell_value)?;
    wb.save(output)?;
    Ok(())
}

fn cmd_convert(
    file: &PathBuf,
    sheet: Option<&str>,
    format: &ConvertFormat,
    output: &Path,
    csv_formula_policy: sheetkit::CsvFormulaPolicy,
) -> Result<(), Box<dyn std::error::Error>> {
    let wb = Workbook::open(file)?;
    let sheet_name = sheet.unwrap_or_else(|| wb.get_active_sheet());
    let rows = wb.get_rows_with_limits(sheet_name, cli_row_limits()?)?;

    match format {
        ConvertFormat::Csv => {
            let csv = sheetkit::serialize_csv_with_policy(&rows, csv_formula_policy)?;
            publish_csv(output, &csv)?;
        }
    }

    Ok(())
}

fn publish_csv(output: &Path, csv: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    let parent = output
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut staged = tempfile::NamedTempFile::new_in(parent)?;
    staged.write_all(csv)?;
    staged.flush()?;
    staged.as_file().sync_all()?;
    staged.persist(output).map_err(|error| error.error)?;
    Ok(())
}

fn write_json(value: &(impl serde::Serialize + ?Sized)) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = json_bytes(value, MAX_CLI_OUTPUT_BYTES)?;
    write_stdout(&bytes)
}

fn cli_row_limits() -> sheetkit::Result<sheetkit::TraversalLimits> {
    let defaults = sheetkit::TraversalLimits::default();
    sheetkit::TraversalLimits::new(
        defaults.max_rows(),
        defaults.max_columns(),
        defaults.max_cells(),
        MAX_CLI_OUTPUT_BYTES,
    )
}

fn write_stdout(bytes: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    std::io::stdout().lock().write_all(bytes)?;
    Ok(())
}

/// Parse a user-provided string into an appropriate CellValue.
/// Recognizes booleans and numbers; otherwise treats as string.
fn parse_cell_input(input: &str) -> sheetkit::CellValue {
    match input.to_uppercase().as_str() {
        "TRUE" => return sheetkit::CellValue::Bool(true),
        "FALSE" => return sheetkit::CellValue::Bool(false),
        _ => {}
    }
    if let Ok(n) = input.parse::<f64>() {
        return sheetkit::CellValue::Number(n);
    }
    sheetkit::CellValue::String(input.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn wide_table_rendering_visits_each_column_and_occupied_cell_once() {
        let rows = vec![(
            1,
            vec![
                (1, sheetkit::CellValue::String("first".to_string())),
                (16_384, sheetkit::CellValue::String("last".to_string())),
            ],
        )];

        let (output, work) = render_table(&rows, 16_384, MAX_CLI_OUTPUT_BYTES).unwrap();

        assert_eq!(work, 16_384 + 2);
        assert!(output.starts_with("first\t"));
        assert!(output.ends_with("\tlast\n"));
    }

    #[test]
    fn table_rendering_rejects_noncanonical_cells() {
        let rows = vec![(
            1,
            vec![
                (2, sheetkit::CellValue::Empty),
                (2, sheetkit::CellValue::Empty),
            ],
        )];

        assert!(render_table(&rows, 2, MAX_CLI_OUTPUT_BYTES).is_err());
    }

    #[test]
    fn table_output_limit_accepts_exact_and_rejects_far_column_amplification() {
        let rows = vec![(
            1,
            vec![
                (1, sheetkit::CellValue::String("한글".to_string())),
                (16_384, sheetkit::CellValue::String("last".to_string())),
            ],
        )];
        let exact = measure_table_output(&rows, 16_384, MAX_CLI_OUTPUT_BYTES).unwrap();
        assert!(render_table(&rows, 16_384, exact).is_ok());

        let error = render_table(&rows, 16_384, exact - 1).unwrap_err();
        let sheetkit::Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.limit(), exact - 1);
        assert_eq!(metadata.observed(), exact);

        let far_rows = (1..=4_097)
            .map(|row| {
                (
                    row,
                    vec![
                        (1, sheetkit::CellValue::Empty),
                        (16_384, sheetkit::CellValue::Empty),
                    ],
                )
            })
            .collect::<Vec<_>>();
        let error = render_table(&far_rows, 16_384, MAX_CLI_OUTPUT_BYTES).unwrap_err();
        assert!(matches!(error, sheetkit::Error::ResourceLimitExceeded(_)));
    }
}
