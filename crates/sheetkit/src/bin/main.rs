use std::io::Write;
use std::path::PathBuf;
use std::process;

use clap::{Parser, Subcommand, ValueEnum};
use sheetkit::Workbook;

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
    },
    /// List all sheet names in the workbook.
    Sheets {
        /// Path to the .xlsx file.
        file: PathBuf,
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
    },
    /// Get a single cell value.
    Get {
        /// Path to the .xlsx file.
        file: PathBuf,
        /// Sheet name.
        sheet: String,
        /// Cell reference (e.g. A1, B2).
        cell: String,
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
    },
}

#[derive(Clone, ValueEnum)]
enum OutputFormat {
    /// Tab-separated table output.
    Table,
    /// Comma-separated values.
    Csv,
}

#[derive(Clone, ValueEnum)]
enum ConvertFormat {
    /// Comma-separated values.
    Csv,
}

fn main() {
    let cli = Cli::parse();
    if let Err(e) = run(cli) {
        eprintln!("Error: {e}");
        process::exit(1);
    }
}

fn run(cli: Cli) -> Result<(), Box<dyn std::error::Error>> {
    match cli.command {
        Commands::Info { file } => cmd_info(&file),
        Commands::Sheets { file } => cmd_sheets(&file),
        Commands::Read {
            file,
            sheet,
            format,
        } => cmd_read(&file, sheet.as_deref(), &format),
        Commands::Get { file, sheet, cell } => cmd_get(&file, &sheet, &cell),
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
        } => cmd_convert(&file, sheet.as_deref(), &format, &output),
    }
}

fn cmd_info(file: &PathBuf) -> Result<(), Box<dyn std::error::Error>> {
    let wb = Workbook::open(file)?;
    let sheets = wb.sheet_names();
    let active = wb.get_active_sheet();
    let props = wb.get_doc_props();

    println!("File: {}", file.display());
    println!("Sheets: {}", sheets.len());
    for (i, name) in sheets.iter().enumerate() {
        let marker = if *name == active { " (active)" } else { "" };
        println!("  {}: {}{}", i + 1, name, marker);
    }

    if let Some(title) = &props.title {
        if !title.is_empty() {
            println!("Title: {title}");
        }
    }
    if let Some(creator) = &props.creator {
        if !creator.is_empty() {
            println!("Creator: {creator}");
        }
    }
    if let Some(modified) = &props.modified {
        if !modified.is_empty() {
            println!("Modified: {modified}");
        }
    }

    Ok(())
}

fn cmd_sheets(file: &PathBuf) -> Result<(), Box<dyn std::error::Error>> {
    let wb = Workbook::open(file)?;
    for name in wb.sheet_names() {
        println!("{name}");
    }
    Ok(())
}

fn cmd_read(
    file: &PathBuf,
    sheet: Option<&str>,
    format: &OutputFormat,
) -> Result<(), Box<dyn std::error::Error>> {
    let wb = Workbook::open(file)?;
    let sheet_name = sheet.unwrap_or_else(|| wb.get_active_sheet());
    let rows = wb.get_rows(sheet_name)?;

    if rows.is_empty() {
        return Ok(());
    }

    // Determine the maximum column index to produce a dense grid.
    let max_col = rows
        .iter()
        .flat_map(|(_, cells)| cells.iter().map(|(c, _)| *c))
        .max()
        .unwrap_or(0);

    let separator = match format {
        OutputFormat::Table => "\t",
        OutputFormat::Csv => ",",
    };

    let stdout = std::io::stdout();
    let mut out = stdout.lock();

    for (_, cells) in &rows {
        let mut line = String::new();
        for col in 1..=max_col {
            if col > 1 {
                line.push_str(separator);
            }
            let val = cells
                .iter()
                .find(|(c, _)| *c == col)
                .map(|(_, v)| format_cell_for_output(v, format))
                .unwrap_or_default();
            line.push_str(&val);
        }
        writeln!(out, "{line}")?;
    }

    Ok(())
}

fn cmd_get(file: &PathBuf, sheet: &str, cell: &str) -> Result<(), Box<dyn std::error::Error>> {
    let wb = Workbook::open(file)?;
    let value = wb.get_cell_value(sheet, cell)?;
    let display = value.to_string();
    if !display.is_empty() {
        println!("{display}");
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
    output: &PathBuf,
) -> Result<(), Box<dyn std::error::Error>> {
    let wb = Workbook::open(file)?;
    let sheet_name = sheet.unwrap_or_else(|| wb.get_active_sheet());
    let rows = wb.get_rows(sheet_name)?;

    match format {
        ConvertFormat::Csv => write_csv(&rows, output)?,
    }

    Ok(())
}

fn write_csv(
    rows: &[(u32, Vec<(u32, sheetkit::CellValue)>)],
    output: &PathBuf,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut file = std::fs::File::create(output)?;

    if rows.is_empty() {
        return Ok(());
    }

    let max_col = rows
        .iter()
        .flat_map(|(_, cells)| cells.iter().map(|(c, _)| *c))
        .max()
        .unwrap_or(0);

    for (_, cells) in rows {
        let mut line = String::new();
        for col in 1..=max_col {
            if col > 1 {
                line.push(',');
            }
            let val = cells
                .iter()
                .find(|(c, _)| *c == col)
                .map(|(_, v)| csv_escape(&v.to_string()))
                .unwrap_or_default();
            line.push_str(&val);
        }
        writeln!(file, "{line}")?;
    }

    Ok(())
}

/// Format a cell value for the given output format.
fn format_cell_for_output(value: &sheetkit::CellValue, format: &OutputFormat) -> String {
    let s = value.to_string();
    match format {
        OutputFormat::Table => s,
        OutputFormat::Csv => csv_escape(&s),
    }
}

/// Escape a string for CSV output. Wraps in quotes if it contains commas,
/// quotes, or newlines.
fn csv_escape(s: &str) -> String {
    if s.contains(',') || s.contains('"') || s.contains('\n') || s.contains('\r') {
        format!("\"{}\"", s.replace('"', "\"\""))
    } else {
        s.to_string()
    }
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
