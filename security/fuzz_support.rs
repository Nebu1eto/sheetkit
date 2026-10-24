//! Shared stable-replay and cargo-fuzz entry points for security-sensitive parsers.

#![allow(dead_code)] // Each cargo-fuzz binary intentionally selects one shared entry point.

use std::io::Cursor;

use sheetkit_core::formula::{
    parse_formula, CellSnapshot, Evaluator, FormulaBudget, FormulaLimits,
};
use sheetkit_core::raw_transfer_write::{cells_to_raw_buffer, raw_buffer_to_cells};
use sheetkit_core::workbook::{OpenOptions, ReadMode, Workbook};
use sheetkit_xml::cell_ref::{parse_a1_reference, CompactCellRef};
use sheetkit_xml::worksheet::WorksheetXml;
use zip::read::Config;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

const MAX_ZIP_INPUT_BYTES: usize = 8 * 1024 * 1024;
const MAX_XML_INPUT_BYTES: usize = 1024 * 1024;
const MAX_FORMULA_INPUT_BYTES: usize = 65_537;
const MAX_TRANSFER_INPUT_BYTES: usize = 1024 * 1024;

/// Normalized result used to prove that replay is deterministic.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Outcome {
    Accepted,
    Rejected,
    HarnessLimit,
}

/// Exercises the workbook ZIP metadata and package-open boundary in both read modes.
pub fn exercise_zip_input(data: &[u8]) -> (Outcome, Outcome, Outcome) {
    if data.len() > MAX_ZIP_INPUT_BYTES {
        return (
            Outcome::HarnessLimit,
            Outcome::HarnessLimit,
            Outcome::HarnessLimit,
        );
    }

    let metadata = zip::ZipArchive::with_config(
        Config {
            max_entries: Some(4_096),
            ..Config::default()
        },
        Cursor::new(data),
    );
    let eager = OpenOptions::new().read_mode(ReadMode::Eager);
    let lazy = OpenOptions::new().read_mode(ReadMode::Lazy);
    (
        outcome(metadata),
        outcome(Workbook::open_from_buffer_with_options(data, &eager)),
        outcome(Workbook::open_from_buffer_with_options(data, &lazy)),
    )
}

/// Exercises direct A1 parsing and worksheet XML deserialization.
pub fn exercise_xml_input(data: &[u8]) -> (Outcome, Outcome, Outcome) {
    if data.len() > MAX_XML_INPUT_BYTES {
        return (
            Outcome::HarnessLimit,
            Outcome::HarnessLimit,
            Outcome::HarnessLimit,
        );
    }
    let Ok(text) = std::str::from_utf8(data) else {
        return (Outcome::Rejected, Outcome::Rejected, Outcome::Rejected);
    };

    (
        outcome(parse_a1_reference(text)),
        outcome(CompactCellRef::new(text)),
        outcome(quick_xml::de::from_str::<WorksheetXml>(text)),
    )
}

/// Exercises formula parsing and the budgeted evaluator entry point.
pub fn exercise_formula_input(data: &[u8]) -> (Outcome, Outcome) {
    if data.len() > MAX_FORMULA_INPUT_BYTES {
        return (Outcome::HarnessLimit, Outcome::HarnessLimit);
    }
    let Ok(text) = std::str::from_utf8(data) else {
        return (Outcome::Rejected, Outcome::Rejected);
    };
    let parsed = parse_formula(text);
    let parse_outcome = if parsed.is_ok() {
        Outcome::Accepted
    } else {
        Outcome::Rejected
    };
    let Ok(expression) = parsed else {
        return (parse_outcome, Outcome::Rejected);
    };
    let Ok(budget) = FormulaBudget::for_calculation(FormulaLimits::secure_default()) else {
        return (parse_outcome, Outcome::Rejected);
    };
    let provider = CellSnapshot::new("Sheet1".to_string());
    let mut evaluator = Evaluator::with_budget(&provider, budget);
    (parse_outcome, outcome(evaluator.eval_expr(&expression)))
}

/// Exercises both transfer versions through the shared Rust decoder.
pub fn exercise_transfer_input(data: &[u8]) -> Outcome {
    if data.len() > MAX_TRANSFER_INPUT_BYTES {
        return Outcome::HarnessLimit;
    }
    let Ok(rows) = raw_buffer_to_cells(data) else {
        return Outcome::Rejected;
    };
    let Ok(canonical) = cells_to_raw_buffer(&rows) else {
        return Outcome::Rejected;
    };
    outcome(raw_buffer_to_cells(&canonical))
}

/// Expands a small text recipe or hexadecimal corpus seed into target bytes.
pub fn materialize_seed(seed: &[u8]) -> Vec<u8> {
    let Ok(text) = std::str::from_utf8(seed) else {
        return seed.to_vec();
    };
    let text = text.trim();
    if let Some(hex) = text.strip_prefix("hex:") {
        return decode_hex(hex).unwrap_or_else(|| seed.to_vec());
    }
    if text == "recipe:zip:minimal-workbook" {
        return Workbook::new()
            .save_to_buffer()
            .expect("the canonical workbook corpus recipe must serialize");
    }
    if let Some(value) = text.strip_prefix("recipe:zip:entries:") {
        if let Ok(count) = value.parse::<usize>() {
            return zip_with_entries(count.min(4_097));
        }
    }
    if let Some(value) = text.strip_prefix("recipe:zip:zip64-sector:") {
        if let Ok(size) = value.parse::<usize>() {
            return zip_with_zip64_sector(size.min(1024 * 1024 + 1));
        }
    }
    if text == "recipe:zip:misleading-declared-size" {
        return zip_with_misleading_declared_size();
    }
    if let Some(value) = text.strip_prefix("recipe:xml:cell-ref:") {
        return worksheet_with_cell_ref(value).into_bytes();
    }
    if let Some(value) = text.strip_prefix("recipe:xml:cell-ref-digits:") {
        if let Ok(count) = value.parse::<usize>() {
            return worksheet_with_cell_ref(&format!("A{}", "9".repeat(count.min(65_536))))
                .into_bytes();
        }
    }
    if let Some(value) = text.strip_prefix("recipe:formula:parentheses:") {
        if let Ok(count) = value.parse::<usize>() {
            let count = count.min(129);
            return format!("{}1{}", "(".repeat(count), ")".repeat(count)).into_bytes();
        }
    }
    if let Some(value) = text.strip_prefix("recipe:formula:binary-chain:") {
        if let Ok(count) = value.parse::<usize>() {
            return std::iter::repeat_n("1", count.min(514).saturating_add(1))
                .collect::<Vec<_>>()
                .join("+")
                .into_bytes();
        }
    }
    if text == "recipe:transfer:v1-max-cell-count" {
        return transfer_max_cell_count(false);
    }
    if text == "recipe:transfer:v2-max-cell-count" {
        return transfer_max_cell_count(true);
    }
    if text == "recipe:transfer:v2-duplicate-rows" {
        return transfer_v2_rows(&[1, 1]);
    }
    if text == "recipe:transfer:v2-descending-rows" {
        return transfer_v2_rows(&[2, 1]);
    }
    if text == "recipe:transfer:v2-duplicate-columns" {
        return transfer_duplicate_columns(true);
    }
    if text == "recipe:transfer:v1-sparse-duplicate-columns" {
        return transfer_duplicate_columns(false);
    }
    seed.to_vec()
}

fn outcome<T, E>(result: Result<T, E>) -> Outcome {
    if result.is_ok() {
        Outcome::Accepted
    } else {
        Outcome::Rejected
    }
}

fn decode_hex(input: &str) -> Option<Vec<u8>> {
    let digits: Vec<_> = input
        .bytes()
        .filter(|byte| !byte.is_ascii_whitespace())
        .collect();
    if digits.len() % 2 != 0 {
        return None;
    }
    digits
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| {
            let high = (pair[0] as char).to_digit(16)?;
            let low = (pair[1] as char).to_digit(16)?;
            Some(((high << 4) | low) as u8)
        })
        .collect()
}

fn zip_with_entries(count: usize) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..count {
        writer
            .start_file(format!("entry-{index}"), SimpleFileOptions::default())
            .expect("bounded corpus ZIP entry must serialize");
    }
    writer
        .finish()
        .expect("bounded corpus ZIP must finish")
        .into_inner()
}

fn zip_with_zip64_sector(size: usize) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    writer.set_raw_zip64_comment(Some(vec![0; size].into_boxed_slice()));
    writer
        .start_file("entry", SimpleFileOptions::default())
        .expect("bounded corpus ZIP entry must serialize");
    writer
        .finish()
        .expect("bounded ZIP64 corpus must finish")
        .into_inner()
}

fn zip_with_misleading_declared_size() -> Vec<u8> {
    let mut bytes = Workbook::new()
        .save_to_buffer()
        .expect("the canonical workbook corpus recipe must serialize");
    let signature = [0x50, 0x4b, 0x01, 0x02];
    let offset = bytes
        .windows(signature.len())
        .position(|window| window == signature)
        .expect("workbook ZIP must contain a central-directory entry");
    bytes[offset + 24..offset + 28].copy_from_slice(&u32::MAX.to_le_bytes());
    bytes
}

fn worksheet_with_cell_ref(cell_ref: &str) -> String {
    let row = parse_a1_reference(cell_ref)
        .map(|reference| reference.row())
        .unwrap_or(1);
    format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="{row}"><c r="{cell_ref}"><v>1</v></c></row></sheetData></worksheet>"#
    )
}

fn transfer_max_cell_count(v2: bool) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(26);
    bytes.extend_from_slice(if v2 { b"2RKS" } else { b"DRKS" });
    bytes.extend_from_slice(&(if v2 { 2_u16 } else { 1_u16 }).to_le_bytes());
    bytes.extend_from_slice(&1_u32.to_le_bytes());
    bytes.extend_from_slice(&1_u16.to_le_bytes());
    bytes.extend_from_slice(&(if v2 { 0_u32 } else { 1_u32 }).to_le_bytes());
    bytes.extend_from_slice(&1_u32.to_le_bytes());
    bytes.extend_from_slice(&0_u32.to_le_bytes());
    if !v2 {
        bytes.extend_from_slice(&0_u32.to_le_bytes());
        bytes.extend_from_slice(&0_u32.to_le_bytes());
    }
    bytes.extend_from_slice(&u16::MAX.to_le_bytes());
    bytes
}

fn transfer_v2_rows(rows: &[u32]) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"2RKS");
    bytes.extend_from_slice(&2_u16.to_le_bytes());
    bytes.extend_from_slice(&(rows.len() as u32).to_le_bytes());
    bytes.extend_from_slice(&1_u16.to_le_bytes());
    bytes.extend_from_slice(&0_u32.to_le_bytes());
    for (index, row) in rows.iter().enumerate() {
        bytes.extend_from_slice(&row.to_le_bytes());
        bytes.extend_from_slice(&((index * 2) as u32).to_le_bytes());
    }
    for _ in rows {
        bytes.extend_from_slice(&0_u16.to_le_bytes());
    }
    bytes
}

fn transfer_duplicate_columns(v2: bool) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(if v2 { b"2RKS" } else { b"DRKS" });
    bytes.extend_from_slice(&(if v2 { 2_u16 } else { 1_u16 }).to_le_bytes());
    bytes.extend_from_slice(&1_u32.to_le_bytes());
    bytes.extend_from_slice(&2_u16.to_le_bytes());
    bytes.extend_from_slice(&(if v2 { 0_u32 } else { 1_u32 }).to_le_bytes());
    bytes.extend_from_slice(&1_u32.to_le_bytes());
    bytes.extend_from_slice(&0_u32.to_le_bytes());
    if !v2 {
        bytes.extend_from_slice(&0_u32.to_le_bytes());
        bytes.extend_from_slice(&0_u32.to_le_bytes());
    }
    bytes.extend_from_slice(&2_u16.to_le_bytes());
    for _ in 0..2 {
        bytes.extend_from_slice(&0_u16.to_le_bytes());
        bytes.push(0);
        if !v2 {
            bytes.extend_from_slice(&[0; 8]);
        }
    }
    bytes
}
