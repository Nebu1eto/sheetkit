use std::fmt::Write as _;
use std::io::{self, Write as IoWrite};
use std::path::Path;

use serde::ser::{SerializeMap, SerializeSeq};
use serde::{Serialize, Serializer};
use sheetkit::{CellValue, DocProperties};

pub(crate) const MAX_CLI_OUTPUT_BYTES: u64 = sheetkit::MAX_CSV_OUTPUT_BYTES;

pub(crate) struct TerminalOutput {
    bytes: Vec<u8>,
    limit: u64,
    used: u64,
}

impl TerminalOutput {
    pub(crate) const fn new(limit: u64) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
            used: 0,
        }
    }

    pub(crate) fn push_literal(&mut self, value: &str) -> sheetkit::Result<()> {
        self.push_bytes(value.as_bytes())
    }

    pub(crate) fn push_safe(&mut self, value: &str) -> sheetkit::Result<()> {
        let additional = terminal_safe_len(value);
        self.charge_and_reserve(additional)?;
        for character in value.chars() {
            if is_terminal_control(character) {
                write!(self.bytes, "\\x{:02X}", u32::from(character))
                    .expect("writing to a Vec cannot fail");
            } else {
                let mut encoded = [0u8; 4];
                self.bytes
                    .extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            }
        }
        Ok(())
    }

    pub(crate) fn push_cell_value(&mut self, value: &CellValue) -> sheetkit::Result<()> {
        for_each_cell_text(value, |segment| self.push_safe(segment))
    }

    pub(crate) fn finish(self) -> Vec<u8> {
        self.bytes
    }

    fn push_bytes(&mut self, value: &[u8]) -> sheetkit::Result<()> {
        let additional = u64::try_from(value.len()).unwrap_or(u64::MAX);
        self.charge_and_reserve(additional)?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn charge_and_reserve(&mut self, additional: u64) -> sheetkit::Result<()> {
        let observed = self.used.saturating_add(additional);
        if observed > self.limit {
            return Err(output_limit_error(self.limit, observed));
        }
        let additional = usize::try_from(additional)
            .map_err(|_| sheetkit::Error::Internal("output size is not addressable".to_string()))?;
        self.bytes
            .try_reserve(additional)
            .map_err(|_| sheetkit::Error::Internal("output allocation failed".to_string()))?;
        self.used = observed;
        Ok(())
    }
}

pub(crate) fn terminal_safe(value: &str) -> String {
    let capacity = usize::try_from(terminal_safe_len(value)).unwrap_or(value.len());
    let mut output = String::with_capacity(capacity);
    append_terminal_safe(&mut output, value);
    output
}

fn append_terminal_safe(output: &mut String, value: &str) {
    for character in value.chars() {
        if is_terminal_control(character) {
            write!(output, "\\x{:02X}", u32::from(character))
                .expect("writing to a String cannot fail");
        } else {
            output.push(character);
        }
    }
}

pub(crate) fn json_bytes(
    value: &(impl Serialize + ?Sized),
    limit: u64,
) -> sheetkit::Result<Vec<u8>> {
    let mut writer = TerminalJsonWriter::new(limit);
    if let Err(error) = serde_json::to_writer(&mut writer, value) {
        if let Some((limit, observed)) = writer.limit_error {
            return Err(output_limit_error(limit, observed));
        }
        return Err(sheetkit::Error::Internal(format!(
            "JSON serialization failed: {error}"
        )));
    }
    writer.finish()
}

pub(crate) fn terminal_safe_len(value: &str) -> u64 {
    value.chars().fold(0u64, |total, character| {
        total.saturating_add(if is_terminal_control(character) {
            4
        } else {
            u64::try_from(character.len_utf8()).unwrap_or(u64::MAX)
        })
    })
}

pub(crate) fn terminal_safe_cell_len(value: &CellValue) -> u64 {
    let mut total = 0u64;
    let _: sheetkit::Result<()> = for_each_cell_text(value, |segment| {
        total = total.saturating_add(terminal_safe_len(segment));
        Ok(())
    });
    total
}

pub(crate) fn append_terminal_safe_cell(output: &mut String, value: &CellValue) {
    let _: sheetkit::Result<()> = for_each_cell_text(value, |segment| {
        append_terminal_safe(output, segment);
        Ok(())
    });
}

fn for_each_cell_text(
    value: &CellValue,
    mut visit: impl FnMut(&str) -> sheetkit::Result<()>,
) -> sheetkit::Result<()> {
    match value {
        CellValue::Empty => Ok(()),
        CellValue::Bool(value) => visit(if *value { "TRUE" } else { "FALSE" }),
        CellValue::Number(_) | CellValue::Date(_) => visit(&value.to_string()),
        CellValue::String(value) | CellValue::Error(value) => visit(value),
        CellValue::Formula {
            result: Some(value),
            ..
        } => for_each_cell_text(value, visit),
        CellValue::Formula { expr, result: None } => {
            visit("=")?;
            visit(expr)
        }
        CellValue::RichString(runs) => {
            for run in runs {
                visit(&run.text)?;
            }
            Ok(())
        }
    }
}

pub(crate) fn is_terminal_control(character: char) -> bool {
    let code = u32::from(character);
    code <= 0x1f || code == 0x7f || (0x80..=0x9f).contains(&code)
}

fn output_limit_error(limit: u64, observed: u64) -> sheetkit::Error {
    sheetkit::ResourceLimitError::exceeded(
        sheetkit::ResourceOperation::Export,
        sheetkit::ResourceCategory::Output,
        sheetkit::ResourceDimension::OutputBytes,
        limit,
        observed.max(limit.saturating_add(1)),
        Some(sheetkit::ResourceStage::Serialize),
    )
    .map(sheetkit::Error::from)
    .unwrap_or_else(|| sheetkit::Error::Internal("invalid output resource limit".to_string()))
}

struct TerminalJsonWriter {
    output: Vec<u8>,
    pending: Vec<u8>,
    limit: u64,
    used: u64,
    limit_error: Option<(u64, u64)>,
}

impl TerminalJsonWriter {
    const fn new(limit: u64) -> Self {
        Self {
            output: Vec::new(),
            pending: Vec::new(),
            limit,
            used: 0,
            limit_error: None,
        }
    }

    fn append_valid(&mut self, value: &str) -> io::Result<()> {
        let additional = value.chars().fold(0u64, |total, character| {
            total.saturating_add(if is_terminal_control(character) {
                6
            } else {
                u64::try_from(character.len_utf8()).unwrap_or(u64::MAX)
            })
        });
        self.reserve(additional)?;
        for character in value.chars() {
            if is_terminal_control(character) {
                write!(self.output, "\\u{:04X}", u32::from(character))
                    .expect("writing to a Vec cannot fail");
            } else {
                let mut encoded = [0u8; 4];
                self.output
                    .extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            }
        }
        Ok(())
    }

    fn reserve(&mut self, additional: u64) -> io::Result<()> {
        let observed = self.used.saturating_add(additional);
        if observed > self.limit {
            self.limit_error = Some((self.limit, observed));
            return Err(io::Error::other("JSON output limit exceeded"));
        }
        let additional = usize::try_from(additional)
            .map_err(|_| io::Error::other("JSON output size is not addressable"))?;
        self.output
            .try_reserve(additional)
            .map_err(|_| io::Error::other("JSON output allocation failed"))?;
        self.used = observed;
        Ok(())
    }

    fn process(&mut self, bytes: &[u8]) -> io::Result<()> {
        let mut remaining = bytes;
        while !self.pending.is_empty() && !remaining.is_empty() {
            self.pending.push(remaining[0]);
            remaining = &remaining[1..];
            match std::str::from_utf8(&self.pending) {
                Ok(value) => {
                    let value = value.to_string();
                    self.pending.clear();
                    self.append_valid(&value)?;
                }
                Err(error) if error.error_len().is_none() && self.pending.len() < 4 => {}
                Err(_) => return Err(invalid_json_utf8()),
            }
        }
        if !self.pending.is_empty() {
            return Ok(());
        }

        match std::str::from_utf8(remaining) {
            Ok(value) => self.append_valid(value),
            Err(error) if error.error_len().is_none() => {
                let valid = error.valid_up_to();
                self.append_valid(
                    std::str::from_utf8(&remaining[..valid]).expect("valid UTF-8 prefix"),
                )?;
                self.pending.extend_from_slice(&remaining[valid..]);
                debug_assert!(self.pending.len() <= 3);
                Ok(())
            }
            Err(_) => Err(invalid_json_utf8()),
        }
    }

    fn finish(mut self) -> sheetkit::Result<Vec<u8>> {
        if !self.pending.is_empty() {
            return Err(sheetkit::Error::Internal(
                "JSON serializer emitted incomplete UTF-8".to_string(),
            ));
        }
        if let Err(error) = self.reserve(1) {
            if let Some((limit, observed)) = self.limit_error {
                return Err(output_limit_error(limit, observed));
            }
            return Err(sheetkit::Error::Internal(error.to_string()));
        }
        self.output.push(b'\n');
        Ok(self.output)
    }
}

fn invalid_json_utf8() -> io::Error {
    io::Error::new(
        io::ErrorKind::InvalidData,
        "JSON serializer emitted invalid UTF-8",
    )
}

impl io::Write for TerminalJsonWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.process(bytes)?;
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

pub(crate) struct InfoJson<'a> {
    file: &'a Path,
    sheets: &'a [&'a str],
    active_sheet: &'a str,
    properties: &'a DocProperties,
}

pub(crate) fn info_json<'a>(
    file: &'a Path,
    sheets: &'a [&'a str],
    active_sheet: &'a str,
    properties: &'a DocProperties,
) -> InfoJson<'a> {
    InfoJson {
        file,
        sheets,
        active_sheet,
        properties,
    }
}

impl Serialize for InfoJson<'_> {
    fn serialize<S>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        let mut map = serializer.serialize_map(Some(5))?;
        map.serialize_entry("file", &self.file.to_string_lossy())?;
        map.serialize_entry("sheetCount", &self.sheets.len())?;
        map.serialize_entry("sheets", self.sheets)?;
        map.serialize_entry("activeSheet", self.active_sheet)?;
        map.serialize_entry("properties", &PropertiesJson(self.properties))?;
        map.end()
    }
}

struct PropertiesJson<'a>(&'a DocProperties);

impl Serialize for PropertiesJson<'_> {
    fn serialize<S>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        let properties = self.0;
        let mut map = serializer.serialize_map(Some(11))?;
        map.serialize_entry("title", &properties.title.as_deref())?;
        map.serialize_entry("subject", &properties.subject.as_deref())?;
        map.serialize_entry("creator", &properties.creator.as_deref())?;
        map.serialize_entry("keywords", &properties.keywords.as_deref())?;
        map.serialize_entry("description", &properties.description.as_deref())?;
        map.serialize_entry("lastModifiedBy", &properties.last_modified_by.as_deref())?;
        map.serialize_entry("revision", &properties.revision.as_deref())?;
        map.serialize_entry("created", &properties.created.as_deref())?;
        map.serialize_entry("modified", &properties.modified.as_deref())?;
        map.serialize_entry("category", &properties.category.as_deref())?;
        map.serialize_entry("contentStatus", &properties.content_status.as_deref())?;
        map.end()
    }
}

pub(crate) struct RowsJson<'a>(&'a [(u32, Vec<(u32, CellValue)>)]);

pub(crate) fn rows_json(rows: &[(u32, Vec<(u32, CellValue)>)]) -> RowsJson<'_> {
    RowsJson(rows)
}

impl Serialize for RowsJson<'_> {
    fn serialize<S>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        let mut sequence = serializer.serialize_seq(Some(self.0.len()))?;
        for (row, cells) in self.0 {
            sequence.serialize_element(&RowJson { row: *row, cells })?;
        }
        sequence.end()
    }
}

struct RowJson<'a> {
    row: u32,
    cells: &'a [(u32, CellValue)],
}

impl Serialize for RowJson<'_> {
    fn serialize<S>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        let mut map = serializer.serialize_map(Some(2))?;
        map.serialize_entry("row", &self.row)?;
        map.serialize_entry("cells", &CellsJson(self.cells))?;
        map.end()
    }
}

struct CellsJson<'a>(&'a [(u32, CellValue)]);

impl Serialize for CellsJson<'_> {
    fn serialize<S>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        let mut sequence = serializer.serialize_seq(Some(self.0.len()))?;
        for (column, value) in self.0 {
            sequence.serialize_element(&CellJson {
                column: *column,
                value,
            })?;
        }
        sequence.end()
    }
}

struct CellJson<'a> {
    column: u32,
    value: &'a CellValue,
}

impl Serialize for CellJson<'_> {
    fn serialize<S>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        let mut map = serializer.serialize_map(Some(2))?;
        map.serialize_entry("column", &self.column)?;
        map.serialize_entry("value", &CellValueJson(self.value))?;
        map.end()
    }
}

pub(crate) struct CellValueJson<'a>(&'a CellValue);

pub(crate) fn cell_value_json(value: &CellValue) -> CellValueJson<'_> {
    CellValueJson(value)
}

impl Serialize for CellValueJson<'_> {
    fn serialize<S>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match self.0 {
            CellValue::Empty => serializer.serialize_none(),
            CellValue::Bool(value) => serializer.serialize_bool(*value),
            CellValue::Number(value) | CellValue::Date(value) if value.is_finite() => {
                serializer.serialize_f64(*value)
            }
            CellValue::Number(value) | CellValue::Date(value) => serializer.collect_str(value),
            CellValue::String(value) | CellValue::Error(value) => serializer.serialize_str(value),
            CellValue::Formula {
                result: Some(value),
                ..
            } => CellValueJson(value).serialize(serializer),
            CellValue::Formula { expr, result: None } => {
                serializer.collect_str(&format_args!("={expr}"))
            }
            CellValue::RichString(runs) => serializer.collect_str(&RichTextJson(runs)),
        }
    }
}

struct RichTextJson<'a>(&'a [sheetkit::RichTextRun]);

impl std::fmt::Display for RichTextJson<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        for run in self.0 {
            formatter.write_str(&run.text)?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::{json, Value};

    #[test]
    fn terminal_controls_are_visible_and_printable_unicode_is_preserved() {
        assert_eq!(
            terminal_safe("한글\0\n\u{001b}\u{007f}\u{0085}\u{009b}"),
            "한글\\x00\\x0A\\x1B\\x7F\\x85\\x9B"
        );
    }

    #[test]
    fn json_escapes_terminal_controls_and_round_trips_original_data() {
        let original = json!({ "value": "한글\u{001b}\u{007f}\u{0085}\u{009b}" });
        let bytes = json_bytes(&original, MAX_CLI_OUTPUT_BYTES).unwrap();
        assert!(!bytes.windows(2).any(|bytes| bytes == [0xc2, 0x9b]));
        assert_eq!(serde_json::from_slice::<Value>(&bytes).unwrap(), original);
    }

    #[test]
    fn json_output_limit_accepts_exact_bytes_and_rejects_plus_one() {
        let value = json!({ "value": "한글\u{009b}" });
        let exact = json_bytes(&value, MAX_CLI_OUTPUT_BYTES).unwrap();
        assert_eq!(json_bytes(&value, exact.len() as u64).unwrap(), exact);

        let error = json_bytes(&value, exact.len() as u64 - 1).unwrap_err();
        let sheetkit::Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.limit(), exact.len() as u64 - 1);
        assert_eq!(metadata.observed(), exact.len() as u64);
    }

    #[test]
    fn json_writer_bounds_large_chunks_and_handles_split_utf8() {
        let mut split = TerminalJsonWriter::new(4);
        split.write_all(&[0xed]).unwrap();
        split.write_all(&[0x95, 0x9c]).unwrap();
        assert_eq!(split.finish().unwrap(), "한\n".as_bytes());

        let mut bounded = TerminalJsonWriter::new(4);
        assert!(bounded.write_all(&[b'x'; 1_024]).is_err());
        assert!(bounded.output.is_empty());
        assert!(bounded.pending.is_empty());
    }
}
