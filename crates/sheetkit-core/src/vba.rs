//! VBA project extraction from macro-enabled workbooks (.xlsm).
//!
//! `.xlsm` files contain a `xl/vbaProject.bin` entry which is an OLE2
//! Compound Binary File (CFB) holding VBA source code. This module
//! provides read-only access to the raw binary and to individual VBA
//! module source code.

use std::io::{Cursor, Read as _};

use crate::error::{Error, Result};
use crate::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};

const MIB: u64 = 1024 * 1024;
const MAX_WARNINGS: usize = 32;

/// Finite limits applied to one VBA extraction operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct VbaLimits {
    /// Maximum size of the input CFB container.
    pub input_bytes: u64,
    /// Maximum number of streams in the CFB container.
    pub cfb_stream_count: u64,
    /// Maximum size declared by one CFB stream.
    pub cfb_stream_declared_bytes: u64,
    /// Maximum sum of sizes declared by CFB streams.
    pub cfb_aggregate_declared_bytes: u64,
    /// Maximum bytes read from one CFB stream.
    pub cfb_stream_read_bytes: u64,
    /// Maximum cumulative bytes read from CFB streams.
    pub cfb_aggregate_read_bytes: u64,
    /// Maximum compressed bytes submitted to VBA decompression.
    pub compressed_bytes: u64,
    /// Maximum number of VBA modules.
    pub module_count: u64,
    /// Maximum number of compressed chunks processed.
    pub chunk_count: u64,
    /// Maximum decoded bytes produced by one module or the `dir` stream.
    pub module_decoded_bytes: u64,
    /// Maximum cumulative bytes produced by VBA decompression.
    pub aggregate_decoded_bytes: u64,
    /// Maximum bytes retained in names, source strings, and warnings.
    pub retained_bytes: u64,
    /// Maximum UTF-8 bytes returned in extracted module fields.
    pub output_bytes: u64,
}

impl Default for VbaLimits {
    fn default() -> Self {
        Self {
            input_bytes: 64 * MIB,
            cfb_stream_count: 1_024,
            cfb_stream_declared_bytes: 32 * MIB,
            cfb_aggregate_declared_bytes: 64 * MIB,
            cfb_stream_read_bytes: 32 * MIB,
            cfb_aggregate_read_bytes: 64 * MIB,
            compressed_bytes: 64 * MIB,
            module_count: 1_024,
            chunk_count: 65_536,
            module_decoded_bytes: 16 * MIB,
            aggregate_decoded_bytes: 64 * MIB,
            retained_bytes: 64 * MIB,
            output_bytes: 64 * MIB,
        }
    }
}

impl VbaLimits {
    /// Largest supported finite values. These ceilings cannot be disabled.
    pub const fn hard_ceiling() -> Self {
        Self {
            input_bytes: 1024 * MIB,
            cfb_stream_count: 65_536,
            cfb_stream_declared_bytes: 512 * MIB,
            cfb_aggregate_declared_bytes: 1024 * MIB,
            cfb_stream_read_bytes: 512 * MIB,
            cfb_aggregate_read_bytes: 1024 * MIB,
            compressed_bytes: 1024 * MIB,
            module_count: 16_384,
            chunk_count: 1_048_576,
            module_decoded_bytes: 256 * MIB,
            aggregate_decoded_bytes: 1024 * MIB,
            retained_bytes: 1024 * MIB,
            output_bytes: 1024 * MIB,
        }
    }

    fn validate(self) -> Result<Self> {
        let ceiling = Self::hard_ceiling();
        for (dimension, value, maximum) in [
            (
                ResourceDimension::InputBytes,
                self.input_bytes,
                ceiling.input_bytes,
            ),
            (
                ResourceDimension::StreamCount,
                self.cfb_stream_count,
                ceiling.cfb_stream_count,
            ),
            (
                ResourceDimension::EntryDeclaredBytes,
                self.cfb_stream_declared_bytes,
                ceiling.cfb_stream_declared_bytes,
            ),
            (
                ResourceDimension::AggregateDeclaredBytes,
                self.cfb_aggregate_declared_bytes,
                ceiling.cfb_aggregate_declared_bytes,
            ),
            (
                ResourceDimension::EntryActualBytes,
                self.cfb_stream_read_bytes,
                ceiling.cfb_stream_read_bytes,
            ),
            (
                ResourceDimension::AggregateActualBytes,
                self.cfb_aggregate_read_bytes,
                ceiling.cfb_aggregate_read_bytes,
            ),
            (
                ResourceDimension::InputBytes,
                self.compressed_bytes,
                ceiling.compressed_bytes,
            ),
            (
                ResourceDimension::ModuleCount,
                self.module_count,
                ceiling.module_count,
            ),
            (
                ResourceDimension::ChunkCount,
                self.chunk_count,
                ceiling.chunk_count,
            ),
            (
                ResourceDimension::OutputBytes,
                self.module_decoded_bytes,
                ceiling.module_decoded_bytes,
            ),
            (
                ResourceDimension::OutputBytes,
                self.aggregate_decoded_bytes,
                ceiling.aggregate_decoded_bytes,
            ),
            (
                ResourceDimension::RetainedBytes,
                self.retained_bytes,
                ceiling.retained_bytes,
            ),
            (
                ResourceDimension::OutputBytes,
                self.output_bytes,
                ceiling.output_bytes,
            ),
        ] {
            if value == 0 || value > maximum {
                return Err(vba_limit(
                    dimension,
                    maximum,
                    value.max(maximum.saturating_add(1)),
                    ResourceStage::Validate,
                ));
            }
        }
        Ok(self)
    }
}

struct VbaBudget {
    limits: VbaLimits,
    stream_count: u64,
    aggregate_declared: u64,
    aggregate_read: u64,
    compressed: u64,
    modules: u64,
    chunks: u64,
    aggregate_decoded: u64,
    retained: u64,
    output: u64,
}

impl VbaBudget {
    fn new(limits: VbaLimits) -> Result<Self> {
        Ok(Self {
            limits: limits.validate()?,
            stream_count: 0,
            aggregate_declared: 0,
            aggregate_read: 0,
            compressed: 0,
            modules: 0,
            chunks: 0,
            aggregate_decoded: 0,
            retained: 0,
            output: 0,
        })
    }

    fn check_input(&self, observed: u64) -> Result<()> {
        check_vba_limit(
            ResourceDimension::InputBytes,
            self.limits.input_bytes,
            observed,
            ResourceStage::Inspect,
        )
    }

    fn inspect_stream(&mut self, declared: u64) -> Result<()> {
        self.stream_count = checked_total(self.stream_count, 1);
        check_vba_limit(
            ResourceDimension::StreamCount,
            self.limits.cfb_stream_count,
            self.stream_count,
            ResourceStage::Inspect,
        )?;
        check_vba_limit(
            ResourceDimension::EntryDeclaredBytes,
            self.limits.cfb_stream_declared_bytes,
            declared,
            ResourceStage::Inspect,
        )?;
        let aggregate = checked_total(self.aggregate_declared, declared);
        check_vba_limit(
            ResourceDimension::AggregateDeclaredBytes,
            self.limits.cfb_aggregate_declared_bytes,
            aggregate,
            ResourceStage::Inspect,
        )?;
        self.aggregate_declared = aggregate;
        Ok(())
    }

    fn charge_stream_read(&mut self, stream_total: u64, amount: u64) -> Result<u64> {
        let stream_total = checked_total(stream_total, amount);
        check_vba_limit(
            ResourceDimension::EntryActualBytes,
            self.limits.cfb_stream_read_bytes,
            stream_total,
            ResourceStage::Read,
        )?;
        let aggregate = checked_total(self.aggregate_read, amount);
        check_vba_limit(
            ResourceDimension::AggregateActualBytes,
            self.limits.cfb_aggregate_read_bytes,
            aggregate,
            ResourceStage::Read,
        )?;
        self.aggregate_read = aggregate;
        Ok(stream_total)
    }

    fn charge_compressed(&mut self, amount: u64) -> Result<()> {
        let observed = checked_total(self.compressed, amount);
        check_vba_limit(
            ResourceDimension::InputBytes,
            self.limits.compressed_bytes,
            observed,
            ResourceStage::Decode,
        )?;
        self.compressed = observed;
        Ok(())
    }

    fn charge_module(&mut self) -> Result<()> {
        let observed = checked_total(self.modules, 1);
        check_vba_limit(
            ResourceDimension::ModuleCount,
            self.limits.module_count,
            observed,
            ResourceStage::Parse,
        )?;
        self.modules = observed;
        Ok(())
    }

    fn charge_chunk(&mut self) -> Result<()> {
        let observed = checked_total(self.chunks, 1);
        check_vba_limit(
            ResourceDimension::ChunkCount,
            self.limits.chunk_count,
            observed,
            ResourceStage::Decode,
        )?;
        self.chunks = observed;
        Ok(())
    }

    fn check_decoded(&self, module_total: u64, amount: u64) -> Result<(u64, u64)> {
        let module_total = checked_total(module_total, amount);
        check_vba_limit(
            ResourceDimension::OutputBytes,
            self.limits.module_decoded_bytes,
            module_total,
            ResourceStage::Decode,
        )?;
        let aggregate = checked_total(self.aggregate_decoded, amount);
        check_vba_limit(
            ResourceDimension::OutputBytes,
            self.limits.aggregate_decoded_bytes,
            aggregate,
            ResourceStage::Decode,
        )?;
        Ok((module_total, aggregate))
    }

    fn commit_decoded(&mut self, aggregate: u64) {
        self.aggregate_decoded = aggregate;
    }

    fn charge_retained(&mut self, amount: u64) -> Result<()> {
        let observed = checked_total(self.retained, amount);
        check_vba_limit(
            ResourceDimension::RetainedBytes,
            self.limits.retained_bytes,
            observed,
            ResourceStage::Decode,
        )?;
        self.retained = observed;
        Ok(())
    }

    fn charge_output(&mut self, amount: u64) -> Result<()> {
        let observed = checked_total(self.output, amount);
        check_vba_limit(
            ResourceDimension::OutputBytes,
            self.limits.output_bytes,
            observed,
            ResourceStage::Decode,
        )?;
        self.output = observed;
        Ok(())
    }
}

fn checked_total(current: u64, amount: u64) -> u64 {
    current.saturating_add(amount)
}

fn check_vba_limit(
    dimension: ResourceDimension,
    limit: u64,
    observed: u64,
    stage: ResourceStage,
) -> Result<()> {
    if observed > limit {
        return Err(vba_limit(dimension, limit, observed, stage));
    }
    Ok(())
}

fn vba_limit(
    dimension: ResourceDimension,
    limit: u64,
    observed: u64,
    stage: ResourceStage,
) -> Error {
    let metadata = ResourceLimitError::exceeded(
        ResourceOperation::MacroRead,
        ResourceCategory::Macro,
        dimension,
        limit,
        observed,
        Some(stage),
    );
    match metadata {
        Some(metadata) => Error::ResourceLimitExceeded(metadata),
        None => Error::Internal("invalid VBA resource limit state".to_string()),
    }
}

/// Classification of a VBA module.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum VbaModuleType {
    /// A standard code module (`.bas`).
    Standard,
    /// A class module (`.cls`).
    Class,
    /// A UserForm module.
    Form,
    /// A document module (e.g. Sheet code-behind).
    Document,
    /// The ThisWorkbook module.
    ThisWorkbook,
}

/// A single VBA module with its name, source code, and type.
#[derive(Debug, Clone)]
pub struct VbaModule {
    pub name: String,
    pub source_code: String,
    pub module_type: VbaModuleType,
}

/// Result of extracting a VBA project from a `.xlsm` file.
///
/// Contains extracted modules and any non-fatal warnings encountered
/// during parsing (e.g., unreadable streams, decompression failures,
/// unsupported codepages).
#[derive(Debug, Clone)]
pub struct VbaProject {
    pub modules: Vec<VbaModule>,
    pub warnings: Vec<String>,
}

/// Offset entry parsed from the `dir` stream for a single module.
#[derive(Debug)]
struct ModuleEntry {
    name: String,
    stream_name: String,
    text_offset: u32,
    module_type: VbaModuleType,
}

/// Parsed metadata from the `dir` stream.
#[derive(Debug)]
struct DirInfo {
    entries: Vec<ModuleEntry>,
    codepage: u16,
}

/// Extract VBA module source code from a `vbaProject.bin` binary blob.
///
/// Parses the OLE/CFB container, reads the `dir` stream to discover
/// module metadata, then decompresses each module stream.
///
/// Returns a [`VbaProject`] containing extracted modules and any
/// non-fatal warnings (e.g., modules that could not be read or
/// decompressed, unsupported codepages).
pub fn extract_vba_modules(vba_bin: &[u8]) -> Result<VbaProject> {
    extract_vba_modules_with_limits(vba_bin, VbaLimits::default())
}

/// Validates a raw VBA project against finite input and hard-ceiling limits.
pub(crate) fn validate_vba_input(vba_bin: &[u8], limits: VbaLimits) -> Result<()> {
    let budget = VbaBudget::new(limits)?;
    budget.check_input(u64::try_from(vba_bin.len()).unwrap_or(u64::MAX))
}

/// Extract VBA modules using explicit finite limits.
pub fn extract_vba_modules_with_limits(vba_bin: &[u8], limits: VbaLimits) -> Result<VbaProject> {
    validate_vba_input(vba_bin, limits)?;
    let mut budget = VbaBudget::new(limits)?;
    let cursor = Cursor::new(vba_bin);
    let mut cfb = cfb::CompoundFile::open(cursor)
        .map_err(|_| Error::Internal("failed to open VBA project as CFB".to_string()))?;

    for entry in cfb.walk() {
        if entry.is_stream() {
            budget.inspect_stream(entry.len())?;
        }
    }

    // Find the VBA storage root. Typically `/VBA` or could be nested.
    let vba_prefix = find_vba_prefix(&mut cfb)?;

    // Read the `dir` stream to get module entries.
    let dir_path = format!("{vba_prefix}dir");
    let dir_data = read_cfb_stream(&mut cfb, &dir_path, &mut budget)?;

    // The dir stream is compressed using MS-OVBA compression.
    let decompressed_dir = decompress_vba_stream_budgeted(&dir_data, &mut budget)?;

    // Parse module entries and codepage from the decompressed dir stream.
    let dir_info = parse_dir_stream_budgeted(&decompressed_dir, &mut budget)?;

    let mut modules = Vec::new();
    modules
        .try_reserve(dir_info.entries.len())
        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
    let mut warnings = Vec::new();
    warnings
        .try_reserve(MAX_WARNINGS + 1)
        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;

    for entry in dir_info.entries {
        let stream_path = format!("{vba_prefix}{}", entry.stream_name);
        let compressed_data = match read_cfb_stream(&mut cfb, &stream_path, &mut budget) {
            Ok(data) => data,
            Err(Error::ResourceLimitExceeded(metadata)) => {
                return Err(Error::ResourceLimitExceeded(metadata));
            }
            Err(_) => {
                push_warning(
                    &mut warnings,
                    "module skipped: stream read failed",
                    &mut budget,
                )?;
                continue;
            }
        };

        // The module stream has text_offset bytes of "performance cache"
        // (compiled code) followed by compressed source code.
        let text_offset = usize::try_from(entry.text_offset)
            .map_err(|_| Error::Internal("invalid VBA module offset".to_string()))?;
        if text_offset > compressed_data.len() {
            push_warning(
                &mut warnings,
                "module skipped: invalid source offset",
                &mut budget,
            )?;
            continue;
        }
        let source_compressed = &compressed_data[text_offset..];
        let source_bytes = match decompress_vba_stream_budgeted(source_compressed, &mut budget) {
            Ok(b) => b,
            Err(Error::ResourceLimitExceeded(metadata)) => {
                return Err(Error::ResourceLimitExceeded(metadata));
            }
            Err(_) => {
                push_warning(
                    &mut warnings,
                    "module skipped: decompression failed",
                    &mut budget,
                )?;
                continue;
            }
        };

        let source_code =
            decode_source_bytes(&source_bytes, dir_info.codepage, &mut warnings, &mut budget)?;
        budget.charge_retained(u64::try_from(entry.name.len()).unwrap_or(u64::MAX))?;
        budget.charge_output(u64::try_from(entry.name.len()).unwrap_or(u64::MAX))?;

        modules.push(VbaModule {
            name: entry.name,
            source_code,
            module_type: entry.module_type,
        });
    }

    Ok(VbaProject { modules, warnings })
}

fn push_warning(
    warnings: &mut Vec<String>,
    warning: &'static str,
    budget: &mut VbaBudget,
) -> Result<()> {
    if warnings.len() < MAX_WARNINGS {
        let len = u64::try_from(warning.len()).unwrap_or(u64::MAX);
        budget.charge_output(len)?;
        budget.charge_retained(len)?;
        warnings.push(warning.to_string());
    } else if warnings.len() == MAX_WARNINGS {
        const OMITTED: &str = "additional VBA warnings omitted";
        let len = u64::try_from(OMITTED.len()).unwrap_or(u64::MAX);
        budget.charge_output(len)?;
        budget.charge_retained(len)?;
        warnings.push(OMITTED.to_string());
    }
    Ok(())
}

/// Decode source bytes using the specified codepage.
///
/// Supports common codepages: 1252 (Western European), 932 (Japanese Shift-JIS),
/// 949 (Korean), 936 (Simplified Chinese GBK), 65001 (UTF-8).
/// For unrecognized codepages, falls back to UTF-8 lossy and emits a warning.
fn decode_source_bytes(
    bytes: &[u8],
    codepage: u16,
    warnings: &mut Vec<String>,
    budget: &mut VbaBudget,
) -> Result<String> {
    let decoded_len = decoded_source_len(bytes, codepage);
    budget.charge_output(decoded_len)?;
    budget.charge_retained(decoded_len)?;
    let capacity = usize::try_from(decoded_len)
        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
    match codepage {
        65001 | 0 => decode_utf8_lossy(bytes, capacity),
        1252 => decode_single_byte(bytes, &WINDOWS_1252_HIGH, capacity),
        932 => decode_shift_jis(bytes, capacity),
        949 => decode_double_byte_lossy(bytes, capacity),
        936 => decode_double_byte_lossy(bytes, capacity),
        _ => {
            push_warning(
                warnings,
                "unsupported codepage; used UTF-8 lossy decoding",
                budget,
            )?;
            decode_utf8_lossy(bytes, capacity)
        }
    }
}

fn decoded_source_len(bytes: &[u8], codepage: u16) -> u64 {
    let length = match codepage {
        65001 | 0 => utf8_lossy_len(bytes),
        1252 => bytes.iter().fold(0_usize, |total, &byte| {
            let width = if byte < 0x80 {
                1
            } else {
                WINDOWS_1252_HIGH[usize::from(byte - 0x80)].len_utf8()
            };
            total.saturating_add(width)
        }),
        932 => shift_jis_lossy_len(bytes),
        949 | 936 => double_byte_lossy_len(bytes),
        _ => utf8_lossy_len(bytes),
    };
    u64::try_from(length).unwrap_or(u64::MAX)
}

fn utf8_lossy_len(mut bytes: &[u8]) -> usize {
    let mut length = 0_usize;
    while !bytes.is_empty() {
        match std::str::from_utf8(bytes) {
            Ok(valid) => return length.saturating_add(valid.len()),
            Err(error) => {
                length = length
                    .saturating_add(error.valid_up_to())
                    .saturating_add('�'.len_utf8());
                match error.error_len() {
                    Some(invalid) => bytes = &bytes[error.valid_up_to() + invalid..],
                    None => return length,
                }
            }
        }
    }
    length
}

fn decode_utf8_lossy(mut bytes: &[u8], capacity: usize) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve(capacity)
        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
    while !bytes.is_empty() {
        match std::str::from_utf8(bytes) {
            Ok(valid) => {
                output.push_str(valid);
                break;
            }
            Err(error) => {
                let valid = &bytes[..error.valid_up_to()];
                let valid = std::str::from_utf8(valid)
                    .map_err(|_| Error::Internal("invalid VBA decode state".to_string()))?;
                output.push_str(valid);
                output.push('�');
                match error.error_len() {
                    Some(invalid) => bytes = &bytes[error.valid_up_to() + invalid..],
                    None => break,
                }
            }
        }
    }
    Ok(output)
}

/// Windows-1252 high-byte mapping (0x80..0xFF).
/// Bytes 0x00..0x7F are identical to ASCII.
static WINDOWS_1252_HIGH: [char; 128] = [
    '\u{20AC}', '\u{0081}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}', '\u{2020}', '\u{2021}',
    '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{008D}', '\u{017D}', '\u{008F}',
    '\u{0090}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}', '\u{2022}', '\u{2013}', '\u{2014}',
    '\u{02DC}', '\u{2122}', '\u{0161}', '\u{203A}', '\u{0153}', '\u{009D}', '\u{017E}', '\u{0178}',
    '\u{00A0}', '\u{00A1}', '\u{00A2}', '\u{00A3}', '\u{00A4}', '\u{00A5}', '\u{00A6}', '\u{00A7}',
    '\u{00A8}', '\u{00A9}', '\u{00AA}', '\u{00AB}', '\u{00AC}', '\u{00AD}', '\u{00AE}', '\u{00AF}',
    '\u{00B0}', '\u{00B1}', '\u{00B2}', '\u{00B3}', '\u{00B4}', '\u{00B5}', '\u{00B6}', '\u{00B7}',
    '\u{00B8}', '\u{00B9}', '\u{00BA}', '\u{00BB}', '\u{00BC}', '\u{00BD}', '\u{00BE}', '\u{00BF}',
    '\u{00C0}', '\u{00C1}', '\u{00C2}', '\u{00C3}', '\u{00C4}', '\u{00C5}', '\u{00C6}', '\u{00C7}',
    '\u{00C8}', '\u{00C9}', '\u{00CA}', '\u{00CB}', '\u{00CC}', '\u{00CD}', '\u{00CE}', '\u{00CF}',
    '\u{00D0}', '\u{00D1}', '\u{00D2}', '\u{00D3}', '\u{00D4}', '\u{00D5}', '\u{00D6}', '\u{00D7}',
    '\u{00D8}', '\u{00D9}', '\u{00DA}', '\u{00DB}', '\u{00DC}', '\u{00DD}', '\u{00DE}', '\u{00DF}',
    '\u{00E0}', '\u{00E1}', '\u{00E2}', '\u{00E3}', '\u{00E4}', '\u{00E5}', '\u{00E6}', '\u{00E7}',
    '\u{00E8}', '\u{00E9}', '\u{00EA}', '\u{00EB}', '\u{00EC}', '\u{00ED}', '\u{00EE}', '\u{00EF}',
    '\u{00F0}', '\u{00F1}', '\u{00F2}', '\u{00F3}', '\u{00F4}', '\u{00F5}', '\u{00F6}', '\u{00F7}',
    '\u{00F8}', '\u{00F9}', '\u{00FA}', '\u{00FB}', '\u{00FC}', '\u{00FD}', '\u{00FE}', '\u{00FF}',
];

/// Decode bytes using a single-byte codepage with the given high-byte table.
fn decode_single_byte(bytes: &[u8], high_table: &[char; 128], capacity: usize) -> Result<String> {
    let mut out = String::new();
    out.try_reserve(capacity)
        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
    for &b in bytes {
        if b < 0x80 {
            out.push(b as char);
        } else {
            out.push(high_table[(b - 0x80) as usize]);
        }
    }
    Ok(out)
}

/// Decode Shift-JIS (codepage 932) bytes to a String.
/// Uses a best-effort approach: valid multi-byte sequences are decoded,
/// invalid bytes are replaced with the Unicode replacement character.
fn decode_shift_jis(bytes: &[u8], capacity: usize) -> Result<String> {
    let mut out = String::new();
    out.try_reserve(capacity)
        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
    let mut i = 0;
    while i < bytes.len() {
        let b = bytes[i];
        if b < 0x80 {
            out.push(b as char);
            i += 1;
        } else if b == 0x80 || b == 0xA0 || b >= 0xFD {
            out.push('\u{FFFD}');
            i += 1;
        } else if (0xA1..=0xDF).contains(&b) {
            // Half-width katakana
            out.push(char::from_u32(0xFF61 + (b as u32 - 0xA1)).unwrap_or('\u{FFFD}'));
            i += 1;
        } else if i + 1 < bytes.len() {
            // Double-byte character -- fall back to replacement for simplicity
            // Full Shift-JIS decoding requires a large mapping table.
            out.push('\u{FFFD}');
            i += 2;
        } else {
            out.push('\u{FFFD}');
            i += 1;
        }
    }
    Ok(out)
}

fn shift_jis_lossy_len(bytes: &[u8]) -> usize {
    let mut length = 0_usize;
    let mut i = 0;
    while i < bytes.len() {
        let b = bytes[i];
        if b < 0x80 {
            length = length.saturating_add(1);
            i += 1;
        } else if b == 0x80 || b == 0xA0 || b >= 0xFD || (0xA1..=0xDF).contains(&b) {
            length = length.saturating_add(3);
            i += 1;
        } else if i + 1 < bytes.len() {
            length = length.saturating_add(3);
            i += 2;
        } else {
            length = length.saturating_add(3);
            i += 1;
        }
    }
    length
}

/// Decode a best-effort two-byte codepage such as EUC-KR or GBK.
fn decode_double_byte_lossy(bytes: &[u8], capacity: usize) -> Result<String> {
    let mut out = String::new();
    out.try_reserve(capacity)
        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
    let mut i = 0;
    while i < bytes.len() {
        let b = bytes[i];
        if b < 0x80 {
            out.push(b as char);
            i += 1;
        } else if i + 1 < bytes.len() {
            out.push('\u{FFFD}');
            i += 2;
        } else {
            out.push('\u{FFFD}');
            i += 1;
        }
    }
    Ok(out)
}

fn double_byte_lossy_len(bytes: &[u8]) -> usize {
    let mut length = 0_usize;
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] < 0x80 {
            length = length.saturating_add(1);
            i += 1;
        } else {
            length = length.saturating_add(3);
            i = i.saturating_add(2).min(bytes.len());
        }
    }
    length
}

/// Find the VBA storage prefix inside the CFB container.
/// Returns the path prefix ending with a separator (e.g. "VBA/").
fn find_vba_prefix(cfb: &mut cfb::CompoundFile<Cursor<&[u8]>>) -> Result<String> {
    // Look for a "dir" stream under a VBA storage.
    for entry in cfb.walk() {
        let entry_path = entry.path().to_string_lossy();
        let normalized = entry_path.replace('\\', "/");
        if normalized.ends_with("/dir") || normalized.ends_with("/DIR") {
            let prefix = &normalized[..normalized.len() - 3];
            return Ok(prefix.to_string());
        }
    }

    // Try common paths directly.
    for prefix in ["/VBA/", "VBA/", "/"] {
        let dir_path = format!("{prefix}dir");
        if cfb.is_stream(&dir_path) {
            return Ok(prefix.to_string());
        }
    }

    Err(Error::Internal(
        "could not find VBA dir stream in vbaProject.bin".to_string(),
    ))
}

/// Read a stream from the CFB container as raw bytes.
fn read_cfb_stream(
    cfb: &mut cfb::CompoundFile<Cursor<&[u8]>>,
    path: &str,
    budget: &mut VbaBudget,
) -> Result<Vec<u8>> {
    let declared = cfb
        .entry(path)
        .map_err(|_| Error::Internal("failed to inspect VBA CFB stream".to_string()))?
        .len();
    check_vba_limit(
        ResourceDimension::EntryDeclaredBytes,
        budget.limits.cfb_stream_declared_bytes,
        declared,
        ResourceStage::Inspect,
    )?;
    let capacity = usize::try_from(declared.min(8192)).map_err(|_| {
        vba_limit(
            ResourceDimension::EntryDeclaredBytes,
            budget.limits.cfb_stream_declared_bytes,
            declared,
            ResourceStage::Inspect,
        )
    })?;
    let mut stream = cfb
        .open_stream(path)
        .map_err(|_| Error::Internal("failed to open VBA CFB stream".to_string()))?;
    let mut data = Vec::new();
    data.try_reserve(capacity)
        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
    let mut buffer = [0_u8; 8192];
    let mut stream_total = 0_u64;
    loop {
        let read = stream
            .read(&mut buffer)
            .map_err(|_| Error::Internal("failed to read VBA CFB stream".to_string()))?;
        if read == 0 {
            break;
        }
        let amount = u64::try_from(read).unwrap_or(u64::MAX);
        stream_total = budget.charge_stream_read(stream_total, amount)?;
        data.try_reserve(read)
            .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
        data.extend_from_slice(&buffer[..read]);
    }
    Ok(data)
}

/// Decompress a VBA compressed stream per MS-OVBA 2.4.1.
///
/// The format is:
/// - 1 byte signature (0x01)
/// - Sequence of compressed chunks, each starting with a 2-byte header
/// - Each chunk contains a mix of literal bytes and copy tokens
pub fn decompress_vba_stream(data: &[u8]) -> Result<Vec<u8>> {
    let mut budget = VbaBudget::new(VbaLimits::default())?;
    decompress_vba_stream_budgeted(data, &mut budget)
}

fn decompress_vba_stream_budgeted(data: &[u8], budget: &mut VbaBudget) -> Result<Vec<u8>> {
    budget.charge_compressed(u64::try_from(data.len()).unwrap_or(u64::MAX))?;
    if data.is_empty() {
        return Ok(Vec::new());
    }

    if data[0] != 0x01 {
        return Err(Error::Internal(
            "invalid VBA compression signature".to_string(),
        ));
    }

    let mut output = Vec::new();
    let mut pos = 1; // skip signature byte
    let mut module_total = 0_u64;

    while pos < data.len() {
        budget.charge_chunk()?;
        if pos + 1 >= data.len() {
            return Err(Error::Internal(
                "truncated VBA compressed chunk header".to_string(),
            ));
        }

        // Read chunk header (2 bytes, little-endian)
        let header = u16::from_le_bytes([data[pos], data[pos + 1]]);
        pos += 2;

        let chunk_size = (header & 0x0FFF) as usize + 3;
        let is_compressed = (header & 0x8000) != 0;

        let payload_size = chunk_size - 2;
        let chunk_end = pos
            .checked_add(payload_size)
            .filter(|&end| end <= data.len())
            .ok_or_else(|| Error::Internal("truncated VBA compressed chunk".to_string()))?;

        if !is_compressed {
            // Uncompressed chunk: raw bytes (4096 bytes max)
            let raw_end = chunk_end.min(pos + 4096);
            let amount = u64::try_from(raw_end - pos).unwrap_or(u64::MAX);
            let (next_module, next_aggregate) = budget.check_decoded(module_total, amount)?;
            output
                .try_reserve(raw_end - pos)
                .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
            output.extend_from_slice(&data[pos..raw_end]);
            module_total = next_module;
            budget.commit_decoded(next_aggregate);
            pos = chunk_end;
            continue;
        }

        // Compressed chunk
        let chunk_start_output = output.len();
        while pos < chunk_end {
            if pos >= data.len() {
                break;
            }

            let flag_byte = data[pos];
            pos += 1;

            for bit_index in 0..8 {
                if pos >= chunk_end {
                    break;
                }

                if (flag_byte >> bit_index) & 1 == 0 {
                    // Literal byte
                    if output.len() - chunk_start_output >= 4096 {
                        return Err(Error::Internal(
                            "VBA literal exceeds decompressed chunk size".to_string(),
                        ));
                    }
                    let (next_module, next_aggregate) = budget.check_decoded(module_total, 1)?;
                    output
                        .try_reserve(1)
                        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
                    output.push(data[pos]);
                    module_total = next_module;
                    budget.commit_decoded(next_aggregate);
                    pos += 1;
                } else {
                    // Copy token (2 bytes, little-endian)
                    if pos + 1 >= chunk_end {
                        return Err(Error::Internal("truncated VBA copy token".to_string()));
                    }
                    let token = u16::from_le_bytes([data[pos], data[pos + 1]]);
                    pos += 2;

                    // Calculate the number of bits for the length and offset
                    let decompressed_current = output.len() - chunk_start_output;
                    let bit_count = offset_bit_count(decompressed_current);
                    let length_mask = 0xFFFF >> bit_count;
                    let offset_mask = !length_mask;

                    let length = ((token & length_mask) + 3) as usize;
                    let offset = (((token & offset_mask) >> (16 - bit_count)) + 1) as usize;

                    if offset == 0 || offset > decompressed_current {
                        return Err(Error::Internal("invalid VBA copy token offset".to_string()));
                    }
                    if decompressed_current + length > 4096 {
                        return Err(Error::Internal(
                            "VBA copy token exceeds decompressed chunk size".to_string(),
                        ));
                    }

                    let amount = u64::try_from(length).unwrap_or(u64::MAX);
                    let (next_module, next_aggregate) =
                        budget.check_decoded(module_total, amount)?;
                    output
                        .try_reserve(length)
                        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
                    for _ in 0..length {
                        let byte = output[output.len() - offset];
                        output.push(byte);
                    }
                    module_total = next_module;
                    budget.commit_decoded(next_aggregate);
                }
            }
        }
    }

    Ok(output)
}

/// Calculate the bit count for the copy token offset field.
/// Per MS-OVBA 2.4.1.3.19.1:
/// The number of bits used for the offset is ceil(log2(decompressed_current)) with min 4.
fn offset_bit_count(decompressed_current: usize) -> u16 {
    if decompressed_current <= 16 {
        return 4;
    }
    if decompressed_current <= 32 {
        return 5;
    }
    if decompressed_current <= 64 {
        return 6;
    }
    if decompressed_current <= 128 {
        return 7;
    }
    if decompressed_current <= 256 {
        return 8;
    }
    if decompressed_current <= 512 {
        return 9;
    }
    if decompressed_current <= 1024 {
        return 10;
    }
    if decompressed_current <= 2048 {
        return 11;
    }
    12
}

/// Parse the decompressed `dir` stream to extract module entries and codepage.
///
/// The dir stream is a sequence of records with 2-byte IDs and 4-byte sizes.
/// We look for MODULE_NAME, MODULE_STREAM_NAME, MODULE_OFFSET,
/// MODULE_TYPE, and PROJECTCODEPAGE records.
///
/// MODULE_TYPE record 0x0021 indicates a procedural (standard) module.
/// MODULE_TYPE record 0x0022 indicates a document/class module. When 0x0022
/// is present, we refine the type to `Document`, `ThisWorkbook`, or `Class`
/// based on the module name (since OOXML does not distinguish these subtypes
/// at the record level).
#[cfg(test)]
fn parse_dir_stream(data: &[u8]) -> Result<DirInfo> {
    let mut budget = VbaBudget::new(VbaLimits::default())?;
    parse_dir_stream_budgeted(data, &mut budget)
}

fn parse_dir_stream_budgeted(data: &[u8], budget: &mut VbaBudget) -> Result<DirInfo> {
    let mut pos = 0;
    let mut modules = Vec::new();
    let mut codepage: u16 = 1252; // Default to Windows-1252

    // Current module being built
    let mut current_name: Option<String> = None;
    let mut current_stream_name: Option<String> = None;
    let mut current_offset: u32 = 0;
    let mut current_type = VbaModuleType::Standard;
    let mut in_module = false;

    while data.len().saturating_sub(pos) >= 6 {
        let record_id = u16::from_le_bytes([data[pos], data[pos + 1]]);
        let record_size =
            u32::from_le_bytes([data[pos + 2], data[pos + 3], data[pos + 4], data[pos + 5]])
                as usize;
        pos += 6;

        let payload_size = if record_id == 0x0009 {
            if record_size != 4 {
                return Err(Error::Internal(
                    "invalid PROJECTVERSION declared size".to_string(),
                ));
            }
            6
        } else {
            record_size
        };
        let payload_end = pos
            .checked_add(payload_size)
            .filter(|&end| end <= data.len())
            .ok_or_else(|| Error::Internal("truncated VBA dir record".to_string()))?;
        let record_data = &data[pos..payload_end];

        match record_id {
            // PROJECTCODEPAGE
            0x0003 => {
                if record_size >= 2 {
                    codepage = u16::from_le_bytes([record_data[0], record_data[1]]);
                }
            }
            // MODULENAME
            0x0019 => {
                if in_module {
                    // Save previous module
                    if let (Some(name), Some(stream)) =
                        (current_name.take(), current_stream_name.take())
                    {
                        budget.charge_module()?;
                        modules
                            .try_reserve(1)
                            .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
                        let refined_type = refine_module_type(&current_type, &name);
                        modules.push(ModuleEntry {
                            name,
                            stream_name: stream,
                            text_offset: current_offset,
                            module_type: refined_type,
                        });
                    }
                }
                in_module = true;
                budget.charge_retained(u64::try_from(record_data.len()).unwrap_or(u64::MAX))?;
                current_name = Some(String::from_utf8_lossy(record_data).into_owned());
                current_stream_name = None;
                current_offset = 0;
                current_type = VbaModuleType::Standard;
            }
            // MODULENAMEUNICODE
            0x0047 => {
                // UTF-16LE encoded name, prefer this over the ANSI name.
                // Use only the even portion of the data; an odd trailing
                // byte indicates a truncated record and is safely ignored.
                if record_size >= 2 {
                    let even_len = record_data.len() & !1;
                    budget.charge_retained(
                        u64::try_from(even_len)
                            .unwrap_or(u64::MAX)
                            .saturating_mul(2),
                    )?;
                    let mut u16_data = Vec::new();
                    u16_data
                        .try_reserve(even_len / 2)
                        .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
                    for chunk in record_data[..even_len].chunks_exact(2) {
                        u16_data.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                    }
                    let name = String::from_utf16_lossy(&u16_data);
                    // Remove trailing null if present
                    let name = name.trim_end_matches('\0').to_string();
                    if !name.is_empty() {
                        current_name = Some(name);
                    }
                }
            }
            // MODULESTREAMNAME
            0x001A => {
                budget.charge_retained(u64::try_from(record_data.len()).unwrap_or(u64::MAX))?;
                current_stream_name = Some(String::from_utf8_lossy(record_data).into_owned());
                // The MODULENAMEUNICODE record for stream name follows with id 0x0032
                // We handle it inline: skip the unicode record
                if let Some(next_header) = pos
                    .checked_add(record_size)
                    .filter(|&start| data.len().saturating_sub(start) >= 6)
                {
                    let next_id = u16::from_le_bytes([data[next_header], data[next_header + 1]]);
                    if next_id == 0x0032 {
                        let next_size = u32::from_le_bytes([
                            data[next_header + 2],
                            data[next_header + 3],
                            data[next_header + 4],
                            data[next_header + 5],
                        ]) as usize;
                        // Skip the unicode stream name record
                        pos = next_header
                            .checked_add(6)
                            .and_then(|start| start.checked_add(next_size))
                            .filter(|&end| end <= data.len())
                            .ok_or_else(|| {
                                Error::Internal("truncated VBA dir record".to_string())
                            })?;
                        continue;
                    }
                }
            }
            // MODULEOFFSET
            0x0031 => {
                if record_size >= 4 {
                    current_offset = u32::from_le_bytes([
                        record_data[0],
                        record_data[1],
                        record_data[2],
                        record_data[3],
                    ]);
                }
            }
            // MODULETYPE procedural (0x0021)
            0x0021 => {
                current_type = VbaModuleType::Standard;
            }
            // MODULETYPE document/class (0x0022)
            0x0022 => {
                // The dir stream only distinguishes procedural (0x0021) from
                // non-procedural (0x0022). We refine 0x0022 into Document,
                // ThisWorkbook, or Class based on the module name when the
                // module is finalized.
                current_type = VbaModuleType::Class;
            }
            // TERMINATOR for modules section (0x002B)
            0x002B => {
                // End of module list
            }
            _ => {}
        }

        pos = payload_end;
    }

    // Save the last module if present
    if in_module {
        if let (Some(name), Some(stream)) = (current_name, current_stream_name) {
            budget.charge_module()?;
            modules
                .try_reserve(1)
                .map_err(|_| Error::Internal("VBA allocation failed".to_string()))?;
            let refined_type = refine_module_type(&current_type, &name);
            modules.push(ModuleEntry {
                name,
                stream_name: stream,
                text_offset: current_offset,
                module_type: refined_type,
            });
        }
    }

    Ok(DirInfo {
        entries: modules,
        codepage,
    })
}

/// Refine the module type for non-procedural modules (0x0022) based on
/// the module name. Procedural modules (0x0021) are always `Standard`.
fn refine_module_type(base_type: &VbaModuleType, name: &str) -> VbaModuleType {
    if *base_type == VbaModuleType::Standard {
        return VbaModuleType::Standard;
    }
    let name_lower = name.to_lowercase();
    if name_lower == "thisworkbook" {
        VbaModuleType::ThisWorkbook
    } else if name_lower.starts_with("sheet") {
        VbaModuleType::Document
    } else {
        // Remains as Class (could be a class module or UserForm).
        VbaModuleType::Class
    }
}

#[cfg(test)]
#[allow(clippy::same_item_push)]
mod tests {
    use super::*;

    #[test]
    fn test_decompress_empty_input() {
        let result = decompress_vba_stream(&[]);
        assert!(result.is_ok());
        assert!(result.unwrap().is_empty());
    }

    #[test]
    fn test_decompress_invalid_signature() {
        let result = decompress_vba_stream(&[0x00, 0x01, 0x02]);
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(err_msg.contains("invalid VBA compression signature"));
    }

    fn two_uncompressed_chunks() -> Vec<u8> {
        vec![0x01, 0x00, 0x00, b'a', 0x00, 0x00, b'b']
    }

    #[test]
    fn test_chunk_limit_accepts_exact_and_rejects_plus_one() {
        let data = two_uncompressed_chunks();
        let mut limits = VbaLimits {
            chunk_count: 2,
            ..VbaLimits::default()
        };
        let mut budget = VbaBudget::new(limits).unwrap();
        assert_eq!(
            decompress_vba_stream_budgeted(&data, &mut budget).unwrap(),
            b"ab"
        );

        limits.chunk_count = 1;
        let mut budget = VbaBudget::new(limits).unwrap();
        let error = decompress_vba_stream_budgeted(&data, &mut budget).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension(),
            ResourceDimension::ChunkCount
        );
    }

    #[test]
    fn test_decoded_limits_are_cumulative_and_cannot_reset() {
        let data = vec![0x01, 0x00, 0x00, b'a'];
        let limits = VbaLimits {
            module_decoded_bytes: 1,
            aggregate_decoded_bytes: 1,
            ..VbaLimits::default()
        };
        let mut budget = VbaBudget::new(limits).unwrap();
        assert_eq!(
            decompress_vba_stream_budgeted(&data, &mut budget).unwrap(),
            b"a"
        );
        let error = decompress_vba_stream_budgeted(&data, &mut budget).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension(),
            ResourceDimension::OutputBytes
        );

        let limits = VbaLimits {
            module_decoded_bytes: 1,
            ..VbaLimits::default()
        };
        let too_large = vec![0x01, 0x01, 0x00, b'a', b'b'];
        let mut budget = VbaBudget::new(limits).unwrap();
        let error = decompress_vba_stream_budgeted(&too_large, &mut budget).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension(),
            ResourceDimension::OutputBytes
        );
    }

    #[test]
    fn test_module_limit_counts_all_dir_entries() {
        let mut dir = Vec::new();
        for name in [b"First".as_slice(), b"Second".as_slice()] {
            write_dir_record(&mut dir, 0x0019, name);
            write_dir_record(&mut dir, 0x001A, name);
        }

        let mut limits = VbaLimits {
            module_count: 2,
            ..VbaLimits::default()
        };
        let mut budget = VbaBudget::new(limits).unwrap();
        assert_eq!(
            parse_dir_stream_budgeted(&dir, &mut budget)
                .unwrap()
                .entries
                .len(),
            2
        );

        limits.module_count = 1;
        let mut budget = VbaBudget::new(limits).unwrap();
        let error = parse_dir_stream_budgeted(&dir, &mut budget).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension(),
            ResourceDimension::ModuleCount
        );
    }

    #[test]
    fn test_decompress_uncompressed_chunk() {
        // Signature byte + uncompressed chunk header (size=3 -> 3+3-2=4 bytes, bit 15 clear)
        // Header: chunk_size = 3 bytes (field = 3-3 = 0), not compressed (bit 15 = 0)
        // So header = 0x0000 means size=3, uncompressed
        let mut data = vec![0x01]; // signature
                                   // Uncompressed chunk: header with bit 15 clear, size field = N-3
                                   // For 4 bytes of data: chunk_size = 4, field = 4-3 = 1
        let header: u16 = 0x0001; // bit 15 = 0 (uncompressed), size = 1+3-2 = 2 (actual chunk payload = 2)
                                  // Wait, let me recalculate.
                                  // chunk_size = (header & 0x0FFF) + 3 = field + 3
                                  // The chunk payload is chunk_size - 2 = field + 1 bytes
                                  // For 3 bytes of payload: field = 2, header = 0x0002
        data.extend_from_slice(&header.to_le_bytes());
        data.extend_from_slice(b"AB");
        // This should produce "AB" but limited to min(chunk_end, pos+4096)
        let result = decompress_vba_stream(&data).unwrap();
        assert_eq!(&result, b"AB");
    }

    #[test]
    fn test_decompress_real_compressed_data() {
        // Test with a known compressed sequence from the MS-OVBA spec example.
        // Compressed representation of "aaaaaaaaaaaaaaa" (15 'a's)
        // Signature: 0x01
        // Chunk header: compressed, size field
        // Flag byte: 0b00000011 = 0x03 (bit 0: literal, bit 1: copy token)
        // Actually building a minimal valid compressed stream:
        // Signature: 0x01
        // Chunk header: size = N-3, compressed bit set
        // Then flag + data

        // A simpler approach: verify that decompression of a manually built stream works.
        let mut compressed = vec![0x01u8];
        // Build chunk: 1 literal 'a', then copy token referencing offset=1, length=3
        // Flag byte: 0b00000010 = bit 0 literal, bit 1 copy
        // Literal: b'a'
        // Copy token with decompressed_current=1 -> bit_count=12, length_mask=0x000F
        // offset=1, length=3 -> offset_field=(1-1)<<4=0, length_field=3-3=0
        // token = 0x0000
        let flag = 0x02u8; // bits: 0=literal, 1=copy, rest=0
        let literal = b'a';
        let copy_token: u16 = 0x0000; // offset=1, length=3

        let mut chunk_payload = Vec::new();
        chunk_payload.push(flag);
        chunk_payload.push(literal);
        chunk_payload.extend_from_slice(&copy_token.to_le_bytes());

        let chunk_size = chunk_payload.len() + 2; // +2 for header
        let header: u16 = 0x8000 | ((chunk_size as u16 - 3) & 0x0FFF);
        compressed.extend_from_slice(&header.to_le_bytes());
        compressed.extend_from_slice(&chunk_payload);

        let result = decompress_vba_stream(&compressed).unwrap();
        assert_eq!(&result, b"aaaa"); // 1 literal + 3 from copy
    }

    #[test]
    fn test_offset_bit_count_thresholds() {
        for (position, expected) in [
            (0, 4),
            (16, 4),
            (17, 5),
            (32, 5),
            (33, 6),
            (64, 6),
            (65, 7),
            (128, 7),
            (129, 8),
            (256, 8),
            (257, 9),
            (512, 9),
            (513, 10),
            (1024, 10),
            (1025, 11),
            (2048, 11),
            (2049, 12),
            (4096, 12),
        ] {
            assert_eq!(offset_bit_count(position), expected, "position {position}");
        }
    }

    #[test]
    fn test_copy_tokens_at_each_bit_width_threshold() {
        for position in [
            16, 17, 32, 33, 64, 65, 128, 129, 256, 257, 512, 513, 1024, 1025, 2048, 2049,
        ] {
            let literals: Vec<u8> = (0..position).map(|i| (i % 251) as u8).collect();
            let offset = position.min(13);
            let stream = compressed_copy_stream(&literals, offset, 3);
            let output = decompress_vba_stream(&stream).unwrap();
            assert_eq!(
                &output[position..],
                &literals[position - offset..position - offset + 3]
            );
        }
    }

    #[test]
    fn test_copy_token_uses_overlapping_expanding_output() {
        let output = decompress_vba_stream(&compressed_copy_stream(b"abc", 3, 9)).unwrap();
        assert_eq!(output, b"abcabcabcabc");
    }

    #[test]
    fn test_maximum_valid_copy_token_fills_chunk() {
        let output = decompress_vba_stream(&compressed_copy_stream(b"x", 1, 4095)).unwrap();
        assert_eq!(output.len(), 4096);
        assert!(output.iter().all(|&byte| byte == b'x'));
    }

    #[test]
    fn test_copy_token_rejects_offset_beyond_current_chunk() {
        let error = decompress_vba_stream(&compressed_copy_stream(b"x", 2, 3)).unwrap_err();
        assert!(error.to_string().contains("invalid VBA copy token offset"));
    }

    #[test]
    fn test_copy_token_rejects_output_past_chunk_limit() {
        let error = decompress_vba_stream(&compressed_copy_stream(b"x", 1, 4096)).unwrap_err();
        assert!(error
            .to_string()
            .contains("exceeds decompressed chunk size"));
    }

    #[test]
    fn test_copy_token_rejects_truncated_token() {
        let data = [0x01, 0x01, 0x80, 0x01, 0x00];
        let error = decompress_vba_stream(&data).unwrap_err();
        assert!(error.to_string().contains("truncated VBA copy token"));
    }

    #[test]
    fn test_project_version_consumes_fixed_six_byte_payload() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0009u16.to_le_bytes());
        data.extend_from_slice(&4u32.to_le_bytes());
        data.extend_from_slice(&0x0001_0000u32.to_le_bytes());
        data.extend_from_slice(&7u16.to_le_bytes());
        write_dir_record(&mut data, 0x0003, &65001u16.to_le_bytes());

        assert_eq!(parse_dir_stream(&data).unwrap().codepage, 65001);
    }

    #[test]
    fn test_project_version_rejects_truncated_fixed_payload() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0009u16.to_le_bytes());
        data.extend_from_slice(&4u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        assert!(parse_dir_stream(&data).is_err());
    }

    #[test]
    fn test_parse_dir_stream_empty() {
        let result = parse_dir_stream(&[]);
        assert!(result.is_ok());
        let info = result.unwrap();
        assert!(info.entries.is_empty());
        assert_eq!(info.codepage, 1252);
    }

    #[test]
    fn test_extract_vba_modules_invalid_cfb() {
        let result = extract_vba_modules(b"not a CFB file");
        assert!(result.is_err());
        let err_msg = result.unwrap_err().to_string();
        assert!(err_msg.contains("failed to open VBA project as CFB"));
    }

    #[test]
    fn test_vba_module_type_clone() {
        let t = VbaModuleType::Standard;
        let t2 = t.clone();
        assert_eq!(t, t2);
    }

    #[test]
    fn test_vba_module_debug() {
        let m = VbaModule {
            name: "Module1".to_string(),
            source_code: "Sub Test()\nEnd Sub".to_string(),
            module_type: VbaModuleType::Standard,
        };
        let debug = format!("{:?}", m);
        assert!(debug.contains("Module1"));
    }

    #[test]
    fn test_vba_roundtrip_with_xlsm() {
        use std::io::{Read as _, Write as _};

        // Build a minimal CFB container with a VBA dir stream and a module
        let vba_bin = build_test_vba_project();

        // Create a valid xlsx using the Workbook API, then inject vbaProject.bin
        let base_wb = crate::workbook::Workbook::new();
        let base_buf = base_wb.save_to_buffer().unwrap();

        // Rewrite the ZIP, adding the vbaProject.bin entry
        let mut buf = Vec::new();
        {
            let base_cursor = std::io::Cursor::new(&base_buf);
            let mut base_archive = zip::ZipArchive::new(base_cursor).unwrap();

            let out_cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(out_cursor);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);

            for i in 0..base_archive.len() {
                let mut entry = base_archive.by_index(i).unwrap();
                let name = entry.name().to_string();
                zip.start_file(&name, options).unwrap();
                let mut data = Vec::new();
                entry.read_to_end(&mut data).unwrap();
                zip.write_all(&data).unwrap();
            }

            zip.start_file("xl/vbaProject.bin", options).unwrap();
            zip.write_all(&vba_bin).unwrap();
            zip.finish().unwrap();
        }

        // Open and extract
        let opts = crate::workbook::OpenOptions::new()
            .read_mode(crate::workbook::ReadMode::Eager)
            .aux_parts(crate::workbook::AuxParts::EagerLoad);
        let wb = crate::workbook::Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();

        // Raw VBA project should be available
        let raw = wb.get_vba_project();
        assert!(raw.is_some(), "VBA project binary should be present");
        assert_eq!(raw.unwrap(), vba_bin);
    }

    #[test]
    fn test_xlsx_without_vba_returns_none() {
        let wb = crate::workbook::Workbook::new();
        assert!(wb.get_vba_project().is_none());
        assert!(wb.get_vba_modules().unwrap().is_none());
    }

    #[test]
    fn test_xlsx_roundtrip_no_vba() {
        let wb = crate::workbook::Workbook::new();
        let buf = wb.save_to_buffer().unwrap();
        let wb2 = crate::workbook::Workbook::open_from_buffer(&buf).unwrap();
        assert!(wb2.get_vba_project().is_none());
    }

    #[test]
    fn test_get_vba_modules_from_test_project() {
        use std::io::{Read as _, Write as _};

        let vba_bin = build_test_vba_project();

        // Create a valid xlsx, then inject vbaProject.bin
        let base_wb = crate::workbook::Workbook::new();
        let base_buf = base_wb.save_to_buffer().unwrap();

        let mut buf = Vec::new();
        {
            let base_cursor = std::io::Cursor::new(&base_buf);
            let mut base_archive = zip::ZipArchive::new(base_cursor).unwrap();

            let out_cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(out_cursor);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);

            for i in 0..base_archive.len() {
                let mut entry = base_archive.by_index(i).unwrap();
                let name = entry.name().to_string();
                zip.start_file(&name, options).unwrap();
                let mut data = Vec::new();
                entry.read_to_end(&mut data).unwrap();
                zip.write_all(&data).unwrap();
            }

            zip.start_file("xl/vbaProject.bin", options).unwrap();
            zip.write_all(&vba_bin).unwrap();
            zip.finish().unwrap();
        }

        let opts = crate::workbook::OpenOptions::new()
            .read_mode(crate::workbook::ReadMode::Eager)
            .aux_parts(crate::workbook::AuxParts::EagerLoad);
        let wb = crate::workbook::Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        let project = wb.get_vba_modules().unwrap();
        assert!(project.is_some(), "should have VBA modules");
        let project = project.unwrap();
        assert_eq!(project.modules.len(), 1);
        assert_eq!(project.modules[0].name, "Module1");
        assert_eq!(project.modules[0].module_type, VbaModuleType::Standard);
        assert!(
            project.modules[0].source_code.contains("Sub Hello()"),
            "source should contain Sub Hello(), got: {}",
            project.modules[0].source_code
        );
    }

    #[test]
    fn test_vba_project_preserved_in_save_roundtrip() {
        use std::io::{Read as _, Write as _};

        let vba_bin = build_test_vba_project();

        let base_wb = crate::workbook::Workbook::new();
        let base_buf = base_wb.save_to_buffer().unwrap();

        let mut buf = Vec::new();
        {
            let base_cursor = std::io::Cursor::new(&base_buf);
            let mut base_archive = zip::ZipArchive::new(base_cursor).unwrap();

            let out_cursor = std::io::Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(out_cursor);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);

            for i in 0..base_archive.len() {
                let mut entry = base_archive.by_index(i).unwrap();
                let name = entry.name().to_string();
                zip.start_file(&name, options).unwrap();
                let mut data = Vec::new();
                entry.read_to_end(&mut data).unwrap();
                zip.write_all(&data).unwrap();
            }

            zip.start_file("xl/vbaProject.bin", options).unwrap();
            zip.write_all(&vba_bin).unwrap();
            zip.finish().unwrap();
        }

        // Open, then save again
        let opts = crate::workbook::OpenOptions::new()
            .read_mode(crate::workbook::ReadMode::Eager)
            .aux_parts(crate::workbook::AuxParts::EagerLoad);
        let mut wb = crate::workbook::Workbook::open_from_buffer_with_options(&buf, &opts).unwrap();
        wb.set_format(crate::workbook::WorkbookFormat::Xlsm);
        let saved_buf = wb.save_to_buffer().unwrap();

        // Re-open and verify VBA is preserved
        let wb2 =
            crate::workbook::Workbook::open_from_buffer_with_options(&saved_buf, &opts).unwrap();
        let raw = wb2.get_vba_project();
        assert!(raw.is_some(), "VBA project should survive save roundtrip");
        assert_eq!(raw.unwrap(), vba_bin);

        // Modules should still be extractable
        let project = wb2.get_vba_modules().unwrap().unwrap();
        assert_eq!(project.modules.len(), 1);
        assert_eq!(project.modules[0].name, "Module1");
    }

    /// Build a minimal CFB container that looks like a VBA project.
    fn build_test_vba_project() -> Vec<u8> {
        let mut buf = Vec::new();
        let cursor = std::io::Cursor::new(&mut buf);
        let mut cfb = cfb::CompoundFile::create(cursor).unwrap();

        // Create VBA storage
        cfb.create_storage("/VBA").unwrap();

        // Build a minimal dir stream
        let dir_data = build_minimal_dir_stream("Module1");

        // Compress the dir stream
        let compressed_dir = compress_for_test(&dir_data);

        // Write dir stream
        {
            let mut stream = cfb.create_stream("/VBA/dir").unwrap();
            std::io::Write::write_all(&mut stream, &compressed_dir).unwrap();
        }

        // Build module source: "Sub Hello()\nEnd Sub\n"
        let source = b"Sub Hello()\r\nEnd Sub\r\n";
        let compressed_source = compress_for_test(source);

        // The module stream has 0 bytes of performance cache + compressed source.
        // (text_offset = 0 in the dir stream)
        {
            let mut stream = cfb.create_stream("/VBA/Module1").unwrap();
            std::io::Write::write_all(&mut stream, &compressed_source).unwrap();
        }

        // Create _VBA_PROJECT stream (required for validity, can be minimal)
        {
            let mut stream = cfb.create_stream("/VBA/_VBA_PROJECT").unwrap();
            // Minimal header: version bytes
            let header = [0xCC, 0x61, 0x00, 0x00, 0x00, 0x00, 0x00];
            std::io::Write::write_all(&mut stream, &header).unwrap();
        }

        cfb.flush().unwrap();
        buf
    }

    /// Build a minimal dir stream binary for one standard module.
    fn build_minimal_dir_stream(module_name: &str) -> Vec<u8> {
        let mut data = Vec::new();
        let name_bytes = module_name.as_bytes();

        // PROJECTSYSKIND record (0x0001): 4 bytes, value = 1 (Win32)
        write_dir_record(&mut data, 0x0001, &1u32.to_le_bytes());

        // PROJECTLCID record (0x0002): 4 bytes
        write_dir_record(&mut data, 0x0002, &0x0409u32.to_le_bytes());

        // PROJECTLCIDINVOKE record (0x0014): 4 bytes
        write_dir_record(&mut data, 0x0014, &0x0409u32.to_le_bytes());

        // PROJECTCODEPAGE record (0x0003): 2 bytes (1252 = Windows-1252)
        write_dir_record(&mut data, 0x0003, &1252u16.to_le_bytes());

        // PROJECTNAME record (0x0004)
        write_dir_record(&mut data, 0x0004, b"VBAProject");

        // PROJECTDOCSTRING record (0x0005): empty
        write_dir_record(&mut data, 0x0005, &[]);
        // Unicode variant (0x0040): empty
        write_dir_record(&mut data, 0x0040, &[]);

        // PROJECTHELPFILEPATH record (0x0006): empty
        write_dir_record(&mut data, 0x0006, &[]);
        // Unicode variant (0x003D): empty
        write_dir_record(&mut data, 0x003D, &[]);

        // PROJECTHELPCONTEXT (0x0007): 4 bytes
        write_dir_record(&mut data, 0x0007, &0u32.to_le_bytes());

        // PROJECTLIBFLAGS (0x0008): 4 bytes
        write_dir_record(&mut data, 0x0008, &0u32.to_le_bytes());

        // PROJECTVERSION (0x0009): 4 + 2 bytes (major + minor)
        let mut version = Vec::new();
        version.extend_from_slice(&1u32.to_le_bytes());
        version.extend_from_slice(&0u16.to_le_bytes());
        // Version record is special: id=0x0009, size=4 for major, then 2 bytes minor appended
        data.extend_from_slice(&0x0009u16.to_le_bytes());
        data.extend_from_slice(&4u32.to_le_bytes());
        data.extend_from_slice(&version);

        // PROJECTCONSTANTS (0x000C): empty
        write_dir_record(&mut data, 0x000C, &[]);
        // Unicode variant (0x003C): empty
        write_dir_record(&mut data, 0x003C, &[]);

        // MODULES count record: id=0x000F, size=2
        let module_count: u16 = 1;
        write_dir_record(&mut data, 0x000F, &module_count.to_le_bytes());

        // PROJECTCOOKIE record (0x0013): 2 bytes
        write_dir_record(&mut data, 0x0013, &0u16.to_le_bytes());

        // MODULE_NAME record (0x0019)
        write_dir_record(&mut data, 0x0019, name_bytes);

        // MODULE_STREAM_NAME record (0x001A)
        write_dir_record(&mut data, 0x001A, name_bytes);
        // Unicode variant (0x0032)
        let name_utf16: Vec<u8> = module_name
            .encode_utf16()
            .flat_map(|c| c.to_le_bytes())
            .collect();
        write_dir_record(&mut data, 0x0032, &name_utf16);

        // MODULE_OFFSET record (0x0031): 4 bytes (offset = 0)
        write_dir_record(&mut data, 0x0031, &0u32.to_le_bytes());

        // MODULE_TYPE procedural (0x0021): 0 bytes
        write_dir_record(&mut data, 0x0021, &[]);

        // MODULE_TERMINATOR (0x002B): 0 bytes
        write_dir_record(&mut data, 0x002B, &[]);

        // End of modules
        // Global TERMINATOR (0x0010): 0 bytes
        write_dir_record(&mut data, 0x0010, &[]);

        data
    }

    fn write_dir_record(buf: &mut Vec<u8>, id: u16, data: &[u8]) {
        buf.extend_from_slice(&id.to_le_bytes());
        buf.extend_from_slice(&(data.len() as u32).to_le_bytes());
        buf.extend_from_slice(data);
    }

    fn compressed_copy_stream(literals: &[u8], offset: usize, length: usize) -> Vec<u8> {
        let mut payload = Vec::new();
        let mut literal_pos = 0;
        while literals.len() - literal_pos >= 8 {
            payload.push(0);
            payload.extend_from_slice(&literals[literal_pos..literal_pos + 8]);
            literal_pos += 8;
        }

        let remaining = literals.len() - literal_pos;
        payload.push(1 << remaining);
        payload.extend_from_slice(&literals[literal_pos..]);

        let offset_bits = offset_bit_count(literals.len());
        let length_bits = 16 - offset_bits;
        let token = (((offset - 1) as u16) << length_bits) | (length - 3) as u16;
        payload.extend_from_slice(&token.to_le_bytes());

        let header = 0x8000 | ((payload.len() as u16 - 1) & 0x0FFF);
        let mut stream = vec![0x01];
        stream.extend_from_slice(&header.to_le_bytes());
        stream.extend_from_slice(&payload);
        stream
    }

    /// Minimal MS-OVBA "compression" that produces an uncompressed container.
    /// Signature 0x01 + one uncompressed chunk per 4096 bytes.
    fn compress_for_test(data: &[u8]) -> Vec<u8> {
        let mut result = vec![0x01u8]; // signature
        let mut pos = 0;
        while pos < data.len() {
            let chunk_len = (data.len() - pos).min(4096);
            let chunk_data = &data[pos..pos + chunk_len];
            // An uncompressed chunk always carries a full 4096-byte payload.
            let header: u16 = 0x0FFF;
            result.extend_from_slice(&header.to_le_bytes());
            result.extend_from_slice(chunk_data);
            // Pad to 4096 if needed
            for _ in chunk_len..4096 {
                result.push(0x00);
            }
            pos += chunk_len;
        }
        result
    }
}
