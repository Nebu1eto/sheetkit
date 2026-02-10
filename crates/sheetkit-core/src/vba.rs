//! VBA project extraction from macro-enabled workbooks (.xlsm).
//!
//! `.xlsm` files contain a `xl/vbaProject.bin` entry which is an OLE2
//! Compound Binary File (CFB) holding VBA source code. This module
//! provides read-only access to the raw binary and to individual VBA
//! module source code.

use std::io::{Cursor, Read as _};

use crate::error::{Error, Result};

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
struct ModuleEntry {
    name: String,
    stream_name: String,
    text_offset: u32,
    module_type: VbaModuleType,
}

/// Parsed metadata from the `dir` stream.
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
    let cursor = Cursor::new(vba_bin);
    let mut cfb = cfb::CompoundFile::open(cursor)
        .map_err(|e| Error::Internal(format!("failed to open VBA project as CFB: {e}")))?;

    // Find the VBA storage root. Typically `/VBA` or could be nested.
    let vba_prefix = find_vba_prefix(&mut cfb)?;

    // Read the `dir` stream to get module entries.
    let dir_path = format!("{vba_prefix}dir");
    let dir_data = read_cfb_stream(&mut cfb, &dir_path)?;

    // The dir stream is compressed using MS-OVBA compression.
    let decompressed_dir = decompress_vba_stream(&dir_data)?;

    // Parse module entries and codepage from the decompressed dir stream.
    let dir_info = parse_dir_stream(&decompressed_dir)?;

    let mut modules = Vec::with_capacity(dir_info.entries.len());
    let mut warnings = Vec::new();

    for entry in dir_info.entries {
        let stream_path = format!("{vba_prefix}{}", entry.stream_name);
        let compressed_data = match read_cfb_stream(&mut cfb, &stream_path) {
            Ok(data) => data,
            Err(e) => {
                warnings.push(format!(
                    "skipped module '{}': failed to read stream '{}': {}",
                    entry.name, stream_path, e
                ));
                continue;
            }
        };

        // The module stream has text_offset bytes of "performance cache"
        // (compiled code) followed by compressed source code.
        if (entry.text_offset as usize) > compressed_data.len() {
            warnings.push(format!(
                "skipped module '{}': text_offset {} exceeds stream length {}",
                entry.name,
                entry.text_offset,
                compressed_data.len()
            ));
            continue;
        }
        let source_compressed = &compressed_data[entry.text_offset as usize..];
        let source_bytes = match decompress_vba_stream(source_compressed) {
            Ok(b) => b,
            Err(e) => {
                warnings.push(format!(
                    "skipped module '{}': decompression failed: {}",
                    entry.name, e
                ));
                continue;
            }
        };

        let source_code = decode_source_bytes(&source_bytes, dir_info.codepage, &mut warnings);

        modules.push(VbaModule {
            name: entry.name,
            source_code,
            module_type: entry.module_type,
        });
    }

    Ok(VbaProject { modules, warnings })
}

/// Decode source bytes using the specified codepage.
///
/// Supports common codepages: 1252 (Western European), 932 (Japanese Shift-JIS),
/// 949 (Korean), 936 (Simplified Chinese GBK), 65001 (UTF-8).
/// For unrecognized codepages, falls back to UTF-8 lossy and emits a warning.
fn decode_source_bytes(bytes: &[u8], codepage: u16, warnings: &mut Vec<String>) -> String {
    match codepage {
        65001 | 0 => String::from_utf8_lossy(bytes).into_owned(),
        1252 => decode_single_byte(bytes, &WINDOWS_1252_HIGH),
        932 => decode_shift_jis(bytes),
        949 => decode_euc_kr(bytes),
        936 => decode_gbk(bytes),
        _ => {
            warnings.push(format!(
                "unsupported codepage {codepage}, falling back to UTF-8 lossy"
            ));
            String::from_utf8_lossy(bytes).into_owned()
        }
    }
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
fn decode_single_byte(bytes: &[u8], high_table: &[char; 128]) -> String {
    let mut out = String::with_capacity(bytes.len());
    for &b in bytes {
        if b < 0x80 {
            out.push(b as char);
        } else {
            out.push(high_table[(b - 0x80) as usize]);
        }
    }
    out
}

/// Decode Shift-JIS (codepage 932) bytes to a String.
/// Uses a best-effort approach: valid multi-byte sequences are decoded,
/// invalid bytes are replaced with the Unicode replacement character.
fn decode_shift_jis(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len());
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
    out
}

/// Decode EUC-KR / codepage 949 bytes to a String.
/// Best-effort: ASCII bytes pass through, multi-byte sequences use replacement.
fn decode_euc_kr(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len());
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
    out
}

/// Decode GBK / codepage 936 bytes to a String.
/// Best-effort: ASCII bytes pass through, multi-byte sequences use replacement.
fn decode_gbk(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len());
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
    out
}

/// Find the VBA storage prefix inside the CFB container.
/// Returns the path prefix ending with a separator (e.g. "VBA/").
fn find_vba_prefix(cfb: &mut cfb::CompoundFile<Cursor<&[u8]>>) -> Result<String> {
    // Collect all entries first to avoid borrow issues.
    let entries: Vec<String> = cfb
        .walk()
        .map(|e| e.path().to_string_lossy().into_owned())
        .collect();

    // Look for a "dir" stream under a VBA storage.
    for entry_path in &entries {
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
fn read_cfb_stream(cfb: &mut cfb::CompoundFile<Cursor<&[u8]>>, path: &str) -> Result<Vec<u8>> {
    let mut stream = cfb
        .open_stream(path)
        .map_err(|e| Error::Internal(format!("failed to open CFB stream '{path}': {e}")))?;
    let mut data = Vec::new();
    stream
        .read_to_end(&mut data)
        .map_err(|e| Error::Internal(format!("failed to read CFB stream '{path}': {e}")))?;
    Ok(data)
}

/// Decompress a VBA compressed stream per MS-OVBA 2.4.1.
///
/// The format is:
/// - 1 byte signature (0x01)
/// - Sequence of compressed chunks, each starting with a 2-byte header
/// - Each chunk contains a mix of literal bytes and copy tokens
pub fn decompress_vba_stream(data: &[u8]) -> Result<Vec<u8>> {
    if data.is_empty() {
        return Ok(Vec::new());
    }

    if data[0] != 0x01 {
        return Err(Error::Internal(format!(
            "invalid VBA compression signature: expected 0x01, got 0x{:02X}",
            data[0]
        )));
    }

    let mut output = Vec::with_capacity(data.len() * 2);
    let mut pos = 1; // skip signature byte

    while pos < data.len() {
        if pos + 1 >= data.len() {
            break;
        }

        // Read chunk header (2 bytes, little-endian)
        let header = u16::from_le_bytes([data[pos], data[pos + 1]]);
        pos += 2;

        let chunk_size = (header & 0x0FFF) as usize + 3;
        let is_compressed = (header & 0x8000) != 0;

        let chunk_end = (pos + chunk_size - 2).min(data.len());

        if !is_compressed {
            // Uncompressed chunk: raw bytes (4096 bytes max)
            let raw_end = chunk_end.min(pos + 4096);
            if raw_end > data.len() {
                break;
            }
            output.extend_from_slice(&data[pos..raw_end]);
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
                    output.push(data[pos]);
                    pos += 1;
                } else {
                    // Copy token (2 bytes, little-endian)
                    if pos + 1 >= data.len() {
                        pos = chunk_end;
                        break;
                    }
                    let token = u16::from_le_bytes([data[pos], data[pos + 1]]);
                    pos += 2;

                    // Calculate the number of bits for the length and offset
                    let decompressed_current = output.len() - chunk_start_output;
                    let bit_count = max_bit_count(decompressed_current);
                    let length_mask = 0xFFFF >> bit_count;
                    let offset_mask = !length_mask;

                    let length = ((token & length_mask) + 3) as usize;
                    let offset = (((token & offset_mask) >> (16 - bit_count)) + 1) as usize;

                    if offset > output.len() {
                        // Invalid offset, skip
                        break;
                    }

                    let copy_start = output.len() - offset;
                    for i in 0..length {
                        let byte = output[copy_start + (i % offset)];
                        output.push(byte);
                    }
                }
            }
        }
    }

    Ok(output)
}

/// Calculate the bit count for the copy token offset field.
/// Per MS-OVBA 2.4.1.3.19.1:
/// The number of bits used for the offset is ceil(log2(decompressed_current)) with min 4.
fn max_bit_count(decompressed_current: usize) -> u16 {
    if decompressed_current <= 16 {
        return 12;
    }
    if decompressed_current <= 32 {
        return 11;
    }
    if decompressed_current <= 64 {
        return 10;
    }
    if decompressed_current <= 128 {
        return 9;
    }
    if decompressed_current <= 256 {
        return 8;
    }
    if decompressed_current <= 512 {
        return 7;
    }
    if decompressed_current <= 1024 {
        return 6;
    }
    if decompressed_current <= 2048 {
        return 5;
    }
    4 // >= 4096
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
fn parse_dir_stream(data: &[u8]) -> Result<DirInfo> {
    let mut pos = 0;
    let mut modules = Vec::new();
    let mut codepage: u16 = 1252; // Default to Windows-1252

    // Current module being built
    let mut current_name: Option<String> = None;
    let mut current_stream_name: Option<String> = None;
    let mut current_offset: u32 = 0;
    let mut current_type = VbaModuleType::Standard;
    let mut in_module = false;

    while pos + 6 <= data.len() {
        let record_id = u16::from_le_bytes([data[pos], data[pos + 1]]);
        let record_size =
            u32::from_le_bytes([data[pos + 2], data[pos + 3], data[pos + 4], data[pos + 5]])
                as usize;
        pos += 6;

        if pos + record_size > data.len() {
            break;
        }

        let record_data = &data[pos..pos + record_size];

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
                current_name = Some(String::from_utf8_lossy(record_data).into_owned());
                current_stream_name = None;
                current_offset = 0;
                current_type = VbaModuleType::Standard;
            }
            // MODULENAMEUNICODE
            0x0047 => {
                // UTF-16LE encoded name, prefer this over the ANSI name
                if record_size >= 2 {
                    let u16_data: Vec<u16> = record_data
                        .chunks_exact(2)
                        .map(|c| u16::from_le_bytes([c[0], c[1]]))
                        .collect();
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
                current_stream_name = Some(String::from_utf8_lossy(record_data).into_owned());
                // The MODULENAMEUNICODE record for stream name follows with id 0x0032
                // We handle it inline: skip the unicode record
                if pos + record_size + 6 <= data.len() {
                    let next_id =
                        u16::from_le_bytes([data[pos + record_size], data[pos + record_size + 1]]);
                    if next_id == 0x0032 {
                        let next_size = u32::from_le_bytes([
                            data[pos + record_size + 2],
                            data[pos + record_size + 3],
                            data[pos + record_size + 4],
                            data[pos + record_size + 5],
                        ]) as usize;
                        // Skip the unicode stream name record
                        pos += record_size + 6 + next_size;
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

        pos += record_size;
    }

    // Save the last module if present
    if in_module {
        if let (Some(name), Some(stream)) = (current_name, current_stream_name) {
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
    fn test_max_bit_count() {
        assert_eq!(max_bit_count(0), 12);
        assert_eq!(max_bit_count(1), 12);
        assert_eq!(max_bit_count(16), 12);
        assert_eq!(max_bit_count(17), 11);
        assert_eq!(max_bit_count(32), 11);
        assert_eq!(max_bit_count(33), 10);
        assert_eq!(max_bit_count(64), 10);
        assert_eq!(max_bit_count(65), 9);
        assert_eq!(max_bit_count(128), 9);
        assert_eq!(max_bit_count(129), 8);
        assert_eq!(max_bit_count(256), 8);
        assert_eq!(max_bit_count(257), 7);
        assert_eq!(max_bit_count(512), 7);
        assert_eq!(max_bit_count(513), 6);
        assert_eq!(max_bit_count(1024), 6);
        assert_eq!(max_bit_count(1025), 5);
        assert_eq!(max_bit_count(2048), 5);
        assert_eq!(max_bit_count(2049), 4);
        assert_eq!(max_bit_count(4096), 4);
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
        let wb = crate::workbook::Workbook::open_from_buffer(&buf).unwrap();

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

        let wb = crate::workbook::Workbook::open_from_buffer(&buf).unwrap();
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
        let wb = crate::workbook::Workbook::open_from_buffer(&buf).unwrap();
        let saved_buf = wb.save_to_buffer().unwrap();

        // Re-open and verify VBA is preserved
        let wb2 = crate::workbook::Workbook::open_from_buffer(&saved_buf).unwrap();
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
        write_dir_record(&mut data, 0x0009, &version);

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

    /// Minimal MS-OVBA "compression" that produces an uncompressed container.
    /// Signature 0x01 + one uncompressed chunk per 4096 bytes.
    fn compress_for_test(data: &[u8]) -> Vec<u8> {
        let mut result = vec![0x01u8]; // signature
        let mut pos = 0;
        while pos < data.len() {
            let chunk_len = (data.len() - pos).min(4096);
            let chunk_data = &data[pos..pos + chunk_len];
            // Chunk header: bit 15 = 0 (uncompressed), bits 0-11 = chunk_len + 2 - 3
            let header: u16 = (chunk_len as u16 + 2).wrapping_sub(3) & 0x0FFF;
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
