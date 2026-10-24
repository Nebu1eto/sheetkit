use std::collections::HashMap;
use std::fs::File;
use std::io::{BufRead, Cursor, Read, Seek, SeekFrom, Write};
use std::sync::{Arc, Mutex};

use crate::error::{Error, Result};
use crate::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};

use super::package_reader::map_zip_error;
use super::OpenLimits;

type WorksheetMetrics = (u64, u64, u64);
type WorksheetMetricsCache = Arc<Mutex<HashMap<String, WorksheetMetrics>>>;

/// Cloneable, independently positioned reader over immutable package bytes.
///
/// File inputs use positional reads on one shared open file description, so
/// clones never race on a shared file offset.
#[derive(Clone)]
pub(crate) enum PackageInput {
    File { file: Arc<File>, position: u64 },
    Buffer(Cursor<Arc<[u8]>>),
}

impl PackageInput {
    pub(crate) fn file(file: File) -> Self {
        Self::File {
            file: Arc::new(file),
            position: 0,
        }
    }

    pub(crate) fn buffer(bytes: Arc<[u8]>) -> Self {
        Self::Buffer(Cursor::new(bytes))
    }
}

impl Read for PackageInput {
    fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
        match self {
            Self::File { file, position } => {
                let count = read_file_at(file, buffer, *position)?;
                *position = position.saturating_add(count as u64);
                Ok(count)
            }
            Self::Buffer(cursor) => cursor.read(buffer),
        }
    }
}

impl Seek for PackageInput {
    fn seek(&mut self, target: SeekFrom) -> std::io::Result<u64> {
        match self {
            Self::File { file, position } => {
                let next = match target {
                    SeekFrom::Start(offset) => Some(offset),
                    SeekFrom::Current(delta) => position.checked_add_signed(delta),
                    SeekFrom::End(delta) => file.metadata()?.len().checked_add_signed(delta),
                };
                *position = next.ok_or_else(|| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "invalid seek to a negative or overflowing position",
                    )
                })?;
                Ok(*position)
            }
            Self::Buffer(cursor) => cursor.seek(target),
        }
    }
}

#[cfg(unix)]
fn read_file_at(file: &File, buffer: &mut [u8], position: u64) -> std::io::Result<usize> {
    std::os::unix::fs::FileExt::read_at(file, buffer, position)
}

#[cfg(windows)]
fn read_file_at(file: &File, buffer: &mut [u8], position: u64) -> std::io::Result<usize> {
    std::os::windows::fs::FileExt::seek_read(file, buffer, position)
}

#[cfg(not(any(unix, windows)))]
fn read_file_at(file: &File, buffer: &mut [u8], position: u64) -> std::io::Result<usize> {
    let mut file = file;
    file.seek(SeekFrom::Start(position))?;
    file.read(buffer)
}

/// Parse a package central directory under the workbook's entry limit.
pub(crate) fn open_package_archive(
    input: PackageInput,
    limits: OpenLimits,
) -> Result<zip::ZipArchive<PackageInput>> {
    let config = zip::read::Config {
        max_entries: Some(limits.archive_entries),
        ..zip::read::Config::default()
    };
    zip::ZipArchive::with_config(config, input).map_err(map_zip_error)
}

/// Immutable backing storage for an opened xlsx package.
///
/// The central directory is parsed once, under the open-time limits, and the
/// same index serves every later lazy read and clean-part passthrough.
/// File-backed workbooks retain an open file description rather than a path,
/// so renaming or replacing the original path cannot redirect those reads,
/// and an in-place rewrite cannot introduce entries the index never listed.
#[derive(Clone)]
pub(crate) struct PackageSource {
    archive: zip::ZipArchive<PackageInput>,
    limits: OpenLimits,
    worksheet_metrics: WorksheetMetricsCache,
}

impl PackageSource {
    pub(crate) fn new(archive: zip::ZipArchive<PackageInput>, limits: OpenLimits) -> Self {
        Self {
            archive,
            limits,
            worksheet_metrics: Arc::new(Mutex::new(HashMap::new())),
        }
    }

    /// Decompress one XML part into an anonymous disk spool with a finite
    /// actual-byte boundary. Memory usage remains independent of part size.
    pub(crate) fn spool_xml_part(&self, path: &str) -> Result<File> {
        let mut archive = self.archive.clone();
        let mut entry = archive.by_name(path).map_err(map_zip_error)?;
        let declared = entry.size();
        let limit = self.limits.xml_bytes.min(self.limits.entry_actual_bytes);
        let mut spool = tempfile::tempfile()?;
        let observed = std::io::copy(
            &mut entry.by_ref().take(limit.saturating_add(1)),
            &mut spool,
        )
        .map_err(|error| Error::Zip(error.to_string()))?;
        check_xml_part_bytes(limit, observed)?;
        check_declared_size(path, declared, observed)?;
        spool.flush()?;
        spool.seek(SeekFrom::Start(0))?;
        Ok(spool)
    }

    /// Parse one XML part straight from the decompressing entry.
    ///
    /// Results match parsing the [`Self::spool_xml_part`] output: the entry
    /// is still read to its end under the same byte limit, and an archive
    /// read error, then the byte limit, then the declared size take
    /// precedence over the parse result. Nothing is spooled to disk.
    pub(crate) fn parse_xml_part<T>(
        &self,
        path: &str,
        parse: impl FnOnce(&mut std::io::BufReader<&mut dyn Read>) -> Result<T>,
    ) -> Result<T> {
        let mut archive = self.archive.clone();
        let mut entry = archive.by_name(path).map_err(map_zip_error)?;
        let declared = entry.size();
        let limit = self.limits.xml_bytes.min(self.limits.entry_actual_bytes);
        let mut source = CountingReader {
            inner: entry.by_ref().take(limit.saturating_add(1)),
            count: 0,
            error: None,
        };
        let parsed = {
            // Only the refill read is dynamic; the parser's per-event buffer
            // calls stay on a concrete `BufReader`.
            let entry_reader: &mut dyn Read = &mut source;
            let mut reader = std::io::BufReader::with_capacity(PARSE_BUF_CAPACITY, entry_reader);
            let parsed = parse(&mut reader);
            // Errors are recorded and sticky, so this never reads past one.
            let _ = std::io::copy(&mut reader, &mut std::io::sink());
            parsed
        };
        if let Some(message) = source.error {
            return Err(Error::Zip(message));
        }
        check_xml_part_bytes(limit, source.count)?;
        check_declared_size(path, declared, source.count)?;
        parsed
    }

    /// Copy one entry's compressed payload and metadata without decompressing
    /// it. Only entries whose actual size and checksum were verified by
    /// [`Self::worksheet_save_metrics`] may be copied.
    pub(crate) fn raw_copy_part<W: Write + Seek>(
        &self,
        path: &str,
        destination: &mut zip::ZipWriter<W>,
    ) -> Result<()> {
        if !self
            .worksheet_metrics
            .lock()
            .map_err(|_| Error::Internal("worksheet metrics cache lock poisoned".to_string()))?
            .contains_key(path)
        {
            return Err(Error::Internal(format!(
                "package part was not verified before passthrough: {path}"
            )));
        }
        let mut archive = self.archive.clone();
        let entry = archive.by_name(path).map_err(map_zip_error)?;
        destination
            .raw_copy_file(entry)
            .map_err(|error| Error::Zip(error.to_string()))
    }

    /// Inspect an untouched package-backed worksheet under save-time limits.
    ///
    /// The XML limit applies to bytes actually decompressed, never to the
    /// size declared by the archive. A successful result also proves that the
    /// declared size and checksum match the entry contents.
    pub(crate) fn worksheet_save_metrics(
        &self,
        path: &str,
        xml_limit: u64,
    ) -> Result<(u64, u64, u64)> {
        if let Some(metrics) = self
            .worksheet_metrics
            .lock()
            .map_err(|_| Error::Internal("worksheet metrics cache lock poisoned".to_string()))?
            .get(path)
            .copied()
        {
            if metrics.0 > xml_limit {
                return Err(save_xml_limit(xml_limit, metrics.0));
            }
            return Ok(metrics);
        }

        // Fast exact pass. A complete, in-limit read whose markup stays in
        // the mirrored subset yields the quick-xml counts; every other case
        // reruns the original pass below, which owns all error reporting.
        {
            let mut archive = self.archive.clone();
            let mut entry = archive.by_name(path).map_err(map_zip_error)?;
            let declared = entry.size();
            let mut counted = CountingReader {
                inner: entry.by_ref().take(xml_limit.saturating_add(1)),
                count: 0,
                error: None,
            };
            if let Some((rows, cells)) = fast_count_rows_and_cells(&mut counted) {
                if counted.count <= xml_limit {
                    check_declared_size(path, declared, counted.count)?;
                    let metrics = (counted.count, rows, cells);
                    self.worksheet_metrics
                        .lock()
                        .map_err(|_| {
                            Error::Internal("worksheet metrics cache lock poisoned".to_string())
                        })?
                        .insert(path.to_string(), metrics);
                    return Ok(metrics);
                }
            }
        }

        let mut archive = self.archive.clone();
        let mut entry = archive.by_name(path).map_err(map_zip_error)?;
        let declared = entry.size();
        let mut counted = CountingReader {
            inner: entry.by_ref().take(xml_limit.saturating_add(1)),
            count: 0,
            error: None,
        };
        let parsed = count_rows_and_cells(std::io::BufReader::new(&mut counted));
        if counted.count > xml_limit {
            return Err(save_xml_limit(xml_limit, counted.count));
        }
        let (rows, cells) = parsed?;
        std::io::copy(&mut counted, &mut std::io::sink())
            .map_err(|error| Error::Zip(error.to_string()))?;
        let observed = counted.count;
        if observed > xml_limit {
            return Err(save_xml_limit(xml_limit, observed));
        }
        check_declared_size(path, declared, observed)?;

        let metrics = (observed, rows, cells);
        self.worksheet_metrics
            .lock()
            .map_err(|_| Error::Internal("worksheet metrics cache lock poisoned".to_string()))?
            .insert(path.to_string(), metrics);
        Ok(metrics)
    }

    /// Read only the worksheet-level extension list from a package-backed
    /// sheet. This preserves opaque extensions during typed mutations without
    /// allocating a second copy of the complete worksheet.
    pub(crate) fn worksheet_extension_fragment(&self, path: &str) -> Result<Option<Vec<u8>>> {
        let spool = self.spool_xml_part(path)?;
        extract_worksheet_extension_fragment(std::io::BufReader::new(spool))
    }
}

/// Read buffer for streaming XML parts; larger than the 8 KiB default to
/// cut per-read overhead through the decompressor.
const PARSE_BUF_CAPACITY: usize = 64 * 1024;

/// Counts bytes read and records the first read error's message. Once an
/// error is recorded, later reads fail without touching the inner reader.
struct CountingReader<R> {
    inner: R,
    count: u64,
    error: Option<String>,
}

impl<R: Read> Read for CountingReader<R> {
    fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
        if let Some(message) = &self.error {
            return Err(std::io::Error::other(message.clone()));
        }
        let count = match self.inner.read(buffer) {
            Ok(count) => count,
            Err(error) => {
                self.error = Some(error.to_string());
                return Err(error);
            }
        };
        self.count = self.count.saturating_add(count as u64);
        Ok(count)
    }
}

fn check_xml_part_bytes(limit: u64, observed: u64) -> Result<()> {
    if observed > limit {
        let metadata = ResourceLimitError::exceeded(
            ResourceOperation::StreamRead,
            ResourceCategory::Streaming,
            ResourceDimension::XmlBytes,
            limit,
            observed,
            Some(ResourceStage::Parse),
        )
        .ok_or_else(|| Error::Internal("invalid stream XML limit state".to_string()))?;
        return Err(Error::ResourceLimitExceeded(metadata));
    }
    Ok(())
}

fn count_rows_and_cells<R: BufRead>(reader: R) -> Result<(u64, u64)> {
    let mut reader = quick_xml::Reader::from_reader(reader);
    let mut buffer = Vec::with_capacity(4096);
    let mut rows = 0u64;
    let mut cells = 0u64;
    loop {
        buffer.clear();
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| match error {
                quick_xml::Error::Io(error) => Error::Zip(error.to_string()),
                other => Error::XmlParse(other.to_string()),
            })? {
            quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element) => {
                match element.local_name().as_ref() {
                    b"row" => rows = rows.saturating_add(1),
                    b"c" => cells = cells.saturating_add(1),
                    _ => {}
                }
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    Ok((rows, cells))
}

/// Markup found at a `<` by [`fast_count_rows_and_cells`].
enum Markup {
    /// More input is needed to finish the construct.
    Incomplete,
    /// The construct is outside the mirrored subset; use quick-xml.
    Unsupported,
    /// A processing instruction or declaration ending at the index.
    Instruction(usize),
    /// A tag whose `>` is at the index.
    Tag(usize),
}

/// Find the end of the markup starting at `data[lt] == b'<'`, mirroring
/// quick-xml 0.36: `<?` ends at the first `?>`, other tags at the first
/// `>` outside single or double quotes.
fn find_markup_end(data: &[u8], lt: usize) -> Markup {
    let Some(&kind) = data.get(lt + 1) else {
        return Markup::Incomplete;
    };
    match kind {
        b'!' => Markup::Unsupported,
        b'?' => {
            for end in lt + 2..data.len() {
                if data[end] == b'>' && data[end - 1] == b'?' {
                    // `<?>` is an unclosed declaration in quick-xml.
                    return if end == lt + 2 {
                        Markup::Unsupported
                    } else {
                        Markup::Instruction(end)
                    };
                }
            }
            Markup::Incomplete
        }
        _ => {
            let mut quote = 0u8;
            for (end, &byte) in data.iter().enumerate().skip(lt + 1) {
                match (quote, byte) {
                    (0, b'>') => return Markup::Tag(end),
                    (0, b'\'' | b'"') => quote = byte,
                    (open, close) if open == close => quote = 0,
                    _ => {}
                }
            }
            Markup::Incomplete
        }
    }
}

/// Count `row` and `c` start and empty elements with the result
/// [`count_rows_and_cells`] would return, without per-event copies.
///
/// Only text, tags, and `<?...?>` instructions are mirrored. `None` means
/// the input left that subset, was ill-formed, ended inside markup, or
/// failed to read; the caller must then rerun the quick-xml pass, which
/// reports any error exactly as before.
fn fast_count_rows_and_cells<R: Read>(mut reader: R) -> Option<(u64, u64)> {
    const CHUNK: usize = 256 * 1024;
    let is_whitespace = |byte: u8| matches!(byte, b' ' | b'\r' | b'\n' | b'\t');
    let mut buffer = vec![0u8; CHUNK];
    let mut filled = 0usize;
    let mut opened: Vec<u8> = Vec::new();
    let mut opened_starts: Vec<usize> = Vec::new();
    let mut rows = 0u64;
    let mut cells = 0u64;
    loop {
        if filled == buffer.len() {
            // One unfinished construct fills the buffer; quick-xml would
            // also buffer it whole.
            buffer.resize(buffer.len().saturating_mul(2), 0);
        }
        let read = loop {
            match reader.read(&mut buffer[filled..]) {
                Ok(read) => break read,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => {}
                Err(_) => return None,
            }
        };
        filled += read;
        let at_eof = read == 0;
        let data = &buffer[..filled];

        let mut position = 0usize;
        while let Some(offset) = data[position..].iter().position(|&byte| byte == b'<') {
            let lt = position + offset;
            let end = match find_markup_end(data, lt) {
                Markup::Incomplete => break,
                Markup::Unsupported => return None,
                Markup::Instruction(end) => {
                    position = end + 1;
                    continue;
                }
                Markup::Tag(end) => end,
            };
            position = end + 1;
            let content = &data[lt + 1..end];
            if let Some(closing) = content.strip_prefix(b"/") {
                let name = match closing.iter().rposition(|&byte| !is_whitespace(byte)) {
                    Some(last) => &closing[..=last],
                    None => closing,
                };
                let start = opened_starts.pop()?;
                if opened[start..] != *name {
                    return None;
                }
                opened.truncate(start);
                continue;
            }
            let (element, is_empty) = match content.strip_suffix(b"/") {
                Some(inner) => (inner, true),
                None => (content, false),
            };
            let name_end = element
                .iter()
                .position(|&byte| is_whitespace(byte))
                .unwrap_or(element.len());
            let name = &element[..name_end];
            let local = match name.iter().position(|&byte| byte == b':') {
                Some(colon) => &name[colon + 1..],
                None => name,
            };
            match local {
                b"row" => rows = rows.saturating_add(1),
                b"c" => cells = cells.saturating_add(1),
                _ => {}
            }
            if !is_empty {
                opened_starts.push(opened.len());
                opened.extend_from_slice(name);
            }
        }
        if at_eof {
            // quick-xml reports `<` without a closing `>` as a syntax error.
            return (!data[position..].contains(&b'<')).then_some((rows, cells));
        }
        // Text carries no state; markup that is still open is retried from
        // its `<` with more input.
        let keep_from = data[position..]
            .iter()
            .position(|&byte| byte == b'<')
            .map_or(filled, |offset| position + offset);
        buffer.copy_within(keep_from..filled, 0);
        filled -= keep_from;
    }
}

fn save_xml_limit(limit: u64, observed: u64) -> Error {
    super::super::save::resource_limit(
        ResourceOperation::WorkbookSave,
        ResourceCategory::Output,
        ResourceDimension::XmlBytes,
        limit,
        observed,
        ResourceStage::Serialize,
    )
}

fn check_declared_size(path: &str, declared: u64, observed: u64) -> Result<()> {
    if declared != observed {
        return Err(Error::Zip(format!(
            "ZIP entry size mismatch for {path}: declared {declared} bytes, read {observed} bytes"
        )));
    }
    Ok(())
}

/// Extract the first `extLst` that is a direct child of the worksheet root.
///
/// Nested extension lists, such as those inside conditional-formatting
/// rules, are not worksheet-level extensions and are ignored.
pub(crate) fn extract_worksheet_extension_fragment<R: BufRead>(
    source: R,
) -> Result<Option<Vec<u8>>> {
    let mut reader = quick_xml::Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut output = quick_xml::Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::with_capacity(4096);
    let mut document_depth = 0usize;
    let mut capture_depth = 0usize;

    loop {
        buffer.clear();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::XmlParse(error.to_string()))?;
        if capture_depth == 0 {
            match &event {
                quick_xml::events::Event::Start(element) => {
                    if document_depth == 1 && element.local_name().as_ref() == b"extLst" {
                        capture_depth = 1;
                        output
                            .write_event(event.into_owned())
                            .map_err(|error| Error::XmlParse(error.to_string()))?;
                    }
                    document_depth += 1;
                }
                quick_xml::events::Event::Empty(element)
                    if document_depth == 1 && element.local_name().as_ref() == b"extLst" =>
                {
                    output
                        .write_event(event.into_owned())
                        .map_err(|error| Error::XmlParse(error.to_string()))?;
                    break;
                }
                quick_xml::events::Event::End(_) => {
                    document_depth = document_depth.saturating_sub(1);
                }
                quick_xml::events::Event::Eof => break,
                _ => {}
            }
            continue;
        }

        match &event {
            quick_xml::events::Event::Start(_) => capture_depth += 1,
            quick_xml::events::Event::End(_) => capture_depth -= 1,
            quick_xml::events::Event::Eof => {
                return Err(Error::XmlParse(
                    "unexpected EOF in worksheet extension list".to_string(),
                ));
            }
            _ => {}
        }
        output
            .write_event(event.into_owned())
            .map_err(|error| Error::XmlParse(error.to_string()))?;
        if capture_depth == 0 {
            break;
        }
    }

    let fragment = output.into_inner().into_inner();
    Ok((!fragment.is_empty()).then_some(fragment))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extension_fragment_ignores_nested_extension_lists() {
        let xml = br#"<worksheet><conditionalFormatting><cfRule><extLst><ext uri="nested"/></extLst></cfRule></conditionalFormatting><extLst><ext uri="root"/></extLst></worksheet>"#;
        let fragment = extract_worksheet_extension_fragment(&xml[..])
            .unwrap()
            .unwrap();
        let fragment = String::from_utf8(fragment).unwrap();
        assert!(fragment.contains("root"), "{fragment}");
        assert!(!fragment.contains("nested"), "{fragment}");
    }

    fn package_with_part(data: &[u8], corrupt_crc: bool) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    "part.xml",
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .unwrap();
            writer.write_all(data).unwrap();
            writer.finish().unwrap();
        }
        if corrupt_crc {
            // Flip the CRC in the local header and the central directory.
            for (signature, offset) in [(b"PK\x03\x04", 14usize), (b"PK\x01\x02", 16usize)] {
                let start = bytes
                    .windows(4)
                    .position(|window| window == signature)
                    .unwrap();
                bytes[start + offset] ^= 0xff;
            }
        }
        bytes
    }

    fn source_for(bytes: Vec<u8>, xml_bytes: u64) -> PackageSource {
        let limits = OpenLimits {
            xml_bytes,
            ..OpenLimits::default()
        };
        let archive = open_package_archive(PackageInput::buffer(bytes.into()), limits).unwrap();
        PackageSource::new(archive, limits)
    }

    /// Parse as the streaming path does, and as the former spool path did.
    fn both_paths(source: &PackageSource, stop_early: bool) -> (String, String) {
        let parse = |reader: &mut std::io::BufReader<&mut dyn Read>| -> Result<String> {
            let mut text = String::new();
            if stop_early {
                let mut first = [0u8; 1];
                reader.read_exact(&mut first)?;
                text.push(char::from(first[0]));
            } else {
                reader.read_to_string(&mut text)?;
            }
            if text.contains("bad") {
                return Err(Error::XmlParse("bad".to_string()));
            }
            Ok(text)
        };
        let streamed = source.parse_xml_part("part.xml", parse);
        let spooled = source.spool_xml_part("part.xml").and_then(|mut spool| {
            let spool: &mut dyn Read = &mut spool;
            parse(&mut std::io::BufReader::new(spool))
        });
        let show = |result: Result<String>| match result {
            Ok(text) => format!("ok:{text}"),
            Err(error) => format!("err:{error:?}"),
        };
        (show(streamed), show(spooled))
    }

    #[test]
    fn streamed_part_parse_matches_spool_precedence() {
        let cases: [(&[u8], bool, u64, bool); 7] = [
            (b"<a>ok</a>", false, 64, false),
            (b"<a>bad</a>", false, 64, false),
            (b"<a>ok</a>", false, 4, false),
            (b"<a>bad</a>", false, 4, true),
            (b"<a>ok</a>", true, 64, false),
            (b"<a>bad</a>", true, 64, false),
            (b"<a>ok</a>", true, 64, true),
        ];
        for (data, corrupt, limit, stop_early) in cases {
            let source = source_for(package_with_part(data, corrupt), limit);
            let (streamed, spooled) = both_paths(&source, stop_early);
            assert_eq!(streamed, spooled, "corrupt={corrupt} limit={limit}");
        }
        let source = source_for(package_with_part(b"<a>ok</a>", true), 64);
        assert!(both_paths(&source, true).0.starts_with("err:Zip"));
        let source = source_for(package_with_part(b"<a>ok</a>", false), 4);
        assert!(both_paths(&source, true)
            .0
            .starts_with("err:ResourceLimitExceeded"));
    }

    /// Returns at most `sizes[i % len]` bytes per read to split markup.
    struct ChunkedReader<'data> {
        data: &'data [u8],
        sizes: &'data [usize],
        reads: usize,
    }

    impl Read for ChunkedReader<'_> {
        fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
            let size = self.sizes[self.reads % self.sizes.len()]
                .min(buffer.len())
                .min(self.data.len());
            self.reads += 1;
            buffer[..size].copy_from_slice(&self.data[..size]);
            self.data = &self.data[size..];
            Ok(size)
        }
    }

    fn fast_count_in_chunks(data: &[u8], sizes: &[usize]) -> Option<(u64, u64)> {
        fast_count_rows_and_cells(ChunkedReader {
            data,
            sizes,
            reads: 0,
        })
    }

    #[test]
    fn fast_row_and_cell_count_matches_quick_xml_on_edge_cases() {
        let accepted: [(&[u8], (u64, u64)); 10] = [
            (
                b"\xEF\xBB\xBF<?xml version=\"1.0\"?><worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"/></row></sheetData></worksheet>",
                (1, 2),
            ),
            (b"<x:row><x:c/></x:row>", (1, 1)),
            (b"<a:b:c/><c/ ><c\t/>", (0, 1)),
            (b"<c a='>' b=\"'>\">t</c >", (0, 1)),
            (b"<row></row\n>", (1, 0)),
            (b"<row><c>", (1, 1)),
            (b"<>< ></>", (0, 0)),
            (b"<??><?pi a>b?>", (0, 0)),
            (b"text > only", (0, 0)),
            (b"", (0, 0)),
        ];
        for (data, expected) in accepted {
            let label = String::from_utf8_lossy(data);
            assert_eq!(count_rows_and_cells(data).unwrap(), expected, "{label}");
            for sizes in [&[1usize][..], &[3, 1, 7], &[1 << 20]] {
                assert_eq!(fast_count_in_chunks(data, sizes), Some(expected), "{label}");
            }
        }

        // Outside the mirrored subset or ill-formed: the caller falls back.
        let fallback: [&[u8]; 10] = [
            b"<c><!-- <c/> --></c>",
            b"<c><![CDATA[<c/>]]></c>",
            b"<!DOCTYPE x><c/>",
            b"<?>",
            b"<row></c>",
            b"<row></row x='>'>",
            b"</row>",
            b"<c/><",
            b"<c a='>",
            b"<?pi",
        ];
        for data in fallback {
            assert_eq!(fast_count_in_chunks(data, &[2, 5]), None, "{data:?}");
        }
    }

    #[test]
    fn fast_row_and_cell_count_agrees_with_quick_xml_on_random_markup() {
        const TOKENS: [&[u8]; 24] = [
            b"<",
            b">",
            b"/",
            b"?",
            b"!",
            b"\"",
            b"'",
            b" ",
            b"\t",
            b"c",
            b"row",
            b"x:",
            b":",
            b"=",
            b"<c>",
            b"</c>",
            b"<row r=\"1\">",
            b"</row>",
            b"<c/>",
            b"<x:c/>",
            b"<?xml v?>",
            b"text",
            b"</c ",
            b"<c a='>'>",
        ];
        let mut state = 0x9e37_79b9u32;
        let mut next = |bound: usize| {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            state as usize % bound
        };
        let mut accepted = 0;
        for _ in 0..20_000 {
            let length = next(24);
            let mut data = Vec::new();
            for _ in 0..length {
                data.extend_from_slice(TOKENS[next(TOKENS.len())]);
            }
            let sizes = [next(9) + 1, next(5) + 1, next(64) + 1];
            if let Some(counts) = fast_count_in_chunks(&data, &sizes) {
                accepted += 1;
                assert_eq!(
                    count_rows_and_cells(&data[..]).ok(),
                    Some(counts),
                    "{:?}",
                    String::from_utf8_lossy(&data)
                );
            }
        }
        assert!(
            accepted > 1_000,
            "only {accepted} inputs exercised the fast path"
        );
    }

    #[test]
    fn positional_file_clones_keep_independent_offsets() {
        let mut file = tempfile::tempfile().unwrap();
        file.write_all(b"abcdef").unwrap();
        let mut first = PackageInput::file(file);
        let mut second = first.clone();
        let mut buffer = [0u8; 3];
        first.read_exact(&mut buffer).unwrap();
        assert_eq!(&buffer, b"abc");
        second.seek(SeekFrom::End(-2)).unwrap();
        let mut tail = [0u8; 2];
        second.read_exact(&mut tail).unwrap();
        assert_eq!(&tail, b"ef");
        first.read_exact(&mut buffer).unwrap();
        assert_eq!(&buffer, b"def");
    }
}
