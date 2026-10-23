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
        check_declared_size(path, declared, observed)?;
        spool.flush()?;
        spool.seek(SeekFrom::Start(0))?;
        Ok(spool)
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

        let mut archive = self.archive.clone();
        let mut entry = archive.by_name(path).map_err(map_zip_error)?;
        let declared = entry.size();
        let mut counted = CountingReader {
            inner: entry.by_ref().take(xml_limit.saturating_add(1)),
            count: 0,
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

struct CountingReader<R> {
    inner: R,
    count: u64,
}

impl<R: Read> Read for CountingReader<R> {
    fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
        let count = self.inner.read(buffer)?;
        self.count = self.count.saturating_add(count as u64);
        Ok(count)
    }
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
