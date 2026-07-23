use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::sync::{Arc, Mutex, MutexGuard};

use crate::error::{Error, Result};
use crate::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};

use super::OpenLimits;

trait ReadSeek: Read + Seek {}
impl<T: Read + Seek> ReadSeek for T {}

type WorksheetMetrics = (u64, u64, u64);
type WorksheetMetricsCache = Arc<Mutex<HashMap<String, WorksheetMetrics>>>;

/// Immutable backing storage for an opened xlsx package.
///
/// File-backed workbooks retain an open file description rather than a path.
/// Renaming or replacing the original path therefore cannot redirect later
/// lazy reads or clean-part passthrough to different bytes.
#[derive(Clone)]
pub(crate) struct PackageSource {
    backing: PackageBacking,
    limits: OpenLimits,
    worksheet_metrics: WorksheetMetricsCache,
}

#[derive(Clone)]
enum PackageBacking {
    File {
        file: Arc<File>,
        operation_lock: Arc<Mutex<()>>,
    },
    Buffer(Arc<[u8]>),
}

impl PackageSource {
    pub(crate) fn from_file(file: File, limits: OpenLimits) -> Self {
        Self {
            backing: PackageBacking::File {
                file: Arc::new(file),
                operation_lock: Arc::new(Mutex::new(())),
            },
            limits,
            worksheet_metrics: Arc::new(Mutex::new(HashMap::new())),
        }
    }

    pub(crate) fn from_buffer(bytes: Arc<[u8]>, limits: OpenLimits) -> Self {
        Self {
            backing: PackageBacking::Buffer(bytes),
            limits,
            worksheet_metrics: Arc::new(Mutex::new(HashMap::new())),
        }
    }

    fn open_archive(&self) -> Result<zip::ZipArchive<Box<dyn ReadSeek>>> {
        let reader: Box<dyn ReadSeek> = match &self.backing {
            PackageBacking::File { file, .. } => {
                let mut clone = file.try_clone()?;
                clone.seek(SeekFrom::Start(0))?;
                Box::new(clone)
            }
            PackageBacking::Buffer(bytes) => Box::new(Cursor::new(Arc::clone(bytes))),
        };
        zip::ZipArchive::new(reader).map_err(|error| Error::Zip(error.to_string()))
    }

    fn lock_file_operation(&self) -> Result<Option<MutexGuard<'_, ()>>> {
        match &self.backing {
            PackageBacking::File { operation_lock, .. } => operation_lock
                .lock()
                .map(Some)
                .map_err(|_| Error::Internal("package source access lock poisoned".to_string())),
            PackageBacking::Buffer(_) => Ok(None),
        }
    }

    /// Decompress one XML part into an anonymous disk spool with a finite
    /// actual-byte boundary. Memory usage remains independent of part size.
    pub(crate) fn spool_xml_part(&self, path: &str) -> Result<File> {
        let _operation = self.lock_file_operation()?;
        let mut archive = self.open_archive()?;
        let mut entry = archive
            .by_name(path)
            .map_err(|error| Error::Zip(error.to_string()))?;
        let limit = self.limits.xml_bytes.min(self.limits.entry_actual_bytes);
        let mut spool = tempfile::tempfile()?;
        let observed = std::io::copy(
            &mut entry.by_ref().take(limit.saturating_add(1)),
            &mut spool,
        )?;
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
        spool.flush()?;
        spool.seek(SeekFrom::Start(0))?;
        Ok(spool)
    }

    /// Copy one entry's compressed payload and metadata without decompressing
    /// it. This is used for clean package-backed worksheets.
    pub(crate) fn raw_copy_part<W: Write + Seek>(
        &self,
        path: &str,
        destination: &mut zip::ZipWriter<W>,
    ) -> Result<()> {
        let _operation = self.lock_file_operation()?;
        let mut archive = self.open_archive()?;
        let entry = archive
            .by_name(path)
            .map_err(|error| Error::Zip(error.to_string()))?;
        destination
            .raw_copy_file(entry)
            .map_err(|error| Error::Zip(error.to_string()))
    }

    /// Inspect an untouched package-backed worksheet under save-time limits.
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
                return Err(super::super::save::resource_limit(
                    ResourceOperation::WorkbookSave,
                    ResourceCategory::Output,
                    ResourceDimension::XmlBytes,
                    xml_limit,
                    metrics.0,
                    ResourceStage::Serialize,
                ));
            }
            return Ok(metrics);
        }

        let _operation = self.lock_file_operation()?;
        let mut archive = self.open_archive()?;
        let entry = archive
            .by_name(path)
            .map_err(|error| Error::Zip(error.to_string()))?;
        let bytes = entry.size();
        if bytes > xml_limit {
            return Err(super::super::save::resource_limit(
                ResourceOperation::WorkbookSave,
                ResourceCategory::Output,
                ResourceDimension::XmlBytes,
                xml_limit,
                bytes,
                ResourceStage::Serialize,
            ));
        }
        let mut reader = quick_xml::Reader::from_reader(std::io::BufReader::new(entry));
        let mut buffer = Vec::with_capacity(4096);
        let mut rows = 0u64;
        let mut cells = 0u64;
        loop {
            buffer.clear();
            match reader
                .read_event_into(&mut buffer)
                .map_err(|error| Error::XmlParse(error.to_string()))?
            {
                quick_xml::events::Event::Start(element)
                | quick_xml::events::Event::Empty(element) => match element.local_name().as_ref() {
                    b"row" => rows = rows.saturating_add(1),
                    b"c" => cells = cells.saturating_add(1),
                    _ => {}
                },
                quick_xml::events::Event::Eof => break,
                _ => {}
            }
        }
        let metrics = (bytes, rows, cells);
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
        let mut reader = quick_xml::Reader::from_reader(std::io::BufReader::new(spool));
        reader.config_mut().trim_text(false);
        let mut output = quick_xml::Writer::new(Cursor::new(Vec::new()));
        let mut buffer = Vec::with_capacity(4096);
        let mut depth = 0usize;

        loop {
            buffer.clear();
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(|error| Error::XmlParse(error.to_string()))?;
            if depth == 0 {
                match &event {
                    quick_xml::events::Event::Start(element)
                        if element.local_name().as_ref() == b"extLst" =>
                    {
                        depth = 1;
                        output
                            .write_event(event.into_owned())
                            .map_err(|error| Error::XmlParse(error.to_string()))?;
                    }
                    quick_xml::events::Event::Empty(element)
                        if element.local_name().as_ref() == b"extLst" =>
                    {
                        output
                            .write_event(event.into_owned())
                            .map_err(|error| Error::XmlParse(error.to_string()))?;
                        break;
                    }
                    quick_xml::events::Event::Eof => break,
                    _ => {}
                }
                continue;
            }

            match &event {
                quick_xml::events::Event::Start(_) => depth += 1,
                quick_xml::events::Event::End(_) => depth -= 1,
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
            if depth == 0 {
                break;
            }
        }

        let fragment = output.into_inner().into_inner();
        Ok((!fragment.is_empty()).then_some(fragment))
    }
}
