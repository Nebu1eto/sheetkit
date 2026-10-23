use std::io::{Read, Seek};

use super::budget::{limit_error, OpenBudget};
use crate::error::{Error, Result};
use crate::resource::{ResourceDimension, ResourceStage};

pub(crate) struct PackageReader<'budget, R: Read + Seek> {
    archive: zip::ZipArchive<R>,
    budget: &'budget mut OpenBudget,
}

impl<'budget, R: Read + Seek> PackageReader<'budget, R> {
    /// Wraps a parsed archive with an existing operation budget and preflights
    /// every declared entry size. Encryption and other container adapters can
    /// therefore share counters without resetting them.
    pub(crate) fn from_archive(
        archive: zip::ZipArchive<R>,
        budget: &'budget mut OpenBudget,
    ) -> Result<Self> {
        let mut package = Self { archive, budget };
        package.preflight()?;
        Ok(package)
    }

    fn preflight(&mut self) -> Result<()> {
        for index in 0..self.archive.len() {
            let entry = self.archive.by_index(index).map_err(map_zip_error)?;
            self.budget
                .preflight_entry(entry.size(), entry.compressed_size())?;
        }
        Ok(())
    }
}

pub(crate) struct BudgetedEntry<'archive> {
    entry: zip::read::ZipFile<'archive>,
    budget: &'archive mut OpenBudget,
    compressed: u64,
    declared: u64,
    actual: u64,
    xml: bool,
}

impl<'archive> BudgetedEntry<'archive> {
    fn new(
        entry: zip::read::ZipFile<'archive>,
        budget: &'archive mut OpenBudget,
        xml: bool,
    ) -> Self {
        let compressed = entry.compressed_size();
        let declared = entry.size();
        Self {
            entry,
            budget,
            compressed,
            declared,
            actual: 0,
            xml,
        }
    }
}

impl Read for BudgetedEntry<'_> {
    fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
        let count = self.entry.read(buffer)?;
        if count == 0 {
            // A central directory size that disagrees with the decompressed
            // stream would let callers trust an attacker-chosen size.
            if !buffer.is_empty() && self.actual != self.declared {
                self.budget.record_integrity_violation(format!(
                    "ZIP entry size mismatch for {}: declared {} bytes, read {} bytes",
                    self.entry.name(),
                    self.declared,
                    self.actual
                ));
                return Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    "ZIP entry size mismatch",
                ));
            }
            return Ok(0);
        }
        match self
            .budget
            .charge_actual(self.actual, count as u64, self.compressed, self.xml)
        {
            Ok(total) => {
                self.actual = total;
                Ok(count)
            }
            Err(error) => {
                self.budget.record_violation(error);
                Err(std::io::Error::other("workbook resource limit exceeded"))
            }
        }
    }
}

pub(crate) trait ArchiveAccess {
    type Entry<'entry>: Read
    where
        Self: 'entry;

    /// Entry names in central-directory order, read without entry I/O.
    fn entry_names(&self) -> Vec<String>;
    fn entry<'entry>(&'entry mut self, name: &str, xml: bool) -> Result<Self::Entry<'entry>>;
    fn entry_optional<'entry>(
        &'entry mut self,
        name: &str,
        xml: bool,
    ) -> Result<Option<Self::Entry<'entry>>>;
    /// First resource-limit or integrity failure recorded by an entry reader.
    fn pending_error(&mut self) -> Option<Error>;
    fn charge_retained(&mut self, amount: u64) -> Result<()>;
}

impl<R: Read + Seek> ArchiveAccess for PackageReader<'_, R> {
    type Entry<'entry>
        = BudgetedEntry<'entry>
    where
        Self: 'entry;

    fn entry_names(&self) -> Vec<String> {
        self.archive.file_names().map(str::to_string).collect()
    }

    fn entry<'entry>(&'entry mut self, name: &str, xml: bool) -> Result<Self::Entry<'entry>> {
        if let Some(error) = self.budget.pending_error() {
            return Err(error);
        }
        let entry = self.archive.by_name(name).map_err(map_zip_error)?;
        Ok(BudgetedEntry::new(entry, self.budget, xml))
    }

    fn entry_optional<'entry>(
        &'entry mut self,
        name: &str,
        xml: bool,
    ) -> Result<Option<Self::Entry<'entry>>> {
        if let Some(error) = self.budget.pending_error() {
            return Err(error);
        }
        let entry = match self.archive.by_name(name) {
            Ok(entry) => entry,
            Err(zip::result::ZipError::FileNotFound) => return Ok(None),
            Err(error) => return Err(map_zip_error(error)),
        };
        Ok(Some(BudgetedEntry::new(entry, self.budget, xml)))
    }

    fn pending_error(&mut self) -> Option<Error> {
        self.budget.pending_error()
    }

    fn charge_retained(&mut self, amount: u64) -> Result<()> {
        match self.budget.charge_retained(amount) {
            Ok(()) => Ok(()),
            Err(error) => {
                if let Some(metadata) = error.resource_metadata().copied() {
                    self.budget
                        .record_violation(Error::ResourceLimitExceeded(metadata));
                }
                Err(error)
            }
        }
    }
}

impl<R: Read + Seek> ArchiveAccess for zip::ZipArchive<R> {
    type Entry<'entry>
        = zip::read::ZipFile<'entry>
    where
        Self: 'entry;

    fn entry_names(&self) -> Vec<String> {
        self.file_names().map(str::to_string).collect()
    }

    fn entry<'entry>(&'entry mut self, name: &str, _xml: bool) -> Result<Self::Entry<'entry>> {
        self.by_name(name).map_err(map_zip_error)
    }

    fn entry_optional<'entry>(
        &'entry mut self,
        name: &str,
        _xml: bool,
    ) -> Result<Option<Self::Entry<'entry>>> {
        match self.by_name(name) {
            Ok(entry) => Ok(Some(entry)),
            Err(zip::result::ZipError::FileNotFound) => Ok(None),
            Err(error) => Err(map_zip_error(error)),
        }
    }

    fn pending_error(&mut self) -> Option<Error> {
        None
    }

    fn charge_retained(&mut self, _amount: u64) -> Result<()> {
        Ok(())
    }
}

pub(crate) fn map_zip_error(error: zip::result::ZipError) -> Error {
    match error {
        zip::result::ZipError::EntryCountLimitExceeded { count, limit } => limit_error(
            ResourceDimension::EntryCount,
            limit,
            count,
            ResourceStage::Inspect,
        ),
        zip::result::ZipError::DuplicateEntryName => {
            Error::Zip("ZIP archive contains duplicate entry names".to_string())
        }
        other => Error::Zip(other.to_string()),
    }
}
