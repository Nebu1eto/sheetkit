use std::io::{Read, Seek};

use super::budget::{limit_error, OpenBudget};
use crate::error::{Error, Result};
use crate::resource::{ResourceDimension, ResourceStage};

pub(crate) struct PackageReader<'budget, R: Read + Seek> {
    archive: zip::ZipArchive<R>,
    budget: &'budget mut OpenBudget,
}

impl<'budget, R: Read + Seek> PackageReader<'budget, R> {
    /// Opens a ZIP using an existing operation budget. Encryption and other
    /// container adapters can therefore share counters without resetting them.
    pub(crate) fn open(reader: R, budget: &'budget mut OpenBudget) -> Result<Self> {
        let config = zip::read::Config {
            max_entries: Some(budget.limits().archive_entries),
            ..zip::read::Config::default()
        };
        let archive = zip::ZipArchive::with_config(config, reader).map_err(map_zip_error)?;
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
    actual: u64,
    xml: bool,
}

impl Read for BudgetedEntry<'_> {
    fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
        let count = self.entry.read(buffer)?;
        if count == 0 {
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

    fn len(&self) -> usize;
    fn entry_name(&mut self, index: usize) -> Result<String>;
    fn entry<'entry>(&'entry mut self, name: &str, xml: bool) -> Result<Self::Entry<'entry>>;
    fn entry_optional<'entry>(
        &'entry mut self,
        name: &str,
        xml: bool,
    ) -> Result<Option<Self::Entry<'entry>>>;
    fn resource_violation(&mut self) -> Option<crate::resource::ResourceLimitError>;
    fn charge_retained(&mut self, amount: u64) -> Result<()>;
}

impl<R: Read + Seek> ArchiveAccess for PackageReader<'_, R> {
    type Entry<'entry>
        = BudgetedEntry<'entry>
    where
        Self: 'entry;

    fn len(&self) -> usize {
        self.archive.len()
    }

    fn entry_name(&mut self, index: usize) -> Result<String> {
        self.archive
            .by_index(index)
            .map(|entry| entry.name().to_string())
            .map_err(map_zip_error)
    }

    fn entry<'entry>(&'entry mut self, name: &str, xml: bool) -> Result<Self::Entry<'entry>> {
        if let Some(metadata) = self.budget.violation() {
            return Err(Error::ResourceLimitExceeded(metadata));
        }
        let entry = self.archive.by_name(name).map_err(map_zip_error)?;
        let compressed = entry.compressed_size();
        Ok(BudgetedEntry {
            entry,
            budget: self.budget,
            compressed,
            actual: 0,
            xml,
        })
    }

    fn entry_optional<'entry>(
        &'entry mut self,
        name: &str,
        xml: bool,
    ) -> Result<Option<Self::Entry<'entry>>> {
        if let Some(metadata) = self.budget.violation() {
            return Err(Error::ResourceLimitExceeded(metadata));
        }
        let entry = match self.archive.by_name(name) {
            Ok(entry) => entry,
            Err(zip::result::ZipError::FileNotFound) => return Ok(None),
            Err(error) => return Err(map_zip_error(error)),
        };
        let compressed = entry.compressed_size();
        Ok(Some(BudgetedEntry {
            entry,
            budget: self.budget,
            compressed,
            actual: 0,
            xml,
        }))
    }

    fn resource_violation(&mut self) -> Option<crate::resource::ResourceLimitError> {
        self.budget.violation()
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

    fn len(&self) -> usize {
        zip::ZipArchive::len(self)
    }

    fn entry_name(&mut self, index: usize) -> Result<String> {
        self.by_index(index)
            .map(|entry| entry.name().to_string())
            .map_err(map_zip_error)
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

    fn resource_violation(&mut self) -> Option<crate::resource::ResourceLimitError> {
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
        other => Error::Zip(other.to_string()),
    }
}
