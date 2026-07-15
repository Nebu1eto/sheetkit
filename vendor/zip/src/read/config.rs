/// Configuration for reading ZIP archives.
#[derive(Debug, Clone, Copy)]
pub struct Config {
    /// An offset into the reader to use to find the start of the archive.
    pub archive_offset: ArchiveOffset,

    /// Maximum number of entries accepted from the central directory.
    ///
    /// `None` disables the limit.
    pub max_entries: Option<u64>,

    /// Maximum number of bytes retained from the ZIP64 extensible data sector.
    ///
    /// `None` is rejected. Values above [`MAX_ZIP64_EXTENSIBLE_DATA_SIZE`]
    /// are also rejected.
    pub max_zip64_extensible_data_size: Option<u64>,
}

/// Non-overridable maximum retained ZIP64 extensible data sector size.
pub const MAX_ZIP64_EXTENSIBLE_DATA_SIZE: u64 = 1024 * 1024;

impl Default for Config {
    fn default() -> Self {
        Self {
            archive_offset: ArchiveOffset::default(),
            max_entries: None,
            max_zip64_extensible_data_size: Some(MAX_ZIP64_EXTENSIBLE_DATA_SIZE),
        }
    }
}

/// The offset of the start of the archive from the beginning of the reader.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ArchiveOffset {
    /// Try to detect the archive offset automatically.
    ///
    /// This will look at the central directory specified by `FromCentralDirectory` for a header.
    /// If missing, this will behave as if `None` were specified.
    #[default]
    Detect,
    /// Use the central directory length and offset to determine the start of the archive.
    #[deprecated(since = "2.3.0", note = "use `Detect` instead")]
    FromCentralDirectory,
    /// Specify a fixed archive offset.
    Known(u64),
}
