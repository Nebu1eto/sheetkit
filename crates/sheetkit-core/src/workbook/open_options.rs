/// Controls how worksheets and auxiliary parts are parsed during open.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ReadMode {
    /// Parse all parts eagerly. Equivalent to the old `Full` mode.
    #[default]
    Eager,
    /// Skip auxiliary parts (comments, drawings, charts, images, doc props,
    /// pivot tables, slicers, threaded comments, VBA, tables, form controls).
    /// These are stored as raw bytes for on-demand parsing or direct
    /// round-trip preservation. Equivalent to the old `ReadFast` mode.
    /// Will evolve into true lazy on-demand hydration in later workstreams.
    Lazy,
    /// Forward-only streaming read mode (reserved for future use).
    /// Currently behaves the same as `Lazy`.
    Stream,
}

/// Controls when auxiliary parts (comments, charts, images, etc.) are parsed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AuxParts {
    /// Parse auxiliary parts only when accessed.
    Deferred,
    /// Parse all auxiliary parts during open (default).
    #[default]
    EagerLoad,
}

/// Options for controlling how a workbook is opened and parsed.
///
/// All fields default to `None` (no limit / parse everything).
/// Use the builder-style setter methods for convenience.
#[derive(Debug, Clone, Default)]
pub struct OpenOptions {
    /// Maximum number of rows to read per sheet. Rows beyond this limit
    /// are silently discarded during parsing.
    pub sheet_rows: Option<u32>,

    /// Only parse sheets whose names are in this list. Sheets not listed
    /// are represented as empty worksheets (their XML is not parsed).
    /// `None` means parse all sheets.
    pub sheets: Option<Vec<String>>,

    /// Maximum total decompressed size of all ZIP entries in bytes.
    /// Exceeding this limit returns [`Error::ZipSizeExceeded`].
    /// Default when `None`: no limit.
    pub max_unzip_size: Option<u64>,

    /// Maximum number of ZIP entries allowed.
    /// Exceeding this limit returns [`Error::ZipEntryCountExceeded`].
    /// Default when `None`: no limit.
    pub max_zip_entries: Option<usize>,

    /// Read mode: `Eager` (default) parses everything; `Lazy` skips
    /// auxiliary parts for faster read-only workloads; `Stream` is
    /// reserved for future streaming reads.
    pub read_mode: ReadMode,

    /// Controls when auxiliary parts are parsed.
    pub aux_parts: AuxParts,
}

impl OpenOptions {
    /// Create a new `OpenOptions` with all defaults (no limits, parse everything).
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the maximum number of rows to read per sheet.
    pub fn sheet_rows(mut self, rows: u32) -> Self {
        self.sheet_rows = Some(rows);
        self
    }

    /// Only parse sheets whose names are in this list.
    pub fn sheets(mut self, names: Vec<String>) -> Self {
        self.sheets = Some(names);
        self
    }

    /// Set the maximum total decompressed size in bytes.
    pub fn max_unzip_size(mut self, size: u64) -> Self {
        self.max_unzip_size = Some(size);
        self
    }

    /// Set the maximum number of ZIP entries.
    pub fn max_zip_entries(mut self, count: usize) -> Self {
        self.max_zip_entries = Some(count);
        self
    }

    /// Set the read mode. `Lazy` skips auxiliary parts for faster
    /// read-only workloads. `Stream` is reserved for future use.
    pub fn read_mode(mut self, mode: ReadMode) -> Self {
        self.read_mode = mode;
        self
    }

    /// Set the auxiliary parts parsing policy.
    pub fn aux_parts(mut self, policy: AuxParts) -> Self {
        self.aux_parts = policy;
        self
    }

    /// Returns true when auxiliary parts should be skipped during open.
    /// Lazy/Stream modes always skip. Eager mode respects `aux_parts`.
    pub(crate) fn skip_aux_parts(&self) -> bool {
        match self.read_mode {
            ReadMode::Eager => self.aux_parts == AuxParts::Deferred,
            ReadMode::Lazy | ReadMode::Stream => true,
        }
    }

    /// Returns true when mode is `Eager`.
    #[allow(dead_code)]
    pub(crate) fn is_eager(&self) -> bool {
        self.read_mode == ReadMode::Eager
    }

    /// Returns true when mode is `Lazy`.
    #[allow(dead_code)]
    pub(crate) fn is_lazy(&self) -> bool {
        self.read_mode == ReadMode::Lazy
    }

    /// Returns true when mode is `Stream`.
    #[allow(dead_code)]
    pub(crate) fn is_stream(&self) -> bool {
        self.read_mode == ReadMode::Stream
    }

    /// Check whether a given sheet name should be parsed based on the `sheets` filter.
    pub(crate) fn should_parse_sheet(&self, name: &str) -> bool {
        match &self.sheets {
            None => true,
            Some(names) => names.iter().any(|n| n == name),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_options() {
        let opts = OpenOptions::default();
        assert!(opts.sheet_rows.is_none());
        assert!(opts.sheets.is_none());
        assert!(opts.max_unzip_size.is_none());
        assert!(opts.max_zip_entries.is_none());
        assert_eq!(opts.read_mode, ReadMode::Eager);
        assert!(!opts.skip_aux_parts());
    }

    #[test]
    fn test_builder_methods() {
        let opts = OpenOptions::new()
            .sheet_rows(100)
            .sheets(vec!["Sheet1".to_string()])
            .max_unzip_size(1_000_000)
            .max_zip_entries(500);
        assert_eq!(opts.sheet_rows, Some(100));
        assert_eq!(opts.sheets, Some(vec!["Sheet1".to_string()]));
        assert_eq!(opts.max_unzip_size, Some(1_000_000));
        assert_eq!(opts.max_zip_entries, Some(500));
    }

    #[test]
    fn test_read_mode_builder() {
        let opts = OpenOptions::new().read_mode(ReadMode::Lazy);
        assert_eq!(opts.read_mode, ReadMode::Lazy);
        assert!(opts.skip_aux_parts());
    }

    #[test]
    fn test_read_mode_default_is_eager() {
        let mode = ReadMode::default();
        assert_eq!(mode, ReadMode::Eager);
    }

    #[test]
    fn test_read_mode_combined_with_other_options() {
        let opts = OpenOptions::new().sheet_rows(50).read_mode(ReadMode::Lazy);
        assert_eq!(opts.sheet_rows, Some(50));
        assert!(opts.skip_aux_parts());
    }

    #[test]
    fn test_stream_mode_skips_aux_parts() {
        let opts = OpenOptions::new().read_mode(ReadMode::Stream);
        assert!(opts.skip_aux_parts());
        assert!(opts.is_stream());
        assert!(!opts.is_eager());
        assert!(!opts.is_lazy());
    }

    #[test]
    fn test_aux_parts_default_is_eager_load() {
        let opts = OpenOptions::default();
        assert_eq!(opts.aux_parts, AuxParts::EagerLoad);
    }

    #[test]
    fn test_aux_parts_deferred() {
        let opts = OpenOptions::new().aux_parts(AuxParts::Deferred);
        assert_eq!(opts.aux_parts, AuxParts::Deferred);
    }

    #[test]
    fn test_eager_mode_with_deferred_aux_skips_aux() {
        let opts = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::Deferred);
        assert!(opts.skip_aux_parts());
    }

    #[test]
    fn test_eager_mode_with_eager_aux_parses_all() {
        let opts = OpenOptions::new()
            .read_mode(ReadMode::Eager)
            .aux_parts(AuxParts::EagerLoad);
        assert!(!opts.skip_aux_parts());
    }

    #[test]
    fn test_should_parse_sheet_no_filter() {
        let opts = OpenOptions::default();
        assert!(opts.should_parse_sheet("Sheet1"));
        assert!(opts.should_parse_sheet("anything"));
    }

    #[test]
    fn test_should_parse_sheet_with_filter() {
        let opts = OpenOptions::new().sheets(vec!["Sales".to_string(), "Data".to_string()]);
        assert!(opts.should_parse_sheet("Sales"));
        assert!(opts.should_parse_sheet("Data"));
        assert!(!opts.should_parse_sheet("Sheet1"));
        assert!(!opts.should_parse_sheet("Other"));
    }

    #[test]
    fn test_helper_methods() {
        let eager = OpenOptions::new().read_mode(ReadMode::Eager);
        assert!(eager.is_eager());
        assert!(!eager.is_lazy());
        assert!(!eager.is_stream());

        let lazy = OpenOptions::new().read_mode(ReadMode::Lazy);
        assert!(!lazy.is_eager());
        assert!(lazy.is_lazy());
        assert!(!lazy.is_stream());

        let stream = OpenOptions::new().read_mode(ReadMode::Stream);
        assert!(!stream.is_eager());
        assert!(!stream.is_lazy());
        assert!(stream.is_stream());
    }
}
