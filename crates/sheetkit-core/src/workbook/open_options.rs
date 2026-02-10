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
}
