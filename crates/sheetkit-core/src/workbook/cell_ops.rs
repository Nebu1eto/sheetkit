use super::*;

impl Workbook {
    /// Get the value of a cell.
    ///
    /// Returns [`CellValue::Empty`] for cells that have no value or do not
    /// exist in the sheet data.
    pub fn get_cell_value(&self, sheet: &str, cell: &str) -> Result<CellValue> {
        let ws = self
            .worksheets
            .iter()
            .find(|(name, _)| name == sheet)
            .map(|(_, ws)| ws)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })?;

        let (col, row) = cell_name_to_coordinates(cell)?;
        let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(col, row)?;

        // Find the row.
        let xml_row = ws.sheet_data.rows.iter().find(|r| r.r == row);
        let xml_row = match xml_row {
            Some(r) => r,
            None => return Ok(CellValue::Empty),
        };

        // Find the cell.
        let xml_cell = xml_row.cells.iter().find(|c| c.r == cell_ref);
        let xml_cell = match xml_cell {
            Some(c) => c,
            None => return Ok(CellValue::Empty),
        };

        self.xml_cell_to_value(xml_cell)
    }

    /// Set the value of a cell.
    ///
    /// The value can be any type that implements `Into<CellValue>`, including
    /// `&str`, `String`, `f64`, `i32`, `i64`, and `bool`.
    ///
    /// Setting a cell to [`CellValue::Empty`] removes the cell from the row.
    pub fn set_cell_value(
        &mut self,
        sheet: &str,
        cell: &str,
        value: impl Into<CellValue>,
    ) -> Result<()> {
        let value = value.into();

        // Validate string length.
        if let CellValue::String(ref s) = value {
            if s.len() > MAX_CELL_CHARS {
                return Err(Error::CellValueTooLong {
                    length: s.len(),
                    max: MAX_CELL_CHARS,
                });
            }
        }

        let ws = self
            .worksheets
            .iter_mut()
            .find(|(name, _)| name == sheet)
            .map(|(_, ws)| ws)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })?;

        let (col, row_num) = cell_name_to_coordinates(cell)?;
        let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(col, row_num)?;

        // Find or create the row (keep rows sorted by row number).
        let row_idx = match ws.sheet_data.rows.iter().position(|r| r.r >= row_num) {
            Some(idx) if ws.sheet_data.rows[idx].r == row_num => idx,
            Some(idx) => {
                ws.sheet_data.rows.insert(idx, new_row(row_num));
                idx
            }
            None => {
                ws.sheet_data.rows.push(new_row(row_num));
                ws.sheet_data.rows.len() - 1
            }
        };

        let row = &mut ws.sheet_data.rows[row_idx];

        // Handle Empty: remove the cell if present.
        if value == CellValue::Empty {
            row.cells.retain(|c| c.r != cell_ref);
            return Ok(());
        }

        // Find or create the cell.
        let cell_idx = match row.cells.iter().position(|c| c.r == cell_ref) {
            Some(idx) => idx,
            None => {
                // Insert in column order.
                let insert_pos = row
                    .cells
                    .iter()
                    .position(|c| {
                        cell_name_to_coordinates(&c.r)
                            .map(|(c_col, _)| c_col > col)
                            .unwrap_or(false)
                    })
                    .unwrap_or(row.cells.len());
                row.cells.insert(
                    insert_pos,
                    Cell {
                        r: cell_ref,
                        s: None,
                        t: None,
                        v: None,
                        f: None,
                        is: None,
                    },
                );
                insert_pos
            }
        };

        let xml_cell = &mut row.cells[cell_idx];
        value_to_xml_cell(&mut self.sst_runtime, xml_cell, value);

        Ok(())
    }

    /// Convert an XML Cell to a CellValue.
    pub(crate) fn xml_cell_to_value(&self, xml_cell: &Cell) -> Result<CellValue> {
        // Check for formula first.
        if let Some(ref formula) = xml_cell.f {
            let expr = formula.value.clone().unwrap_or_default();
            let result = match (&xml_cell.t, &xml_cell.v) {
                (Some(t), Some(v)) if t == "b" => Some(Box::new(CellValue::Bool(v == "1"))),
                (Some(t), Some(v)) if t == "e" => Some(Box::new(CellValue::Error(v.clone()))),
                (_, Some(v)) => v
                    .parse::<f64>()
                    .ok()
                    .map(|n| Box::new(CellValue::Number(n))),
                _ => None,
            };
            return Ok(CellValue::Formula { expr, result });
        }

        let cell_type = xml_cell.t.as_deref();
        let cell_value = xml_cell.v.as_deref();

        match (cell_type, cell_value) {
            // Shared string
            (Some("s"), Some(v)) => {
                let idx: usize = v
                    .parse()
                    .map_err(|_| Error::Internal(format!("invalid SST index: {v}")))?;
                let s = self.sst_runtime.get(idx).unwrap_or("").to_string();
                Ok(CellValue::String(s))
            }
            // Boolean
            (Some("b"), Some(v)) => Ok(CellValue::Bool(v == "1")),
            // Error
            (Some("e"), Some(v)) => Ok(CellValue::Error(v.to_string())),
            // Inline string
            (Some("inlineStr"), _) => {
                let s = xml_cell
                    .is
                    .as_ref()
                    .and_then(|is| is.t.clone())
                    .unwrap_or_default();
                Ok(CellValue::String(s))
            }
            // Formula string (cached string result)
            (Some("str"), Some(v)) => Ok(CellValue::String(v.to_string())),
            // Number (explicit or default type) -- may be a date if styled.
            (None | Some("n"), Some(v)) => {
                let n: f64 = v
                    .parse()
                    .map_err(|_| Error::Internal(format!("invalid number: {v}")))?;
                // Check whether this cell has a date number format.
                if self.is_date_styled_cell(xml_cell) {
                    return Ok(CellValue::Date(n));
                }
                Ok(CellValue::Number(n))
            }
            // No value
            _ => Ok(CellValue::Empty),
        }
    }

    /// Check whether a cell's style indicates a date/time number format.
    pub(crate) fn is_date_styled_cell(&self, xml_cell: &Cell) -> bool {
        let style_idx = match xml_cell.s {
            Some(idx) => idx as usize,
            None => return false,
        };
        let xf = match self.stylesheet.cell_xfs.xfs.get(style_idx) {
            Some(xf) => xf,
            None => return false,
        };
        let num_fmt_id = xf.num_fmt_id.unwrap_or(0);
        // Check built-in date format IDs.
        if crate::cell::is_date_num_fmt(num_fmt_id) {
            return true;
        }
        // Check custom number formats for date patterns.
        if num_fmt_id >= 164 {
            if let Some(ref num_fmts) = self.stylesheet.num_fmts {
                if let Some(nf) = num_fmts
                    .num_fmts
                    .iter()
                    .find(|nf| nf.num_fmt_id == num_fmt_id)
                {
                    return crate::cell::is_date_format_code(&nf.format_code);
                }
            }
        }
        false
    }

    /// Register a new style and return its ID.
    ///
    /// The style is deduplicated: if an identical style already exists in
    /// the stylesheet, the existing ID is returned.
    pub fn add_style(&mut self, style: &crate::style::Style) -> Result<u32> {
        crate::style::add_style(&mut self.stylesheet, style)
    }

    /// Get the style ID applied to a cell.
    ///
    /// Returns `None` if the cell does not exist or has no explicit style
    /// (i.e. uses the default style 0).
    pub fn get_cell_style(&self, sheet: &str, cell: &str) -> Result<Option<u32>> {
        let ws = self.worksheet_ref(sheet)?;

        let (col, row) = cell_name_to_coordinates(cell)?;
        let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(col, row)?;

        // Find the row.
        let xml_row = match ws.sheet_data.rows.iter().find(|r| r.r == row) {
            Some(r) => r,
            None => return Ok(None),
        };

        // Find the cell.
        let xml_cell = match xml_row.cells.iter().find(|c| c.r == cell_ref) {
            Some(c) => c,
            None => return Ok(None),
        };

        Ok(xml_cell.s)
    }

    /// Set the style ID for a cell.
    ///
    /// If the cell does not exist, an empty cell with just the style is created.
    /// The `style_id` must be a valid index in cellXfs.
    pub fn set_cell_style(&mut self, sheet: &str, cell: &str, style_id: u32) -> Result<()> {
        // Validate the style_id.
        if style_id as usize >= self.stylesheet.cell_xfs.xfs.len() {
            return Err(Error::StyleNotFound { id: style_id });
        }

        let ws = self
            .worksheets
            .iter_mut()
            .find(|(name, _)| name == sheet)
            .map(|(_, ws)| ws)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })?;

        let (col, row_num) = cell_name_to_coordinates(cell)?;
        let cell_ref = crate::utils::cell_ref::coordinates_to_cell_name(col, row_num)?;

        // Find or create the row (keep rows sorted by row number).
        let row_idx = match ws.sheet_data.rows.iter().position(|r| r.r >= row_num) {
            Some(idx) if ws.sheet_data.rows[idx].r == row_num => idx,
            Some(idx) => {
                ws.sheet_data.rows.insert(idx, new_row(row_num));
                idx
            }
            None => {
                ws.sheet_data.rows.push(new_row(row_num));
                ws.sheet_data.rows.len() - 1
            }
        };

        let row = &mut ws.sheet_data.rows[row_idx];

        // Find or create the cell.
        let cell_idx = match row.cells.iter().position(|c| c.r == cell_ref) {
            Some(idx) => idx,
            None => {
                // Insert in column order.
                let insert_pos = row
                    .cells
                    .iter()
                    .position(|c| {
                        cell_name_to_coordinates(&c.r)
                            .map(|(c_col, _)| c_col > col)
                            .unwrap_or(false)
                    })
                    .unwrap_or(row.cells.len());
                row.cells.insert(
                    insert_pos,
                    Cell {
                        r: cell_ref,
                        s: None,
                        t: None,
                        v: None,
                        f: None,
                        is: None,
                    },
                );
                insert_pos
            }
        };

        row.cells[cell_idx].s = Some(style_id);
        Ok(())
    }

    /// Merge a range of cells on the given sheet.
    ///
    /// `top_left` and `bottom_right` are cell references like "A1" and "C3".
    /// Returns an error if the range overlaps with an existing merge region.
    pub fn merge_cells(&mut self, sheet: &str, top_left: &str, bottom_right: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::merge::merge_cells(ws, top_left, bottom_right)
    }

    /// Remove a merged cell range from the given sheet.
    ///
    /// `reference` is the exact range string like "A1:C3".
    pub fn unmerge_cell(&mut self, sheet: &str, reference: &str) -> Result<()> {
        let ws = self.worksheet_mut(sheet)?;
        crate::merge::unmerge_cell(ws, reference)
    }

    /// Get all merged cell ranges on the given sheet.
    ///
    /// Returns a list of range strings like `["A1:B2", "D1:F3"]`.
    pub fn get_merge_cells(&self, sheet: &str) -> Result<Vec<String>> {
        let ws = self.worksheet_ref(sheet)?;
        Ok(crate::merge::get_merge_cells(ws))
    }

    /// Set a formula on a cell.
    ///
    /// This is a convenience wrapper around [`set_cell_value`] with
    /// [`CellValue::Formula`].
    pub fn set_cell_formula(&mut self, sheet: &str, cell: &str, formula: &str) -> Result<()> {
        self.set_cell_value(
            sheet,
            cell,
            CellValue::Formula {
                expr: formula.to_string(),
                result: None,
            },
        )
    }

    /// Fill a range of cells with a formula, adjusting row references for each
    /// row relative to the first cell in the range.
    ///
    /// `range` is an A1-style range like `"D2:D10"`. The `formula` is the base
    /// formula for the first cell of the range. For each subsequent row, the
    /// row references in the formula are shifted by the row offset. Absolute
    /// row references (`$1`) are left unchanged.
    pub fn fill_formula(&mut self, sheet: &str, range: &str, formula: &str) -> Result<()> {
        let parts: Vec<&str> = range.split(':').collect();
        if parts.len() != 2 {
            return Err(Error::InvalidCellReference(format!(
                "invalid range: {range}"
            )));
        }
        let (start_col, start_row) = cell_name_to_coordinates(parts[0])?;
        let (end_col, end_row) = cell_name_to_coordinates(parts[1])?;

        if start_col != end_col {
            return Err(Error::InvalidCellReference(
                "fill_formula only supports single-column ranges".to_string(),
            ));
        }

        for row in start_row..=end_row {
            let row_offset = row as i32 - start_row as i32;
            let adjusted = if row_offset == 0 {
                formula.to_string()
            } else {
                crate::cell_ref_shift::shift_cell_references_with_abs(
                    formula,
                    |col, r, _abs_col, abs_row| {
                        if abs_row {
                            (col, r)
                        } else {
                            (col, (r as i32 + row_offset) as u32)
                        }
                    },
                )?
            };
            let cell_ref =
                crate::utils::cell_ref::coordinates_to_cell_name(start_col, row)?;
            self.set_cell_formula(sheet, &cell_ref, &adjusted)?;
        }
        Ok(())
    }

    /// Set a cell to a rich text value (multiple formatted runs).
    pub fn set_cell_rich_text(
        &mut self,
        sheet: &str,
        cell: &str,
        runs: Vec<crate::rich_text::RichTextRun>,
    ) -> Result<()> {
        self.set_cell_value(sheet, cell, CellValue::RichString(runs))
    }

    /// Get rich text runs for a cell, if it contains rich text.
    ///
    /// Returns `None` if the cell is empty, contains a plain string, or holds
    /// a non-string value.
    pub fn get_cell_rich_text(
        &self,
        sheet: &str,
        cell: &str,
    ) -> Result<Option<Vec<crate::rich_text::RichTextRun>>> {
        let (col, row) = cell_name_to_coordinates(cell)?;
        let sheet_idx = self
            .worksheets
            .iter()
            .position(|(name, _)| name == sheet)
            .ok_or_else(|| Error::SheetNotFound {
                name: sheet.to_string(),
            })?;
        let ws = &self.worksheets[sheet_idx].1;

        for xml_row in &ws.sheet_data.rows {
            if xml_row.r == row {
                for xml_cell in &xml_row.cells {
                    let (cc, cr) = cell_name_to_coordinates(&xml_cell.r)?;
                    if cc == col && cr == row {
                        if xml_cell.t.as_deref() == Some("s") {
                            if let Some(ref v) = xml_cell.v {
                                if let Ok(idx) = v.parse::<usize>() {
                                    return Ok(self.sst_runtime.get_rich_text(idx));
                                }
                            }
                        }
                        return Ok(None);
                    }
                }
            }
        }
        Ok(None)
    }
}

/// Write a CellValue into an XML Cell (mutating it in place).
pub(crate) fn value_to_xml_cell(
    sst: &mut SharedStringTable,
    xml_cell: &mut Cell,
    value: CellValue,
) {
    // Clear previous values.
    xml_cell.t = None;
    xml_cell.v = None;
    xml_cell.f = None;
    xml_cell.is = None;

    match value {
        CellValue::String(s) => {
            let idx = sst.add(&s);
            xml_cell.t = Some("s".to_string());
            xml_cell.v = Some(idx.to_string());
        }
        CellValue::Number(n) => {
            xml_cell.v = Some(n.to_string());
        }
        CellValue::Date(serial) => {
            // Dates are stored as numbers in Excel. The style must apply a
            // date number format for correct display.
            xml_cell.v = Some(serial.to_string());
        }
        CellValue::Bool(b) => {
            xml_cell.t = Some("b".to_string());
            xml_cell.v = Some(if b { "1" } else { "0" }.to_string());
        }
        CellValue::Formula { expr, .. } => {
            xml_cell.f = Some(CellFormula {
                t: None,
                reference: None,
                si: None,
                value: Some(expr),
            });
        }
        CellValue::Error(e) => {
            xml_cell.t = Some("e".to_string());
            xml_cell.v = Some(e);
        }
        CellValue::Empty => {
            // Already cleared above; the caller should have removed the cell.
        }
        CellValue::RichString(runs) => {
            let idx = sst.add_rich_text(&runs);
            xml_cell.t = Some("s".to_string());
            xml_cell.v = Some(idx.to_string());
        }
    }
}

/// Create a new empty row with the given 1-based row number.
pub(crate) fn new_row(row_num: u32) -> Row {
    Row {
        r: row_num,
        spans: None,
        s: None,
        custom_format: None,
        ht: None,
        hidden: None,
        custom_height: None,
        outline_level: None,
        cells: vec![],
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn test_set_and_get_string_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Hello").unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::String("Hello".to_string()));
    }

    #[test]
    fn test_set_and_get_number_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "B2", 42.5f64).unwrap();
        let val = wb.get_cell_value("Sheet1", "B2").unwrap();
        assert_eq!(val, CellValue::Number(42.5));
    }

    #[test]
    fn test_set_and_get_bool_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "C3", true).unwrap();
        let val = wb.get_cell_value("Sheet1", "C3").unwrap();
        assert_eq!(val, CellValue::Bool(true));

        wb.set_cell_value("Sheet1", "D4", false).unwrap();
        let val = wb.get_cell_value("Sheet1", "D4").unwrap();
        assert_eq!(val, CellValue::Bool(false));
    }

    #[test]
    fn test_set_value_sheet_not_found() {
        let mut wb = Workbook::new();
        let result = wb.set_cell_value("NoSuchSheet", "A1", "test");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_get_value_sheet_not_found() {
        let wb = Workbook::new();
        let result = wb.get_cell_value("NoSuchSheet", "A1");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_get_empty_cell_returns_empty() {
        let wb = Workbook::new();
        let val = wb.get_cell_value("Sheet1", "Z99").unwrap();
        assert_eq!(val, CellValue::Empty);
    }

    #[test]
    fn test_cell_value_roundtrip_save_open() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("cell_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Hello").unwrap();
        wb.set_cell_value("Sheet1", "B1", 42.0f64).unwrap();
        wb.set_cell_value("Sheet1", "C1", true).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        assert_eq!(
            wb2.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("Hello".to_string())
        );
        assert_eq!(
            wb2.get_cell_value("Sheet1", "B1").unwrap(),
            CellValue::Number(42.0)
        );
        assert_eq!(
            wb2.get_cell_value("Sheet1", "C1").unwrap(),
            CellValue::Bool(true)
        );
    }

    #[test]
    fn test_set_empty_value_clears_cell() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "test").unwrap();
        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("test".to_string())
        );

        wb.set_cell_value("Sheet1", "A1", CellValue::Empty).unwrap();
        assert_eq!(wb.get_cell_value("Sheet1", "A1").unwrap(), CellValue::Empty);
    }

    #[test]
    fn test_string_too_long_returns_error() {
        let mut wb = Workbook::new();
        let long_string = "x".repeat(MAX_CELL_CHARS + 1);
        let result = wb.set_cell_value("Sheet1", "A1", long_string.as_str());
        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            Error::CellValueTooLong { .. }
        ));
    }

    #[test]
    fn test_set_multiple_cells_same_row() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "first").unwrap();
        wb.set_cell_value("Sheet1", "B1", "second").unwrap();
        wb.set_cell_value("Sheet1", "C1", "third").unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("first".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "B1").unwrap(),
            CellValue::String("second".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "C1").unwrap(),
            CellValue::String("third".to_string())
        );
    }

    #[test]
    fn test_overwrite_cell_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "original").unwrap();
        wb.set_cell_value("Sheet1", "A1", "updated").unwrap();

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("updated".to_string())
        );
    }

    #[test]
    fn test_set_and_get_error_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::Error("#DIV/0!".to_string()))
            .unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::Error("#DIV/0!".to_string()));
    }

    #[test]
    fn test_set_and_get_date_value() {
        use crate::style::{builtin_num_fmts, NumFmtStyle, Style};

        let mut wb = Workbook::new();
        // Create a date style.
        let style_id = wb
            .add_style(&Style {
                num_fmt: Some(NumFmtStyle::Builtin(builtin_num_fmts::DATE_MDY)),
                ..Style::default()
            })
            .unwrap();

        // Set a date value.
        let date_serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 6, 15).unwrap());
        wb.set_cell_value("Sheet1", "A1", CellValue::Date(date_serial))
            .unwrap();
        wb.set_cell_style("Sheet1", "A1", style_id).unwrap();

        // Get the value back -- it should be Date because the cell has a date style.
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::Date(date_serial));
    }

    #[test]
    fn test_date_value_without_style_returns_number() {
        let mut wb = Workbook::new();
        // Set a date value without a date style.
        let date_serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 6, 15).unwrap());
        wb.set_cell_value("Sheet1", "A1", CellValue::Date(date_serial))
            .unwrap();

        // Without a date style, the value is read back as Number.
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::Number(date_serial));
    }

    #[test]
    fn test_date_value_roundtrip_through_save() {
        use crate::style::{builtin_num_fmts, NumFmtStyle, Style};

        let mut wb = Workbook::new();
        let style_id = wb
            .add_style(&Style {
                num_fmt: Some(NumFmtStyle::Builtin(builtin_num_fmts::DATETIME)),
                ..Style::default()
            })
            .unwrap();

        let dt = chrono::NaiveDate::from_ymd_opt(2024, 3, 15)
            .unwrap()
            .and_hms_opt(14, 30, 0)
            .unwrap();
        let serial = crate::cell::datetime_to_serial(dt);
        wb.set_cell_value("Sheet1", "A1", CellValue::Date(serial))
            .unwrap();
        wb.set_cell_style("Sheet1", "A1", style_id).unwrap();

        let tmp = tempfile::NamedTempFile::new().unwrap();
        let path = tmp.path().to_str().unwrap();
        wb.save(path).unwrap();

        let wb2 = Workbook::open(path).unwrap();
        let val = wb2.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::Date(serial));
    }

    #[test]
    fn test_date_from_naive_date_conversion() {
        let date = chrono::NaiveDate::from_ymd_opt(2024, 1, 1).unwrap();
        let cv: CellValue = date.into();
        match cv {
            CellValue::Date(s) => {
                let roundtripped = crate::cell::serial_to_date(s).unwrap();
                assert_eq!(roundtripped, date);
            }
            _ => panic!("expected Date variant"),
        }
    }

    #[test]
    fn test_set_and_get_formula_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value(
            "Sheet1",
            "A1",
            CellValue::Formula {
                expr: "SUM(B1:B10)".to_string(),
                result: None,
            },
        )
        .unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        match val {
            CellValue::Formula { expr, .. } => {
                assert_eq!(expr, "SUM(B1:B10)");
            }
            other => panic!("expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_set_i32_value() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", 100i32).unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::Number(100.0));
    }

    #[test]
    fn test_set_string_at_max_length() {
        let mut wb = Workbook::new();
        let max_string = "x".repeat(MAX_CELL_CHARS);
        wb.set_cell_value("Sheet1", "A1", max_string.as_str())
            .unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::String(max_string));
    }

    #[test]
    fn test_set_cells_different_rows() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "row1").unwrap();
        wb.set_cell_value("Sheet1", "A3", "row3").unwrap();
        wb.set_cell_value("Sheet1", "A2", "row2").unwrap(); // inserted between

        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("row1".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "A2").unwrap(),
            CellValue::String("row2".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "A3").unwrap(),
            CellValue::String("row3".to_string())
        );
    }

    #[test]
    fn test_string_deduplication_in_sst() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "same").unwrap();
        wb.set_cell_value("Sheet1", "A2", "same").unwrap();
        wb.set_cell_value("Sheet1", "A3", "different").unwrap();

        // Both A1 and A2 should point to the same SST index
        assert_eq!(wb.sst_runtime.len(), 2);
        assert_eq!(
            wb.get_cell_value("Sheet1", "A1").unwrap(),
            CellValue::String("same".to_string())
        );
        assert_eq!(
            wb.get_cell_value("Sheet1", "A2").unwrap(),
            CellValue::String("same".to_string())
        );
    }

    #[test]
    fn test_add_style_returns_id() {
        use crate::style::{FontStyle, Style};

        let mut wb = Workbook::new();
        let style = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };
        let id = wb.add_style(&style).unwrap();
        assert!(id > 0);
    }

    #[test]
    fn test_get_cell_style_unstyled_cell_returns_none() {
        let wb = Workbook::new();
        let result = wb.get_cell_style("Sheet1", "A1").unwrap();
        assert!(result.is_none());
    }

    #[test]
    fn test_set_cell_style_on_existing_value() {
        use crate::style::{FontStyle, Style};

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Hello").unwrap();

        let style = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };
        let style_id = wb.add_style(&style).unwrap();
        wb.set_cell_style("Sheet1", "A1", style_id).unwrap();

        let retrieved_id = wb.get_cell_style("Sheet1", "A1").unwrap();
        assert_eq!(retrieved_id, Some(style_id));

        // The value should still be there.
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::String("Hello".to_string()));
    }

    #[test]
    fn test_set_cell_style_on_empty_cell_creates_cell() {
        use crate::style::{FontStyle, Style};

        let mut wb = Workbook::new();
        let style = Style {
            font: Some(FontStyle {
                bold: true,
                ..FontStyle::default()
            }),
            ..Style::default()
        };
        let style_id = wb.add_style(&style).unwrap();

        // Set style on a cell that doesn't exist yet.
        wb.set_cell_style("Sheet1", "B5", style_id).unwrap();

        let retrieved_id = wb.get_cell_style("Sheet1", "B5").unwrap();
        assert_eq!(retrieved_id, Some(style_id));

        // The cell value should be empty.
        let val = wb.get_cell_value("Sheet1", "B5").unwrap();
        assert_eq!(val, CellValue::Empty);
    }

    #[test]
    fn test_set_cell_style_invalid_id() {
        let mut wb = Workbook::new();
        let result = wb.set_cell_style("Sheet1", "A1", 999);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::StyleNotFound { .. }));
    }

    #[test]
    fn test_set_cell_style_sheet_not_found() {
        let mut wb = Workbook::new();
        let style = crate::style::Style::default();
        let style_id = wb.add_style(&style).unwrap();
        let result = wb.set_cell_style("NoSuchSheet", "A1", style_id);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_get_cell_style_sheet_not_found() {
        let wb = Workbook::new();
        let result = wb.get_cell_style("NoSuchSheet", "A1");
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), Error::SheetNotFound { .. }));
    }

    #[test]
    fn test_style_roundtrip_save_open() {
        use crate::style::{
            AlignmentStyle, BorderLineStyle, BorderSideStyle, BorderStyle, FillStyle, FontStyle,
            HorizontalAlign, NumFmtStyle, PatternType, Style, StyleColor, VerticalAlign,
        };

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("style_roundtrip.xlsx");

        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", "Styled").unwrap();

        let style = Style {
            font: Some(FontStyle {
                name: Some("Arial".to_string()),
                size: Some(14.0),
                bold: true,
                italic: true,
                color: Some(StyleColor::Rgb("FFFF0000".to_string())),
                ..FontStyle::default()
            }),
            fill: Some(FillStyle {
                pattern: PatternType::Solid,
                fg_color: Some(StyleColor::Rgb("FFFFFF00".to_string())),
                bg_color: None,
                gradient: None,
            }),
            border: Some(BorderStyle {
                left: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                right: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                top: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                bottom: Some(BorderSideStyle {
                    style: BorderLineStyle::Thin,
                    color: None,
                }),
                diagonal: None,
            }),
            alignment: Some(AlignmentStyle {
                horizontal: Some(HorizontalAlign::Center),
                vertical: Some(VerticalAlign::Center),
                wrap_text: true,
                ..AlignmentStyle::default()
            }),
            num_fmt: Some(NumFmtStyle::Custom("#,##0.00".to_string())),
            protection: None,
        };
        let style_id = wb.add_style(&style).unwrap();
        wb.set_cell_style("Sheet1", "A1", style_id).unwrap();
        wb.save(&path).unwrap();

        // Re-open and verify.
        let wb2 = Workbook::open(&path).unwrap();
        let retrieved_id = wb2.get_cell_style("Sheet1", "A1").unwrap();
        assert_eq!(retrieved_id, Some(style_id));

        // Verify the value is still there.
        let val = wb2.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val, CellValue::String("Styled".to_string()));

        // Reverse-lookup the style to verify components survived the roundtrip.
        let retrieved_style = crate::style::get_style(&wb2.stylesheet, style_id).unwrap();
        assert!(retrieved_style.font.is_some());
        let font = retrieved_style.font.unwrap();
        assert!(font.bold);
        assert!(font.italic);
        assert_eq!(font.name, Some("Arial".to_string()));

        assert!(retrieved_style.fill.is_some());
        let fill = retrieved_style.fill.unwrap();
        assert_eq!(fill.pattern, PatternType::Solid);

        assert!(retrieved_style.alignment.is_some());
        let align = retrieved_style.alignment.unwrap();
        assert_eq!(align.horizontal, Some(HorizontalAlign::Center));
        assert_eq!(align.vertical, Some(VerticalAlign::Center));
        assert!(align.wrap_text);
    }

    #[test]
    fn test_set_and_get_cell_rich_text() {
        use crate::rich_text::RichTextRun;

        let mut wb = Workbook::new();
        let runs = vec![
            RichTextRun {
                text: "Bold".to_string(),
                font: None,
                size: None,
                bold: true,
                italic: false,
                color: None,
            },
            RichTextRun {
                text: " Normal".to_string(),
                font: None,
                size: None,
                bold: false,
                italic: false,
                color: None,
            },
        ];
        wb.set_cell_rich_text("Sheet1", "A1", runs.clone()).unwrap();

        // The cell value should be a shared string whose plain text is "Bold Normal".
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        assert_eq!(val.to_string(), "Bold Normal");

        // get_cell_rich_text should return the runs.
        let got = wb.get_cell_rich_text("Sheet1", "A1").unwrap();
        assert!(got.is_some());
        let got_runs = got.unwrap();
        assert_eq!(got_runs.len(), 2);
        assert_eq!(got_runs[0].text, "Bold");
        assert!(got_runs[0].bold);
        assert_eq!(got_runs[1].text, " Normal");
        assert!(!got_runs[1].bold);
    }

    #[test]
    fn test_get_cell_rich_text_returns_none_for_plain() {
        let mut wb = Workbook::new();
        wb.set_cell_value("Sheet1", "A1", CellValue::String("plain".to_string()))
            .unwrap();
        let got = wb.get_cell_rich_text("Sheet1", "A1").unwrap();
        assert!(got.is_none());
    }

    #[test]
    fn test_rich_text_roundtrip_save_open() {
        use crate::rich_text::RichTextRun;

        let dir = TempDir::new().unwrap();
        let path = dir.path().join("rich_text.xlsx");

        // Note: quick-xml's serde deserializer trims leading and trailing
        // whitespace from text content. To avoid false failures, test text
        // values must not rely on boundary whitespace being preserved.
        let mut wb = Workbook::new();
        let runs = vec![
            RichTextRun {
                text: "Hello".to_string(),
                font: Some("Arial".to_string()),
                size: Some(14.0),
                bold: true,
                italic: false,
                color: Some("#FF0000".to_string()),
            },
            RichTextRun {
                text: "World".to_string(),
                font: None,
                size: None,
                bold: false,
                italic: true,
                color: None,
            },
        ];
        wb.set_cell_rich_text("Sheet1", "B2", runs).unwrap();
        wb.save(&path).unwrap();

        let wb2 = Workbook::open(&path).unwrap();
        let val = wb2.get_cell_value("Sheet1", "B2").unwrap();
        assert_eq!(val.to_string(), "HelloWorld");

        let got = wb2.get_cell_rich_text("Sheet1", "B2").unwrap();
        assert!(got.is_some());
        let got_runs = got.unwrap();
        assert_eq!(got_runs.len(), 2);
        assert_eq!(got_runs[0].text, "Hello");
        assert!(got_runs[0].bold);
        assert_eq!(got_runs[0].font.as_deref(), Some("Arial"));
        assert_eq!(got_runs[0].size, Some(14.0));
        assert_eq!(got_runs[0].color.as_deref(), Some("#FF0000"));
        assert_eq!(got_runs[1].text, "World");
        assert!(got_runs[1].italic);
        assert!(!got_runs[1].bold);
    }

    #[test]
    fn test_set_cell_formula() {
        let mut wb = Workbook::new();
        wb.set_cell_formula("Sheet1", "A1", "SUM(B1:B10)").unwrap();
        let val = wb.get_cell_value("Sheet1", "A1").unwrap();
        match val {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "SUM(B1:B10)"),
            other => panic!("expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_fill_formula_basic() {
        let mut wb = Workbook::new();
        wb.fill_formula("Sheet1", "D2:D5", "SUM(A2:C2)").unwrap();

        // D2 should have the base formula unchanged
        match wb.get_cell_value("Sheet1", "D2").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "SUM(A2:C2)"),
            other => panic!("D2: expected Formula, got {:?}", other),
        }
        // D3 should have row shifted by 1
        match wb.get_cell_value("Sheet1", "D3").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "SUM(A3:C3)"),
            other => panic!("D3: expected Formula, got {:?}", other),
        }
        // D4 should have row shifted by 2
        match wb.get_cell_value("Sheet1", "D4").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "SUM(A4:C4)"),
            other => panic!("D4: expected Formula, got {:?}", other),
        }
        // D5 should have row shifted by 3
        match wb.get_cell_value("Sheet1", "D5").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "SUM(A5:C5)"),
            other => panic!("D5: expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_fill_formula_preserves_absolute_refs() {
        let mut wb = Workbook::new();
        wb.fill_formula("Sheet1", "B1:B3", "$A$1*A1").unwrap();

        match wb.get_cell_value("Sheet1", "B1").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "$A$1*A1"),
            other => panic!("B1: expected Formula, got {:?}", other),
        }
        match wb.get_cell_value("Sheet1", "B2").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "$A$1*A2"),
            other => panic!("B2: expected Formula, got {:?}", other),
        }
        match wb.get_cell_value("Sheet1", "B3").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "$A$1*A3"),
            other => panic!("B3: expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_fill_formula_single_cell() {
        let mut wb = Workbook::new();
        wb.fill_formula("Sheet1", "A1:A1", "B1+C1").unwrap();
        match wb.get_cell_value("Sheet1", "A1").unwrap() {
            CellValue::Formula { expr, .. } => assert_eq!(expr, "B1+C1"),
            other => panic!("expected Formula, got {:?}", other),
        }
    }

    #[test]
    fn test_fill_formula_invalid_range() {
        let mut wb = Workbook::new();
        assert!(wb.fill_formula("Sheet1", "INVALID", "A1").is_err());
    }

    #[test]
    fn test_fill_formula_multi_column_range_rejected() {
        let mut wb = Workbook::new();
        assert!(wb.fill_formula("Sheet1", "A1:B5", "C1").is_err());
    }
}
