use crate::error::{Error, Result};
use crate::utils::cell_ref::{column_name_to_number, column_number_to_name};
use crate::utils::constants::{MAX_COLUMNS, MAX_ROWS};

#[derive(Debug, Clone, Copy)]
struct ParsedCellRef {
    end: usize,
    abs_col: bool,
    abs_row: bool,
    col: u32,
    row: u32,
}

/// One endpoint of a reference. A missing coordinate marks a whole-row or
/// whole-column endpoint, such as `A` in `A:C` or `3` in `3:5`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct RefEndpoint {
    pub(crate) col: Option<u32>,
    pub(crate) row: Option<u32>,
    pub(crate) abs_col: bool,
    pub(crate) abs_row: bool,
}

/// A single cell, a cell range, or a whole-row/column range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct RefArea {
    pub(crate) start: RefEndpoint,
    pub(crate) end: Option<RefEndpoint>,
}

/// How one reference token is rewritten.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum RefEdit {
    /// Emit the original token text unchanged.
    Keep,
    /// Emit the given area.
    Replace(RefArea),
    /// The referenced cells were deleted; emit `#REF!`.
    Delete,
}

/// A row or column insertion/removal applied to one worksheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum StructuralShift {
    InsertRows { start: u32, count: u32 },
    RemoveRows { start: u32, count: u32 },
    InsertCols { start: u32, count: u32 },
    RemoveCols { start: u32, count: u32 },
}

impl StructuralShift {
    /// Returns `(is_row_axis, is_removal, start, count)`.
    fn parts(self) -> (bool, bool, u32, u32) {
        match self {
            Self::InsertRows { start, count } => (true, false, start, count),
            Self::RemoveRows { start, count } => (true, true, start, count),
            Self::InsertCols { start, count } => (false, false, start, count),
            Self::RemoveCols { start, count } => (false, true, start, count),
        }
    }

    /// Whether the edit removes rows or columns.
    pub(crate) fn is_removal(self) -> bool {
        self.parts().1
    }

    /// Whether the edit moves rows (as opposed to columns).
    pub(crate) fn is_row_axis(self) -> bool {
        self.parts().0
    }

    /// Shift one coordinate on the edited axis. Returns `None` when the
    /// coordinate lies inside a removed band.
    pub(crate) fn shift_index(self, value: u32) -> Option<u32> {
        let (_, remove, start, count) = self.parts();
        if value < start {
            return Some(value);
        }
        if !remove {
            return Some(value.saturating_add(count));
        }
        let last_removed = start.saturating_add(count.saturating_sub(1));
        if value <= last_removed {
            None
        } else {
            Some(value - count)
        }
    }

    /// Shift a cell position, or return `None` when the cell was removed.
    pub(crate) fn shift_cell(self, col: u32, row: u32) -> Option<(u32, u32)> {
        if self.is_row_axis() {
            self.shift_index(row).map(|row| (col, row))
        } else {
            self.shift_index(col).map(|col| (col, row))
        }
    }

    /// Shift a position that must survive the edit, such as a drawing anchor
    /// or a selection. Positions inside a removed band collapse onto the first
    /// row or column that follows the band.
    pub(crate) fn shift_point(self, col: u32, row: u32) -> (u32, u32) {
        let (rows, _, start, _) = self.parts();
        let map = |value: u32| self.shift_index(value).unwrap_or(start);
        if rows {
            (col, map(row))
        } else {
            (map(col), row)
        }
    }

    /// Shift an inclusive span on the edited axis. A span that partially
    /// overlaps a removed band shrinks; a span fully inside it is removed.
    fn shift_span(self, lo: u32, hi: u32) -> Option<(u32, u32)> {
        let (_, _, start, _) = self.parts();
        match (self.shift_index(lo), self.shift_index(hi)) {
            (Some(lo), Some(hi)) => Some((lo, hi)),
            (None, Some(hi)) => Some((start, hi)),
            // `lo` survived below the band, so `start` is at least 2.
            (Some(lo), None) => Some((lo, start - 1)),
            (None, None) => None,
        }
    }

    /// Apply the edit to a parsed reference area.
    pub(crate) fn apply(self, area: &RefArea) -> RefEdit {
        let rows = self.is_row_axis();
        let get = |endpoint: &RefEndpoint| if rows { endpoint.row } else { endpoint.col };
        let set = |endpoint: &mut RefEndpoint, value: u32| {
            if rows {
                endpoint.row = Some(value);
            } else {
                endpoint.col = Some(value);
            }
        };
        let Some(end) = area.end else {
            return match get(&area.start) {
                None => RefEdit::Replace(*area),
                Some(value) => match self.shift_index(value) {
                    None => RefEdit::Delete,
                    Some(shifted) => {
                        let mut shifted_area = *area;
                        set(&mut shifted_area.start, shifted);
                        RefEdit::Replace(shifted_area)
                    }
                },
            };
        };
        match (get(&area.start), get(&end)) {
            (Some(first), Some(second)) => {
                let Some((lo, mut hi)) = self.shift_span(first.min(second), first.max(second))
                else {
                    return RefEdit::Delete;
                };
                // Whole-row/column spans stay pinned to the grid edge, as in
                // `1:1048576` or `A:XFD`.
                let whole_line = if rows {
                    area.start.col.is_none() && end.col.is_none()
                } else {
                    area.start.row.is_none() && end.row.is_none()
                };
                let max = if rows { MAX_ROWS } else { MAX_COLUMNS };
                if whole_line && lo <= max {
                    hi = hi.min(max);
                }
                let mut shifted_area = *area;
                let mut shifted_end = end;
                if first <= second {
                    set(&mut shifted_area.start, lo);
                    set(&mut shifted_end, hi);
                } else {
                    set(&mut shifted_area.start, hi);
                    set(&mut shifted_end, lo);
                }
                shifted_area.end = Some(shifted_end);
                RefEdit::Replace(shifted_area)
            }
            (None, None) => RefEdit::Replace(*area),
            _ => RefEdit::Keep,
        }
    }
}

#[derive(Debug, Clone)]
struct ParsedReference {
    ref_start: usize,
    end: usize,
    area: RefArea,
    qualified_sheet: Option<String>,
}

fn is_ref_identifier_char(ch: char) -> bool {
    ch.is_alphanumeric() || matches!(ch, '_' | '.')
}

fn is_ref_start_boundary(s: &str, index: usize) -> bool {
    s[..index]
        .chars()
        .next_back()
        .is_none_or(|ch| !is_ref_identifier_char(ch))
}

fn is_ref_end_boundary(s: &str, index: usize) -> bool {
    s[index..]
        .chars()
        .next()
        .is_none_or(|ch| !is_ref_identifier_char(ch))
}

fn is_unquoted_sheet_name_char(ch: char) -> bool {
    ch.is_alphanumeric() || matches!(ch, '_' | '.')
}

fn unquoted_sheet_name_end(s: &str, start: usize) -> usize {
    let mut end = start;
    for ch in s[start..].chars() {
        if !is_unquoted_sheet_name_char(ch) {
            break;
        }
        end += ch.len_utf8();
    }
    end
}

fn is_function_call_at(s: &str, index: usize) -> bool {
    let bytes = s.as_bytes();
    let mut next = index;
    while next < bytes.len() && bytes[next].is_ascii_whitespace() {
        next += 1;
    }
    next < bytes.len() && bytes[next] == b'('
}

fn parse_cell_ref_at(s: &str, start: usize) -> Option<ParsedCellRef> {
    let bytes = s.as_bytes();
    let len = bytes.len();
    if start >= len {
        return None;
    }

    if !is_ref_start_boundary(s, start) {
        return None;
    }

    let mut i = start;
    let abs_col = if bytes[i] == b'$' {
        i += 1;
        true
    } else {
        false
    };

    let col_start = i;
    while i < len && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    let col_len = i.saturating_sub(col_start);
    if !(1..=3).contains(&col_len) {
        return None;
    }

    let abs_row = if i < len && bytes[i] == b'$' {
        i += 1;
        true
    } else {
        false
    };

    let row_start = i;
    while i < len && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if row_start == i {
        return None;
    }

    // Ignore sheet names like "Sheet1!A1" when scanning at "Sheet1".
    if i < len && bytes[i] == b'!' {
        return None;
    }

    if !is_ref_end_boundary(s, i) {
        return None;
    }

    // A token such as LOG10 is a function name when it is immediately called,
    // not a cell reference.
    if is_function_call_at(s, i) {
        return None;
    }

    let col = column_name_to_number(&s[col_start..col_start + col_len]).ok()?;
    let row = s[row_start..i].parse::<u32>().ok()?;
    if !(1..=MAX_ROWS).contains(&row) {
        return None;
    }

    Some(ParsedCellRef {
        end: i,
        abs_col,
        abs_row,
        col,
        row,
    })
}

fn cell_endpoint(cell: &ParsedCellRef) -> RefEndpoint {
    RefEndpoint {
        col: Some(cell.col),
        row: Some(cell.row),
        abs_col: cell.abs_col,
        abs_row: cell.abs_row,
    }
}

/// Parse `$?LETTERS` (one to three letters) as a whole-column endpoint.
fn parse_col_endpoint_at(s: &str, start: usize) -> Option<(usize, RefEndpoint)> {
    let bytes = s.as_bytes();
    let mut i = start;
    let abs_col = bytes.get(i) == Some(&b'$');
    if abs_col {
        i += 1;
    }
    let col_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if !(1..=3).contains(&(i - col_start)) {
        return None;
    }
    let col = column_name_to_number(&s[col_start..i]).ok()?;
    Some((
        i,
        RefEndpoint {
            col: Some(col),
            row: None,
            abs_col,
            abs_row: false,
        },
    ))
}

/// Parse `$?DIGITS` as a whole-row endpoint.
fn parse_row_endpoint_at(s: &str, start: usize) -> Option<(usize, RefEndpoint)> {
    let bytes = s.as_bytes();
    let mut i = start;
    let abs_row = bytes.get(i) == Some(&b'$');
    if abs_row {
        i += 1;
    }
    let row_start = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if row_start == i {
        return None;
    }
    let row = s[row_start..i].parse::<u32>().ok()?;
    if !(1..=MAX_ROWS).contains(&row) {
        return None;
    }
    Some((
        i,
        RefEndpoint {
            col: None,
            row: Some(row),
            abs_col: false,
            abs_row,
        },
    ))
}

type EndpointParser = fn(&str, usize) -> Option<(usize, RefEndpoint)>;

/// Parse a whole-row (`1:3`) or whole-column (`A:C`) range at `start`.
fn parse_line_range_at(s: &str, start: usize) -> Option<(RefArea, usize)> {
    // A `$` directly before `start` belongs to a token that failed to parse.
    if s[..start].ends_with('$') {
        return None;
    }
    let parsers: [EndpointParser; 2] = [parse_col_endpoint_at, parse_row_endpoint_at];
    for parse in parsers {
        let Some((first_end, first)) = parse(s, start) else {
            continue;
        };
        if s.as_bytes().get(first_end) != Some(&b':') {
            continue;
        }
        let Some((end, second)) = parse(s, first_end + 1) else {
            continue;
        };
        if !is_ref_end_boundary(s, end)
            || s.as_bytes().get(end) == Some(&b'!')
            || is_function_call_at(s, end)
        {
            continue;
        }
        return Some((
            RefArea {
                start: first,
                end: Some(second),
            },
            end,
        ));
    }
    None
}

/// Parse a cell, cell range, or whole-row/column range at `start`.
fn parse_area_at(s: &str, start: usize) -> Option<(RefArea, usize)> {
    if start >= s.len() || !is_ref_start_boundary(s, start) {
        return None;
    }
    if let Some(first) = parse_cell_ref_at(s, start) {
        if s.as_bytes().get(first.end) == Some(&b':') {
            if let Some(second) = parse_cell_ref_at(s, first.end + 1) {
                return Some((
                    RefArea {
                        start: cell_endpoint(&first),
                        end: Some(cell_endpoint(&second)),
                    },
                    second.end,
                ));
            }
        }
        return Some((
            RefArea {
                start: cell_endpoint(&first),
                end: None,
            },
            first.end,
        ));
    }
    parse_line_range_at(s, start)
}

fn parse_quoted_sheet_name(s: &str, start: usize) -> Option<(String, usize)> {
    let bytes = s.as_bytes();
    if bytes.get(start) != Some(&b'\'') {
        return None;
    }

    let mut name = String::new();
    let mut i = start + 1;
    while i < bytes.len() {
        if bytes[i] != b'\'' {
            let ch = s[i..].chars().next()?;
            name.push(ch);
            i += ch.len_utf8();
            continue;
        }

        if bytes.get(i + 1) == Some(&b'\'') {
            name.push('\'');
            i += 2;
            continue;
        }

        return Some((name, i + 1));
    }

    None
}

fn parse_qualified_ref_at(s: &str, start: usize) -> Option<ParsedReference> {
    let bytes = s.as_bytes();
    if start >= bytes.len() || !is_ref_start_boundary(s, start) {
        return None;
    }

    let (sheet_name, ref_start) = if bytes[start] == b'\'' {
        let (sheet_name, after_quote) = parse_quoted_sheet_name(s, start)?;
        if bytes.get(after_quote) != Some(&b'!') {
            return None;
        }
        (sheet_name, after_quote + 1)
    } else {
        let sheet_end = unquoted_sheet_name_end(s, start);
        if sheet_end == start || bytes.get(sheet_end) != Some(&b'!') {
            return None;
        }
        (s[start..sheet_end].to_string(), sheet_end + 1)
    };

    let (area, end) = parse_area_at(s, ref_start)?;
    Some(ParsedReference {
        ref_start,
        end,
        area,
        qualified_sheet: Some(sheet_name),
    })
}

fn parse_external_qualified_ref_at(s: &str, start: usize) -> Option<ParsedReference> {
    let bytes = s.as_bytes();
    if bytes.get(start) != Some(&b'[') || !is_ref_start_boundary(s, start) {
        return None;
    }

    let book_end = start + 1 + s[start + 1..].find(']')?;
    let sheet_start = book_end + 1;
    let sheet_end = unquoted_sheet_name_end(s, sheet_start);
    if sheet_end == sheet_start || bytes.get(sheet_end) != Some(&b'!') {
        return None;
    }

    let ref_start = sheet_end + 1;
    let (area, end) = parse_area_at(s, ref_start)?;
    Some(ParsedReference {
        ref_start,
        end,
        area,
        qualified_sheet: Some(s[start..sheet_end].to_string()),
    })
}

fn parse_reference_at(s: &str, start: usize) -> Option<ParsedReference> {
    parse_external_qualified_ref_at(s, start)
        .or_else(|| parse_qualified_ref_at(s, start))
        .or_else(|| {
            parse_area_at(s, start).map(|(area, end)| ParsedReference {
                ref_start: start,
                end,
                area,
                qualified_sheet: None,
            })
        })
}

fn sheet_qualifier_end(text: &str, start: usize) -> Option<usize> {
    if text.as_bytes().get(start) == Some(&b'\'') {
        parse_quoted_sheet_name(text, start).map(|(_, end)| end)
    } else {
        let end = unquoted_sheet_name_end(text, start);
        (end > start).then_some(end)
    }
}

fn end_of_three_dimensional_reference(text: &str, start: usize) -> Option<usize> {
    if !is_ref_start_boundary(text, start) {
        return None;
    }

    if text.as_bytes().get(start) == Some(&b'\'') {
        let (sheet_name, after_quote) = parse_quoted_sheet_name(text, start)?;
        if sheet_name.contains(':') && text.as_bytes().get(after_quote) == Some(&b'!') {
            return parse_area_at(text, after_quote + 1).map(|(_, end)| end);
        }
    }

    let first_sheet_end = sheet_qualifier_end(text, start)?;
    if text.as_bytes().get(first_sheet_end) != Some(&b':') {
        return None;
    }
    let second_sheet_start = first_sheet_end + 1;
    let second_sheet_end = sheet_qualifier_end(text, second_sheet_start)?;
    if text.as_bytes().get(second_sheet_end) != Some(&b'!') {
        return None;
    }
    parse_area_at(text, second_sheet_end + 1).map(|(_, end)| end)
}

fn end_of_bracketed_reference(text: &str, start: usize) -> Option<usize> {
    if text.as_bytes().get(start) != Some(&b'[') {
        return None;
    }
    text[start + 1..].find(']').map(|offset| start + offset + 2)
}

fn end_of_string_literal(text: &str, start: usize) -> usize {
    let bytes = text.as_bytes();
    let mut i = start + 1;
    while i < bytes.len() {
        if bytes[i] != b'"' {
            let ch = text[i..]
                .chars()
                .next()
                .expect("index must remain on a UTF-8 boundary");
            i += ch.len_utf8();
            continue;
        }
        if bytes.get(i + 1) == Some(&b'"') {
            i += 2;
        } else {
            return i + 1;
        }
    }
    bytes.len()
}

fn format_shifted_ref(col: u32, row: u32, abs_col: bool, abs_row: bool) -> Result<String> {
    if !(1..=MAX_ROWS).contains(&row) {
        return Err(Error::InvalidRowNumber(row));
    }
    let col_name = column_number_to_name(col)?;
    Ok(format!(
        "{}{}{}{}",
        if abs_col { "$" } else { "" },
        col_name,
        if abs_row { "$" } else { "" },
        row
    ))
}

fn format_endpoint(endpoint: &RefEndpoint, out: &mut String) -> Result<()> {
    match (endpoint.col, endpoint.row) {
        (Some(col), Some(row)) => {
            out.push_str(&format_shifted_ref(
                col,
                row,
                endpoint.abs_col,
                endpoint.abs_row,
            )?);
        }
        (Some(col), None) => {
            let name = column_number_to_name(col)?;
            if endpoint.abs_col {
                out.push('$');
            }
            out.push_str(&name);
        }
        (None, Some(row)) => {
            if !(1..=MAX_ROWS).contains(&row) {
                return Err(Error::InvalidRowNumber(row));
            }
            if endpoint.abs_row {
                out.push('$');
            }
            out.push_str(&row.to_string());
        }
        (None, None) => {
            return Err(Error::Internal(
                "reference endpoint has no coordinate".to_string(),
            ))
        }
    }
    Ok(())
}

fn format_area(area: &RefArea) -> Result<String> {
    let mut out = String::new();
    format_endpoint(&area.start, &mut out)?;
    if let Some(end) = &area.end {
        out.push(':');
        format_endpoint(end, &mut out)?;
    }
    Ok(out)
}

fn checked_shifted_len_add(total: usize, value: usize) -> Result<usize> {
    total
        .checked_add(value)
        .ok_or_else(|| Error::Internal("shifted formula length overflow".to_string()))
}

trait RefSink {
    fn emit(&mut self, text: &str) -> Result<()>;
}

impl RefSink for String {
    fn emit(&mut self, text: &str) -> Result<()> {
        self.push_str(text);
        Ok(())
    }
}

struct LenSink(usize);

impl RefSink for LenSink {
    fn emit(&mut self, text: &str) -> Result<()> {
        self.0 = checked_shifted_len_add(self.0, text.len())?;
        Ok(())
    }
}

/// Scan formula text and rewrite every reference token through `edit`.
///
/// String literals, structured references, and 3-D references are copied
/// verbatim. `edit` receives the unquoted sheet qualifier, if any.
fn rewrite_references<S, C>(text: &str, sink: &mut S, mut edit: C) -> Result<()>
where
    S: RefSink,
    C: FnMut(Option<&str>, &RefArea) -> Result<RefEdit>,
{
    let mut i = 0usize;
    let bytes = text.as_bytes();
    let last_closing_bracket = text.rfind(']');

    while i < bytes.len() {
        if bytes[i] == b'"' {
            let end = end_of_string_literal(text, i);
            sink.emit(&text[i..end])?;
            i = end;
        } else if let Some(end) = end_of_three_dimensional_reference(text, i) {
            sink.emit(&text[i..end])?;
            i = end;
        } else if bytes[i] == b'[' && last_closing_bracket.is_none_or(|end| end < i) {
            sink.emit("[")?;
            i += 1;
        } else if let Some(parsed) = parse_reference_at(text, i) {
            sink.emit(&text[i..parsed.ref_start])?;
            match edit(parsed.qualified_sheet.as_deref(), &parsed.area)? {
                RefEdit::Keep => sink.emit(&text[parsed.ref_start..parsed.end])?,
                RefEdit::Replace(area) => sink.emit(&format_area(&area)?)?,
                RefEdit::Delete => sink.emit("#REF!")?,
            }
            i = parsed.end;
        } else if let Some(end) = end_of_bracketed_reference(text, i) {
            sink.emit(&text[i..end])?;
            i = end;
        } else {
            let ch = text[i..]
                .chars()
                .next()
                .expect("index must remain on a UTF-8 boundary");
            sink.emit(&text[i..i + ch.len_utf8()])?;
            i += ch.len_utf8();
        }
    }

    Ok(())
}

/// Adapt a per-cell callback to [`rewrite_references`]. Whole-row and
/// whole-column references have no cell coordinates and are kept verbatim.
fn cell_callback_edit<F, P>(
    shift_local: bool,
    shift_qualified_sheet: P,
    shift_cell: F,
) -> impl FnMut(Option<&str>, &RefArea) -> Result<RefEdit>
where
    F: Fn(u32, u32, bool, bool) -> (u32, u32) + Copy,
    P: Fn(&str) -> bool + Copy,
{
    move |sheet, area| {
        if !sheet.map_or(shift_local, shift_qualified_sheet) {
            return Ok(RefEdit::Keep);
        }
        let shift = |endpoint: RefEndpoint| -> Option<RefEndpoint> {
            let (col, row) = (endpoint.col?, endpoint.row?);
            let (col, row) = shift_cell(col, row, endpoint.abs_col, endpoint.abs_row);
            Some(RefEndpoint {
                col: Some(col),
                row: Some(row),
                ..endpoint
            })
        };
        let Some(start) = shift(area.start) else {
            return Ok(RefEdit::Keep);
        };
        let end = match area.end {
            Some(end) => match shift(end) {
                Some(end) => Some(end),
                None => return Ok(RefEdit::Keep),
            },
            None => None,
        };
        Ok(RefEdit::Replace(RefArea { start, end }))
    }
}

/// Shift cell references, respecting absolute markers.
///
/// The callback receives `(col, row, abs_col, abs_row)` and must return the
/// new `(col, row)` values.  This allows callers to skip shifting when a
/// reference is absolute.
#[cfg(test)]
pub(crate) fn shift_cell_references_with_abs<F>(text: &str, shift_cell: F) -> Result<String>
where
    F: Fn(u32, u32, bool, bool) -> (u32, u32) + Copy,
{
    shift_cell_references_with_abs_and_scope(text, true, |_| true, shift_cell)
}

/// Shifts cell references after the caller has admitted the exact output size.
pub(crate) fn shift_cell_references_with_abs_preallocated<F>(
    text: &str,
    output_len: usize,
    shift_cell: F,
) -> Result<String>
where
    F: Fn(u32, u32, bool, bool) -> (u32, u32) + Copy,
{
    let shifted = shift_cell_references_with_abs_and_scope_capacity(
        text,
        true,
        |_| true,
        output_len,
        shift_cell,
    )?;
    debug_assert_eq!(shifted.len(), output_len);
    Ok(shifted)
}

/// Returns the exact byte length of a shifted formula without allocating the
/// proportional output string.
pub(crate) fn shifted_cell_references_len_with_abs<F>(text: &str, shift_cell: F) -> Result<usize>
where
    F: Fn(u32, u32, bool, bool) -> (u32, u32) + Copy,
{
    let mut sink = LenSink(0);
    rewrite_references(
        text,
        &mut sink,
        cell_callback_edit(true, |_| true, shift_cell),
    )?;
    Ok(sink.0)
}

/// Shift references selected independently by their local or qualified scope.
///
/// Qualified sheet names are unquoted before they are passed to the predicate;
/// escaped apostrophes are restored to a single apostrophe.
pub(crate) fn shift_cell_references_with_abs_and_scope<F, P>(
    text: &str,
    shift_local: bool,
    shift_qualified_sheet: P,
    shift_cell: F,
) -> Result<String>
where
    F: Fn(u32, u32, bool, bool) -> (u32, u32) + Copy,
    P: Fn(&str) -> bool + Copy,
{
    shift_cell_references_with_abs_and_scope_capacity(
        text,
        shift_local,
        shift_qualified_sheet,
        text.len(),
        shift_cell,
    )
}

fn shift_cell_references_with_abs_and_scope_capacity<F, P>(
    text: &str,
    shift_local: bool,
    shift_qualified_sheet: P,
    output_len: usize,
    shift_cell: F,
) -> Result<String>
where
    F: Fn(u32, u32, bool, bool) -> (u32, u32) + Copy,
    P: Fn(&str) -> bool + Copy,
{
    let mut out = String::new();
    out.try_reserve_exact(output_len)
        .map_err(|_| Error::Internal("shifted formula allocation failed".to_string()))?;
    rewrite_references(
        text,
        &mut out,
        cell_callback_edit(shift_local, shift_qualified_sheet, shift_cell),
    )?;
    Ok(out)
}

/// Apply a structural row/column edit to every in-scope reference in a
/// formula. References to removed cells become `#REF!`; ranges that partially
/// overlap a removed band shrink. Whole-row and whole-column references shift
/// on their own axis only.
pub(crate) fn shift_references_structural<P>(
    text: &str,
    shift_local: bool,
    shift_qualified_sheet: P,
    edit: StructuralShift,
) -> Result<String>
where
    P: Fn(&str) -> bool + Copy,
{
    let mut out = String::new();
    out.try_reserve_exact(text.len())
        .map_err(|_| Error::Internal("shifted formula allocation failed".to_string()))?;
    rewrite_references(text, &mut out, |sheet, area| {
        if sheet.map_or(shift_local, shift_qualified_sheet) {
            Ok(edit.apply(area))
        } else {
            Ok(RefEdit::Keep)
        }
    })?;
    Ok(out)
}

/// Apply a structural edit to a space-separated sqref list.
///
/// Entries fully inside a removed band are dropped and partially covered
/// entries shrink. Returns `None` when every entry was removed. Tokens that
/// are not plain references are preserved verbatim.
pub(crate) fn shift_sqref_structural(sqref: &str, edit: StructuralShift) -> Result<Option<String>> {
    let mut kept = Vec::new();
    let mut saw_entry = false;
    for piece in sqref.split_whitespace() {
        saw_entry = true;
        match parse_area_at(piece, 0) {
            Some((area, end)) if end == piece.len() => match edit.apply(&area) {
                RefEdit::Delete => {}
                RefEdit::Keep => kept.push(piece.to_string()),
                RefEdit::Replace(area) => kept.push(format_area(&area)?),
            },
            _ => kept.push(piece.to_string()),
        }
    }
    if !saw_entry {
        return Ok(Some(sqref.to_string()));
    }
    if kept.is_empty() {
        return Ok(None);
    }
    Ok(Some(kept.join(" ")))
}

/// Returns true when `reference` is a single cell or a range whose two
/// endpoints address the same cell.
pub(crate) fn is_single_cell_reference(reference: &str) -> bool {
    match parse_area_at(reference, 0) {
        Some((area, end)) if end == reference.len() => match area.end {
            None => area.start.col.is_some() && area.start.row.is_some(),
            Some(last) => {
                area.start.col.is_some()
                    && area.start.row.is_some()
                    && area.start.col == last.col
                    && area.start.row == last.row
            }
        },
        _ => false,
    }
}

/// Rewrite the formula of a row copied by `duplicate_row` into `target_row`.
///
/// Relative rows move with the copy (`+1`). Absolute rows on the edited sheet
/// follow the structural row insertion at `target_row`; absolute rows on other
/// sheets stay unchanged.
pub(crate) fn shift_duplicated_row_references<P>(
    text: &str,
    is_target_sheet: P,
    target_row: u32,
) -> Result<String>
where
    P: Fn(&str) -> bool + Copy,
{
    let insertion = StructuralShift::InsertRows {
        start: target_row,
        count: 1,
    };
    let mut out = String::new();
    out.try_reserve_exact(text.len())
        .map_err(|_| Error::Internal("shifted formula allocation failed".to_string()))?;
    rewrite_references(text, &mut out, |sheet, area| {
        let on_target = sheet.is_none_or(is_target_sheet);
        let shift = |endpoint: RefEndpoint| -> RefEndpoint {
            let Some(row) = endpoint.row else {
                return endpoint;
            };
            let row = if !endpoint.abs_row {
                row.saturating_add(1)
            } else if on_target {
                insertion.shift_index(row).unwrap_or(row)
            } else {
                row
            };
            RefEndpoint {
                row: Some(row),
                ..endpoint
            }
        };
        Ok(RefEdit::Replace(RefArea {
            start: shift(area.start),
            end: area.end.map(shift),
        }))
    })?;
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_shift_cell_refs_basic() {
        let shifted =
            shift_cell_references_with_abs("SUM(A2:B2)", |col, row, _, _| (col, row + 1)).unwrap();
        assert_eq!(shifted, "SUM(A3:B3)");
    }

    #[test]
    fn test_shift_cell_refs_preserves_absolute() {
        let shifted =
            shift_cell_references_with_abs("$A$1:B2", |col, row, _, _| (col + 2, row)).unwrap();
        assert_eq!(shifted, "$C$1:D2");
    }

    #[test]
    fn test_shift_cell_refs_with_sheet_prefix() {
        let shifted = shift_cell_references_with_abs_and_scope(
            "Sheet1!A1+'Other Sheet'!B2+A1",
            true,
            |sheet| sheet == "Sheet1",
            |col, row, _, _| (col + 1, row + 1),
        )
        .unwrap();
        assert_eq!(shifted, "Sheet1!B2+'Other Sheet'!B2+B2");
    }

    #[test]
    fn test_shift_cell_refs_skips_string_literals_and_function_identifiers() {
        let shifted =
            shift_cell_references_with_abs("=\"A1\"&LOG10(A1)&SUM(B2)", |col, row, _, _| {
                (col + 1, row + 1)
            })
            .unwrap();
        assert_eq!(shifted, "=\"A1\"&LOG10(B2)&SUM(C3)");
    }

    #[test]
    fn test_shift_cell_refs_preserves_non_ascii_text() {
        let shifted = shift_cell_references_with_abs("=SUM(A1)+한글+B2", |col, row, _, _| {
            (col + 1, row + 1)
        })
        .unwrap();
        assert_eq!(shifted, "=SUM(B2)+한글+C3");
    }

    #[test]
    fn test_shift_cell_refs_parses_escaped_quoted_sheet_name() {
        let shifted = shift_cell_references_with_abs_and_scope(
            "'O''Brien'!$A$1+Sheet2!B2",
            true,
            |sheet| sheet == "O'Brien",
            |col, row, abs_col, abs_row| {
                (
                    if abs_col { col } else { col + 1 },
                    if abs_row { row } else { row + 1 },
                )
            },
        )
        .unwrap();
        assert_eq!(shifted, "'O''Brien'!$A$1+Sheet2!B2");
    }

    #[test]
    fn test_shift_cell_refs_selects_local_and_qualified_scopes_independently() {
        let qualified_only = shift_cell_references_with_abs_and_scope(
            "A1+Sheet1!B2",
            false,
            |sheet| sheet == "Sheet1",
            |col, row, _, _| (col + 1, row + 1),
        )
        .unwrap();
        assert_eq!(qualified_only, "A1+Sheet1!C3");

        let local_only = shift_cell_references_with_abs_and_scope(
            "A1+Sheet1!B2",
            true,
            |_| false,
            |col, row, _, _| (col + 1, row + 1),
        )
        .unwrap();
        assert_eq!(local_only, "B2+Sheet1!B2");
    }

    #[test]
    fn test_shift_cell_refs_skips_structured_and_three_dimensional_references() {
        let shifted = shift_cell_references_with_abs(
            "=SUM(Table1[A1])+Sheet1:Sheet3!A1+A1",
            |col, row, _, _| (col + 1, row + 1),
        )
        .unwrap();
        assert_eq!(shifted, "=SUM(Table1[A1])+Sheet1:Sheet3!A1+B2");
    }

    #[test]
    fn test_shift_cell_refs_skips_quoted_and_mixed_three_dimensional_references() {
        let shifted = shift_cell_references_with_abs(
            "'Sheet One':'Sheet Three'!A1+Sheet1:'Sheet Three'!B2+'Sheet One:Sheet Three'!C3+A1",
            |col, row, _, _| (col + 1, row + 1),
        )
        .unwrap();
        assert_eq!(
            shifted,
            "'Sheet One':'Sheet Three'!A1+Sheet1:'Sheet Three'!B2+'Sheet One:Sheet Three'!C3+B2"
        );
    }

    #[test]
    fn test_shift_cell_refs_treats_external_and_unicode_qualifiers_as_qualified() {
        let shifted = shift_cell_references_with_abs_and_scope(
            "[Book.xlsx]Sheet1!A1+한국어!B2+A1",
            false,
            |sheet| sheet == "한국어",
            |col, row, _, _| (col + 1, row + 1),
        )
        .unwrap();
        assert_eq!(shifted, "[Book.xlsx]Sheet1!A1+한국어!C3+A1");
    }

    #[test]
    fn test_shift_cell_refs_treats_quoted_external_reference_as_qualified() {
        let shifted = shift_cell_references_with_abs_and_scope(
            "'[Book.xlsx]Sheet1'!A1+A1",
            false,
            |sheet| sheet == "Sheet1",
            |col, row, _, _| (col + 1, row + 1),
        )
        .unwrap();
        assert_eq!(shifted, "'[Book.xlsx]Sheet1'!A1+A1");
    }

    #[test]
    fn test_shift_cell_refs_inherits_qualified_scope_for_range_endpoints() {
        let shifted = shift_cell_references_with_abs_and_scope(
            "Sheet1!$A$2:$A$3+'한국어'!B2:C2",
            false,
            |sheet| matches!(sheet, "Sheet1" | "한국어"),
            |col, row, _, _| (col, row + 1),
        )
        .unwrap();
        assert_eq!(shifted, "Sheet1!$A$3:$A$4+'한국어'!B3:C3");
    }

    #[test]
    fn test_shift_cell_refs_rejects_out_of_grid_shift_results() {
        let above_max = shift_cell_references_with_abs("A1048576", |col, row, _, _| (col, row + 1));
        assert!(matches!(
            above_max,
            Err(crate::error::Error::InvalidRowNumber(1_048_577))
        ));

        let below_min = shift_cell_references_with_abs("A1", |col, row, _, _| (col, row - 1));
        assert!(matches!(
            below_min,
            Err(crate::error::Error::InvalidRowNumber(0))
        ));
    }

    #[test]
    fn test_shift_cell_refs_rejects_function_with_whitespace_and_invalid_row() {
        let shifted = shift_cell_references_with_abs("LOG10 (A1)+A1048577+A1", |col, row, _, _| {
            (col + 1, row + 1)
        })
        .unwrap();
        assert_eq!(shifted, "LOG10 (B2)+A1048577+B2");
    }

    #[test]
    fn test_shift_cell_refs_respects_unicode_identifier_boundaries() {
        let shifted =
            shift_cell_references_with_abs("이름A1값+A1", |col, row, _, _| (col + 1, row + 1))
                .unwrap();
        assert_eq!(shifted, "이름A1값+B2");
    }

    #[test]
    fn shifted_length_matches_output_without_materializing_it() {
        let formula = "SUM(A1:B2)+'Other Sheet'!$C$3+\u{d55c}\u{ae00}+\u{d568}\u{c218}(A1)";
        let shift = |col, row, abs_col, abs_row| {
            (
                if abs_col { col } else { col + 25 },
                if abs_row { row } else { row + 9_999 },
            )
        };
        let expected = shift_cell_references_with_abs(formula, shift).unwrap();

        let output_len = shifted_cell_references_len_with_abs(formula, shift).unwrap();
        assert_eq!(output_len, expected.len());
        assert_eq!(
            shift_cell_references_with_abs_preallocated(formula, output_len, shift).unwrap(),
            expected
        );
    }

    #[test]
    fn shifted_length_rejects_out_of_grid_results() {
        let error =
            shifted_cell_references_len_with_abs("XFD1048576", |col, row, _, _| (col + 1, row))
                .unwrap_err();

        assert!(matches!(error, Error::InvalidColumnNumber(16_385)));
    }

    #[test]
    fn unmatched_bracket_prefixes_are_scanned_once_and_later_refs_still_shift() {
        let formula = format!("{}A1", "[".repeat(10_000));
        let shift = |col, row, _, _| (col + 1, row + 1);
        let output_len = shifted_cell_references_len_with_abs(&formula, shift).unwrap();
        let shifted =
            shift_cell_references_with_abs_preallocated(&formula, output_len, shift).unwrap();

        assert_eq!(output_len, formula.len());
        assert!(shifted.ends_with("B2"));
        assert_eq!(&shifted[..10_000], "[".repeat(10_000));
    }

    fn structural(text: &str, edit: StructuralShift) -> String {
        shift_references_structural(text, true, |_| true, edit).unwrap()
    }

    const REMOVE_ROW_5: StructuralShift = StructuralShift::RemoveRows { start: 5, count: 1 };
    const REMOVE_COL_B: StructuralShift = StructuralShift::RemoveCols { start: 2, count: 1 };

    #[test]
    fn structural_removal_invalidates_cells_and_shrinks_ranges() {
        assert_eq!(structural("A5", REMOVE_ROW_5), "#REF!");
        assert_eq!(structural("$A$5+A6", REMOVE_ROW_5), "#REF!+A5");
        assert_eq!(structural("SUM(A1:A5)", REMOVE_ROW_5), "SUM(A1:A4)");
        assert_eq!(structural("SUM(A5:A9)", REMOVE_ROW_5), "SUM(A5:A8)");
        assert_eq!(structural("SUM(A9:A5)", REMOVE_ROW_5), "SUM(A8:A5)");
        assert_eq!(structural("SUM(A5:B5)", REMOVE_ROW_5), "SUM(#REF!)");
        assert_eq!(
            structural("SUM(Sheet1!A5)", REMOVE_ROW_5),
            "SUM(Sheet1!#REF!)"
        );
        assert_eq!(structural("SUM(A1:C1)", REMOVE_COL_B), "SUM(A1:B1)");
        assert_eq!(structural("B1+C1", REMOVE_COL_B), "#REF!+B1");
    }

    #[test]
    fn structural_shift_handles_whole_rows_and_columns() {
        let insert_col_a = StructuralShift::InsertCols { start: 1, count: 1 };
        let insert_row_2 = StructuralShift::InsertRows { start: 2, count: 3 };
        assert_eq!(
            structural("SUM(A:A)+SUM($A:$C)", insert_col_a),
            "SUM(B:B)+SUM($B:$D)"
        );
        assert_eq!(
            structural("SUM(1:1)+SUM($2:3)", insert_col_a),
            "SUM(1:1)+SUM($2:3)"
        );
        assert_eq!(
            structural("SUM(1:1)+SUM($2:3)", insert_row_2),
            "SUM(1:1)+SUM($5:6)"
        );
        assert_eq!(structural("SUM(A:A)", insert_row_2), "SUM(A:A)");
        assert_eq!(
            structural("SUM(5:5)+SUM(4:6)", REMOVE_ROW_5),
            "SUM(#REF!)+SUM(4:5)"
        );
        assert_eq!(
            structural("SUM(B:B)+SUM(A:C)", REMOVE_COL_B),
            "SUM(#REF!)+SUM(A:B)"
        );
        assert_eq!(structural("SUM(2:1048576)", insert_row_2), "SUM(5:1048576)");
        assert_eq!(
            structural("'My Sheet'!$A:$A+[1]Data!B:B", insert_col_a),
            "'My Sheet'!$B:$B+[1]Data!C:C"
        );
    }

    #[test]
    fn whole_line_parsing_skips_strings_functions_and_non_references() {
        let insert_col_a = StructuralShift::InsertCols { start: 1, count: 1 };
        let insert_row_1 = StructuralShift::InsertRows { start: 1, count: 1 };
        assert_eq!(
            structural("LEN(\"AB:AC\")+LOG10(2)+Table1[[A]:[B]]", insert_col_a),
            "LEN(\"AB:AC\")+LOG10(2)+Table1[[A]:[B]]"
        );
        assert_eq!(
            structural("TIME(1,2,3)+1.5", insert_row_1),
            "TIME(1,2,3)+1.5"
        );
        assert_eq!(structural("ABCD:ABCE+0:1", insert_row_1), "ABCD:ABCE+0:1");
        assert_eq!(structural("Name1:Name2", insert_col_a), "Name1:Name2");
    }

    #[test]
    fn three_dimensional_ranges_are_skipped_whole() {
        let insert_row_1 = StructuralShift::InsertRows { start: 1, count: 1 };
        assert_eq!(
            structural("Sheet1:Sheet3!A1:B2+Sheet1:Sheet3!1:2+A1", insert_row_1),
            "Sheet1:Sheet3!A1:B2+Sheet1:Sheet3!1:2+A2"
        );
    }

    #[test]
    fn cell_callback_api_keeps_whole_line_references_verbatim() {
        let shifted = shift_cell_references_with_abs("SUM(A:A)+SUM(1:1)+B2", |col, row, _, _| {
            (col + 1, row + 1)
        })
        .unwrap();
        assert_eq!(shifted, "SUM(A:A)+SUM(1:1)+C3");
    }

    #[test]
    fn sqref_shift_drops_and_shrinks_entries() {
        assert_eq!(shift_sqref_structural("A5", REMOVE_ROW_5).unwrap(), None);
        assert_eq!(
            shift_sqref_structural("A5 B4:B6 C7", REMOVE_ROW_5).unwrap(),
            Some("B4:B5 C6".to_string())
        );
        assert_eq!(
            shift_sqref_structural("", REMOVE_ROW_5).unwrap(),
            Some(String::new())
        );
        assert!(is_single_cell_reference("C5:C5"));
        assert!(!is_single_cell_reference("C5:C6"));
    }

    #[test]
    fn duplicated_row_references_follow_copy_and_insertion() {
        let shifted = shift_duplicated_row_references(
            "A$10+A10+Sheet1!A$10+Other!A$10+A$2+A2+$5:$5+B:B",
            |sheet| sheet == "sheet1" || sheet == "Sheet1",
            6,
        )
        .unwrap();
        assert_eq!(shifted, "A$11+A11+Sheet1!A$11+Other!A$10+A$2+A3+$5:$5+B:B");
    }
}
