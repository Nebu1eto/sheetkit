use crate::error::{Error, Result};
use crate::utils::cell_ref::{column_name_to_number, column_number_to_name};
use crate::utils::constants::MAX_ROWS;

#[derive(Debug, Clone, Copy)]
struct ParsedCellRef {
    end: usize,
    abs_col: bool,
    abs_row: bool,
    col: u32,
    row: u32,
}

#[derive(Debug, Clone)]
struct ParsedReference {
    cell_start: usize,
    cell: ParsedCellRef,
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
    let mut next = i;
    while next < len && bytes[next].is_ascii_whitespace() {
        next += 1;
    }
    if next < len && bytes[next] == b'(' {
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

    let (sheet_name, cell_start) = if bytes[start] == b'\'' {
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

    let cell = parse_cell_ref_at(s, cell_start)?;
    Some(ParsedReference {
        cell_start,
        cell,
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

    let cell_start = sheet_end + 1;
    let cell = parse_cell_ref_at(s, cell_start)?;
    Some(ParsedReference {
        cell_start,
        cell,
        qualified_sheet: Some(s[start..sheet_end].to_string()),
    })
}

fn parse_reference_at(s: &str, start: usize) -> Option<ParsedReference> {
    parse_external_qualified_ref_at(s, start)
        .or_else(|| parse_qualified_ref_at(s, start))
        .or_else(|| {
            parse_cell_ref_at(s, start).map(|cell| ParsedReference {
                cell_start: start,
                cell,
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
            return parse_cell_ref_at(text, after_quote + 1).map(|cell| cell.end);
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
    parse_cell_ref_at(text, second_sheet_end + 1).map(|cell| cell.end)
}

fn end_of_bracketed_reference(text: &str, start: usize) -> Option<usize> {
    if text.as_bytes().get(start) != Some(&b'[') {
        return None;
    }
    text[start + 1..].find(']').map(|offset| start + offset + 2)
}

fn qualified_range_endpoint(text: &str, first_end: usize) -> Option<ParsedCellRef> {
    if text.as_bytes().get(first_end) != Some(&b':') {
        return None;
    }
    parse_cell_ref_at(text, first_end + 1)
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

/// Shift cell references, respecting absolute markers.
///
/// The callback receives `(col, row, abs_col, abs_row)` and must return the
/// new `(col, row)` values.  This allows callers to skip shifting when a
/// reference is absolute.
pub(crate) fn shift_cell_references_with_abs<F>(text: &str, shift_cell: F) -> Result<String>
where
    F: Fn(u32, u32, bool, bool) -> (u32, u32) + Copy,
{
    shift_cell_references_with_abs_and_scope(text, true, |_| true, shift_cell)
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
    let mut out = String::with_capacity(text.len());
    let mut i = 0usize;
    let bytes = text.as_bytes();

    while i < bytes.len() {
        if bytes[i] == b'"' {
            let end = end_of_string_literal(text, i);
            out.push_str(&text[i..end]);
            i = end;
        } else if let Some(end) = end_of_three_dimensional_reference(text, i) {
            out.push_str(&text[i..end]);
            i = end;
        } else if let Some(parsed) = parse_reference_at(text, i) {
            let range_endpoint = parsed
                .qualified_sheet
                .as_ref()
                .and_then(|_| qualified_range_endpoint(text, parsed.cell.end));
            let should_shift = parsed
                .qualified_sheet
                .as_deref()
                .map(&shift_qualified_sheet)
                .unwrap_or(shift_local);
            if !should_shift {
                let end = range_endpoint.map_or(parsed.cell.end, |cell| cell.end);
                out.push_str(&text[i..end]);
                i = end;
                continue;
            }

            out.push_str(&text[i..parsed.cell_start]);
            let (new_col, new_row) = shift_cell(
                parsed.cell.col,
                parsed.cell.row,
                parsed.cell.abs_col,
                parsed.cell.abs_row,
            );
            out.push_str(&format_shifted_ref(
                new_col,
                new_row,
                parsed.cell.abs_col,
                parsed.cell.abs_row,
            )?);
            if let Some(endpoint) = range_endpoint {
                out.push_str(&text[parsed.cell.end..parsed.cell.end + 1]);
                let (new_col, new_row) = shift_cell(
                    endpoint.col,
                    endpoint.row,
                    endpoint.abs_col,
                    endpoint.abs_row,
                );
                out.push_str(&format_shifted_ref(
                    new_col,
                    new_row,
                    endpoint.abs_col,
                    endpoint.abs_row,
                )?);
                i = endpoint.end;
            } else {
                i = parsed.cell.end;
            }
        } else if let Some(end) = end_of_bracketed_reference(text, i) {
            out.push_str(&text[i..end]);
            i = end;
        } else {
            let ch = text[i..]
                .chars()
                .next()
                .expect("index must remain on a UTF-8 boundary");
            out.push(ch);
            i += ch.len_utf8();
        }
    }

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
}
