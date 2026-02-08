use crate::error::Result;
use crate::utils::cell_ref::{column_name_to_number, column_number_to_name};

#[derive(Debug, Clone, Copy)]
struct ParsedCellRef {
    end: usize,
    abs_col: bool,
    abs_row: bool,
    col: u32,
    row: u32,
}

fn is_ref_boundary_byte(b: u8) -> bool {
    !(b.is_ascii_alphanumeric() || b == b'_' || b == b'.')
}

fn parse_cell_ref_at(s: &str, start: usize) -> Option<ParsedCellRef> {
    let bytes = s.as_bytes();
    let len = bytes.len();
    if start >= len {
        return None;
    }

    if start > 0 && !is_ref_boundary_byte(bytes[start - 1]) {
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

    if i < len && !is_ref_boundary_byte(bytes[i]) {
        return None;
    }

    let col = column_name_to_number(&s[col_start..col_start + col_len]).ok()?;
    let row = s[row_start..i].parse::<u32>().ok()?;
    if row == 0 {
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

fn format_shifted_ref(col: u32, row: u32, abs_col: bool, abs_row: bool) -> Result<String> {
    let col_name = column_number_to_name(col)?;
    Ok(format!(
        "{}{}{}{}",
        if abs_col { "$" } else { "" },
        col_name,
        if abs_row { "$" } else { "" },
        row
    ))
}

pub(crate) fn shift_cell_references_in_text<F>(text: &str, shift_cell: F) -> Result<String>
where
    F: Fn(u32, u32) -> (u32, u32) + Copy,
{
    if !text.is_ascii() {
        return Ok(text.to_string());
    }

    let mut out = String::with_capacity(text.len());
    let mut i = 0usize;
    let bytes = text.as_bytes();

    while i < bytes.len() {
        if let Some(parsed) = parse_cell_ref_at(text, i) {
            let (new_col, new_row) = shift_cell(parsed.col, parsed.row);
            out.push_str(&format_shifted_ref(
                new_col,
                new_row,
                parsed.abs_col,
                parsed.abs_row,
            )?);
            i = parsed.end;
        } else {
            out.push(bytes[i] as char);
            i += 1;
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
            shift_cell_references_in_text("SUM(A2:B2)", |col, row| (col, row + 1)).unwrap();
        assert_eq!(shifted, "SUM(A3:B3)");
    }

    #[test]
    fn test_shift_cell_refs_preserves_absolute() {
        let shifted = shift_cell_references_in_text("$A$1:B2", |col, row| (col + 2, row)).unwrap();
        assert_eq!(shifted, "$C$1:D2");
    }

    #[test]
    fn test_shift_cell_refs_with_sheet_prefix() {
        let shifted =
            shift_cell_references_in_text("Sheet1!A1+Sheet2!B2", |col, row| (col + 1, row + 1))
                .unwrap();
        assert_eq!(shifted, "Sheet1!B2+Sheet2!C3");
    }
}
