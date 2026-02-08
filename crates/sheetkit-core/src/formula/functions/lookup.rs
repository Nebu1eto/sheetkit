//! Lookup and reference formula functions: VLOOKUP, HLOOKUP, INDEX, MATCH,
//! LOOKUP, ROW, COLUMN, ROWS, COLUMNS, CHOOSE, ADDRESS.

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::formula::ast::{CellReference, Expr};
use crate::formula::eval::{coerce_to_number, coerce_to_string, compare_values, Evaluator};
use crate::formula::functions::check_arg_count;
use crate::utils::cell_ref::{column_name_to_number, column_number_to_name};

/// Extract (start, end) CellReference pair from a Range expression.
fn extract_range(expr: &Expr) -> Result<(&CellReference, &CellReference)> {
    match expr {
        Expr::Range { start, end } => Ok((start, end)),
        _ => Err(Error::FormulaError(
            "expected a range reference".to_string(),
        )),
    }
}

/// Read a range into a flat row-major Vec and return (values, num_cols, num_rows).
fn read_range(expr: &Expr, ctx: &mut Evaluator) -> Result<(Vec<CellValue>, usize, usize)> {
    let (start, end) = extract_range(expr)?;
    let start_col = column_name_to_number(&start.col)?;
    let end_col = column_name_to_number(&end.col)?;
    let min_col = start_col.min(end_col);
    let max_col = start_col.max(end_col);
    let min_row = start.row.min(end.row);
    let max_row = start.row.max(end.row);
    let num_cols = (max_col - min_col + 1) as usize;
    let num_rows = (max_row - min_row + 1) as usize;
    let values = ctx.expand_range(start, end)?;
    Ok((values, num_cols, num_rows))
}

/// VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
pub fn fn_vlookup(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("VLOOKUP", args, 3, 4)?;
    let lookup_val = ctx.eval_expr(&args[0])?;
    let (values, num_cols, num_rows) = read_range(&args[1], ctx)?;
    let col_index = coerce_to_number(&ctx.eval_expr(&args[2])?)? as usize;
    let range_lookup = if args.len() > 3 {
        match ctx.eval_expr(&args[3])? {
            CellValue::Bool(b) => b,
            v => coerce_to_number(&v)? != 0.0,
        }
    } else {
        true
    };

    if col_index < 1 || col_index > num_cols {
        return Ok(CellValue::Error("#REF!".to_string()));
    }

    if range_lookup {
        // Approximate match: find largest value <= lookup_val in first column
        let mut best_row: Option<usize> = None;
        for r in 0..num_rows {
            let cell = &values[r * num_cols];
            if compare_values(cell, &lookup_val) != std::cmp::Ordering::Greater {
                best_row = Some(r);
            } else {
                break;
            }
        }
        match best_row {
            Some(r) => Ok(values[r * num_cols + col_index - 1].clone()),
            None => Ok(CellValue::Error("#N/A".to_string())),
        }
    } else {
        // Exact match
        for r in 0..num_rows {
            let cell = &values[r * num_cols];
            if compare_values(cell, &lookup_val) == std::cmp::Ordering::Equal {
                return Ok(values[r * num_cols + col_index - 1].clone());
            }
        }
        Ok(CellValue::Error("#N/A".to_string()))
    }
}

/// HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
pub fn fn_hlookup(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("HLOOKUP", args, 3, 4)?;
    let lookup_val = ctx.eval_expr(&args[0])?;
    let (values, num_cols, num_rows) = read_range(&args[1], ctx)?;
    let row_index = coerce_to_number(&ctx.eval_expr(&args[2])?)? as usize;
    let range_lookup = if args.len() > 3 {
        match ctx.eval_expr(&args[3])? {
            CellValue::Bool(b) => b,
            v => coerce_to_number(&v)? != 0.0,
        }
    } else {
        true
    };

    if row_index < 1 || row_index > num_rows {
        return Ok(CellValue::Error("#REF!".to_string()));
    }

    if range_lookup {
        let mut best_col: Option<usize> = None;
        for (c, cell) in values.iter().enumerate().take(num_cols) {
            if compare_values(cell, &lookup_val) != std::cmp::Ordering::Greater {
                best_col = Some(c);
            } else {
                break;
            }
        }
        match best_col {
            Some(c) => Ok(values[(row_index - 1) * num_cols + c].clone()),
            None => Ok(CellValue::Error("#N/A".to_string())),
        }
    } else {
        for (c, cell) in values.iter().enumerate().take(num_cols) {
            if compare_values(cell, &lookup_val) == std::cmp::Ordering::Equal {
                return Ok(values[(row_index - 1) * num_cols + c].clone());
            }
        }
        Ok(CellValue::Error("#N/A".to_string()))
    }
}

/// INDEX(array, row_num, [col_num])
pub fn fn_index(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("INDEX", args, 2, 3)?;
    let (values, num_cols, num_rows) = read_range(&args[0], ctx)?;
    let row_num = coerce_to_number(&ctx.eval_expr(&args[1])?)? as usize;
    let col_num = if args.len() > 2 {
        coerce_to_number(&ctx.eval_expr(&args[2])?)? as usize
    } else {
        1
    };

    if row_num < 1 || row_num > num_rows || col_num < 1 || col_num > num_cols {
        return Ok(CellValue::Error("#REF!".to_string()));
    }
    Ok(values[(row_num - 1) * num_cols + (col_num - 1)].clone())
}

/// MATCH(lookup_value, lookup_array, [match_type])
pub fn fn_match(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("MATCH", args, 2, 3)?;
    let lookup_val = ctx.eval_expr(&args[0])?;
    let (values, num_cols, num_rows) = read_range(&args[1], ctx)?;
    let match_type = if args.len() > 2 {
        coerce_to_number(&ctx.eval_expr(&args[2])?)? as i32
    } else {
        1
    };

    // MATCH works on a 1-D array (single row or column)
    let items: &[CellValue] = if num_rows == 1 || num_cols == 1 {
        &values
    } else {
        return Ok(CellValue::Error("#N/A".to_string()));
    };

    match match_type {
        0 => {
            // Exact match
            for (i, v) in items.iter().enumerate() {
                if compare_values(v, &lookup_val) == std::cmp::Ordering::Equal {
                    return Ok(CellValue::Number((i + 1) as f64));
                }
            }
            Ok(CellValue::Error("#N/A".to_string()))
        }
        1 => {
            // Largest value <= lookup_val (data sorted ascending)
            let mut best: Option<usize> = None;
            for (i, v) in items.iter().enumerate() {
                if compare_values(v, &lookup_val) != std::cmp::Ordering::Greater {
                    best = Some(i);
                }
            }
            match best {
                Some(i) => Ok(CellValue::Number((i + 1) as f64)),
                None => Ok(CellValue::Error("#N/A".to_string())),
            }
        }
        -1 => {
            // Smallest value >= lookup_val (data sorted descending)
            let mut best: Option<usize> = None;
            for (i, v) in items.iter().enumerate() {
                if compare_values(v, &lookup_val) != std::cmp::Ordering::Less {
                    best = Some(i);
                }
            }
            match best {
                Some(i) => Ok(CellValue::Number((i + 1) as f64)),
                None => Ok(CellValue::Error("#N/A".to_string())),
            }
        }
        _ => Ok(CellValue::Error("#N/A".to_string())),
    }
}

/// LOOKUP(lookup_value, lookup_vector, [result_vector]) - vector form.
pub fn fn_lookup(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("LOOKUP", args, 2, 3)?;
    let lookup_val = ctx.eval_expr(&args[0])?;
    let (lookup_values, _, _) = read_range(&args[1], ctx)?;

    // Find largest value <= lookup_val (assumes sorted ascending)
    let mut best: Option<usize> = None;
    for (i, v) in lookup_values.iter().enumerate() {
        if compare_values(v, &lookup_val) != std::cmp::Ordering::Greater {
            best = Some(i);
        }
    }
    let idx = match best {
        Some(i) => i,
        None => return Ok(CellValue::Error("#N/A".to_string())),
    };

    if args.len() > 2 {
        let (result_values, _, _) = read_range(&args[2], ctx)?;
        if idx < result_values.len() {
            Ok(result_values[idx].clone())
        } else {
            Ok(CellValue::Error("#N/A".to_string()))
        }
    } else {
        Ok(lookup_values[idx].clone())
    }
}

/// ROW([reference]) - returns the row number of a reference.
pub fn fn_row(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ROW", args, 0, 1)?;
    if args.is_empty() {
        return Ok(CellValue::Number(1.0));
    }
    match &args[0] {
        Expr::CellRef(cell_ref) => Ok(CellValue::Number(cell_ref.row as f64)),
        Expr::Range { start, .. } => Ok(CellValue::Number(start.row as f64)),
        _ => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

/// COLUMN([reference]) - returns the column number of a reference.
pub fn fn_column(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("COLUMN", args, 0, 1)?;
    if args.is_empty() {
        return Ok(CellValue::Number(1.0));
    }
    match &args[0] {
        Expr::CellRef(cell_ref) => {
            let col = column_name_to_number(&cell_ref.col)?;
            Ok(CellValue::Number(col as f64))
        }
        Expr::Range { start, .. } => {
            let col = column_name_to_number(&start.col)?;
            Ok(CellValue::Number(col as f64))
        }
        _ => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

/// ROWS(array) - returns the number of rows in a reference.
pub fn fn_rows(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ROWS", args, 1, 1)?;
    match &args[0] {
        Expr::Range { start, end } => {
            let rows = (end.row.max(start.row) - end.row.min(start.row) + 1) as f64;
            Ok(CellValue::Number(rows))
        }
        Expr::CellRef(_) => Ok(CellValue::Number(1.0)),
        _ => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

/// COLUMNS(array) - returns the number of columns in a reference.
pub fn fn_columns(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("COLUMNS", args, 1, 1)?;
    match &args[0] {
        Expr::Range { start, end } => {
            let start_col = column_name_to_number(&start.col)?;
            let end_col = column_name_to_number(&end.col)?;
            let cols = (start_col.max(end_col) - start_col.min(end_col) + 1) as f64;
            Ok(CellValue::Number(cols))
        }
        Expr::CellRef(_) => Ok(CellValue::Number(1.0)),
        _ => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

/// CHOOSE(index_num, value1, [value2], ...) - returns value at given index.
pub fn fn_choose(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("CHOOSE", args, 2, 255)?;
    let index = coerce_to_number(&ctx.eval_expr(&args[0])?)? as usize;
    if index < 1 || index >= args.len() {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    ctx.eval_expr(&args[index])
}

/// ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])
pub fn fn_address(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ADDRESS", args, 2, 5)?;
    let row = coerce_to_number(&ctx.eval_expr(&args[0])?)? as u32;
    let col = coerce_to_number(&ctx.eval_expr(&args[1])?)? as u32;
    let abs_num = if args.len() > 2 {
        coerce_to_number(&ctx.eval_expr(&args[2])?)? as i32
    } else {
        1
    };
    // a1 parameter (default true = A1 style, false = R1C1 style)
    let a1 = if args.len() > 3 {
        match ctx.eval_expr(&args[3])? {
            CellValue::Bool(b) => b,
            v => coerce_to_number(&v)? != 0.0,
        }
    } else {
        true
    };
    let sheet_text = if args.len() > 4 {
        Some(coerce_to_string(&ctx.eval_expr(&args[4])?))
    } else {
        None
    };

    if col < 1 || row < 1 {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }

    let col_name = column_number_to_name(col)?;

    let address = if a1 {
        match abs_num {
            1 => format!("${col_name}${row}"), // absolute row and column
            2 => format!("{col_name}${row}"),  // absolute row
            3 => format!("${col_name}{row}"),  // absolute column
            4 => format!("{col_name}{row}"),   // relative
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        // R1C1 style
        match abs_num {
            1 => format!("R{row}C{col}"),
            2 => format!("R{row}C[{col}]"),
            3 => format!("R[{row}]C{col}"),
            4 => format!("R[{row}]C[{col}]"),
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    };

    if let Some(sheet) = sheet_text {
        Ok(CellValue::String(format!("{sheet}!{address}")))
    } else {
        Ok(CellValue::String(address))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::eval::{evaluate, CellSnapshot};
    use crate::formula::parser::parse_formula;

    fn eval(formula: &str) -> CellValue {
        let snap = CellSnapshot::new("Sheet1".to_string());
        let expr = parse_formula(formula).unwrap();
        evaluate(&expr, &snap).unwrap()
    }

    fn eval_with_data(formula: &str, snap: &CellSnapshot) -> CellValue {
        let expr = parse_formula(formula).unwrap();
        evaluate(&expr, snap).unwrap()
    }

    #[test]
    fn test_vlookup_exact() {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        // Column A: keys, Column B: values
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(1.0));
        snap.set_cell("Sheet1", 2, 1, CellValue::String("one".to_string()));
        snap.set_cell("Sheet1", 1, 2, CellValue::Number(2.0));
        snap.set_cell("Sheet1", 2, 2, CellValue::String("two".to_string()));
        snap.set_cell("Sheet1", 1, 3, CellValue::Number(3.0));
        snap.set_cell("Sheet1", 2, 3, CellValue::String("three".to_string()));

        assert_eq!(
            eval_with_data("VLOOKUP(2,A1:B3,2,FALSE)", &snap),
            CellValue::String("two".to_string())
        );
    }

    #[test]
    fn test_vlookup_not_found() {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(1.0));
        snap.set_cell("Sheet1", 2, 1, CellValue::String("one".to_string()));

        assert_eq!(
            eval_with_data("VLOOKUP(99,A1:B1,2,FALSE)", &snap),
            CellValue::Error("#N/A".to_string())
        );
    }

    #[test]
    fn test_vlookup_approximate() {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(10.0));
        snap.set_cell("Sheet1", 2, 1, CellValue::String("ten".to_string()));
        snap.set_cell("Sheet1", 1, 2, CellValue::Number(20.0));
        snap.set_cell("Sheet1", 2, 2, CellValue::String("twenty".to_string()));
        snap.set_cell("Sheet1", 1, 3, CellValue::Number(30.0));
        snap.set_cell("Sheet1", 2, 3, CellValue::String("thirty".to_string()));

        assert_eq!(
            eval_with_data("VLOOKUP(25,A1:B3,2,TRUE)", &snap),
            CellValue::String("twenty".to_string())
        );
    }

    #[test]
    fn test_hlookup_exact() {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        // Row 1: keys, Row 2: values
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(1.0));
        snap.set_cell("Sheet1", 2, 1, CellValue::Number(2.0));
        snap.set_cell("Sheet1", 3, 1, CellValue::Number(3.0));
        snap.set_cell("Sheet1", 1, 2, CellValue::String("one".to_string()));
        snap.set_cell("Sheet1", 2, 2, CellValue::String("two".to_string()));
        snap.set_cell("Sheet1", 3, 2, CellValue::String("three".to_string()));

        assert_eq!(
            eval_with_data("HLOOKUP(2,A1:C2,2,FALSE)", &snap),
            CellValue::String("two".to_string())
        );
    }

    #[test]
    fn test_index_basic() {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(1.0));
        snap.set_cell("Sheet1", 2, 1, CellValue::Number(2.0));
        snap.set_cell("Sheet1", 1, 2, CellValue::Number(3.0));
        snap.set_cell("Sheet1", 2, 2, CellValue::Number(4.0));

        assert_eq!(
            eval_with_data("INDEX(A1:B2,2,2)", &snap),
            CellValue::Number(4.0)
        );
    }

    #[test]
    fn test_match_exact() {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        snap.set_cell("Sheet1", 1, 1, CellValue::String("apple".to_string()));
        snap.set_cell("Sheet1", 1, 2, CellValue::String("banana".to_string()));
        snap.set_cell("Sheet1", 1, 3, CellValue::String("cherry".to_string()));

        assert_eq!(
            eval_with_data(r#"MATCH("banana",A1:A3,0)"#, &snap),
            CellValue::Number(2.0)
        );
    }

    #[test]
    fn test_match_not_found() {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(1.0));
        snap.set_cell("Sheet1", 1, 2, CellValue::Number(2.0));

        assert_eq!(
            eval_with_data("MATCH(99,A1:A2,0)", &snap),
            CellValue::Error("#N/A".to_string())
        );
    }

    #[test]
    fn test_lookup_vector() {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(1.0));
        snap.set_cell("Sheet1", 1, 2, CellValue::Number(2.0));
        snap.set_cell("Sheet1", 1, 3, CellValue::Number(3.0));
        snap.set_cell("Sheet1", 2, 1, CellValue::String("one".to_string()));
        snap.set_cell("Sheet1", 2, 2, CellValue::String("two".to_string()));
        snap.set_cell("Sheet1", 2, 3, CellValue::String("three".to_string()));

        assert_eq!(
            eval_with_data("LOOKUP(2,A1:A3,B1:B3)", &snap),
            CellValue::String("two".to_string())
        );
    }

    #[test]
    fn test_row() {
        assert_eq!(eval("ROW(B5)"), CellValue::Number(5.0));
    }

    #[test]
    fn test_column() {
        assert_eq!(eval("COLUMN(C1)"), CellValue::Number(3.0));
    }

    #[test]
    fn test_rows() {
        assert_eq!(eval("ROWS(A1:A10)"), CellValue::Number(10.0));
    }

    #[test]
    fn test_columns() {
        assert_eq!(eval("COLUMNS(A1:D1)"), CellValue::Number(4.0));
    }

    #[test]
    fn test_choose() {
        assert_eq!(
            eval(r#"CHOOSE(2,"a","b","c")"#),
            CellValue::String("b".to_string())
        );
    }

    #[test]
    fn test_choose_out_of_range() {
        assert_eq!(
            eval(r#"CHOOSE(5,"a","b","c")"#),
            CellValue::Error("#VALUE!".to_string())
        );
    }

    #[test]
    fn test_address_absolute() {
        assert_eq!(eval("ADDRESS(1,1)"), CellValue::String("$A$1".to_string()));
    }

    #[test]
    fn test_address_relative() {
        assert_eq!(eval("ADDRESS(1,1,4)"), CellValue::String("A1".to_string()));
    }

    #[test]
    fn test_address_with_sheet() {
        assert_eq!(
            eval(r#"ADDRESS(1,1,1,TRUE,"Sheet2")"#),
            CellValue::String("Sheet2!$A$1".to_string())
        );
    }

    #[test]
    fn test_row_no_args() {
        assert_eq!(eval("ROW()"), CellValue::Number(1.0));
    }

    #[test]
    fn test_column_no_args() {
        assert_eq!(eval("COLUMN()"), CellValue::Number(1.0));
    }
}
