//! Information formula functions: ISERR, ISNA, ISLOGICAL, ISEVEN, ISODD,
//! TYPE, N, NA, ERROR.TYPE.

use crate::cell::CellValue;
use crate::error::Result;
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_number, Evaluator};
use crate::formula::functions::check_arg_count;

/// ISERR(value) - TRUE if value is an error other than #N/A
pub fn fn_iserr(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ISERR", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let result = match &v {
        CellValue::Error(e) => e != "#N/A",
        _ => false,
    };
    Ok(CellValue::Bool(result))
}

/// ISNA(value) - TRUE if value is #N/A
pub fn fn_isna(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ISNA", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let result = matches!(&v, CellValue::Error(e) if e == "#N/A");
    Ok(CellValue::Bool(result))
}

/// ISLOGICAL(value) - TRUE if value is a boolean
pub fn fn_islogical(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ISLOGICAL", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    Ok(CellValue::Bool(matches!(v, CellValue::Bool(_))))
}

/// ISEVEN(number) - TRUE if the integer part of number is even
pub fn fn_iseven(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ISEVEN", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let n = coerce_to_number(&v)?;
    let int_part = n.trunc() as i64;
    Ok(CellValue::Bool(int_part % 2 == 0))
}

/// ISODD(number) - TRUE if the integer part of number is odd
pub fn fn_isodd(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ISODD", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let n = coerce_to_number(&v)?;
    let int_part = n.trunc() as i64;
    Ok(CellValue::Bool(int_part % 2 != 0))
}

/// TYPE(value) - returns type code: 1=number, 2=text, 4=boolean, 16=error, 64=array
pub fn fn_type(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("TYPE", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let code = match v {
        CellValue::Number(_) | CellValue::Date(_) | CellValue::Empty => 1.0,
        CellValue::String(_) | CellValue::RichString(_) => 2.0,
        CellValue::Bool(_) => 4.0,
        CellValue::Error(_) => 16.0,
        CellValue::Formula { .. } => 1.0,
    };
    Ok(CellValue::Number(code))
}

/// N(value) - convert to number: numbers stay, bools become 0/1, errors stay, everything else 0
pub fn fn_n(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("N", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    match v {
        CellValue::Number(n) => Ok(CellValue::Number(n)),
        CellValue::Date(n) => Ok(CellValue::Number(n)),
        CellValue::Bool(b) => Ok(CellValue::Number(if b { 1.0 } else { 0.0 })),
        CellValue::Error(_) => Ok(v),
        _ => Ok(CellValue::Number(0.0)),
    }
}

/// NA() - returns #N/A error
pub fn fn_na(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("NA", args, 0, 0)?;
    Ok(CellValue::Error("#N/A".to_string()))
}

/// ERROR.TYPE(value) - returns error type code.
/// 1=#NULL!, 2=#DIV/0!, 3=#VALUE!, 4=#REF!, 5=#NAME?, 6=#NUM!, 7=#N/A.
/// Returns #N/A if value is not an error.
pub fn fn_error_type(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ERROR.TYPE", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    match v {
        CellValue::Error(ref e) => {
            let code = match e.as_str() {
                "#NULL!" => 1.0,
                "#DIV/0!" => 2.0,
                "#VALUE!" => 3.0,
                "#REF!" => 4.0,
                "#NAME?" => 5.0,
                "#NUM!" => 6.0,
                "#N/A" => 7.0,
                _ => return Ok(CellValue::Error("#N/A".to_string())),
            };
            Ok(CellValue::Number(code))
        }
        _ => Ok(CellValue::Error("#N/A".to_string())),
    }
}

#[cfg(test)]
mod tests {
    use crate::cell::CellValue;
    use crate::formula::eval::{evaluate, CellSnapshot};
    use crate::formula::parser::parse_formula;

    fn eval_with_data(formula: &str, data: &[(&str, u32, u32, CellValue)]) -> CellValue {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        for (sheet, col, row, val) in data {
            snap.set_cell(sheet, *col, *row, val.clone());
        }
        let expr = parse_formula(formula).unwrap();
        evaluate(&expr, &snap).unwrap()
    }

    fn eval(formula: &str) -> CellValue {
        eval_with_data(formula, &[])
    }

    // ISERR tests

    #[test]
    fn iserr_div_zero() {
        let data = vec![("Sheet1", 1, 1, CellValue::Error("#DIV/0!".to_string()))];
        assert_eq!(eval_with_data("ISERR(A1)", &data), CellValue::Bool(true));
    }

    #[test]
    fn iserr_na_is_false() {
        let data = vec![("Sheet1", 1, 1, CellValue::Error("#N/A".to_string()))];
        assert_eq!(eval_with_data("ISERR(A1)", &data), CellValue::Bool(false));
    }

    #[test]
    fn iserr_not_error() {
        assert_eq!(eval("ISERR(42)"), CellValue::Bool(false));
    }

    // ISNA tests

    #[test]
    fn isna_true() {
        assert_eq!(eval("ISNA(NA())"), CellValue::Bool(true));
    }

    #[test]
    fn isna_false_for_other_error() {
        let data = vec![("Sheet1", 1, 1, CellValue::Error("#DIV/0!".to_string()))];
        assert_eq!(eval_with_data("ISNA(A1)", &data), CellValue::Bool(false));
    }

    #[test]
    fn isna_false_for_number() {
        assert_eq!(eval("ISNA(1)"), CellValue::Bool(false));
    }

    // ISLOGICAL tests

    #[test]
    fn islogical_true() {
        assert_eq!(eval("ISLOGICAL(TRUE)"), CellValue::Bool(true));
    }

    #[test]
    fn islogical_false_for_number() {
        assert_eq!(eval("ISLOGICAL(1)"), CellValue::Bool(false));
    }

    // ISEVEN tests

    #[test]
    fn iseven_true() {
        assert_eq!(eval("ISEVEN(4)"), CellValue::Bool(true));
    }

    #[test]
    fn iseven_false() {
        assert_eq!(eval("ISEVEN(3)"), CellValue::Bool(false));
    }

    #[test]
    fn iseven_zero() {
        assert_eq!(eval("ISEVEN(0)"), CellValue::Bool(true));
    }

    // ISODD tests

    #[test]
    fn isodd_true() {
        assert_eq!(eval("ISODD(3)"), CellValue::Bool(true));
    }

    #[test]
    fn isodd_false() {
        assert_eq!(eval("ISODD(4)"), CellValue::Bool(false));
    }

    // TYPE tests

    #[test]
    fn type_number() {
        assert_eq!(eval("TYPE(42)"), CellValue::Number(1.0));
    }

    #[test]
    fn type_text() {
        assert_eq!(eval("TYPE(\"hello\")"), CellValue::Number(2.0));
    }

    #[test]
    fn type_boolean() {
        assert_eq!(eval("TYPE(TRUE)"), CellValue::Number(4.0));
    }

    #[test]
    fn type_error() {
        assert_eq!(eval("TYPE(NA())"), CellValue::Number(16.0));
    }

    // N tests

    #[test]
    fn n_number() {
        assert_eq!(eval("N(42)"), CellValue::Number(42.0));
    }

    #[test]
    fn n_true() {
        assert_eq!(eval("N(TRUE)"), CellValue::Number(1.0));
    }

    #[test]
    fn n_false() {
        assert_eq!(eval("N(FALSE)"), CellValue::Number(0.0));
    }

    #[test]
    fn n_string() {
        assert_eq!(eval("N(\"text\")"), CellValue::Number(0.0));
    }

    #[test]
    fn n_error() {
        assert_eq!(eval("N(NA())"), CellValue::Error("#N/A".to_string()));
    }

    // NA test

    #[test]
    fn na_returns_error() {
        assert_eq!(eval("NA()"), CellValue::Error("#N/A".to_string()));
    }

    // ERROR.TYPE tests

    #[test]
    fn error_type_div_zero() {
        let data = vec![("Sheet1", 1, 1, CellValue::Error("#DIV/0!".to_string()))];
        assert_eq!(
            eval_with_data("ERROR.TYPE(A1)", &data),
            CellValue::Number(2.0)
        );
    }

    #[test]
    fn error_type_na() {
        assert_eq!(eval("ERROR.TYPE(NA())"), CellValue::Number(7.0));
    }

    #[test]
    fn error_type_value_error() {
        let data = vec![("Sheet1", 1, 1, CellValue::Error("#VALUE!".to_string()))];
        assert_eq!(
            eval_with_data("ERROR.TYPE(A1)", &data),
            CellValue::Number(3.0)
        );
    }

    #[test]
    fn error_type_not_error() {
        assert_eq!(eval("ERROR.TYPE(42)"), CellValue::Error("#N/A".to_string()));
    }
}
