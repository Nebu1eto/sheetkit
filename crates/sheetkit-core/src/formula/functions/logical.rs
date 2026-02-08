//! Logical formula functions: TRUE, FALSE, IFERROR, IFNA, IFS, SWITCH, XOR.

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_bool, Evaluator};
use crate::formula::functions::check_arg_count;

/// TRUE() - returns boolean TRUE.
pub fn fn_true(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("TRUE", args, 0, 0)?;
    Ok(CellValue::Bool(true))
}

/// FALSE() - returns boolean FALSE.
pub fn fn_false(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("FALSE", args, 0, 0)?;
    Ok(CellValue::Bool(false))
}

/// IFERROR(value, value_if_error) - returns value_if_error if the first arg is an error.
pub fn fn_iferror(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IFERROR", args, 2, 2)?;
    match ctx.eval_expr(&args[0]) {
        Ok(CellValue::Error(_)) | Err(_) => ctx.eval_expr(&args[1]),
        Ok(v) => Ok(v),
    }
}

/// IFNA(value, value_if_na) - returns value_if_na if the first arg is #N/A.
pub fn fn_ifna(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IFNA", args, 2, 2)?;
    match ctx.eval_expr(&args[0]) {
        Ok(CellValue::Error(ref e)) if e == "#N/A" => ctx.eval_expr(&args[1]),
        Ok(v) => Ok(v),
        Err(e) => {
            let msg = e.to_string();
            if msg.contains("#N/A") {
                ctx.eval_expr(&args[1])
            } else {
                Err(e)
            }
        }
    }
}

/// IFS(condition1, value1, [condition2, value2], ...) - evaluates conditions in order.
pub fn fn_ifs(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    if args.len() < 2 || !args.len().is_multiple_of(2) {
        return Err(Error::WrongArgCount {
            name: "IFS".to_string(),
            expected: "even number >= 2".to_string(),
            got: args.len(),
        });
    }
    let mut i = 0;
    while i < args.len() {
        let cond = ctx.eval_expr(&args[i])?;
        if coerce_to_bool(&cond)? {
            return ctx.eval_expr(&args[i + 1]);
        }
        i += 2;
    }
    Ok(CellValue::Error("#N/A".to_string()))
}

/// SWITCH(expression, value1, result1, [value2, result2], ..., [default]) - matches a value.
pub fn fn_switch(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    if args.len() < 3 {
        return Err(Error::WrongArgCount {
            name: "SWITCH".to_string(),
            expected: "3 or more".to_string(),
            got: args.len(),
        });
    }
    let expr_val = ctx.eval_expr(&args[0])?;
    let pairs_end = if args.len().is_multiple_of(2) {
        // Even: has a default value at the end
        args.len() - 1
    } else {
        args.len()
    };
    let mut i = 1;
    while i + 1 < pairs_end {
        let case_val = ctx.eval_expr(&args[i])?;
        if values_equal(&expr_val, &case_val) {
            return ctx.eval_expr(&args[i + 1]);
        }
        i += 2;
    }
    // Check for default value
    if args.len().is_multiple_of(2) {
        ctx.eval_expr(&args[args.len() - 1])
    } else {
        Ok(CellValue::Error("#N/A".to_string()))
    }
}

fn values_equal(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Number(x), CellValue::Number(y)) => (x - y).abs() < f64::EPSILON,
        (CellValue::String(x), CellValue::String(y)) => x.eq_ignore_ascii_case(y),
        (CellValue::Bool(x), CellValue::Bool(y)) => x == y,
        (CellValue::Empty, CellValue::Empty) => true,
        _ => false,
    }
}

/// XOR(logical1, [logical2], ...) - returns TRUE if an odd number of arguments are TRUE.
pub fn fn_xor(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("XOR", args, 1, 255)?;
    let values = ctx.flatten_args_to_values(args)?;
    let mut true_count = 0u32;
    for v in &values {
        if matches!(v, CellValue::Empty) {
            continue;
        }
        if coerce_to_bool(v)? {
            true_count += 1;
        }
    }
    Ok(CellValue::Bool(
        !true_count.is_multiple_of(2) && true_count > 0,
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::eval::{evaluate, CellSnapshot, Evaluator};
    use crate::formula::parser::parse_formula;

    fn eval(formula: &str) -> CellValue {
        let snap = CellSnapshot::new("Sheet1".to_string());
        let expr = parse_formula(formula).unwrap();
        evaluate(&expr, &snap).unwrap()
    }

    #[test]
    fn test_true_function() {
        let snap = CellSnapshot::new("Sheet1".to_string());
        let mut evaluator = Evaluator::new(&snap);
        let result = fn_true(&[], &mut evaluator).unwrap();
        assert_eq!(result, CellValue::Bool(true));
    }

    #[test]
    fn test_false_function() {
        let snap = CellSnapshot::new("Sheet1".to_string());
        let mut evaluator = Evaluator::new(&snap);
        let result = fn_false(&[], &mut evaluator).unwrap();
        assert_eq!(result, CellValue::Bool(false));
    }

    #[test]
    fn test_iferror_no_error() {
        assert_eq!(eval(r#"IFERROR(42,"error")"#), CellValue::Number(42.0));
    }

    #[test]
    fn test_iferror_with_error() {
        assert_eq!(
            eval(r#"IFERROR(1/0,"error")"#),
            CellValue::String("error".to_string())
        );
    }

    #[test]
    fn test_ifna_no_na() {
        assert_eq!(eval(r#"IFNA(42,"not found")"#), CellValue::Number(42.0));
    }

    #[test]
    fn test_ifs_first_true() {
        assert_eq!(
            eval(r#"IFS(TRUE,"first",TRUE,"second")"#),
            CellValue::String("first".to_string())
        );
    }

    #[test]
    fn test_ifs_second_true() {
        assert_eq!(
            eval(r#"IFS(FALSE,"first",TRUE,"second")"#),
            CellValue::String("second".to_string())
        );
    }

    #[test]
    fn test_ifs_none_true() {
        assert_eq!(
            eval(r#"IFS(FALSE,"first",FALSE,"second")"#),
            CellValue::Error("#N/A".to_string())
        );
    }

    #[test]
    fn test_switch_match() {
        assert_eq!(
            eval(r#"SWITCH(2,1,"one",2,"two",3,"three")"#),
            CellValue::String("two".to_string())
        );
    }

    #[test]
    fn test_switch_default() {
        assert_eq!(
            eval(r#"SWITCH(99,1,"one",2,"two","other")"#),
            CellValue::String("other".to_string())
        );
    }

    #[test]
    fn test_switch_no_match_no_default() {
        assert_eq!(
            eval(r#"SWITCH(99,1,"one",2,"two",3,"three")"#),
            CellValue::Error("#N/A".to_string())
        );
    }

    #[test]
    fn test_xor_odd_true() {
        assert_eq!(eval("XOR(TRUE,FALSE,TRUE,TRUE)"), CellValue::Bool(true));
    }

    #[test]
    fn test_xor_even_true() {
        assert_eq!(eval("XOR(TRUE,TRUE)"), CellValue::Bool(false));
    }
}
