//! Math formula functions: SUMIF, SUMIFS, ROUNDUP, ROUNDDOWN, CEILING, FLOOR,
//! SIGN, RAND, RANDBETWEEN, PI, LOG, LOG10, LN, EXP, PRODUCT, QUOTIENT, FACT.

use crate::cell::CellValue;
use crate::error::Result;
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_number, coerce_to_string, Evaluator};
use crate::formula::functions::{check_arg_count, collect_criteria_range_values, matches_criteria};

/// SUMIF(range, criteria, [sum_range])
pub fn fn_sumif(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SUMIF", args, 2, 3)?;
    let range_vals = collect_criteria_range_values(&args[0], ctx)?;
    let criteria_val = ctx.eval_expr(&args[1])?;
    let criteria = coerce_to_string(&criteria_val);
    let sum_vals = if args.len() == 3 {
        collect_criteria_range_values(&args[2], ctx)?
    } else {
        range_vals.clone()
    };
    let mut total = 0.0;
    for (i, rv) in range_vals.iter().enumerate() {
        if matches_criteria(rv, &criteria) {
            if let Some(sv) = sum_vals.get(i) {
                if let Ok(n) = coerce_to_number(sv) {
                    total += n;
                }
            }
        }
    }
    Ok(CellValue::Number(total))
}

/// SUMIFS(sum_range, criteria_range1, criteria1, ...)
pub fn fn_sumifs(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SUMIFS", args, 3, 255)?;
    if !(args.len() - 1).is_multiple_of(2) {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let sum_vals = collect_criteria_range_values(&args[0], ctx)?;
    let pair_count = (args.len() - 1) / 2;
    let mut criteria_ranges: Vec<Vec<CellValue>> = Vec::with_capacity(pair_count);
    let mut criteria_strings: Vec<String> = Vec::with_capacity(pair_count);
    for i in 0..pair_count {
        let range_vals = collect_criteria_range_values(&args[1 + i * 2], ctx)?;
        let crit_val = ctx.eval_expr(&args[2 + i * 2])?;
        criteria_ranges.push(range_vals);
        criteria_strings.push(coerce_to_string(&crit_val));
    }
    let mut total = 0.0;
    for (idx, sv) in sum_vals.iter().enumerate() {
        let all_match =
            criteria_ranges
                .iter()
                .zip(criteria_strings.iter())
                .all(|(range_vals, crit)| {
                    range_vals
                        .get(idx)
                        .is_some_and(|rv| matches_criteria(rv, crit))
                });
        if all_match {
            if let Ok(n) = coerce_to_number(sv) {
                total += n;
            }
        }
    }
    Ok(CellValue::Number(total))
}

/// ROUNDUP(number, digits) - round away from zero
pub fn fn_roundup(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ROUNDUP", args, 2, 2)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let digits = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    let factor = 10f64.powi(digits);
    let result = if n >= 0.0 {
        (n * factor).ceil() / factor
    } else {
        (n * factor).floor() / factor
    };
    Ok(CellValue::Number(result))
}

/// ROUNDDOWN(number, digits) - round toward zero (truncate)
pub fn fn_rounddown(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ROUNDDOWN", args, 2, 2)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let digits = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    let factor = 10f64.powi(digits);
    let result = (n * factor).trunc() / factor;
    Ok(CellValue::Number(result))
}

/// CEILING(number, significance) - round up to nearest multiple of significance
pub fn fn_ceiling(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("CEILING", args, 2, 2)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let sig = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    if sig == 0.0 {
        return Ok(CellValue::Number(0.0));
    }
    if n > 0.0 && sig < 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let result = (n / sig).ceil() * sig;
    Ok(CellValue::Number(result))
}

/// FLOOR(number, significance) - round down to nearest multiple of significance
pub fn fn_floor(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("FLOOR", args, 2, 2)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let sig = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    if sig == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }
    if n > 0.0 && sig < 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let result = (n / sig).floor() * sig;
    Ok(CellValue::Number(result))
}

/// SIGN(number) - returns -1, 0, or 1
pub fn fn_sign(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SIGN", args, 1, 1)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let result = if n > 0.0 {
        1.0
    } else if n < 0.0 {
        -1.0
    } else {
        0.0
    };
    Ok(CellValue::Number(result))
}

/// RAND() - random number between 0 (inclusive) and 1 (exclusive)
pub fn fn_rand(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("RAND", args, 0, 0)?;
    // Simple deterministic-ish random using system time for non-crypto randomness.
    // In a real spreadsheet this would use a proper RNG.
    let t = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .subsec_nanos();
    let r = (t as f64 % 1_000_000.0) / 1_000_000.0;
    Ok(CellValue::Number(r))
}

/// RANDBETWEEN(bottom, top) - random integer in [bottom, top]
pub fn fn_randbetween(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("RANDBETWEEN", args, 2, 2)?;
    let bottom = coerce_to_number(&ctx.eval_expr(&args[0])?)?.ceil() as i64;
    let top = coerce_to_number(&ctx.eval_expr(&args[1])?)?.floor() as i64;
    if bottom > top {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let t = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .subsec_nanos() as i64;
    let range = top - bottom + 1;
    let result = bottom + (t.abs() % range);
    Ok(CellValue::Number(result as f64))
}

/// PI() - returns the value of pi
pub fn fn_pi(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("PI", args, 0, 0)?;
    Ok(CellValue::Number(std::f64::consts::PI))
}

/// LOG(number, [base]) - logarithm with optional base (default 10)
pub fn fn_log(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("LOG", args, 1, 2)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    if n <= 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let base = if args.len() > 1 {
        coerce_to_number(&ctx.eval_expr(&args[1])?)?
    } else {
        10.0
    };
    if base <= 0.0 || base == 1.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    Ok(CellValue::Number(n.log(base)))
}

/// LOG10(number) - base-10 logarithm
pub fn fn_log10(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("LOG10", args, 1, 1)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    if n <= 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    Ok(CellValue::Number(n.log10()))
}

/// LN(number) - natural logarithm
pub fn fn_ln(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("LN", args, 1, 1)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    if n <= 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    Ok(CellValue::Number(n.ln()))
}

/// EXP(number) - e raised to the power of number
pub fn fn_exp(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("EXP", args, 1, 1)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    Ok(CellValue::Number(n.exp()))
}

/// PRODUCT(args...) - product of all numbers
pub fn fn_product(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("PRODUCT", args, 1, 255)?;
    let nums = ctx.collect_numbers(args)?;
    if nums.is_empty() {
        return Ok(CellValue::Number(0.0));
    }
    let result: f64 = nums.iter().product();
    Ok(CellValue::Number(result))
}

/// QUOTIENT(numerator, denominator) - integer part of a division
pub fn fn_quotient(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("QUOTIENT", args, 2, 2)?;
    let num = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let den = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    if den == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }
    let result = (num / den).trunc();
    Ok(CellValue::Number(result))
}

/// FACT(number) - factorial of a non-negative integer
pub fn fn_fact(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("FACT", args, 1, 1)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    if n < 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let n_int = n.floor() as u64;
    let mut result: f64 = 1.0;
    for i in 2..=n_int {
        result *= i as f64;
    }
    Ok(CellValue::Number(result))
}

#[cfg(test)]
#[allow(clippy::manual_range_contains)]
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

    fn approx_eq(a: f64, b: f64) -> bool {
        (a - b).abs() < 1e-9
    }

    // SUMIF tests

    #[test]
    fn sumif_greater_than() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(1.0)),
            ("Sheet1", 1, 2, CellValue::Number(5.0)),
            ("Sheet1", 1, 3, CellValue::Number(10.0)),
        ];
        let result = eval_with_data("SUMIF(A1:A3,\">3\")", &data);
        assert_eq!(result, CellValue::Number(15.0));
    }

    #[test]
    fn sumif_less_than_or_equal() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(1.0)),
            ("Sheet1", 1, 2, CellValue::Number(5.0)),
            ("Sheet1", 1, 3, CellValue::Number(10.0)),
        ];
        let result = eval_with_data("SUMIF(A1:A3,\"<=5\")", &data);
        assert_eq!(result, CellValue::Number(6.0));
    }

    #[test]
    fn sumif_exact_text_match() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::String("Apple".to_string())),
            ("Sheet1", 2, 1, CellValue::Number(10.0)),
            ("Sheet1", 1, 2, CellValue::String("Banana".to_string())),
            ("Sheet1", 2, 2, CellValue::Number(20.0)),
            ("Sheet1", 1, 3, CellValue::String("Apple".to_string())),
            ("Sheet1", 2, 3, CellValue::Number(30.0)),
        ];
        let result = eval_with_data("SUMIF(A1:A3,\"Apple\",B1:B3)", &data);
        assert_eq!(result, CellValue::Number(40.0));
    }

    #[test]
    fn sumif_not_equal() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(0.0)),
            ("Sheet1", 1, 2, CellValue::Number(5.0)),
            ("Sheet1", 1, 3, CellValue::Number(0.0)),
        ];
        let result = eval_with_data("SUMIF(A1:A3,\"<>0\")", &data);
        assert_eq!(result, CellValue::Number(5.0));
    }

    #[test]
    fn sumif_no_sum_range() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(2.0)),
            ("Sheet1", 1, 2, CellValue::Number(4.0)),
            ("Sheet1", 1, 3, CellValue::Number(6.0)),
        ];
        let result = eval_with_data("SUMIF(A1:A3,\">3\")", &data);
        assert_eq!(result, CellValue::Number(10.0));
    }

    // SUMIFS tests

    #[test]
    fn sumifs_multi_criteria() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::String("A".to_string())),
            ("Sheet1", 2, 1, CellValue::Number(1.0)),
            ("Sheet1", 3, 1, CellValue::Number(10.0)),
            ("Sheet1", 1, 2, CellValue::String("B".to_string())),
            ("Sheet1", 2, 2, CellValue::Number(2.0)),
            ("Sheet1", 3, 2, CellValue::Number(20.0)),
            ("Sheet1", 1, 3, CellValue::String("A".to_string())),
            ("Sheet1", 2, 3, CellValue::Number(3.0)),
            ("Sheet1", 3, 3, CellValue::Number(30.0)),
        ];
        let result = eval_with_data("SUMIFS(C1:C3,A1:A3,\"A\",B1:B3,\">1\")", &data);
        assert_eq!(result, CellValue::Number(30.0));
    }

    // ROUNDUP tests

    #[test]
    fn roundup_positive() {
        let result = eval("ROUNDUP(3.2,0)");
        assert_eq!(result, CellValue::Number(4.0));
    }

    #[test]
    fn roundup_negative() {
        let result = eval("ROUNDUP(-3.2,0)");
        assert_eq!(result, CellValue::Number(-4.0));
    }

    #[test]
    fn roundup_with_digits() {
        let result = eval("ROUNDUP(3.14159,2)");
        assert_eq!(result, CellValue::Number(3.15));
    }

    // ROUNDDOWN tests

    #[test]
    fn rounddown_positive() {
        let result = eval("ROUNDDOWN(3.9,0)");
        assert_eq!(result, CellValue::Number(3.0));
    }

    #[test]
    fn rounddown_negative() {
        let result = eval("ROUNDDOWN(-3.9,0)");
        assert_eq!(result, CellValue::Number(-3.0));
    }

    // CEILING tests

    #[test]
    fn ceiling_basic() {
        let result = eval("CEILING(2.5,1)");
        assert_eq!(result, CellValue::Number(3.0));
    }

    #[test]
    fn ceiling_significance() {
        let result = eval("CEILING(4.42,0.05)");
        if let CellValue::Number(n) = result {
            assert!(approx_eq(n, 4.45));
        } else {
            panic!("expected number");
        }
    }

    // FLOOR tests

    #[test]
    fn floor_basic() {
        let result = eval("FLOOR(2.5,1)");
        assert_eq!(result, CellValue::Number(2.0));
    }

    #[test]
    fn floor_zero_significance() {
        let result = eval("FLOOR(2.5,0)");
        assert_eq!(result, CellValue::Error("#DIV/0!".to_string()));
    }

    // SIGN tests

    #[test]
    fn sign_positive() {
        assert_eq!(eval("SIGN(42)"), CellValue::Number(1.0));
    }

    #[test]
    fn sign_negative() {
        assert_eq!(eval("SIGN(-42)"), CellValue::Number(-1.0));
    }

    #[test]
    fn sign_zero() {
        assert_eq!(eval("SIGN(0)"), CellValue::Number(0.0));
    }

    // PI test

    #[test]
    fn pi_value() {
        if let CellValue::Number(n) = eval("PI()") {
            assert!(approx_eq(n, std::f64::consts::PI));
        } else {
            panic!("expected number");
        }
    }

    // LOG tests

    #[test]
    fn log_base10_default() {
        if let CellValue::Number(n) = eval("LOG(100)") {
            assert!(approx_eq(n, 2.0));
        } else {
            panic!("expected number");
        }
    }

    #[test]
    fn log_base2() {
        if let CellValue::Number(n) = eval("LOG(8,2)") {
            assert!(approx_eq(n, 3.0));
        } else {
            panic!("expected number");
        }
    }

    #[test]
    fn log_negative_input() {
        assert_eq!(eval("LOG(-1)"), CellValue::Error("#NUM!".to_string()));
    }

    // LOG10 test

    #[test]
    fn log10_basic() {
        if let CellValue::Number(n) = eval("LOG10(1000)") {
            assert!(approx_eq(n, 3.0));
        } else {
            panic!("expected number");
        }
    }

    // LN test

    #[test]
    fn ln_basic() {
        if let CellValue::Number(n) = eval("LN(1)") {
            assert!(approx_eq(n, 0.0));
        } else {
            panic!("expected number");
        }
    }

    // EXP test

    #[test]
    fn exp_basic() {
        if let CellValue::Number(n) = eval("EXP(0)") {
            assert!(approx_eq(n, 1.0));
        } else {
            panic!("expected number");
        }
    }

    #[test]
    fn exp_one() {
        if let CellValue::Number(n) = eval("EXP(1)") {
            assert!(approx_eq(n, std::f64::consts::E));
        } else {
            panic!("expected number");
        }
    }

    // PRODUCT test

    #[test]
    fn product_basic() {
        assert_eq!(eval("PRODUCT(2,3,4)"), CellValue::Number(24.0));
    }

    // QUOTIENT test

    #[test]
    fn quotient_basic() {
        assert_eq!(eval("QUOTIENT(7,2)"), CellValue::Number(3.0));
    }

    #[test]
    fn quotient_negative() {
        assert_eq!(eval("QUOTIENT(-7,2)"), CellValue::Number(-3.0));
    }

    #[test]
    fn quotient_div_zero() {
        assert_eq!(
            eval("QUOTIENT(7,0)"),
            CellValue::Error("#DIV/0!".to_string())
        );
    }

    // FACT test

    #[test]
    fn fact_basic() {
        assert_eq!(eval("FACT(5)"), CellValue::Number(120.0));
    }

    #[test]
    fn fact_zero() {
        assert_eq!(eval("FACT(0)"), CellValue::Number(1.0));
    }

    #[test]
    fn fact_negative() {
        assert_eq!(eval("FACT(-1)"), CellValue::Error("#NUM!".to_string()));
    }

    // RAND test (just verify it returns a number in [0, 1))

    #[test]
    fn rand_returns_number() {
        if let CellValue::Number(n) = eval("RAND()") {
            assert!(n >= 0.0 && n < 1.0);
        } else {
            panic!("expected number");
        }
    }

    // RANDBETWEEN test

    #[test]
    fn randbetween_returns_integer_in_range() {
        if let CellValue::Number(n) = eval("RANDBETWEEN(1,10)") {
            assert!(n >= 1.0 && n <= 10.0);
            assert_eq!(n, n.floor());
        } else {
            panic!("expected number");
        }
    }

    #[test]
    fn randbetween_invalid_range() {
        assert_eq!(
            eval("RANDBETWEEN(10,1)"),
            CellValue::Error("#NUM!".to_string())
        );
    }
}
