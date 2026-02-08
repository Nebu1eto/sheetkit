//! Built-in Excel function dispatch.
//!
//! Provides [`lookup_function`] to resolve a function name to its implementation,
//! and helper utilities used by individual function implementations.

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::formula::ast::Expr;
use crate::formula::eval::Evaluator;

/// Signature for a built-in function implementation.
///
/// Functions receive unevaluated argument expressions and a mutable evaluator,
/// allowing short-circuit evaluation (e.g., IF) and range expansion (e.g., SUM).
pub type FunctionFn = fn(&[Expr], &mut Evaluator) -> Result<CellValue>;

/// Resolve a function name (case-insensitive) to its implementation.
pub fn lookup_function(name: &str) -> Option<FunctionFn> {
    match name.to_ascii_uppercase().as_str() {
        "SUM" => Some(fn_sum),
        "AVERAGE" => Some(fn_average),
        "COUNT" => Some(fn_count),
        "COUNTA" => Some(fn_counta),
        "MIN" => Some(fn_min),
        "MAX" => Some(fn_max),
        "IF" => Some(fn_if),
        "ABS" => Some(fn_abs),
        "INT" => Some(fn_int),
        "ROUND" => Some(fn_round),
        "MOD" => Some(fn_mod),
        "POWER" => Some(fn_power),
        "SQRT" => Some(fn_sqrt),
        "LEN" => Some(fn_len),
        "LOWER" => Some(fn_lower),
        "UPPER" => Some(fn_upper),
        "TRIM" => Some(fn_trim),
        "LEFT" => Some(fn_left),
        "RIGHT" => Some(fn_right),
        "MID" => Some(fn_mid),
        "CONCATENATE" => Some(fn_concatenate),
        "AND" => Some(fn_and),
        "OR" => Some(fn_or),
        "NOT" => Some(fn_not),
        "ISNUMBER" => Some(fn_isnumber),
        "ISTEXT" => Some(fn_istext),
        "ISBLANK" => Some(fn_isblank),
        "ISERROR" => Some(fn_iserror),
        "VALUE" => Some(fn_value),
        "TEXT" => Some(fn_text),
        _ => None,
    }
}

/// Verify that `args` has between `min` and `max` entries (inclusive).
pub fn check_arg_count(name: &str, args: &[Expr], min: usize, max: usize) -> Result<()> {
    if args.len() < min || args.len() > max {
        let expected = if min == max {
            format!("{min}")
        } else {
            format!("{min}..{max}")
        };
        return Err(Error::WrongArgCount {
            name: name.to_string(),
            expected,
            got: args.len(),
        });
    }
    Ok(())
}

// -- Aggregate functions --

fn fn_sum(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SUM", args, 1, 255)?;
    let nums = ctx.collect_numbers(args)?;
    Ok(CellValue::Number(nums.iter().sum()))
}

fn fn_average(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("AVERAGE", args, 1, 255)?;
    let nums = ctx.collect_numbers(args)?;
    if nums.is_empty() {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }
    let sum: f64 = nums.iter().sum();
    Ok(CellValue::Number(sum / nums.len() as f64))
}

fn fn_count(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("COUNT", args, 1, 255)?;
    let values = ctx.flatten_args_to_values(args)?;
    let count = values
        .iter()
        .filter(|v| matches!(v, CellValue::Number(_) | CellValue::Date(_)))
        .count();
    Ok(CellValue::Number(count as f64))
}

fn fn_counta(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("COUNTA", args, 1, 255)?;
    let values = ctx.flatten_args_to_values(args)?;
    let count = values
        .iter()
        .filter(|v| !matches!(v, CellValue::Empty))
        .count();
    Ok(CellValue::Number(count as f64))
}

fn fn_min(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("MIN", args, 1, 255)?;
    let nums = ctx.collect_numbers(args)?;
    if nums.is_empty() {
        return Ok(CellValue::Number(0.0));
    }
    let min = nums.iter().copied().fold(f64::INFINITY, f64::min);
    Ok(CellValue::Number(min))
}

fn fn_max(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("MAX", args, 1, 255)?;
    let nums = ctx.collect_numbers(args)?;
    if nums.is_empty() {
        return Ok(CellValue::Number(0.0));
    }
    let max = nums.iter().copied().fold(f64::NEG_INFINITY, f64::max);
    Ok(CellValue::Number(max))
}

// -- Logical functions --

fn fn_if(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IF", args, 1, 3)?;
    let cond = ctx.eval_expr(&args[0])?;
    let truth = crate::formula::eval::coerce_to_bool(&cond)?;
    if truth {
        if args.len() > 1 {
            ctx.eval_expr(&args[1])
        } else {
            Ok(CellValue::Bool(true))
        }
    } else if args.len() > 2 {
        ctx.eval_expr(&args[2])
    } else {
        Ok(CellValue::Bool(false))
    }
}

fn fn_and(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("AND", args, 1, 255)?;
    let values = ctx.flatten_args_to_values(args)?;
    for v in &values {
        if matches!(v, CellValue::Empty) {
            continue;
        }
        if !crate::formula::eval::coerce_to_bool(v)? {
            return Ok(CellValue::Bool(false));
        }
    }
    Ok(CellValue::Bool(true))
}

fn fn_or(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("OR", args, 1, 255)?;
    let values = ctx.flatten_args_to_values(args)?;
    for v in &values {
        if matches!(v, CellValue::Empty) {
            continue;
        }
        if crate::formula::eval::coerce_to_bool(v)? {
            return Ok(CellValue::Bool(true));
        }
    }
    Ok(CellValue::Bool(false))
}

fn fn_not(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("NOT", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let b = crate::formula::eval::coerce_to_bool(&v)?;
    Ok(CellValue::Bool(!b))
}

// -- Math functions --

fn fn_abs(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ABS", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let n = crate::formula::eval::coerce_to_number(&v)?;
    Ok(CellValue::Number(n.abs()))
}

fn fn_int(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("INT", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let n = crate::formula::eval::coerce_to_number(&v)?;
    Ok(CellValue::Number(n.floor()))
}

fn fn_round(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ROUND", args, 2, 2)?;
    let v = ctx.eval_expr(&args[0])?;
    let d = ctx.eval_expr(&args[1])?;
    let n = crate::formula::eval::coerce_to_number(&v)?;
    let digits = crate::formula::eval::coerce_to_number(&d)? as i32;
    let factor = 10f64.powi(digits);
    Ok(CellValue::Number((n * factor).round() / factor))
}

fn fn_mod(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("MOD", args, 2, 2)?;
    let a = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let b = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    if b == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }
    // Excel MOD: result has the sign of the divisor.
    let result = a - (a / b).floor() * b;
    Ok(CellValue::Number(result))
}

fn fn_power(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("POWER", args, 2, 2)?;
    let base = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let exp = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    Ok(CellValue::Number(base.powf(exp)))
}

fn fn_sqrt(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SQRT", args, 1, 1)?;
    let n = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    if n < 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    Ok(CellValue::Number(n.sqrt()))
}

// -- Text functions --

fn fn_len(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("LEN", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let s = crate::formula::eval::coerce_to_string(&v);
    Ok(CellValue::Number(s.len() as f64))
}

fn fn_lower(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("LOWER", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let s = crate::formula::eval::coerce_to_string(&v);
    Ok(CellValue::String(s.to_lowercase()))
}

fn fn_upper(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("UPPER", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let s = crate::formula::eval::coerce_to_string(&v);
    Ok(CellValue::String(s.to_uppercase()))
}

fn fn_trim(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("TRIM", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    let s = crate::formula::eval::coerce_to_string(&v);
    // Excel TRIM removes leading/trailing spaces and collapses internal runs.
    let trimmed: String = s.split_whitespace().collect::<Vec<_>>().join(" ");
    Ok(CellValue::String(trimmed))
}

fn fn_left(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("LEFT", args, 1, 2)?;
    let v = ctx.eval_expr(&args[0])?;
    let s = crate::formula::eval::coerce_to_string(&v);
    let n = if args.len() > 1 {
        crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[1])?)? as usize
    } else {
        1
    };
    let result: String = s.chars().take(n).collect();
    Ok(CellValue::String(result))
}

fn fn_right(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("RIGHT", args, 1, 2)?;
    let v = ctx.eval_expr(&args[0])?;
    let s = crate::formula::eval::coerce_to_string(&v);
    let n = if args.len() > 1 {
        crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[1])?)? as usize
    } else {
        1
    };
    let chars: Vec<char> = s.chars().collect();
    let start = chars.len().saturating_sub(n);
    let result: String = chars[start..].iter().collect();
    Ok(CellValue::String(result))
}

fn fn_mid(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("MID", args, 3, 3)?;
    let v = ctx.eval_expr(&args[0])?;
    let s = crate::formula::eval::coerce_to_string(&v);
    let start = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[1])?)? as usize;
    let count = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[2])?)? as usize;
    if start < 1 {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let result: String = s.chars().skip(start - 1).take(count).collect();
    Ok(CellValue::String(result))
}

fn fn_concatenate(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("CONCATENATE", args, 1, 255)?;
    let mut result = String::new();
    for arg in args {
        let v = ctx.eval_expr(arg)?;
        result.push_str(&crate::formula::eval::coerce_to_string(&v));
    }
    Ok(CellValue::String(result))
}

// -- Information functions --

fn fn_isnumber(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ISNUMBER", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    Ok(CellValue::Bool(matches!(
        v,
        CellValue::Number(_) | CellValue::Date(_)
    )))
}

fn fn_istext(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ISTEXT", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    Ok(CellValue::Bool(matches!(v, CellValue::String(_))))
}

fn fn_isblank(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ISBLANK", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    Ok(CellValue::Bool(matches!(v, CellValue::Empty)))
}

fn fn_iserror(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ISERROR", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    Ok(CellValue::Bool(matches!(v, CellValue::Error(_))))
}

// -- Conversion functions --

fn fn_value(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("VALUE", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    match crate::formula::eval::coerce_to_number(&v) {
        Ok(n) => Ok(CellValue::Number(n)),
        Err(_) => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

fn fn_text(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("TEXT", args, 2, 2)?;
    let v = ctx.eval_expr(&args[0])?;
    let _fmt = ctx.eval_expr(&args[1])?;
    // Simplified: just convert to string representation (full format codes not implemented).
    Ok(CellValue::String(crate::formula::eval::coerce_to_string(
        &v,
    )))
}
