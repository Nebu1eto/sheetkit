//! Built-in Excel function dispatch.
//!
//! Provides [`lookup_function`] to resolve a function name to its implementation,
//! and helper utilities used by individual function implementations.

pub mod date_time;
pub mod information;
pub mod logical;
pub mod lookup;
pub mod math;
pub mod statistical;
pub mod text;

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_number, coerce_to_string, Evaluator};

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
        "SUMIF" => Some(math::fn_sumif),
        "SUMIFS" => Some(math::fn_sumifs),
        "ROUNDUP" => Some(math::fn_roundup),
        "ROUNDDOWN" => Some(math::fn_rounddown),
        "CEILING" => Some(math::fn_ceiling),
        "FLOOR" => Some(math::fn_floor),
        "SIGN" => Some(math::fn_sign),
        "RAND" => Some(math::fn_rand),
        "RANDBETWEEN" => Some(math::fn_randbetween),
        "PI" => Some(math::fn_pi),
        "LOG" => Some(math::fn_log),
        "LOG10" => Some(math::fn_log10),
        "LN" => Some(math::fn_ln),
        "EXP" => Some(math::fn_exp),
        "PRODUCT" => Some(math::fn_product),
        "QUOTIENT" => Some(math::fn_quotient),
        "FACT" => Some(math::fn_fact),
        "AVERAGEIF" => Some(statistical::fn_averageif),
        "AVERAGEIFS" => Some(statistical::fn_averageifs),
        "COUNTBLANK" => Some(statistical::fn_countblank),
        "COUNTIF" => Some(statistical::fn_countif),
        "COUNTIFS" => Some(statistical::fn_countifs),
        "MEDIAN" => Some(statistical::fn_median),
        "MODE" => Some(statistical::fn_mode),
        "LARGE" => Some(statistical::fn_large),
        "SMALL" => Some(statistical::fn_small),
        "RANK" => Some(statistical::fn_rank),
        "ISERR" => Some(information::fn_iserr),
        "ISNA" => Some(information::fn_isna),
        "ISLOGICAL" => Some(information::fn_islogical),
        "ISEVEN" => Some(information::fn_iseven),
        "ISODD" => Some(information::fn_isodd),
        "TYPE" => Some(information::fn_type),
        "N" => Some(information::fn_n),
        "NA" => Some(information::fn_na),
        "ERROR.TYPE" => Some(information::fn_error_type),
        "CONCAT" => Some(text::fn_concat),
        "FIND" => Some(text::fn_find),
        "SEARCH" => Some(text::fn_search),
        "SUBSTITUTE" => Some(text::fn_substitute),
        "REPLACE" => Some(text::fn_replace),
        "REPT" => Some(text::fn_rept),
        "EXACT" => Some(text::fn_exact),
        "T" => Some(text::fn_t),
        "PROPER" => Some(text::fn_proper),
        "TRUE" => Some(logical::fn_true),
        "FALSE" => Some(logical::fn_false),
        "IFERROR" => Some(logical::fn_iferror),
        "IFNA" => Some(logical::fn_ifna),
        "IFS" => Some(logical::fn_ifs),
        "SWITCH" => Some(logical::fn_switch),
        "XOR" => Some(logical::fn_xor),
        "DATE" => Some(date_time::fn_date),
        "TODAY" => Some(date_time::fn_today),
        "NOW" => Some(date_time::fn_now),
        "YEAR" => Some(date_time::fn_year),
        "MONTH" => Some(date_time::fn_month),
        "DAY" => Some(date_time::fn_day),
        "HOUR" => Some(date_time::fn_hour),
        "MINUTE" => Some(date_time::fn_minute),
        "SECOND" => Some(date_time::fn_second),
        "DATEDIF" => Some(date_time::fn_datedif),
        "EDATE" => Some(date_time::fn_edate),
        "EOMONTH" => Some(date_time::fn_eomonth),
        "DATEVALUE" => Some(date_time::fn_datevalue),
        "WEEKDAY" => Some(date_time::fn_weekday),
        "WEEKNUM" => Some(date_time::fn_weeknum),
        "NETWORKDAYS" => Some(date_time::fn_networkdays),
        "WORKDAY" => Some(date_time::fn_workday),
        "VLOOKUP" => Some(lookup::fn_vlookup),
        "HLOOKUP" => Some(lookup::fn_hlookup),
        "INDEX" => Some(lookup::fn_index),
        "MATCH" => Some(lookup::fn_match),
        "LOOKUP" => Some(lookup::fn_lookup),
        "ROW" => Some(lookup::fn_row),
        "COLUMN" => Some(lookup::fn_column),
        "ROWS" => Some(lookup::fn_rows),
        "COLUMNS" => Some(lookup::fn_columns),
        "CHOOSE" => Some(lookup::fn_choose),
        "ADDRESS" => Some(lookup::fn_address),
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

/// Check whether a cell value matches a criteria string.
///
/// Criteria format: ">5", "<=10", "<>text", "=exact", plain "text".
/// Numeric comparisons are used when both sides parse as numbers;
/// otherwise strings are compared case-insensitively. Wildcard `*` and
/// `?` at the start/end of plain text are supported in a simplified form.
pub fn matches_criteria(cell_value: &CellValue, criteria: &str) -> bool {
    if criteria.is_empty() {
        return matches!(cell_value, CellValue::Empty);
    }

    let (op, val_str) = if let Some(rest) = criteria.strip_prefix("<=") {
        ("<=", rest)
    } else if let Some(rest) = criteria.strip_prefix(">=") {
        (">=", rest)
    } else if let Some(rest) = criteria.strip_prefix("<>") {
        ("<>", rest)
    } else if let Some(rest) = criteria.strip_prefix('<') {
        ("<", rest)
    } else if let Some(rest) = criteria.strip_prefix('>') {
        (">", rest)
    } else if let Some(rest) = criteria.strip_prefix('=') {
        ("=", rest)
    } else {
        ("=", criteria)
    };

    let cell_num = coerce_to_number(cell_value).ok();
    let crit_num: Option<f64> = val_str.parse().ok();

    if let (Some(cn), Some(crn)) = (cell_num, crit_num) {
        return match op {
            "<=" => cn <= crn,
            ">=" => cn >= crn,
            "<>" => (cn - crn).abs() > f64::EPSILON,
            "<" => cn < crn,
            ">" => cn > crn,
            "=" => (cn - crn).abs() < f64::EPSILON,
            _ => false,
        };
    }

    let cell_str = coerce_to_string(cell_value).to_ascii_lowercase();
    let crit_lower = val_str.to_ascii_lowercase();

    match op {
        "=" => {
            if crit_lower.contains('*') || crit_lower.contains('?') {
                wildcard_match(&cell_str, &crit_lower)
            } else {
                cell_str == crit_lower
            }
        }
        "<>" => {
            if crit_lower.contains('*') || crit_lower.contains('?') {
                !wildcard_match(&cell_str, &crit_lower)
            } else {
                cell_str != crit_lower
            }
        }
        "<" => cell_str < crit_lower,
        ">" => cell_str > crit_lower,
        "<=" => cell_str <= crit_lower,
        ">=" => cell_str >= crit_lower,
        _ => false,
    }
}

fn wildcard_match(text: &str, pattern: &str) -> bool {
    let t: Vec<char> = text.chars().collect();
    let p: Vec<char> = pattern.chars().collect();
    let (tlen, plen) = (t.len(), p.len());
    let mut dp = vec![vec![false; plen + 1]; tlen + 1];
    dp[0][0] = true;
    for j in 1..=plen {
        if p[j - 1] == '*' {
            dp[0][j] = dp[0][j - 1];
        }
    }
    for i in 1..=tlen {
        for j in 1..=plen {
            if p[j - 1] == '*' {
                dp[i][j] = dp[i][j - 1] || dp[i - 1][j];
            } else if p[j - 1] == '?' || p[j - 1] == t[i - 1] {
                dp[i][j] = dp[i - 1][j - 1];
            }
        }
    }
    dp[tlen][plen]
}

/// Expand a single argument expression into a flat list of CellValues.
pub fn collect_criteria_range_values(arg: &Expr, ctx: &mut Evaluator) -> Result<Vec<CellValue>> {
    match arg {
        Expr::Range { start, end } => ctx.expand_range(start, end),
        _ => {
            let v = ctx.eval_expr(arg)?;
            Ok(vec![v])
        }
    }
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
