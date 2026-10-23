//! Built-in Excel function dispatch.
//!
//! Provides [`lookup_function`] to resolve a function name to its implementation,
//! and helper utilities used by individual function implementations.

pub(crate) mod conversion;
pub mod date_time;
pub mod engineering;
pub mod financial;
pub mod information;
pub mod logical;
pub mod lookup;
pub mod math;
pub mod statistical;
pub mod text;

use std::borrow::Cow;

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
        "FV" => Some(financial::fn_fv),
        "PV" => Some(financial::fn_pv),
        "NPV" => Some(financial::fn_npv),
        "IRR" => Some(financial::fn_irr),
        "PMT" => Some(financial::fn_pmt),
        "IPMT" => Some(financial::fn_ipmt),
        "PPMT" => Some(financial::fn_ppmt),
        "RATE" => Some(financial::fn_rate),
        "NPER" => Some(financial::fn_nper),
        "DB" => Some(financial::fn_db),
        "DDB" => Some(financial::fn_ddb),
        "SLN" => Some(financial::fn_sln),
        "SYD" => Some(financial::fn_syd),
        "EFFECT" => Some(financial::fn_effect),
        "NOMINAL" => Some(financial::fn_nominal),
        "DOLLARDE" => Some(financial::fn_dollarde),
        "DOLLARFR" => Some(financial::fn_dollarfr),
        "CUMIPMT" => Some(financial::fn_cumipmt),
        "CUMPRINC" => Some(financial::fn_cumprinc),
        "XNPV" => Some(financial::fn_xnpv),
        "XIRR" => Some(financial::fn_xirr),
        "BIN2DEC" => Some(engineering::fn_bin2dec),
        "BIN2HEX" => Some(engineering::fn_bin2hex),
        "BIN2OCT" => Some(engineering::fn_bin2oct),
        "DEC2BIN" => Some(engineering::fn_dec2bin),
        "DEC2HEX" => Some(engineering::fn_dec2hex),
        "DEC2OCT" => Some(engineering::fn_dec2oct),
        "HEX2BIN" => Some(engineering::fn_hex2bin),
        "HEX2DEC" => Some(engineering::fn_hex2dec),
        "HEX2OCT" => Some(engineering::fn_hex2oct),
        "OCT2BIN" => Some(engineering::fn_oct2bin),
        "OCT2DEC" => Some(engineering::fn_oct2dec),
        "OCT2HEX" => Some(engineering::fn_oct2hex),
        "DELTA" => Some(engineering::fn_delta),
        "GESTEP" => Some(engineering::fn_gestep),
        "ERF" => Some(engineering::fn_erf),
        "ERFC" => Some(engineering::fn_erfc),
        "COMPLEX" => Some(engineering::fn_complex),
        "IMREAL" => Some(engineering::fn_imreal),
        "IMAGINARY" => Some(engineering::fn_imaginary),
        "IMABS" => Some(engineering::fn_imabs),
        "IMARGUMENT" => Some(engineering::fn_imargument),
        "IMCONJUGATE" => Some(engineering::fn_imconjugate),
        "IMSUM" => Some(engineering::fn_imsum),
        "IMSUB" => Some(engineering::fn_imsub),
        "IMPRODUCT" => Some(engineering::fn_improduct),
        "IMDIV" => Some(engineering::fn_imdiv),
        "IMPOWER" => Some(engineering::fn_impower),
        "IMSQRT" => Some(engineering::fn_imsqrt),
        "CONVERT" => Some(engineering::fn_convert),
        "BESSELI" => Some(engineering::fn_besseli),
        "BESSELJ" => Some(engineering::fn_besselj),
        "BESSELK" => Some(engineering::fn_besselk),
        "BESSELY" => Some(engineering::fn_bessely),
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

fn finite_number_or_num_error(value: f64) -> Result<CellValue> {
    if value.is_finite() {
        Ok(CellValue::Number(value))
    } else {
        Ok(CellValue::Error("#NUM!".to_string()))
    }
}

/// Check whether a cell value matches a criteria string.
///
/// Criteria format: ">5", "<=10", "<>text", "=exact", plain "text".
/// Numeric comparisons are used when both sides parse as numbers;
/// otherwise strings are compared case-insensitively. Wildcards use Excel's
/// `*` and `?` behavior, with `~` escaping `*`, `?`, or `~`.
pub fn matches_criteria(cell_value: &CellValue, criteria: &str, ctx: &Evaluator) -> Result<bool> {
    matches_criteria_with_prepaid(cell_value, criteria, 0, ctx)
}

fn matches_criteria_with_prepaid(
    cell_value: &CellValue,
    criteria: &str,
    prepaid_pattern_bytes: u64,
    ctx: &Evaluator,
) -> Result<bool> {
    if criteria.is_empty() {
        return Ok(matches!(cell_value, CellValue::Empty));
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

    if crit_num.is_some() && matches!(cell_value, CellValue::Empty) {
        return Ok(false);
    }

    if let (Some(cn), Some(crn)) = (cell_num, crit_num) {
        return Ok(match op {
            "<=" => cn <= crn,
            ">=" => cn >= crn,
            "<>" => (cn - crn).abs() > f64::EPSILON,
            "<" => cn < crn,
            ">" => cn > crn,
            "=" => (cn - crn).abs() < f64::EPSILON,
            _ => false,
        });
    }

    let uses_wildcards = val_str.contains(['*', '?', '~']);
    if matches!(op, "=" | "<>") && uses_wildcards {
        let (cell_text, prepaid_text_bytes) = cell_value_as_wildcard_text(cell_value, ctx)?;
        let prepaid_decode_bytes = prepaid_pattern_bytes.saturating_add(prepaid_text_bytes);
        let matched =
            wildcard_match_with_prepaid(cell_text.as_ref(), val_str, prepaid_decode_bytes, ctx)?;
        return Ok(if op == "=" { matched } else { !matched });
    }
    let cell_text = cell_value_as_text(cell_value);
    let cell_lower = cell_text.to_lowercase();
    let crit_lower = val_str.to_lowercase();

    Ok(match op {
        "=" => cell_lower == crit_lower,
        "<>" => cell_lower != crit_lower,
        "<" => cell_lower < crit_lower,
        ">" => cell_lower > crit_lower,
        "<=" => cell_lower <= crit_lower,
        ">=" => cell_lower >= crit_lower,
        _ => false,
    })
}

pub(crate) struct PreparedCriteria {
    text: String,
    prepaid_match_bytes: u64,
}

pub(crate) fn prepare_criteria(
    criteria_value: CellValue,
    ctx: &Evaluator,
) -> Result<PreparedCriteria> {
    let (text, prepaid_match_bytes) = cell_value_into_criteria_text(criteria_value, ctx)?;
    Ok(PreparedCriteria {
        text,
        prepaid_match_bytes,
    })
}

pub(crate) fn matches_prepared_criteria(
    cell_value: &CellValue,
    criteria: &PreparedCriteria,
    ctx: &Evaluator,
) -> Result<bool> {
    matches_criteria_with_prepaid(
        cell_value,
        &criteria.text,
        criteria.prepaid_match_bytes,
        ctx,
    )
}

fn cell_value_as_text(value: &CellValue) -> Cow<'_, str> {
    match value {
        CellValue::String(text) | CellValue::Error(text) => Cow::Borrowed(text),
        CellValue::Formula {
            result: Some(result),
            ..
        } => cell_value_as_text(result),
        CellValue::Empty | CellValue::Formula { result: None, .. } => Cow::Borrowed(""),
        _ => Cow::Owned(coerce_to_string(value)),
    }
}

fn cell_value_into_criteria_text(value: CellValue, ctx: &Evaluator) -> Result<(String, u64)> {
    match value {
        CellValue::String(text) | CellValue::Error(text) => Ok((text, 0)),
        CellValue::Formula {
            result: Some(result),
            ..
        } => cell_value_into_criteria_text(*result, ctx),
        CellValue::Empty | CellValue::Formula { result: None, .. } => Ok((String::new(), 0)),
        CellValue::RichString(runs) => flatten_rich_text_with_budget(&runs, ctx),
        other => Ok((coerce_to_string(&other), 0)),
    }
}

fn cell_value_as_wildcard_text<'a>(
    value: &'a CellValue,
    ctx: &Evaluator,
) -> Result<(Cow<'a, str>, u64)> {
    match value {
        CellValue::String(text) | CellValue::Error(text) => Ok((Cow::Borrowed(text), 0)),
        CellValue::Formula {
            result: Some(result),
            ..
        } => cell_value_as_wildcard_text(result, ctx),
        CellValue::Empty | CellValue::Formula { result: None, .. } => Ok((Cow::Borrowed(""), 0)),
        CellValue::RichString(runs) => {
            let (text, prepaid_bytes) = flatten_rich_text_with_budget(runs, ctx)?;
            Ok((Cow::Owned(text), prepaid_bytes))
        }
        other => Ok((Cow::Owned(coerce_to_string(other)), 0)),
    }
}

fn flatten_rich_text_with_budget(
    runs: &[crate::rich_text::RichTextRun],
    ctx: &Evaluator,
) -> Result<(String, u64)> {
    let byte_len = runs
        .iter()
        .try_fold(0usize, |total, run| total.checked_add(run.text.len()))
        .and_then(|len| u64::try_from(len).ok())
        .unwrap_or(u64::MAX);
    ctx.charge_match_work(byte_len)?;
    let capacity = usize::try_from(byte_len)
        .map_err(|_| Error::Internal("rich criteria length overflow".to_string()))?;
    let mut text = String::new();
    text.try_reserve_exact(capacity)
        .map_err(|_| Error::Internal("rich criteria allocation failed".to_string()))?;
    for run in runs {
        text.push_str(&run.text);
    }
    Ok((text, byte_len))
}

#[derive(Clone, Copy)]
enum WildcardToken {
    Star,
    AnyOne,
    Literal(char),
}

#[cfg(test)]
fn wildcard_match(text: &str, pattern: &str, ctx: &Evaluator) -> Result<bool> {
    wildcard_match_with_prepaid(text, pattern, 0, ctx)
}

fn wildcard_match_with_prepaid(
    text: &str,
    pattern: &str,
    prepaid_decode_bytes: u64,
    ctx: &Evaluator,
) -> Result<bool> {
    let decode_work = u64::try_from(text.len())
        .ok()
        .and_then(|text_bytes| {
            u64::try_from(pattern.len())
                .ok()
                .and_then(|pattern_bytes| text_bytes.checked_add(pattern_bytes))
        })
        .unwrap_or(u64::MAX);
    ctx.charge_match_work(decode_work.saturating_sub(prepaid_decode_bytes))?;

    let text_len = u64::try_from(text.chars().count()).unwrap_or(u64::MAX);
    let pattern_len = wildcard_pattern_token_count(pattern);
    let work = text_len
        .checked_add(1)
        .and_then(|rows| {
            pattern_len
                .checked_add(1)
                .and_then(|cols| rows.checked_mul(cols))
        })
        .unwrap_or(u64::MAX);
    ctx.charge_match_work(work)?;

    let tokens = tokenize_wildcard_pattern(pattern)?;
    debug_assert_eq!(u64::try_from(tokens.len()).unwrap_or(u64::MAX), pattern_len);
    let row_len = tokens
        .len()
        .checked_add(1)
        .ok_or_else(|| Error::Internal("wildcard row length overflow".to_string()))?;
    let mut row = Vec::new();
    row.try_reserve_exact(row_len)
        .map_err(|_| Error::Internal("wildcard row allocation failed".to_string()))?;
    row.resize(row_len, false);
    row[0] = true;
    for index in 1..row_len {
        row[index] = matches!(tokens[index - 1], WildcardToken::Star) && row[index - 1];
    }

    for text_char in text.chars() {
        let mut diagonal = row[0];
        row[0] = false;
        for index in 1..row_len {
            let above = row[index];
            row[index] = match tokens[index - 1] {
                WildcardToken::Star => row[index - 1] || above,
                WildcardToken::AnyOne => diagonal,
                WildcardToken::Literal(pattern_char) => {
                    diagonal && chars_equal_case_insensitive(text_char, pattern_char)
                }
            };
            diagonal = above;
        }
    }

    Ok(row[row_len - 1])
}

pub(crate) fn wildcard_search(text: &str, pattern: &str, ctx: &Evaluator) -> Result<Option<usize>> {
    let decode_work = u64::try_from(text.len())
        .ok()
        .and_then(|text_bytes| {
            u64::try_from(pattern.len())
                .ok()
                .and_then(|pattern_bytes| text_bytes.checked_add(pattern_bytes))
        })
        .unwrap_or(u64::MAX);
    ctx.charge_match_work(decode_work)?;

    let text_len = u64::try_from(text.chars().count()).unwrap_or(u64::MAX);
    let pattern_len = wildcard_pattern_token_count(pattern);
    let work = text_len
        .checked_add(1)
        .and_then(|rows| {
            pattern_len
                .checked_add(1)
                .and_then(|cols| rows.checked_mul(cols))
        })
        .unwrap_or(u64::MAX);
    ctx.charge_match_work(work)?;

    let tokens = tokenize_wildcard_pattern(pattern)?;
    let row_len = tokens
        .len()
        .checked_add(1)
        .ok_or_else(|| Error::Internal("wildcard search row length overflow".to_string()))?;
    let mut row = Vec::new();
    row.try_reserve_exact(row_len)
        .map_err(|_| Error::Internal("wildcard search allocation failed".to_string()))?;
    row.resize(row_len, None);
    let mut next = Vec::new();
    next.try_reserve_exact(row_len)
        .map_err(|_| Error::Internal("wildcard search allocation failed".to_string()))?;
    next.resize(row_len, None);

    for (char_index, text_char) in text.chars().enumerate() {
        row[0] = Some(char_index);
        apply_wildcard_search_epsilon(&mut row, &tokens);
        if let Some(start) = row[row_len - 1] {
            return Ok(Some(start));
        }

        next.fill(None);
        for index in 1..row_len {
            next[index] = match tokens[index - 1] {
                WildcardToken::Star => row[index],
                WildcardToken::AnyOne => row[index - 1],
                WildcardToken::Literal(pattern_char) => {
                    row[index - 1].filter(|_| chars_equal_case_insensitive(text_char, pattern_char))
                }
            };
        }
        std::mem::swap(&mut row, &mut next);
    }

    row[0] = Some(usize::try_from(text_len).unwrap_or(usize::MAX));
    apply_wildcard_search_epsilon(&mut row, &tokens);
    Ok(row[row_len - 1])
}

fn apply_wildcard_search_epsilon(row: &mut [Option<usize>], tokens: &[WildcardToken]) {
    for index in 1..row.len() {
        if matches!(tokens[index - 1], WildcardToken::Star) {
            row[index] = match (row[index], row[index - 1]) {
                (Some(current), Some(candidate)) => Some(current.min(candidate)),
                (current @ Some(_), None) => current,
                (None, candidate) => candidate,
            };
        }
    }
}

fn wildcard_pattern_token_count(pattern: &str) -> u64 {
    let mut count = 0u64;
    let mut previous_was_star = false;
    let mut chars = pattern.chars().peekable();
    while let Some(ch) = chars.next() {
        let is_star = ch == '*';
        if ch == '~' && matches!(chars.peek(), Some('*' | '?' | '~')) {
            chars.next();
        }
        if is_star && previous_was_star {
            continue;
        }
        count = count.saturating_add(1);
        previous_was_star = is_star;
    }
    count
}

fn tokenize_wildcard_pattern(pattern: &str) -> Result<Vec<WildcardToken>> {
    let mut tokens = Vec::new();
    tokens
        .try_reserve_exact(pattern.len())
        .map_err(|_| Error::Internal("wildcard pattern allocation failed".to_string()))?;
    let mut chars = pattern.chars().peekable();
    while let Some(ch) = chars.next() {
        let token = match ch {
            '*' => WildcardToken::Star,
            '?' => WildcardToken::AnyOne,
            '~' => match chars.peek().copied() {
                Some(next @ ('*' | '?' | '~')) => {
                    chars.next();
                    WildcardToken::Literal(next)
                }
                _ => WildcardToken::Literal('~'),
            },
            literal => WildcardToken::Literal(literal),
        };
        if matches!(token, WildcardToken::Star)
            && matches!(tokens.last(), Some(WildcardToken::Star))
        {
            continue;
        }
        tokens.push(token);
    }
    Ok(tokens)
}

fn chars_equal_case_insensitive(left: char, right: char) -> bool {
    left.to_lowercase().eq(right.to_lowercase())
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
    let digits = crate::formula::eval::coerce_to_number(&d)?;
    Ok(math::round_to_digits(n, digits, math::RoundMode::Nearest))
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
    let result = base.powf(exp);
    if !result.is_finite() {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    Ok(CellValue::Number(result))
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
    Ok(CellValue::Number(s.chars().count() as f64))
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
        let count = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[1])?)?;
        let Some(count) = conversion::checked_trunc_count(count) else {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        };
        count
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
        let count = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[1])?)?;
        let Some(count) = conversion::checked_trunc_count(count) else {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        };
        count
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
    let start = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let count = crate::formula::eval::coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let (Some(start), Some(count)) = (
        conversion::checked_trunc_count(start),
        conversion::checked_trunc_count(count),
    ) else {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    };
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
    let value = ctx.eval_expr(&args[0])?;
    let format = ctx.eval_expr(&args[1])?;
    let CellValue::String(format_code) = format else {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    };
    let Ok(number) = crate::formula::eval::coerce_to_number(&value) else {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    };
    Ok(CellValue::String(crate::numfmt::format_number(
        number,
        &format_code,
    )))
}

#[cfg(test)]
mod tests {
    use super::{engineering, financial, matches_criteria, math, text, wildcard_match};
    use crate::cell::CellValue;
    use crate::formula::ast::Expr;
    use crate::formula::budget::{FormulaBudget, FormulaLimits};
    use crate::formula::eval::{evaluate, CellSnapshot, Evaluator};
    use crate::formula::parser::parse_formula;
    use crate::rich_text::RichTextRun;

    fn eval(formula: &str) -> CellValue {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let expr = parse_formula(formula).unwrap();
        evaluate(&expr, &snapshot).unwrap()
    }

    #[test]
    fn numeric_criteria_does_not_match_blank() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let evaluator = Evaluator::new(&snapshot);
        assert!(!matches_criteria(&CellValue::Empty, "0", &evaluator).unwrap());
        assert!(matches_criteria(&CellValue::Number(0.0), "0", &evaluator).unwrap());
        assert!(matches_criteria(&CellValue::Empty, "", &evaluator).unwrap());
        assert!(matches_criteria(&CellValue::Empty, "<>word", &evaluator).unwrap());
    }

    #[test]
    fn wildcard_match_preserves_excel_escaping_and_unicode_semantics() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let evaluator = Evaluator::new(&snapshot);

        assert!(wildcard_match("CAFÉ", "caf?", &evaluator).unwrap());
        assert!(wildcard_match("report.xlsx", "*.xlsx", &evaluator).unwrap());
        assert!(wildcard_match("*", "~*", &evaluator).unwrap());
        assert!(wildcard_match("?", "~?", &evaluator).unwrap());
        assert!(wildcard_match("~", "~~", &evaluator).unwrap());
        assert!(!wildcard_match("ab", "a?c", &evaluator).unwrap());
    }

    #[test]
    fn wildcard_match_work_is_exact_and_cumulative() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let limits = FormulaLimits {
            operation_match_work: 17,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let evaluator = Evaluator::with_budget(&snapshot, budget);
        assert!(wildcard_match("abc", "a*", &evaluator).unwrap());

        let limits = FormulaLimits {
            operation_match_work: 16,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let evaluator = Evaluator::with_budget(&snapshot, budget);
        let error = wildcard_match("abc", "a*", &evaluator).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "match_work");
        assert_eq!((metadata.limit(), metadata.observed()), (16, 17));

        let limits = FormulaLimits {
            operation_match_work: 33,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let evaluator = Evaluator::with_budget(&snapshot, budget);
        assert!(wildcard_match("abc", "a*", &evaluator).unwrap());
        let error = wildcard_match("abc", "a*", &evaluator).unwrap_err();
        assert_eq!(error.resource_metadata().unwrap().observed(), 34);
    }

    #[test]
    fn criteria_families_share_one_match_budget() {
        let mut snapshot = CellSnapshot::new("Sheet1".to_string());
        snapshot.set_cell("Sheet1", 1, 1, CellValue::String("CAFÉ".to_string()));
        snapshot.set_cell("Sheet1", 2, 1, CellValue::Number(2.0));
        let expression = parse_formula(
            "SUM(SUMIF(A1,\"caf?\",B1),SUMIFS(B1,A1,\"caf?\"),AVERAGEIF(A1,\"caf?\",B1),AVERAGEIFS(B1,A1,\"caf?\"),COUNTIFS(A1,\"caf?\"))",
        )
        .unwrap();

        let exact = FormulaLimits {
            operation_match_work: 170,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(exact).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        assert_eq!(
            evaluator.eval_expr(&expression).unwrap(),
            CellValue::Number(9.0)
        );

        let over = FormulaLimits {
            operation_match_work: 169,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(over).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        let error = evaluator.eval_expr(&expression).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "match_work");
        assert_eq!((metadata.limit(), metadata.observed()), (169, 170));
    }

    #[test]
    fn rich_text_criterion_is_precharged_and_flattened_once() {
        let mut snapshot = CellSnapshot::new("Sheet1".to_string());
        for row in 1..=3 {
            snapshot.set_cell("Sheet1", 1, row, CellValue::String("abc".to_string()));
            snapshot.set_cell("Sheet1", 3, row, CellValue::Number(f64::from(row)));
        }
        snapshot.set_cell(
            "Sheet1",
            2,
            1,
            CellValue::RichString(vec![
                RichTextRun {
                    text: "*".to_string(),
                    font: None,
                    size: None,
                    bold: false,
                    italic: false,
                    color: None,
                },
                RichTextRun {
                    text: "a*".to_string(),
                    font: None,
                    size: None,
                    bold: false,
                    italic: false,
                    color: None,
                },
            ]),
        );
        let expression = parse_formula("SUMIF(A1:A3,B1,C1:C3)").unwrap();

        let exact = FormulaLimits {
            operation_match_work: 60,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(exact).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        assert_eq!(
            evaluator.eval_expr(&expression).unwrap(),
            CellValue::Number(6.0)
        );

        let over = FormulaLimits {
            operation_match_work: 59,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(over).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        let error = evaluator.eval_expr(&expression).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "match_work");
        assert_eq!((metadata.limit(), metadata.observed()), (59, 60));
    }

    fn assert_evaluation_work_boundary(
        function: super::FunctionFn,
        args: &[Expr],
        exact_work: u64,
    ) {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let limits = FormulaLimits {
            operation_evaluation_steps: exact_work,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        function(args, &mut evaluator).unwrap();

        let limits = FormulaLimits {
            operation_evaluation_steps: exact_work - 1,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        let error = function(args, &mut evaluator).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "operation_count");
        assert_eq!(
            (metadata.limit(), metadata.observed()),
            (exact_work - 1, exact_work)
        );
    }

    #[test]
    fn iterative_low_formula_functions_charge_exact_work() {
        assert_evaluation_work_boundary(math::fn_fact, &[Expr::Number(5.0)], 5);
        assert_evaluation_work_boundary(
            financial::fn_db,
            &[
                Expr::Number(1_000.0),
                Expr::Number(100.0),
                Expr::Number(4.0),
                Expr::Number(4.0),
            ],
            6,
        );
        assert_evaluation_work_boundary(
            financial::fn_ddb,
            &[
                Expr::Number(1_000.0),
                Expr::Number(100.0),
                Expr::Number(4.0),
                Expr::Number(3.0),
            ],
            7,
        );
        assert_evaluation_work_boundary(
            financial::fn_cumipmt,
            &[
                Expr::Number(0.01),
                Expr::Number(3.0),
                Expr::Number(1_000.0),
                Expr::Number(1.0),
                Expr::Number(3.0),
                Expr::Number(0.0),
            ],
            9,
        );
        for function in [engineering::fn_besselk, engineering::fn_bessely] {
            assert_evaluation_work_boundary(function, &[Expr::Number(16.0), Expr::Number(3.0)], 4);
        }
    }

    #[test]
    fn iterative_formula_work_cannot_reset_between_calls() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let limits = FormulaLimits {
            operation_evaluation_steps: 9,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        let args = [Expr::Number(5.0)];
        math::fn_fact(&args, &mut evaluator).unwrap();
        let error = math::fn_fact(&args, &mut evaluator).unwrap_err();
        assert_eq!(error.resource_metadata().unwrap().observed(), 10);
    }

    #[test]
    fn iterative_formula_inputs_use_checked_conversions() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        for function in [engineering::fn_besselk, engineering::fn_bessely] {
            let mut evaluator = Evaluator::new(&snapshot);
            assert_eq!(
                function(&[Expr::Number(16.0), Expr::Number(1e100)], &mut evaluator,).unwrap(),
                CellValue::Error("#NUM!".to_string())
            );
        }
        let mut evaluator = Evaluator::new(&snapshot);
        assert_eq!(
            financial::fn_cumipmt(
                &[
                    Expr::Number(0.01),
                    Expr::Number(3.0),
                    Expr::Number(1_000.0),
                    Expr::Number(1.0),
                    Expr::Number(1e100),
                    Expr::Number(0.0),
                ],
                &mut evaluator,
            )
            .unwrap(),
            CellValue::Error("#NUM!".to_string())
        );

        let mut evaluator = Evaluator::new(&snapshot);
        let fact_error = math::fn_fact(&[Expr::Number(1e100)], &mut evaluator).unwrap_err();
        assert_eq!(
            fact_error.resource_metadata().unwrap().dimension_code(),
            "operation_count"
        );

        for (function, args) in [
            (
                financial::fn_db as super::FunctionFn,
                vec![
                    Expr::Number(1_000.0),
                    Expr::Number(100.0),
                    Expr::Number(f64::from(i32::MAX - 1)),
                    Expr::Number(f64::from(i32::MAX - 1)),
                ],
            ),
            (
                financial::fn_ddb as super::FunctionFn,
                vec![
                    Expr::Number(1_000.0),
                    Expr::Number(100.0),
                    Expr::Number(f64::from(i32::MAX)),
                    Expr::Number(f64::from(i32::MAX)),
                ],
            ),
        ] {
            let mut evaluator = Evaluator::new(&snapshot);
            let error = function(&args, &mut evaluator).unwrap_err();
            assert_eq!(
                error.resource_metadata().unwrap().dimension_code(),
                "operation_count"
            );
        }
    }

    #[test]
    fn countif_and_search_share_exact_match_work() {
        let mut snapshot = CellSnapshot::new("Sheet1".to_string());
        snapshot.set_cell("Sheet1", 1, 1, CellValue::String("abc".to_string()));
        let countif = parse_formula("COUNTIF(A1,\"a*\")").unwrap();
        let limits = FormulaLimits {
            operation_match_work: 17,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        assert_eq!(
            evaluator.eval_expr(&countif).unwrap(),
            CellValue::Number(1.0)
        );

        let limits = FormulaLimits {
            operation_match_work: 16,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        let error = evaluator.eval_expr(&countif).unwrap_err();
        assert_eq!(error.resource_metadata().unwrap().observed(), 17);

        let args = [
            Expr::String("a*c".to_string()),
            Expr::String("abc".to_string()),
        ];
        let limits = FormulaLimits {
            operation_match_work: 22,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        assert_eq!(
            text::fn_search(&args, &mut evaluator).unwrap(),
            CellValue::Number(1.0)
        );

        let limits = FormulaLimits {
            operation_match_work: 21,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        let error = text::fn_search(&args, &mut evaluator).unwrap_err();
        assert_eq!(error.resource_metadata().unwrap().observed(), 22);
    }

    #[test]
    fn search_preserves_excel_escaping_unicode_and_cumulative_work() {
        assert_eq!(eval(r#"SEARCH("~*","a*b")"#), CellValue::Number(2.0));
        assert_eq!(eval(r#"SEARCH("~?","a?b")"#), CellValue::Number(2.0));
        assert_eq!(eval(r#"SEARCH("~~","a~b")"#), CellValue::Number(2.0));
        assert_eq!(eval(r#"SEARCH("caf?","xxCAFÉ")"#), CellValue::Number(3.0));

        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let limits = FormulaLimits {
            operation_match_work: 43,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        let args = [
            Expr::String("a*c".to_string()),
            Expr::String("abc".to_string()),
        ];
        text::fn_search(&args, &mut evaluator).unwrap();
        let error = text::fn_search(&args, &mut evaluator).unwrap_err();
        assert_eq!(error.resource_metadata().unwrap().observed(), 44);
    }

    #[test]
    fn search_rejects_large_failure_before_state_allocation() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let args = [
            Expr::String(format!("{}z", "*a".repeat(512))),
            Expr::String("a".repeat(1_024)),
        ];
        let mut evaluator = Evaluator::new(&snapshot);
        let error = text::fn_search(&args, &mut evaluator).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "match_work");
        assert!(metadata.observed() > metadata.limit());
    }

    #[test]
    fn substitute_and_width_outputs_are_precharged_exactly() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let substitute = Expr::Function {
            name: "SUBSTITUTE".to_string(),
            args: vec![
                Expr::String("aaaa".to_string()),
                Expr::String("a".to_string()),
                Expr::String("bb".to_string()),
            ],
        };
        let limits = FormulaLimits {
            operation_output_bytes: 15,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        assert_eq!(
            evaluator.eval_expr(&substitute).unwrap(),
            CellValue::String("bbbbbbbb".to_string())
        );

        let limits = FormulaLimits {
            operation_output_bytes: 14,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        let error = evaluator.eval_expr(&substitute).unwrap_err();
        assert_eq!(error.resource_metadata().unwrap().observed(), 15);

        for (function_name, args, exact) in [
            ("DEC2HEX", vec![Expr::Number(1.0), Expr::Number(10.0)], 10),
            (
                "HEX2BIN",
                vec![Expr::String("1".to_string()), Expr::Number(10.0)],
                11,
            ),
            (
                "HEX2OCT",
                vec![Expr::String("1".to_string()), Expr::Number(10.0)],
                11,
            ),
        ] {
            let expression = Expr::Function {
                name: function_name.to_string(),
                args,
            };
            let limits = FormulaLimits {
                operation_output_bytes: exact,
                ..FormulaLimits::secure_default()
            };
            let budget = FormulaBudget::for_calculation(limits).unwrap();
            let mut evaluator = Evaluator::with_budget(&snapshot, budget);
            evaluator.eval_expr(&expression).unwrap();

            let limits = FormulaLimits {
                operation_output_bytes: exact - 1,
                ..FormulaLimits::secure_default()
            };
            let budget = FormulaBudget::for_calculation(limits).unwrap();
            let mut evaluator = Evaluator::with_budget(&snapshot, budget);
            let error = evaluator.eval_expr(&expression).unwrap_err();
            assert_eq!(error.resource_metadata().unwrap().observed(), exact);
        }
    }

    #[test]
    fn substitute_rejects_large_amplification_before_allocation() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let args = [
            Expr::String("a".repeat(32_767)),
            Expr::String("a".to_string()),
            Expr::String("b".repeat(32_767)),
        ];
        let mut evaluator = Evaluator::new(&snapshot);
        let error = text::fn_substitute(&args, &mut evaluator).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "output_bytes");
        assert!(metadata.observed() > metadata.limit());
    }

    #[test]
    fn substitute_output_budget_is_cumulative() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let limits = FormulaLimits {
            operation_output_bytes: 29,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        let expression = Expr::Function {
            name: "SUBSTITUTE".to_string(),
            args: vec![
                Expr::String("aaaa".to_string()),
                Expr::String("a".to_string()),
                Expr::String("bb".to_string()),
            ],
        };
        evaluator.eval_expr(&expression).unwrap();
        let error = evaluator.eval_expr(&expression).unwrap_err();
        assert_eq!(error.resource_metadata().unwrap().observed(), 30);
    }

    #[test]
    fn cumprinc_charges_checked_period_work_before_iteration() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let args = [
            Expr::Number(0.01),
            Expr::Number(3.0),
            Expr::Number(1_000.0),
            Expr::Number(1.0),
            Expr::Number(3.0),
            Expr::Number(0.0),
        ];
        let exact = FormulaLimits {
            operation_evaluation_steps: 9,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(exact).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        assert!(matches!(
            financial::fn_cumprinc(&args, &mut evaluator).unwrap(),
            CellValue::Number(_)
        ));

        let over = FormulaLimits {
            operation_evaluation_steps: 8,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(over).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        let error = financial::fn_cumprinc(&args, &mut evaluator).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "operation_count");
        assert_eq!((metadata.limit(), metadata.observed()), (8, 9));
    }

    #[test]
    fn cumprinc_rejects_unrepresentable_periods_without_iteration() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let args = [
            Expr::Number(0.01),
            Expr::Number(3.0),
            Expr::Number(1_000.0),
            Expr::Number(1.0),
            Expr::Number(1e100),
            Expr::Number(0.0),
        ];
        let limits = FormulaLimits {
            operation_evaluation_steps: 6,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&snapshot, budget);
        assert_eq!(
            financial::fn_cumprinc(&args, &mut evaluator).unwrap(),
            CellValue::Error("#NUM!".to_string())
        );
    }

    #[test]
    fn cumprinc_rejects_non_finite_intermediate_results() {
        let snapshot = CellSnapshot::new("Sheet1".to_string());
        let args = [
            Expr::Number(f64::MAX),
            Expr::Number(1.0),
            Expr::Number(f64::MAX),
            Expr::Number(1.0),
            Expr::Number(1.0),
            Expr::Number(0.0),
        ];
        let mut evaluator = Evaluator::new(&snapshot);
        assert_eq!(
            financial::fn_cumprinc(&args, &mut evaluator).unwrap(),
            CellValue::Error("#NUM!".to_string())
        );
    }

    #[test]
    fn power_rejects_non_finite_results() {
        assert_eq!(eval("POWER(-1,0.5)"), CellValue::Error("#NUM!".to_string()));
        assert_eq!(
            eval("POWER(1E308,2)"),
            CellValue::Error("#NUM!".to_string())
        );
    }

    #[test]
    fn len_counts_unicode_scalars() {
        assert_eq!(eval("LEN(\"café\")"), CellValue::Number(4.0));
        assert_eq!(eval("LEN(\"한글🚀\")"), CellValue::Number(3.0));
    }

    #[test]
    fn text_applies_excel_number_formats() {
        assert_eq!(
            eval("TEXT(12.345,\"0.00\")"),
            CellValue::String("12.35".to_string())
        );
        assert_eq!(
            eval("TEXT(1234.5,\"#,##0.00\")"),
            CellValue::String("1,234.50".to_string())
        );
        assert_eq!(
            eval("TEXT(0.256,\"0.0%\")"),
            CellValue::String("25.6%".to_string())
        );
        assert_eq!(
            eval("TEXT(45292,\"m/d/yyyy\")"),
            CellValue::String("1/1/2024".to_string())
        );
        assert_eq!(
            eval("TEXT(0.625,\"h:mm AM/PM\")"),
            CellValue::String("3:00 PM".to_string())
        );
    }

    #[test]
    fn text_requires_numeric_value_and_string_format() {
        assert_eq!(eval("TEXT(12,0)"), CellValue::Error("#VALUE!".to_string()));
        assert_eq!(
            eval("TEXT(\"word\",\"0.00\")"),
            CellValue::Error("#VALUE!".to_string())
        );
    }
}
