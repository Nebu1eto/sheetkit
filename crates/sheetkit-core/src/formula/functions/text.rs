//! Text formula functions: CONCAT, FIND, SEARCH, SUBSTITUTE, REPLACE, REPT, EXACT, T, PROPER.

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_number, coerce_to_string, Evaluator};
use crate::formula::functions::conversion::checked_trunc_count;
use crate::formula::functions::{check_arg_count, wildcard_search};

/// CONCAT(text1, [text2], ...) - concatenates multiple values.
pub fn fn_concat(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("CONCAT", args, 1, 255)?;
    let mut result = String::new();
    for arg in args {
        let v = ctx.eval_expr(arg)?;
        result.push_str(&coerce_to_string(&v));
    }
    Ok(CellValue::String(result))
}

/// FIND(find_text, within_text, [start_num]) - case-sensitive search.
pub fn fn_find(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("FIND", args, 2, 3)?;
    let find_text = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let within_text = coerce_to_string(&ctx.eval_expr(&args[1])?);
    let start_value = if args.len() > 2 {
        coerce_to_number(&ctx.eval_expr(&args[2])?)?
    } else {
        1.0
    };
    let Some(start_num) = checked_trunc_count(start_value) else {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    };
    let char_count = within_text.chars().count();
    if start_num < 1 || start_num > char_count.saturating_add(1) {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let search_in = &within_text[character_position_to_byte_offset(&within_text, start_num)..];
    match search_in.find(&find_text) {
        Some(pos) => Ok(CellValue::Number(
            (start_num + search_in[..pos].chars().count()) as f64,
        )),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

/// SEARCH(find_text, within_text, [start_num]) - case-insensitive search with wildcards.
pub fn fn_search(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SEARCH", args, 2, 3)?;
    let find_text = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let within_text = coerce_to_string(&ctx.eval_expr(&args[1])?);
    let start_value = if args.len() > 2 {
        coerce_to_number(&ctx.eval_expr(&args[2])?)?
    } else {
        1.0
    };
    let Some(start_num) = checked_trunc_count(start_value) else {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    };
    let char_count = within_text.chars().count();
    if start_num < 1 || start_num > char_count.saturating_add(1) {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let search_in = &within_text[character_position_to_byte_offset(&within_text, start_num)..];
    match wildcard_search(search_in, &find_text, ctx)? {
        Some(char_offset) => Ok(CellValue::Number((char_offset + start_num) as f64)),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

fn character_position_to_byte_offset(text: &str, position: usize) -> usize {
    text.char_indices()
        .nth(position - 1)
        .map_or(text.len(), |(byte_offset, _)| byte_offset)
}

/// SUBSTITUTE(text, old_text, new_text, [instance_num]) - replaces occurrences.
pub fn fn_substitute(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SUBSTITUTE", args, 3, 4)?;
    let text = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let old_text = coerce_to_string(&ctx.eval_expr(&args[1])?);
    let new_text = coerce_to_string(&ctx.eval_expr(&args[2])?);
    if old_text.is_empty() {
        return Ok(CellValue::String(text));
    }
    if args.len() > 3 {
        let instance_value = coerce_to_number(&ctx.eval_expr(&args[3])?)?;
        let Some(instance_num) = checked_trunc_count(instance_value) else {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        };
        if instance_num < 1 {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
        let occurrence = text.match_indices(&old_text).nth(instance_num - 1);
        let replacement_count = usize::from(occurrence.is_some());
        let output_len = checked_substitute_output_len(
            text.len(),
            old_text.len(),
            new_text.len(),
            replacement_count,
        );
        let Some(output_len) = output_len else {
            ctx.check_output_bytes(u64::MAX)?;
            unreachable!("an overflowing substitute result must exceed the formula budget");
        };
        ctx.check_output_bytes(u64::try_from(output_len).unwrap_or(u64::MAX))?;
        let Some((position, _)) = occurrence else {
            return Ok(CellValue::String(text));
        };
        let mut result = String::new();
        result
            .try_reserve_exact(output_len)
            .map_err(|_| Error::Internal("substitute allocation failed".to_string()))?;
        result.push_str(&text[..position]);
        result.push_str(&new_text);
        result.push_str(&text[position + old_text.len()..]);
        Ok(CellValue::String(result))
    } else {
        let replacement_count = text.match_indices(&old_text).count();
        let Some(output_len) = checked_substitute_output_len(
            text.len(),
            old_text.len(),
            new_text.len(),
            replacement_count,
        ) else {
            ctx.check_output_bytes(u64::MAX)?;
            unreachable!("an overflowing substitute result must exceed the formula budget");
        };
        ctx.check_output_bytes(u64::try_from(output_len).unwrap_or(u64::MAX))?;
        let mut result = String::new();
        result
            .try_reserve_exact(output_len)
            .map_err(|_| Error::Internal("substitute allocation failed".to_string()))?;
        let mut copied_until = 0;
        for (position, matched) in text.match_indices(&old_text) {
            result.push_str(&text[copied_until..position]);
            result.push_str(&new_text);
            copied_until = position + matched.len();
        }
        result.push_str(&text[copied_until..]);
        Ok(CellValue::String(result))
    }
}

fn checked_substitute_output_len(
    text_len: usize,
    old_len: usize,
    new_len: usize,
    replacement_count: usize,
) -> Option<usize> {
    let removed = old_len.checked_mul(replacement_count)?;
    let added = new_len.checked_mul(replacement_count)?;
    text_len.checked_sub(removed)?.checked_add(added)
}

/// REPLACE(old_text, start_num, num_chars, new_text) - replaces by position.
pub fn fn_replace(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("REPLACE", args, 4, 4)?;
    let old_text = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let start_value = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let count_value = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let new_text = coerce_to_string(&ctx.eval_expr(&args[3])?);
    let (Some(start_num), Some(num_chars)) = (
        checked_trunc_count(start_value),
        checked_trunc_count(count_value),
    ) else {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    };
    if start_num < 1 {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let chars: Vec<char> = old_text.chars().collect();
    let start = start_num - 1;
    let end = start.saturating_add(num_chars).min(chars.len());
    let mut result = String::new();
    for &ch in &chars[..start.min(chars.len())] {
        result.push(ch);
    }
    result.push_str(&new_text);
    for &ch in &chars[end..] {
        result.push(ch);
    }
    Ok(CellValue::String(result))
}

/// REPT(text, number_times) - repeats text.
pub fn fn_rept(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("REPT", args, 2, 2)?;
    let text = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let times_value = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let Some(times) = checked_trunc_count(times_value) else {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    };
    if text
        .chars()
        .count()
        .checked_mul(times)
        .is_none_or(|length| length > 32_767)
    {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    Ok(CellValue::String(text.repeat(times)))
}

/// EXACT(text1, text2) - case-sensitive comparison.
pub fn fn_exact(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("EXACT", args, 2, 2)?;
    let text1 = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let text2 = coerce_to_string(&ctx.eval_expr(&args[1])?);
    Ok(CellValue::Bool(text1 == text2))
}

/// T(value) - returns text if value is text, empty string otherwise.
pub fn fn_t(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("T", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    match v {
        CellValue::String(s) => Ok(CellValue::String(s)),
        _ => Ok(CellValue::String(String::new())),
    }
}

/// PROPER(text) - capitalizes first letter of each word.
pub fn fn_proper(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("PROPER", args, 1, 1)?;
    let text = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let mut result = String::with_capacity(text.len());
    let mut capitalize_next = true;
    for ch in text.chars() {
        if ch.is_alphabetic() {
            if capitalize_next {
                result.extend(ch.to_uppercase());
                capitalize_next = false;
            } else {
                result.extend(ch.to_lowercase());
            }
        } else {
            result.push(ch);
            capitalize_next = true;
        }
    }
    Ok(CellValue::String(result))
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

    #[test]
    fn test_concat() {
        assert_eq!(
            eval(r#"CONCAT("Hello"," ","World")"#),
            CellValue::String("Hello World".to_string())
        );
    }

    #[test]
    fn test_find_basic() {
        assert_eq!(eval(r#"FIND("B","ABCABC")"#), CellValue::Number(2.0));
    }

    #[test]
    fn test_find_start_num() {
        assert_eq!(eval(r#"FIND("B","ABCABC",3)"#), CellValue::Number(5.0));
    }

    #[test]
    fn test_find_uses_character_positions_for_multibyte_text() {
        assert_eq!(eval(r#"FIND("나","가나다",2)"#), CellValue::Number(2.0));
    }

    #[test]
    fn test_find_not_found() {
        assert_eq!(
            eval(r#"FIND("Z","ABCABC")"#),
            CellValue::Error("#VALUE!".to_string())
        );
    }

    #[test]
    fn test_search_case_insensitive() {
        assert_eq!(eval(r#"SEARCH("b","ABCABC")"#), CellValue::Number(2.0));
    }

    #[test]
    fn test_search_wildcard() {
        assert_eq!(eval(r#"SEARCH("A*C","ABCABC")"#), CellValue::Number(1.0));
        assert_eq!(eval(r#"SEARCH("*","ABC",4)"#), CellValue::Number(4.0));
    }

    #[test]
    fn test_search_uses_character_positions_for_multibyte_text() {
        assert_eq!(eval(r#"SEARCH("나","가나다",2)"#), CellValue::Number(2.0));
    }

    #[test]
    fn test_substitute_all() {
        assert_eq!(
            eval(r#"SUBSTITUTE("aabbcc","b","X")"#),
            CellValue::String("aaXXcc".to_string())
        );
    }

    #[test]
    fn test_substitute_instance() {
        assert_eq!(
            eval(r#"SUBSTITUTE("aabbcc","b","X",2)"#),
            CellValue::String("aabXcc".to_string())
        );
    }

    #[test]
    fn substitute_output_length_detects_integer_overflow() {
        assert_eq!(
            checked_substitute_output_len(usize::MAX, 1, usize::MAX, 2),
            None
        );
    }

    #[test]
    fn test_replace() {
        assert_eq!(
            eval(r#"REPLACE("ABCDEF",3,2,"XY")"#),
            CellValue::String("ABXYEF".to_string())
        );
    }

    #[test]
    fn test_rept() {
        assert_eq!(
            eval(r#"REPT("AB",3)"#),
            CellValue::String("ABABAB".to_string())
        );
    }

    #[test]
    fn test_rept_rejects_results_over_excel_character_limit() {
        assert_eq!(
            eval(r#"REPT("AB",16384)"#),
            CellValue::Error("#VALUE!".to_string())
        );
    }

    #[test]
    fn test_exact_true() {
        assert_eq!(eval(r#"EXACT("ABC","ABC")"#), CellValue::Bool(true));
    }

    #[test]
    fn test_exact_false_case() {
        assert_eq!(eval(r#"EXACT("ABC","abc")"#), CellValue::Bool(false));
    }

    #[test]
    fn test_t_with_text() {
        assert_eq!(
            eval(r#"T("hello")"#),
            CellValue::String("hello".to_string())
        );
    }

    #[test]
    fn test_t_with_number() {
        assert_eq!(eval("T(42)"), CellValue::String(String::new()));
    }

    #[test]
    fn test_proper() {
        assert_eq!(
            eval(r#"PROPER("hello world")"#),
            CellValue::String("Hello World".to_string())
        );
    }
}
