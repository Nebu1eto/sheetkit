//! Text formula functions: CONCAT, FIND, SEARCH, SUBSTITUTE, REPLACE, REPT, EXACT, T, PROPER.

use crate::cell::CellValue;
use crate::error::Result;
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_number, coerce_to_string, Evaluator};
use crate::formula::functions::check_arg_count;

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
    let start_num = if args.len() > 2 {
        coerce_to_number(&ctx.eval_expr(&args[2])?)? as usize
    } else {
        1
    };
    if start_num < 1 || start_num > within_text.len() + 1 {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let search_in = &within_text[(start_num - 1)..];
    match search_in.find(&find_text) {
        Some(pos) => Ok(CellValue::Number((pos + start_num) as f64)),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

/// SEARCH(find_text, within_text, [start_num]) - case-insensitive search with wildcards.
pub fn fn_search(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SEARCH", args, 2, 3)?;
    let find_text = coerce_to_string(&ctx.eval_expr(&args[0])?).to_ascii_lowercase();
    let within_text = coerce_to_string(&ctx.eval_expr(&args[1])?).to_ascii_lowercase();
    let start_num = if args.len() > 2 {
        coerce_to_number(&ctx.eval_expr(&args[2])?)? as usize
    } else {
        1
    };
    if start_num < 1 || start_num > within_text.len() + 1 {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let search_in = &within_text[(start_num - 1)..];
    if find_text.contains('*') || find_text.contains('?') {
        for i in 0..=search_in.len() {
            let substring = &search_in[i..];
            if wildcard_match_prefix(&find_text, substring) {
                return Ok(CellValue::Number((i + start_num) as f64));
            }
        }
        Ok(CellValue::Error("#VALUE!".to_string()))
    } else {
        match search_in.find(&find_text) {
            Some(pos) => Ok(CellValue::Number((pos + start_num) as f64)),
            None => Ok(CellValue::Error("#VALUE!".to_string())),
        }
    }
}

/// Wildcard match where the pattern must be fully consumed but the text may have trailing chars.
fn wildcard_match_prefix(pattern: &str, text: &str) -> bool {
    let p: Vec<char> = pattern.chars().collect();
    let t: Vec<char> = text.chars().collect();
    let mut pi = 0;
    let mut ti = 0;
    let mut star_pi = usize::MAX;
    let mut star_ti = 0;
    while pi < p.len() {
        if pi < p.len() && p[pi] == '*' {
            star_pi = pi;
            star_ti = ti;
            pi += 1;
        } else if ti < t.len() && (p[pi] == '?' || p[pi] == t[ti]) {
            pi += 1;
            ti += 1;
        } else if star_pi != usize::MAX {
            pi = star_pi + 1;
            star_ti += 1;
            ti = star_ti;
            if ti > t.len() {
                return false;
            }
        } else {
            return false;
        }
    }
    true
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
        let instance_num = coerce_to_number(&ctx.eval_expr(&args[3])?)? as usize;
        if instance_num < 1 {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
        let mut count = 0;
        let mut result = String::new();
        let mut remaining = text.as_str();
        while let Some(pos) = remaining.find(&old_text) {
            count += 1;
            if count == instance_num {
                result.push_str(&remaining[..pos]);
                result.push_str(&new_text);
                result.push_str(&remaining[pos + old_text.len()..]);
                return Ok(CellValue::String(result));
            }
            result.push_str(&remaining[..pos + old_text.len()]);
            remaining = &remaining[pos + old_text.len()..];
        }
        result.push_str(remaining);
        Ok(CellValue::String(result))
    } else {
        Ok(CellValue::String(text.replace(&old_text, &new_text)))
    }
}

/// REPLACE(old_text, start_num, num_chars, new_text) - replaces by position.
pub fn fn_replace(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("REPLACE", args, 4, 4)?;
    let old_text = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let start_num = coerce_to_number(&ctx.eval_expr(&args[1])?)? as usize;
    let num_chars = coerce_to_number(&ctx.eval_expr(&args[2])?)? as usize;
    let new_text = coerce_to_string(&ctx.eval_expr(&args[3])?);
    if start_num < 1 {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let chars: Vec<char> = old_text.chars().collect();
    let start = start_num - 1;
    let end = (start + num_chars).min(chars.len());
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
    let times = coerce_to_number(&ctx.eval_expr(&args[1])?)? as usize;
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
