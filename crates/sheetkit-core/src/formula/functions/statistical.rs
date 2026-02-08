//! Statistical formula functions: AVERAGEIF, AVERAGEIFS, COUNTBLANK, COUNTIF,
//! COUNTIFS, MEDIAN, MODE, LARGE, SMALL, RANK.

use crate::cell::CellValue;
use crate::error::Result;
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_number, coerce_to_string, Evaluator};
use crate::formula::functions::{check_arg_count, collect_criteria_range_values, matches_criteria};

/// AVERAGEIF(range, criteria, [average_range])
pub fn fn_averageif(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("AVERAGEIF", args, 2, 3)?;
    let range_vals = collect_criteria_range_values(&args[0], ctx)?;
    let criteria_val = ctx.eval_expr(&args[1])?;
    let criteria = coerce_to_string(&criteria_val);
    let avg_vals = if args.len() == 3 {
        collect_criteria_range_values(&args[2], ctx)?
    } else {
        range_vals.clone()
    };
    let mut sum = 0.0;
    let mut count = 0u64;
    for (i, rv) in range_vals.iter().enumerate() {
        if matches_criteria(rv, &criteria) {
            if let Some(sv) = avg_vals.get(i) {
                if let Ok(n) = coerce_to_number(sv) {
                    sum += n;
                    count += 1;
                }
            }
        }
    }
    if count == 0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }
    Ok(CellValue::Number(sum / count as f64))
}

/// AVERAGEIFS(average_range, criteria_range1, criteria1, ...)
pub fn fn_averageifs(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("AVERAGEIFS", args, 3, 255)?;
    if !(args.len() - 1).is_multiple_of(2) {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let avg_vals = collect_criteria_range_values(&args[0], ctx)?;
    let pair_count = (args.len() - 1) / 2;
    let mut criteria_ranges: Vec<Vec<CellValue>> = Vec::with_capacity(pair_count);
    let mut criteria_strings: Vec<String> = Vec::with_capacity(pair_count);
    for i in 0..pair_count {
        let range_vals = collect_criteria_range_values(&args[1 + i * 2], ctx)?;
        let crit_val = ctx.eval_expr(&args[2 + i * 2])?;
        criteria_ranges.push(range_vals);
        criteria_strings.push(coerce_to_string(&crit_val));
    }
    let mut sum = 0.0;
    let mut count = 0u64;
    for (idx, sv) in avg_vals.iter().enumerate() {
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
                sum += n;
                count += 1;
            }
        }
    }
    if count == 0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }
    Ok(CellValue::Number(sum / count as f64))
}

/// COUNTBLANK(range) - count empty cells in a range
pub fn fn_countblank(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("COUNTBLANK", args, 1, 1)?;
    let values = collect_criteria_range_values(&args[0], ctx)?;
    let count = values
        .iter()
        .filter(|v| matches!(v, CellValue::Empty))
        .count();
    Ok(CellValue::Number(count as f64))
}

/// COUNTIF(range, criteria)
pub fn fn_countif(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("COUNTIF", args, 2, 2)?;
    let range_vals = collect_criteria_range_values(&args[0], ctx)?;
    let criteria_val = ctx.eval_expr(&args[1])?;
    let criteria = coerce_to_string(&criteria_val);
    let count = range_vals
        .iter()
        .filter(|rv| matches_criteria(rv, &criteria))
        .count();
    Ok(CellValue::Number(count as f64))
}

/// COUNTIFS(range1, criteria1, range2, criteria2, ...)
pub fn fn_countifs(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("COUNTIFS", args, 2, 255)?;
    if !args.len().is_multiple_of(2) {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let pair_count = args.len() / 2;
    let mut criteria_ranges: Vec<Vec<CellValue>> = Vec::with_capacity(pair_count);
    let mut criteria_strings: Vec<String> = Vec::with_capacity(pair_count);
    for i in 0..pair_count {
        let range_vals = collect_criteria_range_values(&args[i * 2], ctx)?;
        let crit_val = ctx.eval_expr(&args[i * 2 + 1])?;
        criteria_ranges.push(range_vals);
        criteria_strings.push(coerce_to_string(&crit_val));
    }
    let len = criteria_ranges.first().map_or(0, |r| r.len());
    let mut count = 0u64;
    for idx in 0..len {
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
            count += 1;
        }
    }
    Ok(CellValue::Number(count as f64))
}

/// MEDIAN(args...) - median of numeric values
pub fn fn_median(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("MEDIAN", args, 1, 255)?;
    let mut nums = ctx.collect_numbers(args)?;
    if nums.is_empty() {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let n = nums.len();
    let median = if n % 2 == 1 {
        nums[n / 2]
    } else {
        (nums[n / 2 - 1] + nums[n / 2]) / 2.0
    };
    Ok(CellValue::Number(median))
}

/// MODE(args...) - most frequently occurring value
pub fn fn_mode(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("MODE", args, 1, 255)?;
    let nums = ctx.collect_numbers(args)?;
    if nums.is_empty() {
        return Ok(CellValue::Error("#N/A".to_string()));
    }
    let mut counts: std::collections::HashMap<u64, (f64, usize)> = std::collections::HashMap::new();
    for &n in &nums {
        let key = n.to_bits();
        let entry = counts.entry(key).or_insert((n, 0));
        entry.1 += 1;
    }
    let max_count = counts.values().map(|(_, c)| *c).max().unwrap_or(0);
    if max_count <= 1 {
        return Ok(CellValue::Error("#N/A".to_string()));
    }
    // Return the first value that has max_count occurrences (in order of appearance)
    for &n in &nums {
        let key = n.to_bits();
        if let Some((_, c)) = counts.get(&key) {
            if *c == max_count {
                return Ok(CellValue::Number(n));
            }
        }
    }
    Ok(CellValue::Error("#N/A".to_string()))
}

/// LARGE(array, k) - k-th largest value
pub fn fn_large(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("LARGE", args, 2, 2)?;
    let mut nums = collect_criteria_range_values(&args[0], ctx)?
        .iter()
        .filter_map(|v| coerce_to_number(v).ok())
        .collect::<Vec<f64>>();
    let k = coerce_to_number(&ctx.eval_expr(&args[1])?)? as usize;
    if k == 0 || k > nums.len() {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    nums.sort_by(|a, b| b.partial_cmp(a).unwrap_or(std::cmp::Ordering::Equal));
    Ok(CellValue::Number(nums[k - 1]))
}

/// SMALL(array, k) - k-th smallest value
pub fn fn_small(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SMALL", args, 2, 2)?;
    let mut nums = collect_criteria_range_values(&args[0], ctx)?
        .iter()
        .filter_map(|v| coerce_to_number(v).ok())
        .collect::<Vec<f64>>();
    let k = coerce_to_number(&ctx.eval_expr(&args[1])?)? as usize;
    if k == 0 || k > nums.len() {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    Ok(CellValue::Number(nums[k - 1]))
}

/// RANK(number, ref, [order]) - rank of a number in a list.
/// order=0 or omitted: descending (largest=1). order=nonzero: ascending (smallest=1).
pub fn fn_rank(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("RANK", args, 2, 3)?;
    let number = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let ref_vals = collect_criteria_range_values(&args[1], ctx)?;
    let order = if args.len() > 2 {
        coerce_to_number(&ctx.eval_expr(&args[2])?)? as i64
    } else {
        0
    };
    let nums: Vec<f64> = ref_vals
        .iter()
        .filter_map(|v| coerce_to_number(v).ok())
        .collect();
    let rank = if order == 0 {
        // Descending: count how many are greater than number, +1
        nums.iter().filter(|&&n| n > number).count() + 1
    } else {
        // Ascending: count how many are less than number, +1
        nums.iter().filter(|&&n| n < number).count() + 1
    };
    // If the number is not in the list, return #N/A
    if !nums.iter().any(|&n| (n - number).abs() < f64::EPSILON) {
        return Ok(CellValue::Error("#N/A".to_string()));
    }
    Ok(CellValue::Number(rank as f64))
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

    // AVERAGEIF tests

    #[test]
    fn averageif_greater_than() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(2.0)),
            ("Sheet1", 1, 2, CellValue::Number(4.0)),
            ("Sheet1", 1, 3, CellValue::Number(6.0)),
        ];
        let result = eval_with_data("AVERAGEIF(A1:A3,\">3\")", &data);
        assert_eq!(result, CellValue::Number(5.0));
    }

    #[test]
    fn averageif_with_avg_range() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::String("A".to_string())),
            ("Sheet1", 2, 1, CellValue::Number(10.0)),
            ("Sheet1", 1, 2, CellValue::String("B".to_string())),
            ("Sheet1", 2, 2, CellValue::Number(20.0)),
            ("Sheet1", 1, 3, CellValue::String("A".to_string())),
            ("Sheet1", 2, 3, CellValue::Number(30.0)),
        ];
        let result = eval_with_data("AVERAGEIF(A1:A3,\"A\",B1:B3)", &data);
        assert_eq!(result, CellValue::Number(20.0));
    }

    #[test]
    fn averageif_no_match() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(1.0)),
            ("Sheet1", 1, 2, CellValue::Number(2.0)),
        ];
        let result = eval_with_data("AVERAGEIF(A1:A2,\">100\")", &data);
        assert_eq!(result, CellValue::Error("#DIV/0!".to_string()));
    }

    // AVERAGEIFS test

    #[test]
    fn averageifs_multi_criteria() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::String("A".to_string())),
            ("Sheet1", 2, 1, CellValue::Number(1.0)),
            ("Sheet1", 3, 1, CellValue::Number(100.0)),
            ("Sheet1", 1, 2, CellValue::String("A".to_string())),
            ("Sheet1", 2, 2, CellValue::Number(5.0)),
            ("Sheet1", 3, 2, CellValue::Number(200.0)),
            ("Sheet1", 1, 3, CellValue::String("B".to_string())),
            ("Sheet1", 2, 3, CellValue::Number(3.0)),
            ("Sheet1", 3, 3, CellValue::Number(300.0)),
        ];
        let result = eval_with_data("AVERAGEIFS(C1:C3,A1:A3,\"A\",B1:B3,\">2\")", &data);
        assert_eq!(result, CellValue::Number(200.0));
    }

    // COUNTBLANK tests

    #[test]
    fn countblank_basic() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(1.0)),
            ("Sheet1", 1, 2, CellValue::Empty),
            ("Sheet1", 1, 3, CellValue::String("x".to_string())),
            ("Sheet1", 1, 4, CellValue::Empty),
        ];
        let result = eval_with_data("COUNTBLANK(A1:A4)", &data);
        assert_eq!(result, CellValue::Number(2.0));
    }

    // COUNTIF tests

    #[test]
    fn countif_greater_than() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(1.0)),
            ("Sheet1", 1, 2, CellValue::Number(5.0)),
            ("Sheet1", 1, 3, CellValue::Number(10.0)),
        ];
        let result = eval_with_data("COUNTIF(A1:A3,\">3\")", &data);
        assert_eq!(result, CellValue::Number(2.0));
    }

    #[test]
    fn countif_text_match() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::String("Apple".to_string())),
            ("Sheet1", 1, 2, CellValue::String("Banana".to_string())),
            ("Sheet1", 1, 3, CellValue::String("apple".to_string())),
        ];
        let result = eval_with_data("COUNTIF(A1:A3,\"Apple\")", &data);
        assert_eq!(result, CellValue::Number(2.0));
    }

    // COUNTIFS test

    #[test]
    fn countifs_multi_criteria() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::String("A".to_string())),
            ("Sheet1", 2, 1, CellValue::Number(10.0)),
            ("Sheet1", 1, 2, CellValue::String("A".to_string())),
            ("Sheet1", 2, 2, CellValue::Number(20.0)),
            ("Sheet1", 1, 3, CellValue::String("B".to_string())),
            ("Sheet1", 2, 3, CellValue::Number(30.0)),
        ];
        let result = eval_with_data("COUNTIFS(A1:A3,\"A\",B1:B3,\">15\")", &data);
        assert_eq!(result, CellValue::Number(1.0));
    }

    // MEDIAN tests

    #[test]
    fn median_odd_count() {
        assert_eq!(eval("MEDIAN(1,3,2)"), CellValue::Number(2.0));
    }

    #[test]
    fn median_even_count() {
        assert_eq!(eval("MEDIAN(1,2,3,4)"), CellValue::Number(2.5));
    }

    #[test]
    fn median_single() {
        assert_eq!(eval("MEDIAN(42)"), CellValue::Number(42.0));
    }

    // MODE tests

    #[test]
    fn mode_basic() {
        assert_eq!(eval("MODE(1,2,2,3,3,3)"), CellValue::Number(3.0));
    }

    #[test]
    fn mode_no_repeat() {
        assert_eq!(eval("MODE(1,2,3)"), CellValue::Error("#N/A".to_string()));
    }

    // LARGE tests

    #[test]
    fn large_basic() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(5.0)),
            ("Sheet1", 1, 2, CellValue::Number(3.0)),
            ("Sheet1", 1, 3, CellValue::Number(8.0)),
            ("Sheet1", 1, 4, CellValue::Number(1.0)),
        ];
        let result = eval_with_data("LARGE(A1:A4,2)", &data);
        assert_eq!(result, CellValue::Number(5.0));
    }

    #[test]
    fn large_k_out_of_range() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(1.0)),
            ("Sheet1", 1, 2, CellValue::Number(2.0)),
        ];
        let result = eval_with_data("LARGE(A1:A2,5)", &data);
        assert_eq!(result, CellValue::Error("#NUM!".to_string()));
    }

    // SMALL tests

    #[test]
    fn small_basic() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(5.0)),
            ("Sheet1", 1, 2, CellValue::Number(3.0)),
            ("Sheet1", 1, 3, CellValue::Number(8.0)),
            ("Sheet1", 1, 4, CellValue::Number(1.0)),
        ];
        let result = eval_with_data("SMALL(A1:A4,2)", &data);
        assert_eq!(result, CellValue::Number(3.0));
    }

    // RANK tests

    #[test]
    fn rank_descending() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(5.0)),
            ("Sheet1", 1, 2, CellValue::Number(3.0)),
            ("Sheet1", 1, 3, CellValue::Number(8.0)),
            ("Sheet1", 1, 4, CellValue::Number(1.0)),
        ];
        // 5 is 2nd largest in [5,3,8,1]
        let result = eval_with_data("RANK(5,A1:A4)", &data);
        assert_eq!(result, CellValue::Number(2.0));
    }

    #[test]
    fn rank_ascending() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(5.0)),
            ("Sheet1", 1, 2, CellValue::Number(3.0)),
            ("Sheet1", 1, 3, CellValue::Number(8.0)),
            ("Sheet1", 1, 4, CellValue::Number(1.0)),
        ];
        // 5 is 3rd smallest in [5,3,8,1]
        let result = eval_with_data("RANK(5,A1:A4,1)", &data);
        assert_eq!(result, CellValue::Number(3.0));
    }

    #[test]
    fn rank_not_found() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(1.0)),
            ("Sheet1", 1, 2, CellValue::Number(2.0)),
        ];
        let result = eval_with_data("RANK(99,A1:A2)", &data);
        assert_eq!(result, CellValue::Error("#N/A".to_string()));
    }
}
