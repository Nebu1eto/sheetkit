//! Regression tests for formula-engine and number-format audit findings.

use std::panic::{catch_unwind, AssertUnwindSafe};
use std::path::Path;
use std::time::{Duration, Instant};

use sheetkit_core::cell::CellValue;
use sheetkit_core::error::Error;
use sheetkit_core::formula::budget::{FormulaBudget, FormulaLimits};
use sheetkit_core::formula::eval::{evaluate, CellSnapshot, Evaluator};
use sheetkit_core::formula::parser::parse_formula;
use sheetkit_core::numfmt::format_number;

fn eval_in(snapshot: &CellSnapshot, formula: &str) -> sheetkit_core::error::Result<CellValue> {
    let expr = parse_formula(formula).expect("formula must parse");
    catch_unwind(AssertUnwindSafe(|| evaluate(&expr, snapshot)))
        .unwrap_or_else(|_| panic!("evaluating {formula} panicked"))
}

fn eval(formula: &str) -> CellValue {
    let snapshot = CellSnapshot::new("Sheet1".to_string());
    eval_in(&snapshot, formula).unwrap_or_else(|error| panic!("{formula}: {error}"))
}

fn error(code: &str) -> CellValue {
    CellValue::Error(code.to_string())
}

fn number(value: CellValue) -> f64 {
    match value {
        CellValue::Number(n) | CellValue::Date(n) => n,
        other => panic!("expected a number, got {other:?}"),
    }
}

// F-F1: exponent parsing must not mix uppercased and original byte offsets.

#[test]
fn text_with_multibyte_scientific_format_does_not_panic() {
    let value = eval("TEXT(1,\"\u{149}E+\u{e9}\")");
    assert!(matches!(value, CellValue::String(_)), "{value:?}");
}

#[test]
fn format_number_with_case_expanding_prefix_does_not_panic() {
    for format in [
        "\u{149}E+\u{e9}",
        "\u{149}0.00E+00",
        "\u{df}0.0E+000",
        "\u{130}0E-0",
        "\u{149}\u{149}\u{149}E+",
        "0.0e+00\u{149}",
    ] {
        let result = catch_unwind(|| format_number(12345.0, format));
        assert!(result.is_ok(), "format_number panicked for {format:?}");
    }
}

#[test]
fn date_formats_with_out_of_range_serials_do_not_panic() {
    for value in [1e300, -1e300, 2_958_466.0, f64::MAX] {
        for format in ["yyyy-mm-dd", "hh:mm:ss", "m/d/yy h:mm AM/PM"] {
            let result = catch_unwind(|| format_number(value, format));
            assert!(
                result.is_ok(),
                "format_number({value}, {format:?}) panicked"
            );
        }
    }
    for formula in [
        "TEXT(1E300,\"yyyy\")",
        "YEAR(1E300)",
        "HOUR(1E300)",
        "WEEKDAY(1E300)",
    ] {
        eval(formula);
    }
}

#[test]
fn exponent_digit_count_uses_original_format_positions() {
    // U+00DF uppercases to "SS", which shifted the old offset by one byte.
    let shifted = format_number(12345.0, "\u{df}0.0E+000");
    assert!(shifted.ends_with("E+004"), "{shifted}");
    let plain = format_number(12345.0, "0.0E+000");
    assert_eq!(plain, "1.2E+004");
}

// F-F2: date constructors and arithmetic must reject out-of-domain values.

#[test]
fn date_with_overflowing_day_returns_num_error() {
    assert_eq!(eval("DATE(2024,1,2147483647)"), error("#NUM!"));
    assert_eq!(eval("DATE(2024,1,-2147483648)"), error("#NUM!"));
    assert_eq!(eval("DATE(2024,1,1E300)"), error("#NUM!"));
    assert_eq!(eval("DATE(2024,2147483647,1)"), error("#NUM!"));
    assert_eq!(eval("DATE(1E300,1,1)"), error("#NUM!"));
}

#[test]
fn date_in_range_arguments_are_unchanged() {
    assert_eq!(number(eval("DATE(2024,1,15)")), 45306.0);
    assert_eq!(number(eval("DATE(2024,13,1)")), 45658.0);
    assert_eq!(number(eval("DATE(2024,1,0)")), 45291.0);
    assert_eq!(number(eval("DATE(2024,3,-1)")), 45350.0);
}

#[test]
fn workday_with_huge_day_count_returns_num_error() {
    assert_eq!(eval("WORKDAY(45000,1E17)"), error("#NUM!"));
    assert_eq!(eval("WORKDAY(45000,-1E17)"), error("#NUM!"));
    assert_eq!(eval("WORKDAY(45000,1E300)"), error("#NUM!"));
}

#[test]
fn edate_and_eomonth_with_huge_month_count_return_num_error() {
    for formula in [
        "EDATE(45000,2147483647)",
        "EDATE(45000,-2147483648)",
        "EDATE(45000,1E300)",
        "EOMONTH(45000,2147483647)",
        "EOMONTH(45000,-1E300)",
    ] {
        assert_eq!(eval(formula), error("#NUM!"), "{formula}");
    }
    assert_eq!(number(eval("EDATE(DATE(2024,1,31),1)")), 45351.0);
}

#[test]
fn integer_arguments_derived_from_formulas_never_panic() {
    let huge = ["1E300", "-1E300", "2147483648", "-2147483649", "1E19"];
    let templates = [
        "WEEKDAY(45000,{})",
        "WEEKNUM(45000,{})",
        "LEFT(\"abc\",{})",
        "RIGHT(\"abc\",{})",
        "MID(\"abc\",{},2)",
        "MID(\"abc\",1,{})",
        "FIND(\"b\",\"abc\",{})",
        "REPT(\"a\",{})",
        "REPLACE(\"abc\",{},1,\"x\")",
        "SUBSTITUTE(\"aaa\",\"a\",\"b\",{})",
        "ROUND(1.5,{})",
        "ROUNDUP(1.5,{})",
        "ROUNDDOWN(1.5,{})",
        "CHOOSE({},1,2)",
        "ADDRESS({},1)",
        "ADDRESS(1,{})",
        "ISEVEN({})",
        "ISODD({})",
        "LARGE(5,{})",
        "SMALL(5,{})",
        "RANK(1,1,{})",
        "EFFECT(0.1,{})",
        "NOMINAL(0.1,{})",
        "DOLLARDE(1.5,{})",
        "DOLLARFR(1.5,{})",
        "FV(0.1,1,1,1,{})",
        "PV(0.1,1,1,1,{})",
        "PMT(0.1,1,1,1,{})",
        "NPER(0.1,1,1,1,{})",
        "BESSELI(1,{})",
        "BESSELJ(1,{})",
        "RANDBETWEEN({},1)",
        "RANDBETWEEN(1,{})",
    ];
    let snapshot = CellSnapshot::new("Sheet1".to_string());
    for template in templates {
        for value in huge {
            let formula = template.replace("{}", value);
            let result = eval_in(&snapshot, &formula);
            if let Ok(CellValue::Number(n)) = result {
                assert!(n.is_finite(), "{formula} returned non-finite {n}");
            }
        }
    }
}

// F-F6: REPLACE must not overflow when counts saturate; Excel rejects counts
// above 32767.

#[test]
fn replace_with_huge_count_does_not_overflow() {
    assert_eq!(eval("REPLACE(\"abc\",2,1E30,\"x\")"), error("#VALUE!"));
    assert_eq!(eval("REPLACE(\"abc\",1E30,1E30,\"x\")"), error("#VALUE!"));
    assert_eq!(
        eval("REPLACE(\"abc\",2,32767,\"x\")"),
        CellValue::String("ax".to_string())
    );
    assert_eq!(eval("REPLACE(\"abc\",2,-1,\"x\")"), error("#VALUE!"));
    assert_eq!(
        eval("REPLACE(\"abc\",2,1,\"x\")"),
        CellValue::String("axc".to_string())
    );
}

#[test]
fn text_slicing_with_huge_or_negative_counts() {
    assert_eq!(eval("LEFT(\"abc\",1E30)"), CellValue::String("abc".into()));
    assert_eq!(eval("RIGHT(\"abc\",1E30)"), CellValue::String("abc".into()));
    assert_eq!(eval("MID(\"abc\",2,1E30)"), error("#VALUE!"));
    assert_eq!(eval("MID(\"abc\",1E30,1)"), error("#VALUE!"));
    assert_eq!(eval("LEFT(\"abc\",-1)"), error("#VALUE!"));
    assert_eq!(eval("RIGHT(\"abc\",-1)"), error("#VALUE!"));
    assert_eq!(eval("MID(\"abc\",1,-1)"), error("#VALUE!"));
    assert_eq!(eval("REPT(\"a\",-1)"), error("#VALUE!"));
    assert_eq!(eval("FIND(\"b\",\"abc\",1E30)"), error("#VALUE!"));
}

// F-F7: rounding must never produce NaN or infinity for finite inputs.

#[test]
fn round_with_digits_beyond_f64_precision_keeps_value() {
    assert_eq!(eval("ROUND(1,309)"), CellValue::Number(1.0));
    assert_eq!(eval("ROUND(1.25,400)"), CellValue::Number(1.25));
    assert_eq!(eval("ROUNDUP(1.25,400)"), CellValue::Number(1.25));
    assert_eq!(eval("ROUNDDOWN(1.25,400)"), CellValue::Number(1.25));
    assert_eq!(eval("ROUND(123.456,-400)"), CellValue::Number(0.0));
    assert_eq!(eval("ROUNDDOWN(123.456,-400)"), CellValue::Number(0.0));
    assert_eq!(eval("ROUNDUP(123.456,-400)"), error("#NUM!"));
    assert_eq!(eval("ROUND(1.7E308,-308)"), error("#NUM!"));
}

#[test]
fn round_in_range_digits_are_unchanged() {
    assert_eq!(eval("ROUND(2.345,2)"), CellValue::Number(2.35));
    assert_eq!(eval("ROUND(1234.5,-2)"), CellValue::Number(1200.0));
    assert_eq!(eval("ROUND(-2.5,0)"), CellValue::Number(-3.0));
    assert_eq!(eval("ROUNDUP(1.23456,2)"), CellValue::Number(1.24));
    assert_eq!(eval("ROUNDDOWN(-1.23456,2)"), CellValue::Number(-1.23));
    assert_eq!(eval("ROUND(1.5,2.9)"), CellValue::Number(1.5));
}

// F-F4: referenced formulas evaluate in their own sheet context.

fn two_sheet_snapshot() -> CellSnapshot {
    let mut snapshot = CellSnapshot::new("Sheet1".to_string());
    snapshot.set_cell("Sheet1", 1, 1, CellValue::Number(11.0));
    snapshot.set_cell("Sheet2", 1, 1, CellValue::Number(22.0));
    snapshot.set_cell("Sheet3", 1, 1, CellValue::Number(33.0));
    snapshot.set_cell(
        "Sheet2",
        2,
        1,
        CellValue::Formula {
            expr: "A1".to_string(),
            result: None,
        },
    );
    snapshot
}

fn formula(expr: &str) -> CellValue {
    CellValue::Formula {
        expr: expr.to_string(),
        result: None,
    }
}

#[test]
fn cross_sheet_formula_resolves_unqualified_refs_in_its_own_sheet() {
    let snapshot = two_sheet_snapshot();
    assert_eq!(
        eval_in(&snapshot, "Sheet2!B1").unwrap(),
        CellValue::Number(22.0)
    );
    assert_eq!(
        eval_in(&snapshot, "A1+Sheet2!B1+A1").unwrap(),
        CellValue::Number(44.0)
    );
    assert_eq!(
        eval_in(&snapshot, "SUM(Sheet2!B1:B1)").unwrap(),
        CellValue::Number(22.0)
    );
}

#[test]
fn nested_cross_sheet_chain_restores_each_context() {
    let mut snapshot = two_sheet_snapshot();
    // Sheet1!C1 -> Sheet2!C1 -> Sheet3!C1, each adding its own A1.
    snapshot.set_cell("Sheet3", 3, 1, formula("A1"));
    snapshot.set_cell("Sheet2", 3, 1, formula("Sheet3!C1+A1"));
    snapshot.set_cell("Sheet1", 3, 1, formula("Sheet2!C1+A1"));
    assert_eq!(eval_in(&snapshot, "C1").unwrap(), CellValue::Number(66.0));
    assert_eq!(
        eval_in(&snapshot, "C1+A1").unwrap(),
        CellValue::Number(77.0)
    );
    // Unqualified ranges inside the referenced formula use its sheet too.
    snapshot.set_cell("Sheet2", 4, 1, formula("SUM(A1:A1)"));
    assert_eq!(
        eval_in(&snapshot, "Sheet2!D1").unwrap(),
        CellValue::Number(22.0)
    );
}

#[test]
fn context_is_restored_after_an_error_in_a_referenced_formula() {
    let mut snapshot = two_sheet_snapshot();
    snapshot.set_cell("Sheet2", 5, 1, formula("NOSUCHFUNCTION(1)"));
    let expr = parse_formula("Sheet2!E1").unwrap();
    let mut evaluator = Evaluator::new(&snapshot);
    assert!(evaluator.eval_expr(&expr).is_err());
    assert_eq!(evaluator.current_sheet(), "Sheet1");
    let expr = parse_formula("A1").unwrap();
    assert_eq!(evaluator.eval_expr(&expr).unwrap(), CellValue::Number(11.0));
}

#[test]
fn genuine_cross_sheet_cycle_is_detected() {
    let mut snapshot = two_sheet_snapshot();
    snapshot.set_cell("Sheet1", 6, 1, formula("Sheet2!F1"));
    snapshot.set_cell("Sheet2", 6, 1, formula("Sheet1!F1"));
    assert!(matches!(
        eval_in(&snapshot, "F1"),
        Err(Error::CircularReference { .. })
    ));
    // Unqualified self-reference in another sheet is a cycle on that sheet.
    snapshot.set_cell("Sheet2", 7, 1, formula("G1"));
    assert!(matches!(
        eval_in(&snapshot, "Sheet2!G1"),
        Err(Error::CircularReference { .. })
    ));
}

#[test]
fn same_address_on_different_sheets_is_not_a_cycle() {
    let mut snapshot = two_sheet_snapshot();
    // Sheet1!H1 -> Sheet2!H1 whose unqualified H2 refers to Sheet2!H2.
    snapshot.set_cell("Sheet1", 8, 1, formula("Sheet2!H1"));
    snapshot.set_cell("Sheet2", 8, 1, formula("H2*2"));
    snapshot.set_cell("Sheet2", 8, 2, CellValue::Number(4.0));
    snapshot.set_cell("Sheet1", 8, 2, formula("H1"));
    assert_eq!(eval_in(&snapshot, "H1").unwrap(), CellValue::Number(8.0));
    assert_eq!(eval_in(&snapshot, "H2").unwrap(), CellValue::Number(8.0));
}

// F-F5: numeric aggregates over references only use numeric cells.

fn aggregate_snapshot() -> CellSnapshot {
    let mut snapshot = CellSnapshot::new("Sheet1".to_string());
    snapshot.set_cell("Sheet1", 1, 1, CellValue::Number(11.0));
    // A2 empty, A3 TRUE, A4 numeric text, A5 date, A6 text.
    snapshot.set_cell("Sheet1", 1, 3, CellValue::Bool(true));
    snapshot.set_cell("Sheet1", 1, 4, CellValue::String("5".to_string()));
    snapshot.set_cell("Sheet1", 1, 5, CellValue::Date(3.0));
    snapshot.set_cell("Sheet1", 1, 6, CellValue::String("x".to_string()));
    snapshot
}

#[test]
fn average_over_range_skips_empty_bool_and_text_cells() {
    let snapshot = aggregate_snapshot();
    let cases = [
        ("AVERAGE(A1:A2)", CellValue::Number(11.0)),
        ("AVERAGE(A1:A6)", CellValue::Number(7.0)),
        ("SUM(A1:A6)", CellValue::Number(14.0)),
        ("MIN(A1:A6)", CellValue::Number(3.0)),
        ("MAX(A2:A4)", CellValue::Number(0.0)),
        ("PRODUCT(A1:A6)", CellValue::Number(33.0)),
        ("MEDIAN(A1:A6)", CellValue::Number(7.0)),
        ("COUNT(A1:A6)", CellValue::Number(2.0)),
        ("AVERAGE(A2)", error("#DIV/0!")),
        ("AVERAGE(A2:A4)", error("#DIV/0!")),
        ("SUM(A3)", CellValue::Number(0.0)),
        ("SUM(A6)", CellValue::Number(0.0)),
        ("LARGE(A1:A6,2)", CellValue::Number(3.0)),
        ("SMALL(A1:A6,1)", CellValue::Number(3.0)),
        ("RANK(3,A1:A6)", CellValue::Number(2.0)),
        ("AVERAGEIF(A1:A6,\"<>x\")", CellValue::Number(7.0)),
    ];
    for (formula, expected) in cases {
        let actual = eval_in(&snapshot, formula).unwrap_or_else(|e| panic!("{formula}: {e}"));
        assert_eq!(actual, expected, "{formula}");
    }
}

#[test]
fn direct_scalar_arguments_keep_coercion() {
    let snapshot = aggregate_snapshot();
    assert_eq!(
        eval_in(&snapshot, "AVERAGE(TRUE)").unwrap(),
        CellValue::Number(1.0)
    );
    assert_eq!(
        eval_in(&snapshot, "SUM(\"5\",TRUE)").unwrap(),
        CellValue::Number(6.0)
    );
    assert_eq!(
        eval_in(&snapshot, "SUM(A1:A2,1)").unwrap(),
        CellValue::Number(12.0)
    );
    assert!(eval_in(&snapshot, "SUM(\"x\")").is_err());
}

#[test]
fn errors_inside_aggregated_ranges_still_propagate() {
    let mut snapshot = aggregate_snapshot();
    snapshot.set_cell("Sheet1", 1, 2, error("#DIV/0!"));
    assert!(eval_in(&snapshot, "AVERAGE(A1:A6)").is_err());
    assert!(eval_in(&snapshot, "SUM(A2)").is_err());
}

// F-F3: WORKDAY holiday handling must scale and charge evaluation work.

fn holiday_snapshot(count: u32) -> CellSnapshot {
    let mut snapshot = CellSnapshot::new("Sheet1".to_string());
    for row in 1..=count {
        snapshot.set_cell(
            "Sheet1",
            1,
            row,
            CellValue::Number(45000.0 + f64::from(row)),
        );
    }
    snapshot
}

fn eval_with_limits(
    snapshot: &CellSnapshot,
    formula: &str,
    limits: FormulaLimits,
) -> sheetkit_core::error::Result<CellValue> {
    let budget = FormulaBudget::for_calculation(limits).unwrap();
    Evaluator::with_budget(snapshot, budget).eval_expr(&parse_formula(formula).unwrap())
}

#[test]
fn workday_holiday_work_is_charged_against_evaluation_budget() {
    let snapshot = holiday_snapshot(1000);
    let limits = FormulaLimits {
        operation_evaluation_steps: 3,
        ..FormulaLimits::secure_default()
    };
    for formula in [
        "WORKDAY(45000,1,A1:A1000)",
        "NETWORKDAYS(45000,47000,A1:A1000)",
    ] {
        let error = eval_with_limits(&snapshot, formula, limits).unwrap_err();
        assert!(
            matches!(error, Error::ResourceLimitExceeded(_)),
            "{formula}: {error:?}"
        );
    }
}

#[test]
fn workday_skips_consecutive_holidays_correctly() {
    let snapshot = holiday_snapshot(10);
    // 45000 is Wednesday 2023-03-15; holidays cover 45001..=45010 (through
    // Saturday 2023-03-25), so the next working day is Monday 45012.
    let result = eval_in(&snapshot, "WORKDAY(45000,1,A1:A10)").unwrap();
    assert_eq!(result, CellValue::Date(45012.0));
    let result = eval_in(&snapshot, "WORKDAY(45011,-1,A1:A10)").unwrap();
    assert_eq!(result, CellValue::Date(45000.0));
    let result = eval_in(&snapshot, "NETWORKDAYS(45000,45013,A1:A10)").unwrap();
    assert_eq!(result, CellValue::Number(3.0));
}

fn best_elapsed(snapshot: &CellSnapshot, formula: &str) -> Duration {
    (0..3)
        .map(|_| {
            let start = Instant::now();
            eval_with_limits(snapshot, formula, FormulaLimits::secure_default()).unwrap();
            start.elapsed()
        })
        .min()
        .unwrap()
}

#[test]
fn workday_with_consecutive_holidays_scales_near_linearly() {
    let small = holiday_snapshot(2000);
    let large = holiday_snapshot(8000);
    let small_time = best_elapsed(&small, "WORKDAY(45000,1,A1:A2000)");
    let large_time = best_elapsed(&large, "WORKDAY(45000,1,A1:A8000)");
    // Quadratic behavior would be ~16x; allow generous noise above linear 4x.
    let floor = Duration::from_millis(2);
    assert!(
        large_time <= small_time.max(floor) * 8,
        "small {small_time:?}, large {large_time:?}"
    );
}

// Every seed in the formula fuzz corpus must evaluate without panicking.

#[test]
fn formula_regression_seeds_evaluate_without_panicking() {
    let directory =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../security/fuzz-corpus/formula");
    let mut entries: Vec<_> = std::fs::read_dir(&directory)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .collect();
    entries.sort();
    assert!(
        entries.len() >= 14,
        "expected the full corpus, got {}",
        entries.len()
    );
    let snapshot = holiday_snapshot(64);
    let mut evaluated = 0usize;
    for path in entries {
        let text = std::fs::read_to_string(&path).unwrap();
        let formula = text.trim();
        // Calibration seeds (depth-over, binary-chain-over) are rejected by the
        // default parse budget; a typed rejection is a valid outcome here.
        let Ok(expr) = parse_formula(formula) else {
            continue;
        };
        evaluated += 1;
        let outcome = catch_unwind(AssertUnwindSafe(|| {
            let budget = FormulaBudget::for_calculation(FormulaLimits::secure_default()).unwrap();
            Evaluator::with_budget(&snapshot, budget).eval_expr(&expr)
        }));
        let result = outcome.unwrap_or_else(|_| panic!("{} panicked", path.display()));
        if let Ok(CellValue::Number(n)) = result {
            assert!(n.is_finite(), "{} returned {n}", path.display());
        }
    }
    assert!(
        evaluated >= 8,
        "expected the regression seeds to evaluate, got {evaluated}"
    );
}

// Excel parity: COUNT counts direct booleans and numeric text, but only
// numbers reached through references.

fn count_snapshot() -> CellSnapshot {
    let mut snapshot = CellSnapshot::new("Sheet1".to_string());
    snapshot.set_cell("Sheet1", 1, 1, CellValue::Bool(true));
    snapshot.set_cell("Sheet1", 1, 2, CellValue::String("3".to_string()));
    snapshot.set_cell("Sheet1", 1, 3, CellValue::Number(5.0));
    snapshot.set_cell("Sheet1", 1, 4, error("#DIV/0!"));
    snapshot
}

#[test]
fn count_counts_direct_booleans_and_numeric_text() {
    let snapshot = count_snapshot();
    let cases = [
        ("COUNT(TRUE,\"3\",\"x\")", 2.0),
        ("COUNT(\"3\")", 1.0),
        ("COUNT(TRUE)", 1.0),
        ("COUNT(FALSE)", 1.0),
        ("COUNT(\"TRUE\")", 0.0),
        ("COUNT(\"inf\")", 0.0),
        ("COUNT(\"1e400\")", 0.0),
        ("COUNT(1/0)", 0.0),
        ("COUNT(NA(),1)", 1.0),
        ("COUNT(\"3\"&\"\")", 1.0),
        ("COUNT(A1:A4)", 1.0),
        ("COUNT(A1)", 0.0),
        ("COUNT(A2)", 0.0),
        ("COUNT(A1:A4,TRUE)", 2.0),
        ("COUNTA(A1:A4)", 4.0),
        ("COUNTA(TRUE,\"3\",\"x\")", 3.0),
    ];
    for (formula, expected) in cases {
        let actual = eval_in(&snapshot, formula).unwrap_or_else(|e| panic!("{formula}: {e}"));
        assert_eq!(actual, CellValue::Number(expected), "{formula}");
    }
}

// Excel parity: serial 0 is day zero (1900-01-00); negative serials and
// results before January 1900 are #NUM!.

#[test]
fn edate_and_eomonth_accept_day_zero_and_reject_negative_serials() {
    let numbers = [
        ("EDATE(0,0)", 0.0),
        ("EDATE(0,1)", 31.0),
        ("EDATE(0,2)", 60.0),
        ("EDATE(0,14)", 425.0),
        ("EDATE(0.9,1)", 31.0),
        ("EDATE(0,96000)", 2921941.0),
        ("EDATE(60,0)", 59.0),
        ("EDATE(60,1)", 89.0),
        ("EDATE(60,-1)", 29.0),
        ("EDATE(59,1)", 88.0),
        ("EDATE(31,1)", 59.0),
        ("EDATE(2958465.5,0)", 2958465.0),
        ("EOMONTH(0,0)", 31.0),
        ("EOMONTH(0,1)", 59.0),
        ("EOMONTH(0.5,0)", 31.0),
        ("EOMONTH(32,-1)", 31.0),
        ("EOMONTH(60,0)", 59.0),
        ("EOMONTH(60,1)", 91.0),
    ];
    for (formula, expected) in numbers {
        assert_eq!(number(eval(formula)), expected, "{formula}");
    }
    for formula in [
        "EDATE(-1,0)",
        "EDATE(-0.5,0)",
        "EDATE(0,-1)",
        "EDATE(1,-1)",
        "EDATE(31,-1)",
        "EDATE(2958466,0)",
        "EDATE(1E300,0)",
        "EOMONTH(-1,0)",
        "EOMONTH(-0.5,0)",
        "EOMONTH(0,-1)",
        "EOMONTH(0.9,-1)",
        "EOMONTH(1,-1)",
        "EOMONTH(31,-1)",
        "EOMONTH(2958466,0)",
    ] {
        assert_eq!(eval(formula), error("#NUM!"), "{formula}");
    }
    assert_eq!(eval("EDATE(\"x\",1)"), error("#VALUE!"));
}

// Excel parity: WORKDAY counts weekdays on the serial calendar (serial 1 is a
// Sunday, serial 60 is 1900-02-29) and accepts serial 0.

#[test]
fn workday_uses_excel_serial_calendar_and_day_zero() {
    let numbers = [
        ("WORKDAY(0,1)", 2.0),
        ("WORKDAY(0,0)", 0.0),
        ("WORKDAY(0,5)", 6.0),
        ("WORKDAY(0.9,0)", 0.0),
        ("WORKDAY(0.5,1)", 2.0),
        ("WORKDAY(1,0)", 1.0),
        ("WORKDAY(3,-1)", 2.0),
        ("WORKDAY(6,1)", 9.0),
        ("WORKDAY(10,-5)", 3.0),
        ("WORKDAY(55,10)", 69.0),
        ("WORKDAY(59,1)", 60.0),
        ("WORKDAY(60,1)", 61.0),
        ("WORKDAY(61,-1)", 60.0),
        ("WORKDAY(62,-1)", 61.0),
        ("WORKDAY(70,-15)", 51.0),
        ("WORKDAY(0,1,2)", 3.0),
        ("WORKDAY(1,1,0)", 2.0),
        ("WORKDAY(1,1,0.5)", 2.0),
    ];
    for (formula, expected) in numbers {
        assert_eq!(number(eval(formula)), expected, "{formula}");
    }
    for formula in [
        "WORKDAY(0,-1)",
        "WORKDAY(1,-1)",
        "WORKDAY(2,-1)",
        "WORKDAY(-1,1)",
        "WORKDAY(-1,0)",
        "WORKDAY(-0.5,1)",
        "WORKDAY(-1E300,0)",
        "WORKDAY(2958466,0)",
        "WORKDAY(2958465,1)",
        "WORKDAY(0,1,-1)",
        "WORKDAY(1,1,-0.5)",
        "WORKDAY(1,1,2958466)",
    ] {
        assert_eq!(eval(formula), error("#NUM!"), "{formula}");
    }
    assert_eq!(eval("WORKDAY(\"x\",1)"), error("#VALUE!"));
    assert_eq!(eval("WORKDAY(1,1,\"x\")"), error("#VALUE!"));
    assert_eq!(eval("WORKDAY(1,1,\"\")"), error("#VALUE!"));
}

// Excel parity: text functions propagate an error argument.

#[test]
fn text_functions_propagate_error_arguments() {
    let mut snapshot = CellSnapshot::new("Sheet1".to_string());
    snapshot.set_cell("Sheet1", 1, 1, error("#N/A"));
    let cases = [
        ("LEN(1/0)", "#DIV/0!"),
        ("LEN(REPT(\"a\",32768))", "#VALUE!"),
        ("LEN(A1)", "#N/A"),
        ("LEFT(1/0)", "#DIV/0!"),
        ("RIGHT(NA())", "#N/A"),
        ("MID(1/0,1,1)", "#DIV/0!"),
        ("UPPER(1/0)", "#DIV/0!"),
        ("LOWER(1/0)", "#DIV/0!"),
        ("TRIM(1/0)", "#DIV/0!"),
        ("PROPER(1/0)", "#DIV/0!"),
        ("T(1/0)", "#DIV/0!"),
        ("CONCAT(\"a\",1/0)", "#DIV/0!"),
        ("CONCAT(\"a\",NA(),1/0)", "#N/A"),
        ("CONCATENATE(\"a\",1/0)", "#DIV/0!"),
        ("SUBSTITUTE(1/0,\"a\",\"b\")", "#DIV/0!"),
        ("SUBSTITUTE(\"a\",1/0,\"b\")", "#DIV/0!"),
        ("SUBSTITUTE(\"a\",\"a\",1/0)", "#DIV/0!"),
        ("REPLACE(1/0,1,1,\"x\")", "#DIV/0!"),
        ("REPLACE(\"a\",1,1,1/0)", "#DIV/0!"),
        ("REPT(1/0,2)", "#DIV/0!"),
        ("FIND(\"a\",1/0)", "#DIV/0!"),
        ("FIND(1/0,\"a\")", "#DIV/0!"),
        ("SEARCH(1/0,\"a\")", "#DIV/0!"),
        ("EXACT(1/0,\"a\")", "#DIV/0!"),
        ("EXACT(\"a\",A1)", "#N/A"),
        ("VALUE(1/0)", "#DIV/0!"),
    ];
    for (formula, code) in cases {
        let actual = eval_in(&snapshot, formula).unwrap_or_else(|e| panic!("{formula}: {e}"));
        assert_eq!(actual, error(code), "{formula}");
    }
    assert_eq!(number(eval("LEN(A99)")), 0.0);
}

// Excel parity: MID and REPLACE reject positions and lengths beyond their
// limits instead of clamping.

#[test]
fn mid_and_replace_reject_out_of_range_positions_and_lengths() {
    let text = |value: &str| CellValue::String(value.to_string());
    let strings = [
        ("MID(\"abc\",1,2147483647)", "abc"),
        ("MID(\"abc\",1,2147483647.9)", "abc"),
        ("MID(\"abc\",2,2147483647)", "bc"),
        ("MID(\"abc\",2147483647,1)", ""),
        ("MID(\"abc\",2147483648,1)", ""),
        ("MID(\"abc\",2147483648.5,1)", ""),
        ("MID(\"abc\",2147483648,2147483647)", ""),
        ("REPLACE(\"abc\",32767,1,\"x\")", "abcx"),
        ("REPLACE(\"abc\",32767.9,1,\"x\")", "abcx"),
        ("REPLACE(\"abc\",1,32767,\"x\")", "x"),
        ("REPLACE(\"abc\",1,32767.9,\"x\")", "x"),
    ];
    for (formula, expected) in strings {
        assert_eq!(eval(formula), text(expected), "{formula}");
    }
    for formula in [
        "MID(\"abc\",1,2147483648)",
        "MID(\"abc\",1,1E10)",
        "MID(\"abc\",1,4294967295)",
        "MID(\"abc\",1E10,1)",
        "MID(\"abc\",2147483649,1)",
        "MID(\"abc\",4294967296,1)",
        "MID(\"abc\",0.5,1)",
        "REPLACE(\"abc\",32768,1,\"x\")",
        "REPLACE(\"abc\",2147483647,1,\"x\")",
        "REPLACE(\"abc\",1E10,1,\"x\")",
        "REPLACE(\"abc\",1,32768,\"x\")",
        "REPLACE(\"abc\",1,2147483647,\"x\")",
        "REPLACE(\"abc\",1,2147483648,\"x\")",
        "REPLACE(\"abc\",1,1E10,\"x\")",
        "REPLACE(\"abc\",0.5,1,\"x\")",
    ] {
        assert_eq!(eval(formula), error("#VALUE!"), "{formula}");
    }
}
