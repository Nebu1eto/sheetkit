//! Date/time formula functions: DATE, TODAY, NOW, YEAR, MONTH, DAY, HOUR,
//! MINUTE, SECOND, DATEDIF, EDATE, EOMONTH, DATEVALUE, WEEKDAY, WEEKNUM,
//! NETWORKDAYS, WORKDAY.

use chrono::{Datelike, Duration, Local, NaiveDate, Timelike};

use crate::cell::{
    date_to_serial, datetime_to_serial, serial_to_date, serial_to_datetime, CellValue,
};
use crate::error::Result;
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_number, coerce_to_string, Evaluator};
use crate::formula::functions::check_arg_count;

/// Convert a CellValue to an Excel serial number.
fn to_serial(v: &CellValue) -> std::result::Result<f64, CellValue> {
    match v {
        CellValue::Number(n) => Ok(*n),
        CellValue::Date(n) => Ok(*n),
        CellValue::String(s) => {
            if let Ok(n) = s.parse::<f64>() {
                Ok(n)
            } else if let Some(d) = parse_date_string(s) {
                Ok(date_to_serial(d))
            } else {
                Err(CellValue::Error("#VALUE!".to_string()))
            }
        }
        CellValue::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
        CellValue::Empty => Ok(0.0),
        _ => Err(CellValue::Error("#VALUE!".to_string())),
    }
}

/// Parse common date string formats: YYYY-MM-DD and MM/DD/YYYY.
fn parse_date_string(s: &str) -> Option<NaiveDate> {
    if let Ok(d) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return Some(d);
    }
    if let Ok(d) = NaiveDate::parse_from_str(s, "%m/%d/%Y") {
        return Some(d);
    }
    None
}

/// Add months to a date, clamping to end of month if needed.
fn add_months_to_date(date: NaiveDate, months: i32) -> Option<NaiveDate> {
    let total_months = date.year() * 12 + date.month() as i32 - 1 + months;
    let new_year = total_months.div_euclid(12);
    let new_month = (total_months.rem_euclid(12) + 1) as u32;
    let max_day = last_day_of_month(new_year, new_month);
    let new_day = date.day().min(max_day);
    NaiveDate::from_ymd_opt(new_year, new_month, new_day)
}

/// Return the last day of the given month.
fn last_day_of_month(year: i32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 {
                29
            } else {
                28
            }
        }
        _ => 30,
    }
}

/// DATE(year, month, day) - constructs a date serial number.
pub fn fn_date(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DATE", args, 3, 3)?;
    let year = coerce_to_number(&ctx.eval_expr(&args[0])?)? as i32;
    let month = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    let day = coerce_to_number(&ctx.eval_expr(&args[2])?)? as i32;

    let adj_year = if year < 1900 { year + 1900 } else { year };

    // Handle month overflow/underflow
    let base = NaiveDate::from_ymd_opt(adj_year, 1, 1);
    let base = match base {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let with_months = add_months_to_date(base, month - 1);
    let with_months = match with_months {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let result = with_months + Duration::days(i64::from(day) - 1);
    Ok(CellValue::Date(date_to_serial(result)))
}

/// TODAY() - returns today's date serial number.
pub fn fn_today(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("TODAY", args, 0, 0)?;
    let today = Local::now().date_naive();
    Ok(CellValue::Date(date_to_serial(today)))
}

/// NOW() - returns current date and time serial number.
pub fn fn_now(args: &[Expr], _ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("NOW", args, 0, 0)?;
    let now = Local::now().naive_local();
    Ok(CellValue::Date(datetime_to_serial(now)))
}

/// YEAR(serial_number) - extracts the year from a date serial.
pub fn fn_year(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("YEAR", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    match to_serial(&v) {
        Ok(serial) => match serial_to_date(serial) {
            Some(d) => Ok(CellValue::Number(d.year() as f64)),
            None => Ok(CellValue::Error("#VALUE!".to_string())),
        },
        Err(e) => Ok(e),
    }
}

/// MONTH(serial_number) - extracts the month from a date serial.
pub fn fn_month(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("MONTH", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    match to_serial(&v) {
        Ok(serial) => match serial_to_date(serial) {
            Some(d) => Ok(CellValue::Number(d.month() as f64)),
            None => Ok(CellValue::Error("#VALUE!".to_string())),
        },
        Err(e) => Ok(e),
    }
}

/// DAY(serial_number) - extracts the day from a date serial.
pub fn fn_day(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DAY", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    match to_serial(&v) {
        Ok(serial) => match serial_to_date(serial) {
            Some(d) => Ok(CellValue::Number(d.day() as f64)),
            None => Ok(CellValue::Error("#VALUE!".to_string())),
        },
        Err(e) => Ok(e),
    }
}

/// HOUR(serial_number) - extracts the hour from a datetime serial.
pub fn fn_hour(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("HOUR", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    match to_serial(&v) {
        Ok(serial) => match serial_to_datetime(serial) {
            Some(dt) => Ok(CellValue::Number(dt.hour() as f64)),
            None => Ok(CellValue::Error("#VALUE!".to_string())),
        },
        Err(e) => Ok(e),
    }
}

/// MINUTE(serial_number) - extracts the minute from a datetime serial.
pub fn fn_minute(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("MINUTE", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    match to_serial(&v) {
        Ok(serial) => match serial_to_datetime(serial) {
            Some(dt) => Ok(CellValue::Number(dt.minute() as f64)),
            None => Ok(CellValue::Error("#VALUE!".to_string())),
        },
        Err(e) => Ok(e),
    }
}

/// SECOND(serial_number) - extracts the second from a datetime serial.
pub fn fn_second(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SECOND", args, 1, 1)?;
    let v = ctx.eval_expr(&args[0])?;
    match to_serial(&v) {
        Ok(serial) => match serial_to_datetime(serial) {
            Some(dt) => Ok(CellValue::Number(dt.second() as f64)),
            None => Ok(CellValue::Error("#VALUE!".to_string())),
        },
        Err(e) => Ok(e),
    }
}

/// DATEDIF(start_date, end_date, unit) - difference between two dates.
pub fn fn_datedif(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DATEDIF", args, 3, 3)?;
    let v1 = ctx.eval_expr(&args[0])?;
    let v2 = ctx.eval_expr(&args[1])?;
    let unit = coerce_to_string(&ctx.eval_expr(&args[2])?).to_ascii_uppercase();

    let s1 = match to_serial(&v1) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    let s2 = match to_serial(&v2) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    if s1 > s2 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let d1 = match serial_to_date(s1) {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let d2 = match serial_to_date(s2) {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let result = match unit.as_str() {
        "Y" => {
            let mut years = d2.year() - d1.year();
            if (d2.month(), d2.day()) < (d1.month(), d1.day()) {
                years -= 1;
            }
            years as f64
        }
        "M" => {
            let mut months = (d2.year() - d1.year()) * 12 + d2.month() as i32 - d1.month() as i32;
            if d2.day() < d1.day() {
                months -= 1;
            }
            months as f64
        }
        "D" => s2.floor() - s1.floor(),
        "YM" => {
            let mut months = d2.month() as i32 - d1.month() as i32;
            if d2.day() < d1.day() {
                months -= 1;
            }
            if months < 0 {
                months += 12;
            }
            months as f64
        }
        "YD" => {
            let mut d1_this_year = NaiveDate::from_ymd_opt(
                d2.year(),
                d1.month(),
                d1.day().min(last_day_of_month(d2.year(), d1.month())),
            );
            let d1_this_year = match d1_this_year.take() {
                Some(d) => d,
                None => return Ok(CellValue::Error("#VALUE!".to_string())),
            };
            let days = if d2 >= d1_this_year {
                (d2 - d1_this_year).num_days()
            } else {
                let d1_last_year = NaiveDate::from_ymd_opt(
                    d2.year() - 1,
                    d1.month(),
                    d1.day().min(last_day_of_month(d2.year() - 1, d1.month())),
                )
                .unwrap();
                (d2 - d1_last_year).num_days()
            };
            days as f64
        }
        "MD" => {
            let mut days = d2.day() as i32 - d1.day() as i32;
            if days < 0 {
                let prev_month_end = if d2.month() == 1 {
                    last_day_of_month(d2.year() - 1, 12)
                } else {
                    last_day_of_month(d2.year(), d2.month() - 1)
                };
                days += prev_month_end as i32;
            }
            days as f64
        }
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };
    Ok(CellValue::Number(result))
}

/// EDATE(start_date, months) - date that is N months before/after start_date.
pub fn fn_edate(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("EDATE", args, 2, 2)?;
    let v = ctx.eval_expr(&args[0])?;
    let months = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    let serial = match to_serial(&v) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    let date = match serial_to_date(serial) {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    match add_months_to_date(date, months) {
        Some(d) => Ok(CellValue::Date(date_to_serial(d))),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

/// EOMONTH(start_date, months) - last day of the month N months away.
pub fn fn_eomonth(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("EOMONTH", args, 2, 2)?;
    let v = ctx.eval_expr(&args[0])?;
    let months = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    let serial = match to_serial(&v) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    let date = match serial_to_date(serial) {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    match add_months_to_date(date, months) {
        Some(d) => {
            let eom = last_day_of_month(d.year(), d.month());
            let result = NaiveDate::from_ymd_opt(d.year(), d.month(), eom).unwrap();
            Ok(CellValue::Date(date_to_serial(result)))
        }
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

/// DATEVALUE(date_text) - converts a date string to a serial number.
pub fn fn_datevalue(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DATEVALUE", args, 1, 1)?;
    let text = coerce_to_string(&ctx.eval_expr(&args[0])?);
    match parse_date_string(&text) {
        Some(d) => Ok(CellValue::Date(date_to_serial(d))),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

/// WEEKDAY(serial_number, [return_type]) - day of the week.
pub fn fn_weekday(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("WEEKDAY", args, 1, 2)?;
    let v = ctx.eval_expr(&args[0])?;
    let return_type = if args.len() > 1 {
        coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32
    } else {
        1
    };
    let serial = match to_serial(&v) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    let date = match serial_to_date(serial) {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    // chrono: Mon=1 .. Sun=7
    let weekday_num = date.weekday().num_days_from_monday(); // Mon=0 .. Sun=6
    let result = match return_type {
        1 => ((weekday_num + 1) % 7) + 1, // Sun=1, Mon=2, ..., Sat=7
        2 => weekday_num + 1,             // Mon=1, ..., Sun=7
        3 => weekday_num,                 // Mon=0, ..., Sun=6
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };
    Ok(CellValue::Number(result as f64))
}

/// WEEKNUM(serial_number, [return_type]) - week number of the year.
pub fn fn_weeknum(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("WEEKNUM", args, 1, 2)?;
    let v = ctx.eval_expr(&args[0])?;
    let return_type = if args.len() > 1 {
        coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32
    } else {
        1
    };
    let serial = match to_serial(&v) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    let date = match serial_to_date(serial) {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let jan1 = NaiveDate::from_ymd_opt(date.year(), 1, 1).unwrap();
    let jan1_weekday = jan1.weekday().num_days_from_monday(); // Mon=0..Sun=6
    let day_of_year = date.ordinal() as i32;

    let week_start_offset = match return_type {
        1 => (jan1_weekday + 1) % 7, // Sunday-based
        2 => jan1_weekday,           // Monday-based
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };
    let week = (day_of_year - 1 + week_start_offset as i32) / 7 + 1;
    Ok(CellValue::Number(week as f64))
}

/// NETWORKDAYS(start_date, end_date, [holidays]) - working days between two dates.
pub fn fn_networkdays(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("NETWORKDAYS", args, 2, 3)?;
    let v1 = ctx.eval_expr(&args[0])?;
    let v2 = ctx.eval_expr(&args[1])?;
    let s1 = match to_serial(&v1) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    let s2 = match to_serial(&v2) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    let d1 = match serial_to_date(s1) {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let d2 = match serial_to_date(s2) {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let (start, end, sign) = if d1 <= d2 {
        (d1, d2, 1i32)
    } else {
        (d2, d1, -1i32)
    };

    let mut count = 0i32;
    let mut current = start;
    while current <= end {
        let wd = current.weekday().num_days_from_monday();
        if wd < 5 {
            count += 1;
        }
        current += Duration::days(1);
    }
    Ok(CellValue::Number((count * sign) as f64))
}

/// WORKDAY(start_date, days, [holidays]) - date after N working days.
pub fn fn_workday(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("WORKDAY", args, 2, 3)?;
    let v = ctx.eval_expr(&args[0])?;
    let days = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    let serial = match to_serial(&v) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    let start = match serial_to_date(serial) {
        Some(d) => d,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let step = if days >= 0 { 1i64 } else { -1i64 };
    let mut remaining = days.unsigned_abs() as i32;
    let mut current = start;
    while remaining > 0 {
        current += Duration::days(step);
        let wd = current.weekday().num_days_from_monday();
        if wd < 5 {
            remaining -= 1;
        }
    }
    Ok(CellValue::Date(date_to_serial(current)))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::date_to_serial;
    use crate::formula::eval::{evaluate, CellSnapshot};
    use crate::formula::parser::parse_formula;

    fn eval(formula: &str) -> CellValue {
        let snap = CellSnapshot::new("Sheet1".to_string());
        let expr = parse_formula(formula).unwrap();
        evaluate(&expr, &snap).unwrap()
    }

    #[test]
    fn test_date_basic() {
        let result = eval("DATE(2024,1,15)");
        let expected = date_to_serial(NaiveDate::from_ymd_opt(2024, 1, 15).unwrap());
        assert_eq!(result, CellValue::Date(expected));
    }

    #[test]
    fn test_date_month_overflow() {
        let result = eval("DATE(2024,13,1)");
        let expected = date_to_serial(NaiveDate::from_ymd_opt(2025, 1, 1).unwrap());
        assert_eq!(result, CellValue::Date(expected));
    }

    #[test]
    fn test_year() {
        assert_eq!(eval("YEAR(DATE(2024,6,15))"), CellValue::Number(2024.0));
    }

    #[test]
    fn test_month() {
        assert_eq!(eval("MONTH(DATE(2024,6,15))"), CellValue::Number(6.0));
    }

    #[test]
    fn test_day() {
        assert_eq!(eval("DAY(DATE(2024,6,15))"), CellValue::Number(15.0));
    }

    #[test]
    fn test_hour() {
        // Serial 36526.5 = 2000-01-01 12:00:00
        assert_eq!(eval("HOUR(36526.5)"), CellValue::Number(12.0));
    }

    #[test]
    fn test_minute() {
        // 36526 + 14*3600 + 30*60 = time fraction
        // 14:30 = (14*3600 + 30*60) / 86400 = 52200/86400 = 0.604166...
        let serial = 36526.0 + 52200.0 / 86400.0;
        let formula = format!("MINUTE({serial})");
        assert_eq!(eval(&formula), CellValue::Number(30.0));
    }

    #[test]
    fn test_second() {
        // 36526 + 14*3600 + 30*60 + 45 = time fraction
        let serial = 36526.0 + (14.0 * 3600.0 + 30.0 * 60.0 + 45.0) / 86400.0;
        let formula = format!("SECOND({serial})");
        assert_eq!(eval(&formula), CellValue::Number(45.0));
    }

    #[test]
    fn test_datedif_years() {
        assert_eq!(
            eval("DATEDIF(DATE(2020,1,1),DATE(2024,6,15),\"Y\")"),
            CellValue::Number(4.0)
        );
    }

    #[test]
    fn test_datedif_months() {
        assert_eq!(
            eval("DATEDIF(DATE(2024,1,1),DATE(2024,6,15),\"M\")"),
            CellValue::Number(5.0)
        );
    }

    #[test]
    fn test_datedif_days() {
        assert_eq!(
            eval("DATEDIF(DATE(2024,1,1),DATE(2024,1,31),\"D\")"),
            CellValue::Number(30.0)
        );
    }

    #[test]
    fn test_edate() {
        let result = eval("EDATE(DATE(2024,1,31),1)");
        let expected = date_to_serial(NaiveDate::from_ymd_opt(2024, 2, 29).unwrap());
        assert_eq!(result, CellValue::Date(expected));
    }

    #[test]
    fn test_eomonth() {
        let result = eval("EOMONTH(DATE(2024,1,15),1)");
        let expected = date_to_serial(NaiveDate::from_ymd_opt(2024, 2, 29).unwrap());
        assert_eq!(result, CellValue::Date(expected));
    }

    #[test]
    fn test_datevalue() {
        let result = eval(r#"DATEVALUE("2024-06-15")"#);
        let expected = date_to_serial(NaiveDate::from_ymd_opt(2024, 6, 15).unwrap());
        assert_eq!(result, CellValue::Date(expected));
    }

    #[test]
    fn test_weekday_type1() {
        // 2024-01-15 is a Monday
        assert_eq!(eval("WEEKDAY(DATE(2024,1,15),1)"), CellValue::Number(2.0));
    }

    #[test]
    fn test_weekday_type2() {
        // Monday = 1 in type 2
        assert_eq!(eval("WEEKDAY(DATE(2024,1,15),2)"), CellValue::Number(1.0));
    }

    #[test]
    fn test_weeknum() {
        // 2024-01-15 is in week 3 (Sunday-based)
        assert_eq!(eval("WEEKNUM(DATE(2024,1,15),1)"), CellValue::Number(3.0));
    }

    #[test]
    fn test_networkdays() {
        // Jan 1 2024 (Mon) to Jan 5 2024 (Fri) = 5 working days
        assert_eq!(
            eval("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5))"),
            CellValue::Number(5.0)
        );
    }

    #[test]
    fn test_workday() {
        // Start on Fri Jan 5 2024, add 1 workday = Mon Jan 8 2024
        let result = eval("WORKDAY(DATE(2024,1,5),1)");
        let expected = date_to_serial(NaiveDate::from_ymd_opt(2024, 1, 8).unwrap());
        assert_eq!(result, CellValue::Date(expected));
    }
}
