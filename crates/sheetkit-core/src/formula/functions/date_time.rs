//! Date/time formula functions: DATE, TODAY, NOW, YEAR, MONTH, DAY, HOUR,
//! MINUTE, SECOND, DATEDIF, EDATE, EOMONTH, DATEVALUE, WEEKDAY, WEEKNUM,
//! NETWORKDAYS, WORKDAY.

use chrono::{Datelike, Duration, Local, NaiveDate, NaiveDateTime, Timelike};

use crate::cell::{date_to_serial, datetime_to_serial, CellValue};
use crate::error::Result;
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_number, coerce_to_string, Evaluator};
use crate::formula::functions::conversion::{checked_trunc_i32, checked_trunc_i64};
use crate::formula::functions::{check_arg_count, collect_criteria_range_values};

/// One past the largest Excel date serial (9999-12-31 23:59:59).
const MAX_DATE_SERIAL_EXCLUSIVE: f64 = 2_958_466.0;
/// Largest year representable in an Excel date.
const MAX_EXCEL_YEAR: i32 = 9999;

fn num_error() -> CellValue {
    CellValue::Error("#NUM!".to_string())
}

/// Convert a serial to a date, rejecting non-finite or post-9999 serials
/// before they reach chrono arithmetic.
fn serial_to_date(serial: f64) -> Option<NaiveDate> {
    if !serial.is_finite() || serial >= MAX_DATE_SERIAL_EXCLUSIVE {
        return None;
    }
    crate::cell::serial_to_date(serial)
}

/// Convert a serial to a date-time with the same domain check as
/// [`serial_to_date`].
fn serial_to_datetime(serial: f64) -> Option<NaiveDateTime> {
    if !serial.is_finite() || serial >= MAX_DATE_SERIAL_EXCLUSIVE {
        return None;
    }
    crate::cell::serial_to_datetime(serial)
}

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

/// Add months to a date, clamping to end of month if needed. Returns `None`
/// when the result falls outside the chrono date range.
fn add_months_to_date(date: NaiveDate, months: i64) -> Option<NaiveDate> {
    let total_months = i64::from(date.year())
        .checked_mul(12)?
        .checked_add(i64::from(date.month0()))?
        .checked_add(months)?;
    let new_year = i32::try_from(total_months.div_euclid(12)).ok()?;
    let new_month = u32::try_from(total_months.rem_euclid(12) + 1).ok()?;
    let max_day = last_day_of_month(new_year, new_month);
    let new_day = date.day().min(max_day);
    NaiveDate::from_ymd_opt(new_year, new_month, new_day)
}

/// Return `date` when its year is within the Excel date range.
fn excel_date(date: NaiveDate) -> Option<NaiveDate> {
    (date.year() <= MAX_EXCEL_YEAR).then_some(date)
}

/// Year, month, and day of a serial on Excel's calendar, where serial 0 is
/// 1900-01-00 and serial 60 is 1900-02-29.
///
/// Negative, non-finite, and post-9999 serials return `#NUM!`.
fn excel_ymd(serial: f64) -> std::result::Result<(i32, u32, u32), CellValue> {
    if !(0.0..MAX_DATE_SERIAL_EXCLUSIVE).contains(&serial) {
        return Err(num_error());
    }
    match serial.floor() {
        0.0 => Ok((1900, 1, 0)),
        60.0 => Ok((1900, 2, 29)),
        _ => serial_to_date(serial)
            .map(|date| (date.year(), date.month(), date.day()))
            .ok_or_else(num_error),
    }
}

/// Shift a year and month by `months`, rejecting results outside the Excel
/// years 1900..=9999.
fn shift_month(year: i32, month: u32, months: i64) -> Option<(i32, u32)> {
    let total = i64::from(year)
        .checked_mul(12)?
        .checked_add(i64::from(month) - 1)?
        .checked_add(months)?;
    let new_year = i32::try_from(total.div_euclid(12)).ok()?;
    if !(1900..=MAX_EXCEL_YEAR).contains(&new_year) {
        return None;
    }
    let new_month = u32::try_from(total.rem_euclid(12) + 1).ok()?;
    Some((new_year, new_month))
}

/// Serial of `day` in the given month; day 0 is the day before the first.
fn month_day_serial(year: i32, month: u32, day: u32) -> Option<f64> {
    let first = NaiveDate::from_ymd_opt(year, month, 1)?;
    if day == 0 {
        return Some(date_to_serial(first) - 1.0);
    }
    NaiveDate::from_ymd_opt(year, month, day).map(date_to_serial)
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

fn workday_date(value: &CellValue) -> std::result::Result<NaiveDate, CellValue> {
    if let CellValue::Error(error) = value {
        return Err(CellValue::Error(error.clone()));
    }
    let serial = to_serial(value)?;
    serial_to_date(serial).ok_or_else(|| CellValue::Error("#VALUE!".to_string()))
}

/// Collect weekday holidays as a sorted, deduplicated list.
///
/// Evaluation work proportional to the number of holiday values is charged
/// before they are converted and sorted.
fn holiday_dates(
    args: &[Expr],
    ctx: &mut Evaluator,
) -> Result<std::result::Result<Vec<NaiveDate>, CellValue>> {
    if args.len() < 3 {
        return Ok(Ok(Vec::new()));
    }

    let values = collect_criteria_range_values(&args[2], ctx)?;
    ctx.charge_evaluation_work(u64::try_from(values.len()).unwrap_or(u64::MAX))?;
    let mut holidays = Vec::new();
    for value in values {
        if value == CellValue::Empty {
            continue;
        }
        let date = match workday_date(&value) {
            Ok(date) => date,
            Err(error) => return Ok(Err(error)),
        };
        if date.weekday().num_days_from_monday() < 5 {
            holidays.push(date);
        }
    }
    holidays.sort_unstable();
    holidays.dedup();
    Ok(Ok(holidays))
}

fn weekday_count(start: NaiveDate, end: NaiveDate) -> i64 {
    let days = (end - start).num_days() + 1;
    let full_weeks = days / 7;
    let remaining_days = days % 7;
    let mut count = full_weeks * 5;
    let start_weekday = start.weekday().num_days_from_monday() as i64;
    for offset in 0..remaining_days {
        if (start_weekday + offset) % 7 < 5 {
            count += 1;
        }
    }
    count
}

fn checked_add_days(date: NaiveDate, days: i64) -> Option<NaiveDate> {
    date.checked_add_signed(Duration::try_days(days)?)
}

/// Integer serial day for WORKDAY on Excel's serial calendar.
///
/// Error values propagate, unconvertible text returns `#VALUE!`, and
/// negative, non-finite, or post-9999 serials return `#NUM!`.
fn workday_serial(value: &CellValue) -> std::result::Result<i64, CellValue> {
    if let CellValue::Error(error) = value {
        return Err(CellValue::Error(error.clone()));
    }
    let serial = to_serial(value)?;
    if !(0.0..MAX_DATE_SERIAL_EXCLUSIVE).contains(&serial) {
        return Err(num_error());
    }
    Ok(serial.floor() as i64)
}

/// Whether a serial day is Monday..Friday. Excel's serial calendar is
/// uniform, so serial 0 is a Saturday and serial 1 a Sunday.
fn is_serial_weekday(serial: i64) -> bool {
    serial.rem_euclid(7) >= 2
}

fn add_serial_weekdays(mut serial: i64, mut days: u64, direction: i64) -> Option<i64> {
    while days > 0 && !is_serial_weekday(serial) {
        serial = serial.checked_add(direction)?;
        if is_serial_weekday(serial) {
            days -= 1;
        }
    }
    if days == 0 {
        return Some(serial);
    }

    let weeks = days / 5;
    if weeks > 0 {
        let calendar_days = i64::try_from(weeks).ok()?.checked_mul(7)?;
        serial = serial.checked_add(direction.checked_mul(calendar_days)?)?;
        days %= 5;
    }
    while days > 0 {
        serial = serial.checked_add(direction)?;
        if is_serial_weekday(serial) {
            days -= 1;
        }
    }
    Some(serial)
}

/// Collect WORKDAY holidays as sorted, deduplicated weekday serials.
///
/// Evaluation work proportional to the number of holiday values is charged
/// before they are converted and sorted.
fn holiday_serials(
    args: &[Expr],
    ctx: &mut Evaluator,
) -> Result<std::result::Result<Vec<i64>, CellValue>> {
    if args.len() < 3 {
        return Ok(Ok(Vec::new()));
    }
    let values = collect_criteria_range_values(&args[2], ctx)?;
    ctx.charge_evaluation_work(u64::try_from(values.len()).unwrap_or(u64::MAX))?;
    let mut holidays = Vec::new();
    for value in values {
        if value == CellValue::Empty {
            continue;
        }
        match workday_serial(&value) {
            Ok(serial) if is_serial_weekday(serial) => holidays.push(serial),
            Ok(_) => {}
            Err(error) => return Ok(Err(error)),
        }
    }
    holidays.sort_unstable();
    holidays.dedup();
    Ok(Ok(holidays))
}

/// DATE(year, month, day) - constructs a date serial number.
///
/// Arguments are truncated toward zero. Years below 1900 are offset by 1900;
/// month and day overflow roll into later or earlier periods. Non-finite
/// arguments, arguments outside the `i32` range, a year outside 0..=9999, or
/// a result after 9999-12-31 return `#NUM!`.
pub fn fn_date(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DATE", args, 3, 3)?;
    let year = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let month = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let day = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let (Some(year), Some(month), Some(day)) = (
        checked_trunc_i32(year),
        checked_trunc_i32(month),
        checked_trunc_i32(day),
    ) else {
        return Ok(num_error());
    };
    if !(0..=MAX_EXCEL_YEAR).contains(&year) {
        return Ok(num_error());
    }
    let adj_year = if year < 1900 { year + 1900 } else { year };

    let result = NaiveDate::from_ymd_opt(adj_year, 1, 1)
        .and_then(|base| add_months_to_date(base, i64::from(month) - 1))
        .and_then(|with_months| checked_add_days(with_months, i64::from(day) - 1))
        .and_then(excel_date);
    match result {
        Some(date) => Ok(CellValue::Date(date_to_serial(date))),
        None => Ok(num_error()),
    }
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
                let Some(d1_last_year) = NaiveDate::from_ymd_opt(
                    d2.year() - 1,
                    d1.month(),
                    d1.day().min(last_day_of_month(d2.year() - 1, d1.month())),
                ) else {
                    return Ok(CellValue::Error("#VALUE!".to_string()));
                };
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
///
/// `months` is truncated toward zero. Serial 0 is 1900-01-00, as in Excel.
/// A negative or post-9999 start, a non-finite month count, one outside the
/// `i32` range, or a result month outside 1900-01..9999-12 returns `#NUM!`.
pub fn fn_edate(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("EDATE", args, 2, 2)?;
    let v = ctx.eval_expr(&args[0])?;
    let months = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let Some(months) = checked_trunc_i32(months) else {
        return Ok(num_error());
    };
    let serial = match to_serial(&v) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    let (year, month, day) = match excel_ymd(serial) {
        Ok(ymd) => ymd,
        Err(e) => return Ok(e),
    };
    let result = shift_month(year, month, i64::from(months)).and_then(|(year, month)| {
        month_day_serial(year, month, day.min(last_day_of_month(year, month)))
    });
    match result {
        Some(serial) => Ok(CellValue::Date(serial)),
        None => Ok(num_error()),
    }
}

/// EOMONTH(start_date, months) - last day of the month N months away.
///
/// Uses the same start and month-count domains as EDATE and returns `#NUM!`
/// outside them.
pub fn fn_eomonth(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("EOMONTH", args, 2, 2)?;
    let v = ctx.eval_expr(&args[0])?;
    let months = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let Some(months) = checked_trunc_i32(months) else {
        return Ok(num_error());
    };
    let serial = match to_serial(&v) {
        Ok(n) => n,
        Err(e) => return Ok(e),
    };
    let (year, month, _) = match excel_ymd(serial) {
        Ok(ymd) => ymd,
        Err(e) => return Ok(e),
    };
    let result = shift_month(year, month, i64::from(months))
        .and_then(|(year, month)| month_day_serial(year, month, last_day_of_month(year, month)));
    match result {
        Some(serial) => Ok(CellValue::Date(serial)),
        None => Ok(num_error()),
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
        let value = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
        let Some(return_type) = checked_trunc_i32(value) else {
            return Ok(num_error());
        };
        return_type
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
        let value = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
        let Some(return_type) = checked_trunc_i32(value) else {
            return Ok(num_error());
        };
        return_type
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
    let Some(jan1) = NaiveDate::from_ymd_opt(date.year(), 1, 1) else {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    };
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
    let d1 = match workday_date(&ctx.eval_expr(&args[0])?) {
        Ok(date) => date,
        Err(error) => return Ok(error),
    };
    let d2 = match workday_date(&ctx.eval_expr(&args[1])?) {
        Ok(date) => date,
        Err(error) => return Ok(error),
    };
    let holidays = match holiday_dates(args, ctx)? {
        Ok(holidays) => holidays,
        Err(error) => return Ok(error),
    };

    let (start, end, sign) = if d1 <= d2 {
        (d1, d2, 1i32)
    } else {
        (d2, d1, -1i32)
    };

    let first = holidays.partition_point(|date| *date < start);
    let past_end = holidays.partition_point(|date| *date <= end);
    let holiday_count = i64::try_from(past_end - first).unwrap_or(i64::MAX);
    let count = weekday_count(start, end).saturating_sub(holiday_count);
    Ok(CellValue::Number((count * i64::from(sign)) as f64))
}

/// WORKDAY(start_date, days, [holidays]) - date after N working days.
///
/// Weekdays follow Excel's serial calendar, where serial 0 is a Saturday and
/// serial 60 (1900-02-29) is a Wednesday. `days` is truncated toward zero;
/// a negative or post-9999 start or holiday, a non-finite day count, or a
/// result outside the Excel date range returns `#NUM!`. Holidays are sorted
/// once and walked with a monotone cursor, and the walk is charged as
/// evaluation work before it runs.
pub fn fn_workday(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("WORKDAY", args, 2, 3)?;
    let start = match workday_serial(&ctx.eval_expr(&args[0])?) {
        Ok(serial) => serial,
        Err(error) => return Ok(error),
    };
    let days = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let Some(days) = checked_trunc_i64(days, i64::MIN, i64::MAX) else {
        return Ok(num_error());
    };
    let holidays = match holiday_serials(args, ctx)? {
        Ok(holidays) => holidays,
        Err(error) => return Ok(error),
    };
    if days == 0 {
        return Ok(CellValue::Date(start as f64));
    }
    let direction = if days >= 0 { 1 } else { -1 };
    let Some(mut current) = add_serial_weekdays(start, days.unsigned_abs(), direction) else {
        return Ok(num_error());
    };
    // Each holiday between start and the running result pushes the result
    // one working day further. Holidays are distinct weekdays, so a single
    // ordered pass is exact.
    let pending: Box<dyn Iterator<Item = &i64>> = if direction > 0 {
        let first = holidays.partition_point(|serial| *serial <= start);
        ctx.charge_evaluation_work(u64::try_from(holidays.len() - first).unwrap_or(u64::MAX))?;
        Box::new(holidays[first..].iter())
    } else {
        let past_last = holidays.partition_point(|serial| *serial < start);
        ctx.charge_evaluation_work(u64::try_from(past_last).unwrap_or(u64::MAX))?;
        Box::new(holidays[..past_last].iter().rev())
    };
    for holiday in pending {
        let reached = if direction > 0 {
            *holiday <= current
        } else {
            *holiday >= current
        };
        if !reached {
            break;
        }
        let Some(next) = add_serial_weekdays(current, 1, direction) else {
            return Ok(num_error());
        };
        current = next;
    }
    if (1..MAX_DATE_SERIAL_EXCLUSIVE as i64).contains(&current) {
        Ok(CellValue::Date(current as f64))
    } else {
        Ok(num_error())
    }
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

    fn eval_with_data(formula: &str, data: &[(&str, u32, u32, CellValue)]) -> CellValue {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        for (sheet, col, row, value) in data {
            snap.set_cell(sheet, *col, *row, value.clone());
        }
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

    #[test]
    fn networkdays_excludes_range_holidays_once_and_preserves_reverse_sign() {
        let holidays = [
            (
                "Sheet1",
                1,
                1,
                CellValue::Date(date_to_serial(NaiveDate::from_ymd_opt(2024, 1, 1).unwrap())),
            ),
            (
                "Sheet1",
                1,
                2,
                CellValue::Date(date_to_serial(NaiveDate::from_ymd_opt(2024, 1, 1).unwrap())),
            ),
            (
                "Sheet1",
                1,
                3,
                CellValue::Date(date_to_serial(NaiveDate::from_ymd_opt(2024, 1, 6).unwrap())),
            ),
        ];
        assert_eq!(
            eval_with_data(
                "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7),A1:A3)",
                &holidays
            ),
            CellValue::Number(4.0)
        );
        assert_eq!(
            eval_with_data(
                "NETWORKDAYS(DATE(2024,1,7),DATE(2024,1,1),A1:A3)",
                &holidays
            ),
            CellValue::Number(-4.0)
        );
    }

    #[test]
    fn workday_skips_scalar_and_range_holidays_in_both_directions() {
        assert_eq!(
            eval("WORKDAY(DATE(2024,1,5),1,DATE(2024,1,8))"),
            CellValue::Date(date_to_serial(NaiveDate::from_ymd_opt(2024, 1, 9).unwrap()))
        );
        let holidays = [
            (
                "Sheet1",
                1,
                1,
                CellValue::Date(date_to_serial(NaiveDate::from_ymd_opt(2024, 1, 8).unwrap())),
            ),
            (
                "Sheet1",
                1,
                2,
                CellValue::Date(date_to_serial(NaiveDate::from_ymd_opt(2024, 1, 8).unwrap())),
            ),
            (
                "Sheet1",
                1,
                3,
                CellValue::Date(date_to_serial(NaiveDate::from_ymd_opt(2024, 1, 7).unwrap())),
            ),
        ];
        assert_eq!(
            eval_with_data("WORKDAY(DATE(2024,1,9),-1,A1:A3)", &holidays),
            CellValue::Date(date_to_serial(NaiveDate::from_ymd_opt(2024, 1, 5).unwrap()))
        );
    }

    #[test]
    fn workday_functions_propagate_holiday_errors_and_reject_invalid_dates() {
        let errors = [("Sheet1", 1, 1, CellValue::Error("#N/A".to_string()))];
        assert_eq!(
            eval_with_data("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),A1)", &errors),
            CellValue::Error("#N/A".to_string())
        );
        assert_eq!(
            eval_with_data("WORKDAY(DATE(2024,1,5),1,A1)", &errors),
            CellValue::Error("#N/A".to_string())
        );
        assert_eq!(
            eval("NETWORKDAYS(0,DATE(2024,1,5))"),
            CellValue::Error("#VALUE!".to_string())
        );
        assert_eq!(eval("WORKDAY(0,1)"), CellValue::Date(2.0));
        assert_eq!(eval("WORKDAY(-1,1)"), CellValue::Error("#NUM!".to_string()));
    }
}
