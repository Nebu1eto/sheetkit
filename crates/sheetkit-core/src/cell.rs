//! Cell value representation.
//!
//! Provides the [`CellValue`] enum which represents the typed value of a
//! single cell in a worksheet. This is the high-level counterpart to the
//! raw XML `Cell` element from `sheetkit-xml`.

use std::fmt;

use chrono::{NaiveDate, NaiveDateTime, NaiveTime, Timelike};

/// Represents the value of a cell.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// No value (empty cell).
    Empty,
    /// Boolean value.
    Bool(bool),
    /// Numeric value (integers are stored as f64 in Excel).
    Number(f64),
    /// String value.
    String(String),
    /// Formula with optional cached result.
    Formula {
        expr: String,
        result: Option<Box<CellValue>>,
    },
    /// A date/time value stored as an Excel serial number.
    /// Integer part = days since 1899-12-30 (Excel epoch).
    /// Fractional part = time of day (0.5 = noon).
    Date(f64),
    /// Error value (e.g. #DIV/0!, #N/A, #VALUE!).
    Error(String),
}

impl Default for CellValue {
    fn default() -> Self {
        Self::Empty
    }
}

impl fmt::Display for CellValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CellValue::Empty => write!(f, ""),
            CellValue::Bool(b) => write!(f, "{}", if *b { "TRUE" } else { "FALSE" }),
            CellValue::Number(n) => {
                // Display integers without decimal point
                if n.fract() == 0.0 && n.is_finite() {
                    write!(f, "{}", *n as i64)
                } else {
                    write!(f, "{n}")
                }
            }
            CellValue::Date(serial) => {
                if let Some(dt) = serial_to_datetime(*serial) {
                    if serial.fract() == 0.0 {
                        // Date only, no time component.
                        write!(f, "{}", dt.format("%Y-%m-%d"))
                    } else {
                        write!(f, "{}", dt.format("%Y-%m-%d %H:%M:%S"))
                    }
                } else {
                    write!(f, "{serial}")
                }
            }
            CellValue::String(s) => write!(f, "{s}"),
            CellValue::Formula { result, expr, .. } => {
                if let Some(result) = result {
                    write!(f, "{result}")
                } else {
                    write!(f, "={expr}")
                }
            }
            CellValue::Error(e) => write!(f, "{e}"),
        }
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        CellValue::String(s.to_string())
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        CellValue::String(s)
    }
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

impl From<i32> for CellValue {
    fn from(n: i32) -> Self {
        CellValue::Number(f64::from(n))
    }
}

impl From<i64> for CellValue {
    fn from(n: i64) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        CellValue::Bool(b)
    }
}

impl From<NaiveDate> for CellValue {
    fn from(date: NaiveDate) -> Self {
        CellValue::Date(date_to_serial(date))
    }
}

impl From<NaiveDateTime> for CellValue {
    fn from(datetime: NaiveDateTime) -> Self {
        CellValue::Date(datetime_to_serial(datetime))
    }
}

/// Number of seconds in a day.
const SECONDS_PER_DAY: f64 = 86_400.0;

/// Convert a `NaiveDate` to an Excel serial number.
///
/// Serial number 1 = January 1, 1900. Accounts for the Excel 1900 leap year
/// bug (serial 60 = the non-existent February 29, 1900).
pub fn date_to_serial(date: NaiveDate) -> f64 {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
    let days = (date - epoch).num_days() as f64;
    // Excel incorrectly considers 1900 a leap year. All dates from
    // March 1, 1900 onward have serial numbers one higher than the
    // actual day count from the epoch.
    if days >= 60.0 {
        days + 1.0
    } else {
        days
    }
}

/// Convert a `NaiveDateTime` to an Excel serial number with fractional time.
///
/// The integer part represents the date (see [`date_to_serial`]) and the
/// fractional part represents the time of day (0.5 = noon).
pub fn datetime_to_serial(datetime: NaiveDateTime) -> f64 {
    let date_part = date_to_serial(datetime.date());
    let time = datetime.time();
    let seconds_since_midnight =
        time.hour() as f64 * 3600.0 + time.minute() as f64 * 60.0 + time.second() as f64;
    date_part + seconds_since_midnight / SECONDS_PER_DAY
}

/// Convert an Excel serial number to a `NaiveDate`.
///
/// Returns `None` for invalid serial numbers (< 1).
pub fn serial_to_date(serial: f64) -> Option<NaiveDate> {
    let serial_int = serial.floor() as i64;
    if serial_int < 1 {
        return None;
    }
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
    // Serial 60 = February 29, 1900 (Excel bug -- this date does not exist).
    // Map it to February 28, 1900 for practical purposes.
    if serial_int == 60 {
        return NaiveDate::from_ymd_opt(1900, 2, 28);
    }
    // For serial >= 61, subtract 1 to compensate for the phantom leap day.
    let adjusted = if serial_int >= 61 {
        serial_int - 1
    } else {
        serial_int
    };
    epoch.checked_add_signed(chrono::Duration::days(adjusted))
}

/// Convert an Excel serial number to a `NaiveDateTime`.
///
/// Returns `None` for invalid serial numbers (< 1).
pub fn serial_to_datetime(serial: f64) -> Option<NaiveDateTime> {
    let date = serial_to_date(serial)?;
    let frac = serial.fract().abs();
    let total_seconds = (frac * SECONDS_PER_DAY).round() as u32;
    let hours = total_seconds / 3600;
    let minutes = (total_seconds % 3600) / 60;
    let seconds = total_seconds % 60;
    let time = NaiveTime::from_hms_opt(hours, minutes, seconds)?;
    Some(NaiveDateTime::new(date, time))
}

/// Returns `true` if the given number format ID is a built-in date or time format.
///
/// Built-in date format IDs: 14-22 and 45-47.
pub fn is_date_num_fmt(num_fmt_id: u32) -> bool {
    matches!(num_fmt_id, 14..=22 | 45..=47)
}

/// Returns `true` if a custom number format string looks like a date/time format.
///
/// Checks for common date/time tokens such as y, m, d, h, s in the format code.
/// Ignores text in quoted strings and escaped characters.
pub fn is_date_format_code(code: &str) -> bool {
    let mut in_quotes = false;
    let mut prev_backslash = false;
    for ch in code.chars() {
        if prev_backslash {
            prev_backslash = false;
            continue;
        }
        if ch == '\\' {
            prev_backslash = true;
            continue;
        }
        if ch == '"' {
            in_quotes = !in_quotes;
            continue;
        }
        if in_quotes {
            continue;
        }
        let lower = ch.to_ascii_lowercase();
        if matches!(lower, 'y' | 'd' | 'h' | 's') {
            return true;
        }
        // 'm' is ambiguous (month vs minute); consider it a date token
        // when not preceded by 'h' or followed by 's'. For simplicity,
        // treat any bare 'm' as a date indicator.
        if lower == 'm' {
            return true;
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_value_default_is_empty() {
        let v = CellValue::default();
        assert_eq!(v, CellValue::Empty);
    }

    #[test]
    fn test_cell_value_from_str() {
        let v: CellValue = "hello".into();
        assert_eq!(v, CellValue::String("hello".to_string()));
    }

    #[test]
    fn test_cell_value_from_string() {
        let v: CellValue = String::from("world").into();
        assert_eq!(v, CellValue::String("world".to_string()));
    }

    #[test]
    fn test_cell_value_from_f64() {
        let v: CellValue = 3.14.into();
        assert_eq!(v, CellValue::Number(3.14));
    }

    #[test]
    fn test_cell_value_from_i32() {
        let v: CellValue = 42i32.into();
        assert_eq!(v, CellValue::Number(42.0));
    }

    #[test]
    fn test_cell_value_from_i64() {
        let v: CellValue = 100i64.into();
        assert_eq!(v, CellValue::Number(100.0));
    }

    #[test]
    fn test_cell_value_from_bool() {
        let v: CellValue = true.into();
        assert_eq!(v, CellValue::Bool(true));

        let v2: CellValue = false.into();
        assert_eq!(v2, CellValue::Bool(false));
    }

    #[test]
    fn test_cell_value_display() {
        assert_eq!(CellValue::Empty.to_string(), "");
        assert_eq!(CellValue::Bool(true).to_string(), "TRUE");
        assert_eq!(CellValue::Bool(false).to_string(), "FALSE");
        assert_eq!(CellValue::Number(42.0).to_string(), "42");
        assert_eq!(CellValue::Number(3.14).to_string(), "3.14");
        assert_eq!(CellValue::String("hello".to_string()).to_string(), "hello");
        assert_eq!(
            CellValue::Error("#DIV/0!".to_string()).to_string(),
            "#DIV/0!"
        );
        assert_eq!(
            CellValue::Formula {
                expr: "A1+B1".to_string(),
                result: Some(Box::new(CellValue::Number(42.0))),
            }
            .to_string(),
            "42"
        );
        assert_eq!(
            CellValue::Formula {
                expr: "A1+B1".to_string(),
                result: None,
            }
            .to_string(),
            "=A1+B1"
        );
    }

    // -- Date conversion tests --

    #[test]
    fn test_date_to_serial_jan_1_1900() {
        let date = NaiveDate::from_ymd_opt(1900, 1, 1).unwrap();
        assert_eq!(date_to_serial(date), 1.0);
    }

    #[test]
    fn test_date_to_serial_feb_28_1900() {
        let date = NaiveDate::from_ymd_opt(1900, 2, 28).unwrap();
        assert_eq!(date_to_serial(date), 59.0);
    }

    #[test]
    fn test_date_to_serial_mar_1_1900_accounts_for_leap_year_bug() {
        // March 1, 1900 should be serial 61 (skipping the phantom Feb 29).
        let date = NaiveDate::from_ymd_opt(1900, 3, 1).unwrap();
        assert_eq!(date_to_serial(date), 61.0);
    }

    #[test]
    fn test_date_to_serial_jan_1_2000() {
        let date = NaiveDate::from_ymd_opt(2000, 1, 1).unwrap();
        // Known value: January 1, 2000 = serial 36526 in Excel.
        assert_eq!(date_to_serial(date), 36526.0);
    }

    #[test]
    fn test_date_to_serial_jan_1_1970() {
        let date = NaiveDate::from_ymd_opt(1970, 1, 1).unwrap();
        assert_eq!(date_to_serial(date), 25569.0);
    }

    #[test]
    fn test_serial_to_date_jan_1_1900() {
        let date = serial_to_date(1.0).unwrap();
        assert_eq!(date, NaiveDate::from_ymd_opt(1900, 1, 1).unwrap());
    }

    #[test]
    fn test_serial_to_date_feb_28_1900() {
        let date = serial_to_date(59.0).unwrap();
        assert_eq!(date, NaiveDate::from_ymd_opt(1900, 2, 28).unwrap());
    }

    #[test]
    fn test_serial_to_date_60_phantom_leap_day() {
        // Serial 60 is the phantom Feb 29 1900. We map it to Feb 28.
        let date = serial_to_date(60.0).unwrap();
        assert_eq!(date, NaiveDate::from_ymd_opt(1900, 2, 28).unwrap());
    }

    #[test]
    fn test_serial_to_date_mar_1_1900() {
        let date = serial_to_date(61.0).unwrap();
        assert_eq!(date, NaiveDate::from_ymd_opt(1900, 3, 1).unwrap());
    }

    #[test]
    fn test_serial_to_date_jan_1_2000() {
        let date = serial_to_date(36526.0).unwrap();
        assert_eq!(date, NaiveDate::from_ymd_opt(2000, 1, 1).unwrap());
    }

    #[test]
    fn test_serial_to_date_invalid() {
        assert!(serial_to_date(0.0).is_none());
        assert!(serial_to_date(-1.0).is_none());
    }

    #[test]
    fn test_date_roundtrip() {
        // Test that date_to_serial -> serial_to_date is a roundtrip for dates
        // after March 1, 1900.
        let dates = vec![
            NaiveDate::from_ymd_opt(1900, 3, 1).unwrap(),
            NaiveDate::from_ymd_opt(2024, 6, 15).unwrap(),
            NaiveDate::from_ymd_opt(1999, 12, 31).unwrap(),
            NaiveDate::from_ymd_opt(2100, 1, 1).unwrap(),
        ];
        for date in dates {
            let serial = date_to_serial(date);
            let roundtripped = serial_to_date(serial).unwrap();
            assert_eq!(roundtripped, date, "roundtrip failed for {date}");
        }
    }

    #[test]
    fn test_datetime_to_serial_noon() {
        let dt = NaiveDate::from_ymd_opt(2000, 1, 1)
            .unwrap()
            .and_hms_opt(12, 0, 0)
            .unwrap();
        let serial = datetime_to_serial(dt);
        // 36526 + 0.5 = 36526.5
        assert!((serial - 36526.5).abs() < 1e-9);
    }

    #[test]
    fn test_datetime_to_serial_with_time() {
        let dt = NaiveDate::from_ymd_opt(2000, 1, 1)
            .unwrap()
            .and_hms_opt(6, 0, 0)
            .unwrap();
        let serial = datetime_to_serial(dt);
        // 6 AM = 0.25 of a day
        assert!((serial - 36526.25).abs() < 1e-9);
    }

    #[test]
    fn test_serial_to_datetime_noon() {
        let dt = serial_to_datetime(36526.5).unwrap();
        assert_eq!(dt.date(), NaiveDate::from_ymd_opt(2000, 1, 1).unwrap());
        assert_eq!(dt.time(), NaiveTime::from_hms_opt(12, 0, 0).unwrap());
    }

    #[test]
    fn test_datetime_roundtrip() {
        let dt = NaiveDate::from_ymd_opt(2024, 3, 15)
            .unwrap()
            .and_hms_opt(14, 30, 45)
            .unwrap();
        let serial = datetime_to_serial(dt);
        let roundtripped = serial_to_datetime(serial).unwrap();
        assert_eq!(roundtripped, dt);
    }

    #[test]
    fn test_cell_value_date_display_date_only() {
        let serial = date_to_serial(NaiveDate::from_ymd_opt(2024, 6, 15).unwrap());
        let cv = CellValue::Date(serial);
        assert_eq!(cv.to_string(), "2024-06-15");
    }

    #[test]
    fn test_cell_value_date_display_with_time() {
        let dt = NaiveDate::from_ymd_opt(2024, 6, 15)
            .unwrap()
            .and_hms_opt(14, 30, 0)
            .unwrap();
        let serial = datetime_to_serial(dt);
        let cv = CellValue::Date(serial);
        assert_eq!(cv.to_string(), "2024-06-15 14:30:00");
    }

    #[test]
    fn test_cell_value_from_naive_date() {
        let date = NaiveDate::from_ymd_opt(2024, 1, 1).unwrap();
        let cv: CellValue = date.into();
        match cv {
            CellValue::Date(s) => assert_eq!(s, date_to_serial(date)),
            _ => panic!("expected Date variant"),
        }
    }

    #[test]
    fn test_cell_value_from_naive_datetime() {
        let dt = NaiveDate::from_ymd_opt(2024, 1, 1)
            .unwrap()
            .and_hms_opt(12, 0, 0)
            .unwrap();
        let cv: CellValue = dt.into();
        match cv {
            CellValue::Date(s) => assert_eq!(s, datetime_to_serial(dt)),
            _ => panic!("expected Date variant"),
        }
    }

    #[test]
    fn test_is_date_num_fmt() {
        // Date formats: 14-22
        for id in 14..=22 {
            assert!(is_date_num_fmt(id), "expected {id} to be a date format");
        }
        // Time formats: 45-47
        for id in 45..=47 {
            assert!(is_date_num_fmt(id), "expected {id} to be a date format");
        }
        // Not date formats
        assert!(!is_date_num_fmt(0));
        assert!(!is_date_num_fmt(1));
        assert!(!is_date_num_fmt(13));
        assert!(!is_date_num_fmt(23));
        assert!(!is_date_num_fmt(49));
    }

    #[test]
    fn test_is_date_format_code() {
        assert!(is_date_format_code("yyyy-mm-dd"));
        assert!(is_date_format_code("m/d/yyyy"));
        assert!(is_date_format_code("h:mm:ss"));
        assert!(is_date_format_code("dd/mm/yyyy hh:mm"));
        // Not date formats
        assert!(!is_date_format_code("0.00"));
        assert!(!is_date_format_code("#,##0"));
        assert!(!is_date_format_code("0%"));
        // Quoted text should be ignored
        assert!(!is_date_format_code("\"date\"0.00"));
        // Escaped characters should be ignored
        assert!(!is_date_format_code("\\d0.00"));
    }

    #[test]
    fn test_date_early_dates_before_leap_bug() {
        // January 2, 1900 = serial 2
        let date = NaiveDate::from_ymd_opt(1900, 1, 2).unwrap();
        assert_eq!(date_to_serial(date), 2.0);
        assert_eq!(serial_to_date(2.0).unwrap(), date);
    }
}
