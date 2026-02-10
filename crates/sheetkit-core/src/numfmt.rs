//! Number format renderer: converts a (value, format_code) pair into display text.
//!
//! Supports Excel built-in format IDs 0-49, custom numeric patterns
//! (`0`, `#`, `,`, `.`, `%`, `E+`), date/time patterns (`y`, `m`, `d`,
//! `h`, `s`, `AM/PM`), multi-section formats (up to 4 sections separated
//! by `;`), color codes (`[Red]`, `[Blue]`, etc.), conditional sections
//! (`[>100]`), text format (`@`), and fraction formats (`# ?/?`).

use crate::cell::serial_to_date;

/// Map a built-in number format ID (0-49) to its format code string.
pub fn builtin_format_code(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        5 => Some("#,##0_);(#,##0)"),
        6 => Some("#,##0_);[Red](#,##0)"),
        7 => Some("#,##0.00_);(#,##0.00)"),
        8 => Some("#,##0.00_);[Red](#,##0.00)"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("m/d/yyyy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yyyy h:mm"),
        37 => Some("#,##0_);(#,##0)"),
        38 => Some("#,##0_);[Red](#,##0)"),
        39 => Some("#,##0.00_);(#,##0.00)"),
        40 => Some("#,##0.00_);[Red](#,##0.00)"),
        41 => Some(r#"_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)"#),
        42 => Some(r#"_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)"#),
        43 => Some(r#"_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)"#),
        44 => Some(r#"_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)"#),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mm:ss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

/// Format a numeric value using the given format code string.
///
/// Returns the formatted display text. For format codes that contain date/time
/// tokens, the value is interpreted as an Excel serial number.
pub fn format_number(value: f64, format_code: &str) -> String {
    if format_code.is_empty() || format_code.eq_ignore_ascii_case("General") {
        return format_general(value);
    }

    let sections = parse_sections(format_code);

    let has_any_condition = sections.iter().any(|s| extract_condition(s).is_some());
    let section = pick_section(&sections, value);

    let (cleaned, _color) = strip_color_and_condition(section);

    // When multiple sections handle sign presentation, use absolute value:
    // - Standard sign-based sections (>= 2 sections): the negative section
    //   format includes its own sign (parentheses, literal minus, etc.)
    // - Conditional sections: the format encodes its own sign presentation,
    //   so always pass absolute value to avoid double signs
    let use_abs = if has_any_condition {
        sections.len() >= 2
    } else {
        sections.len() >= 2 && value < 0.0
    };
    let effective_value = if use_abs { value.abs() } else { value };

    if cleaned == "@" {
        return format_general(effective_value);
    }

    if is_date_time_format(&cleaned) {
        return format_date_time(effective_value, &cleaned);
    }

    if cleaned.contains('?') && cleaned.contains('/') {
        return format_fraction(effective_value, &cleaned);
    }

    if format_has_unquoted_char(&cleaned, 'E') || format_has_unquoted_char(&cleaned, 'e') {
        return format_scientific(effective_value, &cleaned);
    }

    format_numeric(effective_value, &cleaned)
}

/// Format a numeric value using a built-in format ID.
/// Returns `None` if the ID is not a recognized built-in format.
pub fn format_with_builtin(value: f64, id: u32) -> Option<String> {
    let code = builtin_format_code(id)?;
    Some(format_number(value, code))
}

fn format_general(value: f64) -> String {
    if value == 0.0 {
        return "0".to_string();
    }
    if value.fract() == 0.0 && value.is_finite() && value.abs() < 1e15 {
        return format!("{}", value as i64);
    }
    // Excel displays up to ~11 significant digits in General format.
    let abs = value.abs();
    if (1e-4..1e15).contains(&abs) {
        let s = format!("{:.10}", value);
        trim_trailing_zeros(&s)
    } else if abs < 1e-4 && abs > 0.0 {
        format!("{:.6E}", value)
    } else {
        format!("{}", value)
    }
}

fn trim_trailing_zeros(s: &str) -> String {
    if let Some(dot) = s.find('.') {
        let trimmed = s.trim_end_matches('0');
        if trimmed.len() == dot + 1 {
            trimmed[..dot].to_string()
        } else {
            trimmed.to_string()
        }
    } else {
        s.to_string()
    }
}

fn parse_sections(format_code: &str) -> Vec<&str> {
    let mut sections = Vec::new();
    let mut start = 0;
    let mut in_quotes = false;
    let mut prev_backslash = false;

    for (i, ch) in format_code.char_indices() {
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
        if !in_quotes && ch == ';' {
            sections.push(&format_code[start..i]);
            start = i + 1;
        }
    }
    sections.push(&format_code[start..]);
    sections
}

/// Comparison operator for conditional format sections.
#[derive(Debug, Clone, Copy, PartialEq)]
enum ConditionOp {
    Gt,
    Ge,
    Lt,
    Le,
    Eq,
    Ne,
}

/// A parsed conditional predicate from a format section (e.g., `[>100]`).
#[derive(Debug, Clone, PartialEq)]
struct Condition {
    op: ConditionOp,
    threshold: f64,
}

impl Condition {
    fn matches(&self, value: f64) -> bool {
        match self.op {
            ConditionOp::Gt => value > self.threshold,
            ConditionOp::Ge => value >= self.threshold,
            ConditionOp::Lt => value < self.threshold,
            ConditionOp::Le => value <= self.threshold,
            ConditionOp::Eq => (value - self.threshold).abs() < 1e-12,
            ConditionOp::Ne => (value - self.threshold).abs() >= 1e-12,
        }
    }
}

/// Parse a bracket's inner content as a conditional predicate.
/// Returns `Some(Condition)` for strings like `>100`, `<=0`, `<>5`, `=0`.
fn parse_condition(content: &str) -> Option<Condition> {
    let s = content.trim();
    if s.is_empty() {
        return None;
    }

    // Try two-character operators first, then single-character
    let (op, rest) = if let Some(r) = s.strip_prefix(">=") {
        (ConditionOp::Ge, r)
    } else if let Some(r) = s.strip_prefix("<=") {
        (ConditionOp::Le, r)
    } else if let Some(r) = s.strip_prefix("<>").or_else(|| s.strip_prefix("!=")) {
        (ConditionOp::Ne, r)
    } else if let Some(r) = s.strip_prefix('>') {
        (ConditionOp::Gt, r)
    } else if let Some(r) = s.strip_prefix('<') {
        (ConditionOp::Lt, r)
    } else if let Some(r) = s.strip_prefix('=') {
        (ConditionOp::Eq, r)
    } else {
        return None;
    };

    let threshold: f64 = rest.trim().parse().ok()?;
    Some(Condition { op, threshold })
}

/// Extract the condition (if any) from a format section's bracket content.
fn extract_condition(section: &str) -> Option<Condition> {
    let mut chars = section.chars().peekable();
    while let Some(&ch) = chars.peek() {
        if ch == '[' {
            chars.next();
            let mut bracket_content = String::new();
            while let Some(&c) = chars.peek() {
                if c == ']' {
                    chars.next();
                    break;
                }
                bracket_content.push(c);
                chars.next();
            }
            let lower = bracket_content.to_ascii_lowercase();
            let is_known_non_condition = is_color_code(&lower)
                || lower.starts_with("dbnum")
                || lower.starts_with('$')
                || lower.starts_with("natnum")
                || (lower.starts_with('h') && lower.contains(':'))
                || lower.starts_with("mm")
                || lower.starts_with("ss");
            if !is_known_non_condition {
                if let Some(cond) = parse_condition(&bracket_content) {
                    return Some(cond);
                }
            }
        } else {
            chars.next();
        }
    }
    None
}

/// Pick the format section to apply for a given value.
///
/// When sections contain explicit conditional predicates (e.g., `[>100]`),
/// those conditions are evaluated against the value. Otherwise, the standard
/// Excel sign-based selection (positive / negative / zero / text) is used.
fn pick_section<'a>(sections: &[&'a str], value: f64) -> &'a str {
    // Gather conditions from each section
    let conditions: Vec<Option<Condition>> =
        sections.iter().map(|s| extract_condition(s)).collect();

    let has_any_condition = conditions.iter().any(|c| c.is_some());

    if has_any_condition {
        // Find the first section whose condition matches
        for (i, cond) in conditions.iter().enumerate() {
            if let Some(c) = cond {
                if c.matches(value) {
                    return sections[i];
                }
            }
        }
        // No conditional section matched: use the first section without a condition
        // as the fallback (this is how Excel handles it)
        for (i, cond) in conditions.iter().enumerate() {
            if cond.is_none() {
                return sections[i];
            }
        }
        // All sections have conditions and none matched: use the last section
        return sections.last().unwrap_or(&"General");
    }

    // Standard sign-based selection (no conditional predicates)
    match sections.len() {
        1 => sections[0],
        2 => {
            if value >= 0.0 {
                sections[0]
            } else {
                sections[1]
            }
        }
        3 | 4.. => {
            if value > 0.0 {
                sections[0]
            } else if value < 0.0 {
                sections[1]
            } else {
                sections[2]
            }
        }
        _ => "General",
    }
}

/// Strip color codes and conditional predicates from a format section,
/// returning the cleaned format string and the color name (if any).
fn strip_color_and_condition(section: &str) -> (String, Option<String>) {
    let mut result = String::with_capacity(section.len());
    let mut color = None;
    let mut chars = section.chars().peekable();

    while let Some(&ch) = chars.peek() {
        if ch == '[' {
            let mut bracket_content = String::new();
            chars.next(); // consume '['
            while let Some(&c) = chars.peek() {
                if c == ']' {
                    chars.next(); // consume ']'
                    break;
                }
                bracket_content.push(c);
                chars.next();
            }
            let lower = bracket_content.to_ascii_lowercase();
            if is_color_code(&lower) {
                color = Some(bracket_content);
            } else if lower.starts_with('h') && lower.contains(':') {
                // Elapsed time bracket like [h] or [hh]
                result.push('[');
                result.push_str(&bracket_content);
                result.push(']');
            } else if lower.starts_with("mm") || lower.starts_with("ss") {
                result.push('[');
                result.push_str(&bracket_content);
                result.push(']');
            } else if parse_condition(&bracket_content).is_some() {
                // Conditional predicate -- strip from format string
                // (condition is evaluated during section selection)
            } else if lower.starts_with("dbnum")
                || lower.starts_with("$")
                || lower.starts_with("natnum")
            {
                // Locale or special -- skip
            } else {
                // Unknown bracket, preserve
                result.push('[');
                result.push_str(&bracket_content);
                result.push(']');
            }
        } else {
            result.push(ch);
            chars.next();
        }
    }

    (result, color)
}

fn is_color_code(lower: &str) -> bool {
    matches!(
        lower,
        "red"
            | "blue"
            | "green"
            | "yellow"
            | "cyan"
            | "magenta"
            | "white"
            | "black"
            | "color1"
            | "color2"
            | "color3"
            | "color4"
            | "color5"
            | "color6"
            | "color7"
            | "color8"
            | "color9"
            | "color10"
    )
}

fn is_date_time_format(format: &str) -> bool {
    let mut in_quotes = false;
    let mut prev_backslash = false;
    for ch in format.chars() {
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
        if lower == 'm' {
            return true;
        }
    }
    false
}

fn format_date_time(value: f64, format: &str) -> String {
    let int_part = value.floor() as i64;
    let frac = value.fract().abs();
    let total_seconds = (frac * 86_400.0).round() as u64;
    let mut hours = (total_seconds / 3600) as u32;
    let minutes = ((total_seconds % 3600) / 60) as u32;
    let seconds = (total_seconds % 60) as u32;
    let subsec_frac = (frac * 86_400.0) - (total_seconds as f64);

    let date_opt = serial_to_date(value);
    let (year, month, day) = if let Some(date) = date_opt {
        (date.year() as u32, date.month(), date.day())
    } else {
        (1900, 1, 1)
    };

    let has_ampm = {
        let lower = format.to_ascii_lowercase();
        lower.contains("am/pm") || lower.contains("a/p")
    };

    let mut ampm_str = "";
    if has_ampm {
        if hours == 0 {
            hours = 12;
            ampm_str = "AM";
        } else if hours < 12 {
            ampm_str = "AM";
        } else if hours == 12 {
            ampm_str = "PM";
        } else {
            hours -= 12;
            ampm_str = "PM";
        }
    }

    let month_names_short = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ];
    let month_names_long = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];

    let mut result = String::with_capacity(format.len() * 2);
    let chars: Vec<char> = format.chars().collect();
    let len = chars.len();
    let mut i = 0;
    let mut in_quotes = false;

    while i < len {
        let ch = chars[i];

        if ch == '"' {
            in_quotes = !in_quotes;
            i += 1;
            continue;
        }
        if in_quotes {
            result.push(ch);
            i += 1;
            continue;
        }
        if ch == '\\' && i + 1 < len {
            result.push(chars[i + 1]);
            i += 2;
            continue;
        }

        // Skip padding/spacing characters: _ followed by a character
        if ch == '_' && i + 1 < len {
            result.push(' ');
            i += 2;
            continue;
        }
        // Skip * repetition fill
        if ch == '*' && i + 1 < len {
            i += 2;
            continue;
        }

        let lower = ch.to_ascii_lowercase();

        if lower == 'y' {
            let count = count_char(&chars, i, 'y');
            if count <= 2 {
                result.push_str(&format!("{:02}", year % 100));
            } else {
                result.push_str(&format!("{:04}", year));
            }
            i += count;
            continue;
        }

        if lower == 'm' {
            let count = count_char(&chars, i, 'm');
            if is_m_minute_context(&chars, i) {
                // Minutes
                if count == 1 {
                    result.push_str(&format!("{}", minutes));
                } else {
                    result.push_str(&format!("{:02}", minutes));
                }
            } else {
                // Month
                match count {
                    1 => result.push_str(&format!("{}", month)),
                    2 => result.push_str(&format!("{:02}", month)),
                    3 => {
                        if (1..=12).contains(&month) {
                            result.push_str(month_names_short[(month - 1) as usize]);
                        }
                    }
                    4 => {
                        if (1..=12).contains(&month) {
                            result.push_str(month_names_long[(month - 1) as usize]);
                        }
                    }
                    _ => {
                        result.push_str(&format!("{:02}", month));
                    }
                }
            }
            i += count;
            continue;
        }

        if lower == 'd' {
            let count = count_char(&chars, i, 'd');
            match count {
                1 => result.push_str(&format!("{}", day)),
                2 => result.push_str(&format!("{:02}", day)),
                3 => {
                    if let Some(date) = date_opt {
                        let day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
                        let wd = date.weekday().num_days_from_monday() as usize;
                        result.push_str(day_names[wd]);
                    }
                }
                _ => {
                    if let Some(date) = date_opt {
                        let day_names = [
                            "Monday",
                            "Tuesday",
                            "Wednesday",
                            "Thursday",
                            "Friday",
                            "Saturday",
                            "Sunday",
                        ];
                        let wd = date.weekday().num_days_from_monday() as usize;
                        result.push_str(day_names[wd]);
                    }
                }
            }
            i += count;
            continue;
        }

        if lower == 'h' {
            let count = count_char(&chars, i, 'h');
            // Check for elapsed hours [h]
            if i > 0 && chars[i - 1] == '[' {
                // Elapsed hours: total hours from the serial number
                let serial_days = value.floor() as i64;
                let elapsed_h = (serial_days as u64) * 24 + total_seconds / 3600;
                // Find the closing bracket after the 'h' tokens
                let mut end = i + count;
                if end < len && chars[end] == ']' {
                    end += 1; // skip the ']'
                }
                result.push_str(&format!("{}", elapsed_h));
                i = end;
                continue;
            }
            if count == 1 {
                result.push_str(&format!("{}", hours));
            } else {
                result.push_str(&format!("{:02}", hours));
            }
            i += count;
            continue;
        }

        if lower == 's' {
            let count = count_char(&chars, i, 's');
            if count == 1 {
                result.push_str(&format!("{}", seconds));
            } else {
                result.push_str(&format!("{:02}", seconds));
            }
            i += count;
            continue;
        }

        // AM/PM or A/P
        if lower == 'a' {
            if i + 4 < len {
                let slice: String = chars[i..i + 5].iter().collect();
                if slice.eq_ignore_ascii_case("AM/PM") {
                    result.push_str(ampm_str);
                    i += 5;
                    continue;
                }
            }
            if i + 2 < len {
                let slice: String = chars[i..i + 3].iter().collect();
                if slice.eq_ignore_ascii_case("A/P") {
                    if ampm_str == "AM" {
                        result.push('A');
                    } else {
                        result.push('P');
                    }
                    i += 3;
                    continue;
                }
            }
            result.push(ch);
            i += 1;
            continue;
        }

        // Elapsed time brackets [h], [mm], [ss]
        if ch == '[' {
            // Look ahead to see if this is [h], [hh], [mm], [ss]
            if i + 2 < len && chars[i + 1].eq_ignore_ascii_case(&'h') {
                // Let the 'h' handler deal with it -- pass the '['
                result.push(ch);
                i += 1;
                continue;
            }
            if i + 2 < len && chars[i + 1].eq_ignore_ascii_case(&'m') {
                let count = count_char(&chars, i + 1, 'm');
                let end = i + 1 + count;
                if end < len && chars[end] == ']' {
                    let elapsed_m = (int_part as u64) * 24 * 60 + total_seconds / 60;
                    result.push_str(&format!("{}", elapsed_m));
                    i = end + 1;
                    continue;
                }
            }
            if i + 2 < len && chars[i + 1].eq_ignore_ascii_case(&'s') {
                let count = count_char(&chars, i + 1, 's');
                let end = i + 1 + count;
                if end < len && chars[end] == ']' {
                    let elapsed_s = (int_part as u64) * 24 * 3600 + total_seconds;
                    result.push_str(&format!("{}", elapsed_s));
                    i = end + 1;
                    continue;
                }
            }
            result.push(ch);
            i += 1;
            continue;
        }

        if ch == '.' && i + 1 < len && chars[i + 1] == '0' {
            // Fractional seconds
            result.push('.');
            let count = count_char(&chars, i + 1, '0');
            let sub = subsec_frac.abs();
            let digits = format!("{:.*}", count, sub);
            // digits is like "0.xxx", take the part after '.'
            if let Some(dot_pos) = digits.find('.') {
                result.push_str(&digits[dot_pos + 1..]);
            }
            i += 1 + count;
            continue;
        }

        // Pass through separators and other literal characters
        result.push(ch);
        i += 1;
    }

    result
}

/// Determine whether an 'm' token at position `pos` in the format chars
/// should be interpreted as minutes (true) or months (false).
/// Minutes if there is a preceding 'h' or a following 's' (skipping separators).
fn is_m_minute_context(chars: &[char], pos: usize) -> bool {
    // Look backwards for 'h', skipping ':', ' ', digits, brackets
    let mut j = pos;
    while j > 0 {
        j -= 1;
        let c = chars[j].to_ascii_lowercase();
        if c == 'h' {
            return true;
        }
        if c == ':' || c == ' ' || c == ']' || c == '[' {
            continue;
        }
        break;
    }
    // Look forwards past the 'm' run for 's', skipping ':', ' ', digits, '.'
    let m_count = count_char(chars, pos, 'm');
    let mut k = pos + m_count;
    while k < chars.len() {
        let c = chars[k].to_ascii_lowercase();
        if c == 's' {
            return true;
        }
        if c == ':' || c == ' ' || c == '0' || c == '.' {
            k += 1;
            continue;
        }
        break;
    }
    false
}

fn count_char(chars: &[char], start: usize, target: char) -> usize {
    let lower_target = target.to_ascii_lowercase();
    let mut count = 0;
    let mut i = start;
    while i < chars.len() && chars[i].to_ascii_lowercase() == lower_target {
        count += 1;
        i += 1;
    }
    count
}

fn format_numeric(value: f64, format: &str) -> String {
    let is_negative = value < 0.0;
    let abs_val = value.abs();

    // Parse the format to understand: digit placeholders, decimal places,
    // comma grouping, percent, underscore/star padding, and literal text.
    let has_percent = format_has_unquoted_char(format, '%');
    let display_val = if has_percent {
        abs_val * 100.0
    } else {
        abs_val
    };

    // Count decimal places from the format
    let decimal_places = count_decimal_places(format);

    // Check for thousands separator (comma grouping)
    let has_comma_grouping = has_thousands_separator(format);

    // Check for trailing commas (divide by 1000 per comma)
    let trailing_comma_count = count_trailing_commas(format);
    let display_val = display_val / 1000f64.powi(trailing_comma_count as i32);

    // Round to the number of decimal places
    let rounded = if decimal_places > 0 {
        let factor = 10f64.powi(decimal_places as i32);
        (display_val * factor).round() / factor
    } else {
        display_val.round()
    };

    let int_part = rounded.trunc() as u64;
    let frac_part =
        ((rounded - rounded.trunc()).abs() * 10f64.powi(decimal_places as i32)).round() as u64;

    // Format integer part
    let int_str = format!("{}", int_part);
    let int_display = if has_comma_grouping {
        add_thousands_separators(&int_str)
    } else {
        int_str.clone()
    };

    // Count required integer digits (from '0' placeholders in integer part)
    let min_int_digits = count_integer_zeros(format);
    let padded_int = if int_display.len() < min_int_digits && int_part == 0 {
        let needed = min_int_digits - int_display.len();
        let mut s = "0".repeat(needed);
        s.push_str(&int_display);
        if has_comma_grouping {
            add_thousands_separators(&s)
        } else {
            s
        }
    } else {
        int_display
    };

    // Build the output by walking through the format pattern
    // For simplicity, output the formatted number with appropriate decorations
    let mut output = String::with_capacity(format.len() + 10);
    let chars: Vec<char> = format.chars().collect();
    let len = chars.len();
    let mut i = 0;
    let mut in_quotes = false;
    let mut number_placed = false;

    while i < len {
        let ch = chars[i];

        if ch == '"' {
            in_quotes = !in_quotes;
            i += 1;
            continue;
        }
        if in_quotes {
            output.push(ch);
            i += 1;
            continue;
        }
        if ch == '\\' && i + 1 < len {
            output.push(chars[i + 1]);
            i += 2;
            continue;
        }
        if ch == '_' && i + 1 < len {
            output.push(' ');
            i += 2;
            continue;
        }
        if ch == '*' && i + 1 < len {
            i += 2;
            continue;
        }

        if (ch == '0' || ch == '#' || ch == ',') && !number_placed {
            // Find the end of the numeric pattern
            let num_end = find_numeric_end(&chars, i);
            // Place the formatted number
            let num_str = if decimal_places > 0 {
                let frac_str = format!("{:0>width$}", frac_part, width = decimal_places);
                format!("{}.{}", padded_int, frac_str)
            } else {
                padded_int.clone()
            };

            if is_negative {
                output.push('-');
            }
            output.push_str(&num_str);
            number_placed = true;
            i = num_end;
            continue;
        }

        if ch == '.' && !number_placed {
            // This dot is part of the numeric pattern, handle together
            continue;
        }

        if ch == '%' {
            output.push('%');
            i += 1;
            continue;
        }

        if ch == '(' || ch == ')' || ch == '-' || ch == '+' || ch == ' ' || ch == ':' || ch == '/' {
            output.push(ch);
            i += 1;
            continue;
        }

        if ch == '0' || ch == '#' || ch == ',' || ch == '.' {
            i += 1;
            continue;
        }

        output.push(ch);
        i += 1;
    }

    // If no number was placed, check whether the format actually has digit placeholders.
    // Formats like `"-"` (pure literal text) should not emit a number at all.
    if !number_placed {
        let has_digit_placeholder = format.chars().any(|c| c == '0' || c == '#');
        if has_digit_placeholder {
            if is_negative {
                output.push('-');
            }
            if decimal_places > 0 {
                let frac_str = format!("{:0>width$}", frac_part, width = decimal_places);
                output.push_str(&format!("{}.{}", padded_int, frac_str));
            } else {
                output.push_str(&padded_int);
            }
        }
    }

    output
}

fn format_has_unquoted_char(format: &str, target: char) -> bool {
    let mut in_quotes = false;
    let mut prev_backslash = false;
    for ch in format.chars() {
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
        if !in_quotes && ch == target {
            return true;
        }
    }
    false
}

fn count_decimal_places(format: &str) -> usize {
    let mut in_quotes = false;
    let mut prev_backslash = false;
    let mut found_dot = false;
    let mut count = 0;

    for ch in format.chars() {
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
        if ch == '.' && !found_dot {
            found_dot = true;
            continue;
        }
        if found_dot && (ch == '0' || ch == '#') {
            count += 1;
        } else if found_dot && ch != '0' && ch != '#' {
            break;
        }
    }
    count
}

fn has_thousands_separator(format: &str) -> bool {
    let mut in_quotes = false;
    let mut prev_backslash = false;
    let chars: Vec<char> = format.chars().collect();

    for (i, &ch) in chars.iter().enumerate() {
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
        // A comma is a thousands separator if it appears between digit placeholders
        if ch == ',' {
            let has_digit_before = chars[..i].iter().rev().any(|&c| c == '0' || c == '#');
            let has_digit_after = chars[i + 1..].iter().any(|&c| c == '0' || c == '#');
            if has_digit_before && has_digit_after {
                return true;
            }
        }
    }
    false
}

fn count_trailing_commas(format: &str) -> usize {
    let mut in_quotes = false;
    let mut prev_backslash = false;
    let chars: Vec<char> = format.chars().collect();
    let mut count = 0;

    // Find the last digit placeholder, then count commas after it
    let mut last_digit_pos = None;
    for (i, &ch) in chars.iter().enumerate() {
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
        if ch == '0' || ch == '#' {
            last_digit_pos = Some(i);
        }
    }

    if let Some(pos) = last_digit_pos {
        for &ch in &chars[pos + 1..] {
            if ch == ',' {
                count += 1;
            } else {
                break;
            }
        }
    }
    count
}

fn count_integer_zeros(format: &str) -> usize {
    let mut in_quotes = false;
    let mut prev_backslash = false;
    let mut count = 0;
    let mut found_dot = false;

    for ch in format.chars() {
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
        if ch == '.' {
            found_dot = true;
            continue;
        }
        if !found_dot && ch == '0' {
            count += 1;
        }
    }
    count
}

fn add_thousands_separators(s: &str) -> String {
    let bytes = s.as_bytes();
    let len = bytes.len();
    if len <= 3 {
        return s.to_string();
    }
    let mut result = String::with_capacity(len + len / 3);
    let remainder = len % 3;
    if remainder > 0 {
        result.push_str(&s[..remainder]);
        if len > remainder {
            result.push(',');
        }
    }
    for (i, chunk) in s.as_bytes()[remainder..].chunks(3).enumerate() {
        if i > 0 {
            result.push(',');
        }
        result.push_str(std::str::from_utf8(chunk).unwrap_or(""));
    }
    result
}

fn find_numeric_end(chars: &[char], start: usize) -> usize {
    let mut i = start;
    let mut in_quotes = false;
    while i < chars.len() {
        let ch = chars[i];
        if ch == '"' {
            in_quotes = !in_quotes;
            i += 1;
            continue;
        }
        if in_quotes {
            i += 1;
            continue;
        }
        if ch == '0' || ch == '#' || ch == ',' || ch == '.' {
            i += 1;
        } else {
            break;
        }
    }
    i
}

fn format_scientific(value: f64, format: &str) -> String {
    let decimal_places = count_decimal_places(format);
    let formatted = format!("{:.*E}", decimal_places, value.abs());

    // Split into mantissa and exponent
    let parts: Vec<&str> = formatted.split('E').collect();
    if parts.len() != 2 {
        return formatted;
    }

    let mantissa = parts[0];
    let exp_str = parts[1];
    let exp: i32 = exp_str.parse().unwrap_or(0);

    // Count the '0' digits after E+/E- in the format to determine exponent width
    let exp_width = count_exponent_zeros(format).max(2);

    // Determine sign character for exponent
    let has_plus = format.contains("E+") || format.contains("e+");
    let exp_sign = if exp >= 0 {
        if has_plus {
            "+"
        } else {
            ""
        }
    } else {
        "-"
    };

    let exp_display = format!(
        "{}{:0>width$}",
        exp_sign,
        exp.unsigned_abs(),
        width = exp_width
    );

    let sign = if value < 0.0 { "-" } else { "" };

    // Determine E vs e
    let e_char = if format.contains('e') { 'e' } else { 'E' };

    format!("{}{}{}{}", sign, mantissa, e_char, exp_display)
}

fn count_exponent_zeros(format: &str) -> usize {
    let upper = format.to_uppercase();
    if let Some(pos) = upper.find("E+").or_else(|| upper.find("E-")) {
        let after = &format[pos + 2..];
        after.chars().take_while(|&c| c == '0').count()
    } else {
        2
    }
}

fn format_fraction(value: f64, format: &str) -> String {
    let abs = value.abs();
    let whole = abs.floor() as i64;
    let frac = abs - whole as f64;

    let sign = if value < 0.0 { "-" } else { "" };

    // Determine the maximum denominator from the denominator width (digits after '/')
    let denom_q_count = format
        .split('/')
        .nth(1)
        .map(|s| s.chars().filter(|&c| c == '?').count())
        .unwrap_or(1);
    let max_denom = if denom_q_count >= 4 {
        9999
    } else if denom_q_count >= 3 {
        999
    } else if denom_q_count >= 2 {
        99
    } else {
        9
    };

    if frac < 1e-10 {
        if format.contains('#') {
            return format!("{}{}", sign, whole);
        }
        return format!("{}{}    ", sign, whole);
    }

    // Find the best rational approximation
    let (num, den) = best_fraction(frac, max_denom);

    let has_whole = format.contains('#');

    if has_whole {
        if whole > 0 {
            format!("{}{} {}/{}", sign, whole, num, den)
        } else {
            format!("{}{}/{}", sign, num, den)
        }
    } else {
        let total_num = whole as u64 * den + num;
        format!("{}{}/{}", sign, total_num, den)
    }
}

fn best_fraction(value: f64, max_denom: u64) -> (u64, u64) {
    if value <= 0.0 {
        return (0, 1);
    }
    let mut best_num = 0u64;
    let mut best_den = 1u64;
    let mut best_err = value.abs();

    for den in 1..=max_denom {
        let num = (value * den as f64).round() as u64;
        if num == 0 {
            continue;
        }
        let err = (value - num as f64 / den as f64).abs();
        if err < best_err {
            best_err = err;
            best_num = num;
            best_den = den;
        }
        if best_err < 1e-10 {
            break;
        }
    }
    (best_num, best_den)
}

use chrono::Datelike;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_builtin_format_code_general() {
        assert_eq!(builtin_format_code(0), Some("General"));
    }

    #[test]
    fn test_builtin_format_code_integer() {
        assert_eq!(builtin_format_code(1), Some("0"));
    }

    #[test]
    fn test_builtin_format_code_decimal() {
        assert_eq!(builtin_format_code(2), Some("0.00"));
    }

    #[test]
    fn test_builtin_format_code_thousands() {
        assert_eq!(builtin_format_code(3), Some("#,##0"));
    }

    #[test]
    fn test_builtin_format_code_date() {
        assert_eq!(builtin_format_code(14), Some("m/d/yyyy"));
    }

    #[test]
    fn test_builtin_format_code_text() {
        assert_eq!(builtin_format_code(49), Some("@"));
    }

    #[test]
    fn test_builtin_format_code_unknown() {
        assert_eq!(builtin_format_code(100), None);
        assert_eq!(builtin_format_code(50), None);
    }

    #[test]
    fn test_format_general_zero() {
        assert_eq!(format_number(0.0, "General"), "0");
    }

    #[test]
    fn test_format_general_integer() {
        assert_eq!(format_number(42.0, "General"), "42");
        assert_eq!(format_number(-100.0, "General"), "-100");
    }

    #[test]
    fn test_format_general_decimal() {
        assert_eq!(format_number(3.14, "General"), "3.14");
    }

    #[test]
    fn test_format_general_large_number() {
        assert_eq!(format_number(1000000.0, "General"), "1000000");
    }

    #[test]
    fn test_format_integer() {
        assert_eq!(format_number(42.0, "0"), "42");
        assert_eq!(format_number(42.7, "0"), "43");
        assert_eq!(format_number(0.0, "0"), "0");
    }

    #[test]
    fn test_format_decimal_2() {
        assert_eq!(format_number(3.14159, "0.00"), "3.14");
        assert_eq!(format_number(3.0, "0.00"), "3.00");
        assert_eq!(format_number(0.5, "0.00"), "0.50");
    }

    #[test]
    fn test_format_thousands() {
        assert_eq!(format_number(1234.0, "#,##0"), "1,234");
        assert_eq!(format_number(1234567.0, "#,##0"), "1,234,567");
        assert_eq!(format_number(999.0, "#,##0"), "999");
        assert_eq!(format_number(0.0, "#,##0"), "0");
    }

    #[test]
    fn test_format_thousands_decimal() {
        assert_eq!(format_number(1234.56, "#,##0.00"), "1,234.56");
        assert_eq!(format_number(0.0, "#,##0.00"), "0.00");
    }

    #[test]
    fn test_format_percent() {
        assert_eq!(format_number(0.75, "0%"), "75%");
        assert_eq!(format_number(0.5, "0%"), "50%");
        assert_eq!(format_number(1.0, "0%"), "100%");
    }

    #[test]
    fn test_format_percent_decimal() {
        assert_eq!(format_number(0.7534, "0.00%"), "75.34%");
        assert_eq!(format_number(0.5, "0.00%"), "50.00%");
    }

    #[test]
    fn test_format_scientific() {
        let result = format_number(1234.5, "0.00E+00");
        assert_eq!(result, "1.23E+03");
    }

    #[test]
    fn test_format_scientific_small() {
        let result = format_number(0.001, "0.00E+00");
        assert_eq!(result, "1.00E-03");
    }

    #[test]
    fn test_format_date_mdy() {
        // 2024-01-15 = serial 45306
        let serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 1, 15).unwrap());
        assert_eq!(format_number(serial, "m/d/yyyy"), "1/15/2024");
    }

    #[test]
    fn test_format_date_dmy() {
        let serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 3, 5).unwrap());
        assert_eq!(format_number(serial, "d-mmm-yy"), "5-Mar-24");
    }

    #[test]
    fn test_format_date_dm() {
        let serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 6, 15).unwrap());
        assert_eq!(format_number(serial, "d-mmm"), "15-Jun");
    }

    #[test]
    fn test_format_date_my() {
        let serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 12, 1).unwrap());
        assert_eq!(format_number(serial, "mmm-yy"), "Dec-24");
    }

    #[test]
    fn test_format_time_hm_ampm() {
        let serial = crate::cell::datetime_to_serial(
            chrono::NaiveDate::from_ymd_opt(2024, 1, 1)
                .unwrap()
                .and_hms_opt(14, 30, 0)
                .unwrap(),
        );
        assert_eq!(format_number(serial, "h:mm AM/PM"), "2:30 PM");
    }

    #[test]
    fn test_format_time_hm_ampm_morning() {
        let serial = crate::cell::datetime_to_serial(
            chrono::NaiveDate::from_ymd_opt(2024, 1, 1)
                .unwrap()
                .and_hms_opt(9, 5, 0)
                .unwrap(),
        );
        assert_eq!(format_number(serial, "h:mm AM/PM"), "9:05 AM");
    }

    #[test]
    fn test_format_time_hms() {
        let serial = crate::cell::datetime_to_serial(
            chrono::NaiveDate::from_ymd_opt(2024, 1, 1)
                .unwrap()
                .and_hms_opt(14, 30, 45)
                .unwrap(),
        );
        assert_eq!(format_number(serial, "h:mm:ss"), "14:30:45");
    }

    #[test]
    fn test_format_time_hm_24h() {
        let serial = crate::cell::datetime_to_serial(
            chrono::NaiveDate::from_ymd_opt(2024, 1, 1)
                .unwrap()
                .and_hms_opt(14, 30, 0)
                .unwrap(),
        );
        assert_eq!(format_number(serial, "h:mm"), "14:30");
    }

    #[test]
    fn test_format_datetime_combined() {
        let serial = crate::cell::datetime_to_serial(
            chrono::NaiveDate::from_ymd_opt(2024, 1, 15)
                .unwrap()
                .and_hms_opt(14, 30, 0)
                .unwrap(),
        );
        assert_eq!(format_number(serial, "m/d/yyyy h:mm"), "1/15/2024 14:30");
    }

    #[test]
    fn test_format_date_yyyy_mm_dd() {
        let serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 6, 15).unwrap());
        assert_eq!(format_number(serial, "yyyy-mm-dd"), "2024-06-15");
    }

    #[test]
    fn test_format_date_dd_mm_yyyy() {
        let serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 6, 15).unwrap());
        assert_eq!(format_number(serial, "dd/mm/yyyy"), "15/06/2024");
    }

    #[test]
    fn test_format_text_passthrough() {
        assert_eq!(format_number(42.0, "@"), "42");
    }

    #[test]
    fn test_format_multi_section_positive_negative() {
        assert_eq!(format_number(42.0, "#,##0;-#,##0"), "42");
        assert_eq!(format_number(-42.0, "#,##0;-#,##0"), "-42");
    }

    #[test]
    fn test_format_multi_section_three_parts() {
        assert_eq!(format_number(42.0, "#,##0;-#,##0;\"zero\""), "42");
        assert_eq!(format_number(-42.0, "#,##0;-#,##0;\"zero\""), "-42");
    }

    #[test]
    fn test_format_color_stripped() {
        assert_eq!(format_number(42.0, "[Red]0"), "42");
        assert_eq!(format_number(42.0, "[Blue]0.00"), "42.00");
    }

    #[test]
    fn test_format_with_builtin_general() {
        assert_eq!(format_with_builtin(42.0, 0), Some("42".to_string()));
    }

    #[test]
    fn test_format_with_builtin_percent() {
        assert_eq!(format_with_builtin(0.5, 9), Some("50%".to_string()));
    }

    #[test]
    fn test_format_with_builtin_unknown() {
        assert_eq!(format_with_builtin(42.0, 100), None);
    }

    #[test]
    fn test_format_fraction_simple() {
        let result = format_number(1.5, "# ?/?");
        assert_eq!(result, "1 1/2");
    }

    #[test]
    fn test_format_fraction_two_digit() {
        let result = format_number(0.333, "# ??/??");
        // Should approximate to something close to 1/3
        assert!(result.contains("/"), "result was: {}", result);
    }

    #[test]
    fn test_format_negative_in_parens() {
        let result = format_number(-1234.0, "#,##0_);(#,##0)");
        assert!(result.contains("1,234"), "result was: {}", result);
        assert!(result.contains("("), "result was: {}", result);
    }

    #[test]
    fn test_format_empty_format_uses_general() {
        assert_eq!(format_number(42.0, ""), "42");
    }

    #[test]
    fn test_format_date_long_month() {
        let serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 1, 15).unwrap());
        assert_eq!(format_number(serial, "d mmmm yyyy"), "15 January 2024");
    }

    #[test]
    fn test_format_time_ampm_midnight() {
        let serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 1, 1).unwrap());
        // Midnight = no fractional part
        assert_eq!(format_number(serial, "h:mm AM/PM"), "12:00 AM");
    }

    #[test]
    fn test_format_time_ampm_noon() {
        let serial = crate::cell::datetime_to_serial(
            chrono::NaiveDate::from_ymd_opt(2024, 1, 1)
                .unwrap()
                .and_hms_opt(12, 0, 0)
                .unwrap(),
        );
        assert_eq!(format_number(serial, "h:mm AM/PM"), "12:00 PM");
    }

    #[test]
    fn test_format_builtin_mmss() {
        let serial = crate::cell::datetime_to_serial(
            chrono::NaiveDate::from_ymd_opt(2024, 1, 1)
                .unwrap()
                .and_hms_opt(0, 5, 30)
                .unwrap(),
        );
        assert_eq!(format_number(serial, "mm:ss"), "05:30");
    }

    #[test]
    fn test_format_general_negative_decimal() {
        let result = format_number(-3.14, "General");
        assert_eq!(result, "-3.14");
    }

    #[test]
    fn test_format_date_two_digit_year() {
        let serial =
            crate::cell::date_to_serial(chrono::NaiveDate::from_ymd_opt(2024, 6, 15).unwrap());
        assert_eq!(format_number(serial, "yy"), "24");
    }

    #[test]
    fn test_format_thousands_separator_with_zero() {
        assert_eq!(add_thousands_separators("0"), "0");
        assert_eq!(add_thousands_separators("100"), "100");
        assert_eq!(add_thousands_separators("1000"), "1,000");
        assert_eq!(add_thousands_separators("1000000"), "1,000,000");
    }

    #[test]
    fn test_parse_sections_single() {
        let sections = parse_sections("0.00");
        assert_eq!(sections, vec!["0.00"]);
    }

    #[test]
    fn test_parse_sections_multi() {
        let sections = parse_sections("0.00;-0.00;\"zero\"");
        assert_eq!(sections, vec!["0.00", "-0.00", "\"zero\""]);
    }

    #[test]
    fn test_parse_sections_quoted_semicolon() {
        let sections = parse_sections("\"a;b\"");
        assert_eq!(sections, vec!["\"a;b\""]);
    }

    #[test]
    fn test_strip_color() {
        let (cleaned, color) = strip_color_and_condition("[Red]0.00");
        assert_eq!(cleaned, "0.00");
        assert_eq!(color, Some("Red".to_string()));
    }

    #[test]
    fn test_strip_condition() {
        let (cleaned, _) = strip_color_and_condition("[>100]0.00");
        assert_eq!(cleaned, "0.00");
    }

    #[test]
    fn test_is_date_time_format_checks() {
        assert!(is_date_time_format("yyyy-mm-dd"));
        assert!(is_date_time_format("h:mm:ss"));
        assert!(is_date_time_format("m/d/yyyy"));
        assert!(!is_date_time_format("0.00"));
        assert!(!is_date_time_format("#,##0"));
        assert!(!is_date_time_format("\"yyyy\"0"));
    }

    #[test]
    fn test_count_decimal_places_none() {
        assert_eq!(count_decimal_places("0"), 0);
        assert_eq!(count_decimal_places("#,##0"), 0);
    }

    #[test]
    fn test_count_decimal_places_two() {
        assert_eq!(count_decimal_places("0.00"), 2);
        assert_eq!(count_decimal_places("#,##0.00"), 2);
    }

    #[test]
    fn test_count_decimal_places_three() {
        assert_eq!(count_decimal_places("0.000"), 3);
    }

    #[test]
    fn test_parse_condition_operators() {
        let c = parse_condition(">100").unwrap();
        assert_eq!(c.op, ConditionOp::Gt);
        assert_eq!(c.threshold, 100.0);

        let c = parse_condition(">=50").unwrap();
        assert_eq!(c.op, ConditionOp::Ge);
        assert_eq!(c.threshold, 50.0);

        let c = parse_condition("<1000").unwrap();
        assert_eq!(c.op, ConditionOp::Lt);
        assert_eq!(c.threshold, 1000.0);

        let c = parse_condition("<=0").unwrap();
        assert_eq!(c.op, ConditionOp::Le);
        assert_eq!(c.threshold, 0.0);

        let c = parse_condition("=0").unwrap();
        assert_eq!(c.op, ConditionOp::Eq);
        assert_eq!(c.threshold, 0.0);

        let c = parse_condition("<>5").unwrap();
        assert_eq!(c.op, ConditionOp::Ne);
        assert_eq!(c.threshold, 5.0);

        let c = parse_condition("!=5").unwrap();
        assert_eq!(c.op, ConditionOp::Ne);
        assert_eq!(c.threshold, 5.0);

        assert!(parse_condition("Red").is_none());
        assert!(parse_condition("").is_none());
    }

    #[test]
    fn test_condition_matches() {
        let c = Condition {
            op: ConditionOp::Gt,
            threshold: 100.0,
        };
        assert!(c.matches(150.0));
        assert!(!c.matches(100.0));
        assert!(!c.matches(50.0));
    }

    #[test]
    fn test_conditional_two_sections_color_and_condition() {
        // [Red][>100]0;[Blue][<=100]0
        let fmt = "[Red][>100]0;[Blue][<=100]0";
        assert_eq!(format_number(150.0, fmt), "150");
        assert_eq!(format_number(50.0, fmt), "50");
        assert_eq!(format_number(100.0, fmt), "100");
    }

    #[test]
    fn test_conditional_three_sections_cascading() {
        // [>1000]#,##0;[>100]0.0;0.00
        let fmt = "[>1000]#,##0;[>100]0.0;0.00";
        assert_eq!(format_number(5000.0, fmt), "5,000");
        assert_eq!(format_number(500.0, fmt), "500.0");
        assert_eq!(format_number(50.0, fmt), "50.00");
    }

    #[test]
    fn test_conditional_equals_zero() {
        // [=0]"zero";General -- value 0 matches first section, value 42 falls through
        let fmt = "[=0]\"zero\";0";
        assert_eq!(format_number(0.0, fmt), "zero");
        assert_eq!(format_number(42.0, fmt), "42");
    }

    #[test]
    fn test_conditional_with_sign_format() {
        // [Red][>0]+0;[Blue][<0]-0;0
        let fmt = "[Red][>0]+0;[Blue][<0]-0;0";
        assert_eq!(format_number(5.0, fmt), "+5");
        assert_eq!(format_number(-3.0, fmt), "-3");
        assert_eq!(format_number(0.0, fmt), "0");
    }

    #[test]
    fn test_extract_condition_from_section() {
        let cond = extract_condition("[Red][>100]0.00").unwrap();
        assert_eq!(cond.op, ConditionOp::Gt);
        assert_eq!(cond.threshold, 100.0);

        let cond = extract_condition("[<=0]0");
        assert!(cond.is_some());
        assert_eq!(cond.unwrap().op, ConditionOp::Le);

        assert!(extract_condition("[Red]0.00").is_none());
        assert!(extract_condition("0.00").is_none());
    }
}
