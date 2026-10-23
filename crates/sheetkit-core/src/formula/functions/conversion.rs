//! Checked conversions from formula-derived numbers to integer domains.
//!
//! Formula arguments are arbitrary `f64` values. Every conversion to an
//! integer type goes through these helpers so non-finite or out-of-range
//! input becomes a spreadsheet error instead of a saturating cast, an
//! arithmetic overflow, or a panic in a downstream constructor.

/// Truncate `value` toward zero and return it when it is finite and lies in
/// the inclusive range `[min, max]`. Returns `None` otherwise.
pub(crate) fn checked_trunc_i64(value: f64, min: i64, max: i64) -> Option<i64> {
    if !value.is_finite() {
        return None;
    }
    let truncated = value.trunc();
    if truncated < min as f64 || truncated > max as f64 {
        return None;
    }
    // In-range after the float comparison; the cast only saturates at the
    // i64::MAX boundary, which the integer comparison below still enforces.
    let integer = truncated as i64;
    (min..=max).contains(&integer).then_some(integer)
}

/// Truncate `value` toward zero into the full `i32` range.
pub(crate) fn checked_trunc_i32(value: f64) -> Option<i32> {
    checked_trunc_i64(value, i64::from(i32::MIN), i64::from(i32::MAX))
        .and_then(|integer| i32::try_from(integer).ok())
}

/// Truncate `value` toward zero into the inclusive `i32` range `[min, max]`.
pub(crate) fn checked_trunc_i32_in(value: f64, min: i32, max: i32) -> Option<i32> {
    checked_trunc_i64(value, i64::from(min), i64::from(max))
        .and_then(|integer| i32::try_from(integer).ok())
}

/// Truncate a non-negative character count or position toward zero.
///
/// Returns `None` for non-finite or negative input. Finite values beyond
/// `usize::MAX` saturate, because every such count already exceeds any
/// string length; callers must still combine counts with checked or
/// saturating arithmetic.
pub(crate) fn checked_trunc_count(value: f64) -> Option<usize> {
    if !value.is_finite() {
        return None;
    }
    let truncated = value.trunc();
    if truncated < 0.0 {
        return None;
    }
    if truncated >= usize::MAX as f64 {
        return Some(usize::MAX);
    }
    Some(truncated as usize)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn trunc_i64_rejects_non_finite_and_out_of_range() {
        assert_eq!(checked_trunc_i64(f64::NAN, -10, 10), None);
        assert_eq!(checked_trunc_i64(f64::INFINITY, -10, 10), None);
        assert_eq!(checked_trunc_i64(10.9, -10, 10), Some(10));
        assert_eq!(checked_trunc_i64(-10.9, -10, 10), Some(-10));
        assert_eq!(checked_trunc_i64(11.0, -10, 10), None);
        assert_eq!(checked_trunc_i64(1e300, i64::MIN, i64::MAX), None);
        assert_eq!(checked_trunc_i64(-0.5, 0, 1), Some(0));
    }

    #[test]
    fn trunc_i32_covers_full_and_bounded_ranges() {
        assert_eq!(checked_trunc_i32(2_147_483_647.9), Some(i32::MAX));
        assert_eq!(checked_trunc_i32(2_147_483_648.0), None);
        assert_eq!(checked_trunc_i32(-2_147_483_648.5), Some(i32::MIN));
        assert_eq!(checked_trunc_i32(-2_147_483_649.0), None);
        assert_eq!(checked_trunc_i32_in(5.0, 0, 4), None);
        assert_eq!(checked_trunc_i32_in(4.2, 0, 4), Some(4));
    }

    #[test]
    fn trunc_count_saturates_large_and_rejects_negative() {
        assert_eq!(checked_trunc_count(-1.0), None);
        assert_eq!(checked_trunc_count(-0.5), Some(0));
        assert_eq!(checked_trunc_count(f64::NAN), None);
        assert_eq!(checked_trunc_count(1e30), Some(usize::MAX));
        assert_eq!(checked_trunc_count(3.7), Some(3));
    }
}
