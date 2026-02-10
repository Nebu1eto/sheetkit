//! Engineering formula functions: base conversion (BIN2DEC, BIN2HEX, etc.),
//! complex numbers (COMPLEX, IMREAL, IMAGINARY, etc.), DELTA, GESTEP,
//! ERF, ERFC, CONVERT, and Bessel functions.

use crate::cell::CellValue;
use crate::error::Result;
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_number, coerce_to_string, Evaluator};
use crate::formula::functions::check_arg_count;

/// BIN2DEC(number)
pub fn fn_bin2dec(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("BIN2DEC", args, 1, 1)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s = s.trim();
    if s.len() > 10 || s.is_empty() || s.chars().any(|c| c != '0' && c != '1') {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let val = i64::from_str_radix(s, 2).unwrap_or(0);
    let result = if s.len() == 10 && s.starts_with('1') {
        val - 1024
    } else {
        val
    };
    Ok(CellValue::Number(result as f64))
}

/// BIN2HEX(number, [places])
pub fn fn_bin2hex(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("BIN2HEX", args, 1, 2)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s = s.trim();
    if s.len() > 10 || s.is_empty() || s.chars().any(|c| c != '0' && c != '1') {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let val = i64::from_str_radix(s, 2).unwrap_or(0);
    let result = if s.len() == 10 && s.starts_with('1') {
        val - 1024
    } else {
        val
    };
    format_hex(result, args, ctx, 1)
}

/// BIN2OCT(number, [places])
pub fn fn_bin2oct(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("BIN2OCT", args, 1, 2)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s = s.trim();
    if s.len() > 10 || s.is_empty() || s.chars().any(|c| c != '0' && c != '1') {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let val = i64::from_str_radix(s, 2).unwrap_or(0);
    let result = if s.len() == 10 && s.starts_with('1') {
        val - 1024
    } else {
        val
    };
    format_oct(result, args, ctx, 1)
}

/// DEC2BIN(number, [places])
pub fn fn_dec2bin(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DEC2BIN", args, 1, 2)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)? as i64;
    if !(-512..=511).contains(&n) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    format_bin(n, args, ctx, 1)
}

/// DEC2HEX(number, [places])
pub fn fn_dec2hex(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DEC2HEX", args, 1, 2)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)? as i64;
    if !(-549_755_813_888..=549_755_813_887).contains(&n) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    format_hex(n, args, ctx, 1)
}

/// DEC2OCT(number, [places])
pub fn fn_dec2oct(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DEC2OCT", args, 1, 2)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)? as i64;
    if !(-536_870_912..=536_870_911).contains(&n) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    format_oct(n, args, ctx, 1)
}

/// HEX2BIN(number, [places])
pub fn fn_hex2bin(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("HEX2BIN", args, 1, 2)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s = s.trim();
    if s.len() > 10 || s.is_empty() || !s.chars().all(|c| c.is_ascii_hexdigit()) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let val = i64::from_str_radix(s, 16).unwrap_or(0);
    let result = if val > 0x7FFFFFFFFF {
        val - 0x10000000000_i64
    } else {
        val
    };
    if !(-512..=511).contains(&result) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    format_bin(result, args, ctx, 1)
}

/// HEX2DEC(number)
pub fn fn_hex2dec(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("HEX2DEC", args, 1, 1)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s = s.trim();
    if s.len() > 10 || s.is_empty() || !s.chars().all(|c| c.is_ascii_hexdigit()) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let val = i64::from_str_radix(s, 16).unwrap_or(0);
    let result = if val > 0x7FFFFFFFFF {
        val - 0x10000000000_i64
    } else {
        val
    };
    Ok(CellValue::Number(result as f64))
}

/// HEX2OCT(number, [places])
pub fn fn_hex2oct(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("HEX2OCT", args, 1, 2)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s = s.trim();
    if s.len() > 10 || s.is_empty() || !s.chars().all(|c| c.is_ascii_hexdigit()) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let val = i64::from_str_radix(s, 16).unwrap_or(0);
    let result = if val > 0x7FFFFFFFFF {
        val - 0x10000000000_i64
    } else {
        val
    };
    if !(-536_870_912..=536_870_911).contains(&result) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    format_oct(result, args, ctx, 1)
}

/// OCT2BIN(number, [places])
pub fn fn_oct2bin(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("OCT2BIN", args, 1, 2)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s = s.trim();
    if s.len() > 10 || s.is_empty() || s.chars().any(|c| !('0'..='7').contains(&c)) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let val = i64::from_str_radix(s, 8).unwrap_or(0);
    // 30-bit two's complement: values above 2^29-1 are negative.
    let result = if val > 0x1FFFFFFF {
        val - 0x40000000
    } else {
        val
    };
    if !(-512..=511).contains(&result) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    format_bin(result, args, ctx, 1)
}

/// OCT2DEC(number)
pub fn fn_oct2dec(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("OCT2DEC", args, 1, 1)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s = s.trim();
    if s.len() > 10 || s.is_empty() || s.chars().any(|c| !('0'..='7').contains(&c)) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let val = i64::from_str_radix(s, 8).unwrap_or(0);
    // 30-bit two's complement: values above 2^29-1 are negative.
    let result = if val > 0x1FFFFFFF {
        val - 0x40000000
    } else {
        val
    };
    Ok(CellValue::Number(result as f64))
}

/// OCT2HEX(number, [places])
pub fn fn_oct2hex(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("OCT2HEX", args, 1, 2)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s = s.trim();
    if s.len() > 10 || s.is_empty() || s.chars().any(|c| !('0'..='7').contains(&c)) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let val = i64::from_str_radix(s, 8).unwrap_or(0);
    // 30-bit two's complement: values above 2^29-1 are negative.
    let result = if val > 0x1FFFFFFF {
        val - 0x40000000
    } else {
        val
    };
    format_hex(result, args, ctx, 1)
}

/// DELTA(number1, [number2])
pub fn fn_delta(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DELTA", args, 1, 2)?;
    let n1 = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let n2 = if args.len() > 1 {
        coerce_to_number(&ctx.eval_expr(&args[1])?)?
    } else {
        0.0
    };
    Ok(CellValue::Number(if (n1 - n2).abs() < f64::EPSILON {
        1.0
    } else {
        0.0
    }))
}

/// GESTEP(number, [step])
pub fn fn_gestep(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("GESTEP", args, 1, 2)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let step = if args.len() > 1 {
        coerce_to_number(&ctx.eval_expr(&args[1])?)?
    } else {
        0.0
    };
    Ok(CellValue::Number(if n >= step { 1.0 } else { 0.0 }))
}

/// ERF(lower_limit, [upper_limit])
pub fn fn_erf(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ERF", args, 1, 2)?;
    let lower = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    if args.len() > 1 {
        let upper = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
        Ok(CellValue::Number(erf_approx(upper) - erf_approx(lower)))
    } else {
        Ok(CellValue::Number(erf_approx(lower)))
    }
}

/// ERFC(x)
pub fn fn_erfc(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("ERFC", args, 1, 1)?;
    let x = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    Ok(CellValue::Number(1.0 - erf_approx(x)))
}

/// COMPLEX(real_num, i_num, [suffix])
pub fn fn_complex(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("COMPLEX", args, 2, 3)?;
    let real = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let imag = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let suffix = if args.len() > 2 {
        let s = coerce_to_string(&ctx.eval_expr(&args[2])?);
        if s != "i" && s != "j" {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
        s
    } else {
        "i".to_string()
    };
    Ok(CellValue::String(format_complex(real, imag, &suffix)))
}

/// IMREAL(inumber)
pub fn fn_imreal(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMREAL", args, 1, 1)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    match parse_complex(&s) {
        Some((real, _)) => Ok(CellValue::Number(real)),
        None => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// IMAGINARY(inumber)
pub fn fn_imaginary(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMAGINARY", args, 1, 1)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    match parse_complex(&s) {
        Some((_, imag)) => Ok(CellValue::Number(imag)),
        None => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// IMABS(inumber)
pub fn fn_imabs(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMABS", args, 1, 1)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    match parse_complex(&s) {
        Some((r, i)) => Ok(CellValue::Number((r * r + i * i).sqrt())),
        None => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// IMARGUMENT(inumber)
pub fn fn_imargument(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMARGUMENT", args, 1, 1)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    match parse_complex(&s) {
        Some((r, i)) => {
            if r == 0.0 && i == 0.0 {
                return Ok(CellValue::Error("#DIV/0!".to_string()));
            }
            Ok(CellValue::Number(i.atan2(r)))
        }
        None => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// IMCONJUGATE(inumber)
pub fn fn_imconjugate(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMCONJUGATE", args, 1, 1)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    match parse_complex(&s) {
        Some((r, i)) => Ok(CellValue::String(format_complex(r, -i, "i"))),
        None => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// IMSUM(inumber1, [inumber2], ...)
pub fn fn_imsum(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMSUM", args, 1, 255)?;
    let mut real_sum = 0.0;
    let mut imag_sum = 0.0;
    let values = ctx.flatten_args_to_values(args)?;
    for v in &values {
        let s = coerce_to_string(v);
        match parse_complex(&s) {
            Some((r, i)) => {
                real_sum += r;
                imag_sum += i;
            }
            None => return Ok(CellValue::Error("#NUM!".to_string())),
        }
    }
    Ok(CellValue::String(format_complex(real_sum, imag_sum, "i")))
}

/// IMSUB(inumber1, inumber2)
pub fn fn_imsub(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMSUB", args, 2, 2)?;
    let s1 = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s2 = coerce_to_string(&ctx.eval_expr(&args[1])?);
    match (parse_complex(&s1), parse_complex(&s2)) {
        (Some((r1, i1)), Some((r2, i2))) => {
            Ok(CellValue::String(format_complex(r1 - r2, i1 - i2, "i")))
        }
        _ => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// IMPRODUCT(inumber1, [inumber2], ...)
pub fn fn_improduct(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMPRODUCT", args, 1, 255)?;
    let values = ctx.flatten_args_to_values(args)?;
    let mut real = 1.0;
    let mut imag = 0.0;
    for v in &values {
        let s = coerce_to_string(v);
        match parse_complex(&s) {
            Some((r, i)) => {
                let new_real = real * r - imag * i;
                let new_imag = real * i + imag * r;
                real = new_real;
                imag = new_imag;
            }
            None => return Ok(CellValue::Error("#NUM!".to_string())),
        }
    }
    Ok(CellValue::String(format_complex(real, imag, "i")))
}

/// IMDIV(inumber1, inumber2)
pub fn fn_imdiv(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMDIV", args, 2, 2)?;
    let s1 = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let s2 = coerce_to_string(&ctx.eval_expr(&args[1])?);
    match (parse_complex(&s1), parse_complex(&s2)) {
        (Some((r1, i1)), Some((r2, i2))) => {
            let denom = r2 * r2 + i2 * i2;
            if denom == 0.0 {
                return Ok(CellValue::Error("#NUM!".to_string()));
            }
            let real = (r1 * r2 + i1 * i2) / denom;
            let imag = (i1 * r2 - r1 * i2) / denom;
            Ok(CellValue::String(format_complex(real, imag, "i")))
        }
        _ => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// IMPOWER(inumber, number)
pub fn fn_impower(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMPOWER", args, 2, 2)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    let n = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    match parse_complex(&s) {
        Some((r, i)) => {
            let modulus = (r * r + i * i).sqrt();
            if modulus == 0.0 {
                if n > 0.0 {
                    return Ok(CellValue::String("0".to_string()));
                }
                return Ok(CellValue::Error("#NUM!".to_string()));
            }
            let arg = i.atan2(r);
            let new_mod = modulus.powf(n);
            let new_arg = arg * n;
            let real = new_mod * new_arg.cos();
            let imag = new_mod * new_arg.sin();
            Ok(CellValue::String(format_complex(real, imag, "i")))
        }
        None => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// IMSQRT(inumber)
pub fn fn_imsqrt(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IMSQRT", args, 1, 1)?;
    let s = coerce_to_string(&ctx.eval_expr(&args[0])?);
    match parse_complex(&s) {
        Some((r, i)) => {
            let modulus = (r * r + i * i).sqrt();
            let arg = i.atan2(r);
            let new_mod = modulus.sqrt();
            let new_arg = arg / 2.0;
            let real = new_mod * new_arg.cos();
            let imag = new_mod * new_arg.sin();
            Ok(CellValue::String(format_complex(real, imag, "i")))
        }
        None => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// CONVERT(number, from_unit, to_unit)
pub fn fn_convert(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("CONVERT", args, 3, 3)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let from = coerce_to_string(&ctx.eval_expr(&args[1])?);
    let to = coerce_to_string(&ctx.eval_expr(&args[2])?);
    match convert_units(n, &from, &to) {
        Some(result) => Ok(CellValue::Number(result)),
        None => Ok(CellValue::Error("#N/A".to_string())),
    }
}

/// BESSELI(x, n)
pub fn fn_besseli(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("BESSELI", args, 2, 2)?;
    let x = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    if n < 0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    Ok(CellValue::Number(bessel_i(x, n)))
}

/// BESSELJ(x, n)
pub fn fn_besselj(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("BESSELJ", args, 2, 2)?;
    let x = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    if n < 0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    Ok(CellValue::Number(bessel_j(x, n)))
}

/// BESSELK(x, n)
pub fn fn_besselk(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("BESSELK", args, 2, 2)?;
    let x = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    if x <= 0.0 || n < 0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    Ok(CellValue::Number(bessel_k(x, n)))
}

/// BESSELY(x, n)
pub fn fn_bessely(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("BESSELY", args, 2, 2)?;
    let x = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let n = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    if x <= 0.0 || n < 0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    Ok(CellValue::Number(bessel_y(x, n)))
}

fn format_bin(n: i64, args: &[Expr], ctx: &mut Evaluator, places_idx: usize) -> Result<CellValue> {
    let s = if n >= 0 {
        format!("{n:b}")
    } else {
        let bits = (n as u64) & 0x3FF;
        format!("{bits:010b}")
    };
    if args.len() > places_idx {
        let places = coerce_to_number(&ctx.eval_expr(&args[places_idx])?)? as usize;
        if places < 1 || (n >= 0 && s.len() > places) {
            return Ok(CellValue::Error("#NUM!".to_string()));
        }
        if n >= 0 {
            return Ok(CellValue::String(format!("{:0>width$}", s, width = places)));
        }
    }
    Ok(CellValue::String(s))
}

fn format_hex(n: i64, args: &[Expr], ctx: &mut Evaluator, places_idx: usize) -> Result<CellValue> {
    let s = if n >= 0 {
        format!("{n:X}")
    } else {
        let bits = (n as u64) & 0xFF_FFFF_FFFF;
        format!("{bits:010X}")
    };
    if args.len() > places_idx {
        let places = coerce_to_number(&ctx.eval_expr(&args[places_idx])?)? as usize;
        if places < 1 || (n >= 0 && s.len() > places) {
            return Ok(CellValue::Error("#NUM!".to_string()));
        }
        if n >= 0 {
            return Ok(CellValue::String(format!("{:0>width$}", s, width = places)));
        }
    }
    Ok(CellValue::String(s))
}

fn format_oct(n: i64, args: &[Expr], ctx: &mut Evaluator, places_idx: usize) -> Result<CellValue> {
    let s = if n >= 0 {
        format!("{n:o}")
    } else {
        // 10 octal digits = 30 bits
        let bits = (n as u64) & 0x3FFFFFFF;
        format!("{bits:010o}")
    };
    if args.len() > places_idx {
        let places = coerce_to_number(&ctx.eval_expr(&args[places_idx])?)? as usize;
        if places < 1 || (n >= 0 && s.len() > places) {
            return Ok(CellValue::Error("#NUM!".to_string()));
        }
        if n >= 0 {
            return Ok(CellValue::String(format!("{:0>width$}", s, width = places)));
        }
    }
    Ok(CellValue::String(s))
}

fn erf_approx(x: f64) -> f64 {
    let sign = if x < 0.0 { -1.0 } else { 1.0 };
    let x = x.abs();
    let t = 1.0 / (1.0 + 0.3275911 * x);
    let poly = t
        * (0.254829592
            + t * (-0.284496736 + t * (1.421413741 + t * (-1.453152027 + t * 1.061405429))));
    sign * (1.0 - poly * (-x * x).exp())
}

fn parse_complex(s: &str) -> Option<(f64, f64)> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }
    if let Ok(n) = s.parse::<f64>() {
        return Some((n, 0.0));
    }
    if s == "i" || s == "j" {
        return Some((0.0, 1.0));
    }
    if s == "-i" || s == "-j" {
        return Some((0.0, -1.0));
    }
    if s == "+i" || s == "+j" {
        return Some((0.0, 1.0));
    }
    let suffix = if s.ends_with('i') || s.ends_with('j') {
        s.len() - 1
    } else {
        return None;
    };
    let body = &s[..suffix];
    let sign_pos = body.rfind('+').or_else(|| {
        let rp = body.rfind('-')?;
        if rp == 0 {
            None
        } else {
            Some(rp)
        }
    });
    match sign_pos {
        Some(pos) => {
            let real_str = &body[..pos];
            let imag_str = &body[pos..];
            let real = if real_str.is_empty() {
                0.0
            } else {
                real_str.parse::<f64>().ok()?
            };
            let imag = if imag_str == "+" || imag_str.is_empty() {
                1.0
            } else if imag_str == "-" {
                -1.0
            } else {
                imag_str.parse::<f64>().ok()?
            };
            Some((real, imag))
        }
        None => {
            let imag = if body == "+" || body.is_empty() {
                1.0
            } else if body == "-" {
                -1.0
            } else {
                body.parse::<f64>().ok()?
            };
            Some((0.0, imag))
        }
    }
}

fn format_complex(real: f64, imag: f64, suffix: &str) -> String {
    let real = clean_float(real);
    let imag = clean_float(imag);
    if imag == 0.0 {
        return format_number(real);
    }
    if real == 0.0 {
        if imag == 1.0 {
            return suffix.to_string();
        }
        if imag == -1.0 {
            return format!("-{suffix}");
        }
        return format!("{}{suffix}", format_number(imag));
    }
    let imag_str = if imag == 1.0 {
        format!("+{suffix}")
    } else if imag == -1.0 {
        format!("-{suffix}")
    } else if imag > 0.0 {
        format!("+{}{suffix}", format_number(imag))
    } else {
        format!("{}{suffix}", format_number(imag))
    };
    format!("{}{imag_str}", format_number(real))
}

fn clean_float(v: f64) -> f64 {
    if v.abs() < 1e-15 {
        0.0
    } else {
        v
    }
}

fn format_number(n: f64) -> String {
    if n.fract() == 0.0 && n.is_finite() && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{n}")
    }
}

fn convert_units(value: f64, from: &str, to: &str) -> Option<f64> {
    let from_base = unit_to_base_factor(from)?;
    let to_base = unit_to_base_factor(to)?;
    if from_base.1 != to_base.1 {
        if from_base.1 == "C" && to_base.1 == "C" {
            return None;
        }
        return None;
    }
    if from_base.1 == "C" {
        let celsius = temp_to_celsius(value, from)?;
        return celsius_to_temp(celsius, to);
    }
    Some(value * from_base.0 / to_base.0)
}

fn temp_to_celsius(value: f64, unit: &str) -> Option<f64> {
    match unit {
        "C" | "cel" => Some(value),
        "F" | "fah" => Some((value - 32.0) * 5.0 / 9.0),
        "K" | "kel" => Some(value - 273.15),
        "Rank" => Some((value - 491.67) * 5.0 / 9.0),
        "Reau" => Some(value * 5.0 / 4.0),
        _ => None,
    }
}

fn celsius_to_temp(celsius: f64, unit: &str) -> Option<f64> {
    match unit {
        "C" | "cel" => Some(celsius),
        "F" | "fah" => Some(celsius * 9.0 / 5.0 + 32.0),
        "K" | "kel" => Some(celsius + 273.15),
        "Rank" => Some((celsius + 273.15) * 9.0 / 5.0),
        "Reau" => Some(celsius * 4.0 / 5.0),
        _ => None,
    }
}

fn unit_to_base_factor(unit: &str) -> Option<(f64, &'static str)> {
    match unit {
        // Weight/Mass (base: kg)
        "g" => Some((0.001, "mass")),
        "kg" => Some((1.0, "mass")),
        "mg" => Some((1e-6, "mass")),
        "lbm" => Some((0.45359237, "mass")),
        "ozm" => Some((0.028349523125, "mass")),
        "stone" => Some((6.35029318, "mass")),
        "ton" => Some((907.18474, "mass")),
        "sg" => Some((14.593903, "mass")),
        "u" => Some((1.66053906660e-27, "mass")),
        "grain" => Some((6.479891e-5, "mass")),
        "cwt" | "shweight" => Some((45.359237, "mass")),
        "uk_cwt" | "lcwt" | "hweight" => Some((50.80234544, "mass")),
        "LTON" | "brton" => Some((1016.0469088, "mass")),

        // Distance (base: m)
        "m" => Some((1.0, "dist")),
        "km" => Some((1000.0, "dist")),
        "cm" => Some((0.01, "dist")),
        "mm" => Some((0.001, "dist")),
        "mi" => Some((1609.344, "dist")),
        "Nmi" => Some((1852.0, "dist")),
        "in" => Some((0.0254, "dist")),
        "ft" => Some((0.3048, "dist")),
        "yd" => Some((0.9144, "dist")),
        "ang" => Some((1e-10, "dist")),
        "ell" => Some((1.143, "dist")),
        "ly" => Some((9.46073047258e15, "dist")),
        "parsec" | "pc" => Some((3.08567758149e16, "dist")),
        "Pica" | "Picapt" => Some((0.00035277778, "dist")),
        "pica" => Some((0.00423333333, "dist")),
        "survey_mi" => Some((1609.3472, "dist")),

        // Time (base: s)
        "sec" | "s" => Some((1.0, "time")),
        "min" | "mn" => Some((60.0, "time")),
        "hr" => Some((3600.0, "time")),
        "day" | "d" => Some((86400.0, "time")),
        "yr" => Some((365.25 * 86400.0, "time")),

        // Speed (base: m/s)
        "m/s" | "m/sec" => Some((1.0, "speed")),
        "m/h" | "m/hr" => Some((1.0 / 3600.0, "speed")),
        "mph" => Some((0.44704, "speed")),
        "kn" | "admkn" => Some((0.514444444, "speed")),

        // Area (base: m^2)
        "ar" => Some((100.0, "area")),
        "ha" => Some((10000.0, "area")),
        "uk_acre" => Some((4046.8564224, "area")),
        "us_acre" => Some((4046.8726, "area")),

        // Volume (base: l)
        "l" | "L" | "lt" => Some((1.0, "vol")),
        "ml" => Some((0.001, "vol")),
        "gal" => Some((3.78541178, "vol")),
        "qt" => Some((0.946352946, "vol")),
        "pt" | "us_pt" => Some((0.473176473, "vol")),
        "cup" => Some((0.236588236, "vol")),
        "oz" | "fl_oz" | "us_oz" => Some((0.0295735296, "vol")),
        "tbs" => Some((0.0147867648, "vol")),
        "tsp" => Some((0.00492892159, "vol")),
        "uk_gal" => Some((4.54609, "vol")),
        "uk_qt" => Some((1.1365225, "vol")),
        "uk_pt" => Some((0.56826125, "vol")),

        // Energy (base: J)
        "J" | "j" => Some((1.0, "energy")),
        "e" => Some((1e-7, "energy")),
        "cal" => Some((4.1868, "energy")),
        "eV" | "ev" => Some((1.602176634e-19, "energy")),
        "HPh" | "hh" => Some((2684519.5, "energy")),
        "Wh" | "wh" => Some((3600.0, "energy")),
        "flb" => Some((1.3558179483, "energy")),
        "BTU" | "btu" => Some((1055.05585262, "energy")),

        // Power (base: W)
        "W" | "w" => Some((1.0, "power")),
        "kW" | "kw" => Some((1000.0, "power")),
        "HP" | "h" => Some((745.69987158, "power")),
        "PS" => Some((735.49875, "power")),

        // Force (base: N)
        "N" => Some((1.0, "force")),
        "dyn" | "dy" => Some((1e-5, "force")),
        "lbf" => Some((4.4482216152605, "force")),
        "pond" => Some((9.80665e-3, "force")),

        // Pressure (base: Pa)
        "Pa" | "p" => Some((1.0, "press")),
        "atm" | "at" => Some((101325.0, "press")),
        "mmHg" => Some((133.322, "press")),
        "psi" => Some((6894.757, "press")),
        "Torr" => Some((133.3224, "press")),

        // Temperature (special handling)
        "C" | "cel" | "F" | "fah" | "K" | "kel" | "Rank" | "Reau" => Some((1.0, "C")),

        // Information (base: bit)
        "bit" => Some((1.0, "info")),
        "byte" => Some((8.0, "info")),

        _ => None,
    }
}

fn bessel_j(x: f64, n: i32) -> f64 {
    let mut sum = 0.0;
    for m in 0_i32..50 {
        let sign = if m % 2 == 0 { 1.0 } else { -1.0 };
        let numer = (x / 2.0).powi(2 * m + n);
        let denom = factorial(m as u64) * factorial((m + n) as u64);
        if denom == 0.0 || !numer.is_finite() {
            break;
        }
        sum += sign * numer / denom;
    }
    sum
}

fn bessel_i(x: f64, n: i32) -> f64 {
    let mut sum = 0.0;
    for m in 0_i32..50 {
        let numer = (x / 2.0).powi(2 * m + n);
        let denom = factorial(m as u64) * factorial((m + n) as u64);
        if denom == 0.0 || !numer.is_finite() {
            break;
        }
        sum += numer / denom;
    }
    sum
}

fn bessel_y(x: f64, n: i32) -> f64 {
    let pi = std::f64::consts::PI;
    let cos_factor = (n as f64 * pi).cos();
    (cos_factor * bessel_j(x, n) - bessel_j(x, -n)) / (n as f64 * pi).sin()
}

fn bessel_k(x: f64, n: i32) -> f64 {
    let pi = std::f64::consts::PI;
    let i_neg = bessel_i(x, -n);
    let i_pos = bessel_i(x, n);
    pi / 2.0 * (i_neg - i_pos) / (n as f64 * pi).sin()
}

fn factorial(n: u64) -> f64 {
    let mut result = 1.0;
    for i in 2..=n {
        result *= i as f64;
    }
    result
}

#[cfg(test)]
mod tests {
    use crate::cell::CellValue;
    use crate::formula::eval::{evaluate, CellSnapshot};
    use crate::formula::parser::parse_formula;

    fn eval(formula: &str) -> CellValue {
        let snap = CellSnapshot::new("Sheet1".to_string());
        let expr = parse_formula(formula).unwrap();
        evaluate(&expr, &snap).unwrap()
    }

    fn assert_approx(result: CellValue, expected: f64, tol: f64) {
        match result {
            CellValue::Number(n) => {
                assert!((n - expected).abs() < tol, "expected ~{expected}, got {n}");
            }
            other => panic!("expected number ~{expected}, got {other:?}"),
        }
    }

    #[test]
    fn bin2dec_positive() {
        assert_approx(eval("BIN2DEC(\"1100100\")"), 100.0, 0.01);
    }

    #[test]
    fn bin2dec_negative() {
        assert_approx(eval("BIN2DEC(\"1111111111\")"), -1.0, 0.01);
    }

    #[test]
    fn bin2hex_basic() {
        assert_eq!(
            eval("BIN2HEX(\"11111011\",4)"),
            CellValue::String("00FB".to_string())
        );
    }

    #[test]
    fn bin2oct_basic() {
        assert_eq!(
            eval("BIN2OCT(\"1001\",4)"),
            CellValue::String("0011".to_string())
        );
    }

    #[test]
    fn dec2bin_basic() {
        assert_eq!(eval("DEC2BIN(9)"), CellValue::String("1001".to_string()));
    }

    #[test]
    fn dec2bin_negative() {
        assert_eq!(
            eval("DEC2BIN(-100)"),
            CellValue::String("1110011100".to_string())
        );
    }

    #[test]
    fn dec2bin_with_places() {
        assert_eq!(
            eval("DEC2BIN(9,8)"),
            CellValue::String("00001001".to_string())
        );
    }

    #[test]
    fn dec2hex_basic() {
        assert_eq!(eval("DEC2HEX(100)"), CellValue::String("64".to_string()));
    }

    #[test]
    fn dec2hex_negative() {
        let result = eval("DEC2HEX(-54)");
        assert_eq!(result, CellValue::String("FFFFFFFFCA".to_string()));
    }

    #[test]
    fn dec2oct_basic() {
        assert_eq!(eval("DEC2OCT(58)"), CellValue::String("72".to_string()));
    }

    #[test]
    fn dec2oct_negative() {
        // DEC2OCT(-1) should produce 7777777777 (10-digit, 30-bit two's complement)
        assert_eq!(
            eval("DEC2OCT(-1)"),
            CellValue::String("7777777777".to_string())
        );
        // DEC2OCT(-536870912) is the minimum value
        assert_eq!(
            eval("DEC2OCT(-536870912)"),
            CellValue::String("4000000000".to_string())
        );
    }

    #[test]
    fn hex2bin_basic() {
        assert_eq!(
            eval("HEX2BIN(\"F\",8)"),
            CellValue::String("00001111".to_string())
        );
    }

    #[test]
    fn hex2dec_basic() {
        assert_approx(eval("HEX2DEC(\"A5\")"), 165.0, 0.01);
    }

    #[test]
    fn hex2dec_negative() {
        assert_approx(eval("HEX2DEC(\"FFFFFFFFFF\")"), -1.0, 0.01);
    }

    #[test]
    fn hex2oct_basic() {
        assert_eq!(
            eval("HEX2OCT(\"F\",3)"),
            CellValue::String("017".to_string())
        );
    }

    #[test]
    fn oct2bin_basic() {
        assert_eq!(
            eval("OCT2BIN(\"3\",4)"),
            CellValue::String("0011".to_string())
        );
    }

    #[test]
    fn oct2bin_negative() {
        // 7777777000 octal = -512 decimal -> 10-bit binary 1000000000
        assert_eq!(
            eval("OCT2BIN(\"7777777000\")"),
            CellValue::String("1000000000".to_string())
        );
        // 7777777776 octal = -2 decimal -> 1111111110
        assert_eq!(
            eval("OCT2BIN(\"7777777776\")"),
            CellValue::String("1111111110".to_string())
        );
        // 7777777777 octal = -1 decimal -> 1111111111
        assert_eq!(
            eval("OCT2BIN(\"7777777777\")"),
            CellValue::String("1111111111".to_string())
        );
    }

    #[test]
    fn oct2dec_basic() {
        assert_approx(eval("OCT2DEC(\"54\")"), 44.0, 0.01);
    }

    #[test]
    fn oct2dec_negative() {
        // 7777777777 octal = -1 decimal (30-bit two's complement)
        assert_approx(eval("OCT2DEC(\"7777777777\")"), -1.0, 0.01);
        // 4000000000 octal = -536870912 decimal
        assert_approx(eval("OCT2DEC(\"4000000000\")"), -536_870_912.0, 0.01);
        // 7777777000 octal = -512 decimal
        assert_approx(eval("OCT2DEC(\"7777777000\")"), -512.0, 0.01);
    }

    #[test]
    fn oct2dec_positive_boundary() {
        // 3777777777 octal = 536870911 (max positive in 30-bit two's complement)
        assert_approx(eval("OCT2DEC(\"3777777777\")"), 536_870_911.0, 0.01);
    }

    #[test]
    fn oct2hex_basic() {
        assert_eq!(
            eval("OCT2HEX(\"100\",4)"),
            CellValue::String("0040".to_string())
        );
    }

    #[test]
    fn oct2hex_negative() {
        // 7777777777 octal = -1 decimal -> FFFFFFFFFF hex
        assert_eq!(
            eval("OCT2HEX(\"7777777777\")"),
            CellValue::String("FFFFFFFFFF".to_string())
        );
        // 4000000000 octal = -536870912 decimal -> FFE0000000 hex
        assert_eq!(
            eval("OCT2HEX(\"4000000000\")"),
            CellValue::String("FFE0000000".to_string())
        );
    }

    #[test]
    fn delta_equal() {
        assert_approx(eval("DELTA(5,5)"), 1.0, 0.01);
    }

    #[test]
    fn delta_not_equal() {
        assert_approx(eval("DELTA(5,4)"), 0.0, 0.01);
    }

    #[test]
    fn delta_default() {
        assert_approx(eval("DELTA(0)"), 1.0, 0.01);
    }

    #[test]
    fn gestep_above() {
        assert_approx(eval("GESTEP(5,4)"), 1.0, 0.01);
    }

    #[test]
    fn gestep_below() {
        assert_approx(eval("GESTEP(3,4)"), 0.0, 0.01);
    }

    #[test]
    fn gestep_equal() {
        assert_approx(eval("GESTEP(4,4)"), 1.0, 0.01);
    }

    #[test]
    fn erf_basic() {
        assert_approx(eval("ERF(1)"), 0.8427, 0.001);
    }

    #[test]
    fn erf_range() {
        assert_approx(eval("ERF(0,1)"), 0.8427, 0.001);
    }

    #[test]
    fn erfc_basic() {
        assert_approx(eval("ERFC(1)"), 0.1573, 0.001);
    }

    #[test]
    fn complex_basic() {
        assert_eq!(eval("COMPLEX(3,4)"), CellValue::String("3+4i".to_string()));
    }

    #[test]
    fn complex_real_only() {
        assert_eq!(eval("COMPLEX(3,0)"), CellValue::String("3".to_string()));
    }

    #[test]
    fn complex_imag_only() {
        assert_eq!(eval("COMPLEX(0,4)"), CellValue::String("4i".to_string()));
    }

    #[test]
    fn complex_negative_imag() {
        assert_eq!(eval("COMPLEX(3,-4)"), CellValue::String("3-4i".to_string()));
    }

    #[test]
    fn imreal_basic() {
        assert_approx(eval("IMREAL(\"3+4i\")"), 3.0, 0.01);
    }

    #[test]
    fn imaginary_basic() {
        assert_approx(eval("IMAGINARY(\"3+4i\")"), 4.0, 0.01);
    }

    #[test]
    fn imabs_basic() {
        assert_approx(eval("IMABS(\"3+4i\")"), 5.0, 0.01);
    }

    #[test]
    fn imargument_basic() {
        assert_approx(eval("IMARGUMENT(\"3+4i\")"), (4.0_f64).atan2(3.0), 0.001);
    }

    #[test]
    fn imconjugate_basic() {
        assert_eq!(
            eval("IMCONJUGATE(\"3+4i\")"),
            CellValue::String("3-4i".to_string())
        );
    }

    #[test]
    fn imsum_basic() {
        assert_eq!(
            eval("IMSUM(\"3+4i\",\"1-2i\")"),
            CellValue::String("4+2i".to_string())
        );
    }

    #[test]
    fn imsub_basic() {
        assert_eq!(
            eval("IMSUB(\"3+4i\",\"1+2i\")"),
            CellValue::String("2+2i".to_string())
        );
    }

    #[test]
    fn improduct_basic() {
        assert_eq!(
            eval("IMPRODUCT(\"1+2i\",\"3+4i\")"),
            CellValue::String("-5+10i".to_string())
        );
    }

    #[test]
    fn imdiv_basic() {
        let result = eval("IMDIV(\"2+4i\",\"1+1i\")");
        assert_eq!(result, CellValue::String("3+i".to_string()));
    }

    #[test]
    fn impower_basic() {
        let result = eval("IMPOWER(\"2+3i\",2)");
        if let CellValue::String(s) = &result {
            let parsed = super::parse_complex(s).unwrap();
            assert!((parsed.0 - (-5.0)).abs() < 0.01);
            assert!((parsed.1 - 12.0).abs() < 0.01);
        } else {
            panic!("expected string, got {result:?}");
        }
    }

    #[test]
    fn imsqrt_basic() {
        let result = eval("IMSQRT(\"4\")");
        if let CellValue::String(s) = &result {
            let parsed = super::parse_complex(s).unwrap();
            assert!((parsed.0 - 2.0).abs() < 0.01);
            assert!(parsed.1.abs() < 0.01);
        } else {
            panic!("expected string, got {result:?}");
        }
    }

    #[test]
    fn convert_length() {
        assert_approx(eval("CONVERT(1,\"in\",\"cm\")"), 2.54, 0.001);
    }

    #[test]
    fn convert_weight() {
        assert_approx(eval("CONVERT(1,\"lbm\",\"kg\")"), 0.453592, 0.001);
    }

    #[test]
    fn convert_temperature() {
        assert_approx(eval("CONVERT(100,\"C\",\"F\")"), 212.0, 0.1);
    }

    #[test]
    fn convert_temperature_k() {
        assert_approx(eval("CONVERT(0,\"C\",\"K\")"), 273.15, 0.01);
    }

    #[test]
    fn convert_incompatible() {
        assert_eq!(
            eval("CONVERT(1,\"in\",\"kg\")"),
            CellValue::Error("#N/A".to_string())
        );
    }

    #[test]
    fn besselj_basic() {
        assert_approx(eval("BESSELJ(1.9,2)"), 0.3295, 0.01);
    }

    #[test]
    fn besseli_basic() {
        assert_approx(eval("BESSELI(1.5,1)"), 0.9817, 0.01);
    }

    #[test]
    fn besselk_zero_x() {
        assert_eq!(eval("BESSELK(0,1)"), CellValue::Error("#NUM!".to_string()));
    }

    #[test]
    fn bessely_zero_x() {
        assert_eq!(eval("BESSELY(0,1)"), CellValue::Error("#NUM!".to_string()));
    }
}
