//! Financial formula functions: FV, PV, NPV, IRR, PMT, IPMT, PPMT, RATE, NPER,
//! DB, DDB, SLN, SYD, EFFECT, NOMINAL, DOLLARDE, DOLLARFR, CUMPRINC, CUMIPMT,
//! XNPV, XIRR.

use crate::cell::CellValue;
use crate::error::Result;
use crate::formula::ast::Expr;
use crate::formula::eval::{coerce_to_number, Evaluator};
use crate::formula::functions::check_arg_count;

/// FV(rate, nper, pmt, [pv], [type])
pub fn fn_fv(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("FV", args, 3, 5)?;
    let rate = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let nper = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let pmt = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let pv = if args.len() > 3 {
        coerce_to_number(&ctx.eval_expr(&args[3])?)?
    } else {
        0.0
    };
    let pmt_type = if args.len() > 4 {
        coerce_to_number(&ctx.eval_expr(&args[4])?)? as i32
    } else {
        0
    };
    let fv = calc_fv(rate, nper, pmt, pv, pmt_type);
    Ok(CellValue::Number(fv))
}

/// PV(rate, nper, pmt, [fv], [type])
pub fn fn_pv(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("PV", args, 3, 5)?;
    let rate = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let nper = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let pmt = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let fv = if args.len() > 3 {
        coerce_to_number(&ctx.eval_expr(&args[3])?)?
    } else {
        0.0
    };
    let pmt_type = if args.len() > 4 {
        coerce_to_number(&ctx.eval_expr(&args[4])?)? as i32
    } else {
        0
    };
    let pv = calc_pv(rate, nper, pmt, fv, pmt_type);
    Ok(CellValue::Number(pv))
}

/// NPV(rate, value1, [value2], ...)
pub fn fn_npv(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("NPV", args, 2, 255)?;
    let rate = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let values = ctx.collect_numbers(&args[1..])?;
    let mut npv = 0.0;
    for (i, v) in values.iter().enumerate() {
        npv += v / (1.0 + rate).powi(i as i32 + 1);
    }
    Ok(CellValue::Number(npv))
}

/// IRR(values, [guess])
pub fn fn_irr(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IRR", args, 1, 2)?;
    let values = ctx.collect_numbers(&args[0..1])?;
    let guess = if args.len() > 1 {
        coerce_to_number(&ctx.eval_expr(&args[1])?)?
    } else {
        0.1
    };
    match calc_irr(&values, guess) {
        Some(irr) => Ok(CellValue::Number(irr)),
        None => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// PMT(rate, nper, pv, [fv], [type])
pub fn fn_pmt(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("PMT", args, 3, 5)?;
    let rate = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let nper = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let pv = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let fv = if args.len() > 3 {
        coerce_to_number(&ctx.eval_expr(&args[3])?)?
    } else {
        0.0
    };
    let pmt_type = if args.len() > 4 {
        coerce_to_number(&ctx.eval_expr(&args[4])?)? as i32
    } else {
        0
    };
    let pmt = calc_pmt(rate, nper, pv, fv, pmt_type);
    Ok(CellValue::Number(pmt))
}

/// IPMT(rate, per, nper, pv, [fv], [type])
pub fn fn_ipmt(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("IPMT", args, 4, 6)?;
    let rate = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let per = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let nper = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let pv = coerce_to_number(&ctx.eval_expr(&args[3])?)?;
    let fv = if args.len() > 4 {
        coerce_to_number(&ctx.eval_expr(&args[4])?)?
    } else {
        0.0
    };
    let pmt_type = if args.len() > 5 {
        coerce_to_number(&ctx.eval_expr(&args[5])?)? as i32
    } else {
        0
    };
    if per < 1.0 || per > nper {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let ipmt = calc_ipmt(rate, per, nper, pv, fv, pmt_type);
    Ok(CellValue::Number(ipmt))
}

/// PPMT(rate, per, nper, pv, [fv], [type])
pub fn fn_ppmt(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("PPMT", args, 4, 6)?;
    let rate = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let per = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let nper = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let pv = coerce_to_number(&ctx.eval_expr(&args[3])?)?;
    let fv = if args.len() > 4 {
        coerce_to_number(&ctx.eval_expr(&args[4])?)?
    } else {
        0.0
    };
    let pmt_type = if args.len() > 5 {
        coerce_to_number(&ctx.eval_expr(&args[5])?)? as i32
    } else {
        0
    };
    if per < 1.0 || per > nper {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let pmt = calc_pmt(rate, nper, pv, fv, pmt_type);
    let ipmt = calc_ipmt(rate, per, nper, pv, fv, pmt_type);
    Ok(CellValue::Number(pmt - ipmt))
}

/// RATE(nper, pmt, pv, [fv], [type], [guess])
pub fn fn_rate(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("RATE", args, 3, 6)?;
    let nper = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let pmt = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let pv = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let fv = if args.len() > 3 {
        coerce_to_number(&ctx.eval_expr(&args[3])?)?
    } else {
        0.0
    };
    let pmt_type = if args.len() > 4 {
        coerce_to_number(&ctx.eval_expr(&args[4])?)? as i32
    } else {
        0
    };
    let guess = if args.len() > 5 {
        coerce_to_number(&ctx.eval_expr(&args[5])?)?
    } else {
        0.1
    };
    match calc_rate(nper, pmt, pv, fv, pmt_type, guess) {
        Some(rate) => Ok(CellValue::Number(rate)),
        None => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

/// NPER(rate, pmt, pv, [fv], [type])
pub fn fn_nper(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("NPER", args, 3, 5)?;
    let rate = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let pmt = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let pv = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let fv = if args.len() > 3 {
        coerce_to_number(&ctx.eval_expr(&args[3])?)?
    } else {
        0.0
    };
    let pmt_type = if args.len() > 4 {
        coerce_to_number(&ctx.eval_expr(&args[4])?)? as i32
    } else {
        0
    };
    if rate == 0.0 {
        if pmt == 0.0 {
            return Ok(CellValue::Error("#NUM!".to_string()));
        }
        return Ok(CellValue::Number(-(pv + fv) / pmt));
    }
    let z = pmt * (1.0 + rate * pmt_type as f64) / rate;
    let numer = -fv + z;
    let denom = pv + z;
    if numer <= 0.0 && denom <= 0.0 {
        return Ok(CellValue::Number((numer / denom).ln() / (1.0 + rate).ln()));
    }
    if numer / denom <= 0.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    Ok(CellValue::Number((numer / denom).ln() / (1.0 + rate).ln()))
}

/// DB(cost, salvage, life, period, [month])
pub fn fn_db(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DB", args, 4, 5)?;
    let cost = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let salvage = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let life = coerce_to_number(&ctx.eval_expr(&args[2])?)? as i32;
    let period = coerce_to_number(&ctx.eval_expr(&args[3])?)? as i32;
    let month = if args.len() > 4 {
        coerce_to_number(&ctx.eval_expr(&args[4])?)? as i32
    } else {
        12
    };
    if life <= 0
        || period <= 0
        || period > life + 1
        || cost < 0.0
        || salvage < 0.0
        || !(1..=12).contains(&month)
    {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    if cost == 0.0 {
        return Ok(CellValue::Number(0.0));
    }
    let rate = (1.0 - (salvage / cost).powf(1.0 / life as f64)) * 1000.0;
    let rate = rate.round() / 1000.0;
    let mut total_depreciation = 0.0;
    let first_year_dep = cost * rate * month as f64 / 12.0;
    if period == 1 {
        return Ok(CellValue::Number(first_year_dep));
    }
    total_depreciation += first_year_dep;
    for p in 2..period {
        let dep = (cost - total_depreciation) * rate;
        total_depreciation += dep;
        if p > life {
            break;
        }
    }
    if period == life + 1 {
        let dep = (cost - total_depreciation) * rate * (12 - month) as f64 / 12.0;
        return Ok(CellValue::Number(dep));
    }
    let dep = (cost - total_depreciation) * rate;
    Ok(CellValue::Number(dep))
}

/// DDB(cost, salvage, life, period, [factor])
pub fn fn_ddb(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DDB", args, 4, 5)?;
    let cost = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let salvage = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let life = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let period = coerce_to_number(&ctx.eval_expr(&args[3])?)?;
    let factor = if args.len() > 4 {
        coerce_to_number(&ctx.eval_expr(&args[4])?)?
    } else {
        2.0
    };
    if life <= 0.0 || period <= 0.0 || period > life || cost < 0.0 || salvage < 0.0 || factor <= 0.0
    {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let rate = factor / life;
    let rate = rate.min(1.0);
    let mut total_dep = 0.0;
    for p in 1..=(period as i32) {
        let current_value = cost - total_dep;
        let mut dep = current_value * rate;
        if current_value - dep < salvage {
            dep = (current_value - salvage).max(0.0);
        }
        if p == period as i32 {
            return Ok(CellValue::Number(dep));
        }
        total_dep += dep;
    }
    Ok(CellValue::Number(0.0))
}

/// SLN(cost, salvage, life)
pub fn fn_sln(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SLN", args, 3, 3)?;
    let cost = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let salvage = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let life = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    if life == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }
    Ok(CellValue::Number((cost - salvage) / life))
}

/// SYD(cost, salvage, life, per)
pub fn fn_syd(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("SYD", args, 4, 4)?;
    let cost = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let salvage = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let life = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let per = coerce_to_number(&ctx.eval_expr(&args[3])?)?;
    if life <= 0.0 || per <= 0.0 || per > life {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let sum_of_years = life * (life + 1.0) / 2.0;
    let dep = (cost - salvage) * (life - per + 1.0) / sum_of_years;
    Ok(CellValue::Number(dep))
}

/// EFFECT(nominal_rate, npery)
pub fn fn_effect(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("EFFECT", args, 2, 2)?;
    let nominal = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let npery = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    if nominal <= 0.0 || npery < 1 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let result = (1.0 + nominal / npery as f64).powi(npery) - 1.0;
    Ok(CellValue::Number(result))
}

/// NOMINAL(effect_rate, npery)
pub fn fn_nominal(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("NOMINAL", args, 2, 2)?;
    let effect = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let npery = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    if effect <= 0.0 || npery < 1 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let result = npery as f64 * ((1.0 + effect).powf(1.0 / npery as f64) - 1.0);
    Ok(CellValue::Number(result))
}

/// DOLLARDE(fractional_dollar, fraction)
pub fn fn_dollarde(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DOLLARDE", args, 2, 2)?;
    let fractional = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let fraction = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    if fraction < 1 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let sign = if fractional < 0.0 { -1.0 } else { 1.0 };
    let abs_val = fractional.abs();
    let int_part = abs_val.floor();
    let frac_part = abs_val - int_part;
    let digits = (fraction as f64).log10().ceil().max(1.0) as u32;
    let power = 10f64.powi(digits as i32);
    let frac_decimal = frac_part * power / fraction as f64;
    Ok(CellValue::Number(sign * (int_part + frac_decimal)))
}

/// DOLLARFR(decimal_dollar, fraction)
pub fn fn_dollarfr(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("DOLLARFR", args, 2, 2)?;
    let decimal = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let fraction = coerce_to_number(&ctx.eval_expr(&args[1])?)? as i32;
    if fraction < 1 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let sign = if decimal < 0.0 { -1.0 } else { 1.0 };
    let abs_val = decimal.abs();
    let int_part = abs_val.floor();
    let frac_part = abs_val - int_part;
    let digits = (fraction as f64).log10().ceil().max(1.0) as u32;
    let power = 10f64.powi(digits as i32);
    let frac_result = frac_part * fraction as f64 / power;
    Ok(CellValue::Number(sign * (int_part + frac_result)))
}

/// CUMIPMT(rate, nper, pv, start_period, end_period, type)
pub fn fn_cumipmt(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("CUMIPMT", args, 6, 6)?;
    let rate = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let nper = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let pv = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let start = coerce_to_number(&ctx.eval_expr(&args[3])?)? as i32;
    let end = coerce_to_number(&ctx.eval_expr(&args[4])?)? as i32;
    let pmt_type = coerce_to_number(&ctx.eval_expr(&args[5])?)? as i32;
    if rate <= 0.0
        || nper <= 0.0
        || pv <= 0.0
        || start < 1
        || end < 1
        || start > end
        || (pmt_type != 0 && pmt_type != 1)
    {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let mut total = 0.0;
    for per in start..=end {
        total += calc_ipmt(rate, per as f64, nper, pv, 0.0, pmt_type);
    }
    Ok(CellValue::Number(total))
}

/// CUMPRINC(rate, nper, pv, start_period, end_period, type)
pub fn fn_cumprinc(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("CUMPRINC", args, 6, 6)?;
    let rate = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let nper = coerce_to_number(&ctx.eval_expr(&args[1])?)?;
    let pv = coerce_to_number(&ctx.eval_expr(&args[2])?)?;
    let start = coerce_to_number(&ctx.eval_expr(&args[3])?)? as i32;
    let end = coerce_to_number(&ctx.eval_expr(&args[4])?)? as i32;
    let pmt_type = coerce_to_number(&ctx.eval_expr(&args[5])?)? as i32;
    if rate <= 0.0
        || nper <= 0.0
        || pv <= 0.0
        || start < 1
        || end < 1
        || start > end
        || (pmt_type != 0 && pmt_type != 1)
    {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let pmt = calc_pmt(rate, nper, pv, 0.0, pmt_type);
    let mut total = 0.0;
    for per in start..=end {
        let ipmt = calc_ipmt(rate, per as f64, nper, pv, 0.0, pmt_type);
        total += pmt - ipmt;
    }
    Ok(CellValue::Number(total))
}

/// XNPV(rate, values, dates)
pub fn fn_xnpv(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("XNPV", args, 3, 3)?;
    let rate = coerce_to_number(&ctx.eval_expr(&args[0])?)?;
    let values = ctx.collect_numbers(&args[1..2])?;
    let dates = ctx.collect_numbers(&args[2..3])?;
    if values.len() != dates.len() || values.is_empty() {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    if rate <= -1.0 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let d0 = dates[0];
    let mut npv = 0.0;
    for (i, v) in values.iter().enumerate() {
        let days = dates[i] - d0;
        npv += v / (1.0 + rate).powf(days / 365.0);
    }
    Ok(CellValue::Number(npv))
}

/// XIRR(values, dates, [guess])
pub fn fn_xirr(args: &[Expr], ctx: &mut Evaluator) -> Result<CellValue> {
    check_arg_count("XIRR", args, 2, 3)?;
    let values = ctx.collect_numbers(&args[0..1])?;
    let dates = ctx.collect_numbers(&args[1..2])?;
    let guess = if args.len() > 2 {
        coerce_to_number(&ctx.eval_expr(&args[2])?)?
    } else {
        0.1
    };
    if values.len() != dates.len() || values.len() < 2 {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    let has_positive = values.iter().any(|v| *v > 0.0);
    let has_negative = values.iter().any(|v| *v < 0.0);
    if !has_positive || !has_negative {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }
    match calc_xirr(&values, &dates, guess) {
        Some(irr) => Ok(CellValue::Number(irr)),
        None => Ok(CellValue::Error("#NUM!".to_string())),
    }
}

fn calc_fv(rate: f64, nper: f64, pmt: f64, pv: f64, pmt_type: i32) -> f64 {
    if rate == 0.0 {
        return -(pv + pmt * nper);
    }
    let pow = (1.0 + rate).powf(nper);
    let factor = if pmt_type != 0 { 1.0 + rate } else { 1.0 };
    -(pv * pow + pmt * factor * (pow - 1.0) / rate)
}

fn calc_pv(rate: f64, nper: f64, pmt: f64, fv: f64, pmt_type: i32) -> f64 {
    if rate == 0.0 {
        return -(fv + pmt * nper);
    }
    let pow = (1.0 + rate).powf(nper);
    let factor = if pmt_type != 0 { 1.0 + rate } else { 1.0 };
    -(fv + pmt * factor * (pow - 1.0) / rate) / pow
}

fn calc_pmt(rate: f64, nper: f64, pv: f64, fv: f64, pmt_type: i32) -> f64 {
    if rate == 0.0 {
        return -(pv + fv) / nper;
    }
    let pow = (1.0 + rate).powf(nper);
    let factor = if pmt_type != 0 { 1.0 + rate } else { 1.0 };
    -(pv * pow + fv) / (factor * (pow - 1.0) / rate)
}

fn calc_ipmt(rate: f64, per: f64, nper: f64, pv: f64, fv: f64, pmt_type: i32) -> f64 {
    let pmt = calc_pmt(rate, nper, pv, fv, pmt_type);
    if pmt_type != 0 {
        let fv_prev = calc_fv(rate, per - 2.0, pmt, pv, pmt_type);
        fv_prev * rate / (1.0 + rate)
    } else {
        let fv_prev = calc_fv(rate, per - 1.0, pmt, pv, pmt_type);
        fv_prev * rate
    }
}

fn calc_irr(values: &[f64], guess: f64) -> Option<f64> {
    let has_positive = values.iter().any(|v| *v > 0.0);
    let has_negative = values.iter().any(|v| *v < 0.0);
    if !has_positive || !has_negative {
        return None;
    }
    let mut rate = guess;
    for _ in 0..100 {
        let mut npv = 0.0;
        let mut dnpv = 0.0;
        for (i, v) in values.iter().enumerate() {
            let denom = (1.0 + rate).powi(i as i32);
            if denom == 0.0 {
                return None;
            }
            npv += v / denom;
            dnpv -= (i as f64) * v / (1.0 + rate).powi(i as i32 + 1);
        }
        if dnpv.abs() < 1e-15 {
            return None;
        }
        let new_rate = rate - npv / dnpv;
        if (new_rate - rate).abs() < 1e-10 {
            return Some(new_rate);
        }
        rate = new_rate;
    }
    None
}

fn calc_rate(nper: f64, pmt: f64, pv: f64, fv: f64, pmt_type: i32, guess: f64) -> Option<f64> {
    let mut rate = guess;
    for _ in 0..100 {
        if rate <= -1.0 {
            return None;
        }
        let pow = (1.0 + rate).powf(nper);
        let factor = if pmt_type != 0 { 1.0 + rate } else { 1.0 };
        let y = pv * pow + pmt * factor * (pow - 1.0) / rate + fv;
        let dy = pv * nper * (1.0 + rate).powf(nper - 1.0)
            + pmt
                * (factor * nper * (1.0 + rate).powf(nper - 1.0) / rate
                    - factor * (pow - 1.0) / (rate * rate)
                    + if pmt_type != 0 {
                        (pow - 1.0) / rate
                    } else {
                        0.0
                    });
        if dy.abs() < 1e-15 {
            return None;
        }
        let new_rate = rate - y / dy;
        if (new_rate - rate).abs() < 1e-10 {
            return Some(new_rate);
        }
        rate = new_rate;
    }
    None
}

fn calc_xirr(values: &[f64], dates: &[f64], guess: f64) -> Option<f64> {
    let d0 = dates[0];
    let mut rate = guess;
    for _ in 0..100 {
        let mut npv = 0.0;
        let mut dnpv = 0.0;
        for (i, v) in values.iter().enumerate() {
            let days = (dates[i] - d0) / 365.0;
            let denom = (1.0 + rate).powf(days);
            if denom == 0.0 || !denom.is_finite() {
                return None;
            }
            npv += v / denom;
            dnpv -= days * v / ((1.0 + rate).powf(days + 1.0));
        }
        if dnpv.abs() < 1e-15 {
            return None;
        }
        let new_rate = rate - npv / dnpv;
        if (new_rate - rate).abs() < 1e-10 {
            return Some(new_rate);
        }
        rate = new_rate;
    }
    None
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

    fn assert_approx(result: CellValue, expected: f64, tol: f64) {
        match result {
            CellValue::Number(n) => {
                assert!((n - expected).abs() < tol, "expected ~{expected}, got {n}");
            }
            other => panic!("expected number ~{expected}, got {other:?}"),
        }
    }

    #[test]
    fn fv_basic() {
        assert_approx(eval("FV(0.06/12,10,-200,-500,1)"), 2581.4033740601, 0.01);
    }

    #[test]
    fn fv_zero_rate() {
        assert_approx(eval("FV(0,10,-200)"), 2000.0, 0.01);
    }

    #[test]
    fn pv_basic() {
        assert_approx(eval("PV(0.08/12,20*12,500,0,0)"), -59777.1458, 0.01);
    }

    #[test]
    fn npv_basic() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(-10000.0)),
            ("Sheet1", 1, 2, CellValue::Number(3000.0)),
            ("Sheet1", 1, 3, CellValue::Number(4200.0)),
            ("Sheet1", 1, 4, CellValue::Number(6800.0)),
        ];
        assert_approx(eval_with_data("NPV(0.1,A1:A4)", &data), 1188.4434, 0.01);
    }

    #[test]
    fn irr_basic() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(-70000.0)),
            ("Sheet1", 1, 2, CellValue::Number(12000.0)),
            ("Sheet1", 1, 3, CellValue::Number(15000.0)),
            ("Sheet1", 1, 4, CellValue::Number(18000.0)),
            ("Sheet1", 1, 5, CellValue::Number(21000.0)),
            ("Sheet1", 1, 6, CellValue::Number(26000.0)),
        ];
        assert_approx(eval_with_data("IRR(A1:A6)", &data), 0.08663, 0.001);
    }

    #[test]
    fn irr_no_sign_change() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(100.0)),
            ("Sheet1", 1, 2, CellValue::Number(200.0)),
        ];
        assert_eq!(
            eval_with_data("IRR(A1:A2)", &data),
            CellValue::Error("#NUM!".to_string())
        );
    }

    #[test]
    fn pmt_basic() {
        assert_approx(eval("PMT(0.08/12,10,10000)"), -1037.0321, 0.01);
    }

    #[test]
    fn pmt_zero_rate() {
        assert_approx(eval("PMT(0,10,10000)"), -1000.0, 0.01);
    }

    #[test]
    fn ipmt_basic() {
        assert_approx(eval("IPMT(0.1/12,1,36,8000)"), -66.6667, 0.01);
    }

    #[test]
    fn ipmt_invalid_period() {
        assert_eq!(
            eval("IPMT(0.1,0,36,8000)"),
            CellValue::Error("#NUM!".to_string())
        );
    }

    #[test]
    fn ppmt_basic() {
        let pmt = eval("PMT(0.1/12,36,8000)");
        let ipmt = eval("IPMT(0.1/12,1,36,8000)");
        let ppmt = eval("PPMT(0.1/12,1,36,8000)");
        if let (CellValue::Number(p), CellValue::Number(i), CellValue::Number(pp)) =
            (pmt, ipmt, ppmt)
        {
            assert!((p - i - pp).abs() < 0.01);
        } else {
            panic!("expected numbers");
        }
    }

    #[test]
    fn rate_basic() {
        assert_approx(eval("RATE(48,-200,8000)"), 0.007701, 0.0001);
    }

    #[test]
    fn nper_basic() {
        assert_approx(eval("NPER(0.01,-100,1000)"), 10.5886, 0.01);
    }

    #[test]
    fn nper_zero_rate() {
        assert_approx(eval("NPER(0,-100,1000)"), 10.0, 0.01);
    }

    #[test]
    fn db_basic() {
        assert_approx(eval("DB(1000000,100000,6,1,7)"), 186083.3333, 0.01);
    }

    #[test]
    fn ddb_basic() {
        assert_approx(eval("DDB(2400,300,10,1)"), 480.0, 0.01);
    }

    #[test]
    fn ddb_later_period() {
        assert_approx(eval("DDB(2400,300,10,2)"), 384.0, 0.01);
    }

    #[test]
    fn sln_basic() {
        assert_approx(eval("SLN(30000,7500,10)"), 2250.0, 0.01);
    }

    #[test]
    fn sln_zero_life() {
        assert_eq!(
            eval("SLN(30000,7500,0)"),
            CellValue::Error("#DIV/0!".to_string())
        );
    }

    #[test]
    fn syd_basic() {
        assert_approx(eval("SYD(30000,7500,10,1)"), 4090.9091, 0.01);
    }

    #[test]
    fn syd_last_period() {
        assert_approx(eval("SYD(30000,7500,10,10)"), 409.0909, 0.01);
    }

    #[test]
    fn effect_basic() {
        assert_approx(eval("EFFECT(0.0525,4)"), 0.053543, 0.0001);
    }

    #[test]
    fn effect_invalid() {
        assert_eq!(
            eval("EFFECT(-0.01,4)"),
            CellValue::Error("#NUM!".to_string())
        );
    }

    #[test]
    fn nominal_basic() {
        assert_approx(eval("NOMINAL(0.053543,4)"), 0.0525, 0.0001);
    }

    #[test]
    fn dollarde_basic() {
        assert_approx(eval("DOLLARDE(1.02,16)"), 1.125, 0.001);
    }

    #[test]
    fn dollarfr_basic() {
        assert_approx(eval("DOLLARFR(1.125,16)"), 1.02, 0.001);
    }

    #[test]
    fn cumipmt_basic() {
        assert_approx(eval("CUMIPMT(0.09/12,30*12,125000,1,1,0)"), -937.5, 0.01);
    }

    #[test]
    fn cumprinc_basic() {
        assert_approx(eval("CUMPRINC(0.09/12,30*12,125000,1,1,0)"), -68.2782, 0.01);
    }

    #[test]
    fn xnpv_basic() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(-10000.0)),
            ("Sheet1", 1, 2, CellValue::Number(2750.0)),
            ("Sheet1", 1, 3, CellValue::Number(4250.0)),
            ("Sheet1", 1, 4, CellValue::Number(3250.0)),
            ("Sheet1", 1, 5, CellValue::Number(2750.0)),
            ("Sheet1", 2, 1, CellValue::Number(39448.0)),
            ("Sheet1", 2, 2, CellValue::Number(39508.0)),
            ("Sheet1", 2, 3, CellValue::Number(39600.0)),
            ("Sheet1", 2, 4, CellValue::Number(39692.0)),
            ("Sheet1", 2, 5, CellValue::Number(39783.0)),
        ];
        let result = eval_with_data("XNPV(0.09,A1:A5,B1:B5)", &data);
        if let CellValue::Number(n) = result {
            assert!(n > 2000.0, "XNPV should be positive, got {n}");
        } else {
            panic!("expected number, got {result:?}");
        }
    }

    #[test]
    fn xirr_basic() {
        let data = vec![
            ("Sheet1", 1, 1, CellValue::Number(-10000.0)),
            ("Sheet1", 1, 2, CellValue::Number(2750.0)),
            ("Sheet1", 1, 3, CellValue::Number(4250.0)),
            ("Sheet1", 1, 4, CellValue::Number(3250.0)),
            ("Sheet1", 1, 5, CellValue::Number(2750.0)),
            ("Sheet1", 2, 1, CellValue::Number(39448.0)),
            ("Sheet1", 2, 2, CellValue::Number(39508.0)),
            ("Sheet1", 2, 3, CellValue::Number(39600.0)),
            ("Sheet1", 2, 4, CellValue::Number(39692.0)),
            ("Sheet1", 2, 5, CellValue::Number(39783.0)),
        ];
        let result = eval_with_data("XIRR(A1:A5,B1:B5)", &data);
        if let CellValue::Number(n) = result {
            assert!(n > 0.0, "XIRR should be positive, got {n}");
        } else {
            panic!("expected number, got {result:?}");
        }
    }
}
