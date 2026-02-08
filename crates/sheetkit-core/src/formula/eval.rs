//! Formula evaluation engine.
//!
//! Evaluates parsed formula ASTs against cell data provided through the
//! [`CellDataProvider`] trait. Supports arithmetic, comparison, string
//! concatenation, cell/range references, and built-in function calls.

use std::collections::HashSet;

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::formula::ast::{BinaryOperator, CellReference, Expr, UnaryOperator};
use crate::formula::functions;
use crate::utils::cell_ref::{column_name_to_number, column_number_to_name};

/// Maximum recursion depth for nested formula evaluation.
const MAX_EVAL_DEPTH: usize = 256;

/// Provides cell data for formula evaluation.
pub trait CellDataProvider {
    /// Return the cell value at the given (1-based) column and row on `sheet`.
    fn get_cell(&self, sheet: &str, col: u32, row: u32) -> CellValue;
    /// Return the name of the sheet that owns the formula being evaluated.
    fn current_sheet(&self) -> &str;
}

/// In-memory snapshot of cell data, decoupled from any Workbook borrow.
pub struct CellSnapshot {
    cells: std::collections::HashMap<(String, u32, u32), CellValue>,
    current_sheet: String,
}

impl CellSnapshot {
    /// Create a new empty snapshot with the given current-sheet context.
    pub fn new(current_sheet: String) -> Self {
        Self {
            cells: std::collections::HashMap::new(),
            current_sheet,
        }
    }

    /// Insert a cell value into the snapshot.
    pub fn set_cell(&mut self, sheet: &str, col: u32, row: u32, value: CellValue) {
        self.cells.insert((sheet.to_string(), col, row), value);
    }
}

impl CellDataProvider for CellSnapshot {
    fn get_cell(&self, sheet: &str, col: u32, row: u32) -> CellValue {
        self.cells
            .get(&(sheet.to_string(), col, row))
            .cloned()
            .unwrap_or(CellValue::Empty)
    }

    fn current_sheet(&self) -> &str {
        &self.current_sheet
    }
}

/// Evaluate a parsed formula expression against the given cell data provider.
pub fn evaluate(expr: &Expr, provider: &dyn CellDataProvider) -> Result<CellValue> {
    let mut evaluator = Evaluator::new(provider);
    evaluator.eval_expr(expr)
}

/// Stateful evaluator that tracks recursion depth and circular references.
pub struct Evaluator<'a> {
    provider: &'a dyn CellDataProvider,
    eval_stack: HashSet<(String, u32, u32)>,
    depth: usize,
}

impl<'a> Evaluator<'a> {
    /// Create a new evaluator backed by the given data provider.
    pub fn new(provider: &'a dyn CellDataProvider) -> Self {
        Self {
            provider,
            eval_stack: HashSet::new(),
            depth: 0,
        }
    }

    /// Evaluate an AST expression node.
    pub fn eval_expr(&mut self, expr: &Expr) -> Result<CellValue> {
        self.depth += 1;
        if self.depth > MAX_EVAL_DEPTH {
            self.depth -= 1;
            return Err(Error::FormulaError(
                "maximum evaluation depth exceeded".to_string(),
            ));
        }
        let result = self.eval_inner(expr);
        self.depth -= 1;
        result
    }

    /// Evaluate a single argument expression at `index`, returning an error if
    /// the index is out of bounds.
    pub fn eval_arg(&mut self, args: &[Expr], index: usize) -> Result<CellValue> {
        if index >= args.len() {
            return Err(Error::FormulaError(format!(
                "missing argument at index {index}"
            )));
        }
        self.eval_expr(&args[index])
    }

    /// Collect all numeric values from the argument list, expanding ranges.
    /// Non-numeric values inside ranges are silently skipped (Excel semantics).
    pub fn collect_numbers(&mut self, args: &[Expr]) -> Result<Vec<f64>> {
        let mut nums = Vec::new();
        for arg in args {
            match arg {
                Expr::Range { start, end } => {
                    let values = self.expand_range(start, end)?;
                    for v in values {
                        if let Ok(n) = coerce_to_number(&v) {
                            nums.push(n);
                        }
                    }
                }
                _ => {
                    let v = self.eval_expr(arg)?;
                    nums.push(coerce_to_number(&v)?);
                }
            }
        }
        Ok(nums)
    }

    /// Flatten all arguments into a Vec of CellValues, expanding ranges.
    pub fn flatten_args_to_values(&mut self, args: &[Expr]) -> Result<Vec<CellValue>> {
        let mut values = Vec::new();
        for arg in args {
            match arg {
                Expr::Range { start, end } => {
                    values.extend(self.expand_range(start, end)?);
                }
                _ => {
                    values.push(self.eval_expr(arg)?);
                }
            }
        }
        Ok(values)
    }

    /// Expand a rectangular cell range into individual CellValues (row-major).
    pub fn expand_range(
        &mut self,
        start: &CellReference,
        end: &CellReference,
    ) -> Result<Vec<CellValue>> {
        let sheet = start
            .sheet
            .as_deref()
            .unwrap_or_else(|| self.provider.current_sheet());
        let start_col = column_name_to_number(&start.col)?;
        let end_col = column_name_to_number(&end.col)?;
        let start_row = start.row;
        let end_row = end.row;
        let min_col = start_col.min(end_col);
        let max_col = start_col.max(end_col);
        let min_row = start_row.min(end_row);
        let max_row = start_row.max(end_row);
        let mut values = Vec::new();
        for r in min_row..=max_row {
            for c in min_col..=max_col {
                values.push(self.resolve_cell(sheet, c, r)?);
            }
        }
        Ok(values)
    }

    /// Return the current sheet name from the provider.
    pub fn current_sheet(&self) -> &str {
        self.provider.current_sheet()
    }

    // -- Private helpers --

    fn eval_inner(&mut self, expr: &Expr) -> Result<CellValue> {
        match expr {
            Expr::Number(n) => Ok(CellValue::Number(*n)),
            Expr::String(s) => Ok(CellValue::String(s.clone())),
            Expr::Bool(b) => Ok(CellValue::Bool(*b)),
            Expr::Error(e) => Ok(CellValue::Error(e.clone())),
            Expr::CellRef(cell_ref) => self.eval_cell_ref(cell_ref),
            Expr::Range { start, end } => {
                // When a range appears in a scalar context, return the first cell.
                let vals = self.expand_range(start, end)?;
                Ok(vals.into_iter().next().unwrap_or(CellValue::Empty))
            }
            Expr::Paren(inner) => self.eval_expr(inner),
            Expr::BinaryOp { op, left, right } => self.eval_binary(*op, left, right),
            Expr::UnaryOp { op, operand } => self.eval_unary(*op, operand),
            Expr::Function { name, args } => self.eval_function(name, args),
        }
    }

    fn eval_cell_ref(&mut self, cell_ref: &CellReference) -> Result<CellValue> {
        let sheet = cell_ref
            .sheet
            .as_deref()
            .unwrap_or_else(|| self.provider.current_sheet())
            .to_string();
        let col = column_name_to_number(&cell_ref.col)?;
        let row = cell_ref.row;
        self.resolve_cell(&sheet, col, row)
    }

    /// Resolve a single cell, following formula values and detecting cycles.
    fn resolve_cell(&mut self, sheet: &str, col: u32, row: u32) -> Result<CellValue> {
        let key = (sheet.to_string(), col, row);
        if self.eval_stack.contains(&key) {
            let col_name = column_number_to_name(col)?;
            return Err(Error::CircularReference {
                cell: format!("{col_name}{row}"),
            });
        }
        let cell_value = self.provider.get_cell(sheet, col, row);
        match cell_value {
            CellValue::Formula { ref expr, .. } => {
                self.eval_stack.insert(key.clone());
                let parsed = crate::formula::parser::parse_formula(expr)?;
                let result = self.eval_expr(&parsed);
                self.eval_stack.remove(&key);
                result
            }
            other => Ok(other),
        }
    }

    fn eval_binary(&mut self, op: BinaryOperator, left: &Expr, right: &Expr) -> Result<CellValue> {
        let lhs = self.eval_expr(left)?;
        let rhs = self.eval_expr(right)?;

        // Propagate errors.
        if let CellValue::Error(ref e) = lhs {
            return Ok(CellValue::Error(e.clone()));
        }
        if let CellValue::Error(ref e) = rhs {
            return Ok(CellValue::Error(e.clone()));
        }

        match op {
            BinaryOperator::Concat => {
                let ls = coerce_to_string(&lhs);
                let rs = coerce_to_string(&rhs);
                Ok(CellValue::String(format!("{ls}{rs}")))
            }
            BinaryOperator::Add
            | BinaryOperator::Sub
            | BinaryOperator::Mul
            | BinaryOperator::Div
            | BinaryOperator::Pow => {
                let ln = coerce_to_number(&lhs)?;
                let rn = coerce_to_number(&rhs)?;
                let result = match op {
                    BinaryOperator::Add => ln + rn,
                    BinaryOperator::Sub => ln - rn,
                    BinaryOperator::Mul => ln * rn,
                    BinaryOperator::Div => {
                        if rn == 0.0 {
                            return Ok(CellValue::Error("#DIV/0!".to_string()));
                        }
                        ln / rn
                    }
                    BinaryOperator::Pow => ln.powf(rn),
                    _ => unreachable!(),
                };
                Ok(CellValue::Number(result))
            }
            BinaryOperator::Eq
            | BinaryOperator::Ne
            | BinaryOperator::Lt
            | BinaryOperator::Le
            | BinaryOperator::Gt
            | BinaryOperator::Ge => {
                let ord = compare_values(&lhs, &rhs);
                let result = match op {
                    BinaryOperator::Eq => ord == std::cmp::Ordering::Equal,
                    BinaryOperator::Ne => ord != std::cmp::Ordering::Equal,
                    BinaryOperator::Lt => ord == std::cmp::Ordering::Less,
                    BinaryOperator::Le => {
                        ord == std::cmp::Ordering::Less || ord == std::cmp::Ordering::Equal
                    }
                    BinaryOperator::Gt => ord == std::cmp::Ordering::Greater,
                    BinaryOperator::Ge => {
                        ord == std::cmp::Ordering::Greater || ord == std::cmp::Ordering::Equal
                    }
                    _ => unreachable!(),
                };
                Ok(CellValue::Bool(result))
            }
        }
    }

    fn eval_unary(&mut self, op: UnaryOperator, operand: &Expr) -> Result<CellValue> {
        let val = self.eval_expr(operand)?;
        if let CellValue::Error(ref e) = val {
            return Ok(CellValue::Error(e.clone()));
        }
        let n = coerce_to_number(&val)?;
        match op {
            UnaryOperator::Neg => Ok(CellValue::Number(-n)),
            UnaryOperator::Pos => Ok(CellValue::Number(n)),
            UnaryOperator::Percent => Ok(CellValue::Number(n / 100.0)),
        }
    }

    fn eval_function(&mut self, name: &str, args: &[Expr]) -> Result<CellValue> {
        let func = functions::lookup_function(name).ok_or_else(|| Error::UnknownFunction {
            name: name.to_string(),
        })?;
        func(args, self)
    }
}

// -- Type coercion helpers (pub for use by function implementations) --

/// Coerce a CellValue to f64. Booleans become 0/1, empty becomes 0,
/// numeric strings are parsed. Non-numeric strings yield an error.
pub fn coerce_to_number(value: &CellValue) -> Result<f64> {
    match value {
        CellValue::Number(n) => Ok(*n),
        CellValue::Date(n) => Ok(*n),
        CellValue::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
        CellValue::Empty => Ok(0.0),
        CellValue::String(s) => s
            .parse::<f64>()
            .map_err(|_| Error::FormulaError(format!("cannot convert \"{s}\" to number"))),
        CellValue::Error(e) => Err(Error::FormulaError(e.clone())),
        CellValue::Formula { result, .. } => {
            if let Some(ref inner) = result {
                coerce_to_number(inner)
            } else {
                Ok(0.0)
            }
        }
    }
}

/// Coerce a CellValue to a string.
pub fn coerce_to_string(value: &CellValue) -> String {
    match value {
        CellValue::String(s) => s.clone(),
        CellValue::Number(n) => {
            if n.fract() == 0.0 && n.is_finite() {
                format!("{}", *n as i64)
            } else {
                format!("{n}")
            }
        }
        CellValue::Bool(b) => {
            if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        CellValue::Empty => String::new(),
        CellValue::Error(e) => e.clone(),
        CellValue::Date(n) => format!("{n}"),
        CellValue::Formula { result, .. } => {
            if let Some(ref inner) = result {
                coerce_to_string(inner)
            } else {
                String::new()
            }
        }
    }
}

/// Coerce a CellValue to a boolean. Numbers: 0 is false, nonzero is true.
/// Strings "TRUE"/"FALSE" are recognized. Empty is false.
pub fn coerce_to_bool(value: &CellValue) -> Result<bool> {
    match value {
        CellValue::Bool(b) => Ok(*b),
        CellValue::Number(n) => Ok(*n != 0.0),
        CellValue::Date(n) => Ok(*n != 0.0),
        CellValue::Empty => Ok(false),
        CellValue::String(s) => match s.to_ascii_uppercase().as_str() {
            "TRUE" => Ok(true),
            "FALSE" => Ok(false),
            _ => Err(Error::FormulaError(format!(
                "cannot convert \"{s}\" to boolean"
            ))),
        },
        CellValue::Error(e) => Err(Error::FormulaError(e.clone())),
        CellValue::Formula { result, .. } => {
            if let Some(ref inner) = result {
                coerce_to_bool(inner)
            } else {
                Ok(false)
            }
        }
    }
}

/// Compare two CellValues for ordering. Uses Excel comparison semantics:
/// numbers compare numerically, strings compare case-insensitively,
/// mixed types rank as: empty < number < string < bool.
pub fn compare_values(lhs: &CellValue, rhs: &CellValue) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    fn type_rank(v: &CellValue) -> u8 {
        match v {
            CellValue::Empty => 0,
            CellValue::Number(_) | CellValue::Date(_) => 1,
            CellValue::String(_) => 2,
            CellValue::Bool(_) => 3,
            CellValue::Error(_) => 4,
            CellValue::Formula { .. } => 5,
        }
    }

    let lr = type_rank(lhs);
    let rr = type_rank(rhs);
    if lr != rr {
        return lr.cmp(&rr);
    }

    match (lhs, rhs) {
        (CellValue::Empty, CellValue::Empty) => Ordering::Equal,
        (CellValue::Number(a), CellValue::Number(b))
        | (CellValue::Date(a), CellValue::Date(b))
        | (CellValue::Number(a), CellValue::Date(b))
        | (CellValue::Date(a), CellValue::Number(b)) => a.partial_cmp(b).unwrap_or(Ordering::Equal),
        (CellValue::String(a), CellValue::String(b)) => {
            a.to_ascii_lowercase().cmp(&b.to_ascii_lowercase())
        }
        (CellValue::Bool(a), CellValue::Bool(b)) => a.cmp(b),
        _ => Ordering::Equal,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::ast::{BinaryOperator, CellReference, Expr, UnaryOperator};
    use crate::formula::parser::parse_formula;

    fn make_snapshot() -> CellSnapshot {
        CellSnapshot::new("Sheet1".to_string())
    }

    // -- Literal evaluation --

    #[test]
    fn eval_number_literal() {
        let snap = make_snapshot();
        let result = evaluate(&Expr::Number(42.0), &snap).unwrap();
        assert_eq!(result, CellValue::Number(42.0));
    }

    #[test]
    fn eval_string_literal() {
        let snap = make_snapshot();
        let result = evaluate(&Expr::String("hello".to_string()), &snap).unwrap();
        assert_eq!(result, CellValue::String("hello".to_string()));
    }

    #[test]
    fn eval_bool_literal() {
        let snap = make_snapshot();
        assert_eq!(
            evaluate(&Expr::Bool(true), &snap).unwrap(),
            CellValue::Bool(true)
        );
        assert_eq!(
            evaluate(&Expr::Bool(false), &snap).unwrap(),
            CellValue::Bool(false)
        );
    }

    #[test]
    fn eval_error_literal() {
        let snap = make_snapshot();
        let result = evaluate(&Expr::Error("#N/A".to_string()), &snap).unwrap();
        assert_eq!(result, CellValue::Error("#N/A".to_string()));
    }

    // -- Binary operations --

    #[test]
    fn eval_add() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Add,
            left: Box::new(Expr::Number(10.0)),
            right: Box::new(Expr::Number(5.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(15.0));
    }

    #[test]
    fn eval_sub() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Sub,
            left: Box::new(Expr::Number(10.0)),
            right: Box::new(Expr::Number(3.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(7.0));
    }

    #[test]
    fn eval_mul() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Mul,
            left: Box::new(Expr::Number(4.0)),
            right: Box::new(Expr::Number(3.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(12.0));
    }

    #[test]
    fn eval_div() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Div,
            left: Box::new(Expr::Number(10.0)),
            right: Box::new(Expr::Number(4.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(2.5));
    }

    #[test]
    fn eval_div_by_zero() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Div,
            left: Box::new(Expr::Number(1.0)),
            right: Box::new(Expr::Number(0.0)),
        };
        assert_eq!(
            evaluate(&expr, &snap).unwrap(),
            CellValue::Error("#DIV/0!".to_string())
        );
    }

    #[test]
    fn eval_pow() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Pow,
            left: Box::new(Expr::Number(2.0)),
            right: Box::new(Expr::Number(10.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(1024.0));
    }

    #[test]
    fn eval_concat() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Concat,
            left: Box::new(Expr::String("hello".to_string())),
            right: Box::new(Expr::String(" world".to_string())),
        };
        assert_eq!(
            evaluate(&expr, &snap).unwrap(),
            CellValue::String("hello world".to_string())
        );
    }

    #[test]
    fn eval_eq() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Eq,
            left: Box::new(Expr::Number(5.0)),
            right: Box::new(Expr::Number(5.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Bool(true));
    }

    #[test]
    fn eval_ne() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Ne,
            left: Box::new(Expr::Number(5.0)),
            right: Box::new(Expr::Number(3.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Bool(true));
    }

    #[test]
    fn eval_lt() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Lt,
            left: Box::new(Expr::Number(3.0)),
            right: Box::new(Expr::Number(5.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Bool(true));
    }

    #[test]
    fn eval_le() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Le,
            left: Box::new(Expr::Number(5.0)),
            right: Box::new(Expr::Number(5.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Bool(true));
    }

    #[test]
    fn eval_gt() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Gt,
            left: Box::new(Expr::Number(5.0)),
            right: Box::new(Expr::Number(3.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Bool(true));
    }

    #[test]
    fn eval_ge() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Ge,
            left: Box::new(Expr::Number(5.0)),
            right: Box::new(Expr::Number(5.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Bool(true));
    }

    // -- Unary operations --

    #[test]
    fn eval_unary_neg() {
        let snap = make_snapshot();
        let expr = Expr::UnaryOp {
            op: UnaryOperator::Neg,
            operand: Box::new(Expr::Number(7.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(-7.0));
    }

    #[test]
    fn eval_unary_pos() {
        let snap = make_snapshot();
        let expr = Expr::UnaryOp {
            op: UnaryOperator::Pos,
            operand: Box::new(Expr::Number(7.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(7.0));
    }

    #[test]
    fn eval_unary_percent() {
        let snap = make_snapshot();
        let expr = Expr::UnaryOp {
            op: UnaryOperator::Percent,
            operand: Box::new(Expr::Number(50.0)),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(0.5));
    }

    // -- Cell reference resolution --

    #[test]
    fn eval_cell_ref() {
        let mut snap = make_snapshot();
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(42.0));
        let expr = Expr::CellRef(CellReference {
            col: "A".to_string(),
            row: 1,
            abs_col: false,
            abs_row: false,
            sheet: None,
        });
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(42.0));
    }

    #[test]
    fn eval_cell_ref_empty() {
        let snap = make_snapshot();
        let expr = Expr::CellRef(CellReference {
            col: "Z".to_string(),
            row: 99,
            abs_col: false,
            abs_row: false,
            sheet: None,
        });
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Empty);
    }

    #[test]
    fn eval_cell_ref_cross_sheet() {
        let mut snap = make_snapshot();
        snap.set_cell("Sheet2", 2, 3, CellValue::String("remote".to_string()));
        let expr = Expr::CellRef(CellReference {
            col: "B".to_string(),
            row: 3,
            abs_col: false,
            abs_row: false,
            sheet: Some("Sheet2".to_string()),
        });
        assert_eq!(
            evaluate(&expr, &snap).unwrap(),
            CellValue::String("remote".to_string())
        );
    }

    // -- Range expansion --

    #[test]
    fn eval_range_expansion() {
        let mut snap = make_snapshot();
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(1.0));
        snap.set_cell("Sheet1", 1, 2, CellValue::Number(2.0));
        snap.set_cell("Sheet1", 1, 3, CellValue::Number(3.0));
        let expr = parse_formula("SUM(A1:A3)").unwrap();
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(6.0));
    }

    #[test]
    fn eval_range_2d() {
        let mut snap = make_snapshot();
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(1.0));
        snap.set_cell("Sheet1", 2, 1, CellValue::Number(2.0));
        snap.set_cell("Sheet1", 1, 2, CellValue::Number(3.0));
        snap.set_cell("Sheet1", 2, 2, CellValue::Number(4.0));
        let expr = parse_formula("SUM(A1:B2)").unwrap();
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(10.0));
    }

    // -- Type coercion --

    #[test]
    fn coerce_bool_to_number() {
        assert_eq!(coerce_to_number(&CellValue::Bool(true)).unwrap(), 1.0);
        assert_eq!(coerce_to_number(&CellValue::Bool(false)).unwrap(), 0.0);
    }

    #[test]
    fn coerce_empty_to_number() {
        assert_eq!(coerce_to_number(&CellValue::Empty).unwrap(), 0.0);
    }

    #[test]
    fn coerce_string_to_number_success() {
        assert_eq!(
            coerce_to_number(&CellValue::String("3.14".to_string())).unwrap(),
            3.14
        );
    }

    #[test]
    fn coerce_string_to_number_fail() {
        assert!(coerce_to_number(&CellValue::String("abc".to_string())).is_err());
    }

    #[test]
    fn coerce_number_to_string() {
        assert_eq!(coerce_to_string(&CellValue::Number(42.0)), "42");
        assert_eq!(coerce_to_string(&CellValue::Number(3.14)), "3.14");
    }

    #[test]
    fn coerce_bool_to_string() {
        assert_eq!(coerce_to_string(&CellValue::Bool(true)), "TRUE");
        assert_eq!(coerce_to_string(&CellValue::Bool(false)), "FALSE");
    }

    #[test]
    fn coerce_number_to_bool() {
        assert!(coerce_to_bool(&CellValue::Number(1.0)).unwrap());
        assert!(!coerce_to_bool(&CellValue::Number(0.0)).unwrap());
    }

    // -- Circular reference detection --

    #[test]
    fn eval_circular_reference() {
        let mut snap = make_snapshot();
        // A1 references itself.
        snap.set_cell(
            "Sheet1",
            1,
            1,
            CellValue::Formula {
                expr: "A1".to_string(),
                result: None,
            },
        );
        let expr = Expr::CellRef(CellReference {
            col: "A".to_string(),
            row: 1,
            abs_col: false,
            abs_row: false,
            sheet: None,
        });
        let result = evaluate(&expr, &snap);
        assert!(result.is_err());
        let err_str = result.unwrap_err().to_string();
        assert!(
            err_str.contains("circular reference"),
            "expected circular reference error, got: {err_str}"
        );
    }

    #[test]
    fn eval_indirect_circular_reference() {
        let mut snap = make_snapshot();
        // A1 = B1, B1 = A1
        snap.set_cell(
            "Sheet1",
            1,
            1,
            CellValue::Formula {
                expr: "B1".to_string(),
                result: None,
            },
        );
        snap.set_cell(
            "Sheet1",
            2,
            1,
            CellValue::Formula {
                expr: "A1".to_string(),
                result: None,
            },
        );
        let expr = Expr::CellRef(CellReference {
            col: "A".to_string(),
            row: 1,
            abs_col: false,
            abs_row: false,
            sheet: None,
        });
        let result = evaluate(&expr, &snap);
        assert!(result.is_err());
        assert!(result
            .unwrap_err()
            .to_string()
            .contains("circular reference"));
    }

    // -- Max depth --

    #[test]
    fn eval_max_depth_exceeded() {
        let snap = make_snapshot();
        // Build a deeply nested expression: ((((... (1) ...))))
        let mut expr = Expr::Number(1.0);
        for _ in 0..300 {
            expr = Expr::Paren(Box::new(expr));
        }
        let result = evaluate(&expr, &snap);
        assert!(result.is_err());
        assert!(result
            .unwrap_err()
            .to_string()
            .contains("maximum evaluation depth"));
    }

    // -- Function evaluation --

    #[test]
    fn eval_sum_function() {
        let snap = make_snapshot();
        let expr = parse_formula("SUM(1,2,3)").unwrap();
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(6.0));
    }

    #[test]
    fn eval_average_function() {
        let snap = make_snapshot();
        let expr = parse_formula("AVERAGE(10,20,30)").unwrap();
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(20.0));
    }

    #[test]
    fn eval_if_true_branch() {
        let snap = make_snapshot();
        let expr = parse_formula("IF(TRUE,10,20)").unwrap();
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(10.0));
    }

    #[test]
    fn eval_if_false_branch() {
        let snap = make_snapshot();
        let expr = parse_formula("IF(FALSE,10,20)").unwrap();
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(20.0));
    }

    #[test]
    fn eval_unknown_function() {
        let snap = make_snapshot();
        let expr = parse_formula("NONEXISTENT(1)").unwrap();
        let result = evaluate(&expr, &snap);
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("unknown function"));
    }

    #[test]
    fn eval_nested_functions() {
        let snap = make_snapshot();
        let expr = parse_formula("SUM(1,MAX(2,3))").unwrap();
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(4.0));
    }

    #[test]
    fn eval_abs_function() {
        let snap = make_snapshot();
        let expr = parse_formula("ABS(-42)").unwrap();
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(42.0));
    }

    #[test]
    fn eval_formula_cell_chain() {
        let mut snap = make_snapshot();
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(10.0));
        snap.set_cell("Sheet1", 2, 1, CellValue::Number(20.0));
        // C1 = A1 + B1
        snap.set_cell(
            "Sheet1",
            3,
            1,
            CellValue::Formula {
                expr: "A1+B1".to_string(),
                result: None,
            },
        );
        let expr = Expr::CellRef(CellReference {
            col: "C".to_string(),
            row: 1,
            abs_col: false,
            abs_row: false,
            sheet: None,
        });
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(30.0));
    }

    #[test]
    fn eval_error_propagation_in_binary() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Add,
            left: Box::new(Expr::Error("#VALUE!".to_string())),
            right: Box::new(Expr::Number(1.0)),
        };
        assert_eq!(
            evaluate(&expr, &snap).unwrap(),
            CellValue::Error("#VALUE!".to_string())
        );
    }

    #[test]
    fn eval_parsed_expression() {
        let mut snap = make_snapshot();
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(5.0));
        snap.set_cell("Sheet1", 2, 1, CellValue::Number(3.0));
        let expr = parse_formula("A1*B1+2").unwrap();
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(17.0));
    }

    #[test]
    fn eval_comparison_strings() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Lt,
            left: Box::new(Expr::String("abc".to_string())),
            right: Box::new(Expr::String("def".to_string())),
        };
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Bool(true));
    }

    #[test]
    fn eval_wrong_arg_count() {
        let snap = make_snapshot();
        let expr = parse_formula("ABS(1,2)").unwrap();
        let result = evaluate(&expr, &snap);
        assert!(result.is_err());
        assert!(result.unwrap_err().to_string().contains("expects"));
    }

    #[test]
    fn eval_concat_number_and_string() {
        let snap = make_snapshot();
        let expr = Expr::BinaryOp {
            op: BinaryOperator::Concat,
            left: Box::new(Expr::Number(42.0)),
            right: Box::new(Expr::String(" items".to_string())),
        };
        assert_eq!(
            evaluate(&expr, &snap).unwrap(),
            CellValue::String("42 items".to_string())
        );
    }
}
