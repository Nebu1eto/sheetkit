//! Formula evaluation engine.
//!
//! Evaluates parsed formula ASTs against cell data provided through the
//! [`CellDataProvider`] trait. Supports arithmetic, comparison, string
//! concatenation, cell/range references, and built-in function calls.

use std::cmp::Reverse;
use std::collections::{BinaryHeap, HashMap, HashSet};

use crate::cell::CellValue;
use crate::error::{Error, Result};
use crate::formula::ast::{BinaryOperator, CellReference, Expr, UnaryOperator};
use crate::formula::budget::FormulaBudget;
use crate::formula::functions;
use crate::utils::cell_ref::{column_name_to_number, column_number_to_name};

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

    /// Change the current-sheet context used for unqualified cell references.
    pub fn set_current_sheet(&mut self, sheet: &str) {
        self.current_sheet = sheet.to_string();
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
///
/// Unqualified references resolve against the evaluator's current sheet. That
/// is the provider's sheet for the top-level formula and the owning sheet of
/// any referenced formula cell while that cell's formula is being evaluated.
pub struct Evaluator<'a> {
    provider: &'a dyn CellDataProvider,
    eval_stack: HashSet<(String, u32, u32)>,
    depth: usize,
    position: (u32, u32),
    sheet_override: Option<String>,
    budget: FormulaBudget,
}

impl<'a> Evaluator<'a> {
    /// Create a new evaluator backed by the given data provider.
    pub fn new(provider: &'a dyn CellDataProvider) -> Self {
        Self::with_budget(provider, FormulaBudget::secure_calculation())
    }

    /// Create an evaluator that shares an existing calculation budget.
    pub fn with_budget(provider: &'a dyn CellDataProvider, budget: FormulaBudget) -> Self {
        Self {
            provider,
            eval_stack: HashSet::new(),
            depth: 0,
            position: (1, 1),
            sheet_override: None,
            budget,
        }
    }

    /// Set the 1-based `(column, row)` coordinate of the formula being evaluated.
    pub fn set_position(&mut self, col: u32, row: u32) {
        self.position = (col, row);
    }

    /// Return the 1-based `(column, row)` coordinate of the formula being evaluated.
    pub fn position(&self) -> (u32, u32) {
        self.position
    }

    /// Evaluate an AST expression node.
    pub fn eval_expr(&mut self, expr: &Expr) -> Result<CellValue> {
        self.budget.charge_evaluation_steps(1)?;
        self.depth += 1;
        if let Err(error) = self.budget.check_evaluation_depth(self.depth as u64) {
            self.depth -= 1;
            return Err(error);
        }
        let result = self.eval_inner(expr).and_then(|value| {
            let output_bytes = match &value {
                CellValue::String(value) | CellValue::Error(value) => value.len() as u64,
                CellValue::RichString(runs) => runs
                    .iter()
                    .map(|run| run.text.len() as u64)
                    .fold(0_u64, u64::saturating_add),
                _ => 0,
            };
            self.budget.charge_output_bytes(output_bytes)?;
            Ok(value)
        });
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

    /// Collect numeric values for aggregate functions such as SUM or AVERAGE.
    ///
    /// Reference arguments (ranges and single-cell references) contribute only
    /// numeric cells (see [`reference_number`]); empty, boolean, and text cells
    /// are skipped, and an error cell aborts with that error. Other arguments
    /// are direct values and use [`coerce_to_number`], so `TRUE` counts as 1
    /// and numeric text is parsed.
    pub fn collect_numbers(&mut self, args: &[Expr]) -> Result<Vec<f64>> {
        let mut nums = Vec::new();
        for arg in args {
            match arg {
                Expr::Range { start, end } => {
                    let values = self.expand_range(start, end)?;
                    for v in values {
                        if let CellValue::Error(error) = v {
                            return Err(Error::FormulaError(error));
                        }
                        if let Some(n) = reference_number(&v) {
                            nums.push(n);
                        }
                    }
                }
                Expr::CellRef(_) => {
                    let v = self.eval_expr(arg)?;
                    if let CellValue::Error(error) = v {
                        return Err(Error::FormulaError(error));
                    }
                    if let Some(n) = reference_number(&v) {
                        nums.push(n);
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
        let sheet = match start.sheet.as_deref() {
            Some(sheet) => sheet.to_string(),
            None => self.current_sheet().to_string(),
        };
        let start_col = column_name_to_number(&start.col)?;
        let end_col = column_name_to_number(&end.col)?;
        let start_row = start.row;
        let end_row = end.row;
        let min_col = start_col.min(end_col);
        let max_col = start_col.max(end_col);
        let min_row = start_row.min(end_row);
        let max_row = start_row.max(end_row);
        let width = u64::from(max_col - min_col)
            .checked_add(1)
            .ok_or_else(|| Error::Internal("formula range width overflow".to_string()))?;
        let height = u64::from(max_row - min_row)
            .checked_add(1)
            .ok_or_else(|| Error::Internal("formula range height overflow".to_string()))?;
        let area = width.saturating_mul(height);
        self.budget.charge_range_cells(area)?;
        let mut values = Vec::new();
        let capacity = usize::try_from(area)
            .map_err(|_| Error::Internal("formula range capacity overflow".to_string()))?;
        values
            .try_reserve_exact(capacity)
            .map_err(|_| Error::Internal("formula range allocation failed".to_string()))?;
        for r in min_row..=max_row {
            for c in min_col..=max_col {
                values.push(self.resolve_cell(&sheet, c, r)?);
            }
        }
        Ok(values)
    }

    /// Return the sheet that unqualified references currently resolve against.
    pub fn current_sheet(&self) -> &str {
        self.sheet_override
            .as_deref()
            .unwrap_or_else(|| self.provider.current_sheet())
    }

    /// Charge wildcard or criteria matching work before performing it.
    pub fn charge_match_work(&self, amount: u64) -> Result<()> {
        self.budget.charge_match_work(amount)
    }

    /// Charge analytically known loop or helper work before executing it.
    pub fn charge_evaluation_work(&self, amount: u64) -> Result<()> {
        self.budget.charge_evaluation_steps(amount)
    }

    /// Check analytically known string output before reserving it.
    pub fn check_output_bytes(&self, amount: u64) -> Result<()> {
        self.budget.check_output_bytes(amount)
    }

    // -- Private helpers --

    fn eval_inner(&mut self, expr: &Expr) -> Result<CellValue> {
        match expr {
            Expr::Number(n) => Ok(CellValue::Number(*n)),
            Expr::String(s) => Ok(CellValue::String(s.clone())),
            Expr::Bool(b) => Ok(CellValue::Bool(*b)),
            Expr::Error(e) => Ok(CellValue::Error(e.clone())),
            Expr::Missing => Ok(CellValue::Empty),
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
            .unwrap_or_else(|| self.current_sheet())
            .to_string();
        let col = column_name_to_number(&cell_ref.col)?;
        let row = cell_ref.row;
        self.resolve_cell(&sheet, col, row)
    }

    /// Resolve a single cell, following formula values and detecting cycles.
    ///
    /// A referenced formula is evaluated with the referenced cell's sheet and
    /// position as context; the caller's context is restored afterwards on
    /// both success and error. Cycle keys are sheet-qualified.
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
                let saved_sheet = self.sheet_override.replace(sheet.to_string());
                let saved_position = std::mem::replace(&mut self.position, (col, row));
                let result = crate::formula::parser::parse_formula_with_budget(expr, &self.budget)
                    .and_then(|parsed| self.eval_expr(&parsed));
                self.sheet_override = saved_sheet;
                self.position = saved_position;
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
                if !result.is_finite() {
                    return Ok(CellValue::Error("#NUM!".to_string()));
                }
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

/// Select the numeric value of a cell reached through a reference.
///
/// Aggregates over references use only numeric cells: numbers and dates yield
/// `Some`, and a cached formula result is inspected recursively. Empty,
/// boolean, text, and error cells yield `None`; callers decide whether an
/// error cell propagates.
pub fn reference_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(n) | CellValue::Date(n) => Some(*n),
        CellValue::Formula {
            result: Some(inner),
            ..
        } => reference_number(inner),
        _ => None,
    }
}

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
        CellValue::RichString(runs) => {
            let plain = crate::rich_text::rich_text_to_plain(runs);
            plain
                .parse::<f64>()
                .map_err(|_| Error::FormulaError(format!("cannot convert \"{plain}\" to number")))
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
        CellValue::RichString(runs) => crate::rich_text::rich_text_to_plain(runs),
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
        CellValue::RichString(runs) => {
            let plain = crate::rich_text::rich_text_to_plain(runs);
            match plain.to_ascii_uppercase().as_str() {
                "TRUE" => Ok(true),
                "FALSE" => Ok(false),
                _ => Err(Error::FormulaError(format!(
                    "cannot convert \"{plain}\" to boolean"
                ))),
            }
        }
    }
}

/// Compare two CellValues for ordering. Uses Excel comparison semantics:
/// numbers compare numerically, strings compare case-insensitively,
/// mixed types rank as: empty < number < string < bool.
pub fn compare_values(lhs: &CellValue, rhs: &CellValue) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    match (lhs, rhs) {
        (CellValue::Empty, CellValue::Number(n) | CellValue::Date(n)) => {
            return 0.0_f64.partial_cmp(n).unwrap_or(Ordering::Equal);
        }
        (CellValue::Number(n) | CellValue::Date(n), CellValue::Empty) => {
            return n.partial_cmp(&0.0).unwrap_or(Ordering::Equal);
        }
        _ => {}
    }

    fn type_rank(v: &CellValue) -> u8 {
        match v {
            CellValue::Empty => 0,
            CellValue::Number(_) | CellValue::Date(_) => 1,
            CellValue::String(_) | CellValue::RichString(_) => 2,
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

/// A cell coordinate used as a node in the dependency graph.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct CellCoord {
    pub sheet: String,
    pub col: u32,
    pub row: u32,
}

/// Build a dependency graph from formula cells.
/// Returns a map from each formula cell to the cells it depends on.
pub fn build_dependency_graph(
    formula_cells: &[(CellCoord, String)],
) -> Result<HashMap<CellCoord, Vec<CellCoord>>> {
    let budget = FormulaBudget::secure_calculation();
    let parsed_cells = formula_cells
        .iter()
        .map(|(coord, formula_str)| {
            crate::formula::parser::parse_formula_with_budget(formula_str, &budget)
                .map(|expr| (coord.clone(), expr))
        })
        .collect::<Result<Vec<_>>>()?;
    build_dependency_graph_from_parsed_with_budget(&parsed_cells, &budget)
}

/// Build a dependency graph from formula cells that have already been parsed.
/// Returns a map from each formula cell to the cells it depends on.
pub fn build_dependency_graph_from_parsed(
    formula_cells: &[(CellCoord, Expr)],
) -> Result<HashMap<CellCoord, Vec<CellCoord>>> {
    let budget = FormulaBudget::secure_calculation();
    build_dependency_graph_from_parsed_with_budget(formula_cells, &budget)
}

/// Build a dependency graph while charging an existing calculation budget.
pub(crate) fn build_dependency_graph_from_parsed_with_budget(
    formula_cells: &[(CellCoord, Expr)],
    budget: &FormulaBudget,
) -> Result<HashMap<CellCoord, Vec<CellCoord>>> {
    let mut deps: HashMap<CellCoord, Vec<CellCoord>> = HashMap::new();
    deps.try_reserve(formula_cells.len())
        .map_err(|_| Error::Internal("formula graph allocation failed".to_string()))?;
    for (coord, expr) in formula_cells {
        let refs = extract_cell_refs(expr, &coord.sheet, budget)?;
        deps.insert(coord.clone(), refs);
    }
    Ok(deps)
}

/// Extract all cell references from a parsed expression.
fn extract_cell_refs(
    expr: &Expr,
    current_sheet: &str,
    budget: &FormulaBudget,
) -> Result<Vec<CellCoord>> {
    let mut refs = Vec::new();
    let mut stack = vec![expr];
    while let Some(expr) = stack.pop() {
        match expr {
            Expr::CellRef(cell_ref) => {
                let sheet = cell_ref
                    .sheet
                    .as_deref()
                    .unwrap_or(current_sheet)
                    .to_string();
                if let Ok(col) = column_name_to_number(&cell_ref.col) {
                    budget.charge_dependency_edges(1)?;
                    refs.try_reserve(1).map_err(|_| {
                        Error::Internal("formula graph allocation failed".to_string())
                    })?;
                    refs.push(CellCoord {
                        sheet,
                        col,
                        row: cell_ref.row,
                    });
                }
            }
            Expr::Range { start, end } => {
                let sheet = start.sheet.as_deref().unwrap_or(current_sheet).to_string();
                if let (Ok(start_col), Ok(end_col)) = (
                    column_name_to_number(&start.col),
                    column_name_to_number(&end.col),
                ) {
                    let min_col = start_col.min(end_col);
                    let max_col = start_col.max(end_col);
                    let min_row = start.row.min(end.row);
                    let max_row = start.row.max(end.row);
                    let width = u64::from(max_col - min_col).checked_add(1).ok_or_else(|| {
                        Error::Internal("formula range width overflow".to_string())
                    })?;
                    let height = u64::from(max_row - min_row).checked_add(1).ok_or_else(|| {
                        Error::Internal("formula range height overflow".to_string())
                    })?;
                    let area = width.saturating_mul(height);
                    budget.charge_dependency_edges(area)?;
                    let capacity = usize::try_from(area).map_err(|_| {
                        Error::Internal("formula graph capacity overflow".to_string())
                    })?;
                    refs.try_reserve(capacity).map_err(|_| {
                        Error::Internal("formula graph allocation failed".to_string())
                    })?;
                    for r in min_row..=max_row {
                        for c in min_col..=max_col {
                            refs.push(CellCoord {
                                sheet: sheet.clone(),
                                col: c,
                                row: r,
                            });
                        }
                    }
                }
            }
            Expr::Function { args, .. } => {
                for arg in args.iter().rev() {
                    stack.push(arg);
                }
            }
            Expr::BinaryOp { left, right, .. } => {
                stack.push(right);
                stack.push(left);
            }
            Expr::UnaryOp { operand, .. } | Expr::Paren(operand) => stack.push(operand),
            _ => {}
        }
    }
    Ok(refs)
}

/// Topological sort of formula cells based on dependencies.
/// Returns cells in evaluation order (dependencies first).
/// Detects cycles and returns an error if found.
pub fn topological_sort(
    formula_cells: &[CellCoord],
    deps: &HashMap<CellCoord, Vec<CellCoord>>,
) -> Result<Vec<CellCoord>> {
    let budget = FormulaBudget::secure_calculation();
    topological_sort_with_budget(formula_cells, deps, &budget)
}

/// Deterministically sort formula cells while charging queue operations.
pub(crate) fn topological_sort_with_budget(
    formula_cells: &[CellCoord],
    deps: &HashMap<CellCoord, Vec<CellCoord>>,
    budget: &FormulaBudget,
) -> Result<Vec<CellCoord>> {
    let mut formula_set: HashSet<CellCoord> = HashSet::new();
    formula_set
        .try_reserve(formula_cells.len())
        .map_err(|_| Error::Internal("formula ordering allocation failed".to_string()))?;
    formula_set.extend(formula_cells.iter().cloned());

    let mut in_degree: HashMap<CellCoord, usize> = HashMap::new();
    let mut dependents: HashMap<CellCoord, Vec<CellCoord>> = HashMap::new();
    in_degree
        .try_reserve(formula_cells.len())
        .map_err(|_| Error::Internal("formula ordering allocation failed".to_string()))?;
    dependents
        .try_reserve(formula_cells.len())
        .map_err(|_| Error::Internal("formula ordering allocation failed".to_string()))?;

    for cell in formula_cells {
        in_degree.entry(cell.clone()).or_insert(0);
    }

    for (cell, cell_deps) in deps {
        let dep_count = cell_deps.iter().filter(|d| formula_set.contains(d)).count();
        in_degree.insert(cell.clone(), dep_count);
        for dep in cell_deps {
            if formula_set.contains(dep) {
                let entries = dependents.entry(dep.clone()).or_default();
                entries.try_reserve(1).map_err(|_| {
                    Error::Internal("formula ordering allocation failed".to_string())
                })?;
                entries.push(cell.clone());
            }
        }
    }

    // Kahn's algorithm
    let mut queue: BinaryHeap<Reverse<(String, u32, u32)>> = BinaryHeap::new();
    queue
        .try_reserve(formula_cells.len())
        .map_err(|_| Error::Internal("formula ordering allocation failed".to_string()))?;
    for (cell, degree) in &in_degree {
        if *degree == 0 {
            budget.charge_ready_queue_ops(1)?;
            queue.push(Reverse((cell.sheet.clone(), cell.row, cell.col)));
        }
    }

    let mut sorted = Vec::new();
    sorted
        .try_reserve_exact(formula_cells.len())
        .map_err(|_| Error::Internal("formula ordering allocation failed".to_string()))?;
    let mut visited = 0;

    while let Some(Reverse((sheet, row, col))) = queue.pop() {
        budget.charge_ready_queue_ops(1)?;
        let cell = CellCoord { sheet, col, row };
        sorted.push(cell.clone());
        visited += 1;
        if let Some(cell_dependents) = dependents.get(&cell) {
            for dep in cell_dependents {
                if let Some(deg) = in_degree.get_mut(dep) {
                    *deg -= 1;
                    if *deg == 0 {
                        budget.charge_ready_queue_ops(1)?;
                        queue.push(Reverse((dep.sheet.clone(), dep.row, dep.col)));
                    }
                }
            }
        }
    }

    if visited != formula_cells.len() {
        return Err(Error::CircularReference {
            cell: "multiple cells".to_string(),
        });
    }

    Ok(sorted)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::ast::{BinaryOperator, CellReference, Expr, UnaryOperator};
    use crate::formula::budget::{FormulaBudget, FormulaLimits};
    use crate::formula::parser::parse_formula;

    fn make_snapshot() -> CellSnapshot {
        CellSnapshot::new("Sheet1".to_string())
    }

    struct CountingProvider(std::cell::Cell<u64>);

    impl CellDataProvider for CountingProvider {
        fn get_cell(&self, _sheet: &str, _col: u32, _row: u32) -> CellValue {
            self.0.set(self.0.get() + 1);
            CellValue::Empty
        }

        fn current_sheet(&self) -> &str {
            "Sheet1"
        }
    }

    #[test]
    fn range_budget_is_checked_before_provider_calls() {
        let provider = CountingProvider(std::cell::Cell::new(0));
        let limits = FormulaLimits {
            operation_range_cells: 4,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let mut evaluator = Evaluator::with_budget(&provider, budget);
        let start = CellReference {
            col: "A".to_string(),
            row: 1,
            abs_col: false,
            abs_row: false,
            sheet: None,
        };
        let end = CellReference {
            col: "B".to_string(),
            row: 2,
            abs_col: false,
            abs_row: false,
            sheet: None,
        };

        assert_eq!(evaluator.expand_range(&start, &end).unwrap().len(), 4);
        assert_eq!(provider.0.get(), 4);
        let error = evaluator.expand_range(&start, &start).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension_code(),
            "range_cell_count"
        );
        assert_eq!(provider.0.get(), 4);
    }

    #[test]
    fn ready_queue_budget_is_cumulative_and_deterministic() {
        let a = CellCoord {
            sheet: "Sheet1".to_string(),
            col: 1,
            row: 1,
        };
        let b = CellCoord {
            sheet: "Sheet1".to_string(),
            col: 2,
            row: 1,
        };
        let deps = HashMap::from([(a.clone(), vec![]), (b.clone(), vec![])]);
        let limits = FormulaLimits {
            operation_ready_queue_ops: 4,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        assert_eq!(
            topological_sort_with_budget(&[b.clone(), a.clone()], &deps, &budget).unwrap(),
            vec![a.clone(), b.clone()]
        );

        let limits = FormulaLimits {
            operation_ready_queue_ops: 3,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let error = topological_sort_with_budget(&[a, b], &deps, &budget).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension_code(),
            "operation_count"
        );
    }

    #[test]
    fn large_independent_graph_has_exact_bounded_queue_work() {
        let cells: Vec<CellCoord> = (1..=1_024)
            .rev()
            .map(|row| CellCoord {
                sheet: "Sheet1".to_string(),
                col: 1,
                row,
            })
            .collect();
        let deps: HashMap<CellCoord, Vec<CellCoord>> = cells
            .iter()
            .cloned()
            .map(|cell| (cell, Vec::new()))
            .collect();
        let limits = FormulaLimits {
            operation_ready_queue_ops: 2_048,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let sorted = topological_sort_with_budget(&cells, &deps, &budget).unwrap();
        assert_eq!((sorted[0].row, sorted[1_023].row), (1, 1_024));

        let limits = FormulaLimits {
            operation_ready_queue_ops: 2_047,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let error = topological_sort_with_budget(&cells, &deps, &budget).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!((metadata.limit(), metadata.observed()), (2_047, 2_048));
    }

    #[test]
    fn dependency_edge_boundaries_are_exact() {
        let coord = CellCoord {
            sheet: "Sheet1".to_string(),
            col: 1,
            row: 1,
        };
        let expression = parse_formula("SUM(A1,B1,C1,D1)").unwrap();
        let limits = FormulaLimits {
            operation_dependency_edges: 4,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let graph = build_dependency_graph_from_parsed_with_budget(
            &[(coord.clone(), expression.clone())],
            &budget,
        )
        .unwrap();
        assert_eq!(graph[&coord].len(), 4);

        let limits = FormulaLimits {
            operation_dependency_edges: 3,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let error = build_dependency_graph_from_parsed_with_budget(&[(coord, expression)], &budget)
            .unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!((metadata.limit(), metadata.observed()), (3, 4));
        assert_eq!(metadata.dimension_code(), "dependency_edge_count");
    }

    #[test]
    fn dependency_range_is_rejected_before_materialization() {
        let limits = FormulaLimits {
            operation_range_cells: 0,
            operation_dependency_edges: 4,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let coord = CellCoord {
            sheet: "Sheet1".to_string(),
            col: 1,
            row: 1,
        };
        let expression = parse_formula("A1:C2").unwrap();

        let error = build_dependency_graph_from_parsed_with_budget(&[(coord, expression)], &budget)
            .unwrap_err();

        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "dependency_edge_count");
        assert_eq!((metadata.limit(), metadata.observed()), (4, 6));
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

    #[test]
    fn eval_missing_argument_as_empty() {
        let snap = make_snapshot();
        let expr = parse_formula("IF(TRUE,,1)").unwrap();
        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Empty);
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
    fn eval_negative_fractional_power_returns_num_error() {
        let snap = make_snapshot();
        let expr = parse_formula("(-8)^0.5").unwrap();

        assert_eq!(
            evaluate(&expr, &snap).unwrap(),
            CellValue::Error("#NUM!".to_string())
        );
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

    #[test]
    fn eval_sum_range_propagates_cell_errors() {
        let mut snap = make_snapshot();
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(1.0));
        snap.set_cell("Sheet1", 1, 2, CellValue::Error("#DIV/0!".to_string()));
        let expr = parse_formula("SUM(A1:A2)").unwrap();

        let error = evaluate(&expr, &snap).unwrap_err();
        assert!(error.to_string().contains("#DIV/0!"));
    }

    #[test]
    fn eval_sum_range_skips_non_numeric_text() {
        let mut snap = make_snapshot();
        snap.set_cell("Sheet1", 1, 1, CellValue::Number(1.0));
        snap.set_cell("Sheet1", 1, 2, CellValue::String("text".to_string()));
        let expr = parse_formula("SUM(A1:A2)").unwrap();

        assert_eq!(evaluate(&expr, &snap).unwrap(), CellValue::Number(1.0));
    }

    #[test]
    fn eval_blank_cell_compares_as_zero_against_numbers() {
        let snap = make_snapshot();

        assert_eq!(
            evaluate(&parse_formula("A1=0").unwrap(), &snap).unwrap(),
            CellValue::Bool(true)
        );
        assert_eq!(
            evaluate(&parse_formula("A1<>0").unwrap(), &snap).unwrap(),
            CellValue::Bool(false)
        );
        assert_eq!(
            evaluate(&parse_formula("A1<1").unwrap(), &snap).unwrap(),
            CellValue::Bool(true)
        );
    }

    #[test]
    fn eval_row_and_column_without_reference_use_evaluator_position() {
        let snap = make_snapshot();
        let mut evaluator = Evaluator::new(&snap);
        evaluator.set_position(3, 5);

        assert_eq!(
            evaluator
                .eval_expr(&parse_formula("ROW()").unwrap())
                .unwrap(),
            CellValue::Number(5.0)
        );
        assert_eq!(
            evaluator
                .eval_expr(&parse_formula("COLUMN()").unwrap())
                .unwrap(),
            CellValue::Number(3.0)
        );
    }

    #[test]
    fn eval_row_and_column_preserve_default_and_explicit_references() {
        let snap = make_snapshot();

        assert_eq!(
            evaluate(&parse_formula("ROW()").unwrap(), &snap).unwrap(),
            CellValue::Number(1.0)
        );
        assert_eq!(
            evaluate(&parse_formula("ROW(A7)").unwrap(), &snap).unwrap(),
            CellValue::Number(7.0)
        );
        assert_eq!(
            evaluate(&parse_formula("COLUMN(C7)").unwrap(), &snap).unwrap(),
            CellValue::Number(3.0)
        );
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
            coerce_to_number(&CellValue::String("3.15".to_string())).unwrap(),
            3.15
        );
    }

    #[test]
    fn coerce_string_to_number_fail() {
        assert!(coerce_to_number(&CellValue::String("abc".to_string())).is_err());
    }

    #[test]
    fn coerce_number_to_string() {
        assert_eq!(coerce_to_string(&CellValue::Number(42.0)), "42");
        assert_eq!(coerce_to_string(&CellValue::Number(3.15)), "3.15");
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
        for _ in 0..1_025 {
            expr = Expr::Paren(Box::new(expr));
        }
        let error = evaluate(&expr, &snap).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "nesting_depth");
        assert_eq!((metadata.limit(), metadata.observed()), (1_024, 1_025));
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

    fn coord(sheet: &str, col: u32, row: u32) -> CellCoord {
        CellCoord {
            sheet: sheet.to_string(),
            col,
            row,
        }
    }

    #[test]
    fn test_build_dependency_graph_simple() {
        // A1 = B1 + C1
        let formula_cells = vec![(coord("Sheet1", 1, 1), "B1+C1".to_string())];
        let deps = build_dependency_graph(&formula_cells).unwrap();
        let a1_deps = deps.get(&coord("Sheet1", 1, 1)).unwrap();
        assert!(a1_deps.contains(&coord("Sheet1", 2, 1))); // B1
        assert!(a1_deps.contains(&coord("Sheet1", 3, 1))); // C1
        assert_eq!(a1_deps.len(), 2);
    }

    #[test]
    fn test_build_dependency_graph_range() {
        // A1 = SUM(B1:B5)
        let formula_cells = vec![(coord("Sheet1", 1, 1), "SUM(B1:B5)".to_string())];
        let deps = build_dependency_graph(&formula_cells).unwrap();
        let a1_deps = deps.get(&coord("Sheet1", 1, 1)).unwrap();
        assert_eq!(a1_deps.len(), 5);
        for r in 1..=5 {
            assert!(a1_deps.contains(&coord("Sheet1", 2, r)));
        }
    }

    #[test]
    fn test_build_dependency_graph_cross_sheet() {
        // Sheet1!A1 = Sheet2!B1
        let formula_cells = vec![(coord("Sheet1", 1, 1), "Sheet2!B1".to_string())];
        let deps = build_dependency_graph(&formula_cells).unwrap();
        let a1_deps = deps.get(&coord("Sheet1", 1, 1)).unwrap();
        assert_eq!(a1_deps.len(), 1);
        assert_eq!(a1_deps[0], coord("Sheet2", 2, 1));
    }

    #[test]
    fn test_extract_cell_refs_from_nested_expr() {
        let expr = parse_formula("A1+SUM(B1:B3,C1)").unwrap();
        let budget = FormulaBudget::secure_calculation();
        let refs = extract_cell_refs(&expr, "Sheet1", &budget).unwrap();
        assert!(refs.contains(&coord("Sheet1", 1, 1))); // A1
        assert!(refs.contains(&coord("Sheet1", 2, 1))); // B1
        assert!(refs.contains(&coord("Sheet1", 2, 2))); // B2
        assert!(refs.contains(&coord("Sheet1", 2, 3))); // B3
        assert!(refs.contains(&coord("Sheet1", 3, 1))); // C1
        assert_eq!(refs.len(), 5);
    }

    #[test]
    fn test_topological_sort_linear_chain() {
        // A2 = A1 + 1, A3 = A2 + 1 (A1 is a value, not a formula)
        let formula_cells = vec![coord("Sheet1", 1, 2), coord("Sheet1", 1, 3)];
        let mut deps = HashMap::new();
        deps.insert(coord("Sheet1", 1, 2), vec![coord("Sheet1", 1, 1)]); // A2 depends on A1 (value)
        deps.insert(
            coord("Sheet1", 1, 3),
            vec![coord("Sheet1", 1, 2)], // A3 depends on A2 (formula)
        );
        let sorted = topological_sort(&formula_cells, &deps).unwrap();
        assert_eq!(sorted.len(), 2);
        // A2 must come before A3
        let pos_a2 = sorted
            .iter()
            .position(|c| *c == coord("Sheet1", 1, 2))
            .unwrap();
        let pos_a3 = sorted
            .iter()
            .position(|c| *c == coord("Sheet1", 1, 3))
            .unwrap();
        assert!(pos_a2 < pos_a3);
    }

    #[test]
    fn test_topological_sort_diamond() {
        // A1 = B1 + C1, B1 = D1, C1 = D1
        // D1 is a value, so formula_cells = [A1, B1, C1]
        let formula_cells = vec![
            coord("Sheet1", 1, 1), // A1
            coord("Sheet1", 2, 1), // B1
            coord("Sheet1", 3, 1), // C1
        ];
        let mut deps = HashMap::new();
        deps.insert(
            coord("Sheet1", 1, 1),
            vec![coord("Sheet1", 2, 1), coord("Sheet1", 3, 1)],
        );
        deps.insert(coord("Sheet1", 2, 1), vec![coord("Sheet1", 4, 1)]); // D1 value
        deps.insert(coord("Sheet1", 3, 1), vec![coord("Sheet1", 4, 1)]); // D1 value

        let sorted = topological_sort(&formula_cells, &deps).unwrap();
        assert_eq!(sorted.len(), 3);
        let pos_a1 = sorted
            .iter()
            .position(|c| *c == coord("Sheet1", 1, 1))
            .unwrap();
        let pos_b1 = sorted
            .iter()
            .position(|c| *c == coord("Sheet1", 2, 1))
            .unwrap();
        let pos_c1 = sorted
            .iter()
            .position(|c| *c == coord("Sheet1", 3, 1))
            .unwrap();
        // B1 and C1 must come before A1
        assert!(pos_b1 < pos_a1);
        assert!(pos_c1 < pos_a1);
    }

    #[test]
    fn test_topological_sort_cycle_detection() {
        // A1 = B1, B1 = A1
        let formula_cells = vec![coord("Sheet1", 1, 1), coord("Sheet1", 2, 1)];
        let mut deps = HashMap::new();
        deps.insert(coord("Sheet1", 1, 1), vec![coord("Sheet1", 2, 1)]);
        deps.insert(coord("Sheet1", 2, 1), vec![coord("Sheet1", 1, 1)]);

        let result = topological_sort(&formula_cells, &deps);
        assert!(result.is_err());
        let err_str = result.unwrap_err().to_string();
        assert!(
            err_str.contains("circular reference"),
            "expected circular reference error, got: {err_str}"
        );
    }

    #[test]
    fn test_topological_sort_no_formulas() {
        let formula_cells: Vec<CellCoord> = vec![];
        let deps: HashMap<CellCoord, Vec<CellCoord>> = HashMap::new();
        let sorted = topological_sort(&formula_cells, &deps).unwrap();
        assert!(sorted.is_empty());
    }

    #[test]
    fn test_set_current_sheet() {
        let mut snap = CellSnapshot::new("Sheet1".to_string());
        assert_eq!(snap.current_sheet(), "Sheet1");
        snap.set_current_sheet("Sheet2");
        assert_eq!(snap.current_sheet(), "Sheet2");
    }
}
