//! Excel formula parser and evaluation engine.
//!
//! Parses formula strings into an AST representation and evaluates them
//! against cell data.

pub mod ast;
pub mod budget;
pub mod eval;
pub mod functions;
pub mod parser;

pub use ast::{BinaryOperator, CellReference, Expr, UnaryOperator};
pub use budget::{FormulaBudget, FormulaLimits};
pub use eval::{
    build_dependency_graph, evaluate, topological_sort, CellCoord, CellDataProvider, CellSnapshot,
    Evaluator,
};
pub use parser::parse_formula;
