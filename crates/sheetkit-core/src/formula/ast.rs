//! AST types for parsed Excel formulas.

/// A parsed Excel formula expression.
#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    /// Numeric literal (e.g., 42, 3.14)
    Number(f64),
    /// String literal (e.g., "hello")
    String(String),
    /// Boolean literal (TRUE/FALSE)
    Bool(bool),
    /// Error literal (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!)
    Error(String),
    /// Cell reference (e.g., A1, $B$2)
    CellRef(CellReference),
    /// Range reference (e.g., A1:B10)
    Range {
        start: CellReference,
        end: CellReference,
    },
    /// Function call (e.g., SUM(A1:A10))
    Function { name: String, args: Vec<Expr> },
    /// Binary operation (e.g., A1 + B1)
    BinaryOp {
        op: BinaryOperator,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    /// Unary operation (e.g., -A1, +5)
    UnaryOp {
        op: UnaryOperator,
        operand: Box<Expr>,
    },
    /// Parenthesized expression
    Paren(Box<Expr>),
}

/// A cell reference with optional absolute markers.
#[derive(Debug, Clone, PartialEq)]
pub struct CellReference {
    /// Column name (e.g., "A", "AB")
    pub col: String,
    /// Row number (1-based)
    pub row: u32,
    /// True if column is absolute ($A)
    pub abs_col: bool,
    /// True if row is absolute ($1)
    pub abs_row: bool,
    /// Optional sheet name (e.g., Sheet1!A1)
    pub sheet: Option<String>,
}

/// Binary operators.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOperator {
    /// Addition (+)
    Add,
    /// Subtraction (-)
    Sub,
    /// Multiplication (*)
    Mul,
    /// Division (/)
    Div,
    /// Exponentiation (^)
    Pow,
    /// Concatenation (&)
    Concat,
    /// Equal (=)
    Eq,
    /// Not equal (<>)
    Ne,
    /// Less than (<)
    Lt,
    /// Less than or equal (<=)
    Le,
    /// Greater than (>)
    Gt,
    /// Greater than or equal (>=)
    Ge,
}

/// Unary operators.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOperator {
    /// Negation (-)
    Neg,
    /// Positive (+)
    Pos,
    /// Percent (%)
    Percent,
}
