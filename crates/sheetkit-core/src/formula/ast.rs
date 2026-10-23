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
    /// Omitted function argument (e.g., the second argument in `IF(A1,,1)`)
    Missing,
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

impl Expr {
    /// Return `(nodes, recursive nesting, binary-chain depth)` iteratively.
    ///
    /// Binary branches are tracked independently because the parser builds
    /// them iteratively; function, unary, and parenthesized nodes consume the
    /// recursive-parser nesting budget.
    pub(crate) fn structural_metrics(&self) -> (u64, u64, u64) {
        let mut nodes = 0_u64;
        let mut max_depth = 0_u64;
        let mut max_binary_chain = 0_u64;
        let mut stack = MetricStack::new(self);
        stack.push((self, 1_u64, 0_u64));

        while let Some((expr, depth, parent_binary_chain)) = stack.pop() {
            nodes = nodes.saturating_add(1);
            max_depth = max_depth.max(depth);
            match expr {
                Self::Function { args, .. } => {
                    for arg in args.iter().rev() {
                        stack.push((arg, depth.saturating_add(1), 0));
                    }
                }
                Self::BinaryOp { left, right, .. } => {
                    let chain = parent_binary_chain.saturating_add(1);
                    max_binary_chain = max_binary_chain.max(chain);
                    stack.push((right, depth, 0));
                    stack.push((left, depth, chain));
                }
                Self::UnaryOp { operand, .. } | Self::Paren(operand) => {
                    stack.push((operand, depth.saturating_add(1), 0));
                }
                _ => {}
            }
        }

        (nodes, max_depth, max_binary_chain)
    }
}

/// LIFO work stack for [`Expr::structural_metrics`] that keeps typical
/// formulas off the heap and spills deeper traversals to a `Vec`.
struct MetricStack<'a> {
    inline: [(&'a Expr, u64, u64); MetricStack::INLINE],
    len: usize,
    spill: Vec<(&'a Expr, u64, u64)>,
}

impl<'a> MetricStack<'a> {
    const INLINE: usize = 16;

    fn new(filler: &'a Expr) -> Self {
        Self {
            inline: [(filler, 0, 0); Self::INLINE],
            len: 0,
            spill: Vec::new(),
        }
    }

    fn push(&mut self, entry: (&'a Expr, u64, u64)) {
        if self.spill.is_empty() && self.len < Self::INLINE {
            self.inline[self.len] = entry;
            self.len += 1;
        } else {
            self.spill.push(entry);
        }
    }

    fn pop(&mut self) -> Option<(&'a Expr, u64, u64)> {
        if let Some(entry) = self.spill.pop() {
            return Some(entry);
        }
        if self.len == 0 {
            return None;
        }
        self.len -= 1;
        Some(self.inline[self.len])
    }
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
