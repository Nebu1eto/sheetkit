//! nom-based parser for Excel formula strings.
//!
//! The formula string does NOT include the leading `=` sign.
//!
//! Operator precedence (lowest to highest):
//! 1. Comparison (=, <>, <, <=, >, >=)
//! 2. Concatenation (&)
//! 3. Addition / Subtraction (+, -)
//! 4. Multiplication / Division (*, /)
//! 5. Exponentiation (^)
//! 6. Unary (+, -, %)
//! 7. Primary (literals, references, functions, parens)

use nom::{
    branch::alt,
    bytes::complete::{tag, tag_no_case, take_while1},
    character::complete::{alpha1, char, multispace0},
    combinator::{map, opt, recognize, value},
    multi::{many0, separated_list0},
    sequence::{delimited, pair, preceded, terminated},
    IResult,
};

use super::ast::{BinaryOperator, CellReference, Expr, UnaryOperator};
use crate::error::{Error, Result};

/// Parse an Excel formula string into an AST expression.
///
/// The input should NOT include the leading `=` sign.
///
/// # Errors
///
/// Returns an error if the formula string cannot be parsed.
pub fn parse_formula(input: &str) -> Result<Expr> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return Err(Error::Internal("empty formula".to_string()));
    }
    match parse_expr(trimmed) {
        Ok(("", expr)) => Ok(expr),
        Ok((remaining, _)) => Err(Error::Internal(format!(
            "unexpected trailing input: {remaining}"
        ))),
        Err(e) => Err(Error::Internal(format!("formula parse error: {e}"))),
    }
}

// ---------------------------------------------------------------------------
// Whitespace helper
// ---------------------------------------------------------------------------

/// Wraps a parser to consume optional surrounding whitespace.
fn ws<'a, F, O>(inner: F) -> impl FnMut(&'a str) -> IResult<&'a str, O>
where
    F: FnMut(&'a str) -> IResult<&'a str, O>,
{
    delimited(multispace0, inner, multispace0)
}

// ---------------------------------------------------------------------------
// Expression parsers (ordered by precedence, lowest first)
// ---------------------------------------------------------------------------

/// Top-level expression: comparison operators (lowest precedence).
fn parse_expr(input: &str) -> IResult<&str, Expr> {
    let (input, left) = parse_concat_expr(input)?;
    let (input, rest) = many0(pair(ws(parse_comparison_op), parse_concat_expr))(input)?;
    Ok((input, fold_binary(left, rest)))
}

fn parse_comparison_op(input: &str) -> IResult<&str, BinaryOperator> {
    alt((
        value(BinaryOperator::Le, tag("<=")),
        value(BinaryOperator::Ge, tag(">=")),
        value(BinaryOperator::Ne, tag("<>")),
        value(BinaryOperator::Lt, tag("<")),
        value(BinaryOperator::Gt, tag(">")),
        value(BinaryOperator::Eq, tag("=")),
    ))(input)
}

/// Concatenation (&).
fn parse_concat_expr(input: &str) -> IResult<&str, Expr> {
    let (input, left) = parse_additive(input)?;
    let (input, rest) = many0(pair(
        ws(value(BinaryOperator::Concat, tag("&"))),
        parse_additive,
    ))(input)?;
    Ok((input, fold_binary(left, rest)))
}

/// Addition and subtraction.
fn parse_additive(input: &str) -> IResult<&str, Expr> {
    let (input, left) = parse_multiplicative(input)?;
    let (input, rest) = many0(pair(ws(parse_add_sub_op), parse_multiplicative))(input)?;
    Ok((input, fold_binary(left, rest)))
}

fn parse_add_sub_op(input: &str) -> IResult<&str, BinaryOperator> {
    alt((
        value(BinaryOperator::Add, tag("+")),
        value(BinaryOperator::Sub, tag("-")),
    ))(input)
}

/// Multiplication and division.
fn parse_multiplicative(input: &str) -> IResult<&str, Expr> {
    let (input, left) = parse_power(input)?;
    let (input, rest) = many0(pair(ws(parse_mul_div_op), parse_power))(input)?;
    Ok((input, fold_binary(left, rest)))
}

fn parse_mul_div_op(input: &str) -> IResult<&str, BinaryOperator> {
    alt((
        value(BinaryOperator::Mul, tag("*")),
        value(BinaryOperator::Div, tag("/")),
    ))(input)
}

/// Exponentiation (^).
fn parse_power(input: &str) -> IResult<&str, Expr> {
    let (input, left) = parse_unary(input)?;
    let (input, rest) = many0(pair(ws(value(BinaryOperator::Pow, tag("^"))), parse_unary))(input)?;
    Ok((input, fold_binary(left, rest)))
}

/// Unary +/- prefix and % postfix.
fn parse_unary(input: &str) -> IResult<&str, Expr> {
    let input = input.trim_start();
    // Try unary minus
    if let Ok((rest, _)) = tag::<&str, &str, nom::error::Error<&str>>("-")(input) {
        let (rest, operand) = parse_unary(rest)?;
        return Ok((
            rest,
            Expr::UnaryOp {
                op: UnaryOperator::Neg,
                operand: Box::new(operand),
            },
        ));
    }
    // Try unary plus
    if let Ok((rest, _)) = tag::<&str, &str, nom::error::Error<&str>>("+")(input) {
        let (rest, operand) = parse_unary(rest)?;
        return Ok((
            rest,
            Expr::UnaryOp {
                op: UnaryOperator::Pos,
                operand: Box::new(operand),
            },
        ));
    }
    // Otherwise parse primary and check for postfix %
    let (input, expr) = parse_primary(input)?;
    let (input, pcts) = many0(ws(value(UnaryOperator::Percent, tag("%"))))(input)?;
    let result = pcts.into_iter().fold(expr, |acc, op| Expr::UnaryOp {
        op,
        operand: Box::new(acc),
    });
    Ok((input, result))
}

/// Primary expressions: literals, references, function calls, parenthesized expressions.
fn parse_primary(input: &str) -> IResult<&str, Expr> {
    let input = input.trim_start();
    alt((
        parse_paren_expr,
        parse_string_literal,
        parse_error_literal,
        parse_bool_literal,
        parse_function_call,
        parse_cell_ref_or_range,
        parse_number_literal,
    ))(input)
}

// ---------------------------------------------------------------------------
// Literal parsers
// ---------------------------------------------------------------------------

/// Parse a numeric literal (integer or decimal).
fn parse_number_literal(input: &str) -> IResult<&str, Expr> {
    let (input, num_str) = recognize(pair(
        take_while1(|c: char| c.is_ascii_digit()),
        opt(pair(tag("."), take_while1(|c: char| c.is_ascii_digit()))),
    ))(input)?;
    let n: f64 = num_str.parse().map_err(|_| {
        nom::Err::Error(nom::error::Error::new(input, nom::error::ErrorKind::Float))
    })?;
    Ok((input, Expr::Number(n)))
}

/// Parse a string literal with `"` delimiters. `""` is an escaped quote (Excel convention).
fn parse_string_literal(input: &str) -> IResult<&str, Expr> {
    let (input, _) = tag("\"")(input)?;
    let mut result = String::new();
    let mut remaining = input;
    loop {
        if remaining.is_empty() {
            return Err(nom::Err::Error(nom::error::Error::new(
                remaining,
                nom::error::ErrorKind::Tag,
            )));
        }
        if remaining.starts_with("\"\"") {
            result.push('"');
            remaining = &remaining[2..];
        } else if remaining.starts_with('"') {
            remaining = &remaining[1..];
            break;
        } else {
            let c = remaining.chars().next().unwrap();
            result.push(c);
            remaining = &remaining[c.len_utf8()..];
        }
    }
    Ok((remaining, Expr::String(result)))
}

/// Parse a boolean literal: TRUE or FALSE (case-insensitive).
fn parse_bool_literal(input: &str) -> IResult<&str, Expr> {
    let (input, val) = alt((
        value(
            true,
            terminated(tag_no_case("TRUE"), not_alnum_or_underscore),
        ),
        value(
            false,
            terminated(tag_no_case("FALSE"), not_alnum_or_underscore),
        ),
    ))(input)?;
    Ok((input, Expr::Bool(val)))
}

/// Succeeds if the next character is NOT alphanumeric or underscore, or if at end of input.
/// This prevents "TRUE1" from being parsed as Bool(true) with leftover "1".
fn not_alnum_or_underscore(input: &str) -> IResult<&str, ()> {
    if input.is_empty() {
        return Ok((input, ()));
    }
    let c = input.chars().next().unwrap();
    if c.is_alphanumeric() || c == '_' {
        Err(nom::Err::Error(nom::error::Error::new(
            input,
            nom::error::ErrorKind::Alpha,
        )))
    } else {
        Ok((input, ()))
    }
}

/// Parse an Excel error literal (#DIV/0!, #VALUE!, etc.).
fn parse_error_literal(input: &str) -> IResult<&str, Expr> {
    let (input, err) = alt((
        tag("#DIV/0!"),
        tag("#VALUE!"),
        tag("#REF!"),
        tag("#NAME?"),
        tag("#NUM!"),
        tag("#NULL!"),
        tag("#N/A"),
    ))(input)?;
    Ok((input, Expr::Error(err.to_string())))
}

// ---------------------------------------------------------------------------
// Cell reference and range parsers
// ---------------------------------------------------------------------------

/// Parse a cell reference (possibly with a sheet prefix) or a range (A1:B10).
fn parse_cell_ref_or_range(input: &str) -> IResult<&str, Expr> {
    let (input, first) = parse_single_cell_ref(input)?;
    // Check for range operator ':'
    if let Ok((input, _)) = tag::<&str, &str, nom::error::Error<&str>>(":")(input) {
        let (input, second) = parse_single_cell_ref(input)?;
        Ok((
            input,
            Expr::Range {
                start: first,
                end: second,
            },
        ))
    } else {
        Ok((input, Expr::CellRef(first)))
    }
}

/// Parse a single cell reference, optionally prefixed with a sheet name.
fn parse_single_cell_ref(input: &str) -> IResult<&str, CellReference> {
    // Try to parse sheet prefix first
    let (input, sheet) = opt(parse_sheet_prefix)(input)?;
    // Parse optional $ for abs col
    let (input, abs_col) = map(opt(tag("$")), |o| o.is_some())(input)?;
    // Parse column letters
    let (input, col) = alpha1(input)?;
    // Make sure column letters are A-Z only
    let col_upper = col.to_uppercase();
    if !col_upper.chars().all(|c| c.is_ascii_uppercase()) {
        return Err(nom::Err::Error(nom::error::Error::new(
            input,
            nom::error::ErrorKind::Alpha,
        )));
    }
    // Parse optional $ for abs row
    let (input, abs_row) = map(opt(tag("$")), |o| o.is_some())(input)?;
    // Parse row digits
    let (input, row_str) = take_while1(|c: char| c.is_ascii_digit())(input)?;
    let row: u32 = row_str.parse().map_err(|_| {
        nom::Err::Error(nom::error::Error::new(input, nom::error::ErrorKind::Digit))
    })?;
    Ok((
        input,
        CellReference {
            col: col_upper,
            row,
            abs_col,
            abs_row,
            sheet,
        },
    ))
}

/// Parse a sheet prefix like `Sheet1!` or `'Sheet Name'!`.
fn parse_sheet_prefix(input: &str) -> IResult<&str, String> {
    alt((parse_quoted_sheet_prefix, parse_unquoted_sheet_prefix))(input)
}

/// Parse an unquoted sheet prefix: `SheetName!`
fn parse_unquoted_sheet_prefix(input: &str) -> IResult<&str, String> {
    let (input, name) = terminated(
        recognize(pair(
            take_while1(|c: char| c.is_alphanumeric() || c == '_'),
            many0(take_while1(|c: char| c.is_alphanumeric() || c == '_')),
        )),
        tag("!"),
    )(input)?;
    Ok((input, name.to_string()))
}

/// Parse a quoted sheet prefix: `'Sheet Name'!`
fn parse_quoted_sheet_prefix(input: &str) -> IResult<&str, String> {
    let (input, _) = tag("'")(input)?;
    let mut result = String::new();
    let mut remaining = input;
    loop {
        if remaining.is_empty() {
            return Err(nom::Err::Error(nom::error::Error::new(
                remaining,
                nom::error::ErrorKind::Tag,
            )));
        }
        if remaining.starts_with("''") {
            result.push('\'');
            remaining = &remaining[2..];
        } else if remaining.starts_with('\'') {
            remaining = &remaining[1..];
            break;
        } else {
            let c = remaining.chars().next().unwrap();
            result.push(c);
            remaining = &remaining[c.len_utf8()..];
        }
    }
    let (remaining, _) = tag("!")(remaining)?;
    Ok((remaining, result))
}

// ---------------------------------------------------------------------------
// Function call parser
// ---------------------------------------------------------------------------

/// Parse a function call: `FUNCNAME(arg1, arg2, ...)`.
fn parse_function_call(input: &str) -> IResult<&str, Expr> {
    // Function name: letters, digits, underscores, dots (e.g., _xlfn.CONCAT)
    let (input, name) = recognize(pair(
        alt((alpha1, tag("_"))),
        many0(alt((
            take_while1(|c: char| c.is_alphanumeric() || c == '_'),
            tag("."),
        ))),
    ))(input)?;
    // Must be followed by '('
    let (input, _) = preceded(multispace0, char('('))(input)?;
    let (input, _) = multispace0(input)?;
    // Parse arguments separated by commas
    let (input, args) = separated_list0(ws(char(',')), parse_expr)(input)?;
    let (input, _) = preceded(multispace0, char(')'))(input)?;
    Ok((
        input,
        Expr::Function {
            name: name.to_uppercase(),
            args,
        },
    ))
}

// ---------------------------------------------------------------------------
// Parenthesized expression
// ---------------------------------------------------------------------------

/// Parse a parenthesized expression.
fn parse_paren_expr(input: &str) -> IResult<&str, Expr> {
    let (input, _) = char('(')(input)?;
    let (input, _) = multispace0(input)?;
    let (input, expr) = parse_expr(input)?;
    let (input, _) = preceded(multispace0, char(')'))(input)?;
    Ok((input, Expr::Paren(Box::new(expr))))
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Fold a left-associative chain of binary operations into a single AST node.
fn fold_binary(first: Expr, rest: Vec<(BinaryOperator, Expr)>) -> Expr {
    rest.into_iter()
        .fold(first, |left, (op, right)| Expr::BinaryOp {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
}

// ===========================================================================
// Tests
// ===========================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::ast::{BinaryOperator, CellReference, Expr, UnaryOperator};

    // -----------------------------------------------------------------------
    // Number parsing
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_integer() {
        let result = parse_formula("42").unwrap();
        assert_eq!(result, Expr::Number(42.0));
    }

    #[test]
    fn test_parse_decimal() {
        let result = parse_formula("3.14").unwrap();
        assert_eq!(result, Expr::Number(3.14));
    }

    #[test]
    fn test_parse_negative_number() {
        let result = parse_formula("-5").unwrap();
        assert_eq!(
            result,
            Expr::UnaryOp {
                op: UnaryOperator::Neg,
                operand: Box::new(Expr::Number(5.0)),
            }
        );
    }

    // -----------------------------------------------------------------------
    // String parsing
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_string() {
        let result = parse_formula("\"hello\"").unwrap();
        assert_eq!(result, Expr::String("hello".to_string()));
    }

    #[test]
    fn test_parse_empty_string() {
        let result = parse_formula("\"\"").unwrap();
        assert_eq!(result, Expr::String(String::new()));
    }

    // -----------------------------------------------------------------------
    // Boolean parsing
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_true() {
        let result = parse_formula("TRUE").unwrap();
        assert_eq!(result, Expr::Bool(true));
    }

    #[test]
    fn test_parse_false() {
        let result = parse_formula("FALSE").unwrap();
        assert_eq!(result, Expr::Bool(false));
    }

    // -----------------------------------------------------------------------
    // Error parsing
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_error_div0() {
        let result = parse_formula("#DIV/0!").unwrap();
        assert_eq!(result, Expr::Error("#DIV/0!".to_string()));
    }

    #[test]
    fn test_parse_error_na() {
        let result = parse_formula("#N/A").unwrap();
        assert_eq!(result, Expr::Error("#N/A".to_string()));
    }

    // -----------------------------------------------------------------------
    // Cell reference parsing
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_cell_ref() {
        let result = parse_formula("A1").unwrap();
        assert_eq!(
            result,
            Expr::CellRef(CellReference {
                col: "A".to_string(),
                row: 1,
                abs_col: false,
                abs_row: false,
                sheet: None,
            })
        );
    }

    #[test]
    fn test_parse_abs_cell_ref() {
        let result = parse_formula("$A$1").unwrap();
        assert_eq!(
            result,
            Expr::CellRef(CellReference {
                col: "A".to_string(),
                row: 1,
                abs_col: true,
                abs_row: true,
                sheet: None,
            })
        );
    }

    #[test]
    fn test_parse_mixed_cell_ref() {
        let result = parse_formula("$A1").unwrap();
        assert_eq!(
            result,
            Expr::CellRef(CellReference {
                col: "A".to_string(),
                row: 1,
                abs_col: true,
                abs_row: false,
                sheet: None,
            })
        );
    }

    // -----------------------------------------------------------------------
    // Range parsing
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_range() {
        let result = parse_formula("A1:B10").unwrap();
        assert_eq!(
            result,
            Expr::Range {
                start: CellReference {
                    col: "A".to_string(),
                    row: 1,
                    abs_col: false,
                    abs_row: false,
                    sheet: None,
                },
                end: CellReference {
                    col: "B".to_string(),
                    row: 10,
                    abs_col: false,
                    abs_row: false,
                    sheet: None,
                },
            }
        );
    }

    // -----------------------------------------------------------------------
    // Arithmetic
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_addition() {
        let result = parse_formula("1+2").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Add,
                left: Box::new(Expr::Number(1.0)),
                right: Box::new(Expr::Number(2.0)),
            }
        );
    }

    #[test]
    fn test_parse_subtraction() {
        let result = parse_formula("5-3").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Sub,
                left: Box::new(Expr::Number(5.0)),
                right: Box::new(Expr::Number(3.0)),
            }
        );
    }

    #[test]
    fn test_parse_multiplication() {
        let result = parse_formula("2*3").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Mul,
                left: Box::new(Expr::Number(2.0)),
                right: Box::new(Expr::Number(3.0)),
            }
        );
    }

    #[test]
    fn test_parse_division() {
        let result = parse_formula("10/2").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Div,
                left: Box::new(Expr::Number(10.0)),
                right: Box::new(Expr::Number(2.0)),
            }
        );
    }

    // -----------------------------------------------------------------------
    // Operator precedence
    // -----------------------------------------------------------------------

    #[test]
    fn test_precedence_mul_over_add() {
        // "1+2*3" should parse as Add(1, Mul(2, 3))
        let result = parse_formula("1+2*3").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Add,
                left: Box::new(Expr::Number(1.0)),
                right: Box::new(Expr::BinaryOp {
                    op: BinaryOperator::Mul,
                    left: Box::new(Expr::Number(2.0)),
                    right: Box::new(Expr::Number(3.0)),
                }),
            }
        );
    }

    #[test]
    fn test_precedence_parens() {
        // "(1+2)*3" should parse as Mul(Paren(Add(1, 2)), 3)
        let result = parse_formula("(1+2)*3").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Mul,
                left: Box::new(Expr::Paren(Box::new(Expr::BinaryOp {
                    op: BinaryOperator::Add,
                    left: Box::new(Expr::Number(1.0)),
                    right: Box::new(Expr::Number(2.0)),
                }))),
                right: Box::new(Expr::Number(3.0)),
            }
        );
    }

    // -----------------------------------------------------------------------
    // Function calls
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_function_no_args() {
        let result = parse_formula("NOW()").unwrap();
        assert_eq!(
            result,
            Expr::Function {
                name: "NOW".to_string(),
                args: vec![],
            }
        );
    }

    #[test]
    fn test_parse_function_one_arg() {
        let result = parse_formula("ABS(-5)").unwrap();
        assert_eq!(
            result,
            Expr::Function {
                name: "ABS".to_string(),
                args: vec![Expr::UnaryOp {
                    op: UnaryOperator::Neg,
                    operand: Box::new(Expr::Number(5.0)),
                }],
            }
        );
    }

    #[test]
    fn test_parse_function_multi_args() {
        let result = parse_formula("SUM(1,2,3)").unwrap();
        assert_eq!(
            result,
            Expr::Function {
                name: "SUM".to_string(),
                args: vec![Expr::Number(1.0), Expr::Number(2.0), Expr::Number(3.0),],
            }
        );
    }

    #[test]
    fn test_parse_nested_function() {
        let result = parse_formula("SUM(A1:A10,MAX(B1:B10))").unwrap();
        assert_eq!(
            result,
            Expr::Function {
                name: "SUM".to_string(),
                args: vec![
                    Expr::Range {
                        start: CellReference {
                            col: "A".to_string(),
                            row: 1,
                            abs_col: false,
                            abs_row: false,
                            sheet: None,
                        },
                        end: CellReference {
                            col: "A".to_string(),
                            row: 10,
                            abs_col: false,
                            abs_row: false,
                            sheet: None,
                        },
                    },
                    Expr::Function {
                        name: "MAX".to_string(),
                        args: vec![Expr::Range {
                            start: CellReference {
                                col: "B".to_string(),
                                row: 1,
                                abs_col: false,
                                abs_row: false,
                                sheet: None,
                            },
                            end: CellReference {
                                col: "B".to_string(),
                                row: 10,
                                abs_col: false,
                                abs_row: false,
                                sheet: None,
                            },
                        }],
                    },
                ],
            }
        );
    }

    // -----------------------------------------------------------------------
    // Comparison operators
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_equal() {
        let result = parse_formula("A1=B1").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Eq,
                left: Box::new(Expr::CellRef(CellReference {
                    col: "A".to_string(),
                    row: 1,
                    abs_col: false,
                    abs_row: false,
                    sheet: None,
                })),
                right: Box::new(Expr::CellRef(CellReference {
                    col: "B".to_string(),
                    row: 1,
                    abs_col: false,
                    abs_row: false,
                    sheet: None,
                })),
            }
        );
    }

    #[test]
    fn test_parse_not_equal() {
        let result = parse_formula("A1<>B1").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Ne,
                left: Box::new(Expr::CellRef(CellReference {
                    col: "A".to_string(),
                    row: 1,
                    abs_col: false,
                    abs_row: false,
                    sheet: None,
                })),
                right: Box::new(Expr::CellRef(CellReference {
                    col: "B".to_string(),
                    row: 1,
                    abs_col: false,
                    abs_row: false,
                    sheet: None,
                })),
            }
        );
    }

    #[test]
    fn test_parse_less_than() {
        let result = parse_formula("A1<B1").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Lt,
                left: Box::new(Expr::CellRef(CellReference {
                    col: "A".to_string(),
                    row: 1,
                    abs_col: false,
                    abs_row: false,
                    sheet: None,
                })),
                right: Box::new(Expr::CellRef(CellReference {
                    col: "B".to_string(),
                    row: 1,
                    abs_col: false,
                    abs_row: false,
                    sheet: None,
                })),
            }
        );
    }

    // -----------------------------------------------------------------------
    // Concatenation
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_concat() {
        let result = parse_formula("A1&B1").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Concat,
                left: Box::new(Expr::CellRef(CellReference {
                    col: "A".to_string(),
                    row: 1,
                    abs_col: false,
                    abs_row: false,
                    sheet: None,
                })),
                right: Box::new(Expr::CellRef(CellReference {
                    col: "B".to_string(),
                    row: 1,
                    abs_col: false,
                    abs_row: false,
                    sheet: None,
                })),
            }
        );
    }

    // -----------------------------------------------------------------------
    // Complex expressions
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_complex_formula() {
        // "SUM(A1:A10)+AVERAGE(B1:B10)*2"
        let result = parse_formula("SUM(A1:A10)+AVERAGE(B1:B10)*2").unwrap();
        assert_eq!(
            result,
            Expr::BinaryOp {
                op: BinaryOperator::Add,
                left: Box::new(Expr::Function {
                    name: "SUM".to_string(),
                    args: vec![Expr::Range {
                        start: CellReference {
                            col: "A".to_string(),
                            row: 1,
                            abs_col: false,
                            abs_row: false,
                            sheet: None,
                        },
                        end: CellReference {
                            col: "A".to_string(),
                            row: 10,
                            abs_col: false,
                            abs_row: false,
                            sheet: None,
                        },
                    }],
                }),
                right: Box::new(Expr::BinaryOp {
                    op: BinaryOperator::Mul,
                    left: Box::new(Expr::Function {
                        name: "AVERAGE".to_string(),
                        args: vec![Expr::Range {
                            start: CellReference {
                                col: "B".to_string(),
                                row: 1,
                                abs_col: false,
                                abs_row: false,
                                sheet: None,
                            },
                            end: CellReference {
                                col: "B".to_string(),
                                row: 10,
                                abs_col: false,
                                abs_row: false,
                                sheet: None,
                            },
                        }],
                    }),
                    right: Box::new(Expr::Number(2.0)),
                }),
            }
        );
    }

    // -----------------------------------------------------------------------
    // Sheet reference
    // -----------------------------------------------------------------------

    #[test]
    fn test_parse_sheet_ref() {
        let result = parse_formula("Sheet1!A1").unwrap();
        assert_eq!(
            result,
            Expr::CellRef(CellReference {
                col: "A".to_string(),
                row: 1,
                abs_col: false,
                abs_row: false,
                sheet: Some("Sheet1".to_string()),
            })
        );
    }
}
