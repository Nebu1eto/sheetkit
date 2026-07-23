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
    character::complete::{alpha1, char, multispace0, one_of},
    combinator::{map, opt, recognize, value},
    multi::many0,
    sequence::{delimited, pair, preceded, terminated, tuple},
    IResult,
};

use super::ast::{BinaryOperator, CellReference, Expr, UnaryOperator};
use super::budget::{FormulaBudget, FormulaLimits};
use crate::error::{Error, Result};

/// Parse an Excel formula string into an AST expression.
///
/// The input should NOT include the leading `=` sign.
///
/// # Errors
///
/// Returns an error if the formula string cannot be parsed.
pub fn parse_formula(input: &str) -> Result<Expr> {
    let budget = FormulaBudget::for_parse(FormulaLimits::secure_default())?;
    parse_formula_with_budget(input, &budget)
}

/// Parse a formula while charging an existing operation budget.
pub(crate) fn parse_formula_with_budget(input: &str, budget: &FormulaBudget) -> Result<Expr> {
    let input_bytes = input.len() as u64;
    budget.check_input(input_bytes)?;
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return Err(Error::Internal("empty formula".to_string()));
    }
    let (tokens, parenthesis_depth, unary_depth, binary_operators) = lexical_metrics(trimmed);
    budget.check_nesting(parenthesis_depth.saturating_add(1))?;
    budget.check_nesting(unary_depth.saturating_add(1))?;
    budget.check_binary_chain(binary_operators)?;
    budget.charge_lexical(input_bytes, tokens)?;

    match parse_expr(trimmed) {
        Ok(("", expr)) => {
            let (nodes, depth, binary_chain) = expr.structural_metrics();
            budget.charge_ast(nodes, depth, binary_chain)?;
            Ok(expr)
        }
        Ok((remaining, _)) => Err(Error::Internal(format!(
            "unexpected trailing input: {remaining}"
        ))),
        Err(e) => Err(Error::Internal(format!("formula parse error: {e}"))),
    }
}

fn lexical_metrics(input: &str) -> (u64, u64, u64, u64) {
    let mut tokens = 0_u64;
    let mut parenthesis_depth = 0_u64;
    let mut max_parenthesis_depth = 0_u64;
    let mut unary_run = 0_u64;
    let mut max_unary_run = 0_u64;
    let mut max_binary_chain = 0_u64;
    let mut binary_chains = [0_u64; 129];
    let mut expects_operand = [true; 129];
    let mut frame = 0_usize;
    let mut chars = input.char_indices().peekable();

    while let Some((_, ch)) = chars.next() {
        if ch.is_whitespace() {
            continue;
        }
        tokens = tokens.saturating_add(1);
        match ch {
            '"' => {
                while let Some((_, current)) = chars.next() {
                    if current == '"' {
                        if chars.peek().is_some_and(|(_, next)| *next == '"') {
                            chars.next();
                        } else {
                            break;
                        }
                    }
                }
                unary_run = 0;
                expects_operand[frame] = false;
            }
            '\'' => {
                while let Some((_, current)) = chars.next() {
                    if current == '\'' {
                        if chars.peek().is_some_and(|(_, next)| *next == '\'') {
                            chars.next();
                        } else {
                            break;
                        }
                    }
                }
                if chars.peek().is_some_and(|(_, next)| *next == '!') {
                    chars.next();
                }
                unary_run = 0;
                expects_operand[frame] = false;
            }
            '(' => {
                parenthesis_depth = parenthesis_depth.saturating_add(1);
                max_parenthesis_depth = max_parenthesis_depth.max(parenthesis_depth);
                if frame < 128 {
                    frame += 1;
                    binary_chains[frame] = 0;
                    expects_operand[frame] = true;
                }
                unary_run = 0;
            }
            ')' => {
                parenthesis_depth = parenthesis_depth.saturating_sub(1);
                max_binary_chain = max_binary_chain.max(binary_chains[frame]);
                frame = frame.saturating_sub(1);
                expects_operand[frame] = false;
                unary_run = 0;
            }
            '+' | '-' => {
                if expects_operand[frame] {
                    unary_run = unary_run.saturating_add(1);
                    max_unary_run = max_unary_run.max(unary_run);
                } else {
                    binary_chains[frame] = binary_chains[frame].saturating_add(1);
                    max_binary_chain = max_binary_chain.max(binary_chains[frame]);
                    expects_operand[frame] = true;
                    unary_run = 0;
                }
            }
            '#' => {
                while chars.peek().is_some_and(|(_, next)| {
                    next.is_alphanumeric() || matches!(*next, '#' | '/' | '!' | '?')
                }) {
                    chars.next();
                }
                unary_run = 0;
                expects_operand[frame] = false;
            }
            c if c.is_alphanumeric() || matches!(c, '_' | '$' | '.') => {
                while chars.peek().is_some_and(|(_, next)| {
                    next.is_alphanumeric()
                        || matches!(*next, '_' | '$' | '.' | '\'' | '!' | '?' | '#')
                }) {
                    chars.next();
                }
                unary_run = 0;
                expects_operand[frame] = false;
            }
            '<' | '>' => {
                binary_chains[frame] = binary_chains[frame].saturating_add(1);
                max_binary_chain = max_binary_chain.max(binary_chains[frame]);
                if chars
                    .peek()
                    .is_some_and(|(_, next)| matches!(*next, '=' | '>'))
                {
                    chars.next();
                }
                unary_run = 0;
                expects_operand[frame] = true;
            }
            '*' | '/' | '^' | '&' | '=' => {
                binary_chains[frame] = binary_chains[frame].saturating_add(1);
                max_binary_chain = max_binary_chain.max(binary_chains[frame]);
                unary_run = 0;
                expects_operand[frame] = true;
            }
            ',' => {
                max_binary_chain = max_binary_chain.max(binary_chains[frame]);
                binary_chains[frame] = 0;
                expects_operand[frame] = true;
                unary_run = 0;
            }
            ':' => {
                expects_operand[frame] = true;
                unary_run = 0;
            }
            '%' => {
                expects_operand[frame] = false;
                unary_run = unary_run.saturating_add(1);
                max_unary_run = max_unary_run.max(unary_run);
            }
            _ => unary_run = 0,
        }
    }
    max_binary_chain = max_binary_chain.max(binary_chains[frame]);

    (
        tokens,
        max_parenthesis_depth,
        max_unary_run,
        max_binary_chain,
    )
}

/// Wraps a parser to consume optional surrounding whitespace.
fn ws<'a, F, O>(inner: F) -> impl FnMut(&'a str) -> IResult<&'a str, O>
where
    F: FnMut(&'a str) -> IResult<&'a str, O>,
{
    delimited(multispace0, inner, multispace0)
}

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
        parse_function_call,
        parse_bool_literal,
        parse_cell_ref_or_range,
        parse_number_literal,
    ))(input)
}

/// Parse a numeric literal (integer, decimal, or scientific notation).
fn parse_number_literal(input: &str) -> IResult<&str, Expr> {
    let (input, num_str) = recognize(pair(
        take_while1(|c: char| c.is_ascii_digit()),
        pair(
            opt(pair(tag("."), take_while1(|c: char| c.is_ascii_digit()))),
            opt(tuple((
                one_of("eE"),
                opt(one_of("+-")),
                take_while1(|c: char| c.is_ascii_digit()),
            ))),
        ),
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
    let (mut input, _) = multispace0(input)?;
    if let Ok((remaining, _)) = char::<&str, nom::error::Error<&str>>(')')(input) {
        return Ok((
            remaining,
            Expr::Function {
                name: name.to_uppercase(),
                args: vec![],
            },
        ));
    }

    let mut args = Vec::new();
    let mut expecting_arg = true;
    loop {
        input = input.trim_start();
        if let Ok((remaining, _)) = char::<&str, nom::error::Error<&str>>(')')(input) {
            if expecting_arg {
                args.push(Expr::Missing);
            }
            input = remaining;
            break;
        }

        if let Ok((remaining, _)) = char::<&str, nom::error::Error<&str>>(',')(input) {
            args.push(Expr::Missing);
            input = remaining;
            expecting_arg = true;
            continue;
        }

        let (remaining, arg) = parse_expr(input)?;
        args.push(arg);
        input = remaining.trim_start();
        if let Ok((remaining, _)) = char::<&str, nom::error::Error<&str>>(',')(input) {
            input = remaining;
            expecting_arg = true;
        } else {
            let (remaining, _) = char(')')(input)?;
            input = remaining;
            break;
        }
    }
    Ok((
        input,
        Expr::Function {
            name: name.to_uppercase(),
            args,
        },
    ))
}

/// Parse a parenthesized expression.
fn parse_paren_expr(input: &str) -> IResult<&str, Expr> {
    let (input, _) = char('(')(input)?;
    let (input, _) = multispace0(input)?;
    let (input, expr) = parse_expr(input)?;
    let (input, _) = preceded(multispace0, char(')'))(input)?;
    Ok((input, Expr::Paren(Box::new(expr))))
}

/// Fold a left-associative chain of binary operations into a single AST node.
fn fold_binary(first: Expr, rest: Vec<(BinaryOperator, Expr)>) -> Expr {
    rest.into_iter()
        .fold(first, |left, (op, right)| Expr::BinaryOp {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::ast::{BinaryOperator, CellReference, Expr, UnaryOperator};

    #[test]
    fn formula_limits_reject_deep_recursive_inputs_before_parsing() {
        let exact = format!("{}1{}", "(".repeat(63), ")".repeat(63));
        assert!(parse_formula(&exact).is_ok());

        let over = format!("{}1{}", "(".repeat(64), ")".repeat(64));
        let error = parse_formula(&over).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension_code(),
            "nesting_depth"
        );

        let unary_over = format!("{}1", "-".repeat(64));
        let error = parse_formula(&unary_over).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension_code(),
            "nesting_depth"
        );

        let postfix_exact = format!("1{}", "%".repeat(63));
        assert!(parse_formula(&postfix_exact).is_ok());
        let postfix_over = format!("1{}", "%".repeat(64));
        assert_eq!(
            parse_formula(&postfix_over)
                .unwrap_err()
                .resource_metadata()
                .unwrap()
                .dimension_code(),
            "nesting_depth"
        );

        let nested = |depth: usize| {
            let mut formula = "'Odd)(Sheet'!A1".to_string();
            for _ in 0..depth {
                formula = format!("SUM({formula})");
            }
            formula
        };
        assert!(parse_formula(&nested(63)).is_ok());
        assert_eq!(
            parse_formula(&nested(64))
                .unwrap_err()
                .resource_metadata()
                .unwrap()
                .dimension_code(),
            "nesting_depth"
        );
    }

    #[test]
    fn formula_input_and_binary_chain_boundaries_are_exact() {
        let exact_input = format!("\"{}\"", "a".repeat(32_766));
        assert!(parse_formula(&exact_input).is_ok());
        let over_input = format!("\"{}\"", "a".repeat(32_767));
        let error = parse_formula(&over_input).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension_code(),
            "input_bytes"
        );
        let padded = format!("{}1", " ".repeat(32_768));
        assert_eq!(
            parse_formula(&padded)
                .unwrap_err()
                .resource_metadata()
                .unwrap()
                .dimension_code(),
            "input_bytes"
        );

        let exact_chain = std::iter::repeat_n("1", 513).collect::<Vec<_>>().join("+");
        assert!(parse_formula(&exact_chain).is_ok());
        let over_chain = std::iter::repeat_n("1", 514).collect::<Vec<_>>().join("+");
        let error = parse_formula(&over_chain).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension_code(),
            "nesting_depth"
        );

        let independent_args = std::iter::repeat_n("1+1", 513)
            .collect::<Vec<_>>()
            .join(",");
        assert!(parse_formula(&format!("F({independent_args})")).is_ok());
    }

    #[test]
    fn token_and_ast_node_boundaries_are_exact() {
        let exact = FormulaLimits {
            per_formula_tokens: 3,
            per_formula_ast_nodes: 3,
            operation_tokens: 3,
            operation_ast_nodes: 3,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_parse(exact).unwrap();
        assert!(parse_formula_with_budget("1+1", &budget).is_ok());

        let token_over = FormulaLimits {
            per_formula_tokens: 2,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_parse(token_over).unwrap();
        let error = parse_formula_with_budget("1+1", &budget).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension_code(),
            "token_count"
        );
        assert_eq!(error.resource_metadata().unwrap().observed(), 3);

        let node_over = FormulaLimits {
            per_formula_ast_nodes: 2,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_parse(node_over).unwrap();
        let error = parse_formula_with_budget("1+1", &budget).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension_code(),
            "node_count"
        );
        assert_eq!(error.resource_metadata().unwrap().observed(), 3);

        let independent = FormulaLimits {
            per_formula_nesting_depth: 1,
            per_formula_binary_chain_depth: 1,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_parse(independent).unwrap();
        assert!(parse_formula_with_budget("1+2", &budget).is_ok());

        let chain_over = FormulaLimits {
            per_formula_binary_chain_depth: 0,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_parse(chain_over).unwrap();
        let error = parse_formula_with_budget("1+2", &budget).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "nesting_depth");
        assert_eq!((metadata.limit(), metadata.observed()), (0, 1));
    }

    #[test]
    fn shared_parse_budget_cannot_be_reset_between_calls() {
        let limits = FormulaLimits {
            operation_input_bytes: 3,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_parse(limits).unwrap();
        assert!(parse_formula_with_budget("1", &budget).is_ok());
        assert!(parse_formula_with_budget("22", &budget).is_ok());
        let error = parse_formula_with_budget("3", &budget).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.dimension_code(), "input_bytes");
        assert_eq!((metadata.limit(), metadata.observed()), (3, 4));
    }

    #[test]
    fn test_parse_integer() {
        let result = parse_formula("42").unwrap();
        assert_eq!(result, Expr::Number(42.0));
    }

    #[test]
    fn test_parse_decimal() {
        let result = parse_formula("3.15").unwrap();
        assert_eq!(result, Expr::Number(3.15));
    }

    #[test]
    fn test_parse_scientific_number_literals() {
        for (formula, expected) in [
            ("1E5", 100_000.0),
            ("1e5", 100_000.0),
            ("1.5E+2", 150.0),
            ("2E-3", 0.002),
        ] {
            assert_eq!(parse_formula(formula).unwrap(), Expr::Number(expected));
        }
    }

    #[test]
    fn test_reject_incomplete_scientific_number_literals() {
        for formula in ["1E", "1E+", "1E-"] {
            assert!(parse_formula(formula).is_err(), "{formula} should fail");
        }
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
    fn test_parse_boolean_function_no_args() {
        assert_eq!(
            parse_formula("TRUE()").unwrap(),
            Expr::Function {
                name: "TRUE".to_string(),
                args: vec![],
            }
        );
        assert_eq!(
            parse_formula("FALSE()").unwrap(),
            Expr::Function {
                name: "FALSE".to_string(),
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
    fn test_parse_omitted_function_args() {
        for (formula, args) in [
            (
                "IF(A1,,1)",
                vec![
                    Expr::CellRef(CellReference {
                        col: "A".to_string(),
                        row: 1,
                        abs_col: false,
                        abs_row: false,
                        sheet: None,
                    }),
                    Expr::Missing,
                    Expr::Number(1.0),
                ],
            ),
            (
                "IF(,1,2)",
                vec![Expr::Missing, Expr::Number(1.0), Expr::Number(2.0)],
            ),
            (
                "IF(A1,1,)",
                vec![
                    Expr::CellRef(CellReference {
                        col: "A".to_string(),
                        row: 1,
                        abs_col: false,
                        abs_row: false,
                        sheet: None,
                    }),
                    Expr::Number(1.0),
                    Expr::Missing,
                ],
            ),
        ] {
            assert_eq!(
                parse_formula(formula).unwrap(),
                Expr::Function {
                    name: "IF".to_string(),
                    args,
                }
            );
        }
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
