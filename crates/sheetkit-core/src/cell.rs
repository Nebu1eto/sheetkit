//! Cell value representation.
//!
//! Provides the [`CellValue`] enum which represents the typed value of a
//! single cell in a worksheet. This is the high-level counterpart to the
//! raw XML `Cell` element from `sheetkit-xml`.

use std::fmt;

/// Represents the value of a cell.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// No value (empty cell).
    Empty,
    /// Boolean value.
    Bool(bool),
    /// Numeric value (integers are stored as f64 in Excel).
    Number(f64),
    /// String value.
    String(String),
    /// Formula with optional cached result.
    Formula {
        expr: String,
        result: Option<Box<CellValue>>,
    },
    /// Error value (e.g. #DIV/0!, #N/A, #VALUE!).
    Error(String),
}

impl Default for CellValue {
    fn default() -> Self {
        Self::Empty
    }
}

impl fmt::Display for CellValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CellValue::Empty => write!(f, ""),
            CellValue::Bool(b) => write!(f, "{}", if *b { "TRUE" } else { "FALSE" }),
            CellValue::Number(n) => {
                // Display integers without decimal point
                if n.fract() == 0.0 && n.is_finite() {
                    write!(f, "{}", *n as i64)
                } else {
                    write!(f, "{n}")
                }
            }
            CellValue::String(s) => write!(f, "{s}"),
            CellValue::Formula { result, expr, .. } => {
                if let Some(result) = result {
                    write!(f, "{result}")
                } else {
                    write!(f, "={expr}")
                }
            }
            CellValue::Error(e) => write!(f, "{e}"),
        }
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        CellValue::String(s.to_string())
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        CellValue::String(s)
    }
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

impl From<i32> for CellValue {
    fn from(n: i32) -> Self {
        CellValue::Number(f64::from(n))
    }
}

impl From<i64> for CellValue {
    fn from(n: i64) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        CellValue::Bool(b)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_value_default_is_empty() {
        let v = CellValue::default();
        assert_eq!(v, CellValue::Empty);
    }

    #[test]
    fn test_cell_value_from_str() {
        let v: CellValue = "hello".into();
        assert_eq!(v, CellValue::String("hello".to_string()));
    }

    #[test]
    fn test_cell_value_from_string() {
        let v: CellValue = String::from("world").into();
        assert_eq!(v, CellValue::String("world".to_string()));
    }

    #[test]
    fn test_cell_value_from_f64() {
        let v: CellValue = 3.14.into();
        assert_eq!(v, CellValue::Number(3.14));
    }

    #[test]
    fn test_cell_value_from_i32() {
        let v: CellValue = 42i32.into();
        assert_eq!(v, CellValue::Number(42.0));
    }

    #[test]
    fn test_cell_value_from_i64() {
        let v: CellValue = 100i64.into();
        assert_eq!(v, CellValue::Number(100.0));
    }

    #[test]
    fn test_cell_value_from_bool() {
        let v: CellValue = true.into();
        assert_eq!(v, CellValue::Bool(true));

        let v2: CellValue = false.into();
        assert_eq!(v2, CellValue::Bool(false));
    }

    #[test]
    fn test_cell_value_display() {
        assert_eq!(CellValue::Empty.to_string(), "");
        assert_eq!(CellValue::Bool(true).to_string(), "TRUE");
        assert_eq!(CellValue::Bool(false).to_string(), "FALSE");
        assert_eq!(CellValue::Number(42.0).to_string(), "42");
        assert_eq!(CellValue::Number(3.14).to_string(), "3.14");
        assert_eq!(CellValue::String("hello".to_string()).to_string(), "hello");
        assert_eq!(
            CellValue::Error("#DIV/0!".to_string()).to_string(),
            "#DIV/0!"
        );
        assert_eq!(
            CellValue::Formula {
                expr: "A1+B1".to_string(),
                result: Some(Box::new(CellValue::Number(42.0))),
            }
            .to_string(),
            "42"
        );
        assert_eq!(
            CellValue::Formula {
                expr: "A1+B1".to_string(),
                result: None,
            }
            .to_string(),
            "=A1+B1"
        );
    }
}
