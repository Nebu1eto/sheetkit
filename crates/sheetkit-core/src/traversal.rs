//! Shared finite limits for traversing worksheet-shaped data.

use crate::error::{Error, Result};
use crate::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};

/// Default maximum number of rows visited by one traversal.
pub const DEFAULT_MAX_ROWS: u64 = 1_048_576;
/// Default maximum number of columns visited by one traversal.
pub const DEFAULT_MAX_COLUMNS: u64 = 16_384;
/// Default maximum number of cells visited by one traversal.
pub const DEFAULT_MAX_CELLS: u64 = 10_000_000;
/// Default maximum number of bytes produced by one traversal.
pub const DEFAULT_MAX_OUTPUT_BYTES: u64 = 512 * 1024 * 1024;

/// Non-overridable maximum number of rows visited by one traversal.
pub const HARD_MAX_ROWS: u64 = 1_048_576;
/// Non-overridable maximum number of columns visited by one traversal.
pub const HARD_MAX_COLUMNS: u64 = 16_384;
/// Non-overridable maximum number of cells visited by one traversal.
pub const HARD_MAX_CELLS: u64 = 100_000_000;
/// Non-overridable maximum number of output bytes produced by one traversal.
pub const HARD_MAX_OUTPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;

/// Traversal layout policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TraversalMode {
    /// Visit and encode only occupied rows and cells.
    #[default]
    Sparse,
    /// Visit the complete bounding rectangle, including empty positions.
    Dense,
}

/// Finite resource limits shared by worksheet traversal operations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TraversalLimits {
    max_rows: u64,
    max_columns: u64,
    max_cells: u64,
    max_output_bytes: u64,
}

impl Default for TraversalLimits {
    fn default() -> Self {
        Self {
            max_rows: DEFAULT_MAX_ROWS,
            max_columns: DEFAULT_MAX_COLUMNS,
            max_cells: DEFAULT_MAX_CELLS,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
        }
    }
}

impl TraversalLimits {
    /// Creates limits, rejecting zero values and values above hard ceilings.
    pub fn new(
        max_rows: u64,
        max_columns: u64,
        max_cells: u64,
        max_output_bytes: u64,
    ) -> Result<Self> {
        validate_config(ResourceDimension::RowCount, max_rows, HARD_MAX_ROWS)?;
        validate_config(
            ResourceDimension::ColumnCount,
            max_columns,
            HARD_MAX_COLUMNS,
        )?;
        validate_config(ResourceDimension::CellCount, max_cells, HARD_MAX_CELLS)?;
        validate_config(
            ResourceDimension::OutputBytes,
            max_output_bytes,
            HARD_MAX_OUTPUT_BYTES,
        )?;
        Ok(Self {
            max_rows,
            max_columns,
            max_cells,
            max_output_bytes,
        })
    }

    /// Returns the maximum row count.
    pub const fn max_rows(self) -> u64 {
        self.max_rows
    }

    /// Returns the maximum column count.
    pub const fn max_columns(self) -> u64 {
        self.max_columns
    }

    /// Returns the maximum cell count.
    pub const fn max_cells(self) -> u64 {
        self.max_cells
    }

    /// Returns the maximum output byte count.
    pub const fn max_output_bytes(self) -> u64 {
        self.max_output_bytes
    }
}

/// Options controlling one worksheet traversal.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct TraversalOptions {
    mode: TraversalMode,
    limits: TraversalLimits,
}

impl TraversalOptions {
    /// Creates options for a traversal mode using default finite limits.
    pub const fn new(mode: TraversalMode) -> Self {
        Self {
            mode,
            limits: TraversalLimits {
                max_rows: DEFAULT_MAX_ROWS,
                max_columns: DEFAULT_MAX_COLUMNS,
                max_cells: DEFAULT_MAX_CELLS,
                max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            },
        }
    }

    /// Replaces the finite traversal limits.
    pub const fn with_limits(mut self, limits: TraversalLimits) -> Self {
        self.limits = limits;
        self
    }

    /// Returns the selected traversal mode.
    pub const fn mode(self) -> TraversalMode {
        self.mode
    }

    /// Returns the finite traversal limits.
    pub const fn limits(self) -> TraversalLimits {
        self.limits
    }
}

/// A validated inclusive worksheet rectangle.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CheckedRect {
    min_row: u32,
    max_row: u32,
    min_col: u32,
    max_col: u32,
    row_count: u64,
    column_count: u64,
    cell_count: u64,
}

impl CheckedRect {
    /// Validates coordinates and checked dimensions against finite limits.
    pub fn new(
        min_row: u32,
        max_row: u32,
        min_col: u32,
        max_col: u32,
        limits: TraversalLimits,
    ) -> Result<Self> {
        if min_row == 0 || max_row < min_row {
            return Err(Error::InvalidRowNumber(min_row));
        }
        if min_col == 0 || max_col < min_col {
            return Err(Error::InvalidColumnNumber(min_col));
        }
        check_limit(
            ResourceDimension::RowCount,
            HARD_MAX_ROWS,
            u64::from(max_row),
        )?;
        check_limit(
            ResourceDimension::ColumnCount,
            HARD_MAX_COLUMNS,
            u64::from(max_col),
        )?;
        let row_count = u64::from(max_row)
            .checked_sub(u64::from(min_row))
            .and_then(|span| span.checked_add(1))
            .ok_or_else(|| limit_error(ResourceDimension::RowCount, limits.max_rows, None))?;
        let column_count = u64::from(max_col)
            .checked_sub(u64::from(min_col))
            .and_then(|span| span.checked_add(1))
            .ok_or_else(|| limit_error(ResourceDimension::ColumnCount, limits.max_columns, None))?;
        check_limit(ResourceDimension::RowCount, limits.max_rows, row_count)?;
        check_limit(
            ResourceDimension::ColumnCount,
            limits.max_columns,
            column_count,
        )?;
        let cell_count = row_count
            .checked_mul(column_count)
            .ok_or_else(|| limit_error(ResourceDimension::CellCount, limits.max_cells, None))?;
        check_limit(ResourceDimension::CellCount, limits.max_cells, cell_count)?;
        Ok(Self {
            min_row,
            max_row,
            min_col,
            max_col,
            row_count,
            column_count,
            cell_count,
        })
    }

    /// Validates sparse bounds without charging the unvisited rectangle area.
    pub fn new_sparse(
        min_row: u32,
        max_row: u32,
        min_col: u32,
        max_col: u32,
        limits: TraversalLimits,
    ) -> Result<Self> {
        if min_row == 0 || max_row < min_row {
            return Err(Error::InvalidRowNumber(min_row));
        }
        if min_col == 0 || max_col < min_col {
            return Err(Error::InvalidColumnNumber(min_col));
        }
        check_limit(
            ResourceDimension::RowCount,
            HARD_MAX_ROWS,
            u64::from(max_row),
        )?;
        check_limit(
            ResourceDimension::ColumnCount,
            HARD_MAX_COLUMNS,
            u64::from(max_col),
        )?;
        let row_count = u64::from(max_row - min_row) + 1;
        let column_count = u64::from(max_col - min_col) + 1;
        check_limit(
            ResourceDimension::ColumnCount,
            limits.max_columns,
            column_count,
        )?;
        let cell_count = row_count
            .checked_mul(column_count)
            .ok_or_else(|| limit_error(ResourceDimension::CellCount, limits.max_cells, None))?;
        Ok(Self {
            min_row,
            max_row,
            min_col,
            max_col,
            row_count,
            column_count,
            cell_count,
        })
    }

    pub const fn min_row(self) -> u32 {
        self.min_row
    }

    pub const fn max_row(self) -> u32 {
        self.max_row
    }

    pub const fn min_col(self) -> u32 {
        self.min_col
    }

    pub const fn max_col(self) -> u32 {
        self.max_col
    }

    pub const fn row_count(self) -> u64 {
        self.row_count
    }

    pub const fn column_count(self) -> u64 {
        self.column_count
    }

    pub const fn cell_count(self) -> u64 {
        self.cell_count
    }
}

/// Adds byte counts without wrapping and checks the output budget.
pub fn checked_output_add(current: u64, additional: u64, limits: TraversalLimits) -> Result<u64> {
    let observed = current.checked_add(additional).ok_or_else(|| {
        limit_error(
            ResourceDimension::OutputBytes,
            limits.max_output_bytes,
            None,
        )
    })?;
    check_limit(
        ResourceDimension::OutputBytes,
        limits.max_output_bytes,
        observed,
    )?;
    Ok(observed)
}

/// Multiplies byte counts without wrapping and checks the output budget.
pub fn checked_output_mul(left: u64, right: u64, limits: TraversalLimits) -> Result<u64> {
    let observed = left.checked_mul(right).ok_or_else(|| {
        limit_error(
            ResourceDimension::OutputBytes,
            limits.max_output_bytes,
            None,
        )
    })?;
    check_limit(
        ResourceDimension::OutputBytes,
        limits.max_output_bytes,
        observed,
    )?;
    Ok(observed)
}

/// Checks a traversal count against its finite limit.
pub fn check_limit(dimension: ResourceDimension, limit: u64, observed: u64) -> Result<()> {
    if observed > limit {
        return Err(limit_error(dimension, limit, Some(observed)));
    }
    Ok(())
}

fn validate_config(dimension: ResourceDimension, value: u64, hard_limit: u64) -> Result<()> {
    if value == 0 {
        return Err(limit_error(dimension, 0, Some(1)));
    }
    check_limit(dimension, hard_limit, value)
}

fn limit_error(dimension: ResourceDimension, limit: u64, observed: Option<u64>) -> Error {
    let observed = observed.unwrap_or_else(|| limit.saturating_add(1));
    match ResourceLimitError::exceeded(
        ResourceOperation::DataTraverse,
        ResourceCategory::Traversal,
        dimension,
        limit,
        observed.max(limit.saturating_add(1)),
        Some(ResourceStage::Serialize),
    ) {
        Some(metadata) => metadata.into(),
        None => Error::Internal("invalid traversal resource limit".to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn metadata(error: Error) -> ResourceLimitError {
        match error {
            Error::ResourceLimitExceeded(metadata) => metadata,
            other => panic!("expected resource limit error, got {other}"),
        }
    }

    #[test]
    fn defaults_are_finite_and_sparse() {
        let options = TraversalOptions::default();
        assert_eq!(options.mode(), TraversalMode::Sparse);
        assert!(options.limits().max_rows() < u64::MAX);
        assert!(options.limits().max_columns() < u64::MAX);
        assert!(options.limits().max_cells() < u64::MAX);
        assert!(options.limits().max_output_bytes() < u64::MAX);
    }

    #[test]
    fn configuration_accepts_hard_ceiling_and_rejects_plus_one() {
        TraversalLimits::new(
            HARD_MAX_ROWS,
            HARD_MAX_COLUMNS,
            HARD_MAX_CELLS,
            HARD_MAX_OUTPUT_BYTES,
        )
        .unwrap();
        let error = TraversalLimits::new(
            HARD_MAX_ROWS,
            HARD_MAX_COLUMNS,
            HARD_MAX_CELLS + 1,
            HARD_MAX_OUTPUT_BYTES,
        )
        .unwrap_err();
        let metadata = metadata(error);
        assert_eq!(metadata.dimension(), ResourceDimension::CellCount);
        assert_eq!(metadata.limit(), HARD_MAX_CELLS);
        assert_eq!(metadata.observed(), HARD_MAX_CELLS + 1);
    }

    #[test]
    fn checked_rectangle_accepts_exact_limit_and_rejects_plus_one() {
        let limits = TraversalLimits::new(2, 3, 6, 1024).unwrap();
        let rect = CheckedRect::new(1, 2, 1, 3, limits).unwrap();
        assert_eq!(rect.row_count(), 2);
        assert_eq!(rect.column_count(), 3);
        assert_eq!(rect.cell_count(), 6);

        let error = CheckedRect::new(1, 3, 1, 3, limits).unwrap_err();
        assert_eq!(metadata(error).dimension(), ResourceDimension::RowCount);
    }

    #[test]
    fn checked_rectangle_reports_each_dimension_boundary() {
        let limits = TraversalLimits::new(2, 3, 6, 1024).unwrap();
        let row_error = CheckedRect::new(1, 3, 1, 1, limits).unwrap_err();
        assert_eq!(metadata(row_error).dimension(), ResourceDimension::RowCount);

        let column_error = CheckedRect::new(1, 1, 1, 4, limits).unwrap_err();
        assert_eq!(
            metadata(column_error).dimension(),
            ResourceDimension::ColumnCount
        );

        let cell_limits = TraversalLimits::new(1, 7, 6, 1024).unwrap();
        let cell_error = CheckedRect::new(1, 1, 1, 7, cell_limits).unwrap_err();
        let cell_metadata = metadata(cell_error);
        assert_eq!(cell_metadata.dimension(), ResourceDimension::CellCount);
        assert_eq!(cell_metadata.limit(), 6);
        assert_eq!(cell_metadata.observed(), 7);
    }

    #[test]
    fn checked_output_arithmetic_reports_overflow_as_bounded_error() {
        let limits = TraversalLimits::default();
        let error = checked_output_add(u64::MAX, 1, limits).unwrap_err();
        let error_metadata = metadata(error);
        assert_eq!(error_metadata.dimension(), ResourceDimension::OutputBytes);
        assert_eq!(error_metadata.limit(), limits.max_output_bytes());
        assert_eq!(error_metadata.observed(), limits.max_output_bytes() + 1);

        let error = checked_output_mul(u64::MAX, 2, limits).unwrap_err();
        assert_eq!(metadata(error).dimension(), ResourceDimension::OutputBytes);
    }

    #[test]
    fn rectangles_reject_coordinates_outside_the_hard_grid() {
        let limits = TraversalLimits::default();
        let row_error = CheckedRect::new(
            HARD_MAX_ROWS as u32 + 1,
            HARD_MAX_ROWS as u32 + 1,
            1,
            1,
            limits,
        )
        .unwrap_err();
        let row_metadata = metadata(row_error);
        assert_eq!(row_metadata.dimension(), ResourceDimension::RowCount);
        assert_eq!(row_metadata.limit(), HARD_MAX_ROWS);

        let column_error = CheckedRect::new_sparse(
            1,
            1,
            HARD_MAX_COLUMNS as u32 + 1,
            HARD_MAX_COLUMNS as u32 + 1,
            limits,
        )
        .unwrap_err();
        let column_metadata = metadata(column_error);
        assert_eq!(column_metadata.dimension(), ResourceDimension::ColumnCount);
        assert_eq!(column_metadata.limit(), HARD_MAX_COLUMNS);
    }
}
