//! Stable metadata for bounded resource failures.

use std::fmt;

/// The public operation that exhausted a resource budget.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ResourceOperation {
    WorkbookOpen,
    WorkbookSave,
    FormulaParse,
    FormulaMutate,
    FormulaCalculate,
    DataTraverse,
    Render,
    Export,
    Encrypt,
    Decrypt,
    MacroRead,
    StreamRead,
    StreamWrite,
}

impl ResourceOperation {
    /// Returns the stable machine-readable operation code.
    pub const fn code(self) -> &'static str {
        match self {
            Self::WorkbookOpen => "workbook_open",
            Self::WorkbookSave => "workbook_save",
            Self::FormulaParse => "formula_parse",
            Self::FormulaMutate => "formula_mutate",
            Self::FormulaCalculate => "formula_calculate",
            Self::DataTraverse => "data_traverse",
            Self::Render => "render",
            Self::Export => "export",
            Self::Encrypt => "encrypt",
            Self::Decrypt => "decrypt",
            Self::MacroRead => "macro_read",
            Self::StreamRead => "stream_read",
            Self::StreamWrite => "stream_write",
        }
    }
}

/// The subsystem whose resource budget was exhausted.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ResourceCategory {
    Archive,
    Container,
    Xml,
    Formula,
    Traversal,
    Output,
    Encryption,
    Macro,
    Streaming,
    Structure,
}

impl ResourceCategory {
    /// Returns the stable machine-readable category code.
    pub const fn code(self) -> &'static str {
        match self {
            Self::Archive => "archive",
            Self::Container => "container",
            Self::Xml => "xml",
            Self::Formula => "formula",
            Self::Traversal => "traversal",
            Self::Output => "output",
            Self::Encryption => "encryption",
            Self::Macro => "macro",
            Self::Streaming => "streaming",
            Self::Structure => "structure",
        }
    }
}

/// The independently bounded resource dimension that was exhausted.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ResourceDimension {
    InputBytes,
    OutputBytes,
    RetainedBytes,
    EntryCount,
    EntryDeclaredBytes,
    EntryActualBytes,
    AggregateDeclaredBytes,
    AggregateActualBytes,
    CompressionRatio,
    CompressionRatioFloorBytes,
    XmlBytes,
    NestingDepth,
    TokenCount,
    NodeCount,
    RangeCellCount,
    DependencyEdgeCount,
    OperationCount,
    MatchWork,
    RowCount,
    ColumnCount,
    CellCount,
    StreamCount,
    ModuleCount,
    ChunkCount,
    IterationCount,
    CollectionCount,
}

impl ResourceDimension {
    /// Returns the stable machine-readable dimension code.
    pub const fn code(self) -> &'static str {
        match self {
            Self::InputBytes => "input_bytes",
            Self::OutputBytes => "output_bytes",
            Self::RetainedBytes => "retained_bytes",
            Self::EntryCount => "entry_count",
            Self::EntryDeclaredBytes => "entry_declared_bytes",
            Self::EntryActualBytes => "entry_actual_bytes",
            Self::AggregateDeclaredBytes => "aggregate_declared_bytes",
            Self::AggregateActualBytes => "aggregate_actual_bytes",
            Self::CompressionRatio => "compression_ratio",
            Self::CompressionRatioFloorBytes => "compression_ratio_floor_bytes",
            Self::XmlBytes => "xml_bytes",
            Self::NestingDepth => "nesting_depth",
            Self::TokenCount => "token_count",
            Self::NodeCount => "node_count",
            Self::RangeCellCount => "range_cell_count",
            Self::DependencyEdgeCount => "dependency_edge_count",
            Self::OperationCount => "operation_count",
            Self::MatchWork => "match_work",
            Self::RowCount => "row_count",
            Self::ColumnCount => "column_count",
            Self::CellCount => "cell_count",
            Self::StreamCount => "stream_count",
            Self::ModuleCount => "module_count",
            Self::ChunkCount => "chunk_count",
            Self::IterationCount => "iteration_count",
            Self::CollectionCount => "collection_count",
        }
    }
}

/// The optional processing stage at which exhaustion was detected.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ResourceStage {
    Inspect,
    Read,
    Parse,
    Decode,
    Validate,
    Calculate,
    Render,
    Serialize,
    Write,
    Replace,
}

impl ResourceStage {
    /// Returns the stable machine-readable stage code.
    pub const fn code(self) -> &'static str {
        match self {
            Self::Inspect => "inspect",
            Self::Read => "read",
            Self::Parse => "parse",
            Self::Decode => "decode",
            Self::Validate => "validate",
            Self::Calculate => "calculate",
            Self::Render => "render",
            Self::Serialize => "serialize",
            Self::Write => "write",
            Self::Replace => "replace",
        }
    }
}

/// Stable, bounded metadata for a resource exhaustion error.
///
/// Instances can only be created when `observed` is strictly greater than
/// `limit`. All textual fields are selected from closed enums, so error output
/// cannot echo attacker-controlled content.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ResourceLimitError {
    operation: ResourceOperation,
    category: ResourceCategory,
    dimension: ResourceDimension,
    limit: u64,
    observed: u64,
    stage: Option<ResourceStage>,
}

impl ResourceLimitError {
    /// Stable machine-readable code for this error kind.
    pub const CODE: &'static str = "resource_limit_exceeded";

    /// Creates metadata for a value that exceeded a finite resource limit.
    pub const fn exceeded(
        operation: ResourceOperation,
        category: ResourceCategory,
        dimension: ResourceDimension,
        limit: u64,
        observed: u64,
        stage: Option<ResourceStage>,
    ) -> Option<Self> {
        if observed <= limit {
            return None;
        }

        Some(Self {
            operation,
            category,
            dimension,
            limit,
            observed,
            stage,
        })
    }

    pub const fn operation(self) -> ResourceOperation {
        self.operation
    }

    pub const fn category(self) -> ResourceCategory {
        self.category
    }

    pub const fn dimension(self) -> ResourceDimension {
        self.dimension
    }

    pub const fn limit(self) -> u64 {
        self.limit
    }

    pub const fn observed(self) -> u64 {
        self.observed
    }

    pub const fn stage(self) -> Option<ResourceStage> {
        self.stage
    }

    pub const fn operation_code(self) -> &'static str {
        self.operation.code()
    }

    pub const fn category_code(self) -> &'static str {
        self.category.code()
    }

    pub const fn dimension_code(self) -> &'static str {
        self.dimension.code()
    }

    pub const fn stage_code(self) -> Option<&'static str> {
        match self.stage {
            Some(stage) => Some(stage.code()),
            None => None,
        }
    }
}

impl fmt::Display for ResourceLimitError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "resource limit exceeded: operation={} category={} dimension={} limit={} observed={}",
            self.operation.code(),
            self.category.code(),
            self.dimension.code(),
            self.limit,
            self.observed,
        )?;
        if let Some(stage) = self.stage {
            write!(formatter, " stage={}", stage.code())?;
        }
        Ok(())
    }
}

impl std::error::Error for ResourceLimitError {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn metadata_contains_only_closed_codes_and_numbers() {
        let metadata = ResourceLimitError::exceeded(
            ResourceOperation::FormulaCalculate,
            ResourceCategory::Formula,
            ResourceDimension::OperationCount,
            100,
            101,
            Some(ResourceStage::Calculate),
        )
        .unwrap();

        assert_eq!(metadata.operation_code(), "formula_calculate");
        assert_eq!(metadata.category_code(), "formula");
        assert_eq!(metadata.dimension_code(), "operation_count");
        assert_eq!(metadata.stage_code(), Some("calculate"));
        assert_eq!(metadata.limit(), 100);
        assert_eq!(metadata.observed(), 101);
    }

    #[test]
    fn metadata_requires_an_exceeded_limit() {
        assert!(ResourceLimitError::exceeded(
            ResourceOperation::WorkbookOpen,
            ResourceCategory::Archive,
            ResourceDimension::InputBytes,
            u64::MAX,
            u64::MAX,
            None,
        )
        .is_none());
    }

    #[test]
    fn archive_dimensions_distinguish_admission_from_actual_reads() {
        assert_eq!(
            ResourceDimension::EntryDeclaredBytes.code(),
            "entry_declared_bytes"
        );
        assert_eq!(
            ResourceDimension::EntryActualBytes.code(),
            "entry_actual_bytes"
        );
        assert_eq!(
            ResourceDimension::AggregateDeclaredBytes.code(),
            "aggregate_declared_bytes"
        );
        assert_eq!(
            ResourceDimension::AggregateActualBytes.code(),
            "aggregate_actual_bytes"
        );
    }
}
