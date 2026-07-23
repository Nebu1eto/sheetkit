//! Finite resource limits for workbook and streamed-sheet serialization.

use crate::error::{Error, Result};
use crate::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};

const MIB: u64 = 1024 * 1024;
const GIB: u64 = 1024 * MIB;

/// Default maximum compressed bytes emitted by one workbook save.
pub const DEFAULT_SAVE_OUTPUT_BYTES: u64 = 512 * MIB;
/// Default maximum bytes emitted by one serialized XML part.
pub const DEFAULT_SAVE_XML_PART_BYTES: u64 = 128 * MIB;
/// Default maximum occupied rows serialized by one workbook save.
pub const DEFAULT_SAVE_ROWS: u64 = 10_000_000;
/// Default maximum occupied cells serialized by one workbook save.
pub const DEFAULT_SAVE_CELLS: u64 = 10_000_000;

/// Default maximum bytes in one streamed row XML fragment.
pub const DEFAULT_STREAM_ROW_XML_BYTES: u64 = 16 * MIB;
/// Default maximum bytes in one streamed worksheet's row fragments.
pub const DEFAULT_STREAM_OUTPUT_BYTES: u64 = 512 * MIB;
/// Default maximum occupied rows accepted by one stream writer.
pub const DEFAULT_STREAM_ROWS: u64 = 1_048_576;
/// Default maximum cell slots accepted by one stream writer.
pub const DEFAULT_STREAM_CELLS: u64 = 10_000_000;
/// Default maximum distinct stream configuration entries.
pub const DEFAULT_STREAM_CONFIG_COUNT: u64 = 100_000;
/// Default maximum retained bytes used by stream configuration.
pub const DEFAULT_STREAM_CONFIG_BYTES: u64 = 16 * MIB;

const HARD_SAVE_OUTPUT_BYTES: u64 = 2 * GIB;
const HARD_SAVE_XML_PART_BYTES: u64 = 512 * MIB;
const HARD_SAVE_ROWS: u64 = 100_000_000;
const HARD_SAVE_CELLS: u64 = 100_000_000;
const HARD_STREAM_ROW_XML_BYTES: u64 = 128 * MIB;
const HARD_STREAM_OUTPUT_BYTES: u64 = 2 * GIB;
const HARD_STREAM_ROWS: u64 = 1_048_576;
const HARD_STREAM_CELLS: u64 = 100_000_000;
const HARD_STREAM_CONFIG_COUNT: u64 = 1_000_000;
const HARD_STREAM_CONFIG_BYTES: u64 = 128 * MIB;

/// Finite limits for one workbook save operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SaveLimits {
    /// Maximum compressed archive or encrypted-container output bytes.
    pub output_bytes: u64,
    /// Maximum serialized bytes in one XML package part.
    pub xml_part_bytes: u64,
    /// Maximum occupied rows serialized across the workbook.
    pub row_count: u64,
    /// Maximum occupied cells serialized across the workbook.
    pub cell_count: u64,
}

impl Default for SaveLimits {
    fn default() -> Self {
        Self {
            output_bytes: DEFAULT_SAVE_OUTPUT_BYTES,
            xml_part_bytes: DEFAULT_SAVE_XML_PART_BYTES,
            row_count: DEFAULT_SAVE_ROWS,
            cell_count: DEFAULT_SAVE_CELLS,
        }
    }
}

impl SaveLimits {
    /// Largest supported finite values. These ceilings cannot be disabled.
    pub const fn hard_ceiling() -> Self {
        Self {
            output_bytes: HARD_SAVE_OUTPUT_BYTES,
            xml_part_bytes: HARD_SAVE_XML_PART_BYTES,
            row_count: HARD_SAVE_ROWS,
            cell_count: HARD_SAVE_CELLS,
        }
    }

    /// Validates the limits against the non-overridable hard ceilings.
    pub fn validate(self) -> Result<Self> {
        validate_limits(
            ResourceOperation::WorkbookSave,
            &[
                (
                    ResourceDimension::OutputBytes,
                    self.output_bytes,
                    HARD_SAVE_OUTPUT_BYTES,
                ),
                (
                    ResourceDimension::XmlBytes,
                    self.xml_part_bytes,
                    HARD_SAVE_XML_PART_BYTES,
                ),
                (ResourceDimension::RowCount, self.row_count, HARD_SAVE_ROWS),
                (
                    ResourceDimension::CellCount,
                    self.cell_count,
                    HARD_SAVE_CELLS,
                ),
            ],
        )?;
        Ok(self)
    }
}

/// Finite limits for one streamed-sheet construction operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamLimits {
    /// Maximum occupied rows accepted by the writer.
    pub row_count: u64,
    /// Maximum cell slots accepted by the writer.
    pub cell_count: u64,
    /// Maximum serialized bytes in one row XML fragment.
    pub row_xml_bytes: u64,
    /// Maximum serialized bytes across all row XML fragments.
    pub output_bytes: u64,
    /// Maximum distinct column, merge, and pane configuration entries.
    pub config_count: u64,
    /// Maximum retained bytes used by stream configuration.
    pub config_bytes: u64,
}

impl Default for StreamLimits {
    fn default() -> Self {
        Self {
            row_count: DEFAULT_STREAM_ROWS,
            cell_count: DEFAULT_STREAM_CELLS,
            row_xml_bytes: DEFAULT_STREAM_ROW_XML_BYTES,
            output_bytes: DEFAULT_STREAM_OUTPUT_BYTES,
            config_count: DEFAULT_STREAM_CONFIG_COUNT,
            config_bytes: DEFAULT_STREAM_CONFIG_BYTES,
        }
    }
}

impl StreamLimits {
    /// Largest supported finite values. These ceilings cannot be disabled.
    pub const fn hard_ceiling() -> Self {
        Self {
            row_count: HARD_STREAM_ROWS,
            cell_count: HARD_STREAM_CELLS,
            row_xml_bytes: HARD_STREAM_ROW_XML_BYTES,
            output_bytes: HARD_STREAM_OUTPUT_BYTES,
            config_count: HARD_STREAM_CONFIG_COUNT,
            config_bytes: HARD_STREAM_CONFIG_BYTES,
        }
    }

    /// Validates the limits against the non-overridable hard ceilings.
    pub fn validate(self) -> Result<Self> {
        validate_limits(
            ResourceOperation::StreamWrite,
            &[
                (
                    ResourceDimension::RowCount,
                    self.row_count,
                    HARD_STREAM_ROWS,
                ),
                (
                    ResourceDimension::CellCount,
                    self.cell_count,
                    HARD_STREAM_CELLS,
                ),
                (
                    ResourceDimension::XmlBytes,
                    self.row_xml_bytes,
                    HARD_STREAM_ROW_XML_BYTES,
                ),
                (
                    ResourceDimension::OutputBytes,
                    self.output_bytes,
                    HARD_STREAM_OUTPUT_BYTES,
                ),
                (
                    ResourceDimension::CollectionCount,
                    self.config_count,
                    HARD_STREAM_CONFIG_COUNT,
                ),
                (
                    ResourceDimension::RetainedBytes,
                    self.config_bytes,
                    HARD_STREAM_CONFIG_BYTES,
                ),
            ],
        )?;
        Ok(self)
    }
}

pub(crate) fn resource_limit(
    operation: ResourceOperation,
    category: ResourceCategory,
    dimension: ResourceDimension,
    limit: u64,
    observed: u64,
    stage: ResourceStage,
) -> Error {
    ResourceLimitError::exceeded(
        operation,
        category,
        dimension,
        limit,
        observed.max(limit.saturating_add(1)),
        Some(stage),
    )
    .map(Error::from)
    .unwrap_or_else(|| Error::Internal("invalid save resource limit state".to_string()))
}

fn validate_limits(
    operation: ResourceOperation,
    limits: &[(ResourceDimension, u64, u64)],
) -> Result<()> {
    for &(dimension, value, maximum) in limits {
        if value == 0 || value > maximum {
            return Err(resource_limit(
                operation,
                ResourceCategory::Output,
                dimension,
                maximum,
                value.max(maximum.saturating_add(1)),
                ResourceStage::Validate,
            ));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_are_finite_and_hard_ceilings_are_enforced() {
        SaveLimits::default().validate().unwrap();
        StreamLimits::default().validate().unwrap();

        let error = SaveLimits {
            output_bytes: HARD_SAVE_OUTPUT_BYTES + 1,
            ..SaveLimits::default()
        }
        .validate()
        .unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension(),
            ResourceDimension::OutputBytes
        );

        let error = StreamLimits {
            row_xml_bytes: HARD_STREAM_ROW_XML_BYTES + 1,
            ..StreamLimits::default()
        }
        .validate()
        .unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().operation(),
            ResourceOperation::StreamWrite
        );
    }
}
