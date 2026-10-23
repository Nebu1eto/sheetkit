use crate::error::{Error, Result};
use crate::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};

const MIB: u64 = 1024 * 1024;

/// Finite limits applied to one workbook-open operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OpenLimits {
    /// Maximum bytes in the outer ZIP or container input.
    pub outer_input_bytes: u64,
    /// Maximum central-directory entry count.
    pub archive_entries: u64,
    /// Maximum size declared by one archive entry.
    pub entry_declared_bytes: u64,
    /// Maximum bytes actually read from one archive entry.
    pub entry_actual_bytes: u64,
    /// Maximum sum of archive entry declared sizes.
    pub aggregate_declared_bytes: u64,
    /// Maximum cumulative bytes actually read from archive entries.
    pub aggregate_actual_bytes: u64,
    /// Maximum uncompressed-to-compressed ratio for one entry.
    pub compression_ratio: u64,
    /// Maximum bytes actually read from one XML part.
    pub xml_bytes: u64,
    /// Maximum raw package bytes retained by the resulting workbook.
    pub retained_package_bytes: u64,
}

impl Default for OpenLimits {
    fn default() -> Self {
        Self {
            outer_input_bytes: 128 * MIB,
            archive_entries: 4_096,
            entry_declared_bytes: 256 * MIB,
            entry_actual_bytes: 256 * MIB,
            aggregate_declared_bytes: 512 * MIB,
            aggregate_actual_bytes: 512 * MIB,
            compression_ratio: 100,
            xml_bytes: 256 * MIB,
            retained_package_bytes: 512 * MIB,
        }
    }
}

impl OpenLimits {
    /// Largest supported finite values. These ceilings cannot be disabled.
    pub const fn hard_ceiling() -> Self {
        Self {
            outer_input_bytes: 2 * 1024 * MIB,
            archive_entries: 65_536,
            entry_declared_bytes: 1024 * MIB,
            entry_actual_bytes: 1024 * MIB,
            aggregate_declared_bytes: 4 * 1024 * MIB,
            aggregate_actual_bytes: 4 * 1024 * MIB,
            compression_ratio: 1_000,
            xml_bytes: 1024 * MIB,
            retained_package_bytes: 2 * 1024 * MIB,
        }
    }

    pub(crate) fn validate(self) -> Result<Self> {
        let ceiling = Self::hard_ceiling();
        for (dimension, value, maximum) in [
            (
                ResourceDimension::InputBytes,
                self.outer_input_bytes,
                ceiling.outer_input_bytes,
            ),
            (
                ResourceDimension::EntryCount,
                self.archive_entries,
                ceiling.archive_entries,
            ),
            (
                ResourceDimension::EntryDeclaredBytes,
                self.entry_declared_bytes,
                ceiling.entry_declared_bytes,
            ),
            (
                ResourceDimension::EntryActualBytes,
                self.entry_actual_bytes,
                ceiling.entry_actual_bytes,
            ),
            (
                ResourceDimension::AggregateDeclaredBytes,
                self.aggregate_declared_bytes,
                ceiling.aggregate_declared_bytes,
            ),
            (
                ResourceDimension::AggregateActualBytes,
                self.aggregate_actual_bytes,
                ceiling.aggregate_actual_bytes,
            ),
            (
                ResourceDimension::CompressionRatio,
                self.compression_ratio,
                ceiling.compression_ratio,
            ),
            (
                ResourceDimension::XmlBytes,
                self.xml_bytes,
                ceiling.xml_bytes,
            ),
            (
                ResourceDimension::RetainedBytes,
                self.retained_package_bytes,
                ceiling.retained_package_bytes,
            ),
        ] {
            if value == 0 || value > maximum {
                return Err(limit_error(
                    dimension,
                    maximum,
                    value.max(maximum.saturating_add(1)),
                    ResourceStage::Validate,
                ));
            }
        }
        Ok(self)
    }
}

pub(crate) struct OpenBudget {
    limits: OpenLimits,
    aggregate_declared: u64,
    aggregate_actual: u64,
    retained: u64,
    violation: Option<ResourceLimitError>,
    integrity_violation: Option<String>,
}

impl OpenBudget {
    pub(crate) fn new(limits: OpenLimits) -> Result<Self> {
        Ok(Self {
            limits: limits.validate()?,
            aggregate_declared: 0,
            aggregate_actual: 0,
            retained: 0,
            violation: None,
            integrity_violation: None,
        })
    }

    pub(crate) const fn limits(&self) -> OpenLimits {
        self.limits
    }

    pub(crate) fn check_outer_input(&self, observed: u64) -> Result<()> {
        check(
            ResourceDimension::InputBytes,
            self.limits.outer_input_bytes,
            observed,
            ResourceStage::Inspect,
        )
    }

    pub(crate) fn charge_retained(&mut self, amount: u64) -> Result<()> {
        let observed = checked_total(self.retained, amount);
        check(
            ResourceDimension::RetainedBytes,
            self.limits.retained_package_bytes,
            observed,
            ResourceStage::Read,
        )?;
        self.retained = observed;
        Ok(())
    }

    pub(crate) fn preflight_entry(&mut self, declared: u64, compressed: u64) -> Result<()> {
        check(
            ResourceDimension::EntryDeclaredBytes,
            self.limits.entry_declared_bytes,
            declared,
            ResourceStage::Inspect,
        )?;
        let aggregate = checked_total(self.aggregate_declared, declared);
        check(
            ResourceDimension::AggregateDeclaredBytes,
            self.limits.aggregate_declared_bytes,
            aggregate,
            ResourceStage::Inspect,
        )?;
        check_ratio(
            declared,
            compressed,
            self.limits.compression_ratio,
            ResourceStage::Inspect,
        )?;
        self.aggregate_declared = aggregate;
        Ok(())
    }

    pub(crate) fn charge_actual(
        &mut self,
        entry_total: u64,
        amount: u64,
        compressed: u64,
        xml: bool,
    ) -> Result<u64> {
        let entry_observed = checked_total(entry_total, amount);
        check(
            ResourceDimension::EntryActualBytes,
            self.limits.entry_actual_bytes,
            entry_observed,
            ResourceStage::Read,
        )?;
        if xml {
            check(
                ResourceDimension::XmlBytes,
                self.limits.xml_bytes,
                entry_observed,
                ResourceStage::Read,
            )?;
        }
        check_ratio(
            entry_observed,
            compressed,
            self.limits.compression_ratio,
            ResourceStage::Read,
        )?;
        let aggregate = checked_total(self.aggregate_actual, amount);
        check(
            ResourceDimension::AggregateActualBytes,
            self.limits.aggregate_actual_bytes,
            aggregate,
            ResourceStage::Read,
        )?;
        self.aggregate_actual = aggregate;
        Ok(entry_observed)
    }

    pub(crate) fn record_violation(&mut self, error: Error) {
        if self.violation.is_none() {
            if let Error::ResourceLimitExceeded(metadata) = error {
                self.violation = Some(metadata);
            }
        }
    }

    /// Record the first archive integrity failure observed while reading.
    pub(crate) fn record_integrity_violation(&mut self, message: String) {
        if self.integrity_violation.is_none() {
            self.integrity_violation = Some(message);
        }
    }

    /// Return the first recorded resource or integrity failure, if any.
    pub(crate) fn pending_error(&self) -> Option<Error> {
        if let Some(metadata) = self.violation {
            return Some(Error::ResourceLimitExceeded(metadata));
        }
        self.integrity_violation
            .as_ref()
            .map(|message| Error::Zip(message.clone()))
    }
}

fn checked_total(current: u64, amount: u64) -> u64 {
    let (total, overflowed) = current.overflowing_add(amount);
    if overflowed {
        u64::MAX
    } else {
        total
    }
}

fn check(
    dimension: ResourceDimension,
    limit: u64,
    observed: u64,
    stage: ResourceStage,
) -> Result<()> {
    if observed > limit {
        return Err(limit_error(dimension, limit, observed, stage));
    }
    Ok(())
}

fn check_ratio(actual: u64, compressed: u64, limit: u64, stage: ResourceStage) -> Result<()> {
    let observed = if actual == 0 {
        0
    } else if compressed == 0 {
        u64::MAX
    } else {
        actual / compressed + u64::from(!actual.is_multiple_of(compressed))
    };
    check(ResourceDimension::CompressionRatio, limit, observed, stage)
}

pub(crate) fn limit_error(
    dimension: ResourceDimension,
    limit: u64,
    observed: u64,
    stage: ResourceStage,
) -> Error {
    match ResourceLimitError::exceeded(
        ResourceOperation::WorkbookOpen,
        if dimension == ResourceDimension::XmlBytes {
            ResourceCategory::Xml
        } else {
            ResourceCategory::Archive
        },
        dimension,
        limit,
        observed,
        Some(stage),
    ) {
        Some(metadata) => Error::ResourceLimitExceeded(metadata),
        None => Error::Internal("invalid workbook resource limit state".to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exact_limit_is_accepted_and_one_over_is_rejected() {
        let mut budget = OpenBudget::new(OpenLimits {
            entry_actual_bytes: 8,
            ..OpenLimits::default()
        })
        .unwrap();
        assert_eq!(budget.charge_actual(0, 8, 8, false).unwrap(), 8);
        let error = budget.charge_actual(8, 1, 8, false).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension(),
            ResourceDimension::EntryActualBytes
        );
    }

    #[test]
    fn checked_aggregate_overflow_fails_closed() {
        let mut budget = OpenBudget::new(OpenLimits::hard_ceiling()).unwrap();
        budget.aggregate_actual = u64::MAX - 1;
        let error = budget.charge_actual(0, 2, 2, false).unwrap_err();
        assert_eq!(error.resource_metadata().unwrap().observed(), u64::MAX);
    }

    #[test]
    fn non_empty_zero_compressed_entry_has_unbounded_ratio() {
        let mut budget = OpenBudget::new(OpenLimits::default()).unwrap();
        let error = budget.preflight_entry(1, 0).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension(),
            ResourceDimension::CompressionRatio
        );
    }

    #[test]
    fn retained_and_ratio_boundaries_are_inclusive() {
        let mut budget = OpenBudget::new(OpenLimits {
            retained_package_bytes: 8,
            compression_ratio: 4,
            ..OpenLimits::default()
        })
        .unwrap();
        budget.charge_retained(8).unwrap();
        assert_eq!(budget.charge_actual(0, 8, 2, false).unwrap(), 8);

        let retained = budget.charge_retained(1).unwrap_err();
        assert_eq!(
            retained.resource_metadata().unwrap().dimension(),
            ResourceDimension::RetainedBytes
        );
        let ratio = budget.charge_actual(0, 9, 2, false).unwrap_err();
        assert_eq!(
            ratio.resource_metadata().unwrap().dimension(),
            ResourceDimension::CompressionRatio
        );
    }

    #[test]
    fn aggregate_actual_is_cumulative_across_entry_reads() {
        let mut budget = OpenBudget::new(OpenLimits {
            aggregate_actual_bytes: 8,
            ..OpenLimits::default()
        })
        .unwrap();
        budget.charge_actual(0, 4, 4, false).unwrap();
        budget.charge_actual(0, 4, 4, false).unwrap();
        let error = budget.charge_actual(0, 1, 1, false).unwrap_err();
        assert_eq!(
            error.resource_metadata().unwrap().dimension(),
            ResourceDimension::AggregateActualBytes
        );
    }
}
