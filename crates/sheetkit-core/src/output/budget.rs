use crate::error::{Error, Result};
use crate::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};

/// A non-streaming output budget shared by context-specific serializers.
pub(crate) struct OutputBudget {
    operation: ResourceOperation,
    limit: u64,
    used: u64,
}

impl OutputBudget {
    pub(crate) const fn new(operation: ResourceOperation, limit: u64) -> Self {
        Self {
            operation,
            limit,
            used: 0,
        }
    }

    /// Charges bytes before a serializer emits them.
    pub(crate) fn charge(&mut self, additional: u64) -> Result<()> {
        let observed = self.used.saturating_add(additional);
        if observed > self.limit {
            return Err(self.limit_error(observed));
        }

        self.used = observed;
        Ok(())
    }

    /// Charges and reserves bytes for a staged byte output.
    pub(crate) fn reserve(&mut self, output: &mut Vec<u8>, additional: u64) -> Result<()> {
        self.charge(additional)?;

        let additional = usize::try_from(additional)
            .map_err(|_| Error::Internal("output size is not addressable".to_string()))?;
        output
            .try_reserve_exact(additional)
            .map_err(|_| Error::Internal("output allocation failed".to_string()))?;
        Ok(())
    }

    fn limit_error(&self, observed: u64) -> Error {
        match ResourceLimitError::exceeded(
            self.operation,
            ResourceCategory::Output,
            ResourceDimension::OutputBytes,
            self.limit,
            observed.max(self.limit.saturating_add(1)),
            Some(ResourceStage::Serialize),
        ) {
            Some(metadata) => metadata.into(),
            None => Error::Internal("invalid output resource limit".to_string()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn overflow_is_reported_as_a_stable_resource_error() {
        let mut budget = OutputBudget::new(ResourceOperation::Export, u64::MAX - 1);
        let mut output = Vec::new();
        budget.used = 1;

        let error = budget.reserve(&mut output, u64::MAX).unwrap_err();
        let Error::ResourceLimitExceeded(metadata) = error else {
            panic!("expected resource limit error");
        };
        assert_eq!(metadata.operation(), ResourceOperation::Export);
        assert_eq!(metadata.category(), ResourceCategory::Output);
        assert_eq!(metadata.dimension(), ResourceDimension::OutputBytes);
        assert_eq!(metadata.observed(), u64::MAX);
        assert!(output.is_empty());
    }
}
