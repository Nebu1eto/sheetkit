//! Finite resource limits shared by formula parsing, mutation, and calculation.

use std::cell::{OnceCell, RefCell};
use std::rc::Rc;

use crate::error::{Error, Result};
use crate::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};

/// Finite limits for one formula parse, mutation, or calculation operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaLimits {
    pub per_formula_input_bytes: u64,
    pub per_formula_nesting_depth: u64,
    pub per_formula_binary_chain_depth: u64,
    pub per_formula_tokens: u64,
    pub per_formula_ast_nodes: u64,
    pub operation_input_bytes: u64,
    pub operation_tokens: u64,
    pub operation_ast_nodes: u64,
    pub operation_range_cells: u64,
    pub operation_dependency_edges: u64,
    pub operation_ready_queue_ops: u64,
    pub operation_evaluation_steps: u64,
    pub operation_match_work: u64,
    pub operation_output_bytes: u64,
}

impl FormulaLimits {
    /// Secure defaults calibrated against the repository's formula corpus.
    pub const fn secure_default() -> Self {
        Self {
            per_formula_input_bytes: 32_768,
            per_formula_nesting_depth: 64,
            per_formula_binary_chain_depth: 512,
            per_formula_tokens: 16_384,
            per_formula_ast_nodes: 16_384,
            operation_input_bytes: 4 * 1024 * 1024,
            operation_tokens: 1_048_576,
            operation_ast_nodes: 1_048_576,
            operation_range_cells: 131_072,
            operation_dependency_edges: 524_288,
            operation_ready_queue_ops: 524_288,
            operation_evaluation_steps: 1_048_576,
            operation_match_work: 1_048_576,
            operation_output_bytes: 16 * 1024 * 1024,
        }
    }
}

impl Default for FormulaLimits {
    fn default() -> Self {
        Self::secure_default()
    }
}

#[derive(Debug, Default, Clone, Copy)]
struct FormulaCounters {
    input_bytes: u64,
    tokens: u64,
    ast_nodes: u64,
    range_cells: u64,
    dependency_edges: u64,
    ready_queue_ops: u64,
    evaluation_steps: u64,
    match_work: u64,
    output_bytes: u64,
}

/// Operation counters kept inline until the owning budget is first cloned.
/// The first clone moves the current values behind an `Rc` that the
/// original and every clone then share, so a budget that is never cloned
/// needs no heap allocation.
#[derive(Debug, Default)]
struct CounterCell {
    inline: RefCell<FormulaCounters>,
    shared: OnceCell<Rc<RefCell<FormulaCounters>>>,
}

impl CounterCell {
    #[inline]
    fn get(&self) -> &RefCell<FormulaCounters> {
        match self.shared.get() {
            Some(shared) => shared,
            None => &self.inline,
        }
    }

    fn share(&self) -> Rc<RefCell<FormulaCounters>> {
        Rc::clone(
            self.shared
                .get_or_init(|| Rc::new(RefCell::new(*self.inline.borrow()))),
        )
    }
}

impl Clone for CounterCell {
    fn clone(&self) -> Self {
        Self {
            inline: RefCell::default(),
            shared: OnceCell::from(self.share()),
        }
    }
}

/// A calculation-owned budget. Clones share counters and cannot reset work.
#[derive(Debug, Clone)]
pub struct FormulaBudget {
    operation: ResourceOperation,
    limits: FormulaLimits,
    counters: CounterCell,
}

impl FormulaBudget {
    const HARD_PER_FORMULA_INPUT_BYTES: u64 = 65_536;
    const HARD_PER_FORMULA_NESTING_DEPTH: u64 = 128;
    const HARD_PER_FORMULA_BINARY_CHAIN_DEPTH: u64 = 4_096;
    const HARD_PER_FORMULA_TOKENS: u64 = 65_536;
    const HARD_PER_FORMULA_AST_NODES: u64 = 32_768;
    const HARD_OPERATION_INPUT_BYTES: u64 = 32 * 1024 * 1024;
    const HARD_OPERATION_TOKENS: u64 = 8_388_608;
    const HARD_OPERATION_AST_NODES: u64 = 8_388_608;
    const HARD_OPERATION_RANGE_CELLS: u64 = 1_048_576;
    const HARD_OPERATION_DEPENDENCY_EDGES: u64 = 4_194_304;
    const HARD_OPERATION_READY_QUEUE_OPS: u64 = 4_194_304;
    const HARD_OPERATION_EVALUATION_STEPS: u64 = 8_388_608;
    const HARD_OPERATION_MATCH_WORK: u64 = 16_777_216;
    const HARD_OPERATION_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;
    const HARD_EVALUATION_RECURSION_DEPTH: u64 = 1_024;

    /// Create a parsing budget, rejecting limits above hard ceilings.
    pub fn for_parse(limits: FormulaLimits) -> Result<Self> {
        Self::new(ResourceOperation::FormulaParse, limits)
    }

    /// Create a formula-mutation budget, rejecting limits above hard ceilings.
    pub fn for_mutation(limits: FormulaLimits) -> Result<Self> {
        Self::new(ResourceOperation::FormulaMutate, limits)
    }

    pub(crate) fn secure_mutation() -> Self {
        Self {
            operation: ResourceOperation::FormulaMutate,
            limits: FormulaLimits::secure_default(),
            counters: CounterCell::default(),
        }
    }

    /// Create a calculation budget, rejecting limits above hard ceilings.
    pub fn for_calculation(limits: FormulaLimits) -> Result<Self> {
        Self::new(ResourceOperation::FormulaCalculate, limits)
    }

    pub(crate) fn secure_calculation() -> Self {
        Self {
            operation: ResourceOperation::FormulaCalculate,
            limits: FormulaLimits::secure_default(),
            counters: CounterCell::default(),
        }
    }

    fn new(operation: ResourceOperation, limits: FormulaLimits) -> Result<Self> {
        for (dimension, limit, ceiling) in [
            (
                ResourceDimension::InputBytes,
                limits.per_formula_input_bytes,
                Self::HARD_PER_FORMULA_INPUT_BYTES,
            ),
            (
                ResourceDimension::NestingDepth,
                limits.per_formula_nesting_depth,
                Self::HARD_PER_FORMULA_NESTING_DEPTH,
            ),
            (
                ResourceDimension::NestingDepth,
                limits.per_formula_binary_chain_depth,
                Self::HARD_PER_FORMULA_BINARY_CHAIN_DEPTH,
            ),
            (
                ResourceDimension::TokenCount,
                limits.per_formula_tokens,
                Self::HARD_PER_FORMULA_TOKENS,
            ),
            (
                ResourceDimension::NodeCount,
                limits.per_formula_ast_nodes,
                Self::HARD_PER_FORMULA_AST_NODES,
            ),
            (
                ResourceDimension::InputBytes,
                limits.operation_input_bytes,
                Self::HARD_OPERATION_INPUT_BYTES,
            ),
            (
                ResourceDimension::TokenCount,
                limits.operation_tokens,
                Self::HARD_OPERATION_TOKENS,
            ),
            (
                ResourceDimension::NodeCount,
                limits.operation_ast_nodes,
                Self::HARD_OPERATION_AST_NODES,
            ),
            (
                ResourceDimension::RangeCellCount,
                limits.operation_range_cells,
                Self::HARD_OPERATION_RANGE_CELLS,
            ),
            (
                ResourceDimension::DependencyEdgeCount,
                limits.operation_dependency_edges,
                Self::HARD_OPERATION_DEPENDENCY_EDGES,
            ),
            (
                ResourceDimension::OperationCount,
                limits.operation_ready_queue_ops,
                Self::HARD_OPERATION_READY_QUEUE_OPS,
            ),
            (
                ResourceDimension::OperationCount,
                limits.operation_evaluation_steps,
                Self::HARD_OPERATION_EVALUATION_STEPS,
            ),
            (
                ResourceDimension::MatchWork,
                limits.operation_match_work,
                Self::HARD_OPERATION_MATCH_WORK,
            ),
            (
                ResourceDimension::OutputBytes,
                limits.operation_output_bytes,
                Self::HARD_OPERATION_OUTPUT_BYTES,
            ),
        ] {
            if limit > ceiling {
                return Err(resource_error(
                    operation,
                    dimension,
                    ceiling,
                    limit,
                    ResourceStage::Validate,
                ));
            }
        }
        Ok(Self {
            operation,
            limits,
            counters: CounterCell::default(),
        })
    }

    /// Charge lexical work for one formula before recursive parsing starts.
    pub(crate) fn charge_lexical(&self, input_bytes: u64, tokens: u64) -> Result<()> {
        self.check_input(input_bytes)?;
        self.check_local(
            ResourceDimension::InputBytes,
            self.limits.per_formula_input_bytes,
            input_bytes,
            ResourceStage::Parse,
        )?;
        self.check_local(
            ResourceDimension::TokenCount,
            self.limits.per_formula_tokens,
            tokens,
            ResourceStage::Parse,
        )?;
        self.charge_counter(
            ResourceDimension::InputBytes,
            self.limits.operation_input_bytes,
            input_bytes,
            ResourceStage::Parse,
            |c| &mut c.input_bytes,
        )?;
        self.charge_counter(
            ResourceDimension::TokenCount,
            self.limits.operation_tokens,
            tokens,
            ResourceStage::Parse,
            |c| &mut c.tokens,
        )
    }

    pub(crate) fn check_input(&self, observed: u64) -> Result<()> {
        self.check_local(
            ResourceDimension::InputBytes,
            self.limits.per_formula_input_bytes,
            observed,
            ResourceStage::Parse,
        )
    }

    /// Check recursive-entry depth before the parser descends.
    pub(crate) fn check_nesting(&self, observed: u64) -> Result<()> {
        self.check_local(
            ResourceDimension::NestingDepth,
            self.limits.per_formula_nesting_depth,
            observed,
            ResourceStage::Parse,
        )
    }

    pub(crate) fn check_binary_chain(&self, observed: u64) -> Result<()> {
        self.check_local(
            ResourceDimension::NestingDepth,
            self.limits.per_formula_binary_chain_depth,
            observed,
            ResourceStage::Parse,
        )
    }

    pub(crate) fn check_evaluation_depth(&self, observed: u64) -> Result<()> {
        self.check_local(
            ResourceDimension::NestingDepth,
            Self::HARD_EVALUATION_RECURSION_DEPTH,
            observed,
            ResourceStage::Calculate,
        )
    }

    /// Charge the completed AST without resetting operation counters.
    pub(crate) fn charge_ast(&self, nodes: u64, depth: u64, binary_chain: u64) -> Result<()> {
        self.check_nesting(depth)?;
        self.check_local(
            ResourceDimension::NestingDepth,
            self.limits.per_formula_binary_chain_depth,
            binary_chain,
            ResourceStage::Parse,
        )?;
        self.check_local(
            ResourceDimension::NodeCount,
            self.limits.per_formula_ast_nodes,
            nodes,
            ResourceStage::Parse,
        )?;
        self.charge_counter(
            ResourceDimension::NodeCount,
            self.limits.operation_ast_nodes,
            nodes,
            ResourceStage::Parse,
            |c| &mut c.ast_nodes,
        )
    }

    pub(crate) fn charge_range_cells(&self, amount: u64) -> Result<()> {
        self.charge_counter(
            ResourceDimension::RangeCellCount,
            self.limits.operation_range_cells,
            amount,
            ResourceStage::Calculate,
            |c| &mut c.range_cells,
        )
    }

    pub(crate) fn charge_dependency_edges(&self, amount: u64) -> Result<()> {
        self.charge_counter(
            ResourceDimension::DependencyEdgeCount,
            self.limits.operation_dependency_edges,
            amount,
            ResourceStage::Calculate,
            |c| &mut c.dependency_edges,
        )
    }

    pub(crate) fn charge_ready_queue_ops(&self, amount: u64) -> Result<()> {
        self.charge_counter(
            ResourceDimension::OperationCount,
            self.limits.operation_ready_queue_ops,
            amount,
            ResourceStage::Calculate,
            |c| &mut c.ready_queue_ops,
        )
    }

    pub(crate) fn charge_evaluation_steps(&self, amount: u64) -> Result<()> {
        self.charge_counter(
            ResourceDimension::OperationCount,
            self.limits.operation_evaluation_steps,
            amount,
            ResourceStage::Calculate,
            |c| &mut c.evaluation_steps,
        )
    }

    pub(crate) fn charge_match_work(&self, amount: u64) -> Result<()> {
        self.charge_counter(
            ResourceDimension::MatchWork,
            self.limits.operation_match_work,
            amount,
            ResourceStage::Calculate,
            |c| &mut c.match_work,
        )
    }

    pub(crate) fn charge_output_bytes(&self, amount: u64) -> Result<()> {
        self.charge_counter(
            ResourceDimension::OutputBytes,
            self.limits.operation_output_bytes,
            amount,
            ResourceStage::Calculate,
            |c| &mut c.output_bytes,
        )
    }

    /// Charge one opaque formula copied or shifted by a mutation operation.
    pub(crate) fn charge_mutation_formula(
        &self,
        input_bytes: u64,
        output_bytes: u64,
    ) -> Result<()> {
        self.charge_mutation_input(input_bytes)?;
        self.charge_mutation_output(output_bytes)
    }

    /// Charge one opaque formula before mutation code scans or copies it.
    pub(crate) fn charge_mutation_input(&self, input_bytes: u64) -> Result<()> {
        self.check_mutation_input(input_bytes)?;
        self.charge_mutation_input_total(input_bytes)
    }

    /// Check one opaque formula before any scan proportional to its bytes.
    pub(crate) fn check_mutation_input(&self, input_bytes: u64) -> Result<()> {
        self.check_local(
            ResourceDimension::InputBytes,
            self.limits.per_formula_input_bytes,
            input_bytes,
            ResourceStage::Validate,
        )
    }

    /// Charge cumulative mutation scans after checked bulk preflight.
    pub(crate) fn charge_mutation_input_total(&self, input_bytes: u64) -> Result<()> {
        self.charge_counter(
            ResourceDimension::InputBytes,
            self.limits.operation_input_bytes,
            input_bytes,
            ResourceStage::Validate,
            |c| &mut c.input_bytes,
        )
    }

    /// Charge exact staged formula bytes before allocating the output.
    pub(crate) fn charge_mutation_output(&self, output_bytes: u64) -> Result<()> {
        self.charge_counter(
            ResourceDimension::OutputBytes,
            self.limits.operation_output_bytes,
            output_bytes,
            ResourceStage::Serialize,
            |c| &mut c.output_bytes,
        )
    }

    /// Charge the number of cells targeted by one formula mutation.
    pub(crate) fn charge_mutation_cells(&self, amount: u64) -> Result<()> {
        self.charge_counter(
            ResourceDimension::RangeCellCount,
            self.limits.operation_range_cells,
            amount,
            ResourceStage::Validate,
            |c| &mut c.range_cells,
        )
    }

    pub(crate) fn check_output_bytes(&self, amount: u64) -> Result<()> {
        let observed = self
            .counters
            .get()
            .borrow()
            .output_bytes
            .saturating_add(amount);
        self.check_local(
            ResourceDimension::OutputBytes,
            self.limits.operation_output_bytes,
            observed,
            ResourceStage::Calculate,
        )
    }

    fn check_local(
        &self,
        dimension: ResourceDimension,
        limit: u64,
        observed: u64,
        stage: ResourceStage,
    ) -> Result<()> {
        if observed > limit {
            return Err(resource_error(
                self.operation,
                dimension,
                limit,
                observed,
                stage,
            ));
        }
        Ok(())
    }

    fn charge_counter(
        &self,
        dimension: ResourceDimension,
        limit: u64,
        amount: u64,
        stage: ResourceStage,
        select: impl FnOnce(&mut FormulaCounters) -> &mut u64,
    ) -> Result<()> {
        let mut counters = self.counters.get().borrow_mut();
        let counter = select(&mut counters);
        let observed = counter.checked_add(amount).unwrap_or(u64::MAX);
        if observed > limit {
            *counter = observed;
            return Err(resource_error(
                self.operation,
                dimension,
                limit,
                observed,
                stage,
            ));
        }
        *counter = observed;
        Ok(())
    }
}

fn resource_error(
    operation: ResourceOperation,
    dimension: ResourceDimension,
    limit: u64,
    observed: u64,
    stage: ResourceStage,
) -> Error {
    match ResourceLimitError::exceeded(
        operation,
        ResourceCategory::Formula,
        dimension,
        limit,
        observed,
        Some(stage),
    ) {
        Some(error) => Error::ResourceLimitExceeded(error),
        None => Error::Internal("invalid formula resource-limit state".to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exact_limit_is_accepted_and_first_unit_over_is_stable() {
        let limits = FormulaLimits {
            operation_evaluation_steps: 2,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        budget.charge_evaluation_steps(1).unwrap();
        budget.charge_evaluation_steps(1).unwrap();
        let error = budget.charge_evaluation_steps(1).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!((metadata.limit(), metadata.observed()), (2, 3));
        assert_eq!(metadata.operation_code(), "formula_calculate");
        assert_eq!(metadata.dimension_code(), "operation_count");
    }

    #[test]
    fn clones_share_counters_charged_before_and_after_cloning() {
        let limits = FormulaLimits {
            operation_evaluation_steps: 3,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        budget.charge_evaluation_steps(1).unwrap();
        let clone = budget.clone();
        clone.charge_evaluation_steps(1).unwrap();
        let second = clone.clone();
        budget.charge_evaluation_steps(1).unwrap();
        let error = second.charge_evaluation_steps(1).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!((metadata.limit(), metadata.observed()), (3, 4));
        let error = budget.charge_evaluation_steps(0).unwrap_err();
        assert_eq!(error.resource_metadata().unwrap().observed(), 4);
    }

    #[test]
    fn checked_counter_overflow_is_an_exceeded_limit() {
        let limits = FormulaLimits {
            operation_match_work: FormulaBudget::HARD_OPERATION_MATCH_WORK,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        let error = budget.charge_match_work(u64::MAX).unwrap_err();
        assert_eq!(error.resource_metadata().unwrap().observed(), u64::MAX);
    }

    #[test]
    fn output_preflight_does_not_consume_the_counter() {
        let limits = FormulaLimits {
            operation_output_bytes: 3,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        budget.check_output_bytes(3).unwrap();
        budget.check_output_bytes(3).unwrap();
        budget.charge_output_bytes(3).unwrap();
        let error = budget.check_output_bytes(1).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!((metadata.limit(), metadata.observed()), (3, 4));
    }

    #[test]
    fn trusted_limits_cannot_exceed_hard_ceiling() {
        let limits = FormulaLimits {
            operation_range_cells: FormulaBudget::HARD_OPERATION_RANGE_CELLS + 1,
            ..FormulaLimits::secure_default()
        };
        let error = FormulaBudget::for_calculation(limits).unwrap_err();
        let metadata = error.resource_metadata().unwrap();
        assert_eq!(metadata.stage_code(), Some("validate"));
        assert_eq!(metadata.dimension_code(), "range_cell_count");
    }

    #[test]
    fn matcher_and_output_counters_are_independent() {
        let limits = FormulaLimits {
            operation_match_work: 2,
            operation_output_bytes: 3,
            ..FormulaLimits::secure_default()
        };
        let budget = FormulaBudget::for_calculation(limits).unwrap();
        budget.charge_match_work(2).unwrap();
        budget.charge_output_bytes(3).unwrap();
        assert_eq!(
            budget
                .charge_match_work(1)
                .unwrap_err()
                .resource_metadata()
                .unwrap()
                .dimension_code(),
            "match_work"
        );
        assert_eq!(
            budget
                .charge_output_bytes(1)
                .unwrap_err()
                .resource_metadata()
                .unwrap()
                .dimension_code(),
            "output_bytes"
        );
    }
}
