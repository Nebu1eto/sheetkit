use std::collections::HashMap;

use criterion::{black_box, criterion_group, criterion_main, Criterion};
use sheetkit_core::formula::{parse_formula, topological_sort, CellCoord};
use sheetkit_core::workbook::Workbook;

fn bench_bounded_formula_parse(c: &mut Criterion) {
    let formula = std::iter::repeat_n("1", 129).collect::<Vec<_>>().join("+");
    c.bench_function("formula_security/bounded_parse", |b| {
        b.iter(|| parse_formula(black_box(&formula)).expect("formula should parse"));
    });
}

fn bench_oversized_range_rejection(c: &mut Criterion) {
    let workbook = Workbook::new();
    c.bench_function("formula_security/range_rejection", |b| {
        b.iter(|| {
            let result = workbook.evaluate_formula("Sheet1", black_box("A1:XFD1048576"));
            assert!(result.is_err());
        });
    });
}

fn bench_large_ready_queue(c: &mut Criterion) {
    let cells: Vec<CellCoord> = (1..=10_000)
        .rev()
        .map(|row| CellCoord {
            sheet: "Sheet1".to_string(),
            col: 1,
            row,
        })
        .collect();
    let deps: HashMap<CellCoord, Vec<CellCoord>> = cells
        .iter()
        .cloned()
        .map(|cell| (cell, Vec::new()))
        .collect();
    c.bench_function("formula_security/large_ready_queue", |b| {
        b.iter(|| {
            topological_sort(black_box(&cells), black_box(&deps))
                .expect("independent graph should sort")
        });
    });
}

fn bench_bounded_wildcard_criteria(c: &mut Criterion) {
    let mut workbook = Workbook::new();
    let text = "a".repeat(256);
    let pattern = "*a".repeat(128);
    let formula = format!("SUMIF(A1,\"{pattern}\",B1)");
    workbook
        .set_cell_value("Sheet1", "A1", text)
        .expect("benchmark input should be accepted");
    workbook
        .set_cell_value("Sheet1", "B1", 1.0)
        .expect("benchmark input should be accepted");
    c.bench_function("formula_security/bounded_wildcard_criteria", |b| {
        b.iter(|| {
            workbook
                .evaluate_formula("Sheet1", black_box(&formula))
                .expect("bounded wildcard should evaluate")
        });
    });
}

fn bench_bounded_cumprinc(c: &mut Criterion) {
    let workbook = Workbook::new();
    c.bench_function("formula_security/bounded_cumprinc", |b| {
        b.iter(|| {
            workbook
                .evaluate_formula("Sheet1", black_box("CUMPRINC(0.01,10000,1000,1,10000,0)"))
                .expect("bounded period interval should evaluate")
        });
    });
}

fn bench_bounded_wildcard_search_failure(c: &mut Criterion) {
    let mut workbook = Workbook::new();
    workbook
        .set_cell_value("Sheet1", "A1", "a".repeat(256))
        .expect("benchmark input should be accepted");
    let pattern = format!("{}z", "*a".repeat(128));
    let formula = format!("SEARCH(\"{pattern}\",A1)");
    c.bench_function("formula_security/bounded_wildcard_search_failure", |b| {
        b.iter(|| {
            workbook
                .evaluate_formula("Sheet1", black_box(&formula))
                .expect("bounded wildcard search should evaluate")
        });
    });
}

fn bench_substitute_amplification_rejection(c: &mut Criterion) {
    let mut workbook = Workbook::new();
    workbook
        .set_cell_value("Sheet1", "A1", "a".repeat(32_767))
        .expect("benchmark input should be accepted");
    workbook
        .set_cell_value("Sheet1", "A2", "a")
        .expect("benchmark input should be accepted");
    workbook
        .set_cell_value("Sheet1", "A3", "b".repeat(32_767))
        .expect("benchmark input should be accepted");
    c.bench_function("formula_security/substitute_amplification_rejection", |b| {
        b.iter(|| {
            let result = workbook.evaluate_formula("Sheet1", black_box("SUBSTITUTE(A1,A2,A3)"));
            assert!(result.is_err());
        });
    });
}

criterion_group!(
    benches,
    bench_bounded_formula_parse,
    bench_oversized_range_rejection,
    bench_large_ready_queue,
    bench_bounded_wildcard_criteria,
    bench_bounded_cumprinc,
    bench_bounded_wildcard_search_failure,
    bench_substitute_amplification_rejection
);
criterion_main!(benches);
