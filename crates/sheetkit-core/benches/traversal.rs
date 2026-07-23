use criterion::{black_box, criterion_group, criterion_main, BatchSize, Criterion};
use sheetkit_core::pivot::{AggregateFunction, PivotDataField, PivotField, PivotTableConfig};
use sheetkit_core::raw_transfer::sheet_to_raw_buffer_with_options;
use sheetkit_core::raw_transfer_v2::sheet_to_raw_buffer_v2;
use sheetkit_core::render::{render_to_svg, RenderOptions};
use sheetkit_core::sst::SharedStringTable;
use sheetkit_core::traversal::{TraversalMode, TraversalOptions};
use sheetkit_core::workbook::Workbook;
use sheetkit_xml::styles::StyleSheet;
use sheetkit_xml::worksheet::{Cell, CellTypeTag, CompactCellRef, Row, SheetData, WorksheetXml};

fn cell(reference: &str, row: u32) -> Cell {
    Cell {
        r: CompactCellRef::new(reference).expect("valid benchmark cell reference"),
        col: 1,
        s: None,
        t: CellTypeTag::Number,
        v: Some(row.to_string()),
        f: None,
        is: None,
    }
}

fn row(number: u32, reference: &str) -> Row {
    Row {
        r: number,
        spans: None,
        s: None,
        custom_format: None,
        ht: None,
        hidden: None,
        custom_height: None,
        outline_level: None,
        cells: vec![cell(reference, number)],
    }
}

fn sparse_sheet(last_row: u32) -> WorksheetXml {
    WorksheetXml {
        sheet_data: SheetData {
            rows: vec![row(1, "A1"), row(last_row, &format!("A{last_row}"))],
        },
        ..WorksheetXml::default()
    }
}

fn pivot_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    workbook.set_cell_value("Sheet1", "A1", "Name").unwrap();
    workbook.set_cell_value("Sheet1", "B1", "Sales").unwrap();
    workbook.set_cell_value("Sheet1", "A2", "Alice").unwrap();
    workbook.set_cell_value("Sheet1", "B2", 100.0).unwrap();
    workbook
}

fn pivot_config() -> PivotTableConfig {
    PivotTableConfig {
        name: "PivotTable1".to_string(),
        source_sheet: "Sheet1".to_string(),
        source_range: "A1:B2".to_string(),
        target_sheet: "Sheet1".to_string(),
        target_cell: "D1".to_string(),
        rows: vec![PivotField {
            name: "Name".to_string(),
        }],
        columns: vec![],
        data: vec![PivotDataField {
            name: "Sales".to_string(),
            function: AggregateFunction::Sum,
            display_name: None,
        }],
    }
}

fn unique_pivot_workbook(value_count: u32) -> Workbook {
    let mut workbook = Workbook::new();
    workbook.set_cell_value("Sheet1", "A1", "Value").unwrap();
    for index in 0..value_count {
        workbook
            .set_cell_value(
                "Sheet1",
                &format!("A{}", index + 2),
                format!("value-{index:05}"),
            )
            .unwrap();
    }
    workbook
}

fn unique_pivot_config(value_count: u32) -> PivotTableConfig {
    PivotTableConfig {
        name: "UniqueValues".to_string(),
        source_sheet: "Sheet1".to_string(),
        source_range: format!("A1:A{}", value_count + 1),
        target_sheet: "Sheet1".to_string(),
        target_cell: "C1".to_string(),
        rows: vec![PivotField {
            name: "Value".to_string(),
        }],
        columns: vec![],
        data: vec![PivotDataField {
            name: "Value".to_string(),
            function: AggregateFunction::Count,
            display_name: None,
        }],
    }
}

fn traversal_benchmarks(criterion: &mut Criterion) {
    let strings = SharedStringTable::new();
    let near = sparse_sheet(2);
    let far = sparse_sheet(1_048_576);

    criterion.bench_function("raw_v1_sparse_near", |bencher| {
        bencher.iter(|| {
            sheet_to_raw_buffer_with_options(
                black_box(&near),
                black_box(&strings),
                TraversalOptions::default(),
            )
        });
    });
    criterion.bench_function("raw_v1_sparse_far", |bencher| {
        bencher.iter(|| {
            sheet_to_raw_buffer_with_options(
                black_box(&far),
                black_box(&strings),
                TraversalOptions::default(),
            )
        });
    });
    criterion.bench_function("raw_v1_dense_near", |bencher| {
        bencher.iter(|| {
            sheet_to_raw_buffer_with_options(
                black_box(&near),
                black_box(&strings),
                TraversalOptions::new(TraversalMode::Dense),
            )
        });
    });
    criterion.bench_function("raw_v2_sparse_far", |bencher| {
        bencher.iter(|| sheet_to_raw_buffer_v2(black_box(&far), black_box(&strings)));
    });

    let render_options = RenderOptions {
        sheet_name: "Sheet1".to_string(),
        ..RenderOptions::default()
    };
    criterion.bench_function("render_lazy_layout_near", |bencher| {
        bencher.iter(|| {
            render_to_svg(
                black_box(&near),
                black_box(&strings),
                black_box(&StyleSheet::default()),
                black_box(&render_options),
            )
        });
    });

    let config = pivot_config();
    criterion.bench_function("pivot_materialize_small", |bencher| {
        bencher.iter_batched(
            pivot_workbook,
            |mut workbook| workbook.add_pivot_table(black_box(&config)),
            BatchSize::SmallInput,
        );
    });

    let unique_count = 50_000;
    let unique_config = unique_pivot_config(unique_count);
    criterion.bench_function("pivot_unique_50k_linear_index", |bencher| {
        bencher.iter_batched(
            || unique_pivot_workbook(unique_count),
            |mut workbook| workbook.add_pivot_table(black_box(&unique_config)),
            BatchSize::LargeInput,
        );
    });
}

criterion_group!(benches, traversal_benchmarks);
criterion_main!(benches);
