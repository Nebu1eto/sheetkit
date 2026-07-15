use std::path::{Path, PathBuf};

use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion};
use sheetkit_core::workbook::{AuxParts, OpenOptions, ReadMode, Workbook};

fn fixtures_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join("benchmarks/node/fixtures")
}

struct Fixture {
    name: &'static str,
    file: &'static str,
    sample_size: usize,
}

const FIXTURES: &[Fixture] = &[
    Fixture {
        name: "scale-1k",
        file: "scale-1k.xlsx",
        sample_size: 50,
    },
    Fixture {
        name: "scale-10k",
        file: "scale-10k.xlsx",
        sample_size: 20,
    },
    Fixture {
        name: "scale-100k",
        file: "scale-100k.xlsx",
        sample_size: 10,
    },
    Fixture {
        name: "large-data",
        file: "large-data.xlsx",
        sample_size: 10,
    },
    Fixture {
        name: "strings",
        file: "strings.xlsx",
        sample_size: 10,
    },
];

fn fixture_path(file: &str) -> PathBuf {
    fixtures_dir().join(file)
}

fn bench_open_default_lazy_latency(c: &mut Criterion) {
    let mut group = c.benchmark_group("open_default_lazy_latency");

    for f in FIXTURES {
        let path = fixture_path(f.file);
        if !path.exists() {
            eprintln!("skip {}: fixture not found at {}", f.name, path.display());
            continue;
        }

        group.sample_size(f.sample_size);
        group.bench_with_input(
            BenchmarkId::new("default-lazy", f.name),
            &path,
            |b, path| {
                b.iter(|| {
                    let wb = Workbook::open(path).expect("open failed");
                    std::hint::black_box(wb);
                });
            },
        );
    }

    group.finish();
}

fn bench_open_explicit_eager_latency(c: &mut Criterion) {
    let mut group = c.benchmark_group("open_explicit_eager_latency");
    let opts = OpenOptions::new()
        .read_mode(ReadMode::Eager)
        .aux_parts(AuxParts::EagerLoad);

    for f in FIXTURES {
        let path = fixture_path(f.file);
        if !path.exists() {
            eprintln!("skip {}: fixture not found at {}", f.name, path.display());
            continue;
        }

        group.sample_size(f.sample_size);
        group.bench_with_input(
            BenchmarkId::new("explicit-eager", f.name),
            &path,
            |b, path| {
                b.iter(|| {
                    let wb = Workbook::open_with_options(path, &opts).expect("open failed");
                    std::hint::black_box(wb);
                });
            },
        );
    }

    group.finish();
}

fn bench_get_rows(c: &mut Criterion) {
    let mut group = c.benchmark_group("get_rows");

    for f in FIXTURES {
        let path = fixture_path(f.file);
        if !path.exists() {
            eprintln!("skip {}: fixture not found at {}", f.name, path.display());
            continue;
        }

        let wb = Workbook::open(&path).expect("open failed");
        group.sample_size(f.sample_size);
        group.bench_with_input(BenchmarkId::new("Sheet1", f.name), &wb, |b, wb| {
            b.iter(|| {
                let rows = wb.get_rows("Sheet1").expect("get_rows failed");
                std::hint::black_box(rows);
            });
        });
    }

    group.finish();
}

fn bench_save_latency(c: &mut Criterion) {
    let mut group = c.benchmark_group("save_latency");

    for f in FIXTURES {
        let path = fixture_path(f.file);
        if !path.exists() {
            eprintln!("skip {}: fixture not found at {}", f.name, path.display());
            continue;
        }

        let wb = Workbook::open(&path).expect("open failed");
        group.sample_size(f.sample_size);
        group.bench_with_input(BenchmarkId::new("save", f.name), &wb, |b, wb| {
            b.iter(|| {
                let tmp = tempfile::Builder::new()
                    .suffix(".xlsx")
                    .tempfile()
                    .expect("tempfile failed");
                wb.save(tmp.path()).expect("save failed");
                std::hint::black_box(tmp);
            });
        });
    }

    group.finish();
}

criterion_group!(
    benches,
    bench_open_default_lazy_latency,
    bench_open_explicit_eager_latency,
    bench_get_rows,
    bench_save_latency,
);
criterion_main!(benches);
