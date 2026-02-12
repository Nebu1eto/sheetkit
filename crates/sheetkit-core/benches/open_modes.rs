use std::path::{Path, PathBuf};

use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion};
use sheetkit_core::workbook::{OpenOptions, ReadMode, Workbook};

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
        name: "large-data",
        file: "large-data.xlsx",
        sample_size: 10,
    },
];

fn fixture_path(file: &str) -> PathBuf {
    fixtures_dir().join(file)
}

fn bench_open_latency(c: &mut Criterion) {
    let mut group = c.benchmark_group("open_latency");

    for f in FIXTURES {
        let path = fixture_path(f.file);
        if !path.exists() {
            eprintln!("skip {}: fixture not found at {}", f.name, path.display());
            continue;
        }

        group.sample_size(f.sample_size);
        group.bench_with_input(BenchmarkId::new("full", f.name), &path, |b, path| {
            b.iter(|| {
                let wb = Workbook::open(path).expect("open failed");
                std::hint::black_box(wb);
            });
        });
    }

    group.finish();
}

fn bench_open_readfast_latency(c: &mut Criterion) {
    let mut group = c.benchmark_group("open_readfast_latency");
    let opts = OpenOptions::new().read_mode(ReadMode::Lazy);

    for f in FIXTURES {
        let path = fixture_path(f.file);
        if !path.exists() {
            eprintln!("skip {}: fixture not found at {}", f.name, path.display());
            continue;
        }

        group.sample_size(f.sample_size);
        group.bench_with_input(BenchmarkId::new("readfast", f.name), &path, |b, path| {
            b.iter(|| {
                let wb = Workbook::open_with_options(path, &opts).expect("open failed");
                std::hint::black_box(wb);
            });
        });
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
                let tmp = tempfile::NamedTempFile::new().expect("tempfile failed");
                wb.save(tmp.path()).expect("save failed");
                std::hint::black_box(tmp);
            });
        });
    }

    group.finish();
}

criterion_group!(
    benches,
    bench_open_latency,
    bench_open_readfast_latency,
    bench_get_rows,
    bench_save_latency,
);
criterion_main!(benches);
