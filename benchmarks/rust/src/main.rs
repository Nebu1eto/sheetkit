use sheetkit::utils::column_number_to_name;
/// Rust native benchmark for SheetKit.
///
/// Mirrors the Node.js benchmark scenarios (28 total) to measure
/// napi-rs overhead by comparing pure Rust performance against
/// the Node.js bindings.
///
/// Each scenario runs multiple iterations to collect statistical
/// metrics: min, max, median, p95, and peak RSS delta.
use sheetkit::{
    CellValue, CommentConfig, DataValidationConfig, Style, ValidationOperator, ValidationType,
    Workbook,
};

use std::fs;
use std::io::Write as _;
use std::path::{Path, PathBuf};
use std::time::Instant;

const WARMUP_RUNS: usize = 1;
const BENCH_RUNS: usize = 5;

fn fixtures_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("node")
        .join("fixtures")
}

fn output_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("output")
}

fn cell_ref(row: u32, col: u32) -> String {
    let col_name = column_number_to_name(col).unwrap();
    format!("{col_name}{row}")
}

fn format_ms(ms: f64) -> String {
    if ms < 1000.0 {
        format!("{:.0}ms", ms)
    } else {
        format!("{:.2}s", ms / 1000.0)
    }
}

fn file_size_kb(path: &Path) -> Option<f64> {
    fs::metadata(path).ok().map(|m| m.len() as f64 / 1024.0)
}

#[cfg(target_os = "macos")]
#[allow(deprecated)]
fn current_rss_bytes() -> usize {
    use std::mem::MaybeUninit;
    unsafe {
        let mut info: MaybeUninit<libc::mach_task_basic_info_data_t> = MaybeUninit::uninit();
        let mut count = (std::mem::size_of::<libc::mach_task_basic_info_data_t>()
            / std::mem::size_of::<libc::natural_t>())
            as libc::mach_msg_type_number_t;
        let kr = libc::task_info(
            libc::mach_task_self(),
            libc::MACH_TASK_BASIC_INFO,
            info.as_mut_ptr() as libc::task_info_t,
            &mut count,
        );
        if kr == libc::KERN_SUCCESS {
            info.assume_init().resident_size as usize
        } else {
            0
        }
    }
}

#[cfg(target_os = "linux")]
fn current_rss_bytes() -> usize {
    fs::read_to_string("/proc/self/statm")
        .ok()
        .and_then(|s| s.split_whitespace().nth(1)?.parse::<usize>().ok())
        .map(|pages| pages * 4096)
        .unwrap_or(0)
}

#[cfg(not(any(target_os = "macos", target_os = "linux")))]
fn current_rss_bytes() -> usize {
    0
}

fn detect_cpu() -> String {
    #[cfg(target_os = "macos")]
    {
        std::process::Command::new("sysctl")
            .args(["-n", "machdep.cpu.brand_string"])
            .output()
            .ok()
            .and_then(|o| String::from_utf8(o.stdout).ok())
            .map(|s| s.trim().to_string())
            .unwrap_or_else(|| "Unknown".to_string())
    }
    #[cfg(target_os = "linux")]
    {
        fs::read_to_string("/proc/cpuinfo")
            .ok()
            .and_then(|s| {
                s.lines()
                    .find(|l| l.starts_with("model name"))
                    .and_then(|l| l.split(':').nth(1))
                    .map(|s| s.trim().to_string())
            })
            .unwrap_or_else(|| "Unknown".to_string())
    }
    #[cfg(not(any(target_os = "macos", target_os = "linux")))]
    {
        "Unknown".to_string()
    }
}

fn detect_ram() -> String {
    #[cfg(target_os = "macos")]
    {
        std::process::Command::new("sysctl")
            .args(["-n", "hw.memsize"])
            .output()
            .ok()
            .and_then(|o| String::from_utf8(o.stdout).ok())
            .and_then(|s| s.trim().parse::<u64>().ok())
            .map(|bytes| format!("{} GB", bytes / (1024 * 1024 * 1024)))
            .unwrap_or_else(|| "Unknown".to_string())
    }
    #[cfg(target_os = "linux")]
    {
        fs::read_to_string("/proc/meminfo")
            .ok()
            .and_then(|s| {
                s.lines()
                    .find(|l| l.starts_with("MemTotal"))
                    .and_then(|l| l.split_whitespace().nth(1))
                    .and_then(|s| s.parse::<u64>().ok())
                    .map(|kb| format!("{} GB", kb / (1024 * 1024)))
            })
            .unwrap_or_else(|| "Unknown".to_string())
    }
    #[cfg(not(any(target_os = "macos", target_os = "linux")))]
    {
        "Unknown".to_string()
    }
}

fn detect_rustc_version() -> String {
    std::process::Command::new("rustc")
        .arg("--version")
        .output()
        .ok()
        .and_then(|o| String::from_utf8(o.stdout).ok())
        .map(|s| s.trim().to_string())
        .unwrap_or_else(|| "Unknown".to_string())
}

#[derive(Clone)]
struct BenchResult {
    scenario: String,
    category: String,
    times_ms: Vec<f64>,
    memory_deltas_mb: Vec<f64>,
    file_size_kb: Option<f64>,
}

impl BenchResult {
    fn min(&self) -> f64 {
        self.times_ms.iter().copied().fold(f64::INFINITY, f64::min)
    }
    fn max(&self) -> f64 {
        self.times_ms.iter().copied().fold(0.0_f64, f64::max)
    }
    fn median(&self) -> f64 {
        let mut sorted = self.times_ms.clone();
        sorted.sort_by(|a, b| a.partial_cmp(b).unwrap());
        let n = sorted.len();
        if n.is_multiple_of(2) {
            (sorted[n / 2 - 1] + sorted[n / 2]) / 2.0
        } else {
            sorted[n / 2]
        }
    }
    fn p95(&self) -> f64 {
        let mut sorted = self.times_ms.clone();
        sorted.sort_by(|a, b| a.partial_cmp(b).unwrap());
        let idx = ((sorted.len() as f64) * 0.95).ceil() as usize;
        sorted[idx.min(sorted.len() - 1)]
    }
    fn peak_mem_mb(&self) -> f64 {
        self.memory_deltas_mb
            .iter()
            .copied()
            .fold(0.0_f64, f64::max)
    }
}

fn bench<F>(scenario: &str, category: &str, output_path: Option<&Path>, make_fn: F) -> BenchResult
where
    F: Fn() -> Box<dyn FnOnce()>,
{
    // Warmup
    for _ in 0..WARMUP_RUNS {
        let f = make_fn();
        f();
        if let Some(p) = output_path {
            let _ = fs::remove_file(p);
        }
    }

    let mut times = Vec::with_capacity(BENCH_RUNS);
    let mut mem_deltas = Vec::with_capacity(BENCH_RUNS);
    let mut last_size = None;

    for _ in 0..BENCH_RUNS {
        let rss_before = current_rss_bytes();
        let f = make_fn();
        let start = Instant::now();
        f();
        let elapsed = start.elapsed().as_secs_f64() * 1000.0;
        let rss_after = current_rss_bytes();
        let mem_delta_mb = if rss_after > rss_before {
            (rss_after - rss_before) as f64 / (1024.0 * 1024.0)
        } else {
            0.0
        };

        times.push(elapsed);
        mem_deltas.push(mem_delta_mb);

        if let Some(p) = output_path {
            last_size = file_size_kb(p);
            let _ = fs::remove_file(p);
        }
    }

    let result = BenchResult {
        scenario: scenario.to_string(),
        category: category.to_string(),
        times_ms: times,
        memory_deltas_mb: mem_deltas,
        file_size_kb: last_size,
    };

    let size_str = match result.file_size_kb {
        Some(kb) => format!(" | {:.1}MB", kb / 1024.0),
        None => String::new(),
    };
    println!(
        "  [Rust    ] {:<40} med={:>8} min={:>8} max={:>8} p95={:>8} mem={:.1}MB{size_str}",
        scenario,
        format_ms(result.median()),
        format_ms(result.min()),
        format_ms(result.max()),
        format_ms(result.p95()),
        result.peak_mem_mb(),
    );

    result
}

fn cleanup(path: &Path) {
    let _ = fs::remove_file(path);
}

fn dv_config(
    sqref: String,
    vtype: ValidationType,
    op: Option<ValidationOperator>,
    f1: Option<String>,
    f2: Option<String>,
) -> DataValidationConfig {
    DataValidationConfig {
        sqref,
        validation_type: vtype,
        operator: op,
        formula1: f1,
        formula2: f2,
        allow_blank: true,
        error_style: None,
        error_title: None,
        error_message: None,
        prompt_title: None,
        prompt_message: None,
        show_input_message: false,
        show_error_message: false,
    }
}

fn bench_read_file(results: &mut Vec<BenchResult>, filename: &str, label: &str, category: &str) {
    let filepath = fixtures_dir().join(filename);
    if !filepath.exists() {
        println!("  SKIP: {filepath:?} not found. Run 'pnpm generate' in benchmarks/node first.");
        return;
    }

    println!("\n--- {label} ---");

    let scenario = format!("Read {label}");
    let fp = filepath.clone();
    results.push(bench(&scenario, category, None, move || {
        let fp = fp.clone();
        Box::new(move || {
            let wb = Workbook::open(&fp).unwrap();
            for name in wb.sheet_names() {
                let _ = wb.get_rows(name).unwrap();
            }
        })
    }));
}

fn bench_write_large_data(results: &mut Vec<BenchResult>) {
    let rows: u32 = 50_000;
    let cols: u32 = 20;
    let label = format!("Write {rows} rows x {cols} cols");
    println!("\n--- {label} ---");

    let out = output_dir().join("write-large-rust.xlsx");

    results.push(bench(&label, "Write", Some(&out), || {
        let out = output_dir().join("write-large-rust.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            let sheet = "Sheet1";
            for r in 1..=rows {
                for c in 1..=cols {
                    let cr = cell_ref(r, c);
                    if (c - 1) % 3 == 0 {
                        wb.set_cell_value(sheet, &cr, (r as f64) * (c as f64))
                            .unwrap();
                    } else if (c - 1) % 3 == 1 {
                        wb.set_cell_value(sheet, &cr, format!("R{r}C{}", c - 1))
                            .unwrap();
                    } else {
                        wb.set_cell_value(sheet, &cr, (r as f64) * ((c - 1) as f64) / 100.0)
                            .unwrap();
                    }
                }
            }
            wb.save(&out).unwrap();
        })
    }));

    cleanup(&out);
}

fn bench_write_with_styles(results: &mut Vec<BenchResult>) {
    let rows: u32 = 5_000;
    let label = format!("Write {rows} styled rows");
    println!("\n--- {label} ---");

    let out = output_dir().join("write-styles-rust.xlsx");

    results.push(bench(&label, "Write", Some(&out), || {
        let out = output_dir().join("write-styles-rust.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            let sheet = "Sheet1";

            let bold_id = wb
                .add_style(&Style {
                    font: Some(sheetkit::FontStyle {
                        bold: true,
                        size: Some(12.0),
                        name: Some("Arial".to_string()),
                        color: Some(sheetkit::StyleColor::Rgb("FFFFFFFF".to_string())),
                        ..Default::default()
                    }),
                    fill: Some(sheetkit::FillStyle {
                        pattern: sheetkit::PatternType::Solid,
                        fg_color: Some(sheetkit::StyleColor::Rgb("FF4472C4".to_string())),
                        bg_color: None,
                        gradient: None,
                    }),
                    border: Some(sheetkit::BorderStyle {
                        top: Some(sheetkit::BorderSideStyle {
                            style: sheetkit::BorderLineStyle::Thin,
                            color: Some(sheetkit::StyleColor::Rgb("FF000000".to_string())),
                        }),
                        bottom: Some(sheetkit::BorderSideStyle {
                            style: sheetkit::BorderLineStyle::Thin,
                            color: Some(sheetkit::StyleColor::Rgb("FF000000".to_string())),
                        }),
                        left: Some(sheetkit::BorderSideStyle {
                            style: sheetkit::BorderLineStyle::Thin,
                            color: Some(sheetkit::StyleColor::Rgb("FF000000".to_string())),
                        }),
                        right: Some(sheetkit::BorderSideStyle {
                            style: sheetkit::BorderLineStyle::Thin,
                            color: Some(sheetkit::StyleColor::Rgb("FF000000".to_string())),
                        }),
                        ..Default::default()
                    }),
                    alignment: Some(sheetkit::AlignmentStyle {
                        horizontal: Some(sheetkit::HorizontalAlign::Center),
                        ..Default::default()
                    }),
                    ..Default::default()
                })
                .unwrap();

            let num_id = wb
                .add_style(&Style {
                    num_fmt: Some(sheetkit::NumFmtStyle::Builtin(4)),
                    font: Some(sheetkit::FontStyle {
                        name: Some("Calibri".to_string()),
                        size: Some(11.0),
                        ..Default::default()
                    }),
                    ..Default::default()
                })
                .unwrap();

            let pct_id = wb
                .add_style(&Style {
                    num_fmt: Some(sheetkit::NumFmtStyle::Builtin(10)),
                    font: Some(sheetkit::FontStyle {
                        italic: true,
                        ..Default::default()
                    }),
                    ..Default::default()
                })
                .unwrap();

            for c in 1..=10u32 {
                let cr = cell_ref(1, c);
                wb.set_cell_value(sheet, &cr, format!("Header{c}")).unwrap();
                wb.set_cell_style(sheet, &cr, bold_id).unwrap();
            }

            for r in 2..=rows + 1 {
                for c in 1..=10u32 {
                    let cr = cell_ref(r, c);
                    if (c - 1) % 3 == 0 {
                        wb.set_cell_value(sheet, &cr, (r as f64) * ((c - 1) as f64))
                            .unwrap();
                        wb.set_cell_style(sheet, &cr, num_id).unwrap();
                    } else if (c - 1) % 3 == 1 {
                        wb.set_cell_value(sheet, &cr, format!("Data_{r}_{}", c - 1))
                            .unwrap();
                    } else {
                        wb.set_cell_value(sheet, &cr, ((r % 100) as f64) / 100.0)
                            .unwrap();
                        wb.set_cell_style(sheet, &cr, pct_id).unwrap();
                    }
                }
            }
            wb.save(&out).unwrap();
        })
    }));

    cleanup(&out);
}

fn bench_write_multi_sheet(results: &mut Vec<BenchResult>) {
    let sheets: u32 = 10;
    let rows: u32 = 5_000;
    let cols: u32 = 10;
    let label = format!("Write {sheets} sheets x {rows} rows");
    println!("\n--- {label} ---");

    let out = output_dir().join("write-multi-rust.xlsx");

    results.push(bench(&label, "Write", Some(&out), || {
        let out = output_dir().join("write-multi-rust.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            for s in 0..sheets {
                let name = if s == 0 {
                    "Sheet1".to_string()
                } else {
                    let n = format!("Sheet{}", s + 1);
                    wb.new_sheet(&n).unwrap();
                    n
                };
                for r in 1..=rows {
                    for c in 1..=cols {
                        let cr = cell_ref(r, c);
                        if (c - 1) % 2 == 0 {
                            wb.set_cell_value(&name, &cr, (r as f64) * (c as f64))
                                .unwrap();
                        } else {
                            wb.set_cell_value(&name, &cr, format!("S{s}R{r}C{}", c - 1))
                                .unwrap();
                        }
                    }
                }
            }
            wb.save(&out).unwrap();
        })
    }));

    cleanup(&out);
}

fn bench_write_formulas(results: &mut Vec<BenchResult>) {
    let rows: u32 = 10_000;
    let label = format!("Write {rows} rows with formulas");
    println!("\n--- {label} ---");

    let out = output_dir().join("write-formulas-rust.xlsx");

    results.push(bench(&label, "Write", Some(&out), || {
        let out = output_dir().join("write-formulas-rust.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            let sheet = "Sheet1";
            for r in 1..=rows {
                wb.set_cell_value(sheet, &format!("A{r}"), (r as f64) * 1.5)
                    .unwrap();
                wb.set_cell_value(sheet, &format!("B{r}"), ((r % 100) as f64) + 0.5)
                    .unwrap();
                wb.set_cell_formula(sheet, &format!("C{r}"), &format!("A{r}+B{r}"))
                    .unwrap();
                wb.set_cell_formula(sheet, &format!("D{r}"), &format!("A{r}*B{r}"))
                    .unwrap();
                wb.set_cell_formula(
                    sheet,
                    &format!("E{r}"),
                    &format!("IF(A{r}>B{r},\"A\",\"B\")"),
                )
                .unwrap();
            }
            wb.save(&out).unwrap();
        })
    }));

    cleanup(&out);
}

fn bench_write_strings(results: &mut Vec<BenchResult>) {
    let rows: u32 = 20_000;
    let label = format!("Write {rows} text-heavy rows");
    println!("\n--- {label} ---");

    let out = output_dir().join("write-strings-rust.xlsx");

    let words: &[&str] = &[
        "alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel", "india", "juliet",
    ];

    results.push(bench(&label, "Write", Some(&out), move || {
        let out = output_dir().join("write-strings-rust.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            let sheet = "Sheet1";
            for r in 1..=rows {
                let ru = r as usize;
                let w1 = words[(ru * 3) % words.len()];
                let w2 = words[(ru * 7) % words.len()];
                wb.set_cell_value(sheet, &cell_ref(r, 1), format!("{w1} {w2}"))
                    .unwrap();
                wb.set_cell_value(sheet, &cell_ref(r, 2), format!("{w1}.{w2}@example.com"))
                    .unwrap();
                wb.set_cell_value(sheet, &cell_ref(r, 3), format!("Dept_{}", r % 20))
                    .unwrap();
                wb.set_cell_value(sheet, &cell_ref(r, 4), format!("{r} {w1} Street"))
                    .unwrap();
                wb.set_cell_value(
                    sheet,
                    &cell_ref(r, 5),
                    format!("Description for {r}: {w1} {w2}"),
                )
                .unwrap();
                wb.set_cell_value(sheet, &cell_ref(r, 6), format!("City_{w1}"))
                    .unwrap();
                wb.set_cell_value(sheet, &cell_ref(r, 7), format!("Country_{w2}"))
                    .unwrap();
                wb.set_cell_value(sheet, &cell_ref(r, 8), format!("+1-555-{:04}", r % 10_000))
                    .unwrap();
                wb.set_cell_value(sheet, &cell_ref(r, 9), format!("{w1} Specialist"))
                    .unwrap();
                wb.set_cell_value(
                    sheet,
                    &cell_ref(r, 10),
                    format!("Experienced {w2} professional in {w1} domain."),
                )
                .unwrap();
            }
            wb.save(&out).unwrap();
        })
    }));

    cleanup(&out);
}

fn bench_write_data_validation(results: &mut Vec<BenchResult>) {
    let rows: u32 = 5_000;
    let label = format!("Write {rows} rows + 8 validation rules");
    println!("\n--- {label} ---");

    let out = output_dir().join("write-dv-rust.xlsx");

    results.push(bench(&label, "Write (DV)", Some(&out), || {
        let out = output_dir().join("write-dv-rust.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            let sheet = "Sheet1";

            wb.add_data_validation(
                sheet,
                &dv_config(
                    format!("A2:A{}", rows + 1),
                    ValidationType::List,
                    None,
                    Some("\"Active,Inactive,Pending,Closed\"".to_string()),
                    None,
                ),
            )
            .unwrap();
            wb.add_data_validation(
                sheet,
                &dv_config(
                    format!("B2:B{}", rows + 1),
                    ValidationType::Whole,
                    Some(ValidationOperator::Between),
                    Some("0".to_string()),
                    Some("100".to_string()),
                ),
            )
            .unwrap();
            wb.add_data_validation(
                sheet,
                &dv_config(
                    format!("C2:C{}", rows + 1),
                    ValidationType::Decimal,
                    Some(ValidationOperator::Between),
                    Some("0".to_string()),
                    Some("1".to_string()),
                ),
            )
            .unwrap();
            wb.add_data_validation(
                sheet,
                &dv_config(
                    format!("D2:D{}", rows + 1),
                    ValidationType::TextLength,
                    Some(ValidationOperator::LessThanOrEqual),
                    Some("50".to_string()),
                    None,
                ),
            )
            .unwrap();

            let statuses = ["Active", "Inactive", "Pending", "Closed"];
            for r in 2..=rows + 1 {
                let ru = r as usize;
                wb.set_cell_value(sheet, &format!("A{r}"), statuses[ru % 4])
                    .unwrap();
                wb.set_cell_value(sheet, &format!("B{r}"), (ru % 101) as f64)
                    .unwrap();
                wb.set_cell_value(sheet, &format!("C{r}"), ((ru % 100) as f64) / 100.0)
                    .unwrap();
                wb.set_cell_value(sheet, &format!("D{r}"), format!("Item_{r}"))
                    .unwrap();
            }
            wb.save(&out).unwrap();
        })
    }));

    cleanup(&out);
}

fn bench_write_comments(results: &mut Vec<BenchResult>) {
    let rows: u32 = 2_000;
    let label = format!("Write {rows} rows with comments");
    println!("\n--- {label} ---");

    let out = output_dir().join("write-comments-rust.xlsx");

    results.push(bench(&label, "Write (Comments)", Some(&out), || {
        let out = output_dir().join("write-comments-rust.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            let sheet = "Sheet1";
            for r in 1..=rows {
                wb.set_cell_value(sheet, &format!("A{r}"), r as f64)
                    .unwrap();
                wb.set_cell_value(sheet, &format!("B{r}"), format!("Name_{r}"))
                    .unwrap();
                wb.add_comment(
                    sheet,
                    &CommentConfig {
                        cell: format!("A{r}"),
                        author: "Reviewer".to_string(),
                        text: format!("Comment for row {r}: review completed."),
                    },
                )
                .unwrap();
            }
            wb.save(&out).unwrap();
        })
    }));

    cleanup(&out);
}

fn bench_write_merged_cells(results: &mut Vec<BenchResult>) {
    let regions: u32 = 500;
    let label = format!("Write {regions} merged regions");
    println!("\n--- {label} ---");

    let out = output_dir().join("write-merge-rust.xlsx");

    results.push(bench(&label, "Write (Merge)", Some(&out), || {
        let out = output_dir().join("write-merge-rust.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            let sheet = "Sheet1";
            for i in 0..regions {
                let row = i * 3 + 1;
                wb.set_cell_value(sheet, &format!("A{row}"), format!("Section {}", i + 1))
                    .unwrap();
                wb.set_cell_value(sheet, &format!("A{}", row + 1), (i * 100) as f64)
                    .unwrap();
                wb.set_cell_value(sheet, &format!("B{}", row + 1), format!("Data_{i}"))
                    .unwrap();
                wb.merge_cells(sheet, &format!("A{row}"), &format!("D{row}"))
                    .unwrap();
            }
            wb.save(&out).unwrap();
        })
    }));

    cleanup(&out);
}

fn bench_write_scale(results: &mut Vec<BenchResult>, rows: u32) {
    let cols: u32 = 10;
    let tag = if rows >= 1000 {
        format!("{}k", rows / 1000)
    } else {
        format!("{rows}")
    };
    let label = format!("Write {tag} rows x {cols} cols");
    println!("\n--- {label} ---");

    let out = output_dir().join(format!("write-scale-{tag}-rust.xlsx"));
    let out_for_bench = out.clone();

    results.push(bench(
        &label,
        "Write (Scale)",
        Some(&out_for_bench),
        move || {
            let tag2 = if rows >= 1000 {
                format!("{}k", rows / 1000)
            } else {
                format!("{rows}")
            };
            let out = output_dir().join(format!("write-scale-{tag2}-rust.xlsx"));
            Box::new(move || {
                let mut wb = Workbook::new();
                let sheet = "Sheet1";
                for r in 1..=rows {
                    for c in 1..=cols {
                        let cr = cell_ref(r, c);
                        if (c - 1) % 3 == 0 {
                            wb.set_cell_value(sheet, &cr, (r as f64) * (c as f64))
                                .unwrap();
                        } else if (c - 1) % 3 == 1 {
                            wb.set_cell_value(sheet, &cr, format!("R{r}C{}", c - 1))
                                .unwrap();
                        } else {
                            wb.set_cell_value(sheet, &cr, (r as f64) * ((c - 1) as f64) / 100.0)
                                .unwrap();
                        }
                    }
                }
                wb.save(&out).unwrap();
            })
        },
    ));

    cleanup(&out);
}

fn bench_buffer_round_trip(results: &mut Vec<BenchResult>) {
    let rows: u32 = 10_000;
    let cols: u32 = 10;
    let label = format!("Buffer round-trip ({rows} rows)");
    println!("\n--- {label} ---");

    results.push(bench(&label, "Round-Trip", None, move || {
        Box::new(move || {
            let mut wb = Workbook::new();
            let sheet = "Sheet1";
            for r in 1..=rows {
                for c in 1..=cols {
                    let cr = cell_ref(r, c);
                    wb.set_cell_value(sheet, &cr, (r as f64) * (c as f64))
                        .unwrap();
                }
            }
            let buf = wb.save_to_buffer().unwrap();
            let wb2 = Workbook::open_from_buffer(&buf).unwrap();
            let _ = wb2.get_rows("Sheet1").unwrap();
        })
    }));
}

fn bench_streaming_write(results: &mut Vec<BenchResult>) {
    let rows: u32 = 50_000;
    let cols: u32 = 20;
    let label = format!("Streaming write ({rows} rows)");
    println!("\n--- {label} ---");

    let out = output_dir().join("stream-rust.xlsx");

    results.push(bench(&label, "Streaming", Some(&out), move || {
        let out = output_dir().join("stream-rust.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            let mut sw = wb.new_stream_writer("StreamSheet").unwrap();
            for r in 1..=rows {
                let mut vals = Vec::with_capacity(cols as usize);
                for c in 1..=cols {
                    if (c - 1) % 3 == 0 {
                        vals.push(CellValue::Number((r as f64) * (c as f64)));
                    } else if (c - 1) % 3 == 1 {
                        vals.push(CellValue::String(format!("R{r}C{}", c - 1)));
                    } else {
                        vals.push(CellValue::Number((r as f64) * ((c - 1) as f64) / 100.0));
                    }
                }
                sw.write_row(r, &vals).unwrap();
            }
            wb.apply_stream_writer(sw).unwrap();
            wb.save(&out).unwrap();
        })
    }));

    cleanup(&out);
}

fn random_cell_addresses(count: usize) -> Vec<String> {
    let mut cells = Vec::with_capacity(count);
    let mut seed: u64 = 42;
    for _ in 0..count {
        seed = seed
            .wrapping_mul(6364136223846793005)
            .wrapping_add(1442695040888963407);
        let r = ((seed >> 33) % 50_000) as u32 + 2;
        seed = seed
            .wrapping_mul(6364136223846793005)
            .wrapping_add(1442695040888963407);
        let c = ((seed >> 33) % 20) as u32 + 1;
        cells.push(cell_ref(r, c));
    }
    cells
}

fn bench_random_access_read(results: &mut Vec<BenchResult>) {
    let filepath = fixtures_dir().join("large-data.xlsx");
    if !filepath.exists() {
        println!("  SKIP: large-data.xlsx not found. Run pnpm generate first.");
        return;
    }

    let lookups: usize = 1_000;
    let cells = random_cell_addresses(lookups);

    // Open+lookup: measures file open overhead together with cell lookups.
    let label_open = format!("Random-access (open+{lookups} lookups)");
    println!("\n--- {label_open} ---");

    let fp = filepath.clone();
    let cells_clone = cells.clone();
    results.push(bench(&label_open, "Random Access", None, move || {
        let fp = fp.clone();
        let cells_clone = cells_clone.clone();
        Box::new(move || {
            let wb = Workbook::open(&fp).unwrap();
            for cell in &cells_clone {
                let _ = wb.get_cell_value("Sheet1", cell);
            }
        })
    }));

    // Lookup-only: opens a fresh workbook per run inside `make_fn` (not timed),
    // then the returned closure only performs cell lookups (timed). Each run gets
    // its own Workbook instance so internal caches from previous runs do not
    // accumulate and distort the measurement.
    let label_lookup = format!("Random-access (lookup-only, {lookups} cells)");
    println!("\n--- {label_lookup} ---");

    let fp = filepath.clone();
    results.push(bench(&label_lookup, "Random Access", None, move || {
        let wb = Workbook::open(&fp).unwrap();
        let cells = cells.clone();
        Box::new(move || {
            for cell in &cells {
                let _ = wb.get_cell_value("Sheet1", cell);
            }
        })
    }));
}

fn bench_mixed_workload_write(results: &mut Vec<BenchResult>) {
    let label = "Mixed workload write (ERP-style)";
    println!("\n--- {label} ---");

    let out = output_dir().join("write-mixed-rust.xlsx");

    results.push(bench(label, "Mixed Write", Some(&out), || {
        let out = output_dir().join("write-mixed-rust.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            let sheet = "Sheet1";

            let bold_id = wb
                .add_style(&Style {
                    font: Some(sheetkit::FontStyle {
                        bold: true,
                        size: Some(11.0),
                        ..Default::default()
                    }),
                    fill: Some(sheetkit::FillStyle {
                        pattern: sheetkit::PatternType::Solid,
                        fg_color: Some(sheetkit::StyleColor::Rgb("FFD9E2F3".to_string())),
                        bg_color: None,
                        gradient: None,
                    }),
                    ..Default::default()
                })
                .unwrap();

            let num_id = wb
                .add_style(&Style {
                    num_fmt: Some(sheetkit::NumFmtStyle::Custom("$#,##0.00".to_string())),
                    ..Default::default()
                })
                .unwrap();

            wb.set_cell_value(sheet, "A1", "Sales Report Q4").unwrap();
            wb.merge_cells(sheet, "A1", "F1").unwrap();
            wb.set_cell_style(sheet, "A1", bold_id).unwrap();

            let headers = [
                "Order_ID",
                "Product",
                "Quantity",
                "Unit_Price",
                "Total",
                "Region",
            ];
            for (c, h) in headers.iter().enumerate() {
                let cr = cell_ref(2, (c + 1) as u32);
                wb.set_cell_value(sheet, &cr, *h).unwrap();
                wb.set_cell_style(sheet, &cr, bold_id).unwrap();
            }

            wb.add_data_validation(
                sheet,
                &dv_config(
                    "F3:F5002".to_string(),
                    ValidationType::List,
                    None,
                    Some("\"North,South,East,West\"".to_string()),
                    None,
                ),
            )
            .unwrap();
            wb.add_data_validation(
                sheet,
                &dv_config(
                    "C3:C5002".to_string(),
                    ValidationType::Whole,
                    Some(ValidationOperator::GreaterThan),
                    Some("0".to_string()),
                    None,
                ),
            )
            .unwrap();

            let regions = ["North", "South", "East", "West"];
            let products = ["Widget A", "Widget B", "Gadget X", "Gadget Y", "Service Z"];

            for i in 1u32..=5000 {
                let r = i + 2;
                wb.set_cell_value(sheet, &format!("A{r}"), format!("ORD-{i:05}"))
                    .unwrap();
                wb.set_cell_value(
                    sheet,
                    &format!("B{r}"),
                    products[(i as usize) % products.len()],
                )
                .unwrap();
                wb.set_cell_value(sheet, &format!("C{r}"), ((i % 50) + 1) as f64)
                    .unwrap();
                wb.set_cell_value(sheet, &format!("D{r}"), (((i * 19) % 500) + 10) as f64)
                    .unwrap();
                wb.set_cell_style(sheet, &format!("D{r}"), num_id).unwrap();
                wb.set_cell_formula(sheet, &format!("E{r}"), &format!("C{r}*D{r}"))
                    .unwrap();
                wb.set_cell_style(sheet, &format!("E{r}"), num_id).unwrap();
                wb.set_cell_value(
                    sheet,
                    &format!("F{r}"),
                    regions[(i as usize) % regions.len()],
                )
                .unwrap();

                if i % 50 == 0 {
                    wb.add_comment(
                        sheet,
                        &CommentConfig {
                            cell: format!("A{r}"),
                            author: "Sales".to_string(),
                            text: format!("Bulk order - special pricing applied for order {i}."),
                        },
                    )
                    .unwrap();
                }
            }

            wb.save(&out).unwrap();
        })
    }));

    cleanup(&out);
}

fn print_summary_table(results: &[BenchResult]) {
    println!("\n\n========================================");
    println!(" RUST BENCHMARK RESULTS SUMMARY");
    println!(" ({BENCH_RUNS} runs per scenario, {WARMUP_RUNS} warmup)");
    println!("========================================\n");

    println!(
        "| {:<50}| {:>8} | {:>8} | {:>8} | {:>8} | {:>8} |",
        "Scenario", "Median", "Min", "Max", "P95", "Mem(MB)"
    );
    println!(
        "|{:-<51}|{:-<10}|{:-<10}|{:-<10}|{:-<10}|{:-<10}|",
        "", "", "", "", "", ""
    );

    for r in results {
        println!(
            "| {:<50}| {:>8} | {:>8} | {:>8} | {:>8} | {:>8.1} |",
            r.scenario,
            format_ms(r.median()),
            format_ms(r.min()),
            format_ms(r.max()),
            format_ms(r.p95()),
            r.peak_mem_mb(),
        );
    }
}

fn generate_markdown_report(results: &[BenchResult]) -> String {
    let mut lines = Vec::new();
    lines.push("# SheetKit Rust Native Benchmark".to_string());
    lines.push(String::new());
    lines.push(format!("Benchmark run: {}", chrono_free_timestamp()));
    lines.push(String::new());
    lines.push("## Environment".to_string());
    lines.push(String::new());
    lines.push("| Item | Value |".to_string());
    lines.push("|------|-------|".to_string());
    lines.push(format!("| CPU | {} |", detect_cpu()));
    lines.push(format!("| RAM | {} |", detect_ram()));
    lines.push(format!(
        "| OS | {} {} |",
        std::env::consts::OS,
        std::env::consts::ARCH
    ));
    lines.push(format!("| Rust | {} |", detect_rustc_version()));
    lines.push("| Profile | release (opt-level=3, LTO=fat) |".to_string());
    lines.push(format!(
        "| Iterations | {BENCH_RUNS} runs per scenario, {WARMUP_RUNS} warmup |"
    ));
    lines.push(String::new());

    let categories: Vec<String> = {
        let mut seen = Vec::new();
        for r in results {
            if !seen.contains(&r.category) {
                seen.push(r.category.clone());
            }
        }
        seen
    };

    for cat in &categories {
        let subset: Vec<&BenchResult> = results.iter().filter(|r| &r.category == cat).collect();
        if subset.is_empty() {
            continue;
        }
        lines.push(format!("## {cat}"));
        lines.push(String::new());
        lines.push("| Scenario | Median | Min | Max | P95 | Peak Mem (MB) |".to_string());
        lines.push("|----------|--------|-----|-----|-----|---------------|".to_string());
        for r in &subset {
            lines.push(format!(
                "| {} | {} | {} | {} | {} | {:.1} |",
                r.scenario,
                format_ms(r.median()),
                format_ms(r.min()),
                format_ms(r.max()),
                format_ms(r.p95()),
                r.peak_mem_mb(),
            ));
        }
        lines.push(String::new());
    }

    lines.join("\n")
}

fn chrono_free_timestamp() -> String {
    let dur = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap();
    let secs = dur.as_secs();
    let days = secs / 86400;
    let remaining = secs % 86400;
    let hours = remaining / 3600;
    let minutes = (remaining % 3600) / 60;
    let seconds = remaining % 60;
    let (year, month, day) = days_to_date(days);
    format!("{year:04}-{month:02}-{day:02}T{hours:02}:{minutes:02}:{seconds:02}Z")
}

fn days_to_date(days_since_epoch: u64) -> (u64, u64, u64) {
    let z = days_since_epoch + 719468;
    let era = z / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if m <= 2 { y + 1 } else { y };
    (y, m, d)
}

fn main() {
    println!("SheetKit Rust Native Benchmark");
    println!(
        "Platform: {} {} | Profile: {}",
        std::env::consts::OS,
        std::env::consts::ARCH,
        if cfg!(debug_assertions) {
            "debug (WARNING: run with --release!)"
        } else {
            "release"
        }
    );
    println!("Iterations: {BENCH_RUNS} runs per scenario, {WARMUP_RUNS} warmup");
    println!();

    fs::create_dir_all(output_dir()).unwrap();

    let mut results = Vec::new();

    // READ benchmarks
    println!("=== READ BENCHMARKS ===");
    bench_read_file(
        &mut results,
        "large-data.xlsx",
        "Large Data (50k rows x 20 cols)",
        "Read",
    );
    bench_read_file(
        &mut results,
        "heavy-styles.xlsx",
        "Heavy Styles (5k rows, formatted)",
        "Read",
    );
    bench_read_file(
        &mut results,
        "multi-sheet.xlsx",
        "Multi-Sheet (10 sheets x 5k rows)",
        "Read",
    );
    bench_read_file(&mut results, "formulas.xlsx", "Formulas (10k rows)", "Read");
    bench_read_file(
        &mut results,
        "strings.xlsx",
        "Strings (20k rows text-heavy)",
        "Read",
    );
    bench_read_file(
        &mut results,
        "data-validation.xlsx",
        "Data Validation (5k rows, 8 rules)",
        "Read",
    );
    bench_read_file(
        &mut results,
        "comments.xlsx",
        "Comments (2k rows with comments)",
        "Read",
    );
    bench_read_file(
        &mut results,
        "merged-cells.xlsx",
        "Merged Cells (500 regions)",
        "Read",
    );
    bench_read_file(
        &mut results,
        "mixed-workload.xlsx",
        "Mixed Workload (ERP document)",
        "Read",
    );

    // READ scaling
    println!("\n\n=== READ SCALING ===");
    bench_read_file(
        &mut results,
        "scale-1k.xlsx",
        "Scale 1k rows",
        "Read (Scale)",
    );
    bench_read_file(
        &mut results,
        "scale-10k.xlsx",
        "Scale 10k rows",
        "Read (Scale)",
    );
    bench_read_file(
        &mut results,
        "scale-100k.xlsx",
        "Scale 100k rows",
        "Read (Scale)",
    );

    // WRITE benchmarks
    println!("\n\n=== WRITE BENCHMARKS ===");
    bench_write_large_data(&mut results);
    bench_write_with_styles(&mut results);
    bench_write_multi_sheet(&mut results);
    bench_write_formulas(&mut results);
    bench_write_strings(&mut results);
    bench_write_data_validation(&mut results);
    bench_write_comments(&mut results);
    bench_write_merged_cells(&mut results);

    // WRITE scaling
    println!("\n\n=== WRITE SCALING ===");
    bench_write_scale(&mut results, 1_000);
    bench_write_scale(&mut results, 10_000);
    bench_write_scale(&mut results, 50_000);
    bench_write_scale(&mut results, 100_000);

    // Buffer round-trip
    println!("\n\n=== BUFFER ROUND-TRIP ===");
    bench_buffer_round_trip(&mut results);

    // Streaming
    println!("\n\n=== STREAMING ===");
    bench_streaming_write(&mut results);

    // Random access
    println!("\n\n=== RANDOM ACCESS ===");
    bench_random_access_read(&mut results);

    // Mixed workload write
    println!("\n\n=== MIXED WORKLOAD ===");
    bench_mixed_workload_write(&mut results);

    // Summary
    print_summary_table(&results);

    // Write markdown report
    let report = generate_markdown_report(&results);
    let report_path = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("RESULTS.md");
    let mut file = fs::File::create(&report_path).unwrap();
    file.write_all(report.as_bytes()).unwrap();
    println!("\nMarkdown report written to: {report_path:?}");
}
