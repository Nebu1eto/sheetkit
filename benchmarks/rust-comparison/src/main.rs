/// Rust ecosystem comparison benchmark.
///
/// Compares SheetKit against other Rust Excel libraries:
///   - calamine (read-only)
///   - rust_xlsxwriter (write-only)
///   - edit-xlsx (read/modify/write)
///
/// Mirrors the scenario structure of the existing Node.js and Rust benchmarks.
use sheetkit::utils::column_number_to_name;
use sheetkit::{CellValue, Style, Workbook};

use std::fs;
use std::io::Write as _;
use std::path::{Path, PathBuf};
use std::time::Instant;

const WARMUP_RUNS: usize = 1;
const BENCH_RUNS: usize = 5;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// BenchResult
// ---------------------------------------------------------------------------

#[derive(Clone)]
struct BenchResult {
    scenario: String,
    library: String,
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

// ---------------------------------------------------------------------------
// Bench harness
// ---------------------------------------------------------------------------

fn bench<F>(
    scenario: &str,
    library: &str,
    category: &str,
    output_path: Option<&Path>,
    make_fn: F,
) -> BenchResult
where
    F: Fn() -> Box<dyn FnOnce()>,
{
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
        library: library.to_string(),
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
        "  [{:<14}] {:<40} med={:>8} min={:>8} max={:>8} p95={:>8} mem={:.1}MB{size_str}",
        library,
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

// ===================================================================
// READ BENCHMARKS: SheetKit vs calamine
// ===================================================================

fn bench_read_file(results: &mut Vec<BenchResult>, filename: &str, label: &str, category: &str) {
    let filepath = fixtures_dir().join(filename);
    if !filepath.exists() {
        println!("  SKIP: {filepath:?} not found. Run 'pnpm generate' in benchmarks/node first.");
        return;
    }

    println!("\n--- Read: {label} ---");

    // SheetKit
    let fp = filepath.clone();
    results.push(bench(
        &format!("Read {label}"),
        "SheetKit",
        category,
        None,
        move || {
            let fp = fp.clone();
            Box::new(move || {
                let wb = Workbook::open(&fp).unwrap();
                for name in wb.sheet_names() {
                    let _ = wb.get_rows(name).unwrap();
                }
            })
        },
    ));

    // calamine
    let fp = filepath.clone();
    results.push(bench(
        &format!("Read {label}"),
        "calamine",
        category,
        None,
        move || {
            let fp = fp.clone();
            Box::new(move || {
                use calamine::{open_workbook, Reader, Xlsx};
                let mut wb: Xlsx<_> = open_workbook(&fp).unwrap();
                let names: Vec<String> = wb.sheet_names().to_vec();
                for name in &names {
                    if let Ok(range) = wb.worksheet_range(name) {
                        for _row in range.rows() {
                            // iterate all rows
                        }
                    }
                }
            })
        },
    ));

    // edit-xlsx (read support)
    let fp = filepath.clone();
    results.push(bench(
        &format!("Read {label}"),
        "edit-xlsx",
        category,
        None,
        move || {
            let fp = fp.clone();
            Box::new(move || {
                use edit_xlsx::{Read as _, Workbook as EditWorkbook};
                let wb = EditWorkbook::from_path(&fp).unwrap();
                // edit-xlsx worksheets are 1-indexed; iterate up to a
                // reasonable count (stop on first error).
                for id in 1..=100u32 {
                    match wb.get_worksheet(id) {
                        Ok(ws) => {
                            let max_row = ws.max_row();
                            let max_col = ws.max_column();
                            for r in 1..=max_row {
                                for c in 1..=max_col {
                                    let _ = ws.read_cell((r, c));
                                }
                            }
                        }
                        Err(_) => break,
                    }
                }
            })
        },
    ));
}

// ===================================================================
// WRITE BENCHMARKS: SheetKit vs rust_xlsxwriter vs edit-xlsx
// ===================================================================

fn bench_write_large_data(results: &mut Vec<BenchResult>) {
    let rows: u32 = 50_000;
    let cols: u32 = 20;
    let label = format!("Write {rows} rows x {cols} cols");
    println!("\n--- {label} ---");

    // SheetKit
    let out = output_dir().join("cmp-write-large-sheetkit.xlsx");
    results.push(bench(&label, "SheetKit", "Write", Some(&out), || {
        let out = output_dir().join("cmp-write-large-sheetkit.xlsx");
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

    // rust_xlsxwriter
    let out = output_dir().join("cmp-write-large-xlsxwriter.xlsx");
    results.push(bench(
        &label,
        "rust_xlsxwriter",
        "Write",
        Some(&out),
        || {
            let out = output_dir().join("cmp-write-large-xlsxwriter.xlsx");
            Box::new(move || {
                let mut wb = rust_xlsxwriter::Workbook::new();
                let ws = wb.add_worksheet();
                for r in 0..rows {
                    for c in 0..cols {
                        if c % 3 == 0 {
                            ws.write_number(r, c as u16, ((r + 1) as f64) * ((c + 1) as f64))
                                .unwrap();
                        } else if c % 3 == 1 {
                            ws.write_string(r, c as u16, format!("R{}C{c}", r + 1))
                                .unwrap();
                        } else {
                            ws.write_number(r, c as u16, ((r + 1) as f64) * (c as f64) / 100.0)
                                .unwrap();
                        }
                    }
                }
                wb.save(&out).unwrap();
            })
        },
    ));
    cleanup(&out);

    // edit-xlsx
    let out = output_dir().join("cmp-write-large-editxlsx.xlsx");
    results.push(bench(&label, "edit-xlsx", "Write", Some(&out), || {
        let out = output_dir().join("cmp-write-large-editxlsx.xlsx");
        Box::new(move || {
            use edit_xlsx::{Workbook as EditWorkbook, Write as _};
            let mut wb = EditWorkbook::new();
            let ws = wb.get_worksheet_mut(1).unwrap();
            for r in 1..=rows {
                for c in 1..=cols {
                    if (c - 1) % 3 == 0 {
                        ws.write((r, c), (r as f64) * (c as f64)).unwrap();
                    } else if (c - 1) % 3 == 1 {
                        ws.write((r, c), format!("R{r}C{}", c - 1)).unwrap();
                    } else {
                        ws.write((r, c), (r as f64) * ((c - 1) as f64) / 100.0)
                            .unwrap();
                    }
                }
            }
            wb.save_as(&out).unwrap();
        })
    }));
    cleanup(&out);
}

fn bench_write_with_styles(results: &mut Vec<BenchResult>) {
    let rows: u32 = 5_000;
    let label = format!("Write {rows} styled rows");
    println!("\n--- {label} ---");

    // SheetKit
    let out = output_dir().join("cmp-write-styles-sheetkit.xlsx");
    results.push(bench(&label, "SheetKit", "Write", Some(&out), || {
        let out = output_dir().join("cmp-write-styles-sheetkit.xlsx");
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

    // rust_xlsxwriter
    let out = output_dir().join("cmp-write-styles-xlsxwriter.xlsx");
    results.push(bench(
        &label,
        "rust_xlsxwriter",
        "Write",
        Some(&out),
        || {
            let out = output_dir().join("cmp-write-styles-xlsxwriter.xlsx");
            Box::new(move || {
                use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder};
                let mut wb = rust_xlsxwriter::Workbook::new();
                let bold_fmt = Format::new()
                    .set_bold()
                    .set_font_size(12)
                    .set_font_name("Arial")
                    .set_font_color(Color::White)
                    .set_background_color(Color::RGB(0x4472C4))
                    .set_border(FormatBorder::Thin)
                    .set_border_color(Color::Black)
                    .set_align(FormatAlign::Center);
                let num_fmt = Format::new()
                    .set_num_format("#,##0.00")
                    .set_font_name("Calibri")
                    .set_font_size(11);
                let pct_fmt = Format::new().set_num_format("0.00%").set_italic();

                let ws = wb.add_worksheet();
                for c in 0..10u16 {
                    ws.write_string_with_format(0, c, format!("Header{}", c + 1), &bold_fmt)
                        .unwrap();
                }
                for r in 1..=rows {
                    for c in 0..10u32 {
                        if c % 3 == 0 {
                            ws.write_number_with_format(
                                r,
                                c as u16,
                                (r as f64) * (c as f64),
                                &num_fmt,
                            )
                            .unwrap();
                        } else if c % 3 == 1 {
                            ws.write_string(r, c as u16, format!("Data_{}{c}", r + 1))
                                .unwrap();
                        } else {
                            ws.write_number_with_format(
                                r,
                                c as u16,
                                ((r % 100) as f64) / 100.0,
                                &pct_fmt,
                            )
                            .unwrap();
                        }
                    }
                }
                wb.save(&out).unwrap();
            })
        },
    ));
    cleanup(&out);

    // edit-xlsx
    let out = output_dir().join("cmp-write-styles-editxlsx.xlsx");
    results.push(bench(&label, "edit-xlsx", "Write", Some(&out), || {
        let out = output_dir().join("cmp-write-styles-editxlsx.xlsx");
        Box::new(move || {
            use edit_xlsx::{
                Format as EditFormat, FormatAlignType, FormatBorderType, FormatColor,
                Workbook as EditWorkbook, Write as _,
            };
            let mut wb = EditWorkbook::new();
            let ws = wb.get_worksheet_mut(1).unwrap();
            let bold_fmt = EditFormat::default()
                .set_bold()
                .set_size(12u8)
                .set_font("Arial")
                .set_color(FormatColor::RGB(255, 255, 255))
                .set_background_color(FormatColor::RGB(0x44, 0x72, 0xC4))
                .set_border(FormatBorderType::Thin)
                .set_align(FormatAlignType::Center);
            let num_fmt = EditFormat::default().set_font("Calibri").set_size(11u8);
            let pct_fmt = EditFormat::default().set_italic();
            for c in 1..=10u32 {
                ws.write_with_format((1, c), format!("Header{c}"), &bold_fmt)
                    .unwrap();
            }
            for r in 2..=rows + 1 {
                for c in 1..=10u32 {
                    if (c - 1) % 3 == 0 {
                        ws.write_with_format((r, c), (r as f64) * ((c - 1) as f64), &num_fmt)
                            .unwrap();
                    } else if (c - 1) % 3 == 1 {
                        ws.write((r, c), format!("Data_{r}_{}", c - 1)).unwrap();
                    } else {
                        ws.write_with_format((r, c), ((r % 100) as f64) / 100.0, &pct_fmt)
                            .unwrap();
                    }
                }
            }
            wb.save_as(&out).unwrap();
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

    // SheetKit
    let out = output_dir().join("cmp-write-multi-sheetkit.xlsx");
    results.push(bench(&label, "SheetKit", "Write", Some(&out), || {
        let out = output_dir().join("cmp-write-multi-sheetkit.xlsx");
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

    // rust_xlsxwriter
    let out = output_dir().join("cmp-write-multi-xlsxwriter.xlsx");
    results.push(bench(
        &label,
        "rust_xlsxwriter",
        "Write",
        Some(&out),
        || {
            let out = output_dir().join("cmp-write-multi-xlsxwriter.xlsx");
            Box::new(move || {
                let mut wb = rust_xlsxwriter::Workbook::new();
                for s in 0..sheets {
                    let ws = wb.add_worksheet();
                    ws.set_name(format!("Sheet{}", s + 1)).unwrap();
                    for r in 0..rows {
                        for c in 0..cols {
                            if c % 2 == 0 {
                                ws.write_number(r, c as u16, ((r + 1) as f64) * ((c + 1) as f64))
                                    .unwrap();
                            } else {
                                ws.write_string(r, c as u16, format!("S{s}R{}C{c}", r + 1))
                                    .unwrap();
                            }
                        }
                    }
                }
                wb.save(&out).unwrap();
            })
        },
    ));
    cleanup(&out);

    // edit-xlsx
    let out = output_dir().join("cmp-write-multi-editxlsx.xlsx");
    results.push(bench(&label, "edit-xlsx", "Write", Some(&out), || {
        let out = output_dir().join("cmp-write-multi-editxlsx.xlsx");
        Box::new(move || {
            use edit_xlsx::{Workbook as EditWorkbook, Write as _};
            let mut wb = EditWorkbook::new();
            for s in 0..sheets {
                let ws = if s == 0 {
                    wb.get_worksheet_mut(1).unwrap()
                } else {
                    wb.add_worksheet_by_name(&format!("Sheet{}", s + 1))
                        .unwrap()
                };
                for r in 1..=rows {
                    for c in 1..=cols {
                        if (c - 1) % 2 == 0 {
                            ws.write((r, c), (r as f64) * (c as f64)).unwrap();
                        } else {
                            ws.write((r, c), format!("S{s}R{r}C{}", c - 1)).unwrap();
                        }
                    }
                }
            }
            wb.save_as(&out).unwrap();
        })
    }));
    cleanup(&out);
}

fn bench_write_formulas(results: &mut Vec<BenchResult>) {
    let rows: u32 = 10_000;
    let label = format!("Write {rows} rows with formulas");
    println!("\n--- {label} ---");

    // SheetKit
    let out = output_dir().join("cmp-write-formulas-sheetkit.xlsx");
    results.push(bench(&label, "SheetKit", "Write", Some(&out), || {
        let out = output_dir().join("cmp-write-formulas-sheetkit.xlsx");
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

    // rust_xlsxwriter
    let out = output_dir().join("cmp-write-formulas-xlsxwriter.xlsx");
    results.push(bench(
        &label,
        "rust_xlsxwriter",
        "Write",
        Some(&out),
        || {
            let out = output_dir().join("cmp-write-formulas-xlsxwriter.xlsx");
            Box::new(move || {
                let mut wb = rust_xlsxwriter::Workbook::new();
                let ws = wb.add_worksheet();
                for r in 0..rows {
                    let row = r + 1; // 1-based for formula references
                    ws.write_number(r, 0, ((r + 1) as f64) * 1.5).unwrap();
                    ws.write_number(r, 1, (((r + 1) % 100) as f64) + 0.5)
                        .unwrap();
                    ws.write_formula(r, 2, format!("A{row}+B{row}").as_str())
                        .unwrap();
                    ws.write_formula(r, 3, format!("A{row}*B{row}").as_str())
                        .unwrap();
                    ws.write_formula(r, 4, format!("IF(A{row}>B{row},\"A\",\"B\")").as_str())
                        .unwrap();
                }
                wb.save(&out).unwrap();
            })
        },
    ));
    cleanup(&out);

    // edit-xlsx
    let out = output_dir().join("cmp-write-formulas-editxlsx.xlsx");
    results.push(bench(&label, "edit-xlsx", "Write", Some(&out), || {
        let out = output_dir().join("cmp-write-formulas-editxlsx.xlsx");
        Box::new(move || {
            use edit_xlsx::{Workbook as EditWorkbook, Write as _};
            let mut wb = EditWorkbook::new();
            let ws = wb.get_worksheet_mut(1).unwrap();
            for r in 1..=rows {
                ws.write((r, 1), (r as f64) * 1.5).unwrap();
                ws.write((r, 2), ((r % 100) as f64) + 0.5).unwrap();
                ws.write_formula((r, 3), &format!("A{r}+B{r}")).unwrap();
                ws.write_formula((r, 4), &format!("A{r}*B{r}")).unwrap();
                ws.write_formula((r, 5), &format!("IF(A{r}>B{r},\"A\",\"B\")"))
                    .unwrap();
            }
            wb.save_as(&out).unwrap();
        })
    }));
    cleanup(&out);
}

fn bench_write_strings(results: &mut Vec<BenchResult>) {
    let rows: u32 = 20_000;
    let label = format!("Write {rows} text-heavy rows");
    println!("\n--- {label} ---");

    let words: &[&str] = &[
        "alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel", "india", "juliet",
    ];

    // SheetKit
    let out = output_dir().join("cmp-write-strings-sheetkit.xlsx");
    results.push(bench(&label, "SheetKit", "Write", Some(&out), move || {
        let out = output_dir().join("cmp-write-strings-sheetkit.xlsx");
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
            }
            wb.save(&out).unwrap();
        })
    }));
    cleanup(&out);

    // rust_xlsxwriter
    let out = output_dir().join("cmp-write-strings-xlsxwriter.xlsx");
    results.push(bench(
        &label,
        "rust_xlsxwriter",
        "Write",
        Some(&out),
        move || {
            let out = output_dir().join("cmp-write-strings-xlsxwriter.xlsx");
            Box::new(move || {
                let mut wb = rust_xlsxwriter::Workbook::new();
                let ws = wb.add_worksheet();
                for r in 0..rows {
                    let ru = (r + 1) as usize;
                    let w1 = words[(ru * 3) % words.len()];
                    let w2 = words[(ru * 7) % words.len()];
                    ws.write_string(r, 0, format!("{w1} {w2}")).unwrap();
                    ws.write_string(r, 1, format!("{w1}.{w2}@example.com"))
                        .unwrap();
                    ws.write_string(r, 2, format!("Dept_{}", (r + 1) % 20))
                        .unwrap();
                    ws.write_string(r, 3, format!("{} {w1} Street", r + 1))
                        .unwrap();
                    ws.write_string(r, 4, format!("Description for {}: {w1} {w2}", r + 1))
                        .unwrap();
                }
                wb.save(&out).unwrap();
            })
        },
    ));
    cleanup(&out);

    // edit-xlsx
    let out = output_dir().join("cmp-write-strings-editxlsx.xlsx");
    results.push(bench(&label, "edit-xlsx", "Write", Some(&out), move || {
        let out = output_dir().join("cmp-write-strings-editxlsx.xlsx");
        Box::new(move || {
            use edit_xlsx::{Workbook as EditWorkbook, Write as _};
            let mut wb = EditWorkbook::new();
            let ws = wb.get_worksheet_mut(1).unwrap();
            for r in 1..=rows {
                let ru = r as usize;
                let w1 = words[(ru * 3) % words.len()];
                let w2 = words[(ru * 7) % words.len()];
                ws.write((r, 1), format!("{w1} {w2}")).unwrap();
                ws.write((r, 2), format!("{w1}.{w2}@example.com")).unwrap();
                ws.write((r, 3), format!("Dept_{}", r % 20)).unwrap();
                ws.write((r, 4), format!("{r} {w1} Street")).unwrap();
                ws.write((r, 5), format!("Description for {r}: {w1} {w2}"))
                    .unwrap();
            }
            wb.save_as(&out).unwrap();
        })
    }));
    cleanup(&out);
}

fn bench_write_merged_cells(results: &mut Vec<BenchResult>) {
    let regions: u32 = 500;
    let label = format!("Write {regions} merged regions");
    println!("\n--- {label} ---");

    // SheetKit
    let out = output_dir().join("cmp-write-merge-sheetkit.xlsx");
    results.push(bench(&label, "SheetKit", "Write", Some(&out), || {
        let out = output_dir().join("cmp-write-merge-sheetkit.xlsx");
        Box::new(move || {
            let mut wb = Workbook::new();
            let sheet = "Sheet1";
            for i in 0..regions {
                let row = i * 3 + 1;
                wb.set_cell_value(sheet, &format!("A{row}"), format!("Section {}", i + 1))
                    .unwrap();
                wb.set_cell_value(sheet, &format!("A{}", row + 1), (i * 100) as f64)
                    .unwrap();
                wb.merge_cells(sheet, &format!("A{row}"), &format!("D{row}"))
                    .unwrap();
            }
            wb.save(&out).unwrap();
        })
    }));
    cleanup(&out);

    // rust_xlsxwriter
    let out = output_dir().join("cmp-write-merge-xlsxwriter.xlsx");
    results.push(bench(
        &label,
        "rust_xlsxwriter",
        "Write",
        Some(&out),
        || {
            let out = output_dir().join("cmp-write-merge-xlsxwriter.xlsx");
            Box::new(move || {
                let mut wb = rust_xlsxwriter::Workbook::new();
                let ws = wb.add_worksheet();
                let fmt = rust_xlsxwriter::Format::new();
                for i in 0..regions {
                    let row = i * 3;
                    ws.merge_range(row, 0, row, 3, &format!("Section {}", i + 1), &fmt)
                        .unwrap();
                    ws.write_number(row + 1, 0, (i * 100) as f64).unwrap();
                }
                wb.save(&out).unwrap();
            })
        },
    ));
    cleanup(&out);

    // edit-xlsx
    let out = output_dir().join("cmp-write-merge-editxlsx.xlsx");
    results.push(bench(&label, "edit-xlsx", "Write", Some(&out), || {
        let out = output_dir().join("cmp-write-merge-editxlsx.xlsx");
        Box::new(move || {
            use edit_xlsx::{Workbook as EditWorkbook, Write as _};
            let mut wb = EditWorkbook::new();
            let ws = wb.get_worksheet_mut(1).unwrap();
            for i in 0..regions {
                let row = i * 3 + 1;
                ws.merge_range(&format!("A{row}:D{row}"), format!("Section {}", i + 1))
                    .unwrap();
                ws.write((row + 1, 1), (i * 100) as f64).unwrap();
            }
            wb.save_as(&out).unwrap();
        })
    }));
    cleanup(&out);
}

// ===================================================================
// WRITE SCALING
// ===================================================================

fn bench_write_scale(results: &mut Vec<BenchResult>, rows: u32) {
    let cols: u32 = 10;
    let tag = if rows >= 1000 {
        format!("{}k", rows / 1000)
    } else {
        format!("{rows}")
    };
    let label = format!("Write {tag} rows x {cols} cols");
    println!("\n--- {label} ---");

    // SheetKit
    let out = output_dir().join(format!("cmp-scale-{tag}-sheetkit.xlsx"));
    let out_bench = out.clone();
    results.push(bench(
        &label,
        "SheetKit",
        "Write (Scale)",
        Some(&out_bench),
        move || {
            let tag2 = if rows >= 1000 {
                format!("{}k", rows / 1000)
            } else {
                format!("{rows}")
            };
            let out = output_dir().join(format!("cmp-scale-{tag2}-sheetkit.xlsx"));
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

    // rust_xlsxwriter
    let out = output_dir().join(format!("cmp-scale-{tag}-xlsxwriter.xlsx"));
    let out_bench = out.clone();
    results.push(bench(
        &label,
        "rust_xlsxwriter",
        "Write (Scale)",
        Some(&out_bench),
        move || {
            let tag2 = if rows >= 1000 {
                format!("{}k", rows / 1000)
            } else {
                format!("{rows}")
            };
            let out = output_dir().join(format!("cmp-scale-{tag2}-xlsxwriter.xlsx"));
            Box::new(move || {
                let mut wb = rust_xlsxwriter::Workbook::new();
                let ws = wb.add_worksheet();
                for r in 0..rows {
                    for c in 0..cols {
                        if c % 3 == 0 {
                            ws.write_number(r, c as u16, ((r + 1) as f64) * ((c + 1) as f64))
                                .unwrap();
                        } else if c % 3 == 1 {
                            ws.write_string(r, c as u16, format!("R{}C{c}", r + 1))
                                .unwrap();
                        } else {
                            ws.write_number(r, c as u16, ((r + 1) as f64) * (c as f64) / 100.0)
                                .unwrap();
                        }
                    }
                }
                wb.save(&out).unwrap();
            })
        },
    ));
    cleanup(&out);

    // edit-xlsx
    let out = output_dir().join(format!("cmp-scale-{tag}-editxlsx.xlsx"));
    let out_bench = out.clone();
    results.push(bench(
        &label,
        "edit-xlsx",
        "Write (Scale)",
        Some(&out_bench),
        move || {
            let tag2 = if rows >= 1000 {
                format!("{}k", rows / 1000)
            } else {
                format!("{rows}")
            };
            let out = output_dir().join(format!("cmp-scale-{tag2}-editxlsx.xlsx"));
            Box::new(move || {
                use edit_xlsx::{Workbook as EditWorkbook, Write as _};
                let mut wb = EditWorkbook::new();
                let ws = wb.get_worksheet_mut(1).unwrap();
                for r in 1..=rows {
                    for c in 1..=cols {
                        if (c - 1) % 3 == 0 {
                            ws.write((r, c), (r as f64) * (c as f64)).unwrap();
                        } else if (c - 1) % 3 == 1 {
                            ws.write((r, c), format!("R{r}C{}", c - 1)).unwrap();
                        } else {
                            ws.write((r, c), (r as f64) * ((c - 1) as f64) / 100.0)
                                .unwrap();
                        }
                    }
                }
                wb.save_as(&out).unwrap();
            })
        },
    ));
    cleanup(&out);
}

// ===================================================================
// BUFFER ROUND-TRIP (SheetKit only -- others lack buffer API or read)
// ===================================================================

fn bench_buffer_round_trip(results: &mut Vec<BenchResult>) {
    let rows: u32 = 10_000;
    let cols: u32 = 10;
    let label = format!("Buffer round-trip ({rows} rows)");
    println!("\n--- {label} ---");

    // SheetKit (full round-trip)
    results.push(bench(&label, "SheetKit", "Round-Trip", None, move || {
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

    // rust_xlsxwriter write -> calamine read
    results.push(bench(
        &label,
        "xlsxwriter+calamine",
        "Round-Trip",
        None,
        move || {
            Box::new(move || {
                // Write with rust_xlsxwriter
                let mut wb = rust_xlsxwriter::Workbook::new();
                let ws = wb.add_worksheet();
                for r in 0..rows {
                    for c in 0..cols {
                        ws.write_number(r, c as u16, ((r + 1) as f64) * ((c + 1) as f64))
                            .unwrap();
                    }
                }
                let buf = wb.save_to_buffer().unwrap();

                // Read with calamine
                use calamine::{Reader, Xlsx};
                use std::io::Cursor;
                let cursor = Cursor::new(buf);
                let mut reader: Xlsx<_> = Xlsx::new(cursor).unwrap();
                let names: Vec<String> = reader.sheet_names().to_vec();
                if let Ok(range) = reader.worksheet_range(&names[0]) {
                    for _row in range.rows() {}
                }
            })
        },
    ));
}

// ===================================================================
// STREAMING WRITE (SheetKit only -- others have no streaming API)
// ===================================================================

fn bench_streaming_write(results: &mut Vec<BenchResult>) {
    let rows: u32 = 50_000;
    let cols: u32 = 20;
    let label = format!("Streaming write ({rows} rows)");
    println!("\n--- {label} ---");

    let out = output_dir().join("cmp-stream-sheetkit.xlsx");
    results.push(bench(
        &label,
        "SheetKit",
        "Streaming",
        Some(&out),
        move || {
            let out = output_dir().join("cmp-stream-sheetkit.xlsx");
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
        },
    ));
    cleanup(&out);

    // rust_xlsxwriter normal write for comparison (no streaming API)
    let out = output_dir().join("cmp-stream-xlsxwriter.xlsx");
    results.push(bench(
        &label,
        "rust_xlsxwriter",
        "Streaming",
        Some(&out),
        move || {
            let out = output_dir().join("cmp-stream-xlsxwriter.xlsx");
            Box::new(move || {
                let mut wb = rust_xlsxwriter::Workbook::new();
                let ws = wb.add_worksheet();
                for r in 0..rows {
                    for c in 0..cols {
                        if c % 3 == 0 {
                            ws.write_number(r, c as u16, ((r + 1) as f64) * ((c + 1) as f64))
                                .unwrap();
                        } else if c % 3 == 1 {
                            ws.write_string(r, c as u16, format!("R{}C{c}", r + 1))
                                .unwrap();
                        } else {
                            ws.write_number(r, c as u16, ((r + 1) as f64) * (c as f64) / 100.0)
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

// ===================================================================
// RANDOM ACCESS READ
// ===================================================================

fn bench_random_access_read(results: &mut Vec<BenchResult>) {
    let filepath = fixtures_dir().join("large-data.xlsx");
    if !filepath.exists() {
        println!("  SKIP: large-data.xlsx not found. Run pnpm generate first.");
        return;
    }

    let lookups: usize = 1_000;
    let label = format!("Random-access read ({lookups} cells)");
    println!("\n--- {label} ---");

    // Deterministic random cell addresses
    let mut cells_rc: Vec<(u32, u32)> = Vec::with_capacity(lookups);
    let mut seed: u64 = 42;
    for _ in 0..lookups {
        seed = seed
            .wrapping_mul(6364136223846793005)
            .wrapping_add(1442695040888963407);
        let r = ((seed >> 33) % 50_000) as u32 + 2;
        seed = seed
            .wrapping_mul(6364136223846793005)
            .wrapping_add(1442695040888963407);
        let c = ((seed >> 33) % 20) as u32 + 1;
        cells_rc.push((r, c));
    }
    let cells_str: Vec<String> = cells_rc.iter().map(|(r, c)| cell_ref(*r, *c)).collect();

    // SheetKit
    let fp = filepath.clone();
    let cells = cells_str.clone();
    results.push(bench(
        &label,
        "SheetKit",
        "Random Access",
        None,
        move || {
            let fp = fp.clone();
            let cells = cells.clone();
            Box::new(move || {
                let wb = Workbook::open(&fp).unwrap();
                for cell in &cells {
                    let _ = wb.get_cell_value("Sheet1", cell);
                }
            })
        },
    ));

    // calamine
    let fp = filepath.clone();
    let cells_rc2 = cells_rc.clone();
    results.push(bench(
        &label,
        "calamine",
        "Random Access",
        None,
        move || {
            let fp = fp.clone();
            let cells_rc2 = cells_rc2.clone();
            Box::new(move || {
                use calamine::{open_workbook, Reader, Xlsx};
                let mut wb: Xlsx<_> = open_workbook(&fp).unwrap();
                if let Ok(range) = wb.worksheet_range("Sheet1") {
                    for (r, c) in &cells_rc2 {
                        // calamine uses 0-indexed
                        let _ = range.get_value((*r - 1, *c - 1));
                    }
                }
            })
        },
    ));
}

// ===================================================================
// MODIFY BENCHMARK (SheetKit vs edit-xlsx)
// ===================================================================

fn bench_modify_file(results: &mut Vec<BenchResult>) {
    let filepath = fixtures_dir().join("large-data.xlsx");
    if !filepath.exists() {
        println!("  SKIP: large-data.xlsx not found. Run pnpm generate first.");
        return;
    }

    let label = "Modify 1000 cells in 50k-row file";
    println!("\n--- {label} ---");

    // SheetKit
    let fp = filepath.clone();
    let out = output_dir().join("cmp-modify-sheetkit.xlsx");
    results.push(bench(label, "SheetKit", "Modify", Some(&out), move || {
        let fp = fp.clone();
        let out = output_dir().join("cmp-modify-sheetkit.xlsx");
        Box::new(move || {
            let mut wb = Workbook::open(&fp).unwrap();
            for i in 0..1000u32 {
                let r = i + 2;
                wb.set_cell_value("Sheet1", &format!("A{r}"), format!("Modified_{i}"))
                    .unwrap();
            }
            wb.save(&out).unwrap();
        })
    }));
    cleanup(&out);

    // edit-xlsx
    let fp = filepath.clone();
    let out = output_dir().join("cmp-modify-editxlsx.xlsx");
    results.push(bench(label, "edit-xlsx", "Modify", Some(&out), move || {
        let fp = fp.clone();
        let out = output_dir().join("cmp-modify-editxlsx.xlsx");
        Box::new(move || {
            use edit_xlsx::{Workbook as EditWorkbook, Write as _};
            let mut wb = EditWorkbook::from_path(&fp).unwrap();
            let ws = wb.get_worksheet_mut(1).unwrap();
            for i in 0..1000u32 {
                let r = i + 2;
                ws.write((r, 1), format!("Modified_{i}")).unwrap();
            }
            wb.save_as(&out).unwrap();
        })
    }));
    cleanup(&out);

    // calamine + rust_xlsxwriter cannot modify
    println!("  [calamine       ] N/A (read-only)");
    println!("  [rust_xlsxwriter] N/A (write-only, cannot open existing files)");
}

// ===================================================================
// Summary and report
// ===================================================================

fn print_summary_table(results: &[BenchResult]) {
    println!("\n\n========================================");
    println!(" RUST ECOSYSTEM COMPARISON RESULTS");
    println!(" ({BENCH_RUNS} runs per scenario, {WARMUP_RUNS} warmup)");
    println!("========================================\n");

    let libraries: Vec<String> = {
        let mut seen = Vec::new();
        for r in results {
            if !seen.contains(&r.library) {
                seen.push(r.library.clone());
            }
        }
        seen
    };
    let scenarios: Vec<String> = {
        let mut seen = Vec::new();
        for r in results {
            if !seen.contains(&r.scenario) {
                seen.push(r.scenario.clone());
            }
        }
        seen
    };

    // Header
    print!("| {:<46}", "Scenario");
    for lib in &libraries {
        print!("| {:>16} ", lib);
    }
    println!("| Winner          |");

    print!("|{:-<47}", "");
    for _ in &libraries {
        print!("|{:-<17}", "");
    }
    println!("|{:-<17}|", "");

    // Rows
    for scenario in &scenarios {
        print!("| {:<46}", scenario);
        let mut best_lib = String::new();
        let mut best_ms = f64::INFINITY;
        for lib in &libraries {
            if let Some(r) = results
                .iter()
                .find(|r| &r.scenario == scenario && &r.library == lib)
            {
                let med = r.median();
                print!("| {:>16} ", format_ms(med));
                if med < best_ms {
                    best_ms = med;
                    best_lib = lib.clone();
                }
            } else {
                print!("| {:>16} ", "N/A");
            }
        }
        println!("| {:<16}|", best_lib);
    }
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

fn generate_markdown_report(results: &[BenchResult]) -> String {
    let mut lines = Vec::new();
    lines.push("# Rust Excel Library Comparison Benchmark".to_string());
    lines.push(String::new());
    lines.push(format!("Benchmark run: {}", chrono_free_timestamp()));
    lines.push(String::new());

    lines.push("## Libraries".to_string());
    lines.push(String::new());
    lines.push("| Library | Description | Capability |".to_string());
    lines.push("|---------|-------------|------------|".to_string());
    lines.push(
        "| **SheetKit** | Rust Excel library (this project) | Read + Write + Modify |".to_string(),
    );
    lines.push("| **calamine** | Fast Excel/ODS reader | Read-only |".to_string());
    lines.push(
        "| **rust_xlsxwriter** | Excel writer (port of libxlsxwriter) | Write-only |".to_string(),
    );
    lines.push(
        "| **edit-xlsx** | Excel read/modify/write library | Read + Write + Modify |".to_string(),
    );
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

    // Results by category
    let categories: Vec<String> = {
        let mut seen = Vec::new();
        for r in results {
            if !seen.contains(&r.category) {
                seen.push(r.category.clone());
            }
        }
        seen
    };
    let libraries: Vec<String> = {
        let mut seen = Vec::new();
        for r in results {
            if !seen.contains(&r.library) {
                seen.push(r.library.clone());
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

        // Table header
        let mut header = "| Scenario ".to_string();
        for lib in &libraries {
            header.push_str(&format!("| {lib} "));
        }
        header.push_str("| Winner |");
        lines.push(header);

        let mut sep = "|----------".to_string();
        for _ in &libraries {
            sep.push_str("|--------");
        }
        sep.push_str("|--------|");
        lines.push(sep);

        let scenarios: Vec<String> = {
            let mut seen = Vec::new();
            for r in &subset {
                if !seen.contains(&r.scenario) {
                    seen.push(r.scenario.clone());
                }
            }
            seen
        };

        for scenario in &scenarios {
            let mut row = format!("| {scenario} ");
            let mut best_lib = String::new();
            let mut best_ms = f64::INFINITY;
            for lib in &libraries {
                if let Some(r) = subset
                    .iter()
                    .find(|r| &r.scenario == scenario && &r.library == lib)
                {
                    let med = r.median();
                    row.push_str(&format!("| {} ", format_ms(med)));
                    if med < best_ms {
                        best_ms = med;
                        best_lib = lib.clone();
                    }
                } else {
                    row.push_str("| N/A ");
                }
            }
            row.push_str(&format!("| {best_lib} |"));
            lines.push(row);
        }
        lines.push(String::new());
    }

    // Detailed stats
    lines.push("## Detailed Statistics".to_string());
    lines.push(String::new());
    lines.push("| Scenario | Library | Median | Min | Max | P95 | Peak Mem (MB) |".to_string());
    lines.push("|----------|---------|--------|-----|-----|-----|---------------|".to_string());
    for r in results {
        lines.push(format!(
            "| {} | {} | {} | {} | {} | {} | {:.1} |",
            r.scenario,
            r.library,
            format_ms(r.median()),
            format_ms(r.min()),
            format_ms(r.max()),
            format_ms(r.p95()),
            r.peak_mem_mb(),
        ));
    }
    lines.push(String::new());

    // Win summary
    let all_scenarios: Vec<String> = {
        let mut seen = Vec::new();
        for r in results {
            if !seen.contains(&r.scenario) {
                seen.push(r.scenario.clone());
            }
        }
        seen
    };
    let mut wins: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
    for scenario in &all_scenarios {
        let mut best_lib = String::new();
        let mut best_ms = f64::INFINITY;
        for r in results.iter().filter(|r| &r.scenario == scenario) {
            let med = r.median();
            if med < best_ms {
                best_ms = med;
                best_lib = r.library.clone();
            }
        }
        if !best_lib.is_empty() {
            *wins.entry(best_lib).or_default() += 1;
        }
    }
    lines.push("## Win Summary".to_string());
    lines.push(String::new());
    lines.push("| Library | Wins |".to_string());
    lines.push("|---------|------|".to_string());
    let mut sorted_wins: Vec<_> = wins.into_iter().collect();
    sorted_wins.sort_by(|a, b| b.1.cmp(&a.1));
    for (lib, count) in &sorted_wins {
        lines.push(format!("| {lib} | {count}/{} |", all_scenarios.len()));
    }
    lines.push(String::new());

    lines.join("\n")
}

// ===================================================================
// Main
// ===================================================================

fn main() {
    println!("Rust Excel Library Comparison Benchmark");
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
    println!("Libraries: SheetKit, calamine, rust_xlsxwriter, edit-xlsx");
    println!("Iterations: {BENCH_RUNS} runs per scenario, {WARMUP_RUNS} warmup\n");

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

    // Modify
    println!("\n\n=== MODIFY ===");
    bench_modify_file(&mut results);

    // Summary
    print_summary_table(&results);

    // Write markdown report
    let report = generate_markdown_report(&results);
    let report_path = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("RESULTS.md");
    let mut file = fs::File::create(&report_path).unwrap();
    file.write_all(report.as_bytes()).unwrap();
    println!("\nMarkdown report written to: {report_path:?}");
}
