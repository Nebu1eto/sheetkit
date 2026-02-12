/**
 * Focused baseline benchmark for SheetKit open/getRows performance.
 *
 * Measures open latency (path and buffer), openSync vs open (async),
 * getRows and getRowsRaw across scaling fixtures. Reports timing
 * statistics (min/max/median/p95) and peak RSS delta per operation.
 *
 * Usage:
 *   node --expose-gc --import tsx benchmarks/node/bench-baseline.ts
 */

import { existsSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { cpus, totalmem } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook } from "@sheetkit/node";

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURES_DIR = join(__dirname, "fixtures");
const WARMUP_RUNS = 1;
const BENCH_RUNS = 5;

interface FixtureSpec {
  name: string;
  label: string;
}

interface MeasureResult {
  label: string;
  min: number;
  max: number;
  median: number;
  p95: number;
  rssMedian: number;
  timesMs: number[];
  rssDeltas: number[];
}

interface BenchResult extends MeasureResult {
  category: string;
  fixture: string;
}

const FIXTURES: FixtureSpec[] = [
  { name: "scale-1k.xlsx", label: "scale-1k" },
  { name: "scale-10k.xlsx", label: "scale-10k" },
  { name: "scale-100k.xlsx", label: "scale-100k" },
  { name: "large-data.xlsx", label: "large-data" },
];

function fixturePath(name: string): string {
  return join(FIXTURES_DIR, name);
}

function fixtureExists(name: string): boolean {
  return existsSync(fixturePath(name));
}

function fileSizeKb(name: string): number | null {
  try {
    return statSync(fixturePath(name)).size / 1024;
  } catch {
    return null;
  }
}

function rssMb(): number {
  if (globalThis.gc) globalThis.gc();
  return process.memoryUsage().rss / 1024 / 1024;
}

function median(arr: number[]): number {
  if (arr.length === 0) return 0;
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? (sorted[mid - 1] + sorted[mid]) / 2
    : sorted[mid];
}

function p95(arr: number[]): number {
  if (arr.length === 0) return 0;
  const sorted = [...arr].sort((a, b) => a - b);
  const idx = Math.ceil(0.95 * sorted.length) - 1;
  return sorted[Math.max(0, idx)];
}

function formatMs(ms: number): string {
  if (ms < 1000) return `${ms.toFixed(1)}ms`;
  return `${(ms / 1000).toFixed(3)}s`;
}

async function measure(
  label: string,
  fn: () => void | Promise<void>,
): Promise<MeasureResult> {
  for (let i = 0; i < WARMUP_RUNS; i++) {
    if (globalThis.gc) globalThis.gc();
    await fn();
  }

  const timesMs: number[] = [];
  const rssDeltas: number[] = [];

  for (let i = 0; i < BENCH_RUNS; i++) {
    if (globalThis.gc) globalThis.gc();
    const rssBefore = rssMb();
    const start = performance.now();
    await fn();
    const elapsed = performance.now() - start;
    const rssAfter = rssMb();

    timesMs.push(elapsed);
    rssDeltas.push(Math.max(0, rssAfter - rssBefore));
  }

  const sorted = [...timesMs].sort((a, b) => a - b);
  const stats: MeasureResult = {
    label,
    min: sorted[0],
    max: sorted[sorted.length - 1],
    median: median(timesMs),
    p95: p95(timesMs),
    rssMedian: median(rssDeltas),
    timesMs,
    rssDeltas,
  };

  console.log(
    `  ${label.padEnd(52)} ` +
      `med=${formatMs(stats.median).padStart(10)} ` +
      `min=${formatMs(stats.min).padStart(10)} ` +
      `max=${formatMs(stats.max).padStart(10)} ` +
      `p95=${formatMs(stats.p95).padStart(10)} ` +
      `rss=${stats.rssMedian.toFixed(1).padStart(6)}MB`,
  );

  return stats;
}

function round3(n: number): number {
  return Math.round(n * 1000) / 1000;
}

function round1(n: number): number {
  return Math.round(n * 10) / 10;
}

function printSummaryTable(allResults: BenchResult[]): void {
  console.log("=== SUMMARY ===");
  console.log("");

  const categories = [...new Set(allResults.map((r) => r.category))];

  console.log(
    `| ${"Category".padEnd(22)}` +
      `| ${"Fixture".padEnd(14)}` +
      `| ${"Median".padEnd(12)}` +
      `| ${"Min".padEnd(12)}` +
      `| ${"Max".padEnd(12)}` +
      `| ${"P95".padEnd(12)}` +
      `| ${"RSS Delta".padEnd(10)}|`,
  );
  console.log(
    `|${"-".repeat(23)}` +
      `|${"-".repeat(15)}` +
      `|${"-".repeat(13)}` +
      `|${"-".repeat(13)}` +
      `|${"-".repeat(13)}` +
      `|${"-".repeat(13)}` +
      `|${"-".repeat(11)}|`,
  );

  for (const cat of categories) {
    const catResults = allResults.filter((r) => r.category === cat);
    for (const r of catResults) {
      console.log(
        `| ${r.category.padEnd(22)}` +
          `| ${r.fixture.padEnd(14)}` +
          `| ${formatMs(r.median).padEnd(12)}` +
          `| ${formatMs(r.min).padEnd(12)}` +
          `| ${formatMs(r.max).padEnd(12)}` +
          `| ${formatMs(r.p95).padEnd(12)}` +
          `| ${`${r.rssMedian.toFixed(1)}MB`.padEnd(10)}|`,
      );
    }
  }
  console.log("");
}

async function main(): Promise<void> {
  console.log("SheetKit Baseline Benchmark (open / getRows)");
  console.log(
    `Platform: ${process.platform} ${process.arch} | Node.js: ${process.version}`,
  );
  console.log(`CPU: ${cpus()[0]?.model ?? "Unknown"}`);
  console.log(`RAM: ${Math.round(totalmem() / 1024 ** 3)} GB`);
  console.log(
    `Config: ${WARMUP_RUNS} warmup + ${BENCH_RUNS} measured runs per scenario`,
  );
  if (!globalThis.gc) {
    console.log(
      "WARNING: --expose-gc not enabled. RSS measurements may be less accurate.",
    );
  }
  console.log("");

  const available = FIXTURES.filter((f) => {
    if (!fixtureExists(f.name)) {
      console.log(
        `SKIP: ${f.name} not found. Run 'pnpm generate' in benchmarks/node first.`,
      );
      return false;
    }
    return true;
  });

  if (available.length === 0) {
    console.log("No fixtures available. Exiting.");
    process.exit(1);
  }

  console.log("Fixtures:");
  for (const f of available) {
    const sizeKb = fileSizeKb(f.name);
    const sizeStr =
      sizeKb != null ? `${(sizeKb / 1024).toFixed(1)}MB` : "unknown size";
    console.log(`  ${f.label.padEnd(20)} ${sizeStr}`);
  }
  console.log("");

  const allResults: BenchResult[] = [];

  // 1. Open latency (path, sync)
  console.log("--- openSync (path) ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const result = await measure(`openSync(${f.label})`, () => {
      Workbook.openSync(path);
    });
    allResults.push({ ...result, category: "openSync-path", fixture: f.label });
  }
  console.log("");

  // 2. Open latency (path, async)
  console.log("--- open (path, async) ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const result = await measure(`open(${f.label})`, async () => {
      await Workbook.open(path);
    });
    allResults.push({
      ...result,
      category: "open-path-async",
      fixture: f.label,
    });
  }
  console.log("");

  // 3. Open latency (buffer, sync)
  console.log("--- openBufferSync ---");
  for (const f of available) {
    const buf = readFileSync(fixturePath(f.name));
    const result = await measure(`openBufferSync(${f.label})`, () => {
      Workbook.openBufferSync(buf);
    });
    allResults.push({
      ...result,
      category: "openBufferSync",
      fixture: f.label,
    });
  }
  console.log("");

  // 4. Open latency (buffer, async)
  console.log("--- openBuffer (async) ---");
  for (const f of available) {
    const buf = readFileSync(fixturePath(f.name));
    const result = await measure(`openBuffer(${f.label})`, async () => {
      await Workbook.openBuffer(buf);
    });
    allResults.push({
      ...result,
      category: "openBuffer-async",
      fixture: f.label,
    });
  }
  console.log("");

  // 5. openSync vs open comparison
  console.log("--- openSync vs open (comparison) ---");
  for (const f of available) {
    const syncResult = allResults.find(
      (r) => r.category === "openSync-path" && r.fixture === f.label,
    );
    const asyncResult = allResults.find(
      (r) => r.category === "open-path-async" && r.fixture === f.label,
    );
    if (syncResult && asyncResult) {
      const ratio = asyncResult.median / syncResult.median;
      console.log(
        `  ${f.label.padEnd(20)} ` +
          `sync=${formatMs(syncResult.median).padStart(10)} ` +
          `async=${formatMs(asyncResult.median).padStart(10)} ` +
          `ratio=${ratio.toFixed(2)}x`,
      );
    }
  }
  console.log("");

  // 6. getRows latency
  console.log("--- getRows (Sheet1) ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const result = await measure(`getRows(${f.label})`, () => {
      const wb = Workbook.openSync(path);
      wb.getRows("Sheet1");
    });
    allResults.push({ ...result, category: "getRows", fixture: f.label });
  }
  console.log("");

  // 7. getRows on pre-opened workbook
  console.log("--- getRows only (pre-opened) ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const wb = Workbook.openSync(path);
    const result = await measure(`getRows-only(${f.label})`, () => {
      wb.getRows("Sheet1");
    });
    allResults.push({
      ...result,
      category: "getRows-only",
      fixture: f.label,
    });
  }
  console.log("");

  // 8. getRowsRaw latency
  console.log("--- getRowsRaw (Sheet1) ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const result = await measure(`getRowsRaw(${f.label})`, () => {
      const wb = Workbook.openSync(path);
      wb.getRowsRaw("Sheet1");
    });
    allResults.push({ ...result, category: "getRowsRaw", fixture: f.label });
  }
  console.log("");

  // 9. getRowsRaw on pre-opened workbook
  console.log("--- getRowsRaw only (pre-opened) ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const wb = Workbook.openSync(path);
    const result = await measure(`getRowsRaw-only(${f.label})`, () => {
      wb.getRowsRaw("Sheet1");
    });
    allResults.push({
      ...result,
      category: "getRowsRaw-only",
      fixture: f.label,
    });
  }
  console.log("");

  // 10. Lazy open latency (path, async)
  console.log("--- open lazy (path, async) ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const result = await measure(`open-lazy(${f.label})`, async () => {
      await Workbook.open(path, { readMode: "lazy" });
    });
    allResults.push({
      ...result,
      category: "open-lazy-path",
      fixture: f.label,
    });
  }
  console.log("");

  // 11. Lazy open + getRows (measures on-demand hydration)
  console.log("--- lazy open + getRows ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const result = await measure(`lazy+getRows(${f.label})`, async () => {
      const wb = await Workbook.open(path, { readMode: "lazy" });
      wb.getRows("Sheet1");
    });
    allResults.push({
      ...result,
      category: "lazy-open-getRows",
      fixture: f.label,
    });
  }
  console.log("");

  // 12. Lazy open vs eager open comparison
  console.log("--- lazy vs eager open comparison ---");
  for (const f of available) {
    const eagerResult = allResults.find(
      (r) => r.category === "open-path-async" && r.fixture === f.label,
    );
    const lazyResult = allResults.find(
      (r) => r.category === "open-lazy-path" && r.fixture === f.label,
    );
    if (eagerResult && lazyResult) {
      const ratio = lazyResult.median / eagerResult.median;
      console.log(
        `  ${f.label.padEnd(20)} ` +
          `eager=${formatMs(eagerResult.median).padStart(10)} ` +
          `lazy=${formatMs(lazyResult.median).padStart(10)} ` +
          `ratio=${ratio.toFixed(2)}x`,
      );
    }
  }
  console.log("");

  // 13. openSheetReader streaming
  console.log("--- openSheetReader (stream mode) ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const result = await measure(
      `openSheetReader(${f.label})`,
      async () => {
        const wb = await Workbook.open(path, { readMode: "stream" });
        const reader = await wb.openSheetReader("Sheet1", {
          batchSize: 1000,
        });
        for await (const _batch of reader) {
          // consume all batches
        }
      },
    );
    allResults.push({
      ...result,
      category: "openSheetReader",
      fixture: f.label,
    });
  }
  console.log("");

  // 14. getRowsBufferV2 vs getRowsBuffer (v1)
  console.log("--- getRowsBufferV2 vs getRowsBuffer ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const wb = Workbook.openSync(path);

    const v1Result = await measure(`getRowsBuffer-v1(${f.label})`, () => {
      wb.getRowsBuffer("Sheet1");
    });
    allResults.push({
      ...v1Result,
      category: "getRowsBuffer-v1",
      fixture: f.label,
    });

    const v2Result = await measure(`getRowsBufferV2(${f.label})`, () => {
      wb.getRowsBufferV2("Sheet1");
    });
    allResults.push({
      ...v2Result,
      category: "getRowsBufferV2",
      fixture: f.label,
    });
  }
  console.log("");

  // 15. Copy-on-write save (lazy open + save without modifications)
  console.log("--- copy-on-write save (lazy, no modifications) ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const tmpOut = join(__dirname, "output", `cow-save-${f.label}.xlsx`);
    const result = await measure(`cow-save(${f.label})`, async () => {
      const wb = await Workbook.open(path, { readMode: "lazy" });
      wb.saveSync(tmpOut);
    });
    allResults.push({
      ...result,
      category: "cow-save-untouched",
      fixture: f.label,
    });
    try {
      const { unlinkSync } = await import("node:fs");
      unlinkSync(tmpOut);
    } catch {
      /* ignore cleanup errors */
    }
  }
  console.log("");

  // 16. Copy-on-write save (lazy open + single cell edit + save)
  console.log("--- copy-on-write save (lazy, single-cell edit) ---");
  for (const f of available) {
    const path = fixturePath(f.name);
    const tmpOut = join(__dirname, "output", `cow-edit-${f.label}.xlsx`);
    const result = await measure(`cow-edit-save(${f.label})`, async () => {
      const wb = await Workbook.open(path, { readMode: "lazy" });
      wb.setCellValue("Sheet1", "A1", "edited");
      wb.saveSync(tmpOut);
    });
    allResults.push({
      ...result,
      category: "cow-save-edited",
      fixture: f.label,
    });
    try {
      const { unlinkSync } = await import("node:fs");
      unlinkSync(tmpOut);
    } catch {
      /* ignore cleanup errors */
    }
  }
  console.log("");

  // Summary table
  printSummaryTable(allResults);

  // Write JSON baseline
  const outputPath = join(__dirname, "baseline-results.json");
  const jsonOutput = {
    timestamp: new Date().toISOString(),
    platform: `${process.platform} ${process.arch}`,
    nodeVersion: process.version,
    cpu: cpus()[0]?.model ?? "Unknown",
    ramGb: Math.round(totalmem() / 1024 ** 3),
    config: { warmupRuns: WARMUP_RUNS, benchRuns: BENCH_RUNS },
    results: allResults.map((r) => ({
      category: r.category,
      fixture: r.fixture,
      label: r.label,
      min: round3(r.min),
      max: round3(r.max),
      median: round3(r.median),
      p95: round3(r.p95),
      rssMedianMb: round1(r.rssMedian),
      timesMs: r.timesMs.map(round3),
      rssDeltas: r.rssDeltas.map(round1),
    })),
  };
  writeFileSync(outputPath, `${JSON.stringify(jsonOutput, null, 2)}\n`);
  console.log(`\nBaseline JSON written to: ${outputPath}`);
}

main().catch((err) => {
  console.error("Benchmark failed:", err);
  process.exit(1);
});
