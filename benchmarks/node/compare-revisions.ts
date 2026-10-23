/**
 * Compare release-gate performance between two git revisions.
 *
 * Creates a detached worktree per revision under a scratch directory, copies
 * the current Cargo.lock into each (Cargo.lock is gitignored), builds the Node
 * binding in release profile, runs release-gate.ts from THIS checkout against
 * each revision for every operation and fixture, and writes a markdown table of
 * median latency with the relative change. The script is CPU heavy and runs
 * everything sequentially; run it on an otherwise idle machine.
 *
 * Usage:
 *   SHEETKIT_COMPARE_BASE=<rev> [SHEETKIT_COMPARE_HEAD=HEAD] \
 *   [SHEETKIT_COMPARE_SAMPLES=20] [SHEETKIT_COMPARE_WARMUPS=3] \
 *   [SHEETKIT_COMPARE_THRESHOLD_PERCENT=10] \
 *   [SHEETKIT_COMPARE_SCRATCH=/tmp/sheetkit-compare] \
 *   [SHEETKIT_COMPARE_OPERATIONS=ordinary-open,dense-traversal,...] \
 *   [SHEETKIT_COMPARE_FIXTURES=large-data.xlsx,...] \
 *   [SHEETKIT_COMPARE_PASSWORD=<password>] \
 *   [SHEETKIT_COMPARE_KEEP=1] \
 *   node --import tsx benchmarks/node/compare-revisions.ts
 *
 * Output: benchmarks/node/output/compare-<base>-<head>.md and the raw
 * release-gate JSON files under benchmarks/node/output/compare-<base>-<head>/.
 * Exit code is 1 when any operation regresses beyond the threshold.
 */

import assert from "node:assert";
import { execFileSync, spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const OPERATIONS = [
  "ordinary-open",
  "encrypted-open",
  "scalar-calculation",
  "dense-traversal",
  "streaming",
  "output-csv",
  "output-json",
  "output-html",
  "output-svg",
] as const;

type Operation = (typeof OPERATIONS)[number];

interface GateResult {
  operation: Operation;
  fixture: string;
  medianMs: number;
  madMs: number;
  samples: number;
}

interface Comparison {
  operation: Operation;
  fixture: string;
  base: GateResult | undefined;
  head: GateResult | undefined;
  deltaPercent: number | undefined;
}

const scriptDirectory = dirname(fileURLToPath(import.meta.url));
const repositoryRoot = resolve(scriptDirectory, "../..");
const fixturesDirectory = join(scriptDirectory, "fixtures");

function env(name: string): string | undefined {
  const value = process.env[name];
  return value === undefined || value.length === 0 ? undefined : value;
}

function integerEnv(name: string, fallback: number): number {
  const raw = env(name);
  if (raw === undefined) return fallback;
  const value = Number(raw);
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new Error(`${name} must be a non-negative safe integer`);
  }
  return value;
}

function listEnv(name: string, fallback: readonly string[]): string[] {
  const raw = env(name);
  return raw === undefined ? [...fallback] : raw.split(",").map((item) => item.trim()).filter(Boolean);
}

function git(args: readonly string[], cwd: string = repositoryRoot): string {
  return execFileSync("git", args, { cwd, encoding: "utf8" }).trim();
}

function run(command: string, args: readonly string[], cwd: string, extraEnv: NodeJS.ProcessEnv = {}): void {
  const result = spawnSync(command, args, {
    cwd,
    stdio: "inherit",
    env: { ...process.env, ...extraEnv },
  });
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(" ")} failed in ${cwd} with status ${result.status}`);
  }
}

function prepareWorktree(
  revision: string,
  sha: string,
  scratch: string,
  owned: string[],
): string {
  const path = join(scratch, `worktree-${owned.length}-${sha.slice(0, 12)}`);
  if (existsSync(path)) {
    git(["worktree", "remove", "--force", path]);
  }
  git(["worktree", "add", "--detach", path, sha]);
  owned.push(path);
  const lock = join(repositoryRoot, "Cargo.lock");
  if (existsSync(lock)) {
    copyFileSync(lock, join(path, "Cargo.lock"));
  }
  console.log(`[compare] ${revision} -> ${sha} at ${path}`);
  run("pnpm", ["install", "--frozen-lockfile", "--prefer-offline"], path);
  run("pnpm", ["--filter", "@sheetkit/node", "build"], path, {
    CARGO_TARGET_DIR: join(scratch, `target-${sha.slice(0, 12)}`),
  });
  return path;
}

function gateScriptFor(worktree: string): string {
  const target = join(worktree, "benchmarks", "node", "release-gate.ts");
  copyFileSync(join(scriptDirectory, "release-gate.ts"), target);
  return target;
}

function runGate(
  worktree: string,
  operation: Operation,
  fixture: string,
  output: string,
  samples: number,
  warmups: number,
  encryptedFixture: string | undefined,
  password: string | undefined,
): GateResult | undefined {
  if (existsSync(output)) rmSync(output);
  const gateEnv: NodeJS.ProcessEnv = {
    SHEETKIT_GATE_OPERATION: operation,
    SHEETKIT_GATE_FIXTURE: fixture,
    SHEETKIT_GATE_OUTPUT: output,
    SHEETKIT_GATE_SAMPLES: String(samples),
    SHEETKIT_GATE_WARMUPS: String(warmups),
  };
  if (encryptedFixture !== undefined) gateEnv.SHEETKIT_GATE_ENCRYPTED_FIXTURE = encryptedFixture;
  if (password !== undefined) gateEnv.SHEETKIT_GATE_PASSWORD = password;
  const result = spawnSync(
    process.execPath,
    ["--expose-gc", "--import", "tsx", gateScriptFor(worktree)],
    { cwd: join(worktree, "benchmarks", "node"), stdio: "inherit", env: { ...process.env, ...gateEnv } },
  );
  if (result.status !== 0 || !existsSync(output)) {
    console.error(`[compare] ${operation} on ${fixture} failed for ${worktree}`);
    return undefined;
  }
  const parsed: unknown = JSON.parse(readFileSync(output, "utf8"));
  assert(typeof parsed === "object" && parsed !== null);
  return parsed as GateResult;
}

function formatMs(value: number | undefined): string {
  return value === undefined ? "n/a" : value.toFixed(3);
}

function formatDelta(value: number | undefined): string {
  if (value === undefined) return "n/a";
  const sign = value > 0 ? "+" : "";
  return `${sign}${value.toFixed(1)}%`;
}

function main(): void {
  const baseRevision = env("SHEETKIT_COMPARE_BASE");
  if (baseRevision === undefined) {
    throw new Error("SHEETKIT_COMPARE_BASE is required");
  }
  const headRevision = env("SHEETKIT_COMPARE_HEAD") ?? "HEAD";
  const samples = integerEnv("SHEETKIT_COMPARE_SAMPLES", 20);
  const warmups = integerEnv("SHEETKIT_COMPARE_WARMUPS", 3);
  const threshold = integerEnv("SHEETKIT_COMPARE_THRESHOLD_PERCENT", 10);
  const scratch = env("SHEETKIT_COMPARE_SCRATCH") ?? "/tmp/sheetkit-compare";
  const password = env("SHEETKIT_COMPARE_PASSWORD");
  const keep = env("SHEETKIT_COMPARE_KEEP") === "1";
  const operations = listEnv("SHEETKIT_COMPARE_OPERATIONS", OPERATIONS).map((name) => {
    if (!OPERATIONS.includes(name as Operation)) {
      throw new Error(`Unsupported operation: ${name}`);
    }
    return name as Operation;
  });
  const fixtures = listEnv(
    "SHEETKIT_COMPARE_FIXTURES",
    readdirSync(fixturesDirectory).filter((name) => name.endsWith(".xlsx")),
  ).map((name) => resolve(fixturesDirectory, name));

  const baseSha = git(["rev-parse", baseRevision]);
  const headSha = git(["rev-parse", headRevision]);
  const label = `compare-${baseSha.slice(0, 7)}-${headSha.slice(0, 7)}`;
  const outputDirectory = join(scriptDirectory, "output", label);
  mkdirSync(outputDirectory, { recursive: true });
  // Everything this run creates lives in one run-owned child directory so
  // cleanup never touches unrelated contents of the scratch root.
  const runScratch = join(scratch, `${label}-${process.pid}`);
  mkdirSync(runScratch, { recursive: true });
  const ownedWorktrees: string[] = [];
  const cleanup = (): void => {
    if (keep) {
      console.log(`[compare] keeping ${runScratch}`);
      return;
    }
    for (const path of ownedWorktrees) {
      try {
        git(["worktree", "remove", "--force", path]);
      } catch (error) {
        console.error(`[compare] failed to remove worktree ${path}: ${String(error)}`);
      }
    }
    rmSync(runScratch, { recursive: true, force: true });
  };

  try {
    runComparison({
      baseRevision,
      headRevision,
      baseSha,
      headSha,
      label,
      outputDirectory,
      runScratch,
      ownedWorktrees,
      samples,
      warmups,
      threshold,
      password,
      operations,
      fixtures,
    });
  } finally {
    cleanup();
  }
}

interface ComparisonRun {
  baseRevision: string;
  headRevision: string;
  baseSha: string;
  headSha: string;
  label: string;
  outputDirectory: string;
  runScratch: string;
  ownedWorktrees: string[];
  samples: number;
  warmups: number;
  threshold: number;
  password: string | undefined;
  operations: Operation[];
  fixtures: string[];
}

function runComparison(run: ComparisonRun): void {
  const {
    baseRevision,
    headRevision,
    baseSha,
    headSha,
    label,
    outputDirectory,
    runScratch,
    ownedWorktrees,
    samples,
    warmups,
    threshold,
    password,
    operations,
    fixtures,
  } = run;
  const worktrees = {
    base: prepareWorktree(baseRevision, baseSha, runScratch, ownedWorktrees),
    head: prepareWorktree(headRevision, headSha, runScratch, ownedWorktrees),
  };

  const encryptedFixture = password === undefined ? undefined : join(outputDirectory, "encrypted-fixture.xlsx");
  if (encryptedFixture !== undefined && !existsSync(encryptedFixture) && fixtures.length > 0) {
    run(process.execPath, ["--import", "tsx", gateScriptFor(worktrees.head), "--generate-encrypted"], join(worktrees.head, "benchmarks", "node"), {
      SHEETKIT_GATE_FIXTURE: fixtures[0],
      SHEETKIT_GATE_ENCRYPTED_FIXTURE: encryptedFixture,
      SHEETKIT_GATE_PASSWORD: password,
    });
  }

  const comparisons: Comparison[] = [];
  for (const fixture of fixtures) {
    const fixtureName = fixture.slice(fixture.lastIndexOf("/") + 1);
    for (const operation of operations) {
      if (operation === "encrypted-open" && encryptedFixture === undefined) continue;
      const results: Partial<Record<"base" | "head", GateResult | undefined>> = {};
      for (const side of ["base", "head"] as const) {
        results[side] = runGate(
          worktrees[side],
          operation,
          fixture,
          join(outputDirectory, `${side}-${operation}-${fixtureName}.json`),
          samples,
          warmups,
          encryptedFixture,
          password,
        );
      }
      const base = results.base;
      const head = results.head;
      const deltaPercent =
        base !== undefined && head !== undefined && base.medianMs > 0
          ? ((head.medianMs - base.medianMs) / base.medianMs) * 100
          : undefined;
      comparisons.push({ operation, fixture: fixtureName, base, head, deltaPercent });
    }
  }

  const lines = [
    `# Release gate comparison: ${baseSha.slice(0, 7)} -> ${headSha.slice(0, 7)}`,
    "",
    `Samples: ${samples}, warmups: ${warmups}, regression threshold: +${threshold}% median.`,
    "",
    "| Operation | Fixture | Base median (ms) | Base MAD | Head median (ms) | Head MAD | Delta | Verdict |",
    "|---|---|---|---|---|---|---|---|",
  ];
  let regressions = 0;
  let missingResults = 0;
  for (const item of comparisons) {
    const regressed = item.deltaPercent !== undefined && item.deltaPercent > threshold;
    const missing = item.base === undefined || item.head === undefined;
    if (regressed) regressions += 1;
    if (missing) missingResults += 1;
    const verdict = missing ? "missing" : regressed ? "REGRESSION" : "ok";
    lines.push(
      `| ${item.operation} | ${item.fixture} | ${formatMs(item.base?.medianMs)} | ${formatMs(item.base?.madMs)} | ${formatMs(item.head?.medianMs)} | ${formatMs(item.head?.madMs)} | ${formatDelta(item.deltaPercent)} | ${verdict} |`,
    );
  }
  const report = `${lines.join("\n")}\n`;
  const reportPath = join(scriptDirectory, "output", `${label}.md`);
  writeFileSync(reportPath, report);
  console.log(report);
  console.log(`[compare] report written to ${reportPath}`);

  if (missingResults > 0) {
    console.error(`[compare] ${missingResults} measurement(s) missing; comparison incomplete`);
    process.exitCode = 1;
  }
  if (regressions > 0) {
    console.error(`[compare] ${regressions} operation(s) regressed beyond +${threshold}%`);
    process.exitCode = 1;
  }
}

main();
