/**
 * Single-workload runner for the SheetKit 1.0 release performance gate.
 *
 * The file intentionally uses APIs shared by v0.5.1, the remediation baseline,
 * and 1.0. Copy this source unchanged into each isolated revision.
 */

import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { cpus, totalmem } from 'node:os';
import { dirname, resolve } from 'node:path';
import { performance } from 'node:perf_hooks';
import { Workbook } from '@sheetkit/node';
import { decodeRowsBuffer } from '@sheetkit/node/buffer-codec.js';

const OPERATIONS = [
  'ordinary-open',
  'encrypted-open',
  'scalar-calculation',
  'dense-traversal',
  'streaming',
  'output-csv',
  'output-json',
  'output-html',
  'output-svg',
] as const;

type Operation = (typeof OPERATIONS)[number];

interface GateResult {
  schemaVersion: 1;
  operation: Operation;
  fixture: string;
  samples: number;
  warmups: number;
  timesMs: number[];
  medianMs: number;
  madMs: number;
  environment: {
    platform: NodeJS.Platform;
    architecture: string;
    node: string;
    cpu: string;
    ramBytes: number;
  };
}

/** Keeps each sample result reachable so the work is not optimized away. */
const retained: { value: unknown } = { value: undefined };

function usage(): string {
  return [
    'Usage:',
    '  SHEETKIT_GATE_OPERATION=<operation> \\',
    '  SHEETKIT_GATE_FIXTURE=<xlsx> \\',
    '  SHEETKIT_GATE_OUTPUT=<json> \\',
    '  [SHEETKIT_GATE_SAMPLES=20] [SHEETKIT_GATE_WARMUPS=3] \\',
    '  [SHEETKIT_GATE_ENCRYPTED_FIXTURE=<xlsx>] \\',
    '  [SHEETKIT_GATE_PASSWORD=<password>] \\',
    '  node --expose-gc --import tsx release-gate.ts',
    '',
    'Encrypted fixture generation (run once from the release candidate):',
    '  SHEETKIT_GATE_FIXTURE=<plain-xlsx> \\',
    '  SHEETKIT_GATE_ENCRYPTED_FIXTURE=<encrypted-xlsx> \\',
    '  SHEETKIT_GATE_PASSWORD=<password> \\',
    '  node --import tsx release-gate.ts --generate-encrypted',
    '',
    `Operations: ${OPERATIONS.join(', ')}`,
  ].join('\n');
}

function requiredEnvironment(name: string): string {
  const value = process.env[name];
  if (value === undefined || value.length === 0) {
    throw new Error(`${name} is required`);
  }
  return value;
}

function positiveIntegerEnvironment(name: string, fallback: number): number {
  const raw = process.env[name];
  if (raw === undefined) return fallback;
  const value = Number(raw);
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new Error(`${name} must be a non-negative safe integer`);
  }
  return value;
}

function operationEnvironment(): Operation {
  const value = requiredEnvironment('SHEETKIT_GATE_OPERATION');
  if (!OPERATIONS.includes(value as Operation)) {
    throw new Error(`Unsupported SHEETKIT_GATE_OPERATION: ${value}`);
  }
  return value as Operation;
}

function median(values: readonly number[]): number {
  const sorted = [...values].sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
}

function mad(values: readonly number[], center: number): number {
  return median(values.map((value) => Math.abs(value - center)));
}

async function consumeStream(workbook: Workbook): Promise<number> {
  const reader = await workbook.openSheetReader('Sheet1', { batchSize: 1_000 });
  let rowCount = 0;
  for await (const batch of reader) {
    rowCount += batch.length;
  }
  return rowCount;
}

function scalarCalculation(workbook: Workbook): number {
  const formulas = ['SUM(1,2,3,4,5)', 'IF(10>5,10,5)', 'ROUND(1234.5678,2)', 'ABS(-42)+SQRT(81)'];
  for (let iteration = 0; iteration < 250; iteration += 1) {
    for (const formula of formulas) {
      retained.value = workbook.evaluateFormula('Sheet1', formula);
    }
  }
  return formulas.length * 250;
}

async function prepareOperation(
  operation: Operation,
  fixture: string,
): Promise<() => void | Promise<void>> {
  switch (operation) {
    case 'ordinary-open':
      return () => {
        retained.value = Workbook.openSync(fixture);
      };
    case 'encrypted-open': {
      const encryptedFixture = resolve(requiredEnvironment('SHEETKIT_GATE_ENCRYPTED_FIXTURE'));
      const password = requiredEnvironment('SHEETKIT_GATE_PASSWORD');
      return () => {
        retained.value = Workbook.openWithPasswordSync(encryptedFixture, password);
      };
    }
    case 'scalar-calculation': {
      const workbook = new Workbook();
      return () => {
        retained.value = scalarCalculation(workbook);
      };
    }
    case 'dense-traversal': {
      const workbook = Workbook.openSync(fixture);
      return () => {
        retained.value = decodeRowsBuffer(workbook.getRowsBufferV2('Sheet1'));
      };
    }
    case 'streaming': {
      const workbook = await Workbook.open(fixture, { readMode: 'stream' });
      return async () => {
        retained.value = await consumeStream(workbook);
      };
    }
    case 'output-csv': {
      const workbook = Workbook.openSync(fixture);
      return () => {
        retained.value = workbook.toCSV('Sheet1');
      };
    }
    case 'output-json': {
      const workbook = Workbook.openSync(fixture);
      return () => {
        retained.value = workbook.toJSON('Sheet1');
      };
    }
    case 'output-html': {
      const workbook = Workbook.openSync(fixture);
      return () => {
        retained.value = workbook.toHTML('Sheet1');
      };
    }
    case 'output-svg': {
      const workbook = Workbook.openSync(fixture);
      return () => {
        retained.value = workbook.renderToSvg({ sheetName: 'Sheet1' });
      };
    }
  }
}

async function main(): Promise<void> {
  if (process.argv.includes('--help')) {
    console.log(usage());
    return;
  }
  if (process.argv.includes('--generate-encrypted')) {
    const fixture = resolve(requiredEnvironment('SHEETKIT_GATE_FIXTURE'));
    const encryptedFixture = resolve(requiredEnvironment('SHEETKIT_GATE_ENCRYPTED_FIXTURE'));
    const password = requiredEnvironment('SHEETKIT_GATE_PASSWORD');
    if (existsSync(encryptedFixture)) {
      throw new Error(`Encrypted fixture already exists: ${encryptedFixture}`);
    }
    mkdirSync(dirname(encryptedFixture), { recursive: true });
    Workbook.openSync(fixture).saveWithPasswordSync(encryptedFixture, password);
    console.log(encryptedFixture);
    return;
  }

  const operation = operationEnvironment();
  const fixture = resolve(requiredEnvironment('SHEETKIT_GATE_FIXTURE'));
  const output = resolve(requiredEnvironment('SHEETKIT_GATE_OUTPUT'));
  const samples = positiveIntegerEnvironment('SHEETKIT_GATE_SAMPLES', 20);
  const warmups = positiveIntegerEnvironment('SHEETKIT_GATE_WARMUPS', 3);
  if (samples === 0) {
    throw new Error('SHEETKIT_GATE_SAMPLES must be greater than zero');
  }

  const run = await prepareOperation(operation, fixture);
  for (let index = 0; index < warmups; index += 1) {
    retained.value = undefined;
    globalThis.gc?.();
    await run();
  }

  const timesMs: number[] = [];
  for (let index = 0; index < samples; index += 1) {
    retained.value = undefined;
    globalThis.gc?.();
    const started = performance.now();
    await run();
    timesMs.push(performance.now() - started);
  }

  const medianMs = median(timesMs);
  const result: GateResult = {
    schemaVersion: 1,
    operation,
    fixture,
    samples,
    warmups,
    timesMs,
    medianMs,
    madMs: mad(timesMs, medianMs),
    environment: {
      platform: process.platform,
      architecture: process.arch,
      node: process.version,
      cpu: cpus()[0]?.model ?? 'unknown',
      ramBytes: totalmem(),
    },
  };
  mkdirSync(dirname(output), { recursive: true });
  writeFileSync(output, `${JSON.stringify(result, null, 2)}\n`, { flag: 'wx' });
  console.log(JSON.stringify(result));
}

await main();
