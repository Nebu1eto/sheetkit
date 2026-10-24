import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join } from 'node:path';

const BENCHMARK_DIR = join(import.meta.dirname, '..', '..', '..', 'benchmarks', 'node');
const FIXTURE_DIR = join(BENCHMARK_DIR, 'fixtures');
const FIXTURES = [
  'comments.xlsx',
  'data-validation.xlsx',
  'formulas.xlsx',
  'heavy-styles.xlsx',
  'large-data.xlsx',
  'merged-cells.xlsx',
  'mixed-workload.xlsx',
  'multi-sheet.xlsx',
  'scale-100k.xlsx',
  'scale-10k.xlsx',
  'scale-1k.xlsx',
  'strings.xlsx',
];

/**
 * Generate the deterministic benchmark fixtures that the parity and identity
 * suites read, so a clean checkout runs them instead of failing or skipping.
 */
export default function setup(): void {
  if (FIXTURES.every((name) => existsSync(join(FIXTURE_DIR, name)))) return;
  execFileSync(process.execPath, ['--import', 'tsx', 'generate-fixtures.ts'], {
    cwd: BENCHMARK_DIR,
    stdio: 'inherit',
  });
}
