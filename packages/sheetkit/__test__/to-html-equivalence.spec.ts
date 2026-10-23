import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { isResourceLimitError, type ToHtmlOptions, Workbook } from '../index.js';
import { type CellValue, SheetData } from '../sheet-data.js';

const FIXTURE_DIR = join(import.meta.dirname, '..', '..', '..', 'benchmarks', 'node', 'fixtures');
const FIXTURES = ['mixed-workload.xlsx', 'strings.xlsx', 'merged-cells.xlsx', 'comments.xlsx'];
const OPTION_CASES: readonly [string, ToHtmlOptions | undefined][] = [
  ['default', undefined],
  ['default with className', { className: 'grid "x" <y>' }],
  ['dense', { mode: 'dense' }],
  ['dense with className', { mode: 'dense', className: 'grid' }],
];

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function cellValueToString(v: CellValue): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v.type === 'date') return v.iso ?? String(v.serial);
  return String(v);
}

/** Reference toHTML: the pre-limit loop from 73819d8, without output-byte accounting. */
function referenceToHtml(workbook: Workbook, sheet: string, options?: ToHtmlOptions): string {
  const className = options?.className;
  const buf =
    options?.mode === 'dense'
      ? workbook.getRowsBuffer(sheet, { mode: options.mode })
      : workbook.getRowsBufferV2(sheet, options?.mode === undefined ? undefined : options);
  const sd = new SheetData(buf);
  const rows =
    options?.mode === 'dense'
      ? sd.toArray()
      : [...sd.occupiedRows()].map(({ cells }) => cells.map(({ value }) => value));

  const classAttr = className ? ` class="${escapeHtml(className)}"` : '';
  const parts: string[] = [`<table${classAttr}>`];

  for (const row of rows) {
    parts.push('<tr>');
    for (const cell of row) {
      const text = escapeHtml(cellValueToString(cell));
      parts.push(`<td>${text}</td>`);
    }
    parts.push('</tr>');
  }
  parts.push('</table>');
  return parts.join('');
}

describe('toHTML byte-for-byte equivalence with the reference loop', () => {
  for (const fixture of FIXTURES) {
    const path = join(FIXTURE_DIR, fixture);
    it.skipIf(!existsSync(path))(`renders ${fixture} identically`, () => {
      const workbook = Workbook.openSync(path);
      for (const sheet of workbook.sheetNames) {
        for (const [label, options] of OPTION_CASES) {
          const expected = referenceToHtml(workbook, sheet, options);
          expect(workbook.toHTML(sheet, options), `${fixture}/${sheet}/${label}`).toBe(expected);
        }
      }
    });
  }

  it('renders escaped, multibyte, sparse, and empty-row content identically', () => {
    const workbook = new Workbook();
    workbook.setSheetData('Sheet1', [
      ['<&>"\'', 'caf\u00e9 \u4e2d\u6587 \u{1F600}', 42],
      [],
      [true],
    ]);
    workbook.setCellValue('Sheet1', 'E7', 'tail');
    for (const [label, options] of OPTION_CASES) {
      expect(workbook.toHTML('Sheet1', options), label).toBe(
        referenceToHtml(workbook, 'Sheet1', options),
      );
    }
  });
});

describe('toHTML output limit boundaries', () => {
  it('throws at 200 bytes and succeeds at exactly the full output byte length', () => {
    const workbook = new Workbook();
    const sheet = 'Sheet1';
    // Escaping expands each '&' fivefold, so HTML exceeds 200 bytes while the native buffer does not.
    workbook.setSheetData(sheet, [['&'.repeat(20)], ['\u00e9&'.repeat(5)], ['tail']]);
    const html = workbook.toHTML(sheet);
    const exact = Buffer.byteLength(html);
    expect(exact).toBeGreaterThan(200);

    let error: unknown;
    try {
      workbook.toHTML(sheet, { limits: { maxOutputBytes: 200 } });
    } catch (caught) {
      error = caught;
    }
    expect(isResourceLimitError(error)).toBe(true);
    if (!isResourceLimitError(error)) throw new Error('Expected a resource limit error');
    expect(error).toMatchObject({
      operation: 'export',
      category: 'output',
      dimension: 'output_bytes',
      limit: 200n,
      stage: 'serialize',
    });
    expect(error.observed).toBeGreaterThan(200n);

    expect(workbook.toHTML(sheet, { limits: { maxOutputBytes: exact } })).toBe(html);
  });
});
