import { readFile, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { Workbook } from '../index.js';

const TEST_DIR = import.meta.dirname;

function tmpFile(name: string) {
  return join(TEST_DIR, name);
}

async function cleanup(...files: string[]) {
  for (const f of files) {
    await unlink(f).catch(() => {});
  }
}

describe('VML Comment Compatibility', () => {
  const out = tmpFile('vml-comments-test.xlsx');
  const out2 = tmpFile('vml-comments-test2.xlsx');
  afterEach(async () => cleanup(out, out2));

  it('should produce VML parts when adding comments', async () => {
    const wb = new Workbook();
    wb.addComment('Sheet1', { cell: 'A1', author: 'Test', text: 'Hello' });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const comments = wb2.getComments('Sheet1');
    expect(comments.length).toBe(1);
    expect(comments[0].cell).toBe('A1');
    expect(comments[0].text).toBe('Hello');
    expect(comments[0].author).toBe('Test');
  });

  it('should roundtrip comments through save/open cycle', async () => {
    const wb = new Workbook();
    wb.addComment('Sheet1', { cell: 'B3', author: 'Alice', text: 'Comment 1' });
    wb.addComment('Sheet1', { cell: 'D7', author: 'Bob', text: 'Comment 2' });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const comments = wb2.getComments('Sheet1');
    expect(comments.length).toBe(2);

    const c1 = comments.find((c: { cell: string }) => c.cell === 'B3');
    const c2 = comments.find((c: { cell: string }) => c.cell === 'D7');
    expect(c1).toBeDefined();
    expect(c1?.author).toBe('Alice');
    expect(c1?.text).toBe('Comment 1');
    expect(c2).toBeDefined();
    expect(c2?.author).toBe('Bob');
    expect(c2?.text).toBe('Comment 2');
  });

  it('should preserve comments through double save/open roundtrip', async () => {
    const wb = new Workbook();
    wb.addComment('Sheet1', { cell: 'A1', author: 'Author', text: 'Persist' });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    await wb2.save(out2);

    const wb3 = await Workbook.open(out2);
    const comments = wb3.getComments('Sheet1');
    expect(comments.length).toBe(1);
    expect(comments[0].text).toBe('Persist');
  });

  it('should handle removing all comments gracefully', async () => {
    const wb = new Workbook();
    wb.addComment('Sheet1', { cell: 'A1', author: 'Author', text: 'Temp' });
    wb.removeComment('Sheet1', 'A1');
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const comments = wb2.getComments('Sheet1');
    expect(comments.length).toBe(0);
  });

  it('should add comments after opening an existing workbook', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Data');
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    wb2.addComment('Sheet1', { cell: 'A1', author: 'Reviewer', text: 'Needs review' });
    await wb2.save(out2);

    const wb3 = await Workbook.open(out2);
    const comments = wb3.getComments('Sheet1');
    expect(comments.length).toBe(1);
    expect(comments[0].text).toBe('Needs review');
  });

  it('should handle comments on multiple sheets', async () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    wb.addComment('Sheet1', { cell: 'A1', author: 'Author', text: 'Sheet1 comment' });
    wb.addComment('Sheet2', { cell: 'B2', author: 'Author', text: 'Sheet2 comment' });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const c1 = wb2.getComments('Sheet1');
    const c2 = wb2.getComments('Sheet2');
    expect(c1.length).toBe(1);
    expect(c1[0].text).toBe('Sheet1 comment');
    expect(c2.length).toBe(1);
    expect(c2[0].text).toBe('Sheet2 comment');
  });

  it('should partially remove comments while preserving others', async () => {
    const wb = new Workbook();
    wb.addComment('Sheet1', { cell: 'A1', author: 'Alice', text: 'Keep' });
    wb.addComment('Sheet1', { cell: 'B2', author: 'Bob', text: 'Remove' });
    wb.removeComment('Sheet1', 'B2');
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const comments = wb2.getComments('Sheet1');
    expect(comments.length).toBe(1);
    expect(comments[0].cell).toBe('A1');
    expect(comments[0].text).toBe('Keep');
  });

  it('should verify VML content is included in saved file', async () => {
    const wb = new Workbook();
    wb.addComment('Sheet1', { cell: 'C5', author: 'Tester', text: 'VML test' });
    await wb.save(out);

    const bytes = await readFile(out);
    const content = bytes.toString('latin1');
    expect(content).toContain('vmlDrawing');
  });
});
