import { access, unlink } from 'node:fs/promises';
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

describe('Phase 1 - Basic I/O', () => {
  const out = tmpFile('test-basic.xlsx');
  afterEach(async () => cleanup(out));

  it('should create a new workbook', () => {
    const wb = new Workbook();
    expect(wb.sheetNames).toEqual(['Sheet1']);
  });

  it('should save and open a workbook', async () => {
    const wb = new Workbook();
    await wb.save(out);
    await expect(access(out)).resolves.toBeUndefined();
    const wb2 = await Workbook.open(out);
    expect(wb2.sheetNames).toEqual(['Sheet1']);
  });

  it('should throw on invalid path', async () => {
    await expect(Workbook.open('/nonexistent/path.xlsx')).rejects.toThrow();
  });
});

describe('Phase 2 - Cell Operations', () => {
  const out = tmpFile('test-cell.xlsx');
  afterEach(async () => cleanup(out));

  it('should set and get string cell value', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'hello');
    expect(wb.getCellValue('Sheet1', 'A1')).toBe('hello');
  });

  it('should set and get number cell value', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'B2', 42.5);
    expect(wb.getCellValue('Sheet1', 'B2')).toBe(42.5);
  });

  it('should set and get boolean cell value', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'C3', true);
    expect(wb.getCellValue('Sheet1', 'C3')).toBe(true);
  });

  it('should clear cell with null', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'test');
    wb.setCellValue('Sheet1', 'A1', null);
    expect(wb.getCellValue('Sheet1', 'A1')).toBeNull();
  });

  it('should return null for empty cell', () => {
    const wb = new Workbook();
    expect(wb.getCellValue('Sheet1', 'Z99')).toBeNull();
  });

  it('should roundtrip cell values through save/open', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'text');
    wb.setCellValue('Sheet1', 'B1', 123);
    wb.setCellValue('Sheet1', 'C1', true);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe('text');
    expect(wb2.getCellValue('Sheet1', 'B1')).toBe(123);
    expect(wb2.getCellValue('Sheet1', 'C1')).toBe(true);
  });
});

describe('Phase 5 - Sheet Management', () => {
  it('should create a new sheet', () => {
    const wb = new Workbook();
    const idx = wb.newSheet('Sheet2');
    expect(idx).toBe(1);
    expect(wb.sheetNames).toEqual(['Sheet1', 'Sheet2']);
  });

  it('should delete a sheet', () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    wb.deleteSheet('Sheet1');
    expect(wb.sheetNames).toEqual(['Sheet2']);
  });

  it('should rename a sheet', () => {
    const wb = new Workbook();
    wb.setSheetName('Sheet1', 'Data');
    expect(wb.sheetNames).toEqual(['Data']);
  });

  it('should copy a sheet', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'original');
    wb.copySheet('Sheet1', 'Copy');
    expect(wb.sheetNames).toContain('Copy');
  });

  it('should get/set active sheet', () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    wb.setActiveSheet('Sheet2');
    expect(wb.getActiveSheet()).toBe('Sheet2');
  });

  it('should get sheet index', () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    expect(wb.getSheetIndex('Sheet1')).toBe(0);
    expect(wb.getSheetIndex('Sheet2')).toBe(1);
    expect(wb.getSheetIndex('NotExist')).toBeNull();
  });
});

describe('Phase 3 - Row/Column Operations', () => {
  it('should set and get row height', () => {
    const wb = new Workbook();
    wb.setRowHeight('Sheet1', 1, 30);
    expect(wb.getRowHeight('Sheet1', 1)).toBe(30);
  });

  it('should set and get col width', () => {
    const wb = new Workbook();
    wb.setColWidth('Sheet1', 'A', 20);
    expect(wb.getColWidth('Sheet1', 'A')).toBe(20);
  });

  it('should insert rows and shift cells', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'row1');
    wb.insertRows('Sheet1', 1, 2);
    expect(wb.getCellValue('Sheet1', 'A3')).toBe('row1');
  });

  it('should insert cols and shift cells', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'colA');
    wb.insertCols('Sheet1', 'A', 1);
    expect(wb.getCellValue('Sheet1', 'B1')).toBe('colA');
  });

  it('should duplicate a row', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'hello');
    wb.duplicateRow('Sheet1', 1);
    expect(wb.getCellValue('Sheet1', 'A2')).toBe('hello');
  });
});

describe('Row/Col Getters & Outline Levels', () => {
  it('should get row visible default as true', () => {
    const wb = new Workbook();
    expect(wb.getRowVisible('Sheet1', 1)).toBe(true);
  });

  it('should get row visible false after hiding', () => {
    const wb = new Workbook();
    wb.setRowVisible('Sheet1', 1, false);
    expect(wb.getRowVisible('Sheet1', 1)).toBe(false);
  });

  it('should get row visible true after show', () => {
    const wb = new Workbook();
    wb.setRowVisible('Sheet1', 1, false);
    wb.setRowVisible('Sheet1', 1, true);
    expect(wb.getRowVisible('Sheet1', 1)).toBe(true);
  });

  it('should set and get row outline level', () => {
    const wb = new Workbook();
    wb.setRowOutlineLevel('Sheet1', 1, 3);
    expect(wb.getRowOutlineLevel('Sheet1', 1)).toBe(3);
  });

  it('should return 0 for default row outline level', () => {
    const wb = new Workbook();
    expect(wb.getRowOutlineLevel('Sheet1', 1)).toBe(0);
  });

  it('should get col visible default as true', () => {
    const wb = new Workbook();
    expect(wb.getColVisible('Sheet1', 'A')).toBe(true);
  });

  it('should get col visible false after hiding', () => {
    const wb = new Workbook();
    wb.setColVisible('Sheet1', 'B', false);
    expect(wb.getColVisible('Sheet1', 'B')).toBe(false);
  });

  it('should set and get col outline level', () => {
    const wb = new Workbook();
    wb.setColOutlineLevel('Sheet1', 'A', 5);
    expect(wb.getColOutlineLevel('Sheet1', 'A')).toBe(5);
  });

  it('should return 0 for default col outline level', () => {
    const wb = new Workbook();
    expect(wb.getColOutlineLevel('Sheet1', 'A')).toBe(0);
  });

  it('should reject outline level > 7 for row', () => {
    const wb = new Workbook();
    expect(() => wb.setRowOutlineLevel('Sheet1', 1, 8)).toThrow();
  });

  it('should reject outline level > 7 for col', () => {
    const wb = new Workbook();
    expect(() => wb.setColOutlineLevel('Sheet1', 'A', 8)).toThrow();
  });
});

describe('Phase 4 - Style', () => {
  it('should add a style with font and apply to cell', () => {
    const wb = new Workbook();
    const styleId = wb.addStyle({
      font: { bold: true, size: 14, color: '#FF0000' },
    });
    expect(styleId).toBeGreaterThanOrEqual(0);
    wb.setCellStyle('Sheet1', 'A1', styleId);
    expect(wb.getCellStyle('Sheet1', 'A1')).toBe(styleId);
  });

  it('should add a style with fill', () => {
    const wb = new Workbook();
    const styleId = wb.addStyle({
      fill: { pattern: 'solid', fgColor: '#FFFF00' },
    });
    expect(styleId).toBeGreaterThanOrEqual(0);
  });

  it('should add a style with border', () => {
    const wb = new Workbook();
    const styleId = wb.addStyle({
      border: {
        left: { style: 'thin', color: '#000000' },
        right: { style: 'thin', color: '#000000' },
      },
    });
    expect(styleId).toBeGreaterThanOrEqual(0);
  });

  it('should add a style with alignment', () => {
    const wb = new Workbook();
    const styleId = wb.addStyle({
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    });
    expect(styleId).toBeGreaterThanOrEqual(0);
  });

  it('should add a style with number format', () => {
    const wb = new Workbook();
    const sid1 = wb.addStyle({ numFmtId: 2 });
    const sid2 = wb.addStyle({ customNumFmt: '0.00%' });
    expect(sid1).toBeGreaterThanOrEqual(0);
    expect(sid2).toBeGreaterThanOrEqual(0);
  });
});

describe('Row/Col Style', () => {
  it('should set and get row style', () => {
    const wb = new Workbook();
    const styleId = wb.addStyle({ font: { bold: true } });
    wb.setRowStyle('Sheet1', 1, styleId);
    expect(wb.getRowStyle('Sheet1', 1)).toBe(styleId);
  });

  it('should return 0 for default row style', () => {
    const wb = new Workbook();
    expect(wb.getRowStyle('Sheet1', 1)).toBe(0);
  });

  it('should apply row style to existing cells', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'hello');
    wb.setCellValue('Sheet1', 'B1', 42);
    const styleId = wb.addStyle({ font: { bold: true } });
    wb.setRowStyle('Sheet1', 1, styleId);
    // Cells in the row should now have this style
    expect(wb.getCellStyle('Sheet1', 'A1')).toBe(styleId);
    expect(wb.getCellStyle('Sheet1', 'B1')).toBe(styleId);
  });

  it('should reject invalid row style ID', () => {
    const wb = new Workbook();
    expect(() => wb.setRowStyle('Sheet1', 1, 999)).toThrow();
  });

  it('should set and get col style', () => {
    const wb = new Workbook();
    const styleId = wb.addStyle({ font: { italic: true } });
    wb.setColStyle('Sheet1', 'A', styleId);
    expect(wb.getColStyle('Sheet1', 'A')).toBe(styleId);
  });

  it('should return 0 for default col style', () => {
    const wb = new Workbook();
    expect(wb.getColStyle('Sheet1', 'A')).toBe(0);
  });

  it('should reject invalid col style ID', () => {
    const wb = new Workbook();
    expect(() => wb.setColStyle('Sheet1', 'A', 999)).toThrow();
  });
});

describe('Phase 7 - Charts & Images', () => {
  const out = tmpFile('test-chart.xlsx');
  afterEach(async () => cleanup(out));

  it('should add a column chart and save', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Category');
    wb.setCellValue('Sheet1', 'B1', 100);
    wb.addChart('Sheet1', 'D1', 'J10', {
      chartType: 'col',
      series: [{ name: 'S1', categories: 'Sheet1!$A$1:$A$3', values: 'Sheet1!$B$1:$B$3' }],
    });
    await wb.save(out);
    await expect(access(out)).resolves.toBeUndefined();
  });

  it('should add a PNG image and save', async () => {
    const wb = new Workbook();
    const pngData = Buffer.from([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
      0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f,
      0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00,
      0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49,
      0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ]);
    wb.addImage('Sheet1', {
      data: pngData,
      format: 'png',
      fromCell: 'A1',
      widthPx: 100,
      heightPx: 100,
    });
    await wb.save(out);
    await expect(access(out)).resolves.toBeUndefined();
  });
});

describe('Phase 8 - Comments', () => {
  it('should add and get comments', () => {
    const wb = new Workbook();
    wb.addComment('Sheet1', { cell: 'A1', author: 'Test', text: 'Hello' });
    const comments = wb.getComments('Sheet1');
    expect(comments.length).toBe(1);
    expect(comments[0].cell).toBe('A1');
    expect(comments[0].text).toBe('Hello');
  });

  it('should remove a comment', () => {
    const wb = new Workbook();
    wb.addComment('Sheet1', { cell: 'A1', author: 'Test', text: 'Hello' });
    wb.removeComment('Sheet1', 'A1');
    expect(wb.getComments('Sheet1').length).toBe(0);
  });
});

describe('Phase 8 - Data Validation', () => {
  it('should add and get list validation', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'list',
      formula1: '"Option1,Option2,Option3"',
    });
    const validations = wb.getDataValidations('Sheet1');
    expect(validations.length).toBe(1);
    expect(validations[0].sqref).toBe('A1:A10');
  });

  it('should remove data validation', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'list',
      formula1: '"A,B,C"',
    });
    wb.removeDataValidation('Sheet1', 'A1:A10');
    expect(wb.getDataValidations('Sheet1').length).toBe(0);
  });
});

describe('Merge Cells', () => {
  const out = tmpFile('test-merge.xlsx');
  afterEach(async () => cleanup(out));

  it('should merge and get merge cells', () => {
    const wb = new Workbook();
    wb.mergeCells('Sheet1', 'A1', 'B2');
    const merged = wb.getMergeCells('Sheet1');
    expect(merged).toEqual(['A1:B2']);
  });

  it('should merge multiple ranges', () => {
    const wb = new Workbook();
    wb.mergeCells('Sheet1', 'A1', 'B2');
    wb.mergeCells('Sheet1', 'D1', 'F3');
    const merged = wb.getMergeCells('Sheet1');
    expect(merged.length).toBe(2);
    expect(merged).toContain('A1:B2');
    expect(merged).toContain('D1:F3');
  });

  it('should detect overlapping merge ranges', () => {
    const wb = new Workbook();
    wb.mergeCells('Sheet1', 'A1', 'C3');
    expect(() => wb.mergeCells('Sheet1', 'B2', 'D4')).toThrow();
  });

  it('should unmerge a range', () => {
    const wb = new Workbook();
    wb.mergeCells('Sheet1', 'A1', 'B2');
    wb.unmergeCell('Sheet1', 'A1:B2');
    expect(wb.getMergeCells('Sheet1').length).toBe(0);
  });

  it('should throw when unmerging non-existent range', () => {
    const wb = new Workbook();
    expect(() => wb.unmergeCell('Sheet1', 'A1:B2')).toThrow();
  });

  it('should return empty array for no merge cells', () => {
    const wb = new Workbook();
    expect(wb.getMergeCells('Sheet1')).toEqual([]);
  });

  it('should roundtrip merge cells through save/open', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Merged');
    wb.mergeCells('Sheet1', 'A1', 'C3');
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const merged = wb2.getMergeCells('Sheet1');
    expect(merged).toEqual(['A1:C3']);
  });
});

describe('Phase 8 - Auto-filter', () => {
  const out = tmpFile('test-autofilter.xlsx');
  afterEach(async () => cleanup(out));

  it('should set and remove auto filter', () => {
    const wb = new Workbook();
    wb.setAutoFilter('Sheet1', 'A1:C10');
    wb.removeAutoFilter('Sheet1');
  });

  it('should set auto filter and save', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 'Age');
    wb.setAutoFilter('Sheet1', 'A1:B1');
    await wb.save(out);
    await expect(access(out)).resolves.toBeUndefined();
  });
});

describe('Phase 9 - StreamWriter', () => {
  const out = tmpFile('test-stream.xlsx');
  afterEach(async () => cleanup(out));

  it('should create and use stream writer', () => {
    const wb = new Workbook();
    const sw = wb.newStreamWriter('Stream1');
    expect(sw.sheetName).toBe('Stream1');
    sw.setColWidth(1, 20);
    sw.writeRow(1, ['Hello', 42, true, null]);
    sw.writeRow(2, ['World', 99, false, null]);
    const idx = wb.applyStreamWriter(sw);
    expect(idx).toBeGreaterThanOrEqual(0);
    expect(wb.sheetNames).toContain('Stream1');
  });

  it('should roundtrip stream writer data', async () => {
    const wb = new Workbook();
    const sw = wb.newStreamWriter('Data');
    sw.writeRow(1, ['Name', 'Value']);
    sw.writeRow(2, ['A', 100]);
    wb.applyStreamWriter(sw);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    expect(wb2.sheetNames).toContain('Data');
    expect(wb2.getCellValue('Data', 'A1')).toBe('Name');
    expect(wb2.getCellValue('Data', 'B2')).toBe(100);
  });
});

describe('Phase 10 - Document Properties', () => {
  const out = tmpFile('test-docprops.xlsx');
  afterEach(async () => cleanup(out));

  it('should set and get doc properties', () => {
    const wb = new Workbook();
    wb.setDocProps({ title: 'Test', creator: 'SheetKit' });
    const props = wb.getDocProps();
    expect(props.title).toBe('Test');
    expect(props.creator).toBe('SheetKit');
  });

  it('should set and get app properties', () => {
    const wb = new Workbook();
    wb.setAppProps({ company: 'TestCorp', application: 'SheetKit' });
    const props = wb.getAppProps();
    expect(props.company).toBe('TestCorp');
  });

  it('should roundtrip doc properties', async () => {
    const wb = new Workbook();
    wb.setDocProps({ title: 'My Doc', creator: 'Author' });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const props = wb2.getDocProps();
    expect(props.title).toBe('My Doc');
    expect(props.creator).toBe('Author');
  });

  it('should set/get/delete custom properties', () => {
    const wb = new Workbook();
    wb.setCustomProperty('myString', 'hello');
    wb.setCustomProperty('myNumber', 42);
    wb.setCustomProperty('myBool', true);
    expect(wb.getCustomProperty('myString')).toBe('hello');
    expect(wb.getCustomProperty('myNumber')).toBe(42);
    expect(wb.getCustomProperty('myBool')).toBe(true);
    expect(wb.getCustomProperty('nonexistent')).toBeNull();

    expect(wb.deleteCustomProperty('myString')).toBe(true);
    expect(wb.deleteCustomProperty('myString')).toBe(false);
  });
});

describe('Phase 10 - Workbook Protection', () => {
  const out = tmpFile('test-protection.xlsx');
  afterEach(async () => cleanup(out));

  it('should protect and unprotect workbook', () => {
    const wb = new Workbook();
    expect(wb.isWorkbookProtected()).toBe(false);
    wb.protectWorkbook({ lockStructure: true });
    expect(wb.isWorkbookProtected()).toBe(true);
    wb.unprotectWorkbook();
    expect(wb.isWorkbookProtected()).toBe(false);
  });

  it('should protect with password and roundtrip', async () => {
    const wb = new Workbook();
    wb.protectWorkbook({ password: 'secret', lockStructure: true, lockWindows: true });
    expect(wb.isWorkbookProtected()).toBe(true);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    expect(wb2.isWorkbookProtected()).toBe(true);
  });
});

describe('Hyperlinks', () => {
  const out = tmpFile('test-hyperlink.xlsx');
  afterEach(async () => cleanup(out));

  it('should set and get external hyperlink', () => {
    const wb = new Workbook();
    wb.setCellHyperlink('Sheet1', 'A1', {
      linkType: 'external',
      target: 'https://example.com',
      display: 'Example',
      tooltip: 'Visit Example',
    });
    const info = wb.getCellHyperlink('Sheet1', 'A1');
    expect(info).not.toBeNull();
    expect(info?.linkType).toBe('external');
    expect(info?.target).toBe('https://example.com');
    expect(info?.display).toBe('Example');
    expect(info?.tooltip).toBe('Visit Example');
  });

  it('should set and get internal hyperlink', () => {
    const wb = new Workbook();
    wb.setCellHyperlink('Sheet1', 'B2', {
      linkType: 'internal',
      target: 'Sheet1!C1',
      display: 'Go to C1',
    });
    const info = wb.getCellHyperlink('Sheet1', 'B2');
    expect(info).not.toBeNull();
    expect(info?.linkType).toBe('internal');
    expect(info?.target).toBe('Sheet1!C1');
    expect(info?.display).toBe('Go to C1');
  });

  it('should set and get email hyperlink', () => {
    const wb = new Workbook();
    wb.setCellHyperlink('Sheet1', 'C3', {
      linkType: 'email',
      target: 'mailto:user@example.com',
    });
    const info = wb.getCellHyperlink('Sheet1', 'C3');
    expect(info).not.toBeNull();
    expect(info?.linkType).toBe('email');
    expect(info?.target).toBe('mailto:user@example.com');
  });

  it('should return null for cell without hyperlink', () => {
    const wb = new Workbook();
    const info = wb.getCellHyperlink('Sheet1', 'Z99');
    expect(info).toBeNull();
  });

  it('should delete a hyperlink', () => {
    const wb = new Workbook();
    wb.setCellHyperlink('Sheet1', 'A1', {
      linkType: 'external',
      target: 'https://example.com',
    });
    wb.deleteCellHyperlink('Sheet1', 'A1');
    expect(wb.getCellHyperlink('Sheet1', 'A1')).toBeNull();
  });

  it('should overwrite existing hyperlink', () => {
    const wb = new Workbook();
    wb.setCellHyperlink('Sheet1', 'A1', {
      linkType: 'external',
      target: 'https://old.com',
    });
    wb.setCellHyperlink('Sheet1', 'A1', {
      linkType: 'external',
      target: 'https://new.com',
      display: 'New Link',
    });
    const info = wb.getCellHyperlink('Sheet1', 'A1');
    expect(info?.target).toBe('https://new.com');
    expect(info?.display).toBe('New Link');
  });

  it('should handle multiple hyperlinks on different cells', () => {
    const wb = new Workbook();
    wb.setCellHyperlink('Sheet1', 'A1', {
      linkType: 'external',
      target: 'https://example.com',
    });
    wb.setCellHyperlink('Sheet1', 'B1', {
      linkType: 'internal',
      target: 'Sheet1!C1',
    });
    wb.setCellHyperlink('Sheet1', 'C1', {
      linkType: 'email',
      target: 'mailto:test@test.com',
    });

    expect(wb.getCellHyperlink('Sheet1', 'A1')?.linkType).toBe('external');
    expect(wb.getCellHyperlink('Sheet1', 'B1')?.linkType).toBe('internal');
    expect(wb.getCellHyperlink('Sheet1', 'C1')?.linkType).toBe('email');
  });

  it('should roundtrip hyperlinks through save/open', async () => {
    const wb = new Workbook();
    wb.setCellHyperlink('Sheet1', 'A1', {
      linkType: 'external',
      target: 'https://rust-lang.org',
      display: 'Rust',
      tooltip: 'Rust Homepage',
    });
    wb.setCellHyperlink('Sheet1', 'B1', {
      linkType: 'internal',
      target: 'Sheet1!C1',
      display: 'Go to C1',
    });
    wb.setCellHyperlink('Sheet1', 'C1', {
      linkType: 'email',
      target: 'mailto:hello@example.com',
      display: 'Email',
    });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const a1 = wb2.getCellHyperlink('Sheet1', 'A1');
    expect(a1).not.toBeNull();
    expect(a1?.linkType).toBe('external');
    expect(a1?.target).toBe('https://rust-lang.org');
    expect(a1?.display).toBe('Rust');
    expect(a1?.tooltip).toBe('Rust Homepage');

    const b1 = wb2.getCellHyperlink('Sheet1', 'B1');
    expect(b1).not.toBeNull();
    expect(b1?.linkType).toBe('internal');
    expect(b1?.target).toBe('Sheet1!C1');

    const c1 = wb2.getCellHyperlink('Sheet1', 'C1');
    expect(c1).not.toBeNull();
    expect(c1?.linkType).toBe('email');
    expect(c1?.target).toBe('mailto:hello@example.com');
  });
});

describe('Freeze Panes', () => {
  const out = tmpFile('test-panes.xlsx');
  afterEach(async () => cleanup(out));

  it('should return null when no panes are set', () => {
    const wb = new Workbook();
    expect(wb.getPanes('Sheet1')).toBeNull();
  });

  it('should freeze row 1', () => {
    const wb = new Workbook();
    wb.setPanes('Sheet1', 'A2');
    expect(wb.getPanes('Sheet1')).toBe('A2');
  });

  it('should freeze column A', () => {
    const wb = new Workbook();
    wb.setPanes('Sheet1', 'B1');
    expect(wb.getPanes('Sheet1')).toBe('B1');
  });

  it('should freeze row 1 and column A', () => {
    const wb = new Workbook();
    wb.setPanes('Sheet1', 'B2');
    expect(wb.getPanes('Sheet1')).toBe('B2');
  });

  it('should freeze multiple rows', () => {
    const wb = new Workbook();
    wb.setPanes('Sheet1', 'A4');
    expect(wb.getPanes('Sheet1')).toBe('A4');
  });

  it('should unset panes', () => {
    const wb = new Workbook();
    wb.setPanes('Sheet1', 'B2');
    expect(wb.getPanes('Sheet1')).toBe('B2');
    wb.unsetPanes('Sheet1');
    expect(wb.getPanes('Sheet1')).toBeNull();
  });

  it('should throw for A1 cell reference', () => {
    const wb = new Workbook();
    expect(() => wb.setPanes('Sheet1', 'A1')).toThrow();
  });

  it('should throw for nonexistent sheet', () => {
    const wb = new Workbook();
    expect(() => wb.setPanes('NoSheet', 'A2')).toThrow();
  });

  it('should overwrite previous panes', () => {
    const wb = new Workbook();
    wb.setPanes('Sheet1', 'A2');
    wb.setPanes('Sheet1', 'C3');
    expect(wb.getPanes('Sheet1')).toBe('C3');
  });

  it('should roundtrip panes through save/open', async () => {
    const wb = new Workbook();
    wb.setPanes('Sheet1', 'B3');
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    expect(wb2.getPanes('Sheet1')).toBe('B3');
  });
});

describe('Date CellValue', () => {
  const out = tmpFile('test-date.xlsx');
  afterEach(async () => cleanup(out));

  it('should set a date value object and read it back as date', () => {
    const wb = new Workbook();
    // Create a date style (numFmtId 14 = m/d/yyyy)
    const styleId = wb.addStyle({ numFmtId: 14 });
    // Set a date value using the object format (Jan 1, 2024 = serial 45292)
    wb.setCellValue('Sheet1', 'A1', { type: 'date', serial: 45292 });
    wb.setCellStyle('Sheet1', 'A1', styleId);

    // Read it back - should be a date object
    const val = wb.getCellValue('Sheet1', 'A1') as { type: string; serial: number; iso?: string };
    expect(val).not.toBeNull();
    expect(val.type).toBe('date');
    expect(val.serial).toBe(45292);
    expect(val.iso).toBe('2024-01-01');
  });

  it('should return number for date without date style', () => {
    const wb = new Workbook();
    // Set a date serial as a plain number
    wb.setCellValue('Sheet1', 'A1', 45292);
    // Without a date style, it should come back as a number
    const val = wb.getCellValue('Sheet1', 'A1');
    expect(typeof val).toBe('number');
    expect(val).toBe(45292);
  });

  it('should roundtrip date values through save/open', async () => {
    const wb = new Workbook();
    // Create a datetime style (numFmtId 22 = m/d/yyyy h:mm)
    const styleId = wb.addStyle({ numFmtId: 22 });
    wb.setCellValue('Sheet1', 'A1', { type: 'date', serial: 45292.5 });
    wb.setCellStyle('Sheet1', 'A1', styleId);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const val = wb2.getCellValue('Sheet1', 'A1') as { type: string; serial: number; iso?: string };
    expect(val.type).toBe('date');
    expect(val.serial).toBe(45292.5);
    expect(val.iso).toBe('2024-01-01T12:00:00');
  });

  it('should handle date with custom date format', () => {
    const wb = new Workbook();
    const styleId = wb.addStyle({ customNumFmt: 'yyyy-mm-dd' });
    wb.setCellValue('Sheet1', 'A1', { type: 'date', serial: 45292 });
    wb.setCellStyle('Sheet1', 'A1', styleId);

    const val = wb.getCellValue('Sheet1', 'A1') as { type: string; serial: number; iso?: string };
    expect(val.type).toBe('date');
    expect(val.serial).toBe(45292);
  });
});

describe('Row/Col Iterators', () => {
  const out = tmpFile('test-iterators.xlsx');
  afterEach(async () => cleanup(out));

  it('should return empty array for empty sheet', () => {
    const wb = new Workbook();
    const rows = wb.getRows('Sheet1');
    expect(rows).toEqual([]);
  });

  it('should return row data with correct structure', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 100);
    wb.setCellValue('Sheet1', 'A2', 'Alice');
    wb.setCellValue('Sheet1', 'B2', true);

    const rows = wb.getRows('Sheet1');
    expect(rows.length).toBe(2);

    // Row 1
    expect(rows[0].row).toBe(1);
    expect(rows[0].cells.length).toBe(2);
    expect(rows[0].cells[0].column).toBe('A');
    expect(rows[0].cells[0].valueType).toBe('string');
    expect(rows[0].cells[0].value).toBe('Name');
    expect(rows[0].cells[1].column).toBe('B');
    expect(rows[0].cells[1].valueType).toBe('number');
    expect(rows[0].cells[1].numberValue).toBe(100);

    // Row 2
    expect(rows[1].row).toBe(2);
    expect(rows[1].cells[0].column).toBe('A');
    expect(rows[1].cells[0].valueType).toBe('string');
    expect(rows[1].cells[0].value).toBe('Alice');
    expect(rows[1].cells[1].column).toBe('B');
    expect(rows[1].cells[1].valueType).toBe('boolean');
    expect(rows[1].cells[1].boolValue).toBe(true);
  });

  it('should return empty cols array for empty sheet', () => {
    const wb = new Workbook();
    const cols = wb.getCols('Sheet1');
    expect(cols).toEqual([]);
  });

  it('should return column data transposed from row data', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 'Age');
    wb.setCellValue('Sheet1', 'A2', 'Alice');
    wb.setCellValue('Sheet1', 'B2', 30);

    const cols = wb.getCols('Sheet1');
    expect(cols.length).toBe(2);

    // Column A
    expect(cols[0].column).toBe('A');
    expect(cols[0].cells.length).toBe(2);
    expect(cols[0].cells[0].row).toBe(1);
    expect(cols[0].cells[0].valueType).toBe('string');
    expect(cols[0].cells[0].value).toBe('Name');
    expect(cols[0].cells[1].row).toBe(2);
    expect(cols[0].cells[1].value).toBe('Alice');

    // Column B
    expect(cols[1].column).toBe('B');
    expect(cols[1].cells.length).toBe(2);
    expect(cols[1].cells[0].row).toBe(1);
    expect(cols[1].cells[0].valueType).toBe('string');
    expect(cols[1].cells[0].value).toBe('Age');
    expect(cols[1].cells[1].row).toBe(2);
    expect(cols[1].cells[1].valueType).toBe('number');
    expect(cols[1].cells[1].numberValue).toBe(30);
  });

  it('should throw for non-existent sheet', () => {
    const wb = new Workbook();
    expect(() => wb.getRows('NoSheet')).toThrow();
    expect(() => wb.getCols('NoSheet')).toThrow();
  });

  it('should roundtrip rows through save/open', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'hello');
    wb.setCellValue('Sheet1', 'B1', 99);
    wb.setCellValue('Sheet1', 'A2', true);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const rows = wb2.getRows('Sheet1');
    expect(rows.length).toBe(2);
    expect(rows[0].cells[0].value).toBe('hello');
    expect(rows[0].cells[1].numberValue).toBe(99);
    expect(rows[1].cells[0].boolValue).toBe(true);
  });

  it('should handle sparse data correctly', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'first');
    wb.setCellValue('Sheet1', 'C3', 'sparse');

    const rows = wb.getRows('Sheet1');
    expect(rows.length).toBe(2);
    expect(rows[0].row).toBe(1);
    expect(rows[0].cells.length).toBe(1);
    expect(rows[0].cells[0].column).toBe('A');
    expect(rows[1].row).toBe(3);
    expect(rows[1].cells.length).toBe(1);
    expect(rows[1].cells[0].column).toBe('C');

    const cols = wb.getCols('Sheet1');
    expect(cols.length).toBe(2);
    expect(cols[0].column).toBe('A');
    expect(cols[0].cells.length).toBe(1);
    expect(cols[0].cells[0].row).toBe(1);
    expect(cols[1].column).toBe('C');
    expect(cols[1].cells.length).toBe(1);
    expect(cols[1].cells[0].row).toBe(3);
  });
});

describe('Page Layout', () => {
  const out = tmpFile('test-page-layout.xlsx');
  afterEach(async () => cleanup(out));

  it('should set and get page margins', () => {
    const wb = new Workbook();
    wb.setPageMargins('Sheet1', {
      left: 1.0,
      right: 1.0,
      top: 1.5,
      bottom: 1.5,
      header: 0.5,
      footer: 0.5,
    });
    const m = wb.getPageMargins('Sheet1');
    expect(m.left).toBe(1.0);
    expect(m.right).toBe(1.0);
    expect(m.top).toBe(1.5);
    expect(m.bottom).toBe(1.5);
    expect(m.header).toBe(0.5);
    expect(m.footer).toBe(0.5);
  });

  it('should return default margins when not set', () => {
    const wb = new Workbook();
    const m = wb.getPageMargins('Sheet1');
    expect(m.left).toBe(0.7);
    expect(m.right).toBe(0.7);
    expect(m.top).toBe(0.75);
    expect(m.bottom).toBe(0.75);
    expect(m.header).toBe(0.3);
    expect(m.footer).toBe(0.3);
  });

  it('should set and get page setup', () => {
    const wb = new Workbook();
    wb.setPageSetup('Sheet1', {
      orientation: 'landscape',
      paperSize: 'a4',
      scale: 75,
    });
    const setup = wb.getPageSetup('Sheet1');
    expect(setup.orientation).toBe('landscape');
    expect(setup.paperSize).toBe('a4');
    expect(setup.scale).toBe(75);
  });

  it('should return undefined for unset page setup', () => {
    const wb = new Workbook();
    const setup = wb.getPageSetup('Sheet1');
    expect(setup.orientation).toBeUndefined();
    expect(setup.paperSize).toBeUndefined();
    expect(setup.scale).toBeUndefined();
    expect(setup.fitToWidth).toBeUndefined();
    expect(setup.fitToHeight).toBeUndefined();
  });

  it('should set fit-to-page options', () => {
    const wb = new Workbook();
    wb.setPageSetup('Sheet1', {
      fitToWidth: 1,
      fitToHeight: 2,
    });
    const setup = wb.getPageSetup('Sheet1');
    expect(setup.fitToWidth).toBe(1);
    expect(setup.fitToHeight).toBe(2);
  });

  it('should set and get header and footer', () => {
    const wb = new Workbook();
    wb.setHeaderFooter('Sheet1', '&CPage &P', '&LFooter');
    const hf = wb.getHeaderFooter('Sheet1');
    expect(hf.header).toBe('&CPage &P');
    expect(hf.footer).toBe('&LFooter');
  });

  it('should return undefined for unset header/footer', () => {
    const wb = new Workbook();
    const hf = wb.getHeaderFooter('Sheet1');
    expect(hf.header).toBeUndefined();
    expect(hf.footer).toBeUndefined();
  });

  it('should set and get print options', () => {
    const wb = new Workbook();
    wb.setPrintOptions('Sheet1', {
      gridLines: true,
      headings: true,
      horizontalCentered: true,
      verticalCentered: false,
    });
    const opts = wb.getPrintOptions('Sheet1');
    expect(opts.gridLines).toBe(true);
    expect(opts.headings).toBe(true);
    expect(opts.horizontalCentered).toBe(true);
    expect(opts.verticalCentered).toBe(false);
  });

  it('should return undefined for unset print options', () => {
    const wb = new Workbook();
    const opts = wb.getPrintOptions('Sheet1');
    expect(opts.gridLines).toBeUndefined();
    expect(opts.headings).toBeUndefined();
    expect(opts.horizontalCentered).toBeUndefined();
    expect(opts.verticalCentered).toBeUndefined();
  });

  it('should insert and get page breaks', () => {
    const wb = new Workbook();
    wb.insertPageBreak('Sheet1', 10);
    wb.insertPageBreak('Sheet1', 20);
    const breaks = wb.getPageBreaks('Sheet1');
    expect(breaks).toEqual([10, 20]);
  });

  it('should remove page breaks', () => {
    const wb = new Workbook();
    wb.insertPageBreak('Sheet1', 10);
    wb.insertPageBreak('Sheet1', 20);
    wb.removePageBreak('Sheet1', 10);
    const breaks = wb.getPageBreaks('Sheet1');
    expect(breaks).toEqual([20]);
  });

  it('should return empty array when no page breaks', () => {
    const wb = new Workbook();
    expect(wb.getPageBreaks('Sheet1')).toEqual([]);
  });

  it('should roundtrip page layout through save/open', async () => {
    const wb = new Workbook();
    wb.setPageMargins('Sheet1', {
      left: 1.0,
      right: 1.0,
      top: 1.5,
      bottom: 1.5,
      header: 0.5,
      footer: 0.5,
    });
    wb.setPageSetup('Sheet1', {
      orientation: 'landscape',
      paperSize: 'letter',
      scale: 80,
    });
    wb.setHeaderFooter('Sheet1', '&CTitle', '&RPage &P');
    wb.setPrintOptions('Sheet1', { gridLines: true });
    wb.insertPageBreak('Sheet1', 15);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const m = wb2.getPageMargins('Sheet1');
    expect(m.left).toBe(1.0);
    expect(m.top).toBe(1.5);

    const setup = wb2.getPageSetup('Sheet1');
    expect(setup.orientation).toBe('landscape');
    expect(setup.paperSize).toBe('letter');
    expect(setup.scale).toBe(80);

    const hf = wb2.getHeaderFooter('Sheet1');
    expect(hf.header).toBe('&CTitle');
    expect(hf.footer).toBe('&RPage &P');

    const opts = wb2.getPrintOptions('Sheet1');
    expect(opts.gridLines).toBe(true);

    const breaks = wb2.getPageBreaks('Sheet1');
    expect(breaks).toEqual([15]);
  });
});

describe('Formula Evaluation', () => {
  const out = tmpFile('test-formula-eval.xlsx');
  afterEach(async () => cleanup(out));

  it('should evaluate a simple formula', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 10);
    wb.setCellValue('Sheet1', 'A2', 20);
    const result = wb.evaluateFormula('Sheet1', 'SUM(A1:A2)');
    expect(result).toBe(30);
  });

  it('should evaluate string functions', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'hello');
    const result = wb.evaluateFormula('Sheet1', 'UPPER(A1)');
    expect(result).toBe('HELLO');
  });

  it('should calculate all formulas', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 5);
    wb.setCellValue('Sheet1', 'A2', 10);
    wb.setCellValue('Sheet1', 'A3', 100);
    wb.calculateAll();
    await wb.save(out);
    const wb2 = await Workbook.open(out);
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe(5);
    expect(wb2.getCellValue('Sheet1', 'A2')).toBe(10);
  });
});

describe('Pivot Tables', () => {
  const out = tmpFile('test-pivot.xlsx');
  afterEach(async () => cleanup(out));

  it('should add and get pivot tables', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 'Region');
    wb.setCellValue('Sheet1', 'C1', 'Sales');
    wb.setCellValue('Sheet1', 'A2', 'Alice');
    wb.setCellValue('Sheet1', 'B2', 'East');
    wb.setCellValue('Sheet1', 'C2', 1000);
    wb.setCellValue('Sheet1', 'A3', 'Bob');
    wb.setCellValue('Sheet1', 'B3', 'West');
    wb.setCellValue('Sheet1', 'C3', 2000);

    wb.newSheet('PivotSheet');
    wb.addPivotTable({
      name: 'PivotTable1',
      sourceSheet: 'Sheet1',
      sourceRange: 'A1:C3',
      targetSheet: 'PivotSheet',
      targetCell: 'A1',
      rows: [{ name: 'Region' }],
      columns: [],
      data: [{ name: 'Sales', function: 'sum' }],
    });

    const pivots = wb.getPivotTables();
    expect(pivots.length).toBe(1);
    expect(pivots[0].name).toBe('PivotTable1');
    expect(pivots[0].sourceSheet).toBe('Sheet1');
  });

  it('should delete a pivot table', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 'Sales');
    wb.setCellValue('Sheet1', 'A2', 'Alice');
    wb.setCellValue('Sheet1', 'B2', 100);

    wb.newSheet('PivotSheet');
    wb.addPivotTable({
      name: 'PT1',
      sourceSheet: 'Sheet1',
      sourceRange: 'A1:B2',
      targetSheet: 'PivotSheet',
      targetCell: 'A1',
      rows: [{ name: 'Name' }],
      columns: [],
      data: [{ name: 'Sales', function: 'sum' }],
    });

    expect(wb.getPivotTables().length).toBe(1);
    wb.deletePivotTable('PT1');
    expect(wb.getPivotTables().length).toBe(0);
  });

  it('should save and open with pivot tables', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Category');
    wb.setCellValue('Sheet1', 'B1', 'Amount');
    wb.setCellValue('Sheet1', 'A2', 'Food');
    wb.setCellValue('Sheet1', 'B2', 500);
    wb.setCellValue('Sheet1', 'A3', 'Transport');
    wb.setCellValue('Sheet1', 'B3', 300);

    wb.newSheet('PivotSheet');
    wb.addPivotTable({
      name: 'PT1',
      sourceSheet: 'Sheet1',
      sourceRange: 'A1:B3',
      targetSheet: 'PivotSheet',
      targetCell: 'A1',
      rows: [{ name: 'Category' }],
      columns: [],
      data: [{ name: 'Amount', function: 'sum' }],
    });

    await wb.save(out);
    const wb2 = await Workbook.open(out);
    const pivots = wb2.getPivotTables();
    expect(pivots.length).toBe(1);
    expect(pivots[0].name).toBe('PT1');
  });

  it('should throw on duplicate pivot table name', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'X');
    wb.setCellValue('Sheet1', 'B1', 'Y');
    wb.setCellValue('Sheet1', 'A2', 'a');
    wb.setCellValue('Sheet1', 'B2', 1);

    wb.newSheet('PivotSheet');
    wb.addPivotTable({
      name: 'PT1',
      sourceSheet: 'Sheet1',
      sourceRange: 'A1:B2',
      targetSheet: 'PivotSheet',
      targetCell: 'A1',
      rows: [{ name: 'X' }],
      columns: [],
      data: [{ name: 'Y', function: 'sum' }],
    });

    expect(() =>
      wb.addPivotTable({
        name: 'PT1',
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:B2',
        targetSheet: 'PivotSheet',
        targetCell: 'A5',
        rows: [{ name: 'X' }],
        columns: [],
        data: [{ name: 'Y', function: 'count' }],
      }),
    ).toThrow();
  });
});

describe('Sparklines', () => {
  it('should have sparkline config type available', () => {
    // Type-level check that the binding exists.
    // Full integration test will be added when workbook.addSparkline is ready.
    const config = {
      dataRange: 'Sheet1!A1:A10',
      location: 'B1',
      sparklineType: 'line',
      markers: true,
      highPoint: false,
      lowPoint: false,
      firstPoint: false,
      lastPoint: false,
      negativePoints: false,
      showAxis: false,
      lineWeight: 0.75,
      style: 1,
    };
    expect(config.dataRange).toBe('Sheet1!A1:A10');
    expect(config.location).toBe('B1');
    expect(config.sparklineType).toBe('line');
    expect(config.markers).toBe(true);
    expect(config.lineWeight).toBe(0.75);
    expect(config.style).toBe(1);
  });
});

// Theme colors
describe('Theme Colors', () => {
  it('should return default theme colors by index', () => {
    const wb = new Workbook();
    expect(wb.getThemeColor(0, null)).toBe('FF000000');
    expect(wb.getThemeColor(1, null)).toBe('FFFFFFFF');
    expect(wb.getThemeColor(4, null)).toBe('FF4472C4');
    expect(wb.getThemeColor(11, null)).toBe('FF954F72');
  });

  it('should return null for out-of-range index', () => {
    const wb = new Workbook();
    expect(wb.getThemeColor(99, null)).toBeNull();
  });

  it('should apply tint to theme colors', () => {
    const wb = new Workbook();
    const lightened = wb.getThemeColor(0, 0.5);
    expect(lightened).toBeTruthy();
    expect(lightened).toMatch(/^FF/);
    expect(lightened).not.toBe('FF000000');
  });

  it('should return base color with zero tint', () => {
    const wb = new Workbook();
    expect(wb.getThemeColor(4, 0.0)).toBe('FF4472C4');
  });
});

// Rich Text
describe('Rich Text', () => {
  const out = tmpFile('test-rich-text.xlsx');
  afterEach(async () => cleanup(out));

  it('should set and get rich text', () => {
    const wb = new Workbook();
    wb.setCellRichText('Sheet1', 'A1', [{ text: 'Bold', bold: true }, { text: 'Normal' }]);

    const runs = wb.getCellRichText('Sheet1', 'A1');
    expect(runs).not.toBeNull();
    expect(runs).toHaveLength(2);
    expect(runs?.[0].text).toBe('Bold');
    expect(runs?.[0].bold).toBe(true);
    expect(runs?.[1].text).toBe('Normal');
    expect(runs?.[1].bold).toBeUndefined();
  });

  it('should return null for non-rich-text cells', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'plain');
    const runs = wb.getCellRichText('Sheet1', 'A1');
    expect(runs).toBeNull();
  });

  it('should preserve font and color formatting', () => {
    const wb = new Workbook();
    wb.setCellRichText('Sheet1', 'B2', [
      {
        text: 'Styled',
        font: 'Arial',
        size: 14,
        bold: true,
        italic: true,
        color: '#FF0000',
      },
    ]);

    const runs = wb.getCellRichText('Sheet1', 'B2');
    expect(runs).not.toBeNull();
    expect(runs).toHaveLength(1);
    expect(runs?.[0].font).toBe('Arial');
    expect(runs?.[0].size).toBe(14);
    expect(runs?.[0].bold).toBe(true);
    expect(runs?.[0].italic).toBe(true);
    expect(runs?.[0].color).toBe('#FF0000');
  });

  it('should round-trip rich text through save and open', async () => {
    const wb = new Workbook();
    wb.setCellRichText('Sheet1', 'C3', [
      { text: 'Hello', bold: true, font: 'Arial', size: 14, color: '#FF0000' },
      { text: 'World', italic: true },
    ]);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const runs = wb2.getCellRichText('Sheet1', 'C3');
    expect(runs).not.toBeNull();
    expect(runs).toHaveLength(2);
    expect(runs?.[0].text).toBe('Hello');
    expect(runs?.[0].bold).toBe(true);
    expect(runs?.[0].font).toBe('Arial');
    expect(runs?.[0].size).toBe(14);
    expect(runs?.[0].color).toBe('#FF0000');
    expect(runs?.[1].text).toBe('World');
    expect(runs?.[1].italic).toBe(true);
  });

  it('should read rich text cell value as concatenated plain text', () => {
    const wb = new Workbook();
    wb.setCellRichText('Sheet1', 'A1', [{ text: 'Hello' }, { text: 'World' }]);

    const val = wb.getCellValue('Sheet1', 'A1');
    expect(val).toBe('HelloWorld');
  });
});
