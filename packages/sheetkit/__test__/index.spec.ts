import { access, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { decodeRowsBuffer } from '../buffer-codec.js';
import { Workbook } from '../index.js';
import { SheetData } from '../sheet-data.js';

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
  const out = tmpFile('test-validation.xlsx');
  afterEach(async () => cleanup(out));

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

  it('should return all properties for list validation', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'B2:B100',
      validationType: 'list',
      formula1: '"Red,Green,Blue"',
      allowBlank: true,
      errorStyle: 'stop',
      errorTitle: 'Invalid color',
      errorMessage: 'Please select from the list.',
      promptTitle: 'Color',
      promptMessage: 'Choose a color.',
      showInputMessage: true,
      showErrorMessage: true,
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].validationType).toBe('list');
    expect(v[0].formula1).toBe('"Red,Green,Blue"');
    expect(v[0].allowBlank).toBe(true);
    expect(v[0].errorStyle).toBe('stop');
    expect(v[0].errorTitle).toBe('Invalid color');
    expect(v[0].errorMessage).toBe('Please select from the list.');
    expect(v[0].promptTitle).toBe('Color');
    expect(v[0].promptMessage).toBe('Choose a color.');
    expect(v[0].showInputMessage).toBe(true);
    expect(v[0].showErrorMessage).toBe(true);
  });

  it('should add whole number between validation', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'C1:C50',
      validationType: 'whole',
      operator: 'between',
      formula1: '1',
      formula2: '100',
      errorMessage: 'Enter a number between 1 and 100',
      showErrorMessage: true,
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].validationType).toBe('whole');
    expect(v[0].operator).toBe('between');
    expect(v[0].formula1).toBe('1');
    expect(v[0].formula2).toBe('100');
  });

  it('should add decimal greaterThan validation', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'D1:D50',
      validationType: 'decimal',
      operator: 'greaterThan',
      formula1: '0.5',
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].validationType).toBe('decimal');
    expect(v[0].operator).toBe('greaterThan');
    expect(v[0].formula1).toBe('0.5');
  });

  it('should add textLength validation', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'E1:E50',
      validationType: 'textLength',
      operator: 'lessThanOrEqual',
      formula1: '255',
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].validationType).toBe('textLength');
    expect(v[0].operator).toBe('lessThanOrEqual');
    expect(v[0].formula1).toBe('255');
  });

  it('should add custom formula validation', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'F1:F50',
      validationType: 'custom',
      formula1: 'AND(LEN(F1)>0,LEN(F1)<100)',
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].validationType).toBe('custom');
    expect(v[0].formula1).toBe('AND(LEN(F1)>0,LEN(F1)<100)');
  });

  it('should support warning and information error styles', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'whole',
      operator: 'greaterThan',
      formula1: '0',
      errorStyle: 'warning',
    });
    wb.addDataValidation('Sheet1', {
      sqref: 'B1:B10',
      validationType: 'whole',
      operator: 'greaterThan',
      formula1: '0',
      errorStyle: 'information',
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(2);
    const styles = v.map((x) => x.errorStyle).sort();
    expect(styles).toEqual(['information', 'warning']);
  });

  it('should handle multiple validations on one sheet', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'list',
      formula1: '"X,Y,Z"',
    });
    wb.addDataValidation('Sheet1', {
      sqref: 'B1:B10',
      validationType: 'whole',
      operator: 'between',
      formula1: '1',
      formula2: '999',
    });
    wb.addDataValidation('Sheet1', {
      sqref: 'C1:C10',
      validationType: 'decimal',
      operator: 'greaterThanOrEqual',
      formula1: '0',
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(3);
  });

  it('should remove one validation and keep others', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'list',
      formula1: '"A,B"',
    });
    wb.addDataValidation('Sheet1', {
      sqref: 'B1:B10',
      validationType: 'whole',
      operator: 'between',
      formula1: '1',
      formula2: '10',
    });
    wb.removeDataValidation('Sheet1', 'A1:A10');
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].sqref).toBe('B1:B10');
  });

  it('should persist validations through save/open cycle', async () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'list',
      formula1: '"Red,Green,Blue"',
      allowBlank: true,
      errorStyle: 'stop',
      errorTitle: 'Invalid',
      errorMessage: 'Pick from list',
      promptTitle: 'Color',
      promptMessage: 'Choose color',
      showInputMessage: true,
      showErrorMessage: true,
    });
    wb.addDataValidation('Sheet1', {
      sqref: 'B1:B10',
      validationType: 'whole',
      operator: 'between',
      formula1: '0',
      formula2: '100',
    });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const v = wb2.getDataValidations('Sheet1');
    expect(v.length).toBe(2);
    const listV = v.find((x: { sqref: string }) => x.sqref === 'A1:A10');
    expect(listV).toBeDefined();
    expect(listV?.validationType).toBe('list');
    expect(listV?.formula1).toBe('"Red,Green,Blue"');
    expect(listV?.errorTitle).toBe('Invalid');
    expect(listV?.promptTitle).toBe('Color');
    const wholeV = v.find((x: { sqref: string }) => x.sqref === 'B1:B10');
    expect(wholeV).toBeDefined();
    expect(wholeV?.validationType).toBe('whole');
    expect(wholeV?.operator).toBe('between');
  });

  it('should return empty array when no validations exist', () => {
    const wb = new Workbook();
    const v = wb.getDataValidations('Sheet1');
    expect(v).toEqual([]);
  });

  it('should handle removing nonexistent validation gracefully', () => {
    const wb = new Workbook();
    wb.removeDataValidation('Sheet1', 'Z1:Z99');
    expect(wb.getDataValidations('Sheet1').length).toBe(0);
  });

  it('should support none validation type for prompt-only rules', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'none',
      promptTitle: 'Hint',
      promptMessage: 'Enter any value',
      showInputMessage: true,
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].validationType).toBe('none');
    expect(v[0].promptTitle).toBe('Hint');
  });

  it('should default allowBlank to false per OOXML spec', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'list',
      formula1: '"X,Y"',
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v[0].allowBlank).toBe(false);
  });

  it('should reject invalid sqref', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addDataValidation('Sheet1', {
        sqref: '',
        validationType: 'list',
        formula1: '"A,B"',
      }),
    ).toThrow();
  });

  it('should reject list validation without formula1', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addDataValidation('Sheet1', {
        sqref: 'A1:A10',
        validationType: 'list',
      }),
    ).toThrow();
  });

  it('should reject between operator without formula2', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addDataValidation('Sheet1', {
        sqref: 'A1:A10',
        validationType: 'whole',
        operator: 'between',
        formula1: '1',
      }),
    ).toThrow();
  });

  it('should accept case-insensitive operator input', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'whole',
      operator: 'GREATERTHAN',
      formula1: '0',
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v[0].operator).toBe('greaterThan');
  });

  it('should return camelCase operator strings', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'whole',
      operator: 'notbetween',
      formula1: '1',
      formula2: '10',
    });
    wb.addDataValidation('Sheet1', {
      sqref: 'B1:B10',
      validationType: 'textLength',
      operator: 'lessthanorequal',
      formula1: '255',
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v[0].operator).toBe('notBetween');
    expect(v[1].operator).toBe('lessThanOrEqual');
    expect(v[1].validationType).toBe('textLength');
  });

  it('should add date validation with between operator', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A100',
      validationType: 'date',
      operator: 'between',
      formula1: '44927',
      formula2: '45291',
      errorStyle: 'stop',
      errorTitle: 'Invalid date',
      errorMessage: 'Enter a date in 2023',
      showErrorMessage: true,
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].validationType).toBe('date');
    expect(v[0].operator).toBe('between');
    expect(v[0].formula1).toBe('44927');
    expect(v[0].formula2).toBe('45291');
    expect(v[0].errorTitle).toBe('Invalid date');
  });

  it('should add time validation with greaterThan operator', () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'B1:B50',
      validationType: 'time',
      operator: 'greaterThan',
      formula1: '0.375',
      errorStyle: 'warning',
      errorMessage: 'Time must be after 9:00 AM',
      showErrorMessage: true,
    });
    const v = wb.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].validationType).toBe('time');
    expect(v[0].operator).toBe('greaterThan');
    expect(v[0].formula1).toBe('0.375');
    expect(v[0].errorStyle).toBe('warning');
  });

  it('should persist none type validation through save/open cycle', async () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'none',
      promptTitle: 'Instructions',
      promptMessage: 'Enter any value here',
      showInputMessage: true,
    });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const v = wb2.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].validationType).toBe('none');
    expect(v[0].promptTitle).toBe('Instructions');
    expect(v[0].promptMessage).toBe('Enter any value here');
    expect(v[0].showInputMessage).toBe(true);
  });

  it('should persist all boolean properties through save/open cycle', async () => {
    const wb = new Workbook();
    wb.addDataValidation('Sheet1', {
      sqref: 'A1:A10',
      validationType: 'list',
      formula1: '"Yes,No"',
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      errorStyle: 'stop',
      errorTitle: 'Error',
      errorMessage: 'Select Yes or No',
      promptTitle: 'Choice',
      promptMessage: 'Pick one',
    });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const v = wb2.getDataValidations('Sheet1');
    expect(v.length).toBe(1);
    expect(v[0].allowBlank).toBe(true);
    expect(v[0].showInputMessage).toBe(true);
    expect(v[0].showErrorMessage).toBe(true);
    expect(v[0].errorStyle).toBe('stop');
    expect(v[0].errorTitle).toBe('Error');
    expect(v[0].errorMessage).toBe('Select Yes or No');
    expect(v[0].promptTitle).toBe('Choice');
    expect(v[0].promptMessage).toBe('Pick one');
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

describe('Buffer I/O', () => {
  it('should writeBufferSync and openBufferSync roundtrip', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'hello');
    wb.setCellValue('Sheet1', 'B1', 42);
    wb.setCellValue('Sheet1', 'C1', true);

    const buf = wb.writeBufferSync();
    expect(buf).toBeInstanceOf(Buffer);
    expect(buf.length).toBeGreaterThan(0);

    const wb2 = Workbook.openBufferSync(buf);
    expect(wb2.sheetNames).toEqual(['Sheet1']);
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe('hello');
    expect(wb2.getCellValue('Sheet1', 'B1')).toBe(42);
    expect(wb2.getCellValue('Sheet1', 'C1')).toBe(true);
  });

  it('should writeBuffer and openBuffer async roundtrip', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'async test');

    const buf = await wb.writeBuffer();
    expect(buf).toBeInstanceOf(Buffer);

    const wb2 = await Workbook.openBuffer(buf);
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe('async test');
  });

  it('should throw on invalid buffer', () => {
    expect(() => Workbook.openBufferSync(Buffer.from('not a zip'))).toThrow();
  });
});

describe('setCellFormula / fillFormula', () => {
  it('should set a cell formula', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 10);
    wb.setCellValue('Sheet1', 'A2', 20);
    wb.setCellFormula('Sheet1', 'A3', 'SUM(A1:A2)');
    const result = wb.evaluateFormula('Sheet1', 'SUM(A1:A2)');
    expect(result).toBe(30);
  });

  it('should fill formula across a column range', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 10);
    wb.setCellValue('Sheet1', 'B1', 20);
    wb.setCellValue('Sheet1', 'A2', 30);
    wb.setCellValue('Sheet1', 'B2', 40);

    wb.fillFormula('Sheet1', 'C1:C2', 'A1+B1');

    // C1 should have A1+B1
    wb.calculateAll();
    const buf = wb.writeBufferSync();
    const wb2 = Workbook.openBufferSync(buf);
    // After calculateAll, the formula results should be stored
    // C1 = 10 + 20 = 30, C2 = 30 + 40 = 70
    expect(wb2.sheetNames).toContain('Sheet1');
  });

  it('should throw on invalid range', () => {
    const wb = new Workbook();
    expect(() => wb.fillFormula('Sheet1', 'INVALID', 'A1')).toThrow();
  });

  it('should throw on multi-column range', () => {
    const wb = new Workbook();
    expect(() => wb.fillFormula('Sheet1', 'A1:B5', 'C1')).toThrow();
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

describe('Batch APIs', () => {
  const out = tmpFile('test-batch.xlsx');
  afterEach(async () => cleanup(out));

  it('should set multiple cell values at once with setCellValues', () => {
    const wb = new Workbook();
    wb.setCellValues('Sheet1', [
      { cell: 'A1', value: 'Hello' },
      { cell: 'B1', value: 42 },
      { cell: 'C1', value: true },
      { cell: 'A2', value: 'World' },
    ]);

    expect(wb.getCellValue('Sheet1', 'A1')).toBe('Hello');
    expect(wb.getCellValue('Sheet1', 'B1')).toBe(42);
    expect(wb.getCellValue('Sheet1', 'C1')).toBe(true);
    expect(wb.getCellValue('Sheet1', 'A2')).toBe('World');
  });

  it('should clear cells with null in setCellValues', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'existing');
    wb.setCellValues('Sheet1', [{ cell: 'A1', value: null }]);
    expect(wb.getCellValue('Sheet1', 'A1')).toBeNull();
  });

  it('should set row values with setRowValues', () => {
    const wb = new Workbook();
    wb.setRowValues('Sheet1', 1, 'A', ['Name', 'Age', 'Active']);

    expect(wb.getCellValue('Sheet1', 'A1')).toBe('Name');
    expect(wb.getCellValue('Sheet1', 'B1')).toBe('Age');
    expect(wb.getCellValue('Sheet1', 'C1')).toBe('Active');
  });

  it('should set row values with column offset', () => {
    const wb = new Workbook();
    wb.setRowValues('Sheet1', 2, 'C', [100, 200, 300]);

    expect(wb.getCellValue('Sheet1', 'C2')).toBe(100);
    expect(wb.getCellValue('Sheet1', 'D2')).toBe(200);
    expect(wb.getCellValue('Sheet1', 'E2')).toBe(300);
    // A2 and B2 should be empty
    expect(wb.getCellValue('Sheet1', 'A2')).toBeNull();
  });

  it('should set entire sheet data with setSheetData', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Name', 'Score', 'Pass'],
      ['Alice', 95, true],
      ['Bob', 80, true],
      ['Charlie', 55, false],
    ]);

    expect(wb.getCellValue('Sheet1', 'A1')).toBe('Name');
    expect(wb.getCellValue('Sheet1', 'B1')).toBe('Score');
    expect(wb.getCellValue('Sheet1', 'C1')).toBe('Pass');
    expect(wb.getCellValue('Sheet1', 'A2')).toBe('Alice');
    expect(wb.getCellValue('Sheet1', 'B2')).toBe(95);
    expect(wb.getCellValue('Sheet1', 'C2')).toBe(true);
    expect(wb.getCellValue('Sheet1', 'A4')).toBe('Charlie');
    expect(wb.getCellValue('Sheet1', 'B4')).toBe(55);
    expect(wb.getCellValue('Sheet1', 'C4')).toBe(false);
  });

  it('should set sheet data with start cell offset', () => {
    const wb = new Workbook();
    wb.setSheetData(
      'Sheet1',
      [
        [1, 2],
        [3, 4],
      ],
      'C3',
    );

    expect(wb.getCellValue('Sheet1', 'C3')).toBe(1);
    expect(wb.getCellValue('Sheet1', 'D3')).toBe(2);
    expect(wb.getCellValue('Sheet1', 'C4')).toBe(3);
    expect(wb.getCellValue('Sheet1', 'D4')).toBe(4);
    // A1 should be empty
    expect(wb.getCellValue('Sheet1', 'A1')).toBeNull();
  });

  it('should roundtrip batch data through save/open', async () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Header1', 'Header2'],
      [100, true],
      ['text', 42.5],
    ]);
    await wb.save(out);

    const wb2 = Workbook.openSync(out);
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe('Header1');
    expect(wb2.getCellValue('Sheet1', 'B1')).toBe('Header2');
    expect(wb2.getCellValue('Sheet1', 'A2')).toBe(100);
    expect(wb2.getCellValue('Sheet1', 'B2')).toBe(true);
    expect(wb2.getCellValue('Sheet1', 'A3')).toBe('text');
    expect(wb2.getCellValue('Sheet1', 'B3')).toBe(42.5);
  });

  it('should handle large batch efficiently', () => {
    const wb = new Workbook();
    const rows: (string | number)[][] = [];
    for (let r = 0; r < 1000; r++) {
      const row: (string | number)[] = [];
      for (let c = 0; c < 20; c++) {
        row.push(r * 20 + c);
      }
      rows.push(row);
    }
    wb.setSheetData('Sheet1', rows);

    // Spot-check
    expect(wb.getCellValue('Sheet1', 'A1')).toBe(0);
    expect(wb.getCellValue('Sheet1', 'T1')).toBe(19);
    expect(wb.getCellValue('Sheet1', 'A1000')).toBe(999 * 20);
    expect(wb.getCellValue('Sheet1', 'T1000')).toBe(999 * 20 + 19);
  });

  it('should handle mixed types in setSheetData', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [['text', 42, true, null]]);

    expect(wb.getCellValue('Sheet1', 'A1')).toBe('text');
    expect(wb.getCellValue('Sheet1', 'B1')).toBe(42);
    expect(wb.getCellValue('Sheet1', 'C1')).toBe(true);
    expect(wb.getCellValue('Sheet1', 'D1')).toBeNull();
  });

  it('should handle ragged rows in setSheetData', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [['a', 'b', 'c'], ['d'], ['e', 'f']]);

    expect(wb.getCellValue('Sheet1', 'A1')).toBe('a');
    expect(wb.getCellValue('Sheet1', 'C1')).toBe('c');
    expect(wb.getCellValue('Sheet1', 'A2')).toBe('d');
    expect(wb.getCellValue('Sheet1', 'B2')).toBeNull();
    expect(wb.getCellValue('Sheet1', 'A3')).toBe('e');
    expect(wb.getCellValue('Sheet1', 'B3')).toBe('f');
  });
});

// getRowsBuffer API
describe('getRowsBuffer', () => {
  it('should return a Buffer with valid SKRD magic', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'test');
    const buf = wb.getRowsBuffer('Sheet1');
    expect(buf).toBeInstanceOf(Buffer);
    expect(buf.readUInt32LE(0)).toBe(0x534b5244);
  });

  it('should return a small header-only buffer for an empty sheet', () => {
    const wb = new Workbook();
    const buf = wb.getRowsBuffer('Sheet1');
    expect(buf).toBeInstanceOf(Buffer);
    expect(buf.length).toBe(16);
    const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    expect(view.getUint32(6, true)).toBe(0);
    expect(view.getUint16(10, true)).toBe(0);
  });

  it('should contain correct row and col counts', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'r1c1');
    wb.setCellValue('Sheet1', 'C1', 'r1c3');
    wb.setCellValue('Sheet1', 'A2', 'r2c1');
    const buf = wb.getRowsBuffer('Sheet1');
    const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    expect(view.getUint32(6, true)).toBe(2);
    expect(view.getUint16(10, true)).toBe(3);
  });

  it('should throw on non-existent sheet', () => {
    const wb = new Workbook();
    expect(() => wb.getRowsBuffer('NoSheet')).toThrow();
  });
});

// setSheetDataBuffer API
describe('setSheetDataBuffer', () => {
  it('should round-trip data through getRowsBuffer and setSheetDataBuffer', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 100);
    wb.setCellValue('Sheet1', 'C1', true);
    wb.setCellValue('Sheet1', 'A2', 'Alice');
    wb.setCellValue('Sheet1', 'B2', 42.5);
    const buf = wb.getRowsBuffer('Sheet1');

    const wb2 = new Workbook();
    wb2.newSheet('Target');
    wb2.setSheetDataBuffer('Target', buf);
    expect(wb2.getCellValue('Target', 'A1')).toBe('Name');
    expect(wb2.getCellValue('Target', 'B1')).toBe(100);
    expect(wb2.getCellValue('Target', 'C1')).toBe(true);
    expect(wb2.getCellValue('Target', 'A2')).toBe('Alice');
    expect(wb2.getCellValue('Target', 'B2')).toBe(42.5);
  });

  it('should use default startCell A1', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'hello');
    const buf = wb.getRowsBuffer('Sheet1');

    const wb2 = new Workbook();
    wb2.setSheetDataBuffer('Sheet1', buf);
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe('hello');
  });

  it('should apply data at a custom startCell offset', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'val1');
    wb.setCellValue('Sheet1', 'B1', 'val2');
    wb.setCellValue('Sheet1', 'A2', 'val3');
    const buf = wb.getRowsBuffer('Sheet1');

    const wb2 = new Workbook();
    wb2.setSheetDataBuffer('Sheet1', buf, 'C3');
    expect(wb2.getCellValue('Sheet1', 'C3')).toBe('val1');
    expect(wb2.getCellValue('Sheet1', 'D3')).toBe('val2');
    expect(wb2.getCellValue('Sheet1', 'C4')).toBe('val3');
    expect(wb2.getCellValue('Sheet1', 'A1')).toBeNull();
  });

  it('should handle empty buffer without error', () => {
    const wb = new Workbook();
    const emptyBuf = wb.getRowsBuffer('Sheet1');
    const wb2 = new Workbook();
    wb2.setSheetDataBuffer('Sheet1', emptyBuf);
    expect(wb2.getCellValue('Sheet1', 'A1')).toBeNull();
  });
});

// Buffer round-trip with all cell types
describe('Buffer round-trip cell types', () => {
  it('should transfer string cells', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'hello world');
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCell(1, 1)).toBe('hello world');
    expect(sd.getCellType(1, 1)).toBe('string');
  });

  it('should transfer integer and float number cells', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 42);
    wb.setCellValue('Sheet1', 'B1', 3.14);
    wb.setCellValue('Sheet1', 'C1', -100.5);
    wb.setCellValue('Sheet1', 'D1', 0);
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCell(1, 1)).toBe(42);
    expect(sd.getCell(1, 2)).toBe(3.14);
    expect(sd.getCell(1, 3)).toBe(-100.5);
    expect(sd.getCell(1, 4)).toBe(0);
  });

  it('should transfer boolean cells', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', true);
    wb.setCellValue('Sheet1', 'B1', false);
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCell(1, 1)).toBe(true);
    expect(sd.getCell(1, 2)).toBe(false);
    expect(sd.getCellType(1, 1)).toBe('boolean');
    expect(sd.getCellType(1, 2)).toBe('boolean');
  });

  it('should transfer date cells as numbers', () => {
    const wb = new Workbook();
    const styleId = wb.addStyle({ numFmtId: 14 });
    wb.setCellValue('Sheet1', 'A1', { type: 'date', serial: 45292 });
    wb.setCellStyle('Sheet1', 'A1', styleId);
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCell(1, 1)).toBe(45292);
  });

  it('should transfer formula cells with formula text', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 10);
    wb.setCellValue('Sheet1', 'A2', 20);
    wb.setCellFormula('Sheet1', 'A3', 'SUM(A1:A2)');
    wb.calculateAll();
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCellType(3, 1)).toBe('formula');
    expect(sd.getCell(3, 1)).toBe('SUM(A1:A2)');
  });

  it('should transfer rich text cells as concatenated strings', () => {
    const wb = new Workbook();
    wb.setCellRichText('Sheet1', 'A1', [{ text: 'Bold', bold: true }, { text: 'Normal' }]);
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCellType(1, 1)).toBe('string');
    expect(sd.getCell(1, 1)).toBe('BoldNormal');
  });
});

// SheetData class
describe('SheetData', () => {
  it('should construct from a valid getRowsBuffer result', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'test');
    wb.setCellValue('Sheet1', 'B1', 42);
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.rowCount).toBe(1);
    expect(sd.colCount).toBe(2);
  });

  it('should handle null buffer without crashing', () => {
    const sd = new SheetData(null);
    expect(sd.rowCount).toBe(0);
    expect(sd.colCount).toBe(0);
  });

  it('should handle empty buffer without crashing', () => {
    const sd = new SheetData(Buffer.alloc(0));
    expect(sd.rowCount).toBe(0);
    expect(sd.colCount).toBe(0);
  });

  it('should return correct rowCount and colCount', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'a');
    wb.setCellValue('Sheet1', 'C2', 'c');
    wb.setCellValue('Sheet1', 'D3', 'd');
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.rowCount).toBe(3);
    expect(sd.colCount).toBe(4);
  });

  it('should return correct values from getCell for all types', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'text');
    wb.setCellValue('Sheet1', 'B1', 99);
    wb.setCellValue('Sheet1', 'C1', true);
    wb.setCellValue('Sheet1', 'D1', false);
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCell(1, 1)).toBe('text');
    expect(sd.getCell(1, 2)).toBe(99);
    expect(sd.getCell(1, 3)).toBe(true);
    expect(sd.getCell(1, 4)).toBe(false);
  });

  it('should return null from getCell for non-existent row/col', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'x');
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCell(99, 1)).toBeNull();
    expect(sd.getCell(1, 99)).toBeNull();
  });

  it('should return correct type strings from getCellType', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'text');
    wb.setCellValue('Sheet1', 'B1', 42);
    wb.setCellValue('Sheet1', 'C1', true);
    wb.setCellValue('Sheet1', 'D1', false);
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCellType(1, 1)).toBe('string');
    expect(sd.getCellType(1, 2)).toBe('number');
    expect(sd.getCellType(1, 3)).toBe('boolean');
    expect(sd.getCellType(1, 4)).toBe('boolean');
    expect(sd.getCellType(99, 1)).toBe('empty');
  });

  it('should return correct array from getRow', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 100);
    wb.setCellValue('Sheet1', 'C1', true);
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    const row = sd.getRow(1);
    expect(row).toEqual(['Name', 100, true]);
  });

  it('should return empty array from getRow for non-existent row', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'x');
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getRow(99)).toEqual([]);
  });

  it('should return full 2D array from toArray', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'r1');
    wb.setCellValue('Sheet1', 'B1', 10);
    wb.setCellValue('Sheet1', 'A2', 'r2');
    wb.setCellValue('Sheet1', 'B2', 20);
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    const arr = sd.toArray();
    expect(arr).toEqual([
      ['r1', 10],
      ['r2', 20],
    ]);
  });

  it('should yield row objects from rows() generator', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'first');
    wb.setCellValue('Sheet1', 'A2', 'second');
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    const collected = [];
    for (const r of sd.rows()) {
      collected.push(r);
    }
    expect(collected.length).toBe(2);
    expect(collected[0].row).toBe(1);
    expect(collected[0].values).toEqual(['first']);
    expect(collected[1].row).toBe(2);
    expect(collected[1].values).toEqual(['second']);
  });

  it('should convert column indices to names with columnName', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'x');
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.columnName(0)).toBe('A');
    expect(sd.columnName(25)).toBe('Z');
    expect(sd.columnName(26)).toBe('AA');
  });
});

// decodeRowsBuffer function
describe('decodeRowsBuffer', () => {
  it('should decode buffer to JsRowData array matching getRows output', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 100);
    wb.setCellValue('Sheet1', 'A2', 'Alice');
    wb.setCellValue('Sheet1', 'B2', true);
    const buf = wb.getRowsBuffer('Sheet1');
    const decoded = decodeRowsBuffer(buf);
    const rows = wb.getRows('Sheet1');
    expect(decoded).toEqual(rows);
  });

  it('should include correct row number and cell properties', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'hello');
    wb.setCellValue('Sheet1', 'B1', 42);
    wb.setCellValue('Sheet1', 'C1', true);
    const buf = wb.getRowsBuffer('Sheet1');
    const decoded = decodeRowsBuffer(buf);
    expect(decoded.length).toBe(1);
    expect(decoded[0].row).toBe(1);
    expect(decoded[0].cells.length).toBe(3);
    expect(decoded[0].cells[0]).toEqual({ column: 'A', valueType: 'string', value: 'hello' });
    expect(decoded[0].cells[1]).toEqual({ column: 'B', valueType: 'number', numberValue: 42 });
    expect(decoded[0].cells[2]).toEqual({ column: 'C', valueType: 'boolean', boolValue: true });
  });

  it('should return empty array for null or empty buffer', () => {
    expect(decodeRowsBuffer(null)).toEqual([]);
    expect(decodeRowsBuffer(Buffer.alloc(0))).toEqual([]);
    expect(decodeRowsBuffer(Buffer.alloc(5))).toEqual([]);
  });

  it('should return correct valueType strings', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'text');
    wb.setCellValue('Sheet1', 'B1', 42);
    wb.setCellValue('Sheet1', 'C1', true);
    wb.setCellFormula('Sheet1', 'D1', 'A1');
    wb.calculateAll();
    const buf = wb.getRowsBuffer('Sheet1');
    const decoded = decodeRowsBuffer(buf);
    expect(decoded[0].cells[0].valueType).toBe('string');
    expect(decoded[0].cells[1].valueType).toBe('number');
    expect(decoded[0].cells[2].valueType).toBe('boolean');
    expect(decoded[0].cells[3].valueType).toBe('formula');
  });
});

// Sparse data handling
describe('Buffer sparse data', () => {
  it('should handle cells at A1 and Z50', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'first');
    wb.setCellValue('Sheet1', 'Z50', 'last');
    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCell(1, 1)).toBe('first');
    expect(sd.getCell(50, 26)).toBe('last');
    expect(sd.getCell(1, 2)).toBeNull();
    expect(sd.getCell(25, 1)).toBeNull();
  });

  it('should decode sparse cells correctly with getRows', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'first');
    wb.setCellValue('Sheet1', 'Z50', 'last');
    const rows = wb.getRows('Sheet1');
    expect(rows.length).toBe(2);
    expect(rows[0].row).toBe(1);
    expect(rows[0].cells[0].column).toBe('A');
    expect(rows[0].cells[0].value).toBe('first');
    expect(rows[1].row).toBe(50);
    expect(rows[1].cells[0].column).toBe('Z');
    expect(rows[1].cells[0].value).toBe('last');
  });

  it('should decode sparse cells correctly with decodeRowsBuffer', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'first');
    wb.setCellValue('Sheet1', 'Z50', 'last');
    const buf = wb.getRowsBuffer('Sheet1');
    const decoded = decodeRowsBuffer(buf);
    expect(decoded.length).toBe(2);
    expect(decoded[0].cells[0].column).toBe('A');
    expect(decoded[0].cells[0].value).toBe('first');
    expect(decoded[1].cells[0].column).toBe('Z');
    expect(decoded[1].cells[0].value).toBe('last');
  });
});

// Save/open round-trip with buffer
describe('Buffer save/open round-trip', () => {
  const out = tmpFile('test-buffer-roundtrip.xlsx');
  afterEach(async () => cleanup(out));

  it('should preserve values through save, open, and buffer read', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 100);
    wb.setCellValue('Sheet1', 'C1', true);
    wb.setCellValue('Sheet1', 'A2', 'Bob');
    wb.setCellValue('Sheet1', 'B2', 55.5);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const buf = wb2.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.getCell(1, 1)).toBe('Name');
    expect(sd.getCell(1, 2)).toBe(100);
    expect(sd.getCell(1, 3)).toBe(true);
    expect(sd.getCell(2, 1)).toBe('Bob');
    expect(sd.getCell(2, 2)).toBe(55.5);
  });

  it('should preserve values through setSheetData, save, open, and buffer verify', async () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Header', 'Value'],
      ['Row1', 10],
      ['Row2', 20],
    ]);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const buf = wb2.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.toArray()).toEqual([
      ['Header', 'Value'],
      ['Row1', 10],
      ['Row2', 20],
    ]);
  });
});

// Large dataset
describe('Buffer large dataset', () => {
  it('should handle 1000 rows x 20 cols via SheetData', () => {
    const wb = new Workbook();
    const data: number[][] = [];
    for (let r = 0; r < 1000; r++) {
      const row: number[] = [];
      for (let c = 0; c < 20; c++) {
        row.push(r * 20 + c);
      }
      data.push(row);
    }
    wb.setSheetData('Sheet1', data);

    const buf = wb.getRowsBuffer('Sheet1');
    const sd = new SheetData(buf);
    expect(sd.rowCount).toBe(1000);
    expect(sd.colCount).toBe(20);
    expect(sd.getCell(1, 1)).toBe(0);
    expect(sd.getCell(1, 20)).toBe(19);
    expect(sd.getCell(500, 1)).toBe(499 * 20);
    expect(sd.getCell(500, 10)).toBe(499 * 20 + 9);
    expect(sd.getCell(1000, 1)).toBe(999 * 20);
    expect(sd.getCell(1000, 20)).toBe(999 * 20 + 19);
  });

  it('should produce consistent results between getRows and decodeRowsBuffer for large data', () => {
    const wb = new Workbook();
    const data: (string | number)[][] = [];
    for (let r = 0; r < 100; r++) {
      const row: (string | number)[] = [];
      for (let c = 0; c < 10; c++) {
        row.push(r % 2 === 0 ? r * 10 + c : `val_${r}_${c}`);
      }
      data.push(row);
    }
    wb.setSheetData('Sheet1', data);

    const rows = wb.getRows('Sheet1');
    const buf = wb.getRowsBuffer('Sheet1');
    const decoded = decodeRowsBuffer(buf);
    expect(decoded).toEqual(rows);
  });
});

describe('toJSON', () => {
  it('should convert sheet data to JSON with first row as headers', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Name', 'Age', 'Active'],
      ['Alice', 30, true],
      ['Bob', 25, false],
    ]);
    const result = wb.toJSON('Sheet1');
    expect(result).toEqual([
      { Name: 'Alice', Age: 30, Active: true },
      { Name: 'Bob', Age: 25, Active: false },
    ]);
  });

  it('should convert with custom header keys', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Alice', 30],
      ['Bob', 25],
    ]);
    const result = wb.toJSON('Sheet1', { header: ['name', 'age'] });
    expect(result).toEqual([
      { name: 'Alice', age: 30 },
      { name: 'Bob', age: 25 },
    ]);
  });

  it('should use column letters when header is false', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Alice', 30],
      ['Bob', 25],
    ]);
    const result = wb.toJSON('Sheet1', { header: false });
    expect(result).toEqual([
      { A: 'Alice', B: 30 },
      { B: 25, A: 'Bob' },
    ]);
  });

  it('should return empty array for empty sheet', () => {
    const wb = new Workbook();
    const result = wb.toJSON('Sheet1');
    expect(result).toEqual([]);
  });

  it('should handle mixed types including null', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Key', 'Value'],
      ['number', 42],
      ['text', 'hello'],
      ['bool', true],
      ['empty', null],
    ]);
    const result = wb.toJSON('Sheet1');
    expect(result).toEqual([
      { Key: 'number', Value: 42 },
      { Key: 'text', Value: 'hello' },
      { Key: 'bool', Value: true },
      { Key: 'empty', Value: null },
    ]);
  });

  it('should handle rows shorter than header', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [['A', 'B', 'C'], ['only_one']]);
    const result = wb.toJSON('Sheet1');
    expect(result).toEqual([{ A: 'only_one', B: null, C: null }]);
  });

  it('should handle single-row sheet (header only)', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [['Name', 'Age']]);
    const result = wb.toJSON('Sheet1');
    expect(result).toEqual([]);
  });
});

describe('toCSV', () => {
  it('should convert sheet data to CSV', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Name', 'Age'],
      ['Alice', 30],
      ['Bob', 25],
    ]);
    const csv = wb.toCSV('Sheet1');
    expect(csv).toBe('Name,Age\nAlice,30\nBob,25');
  });

  it('should escape values containing delimiter', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Name', 'Notes'],
      ['Alice', 'likes cats, dogs'],
    ]);
    const csv = wb.toCSV('Sheet1');
    expect(csv).toBe('Name,Notes\nAlice,"likes cats, dogs"');
  });

  it('should escape values containing quotes', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Name', 'Quote'],
      ['Alice', 'She said "hello"'],
    ]);
    const csv = wb.toCSV('Sheet1');
    expect(csv).toBe('Name,Quote\nAlice,"She said ""hello"""');
  });

  it('should escape values containing newlines', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Name', 'Bio'],
      ['Alice', 'Line 1\nLine 2'],
    ]);
    const csv = wb.toCSV('Sheet1');
    expect(csv).toBe('Name,Bio\nAlice,"Line 1\nLine 2"');
  });

  it('should use custom delimiter and line ending', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Name', 'Age'],
      ['Alice', 30],
    ]);
    const csv = wb.toCSV('Sheet1', { delimiter: '\t', lineEnding: '\r\n' });
    expect(csv).toBe('Name\tAge\r\nAlice\t30');
  });

  it('should return empty string for empty sheet', () => {
    const wb = new Workbook();
    const csv = wb.toCSV('Sheet1');
    expect(csv).toBe('');
  });

  it('should handle boolean and null values', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [[true, false, null, 0]]);
    const csv = wb.toCSV('Sheet1');
    expect(csv).toBe('true,false,,0');
  });

  it('should escape custom quote character', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [["it's a test", 'normal']]);
    const csv = wb.toCSV('Sheet1', { quote: "'" });
    expect(csv).toBe("'it''s a test',normal");
  });
});

describe('toHTML', () => {
  it('should generate an HTML table', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [
      ['Name', 'Age'],
      ['Alice', 30],
    ]);
    const html = wb.toHTML('Sheet1');
    expect(html).toBe(
      '<table><tr><td>Name</td><td>Age</td></tr><tr><td>Alice</td><td>30</td></tr></table>',
    );
  });

  it('should apply className to table element', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [['A']]);
    const html = wb.toHTML('Sheet1', { className: 'my-table' });
    expect(html).toContain('<table class="my-table">');
  });

  it('should escape HTML special characters', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [['<script>alert("xss")</script>']]);
    const html = wb.toHTML('Sheet1');
    expect(html).not.toContain('<script>');
    expect(html).toContain('&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;');
  });

  it('should escape ampersands', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [['A & B']]);
    const html = wb.toHTML('Sheet1');
    expect(html).toContain('A &amp; B');
  });

  it('should escape single quotes in content', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [["it's"]]);
    const html = wb.toHTML('Sheet1');
    expect(html).toContain('it&#39;s');
  });

  it('should escape className to prevent attribute injection', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [['A']]);
    const html = wb.toHTML('Sheet1', { className: 'foo" onclick="alert(1)' });
    expect(html).toContain('class="foo&quot; onclick=&quot;alert(1)"');
    expect(html).not.toContain('onclick="alert(1)"');
  });

  it('should return empty table for empty sheet', () => {
    const wb = new Workbook();
    const html = wb.toHTML('Sheet1');
    expect(html).toBe('<table></table>');
  });

  it('should handle null cells', () => {
    const wb = new Workbook();
    wb.setSheetData('Sheet1', [['', 'value', '']]);
    const html = wb.toHTML('Sheet1');
    expect(html).toContain('<td></td><td>value</td><td></td>');
  });
});

describe('fromJSON', () => {
  it('should write JSON objects to a sheet with auto-detected headers', () => {
    const wb = new Workbook();
    wb.fromJSON('Sheet1', [
      { Name: 'Alice', Age: 30 },
      { Name: 'Bob', Age: 25 },
    ]);
    expect(wb.getCellValue('Sheet1', 'A1')).toBe('Name');
    expect(wb.getCellValue('Sheet1', 'B1')).toBe('Age');
    expect(wb.getCellValue('Sheet1', 'A2')).toBe('Alice');
    expect(wb.getCellValue('Sheet1', 'B2')).toBe(30);
    expect(wb.getCellValue('Sheet1', 'A3')).toBe('Bob');
    expect(wb.getCellValue('Sheet1', 'B3')).toBe(25);
  });

  it('should write with custom headers', () => {
    const wb = new Workbook();
    wb.fromJSON(
      'Sheet1',
      [
        { name: 'Alice', age: 30 },
        { name: 'Bob', age: 25 },
      ],
      { header: ['name', 'age'] },
    );
    expect(wb.getCellValue('Sheet1', 'A1')).toBe('name');
    expect(wb.getCellValue('Sheet1', 'B1')).toBe('age');
    expect(wb.getCellValue('Sheet1', 'A2')).toBe('Alice');
  });

  it('should skip header row when header is false', () => {
    const wb = new Workbook();
    wb.fromJSON(
      'Sheet1',
      [
        { Name: 'Alice', Age: 30 },
        { Name: 'Bob', Age: 25 },
      ],
      { header: false },
    );
    expect(wb.getCellValue('Sheet1', 'A1')).toBe('Alice');
    expect(wb.getCellValue('Sheet1', 'B1')).toBe(30);
    expect(wb.getCellValue('Sheet1', 'A2')).toBe('Bob');
  });

  it('should write to a custom start cell', () => {
    const wb = new Workbook();
    wb.fromJSON('Sheet1', [{ X: 1, Y: 2 }], { startCell: 'C3' });
    expect(wb.getCellValue('Sheet1', 'C3')).toBe('X');
    expect(wb.getCellValue('Sheet1', 'D3')).toBe('Y');
    expect(wb.getCellValue('Sheet1', 'C4')).toBe(1);
    expect(wb.getCellValue('Sheet1', 'D4')).toBe(2);
  });

  it('should handle empty data array', () => {
    const wb = new Workbook();
    wb.fromJSON('Sheet1', []);
    expect(wb.getCellValue('Sheet1', 'A1')).toBeNull();
  });

  it('should handle null values in objects', () => {
    const wb = new Workbook();
    wb.fromJSON('Sheet1', [{ A: 'value', B: null }]);
    expect(wb.getCellValue('Sheet1', 'A2')).toBe('value');
    expect(wb.getCellValue('Sheet1', 'B2')).toBeNull();
  });

  it('should handle missing keys in records (undefined becomes null)', () => {
    const wb = new Workbook();
    wb.fromJSON('Sheet1', [{ A: 1, B: 2, C: 3 }, { A: 4 }]);
    expect(wb.getCellValue('Sheet1', 'A3')).toBe(4);
    expect(wb.getCellValue('Sheet1', 'B3')).toBeNull();
    expect(wb.getCellValue('Sheet1', 'C3')).toBeNull();
  });

  it('should roundtrip with toJSON', () => {
    const wb = new Workbook();
    const original = [
      { Name: 'Alice', Score: 95, Pass: true },
      { Name: 'Bob', Score: 80, Pass: true },
      { Name: 'Charlie', Score: 55, Pass: false },
    ];
    wb.fromJSON('Sheet1', original);
    const result = wb.toJSON('Sheet1');
    expect(result).toEqual(original);
  });

  it('should roundtrip with toJSON using header: false on both sides', () => {
    const wb = new Workbook();
    wb.fromJSON(
      'Sheet1',
      [
        { A: 'x', B: 1 },
        { A: 'y', B: 2 },
      ],
      { header: false },
    );
    const result = wb.toJSON('Sheet1', { header: false });
    expect(result).toEqual([
      { A: 'x', B: 1 },
      { A: 'y', B: 2 },
    ]);
  });

  it('should collect keys from all records, not just the first', () => {
    const wb = new Workbook();
    wb.fromJSON('Sheet1', [{ A: 1 }, { B: 2 }, { A: 3, C: 4 }]);
    expect(wb.getCellValue('Sheet1', 'A1')).toBe('A');
    expect(wb.getCellValue('Sheet1', 'B1')).toBe('B');
    expect(wb.getCellValue('Sheet1', 'C1')).toBe('C');
    expect(wb.getCellValue('Sheet1', 'A2')).toBe(1);
    expect(wb.getCellValue('Sheet1', 'B2')).toBeNull();
    expect(wb.getCellValue('Sheet1', 'C2')).toBeNull();
    expect(wb.getCellValue('Sheet1', 'A3')).toBeNull();
    expect(wb.getCellValue('Sheet1', 'B3')).toBe(2);
    expect(wb.getCellValue('Sheet1', 'C3')).toBeNull();
    expect(wb.getCellValue('Sheet1', 'A4')).toBe(3);
    expect(wb.getCellValue('Sheet1', 'B4')).toBeNull();
    expect(wb.getCellValue('Sheet1', 'C4')).toBe(4);
  });
});

describe('Sheet View Options', () => {
  const out = tmpFile('test-sheet-view.xlsx');
  afterEach(async () => cleanup(out));

  it('should return defaults for a new sheet', () => {
    const wb = new Workbook();
    const opts = wb.getSheetViewOptions('Sheet1');
    expect(opts.showGridlines).toBe(true);
    expect(opts.showFormulas).toBe(false);
    expect(opts.showRowColHeaders).toBe(true);
    expect(opts.zoomScale).toBe(100);
    expect(opts.viewMode).toBe('normal');
    expect(opts.topLeftCell).toBeUndefined();
  });

  it('should set and get gridlines off', () => {
    const wb = new Workbook();
    wb.setSheetViewOptions('Sheet1', { showGridlines: false });
    const opts = wb.getSheetViewOptions('Sheet1');
    expect(opts.showGridlines).toBe(false);
  });

  it('should set and get zoom scale', () => {
    const wb = new Workbook();
    wb.setSheetViewOptions('Sheet1', { zoomScale: 150 });
    const opts = wb.getSheetViewOptions('Sheet1');
    expect(opts.zoomScale).toBe(150);
  });

  it('should reject invalid zoom scale', () => {
    const wb = new Workbook();
    expect(() => wb.setSheetViewOptions('Sheet1', { zoomScale: 5 })).toThrow();
    expect(() => wb.setSheetViewOptions('Sheet1', { zoomScale: 500 })).toThrow();
  });

  it('should set and get view mode', () => {
    const wb = new Workbook();
    wb.setSheetViewOptions('Sheet1', { viewMode: 'pageBreak' });
    const opts = wb.getSheetViewOptions('Sheet1');
    expect(opts.viewMode).toBe('pageBreakPreview');
  });

  it('should set and get page layout view mode', () => {
    const wb = new Workbook();
    wb.setSheetViewOptions('Sheet1', { viewMode: 'pageLayout' });
    const opts = wb.getSheetViewOptions('Sheet1');
    expect(opts.viewMode).toBe('pageLayout');
  });

  it('should set and get show formulas', () => {
    const wb = new Workbook();
    wb.setSheetViewOptions('Sheet1', { showFormulas: true });
    const opts = wb.getSheetViewOptions('Sheet1');
    expect(opts.showFormulas).toBe(true);
  });

  it('should set and get top left cell', () => {
    const wb = new Workbook();
    wb.setSheetViewOptions('Sheet1', { topLeftCell: 'C5' });
    const opts = wb.getSheetViewOptions('Sheet1');
    expect(opts.topLeftCell).toBe('C5');
  });

  it('should set multiple options at once', () => {
    const wb = new Workbook();
    wb.setSheetViewOptions('Sheet1', {
      showGridlines: false,
      zoomScale: 75,
      viewMode: 'pageLayout',
      topLeftCell: 'B10',
    });
    const opts = wb.getSheetViewOptions('Sheet1');
    expect(opts.showGridlines).toBe(false);
    expect(opts.zoomScale).toBe(75);
    expect(opts.viewMode).toBe('pageLayout');
    expect(opts.topLeftCell).toBe('B10');
  });

  it('should roundtrip view options through save/open', async () => {
    const wb = new Workbook();
    wb.setSheetViewOptions('Sheet1', {
      showGridlines: false,
      zoomScale: 200,
      viewMode: 'pageBreak',
      showFormulas: true,
      showRowColHeaders: false,
    });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const opts = wb2.getSheetViewOptions('Sheet1');
    expect(opts.showGridlines).toBe(false);
    expect(opts.zoomScale).toBe(200);
    expect(opts.viewMode).toBe('pageBreakPreview');
    expect(opts.showFormulas).toBe(true);
    expect(opts.showRowColHeaders).toBe(false);
  });

  it('should preserve panes when setting view options', () => {
    const wb = new Workbook();
    wb.setPanes('Sheet1', 'A3');
    wb.setSheetViewOptions('Sheet1', { zoomScale: 80 });
    expect(wb.getPanes('Sheet1')).toBe('A3');
    const opts = wb.getSheetViewOptions('Sheet1');
    expect(opts.zoomScale).toBe(80);
  });

  it('should throw for non-existent sheet', () => {
    const wb = new Workbook();
    expect(() => wb.getSheetViewOptions('NoSheet')).toThrow();
    expect(() => wb.setSheetViewOptions('NoSheet', { zoomScale: 100 })).toThrow();
  });
});

describe('Sheet Visibility', () => {
  const out = tmpFile('test-sheet-visibility.xlsx');
  afterEach(async () => cleanup(out));

  it('should default to visible', () => {
    const wb = new Workbook();
    expect(wb.getSheetVisibility('Sheet1')).toBe('visible');
  });

  it('should set and get hidden', () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    wb.setSheetVisibility('Sheet1', 'hidden');
    expect(wb.getSheetVisibility('Sheet1')).toBe('hidden');
  });

  it('should set and get veryHidden', () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    wb.setSheetVisibility('Sheet1', 'veryHidden');
    expect(wb.getSheetVisibility('Sheet1')).toBe('veryHidden');
  });

  it('should not allow hiding the last visible sheet', () => {
    const wb = new Workbook();
    expect(() => wb.setSheetVisibility('Sheet1', 'hidden')).toThrow();
  });

  it('should roundtrip visibility through save/open', async () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    wb.newSheet('Sheet3');
    wb.setSheetVisibility('Sheet2', 'hidden');
    wb.setSheetVisibility('Sheet3', 'veryHidden');
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    expect(wb2.getSheetVisibility('Sheet1')).toBe('visible');
    expect(wb2.getSheetVisibility('Sheet2')).toBe('hidden');
    expect(wb2.getSheetVisibility('Sheet3')).toBe('veryHidden');
  });

  it('should throw for non-existent sheet', () => {
    const wb = new Workbook();
    expect(() => wb.getSheetVisibility('NoSheet')).toThrow();
    expect(() => wb.setSheetVisibility('NoSheet', 'hidden')).toThrow();
  });

  it('should throw for invalid visibility string', () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    expect(() => wb.setSheetVisibility('Sheet1', 'hiddden')).toThrow(/Invalid visibility/);
    expect(() => wb.setSheetVisibility('Sheet1', 'HIDDEN')).toThrow(/Invalid visibility/);
    expect(() => wb.setSheetVisibility('Sheet1', '')).toThrow(/Invalid visibility/);
  });
});

describe('Workbook Format', () => {
  const out = tmpFile('test-format.xlsx');
  afterEach(async () => cleanup(out));

  it('should return xlsx for a new workbook', () => {
    const wb = new Workbook();
    expect(wb.getFormat()).toBe('xlsx');
  });

  it('should detect xlsx format after save and open', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'test');
    await wb.save(out);
    const wb2 = await Workbook.open(out);
    expect(wb2.getFormat()).toBe('xlsx');
  });

  it('should detect xlsm format from buffer', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'data');
    const buf = wb.writeBufferSync();
    const wb2 = Workbook.openBufferSync(buf);
    expect(wb2.getFormat()).toBe('xlsx');
  });
});

describe('Table CRUD', () => {
  const out = tmpFile('test-table.xlsx');
  afterEach(async () => cleanup(out));

  it('should add a table and retrieve it', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 'Score');
    wb.setCellValue('Sheet1', 'A2', 'Alice');
    wb.setCellValue('Sheet1', 'B2', 95);

    wb.addTable('Sheet1', {
      name: 'Table1',
      displayName: 'Table1',
      range: 'A1:B2',
      columns: [{ name: 'Name' }, { name: 'Score' }],
    });

    const tables = wb.getTables('Sheet1');
    expect(tables).toHaveLength(1);
    expect(tables[0].name).toBe('Table1');
    expect(tables[0].displayName).toBe('Table1');
    expect(tables[0].range).toBe('A1:B2');
    expect(tables[0].columns).toEqual(['Name', 'Score']);
    expect(tables[0].showHeaderRow).toBe(true);
    expect(tables[0].autoFilter).toBe(true);
  });

  it('should add a table with style options', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Item');
    wb.setCellValue('Sheet1', 'B1', 'Price');

    wb.addTable('Sheet1', {
      name: 'StyledTable',
      displayName: 'StyledTable',
      range: 'A1:B5',
      columns: [{ name: 'Item' }, { name: 'Price' }],
      styleName: 'TableStyleMedium2',
      showRowStripes: true,
      showColumnStripes: false,
      showFirstColumn: true,
      showLastColumn: false,
    });

    const tables = wb.getTables('Sheet1');
    expect(tables).toHaveLength(1);
    expect(tables[0].styleName).toBe('TableStyleMedium2');
  });

  it('should add a table without header row', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'NoHeader',
      displayName: 'NoHeader',
      range: 'A1:B5',
      columns: [{ name: 'Col1' }, { name: 'Col2' }],
      showHeaderRow: false,
    });

    const tables = wb.getTables('Sheet1');
    expect(tables[0].showHeaderRow).toBe(false);
  });

  it('should delete a table', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'ToDelete',
      displayName: 'ToDelete',
      range: 'A1:B5',
      columns: [{ name: 'Col1' }, { name: 'Col2' }],
    });

    expect(wb.getTables('Sheet1')).toHaveLength(1);
    wb.deleteTable('Sheet1', 'ToDelete');
    expect(wb.getTables('Sheet1')).toHaveLength(0);
  });

  it('should throw when adding duplicate table name', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'Dup',
      displayName: 'Dup',
      range: 'A1:B5',
      columns: [{ name: 'Col1' }, { name: 'Col2' }],
    });

    expect(() =>
      wb.addTable('Sheet1', {
        name: 'Dup',
        displayName: 'Dup',
        range: 'C1:D5',
        columns: [{ name: 'Col3' }, { name: 'Col4' }],
      }),
    ).toThrow();
  });

  it('should throw when deleting non-existent table', () => {
    const wb = new Workbook();
    expect(() => wb.deleteTable('Sheet1', 'NoSuchTable')).toThrow();
  });

  it('should throw for non-existent sheet', () => {
    const wb = new Workbook();
    expect(() => wb.getTables('NoSheet')).toThrow();
  });

  it('should support multiple tables on different sheets', () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');

    wb.addTable('Sheet1', {
      name: 'T1',
      displayName: 'T1',
      range: 'A1:B5',
      columns: [{ name: 'X' }, { name: 'Y' }],
    });

    wb.addTable('Sheet2', {
      name: 'T2',
      displayName: 'T2',
      range: 'A1:C3',
      columns: [{ name: 'A' }, { name: 'B' }, { name: 'C' }],
    });

    expect(wb.getTables('Sheet1')).toHaveLength(1);
    expect(wb.getTables('Sheet2')).toHaveLength(1);
    expect(wb.getTables('Sheet1')[0].name).toBe('T1');
    expect(wb.getTables('Sheet2')[0].name).toBe('T2');
  });

  it('should roundtrip tables through save/open', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Product');
    wb.setCellValue('Sheet1', 'B1', 'Quantity');
    wb.setCellValue('Sheet1', 'A2', 'Widget');
    wb.setCellValue('Sheet1', 'B2', 100);

    wb.addTable('Sheet1', {
      name: 'Products',
      displayName: 'Products',
      range: 'A1:B2',
      columns: [{ name: 'Product' }, { name: 'Quantity' }],
      styleName: 'TableStyleLight1',
    });

    await wb.save(out);
    const wb2 = await Workbook.open(out);
    const tables = wb2.getTables('Sheet1');
    expect(tables).toHaveLength(1);
    expect(tables[0].name).toBe('Products');
    expect(tables[0].range).toBe('A1:B2');
    expect(tables[0].columns).toEqual(['Product', 'Quantity']);
  });
});

describe('Shapes', () => {
  const out = tmpFile('test-shape.xlsx');
  afterEach(async () => cleanup(out));

  it('should add a basic rectangle shape', async () => {
    const wb = new Workbook();
    wb.addShape('Sheet1', {
      shapeType: 'rect',
      fromCell: 'B2',
      toCell: 'F10',
    });
    await wb.save(out);
    await expect(access(out)).resolves.toBeUndefined();
  });

  it('should add a shape with text', async () => {
    const wb = new Workbook();
    wb.addShape('Sheet1', {
      shapeType: 'ellipse',
      fromCell: 'A1',
      toCell: 'D5',
      text: 'Hello World',
    });
    await wb.save(out);
    const wb2 = await Workbook.open(out);
    expect(wb2.sheetNames).toEqual(['Sheet1']);
  });

  it('should add a shape with fill and line styling', async () => {
    const wb = new Workbook();
    wb.addShape('Sheet1', {
      shapeType: 'roundRect',
      fromCell: 'B2',
      toCell: 'H12',
      text: 'Styled Shape',
      fillColor: '4472C4',
      lineColor: '2F528F',
      lineWidth: 2.0,
    });
    await wb.save(out);
    await expect(access(out)).resolves.toBeUndefined();
  });

  it('should add multiple shapes on one sheet', async () => {
    const wb = new Workbook();
    wb.addShape('Sheet1', {
      shapeType: 'rect',
      fromCell: 'A1',
      toCell: 'C3',
    });
    wb.addShape('Sheet1', {
      shapeType: 'diamond',
      fromCell: 'E1',
      toCell: 'H5',
      fillColor: 'FF0000',
    });
    wb.addShape('Sheet1', {
      shapeType: 'star5',
      fromCell: 'A6',
      toCell: 'D10',
      text: 'Star',
    });
    await wb.save(out);
    await expect(access(out)).resolves.toBeUndefined();
  });

  it('should add shapes on different sheets', async () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    wb.addShape('Sheet1', {
      shapeType: 'rect',
      fromCell: 'A1',
      toCell: 'C3',
      text: 'Sheet1 Shape',
    });
    wb.addShape('Sheet2', {
      shapeType: 'ellipse',
      fromCell: 'B2',
      toCell: 'E6',
      text: 'Sheet2 Shape',
      fillColor: '00FF00',
    });
    await wb.save(out);
    const wb2 = await Workbook.open(out);
    expect(wb2.sheetNames).toContain('Sheet1');
    expect(wb2.sheetNames).toContain('Sheet2');
  });

  it('should throw for unknown shape type', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addShape('Sheet1', {
        shapeType: 'nonexistent',
        fromCell: 'A1',
        toCell: 'C3',
      }),
    ).toThrow();
  });

  it('should throw for nonexistent sheet', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addShape('NoSheet', {
        shapeType: 'rect',
        fromCell: 'A1',
        toCell: 'C3',
      }),
    ).toThrow();
  });

  it('should coexist with charts on the same sheet', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Category');
    wb.setCellValue('Sheet1', 'B1', 'Value');
    wb.setCellValue('Sheet1', 'A2', 'A');
    wb.setCellValue('Sheet1', 'B2', 10);
    wb.addChart('Sheet1', 'E1', 'L10', {
      chartType: 'col',
      series: [{ name: 'S1', categories: 'Sheet1!$A$2:$A$2', values: 'Sheet1!$B$2:$B$2' }],
    });
    wb.addShape('Sheet1', {
      shapeType: 'rect',
      fromCell: 'A12',
      toCell: 'D16',
      text: 'Annotation',
    });
    await wb.save(out);
    await expect(access(out)).resolves.toBeUndefined();
  });
});

describe('Cross-feature integration', () => {
  const out = tmpFile('test-integration.xlsx');
  afterEach(async () => cleanup(out));

  it('should combine format detection with tables and views', async () => {
    const wb = new Workbook();
    expect(wb.getFormat()).toBe('xlsx');

    wb.setCellValue('Sheet1', 'A1', 'Region');
    wb.setCellValue('Sheet1', 'B1', 'Revenue');
    wb.setCellValue('Sheet1', 'A2', 'North');
    wb.setCellValue('Sheet1', 'B2', 50000);
    wb.setCellValue('Sheet1', 'A3', 'South');
    wb.setCellValue('Sheet1', 'B3', 75000);

    wb.addTable('Sheet1', {
      name: 'SalesData',
      displayName: 'SalesData',
      range: 'A1:B3',
      columns: [{ name: 'Region' }, { name: 'Revenue' }],
      styleName: 'TableStyleMedium2',
    });

    wb.setSheetViewOptions('Sheet1', {
      zoomScale: 120,
      showGridlines: false,
    });

    wb.setPanes('Sheet1', 'A2');

    await wb.save(out);
    const wb2 = await Workbook.open(out);

    expect(wb2.getFormat()).toBe('xlsx');

    const tables = wb2.getTables('Sheet1');
    expect(tables).toHaveLength(1);
    expect(tables[0].name).toBe('SalesData');

    const view = wb2.getSheetViewOptions('Sheet1');
    expect(view.zoomScale).toBe(120);
    expect(view.showGridlines).toBe(false);

    const panes = wb2.getPanes('Sheet1');
    expect(panes).toBe('A2');
  });

  it('should combine tables with cell styles and defined names', async () => {
    const wb = new Workbook();

    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 'Amount');
    wb.setCellValue('Sheet1', 'A2', 'Alice');
    wb.setCellValue('Sheet1', 'B2', 1000);
    wb.setCellValue('Sheet1', 'A3', 'Bob');
    wb.setCellValue('Sheet1', 'B3', 2000);

    wb.addTable('Sheet1', {
      name: 'Budget',
      displayName: 'Budget',
      range: 'A1:B3',
      columns: [{ name: 'Name' }, { name: 'Amount' }],
    });

    const boldStyle = wb.addStyle({
      font: { bold: true, size: 12 },
    });
    wb.setCellStyle('Sheet1', 'A1', boldStyle);
    wb.setCellStyle('Sheet1', 'B1', boldStyle);

    wb.setDefinedName({
      name: 'BudgetRange',
      value: 'Sheet1!$A$1:$B$3',
    });

    await wb.save(out);
    const wb2 = await Workbook.open(out);

    const tables = wb2.getTables('Sheet1');
    expect(tables).toHaveLength(1);
    expect(tables[0].name).toBe('Budget');

    const name = wb2.getDefinedName('BudgetRange');
    expect(name).not.toBeNull();
    expect(name?.value).toBe('Sheet1!$A$1:$B$3');

    expect(wb2.getCellValue('Sheet1', 'A2')).toBe('Alice');
    expect(wb2.getCellValue('Sheet1', 'B3')).toBe(2000);
  });

  it('should combine tables with sheet visibility and protection', () => {
    const wb = new Workbook();
    wb.newSheet('Data');
    wb.newSheet('Config');

    wb.setCellValue('Data', 'A1', 'Key');
    wb.setCellValue('Data', 'B1', 'Value');
    wb.setCellValue('Data', 'A2', 'alpha');
    wb.setCellValue('Data', 'B2', 42);

    wb.addTable('Data', {
      name: 'DataTable',
      displayName: 'DataTable',
      range: 'A1:B2',
      columns: [{ name: 'Key' }, { name: 'Value' }],
    });

    wb.setSheetVisibility('Config', 'hidden');
    wb.protectSheet('Data', { password: 'test123' });

    expect(wb.getSheetVisibility('Config')).toBe('hidden');
    expect(wb.isSheetProtected('Data')).toBe(true);
    expect(wb.getTables('Data')).toHaveLength(1);
  });

  it('should combine tables with data validation and comments', () => {
    const wb = new Workbook();

    wb.setCellValue('Sheet1', 'A1', 'Status');
    wb.setCellValue('Sheet1', 'B1', 'Notes');
    wb.setCellValue('Sheet1', 'A2', 'Active');
    wb.setCellValue('Sheet1', 'B2', 'First entry');

    wb.addTable('Sheet1', {
      name: 'StatusTable',
      displayName: 'StatusTable',
      range: 'A1:B2',
      columns: [{ name: 'Status' }, { name: 'Notes' }],
    });

    wb.addDataValidation('Sheet1', {
      sqref: 'A2:A100',
      validationType: 'list',
      formula1: '"Active,Inactive,Pending"',
    });

    wb.addComment('Sheet1', {
      cell: 'A1',
      author: 'System',
      text: 'This column uses a dropdown validation.',
    });

    const tables = wb.getTables('Sheet1');
    expect(tables).toHaveLength(1);

    const validations = wb.getDataValidations('Sheet1');
    expect(validations).toHaveLength(1);
    expect(validations[0].validationType).toBe('list');

    const comments = wb.getComments('Sheet1');
    expect(comments).toHaveLength(1);
    expect(comments[0].author).toBe('System');
  });

  it('should combine format detection with multi-sheet tables and page setup', async () => {
    const wb = new Workbook();
    wb.newSheet('Summary');

    wb.setCellValue('Sheet1', 'A1', 'Item');
    wb.setCellValue('Sheet1', 'B1', 'Count');
    wb.addTable('Sheet1', {
      name: 'Items',
      displayName: 'Items',
      range: 'A1:B5',
      columns: [{ name: 'Item' }, { name: 'Count' }],
    });

    wb.setCellValue('Summary', 'A1', 'Metric');
    wb.setCellValue('Summary', 'B1', 'Value');
    wb.addTable('Summary', {
      name: 'Metrics',
      displayName: 'Metrics',
      range: 'A1:B3',
      columns: [{ name: 'Metric' }, { name: 'Value' }],
    });

    wb.setPageSetup('Sheet1', { orientation: 'landscape', paperSize: 'a4' });
    wb.setPageSetup('Summary', { orientation: 'portrait', scale: 80 });

    wb.setDocProps({ title: 'Integration Test', creator: 'SheetKit' });

    await wb.save(out);
    const wb2 = await Workbook.open(out);

    expect(wb2.getFormat()).toBe('xlsx');
    expect(wb2.getTables('Sheet1')).toHaveLength(1);
    expect(wb2.getTables('Summary')).toHaveLength(1);
    expect(wb2.getTables('Sheet1')[0].name).toBe('Items');
    expect(wb2.getTables('Summary')[0].name).toBe('Metrics');

    const setup1 = wb2.getPageSetup('Sheet1');
    expect(setup1.orientation).toBe('landscape');

    const props = wb2.getDocProps();
    expect(props.title).toBe('Integration Test');
    expect(props.creator).toBe('SheetKit');
  });
});

describe('Open Options', () => {
  const out = tmpFile('test-open-options.xlsx');
  afterEach(async () => cleanup(out));

  it('should limit rows with sheetRows option', () => {
    const wb = new Workbook();
    for (let i = 1; i <= 20; i++) {
      wb.setCellValue('Sheet1', `A${i}`, i);
    }
    const buf = wb.writeBufferSync();

    const wb2 = Workbook.openBufferSync(buf, { sheetRows: 5 });
    for (let i = 1; i <= 5; i++) {
      expect(wb2.getCellValue('Sheet1', `A${i}`)).toBe(i);
    }
    for (let i = 6; i <= 20; i++) {
      expect(wb2.getCellValue('Sheet1', `A${i}`)).toBeNull();
    }
  });

  it('should limit rows with sheetRows option via file', async () => {
    const wb = new Workbook();
    for (let i = 1; i <= 15; i++) {
      wb.setCellValue('Sheet1', `A${i}`, i * 10);
    }
    await wb.save(out);

    const wb2 = Workbook.openSync(out, { sheetRows: 3 });
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe(10);
    expect(wb2.getCellValue('Sheet1', 'A3')).toBe(30);
    expect(wb2.getCellValue('Sheet1', 'A4')).toBeNull();
  });

  it('should parse only selected sheets', () => {
    const wb = new Workbook();
    wb.newSheet('Sales');
    wb.newSheet('Data');
    wb.setCellValue('Sheet1', 'A1', 'sheet1');
    wb.setCellValue('Sales', 'A1', 'sales');
    wb.setCellValue('Data', 'A1', 'data');
    const buf = wb.writeBufferSync();

    const wb2 = Workbook.openBufferSync(buf, { sheets: ['Sales'] });
    expect(wb2.sheetNames).toEqual(['Sheet1', 'Sales', 'Data']);
    expect(wb2.getCellValue('Sales', 'A1')).toBe('sales');
    expect(wb2.getCellValue('Sheet1', 'A1')).toBeNull();
    expect(wb2.getCellValue('Data', 'A1')).toBeNull();
  });

  it('should parse multiple selected sheets', () => {
    const wb = new Workbook();
    wb.newSheet('Alpha');
    wb.newSheet('Beta');
    wb.setCellValue('Sheet1', 'A1', 1);
    wb.setCellValue('Alpha', 'A1', 2);
    wb.setCellValue('Beta', 'A1', 3);
    const buf = wb.writeBufferSync();

    const wb2 = Workbook.openBufferSync(buf, { sheets: ['Sheet1', 'Beta'] });
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe(1);
    expect(wb2.getCellValue('Alpha', 'A1')).toBeNull();
    expect(wb2.getCellValue('Beta', 'A1')).toBe(3);
  });

  it('should reject when maxZipEntries is exceeded', () => {
    const wb = new Workbook();
    const buf = wb.writeBufferSync();
    expect(() => Workbook.openBufferSync(buf, { maxZipEntries: 2 })).toThrow(/entry count/i);
  });

  it('should accept when maxZipEntries is within limit', () => {
    const wb = new Workbook();
    const buf = wb.writeBufferSync();
    const wb2 = Workbook.openBufferSync(buf, { maxZipEntries: 1000 });
    expect(wb2.sheetNames).toEqual(['Sheet1']);
  });

  it('should reject when maxUnzipSize is exceeded', () => {
    const wb = new Workbook();
    for (let i = 1; i <= 100; i++) {
      wb.setCellValue('Sheet1', `A${i}`, 'x'.repeat(200));
    }
    const buf = wb.writeBufferSync();
    expect(() => Workbook.openBufferSync(buf, { maxUnzipSize: 50 })).toThrow(/decompressed size/i);
  });

  it('should accept when maxUnzipSize is within limit', () => {
    const wb = new Workbook();
    const buf = wb.writeBufferSync();
    const wb2 = Workbook.openBufferSync(buf, { maxUnzipSize: 100_000_000 });
    expect(wb2.sheetNames).toEqual(['Sheet1']);
  });

  it('should combine sheetRows and sheets options', () => {
    const wb = new Workbook();
    wb.newSheet('Parsed');
    wb.newSheet('Skipped');
    for (let i = 1; i <= 10; i++) {
      wb.setCellValue('Parsed', `A${i}`, i);
      wb.setCellValue('Skipped', `A${i}`, i * 100);
    }
    const buf = wb.writeBufferSync();

    const wb2 = Workbook.openBufferSync(buf, {
      sheets: ['Parsed'],
      sheetRows: 3,
    });
    expect(wb2.getCellValue('Parsed', 'A3')).toBe(3);
    expect(wb2.getCellValue('Parsed', 'A4')).toBeNull();
    expect(wb2.getCellValue('Skipped', 'A1')).toBeNull();
  });

  it('should work with no options (backward compatible)', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'hello');
    const buf = wb.writeBufferSync();
    const wb2 = Workbook.openBufferSync(buf);
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe('hello');
  });

  it('should work with async open', async () => {
    const wb = new Workbook();
    for (let i = 1; i <= 10; i++) {
      wb.setCellValue('Sheet1', `A${i}`, i);
    }
    await wb.save(out);

    const wb2 = await Workbook.open(out, { sheetRows: 2 });
    expect(wb2.getCellValue('Sheet1', 'A2')).toBe(2);
    expect(wb2.getCellValue('Sheet1', 'A3')).toBeNull();
  });

  it('should work with async openBuffer', async () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'test');
    const buf = wb.writeBufferSync();

    const wb2 = await Workbook.openBuffer(buf, { sheets: ['Sheet1'] });
    expect(wb2.getCellValue('Sheet1', 'A1')).toBe('test');
  });

  it('sheetRows zero means no rows', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 42);
    const buf = wb.writeBufferSync();

    const wb2 = Workbook.openBufferSync(buf, { sheetRows: 0 });
    expect(wb2.getCellValue('Sheet1', 'A1')).toBeNull();
  });
});

describe('Additional Chart Types', () => {
  const out = tmpFile('test-additional-charts.xlsx');
  afterEach(async () => cleanup(out));

  const chartSeries = [{ name: 'S1', categories: 'Sheet1!$A$1:$A$3', values: 'Sheet1!$B$1:$B$3' }];

  const newChartTypes = [
    'pieOfPie',
    'barOfPie',
    'col3DCone',
    'col3DConeStacked',
    'col3DConePercentStacked',
    'col3DPyramid',
    'col3DPyramidStacked',
    'col3DPyramidPercentStacked',
    'col3DCylinder',
    'col3DCylinderStacked',
    'col3DCylinderPercentStacked',
    'contour',
    'wireframeContour',
    'bubble3D',
  ];

  for (const chartType of newChartTypes) {
    it(`should create a ${chartType} chart and save`, async () => {
      const wb = new Workbook();
      wb.setCellValue('Sheet1', 'A1', 'Category');
      wb.setCellValue('Sheet1', 'B1', 100);
      wb.addChart('Sheet1', 'D1', 'J10', { chartType, series: chartSeries });
      await wb.save(out);
      await expect(access(out)).resolves.toBeUndefined();
    });
  }

  it('should reject unknown chart type', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addChart('Sheet1', 'D1', 'J10', { chartType: 'unknownType', series: chartSeries }),
    ).toThrow();
  });
});

describe('Slicers', () => {
  const out = tmpFile('test-slicer.xlsx');
  afterEach(async () => cleanup(out));

  it('should add and get slicers', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'Table1',
      displayName: 'Table1',
      range: 'A1:D10',
      columns: [{ name: 'Status' }, { name: 'Region' }, { name: 'Category' }, { name: 'Sales' }],
      showHeaderRow: true,
      autoFilter: false,
    });
    wb.addSlicer('Sheet1', {
      name: 'StatusFilter',
      cell: 'F1',
      tableName: 'Table1',
      columnName: 'Status',
    });

    const slicers = wb.getSlicers('Sheet1');
    expect(slicers.length).toBe(1);
    expect(slicers[0].name).toBe('StatusFilter');
    expect(slicers[0].columnName).toBe('Status');
  });

  it('should add slicer with options', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'Table1',
      displayName: 'Table1',
      range: 'A1:D10',
      columns: [{ name: 'Status' }, { name: 'Region' }, { name: 'Category' }, { name: 'Sales' }],
      showHeaderRow: true,
      autoFilter: false,
    });
    wb.addSlicer('Sheet1', {
      name: 'RegionSlicer',
      cell: 'G2',
      tableName: 'Table1',
      columnName: 'Region',
      caption: 'Filter by Region',
      style: 'SlicerStyleLight1',
      width: 300,
      height: 250,
      showCaption: true,
      columnCount: 2,
    });

    const slicers = wb.getSlicers('Sheet1');
    expect(slicers.length).toBe(1);
    expect(slicers[0].caption).toBe('Filter by Region');
    expect(slicers[0].style).toBe('SlicerStyleLight1');
  });

  it('should delete a slicer', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'T1',
      displayName: 'T1',
      range: 'A1:B10',
      columns: [{ name: 'Col1' }, { name: 'Col2' }],
      showHeaderRow: true,
      autoFilter: false,
    });
    wb.addSlicer('Sheet1', {
      name: 'S1',
      cell: 'F1',
      tableName: 'T1',
      columnName: 'Col1',
    });

    expect(wb.getSlicers('Sheet1').length).toBe(1);
    wb.deleteSlicer('Sheet1', 'S1');
    expect(wb.getSlicers('Sheet1').length).toBe(0);
  });

  it('should save and open with slicers', async () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'Table1',
      displayName: 'Table1',
      range: 'A1:D10',
      columns: [{ name: 'Status' }, { name: 'Region' }, { name: 'Category' }, { name: 'Sales' }],
      showHeaderRow: true,
      autoFilter: false,
    });
    wb.addSlicer('Sheet1', {
      name: 'CategoryFilter',
      cell: 'F1',
      tableName: 'Table1',
      columnName: 'Category',
      caption: 'Category',
      style: 'SlicerStyleLight1',
    });

    await wb.save(out);
    const wb2 = await Workbook.open(out);
    const slicers = wb2.getSlicers('Sheet1');
    expect(slicers.length).toBe(1);
    expect(slicers[0].name).toBe('CategoryFilter');
    expect(slicers[0].columnName).toBe('Category');
  });

  it('should throw on duplicate slicer name', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'T1',
      displayName: 'T1',
      range: 'A1:B10',
      columns: [{ name: 'Col1' }, { name: 'Col2' }],
      showHeaderRow: true,
      autoFilter: false,
    });
    wb.addSlicer('Sheet1', {
      name: 'S1',
      cell: 'F1',
      tableName: 'T1',
      columnName: 'Col1',
    });

    expect(() =>
      wb.addSlicer('Sheet1', {
        name: 'S1',
        cell: 'G1',
        tableName: 'T1',
        columnName: 'Col2',
      }),
    ).toThrow(/already exists/);
  });

  it('should throw on delete non-existent slicer', () => {
    const wb = new Workbook();
    expect(() => wb.deleteSlicer('Sheet1', 'NoSuch')).toThrow(/not found/);
  });

  it('should support multiple slicers on same sheet', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'T1',
      displayName: 'T1',
      range: 'A1:B10',
      columns: [{ name: 'Col1' }, { name: 'Col2' }],
      showHeaderRow: true,
      autoFilter: false,
    });
    wb.addSlicer('Sheet1', {
      name: 'S1',
      cell: 'F1',
      tableName: 'T1',
      columnName: 'Col1',
    });
    wb.addSlicer('Sheet1', {
      name: 'S2',
      cell: 'G1',
      tableName: 'T1',
      columnName: 'Col2',
    });

    const slicers = wb.getSlicers('Sheet1');
    expect(slicers.length).toBe(2);
  });

  it('should throw on non-existent table', () => {
    const wb = new Workbook();
    expect(() =>
      wb.addSlicer('Sheet1', {
        name: 'S1',
        cell: 'F1',
        tableName: 'NoSuchTable',
        columnName: 'Col1',
      }),
    ).toThrow(/table.*not found/i);
  });

  it('should throw on non-existent column', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'T1',
      displayName: 'T1',
      range: 'A1:B10',
      columns: [{ name: 'Col1' }, { name: 'Col2' }],
      showHeaderRow: true,
      autoFilter: false,
    });
    expect(() =>
      wb.addSlicer('Sheet1', {
        name: 'S1',
        cell: 'F1',
        tableName: 'T1',
        columnName: 'NoSuchCol',
      }),
    ).toThrow(/column.*not found/i);
  });

  it('should resolve table_name in getSlicers', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'SalesTable',
      displayName: 'SalesTable',
      range: 'A1:C10',
      columns: [{ name: 'Region' }, { name: 'Product' }, { name: 'Amount' }],
      showHeaderRow: true,
      autoFilter: false,
    });
    wb.addSlicer('Sheet1', {
      name: 'RegionFilter',
      cell: 'E1',
      tableName: 'SalesTable',
      columnName: 'Region',
    });

    const slicers = wb.getSlicers('Sheet1');
    expect(slicers.length).toBe(1);
    expect(slicers[0].tableName).toBe('SalesTable');
    expect(slicers[0].columnName).toBe('Region');
  });
});

describe('Tables', () => {
  it('should add and get tables', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'Sales',
      displayName: 'Sales',
      range: 'A1:D10',
      columns: [{ name: 'Name' }, { name: 'Region' }, { name: 'Product' }, { name: 'Amount' }],
      showHeaderRow: true,
      autoFilter: false,
    });

    const tables = wb.getTables('Sheet1');
    expect(tables.length).toBe(1);
    expect(tables[0].name).toBe('Sales');
    expect(tables[0].range).toBe('A1:D10');
    expect(tables[0].columns).toEqual(['Name', 'Region', 'Product', 'Amount']);
  });

  it('should add multiple tables', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'T1',
      displayName: 'T1',
      range: 'A1:B10',
      columns: [{ name: 'Col1' }, { name: 'Col2' }],
      showHeaderRow: true,
      autoFilter: false,
    });
    wb.addTable('Sheet1', {
      name: 'T2',
      displayName: 'T2',
      range: 'D1:F10',
      columns: [{ name: 'X' }, { name: 'Y' }, { name: 'Z' }],
      showHeaderRow: true,
      autoFilter: false,
    });

    const tables = wb.getTables('Sheet1');
    expect(tables.length).toBe(2);
    expect(tables[0].name).toBe('T1');
    expect(tables[1].name).toBe('T2');
  });

  it('should throw on duplicate table name', () => {
    const wb = new Workbook();
    wb.addTable('Sheet1', {
      name: 'T1',
      displayName: 'T1',
      range: 'A1:B10',
      columns: [{ name: 'Col1' }, { name: 'Col2' }],
      showHeaderRow: true,
      autoFilter: false,
    });
    expect(() =>
      wb.addTable('Sheet1', {
        name: 'T1',
        displayName: 'T1',
        range: 'D1:E10',
        columns: [{ name: 'X' }, { name: 'Y' }],
        showHeaderRow: true,
        autoFilter: false,
      }),
    ).toThrow(/already exists/i);
  });
});

describe('SVG Renderer', () => {
  it('should render a sheet to valid SVG', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 'Score');
    wb.setCellValue('Sheet1', 'A2', 'Alice');
    wb.setCellValue('Sheet1', 'B2', 95);

    const svg = wb.renderToSvg({ sheetName: 'Sheet1' });

    expect(svg).toMatch(/^<svg/);
    expect(svg).toMatch(/<\/svg>$/);
    expect(svg).toContain('xmlns="http://www.w3.org/2000/svg"');
  });

  it('should include cell text in SVG output', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Hello');
    wb.setCellValue('Sheet1', 'B1', 42);

    const svg = wb.renderToSvg({ sheetName: 'Sheet1' });

    expect(svg).toContain('>Hello<');
    expect(svg).toContain('>42<');
  });

  it('should render with gridlines by default', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Test');

    const svg = wb.renderToSvg({ sheetName: 'Sheet1' });

    expect(svg).toContain('stroke="#D0D0D0"');
  });

  it('should omit gridlines when disabled', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Test');

    const svg = wb.renderToSvg({ sheetName: 'Sheet1', showGridlines: false });

    expect(svg).not.toContain('stroke="#D0D0D0"');
  });

  it('should render headers by default', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Test');

    const svg = wb.renderToSvg({ sheetName: 'Sheet1' });

    expect(svg).toContain('>A<');
    expect(svg).toContain('>1<');
  });

  it('should omit headers when disabled', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Test');

    const svg = wb.renderToSvg({ sheetName: 'Sheet1', showHeaders: false });

    expect(svg).not.toContain('fill="#F0F0F0"');
  });

  it('should render a sub-range', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Name');
    wb.setCellValue('Sheet1', 'B1', 'Score');
    wb.setCellValue('Sheet1', 'A2', 'Alice');

    const svg = wb.renderToSvg({ sheetName: 'Sheet1', range: 'A1:A2' });

    expect(svg).toContain('>Name<');
    expect(svg).toContain('>Alice<');
    expect(svg).not.toContain('>Score<');
  });

  it('should apply scale factor to dimensions', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Test');

    const svg1 = wb.renderToSvg({ sheetName: 'Sheet1', scale: 1 });
    const svg2 = wb.renderToSvg({ sheetName: 'Sheet1', scale: 2 });

    const widthRegex = /width="([^"]+)"/;
    const w1 = Number.parseFloat(svg1.match(widthRegex)![1]);
    const w2 = Number.parseFloat(svg2.match(widthRegex)![1]);

    expect(w2).toBeCloseTo(w1 * 2, 1);
  });

  it('should throw for invalid sheet name', () => {
    const wb = new Workbook();
    expect(() => wb.renderToSvg({ sheetName: 'NonExistent' })).toThrow();
  });

  it('should render styled cells with bold text', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Bold');
    const styleId = wb.addStyle({ font: { bold: true } });
    wb.setCellStyle('Sheet1', 'A1', styleId);

    const svg = wb.renderToSvg({ sheetName: 'Sheet1' });

    expect(svg).toContain('font-weight="bold"');
  });

  it('should render cell fill colors', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'Yellow');
    const styleId = wb.addStyle({ fill: { pattern: 'solid', fgColor: 'FFFFFF00' } });
    wb.setCellStyle('Sheet1', 'A1', styleId);

    const svg = wb.renderToSvg({ sheetName: 'Sheet1' });

    expect(svg).toContain('fill="#FFFF00"');
  });
});

describe('Threaded Comments', () => {
  const out = tmpFile('test-threaded-comments.xlsx');
  afterEach(async () => cleanup(out));

  it('should add and get a threaded comment', () => {
    const wb = new Workbook();
    const id = wb.addThreadedComment('Sheet1', 'A1', {
      author: 'Alice',
      text: 'Hello thread',
    });
    expect(id).toBeTruthy();

    const comments = wb.getThreadedComments('Sheet1');
    expect(comments).toHaveLength(1);
    expect(comments[0].cellRef).toBe('A1');
    expect(comments[0].text).toBe('Hello thread');
    expect(comments[0].author).toBe('Alice');
    expect(comments[0].done).toBe(false);
  });

  it('should add a reply to a threaded comment', () => {
    const wb = new Workbook();
    const parentId = wb.addThreadedComment('Sheet1', 'A1', {
      author: 'Alice',
      text: 'Initial comment',
    });
    wb.addThreadedComment('Sheet1', 'A1', {
      author: 'Bob',
      text: 'Reply to Alice',
      parentId,
    });

    const comments = wb.getThreadedComments('Sheet1');
    expect(comments).toHaveLength(2);
    expect(comments[1].parentId).toBe(parentId);
    expect(comments[1].text).toBe('Reply to Alice');
  });

  it('should get threaded comments by cell', () => {
    const wb = new Workbook();
    wb.addThreadedComment('Sheet1', 'A1', {
      author: 'Alice',
      text: 'On A1',
    });
    wb.addThreadedComment('Sheet1', 'B2', {
      author: 'Bob',
      text: 'On B2',
    });

    const a1 = wb.getThreadedCommentsByCell('Sheet1', 'A1');
    expect(a1).toHaveLength(1);
    expect(a1[0].text).toBe('On A1');

    const b2 = wb.getThreadedCommentsByCell('Sheet1', 'B2');
    expect(b2).toHaveLength(1);
    expect(b2[0].text).toBe('On B2');

    const c3 = wb.getThreadedCommentsByCell('Sheet1', 'C3');
    expect(c3).toHaveLength(0);
  });

  it('should delete a threaded comment', () => {
    const wb = new Workbook();
    const id = wb.addThreadedComment('Sheet1', 'A1', {
      author: 'Alice',
      text: 'Delete me',
    });
    wb.deleteThreadedComment('Sheet1', id);

    const comments = wb.getThreadedComments('Sheet1');
    expect(comments).toHaveLength(0);
  });

  it('should resolve and unresolve a threaded comment', () => {
    const wb = new Workbook();
    const id = wb.addThreadedComment('Sheet1', 'A1', {
      author: 'Alice',
      text: 'Resolve me',
    });

    wb.resolveThreadedComment('Sheet1', id, true);
    let comments = wb.getThreadedComments('Sheet1');
    expect(comments[0].done).toBe(true);

    wb.resolveThreadedComment('Sheet1', id, false);
    comments = wb.getThreadedComments('Sheet1');
    expect(comments[0].done).toBe(false);
  });

  it('should add and get persons', () => {
    const wb = new Workbook();
    const id = wb.addPerson({
      displayName: 'Alice',
      userId: 'alice@example.com',
      providerId: 'ADAL',
    });
    expect(id).toBeTruthy();

    const persons = wb.getPersons();
    expect(persons).toHaveLength(1);
    expect(persons[0].displayName).toBe('Alice');
    expect(persons[0].userId).toBe('alice@example.com');
  });

  it('should auto-create persons when adding comments', () => {
    const wb = new Workbook();
    wb.addThreadedComment('Sheet1', 'A1', {
      author: 'Alice',
      text: 'Auto person',
    });

    const persons = wb.getPersons();
    expect(persons).toHaveLength(1);
    expect(persons[0].displayName).toBe('Alice');
  });

  it('should round-trip threaded comments through save/open', async () => {
    const wb = new Workbook();
    const id = wb.addThreadedComment('Sheet1', 'A1', {
      author: 'Alice',
      text: 'Persisted comment',
    });
    wb.resolveThreadedComment('Sheet1', id, true);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const comments = wb2.getThreadedComments('Sheet1');
    expect(comments).toHaveLength(1);
    expect(comments[0].cellRef).toBe('A1');
    expect(comments[0].text).toBe('Persisted comment');
    expect(comments[0].author).toBe('Alice');
    expect(comments[0].done).toBe(true);

    const persons = wb2.getPersons();
    expect(persons).toHaveLength(1);
    expect(persons[0].displayName).toBe('Alice');
  });

  it('should round-trip threaded comments with replies', async () => {
    const wb = new Workbook();
    const parentId = wb.addThreadedComment('Sheet1', 'B2', {
      author: 'Bob',
      text: 'Parent comment',
    });
    wb.addThreadedComment('Sheet1', 'B2', {
      author: 'Alice',
      text: 'Reply comment',
      parentId,
    });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const comments = wb2.getThreadedComments('Sheet1');
    expect(comments).toHaveLength(2);
    expect(comments[0].text).toBe('Parent comment');
    expect(comments[1].text).toBe('Reply comment');
    expect(comments[1].parentId).toBe(parentId);
  });

  it('should support threaded comments on multiple sheets', async () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    wb.addThreadedComment('Sheet1', 'A1', {
      author: 'Alice',
      text: 'Sheet1 comment',
    });
    wb.addThreadedComment('Sheet2', 'C3', {
      author: 'Bob',
      text: 'Sheet2 comment',
    });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const s1 = wb2.getThreadedComments('Sheet1');
    expect(s1).toHaveLength(1);
    expect(s1[0].text).toBe('Sheet1 comment');

    const s2 = wb2.getThreadedComments('Sheet2');
    expect(s2).toHaveLength(1);
    expect(s2[0].text).toBe('Sheet2 comment');
  });
});
