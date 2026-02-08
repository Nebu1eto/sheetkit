import { existsSync, unlinkSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { Workbook } from '../index.js';

const TEST_DIR = import.meta.dirname;

function tmpFile(name: string) {
  return join(TEST_DIR, name);
}

function cleanup(...files: string[]) {
  for (const f of files) {
    if (existsSync(f)) unlinkSync(f);
  }
}

// Defined names tests

describe('Defined Names', () => {
  const out = tmpFile('test-defined-names.xlsx');
  afterEach(() => cleanup(out));

  it('should set and get a workbook-scoped defined name', () => {
    const wb = new Workbook();
    wb.setDefinedName({
      name: 'SalesData',
      value: 'Sheet1!$A$1:$D$10',
    });
    const info = wb.getDefinedName('SalesData');
    expect(info).not.toBeNull();
    expect(info!.name).toBe('SalesData');
    expect(info!.value).toBe('Sheet1!$A$1:$D$10');
    expect(info!.scope).toBeUndefined();
    expect(info!.comment).toBeUndefined();
  });

  it('should set and get a sheet-scoped defined name', () => {
    const wb = new Workbook();
    wb.setDefinedName({
      name: 'LocalRange',
      value: 'Sheet1!$B$2:$C$5',
      scope: 'Sheet1',
      comment: 'A local range',
    });
    const info = wb.getDefinedName('LocalRange', 'Sheet1');
    expect(info).not.toBeNull();
    expect(info!.name).toBe('LocalRange');
    expect(info!.value).toBe('Sheet1!$B$2:$C$5');
    expect(info!.scope).toBe('Sheet1');
    expect(info!.comment).toBe('A local range');
  });

  it('should return null for non-existent defined name', () => {
    const wb = new Workbook();
    const info = wb.getDefinedName('NonExistent');
    expect(info).toBeNull();
  });

  it('should get all defined names', () => {
    const wb = new Workbook();
    wb.setDefinedName({ name: 'Alpha', value: 'Sheet1!$A$1' });
    wb.setDefinedName({
      name: 'Beta',
      value: 'Sheet1!$B$1',
      scope: 'Sheet1',
    });
    const all = wb.getDefinedNames();
    expect(all.length).toBe(2);
    expect(all[0].name).toBe('Alpha');
    expect(all[0].scope).toBeUndefined();
    expect(all[1].name).toBe('Beta');
    expect(all[1].scope).toBe('Sheet1');
  });

  it('should update existing defined name (no duplication)', () => {
    const wb = new Workbook();
    wb.setDefinedName({ name: 'Range', value: 'Sheet1!$A$1:$A$10' });
    wb.setDefinedName({
      name: 'Range',
      value: 'Sheet1!$A$1:$A$50',
      comment: 'Updated',
    });
    const all = wb.getDefinedNames();
    expect(all.length).toBe(1);
    expect(all[0].value).toBe('Sheet1!$A$1:$A$50');
    expect(all[0].comment).toBe('Updated');
  });

  it('should delete a defined name', () => {
    const wb = new Workbook();
    wb.setDefinedName({ name: 'ToRemove', value: 'Sheet1!$A$1' });
    expect(wb.getDefinedNames().length).toBe(1);
    wb.deleteDefinedName('ToRemove');
    expect(wb.getDefinedNames().length).toBe(0);
    expect(wb.getDefinedName('ToRemove')).toBeNull();
  });

  it('should allow same name in different scopes', () => {
    const wb = new Workbook();
    wb.setDefinedName({ name: 'Total', value: 'Sheet1!$A$1' });
    wb.setDefinedName({
      name: 'Total',
      value: 'Sheet1!$B$1',
      scope: 'Sheet1',
    });
    const all = wb.getDefinedNames();
    expect(all.length).toBe(2);

    const wbScoped = wb.getDefinedName('Total');
    expect(wbScoped!.value).toBe('Sheet1!$A$1');

    const sheetScoped = wb.getDefinedName('Total', 'Sheet1');
    expect(sheetScoped!.value).toBe('Sheet1!$B$1');
  });

  it('should preserve defined names across save/open roundtrip', async () => {
    const wb = new Workbook();
    wb.setDefinedName({
      name: 'Revenue',
      value: 'Sheet1!$E$1:$E$100',
      comment: 'Revenue column',
    });
    wb.setDefinedName({
      name: 'Local',
      value: 'Sheet1!$A$1',
      scope: 'Sheet1',
    });
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    const all = wb2.getDefinedNames();
    expect(all.length).toBe(2);

    const revenue = wb2.getDefinedName('Revenue');
    expect(revenue).not.toBeNull();
    expect(revenue!.value).toBe('Sheet1!$E$1:$E$100');
    expect(revenue!.comment).toBe('Revenue column');

    const local = wb2.getDefinedName('Local', 'Sheet1');
    expect(local).not.toBeNull();
    expect(local!.value).toBe('Sheet1!$A$1');
    expect(local!.scope).toBe('Sheet1');
  });
});

// Sheet protection tests

describe('Sheet Protection', () => {
  const out = tmpFile('test-sheet-protection.xlsx');
  afterEach(() => cleanup(out));

  it('should protect a sheet with default config', () => {
    const wb = new Workbook();
    expect(wb.isSheetProtected('Sheet1')).toBe(false);
    wb.protectSheet('Sheet1');
    expect(wb.isSheetProtected('Sheet1')).toBe(true);
  });

  it('should protect a sheet with password and permissions', () => {
    const wb = new Workbook();
    wb.protectSheet('Sheet1', {
      password: 'secret',
      formatCells: true,
      insertRows: true,
      sort: true,
    });
    expect(wb.isSheetProtected('Sheet1')).toBe(true);
  });

  it('should unprotect a sheet', () => {
    const wb = new Workbook();
    wb.protectSheet('Sheet1', { password: 'test' });
    expect(wb.isSheetProtected('Sheet1')).toBe(true);
    wb.unprotectSheet('Sheet1');
    expect(wb.isSheetProtected('Sheet1')).toBe(false);
  });

  it('should throw for non-existent sheet', () => {
    const wb = new Workbook();
    expect(() => wb.protectSheet('NonExistent')).toThrow();
    expect(() => wb.unprotectSheet('NonExistent')).toThrow();
    expect(() => wb.isSheetProtected('NonExistent')).toThrow();
  });

  it('should preserve sheet protection across save/open roundtrip', async () => {
    const wb = new Workbook();
    wb.protectSheet('Sheet1', {
      password: 'pass123',
      selectLockedCells: true,
      selectUnlockedCells: true,
      formatCells: true,
    });
    expect(wb.isSheetProtected('Sheet1')).toBe(true);
    await wb.save(out);

    const wb2 = await Workbook.open(out);
    expect(wb2.isSheetProtected('Sheet1')).toBe(true);
  });

  it('should protect only the specified sheet', () => {
    const wb = new Workbook();
    wb.newSheet('Sheet2');
    wb.protectSheet('Sheet1');
    expect(wb.isSheetProtected('Sheet1')).toBe(true);
    expect(wb.isSheetProtected('Sheet2')).toBe(false);
  });
});
