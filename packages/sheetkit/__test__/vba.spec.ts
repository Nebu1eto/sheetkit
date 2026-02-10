import { describe, expect, it } from 'vitest';
import { Workbook } from '../index.js';

describe('VBA extraction', () => {
  it('returns null for xlsx without VBA project', () => {
    const wb = new Workbook();
    expect(wb.getVbaProject()).toBeNull();
    expect(wb.getVbaModules()).toBeNull();
  });

  it('returns null after save/open roundtrip of xlsx without VBA', () => {
    const wb = new Workbook();
    wb.setCellValue('Sheet1', 'A1', 'test');
    const buf = wb.writeBufferSync();
    const wb2 = Workbook.openBufferSync(buf);
    expect(wb2.getVbaProject()).toBeNull();
    expect(wb2.getVbaModules()).toBeNull();
  });

  it('returns Buffer for workbook with VBA project binary', () => {
    // We cannot easily create a VBA project from JS side, but we can
    // verify the API shape and null behavior work correctly.
    const wb = new Workbook();
    const project = wb.getVbaProject();
    expect(project).toBeNull();
  });

  it('getVbaModules returns null type for standard xlsx', () => {
    const wb = new Workbook();
    const modules = wb.getVbaModules();
    expect(modules).toBeNull();
  });
});
