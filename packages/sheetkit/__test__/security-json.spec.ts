import assert from 'node:assert';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../index.js';

describe('toJSON property safety', () => {
  it('preserves special headers as own data properties', () => {
    const wb = new Workbook();
    const dateValue = { type: 'date' as const, serial: 45292, iso: '2024-01-01' };
    wb.setCellValue('Sheet1', 'A1', 45292);
    wb.setCellStyle('Sheet1', 'A1', wb.addStyle({ numFmtId: 14 }));
    wb.setRowValues('Sheet1', 1, 'B', [
      'constructor-value',
      'prototype-value',
      'first-value',
      'last-value',
      'Unicode value',
    ]);

    const [row] = wb.toJSON('Sheet1', {
      header: ['__proto__', 'constructor', 'prototype', 'duplicate', 'duplicate', '顧客名'],
    });
    assert(row != null);

    expect(Object.getPrototypeOf(row)).toBe(Object.prototype);
    expect(Object.prototype).not.toHaveProperty('serial');
    expect(row).toEqual({
      ['__proto__']: dateValue,
      constructor: 'constructor-value',
      prototype: 'prototype-value',
      duplicate: 'last-value',
      顧客名: 'Unicode value',
    });

    for (const key of ['__proto__', 'constructor', 'prototype', 'duplicate', '顧客名']) {
      const descriptor = Object.getOwnPropertyDescriptor(row, key);
      assert(descriptor != null);
      expect(descriptor.writable).toBe(true);
      expect(descriptor.enumerable).toBe(true);
      expect(descriptor.configurable).toBe(true);
    }

    const parsed = JSON.parse(JSON.stringify(row)) as Record<string, unknown>;
    expect(Object.getPrototypeOf(parsed)).toBe(Object.prototype);
    expect(Object.hasOwn(parsed, '__proto__')).toBe(true);
    expect(parsed).toEqual(row);
  });

  it('does not invoke inherited setters for workbook-controlled headers', () => {
    const key = '__sheetkit_inherited_setter__';
    let setterCalls = 0;
    Object.defineProperty(Object.prototype, key, {
      configurable: true,
      set: () => {
        setterCalls++;
      },
    });

    try {
      const workbook = new Workbook();
      workbook.setCellValue('Sheet1', 'A1', 'value');
      const [row] = workbook.toJSON('Sheet1', { header: [key] });
      assert(row != null);

      expect(setterCalls).toBe(0);
      expect(Object.hasOwn(row, key)).toBe(true);
      expect(Object.getOwnPropertyDescriptor(row, key)?.value).toBe('value');
    } finally {
      Reflect.deleteProperty(Object.prototype, key);
    }
  });
});
