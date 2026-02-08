import { describe, it, expect, afterEach } from 'vitest';
import { existsSync, unlinkSync } from 'node:fs';
import { join } from 'node:path';
import { Workbook } from '../index.js';

const TEST_OUTPUT = join(import.meta.dirname, 'test-output.xlsx');

afterEach(() => {
	if (existsSync(TEST_OUTPUT)) {
		unlinkSync(TEST_OUTPUT);
	}
});

describe('Workbook', () => {
	it('should create a new workbook', () => {
		const wb = new Workbook();
		expect(wb).toBeDefined();
	});

	it('should have default sheet names', () => {
		const wb = new Workbook();
		expect(wb.sheetNames).toEqual(['Sheet1']);
	});

	it('should save and open a workbook', () => {
		const wb = new Workbook();
		wb.save(TEST_OUTPUT);
		expect(existsSync(TEST_OUTPUT)).toBe(true);

		const wb2 = Workbook.open(TEST_OUTPUT);
		expect(wb2.sheetNames).toEqual(['Sheet1']);
	});

	it('should throw on opening non-existent file', () => {
		expect(() => Workbook.open('/tmp/nonexistent.xlsx')).toThrow();
	});
});
