import { describe, it, expect, afterEach } from 'vitest';
import { existsSync, unlinkSync } from 'node:fs';
import { join } from 'node:path';
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

describe('Phase 1 - Basic I/O', () => {
    const out = tmpFile('test-basic.xlsx');
    afterEach(() => cleanup(out));

    it('should create a new workbook', () => {
        const wb = new Workbook();
        expect(wb.sheetNames).toEqual(['Sheet1']);
    });

    it('should save and open a workbook', () => {
        const wb = new Workbook();
        wb.save(out);
        expect(existsSync(out)).toBe(true);
        const wb2 = Workbook.open(out);
        expect(wb2.sheetNames).toEqual(['Sheet1']);
    });

    it('should throw on invalid path', () => {
        expect(() => Workbook.open('/nonexistent/path.xlsx')).toThrow();
    });
});

describe('Phase 2 - Cell Operations', () => {
    const out = tmpFile('test-cell.xlsx');
    afterEach(() => cleanup(out));

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

    it('should roundtrip cell values through save/open', () => {
        const wb = new Workbook();
        wb.setCellValue('Sheet1', 'A1', 'text');
        wb.setCellValue('Sheet1', 'B1', 123);
        wb.setCellValue('Sheet1', 'C1', true);
        wb.save(out);

        const wb2 = Workbook.open(out);
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

describe('Phase 7 - Charts & Images', () => {
    const out = tmpFile('test-chart.xlsx');
    afterEach(() => cleanup(out));

    it('should add a column chart and save', () => {
        const wb = new Workbook();
        wb.setCellValue('Sheet1', 'A1', 'Category');
        wb.setCellValue('Sheet1', 'B1', 100);
        wb.addChart('Sheet1', 'D1', 'J10', {
            chartType: 'col',
            series: [{ name: 'S1', categories: 'Sheet1!$A$1:$A$3', values: 'Sheet1!$B$1:$B$3' }],
        });
        wb.save(out);
        expect(existsSync(out)).toBe(true);
    });

    it('should add a PNG image and save', () => {
        const wb = new Workbook();
        const pngData = Buffer.from([
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
            0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
            0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41,
            0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
            0x42, 0x60, 0x82,
        ]);
        wb.addImage('Sheet1', {
            data: pngData,
            format: 'png',
            fromCell: 'A1',
            widthPx: 100,
            heightPx: 100,
        });
        wb.save(out);
        expect(existsSync(out)).toBe(true);
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
    afterEach(() => cleanup(out));

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

    it('should roundtrip merge cells through save/open', () => {
        const wb = new Workbook();
        wb.setCellValue('Sheet1', 'A1', 'Merged');
        wb.mergeCells('Sheet1', 'A1', 'C3');
        wb.save(out);

        const wb2 = Workbook.open(out);
        const merged = wb2.getMergeCells('Sheet1');
        expect(merged).toEqual(['A1:C3']);
    });
});

describe('Phase 8 - Auto-filter', () => {
    const out = tmpFile('test-autofilter.xlsx');
    afterEach(() => cleanup(out));

    it('should set and remove auto filter', () => {
        const wb = new Workbook();
        wb.setAutoFilter('Sheet1', 'A1:C10');
        wb.removeAutoFilter('Sheet1');
    });

    it('should set auto filter and save', () => {
        const wb = new Workbook();
        wb.setCellValue('Sheet1', 'A1', 'Name');
        wb.setCellValue('Sheet1', 'B1', 'Age');
        wb.setAutoFilter('Sheet1', 'A1:B1');
        wb.save(out);
        expect(existsSync(out)).toBe(true);
    });
});

describe('Phase 9 - StreamWriter', () => {
    const out = tmpFile('test-stream.xlsx');
    afterEach(() => cleanup(out));

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

    it('should roundtrip stream writer data', () => {
        const wb = new Workbook();
        const sw = wb.newStreamWriter('Data');
        sw.writeRow(1, ['Name', 'Value']);
        sw.writeRow(2, ['A', 100]);
        wb.applyStreamWriter(sw);
        wb.save(out);

        const wb2 = Workbook.open(out);
        expect(wb2.sheetNames).toContain('Data');
        expect(wb2.getCellValue('Data', 'A1')).toBe('Name');
        expect(wb2.getCellValue('Data', 'B2')).toBe(100);
    });
});

describe('Phase 10 - Document Properties', () => {
    const out = tmpFile('test-docprops.xlsx');
    afterEach(() => cleanup(out));

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

    it('should roundtrip doc properties', () => {
        const wb = new Workbook();
        wb.setDocProps({ title: 'My Doc', creator: 'Author' });
        wb.save(out);

        const wb2 = Workbook.open(out);
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
    afterEach(() => cleanup(out));

    it('should protect and unprotect workbook', () => {
        const wb = new Workbook();
        expect(wb.isWorkbookProtected()).toBe(false);
        wb.protectWorkbook({ lockStructure: true });
        expect(wb.isWorkbookProtected()).toBe(true);
        wb.unprotectWorkbook();
        expect(wb.isWorkbookProtected()).toBe(false);
    });

    it('should protect with password and roundtrip', () => {
        const wb = new Workbook();
        wb.protectWorkbook({ password: 'secret', lockStructure: true, lockWindows: true });
        expect(wb.isWorkbookProtected()).toBe(true);
        wb.save(out);

        const wb2 = Workbook.open(out);
        expect(wb2.isWorkbookProtected()).toBe(true);
    });
});

describe('Hyperlinks', () => {
    const out = tmpFile('test-hyperlink.xlsx');
    afterEach(() => cleanup(out));

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
        expect(info!.linkType).toBe('external');
        expect(info!.target).toBe('https://example.com');
        expect(info!.display).toBe('Example');
        expect(info!.tooltip).toBe('Visit Example');
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
        expect(info!.linkType).toBe('internal');
        expect(info!.target).toBe('Sheet1!C1');
        expect(info!.display).toBe('Go to C1');
    });

    it('should set and get email hyperlink', () => {
        const wb = new Workbook();
        wb.setCellHyperlink('Sheet1', 'C3', {
            linkType: 'email',
            target: 'mailto:user@example.com',
        });
        const info = wb.getCellHyperlink('Sheet1', 'C3');
        expect(info).not.toBeNull();
        expect(info!.linkType).toBe('email');
        expect(info!.target).toBe('mailto:user@example.com');
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
        expect(info!.target).toBe('https://new.com');
        expect(info!.display).toBe('New Link');
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

        expect(wb.getCellHyperlink('Sheet1', 'A1')!.linkType).toBe('external');
        expect(wb.getCellHyperlink('Sheet1', 'B1')!.linkType).toBe('internal');
        expect(wb.getCellHyperlink('Sheet1', 'C1')!.linkType).toBe('email');
    });

    it('should roundtrip hyperlinks through save/open', () => {
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
        wb.save(out);

        const wb2 = Workbook.open(out);
        const a1 = wb2.getCellHyperlink('Sheet1', 'A1');
        expect(a1).not.toBeNull();
        expect(a1!.linkType).toBe('external');
        expect(a1!.target).toBe('https://rust-lang.org');
        expect(a1!.display).toBe('Rust');
        expect(a1!.tooltip).toBe('Rust Homepage');

        const b1 = wb2.getCellHyperlink('Sheet1', 'B1');
        expect(b1).not.toBeNull();
        expect(b1!.linkType).toBe('internal');
        expect(b1!.target).toBe('Sheet1!C1');

        const c1 = wb2.getCellHyperlink('Sheet1', 'C1');
        expect(c1).not.toBeNull();
        expect(c1!.linkType).toBe('email');
        expect(c1!.target).toBe('mailto:hello@example.com');
    });
});
