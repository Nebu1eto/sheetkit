import { Workbook } from 'sheetkit';

console.log('=== SheetKit Node.js Example ===\n');

// ── Phase 1: Create workbook ──
const wb = new Workbook();
console.log('[Phase 1] New workbook created. Sheets:', wb.sheetNames);

// ── Phase 2: Read/write cell values ──
wb.setCellValue('Sheet1', 'A1', 'Name');
wb.setCellValue('Sheet1', 'B1', 'Age');
wb.setCellValue('Sheet1', 'C1', 'Active');
wb.setCellValue('Sheet1', 'A2', 'John Doe');
wb.setCellValue('Sheet1', 'B2', 30);
wb.setCellValue('Sheet1', 'C2', true);
wb.setCellValue('Sheet1', 'A3', 'Jane Smith');
wb.setCellValue('Sheet1', 'B3', 25);
wb.setCellValue('Sheet1', 'C3', false);

const val = wb.getCellValue('Sheet1', 'A1');
console.log('[Phase 2] A1 cell value:', val);

// ── Phase 5: Sheet management ──
const idx = wb.newSheet('SalesData');
console.log(`[Phase 5] 'SalesData' sheet added (index: ${idx})`);
wb.setSheetName('SalesData', 'Sales');
wb.copySheet('Sheet1', 'Sheet1_Copy');
wb.setActiveSheet('Sheet1');
console.log('[Phase 5] Sheet list:', wb.sheetNames);

// ── Phase 3: Row/column operations ──
wb.setRowHeight('Sheet1', 1, 25);
wb.setColWidth('Sheet1', 'A', 20);
wb.setColWidth('Sheet1', 'B', 15);
wb.setColWidth('Sheet1', 'C', 12);
wb.insertRows('Sheet1', 1, 1); // Insert title row above header
wb.setCellValue('Sheet1', 'A1', 'Employee List');
console.log('[Phase 3] Row/column sizing and row insertion complete');

// ── Phase 4: Styles ──
// Title style
const titleStyleId = wb.addStyle({
	font: { name: 'Arial', size: 16, bold: true, color: '#FFFFFF' },
	fill: { pattern: 'solid', fgColor: '#4472C4' },
	alignment: { horizontal: 'center', vertical: 'center' },
});
wb.setCellStyle('Sheet1', 'A1', titleStyleId);

// Header style
const headerStyleId = wb.addStyle({
	font: { bold: true, size: 11, color: '#FFFFFF' },
	fill: { pattern: 'solid', fgColor: '#5B9BD5' },
	border: {
		bottom: { style: 'thin', color: '#000000' },
	},
	alignment: { horizontal: 'center' },
});
wb.setCellStyle('Sheet1', 'A2', headerStyleId);
wb.setCellStyle('Sheet1', 'B2', headerStyleId);
wb.setCellStyle('Sheet1', 'C2', headerStyleId);
console.log('[Phase 4] Styles applied (title + header)');

// ── Phase 7: Chart ──
// Chart data on Sales sheet
wb.setCellValue('Sales', 'A1', 'Quarter');
wb.setCellValue('Sales', 'B1', 'Revenue');
wb.setCellValue('Sales', 'A2', 'Q1');
wb.setCellValue('Sales', 'B2', 1500);
wb.setCellValue('Sales', 'A3', 'Q2');
wb.setCellValue('Sales', 'B3', 2300);
wb.setCellValue('Sales', 'A4', 'Q3');
wb.setCellValue('Sales', 'B4', 1800);
wb.setCellValue('Sales', 'A5', 'Q4');
wb.setCellValue('Sales', 'B5', 2700);

wb.addChart('Sales', 'D1', 'K15', {
	chartType: 'col',
	title: 'Quarterly Revenue',
	series: [
		{
			name: 'Revenue',
			categories: 'Sales!$A$2:$A$5',
			values: 'Sales!$B$2:$B$5',
		},
	],
	showLegend: true,
});
console.log('[Phase 7] Chart added (Sales sheet)');

// ── Phase 7: Image (1x1 PNG placeholder) ──
const pngData = Buffer.from([
	0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
	0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
	0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
	0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
	0x42, 0x60, 0x82,
]);
wb.addImage('Sheet1', {
	data: pngData,
	format: 'png',
	fromCell: 'E1',
	widthPx: 64,
	heightPx: 64,
});
console.log('[Phase 7] Image added');

// ── Phase 8: Data validation ──
wb.addDataValidation('Sales', {
	sqref: 'C2:C5',
	validationType: 'list',
	formula1: '"Achieved,Not Achieved,In Progress"',
	allowBlank: true,
	showInputMessage: true,
	promptTitle: 'Select Status',
	promptMessage: 'Select a status from the dropdown',
	showErrorMessage: true,
	errorStyle: 'stop',
	errorTitle: 'Error',
	errorMessage: 'Please select from the list',
});
console.log('[Phase 8] Data validation added');

// ── Phase 8: Comment ──
wb.addComment('Sheet1', {
	cell: 'A1',
	author: 'Admin',
	text: 'This sheet contains the employee list.',
});
console.log('[Phase 8] Comment added');

// ── Phase 8: Auto filter ──
wb.setAutoFilter('Sheet1', 'A2:C4');
console.log('[Phase 8] Auto filter set');

// ── Phase 9: StreamWriter ──
const sw = wb.newStreamWriter('LargeSheet');
sw.setColWidth(1, 15);
sw.setColWidth(2, 10);
sw.writeRow(1, ['Item', 'Value']);
for (let i = 2; i <= 100; i++) {
	sw.writeRow(i, [`Item_${i - 1}`, i * 10]);
}
sw.addMergeCell('A1:B1');
wb.applyStreamWriter(sw);
console.log('[Phase 9] StreamWriter wrote 100 rows');

// ── Phase 10: Document properties ──
wb.setDocProps({
	title: 'SheetKit Example Document',
	creator: 'SheetKit Node.js Example',
	description: 'An example file demonstrating all SheetKit features',
});
wb.setAppProps({
	application: 'SheetKit',
	company: 'SheetKit Project',
});
wb.setCustomProperty('Project', 'SheetKit');
wb.setCustomProperty('Version', 1);
wb.setCustomProperty('Release', false);
console.log('[Phase 10] Document properties set');

// ── Phase 10: Workbook protection ──
wb.protectWorkbook({
	password: 'demo',
	lockStructure: true,
	lockWindows: false,
	lockRevision: false,
});
console.log('[Phase 10] Workbook protection set');

// ── Save ──
wb.save('output.xlsx');
console.log('\noutput.xlsx has been created!');
console.log('Sheet list:', wb.sheetNames);
