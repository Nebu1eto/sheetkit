import { Workbook } from 'sheetkit';

console.log('=== SheetKit Node.js 예제 ===\n');

// ── Phase 1: 워크북 생성 ──
const wb = new Workbook();
console.log('[Phase 1] 새 워크북 생성 완료. 시트:', wb.sheetNames);

// ── Phase 2: 셀 값 읽기/쓰기 ──
wb.setCellValue('Sheet1', 'A1', '이름');
wb.setCellValue('Sheet1', 'B1', '나이');
wb.setCellValue('Sheet1', 'C1', '활성');
wb.setCellValue('Sheet1', 'A2', '홍길동');
wb.setCellValue('Sheet1', 'B2', 30);
wb.setCellValue('Sheet1', 'C2', true);
wb.setCellValue('Sheet1', 'A3', '김철수');
wb.setCellValue('Sheet1', 'B3', 25);
wb.setCellValue('Sheet1', 'C3', false);

const val = wb.getCellValue('Sheet1', 'A1');
console.log('[Phase 2] A1 셀 값:', val);

// ── Phase 5: 시트 관리 ──
const idx = wb.newSheet('매출데이터');
console.log(`[Phase 5] '매출데이터' 시트 추가 (인덱스: ${idx})`);
wb.setSheetName('매출데이터', 'Sales');
wb.copySheet('Sheet1', 'Sheet1_복사');
wb.setActiveSheet('Sheet1');
console.log('[Phase 5] 시트 목록:', wb.sheetNames);

// ── Phase 3: 행/열 조작 ──
wb.setRowHeight('Sheet1', 1, 25);
wb.setColWidth('Sheet1', 'A', 20);
wb.setColWidth('Sheet1', 'B', 15);
wb.setColWidth('Sheet1', 'C', 12);
wb.insertRows('Sheet1', 1, 1); // 헤더 위에 제목 행 삽입
wb.setCellValue('Sheet1', 'A1', '직원 목록');
console.log('[Phase 3] 행/열 크기 조정 및 행 삽입 완료');

// ── Phase 4: 스타일 ──
// 제목 스타일
const titleStyleId = wb.addStyle({
  font: { name: 'Arial', size: 16, bold: true, color: '#FFFFFF' },
  fill: { pattern: 'solid', fgColor: '#4472C4' },
  alignment: { horizontal: 'center', vertical: 'center' },
});
wb.setCellStyle('Sheet1', 'A1', titleStyleId);

// 헤더 스타일
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
console.log('[Phase 4] 스타일 적용 완료 (제목 + 헤더)');

// ── Phase 7: 차트 ──
// Sales 시트에 차트 데이터
wb.setCellValue('Sales', 'A1', '분기');
wb.setCellValue('Sales', 'B1', '매출');
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
  title: '분기별 매출',
  series: [
    {
      name: '매출',
      categories: 'Sales!$A$2:$A$5',
      values: 'Sales!$B$2:$B$5',
    },
  ],
  showLegend: true,
});
console.log('[Phase 7] 차트 추가 완료 (Sales 시트)');

// ── Phase 7: 이미지 (1x1 PNG placeholder) ──
const pngData = Buffer.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d,
  0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
  0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
  0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49,
  0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);
wb.addImage('Sheet1', {
  data: pngData,
  format: 'png',
  fromCell: 'E1',
  widthPx: 64,
  heightPx: 64,
});
console.log('[Phase 7] 이미지 추가 완료');

// ── Phase 8: 데이터 유효성 검사 ──
wb.addDataValidation('Sales', {
  sqref: 'C2:C5',
  validationType: 'list',
  formula1: '"달성,미달성,진행중"',
  allowBlank: true,
  showInputMessage: true,
  promptTitle: '상태 선택',
  promptMessage: '드롭다운에서 상태를 선택하세요',
  showErrorMessage: true,
  errorStyle: 'stop',
  errorTitle: '오류',
  errorMessage: '목록에서 선택해주세요',
});
console.log('[Phase 8] 데이터 유효성 검사 추가 완료');

// ── Phase 8: 코멘트 ──
wb.addComment('Sheet1', {
  cell: 'A1',
  author: '관리자',
  text: '이 시트는 직원 목록을 포함합니다.',
});
console.log('[Phase 8] 코멘트 추가 완료');

// ── Phase 8: 자동 필터 ──
wb.setAutoFilter('Sheet1', 'A2:C4');
console.log('[Phase 8] 자동 필터 설정 완료');

// ── Phase 9: StreamWriter ──
const sw = wb.newStreamWriter('대용량시트');
sw.setColWidth(1, 15);
sw.setColWidth(2, 10);
sw.writeRow(1, ['항목', '값']);
for (let i = 2; i <= 100; i++) {
  sw.writeRow(i, [`항목_${i - 1}`, i * 10]);
}
sw.addMergeCell('A1:B1');
wb.applyStreamWriter(sw);
console.log('[Phase 9] StreamWriter로 100행 작성 완료');

// ── Phase 10: 문서 속성 ──
wb.setDocProps({
  title: 'SheetKit 예제 문서',
  creator: 'SheetKit Node.js Example',
  description: 'SheetKit의 모든 기능을 보여주는 예제 파일',
});
wb.setAppProps({
  application: 'SheetKit',
  company: 'SheetKit Project',
});
wb.setCustomProperty('프로젝트', 'SheetKit');
wb.setCustomProperty('버전', 1);
wb.setCustomProperty('릴리즈', false);
console.log('[Phase 10] 문서 속성 설정 완료');

// ── Phase 10: 워크북 보호 ──
wb.protectWorkbook({
  password: 'demo',
  lockStructure: true,
  lockWindows: false,
  lockRevision: false,
});
console.log('[Phase 10] 워크북 보호 설정 완료');

// ── 저장 ──
wb.save('output.xlsx');
console.log('\n✅ output.xlsx 파일이 생성되었습니다!');
console.log('시트 목록:', wb.sheetNames);
