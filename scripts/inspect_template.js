const XLSX = require('../libs/xlsx.full.min.js');
const fs = require('fs');

const path = require('path');
const filePath = path.resolve(__dirname, '../checklist-pretrip_inspect-progress-lvl2.1 full.xls');
const buf = fs.readFileSync(filePath);
const wb = XLSX.read(buf, { type: 'buffer', cellStyles: true, cellNF: true });
console.log('Sheet names:', wb.SheetNames);
const sheetName = wb.SheetNames[0];
const sheet = wb.Sheets[sheetName];
console.log('Range:', sheet['!ref']);
console.log('Merges count:', (sheet['!merges'] || []).length);
console.log('Merges first 20:', JSON.stringify((sheet['!merges'] || []).slice(0, 20), null, 2));
console.log('Cols:', JSON.stringify(sheet['!cols'], null, 2));
console.log('Rows (first 15):', JSON.stringify((sheet['!rows'] || []).slice(0, 15), null, 2));

console.log('\n--- FIRST 15 ROWS CELL DUMP ---');
for (let r = 0; r <= 15; r++) {
  let cells = [];
  for (let c = 0; c < 16; c++) {
    const addr = XLSX.utils.encode_cell({r, c});
    if (sheet[addr]) {
      cells.push(`${addr}(t:${sheet[addr].t}): ${JSON.stringify(sheet[addr].v)}`);
    }
  }
  if (cells.length > 0) {
    console.log(`Row ${r + 1}: ${cells.join(' | ')}`);
  }
}

// Let's also check sample row 9 (first data row) and row 10
console.log('\n--- SAMPLE DATA ROW 9 DETAILED ---');
for (let c = 0; c < 16; c++) {
  const addr = XLSX.utils.encode_cell({r: 8, c});
  if (sheet[addr]) {
    console.log(`Col ${c} (${addr}):`, sheet[addr]);
  }
}
