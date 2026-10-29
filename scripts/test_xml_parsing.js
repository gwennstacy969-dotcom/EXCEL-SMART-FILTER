const fs = require('fs');

const rawXml = fs.readFileSync('checklist-pretrip_inspect-progress-lvl2.1 full.xls', 'utf8');

console.time('splitRows');
const rowTag = '<Row';
const rows = [];
let pos = 0;
while (true) {
  const start = rawXml.indexOf(rowTag, pos);
  if (start === -1) break;
  const end = rawXml.indexOf('</Row>', start);
  if (end === -1) break;
  rows.push(rawXml.slice(start, end + 6));
  pos = end + 6;
}
console.timeEnd('splitRows');
console.log('Total XML rows found:', rows.length);
console.log('Row 0 (Excel row 1):', rows[0].slice(0, 80));
console.log('Row 7 (Excel row 8 header):', rows[7].slice(0, 80));
console.log('Row 8 (Excel row 9 first data row):', rows[8].slice(0, 80));
console.log('Row 9 (Excel row 10 second data row):', rows[9].slice(0, 80));

