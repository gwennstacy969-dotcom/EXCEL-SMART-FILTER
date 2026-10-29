const fs = require('fs');
const content = fs.readFileSync('checklist-pretrip_inspect-progress-lvl2.1 full.xls', 'utf8');

const row8Pos = content.indexOf('>NO.<');
const row8Start = content.lastIndexOf('<Row', row8Pos);
const row8End = content.indexOf('</Row>', row8Start) + 6;
console.log('--- ROW 8 (HEADER) ---');
console.log(content.slice(row8Start, row8End));

const row9Start = content.indexOf('<Row', row8End);
const row9End = content.indexOf('</Row>', row9Start) + 6;
console.log('\n--- ROW 9 (ODD DATA ROW) ---');
console.log(content.slice(row9Start, row9End));

const row10Start = content.indexOf('<Row', row9End);
const row10End = content.indexOf('</Row>', row10Start) + 6;
console.log('\n--- ROW 10 (EVEN DATA ROW) ---');
console.log(content.slice(row10Start, row10End));

const row11Start = content.indexOf('<Row', row10End);
const row11End = content.indexOf('</Row>', row11Start) + 6;
console.log('\n--- ROW 11 (ODD DATA ROW) ---');
console.log(content.slice(row11Start, row11End));

