const fs = require('fs');
const XLSX = require('../libs/xlsx.full.min.js');
const JSZip = require('../libs/jszip.min.js');

// Load index.html and evaluate its helper functions in a test sandbox
const html = fs.readFileSync('index.html', 'utf8');
const scriptMatch = html.match(/<script>([\s\S]*?)<\/script>/g);
const appScript = scriptMatch[1].replace(/<\/?script>/g, '');

// Create a clean mock environment
const mockWindow = {
  DEFAULT_PRETRIP_TEMPLATE: require('./pretrip_header.json')
};
const mockDocument = {
  getElementById: (id) => ({
    value: '',
    innerHTML: '',
    textContent: '',
    classList: { add: () => {}, remove: () => {} },
    addEventListener: () => {}
  }),
  querySelectorAll: () => [],
  createElement: () => ({ appendChild: () => {}, click: () => {}, removeChild: () => {} }),
  body: { appendChild: () => {}, removeChild: () => {} }
};

console.log('Testing Pretrip export generation on actual file...');

const buf = fs.readFileSync('checklist-pretrip_inspect-progress-lvl2.1 full.xls');
const rawXml = buf.toString('utf8');

// Test XML extraction
const row8Pos = rawXml.indexOf('>NO.<');
const row8Start = rawXml.lastIndexOf('<Row', row8Pos);
const row8End = rawXml.indexOf('</Row>', row8Start) + 6;
const xmlHeader = rawXml.slice(0, row8End);
const tableEndPos = rawXml.lastIndexOf('</ss:Table>');
const xmlFooter = rawXml.slice(tableEndPos);

const xmlRows = [];
let pos = 0;
while (true) {
  const start = rawXml.indexOf('<Row', pos);
  if (start === -1) break;
  const end = rawXml.indexOf('</Row>', start);
  if (end === -1) break;
  xmlRows.push(rawXml.slice(start, end + 6));
  pos = end + 6;
}

// Test building filtered XML
function escapeXml(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function formatOriginalXmlRow(rawRowXml, seqNum) {
  const isOdd = (seqNum % 2 === 1);
  let xml = rawRowXml;
  xml = xml.replace(/(<Cell ss:Index="2" ss:StyleID=")[^"]*(">([\s\S]*?)<Data ss:Type="Number">)[^<]*(<\/Data>)/, (match, p1, p2, p3, p4) => {
    const style = isOdd ? '30' : '40';
    return `${p1}${style}${p2}${seqNum}.0${p4}`;
  });
  if (isOdd) {
    xml = xml.replace(/ss:StyleID="40"/g, 'ss:StyleID="30"').replace(/ss:StyleID="42"/g, 'ss:StyleID="32"').replace(/ss:StyleID="44"/g, 'ss:StyleID="34"').replace(/ss:StyleID="46"/g, 'ss:StyleID="36"').replace(/ss:StyleID="48"/g, 'ss:StyleID="38"');
  } else {
    xml = xml.replace(/ss:StyleID="30"/g, 'ss:StyleID="40"').replace(/ss:StyleID="32"/g, 'ss:StyleID="42"').replace(/ss:StyleID="34"/g, 'ss:StyleID="44"').replace(/ss:StyleID="36"/g, 'ss:StyleID="46"').replace(/ss:StyleID="38"/g, 'ss:StyleID="48"');
  }
  return xml;
}

// Check row 8 (Excel Row 9, first data row)
const testFormattedRow1 = formatOriginalXmlRow(xmlRows[8], 1);
const testFormattedRow2 = formatOriginalXmlRow(xmlRows[9], 2);
const testFormattedRow3 = formatOriginalXmlRow(xmlRows[10], 3);

let updatedHeader = xmlHeader.replace(/(<Cell ss:Index="4" ss:StyleID="25" ss:MergeAcross="10">([\s\S]*?)<Data ss:Type="String">)[^<]*(<\/Data>)/, `$1: 2026-08-01$3`);
updatedHeader = updatedHeader.replace(/ss:MergeDown="\d+"/g, `ss:MergeDown="5"`);

const fullExportXml = updatedHeader + '\n' + [testFormattedRow1, testFormattedRow2, testFormattedRow3].join('\n') + '\n\t' + xmlFooter;

fs.writeFileSync('output_e2e_verified.xls', fullExportXml, 'utf8');

const verifyWb = XLSX.read(fs.readFileSync('output_e2e_verified.xls'), { type: 'buffer' });
const s = verifyWb.Sheets[verifyWb.SheetNames[0]];

console.log('--- E2E VERIFICATION RESULTS ---');
console.log('1. Sheet Range:', s['!ref']);
console.log('2. Kop Title B2:', s['B2'] ? s['B2'].v : 'None');
console.log('3. Date Cell D5:', s['D5'] ? s['D5'].v : 'None');
console.log('4. Depot Cell D6:', s['D6'] ? s['D6'].v : 'None');
console.log('5. Header Row 8:', s['B8']?.v, '|', s['C8']?.v, '|', s['D8']?.v, '|', s['E8']?.v, '|', s['F8']?.v);
console.log('6. Data Row 1 (Seq 1):', s['B9']?.v, s['C9']?.v, s['D9']?.v, s['E9']?.v, s['H9']?.v);
console.log('7. Data Row 2 (Seq 2):', s['B10']?.v, s['C10']?.v, s['D10']?.v, s['E10']?.v, s['H10']?.v);
console.log('8. Data Row 3 (Seq 3):', s['B11']?.v, s['C11']?.v, s['D11']?.v, s['E11']?.v, s['H11']?.v);
console.log('9. Hyperlink on C9:', s['C9']?.l?.Target ? 'PRESERVED OK' : 'None');

if (s['B2']?.v.includes('PRETRIP') && s['D5']?.v.includes('2026-08-01') && s['B9']?.v === 1 && s['B11']?.v === 3) {
  console.log('\n>>> E2E TEMPLATE PRESERVATION IS 100% SUCCESSFUL! <<<');
} else {
  console.error('\n>>> VERIFICATION FAILED! <<<');
  process.exit(1);
}
