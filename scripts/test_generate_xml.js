const fs = require('fs');
const XLSX = require('../libs/xlsx.full.min.js');

const rawXml = fs.readFileSync('checklist-pretrip_inspect-progress-lvl2.1 full.xls', 'utf8');

// Find header part up to Row 8 (before Row 9)
const row8Pos = rawXml.indexOf('>NO.<');
const row8Start = rawXml.lastIndexOf('<Row', row8Pos);
const row8End = rawXml.indexOf('</Row>', row8Start) + 6;

const xmlHeader = rawXml.slice(0, row8End);

// Find footer part from </Table>
const tableEndPos = rawXml.lastIndexOf('</ss:Table>');
const xmlFooter = rawXml.slice(tableEndPos);

console.log('xmlHeader length:', xmlHeader.length);
console.log('xmlFooter length:', xmlFooter.length);

// Sample rows to test
const sampleRows = [
  {
    no: 1,
    nopol: 'E9556YC',
    jenisSewa: 'allin',
    kl: '24.0',
    tglPeriksa: '2026-08-01T07:03:31.535',
    action: 'GO',
    amt: 'Kuat Sutrisno',
    itemRusak: '',
    keterangan: 'go',
    href: 'https://bigdata2.smart-leaders.com:443/WebViewer/frameset?report=test1'
  },
  {
    no: 2,
    nopol: 'B9810TEI',
    jenisSewa: 'allin',
    kl: '24.0',
    tglPeriksa: '2026-08-01T07:07:14.577',
    action: 'GO',
    amt: 'Noaf Dwi Prakoso',
    itemRusak: '',
    keterangan: 'ok',
    href: 'https://bigdata2.smart-leaders.com:443/WebViewer/frameset?report=test2'
  }
];

function escapeXml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildPretripXmlRows(rows) {
  let xml = '';
  rows.forEach((r, idx) => {
    const isOdd = (idx % 2 === 0); // 1st row (idx 0) is odd in 1-based numbering
    const numStyle = isOdd ? '30' : '40';
    const nopolStyle = isOdd ? '32' : '42';
    const generalStyle = isOdd ? '34' : '44';
    const dateStyle = isOdd ? '36' : '46';
    const textStyle = isOdd ? '38' : '48';

    const hrefAttr = r.href ? ` ss:HRef="${escapeXml(r.href)}"` : '';

    xml += `\n\t\t<Row ss:AutoFitHeight="1">
\t\t\t<Cell ss:Index="2" ss:StyleID="${numStyle}">
\t\t\t\t<Data ss:Type="Number">${idx + 1}.0</Data>
\t\t\t</Cell>
\t\t\t<Cell ss:Index="3" ss:StyleID="${nopolStyle}"${hrefAttr}>
\t\t\t\t<Data ss:Type="String">${escapeXml(r.nopol)}</Data>
\t\t\t</Cell>
\t\t\t<Cell ss:Index="4" ss:StyleID="${generalStyle}">
\t\t\t\t<Data ss:Type="String">${escapeXml(r.jenisSewa)}</Data>
\t\t\t</Cell>
\t\t\t<Cell ss:Index="5" ss:StyleID="${generalStyle}">
\t\t\t\t<Data ss:Type="String">${escapeXml(r.kl)}</Data>
\t\t\t</Cell>
\t\t\t<Cell ss:Index="6" ss:StyleID="${dateStyle}" ss:MergeAcross="1">
\t\t\t\t<Data ss:Type="${r.tglPeriksa && r.tglPeriksa.includes('T') ? 'DateTime' : 'String'}">${escapeXml(r.tglPeriksa)}</Data>
\t\t\t</Cell>
\t\t\t<Cell ss:Index="8" ss:StyleID="${generalStyle}">
\t\t\t\t<Data ss:Type="String">${escapeXml(r.action)}</Data>
\t\t\t</Cell>
\t\t\t<Cell ss:Index="9" ss:StyleID="${textStyle}" ss:MergeAcross="1">
\t\t\t\t<Data ss:Type="String">${escapeXml(r.amt)}</Data>
\t\t\t</Cell>
\t\t\t<Cell ss:Index="11" ss:StyleID="${textStyle}">
\t\t\t\t<Data ss:Type="String">${escapeXml(r.itemRusak)}</Data>
\t\t\t</Cell>
\t\t\t<Cell ss:Index="12" ss:StyleID="${textStyle}" ss:MergeAcross="2">
\t\t\t\t<Data ss:Type="String">${escapeXml(r.keterangan)}</Data>
\t\t\t</Cell>
\t\t</Row>`;
  });
  return xml;
}

// Update Date in header
let customHeader = xmlHeader;
const targetDate = '2026-08-01';
// Update MergeDown in Row 5
const totalRows = sampleRows.length;
const mergeDownVal = Math.max(1, totalRows + 2);
customHeader = customHeader.replace(/ss:MergeDown="\d+"/g, `ss:MergeDown="${mergeDownVal}"`);
customHeader = customHeader.replace(/<Data ss:Type="String">: [^<]*<\/Data>/, `<Data ss:Type="String">: ${targetDate}</Data>`);

const fullXml = customHeader + buildPretripXmlRows(sampleRows) + '\n\t' + xmlFooter;

fs.writeFileSync('output_test.xls', fullXml, 'utf8');
console.log('Saved output_test.xls! Size:', fullXml.length);

// Verify with SheetJS
const verifyWb = XLSX.read(fs.readFileSync('output_test.xls'), { type: 'buffer' });
const verifySheet = verifyWb.Sheets[verifyWb.SheetNames[0]];
console.log('SheetJS read verified! Range:', verifySheet['!ref']);
console.log('Cell B9:', verifySheet['B9']);
console.log('Cell C9:', verifySheet['C9']);
console.log('Cell D5:', verifySheet['D5']);
