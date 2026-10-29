const fs = require('fs');

const rawXml = fs.readFileSync('checklist-pretrip_inspect-progress-lvl2.1 full.xls', 'utf8');

// Header part before row 9
const row8Pos = rawXml.indexOf('>NO.<');
const row8Start = rawXml.lastIndexOf('<Row', row8Pos);
const row8End = rawXml.indexOf('</Row>', row8Start) + 6;
const headerPart = rawXml.slice(0, row8End);

// Footer part after last row
const tableEndPos = rawXml.lastIndexOf('</ss:Table>');
const footerPart = rawXml.slice(tableEndPos);

// Extract all rows
const rows = [];
let pos = 0;
while (true) {
  const start = rawXml.indexOf('<Row', pos);
  if (start === -1) break;
  const end = rawXml.indexOf('</Row>', start);
  if (end === -1) break;
  rows.push(rawXml.slice(start, end + 6));
  pos = end + 6;
}

// Suppose filtered rows are row 8, 10, 15 (indices in rows array)
const filteredRowIndices = [8, 10, 15];

function formatXmlRow(rawRowXml, seqNum) {
  const isOdd = (seqNum % 2 === 1);
  let xml = rawRowXml;

  // Replace sequential number in Cell Index="2"
  xml = xml.replace(/(<Cell ss:Index="2" ss:StyleID=")[^"]*(">[\s\S]*?<Data ss:Type="Number">)[^<]*(<\/Data>)/, (match, p1, p2, p3) => {
    const style = isOdd ? '30' : '40';
    return `${p1}${style}${p2}${seqNum}.0${p3}`;
  });

  // Normalize zebra styling
  if (isOdd) {
    xml = xml
      .replace(/ss:StyleID="40"/g, 'ss:StyleID="30"')
      .replace(/ss:StyleID="42"/g, 'ss:StyleID="32"')
      .replace(/ss:StyleID="44"/g, 'ss:StyleID="34"')
      .replace(/ss:StyleID="46"/g, 'ss:StyleID="36"')
      .replace(/ss:StyleID="48"/g, 'ss:StyleID="38"');
  } else {
    xml = xml
      .replace(/ss:StyleID="30"/g, 'ss:StyleID="40"')
      .replace(/ss:StyleID="32"/g, 'ss:StyleID="42"')
      .replace(/ss:StyleID="34"/g, 'ss:StyleID="44"')
      .replace(/ss:StyleID="36"/g, 'ss:StyleID="46"')
      .replace(/ss:StyleID="38"/g, 'ss:StyleID="48"');
  }

  return xml;
}

const formattedRows = filteredRowIndices.map((rIdx, i) => formatXmlRow(rows[rIdx], i + 1));

// Update header
let updatedHeader = headerPart;
const dateLabel = '01/08/2026';
updatedHeader = updatedHeader.replace(/(<Cell ss:Index="4" ss:StyleID="25" ss:MergeAcross="10">[\s\S]*?<Data ss:Type="String">)[^<]*(<\/Data>)/, `$1: ${dateLabel}$2`);
const mergeDownVal = Math.max(1, filteredRowIndices.length + 2);
updatedHeader = updatedHeader.replace(/ss:MergeDown="\d+"/g, `ss:MergeDown="${mergeDownVal}"`);

const finalXml = updatedHeader + '\n' + formattedRows.join('\n') + '\n\t' + footerPart;

fs.writeFileSync('output_filtered_test.xls', finalXml, 'utf8');
console.log('Saved output_filtered_test.xls successfully!');

const XLSX = require('../libs/xlsx.full.min.js');
const wb = XLSX.read(fs.readFileSync('output_filtered_test.xls'), { type: 'buffer' });
const s = wb.Sheets[wb.SheetNames[0]];
console.log('Verified ref:', s['!ref']);
console.log('Row 1 (B9):', s['B9'].v, s['C9'].v, s['D5'].v);
console.log('Row 2 (B10):', s['B10'].v, s['C10'].v);
console.log('Row 3 (B11):', s['B11'].v, s['C11'].v);

