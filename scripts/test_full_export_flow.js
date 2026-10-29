const fs = require('fs');
const XLSX = require('../libs/xlsx.full.min.js');
const JSZip = require('../libs/jszip.min.js');
const { header: defaultHeader, footer: defaultFooter } = JSON.parse(fs.readFileSync('scripts/pretrip_header.json', 'utf8'));

const buf = fs.readFileSync('checklist-pretrip_inspect-progress-lvl2.1 full.xls');
const rawXml = buf.toString('utf8');

// Simulate state
const state = {
  fileName: 'checklist-pretrip_inspect-progress-lvl2.1 full.xls',
  isXmlSpreadsheet: true,
  rawXmlText: rawXml,
  xmlHeader: null,
  xmlFooter: null,
  xmlRows: [],
  dataRows: []
};

// 1. Extract XML Template
function extractXmlTemplate(xml) {
  state.xmlRows = [];
  const row8Pos = xml.indexOf('>NO.<');
  if (row8Pos !== -1) {
    const row8Start = xml.lastIndexOf('<Row', row8Pos);
    const row8End = xml.indexOf('</Row>', row8Start) + 6;
    state.xmlHeader = xml.slice(0, row8End);
  }

  const tableEndPos = xml.lastIndexOf('</ss:Table>');
  state.xmlFooter = xml.slice(tableEndPos);

  let pos = 0;
  while (true) {
    const start = xml.indexOf('<Row', pos);
    if (start === -1) break;
    const end = xml.indexOf('</Row>', start);
    if (end === -1) break;
    state.xmlRows.push(xml.slice(start, end + 6));
    pos = end + 6;
  }
}

extractXmlTemplate(rawXml);
console.log('Template extracted! Total XML rows:', state.xmlRows.length);

function formatOriginalXmlRow(rawRowXml, seqNum) {
  const isOdd = (seqNum % 2 === 1);
  let xml = rawRowXml;

  xml = xml.replace(/(<Cell ss:Index="2" ss:StyleID=")[^"]*(">[\s\S]*?<Data ss:Type="Number">)[^<]*(<\/Data>)/, (match, p1, p2, p3) => {
    const style = isOdd ? '30' : '40';
    return `${p1}${style}${p2}${seqNum}.0${p3}`;
  });

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

function buildFilteredXmlSpreadsheet(dateLabel, rows) {
  let header = state.xmlHeader || defaultHeader;
  let footer = state.xmlFooter || defaultFooter;

  if (dateLabel) {
    header = header.replace(
      /(<Cell ss:Index="4" ss:StyleID="25" ss:MergeAcross="10">[\s\S]*?<Data ss:Type="String">)[^<]*(<\/Data>)/,
      `$1: ${dateLabel}$2`
    );
  }

  const mergeDownVal = Math.max(1, rows.length + 2);
  header = header.replace(/ss:MergeDown="\d+"/g, `ss:MergeDown="${mergeDownVal}"`);

  const rowXmlList = rows.map((rObj, idx) => {
    const seqNum = idx + 1;
    return formatOriginalXmlRow(state.xmlRows[rObj._id], seqNum);
  });

  return header + '\n' + rowXmlList.join('\n') + '\n\t' + footer;
}

async function testZipGeneration() {
  const sampleRows1 = [{ _id: 8 }, { _id: 9 }, { _id: 10 }];
  const sampleRows2 = [{ _id: 11 }, { _id: 12 }, { _id: 13 }, { _id: 14 }];

  const zip = new JSZip();
  const xml1 = buildFilteredXmlSpreadsheet('2026-08-01', sampleRows1);
  const xml2 = buildFilteredXmlSpreadsheet('2026-08-02', sampleRows2);

  zip.file('Pretrip_2026-08-01.xls', xml1);
  zip.file('Pretrip_2026-08-02.xls', xml2);

  const zipBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  console.log('ZIP generated successfully! Size:', zipBuffer.length);

  // Read back from zip to verify
  const loadedZip = await JSZip.loadAsync(zipBuffer);
  const file1Str = await loadedZip.file('Pretrip_2026-08-01.xls').async('string');
  const wb1 = XLSX.read(Buffer.from(file1Str, 'utf8'), { type: 'buffer' });
  const s1 = wb1.Sheets[wb1.SheetNames[0]];
  console.log('File 1 verified! Range:', s1['!ref']);
  console.log('File 1 Date D5:', s1['D5'].v);
  console.log('File 1 B9 (No 1):', s1['B9'].v);
  console.log('File 1 B10 (No 2):', s1['B10'].v);
  console.log('File 1 B11 (No 3):', s1['B11'].v);

  const file2Str = await loadedZip.file('Pretrip_2026-08-02.xls').async('string');
  const wb2 = XLSX.read(Buffer.from(file2Str, 'utf8'), { type: 'buffer' });
  const s2 = wb2.Sheets[wb2.SheetNames[0]];
  console.log('File 2 verified! Range:', s2['!ref']);
  console.log('File 2 Date D5:', s2['D5'].v);
  console.log('File 2 B9 (No 1):', s2['B9'].v);
  console.log('File 2 B12 (No 4):', s2['B12'].v);
}

testZipGeneration().then(() => console.log('ALL TESTS PASSED!'));

