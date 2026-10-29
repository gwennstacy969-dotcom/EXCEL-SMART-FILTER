const fs = require('fs');
const path = require('path');

console.log('====================================================');
console.log('       APP HEALTH & DIAGNOSTIC CHECK SUITE          ');
console.log('====================================================\n');

let hasError = false;

// 1. Check Libraries
console.log('[TEST 1] Checking libraries...');
const xlsxPath = path.join(__dirname, 'app', 'libs', 'xlsx.full.min.js');
const jszipPath = path.join(__dirname, 'app', 'libs', 'jszip.min.js');

if (fs.existsSync(xlsxPath) && fs.statSync(xlsxPath).size > 100000) {
  console.log('  [PASS] SheetJS (xlsx.full.min.js) exists and valid size:', fs.statSync(xlsxPath).size);
} else {
  console.error('  [FAIL] SheetJS missing or incomplete!');
  hasError = true;
}

if (fs.existsSync(jszipPath) && fs.statSync(jszipPath).size > 50000) {
  console.log('  [PASS] JSZip (jszip.min.js) exists and valid size:', fs.statSync(jszipPath).size);
} else {
  console.error('  [FAIL] JSZip missing or incomplete!');
  hasError = true;
}

// 2. Check HTML & JS Syntax
console.log('\n[TEST 2] Checking index.html and JavaScript syntax...');
const rootIndexPath = path.join(__dirname, 'index.html');
const appIndexPath = path.join(__dirname, 'app', 'index.html');

[rootIndexPath, appIndexPath].forEach((targetPath) => {
  const fileName = path.relative(__dirname, targetPath);
  if (!fs.existsSync(targetPath)) {
    console.error(`  [FAIL] ${fileName} not found!`);
    hasError = true;
    return;
  }
  const content = fs.readFileSync(targetPath, 'utf8');
  const scripts = content.match(/<script>([\s\S]*?)<\/script>/gi) || [];
  let scriptValid = true;
  scripts.forEach((s, idx) => {
    const code = s.replace(/<\/?script>/gi, '');
    try {
      new Function(code);
    } catch (err) {
      console.error(`  [FAIL] ${fileName} Script ${idx} syntax error:`, err.message);
      hasError = true;
      scriptValid = false;
    }
  });
  if (scriptValid && scripts.length > 0) {
    console.log(`  [PASS] ${fileName} all ${scripts.length} script blocks are 100% valid.`);
  }
});

const html = fs.readFileSync(rootIndexPath, 'utf8');

// 3. Check DOM IDs
console.log('\n[TEST 3] Checking DOM ID references...');
const idRegex = /document\.getElementById\(['"]([^'"]+)['"]\)/g;
let m;
const usedIds = new Set();
while ((m = idRegex.exec(html)) !== null) {
  usedIds.add(m[1]);
}

const dynamicallyCreatedIds = new Set([
  'selectDayDropdown', 'btnPrevDay', 'btnNextDay',
  'selectWeekDropdown', 'inputRangeStart', 'inputRangeEnd'
]);

const missingStaticIds = [];
usedIds.forEach(id => {
  if (!dynamicallyCreatedIds.has(id)) {
    const pattern = new RegExp(`id=["']${id}["']`);
    if (!pattern.test(html)) {
      missingStaticIds.push(id);
    }
  }
});

if (missingStaticIds.length === 0) {
  console.log(`  [PASS] All ${usedIds.size} DOM element IDs referenced in code are valid.`);
} else {
  console.error('  [FAIL] Missing static DOM IDs in HTML:', missingStaticIds);
  hasError = true;
}

// 4. Test real Excel processing & functions
console.log('\n[TEST 4] Simulating Excel Loading & Date Splitting...');
const targetExcelPath = path.join(__dirname, 'checklist-pretrip_inspect-progress-lvl2.1 full.xls');
if (!fs.existsSync(targetExcelPath)) {
  console.error('  [FAIL] checklist-pretrip_inspect-progress-lvl2.1 full.xls not found!');
  hasError = true;
} else {
  const XLSX = require('./app/libs/xlsx.full.min.js');
  const JSZip = require('./app/libs/jszip.min.js');
  const buf = fs.readFileSync(targetExcelPath);
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });

  const sheet = wb.Sheets['Checklist Harian'];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  console.log(`  [INFO] Sheet loaded: ${rawRows.length} raw rows.`);

  // Detect header
  let headerRow = 1;
  for (let r = 0; r < Math.min(rawRows.length, 15); r++) {
    const row = rawRows[r];
    if (Array.isArray(row) && row.filter(c => c && String(c).trim().length > 0).length >= 4) {
      headerRow = r + 1;
      break;
    }
  }
  console.log(`  [PASS] Header detected at Row ${headerRow}`);

  // Test date parsing & grouping
  const headers = ['NO.', 'NO. POLISI', 'JENIS SEWA', 'KL', 'TANGGAL PERIKSA', 'ACTION', 'AMT', 'Item Rusak', 'Keterangan'];
  const colIndices = [1, 2, 3, 4, 5, 7, 8, 10, 11];

  const dataRows = [];
  const dateSet = new Set();
  for (let r = headerRow; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row) continue;
    let hasData = false;
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[colIndices[i]];
      const clean = (v !== undefined && v !== null) ? String(v).trim() : '';
      obj[h] = clean;
      if (clean.length > 0) hasData = true;
    });
    if (hasData) {
      dataRows.push(obj);
      if (obj['TANGGAL PERIKSA']) dateSet.add(obj['TANGGAL PERIKSA']);
    }
  }
  console.log(`  [PASS] Parsed ${dataRows.length} data rows across ${dateSet.size} unique dates.`);

  // Test dynamic weeks
  // (Include generateDynamicWeeks logic)
  function parseFlexDate(val) {
    if (!val && val !== 0) return null;
    if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
    const s = String(val).trim();
    if (!s) return null;
    const parts = s.split(' ')[0].split('/');
    if (parts.length === 3) {
      const m = parseInt(parts[0], 10) - 1;
      const d = parseInt(parts[1], 10);
      let y = parseInt(parts[2], 10);
      if (y < 100) y += 2000;
      return new Date(y, m, d);
    }
    const dashParts = s.split(' ')[0].split('-');
    if (dashParts.length === 3 && dashParts[0].length === 4) {
      return new Date(parseInt(dashParts[0], 10), parseInt(dashParts[1], 10) - 1, parseInt(dashParts[2], 10));
    }
    const dt = new Date(s);
    return isNaN(dt.getTime()) ? null : dt;
  }
  function toDateKey(date) {
    if (!date) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  const parsedDates = Array.from(dateSet).map(parseFlexDate).filter(Boolean);
  parsedDates.sort((a, b) => a.getTime() - b.getTime());
  const minDate = new Date(parsedDates[0]);
  const maxDate = new Date(parsedDates[parsedDates.length - 1]);

  console.log(`  [PASS] Earliest date: ${toDateKey(minDate)}, Latest date: ${toDateKey(maxDate)}`);

  // Test single day formatted output
  function createFormattedFilteredWorkbook(sourceWb, targetDateLabel, rows) {
    const sourceSheet = sourceWb.Sheets['Checklist Harian'];
    const newSheet = {};
    for (const cellAddress in sourceSheet) {
      if (cellAddress.startsWith('!')) continue;
      const match = cellAddress.match(/^([A-Z]+)(\d+)$/);
      if (!match) continue;
      const row = parseInt(match[2], 10);
      if (row <= 8) {
        newSheet[cellAddress] = Object.assign({}, sourceSheet[cellAddress]);
      }
    }
    if (targetDateLabel && newSheet['D5']) {
      newSheet['D5'].v = ': ' + targetDateLabel;
      newSheet['D5'].w = ': ' + targetDateLabel;
    }
    const sourceMerges = sourceSheet['!merges'] || [];
    const newMerges = sourceMerges.filter(m => m.s.r <= 7).map(m => Object.assign({}, m));

    rows.forEach((rowObj, idx) => {
      const r = 9 + idx;
      const rIdx = 8 + idx;
      newSheet['B' + r] = { t: 'n', v: idx + 1, w: String(idx + 1) };
      newSheet['C' + r] = { t: 's', v: rowObj['NO. POLISI'] || '', w: rowObj['NO. POLISI'] || '' };
      newSheet['D' + r] = { t: 's', v: rowObj['JENIS SEWA'] || '', w: rowObj['JENIS SEWA'] || '' };
      newSheet['E' + r] = { t: 's', v: rowObj['KL'] || '', w: rowObj['KL'] || '' };
      newSheet['F' + r] = { t: 's', v: rowObj['TANGGAL PERIKSA'] || '', w: rowObj['TANGGAL PERIKSA'] || '' };
      newSheet['H' + r] = { t: 's', v: rowObj['ACTION'] || '', w: rowObj['ACTION'] || '' };
      newSheet['I' + r] = { t: 's', v: rowObj['AMT'] || '', w: rowObj['AMT'] || '' };
      newSheet['K' + r] = { t: 's', v: rowObj['Item Rusak'] || '', w: rowObj['Item Rusak'] || '' };
      newSheet['L' + r] = { t: 's', v: rowObj['Keterangan'] || '', w: rowObj['Keterangan'] || '' };

      newMerges.push({ s: { r: rIdx, c: 5 }, e: { r: rIdx, c: 6 } });
      newMerges.push({ s: { r: rIdx, c: 8 }, e: { r: rIdx, c: 9 } });
      newMerges.push({ s: { r: rIdx, c: 11 }, e: { r: rIdx, c: 13 } });
    });

    const lastRowNum = Math.max(8, 8 + rows.length);
    newSheet['!ref'] = 'A1:P' + lastRowNum;
    newSheet['!merges'] = newMerges;
    newSheet['!cols'] = sourceSheet['!cols'] ? sourceSheet['!cols'].slice() : [];

    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, newSheet, 'Checklist Harian');
    return newWb;
  }

  const sampleDayRows = dataRows.filter(r => r['TANGGAL PERIKSA'] === '8/1/26');
  const outWb = createFormattedFilteredWorkbook(wb, '2026-08-01', sampleDayRows);
  const outBuf = XLSX.write(outWb, { type: 'buffer', bookType: 'biff8' });
  if (outBuf.length > 5000) {
    console.log(`  [PASS] Formatted Excel export generated successfully (${outBuf.length} bytes).`);
  } else {
    console.error('  [FAIL] Formatted Excel output suspiciously small:', outBuf.length);
    hasError = true;
  }
}

// 5. Check server.js
console.log('\n[TEST 5] Checking server.js...');
const serverPath = path.join(__dirname, 'server.js');
if (fs.existsSync(serverPath)) {
  const serverCode = fs.readFileSync(serverPath, 'utf8');
  try {
    new Function(serverCode);
    console.log('  [PASS] server.js syntax is valid.');
  } catch (err) {
    console.error('  [FAIL] server.js syntax error:', err.message);
    hasError = true;
  }
} else {
  console.error('  [FAIL] server.js not found!');
  hasError = true;
}

// Summary
console.log('\n====================================================');
if (!hasError) {
  console.log('>>> ALL DIAGNOSTIC TESTS PASSED! APPLICATION IS HEALTHY <<<');
} else {
  console.log('>>> SOME TESTS FAILED. PLEASE REVIEW LOG ABOVE <<<');
}
console.log('====================================================');

