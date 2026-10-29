const fs = require('fs');

function updateHtmlFile(filePath) {
  let html = fs.readFileSync(filePath, 'utf8');

  // 1. Add <script src="libs/pretrip_template.js"></script> if not present
  if (!html.includes('pretrip_template.js')) {
    html = html.replace(
      '<script src="libs/jszip.min.js"></script>',
      '<script src="libs/jszip.min.js"></script>\n  <!-- Pretrip Excel XML Template for 100% Identical Kop & Styling -->\n  <script src="libs/pretrip_template.js"></script>'
    );
  }

  // 2. Add XML state fields to state object
  if (!html.includes('isXmlSpreadsheet:')) {
    html = html.replace(
      "selectedRows: new Set()",
      "selectedRows: new Set(),\n      // XML Spreadsheet 2003 Template State (Format Asli BIRT)\n      isXmlSpreadsheet: false,\n      rawXmlText: null,\n      xmlHeader: null,\n      xmlFooter: null,\n      xmlRows: []"
    );
  }

  // 3. Add checkAndParseXmlSpreadsheet and extractXmlTemplate helper functions before loadPretripInspectionFile
  if (!html.includes('function checkAndParseXmlSpreadsheet')) {
    const xmlParseHelpers = `
    // =========================================================================
    // XML SPREADSHEET 2003 TEMPLATE PARSER (FORMAT ASLI BIRT / SMART LEADERS)
    // =========================================================================
    function checkAndParseXmlSpreadsheet(uint8Arr) {
      try {
        const preview = new TextDecoder('utf-8').decode(uint8Arr.slice(0, 1500));
        if (preview.includes('<?xml') && (preview.includes('spreadsheet') || preview.includes('Workbook'))) {
          const fullText = new TextDecoder('utf-8').decode(uint8Arr);
          state.isXmlSpreadsheet = true;
          state.rawXmlText = fullText;
          extractXmlTemplate(fullText);
          return true;
        }
      } catch (e) {
        console.warn('Gagal membaca format XML spreadsheet:', e);
      }
      state.isXmlSpreadsheet = false;
      state.rawXmlText = null;
      state.xmlHeader = null;
      state.xmlFooter = null;
      state.xmlRows = [];
      return false;
    }

    function extractXmlTemplate(xml) {
      state.xmlRows = [];
      
      // 1. Temukan header Kop sampai baris 8 (yang memuat '>NO.<')
      const row8Pos = xml.indexOf('>NO.<');
      if (row8Pos !== -1) {
        const row8Start = xml.lastIndexOf('<Row', row8Pos);
        const row8End = xml.indexOf('</Row>', row8Start) + 6;
        state.xmlHeader = xml.slice(0, row8End);
      } else {
        let pos = 0;
        let rCount = 0;
        let r8End = 0;
        while (rCount < 8) {
          const start = xml.indexOf('<Row', pos);
          if (start === -1) break;
          const end = xml.indexOf('</Row>', start);
          if (end === -1) break;
          rCount++;
          r8End = end + 6;
          pos = r8End;
        }
        state.xmlHeader = xml.slice(0, r8End);
      }

      // 2. Temukan Footer mulai dari </ss:Table>
      const tableEndPos = xml.lastIndexOf('</ss:Table>');
      if (tableEndPos !== -1) {
        state.xmlFooter = xml.slice(tableEndPos);
      } else {
        state.xmlFooter = '\\t</ss:Table>\\n</Worksheet>\\n</Workbook>';
      }

      // 3. Ekstrak seluruh elemen baris XML untuk referensi cepat
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
`;
    html = html.replace(
      '// LOAD PRETRIP FILE & PARSE HEADERS',
      xmlParseHelpers + '\n    // LOAD PRETRIP FILE & PARSE HEADERS'
    );
  }

  // 4. Hook checkAndParseXmlSpreadsheet into loadPretripInspectionFile and handleFileUpload
  if (html.includes('state.workbook = XLSX.read(data, { type: \'array\', cellDates: true });')) {
    html = html.replace(
      /const data = new Uint8Array\(buf\);\s*state\.fileName = fileName;\s*state\.workbook = XLSX\.read\(data/,
      "const data = new Uint8Array(buf);\n          checkAndParseXmlSpreadsheet(data);\n          state.fileName = fileName;\n          state.workbook = XLSX.read(data"
    );
    html = html.replace(
      /const data = new Uint8Array\(e\.target\.result\);\s*state\.workbook = XLSX\.read\(data/,
      "const data = new Uint8Array(e.target.result);\n          checkAndParseXmlSpreadsheet(data);\n          state.workbook = XLSX.read(data"
    );
  }

  // 5. In parseWorkbookSheet(), extract _href for nopol
  if (!html.includes('rowObj._href =')) {
    html = html.replace(
      "rowObj[h] = clean;\n          if (clean.length > 0) hasData = true;",
      "rowObj[h] = clean;\n          if (clean.length > 0) hasData = true;\n          // Simpan hyperlink nopol jika ada\n          const cAddr = XLSX.utils.encode_cell({ r, c: realCol });\n          if (sheet[cAddr] && sheet[cAddr].l && sheet[cAddr].l.Target) {\n            rowObj._href = sheet[cAddr].l.Target;\n          }"
    );
  }

  // 6. Replace Export Section with complete XML Spreadsheet & Template Generator
  const oldExportSectionStart = html.indexOf('function createFormattedFilteredWorkbook');
  const oldExportSectionEnd = html.indexOf('function exportToCSV');

  if (oldExportSectionStart !== -1 && oldExportSectionEnd !== -1) {
    const newExportSection = `function isPretripInspection() {
      if (state.isXmlSpreadsheet && state.xmlHeader && state.xmlHeader.includes('PRETRIP')) return true;
      if (state.workbook && state.currentSheet) {
        const sheet = state.workbook.Sheets[state.currentSheet];
        if (sheet && sheet['B2'] && String(sheet['B2'].v || '').toUpperCase().includes('PRETRIP')) return true;
      }
      if (state.headers && state.headers.includes('NO. POLISI') && state.headers.includes('JENIS SEWA')) return true;
      return false;
    }

    function escapeXml(str) {
      if (str === null || str === undefined) return '';
      return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    }

    function formatOriginalXmlRow(rawRowXml, seqNum) {
      const isOdd = (seqNum % 2 === 1);
      let xml = rawRowXml;

      // Update nomor urut di Cell Index="2"
      xml = xml.replace(/(<Cell ss:Index="2" ss:StyleID=")[^"]*(">[\s\S]*?<Data ss:Type="Number">)[^<]*(<\/Data>)/, (match, p1, p2, p3) => {
        const style = isOdd ? '30' : '40';
        return \`\${p1}\${style}\${p2}\${seqNum}.0\${p3}\`;
      });

      // Normalisasi zebra styling (selang-seling baris ganjil & genap rapi)
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

    function synthesizePretripXmlRow(rowObj, seqNum) {
      const isOdd = (seqNum % 2 === 1);
      const numStyle = isOdd ? '30' : '40';
      const nopolStyle = isOdd ? '32' : '42';
      const generalStyle = isOdd ? '34' : '44';
      const dateStyle = isOdd ? '36' : '46';
      const textStyle = isOdd ? '38' : '48';

      const nopol = rowObj['NO. POLISI'] || '';
      const jenis = rowObj['JENIS SEWA'] || '';
      const kl = rowObj['KL'] || '';
      const tgl = rowObj['TANGGAL PERIKSA'] || '';
      const action = rowObj['ACTION'] || '';
      const amt = rowObj['AMT'] || '';
      const itemRusak = rowObj['Item Rusak'] || rowObj['ITEM RUSAK'] || '';
      const ket = rowObj['Keterangan'] || rowObj['KETERANGAN'] || '';
      const href = rowObj._href ? \` ss:HRef="\${escapeXml(rowObj._href)}"\` : '';

      const dateType = (typeof tgl === 'string' && tgl.includes('T')) ? 'DateTime' : 'String';

      return \`\\t\\t<Row ss:AutoFitHeight="1">
\\t\\t\\t<Cell ss:Index="2" ss:StyleID="\${numStyle}">
\\t\\t\\t\\t<Data ss:Type="Number">\${seqNum}.0</Data>
\\t\\t\\t</Cell>
\\t\\t\\t<Cell ss:Index="3" ss:StyleID="\${nopolStyle}"\${href}>
\\t\\t\\t\\t<Data ss:Type="String">\${escapeXml(nopol)}</Data>
\\t\\t\\t</Cell>
\\t\\t\\t<Cell ss:Index="4" ss:StyleID="\${generalStyle}">
\\t\\t\\t\\t<Data ss:Type="String">\${escapeXml(jenis)}</Data>
\\t\\t\\t</Cell>
\\t\\t\\t<Cell ss:Index="5" ss:StyleID="\${generalStyle}">
\\t\\t\\t\\t<Data ss:Type="String">\${escapeXml(kl)}</Data>
\\t\\t\\t</Cell>
\\t\\t\\t<Cell ss:Index="6" ss:StyleID="\${dateStyle}" ss:MergeAcross="1">
\\t\\t\\t\\t<Data ss:Type="\${dateType}">\${escapeXml(tgl)}</Data>
\\t\\t\\t</Cell>
\\t\\t\\t<Cell ss:Index="8" ss:StyleID="\${generalStyle}">
\\t\\t\\t\\t<Data ss:Type="String">\${escapeXml(action)}</Data>
\\t\\t\\t</Cell>
\\t\\t\\t<Cell ss:Index="9" ss:StyleID="\${textStyle}" ss:MergeAcross="1">
\\t\\t\\t\\t<Data ss:Type="String">\${escapeXml(amt)}</Data>
\\t\\t\\t</Cell>
\\t\\t\\t<Cell ss:Index="11" ss:StyleID="\${textStyle}">
\\t\\t\\t\\t<Data ss:Type="String">\${escapeXml(itemRusak)}</Data>
\\t\\t\\t</Cell>
\\t\\t\\t<Cell ss:Index="12" ss:StyleID="\${textStyle}" ss:MergeAcross="2">
\\t\\t\\t\\t<Data ss:Type="String">\${escapeXml(ket)}</Data>
\\t\\t\\t</Cell>
\\t\\t</Row>\`;
    }

    function buildFilteredXmlSpreadsheet(dateLabel, rows) {
      let header = state.xmlHeader;
      let footer = state.xmlFooter;

      if (!header && window.DEFAULT_PRETRIP_TEMPLATE) {
        header = window.DEFAULT_PRETRIP_TEMPLATE.header;
        footer = window.DEFAULT_PRETRIP_TEMPLATE.footer;
      }

      if (!header) return null;

      // 1. Update nilai tanggal di Kop (Baris 5: DATE)
      if (dateLabel) {
        header = header.replace(
          /(<Cell ss:Index="4" ss:StyleID="25" ss:MergeAcross="10">[\s\S]*?<Data ss:Type="String">)[^<]*(<\/Data>)/,
          \`$1: \${escapeXml(dateLabel)}$2\`
        );
      }

      // 2. Update border garis vertikal (MergeDown) agar membentang pas sesuai jumlah baris
      const mergeDownVal = Math.max(1, rows.length + 2);
      header = header.replace(/ss:MergeDown="\\d+"/g, \`ss:MergeDown="\${mergeDownVal}"\`);

      // 3. Format atau sintesis seluruh baris data terfilter
      const rowXmlList = rows.map((rowObj, idx) => {
        const seqNum = idx + 1;
        if (state.xmlRows && rowObj._id !== undefined && state.xmlRows[rowObj._id]) {
          return formatOriginalXmlRow(state.xmlRows[rowObj._id], seqNum);
        } else {
          return synthesizePretripXmlRow(rowObj, seqNum);
        }
      });

      return header + '\\n' + rowXmlList.join('\\n') + '\\n\\t' + footer;
    }

    function triggerFileDownload(blobOrString, fileName, mimeType = 'application/vnd.ms-excel;charset=utf-8') {
      const blob = (typeof blobOrString === 'string') 
        ? new Blob([blobOrString], { type: mimeType }) 
        : blobOrString;
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function createFormattedFilteredWorkbook(sourceWb, targetDateLabel, rows) {
      const sourceSheetName = state.currentSheet || 'Checklist Harian';
      const sourceSheet = sourceWb.Sheets[sourceSheetName] || sourceWb.Sheets[sourceWb.SheetNames[0]];
      const newSheet = {};

      const isPretripTemplate = isPretripInspection();

      if (isPretripTemplate) {
        for (const cellAddress in sourceSheet) {
          if (cellAddress.startsWith('!')) continue;
          const match = cellAddress.match(/^([A-Z]+)(\\d+)$/);
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
      } else {
        const headerRowLimit = state.headerRow || 1;
        for (const cellAddress in sourceSheet) {
          if (cellAddress.startsWith('!')) continue;
          const match = cellAddress.match(/^([A-Z]+)(\\d+)$/);
          if (!match) continue;
          const row = parseInt(match[2], 10);
          if (row <= headerRowLimit) {
            newSheet[cellAddress] = Object.assign({}, sourceSheet[cellAddress]);
          }
        }

        const sourceMerges = sourceSheet['!merges'] || [];
        const newMerges = sourceMerges.filter(m => m.s.r < headerRowLimit).map(m => Object.assign({}, m));

        rows.forEach((rowObj, idx) => {
          const r = headerRowLimit + 1 + idx;
          state.headers.forEach((h, colIdx) => {
            const realCol = state.colIndices[colIdx];
            const cellAddr = XLSX.utils.encode_cell({ r: r - 1, c: realCol });
            const val = rowObj[h] || '';
            newSheet[cellAddr] = { t: 's', v: val, w: String(val) };
          });
        });

        const maxCol = state.colIndices && state.colIndices.length > 0 ? Math.max(...state.colIndices) : 10;
        const lastRowNum = headerRowLimit + rows.length;
        newSheet['!ref'] = 'A1:' + XLSX.utils.encode_cell({ r: lastRowNum, c: maxCol });
        newSheet['!merges'] = newMerges;
        newSheet['!cols'] = sourceSheet['!cols'] ? sourceSheet['!cols'].slice() : [];
      }

      const newWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWb, newSheet, sourceSheetName);
      return newWb;
    }

    // 1. Ekspor Single Terfilter (Format Asli 100% Identik)
    function exportFormattedWorkbookSingle() {
      if (!state.filteredRows || state.filteredRows.length === 0) {
        alert('Tidak ada baris data yang cocok untuk diekspor.');
        return;
      }

      let dateLabel = '';
      if (state.dateMode === 'DAY') dateLabel = state.dateDayVal;
      else if (state.dateMode === 'WEEK') {
        const w = (state.dynamicWeeks && state.dynamicWeeks[state.dateWeekIdx]) || {};
        dateLabel = w.label || '';
      }
      else if (state.dateMode === 'RANGE') dateLabel = \`\${state.dateRangeStart} s/d \${state.dateRangeEnd}\`;
      else dateLabel = 'Semua Periode';

      const baseName = state.fileName ? state.fileName.replace(/\\.[^/.]+$/, "") : 'Checklist_Pretrip';
      const dateTag = state.dateDayVal ? state.dateDayVal.replace(/\\//g, '-') : new Date().toISOString().slice(0, 10);

      // Jika format Pretrip / XML Template, ekspor dengan format asli BIRT (Kop, style, font, warna, border)
      if (isPretripInspection()) {
        const xmlStr = buildFilteredXmlSpreadsheet(dateLabel, state.filteredRows);
        if (xmlStr) {
          const outFileName = \`\${baseName}_Filtered_\${dateTag}.xls\`;
          triggerFileDownload(xmlStr, outFileName, 'application/vnd.ms-excel;charset=utf-8');
          return;
        }
      }

      // Fallback SheetJS untuk file Excel umum lainnya
      const newWb = createFormattedFilteredWorkbook(state.workbook, dateLabel, state.filteredRows);
      const isXlsx = state.fileName && state.fileName.toLowerCase().endsWith('.xlsx');
      const ext = isXlsx ? 'xlsx' : 'xls';
      const bType = isXlsx ? 'xlsx' : 'biff8';
      const outFileName = \`\${baseName}_Filtered_\${dateTag}.\${ext}\`;

      XLSX.writeFile(newWb, outFileName, { bookType: bType });
    }

    // 2. Batch Split Per Hari -> Download ZIP (Format Asli 100% Identik)
    async function exportSplitDailyZip() {
      if (!state.dataRows || state.dataRows.length === 0) return;

      const btn = document.getElementById('btnSplitDailyZip');
      const originalText = btn.innerHTML;
      btn.innerHTML = \`<span>⏳</span><span>Membuat ZIP Harian...</span>\`;
      btn.disabled = true;

      try {
        const rowsByDate = new Map();

        state.dataRows.forEach(row => {
          for (const col in state.activeFilters) {
            const f = state.activeFilters[col];
            const rawVal = row[col];
            const isBlank = rawVal === undefined || rawVal === null || String(rawVal).trim() === '';
            const val = isBlank ? '(Kosong)' : String(rawVal).trim();

            if (f.mode === 'VALUES') {
              if (!f.selectedValues.has(val)) return;
            } else if (f.mode === 'TEXT') {
              const vLow = val.toLowerCase();
              const kwLow = f.keyword.toLowerCase();
              if (f.op === 'CONTAINS' && !vLow.includes(kwLow)) return;
            }
          }

          const dVal = row[state.dateColumn];
          if (!dVal) return;
          if (!rowsByDate.has(dVal)) rowsByDate.set(dVal, []);
          rowsByDate.get(dVal).push(row);
        });

        if (rowsByDate.size === 0) {
          alert('Tidak ada data yang cocok dengan kriteria filter.');
          return;
        }

        const zip = new JSZip();
        const baseName = state.fileName ? state.fileName.replace(/\\.[^/.]+$/, "") : 'Checklist_Pretrip';
        const isPretrip = isPretripInspection();
        const isXlsx = state.fileName && state.fileName.toLowerCase().endsWith('.xlsx');
        const ext = isPretrip ? 'xls' : (isXlsx ? 'xlsx' : 'xls');
        const bType = isXlsx ? 'xlsx' : 'biff8';

        rowsByDate.forEach((dayRows, dStr) => {
          const parts = dStr.split('/');
          let iso = dStr;
          if (parts.length === 3) {
            const m = parts[0].padStart(2, '0');
            const d = parts[1].padStart(2, '0');
            let y = parts[2];
            if (y.length === 2) y = '20' + y;
            iso = \`\${y}-\${m}-\${d}\`;
          }

          if (isPretrip) {
            const xmlStr = buildFilteredXmlSpreadsheet(iso, dayRows);
            if (xmlStr) {
              zip.file(\`\${baseName}_\${iso}.\${ext}\`, xmlStr);
              return;
            }
          }

          const dayWb = createFormattedFilteredWorkbook(state.workbook, iso, dayRows);
          const dayBytes = XLSX.write(dayWb, { type: 'array', bookType: bType });
          zip.file(\`\${baseName}_\${iso}.\${ext}\`, dayBytes);
        });

        const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
        triggerFileDownload(zipBlob, \`\${baseName}_Harian_ZIP.zip\`, 'application/zip');

      } catch (err) {
        console.error('Gagal membuat ZIP harian:', err);
        alert('Terjadi kesalahan saat membuat ZIP harian: ' + err.message);
      } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
      }
    }

    // 3. Batch Split Per Minggu -> Download ZIP (Format Asli 100% Identik)
    async function exportSplitWeeklyZip() {
      if (!state.dataRows || state.dataRows.length === 0) return;

      const btn = document.getElementById('btnSplitWeeklyZip');
      const originalText = btn.innerHTML;
      btn.innerHTML = \`<span>⏳</span><span>Membuat ZIP Mingguan...</span>\`;
      btn.disabled = true;

      try {
        const zip = new JSZip();
        const baseName = state.fileName ? state.fileName.replace(/\\.[^/.]+$/, "") : 'Checklist_Pretrip';
        const isPretrip = isPretripInspection();
        const isXlsx = state.fileName && state.fileName.toLowerCase().endsWith('.xlsx');
        const ext = isPretrip ? 'xls' : (isXlsx ? 'xlsx' : 'xls');
        const bType = isXlsx ? 'xlsx' : 'biff8';
        const weeksList = state.dynamicWeeks && state.dynamicWeeks.length > 0 ? state.dynamicWeeks : [];

        weeksList.forEach((w) => {
          const weekRows = state.dataRows.filter(row => {
            for (const col in state.activeFilters) {
              const f = state.activeFilters[col];
              const rawVal = row[col];
              const isBlank = rawVal === undefined || rawVal === null || String(rawVal).trim() === '';
              const val = isBlank ? '(Kosong)' : String(rawVal).trim();
              if (f.mode === 'VALUES' && !f.selectedValues.has(val)) return false;
            }

            const rowD = parseFlexDate(row[state.dateColumn]);
            if (!rowD) return false;
            const rowIso = toDateKey(rowD);
            return rowIso >= w.start && rowIso <= w.end;
          });

          if (weekRows.length > 0) {
            if (isPretrip) {
              const xmlStr = buildFilteredXmlSpreadsheet(w.label, weekRows);
              if (xmlStr) {
                zip.file(\`\${baseName}_\${w.name.replace(/\\s+/g, '_')}_\${w.start}_sd_\${w.end}.\${ext}\`, xmlStr);
                return;
              }
            }

            const weekWb = createFormattedFilteredWorkbook(state.workbook, w.label, weekRows);
            const weekBytes = XLSX.write(weekWb, { type: 'array', bookType: bType });
            zip.file(\`\${baseName}_\${w.name.replace(/\\s+/g, '_')}_\${w.start}_sd_\${w.end}.\${ext}\`, weekBytes);
          }
        });

        const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
        triggerFileDownload(zipBlob, \`\${baseName}_Mingguan_ZIP.zip\`, 'application/zip');

      } catch (err) {
        console.error('Gagal membuat ZIP mingguan:', err);
        alert('Terjadi kesalahan saat membuat ZIP mingguan: ' + err.message);
      } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
      }
    }

    `;

    html = html.slice(0, oldExportSectionStart) + newExportSection + html.slice(oldExportSectionEnd);
  }

  fs.writeFileSync(filePath, html, 'utf8');
  console.log(`Updated ${filePath} successfully!`);
}

updateHtmlFile('index.html');
if (fs.existsSync('app/index.html')) {
  updateHtmlFile('app/index.html');
}

