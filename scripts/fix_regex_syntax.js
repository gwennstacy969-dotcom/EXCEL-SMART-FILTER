const fs = require('fs');

function fixFile(filePath) {
  let content = fs.readFileSync(filePath, 'utf8');

  // Fix 1: formatOriginalXmlRow regex
  content = content.replace(
    'xml = xml.replace(/(<Cell ss:Index="2" ss:StyleID=")[^"]*(">[sS]*?<Data ss:Type="Number">)[^<]*(</Data>)/',
    'xml = xml.replace(/(<Cell ss:Index="2" ss:StyleID=")[^"]*(">([\\s\\S]*?)<Data ss:Type="Number">)[^<]*(<\\/Data>)/'
  );

  // Fix 2: date replace regex in buildFilteredXmlSpreadsheet
  content = content.replace(
    '/(<Cell ss:Index="4" ss:StyleID="25" ss:MergeAcross="10">[sS]*?<Data ss:Type="String">)[^<]*(</Data>)/',
    '/(<Cell ss:Index="4" ss:StyleID="25" ss:MergeAcross="10">([\\s\\S]*?)<Data ss:Type="String">)[^<]*(<\\/Data>)/'
  );

  fs.writeFileSync(filePath, content, 'utf8');
  console.log('Fixed', filePath);
}

fixFile('index.html');
if (fs.existsSync('app/index.html')) {
  fixFile('app/index.html');
}

