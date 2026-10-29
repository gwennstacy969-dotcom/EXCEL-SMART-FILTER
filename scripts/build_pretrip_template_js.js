const fs = require('fs');
const { header, footer } = JSON.parse(fs.readFileSync('scripts/pretrip_header.json', 'utf8'));

const js = `// Pretrip Inspection BIRT Excel Template (Kop PTI, TBBM/DEPOT MAOS, Styles, Borders, Column Widths)
window.DEFAULT_PRETRIP_TEMPLATE = {
  header: ${JSON.stringify(header)},
  footer: ${JSON.stringify(footer)}
};
`;

fs.writeFileSync('libs/pretrip_template.js', js, 'utf8');
if (fs.existsSync('app/libs')) {
  fs.writeFileSync('app/libs/pretrip_template.js', js, 'utf8');
}
console.log('libs/pretrip_template.js written successfully! Size:', fs.statSync('libs/pretrip_template.js').size);

