// Rebuild the offline HTML; third-party scripts remain separate in vendor/.
const fs = require('node:fs');
const path = require('node:path');
const root = __dirname;
let html = fs.readFileSync(path.join(root, 'src/template.html'), 'utf8');
for (const [marker, filename] of Object.entries({ STYLES: 'src/style.css', SHEET: 'src/sheet-reader.js', PLOT: 'src/plot.js', CORE: 'src/core.js', STORE: 'src/config-store.js', APP: 'src/app.js' })) {
  let code = fs.readFileSync(path.join(root, filename), 'utf8');
  if (marker !== 'STYLES') code = code.replace(/<\/script/gi, '<\\/script').replace(/^\/\/# sourceMappingURL=.*$/gm, '');
  html = html.replace(`/*__${marker}__*/`, () => code);
}
fs.writeFileSync(path.join(root, 'index.html'), html, 'utf8');
console.log('Built index.html (' + Buffer.byteLength(html).toLocaleString() + ' bytes)');
