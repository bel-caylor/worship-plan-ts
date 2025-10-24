// scripts/wrap-css.js
const fs = require('fs');

const inPath = 'dist/html/styles.css';
const outPath = 'dist/html/styles.html';

if (!fs.existsSync(inPath)) {
  console.log('[wrap-css] Missing', inPath, '- nothing to do yet.');
  process.exit(0);
}

const css = fs.readFileSync(inPath, 'utf8');
fs.writeFileSync(outPath, `<style>${css}</style>`);
console.log('[wrap-css] Wrote', outPath);
