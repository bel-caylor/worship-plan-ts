// scripts/wrap-css.js
const fs = require('fs');
const cssPath = 'dist/styles.css';
const outPath = 'dist/styles.html';

if (!fs.existsSync(cssPath)) process.exit(0);
const css = fs.readFileSync(cssPath, 'utf8');
fs.writeFileSync(outPath, `<style>${css}</style>`);
console.log('Wrapped CSS ->', outPath);
