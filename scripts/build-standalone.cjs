/* eslint-disable no-console */
const fs = require('fs');
const path = require('path');

const ROOT = process.cwd();
const DIST = path.join(ROOT, 'dist');
const OUT_DIR = path.join(ROOT, 'dist-standalone');
const INDEX_TEMPLATE = path.join(ROOT, 'src', 'html', 'index.html');

const INCLUDE_FILES = [
  'styles',
  'util',
  'context',
  'apps-service-viewer',
  'apps-songs',
  'apps-team',
  'apps-weekly-plan',
  'apps-availability',
  'login',
  'service-viewer',
  'topbar',
  'songs',
  'songs-browser',
  'weekly-plan',
  'team',
  'availability'
];

const readFile = (filePath) => fs.readFileSync(filePath, 'utf8');

const replaceInclude = (buffer, name, content) => {
  const pattern = new RegExp(
    `<\\?!=\\s*HtmlService\\.createHtmlOutputFromFile\\('${name}'\\)\\.getContent\\(\\);?\\s*\\?>`,
    'g'
  );
  return buffer.replace(pattern, () => content);
};

const ensureDistFile = (name) => {
  const filePath = path.join(DIST, `${name}.html`);
  if (!fs.existsSync(filePath)) {
    throw new Error(`Standalone build missing dist/${name}.html. Run "npm run build" first.`);
  }
  return filePath;
};

function main() {
  if (!fs.existsSync(DIST)) {
    throw new Error('dist/ not found. Run "npm run build" before building the standalone bundle.');
  }

  let html = readFile(INDEX_TEMPLATE);
  INCLUDE_FILES.forEach((name) => {
    const filePath = ensureDistFile(name);
    const content = readFile(filePath);
    html = replaceInclude(html, name, content);
  });

  const base = process.env.APPS_SCRIPT_BASE || '';
  if (base) {
    const metaTag = `  <meta name="app-script-base" content="${base}">`;
    if (html.includes('<meta charset="utf-8" />')) {
      html = html.replace('<meta charset="utf-8" />', `<meta charset="utf-8" />\n${metaTag}`);
    } else {
      html = html.replace('<head>', `<head>\n${metaTag}`);
    }
  } else {
    console.warn('[standalone] APPS_SCRIPT_BASE env not set. Configure window.APP_RPC_BASE manually at runtime.');
  }

  fs.rmSync(OUT_DIR, { recursive: true, force: true });
  fs.mkdirSync(OUT_DIR, { recursive: true });
  fs.writeFileSync(path.join(OUT_DIR, 'index.html'), html);
  console.log('Standalone web app written to dist-standalone/index.html');
}

try {
  main();
} catch (err) {
  console.error('[standalone] build failed:', err);
  process.exitCode = 1;
}
