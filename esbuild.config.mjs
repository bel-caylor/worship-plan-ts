// esbuild.config.mjs
import esbuild from 'esbuild';

const isWatch = process.argv.includes('--watch');

const common = {
  entryPoints: ['src/main.ts'],
  outfile: 'dist/Code.gs',
  bundle: true,
  platform: 'browser',
  format: 'iife',
  target: ['es2019'],
  banner: { js: 'var global=this;' },
  footer: { js: `
// Top-level wrappers so google.script.run can find callable server functions
function rpc(input){ return global.rpc.apply(global, arguments); }
function getFilesForFolderUrl(url, limit){ return global.getFilesForFolderUrl.apply(global, arguments); }
` },
  sourcemap: false,
  minify: false,
  logLevel: 'info'
};

async function run() {
  if (isWatch) {
    const ctx = await esbuild.context(common);
    await ctx.watch();
  } else {
    await esbuild.build(common);
  }
}

run().catch((e) => { console.error(e); process.exit(1); });
