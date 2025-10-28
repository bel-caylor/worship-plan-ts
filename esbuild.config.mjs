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
