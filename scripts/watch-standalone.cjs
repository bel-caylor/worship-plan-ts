/* eslint-disable no-console */
const chokidar = require('chokidar');
const { spawn } = require('child_process');

const WATCH_GLOB = 'src/**/*';
const BUILD_COMMAND = process.platform === 'win32' ? 'npm run build:standalone' : 'npm';
const BUILD_ARGS = process.platform === 'win32' ? [] : ['run', 'build:standalone'];

let running = false;
let queued = false;

function runBuild(triggerLabel) {
  if (running) {
    queued = true;
    return;
  }

  running = true;
  console.log(`[watch:standalone] starting rebuild (${triggerLabel})`);

  const child = spawn(BUILD_COMMAND, BUILD_ARGS, {
    cwd: process.cwd(),
    stdio: 'inherit',
    shell: process.platform === 'win32'
  });

  child.on('exit', (code) => {
    running = false;
    if (code === 0) {
      console.log('[watch:standalone] rebuild complete');
    } else {
      console.error(`[watch:standalone] rebuild failed with exit code ${code}`);
    }

    if (queued) {
      queued = false;
      runBuild('queued change');
    }
  });
}

const watcher = chokidar.watch(WATCH_GLOB, {
  ignoreInitial: true
});

watcher
  .on('ready', () => {
    if (!running) runBuild('initial');
  })
  .on('all', (event, filePath) => {
    if (event === 'add' || event === 'change' || event === 'unlink' || event === 'addDir' || event === 'unlinkDir') {
      runBuild(`${event}:${filePath}`);
    }
  })
  .on('error', (error) => {
    console.error('[watch:standalone] watcher error:', error);
  });
