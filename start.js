// Start the MyExcel web app and open it in the browser.
//   node start.js              serve the production build (rebuilds first if the code changed)
//   node start.js --dev        development server with hot reload
//   node start.js --build      force a fresh build before serving
//   node start.js --port 4000  use another port (default 3100)
//   node start.js --no-open    don't open the browser
const { execSync, spawn } = require('node:child_process');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');

const root = __dirname;
const args = process.argv.slice(2);
const has = (f) => args.includes(f);
const port = Number(args[args.indexOf('--port') + 1]) || Number(process.env.PORT) || 3100;
const url = `http://localhost:${port}`;

/** Newest modification time of the files under the given folders. */
function newest(dirs) {
  let t = 0;
  const walk = (d) => {
    for (const e of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, e.name);
      if (e.isDirectory()) walk(p);
      else t = Math.max(t, fs.statSync(p).mtimeMs);
    }
  };
  for (const d of dirs) if (fs.existsSync(path.join(root, d))) walk(path.join(root, d));
  for (const f of ['package.json', 'next.config.mjs', 'tailwind.config.ts']) {
    if (fs.existsSync(path.join(root, f))) t = Math.max(t, fs.statSync(path.join(root, f)).mtimeMs);
  }
  return t;
}

/** Resolves true when something answers on the port. */
function responds() {
  return new Promise((resolve) => {
    const req = http.get(url, (res) => {
      res.resume();
      resolve(true);
    });
    req.on('error', () => resolve(false));
    req.setTimeout(1000, () => {
      req.destroy();
      resolve(false);
    });
  });
}

function openBrowser() {
  if (has('--no-open')) return;
  const cmd = process.platform === 'win32' ? `start "" "${url}"` : process.platform === 'darwin' ? `open "${url}"` : `xdg-open "${url}"`;
  try {
    execSync(cmd, { stdio: 'ignore' });
  } catch {
    /* no browser available: the URL is printed anyway */
  }
}

(async () => {
  if (await responds()) {
    console.log(`MyExcel is already running on ${url}`);
    openBrowser();
    return;
  }
  if (!fs.existsSync(path.join(root, 'node_modules'))) {
    console.log('Installing dependencies...');
    execSync('npm install', { cwd: root, stdio: 'inherit' });
  }

  let server;
  if (has('--dev')) {
    server = spawn('npx', ['cross-env', 'NEXT_TELEMETRY_DISABLED=1', 'next', 'dev', '-p', String(port)], { cwd: root, stdio: 'inherit', shell: true });
  } else {
    const index = path.join(root, 'out', 'index.html');
    const stale = !fs.existsSync(index) || newest(['src', 'app', 'public']) > fs.statSync(index).mtimeMs;
    if (has('--build') || stale) {
      console.log(stale ? 'Code changed since the last build - building (about a minute)...' : 'Building...');
      execSync('npm run build', { cwd: root, stdio: 'inherit' });
    }
    server = spawn(process.execPath, [path.join(root, 'scripts', 'serve-static.mjs'), String(port)], { cwd: root, stdio: 'inherit' });
  }
  server.on('exit', (code) => process.exit(code ?? 0));
  process.on('SIGINT', () => server.kill());

  // open the browser once the server answers (the dev server needs a few seconds for its first build)
  for (let i = 0; i < 120; i++) {
    if (await responds()) {
      console.log(`\nMyExcel is running on ${url} - press Ctrl+C to stop.`);
      openBrowser();
      return;
    }
    await new Promise((r) => setTimeout(r, 500));
  }
  console.log(`Server did not answer on ${url} yet; open it manually when it's ready.`);
})().catch((e) => {
  console.error('Start failed:', e.message);
  process.exit(1);
});
