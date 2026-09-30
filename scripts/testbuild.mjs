// Test build channel ("MyExcel (Test)"): a side-by-side desktop app whose UI is read from
// builds/test/MyExcel Test/ui, so UI fixes ship without recompiling Rust.
//
//   node scripts/testbuild.mjs app            build the test exe (first time / Rust changes, ~2-8 min)
//   node scripts/testbuild.mjs ui "note"      build the UI and deploy it (~1 min); then Help > Reload
//
// Both append to builds/test/CHANGELOG.md. The release build (npm run tauri:build) is unaffected.
import { execSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const testDir = path.join(root, 'builds', 'test');
const appDir = path.join(testDir, 'MyExcel Test');
const numFile = path.join(testDir, 'build-number.txt');
const changelog = path.join(testDir, 'CHANGELOG.md');
const [mode, ...noteParts] = process.argv.slice(2);
const note = noteParts.join(' ').trim();

fs.mkdirSync(appDir, { recursive: true });
const run = (cmd, env = {}) => execSync(cmd, { cwd: root, stdio: 'inherit', env: { ...process.env, ...env } });
const stamp = () => new Date().toLocaleString('sv-SE').slice(0, 16);
const HEADER = '# MyExcel (Test) changelog\n\nNewest first. UI builds need Help > Reload; App rebuilds need the app restarted.\n\n';
function log(line) {
  const old = fs.existsSync(changelog) ? fs.readFileSync(changelog, 'utf8') : HEADER;
  fs.writeFileSync(changelog, `${HEADER}- ${stamp()} ${line}\n${old.slice(HEADER.length)}`);
}

if (mode === 'app') {
  // Separate target dir: the test feature set would otherwise invalidate the release build cache.
  run('npx tauri build --no-bundle --features test-build --config src-tauri/tauri.test.conf.json', {
    CARGO_TARGET_DIR: path.join(root, 'src-tauri', 'target-test'),
  });
  const rel = path.join(root, 'src-tauri', 'target-test', 'release');
  const exe = ['MyExcel-Test.exe', 'myexcel.exe'].map((f) => path.join(rel, f)).find((f) => fs.existsSync(f));
  if (!exe) throw new Error('test exe not found in ' + rel);
  // A running test app locks its exe; Windows still allows renaming it, so move it aside
  // (the running copy keeps working; the next start uses the new exe).
  const dest = path.join(appDir, 'MyExcel-Test.exe');
  for (const f of fs.readdirSync(appDir).filter((n) => n.startsWith('MyExcel-Test.old-'))) {
    try {
      fs.rmSync(path.join(appDir, f));
    } catch {
      /* still running */
    }
  }
  if (fs.existsSync(dest)) {
    try {
      fs.rmSync(dest);
    } catch {
      fs.renameSync(dest, path.join(appDir, `MyExcel-Test.old-${Date.now()}.exe`));
    }
  }
  fs.copyFileSync(exe, dest);
  log(`**App** rebuilt${note ? ': ' + note : ''} (restart the app)`);
  console.log(`\nTest app: ${path.join(appDir, 'MyExcel-Test.exe')}`);
} else if (mode === 'ui') {
  const n = (Number(fs.existsSync(numFile) ? fs.readFileSync(numFile, 'utf8') : '0') || 0) + 1;
  run('npm run build', { NEXT_PUBLIC_TEST_BUILD: String(n) });
  // Swap folders so a reload never sees a half-copied UI.
  const next = path.join(appDir, 'ui.new');
  const cur = path.join(appDir, 'ui');
  const old = path.join(appDir, 'ui.old');
  fs.rmSync(next, { recursive: true, force: true });
  fs.cpSync(path.join(root, 'out'), next, { recursive: true });
  fs.rmSync(old, { recursive: true, force: true });
  if (fs.existsSync(cur)) fs.renameSync(cur, old);
  fs.renameSync(next, cur);
  fs.rmSync(old, { recursive: true, force: true });
  fs.writeFileSync(numFile, String(n));
  log(`**UI build ${n}**${note ? ': ' + note : ''}`);
  console.log(`\nDeployed UI build ${n}. In MyExcel (Test): Help > Reload.`);
} else {
  console.log('usage: node scripts/testbuild.mjs app|ui ["what changed"]');
  process.exit(1);
}
