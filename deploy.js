// Push the project to GitHub: commits any changes, then pushes the current branch.
//   node deploy.js                  commit message defaults to "Update <date time>"
//   node deploy.js "Fix wrap text"  your own commit message
const { execSync } = require('node:child_process');

const git = (args, opts = {}) => execSync(`git ${args}`, { cwd: __dirname, encoding: 'utf8', ...opts }).trim();
const run = (args) => execSync(`git ${args}`, { cwd: __dirname, stdio: 'inherit' });

try {
  const branch = git('branch --show-current');
  const message = process.argv.slice(2).join(' ').trim() || `Update ${new Date().toLocaleString('sv-SE').slice(0, 16)}`;

  if (git('status --porcelain')) {
    run('add -A');
    execSync('git commit -F -', { cwd: __dirname, input: message, stdio: ['pipe', 'inherit', 'inherit'] });
  } else {
    console.log('No changes to commit.');
  }

  // push; sets the upstream the first time
  const hasUpstream = (() => {
    try {
      git('rev-parse --abbrev-ref @{u}', { stdio: 'pipe' });
      return true;
    } catch {
      return false;
    }
  })();
  run(hasUpstream ? 'push' : `push -u origin ${branch}`);

  console.log(`\nPushed ${branch} @ ${git('rev-parse --short HEAD')} to ${git('remote get-url origin')}`);
} catch (e) {
  console.error('\nDeploy failed:', e.message.split('\n')[0]);
  process.exit(1);
}
