import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const script = path.join(root, '..', 'scripts', 'docs', 'sync_starlight.py');
const candidates = process.platform === 'win32'
  ? [path.join(root, '..', '.venv-docs', 'Scripts', 'python.exe'), 'python', 'py']
  : ['python3', 'python'];

for (const command of candidates) {
  const args = command.endsWith('py') && process.platform === 'win32' ? ['-3', script] : [script];
  const result = spawnSync(command, args, { cwd: root, stdio: 'inherit' });
  if (result.error?.code === 'ENOENT') continue;
  process.exit(result.status ?? 1);
}

console.error('Python was not found. Install Python 3 and documentation dependencies first.');
process.exit(1);
