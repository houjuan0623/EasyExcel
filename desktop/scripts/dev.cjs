// Copyright (c) 2026 AbeLin · MIT License
const { spawn } = require('node:child_process');
const { join } = require('node:path');
const { once } = require('node:events');

const root = join(__dirname, '..');
const children = new Set();
function start(exe, args, env = {}) {
  const child = spawn(exe, args, { cwd: root, stdio: 'inherit', shell: false, env: { ...process.env, ...env } });
  children.add(child);
  child.on('exit', () => children.delete(child));
  return child;
}
async function main() {
  const compiler = start(process.execPath, [require.resolve('typescript/bin/tsc'), '-p', 'tsconfig.electron.json']);
  const [code] = await once(compiler, 'exit');
  if (code) process.exit(code);
  const { createServer } = await import('vite');
  const server = await createServer({ configFile: join(root, 'vite.config.mts') });
  await server.listen();
  const env = { ...process.env, ELECTRON_RENDERER_URL: 'http://127.0.0.1:5173' };
  delete env.ELECTRON_RUN_AS_NODE;
  const electron = spawn(require('electron'), ['.'], { cwd: root, stdio: 'inherit', env, shell: false });
  children.add(electron);
  electron.on('exit', async (code) => { children.delete(electron); await server.close(); process.exit(code || 0); });
  electron.on('error', async (error) => { console.error(error); await server.close(); process.exit(1); });
}
process.on('exit', () => { for (const child of children) child.kill(); });
process.on('SIGINT', () => process.exit(0));
process.on('SIGTERM', () => process.exit(0));
main().catch((error) => { console.error(error); process.exit(1); });
