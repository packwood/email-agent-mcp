import { describe, it, expect } from 'vitest';
import { spawn } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { STDIO_MAX_BUFFER_BYTES } from './server.js';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, '..', '..', '..');

// Regression: the SDK's default 10MB stdio cap closed the transport on a large
// attachment payload and the server exited without a word.
describe('email-mcp/stdio message size', () => {
  it('accepts a 35MB attachment set once base64-encoded', () => {
    const encoded = 4 * Math.ceil((35 * 1024 * 1024) / 3);
    expect(STDIO_MAX_BUFFER_BYTES).toBeGreaterThan(encoded + 1024 * 1024);
  });

  it('answers a 12MB request instead of exiting', async () => {
    const home = await mkdtemp(join(tmpdir(), 'email-mcp-stdio-'));
    const child = spawn(
      join(repoRoot, 'node_modules', '.bin', 'tsx'),
      [join(here, 'serve-entry.ts')],
      {
        env: { PATH: process.env.PATH ?? '', HOME: home, EMAIL_AGENT_MCP_HOME: home },
        stdio: ['pipe', 'pipe', 'ignore'],
      },
    );
    try {
      const response = await new Promise<Record<string, unknown>>((resolve, reject) => {
        let buffered = '';
        const timer = setTimeout(() => reject(new Error('no response within 30s')), 30_000);
        child.on('exit', code => {
          clearTimeout(timer);
          reject(new Error(`server exited with code ${code}`));
        });
        child.stdout.on('data', (chunk: Buffer) => {
          buffered += chunk.toString('utf8');
          let index: number;
          while ((index = buffered.indexOf('\n')) !== -1) {
            const message = JSON.parse(buffered.slice(0, index)) as Record<string, unknown>;
            buffered = buffered.slice(index + 1);
            if (message.id === 2) {
              clearTimeout(timer);
              resolve(message);
            }
          }
        });
        const send = (message: unknown) => child.stdin.write(JSON.stringify(message) + '\n');
        send({
          jsonrpc: '2.0',
          id: 1,
          method: 'initialize',
          params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '0' } },
        });
        send({ jsonrpc: '2.0', method: 'notifications/initialized' });
        send({
          jsonrpc: '2.0',
          id: 2,
          method: 'tools/call',
          params: { name: 'no_such_tool', arguments: { padding: 'x'.repeat(12 * 1024 * 1024) } },
        });
      });
      expect(response.id).toBe(2);
    } finally {
      child.kill();
      await rm(home, { recursive: true, force: true });
    }
  }, 40_000);
});
