import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

test('installed profile resolver activates the plugin and converts PPTX', { timeout: 60000 }, () => {
  const result = spawnSync(process.execPath, [fileURLToPath(new URL('fixtures/profile-activation.mjs', import.meta.url))], {
    encoding: 'utf8',
    timeout: 55000,
  });
  assert.equal(result.status, 0, `${result.error ?? ''}\n${result.stdout}\n${result.stderr}`);
  assert.match(result.stdout, /Profile plugin activation and PPTX conversion passed/);
});
