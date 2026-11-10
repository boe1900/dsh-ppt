import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import { promisify } from 'node:util';
import { checkUpstream, sha256 } from '../scripts/check-upstream.mjs';
import { updateUpstreamPatches } from '../scripts/update-upstream-patches.mjs';

const exec = promisify(execFile);

async function fixture(t) {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'dsh-upstream-test-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const checkout = path.join(directory, 'upstream');
  const root = path.join(directory, 'project');
  const core = 'packages/ppt-runtime/core';
  for (const destination of [path.join(checkout, core, 'lib'), path.join(root, 'lib'), path.join(root, 'scripts')]) await mkdir(destination, { recursive: true });
  for (const [name, content] of Object.entries({ 'index.js': 'export const host = "desktop";\n', 'unchanged.js': 'import { JSDOM } from "jsdom";\n' })) {
    await writeFile(path.join(checkout, core, 'lib', name), content);
    await writeFile(path.join(root, 'lib', name), content);
  }
  for (const file of ['SKILL.md', 'references/pptd.md', 'references/composition.md']) {
    for (const base of [path.join(checkout, core), root]) {
      const destination = path.join(base, 'skills/dsh-ppt', file);
      await mkdir(path.dirname(destination), { recursive: true });
      await writeFile(destination, `# ${file}\n`);
    }
  }
  await exec('git', ['init', '--quiet'], { cwd: checkout });
  await exec('git', ['add', '.'], { cwd: checkout });
  await exec('git', ['-c', 'user.name=Fixture', '-c', 'user.email=fixture@example.invalid', '-c', 'commit.gpgsign=false', 'commit', '--quiet', '-m', 'Upstream fixture'], { cwd: checkout });
  const commit = (await exec('git', ['rev-parse', 'HEAD'], { cwd: checkout })).stdout.trim();
  const manifest = { version: 1, commit, patches: { 'index.js': { reason: 'Standalone host integration.', patch: 'scripts/upstream-patches/index.js.patch' } }, generated: { 'client.js': { reason: 'Generated standalone browser client.' } }, local: { 'store.js': { reason: 'Standalone persistence.' } } };
  await writeFile(path.join(root, 'lib/index.js'), 'export const host = "standalone";\n');
  await writeFile(path.join(root, 'lib/unchanged.js'), 'import { JSDOM } from "./jsdom.js";\n');
  await writeFile(path.join(root, 'lib/client.js'), 'export const client = true;\n');
  await writeFile(path.join(root, 'lib/store.js'), 'export const store = true;\n');
  await writeFile(path.join(root, 'scripts/upstream-manifest.json'), JSON.stringify(manifest));
  await updateUpstreamPatches(checkout, root);
  return { checkout, root, core };
}

test('upstream guard replays reviewed patches and accepts only the jsdom import normalization', async t => {
  const { checkout, root } = await fixture(t);
  assert.match(await checkUpstream(checkout, root), /replayed 1 reviewed patches/);
  await writeFile(path.join(root, 'lib/unchanged.js'), 'import { JSDOM } from "./jsdom.js";\nexport const changed = true;\n');
  await assert.rejects(checkUpstream(checkout, root), /Unreviewed upstream runtime change: unchanged.js/);
  await assert.rejects(updateUpstreamPatches(checkout, root), /Cannot refresh unapproved module: unchanged.js/);
});

test('upstream guard detects accidental edits to a reviewed module and patch', async t => {
  const { checkout, root } = await fixture(t);
  const source = path.join(root, 'lib/index.js');
  const original = await readFile(source);
  await writeFile(source, 'export const host = "unexpected";\n');
  await assert.rejects(checkUpstream(checkout, root), /Local hash differs; review and refresh patch: index.js/);
  await writeFile(source, original);
  await writeFile(path.join(root, 'scripts/upstream-patches/index.js.patch'), 'corrupted patch\n');
  await assert.rejects(checkUpstream(checkout, root), /Reviewed patch hash differs: index.js/);
});

test('patch replay must reproduce the local bytes even when recorded hashes are updated', async t => {
  const { checkout, root } = await fixture(t);
  const source = path.join(root, 'lib/index.js');
  const changed = Buffer.from('export const host = "unreviewed";\n');
  await writeFile(source, changed);
  const manifestPath = path.join(root, 'scripts/upstream-manifest.json');
  const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
  manifest.patches['index.js'].localSha256 = sha256(changed);
  await writeFile(manifestPath, JSON.stringify(manifest));
  await assert.rejects(checkUpstream(checkout, root), /Reviewed patch does not reproduce local module: index.js/);
});

test('upstream guard rejects unknown local modules and refresh never whitelists them', async t => {
  const { checkout, root } = await fixture(t);
  await writeFile(path.join(root, 'lib/surprise.js'), 'export default true;\n');
  await assert.rejects(checkUpstream(checkout, root), /Local runtime module list differs from explicit manifest/);
  await assert.rejects(updateUpstreamPatches(checkout, root), /Local runtime module list differs from explicit manifest/);
});

test('upstream guard checks checkout content and the pinned commit', async t => {
  const { checkout, root, core } = await fixture(t);
  const source = path.join(checkout, core, 'lib/index.js');
  await writeFile(source, 'export const host = "modified";\n');
  await assert.rejects(checkUpstream(checkout, root), /Upstream checkout is modified: index.js/);
  await exec('git', ['add', '.'], { cwd: checkout });
  await exec('git', ['-c', 'user.name=Fixture', '-c', 'user.email=fixture@example.invalid', '-c', 'commit.gpgsign=false', 'commit', '--quiet', '-m', 'Unexpected upstream'], { cwd: checkout });
  await assert.rejects(checkUpstream(checkout, root), /Upstream checkout must be at the pinned commit/);
});

test('refresh is deterministic and retains explicit review reasons', async t => {
  const { checkout, root } = await fixture(t);
  const manifestPath = path.join(root, 'scripts/upstream-manifest.json');
  const original = await readFile(manifestPath, 'utf8');
  await updateUpstreamPatches(checkout, root);
  assert.equal(await readFile(manifestPath, 'utf8'), original);
  assert.equal(JSON.parse(original).patches['index.js'].reason, 'Standalone host integration.');
});
