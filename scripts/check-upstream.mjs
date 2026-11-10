import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, readdir, mkdir, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { promisify } from 'node:util';

const exec = promisify(execFile);
const root = fileURLToPath(new URL('../', import.meta.url));
const core = 'packages/ppt-runtime/core';
export const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
export const normalizeJsdom = bytes => Buffer.from(bytes.toString('utf8').replace(/from (['"])\.\/jsdom\.js\1/g, 'from $1jsdom$1'));
export const git = async (cwd, args) => (await exec('git', args, { cwd, encoding: 'buffer', maxBuffer: 32 * 1024 * 1024 })).stdout;

export async function loadUpstream(checkout, projectRoot = root) {
  const manifest = JSON.parse(await readFile(path.join(projectRoot, 'scripts/upstream-manifest.json'), 'utf8'));
  assert.equal(manifest.version, 1, 'Unsupported upstream manifest version');
  assert.match(manifest.commit, /^[0-9a-f]{40}$/, 'Manifest must pin a full upstream commit');
  assert.equal((await git(checkout, ['rev-parse', 'HEAD'])).toString().trim(), manifest.commit, 'Upstream checkout must be at the pinned commit');
  const upstreamNames = (await git(checkout, ['ls-tree', '--name-only', `${manifest.commit}:${core}/lib`])).toString().trim().split('\n').filter(name => name.endsWith('.js')).sort();
  assert(upstreamNames.length > 0, 'No upstream runtime modules found');
  const allNames = new Set();
  for (const [kind, entries] of Object.entries({ patches: manifest.patches, generated: manifest.generated, local: manifest.local })) {
    assert(entries && typeof entries === 'object' && !Array.isArray(entries), `Invalid manifest category: ${kind}`);
    for (const [name, entry] of Object.entries(entries)) {
      assert.match(name, /^[a-zA-Z0-9_-]+\.js$/, `Invalid module name: ${name}`);
      assert(!allNames.has(name), `Module has multiple manifest categories: ${name}`);
      allNames.add(name);
      assert(typeof entry.reason === 'string' && entry.reason.trim(), `Missing reason for ${name}`);
      if (kind === 'patches') {
        assert(upstreamNames.includes(name), `Patched module is absent upstream: ${name}`);
        assert.equal(entry.patch, `scripts/upstream-patches/${name}.patch`, `Unexpected patch path for ${name}`);
      }
      if (kind === 'local') assert(!upstreamNames.includes(name), `Local-only module exists upstream: ${name}`);
    }
  }
  const jsNames = async directory => (await readdir(directory)).filter(name => name.endsWith('.js')).sort();
  assert.deepEqual(await jsNames(path.join(checkout, core, 'lib')), upstreamNames, 'Upstream checkout runtime module list differs from pinned commit');
  assert.deepEqual(await jsNames(path.join(projectRoot, 'lib')), [...new Set([...upstreamNames, ...Object.keys(manifest.local), ...Object.keys(manifest.generated)])].sort(), 'Local runtime module list differs from explicit manifest');
  const upstream = new Map();
  for (const name of upstreamNames) {
    const committed = await git(checkout, ['show', `${manifest.commit}:${core}/lib/${name}`]);
    assert.deepEqual(await readFile(path.join(checkout, core, 'lib', name)), committed, `Upstream checkout is modified: ${name}`);
    upstream.set(name, committed);
  }
  return { manifest, upstream };
}

async function replayPatch(name, upstream, patch) {
  if (!patch.length) return upstream;
  const temporary = await mkdtemp(path.join(os.tmpdir(), 'dsh-upstream-replay-'));
  try {
    await mkdir(path.join(temporary, 'lib'));
    await writeFile(path.join(temporary, 'lib', name), upstream);
    await writeFile(path.join(temporary, 'reviewed.patch'), patch);
    const summary = (await git(temporary, ['apply', '--numstat', 'reviewed.patch'])).toString().trim().split('\n');
    assert.equal(summary.length, 1, `Patch must change exactly one module: ${name}`);
    assert.equal(summary[0].split('\t')[2], `lib/${name}`, `Patch targets another module: ${name}`);
    await git(temporary, ['apply', '--whitespace=nowarn', 'reviewed.patch']);
    return await readFile(path.join(temporary, 'lib', name));
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
}

export async function checkUpstream(checkout, projectRoot = root) {
  const { manifest, upstream } = await loadUpstream(checkout, projectRoot);
  let identical = 0;
  for (const [name, original] of upstream) {
    if (Object.hasOwn(manifest.generated, name)) continue;
    const local = await readFile(path.join(projectRoot, 'lib', name));
    const record = manifest.patches[name];
    if (!record) {
      assert.deepEqual(normalizeJsdom(local), original, `Unreviewed upstream runtime change: ${name}`);
      identical++;
      continue;
    }
    assert.equal(sha256(original), record.upstreamSha256, `Upstream hash differs: ${name}`);
    assert.equal(sha256(local), record.localSha256, `Local hash differs; review and refresh patch: ${name}`);
    const patch = await readFile(path.join(projectRoot, record.patch));
    assert.equal(sha256(patch), record.patchSha256, `Reviewed patch hash differs: ${name}`);
    assert.deepEqual(await replayPatch(name, original, patch), normalizeJsdom(local), `Reviewed patch does not reproduce local module: ${name}`);
  }
  for (const file of ['SKILL.md', 'references/pptd.md', 'references/composition.md']) {
    const original = await git(checkout, ['show', `${manifest.commit}:${core}/skills/dsh-ppt/${file}`]);
    assert.deepEqual(await readFile(path.join(checkout, core, 'skills/dsh-ppt', file)), original, `Upstream Skill checkout is modified: ${file}`);
    assert.deepEqual(await readFile(path.join(projectRoot, 'skills/dsh-ppt', file)), original, `Upstream Skill differs: ${file}`);
  }
  return `Matched ${identical} unchanged upstream modules, replayed ${Object.keys(manifest.patches).length} reviewed patches, checked explicit generated/local module inventory and 3 Skill documents.`;
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  assert(process.argv[2], 'Usage: node scripts/check-upstream.mjs /path/to/dsh-desktop');
  console.log(await checkUpstream(path.resolve(process.argv[2])));
}
