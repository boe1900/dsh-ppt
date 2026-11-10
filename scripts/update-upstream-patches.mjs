import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { checkUpstream, git, loadUpstream, normalizeJsdom, sha256 } from './check-upstream.mjs';

const root = fileURLToPath(new URL('../', import.meta.url));

export async function updateUpstreamPatches(checkout, projectRoot = root) {
  const { manifest, upstream } = await loadUpstream(checkout, projectRoot);
  for (const [name, original] of upstream) {
    if (Object.hasOwn(manifest.patches, name) || Object.hasOwn(manifest.generated, name)) continue;
    assert.deepEqual(normalizeJsdom(await readFile(path.join(projectRoot, 'lib', name))), original, `Cannot refresh unapproved module: ${name}`);
  }
  const patches = new Map();
  const temporary = await mkdtemp(path.join(os.tmpdir(), 'dsh-upstream-diff-'));
  try {
    await mkdir(path.join(temporary, 'a/lib'), { recursive: true });
    await mkdir(path.join(temporary, 'b/lib'), { recursive: true });
    for (const [name, record] of Object.entries(manifest.patches)) {
      const original = upstream.get(name);
      const local = await readFile(path.join(projectRoot, 'lib', name));
      await writeFile(path.join(temporary, 'a/lib', name), original);
      await writeFile(path.join(temporary, 'b/lib', name), normalizeJsdom(local));
      let diff;
      try {
        diff = await git(temporary, ['diff', '--no-index', '--no-ext-diff', '--no-prefix', '--', `a/lib/${name}`, `b/lib/${name}`]);
      } catch (error) {
        if (error.code !== 1) throw error;
        diff = error.stdout;
      }
      patches.set(record.patch, diff);
      record.upstreamSha256 = sha256(original);
      record.localSha256 = sha256(local);
      record.patchSha256 = sha256(diff);
    }
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
  await mkdir(path.join(projectRoot, 'scripts/upstream-patches'), { recursive: true });
  for (const [file, patch] of patches) await writeFile(path.join(projectRoot, file), patch);
  await writeFile(path.join(projectRoot, 'scripts/upstream-manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
  return await checkUpstream(checkout, projectRoot);
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  assert(process.argv[2], 'Usage: node scripts/update-upstream-patches.mjs /path/to/dsh-desktop');
  console.log(await updateUpstreamPatches(path.resolve(process.argv[2])));
}
