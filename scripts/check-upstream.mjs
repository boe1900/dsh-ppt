import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const checkout = process.argv[2];
assert(checkout, 'Usage: node scripts/check-upstream.mjs /path/to/dsh-desktop');
const root = fileURLToPath(new URL('../', import.meta.url));
const upstream = path.resolve(checkout, 'packages/ppt-runtime/core/lib');
const adapters = new Set(['index.js', 'invariant.js', 'office-template-preview.js', 'client.js', 'catalog.js', 'preview-manifest.js']);
const runtimeFiles = names => names.filter(name => name.endsWith('.js') && !adapters.has(name) && name !== 'preset.js').sort();
const files = runtimeFiles(await readdir(upstream));
assert.deepEqual(runtimeFiles(await readdir(path.join(root, 'lib'))), files, 'Upstream runtime module list differs');
let compared = 0;
for (const name of files) {
  assert.deepEqual(await readFile(path.join(root, 'lib', name)), await readFile(path.join(upstream, name)), `Upstream runtime differs: ${name}`);
  compared++;
}
assert(compared > 0, 'No upstream runtime modules found');
for (const file of ['SKILL.md', 'references/pptd.md', 'references/composition.md']) {
  assert.deepEqual(await readFile(path.join(root, 'skills/dsh-ppt', file)), await readFile(path.resolve(checkout, 'packages/ppt-runtime/core/skills/dsh-ppt', file)), `Upstream Skill differs: ${file}`);
}
console.log(`Matched ${compared} upstream runtime modules and 3 Skill documents. Local adapters require separate integration checks.`);
