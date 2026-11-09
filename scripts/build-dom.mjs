import assert from 'node:assert/strict';
import { createRequire, isBuiltin } from 'node:module';
import { copyFile, mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

const root = fileURLToPath(new URL('../', import.meta.url));
const require = createRequire(import.meta.url);
const jsdom = path.dirname(require.resolve('jsdom/package.json'));
const vendor = path.join(root, 'lib/vendor');
const styleHelper = path.join(jsdom, 'lib/jsdom/living/css/helpers/computed-style.js');
const xhr = path.join(jsdom, 'lib/jsdom/living/xhr/XMLHttpRequest-impl.js');

await mkdir(vendor, { recursive: true });
const result = await build({
  absWorkingDir: root,
  entryPoints: {
    jsdom: path.join(jsdom, 'lib/api.js'),
    'xhr-sync-worker': path.join(jsdom, 'lib/jsdom/living/xhr/xhr-sync-worker.js'),
  },
  outdir: vendor,
  outExtension: { '.js': '.cjs' },
  bundle: true,
  platform: 'node',
  target: 'node22.19',
  format: 'cjs',
  external: ['canvas', 'saxes'],
  legalComments: 'linked',
  metafile: true,
  plugins: [{
    name: 'jsdom-runtime-assets',
    setup(context) {
      // css-tree's CJS entry lets esbuild include JSON data hidden behind createRequire in its ESM entry.
      context.onResolve({ filter: /^css-tree(?:\/|$)/ }, ({ path: specifier }) => ({ path: require.resolve(specifier) }));
      context.onLoad({ filter: /[\\/]jsdom[\\/]lib[\\/]jsdom[\\/]living[\\/]/ }, async ({ path: filename }) => {
        if (filename !== styleHelper && filename !== xhr) return;
        let contents = await readFile(filename, 'utf8');
        const before = filename === styleHelper
          ? '"../../../browser/default-stylesheet.css"'
          : 'require.resolve("./xhr-sync-worker.js")';
        const after = filename === styleHelper
          ? '"default-stylesheet.css"'
          : 'require("node:path").join(__dirname, "xhr-sync-worker.cjs")';
        assert.equal(contents.split(before).length, 2, `JSdom asset reference changed: ${filename}`);
        contents = contents.replace(before, after);
        return { contents, loader: 'js' };
      });
    },
  }],
});
for (const output of Object.values(result.metafile.outputs)) {
  for (const dependency of output.imports) {
    assert(dependency.external && (isBuiltin(dependency.path) || ['canvas', 'saxes'].includes(dependency.path)),
      `Unexpected bundled DOM runtime dependency: ${dependency.path}`);
  }
}
await copyFile(path.join(jsdom, 'lib/jsdom/browser/default-stylesheet.css'), path.join(vendor, 'default-stylesheet.css'));

const packages = new Map();
for (const input of Object.keys(result.metafile.inputs)) {
  let directory = path.dirname(path.resolve(root, input));
  while (directory !== root && directory !== path.dirname(directory)) {
    try {
      const manifest = JSON.parse(await readFile(path.join(directory, 'package.json'), 'utf8'));
      if (manifest.name && manifest.version) {
        packages.set(directory, manifest);
        break;
      }
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
    directory = path.dirname(directory);
  }
}
const notices = [];
for (const [directory, manifest] of [...packages].sort((a, b) => a[1].name.localeCompare(b[1].name))) {
  const files = (await readdir(directory, { recursive: true, withFileTypes: true }))
    .filter(entry => entry.isFile() && /^(?:licen[cs]e|copying|notice)(?:[.-]|$)/i.test(entry.name)
      && !path.relative(directory, entry.parentPath).split(path.sep).includes('node_modules'))
    .map(entry => path.relative(directory, path.join(entry.parentPath, entry.name))).sort();
  assert(files.length > 0, `Missing bundled dependency license: ${manifest.name}`);
  const licenses = await Promise.all(files.map(file => readFile(path.join(directory, file), 'utf8')));
  notices.push(`${manifest.name}@${manifest.version} (${manifest.license})\n\n${licenses.join('\n\n')}`);
}
await writeFile(path.join(vendor, 'LICENSES.txt'), notices.join('\n\n----------------------------------------\n\n'));
console.error(`Bundled JSdom with ${packages.size} dependency licenses and its stylesheet/XHR worker.`);
