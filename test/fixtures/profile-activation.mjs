import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { mkdtemp, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = path.resolve(process.argv[2] ?? fileURLToPath(new URL('../../', import.meta.url)));
const pluginRequire = createRequire(path.join(root, 'package.json'));
const installAnchor = process.argv[3] ?? pluginRequire.resolve('@deepseek-ai/dsh/package.json');
const hostRequire = createRequire(installAnchor);
const { Context } = await import(pathToFileURL(hostRequire.resolve('@deepseek-ai/cordis')));
const { PluginPackages, createRuntimeResolution } = await import(pathToFileURL(hostRequire.resolve('@deepseek-ai/dsh-app-boot')));
const { default: PptxGenJS } = await import(pathToFileURL(pluginRequire.resolve('pptxgenjs')));
const base = await createRuntimeResolution({ installAnchor });
const ctx = new Context();
const scratch = await mkdtemp(path.join(os.tmpdir(), 'dsh-ppt-profile-'));

try {
  new PluginPackages(ctx, { resolution: {
    ...base,
    profilesDir: path.dirname(root),
    profileDir: root,
    localPackageNames: ['@cola1900/dsh-ppt'],
    linkedRoots: [],
  } });
  const plugin = await import(pathToFileURL(path.join(root, 'lib/index.js')));
  let rpc;
  ctx.provide('connection', { rpc: { handle: (_channel, handler) => { rpc = handler; } } });
  await ctx.plugin(plugin, { root: scratch });
  const state = await rpc('state', { sessionId: 'profile-activation' });
  assert.equal(state.ok, true);
  assert.equal(state.value.data.templates.length, 16);

  const { convertPptxToPptd } = await import(pathToFileURL(path.join(root, 'lib/pptx-converter.js')));
  const pptx = new PptxGenJS();
  pptx.addSlide().addText('Profile import regression', { x: 1, y: 1, w: 8, h: 1 });
  const bytes = Buffer.from(await pptx.write({ outputType: 'nodebuffer' }));
  const result = await convertPptxToPptd(bytes, 'profile.pptx');
  assert.equal(result.slideCount, 1);
  assert.equal(result.outputElementCount, 1);
  assert.match([...result.project.pages.values()].join('\n'), /Profile import regression/);
  console.log('Profile plugin activation and PPTX conversion passed');
} finally {
  await ctx.fiber.dispose();
  await rm(scratch, { recursive: true, force: true });
}
