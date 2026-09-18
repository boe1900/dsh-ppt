import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import { createHash, randomUUID } from 'node:crypto';
import { createServer } from 'node:http';
import { createRequire } from 'node:module';
import { cp, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';
import test from 'node:test';
import { unzipSync } from 'fflate';
import { JSDOM } from 'jsdom';
import yaml from 'js-yaml';
import { Context } from '@deepseek-ai/cordis';
import { SlotCore } from '@deepseek-ai/dsh-client-ui-slots';
import { SkillRegistry } from '@deepseek-ai/dsh-skill';
import { SystemPrompt } from '@deepseek-ai/dsh-system-prompt';
import manifest from '../package.json' with { type: 'json' };
import { definitions } from '../lib/catalog.js';
import { previewFiles } from '../lib/preview-manifest.js';
import { createPreviewHandler } from '../lib/preview-assets.js';
import { checkPptdProject, loadPptdProject } from '../lib/pptd.js';

const root = fileURLToPath(new URL('../', import.meta.url));
const references = path.join(root, 'skills/dsh-ppt/references');
const cli = path.join(root, 'lib/bin.js');

test('all 16 templates include valid bilingual source and matching previews', async () => {
  assert.equal(definitions.length, 16);
  assert.equal(definitions.reduce((sum, template) => sum + template.referencePageCount, 0), 192);
  assert.equal(Object.keys(previewFiles).length, 192);
  for (const template of definitions) {
    for (const language of ['source', 'source-zh']) {
      const project = await loadPptdProject(path.join(references, template.referenceDirectory, language));
      const check = checkPptdProject(project);
      assert.equal(project.pages.length, template.referencePageCount);
      assert.equal(check.errorCount, 0, `${template.id}/${language}: ${JSON.stringify(check.issues)}`);
    }
  }
  for (const [hash, relative] of Object.entries(previewFiles)) {
    const bytes = await readFile(path.join(references, relative));
    assert.equal(createHash('sha256').update(bytes).digest('hex'), hash);
  }
});

test('Cordis loads the single bundle, scopes PPT mode, and exports an editable deck', async t => {
  const schema = yaml.DEFAULT_SCHEMA.extend(new yaml.Type('tag:yaml.org,2002:js', { kind: 'scalar' }));
  const [{ insert: [entry] }] = yaml.load(await readFile(path.join(root, 'cordis.patch.yml'), 'utf8'), { schema });
  assert.equal(entry.name, manifest.name);
  const plugin = await import(entry.name);
  const scratch = await mkdtemp(path.join(os.tmpdir(), 'dsh-ppt-test-'));
  t.after(() => rm(scratch, { recursive: true, force: true }));
  const workspace = path.join(scratch, 'workspace');
  await mkdir(workspace);
  await cp(path.join(references, 'business/dsh-blue-professional/source-zh'), path.join(workspace, 'deck'), { recursive: true });

  const ctx = new Context();
  const tools = new Map();
  const routes = new Map();
  let rpc;
  ctx.provide('connection', { rpc: { handle: (_channel, handler) => { rpc = handler; } } });
  ctx.provide('tools', { register: tool => tools.set(tool.name, tool) });
  ctx.provide('webServer', { register: route => {
    routes.set(route.path, route);
    return () => routes.delete(route.path);
  } });
  const prompt = ctx.plugin(SystemPrompt, { includeHarnessIdentity: false, personaPrefix: 'Default persona.' });
  const skills = ctx.plugin(SkillRegistry);
  await prompt;
  await skills;
  const runtime = ctx.plugin(plugin, { root: path.join(scratch, 'storage') });
  t.after(async () => { await runtime.dispose(); await skills.dispose(); await prompt.dispose(); });
  await runtime;
  assert(tools.has('pptd_render'));
  assert(routes.has('/dsh-ppt/previews'));
  assert(routes.has('/dsh-ppt'));
  assert(await ctx.skills.get('dsh-ppt', {}));

  const sessionId = randomUUID();
  assert.equal((await rpc('state', { sessionId })).value.data.templates.length, 16);
  assert.equal((await rpc('template/select', { sessionId, templateId: 'dsh-blue-professional', mode: 'ppt' })).value.status, 'ok');
  const active = await ctx.systemPrompt.assemble({ agent: { id: sessionId } });
  const inactive = await ctx.systemPrompt.assemble({ agent: { id: randomUUID() } });
  assert(active.sections.some(section => section.name === 'tool:dsh-ppt'));
  assert(!inactive.sections.some(section => section.name === 'tool:dsh-ppt'));
  assert.equal((await rpc('presentation/mode', { sessionId, mode: null })).value.status, 'ok');
  assert(!(await ctx.systemPrompt.assemble({ agent: { id: sessionId } })).sections.some(section => section.name === 'tool:dsh-ppt'));

  const execution = { agent: { id: sessionId, session: { header: { cwd: workspace } } }, signal: new AbortController().signal };
  const run = (name, args) => tools.get(name).execute(args, execution);
  assert.equal((await run('pptd_check', { project_path: 'deck' })).errorCount, 0);
  await assert.rejects(run('pptd_check', { project_path: '../outside' }), /inside the active workspace/);
  const output = await run('pptd_render', { project_path: 'deck', output_file: 'example.pptx' });
  assert.equal(output.status, 'exported');
  assert.equal(output.pageCount, 12);
  const zip = unzipSync(await readFile(path.join(workspace, output.outputPath)));
  const slides = Object.keys(zip).filter(file => /^ppt\/slides\/slide\d+\.xml$/.test(file));
  assert.equal(slides.length, 12);
  assert(slides.some(file => Buffer.from(zip[file]).toString().includes('<a:t>')), 'Export must retain editable text');
  assert.equal((await rpc('state', { sessionId })).value.data.decks.length, 1);

  const args = { project_path: 'invalid', file_path: 'deck.pptd', content: 'version: v2\nsize: [960, 540]\npages: [pages/1.page]\n' };
  await run('pptd_write_file', args);
  await assert.rejects(run('pptd_write_file', { ...args, content: 'overwritten' }));
  const previous = await run('pptd_read_file', { project_path: 'invalid', file_path: 'deck.pptd' });
  await run('pptd_write_file', { ...args, expected_sha256: previous.sha256 });
  await run('pptd_write_file', { project_path: 'invalid', file_path: 'pages/1.page', content: JSON.stringify({ elements: [
    { elementId: 'caption', elementType: 'text', bounds: [48, 80, 800, 1], content: { text: 'Overflow must block export', fontSize: 32 } }
  ] }) });
  const invalid = await run('pptd_render', { project_path: 'invalid', output_file: 'invalid.pptx' });
  assert.equal(invalid.status, 'needs_revision');
  assert(invalid.check.issues.some(issue => issue.code === 'text-overflow'));
  assert.equal((await rpc('state', { sessionId })).value.data.decks.length, 1);
  const failedCli = spawnSync(process.execPath, [cli, 'render', path.join(workspace, 'invalid'), '-o', path.join(scratch, 'invalid.pptx'), '--json'], { encoding: 'utf8' });
  assert.equal(failedCli.status, 1, failedCli.stderr);
  assert.equal(JSON.parse(failedCli.stdout).exported, false);

  const screenshots = path.join(scratch, 'screenshots');
  execFileSync(process.execPath, [cli, 'screenshot', path.join(workspace, 'deck'), '-p', '1', '-o', screenshots, '--json']);
  const screenshot = JSON.parse(await readFile(path.join(screenshots, 'index.json'), 'utf8'));
  const png = await readFile(path.join(screenshots, screenshot.pages[0].file));
  assert.equal(png.subarray(1, 4).toString(), 'PNG');
  await runtime.dispose();
  assert.equal(routes.size, 0, 'Plugin disposal must release HTTP routes');
});

test('toolbar picker selects, closes, retries failures, and exits without toggling on browse', async t => {
  const dom = new JSDOM('<!doctype html><div id="root"></div>', { pretendToBeVisual: true, url: 'http://localhost' });
  const saved = ['window', 'document', 'IS_REACT_ACT_ENVIRONMENT'].map(key => [key, Object.getOwnPropertyDescriptor(globalThis, key)]);
  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  let unmount;
  t.after(async () => {
    await unmount?.();
    dom.window.close();
    for (const [key, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else delete globalThis[key];
    }
  });
  let entry;
  dom.window.__ModuleLoader__ = { load: value => { entry = value; } };
  vm.runInNewContext(await readFile(path.join(root, 'lib/client.js'), 'utf8'), {
    window: dom.window, document: dom.window.document, AbortController, Node: dom.window.Node, performance,
  });
  assert.equal(entry.id, manifest.name);
  const require = createRequire(import.meta.url);
  const React = require('react');
  const { createRoot } = require('react-dom/client');
  const client = entry.factory(require);
  const slots = new SlotCore();
  slots.register({ name: 'root', children: {
    'conversation.input.left': { kind: 'list', scope: 'session' },
    'conversation.input.dock': { kind: 'list', scope: 'session' },
  } }, () => null);
  const seats = new Map();
  let dictionary, failSelect = false, failExit = false, blank = true;
  const state = { templates: definitions, selectedTemplateId: null, presentationMode: null };
  const requests = [];
  const rpc = { async call(channel, endpoint, payload) {
    assert.equal(channel, '/dsh-ppt');
    requests.push({ endpoint, payload });
    if ((endpoint === 'template/select' && failSelect) || (endpoint === 'presentation/mode' && failExit)) {
      throw new Error('Connection unavailable');
    }
    if (endpoint === 'presentation/mode') state.presentationMode = payload.mode;
    if (endpoint === 'template/select') { state.selectedTemplateId = payload.templateId; state.presentationMode = payload.mode; }
    return { ok: true, value: { status: 'ok', data: state } };
  } };
  client.apply({
    effect: run => run(),
    locale: { register: (_name, value) => { dictionary = value.en; return () => {}; } },
    get: () => ({ rpc }),
    slots: { inject: (_name, callback) => callback(), register: (options, component) => {
      slots.register(options, component);
      seats.set(options.name, { options, component, injected: options.inject('ui-test') });
    } },
  });
  const mount = createRoot(dom.window.document.getElementById('root'));
  unmount = () => React.act(async () => mount.unmount());
  const render = () => mount.render(React.createElement('div', { style: { display: 'flex', flexDirection: 'column' } },
    ...[...seats].map(([name, seat]) => React.createElement(seat.component, {
      key: name, ...seat.injected, sessionId: 'ui-test', useSession: select => select({ blank }),
      t: key => dictionary[key] ?? key,
    }))));
  await React.act(async () => render());
  const doc = dom.window.document;
  const button = doc.querySelector('button[data-desktop-ppt]');
  const panel = () => doc.querySelector('[data-office-ppt-template-panel]');
  const choice = () => [...doc.querySelectorAll('button[aria-label]')].find(node => node.getAttribute('aria-label') === definitions[0].name);
  assert(button, 'Toolbar needs no owner props beyond the standard session hook');
  assert.equal(requests.filter(request => request.endpoint === 'state').length, 1);
  assert.equal(button.getAttribute('aria-pressed'), 'false');
  assert.equal(panel(), null);
  await React.act(async () => button.click());
  assert(panel());
  assert.equal(state.presentationMode, null, 'Browsing must not enable PPT mode');
  assert.equal(panel().dataset.placement, 'bottom');
  failSelect = true;
  await React.act(async () => choice().click());
  assert.match(doc.querySelector('[role="alert"]').textContent, /Connection unavailable/);
  assert.equal(state.presentationMode, null);
  assert(panel(), 'Failed selections keep the picker open');
  failSelect = false;
  await React.act(async () => choice().click());
  assert.equal(state.selectedTemplateId, definitions[0].id);
  assert.equal(state.presentationMode, 'ppt');
  assert.equal(panel(), null);
  assert.equal(doc.querySelector('img'), null, 'No persistent thumbnail remains');
  assert(button.textContent.includes(definitions[0].name));
  assert.equal(button.getAttribute('aria-pressed'), 'true');
  assert.equal(doc.activeElement, button);
  assert(requests.some(request => request.endpoint === 'template/select' && request.payload.sessionId === 'ui-test'));
  await React.act(async () => button.click());
  assert.equal(choice().getAttribute('aria-pressed'), 'true');
  assert(doc.querySelector('.dsh-ppt-selected'));
  await React.act(async () => panel().dispatchEvent(new dom.window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true })));
  assert.equal(panel(), null);
  assert.equal(state.presentationMode, 'ppt', 'Escape only closes the picker');
  await React.act(async () => button.click());
  const before = requests.length;
  await React.act(async () => choice().click());
  assert.equal(requests.length, before, 'Re-selecting the active template just closes the picker');
  assert.equal(state.presentationMode, 'ppt');
  await React.act(async () => button.click());
  failExit = true;
  await React.act(async () => doc.querySelector('.dsh-ppt-exit').click());
  assert.equal(state.presentationMode, 'ppt', 'A failed exit must retain the active mode');
  assert(panel());
  failExit = false;
  await React.act(async () => doc.querySelector('.dsh-ppt-exit').click());
  assert.equal(state.presentationMode, null);
  assert.equal(panel(), null);
  assert.equal(button.getAttribute('aria-pressed'), 'false');
  blank = false;
  await React.act(async () => render());
  assert.equal(doc.querySelector('button[data-desktop-ppt]'), null);
});

test('packaged previews are served locally and unknown paths stay inaccessible', async t => {
  const handler = createPreviewHandler(previewFiles, references);
  const server = createServer(handler);
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  t.after(() => new Promise(resolve => { server.closeAllConnections(); server.close(resolve); }));
  const base = `http://127.0.0.1:${server.address().port}/dsh-ppt/previews/`;
  const [hash, relative] = Object.entries(previewFiles)[0];
  const response = await fetch(`${base}${hash}.jpg`);
  assert.equal(response.status, 200);
  assert.deepEqual(Buffer.from(await response.arrayBuffer()), await readFile(path.join(references, relative)));
  assert.equal((await fetch(`${base}${hash}.jpg`, { headers: { 'If-None-Match': response.headers.get('etag') } })).status, 304);
  assert.equal((await fetch(`${base}${'f'.repeat(64)}.jpg`)).status, 404);
  assert.equal((await fetch(`${base}%2e%2e%2fpackage.json`)).status, 404);
  assert.equal((await fetch(`${base}${hash}.jpg`, { method: 'POST' })).status, 405);
});
