import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { setTimeout as delay } from 'node:timers/promises';
import test from 'node:test';
import vm from 'node:vm';
import { JSDOM } from 'jsdom';
import { definitions } from '../lib/catalog.js';

const require = createRequire(import.meta.url);
const React = require('react');
const personal = {
  ...definitions[0], id: 'personal-imported', origin: 'personal', name: 'Imported slides',
  slideCount: 2, createdAt: '2026-11-09T00:00:00Z', previewImages: ['data:image/png;base64,aA=='],
};

async function fixture(t, handlers = {}) {
  const dom = new JSDOM('<!doctype html><div id="root"></div>', { pretendToBeVisual: true, url: 'http://localhost' });
  const globals = ['window', 'document', 'IS_REACT_ACT_ENVIRONMENT'].map(key => [key, Object.getOwnPropertyDescriptor(globalThis, key)]);
  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  dom.window.HTMLDialogElement.prototype.showModal = function () { this.open = true; };
  dom.window.HTMLDialogElement.prototype.close = function () { this.open = false; };
  let entry, seat, dictionary;
  let blank = true;
  let agentPreset = 'ppt';
  const state = { templates: [...definitions], selectedTemplateId: null, presentationMode: 'ppt' };
  const requests = [];
  dom.window.__ModuleLoader__ = { load: value => { entry = value; } };
  vm.runInNewContext(await readFile(new URL('../lib/client.js', import.meta.url), 'utf8'), {
    window: dom.window, document: dom.window.document, AbortController, Node: dom.window.Node,
    FileReader: dom.window.FileReader, DOMException: dom.window.DOMException, crypto,
    getComputedStyle: dom.window.getComputedStyle.bind(dom.window), performance, setTimeout, clearTimeout,
  });
  const rpc = { async call(channel, endpoint, payload) {
    assert.equal(channel, '/dsh-ppt');
    assert.equal(payload.sessionId, 'custom-template-test');
    requests.push({ endpoint, payload });
    let data;
    if (handlers[endpoint]) data = await handlers[endpoint](payload, state);
    else if (endpoint === 'state') data = { ...state, templates: [...state.templates] };
    else if (endpoint === 'template/select') state.selectedTemplateId = payload.templateId;
    return { ok: true, value: { status: 'ok', data } };
  } };
  entry.factory(require).apply({
    effect: run => run(), locale: { register: (_name, value) => { dictionary = value.en; } }, get: () => ({ rpc }),
    slots: { inject: (_name, callback) => callback(), register: (options, component) => {
      assert.equal(options.name, 'conversation.input.dock');
      seat = { component, injected: options.inject('custom-template-test') };
    } },
  });
  const { createRoot } = require('react-dom/client');
  const mount = createRoot(dom.window.document.getElementById('root'));
  const render = () => mount.render(React.createElement(seat.component, {
    ...seat.injected, sessionId: 'custom-template-test',
    useSession: select => select({ blank }),
    useSessions: select => select({ byId: { 'custom-template-test': { projectionValues: { agentPreset } } } }),
    t: key => dictionary[key] ?? key,
  }));
  t.after(async () => {
    await React.act(async () => mount.unmount());
    dom.window.close();
    for (const [key, descriptor] of globals) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else delete globalThis[key];
    }
  });
  await React.act(async () => render());
  const doc = dom.window.document;
  async function click(node) { assert(node, 'Expected a button'); await React.act(async () => node.click()); }
  async function until(predicate) {
    for (let attempt = 0; attempt < 120; attempt++) {
      if (predicate()) return;
      await React.act(async () => delay(20));
    }
    assert.fail('The expected client state did not appear');
  }
  async function upload(name = 'example.pptx') {
    const input = doc.querySelector('input[type=file]');
    Object.defineProperty(input, 'files', { configurable: true, value: [new dom.window.File(['abcdef'], name)] });
    await React.act(async () => input.dispatchEvent(new dom.window.Event('change', { bubbles: true })));
  }
  return { doc, state, requests, click, until, upload,
    async startConversation() { blank = false; await React.act(async () => render()); },
    async setPreset(value) { agentPreset = value; await React.act(async () => render()); },
  };
}

test('reopening the picker refreshes profile templates changed in another conversation', async t => {
  const ui = await fixture(t);
  assert.equal(ui.doc.querySelector('[data-personal-card]'), null);
  await ui.setPreset('standard');
  ui.state.templates.push(personal);
  await ui.setPreset('ppt');
  assert(ui.doc.querySelector('[data-personal-card="personal-imported"]'));
  await ui.setPreset('standard');
  ui.state.templates = [...definitions];
  await ui.setPreset('ppt');
  assert.equal(ui.doc.querySelector('[data-personal-card]'), null);
  assert.equal(ui.requests.filter(item => item.endpoint === 'state').length, 3);
});

test('custom template import, preview, save, edit and delete stay inside the native picker', async t => {
  const chunks = [];
  const ui = await fixture(t, {
    'template/upload-start': () => ({ uploadId: 'upload-one', chunkBytes: 2 }),
    'template/upload-chunk': payload => { chunks.push(payload); },
    'template/import-progress': () => ({ total: 2, completed: 0, pages: [] }),
    'template/prepare': () => ({ draftId: 'draft-one', template: personal, preview: personal.previewImages[0] }),
    'template/preview-page': () => ({ preview: 'data:image/png;base64,cGFnZTI=' }),
    'template/preview-saved-page': () => ({ preview: personal.previewImages[0] }),
    'template/save': (_payload, state) => { state.templates.push(personal); return personal; },
    'template/update': (payload, state) => { state.templates = state.templates.map(item => item.id === payload.templateId ? { ...item, description: payload.description } : item); },
    'template/delete': (payload, state) => {
      state.templates = state.templates.filter(item => item.id !== payload.templateId);
      state.selectedTemplateId = definitions[0].id;
      state.templateMigration = { reason: 'personal-template-deleted' };
    },
  });
  const { doc, requests, click, until, upload } = ui;
  assert.equal(requests.filter(item => item.endpoint === 'state').length, 1);
  assert.equal(doc.querySelectorAll('.ppt-preview-open').length, 0, 'Built-in cards keep the existing interaction');
  await click(doc.querySelector('[data-ppt-category=personal]'));
  assert.match(doc.querySelector('.personal-create').textContent, /Create custom template/);
  await upload();
  await until(() => doc.querySelector('.personal-dialog input[required]'));
  assert.deepEqual(chunks.map(item => item.offset), [0, 2, 4]);
  assert.equal(Buffer.concat(chunks.map(item => Buffer.from(item.base64, 'base64'))).toString(), 'abcdef');
  await click(doc.querySelector('.personal-pagination button[aria-label="Next page"]'));
  await until(() => doc.querySelector('.personal-zoom-stage img')?.src.endsWith('cGFnZTI='));
  await click(doc.querySelector('.personal-dialog button[type=submit]'));
  await until(() => doc.querySelector('[data-personal-card="personal-imported"]'));
  assert.equal(ui.state.selectedTemplateId, personal.id);
  assert.equal(doc.querySelector('[data-personal-card] button[aria-pressed]').getAttribute('aria-pressed'), 'true');
  assert.equal(doc.querySelector('.personal-dialog').open, false);
  await click(doc.querySelector('[data-personal-card] .ppt-preview-open'));
  await until(() => doc.querySelector('.ppt-template-viewer img'));
  assert(requests.some(item => item.endpoint === 'template/preview-saved-page' && item.payload.page === 1));
  await click(doc.querySelector('.ppt-template-viewer button[aria-label=Close]'));
  await click(doc.querySelector('[data-personal-card] button[aria-label="Edit template"]'));
  const description = doc.querySelector('.personal-dialog textarea');
  const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
  await React.act(async () => {
    setter.call(description, 'Updated use case');
    description.dispatchEvent(new window.Event('input', { bubbles: true }));
  });
  await click(doc.querySelector('.personal-dialog button[type=submit]'));
  assert(requests.some(item => item.endpoint === 'template/update' && item.payload.description === 'Updated use case'));
  await click(doc.querySelector('[data-personal-card] button[aria-label=Delete]'));
  await click(doc.querySelector('.personal-dialog .personal-danger'));
  assert.equal(doc.querySelector('[data-personal-card]'), null);
  assert.equal(ui.state.selectedTemplateId, definitions[0].id);
  assert.match(doc.querySelector('.dsh-ppt-panel-notice').textContent, /was deleted/);
  assert.equal(doc.querySelector('button[data-desktop-ppt]'), null);
  await ui.startConversation();
  assert.equal(doc.querySelector('[data-office-ppt-template-panel]'), null);
});

test('starting the conversation cancels a late custom template upload', async t => {
  let finishStart;
  const pending = new Promise(resolve => { finishStart = resolve; });
  const ui = await fixture(t, { 'template/upload-start': () => pending });
  await ui.upload();
  await ui.until(() => ui.requests.some(item => item.endpoint === 'template/upload-start'));
  await ui.startConversation();
  await React.act(async () => finishStart({ uploadId: 'late-upload', chunkBytes: 2 }));
  await ui.until(() => ui.requests.some(item => item.endpoint === 'template/upload-cancel'));
  assert.equal(ui.requests.find(item => item.endpoint === 'template/upload-cancel').payload.uploadId, 'late-upload');
  assert.equal(ui.requests.some(item => item.endpoint === 'template/prepare'), false);
  assert.equal(ui.doc.querySelector('[data-office-ppt-template-panel]'), null);
});

test('invalid custom input can be retried without saving a template', async t => {
  const ui = await fixture(t);
  await ui.upload('notes.txt');
  assert.match(ui.doc.querySelector('[role=alert]').textContent, /Choose a PPTX file/);
  assert.equal(ui.requests.some(item => item.endpoint === 'template/upload-start'), false);
  await ui.click(ui.doc.querySelector('.personal-dialog button[aria-label=Close]'));
  assert.equal(ui.doc.querySelector('.personal-dialog').open, false);
  assert.equal(ui.doc.querySelector('.personal-create').disabled, false);
});
