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
  let blank = true, running = false, agentPreset = 'ppt', sessionId = 'custom-template-test';
  const states = new Map();
  const initial = () => ({ templates: [...definitions], selectedTemplateId: null, presentationMode: 'ppt' });
  const stateFor = id => { if (!states.has(id)) states.set(id, initial()); return states.get(id); };
  const requests = [], sent = [];
  const composer = { text: 'Keep my existing draft', attachments: [{ name: 'brand.pptx' }] };
  const bindings = new Map();
  const bindingFor = id => { if (!bindings.has(id)) bindings.set(id, {}); return bindings.get(id); };
  dom.window.__ModuleLoader__ = { load: value => { entry = value; } };
  vm.runInNewContext(await readFile(new URL('../lib/client.js', import.meta.url), 'utf8'), {
    window: dom.window, document: dom.window.document, AbortController, Node: dom.window.Node,
    FileReader: dom.window.FileReader, DOMException: dom.window.DOMException, crypto,
    getComputedStyle: dom.window.getComputedStyle.bind(dom.window), performance, setTimeout, clearTimeout,
  });
  const rpc = { async call(channel, endpoint, payload) {
    assert.equal(channel, '/dsh-ppt');
    requests.push({ endpoint, payload });
    const state = stateFor(payload.sessionId);
    let data;
    if (handlers[endpoint]) data = await handlers[endpoint](payload, state);
    else if (endpoint === 'state') data = { ...state, templates: [...state.templates] };
    else if (endpoint === 'template/authoring-state') data = state.templateAuthoring ?? null;
    else if (endpoint === 'template/select') state.selectedTemplateId = payload.templateId;
    else if (endpoint === 'template/authoring-start') data = state.templateAuthoring ??= { phase: 'creating', workflowId: 'workflow-one', revision: 0, guideSent: false };
    else if (endpoint === 'template/authoring-guided') { state.templateAuthoring.guideSent = true; data = state.templateAuthoring; }
    else if (endpoint === 'template/authoring-exit') delete state.templateAuthoring;
    return { ok: true, value: { status: 'ok', data } };
  } };
  entry.factory(require).apply({
    effect: run => run(), locale: { register: (_name, value) => { dictionary = value.en; } }, get: () => ({ rpc }),
    sessions: { binding: bindingFor, scope: id => ({ get: name => {
      assert.equal(name, 'conversation');
      return { send: async text => { if (handlers.send) await handlers.send(text); sent.push({ sessionId: id, text }); } };
    } }) },
    slots: { inject: (_name, callback) => callback(), register: (options, component) => {
      assert.equal(options.name, 'conversation.input.dock');
      seat = { options, component, injected: options.inject(sessionId) };
    } },
  });
  const { createRoot } = require('react-dom/client');
  const mount = createRoot(dom.window.document.getElementById('root'));
  const render = () => mount.render(React.createElement(seat.component, {
    ...seat.injected, sessionId,
    useSession: select => select({ blank, running, pendingSubmissions: [] }),
    useSessions: select => select({ byId: { [sessionId]: { projectionValues: { agentPreset } } } }),
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
    for (let attempt = 0; attempt < 160; attempt++) {
      if (predicate()) return;
      await React.act(async () => delay(20));
    }
    assert.fail('The expected client state did not appear');
  }
  return { doc, get state() { return stateFor(sessionId); }, requests, click, until, sent, composer,
    async startConversation() { blank = false; await React.act(async () => render()); },
    async setRunning(value) { running = value; await React.act(async () => render()); },
    async setPreset(value) { agentPreset = value; await React.act(async () => render()); },
    async switchSession(value) { sessionId = value; seat.injected = seat.options.inject(sessionId); await React.act(async () => render()); },
    async reload() { await React.act(async () => mount.render(null)); await React.act(async () => render()); },
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

test('new template invokes the guide without changing composer text or attachments', async t => {
  const ui = await fixture(t);
  assert.equal(ui.doc.querySelector('input[type=file]'), null);
  await ui.click(ui.doc.querySelector('.personal-create'));
  assert.equal(ui.sent.length, 1);
  assert.match(ui.sent[0].text, /dsh-ppt-template/);
  assert.deepEqual(ui.composer, { text: 'Keep my existing draft', attachments: [{ name: 'brand.pptx' }] });
  assert(ui.doc.querySelector('[data-template-authoring=creating]'));
  assert.equal(ui.doc.querySelector('[data-template-authoring] button[type=submit]'), null);
  await ui.startConversation();
  assert(ui.doc.querySelector('[data-template-authoring=creating]'), 'The authoring panel remains after the first message');
  await ui.reload();
  assert(ui.doc.querySelector('[data-template-authoring=creating]'), 'Reload restores the workflow');
  assert.equal(ui.sent.length, 1, 'Reload must not send the guide twice');
  await ui.switchSession('other-session');
  assert.equal(ui.doc.querySelector('[data-template-authoring]'), null);
  await ui.switchSession('custom-template-test');
  assert(ui.doc.querySelector('[data-template-authoring=creating]'));
  assert.equal(ui.requests.some(item => /upload|template\/cancel/.test(item.endpoint)), false);
});

test('guide admission failure exposes a retry and resumes the same workflow', async t => {
  let attempts = 0;
  const ui = await fixture(t, { send: () => { if (++attempts === 1) throw new Error('Connection unavailable'); } });
  await ui.click(ui.doc.querySelector('.personal-create'));
  assert.match(ui.doc.querySelector('[role=alert]').textContent, /Connection unavailable/);
  assert.equal(ui.state.templateAuthoring.guideSent, false);
  await ui.click([...ui.doc.querySelectorAll('[data-template-authoring] button')].find(node => /Continue the AI/.test(node.textContent)));
  assert.equal(ui.sent.length, 1);
  assert.equal(ui.state.templateAuthoring.workflowId, 'workflow-one');
  assert.equal(ui.state.templateAuthoring.guideSent, true);
});

test('review previews every page, blocks save while running, and saves without changing selection', async t => {
  const ui = await fixture(t, {
    'template/authoring-preview': payload => ({ page: payload.page, preview: `data:image/png;base64,${Buffer.from('page'+payload.page).toString('base64')}` }),
    'template/authoring-save': (payload, state) => {
      assert.equal(payload.revision, 3); assert.equal(payload.draftId, 'draft-one');
      state.templateAuthoring = { ...state.templateAuthoring, phase: 'saved', name: payload.name };
      state.templates.push(personal);
      return state.templateAuthoring;
    },
  });
  ui.state.selectedTemplateId = definitions[0].id;
  ui.state.templateAuthoring = { phase: 'review', workflowId: 'workflow-one', revision: 3, draftId: 'draft-one', slideCount: 2, name: 'Weekly report', guideSent: true };
  await ui.startConversation(); await ui.reload();
  await ui.until(() => ui.doc.querySelector('.dsh-ppt-authoring-thumbnail img'));
  await ui.click(ui.doc.querySelector('.dsh-ppt-authoring-thumbnail'));
  assert(ui.doc.querySelector('.ppt-template-viewer').open);
  await ui.click(ui.doc.querySelector('button[aria-label="Next page"]'));
  await ui.until(() => ui.doc.querySelector('.ppt-template-viewer img')?.src.endsWith(Buffer.from('page2').toString('base64')));
  await ui.click(ui.doc.querySelector('.ppt-template-viewer button[aria-label=Close]'));
  await ui.setRunning(true);
  assert(ui.doc.querySelector('[data-template-authoring] button[type=submit]').disabled);
  await ui.setRunning(false);
  await ui.click(ui.doc.querySelector('[data-template-authoring] button[type=submit]'));
  assert(ui.doc.querySelector('[data-template-authoring=saved]'));
  assert.equal(ui.state.selectedTemplateId, definitions[0].id);
  assert.equal(ui.requests.some(item => item.endpoint === 'template/select'), false);
  await ui.click(ui.doc.querySelector('[data-template-authoring] header button'));
  assert.equal(ui.doc.querySelector('[data-template-authoring]'), null);
});

test('saved library retains preview, editing, deletion and ordinary first-message behavior', async t => {
  const ui = await fixture(t, {
    'template/preview-saved-page': () => ({ preview: personal.previewImages[0] }),
    'template/update': (payload, state) => { state.templates = state.templates.map(item => item.id === payload.templateId ? { ...item, description: payload.description } : item); },
    'template/delete': (payload, state) => { state.templates = state.templates.filter(item => item.id !== payload.templateId); },
  });
  ui.state.templates.push(personal); await ui.reload();
  const { doc, click } = ui;
  await click(doc.querySelector('[data-personal-card] .ppt-preview-open'));
  await ui.until(() => doc.querySelector('.ppt-template-viewer img'));
  await click(doc.querySelector('.ppt-template-viewer button[aria-label=Close]'));
  await click(doc.querySelector('[data-personal-card] button[aria-label="Edit template"]'));
  const description = doc.querySelector('.personal-dialog textarea');
  const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
  await React.act(async () => { setter.call(description, 'Updated use case'); description.dispatchEvent(new window.Event('input', { bubbles: true })); });
  await click(doc.querySelector('.personal-dialog button[type=submit]'));
  assert(ui.requests.some(item => item.endpoint === 'template/update' && item.payload.description === 'Updated use case'));
  await click(doc.querySelector('[data-personal-card] button[aria-label=Delete]'));
  await click(doc.querySelector('.personal-dialog .personal-danger'));
  assert.equal(doc.querySelector('[data-personal-card]'), null);
  await ui.startConversation();
  assert.equal(doc.querySelector('[data-office-ppt-template-panel]'), null);
  assert.equal(doc.querySelector('[data-template-authoring]'), null);
});
