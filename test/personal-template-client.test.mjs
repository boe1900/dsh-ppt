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
  assert.doesNotMatch(ui.sent[0].text, /dsh-ppt-template|skill/i);
  assert.match(ui.sent[0].text, /purpose, style or branding/);
  assert.match(ui.sent[0].text, /preview card beside the input.*view every page, enter a name, and click Save to custom templates/);
  assert.match(ui.sent[0].text, /Custom in a new PPT conversation/);
  assert.deepEqual(ui.composer, { text: 'Keep my existing draft', attachments: [{ name: 'brand.pptx' }] });
  assert.equal(ui.doc.querySelector('[data-template-authoring]'), null);
  assert.equal(Boolean(ui.doc.querySelector('button')), false, 'Creating templates must leave no standalone controls');
  assert.equal(ui.doc.querySelector('form'), null);
  await ui.startConversation();
  assert.equal(Boolean(ui.doc.querySelector('button')), false, 'The first message must not add authoring controls');
  await ui.setRunning(true);
  assert.equal(Boolean(ui.doc.querySelector('button')), false, 'Generation must not add authoring controls');
  await ui.setRunning(false);
  await ui.reload();
  assert.equal(Boolean(ui.doc.querySelector('button')), false, 'Reload must keep the creating conversation clean');
  assert.equal(ui.doc.querySelector('[data-template-authoring]'), null);
  assert.equal(ui.sent.length, 1, 'Reload must not send the guide twice');
  await ui.switchSession('other-session');
  assert.equal(ui.doc.querySelector('[data-template-authoring-actions]'), null);
  await ui.switchSession('custom-template-test');
  assert.equal(Boolean(ui.doc.querySelector('button')), false, 'Returning to the workflow must not restore authoring controls');
  assert.equal(ui.state.templateAuthoring.workflowId, 'workflow-one');
  assert.equal(ui.state.templateAuthoring.guideSent, true);
  assert.equal(ui.requests.some(item => /upload|template\/cancel/.test(item.endpoint)), false);
  assert.equal(ui.requests.some(item => item.endpoint === 'template/authoring-exit'), false);
});

test('guide admission failure keeps its explanation and retry together across reload and navigation', async t => {
  let attempts = 0;
  const ui = await fixture(t, { send: () => { if (++attempts < 3) throw new Error('Connection unavailable'); } });
  await ui.click(ui.doc.querySelector('.personal-create'));
  let alert = ui.doc.querySelector('[role=alert]');
  assert.match(alert.textContent, /Connection unavailable/);
  assert(alert.querySelector('button'), 'Guide admission errors must include their retry action');
  assert.equal(ui.doc.querySelectorAll('button').length, 1, 'Only the contextual retry is shown');
  assert.equal(ui.state.templateAuthoring.guideSent, false);
  assert.equal(ui.doc.querySelector('[data-template-authoring]'), null);
  await ui.reload();
  assert.equal(attempts, 1, 'Reload must leave guide retry to the user');
  alert = ui.doc.querySelector('[role=alert]');
  assert(alert, 'An unsent guide must explain the recovery after reload');
  let retry = alert.querySelector('button');
  assert(retry);
  assert.match(alert.textContent.replace(retry.textContent, ''), /template guide/i);
  await ui.switchSession('other-session');
  assert.equal(ui.doc.querySelector('[role=alert]'), null);
  await ui.switchSession('custom-template-test');
  alert = ui.doc.querySelector('[role=alert]');
  assert(alert, 'Returning to an unsent guide must restore contextual recovery');
  retry = alert.querySelector('button');
  assert(retry);
  assert.match(alert.textContent.replace(retry.textContent, ''), /template guide/i);
  await ui.setRunning(true);
  assert(retry.disabled);
  await ui.setRunning(false);
  await ui.click(retry);
  assert.match(ui.doc.querySelector('[role=alert]').textContent, /Connection unavailable/);
  await ui.click(ui.doc.querySelector('[role=alert] button'));
  assert.equal(ui.sent.length, 1);
  assert.equal(ui.state.templateAuthoring.workflowId, 'workflow-one');
  assert.equal(ui.state.templateAuthoring.guideSent, true);
  assert.equal(ui.doc.querySelector('[role=alert]'), null);
  assert.equal(Boolean(ui.doc.querySelector('button')), false, 'Successful recovery must remove authoring controls');
  await ui.reload();
  assert.equal(Boolean(ui.doc.querySelector('button')), false);
  assert.equal(attempts, 3);
});

test('a creating workflow does not add controls when polling fails', async t => {
  const ui = await fixture(t, { 'template/authoring-state': () => { throw new Error('Status unavailable'); } });
  ui.state.templateAuthoring = { phase: 'creating', workflowId: 'workflow-one', revision: 2, guideSent: true };
  await ui.startConversation(); await ui.reload();
  await ui.until(() => ui.doc.querySelector('[role=alert]')?.textContent.includes('Status unavailable'));
  assert.equal(Boolean(ui.doc.querySelector('button')), false, 'A status error must not restore authoring controls');
  assert.equal(ui.state.templateAuthoring.guideSent, true);
  assert.equal(ui.sent.length, 0);
});

test('polling failures keep an in-flight template guide free of recovery controls', async t => {
  let release, polls = 0;
  const ui = await fixture(t, {
    send: () => new Promise(resolve => { release = resolve; }),
    'template/authoring-state': () => { polls++; throw new Error('Status unavailable'); },
  });
  t.after(() => release?.());
  await ui.click(ui.doc.querySelector('.personal-create'));
  await ui.until(() => polls > 0);
  assert.equal(Boolean(ui.doc.querySelector('[role=alert]')), false, 'Polling must not expose recovery while the guide is still being sent');
  assert.equal(Boolean(ui.doc.querySelector('button')), false);
  release();
  await ui.until(() => ui.state.templateAuthoring.guideSent);
  assert.equal(ui.doc.querySelector('[role=alert]'), null);
  assert.equal(ui.sent.length, 1);
});

test('review previews every page, blocks save while running, and saves without changing selection', async t => {
  let finishSave;
  const ui = await fixture(t, {
    'template/authoring-preview': payload => ({ page: payload.page, preview: `data:image/png;base64,${Buffer.from('page'+payload.page).toString('base64')}` }),
    'template/authoring-save': async (payload, state) => {
      assert.equal(payload.revision, 3); assert.equal(payload.draftId, 'draft-one');
      await new Promise(resolve => { finishSave = resolve; });
      state.templateAuthoring = { ...state.templateAuthoring, phase: 'saved', name: payload.name };
      state.templates.push(personal);
      return state.templateAuthoring;
    },
  });
  ui.state.selectedTemplateId = definitions[0].id;
  ui.state.templateAuthoring = { phase: 'review', workflowId: 'workflow-one', revision: 3, draftId: 'draft-one', slideCount: 2, name: 'Weekly report', guideSent: true };
  await ui.startConversation(); await ui.reload();
  assert.equal(ui.doc.querySelector('[data-template-authoring=review] strong').textContent, 'Template preview');
  assert.match(ui.doc.querySelector('[data-template-authoring]').textContent, /2 template pages/);
  assert.match(ui.doc.querySelector('[data-template-authoring]').textContent, /tell AI what to change in the conversation/);
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
  assert(ui.doc.querySelector('[data-template-authoring] button[type=submit]').disabled);
  assert(ui.doc.querySelector('[data-template-authoring] input').disabled);
  assert(ui.doc.querySelector('[data-template-authoring] header button').disabled);
  await React.act(async () => finishSave());
  assert.equal(ui.doc.querySelector('[data-template-authoring]'), null);
  assert.match(ui.doc.querySelector('[data-template-authoring-feedback] [role=status]').textContent, /Weekly report.*saved.*new PPT conversation.*Custom/);
  assert.equal(ui.state.selectedTemplateId, definitions[0].id);
  assert.equal(ui.requests.some(item => item.endpoint === 'template/select'), false);
  await ui.click(ui.doc.querySelector('[data-template-authoring-feedback] button[aria-label=Close]'));
  assert.equal(ui.doc.querySelector('[data-template-authoring-feedback]'), null);
  await ui.reload();
  assert.equal(ui.doc.querySelector('[data-template-authoring-feedback]'), null, 'Reload does not replay save feedback');
  assert.equal(ui.doc.querySelector('[data-template-authoring]'), null);
});

test('saved feedback disappears on navigation and persisted saved workflows stay quiet', async t => {
  const ui = await fixture(t, {
    'template/authoring-preview': () => ({ preview: personal.previewImages[0] }),
    'template/authoring-save': (payload, state) => state.templateAuthoring = { ...state.templateAuthoring, phase: 'saved', name: payload.name },
  });
  ui.state.templateAuthoring = { phase: 'review', workflowId: 'workflow-one', revision: 1, draftId: 'draft-one', slideCount: 2, name: 'Brand slides', guideSent: true };
  await ui.startConversation(); await ui.reload();
  await ui.until(() => ui.doc.querySelector('.dsh-ppt-authoring-thumbnail img'));
  await ui.click(ui.doc.querySelector('[data-template-authoring] button[type=submit]'));
  assert(ui.doc.querySelector('[data-template-authoring-feedback]'));
  await ui.switchSession('other-session');
  assert.equal(ui.doc.querySelector('[data-template-authoring-feedback]'), null);
  await ui.switchSession('custom-template-test');
  assert.equal(ui.doc.querySelector('[data-template-authoring-feedback]'), null);
  assert.equal(ui.doc.querySelector('[data-template-authoring]'), null);
  assert.equal(ui.sent.length, 0);
});

test('review keeps preview retries and explicit exit while creating hides the result card', async t => {
  let attempts = 0;
  const ui = await fixture(t, {
    'template/authoring-preview': payload => {
      assert.equal(payload.revision, 4); assert.equal(payload.draftId, 'draft-four');
      if (++attempts === 1) throw new Error('Preview unavailable');
      return { preview: personal.previewImages[0] };
    },
  });
  ui.state.templateAuthoring = { phase: 'review', workflowId: 'workflow-one', revision: 4, draftId: 'draft-four', slideCount: 2, name: 'Brand slides', guideSent: true };
  await ui.startConversation(); await ui.reload();
  assert.match(ui.doc.querySelector('[role=alert]').textContent, /Preview unavailable/);
  assert(ui.doc.querySelector('[data-template-authoring] button[type=submit]').disabled);
  await ui.click(ui.doc.querySelector('[role=alert] button'));
  await ui.until(() => ui.doc.querySelector('.dsh-ppt-authoring-thumbnail img'));
  assert.equal(attempts, 2);
  assert.equal(ui.doc.querySelector('[data-template-authoring] button[type=submit]').disabled, false);
  ui.state.templateAuthoring = { ...ui.state.templateAuthoring, phase: 'creating', revision: 5 };
  await ui.reload();
  assert.equal(ui.doc.querySelector('[data-template-authoring]'), null);
  assert.equal(Boolean(ui.doc.querySelector('button')), false, 'Revising a template must return to the clean conversation');
  ui.state.templateAuthoring = { ...ui.state.templateAuthoring, phase: 'review', revision: 4 };
  await ui.reload();
  await ui.click(ui.doc.querySelector('[data-template-authoring] header button'));
  assert.equal(ui.state.templateAuthoring, undefined);
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
