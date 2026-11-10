import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, readdir, rm, symlink, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Context } from '@deepseek-ai/cordis';
import { agentEvents } from '@deepseek-ai/dsh-agent';
import { createUserMessage } from '@deepseek-ai/dsh-llm';
import { Session } from '@deepseek-ai/dsh-session';
import { SkillRegistry } from '@deepseek-ai/dsh-skill';
import { SystemPrompt } from '@deepseek-ai/dsh-system-prompt';
import sharp from 'sharp';
import * as plugin from '../lib/index.js';
import * as preset from '../lib/preset.js';
import { normalizeTemplateAuthoring } from '../lib/template-authoring.js';

async function fixture(t) {
  const scratch = await mkdtemp(path.join(os.tmpdir(), 'dsh-template-authoring-'));
  const workspace = path.join(scratch, 'workspace');
  await mkdir(workspace);
  const ctx = new Context(), tools = new Map();
  let rpc;
  ctx.provide('connection', { rpc: { handle: (_channel, value) => { rpc = value; } } });
  ctx.provide('tools', { register: tool => tools.set(tool.name, tool) });
  ctx.provide('workingDirectory', { ensure: async () => workspace });
  ctx.provide('sandboxPolicy', { resolve: () => ({ mode: 'workspace-write', workspaceRoot: workspace }) });
  ctx.provide('webServer', { register: () => () => {} });
  const prompt = ctx.plugin(SystemPrompt, { includeHarnessIdentity: false });
  const skills = ctx.plugin(SkillRegistry);
  await prompt; await skills;
  const runtime = ctx.plugin(plugin, { root: path.join(scratch, 'storage') });
  await runtime;
  const child = ctx.plugin(preset); await child;
  t.after(async () => {
    await child.dispose(); await runtime.dispose(); await skills.dispose(); await prompt.dispose();
    await rm(scratch, { recursive: true, force: true });
  });
  const request = async (endpoint, payload = {}, sessionId = 'session-a') => {
    const result = await rpc(endpoint, { ...payload, sessionId });
    assert.equal(result.ok, true, JSON.stringify(result));
    if (result.value.status === 'error') throw new Error(result.value.error.message);
    return result.value.data;
  };
  const agents = new Map();
  const agentFor = id => {
    if (!agents.has(id)) agents.set(id, { id, session: Session.create(id, undefined, {
      version: 4, id, createdAt: Date.now(), cwd: workspace, isSeeded: false, agentPreset: 'ppt',
    }) });
    return agents.get(id);
  };
  const user = (text, id = 'session-a') => {
    const message = createUserMessage({ content: [{ type: 'text', text }], source: { kind: 'user' } });
    agentFor(id).session.append('user/message', message, { surfaceOp: 'append' });
    return message;
  };
  const run = (name, args, id = 'session-a') => tools.get(name).execute(args, {
    agent: agentFor(id), signal: new AbortController().signal,
  });
  const step = async (number = 1, id = 'session-a') => {
    const agent = agentFor(id);
    return agentEvents(ctx, agent).waterfall('agent/pre-step', {
      messages: agent.session.deriveMessages(), turn: 1, step: number, signal: new AbortController().signal,
    }, async () => ({ kind: 'enter', messages: [] }));
  };
  const project = async (name = 'template') => {
    await run('pptd_write_file', { project_path: name, file_path: 'deck.pptd', content: 'version: v2\ntitle: Reusable template\nsize: [960, 540]\npages: [pages/cover.page, pages/data.page]\n' });
    for (const [file, title] of [['cover', '报告标题'], ['data', '示例数据 42']]) {
      await run('pptd_write_file', { project_path: name, file_path: `pages/${file}.page`, content: JSON.stringify({
        background: { color: '#F4F2ED' }, notes: 'Illustrative placeholders. Replace with current task material.',
        elements: [{ elementId: 'title', elementType: 'text', bounds: [60, 50, 820, 90], content: { text: title, fontFamily: 'Arial', fontSize: 36, color: '#123456' } }],
      }) });
    }
    return name;
  };
  return { scratch, workspace, ctx, tools, request, run, user, step, agentFor, project, service: ctx.get(plugin.PPT_SERVICE) };
}

test('guided template preparation previews, saves an immutable native source, and reuses it', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await writeFile(path.join(f.workspace, 'template/private-notes.txt'), 'not part of the template');
  const asset = await sharp({ create: { width: 8, height: 8, channels: 3, background: '#12AABB' } }).png().toBuffer();
  await mkdir(path.join(f.workspace, 'template/assets'));
  await writeFile(path.join(f.workspace, 'template/assets/brand.png'), asset);
  const assetPagePath = path.join(f.workspace, 'template/pages/cover.page');
  const assetPage = JSON.parse(await readFile(assetPagePath, 'utf8'));
  assetPage.elements.push({ elementId: 'brand-image', elementType: 'image', bounds: [60, 180, 120, 120], src: 'assets/brand.png' });
  await writeFile(assetPagePath, JSON.stringify(assetPage));
  await f.request('template/select', { templateId: 'dsh-blue-professional', mode: 'ppt' });
  const start = await f.request('template/authoring-start');
  assert.equal((await f.request('template/authoring-start')).workflowId, start.workflowId);
  f.user('Blue business weekly report template');
  const initial = await f.step();
  assert(initial.messages.some(message => message.source.kind === 'skill-invocation' && message.source.name === 'dsh-ppt-template'));
  for (const message of initial.messages) f.agentFor('session-a').session.append('user/message', message, { surfaceOp: 'append' });
  const prepared = await f.run('ppt_template_prepare', { project_path: 'template', name: ` ${'W'.repeat(80)} `, description: 'Blue business layouts' });
  assert.equal(prepared.slideCount, 2);
  const review = await f.request('template/authoring-state');
  assert.notEqual(review, null, 'A valid trimmed template name must retain the review');
  assert.equal(review.phase, 'review');
  assert.equal(review.name, 'W'.repeat(80));
  assert.equal((await f.request('state')).decks.length, 0, 'Preparing a template must not consume a delivery slot');
  await f.step(2);
  assert.equal((await f.request('template/authoring-state')).phase, 'review', 'Tool pre-steps preserve review');
  await f.step();
  assert.equal((await f.request('template/authoring-state')).phase, 'review', 'Repeated assembly without a new real user message preserves review');
  const preview = await f.request('template/authoring-preview', { ...review, page: 2 });
  const info = await sharp(Buffer.from(preview.preview.split(',')[1], 'base64')).metadata();
  assert.equal(info.width, 1920); assert.equal(info.height, 1080);
  await assert.rejects(f.request('template/authoring-preview', { ...review, page: 3 }), /页码/);
  await assert.rejects(f.request('template/authoring-preview', { ...review, page: 1 }, 'other-session'), /版本已更新/);
  await assert.rejects(f.request('template/authoring-save', { ...review, revision: review.revision - 1, name: 'Stale' }), /版本已更新/);
  const originalPage = await readFile(path.join(f.workspace, 'template/pages/data.page'));
  const draftRoot = path.join(f.scratch, 'storage/personal-templates/drafts', review.draftId, 'project');
  assert.deepEqual(await readFile(path.join(draftRoot, 'pages/data.page')), originalPage);
  assert(!((await readdir(draftRoot)).includes('private-notes.txt')));
  assert.deepEqual(await readFile(path.join(draftRoot, 'assets/brand.png')), asset);
  await writeFile(path.join(f.workspace, 'template/pages/data.page'), 'Changed after preview');
  const saved = await f.request('template/authoring-save', { ...review, name: '我的周报模板' });
  assert.equal(saved.phase, 'saved');
  assert.equal((await f.request('template/authoring-save', { ...review, name: '我的周报模板' })).templateId, saved.templateId);
  const state = await f.request('state');
  assert.equal(state.selectedTemplateId, 'dsh-blue-professional', 'Saving does not change the locked selection');
  const record = state.templates.find(item => item.id === saved.templateId);
  assert.equal(record.authoringKind, 'native-pptd');
  assert.equal(record.templateSamples, undefined, 'Authored examples do not pass through heuristic replacement');
  assert.equal(record.description, 'Blue business layouts');
  assert.deepEqual(await readFile(path.join(f.scratch, 'storage/personal-templates/saved', saved.templateId, 'project/pages/data.page')), originalPage);
  await f.request('template/select', { templateId: saved.templateId, mode: 'ppt' }, 'session-b');
  await f.run('ppt_template_create_project', { template_id: saved.templateId, output_directory: 'reused' }, 'session-b');
  assert.deepEqual(await readFile(path.join(f.workspace, 'reused/pages/data.page')), originalPage);
  assert.deepEqual(await readFile(path.join(f.workspace, 'reused/assets/brand.png')), asset);
  await f.step();
  const active = f.agentFor('session-a').session.deriveMessages();
  assert(!active.some(message => message.source.kind === 'skill-invocation' && message.source.name === 'dsh-ppt-template'), 'Saved authoring retires the template-specific Skill');
  await f.request('template/authoring-exit', { workflowId: review.workflowId });
  assert.equal(await f.request('template/authoring-state'), null);
  assert((await f.request('state')).templates.some(item => item.id === saved.templateId));
});

test('an existing project prepares, previews and saves without the template entry button', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  f.user('Use this finished project as a reusable template');
  await f.step();
  assert.equal(await f.request('template/authoring-state'), null);
  const prepared = await f.run('ppt_template_prepare', { project_path: 'template', name: 'Existing project' });
  const review = await f.request('template/authoring-state');
  assert.equal(review.phase, 'review');
  assert.equal(prepared.projectPath, 'template');
  assert.equal(prepared.slideCount, 2);
  assert.equal((await f.request('state')).templates.some(item => item.id === review.templateId), false, 'Preparing previews does not save the template');
  for (let page = 1; page <= prepared.slideCount; page++) {
    const preview = await f.request('template/authoring-preview', { ...review, page });
    const info = await sharp(Buffer.from(preview.preview.split(',')[1], 'base64')).metadata();
    assert.equal(info.width, 1920);
    assert.equal(info.height, 1080);
  }
  const saved = await f.request('template/authoring-save', { ...review, name: 'Named by the user' });
  assert.equal(saved.phase, 'saved');
  await f.request('template/select', { templateId: saved.templateId, mode: 'ppt' }, 'session-b');
  await f.run('ppt_template_create_project', { template_id: saved.templateId, output_directory: 'reused' }, 'session-b');
  assert.deepEqual(await readFile(path.join(f.workspace, 'reused/pages/data.page')), await readFile(path.join(f.workspace, 'template/pages/data.page')));
  await f.project('another-template');
  await f.run('ppt_template_prepare', { project_path: 'another-template', name: 'Another template' });
  const next = await f.request('template/authoring-state');
  assert.equal(next.phase, 'review');
  assert.notEqual(next.workflowId, saved.workflowId, 'A saved workflow admits a new template without the entry button');
  await assert.rejects(f.request('template/authoring-save', { ...review, name: 'Old preview' }), /版本已更新/);
  await f.request('template/authoring-exit', { workflowId: saved.workflowId });
  assert.equal((await f.request('template/authoring-state')).workflowId, next.workflowId, 'A stale exit cannot cancel the newly admitted workflow');
  const record = (await f.request('state')).templates.find(item => item.id === saved.templateId);
  assert.equal(record.name, 'Named by the user', 'Preparing another template preserves the saved template');
});

test('cancelled direct prepares do not admit or publish review', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  const controller = new AbortController();
  controller.abort();
  const execute = signal => f.tools.get('ppt_template_prepare').execute({ project_path: 'template', name: 'Cancelled' }, {
    agent: f.agentFor('session-a'), signal,
  });
  await assert.rejects(execute(controller.signal), /abort/i);
  assert.equal(await f.request('template/authoring-state'), null, 'A previously cancelled call must not begin template authoring');

  const active = new AbortController();
  const library = f.service.store.personalTemplates, native = library.prepareAuthored.bind(library);
  library.prepareAuthored = async (...args) => {
    const prepared = await native(...args);
    active.abort();
    return prepared;
  };
  await assert.rejects(execute(active.signal), /abort/i);
  assert.equal((await f.request('template/authoring-state')).phase, 'cancelled');
  assert.deepEqual(await readdir(path.join(f.scratch, 'storage/personal-templates/drafts')), [], 'Cancellation before publication discards the snapshot');
});

test('new user revisions invalidate reviewed snapshots and cancellation prevents late publication', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await f.request('template/authoring-start'); f.user('Create a template');
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Initial' });
  const review = await f.request('template/authoring-state');
  f.user('Make it green');
  await f.step(2);
  const creating = await f.request('template/authoring-state');
  assert.equal(creating.phase, 'creating'); assert.equal(creating.projectPath, 'template');
  assert(creating.revision > review.revision);
  await assert.rejects(f.request('template/authoring-save', { ...review, name: 'Old' }), /版本已更新/);
  const library = f.service.store.personalTemplates;
  const native = library.prepareAuthored.bind(library);
  let started;
  const waiting = new Promise(resolve => { started = resolve; });
  library.prepareAuthored = async (_sessionId, _input, signal) => {
    started();
    return new Promise((resolve, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true }));
  };
  const pending = f.run('ppt_template_prepare', { project_path: 'template', name: 'Cancelled' });
  const failure = assert.rejects(pending, /abort/i);
  await waiting;
  await f.request('template/authoring-exit', { workflowId: creating.workflowId });
  await failure;
  assert.equal(await f.request('template/authoring-state'), null);
  library.prepareAuthored = native;
  const restart = await f.request('template/authoring-start');
  assert.notEqual(restart.workflowId, creating.workflowId);
  await f.request('template/authoring-exit', { workflowId: creating.workflowId });
  assert.equal((await f.request('template/authoring-state')).workflowId, restart.workflowId, 'A stale exit cannot cancel a new workflow');
});

test('save retries recover after draft publication but failed session persistence', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project(); await f.request('template/authoring-start'); f.user('Weekly report');
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Recovery' });
  const review = await f.request('template/authoring-state');
  const store = f.service.store, writeState = store.writeState.bind(store);
  let failed = false;
  store.writeState = async state => {
    if (!failed && state.templateAuthoring?.phase === 'saved') { failed = true; throw new Error('simulated write failure'); }
    return writeState(state);
  };
  await assert.rejects(f.request('template/authoring-save', { ...review, name: 'Recovery' }), /simulated write failure/);
  assert.equal((await f.request('template/authoring-state')).phase, 'review');
  assert.equal((await f.request('template/authoring-preview', { ...review, page: 1 })).page, 1, 'Recovery review can still display the published image');
  const saved = await f.request('template/authoring-save', { ...review, name: 'Recovery' });
  assert.equal(saved.templateId, review.templateId);
  assert.equal((await f.request('state')).templates.filter(template => template.id === saved.templateId).length, 1);
});

test('native template boundaries reject invalid layouts, foreign paths and symbolic links', async t => {
  const f = await fixture(t);
  await f.project(); await f.request('template/authoring-start');
  await assert.rejects(f.run('ppt_template_prepare', { project_path: '../outside', name: 'No' }), /workspace/);
  await assert.rejects(f.run('ppt_template_prepare', { project_path: 'template', name: '\n' }), /名称/);
  await assert.rejects(f.run('ppt_template_prepare', { project_path: 'template', name: 'No', description: 'x'.repeat(1001) }), /描述/);
  await symlink('data.page', path.join(f.workspace, 'template/pages/cover-linked.page'));
  await writeFile(path.join(f.workspace, 'template/deck.pptd'), 'version: v2\nsize: [960, 540]\npages: [pages/cover-linked.page]\n');
  await assert.rejects(f.run('ppt_template_prepare', { project_path: 'template', name: 'No' }), /符号链接/);
  const drafts = path.join(f.scratch, 'storage/personal-templates/drafts');
  assert.deepEqual(await readdir(drafts), []);
});

test('persisted authoring state keeps only valid workflow fields', () => {
  const workflowId = '12345678-1234-1234-1234-123456789abc';
  assert.equal(normalizeTemplateAuthoring({ workflowId, revision: 1, phase: 'review' }), undefined);
  assert.equal(normalizeTemplateAuthoring({ workflowId, revision: -1, phase: 'creating' }), undefined);
  assert.deepEqual(normalizeTemplateAuthoring({ workflowId, revision: 1, phase: 'creating', guideSent: true, projectPath: '../escape', secret: 'drop' }),
    { workflowId, revision: 1, guideSent: true, phase: 'creating' });
});

test('each model step receives the current template and authoring context without duplicating Skills', async t => {
  const f = await fixture(t);
  f.user('Create slides with the selected template');
  const assemble = async number => {
    const decision = await f.step(number);
    for (const message of decision.messages) f.agentFor('session-a').session.append('user/message', message, { surfaceOp: 'append' });
    return f.agentFor('session-a').session.deriveMessages();
  };
  const snapshots = messages => messages.filter(message => message.source?.kind === 'dsh-ppt-composer');
  await f.request('template/select', { templateId: 'dsh-blue-professional', mode: 'ppt' });
  await assemble(1);
  let active = await assemble(2);
  assert.equal(snapshots(active).length, 1, 'A subsequent tool step must retain a current snapshot');
  assert.match(snapshots(active)[0].content[0].text, /selected_template_id: dsh-blue-professional/);
  await f.request('template/select', { templateId: 'dsh-broadside', mode: 'ppt' });
  await f.request('template/authoring-start');
  active = await assemble(3);
  assert.equal(snapshots(active).length, 1);
  assert.match(snapshots(active)[0].content[0].text, /selected_template_id: dsh-broadside/);
  assert.match(snapshots(active)[0].content[0].text, /template_authoring:/);
  assert.equal(active.filter(message => message.source?.kind === 'skill-invocation' && message.source.name === 'dsh-ppt').length, 1);
  const authoring = await f.request('template/authoring-state');
  await f.request('template/authoring-exit', { workflowId: authoring.workflowId });
  active = await assemble(4);
  assert.equal(snapshots(active).length, 1);
  assert.doesNotMatch(snapshots(active)[0].content[0].text, /template_authoring:/);
  assert(!active.some(message => message.source?.kind === 'skill-invocation' && message.source.name === 'dsh-ppt-template'));
});

test('preview failure is explicit, retry reuses the project, and audit failure cannot undo a saved template', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  const library = f.service.store.personalTemplates, prepare = library.prepareAuthored.bind(library);
  library.prepareAuthored = async () => { throw new Error('renderer unavailable'); };
  await assert.rejects(f.run('ppt_template_prepare', { project_path: 'template', name: 'Retry me' }), /renderer unavailable/);
  const failed = await f.request('template/authoring-state');
  assert.equal(failed.phase, 'failed');
  assert.equal(failed.projectPath, 'template');
  assert.equal(failed.error, 'renderer unavailable');
  const decision = await f.step(2);
  assert.match(decision.messages.at(-1).content[0].text, /renderer unavailable/);
  assert.equal(f.service.templateAuthoring.jobs.size, 0);
  library.prepareAuthored = prepare;
  await f.run('ppt_template_prepare', { project_path: failed.projectPath, name: failed.name });
  const review = await f.request('template/authoring-state');
  assert.equal(review.phase, 'review');
  assert.equal(review.workflowId, failed.workflowId);
  assert(review.revision > failed.revision);
  f.service.store.appendAudit = async () => { throw new Error('audit unavailable'); };
  const saved = await f.request('template/authoring-save', { ...review, name: 'Saved despite audit outage' });
  assert.equal(saved.phase, 'saved');
  assert.equal((await f.request('template/authoring-state')).phase, 'saved');
  assert.equal((await f.request('state')).templates.find(template => template.id === saved.templateId).name, saved.name);
  assert((await readdir(path.join(f.service.store.root, 'audit-pending'))).length > 0);
});

test('failed preparation admission leaves no live job and interrupted preparations offer retry', async t => {
  const f = await fixture(t);
  await f.project();
  const current = await f.request('template/authoring-start');
  const store = f.service.store, writeState = store.writeState.bind(store);
  store.writeState = async state => {
    if (state.templateAuthoring?.phase === 'preparing') throw new Error('reservation write failed');
    return writeState(state);
  };
  await assert.rejects(f.run('ppt_template_prepare', { project_path: 'template', name: 'Never admitted' }), /reservation write failed/);
  assert.equal(f.service.templateAuthoring.jobs.size, 0);
  assert.equal((await f.request('template/authoring-state')).phase, 'creating');
  store.writeState = writeState;
  await writeState({ ...await store.readState('session-a'), templateAuthoring: {
    ...current, revision: 1, phase: 'preparing', projectPath: 'template', name: 'Interrupted', guideSent: true,
  } });
  const interrupted = await f.request('template/authoring-state');
  assert.equal(interrupted.phase, 'failed');
  assert.match(interrupted.error, /中断/);
  const cancelled = await f.run('ppt_template_cancel', { workflow_id: interrupted.workflowId });
  assert.equal(cancelled.cancelled, true);
  assert.equal(await f.request('template/authoring-state'), null);
  assert.equal((await f.run('ppt_template_cancel', { workflow_id: interrupted.workflowId })).cancelled, false);
  assert((await readFile(path.join(f.workspace, 'template/deck.pptd'), 'utf8')).includes('version: v2'));
});

test('failed exit persistence keeps the reviewed snapshot available and conversation cancellation fences stale workflows', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Do not lose this preview' });
  const review = await f.request('template/authoring-state');
  const store = f.service.store, writeState = store.writeState.bind(store);
  store.writeState = async state => { if (!state.templateAuthoring) throw new Error('exit write failed'); return writeState(state); };
  await assert.rejects(f.run('ppt_template_cancel', { workflow_id: review.workflowId }), /exit write failed/);
  assert.equal((await f.request('template/authoring-preview', { ...review, page: 1 })).page, 1);
  store.writeState = writeState;
  assert.equal((await f.run('ppt_template_cancel', { workflow_id: review.workflowId })).cancelled, true);
  const next = await f.request('template/authoring-start');
  assert.equal((await f.run('ppt_template_cancel', { workflow_id: review.workflowId })).cancelled, false);
  assert.equal((await f.request('template/authoring-state')).workflowId, next.workflowId);
});

test('a missing unsaved draft cannot claim an unrelated saved template with identical content', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Original' });
  const first = await f.request('template/authoring-state');
  await f.request('template/authoring-save', { ...first, name: 'Original' });
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Another review' });
  const second = await f.request('template/authoring-state');
  assert.equal(second.templateId, first.templateId);
  await f.service.store.personalTemplates.cancel('session-a', second.draftId);
  await assert.rejects(f.request('template/authoring-save', { ...second, name: 'Should not save' }), /ENOENT/);
  await assert.rejects(f.request('template/authoring-preview', { ...second, page: 1 }), /ENOENT/);
  assert.equal((await f.request('template/authoring-state')).phase, 'review');
});

test('conversation cancellation remains responsive while a renderer holds the library processing queue', { timeout: 5000 }, async t => {
  const f = await fixture(t);
  await f.project();
  const current = await f.request('template/authoring-start');
  let enter;
  const entered = new Promise(resolve => { enter = resolve; });
  const library = f.service.store.personalTemplates;
  library.prepareAuthored = (_session, _input, signal) => library.locked(async () => {
    enter();
    await new Promise((resolve, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true }));
  });
  const pending = f.run('ppt_template_prepare', { project_path: 'template', name: 'Cancel busy renderer' });
  const rejected = assert.rejects(pending, /abort/i);
  await entered;
  assert.equal((await f.request('template/authoring-state')).phase, 'preparing', 'Reading state does not wait for rendering');
  assert.equal((await f.run('ppt_template_cancel', { workflow_id: current.workflowId })).cancelled, true);
  await rejected;
  assert.equal(await f.request('template/authoring-state'), null);
});

test('bilingual native templates remain listed and reusable after saving', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  const file = path.join(f.workspace, 'template/pages/cover.page');
  const page = JSON.parse(await readFile(file, 'utf8'));
  page.elements[0].content.fontFamily = { latin: 'Arial', ea: 'Microsoft YaHei', mac: 'Microsoft YaHei', win: 'Microsoft YaHei' };
  await writeFile(file, JSON.stringify(page));
  const original = await readFile(file);
  await f.run('ppt_template_prepare', { project_path: 'template', name: '双语品牌模板' });
  const review = await f.request('template/authoring-state');
  const saved = await f.request('template/authoring-save', { ...review, name: '双语品牌模板' });
  const brokenId = `personal-${'b'.repeat(64)}`;
  const brokenDirectory = path.join(f.service.store.personalTemplates.root, 'saved', brokenId);
  await mkdir(brokenDirectory);
  await writeFile(path.join(brokenDirectory, 'template.json'), '{broken');
  const state = await f.request('state');
  assert.equal(state.templates.filter(template => template.origin !== 'personal').length, 16, 'One broken custom record must not block the built-in catalog');
  assert(state.templates.some(template => template.id === saved.templateId), 'A saved bilingual template must remain in the catalog');
  const record = await f.service.store.personalTemplates.readRecord(saved.templateId);
  assert.equal(record.templateProfile.pages[0].slots[0].fontFace, 'Microsoft YaHei');
  await f.request('template/select', { templateId: saved.templateId, mode: 'ppt' }, 'session-b');
  await f.run('ppt_template_create_project', { template_id: saved.templateId, output_directory: 'reused-bilingual' }, 'session-b');
  assert.deepEqual(await readFile(path.join(f.workspace, 'reused-bilingual/pages/cover.page')), original);
});
