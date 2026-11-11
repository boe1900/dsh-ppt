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
import { AuthoredTemplateLibrary } from '../lib/authored-templates.js';
import { MAX_TEMPLATE_DESIGN_BYTES, parseTemplateDesign } from '../lib/template-design.js';

test('a content update between source capture and checkout cannot publish a mixed template copy', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Original two-page template' });
  const first = await f.request('template/authoring-state');
  const saved = await f.request('template/authoring-save', { ...first, name: 'Original two-page template' });
  await f.request('template/select', { templateId: saved.templateId, mode: 'ppt' }, 'reader');
  const design = designSpec(); design.layouts = design.layouts.slice(0, 1);
  await writeFile(path.join(f.workspace, 'template/deck.pptd'), 'version: v2\ntitle: Updated one-page template\nsize: [960, 540]\npages: [pages/cover.page]\n');
  await writeFile(path.join(f.workspace, 'template/template-design.json'), JSON.stringify(design));
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Updated one-page template' });
  const review = await f.request('template/authoring-state');
  const library = f.service.store.personalTemplates, readSnapshot = library.readSnapshot.bind(library);
  let injected = false;
  library.readSnapshot = async id => {
    const snapshot = await readSnapshot(id);
    if (!injected && id === saved.templateId) {
      injected = true;
      await f.request('template/authoring-save', { ...review, name: 'Updated one-page template' });
    }
    return snapshot;
  };
  await assert.rejects(f.run('ppt_template_create_project', { template_id: saved.templateId, output_directory: 'mixed-copy' }, 'reader'), /原模板已更新/);
  assert(injected, 'The content update must occur after the old source snapshot was captured');
  assert(!(await readdir(f.workspace)).includes('mixed-copy'), 'A stale checkout must not leave a mixed published directory');
  library.readSnapshot = readSnapshot;
  const copied = await f.run('ppt_template_create_project', { template_id: saved.templateId, output_directory: 'current-copy' }, 'reader');
  assert.equal(copied.slideCount, 1);
  assert.match(await readFile(path.join(f.workspace, 'current-copy/deck.pptd'), 'utf8'), /Updated one-page template/);
});

test('page and reference reads retain one complete revision across a concurrent template update', { timeout: 180000 }, async t => {
  for (const tool of ['ppt_get_template_pages', 'ppt_get_template_reference']) {
    const f = await fixture(t);
    await f.project();
    await f.run('ppt_template_prepare', { project_path: 'template', name: 'Original two-page template' });
    const first = await f.request('template/authoring-state');
    const saved = await f.request('template/authoring-save', { ...first, name: 'Original two-page template' });
    await f.request('template/select', { templateId: saved.templateId, mode: 'ppt' }, 'reader');
    const design = designSpec(); design.layouts = [{ ...design.layouts[0], name: 'Updated single cover' }];
    await writeFile(path.join(f.workspace, 'template/deck.pptd'), 'version: v2\ntitle: Updated one-page template\nsize: [960, 540]\npages: [pages/cover.page]\n');
    await writeFile(path.join(f.workspace, 'template/template-design.json'), JSON.stringify(design));
    await f.run('ppt_template_prepare', { project_path: 'template', name: 'Updated one-page template' });
    const review = await f.request('template/authoring-state');
    const library = f.service.store.personalTemplates, readSnapshot = library.readSnapshot.bind(library);
    let injected = false;
    library.readSnapshot = async id => {
      const snapshot = await readSnapshot(id);
      if (!injected && id === saved.templateId) {
        injected = true;
        await f.request('template/authoring-save', { ...review, name: 'Updated one-page template' });
      }
      return snapshot;
    };
    const args = { template_id: saved.templateId, ...(tool === 'ppt_get_template_pages' ? { slide_numbers: [1, 2] } : {}) };
    const result = await f.run(tool, args, 'reader');
    assert(injected, 'The concurrent update must occur after snapshot resolution');
    const pages = tool === 'ppt_get_template_pages' ? result : result.pageContracts;
    assert.equal(pages.length, 2, tool);
    assert.equal(pages[1].sourceTitle, '指标结论', tool);
    if (tool === 'ppt_get_template_pages') assert.match(pages[1].pptdLayoutReference, /示例数据 42/);
    else assert.equal(result.templateName, 'Original two-page template');
    library.readSnapshot = readSnapshot;
    const current = await f.run(tool, { template_id: saved.templateId }, 'reader');
    const currentPages = tool === 'ppt_get_template_pages' ? current : current.pageContracts;
    assert.equal(currentPages.length, 1, tool);
    assert.equal(currentPages[0].sourceTitle, 'Updated single cover', tool);
  }
});

test('a deleted legacy saved template cannot block the next chat turn or cancellation', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Legacy saved template' });
  const review = await f.request('template/authoring-state');
  const saved = await f.request('template/authoring-save', { ...review, name: 'Legacy saved template' });
  const state = await f.service.store.readState('session-a');
  await f.service.store.writeState({ ...state, templateAuthoring: {
    phase: 'saved', workflowId: saved.workflowId, revision: saved.revision, guideSent: true,
    templateId: saved.templateId, name: saved.name,
  } });
  await f.service.store.personalTemplates.remove('other', saved.templateId);
  f.user('退出模板制作，改为普通 PPT');
  const step = await f.step();
  assert.equal(step.kind, 'enter');
  const authoring = await f.request('template/authoring-state');
  assert.equal(authoring.templateId, saved.templateId, 'Keep the missing legacy identity until an explicit recovery or cancellation');
  assert.deepEqual(authoring.origin, { kind: 'new' }, 'A deleted entry must not gain a fabricated version binding');
  assert.deepEqual(await f.run('ppt_template_cancel', { workflow_id: saved.workflowId }), { cancelled: true });
  assert.equal(await f.request('template/authoring-state'), null);
});

function designSpec() {
  return {
    version: 1, purpose: '技术团队向管理层汇报进度与风险',
    rules: {
      palette: '米白底 #F4F2ED；深蓝 #123456 正文；琥珀色仅标记风险。',
      typography: '标题 36 pt，正文 20 pt；中文用微软雅黑；最低 18 pt。',
      spacing: '960 × 540 pt；左右留白 60 pt；同级内容间距 24 pt。',
      imagery: '保留品牌图标，业务配图替换为经核实且可用的素材。',
      data: '图表保留单位、时间范围和来源；示例数据必须替换。',
      overflow: '长标题允许两行；条目过多拆页，不缩小字号。',
    },
    layouts: [
      { file: 'pages/cover.page', name: '汇报封面', purpose: '说明主题与汇报人', relationship: '主题与署名', capacity: '标题最多 24 个汉字，署名一行', adaptation: '长标题分两行并下移署名', preserve: '品牌图标、配色和左对齐基线', replace: '报告标题、部门与日期' },
      { file: 'pages/data.page', name: '指标结论', purpose: '展示单一指标及含义', relationship: '结论与证据', capacity: '一项主指标和最多三条解释', adaptation: '四项以上解释拆到下一页', preserve: '指标强调方式和来源区', replace: '指标数值、时间范围、来源与解释' },
    ],
  };
}

async function fixture(t, services = {}) {
  const scratch = await mkdtemp(path.join(os.tmpdir(), 'dsh-template-authoring-'));
  const workspace = path.join(scratch, 'workspace');
  await mkdir(workspace);
  const ctx = new Context(), tools = new Map();
  let rpc;
  ctx.provide('connection', { rpc: { handle: (_channel, value) => { rpc = value; } } });
  ctx.provide('tools', { register: tool => tools.set(tool.name, tool) });
  if (services.attachments) ctx.provide('attachments', services.attachments);
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
    await run('pptd_write_file', { project_path: name, file_path: 'template-design.json', content: JSON.stringify(designSpec(), null, 2) + '\n' });
    return name;
  };
  return { scratch, workspace, ctx, tools, request, run, user, step, agentFor, project, service: ctx.get(plugin.PPT_SERVICE) };
}

test('saved templates retain their editing project and resume authoring on the next chat turn', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  f.user('制作可复用模板');
  await f.run('ppt_template_prepare', { project_path: 'template', name: '持续编辑' });
  const review = await f.request('template/authoring-state');
  const saved = await f.request('template/authoring-save', { ...review, name: '持续编辑' });
  assert.equal(saved.projectPath, 'template', 'Saving must retain the editable workspace project');
  f.user('把刚保存的模板主色改成红色');
  const next = await f.step();
  const context = next.messages.find(message => message.source.kind === 'dsh-ppt-composer');
  assert(context?.content.some(block => block.type === 'text' && block.text.includes('template_authoring:')),
    'The next chat turn must retain authoritative template authoring context');
  assert(next.messages.some(message => message.source.kind === 'skill-invocation' && message.source.name === 'dsh-ppt-template'),
    'The template authoring skill must be available for saved-template revisions');
});

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
  assert(active.some(message => message.source.kind === 'skill-invocation' && message.source.name === 'dsh-ppt-template'), 'Saved authoring retains the template-specific Skill');
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
  assert.equal(saved.origin.base.templateId, saved.templateId);
  assert.notEqual(saved.templateId, review.templateId, 'A new authored copy receives a logical identity independent of its content');
  assert.equal((await f.request('state')).templates.filter(template => template.id === saved.templateId).length, 1);
});

test('reopening captures the source version, preserves local edits, and binds legacy saved workflows before revision', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Original' });
  const review = await f.request('template/authoring-state');
  const saved = await f.request('template/authoring-save', { ...review, name: 'Original' });
  const authoring = f.service.templateAuthoring, library = f.service.store.personalTemplates;
  await f.request('template/select', { templateId: 'dsh-blue-professional', mode: 'ppt' }, 'editor');
  const started = await authoring.start('editor', saved.templateId);
  await library.update('other', saved.templateId, { name: 'Renamed elsewhere' });
  const staleDirectory = path.join(f.workspace, 'stale');
  await mkdir(staleDirectory);
  await assert.rejects(authoring.copySource('editor', { templateId: saved.templateId, projectPath: 'stale' },
    base => library.checkout(base, staleDirectory)), /已更新/);
  assert.equal((await f.request('template/authoring-state', {}, 'editor')).workflowId, started.workflowId);
  assert.deepEqual(await readdir(staleDirectory), []);
  await authoring.exit('editor', started.workflowId);

  const output = path.join(f.workspace, 'edit-original');
  await mkdir(output);
  const editing = await authoring.copySource('editor', { templateId: saved.templateId, projectPath: 'edit-original', turnId: 'edit-turn' },
    base => library.checkout(base, output));
  assert.deepEqual(await readFile(path.join(output, 'template-design.json')), await readFile(path.join(f.workspace, 'template/template-design.json')));
  assert.deepEqual(await readFile(path.join(output, 'pages/data.page')), await readFile(path.join(f.workspace, 'template/pages/data.page')));
  await writeFile(path.join(output, 'local-notes.txt'), 'Keep this unfinished work');
  assert.deepEqual(await authoring.copySource('editor', { templateId: saved.templateId, projectPath: 'ignored-output' },
    () => { throw new Error('An existing workspace must not be copied again'); }), editing);
  assert.equal((await f.request('state', {}, 'editor')).selectedTemplateId, 'dsh-blue-professional');
  await authoring.exit('editor', editing.workflowId);
  assert.equal((await library.readRecord(saved.templateId)).version, editing.origin.base.version);
  assert.equal(await readFile(path.join(output, 'local-notes.txt'), 'utf8'), 'Keep this unfinished work');

  const state = await f.service.store.readState('legacy');
  await f.service.store.writeState({ ...state, templateAuthoring: {
    phase: 'saved', workflowId: saved.workflowId, revision: saved.revision, guideSent: true,
    templateId: saved.templateId, name: saved.name,
  } });
  const legacy = await authoring.revise('legacy', 'legacy-turn');
  assert.equal(legacy.phase, 'creating');
  assert.deepEqual(legacy.origin, editing.origin, 'Raw persisted saved state must capture the actual catalog version before losing its templateId');
  const legacyOutput = path.join(f.workspace, 'legacy-edit');
  await mkdir(legacyOutput);
  const continued = await authoring.copySource('legacy', { templateId: saved.templateId, projectPath: 'legacy-edit' },
    base => library.checkout(base, legacyOutput));
  assert.equal(continued.workflowId, legacy.workflowId);
  assert.deepEqual(await readFile(path.join(legacyOutput, 'pages/data.page')), await readFile(path.join(f.workspace, 'template/pages/data.page')));
});

test('existing authoring saves to the stable original and leaves a conflicting review available as a copy', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Original' });
  const first = await f.request('template/authoring-state');
  const saved = await f.request('template/authoring-save', { ...first, name: 'Original' });
  const library = f.service.store.personalTemplates, authoring = f.service.templateAuthoring;
  const copyDirectory = path.join(f.workspace, 'other-editor');
  await mkdir(copyDirectory);
  await authoring.copySource('editor-b', { templateId: saved.templateId, projectPath: 'other-editor' }, base => library.checkout(base, copyDirectory));
  const file = path.join(f.workspace, 'template/pages/data.page');
  const page = JSON.parse(await readFile(file, 'utf8'));
  page.elements[0].content.text = 'Updated by the first editor';
  await writeFile(file, JSON.stringify(page));
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Updated original' });
  const changed = await f.request('template/authoring-state');
  assert.deepEqual(changed.origin, saved.origin);
  const updated = await f.request('template/authoring-save', { ...changed, name: 'Updated original' });
  assert.equal(updated.templateId, saved.templateId);
  assert.notEqual(updated.origin.base.version, saved.origin.base.version);
  assert.deepEqual(await f.request('template/authoring-save', { ...changed, name: 'Updated original' }), updated,
    'An original retry must retain its submitted base instead of using the new saved version');
  await f.run('ppt_template_prepare', { project_path: 'other-editor', name: 'Other edit' }, 'editor-b');
  const conflict = await f.request('template/authoring-state', {}, 'editor-b');
  await assert.rejects(f.request('template/authoring-save', { ...conflict, name: 'Other edit' }, 'editor-b'), /已更新/);
  assert.equal((await f.request('template/authoring-preview', { ...conflict, page: 1 }, 'editor-b')).page, 1);
  const copied = await f.request('template/authoring-save', { ...conflict, name: 'Other edit', destination: 'copy' }, 'editor-b');
  assert.notEqual(copied.templateId, saved.templateId);
  assert.equal((await library.list()).length, 2);
  assert.deepEqual(await readFile(path.join(await library.projectSource(saved.templateId), 'pages/data.page')), await readFile(file));
});

test('an original save with failed session persistence retries exactly after later rename and deletion', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Original' });
  const first = await f.request('template/authoring-state');
  const saved = await f.request('template/authoring-save', { ...first, name: 'Original' });
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Committed update' });
  const review = await f.request('template/authoring-state');
  const store = f.service.store, library = store.personalTemplates, writeState = store.writeState.bind(store);
  let failed = false;
  store.writeState = async state => {
    if (!failed && state.templateAuthoring?.phase === 'saved') { failed = true; throw new Error('session persistence failed'); }
    return writeState(state);
  };
  await assert.rejects(f.request('template/authoring-save', { ...review, name: 'Committed update' }), /session persistence failed/);
  const committed = await library.savedDraft('session-a', review.draftId);
  await library.update('other', saved.templateId, { name: 'A later rename' });
  await library.remove('other', saved.templateId);
  const recovered = await f.request('template/authoring-save', { ...review, name: 'Committed update' });
  assert.equal(recovered.templateId, saved.templateId);
  assert.equal(recovered.name, committed.name);
  assert.deepEqual(recovered.origin.base, { templateId: committed.id, version: committed.version });
  assert.deepEqual(await f.request('template/authoring-save', { ...review, name: 'Committed update' }), recovered);
  await assert.rejects(f.request('template/authoring-save', { ...review, name: 'Different' }), /其他名称/);
  await assert.rejects(f.request('template/authoring-save', { ...review, name: 'Committed update', destination: 'copy' }), /其他保存目标/);
  assert.deepEqual(await library.list(), [], 'A historical retry must not resurrect a deleted template');
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
    { workflowId, revision: 1, guideSent: true, origin: { kind: 'new' }, phase: 'creating' });
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
  await assert.rejects(f.request('template/authoring-preview', { ...second, page: 1 }), /回执不存在/);
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

test('design rules and layout intent survive reviewed snapshot, save, reopen and reuse exactly', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  const designFile = path.join(f.workspace, 'template/template-design.json');
  const bytes = await readFile(designFile);
  await f.run('ppt_template_prepare', { project_path: 'template', name: '设计系统' });
  const review = await f.request('template/authoring-state');
  const draftRoot = path.join(f.service.store.personalTemplates.root, 'drafts', review.draftId, 'project');
  assert.deepEqual(await readFile(path.join(draftRoot, 'template-design.json')), bytes);
  await writeFile(designFile, JSON.stringify({ ...designSpec(), purpose: '仅工作区改变' }));
  const preview = await f.request('template/authoring-preview', { ...review, page: 1 });
  assert(preview.preview.startsWith('data:image/png;base64,'));
  const saved = await f.request('template/authoring-save', { ...review, name: '设计系统' });
  const library = f.service.store.personalTemplates;
  const reopened = new AuthoredTemplateLibrary(f.service.store, library.convert, library.writeSource);
  assert.deepEqual(await reopened.readDesign(saved.templateId), designSpec());
  const savedRoot = await reopened.projectSource(saved.templateId);
  assert.deepEqual(await readFile(path.join(savedRoot, 'template-design.json')), bytes);
  const record = await reopened.readRecord(saved.templateId);
  assert.equal(record.design, undefined, 'The sidecar is the only saved copy of the design rules');
  await f.request('template/select', { templateId: saved.templateId, mode: 'ppt' }, 'session-b');
  const reference = await f.run('ppt_get_template_reference', { template_id: saved.templateId }, 'session-b');
  assert.match(reference.designProfile, /技术团队向管理层/);
  assert.match(reference.designProfile, /最低 18 pt/);
  assert.match(reference.designProfile, /几何推断/);
  assert.equal(reference.pageContracts[1].sourceTitle, '指标结论');
  assert.equal(reference.pageContracts[1].relationship, '结论与证据');
  for (const text of ['一项主指标', '四项以上解释', '指标强调方式', '指标数值']) assert(reference.pageContracts[1].structureSummary.includes(text));
  const pages = await f.run('ppt_get_template_pages', { template_id: saved.templateId, slide_numbers: [2] }, 'session-b');
  assert.equal(pages[0].sourceTitle, '指标结论');
  assert.equal(pages[0].relationship, '结论与证据');
  assert.match(pages[0].structureSummary, /一项主指标/);
  assert.match(pages[0].simplificationGuidance, /四项以上解释/);
  assert.match(pages[0].pptdLayoutReference, /示例数据 42/);
  const copied = await f.run('ppt_template_create_project', { template_id: saved.templateId, output_directory: 'design-reuse' }, 'session-b');
  assert.equal(copied.designPath, 'design-reuse/template-design.json');
  assert.deepEqual(await readFile(path.join(f.workspace, copied.designPath)), bytes);
  await f.run('ppt_template_prepare', { project_path: 'template', name: '同版式不同规范' });
  const changed = await f.request('template/authoring-state');
  assert.notEqual(changed.templateId, saved.templateId, 'A guidance-only edit must receive a new content identity');
});

test('design sidecar uses the bounded text tools and preserves SHA-256 replacement guards', async t => {
  const f = await fixture(t);
  await f.project();
  const args = { project_path: 'template', file_path: 'template-design.json' };
  const read = await f.run('pptd_read_file', args);
  const changed = JSON.stringify({ ...designSpec(), purpose: '年度经营汇报' });
  await assert.rejects(f.run('pptd_write_file', { ...args, content: changed }), /expected_sha256/);
  await f.run('pptd_write_file', { ...args, content: changed, expected_sha256: read.sha256 });
  await assert.rejects(f.run('pptd_write_file', { ...args, content: changed, expected_sha256: read.sha256 }), /changed after/);
  assert.equal(JSON.parse((await f.run('pptd_read_file', args)).content).purpose, '年度经营汇报');
  assert.equal((await f.run('pptd_list_files', { project_path: 'template' })).files.find(file => file.path === 'template-design.json').kind, 'design');
  for (const file_path of ['other.json', 'pages/template-design.json', 'Template-design.json', '../template-design.json']) {
    await assert.rejects(f.run('pptd_write_file', { ...args, file_path, content: '{}' }), /unsupported|inside/);
    await assert.rejects(f.run('pptd_read_file', { ...args, file_path }), /unsupported|inside/);
  }
  await rm(path.join(f.workspace, 'template/template-design.json'));
  await symlink('deck.pptd', path.join(f.workspace, 'template/template-design.json'));
  await assert.rejects(f.run('pptd_read_file', args), /regular project file/);
  await assert.rejects(f.run('pptd_write_file', { ...args, content: changed, expected_sha256: read.sha256 }), /regular PPTD project file/);
  await assert.rejects(f.run('ppt_template_prepare', { project_path: 'template', name: 'Linked' }), /符号链接/);
});

test('new template preparation requires complete, bounded design guidance tied to every actual page', async t => {
  const f = await fixture(t);
  await f.project();
  const file = path.join(f.workspace, 'template/template-design.json');
  await rm(file);
  await assert.rejects(f.run('ppt_template_prepare', { project_path: 'template', name: 'Missing rules' }), /新建模板需要 template-design.json.*pptd_write_file/);
  const variants = [
    ['invalid JSON', '{', /UTF-8 JSON/],
    ['missing page', { ...designSpec(), layouts: [designSpec().layouts[0]] }, /一一对应/],
    ['unknown page', { ...designSpec(), layouts: [designSpec().layouts[0], { ...designSpec().layouts[1], file: 'pages/unknown.page' }] }, /一一对应/],
    ['duplicate page', { ...designSpec(), layouts: [designSpec().layouts[0], designSpec().layouts[0]] }, /重复/],
    ['foreign path', { ...designSpec(), layouts: [{ ...designSpec().layouts[0], file: '../private.page' }] }, /pages\//],
    ['unknown field', { ...designSpec(), executable: 'do something' }, /必须且只能/],
    ['unknown rule', { ...designSpec(), rules: { ...designSpec().rules, fontSzie: 12 } }, /必须且只能/],
    ['missing rule', { ...designSpec(), rules: {} }, /必须且只能/],
    ['empty rule', { ...designSpec(), rules: { ...designSpec().rules, overflow: ' ' } }, /具体说明/],
    ['long field', { ...designSpec(), purpose: 'x'.repeat(2001) }, /2000/],
    ['long name', { ...designSpec(), layouts: [{ ...designSpec().layouts[0], name: 'x'.repeat(121) }, designSpec().layouts[1]] }, /120/],
    ['unsupported version', { ...designSpec(), version: 2 }, /version/],
  ];
  for (const [label, value, expected] of variants) {
    await writeFile(file, typeof value === 'string' ? value : JSON.stringify(value));
    await assert.rejects(f.run('ppt_template_prepare', { project_path: 'template', name: label }), expected, label);
  }
  await writeFile(file, ' '.repeat(MAX_TEMPLATE_DESIGN_BYTES + 1));
  await assert.rejects(f.run('ppt_template_prepare', { project_path: 'template', name: 'Too large' }), /64 KiB/);
  assert.throws(() => parseTemplateDesign(Buffer.from([0xff, 0xfe])), /UTF-8 JSON/);
  assert.deepEqual(await readdir(path.join(f.service.store.personalTemplates.root, 'drafts')), [], 'Invalid guidance never publishes a draft');
});

test('old saved templates without guidance reuse geometric profiles while bad guidance stays isolated', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Existing saved template' });
  const review = await f.request('template/authoring-state');
  const saved = await f.request('template/authoring-save', { ...review, name: 'Existing saved template' });
  const library = f.service.store.personalTemplates;
  const guide = path.join(await library.projectSource(saved.templateId), 'template-design.json');
  await rm(guide);
  assert.equal(await library.readDesign(saved.templateId), undefined);
  await f.request('template/select', { templateId: saved.templateId, mode: 'ppt' }, 'session-b');
  const reference = await f.run('ppt_get_template_reference', { template_id: saved.templateId }, 'session-b');
  assert.match(reference.designProfile, /未保存作者设计规范/);
  assert.match(reference.designProfile, /几何推断/);
  const copied = await f.run('ppt_template_create_project', { template_id: saved.templateId, output_directory: 'legacy-reuse' }, 'session-b');
  assert.equal(copied.designPath, undefined);
  assert.match(await readFile(path.join(f.workspace, 'legacy-reuse/pages/data.page'), 'utf8'), /示例数据 42/);
  const oldPages = await f.run('ppt_get_template_pages', { template_id: saved.templateId, slide_numbers: [2] }, 'session-b');
  assert.equal(oldPages[0].sourceTitle, '模板第 2 页');
  assert.match(oldPages[0].structureSummary, /几何推断/);
  await writeFile(guide, '{bad');
  assert((await f.request('state')).templates.some(template => template.id === saved.templateId), 'Malformed guidance does not hide catalog records');
  assert.equal((await f.request('state')).templates.filter(template => template.origin !== 'personal').length, 16);
  await assert.rejects(library.readDesign(saved.templateId), /template-design.json 无效/);
  await assert.rejects(f.run('ppt_get_template_reference', { template_id: saved.templateId }, 'session-b'), /template-design.json 无效/);
  await assert.rejects(f.run('ppt_get_template_pages', { template_id: saved.templateId }, 'session-b'), /template-design.json 无效/);
  await assert.rejects(f.run('ppt_template_create_project', { template_id: saved.templateId, output_directory: 'bad-reuse' }, 'session-b'), /template-design.json 无效/);
  await f.request('template/select', { templateId: 'dsh-blue-professional', mode: 'ppt' }, 'session-c');
  assert((await f.run('ppt_get_template_pages', { template_id: 'dsh-blue-professional' }, 'session-c')).length > 0);
});

test('model template preview attaches the actual reviewed Office PNG with matching page labels and revision guards', { timeout: 180000 }, async t => {
  const received = [];
  let afterImage;
  const f = await fixture(t, { attachments: { saveImage: async value => {
    received.push(value);
    const info = await sharp(value.data).metadata();
    if (afterImage) await afterImage();
    return { attachmentId: `preview-${received.length}`, mediaType: value.mediaType, bytes: value.data.length, width: info.width, height: info.height, name: value.name };
  } } });
  await f.project();
  const prepared = await f.run('ppt_template_prepare', { project_path: 'template', name: 'Actual preview' });
  const args = { workflow_id: prepared.workflowId, revision: prepared.revision, slide_numbers: [2, 1] };
  const result = await f.run('ppt_template_preview', args);
  assert.equal(result.workflowId, prepared.workflowId);
  assert.equal(result.revision, prepared.revision);
  assert.deepEqual(result.pages.map(page => page.slideNumber), [2, 1]);
  assert.equal(received.length, 2);
  const review = await f.request('template/authoring-state');
  for (const [index, page] of result.pages.entries()) {
    assert.equal(page.image.width, 1920);
    assert.equal(page.image.height, 1080);
    const preview = await f.request('template/authoring-preview', { ...review, page: page.slideNumber });
    assert.deepEqual(received[index].data, Buffer.from(preview.preview.split(',')[1], 'base64'), 'The model receives the same immutable Office rendering as the user');
    assert.match(received[index].name, new RegExp(`r${prepared.revision}-page-${page.slideNumber}\\.png$`));
  }
  const rendered = f.tools.get('ppt_template_preview').output.render(args, result);
  assert.match(rendered[0].text, /第 2 页/);
  assert.equal(rendered[1].type, 'image');
  assert.equal(rendered[1].attachment.attachmentId, 'preview-1');
  assert.match(rendered[2].text, /第 1 页/);
  assert.equal(rendered[3].attachment.attachmentId, 'preview-2');
  for (const slide_numbers of [[], [1, 1], [0], [3], [1, 2, 1, 2, 1]])
    await assert.rejects(f.run('ppt_template_preview', { ...args, slide_numbers }), /1–4/);
  await assert.rejects(f.run('ppt_template_preview', { ...args, revision: prepared.revision - 1 }), /版本已更新/);
  await assert.rejects(f.run('ppt_template_preview', args, 'session-b'), /先调用 ppt_template_prepare/);
  assert.equal(received.length, 2, 'Invalid, stale and cross-session requests cannot produce image attachments');
  const previewBefore = received[0].data;
  await writeFile(path.join(f.workspace, 'template/pages/data.page'), 'Live project changed after prepare');
  const immutable = await f.run('ppt_template_preview', { ...args, slide_numbers: [2] });
  assert.equal(immutable.pages[0].slideNumber, 2);
  assert.deepEqual(received.at(-1).data, previewBefore);
  afterImage = () => f.run('ppt_template_cancel', { workflow_id: args.workflow_id });
  await assert.rejects(f.run('ppt_template_preview', { ...args, slide_numbers: [1] }), /版本已更新/);
});

test('model preview reports a missing image service without claiming visual inspection', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  const prepared = await f.run('ppt_template_prepare', { project_path: 'template', name: 'User review' });
  await assert.rejects(f.run('ppt_template_preview', { workflow_id: prepared.workflowId, revision: prepared.revision, slide_numbers: [1] }), /没有图片附件服务.*预览卡/);
  assert.equal((await f.request('template/authoring-state')).phase, 'review', 'A missing attachment service leaves the user preview available');
});

test('save revalidates optional draft guidance before admitting a duplicate or new record', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  await f.project();
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Original design' });
  const original = await f.request('template/authoring-state');
  const originalSaved = await f.request('template/authoring-save', { ...original, name: 'Original design' });
  await f.run('ppt_template_prepare', { project_path: 'template', name: 'Duplicate review' });
  const duplicate = await f.request('template/authoring-state');
  assert.equal(duplicate.templateId, original.templateId);
  const guide = path.join(f.service.store.personalTemplates.root, 'drafts', duplicate.draftId, 'project/template-design.json');
  const bytes = await readFile(guide);
  await writeFile(guide, JSON.stringify({ ...designSpec(), layouts: [designSpec().layouts[0]] }));
  await assert.rejects(f.request('template/authoring-save', { ...duplicate, name: 'Invalid review' }), /一一对应/);
  assert.equal((await f.request('template/authoring-state')).phase, 'review');
  assert.equal((await f.service.store.personalTemplates.readRecord(originalSaved.templateId)).name, 'Original design');
  await writeFile(guide, bytes);
  const saved = await f.request('template/authoring-save', { ...duplicate, name: 'Duplicate review' });
  assert.equal(saved.templateId, originalSaved.templateId);
});
