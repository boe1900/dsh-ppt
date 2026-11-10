import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Context } from '@deepseek-ai/cordis';
import { SkillRegistry } from '@deepseek-ai/dsh-skill';
import { SystemPrompt } from '@deepseek-ai/dsh-system-prompt';
import * as plugin from '../lib/index.js';
import * as preset from '../lib/preset.js';
import { loadPptdProject } from '../lib/pptd.js';

async function fixture(t, scratch) {
  scratch ??= await mkdtemp(path.join(os.tmpdir(), 'dsh-export-recovery-'));
  const workspace = path.join(scratch, 'workspace');
  await mkdir(workspace, { recursive: true });
  const ctx = new Context(), tools = new Map();
  ctx.provide('connection', { rpc: { handle() {} } });
  ctx.provide('tools', { register: tool => tools.set(tool.name, tool) });
  ctx.provide('workingDirectory', { ensure: async () => workspace });
  ctx.provide('sandboxPolicy', { resolve: () => ({ mode: 'workspace-write', workspaceRoot: workspace }) });
  await ctx.plugin(SystemPrompt, { includeHarnessIdentity: false });
  await ctx.plugin(SkillRegistry);
  await ctx.plugin(plugin, { root: path.join(scratch, 'storage') });
  await ctx.plugin(preset);
  t.after(async () => { await ctx.fiber.dispose(); await rm(scratch, { recursive: true, force: true }); });
  await mkdir(path.join(workspace, 'source'), { recursive: true });
  await writeFile(path.join(workspace, 'source/deck.pptd'), 'version: v2\ntitle: Reliable report\nsize: [960, 540]\npages: [one.page]\n');
  await writeFile(path.join(workspace, 'source/one.page'), JSON.stringify({ elements: [
    { elementId: 'title', elementType: 'text', bounds: [40, 40, 800, 100], content: { text: 'Report', fontSize: 36 } },
  ] }));
  const service = ctx.get(plugin.PPT_SERVICE);
  const run = (callId = 'call-1') => tools.get('pptd_render').execute({ project_path: 'source', output_file: 'report.pptx' }, {
    callId, agent: { id: 'session', session: { header: { cwd: workspace, agentPreset: 'ppt' } } }, signal: new AbortController().signal,
  });
  const deliveries = async () => (await readdir(workspace)).filter(name => name !== 'source' && !name.endsWith('.stage'));
  return { scratch, workspace, ctx, service, run, deliveries };
}

test('committed exports survive audit failure, keep a pending audit and deduplicate the same tool call', async t => {
  const f = await fixture(t), store = f.service.store;
  const appendAudit = store.appendAudit.bind(store);
  store.appendAudit = async () => { throw Object.assign(new Error('audit disk full'), { code: 'ENOSPC' }); };
  const first = await f.run();
  assert.equal(first.status, 'exported');
  assert.equal((await f.service.state('session')).decks.length, 1);
  assert.equal((await readdir(path.join(store.root, 'audit-pending'))).length, 1);
  const retry = await f.run();
  assert.equal(retry.id, first.id);
  assert.deepEqual(await f.deliveries(), ['Reliable-report']);
  store.appendAudit = appendAudit;
  await store.flushAudits();
  assert.deepEqual(await readdir(path.join(store.root, 'audit-pending')), []);
  assert.match(await readFile(path.join(store.root, 'audit.ndjson'), 'utf8'), new RegExp(first.id));
  const intentional = await f.run('new-call');
  assert.notEqual(intentional.id, first.id, 'A new successful request can intentionally make another delivery');
});

test('a published export with failed state persistence resumes after restart and merges current state', async t => {
  const f = await fixture(t), store = f.service.store;
  const writeState = store.writeState.bind(store);
  store.writeState = async state => {
    if (state.decks.length) throw new Error('state volume unavailable');
    return writeState(state);
  };
  await assert.rejects(f.run(), /state volume unavailable/);
  assert.equal((await f.service.state('session')).decks.length, 0);
  assert.deepEqual(await f.deliveries(), ['Reliable-report']);
  const fileBefore = await readFile(path.join(f.workspace, 'Reliable-report/report.pptx'));
  await f.ctx.fiber.dispose();
  const restarted = await fixture(t, f.scratch);
  await restarted.service.selectTemplate('session', 'dsh-broadside', { kind: 'user' });
  // A newer selection is independent state and must survive recovery of the original export.
  await restarted.service.templateAuthoring.start('session');
  const receipts = await restarted.service.exports.receipts('session');
  const receipt = receipts[0];
  const deck = await restarted.run();
  assert.equal(deck.id, receipt.deck.id);
  assert.equal((await restarted.service.state('session')).selectedTemplateId, 'dsh-broadside');
  assert.equal((await restarted.service.state('session')).templateAuthoring.phase, 'creating');
  assert.equal((await restarted.service.state('session')).decks.length, 1);
  assert.deepEqual(await restarted.deliveries(), ['Reliable-report']);
  assert.deepEqual(await readFile(path.join(restarted.workspace, 'Reliable-report/report.pptx')), fileBefore);
});

test('new-call retry resumes pending publication and changed input never overwrites the old delivery', async t => {
  const f = await fixture(t), store = f.service.store;
  const writeState = store.writeState.bind(store);
  let fail = true;
  store.writeState = async state => {
    if (fail && state.decks.length) { fail = false; throw new Error('transient state write'); }
    return writeState(state);
  };
  await assert.rejects(f.run(), /transient state write/);
  const retry = await f.run('retry-with-new-call');
  assert.equal((await f.service.state('session')).decks.length, 1);
  assert.deepEqual(await f.deliveries(), ['Reliable-report']);
  assert.equal((await f.run('retry-with-new-call')).id, retry.id, 'Adopted call aliases also remain idempotent');
  await writeFile(path.join(f.workspace, 'source/one.page'), JSON.stringify({ elements: [] }));
  await assert.rejects(f.run(), /input has changed/);
  assert.equal((await f.service.state('session')).decks[0].id, retry.id);
});

test('a failed receipt never publishes and a modified uncommitted delivery is preserved', async t => {
  const f = await fixture(t), store = f.service.store;
  const writeJson = store.writeJson.bind(store);
  store.writeJson = async (target, value) => {
    if (target.includes('/exports/')) throw new Error('cannot reserve export');
    return writeJson(target, value);
  };
  await assert.rejects(f.run(), /cannot reserve export/);
  assert.deepEqual(await f.deliveries(), []);
  store.writeJson = writeJson;
  const writeState = store.writeState.bind(store);
  store.writeState = async state => { if (state.decks.length) throw new Error('state failed'); return writeState(state); };
  await assert.rejects(f.run(), /state failed/);
  const target = path.join(f.workspace, 'Reliable-report/report.pptx');
  await writeFile(target, 'user edited file');
  store.writeState = writeState;
  await assert.rejects(f.run(), /has changed/);
  assert.equal(await readFile(target, 'utf8'), 'user edited file');
  assert.equal((await f.service.state('session')).decks.length, 0);
});

test('concurrent sessions sharing a workspace receive separate deliveries and recover their own receipts', async t => {
  const f = await fixture(t);
  const source = await loadPptdProject(path.join(f.workspace, 'source'));
  const exports = f.service.exports, reserve = exports.reserve.bind(exports);
  let arrivals = 0, release;
  const bothReserved = new Promise(resolve => { release = resolve; });
  exports.reserve = async (...args) => {
    const receipt = await reserve(...args);
    if (++arrivals === 2) release();
    await bothReserved;
    return receipt;
  };
  const run = id => f.service.createPptdDeck(id, source, 'report.pptx', f.workspace, { kind: 'agent' }, new AbortController().signal, 'same-call-id');
  const [a, b] = await Promise.all([run('session-a'), run('session-b')]);
  assert.notEqual(a.output.workspaceDirectoryPath, b.output.workspaceDirectoryPath);
  assert.equal((await f.service.state('session-a')).decks.length, 1);
  assert.equal((await f.service.state('session-b')).decks.length, 1);
  assert.equal((await run('session-b')).id, b.id);
  assert.equal((await f.deliveries()).length, 2);
});

test('startup recovers an audit prepared before the process stopped just after state commit', async t => {
  const f = await fixture(t), store = f.service.store;
  const recordAudit = store.recordAudit.bind(store);
  store.recordAudit = async () => { throw new Error('simulated process stop after commit'); };
  await assert.rejects(f.run(), /simulated process stop/);
  assert.equal((await f.service.state('session')).decks.length, 1);
  store.recordAudit = recordAudit;
  await f.ctx.fiber.dispose();
  const restarted = await fixture(t, f.scratch);
  assert.match(await readFile(path.join(store.root, 'audit.ndjson'), 'utf8'), /"operation":"create"/);
  assert.equal((await restarted.run()).status, 'exported');
  assert.deepEqual(await restarted.deliveries(), ['Reliable-report']);
});

test('a verified committed audit remains retryable after bounded history is trimmed', async t => {
  const f = await fixture(t), store = f.service.store;
  store.limits.maxActivities = 1;
  const activity = { id: '12345678-1234-1234-1234-123456789abc', operation: 'select-template', actor: 'user',
    status: 'completed', startedAt: new Date().toISOString(), completedAt: new Date().toISOString(), summary: 'Selected' };
  const state = await f.service.state('session');
  await store.prepareAudit('session', activity, {});
  await store.writeState({ ...state, activities: [activity] });
  const appendAudit = store.appendAudit.bind(store);
  store.appendAudit = async () => { throw new Error('audit still unavailable after restart'); };
  await store.flushAudits();
  const ready = JSON.parse(await readFile(path.join(store.root, 'audit-pending', `${activity.id}.json`), 'utf8'));
  assert.equal(ready.requiresCommit, undefined);
  await store.writeState({ ...state, activities: [{ ...activity, id: '22345678-1234-1234-1234-123456789abc' }, activity] });
  store.appendAudit = appendAudit;
  await store.flushAudits();
  assert.deepEqual(await readdir(path.join(store.root, 'audit-pending')), []);
  assert.match(await readFile(path.join(store.root, 'audit.ndjson'), 'utf8'), new RegExp(activity.id));
});
