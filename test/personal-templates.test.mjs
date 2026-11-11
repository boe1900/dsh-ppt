import assert from 'node:assert/strict';
import { randomUUID } from 'node:crypto';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { Readable } from 'node:stream';
import test from 'node:test';
import { Context } from '@deepseek-ai/cordis';
import { SkillRegistry } from '@deepseek-ai/dsh-skill';
import { SystemPrompt } from '@deepseek-ai/dsh-system-prompt';
import PptxGenJS from 'pptxgenjs';
import sharp from 'sharp';
import { unzipSync, strFromU8 } from 'fflate';
import * as plugin from '../lib/index.js';
import * as preset from '../lib/preset.js';

async function fixture(t, storageRoot) {
  const scratch = storageRoot ?? await mkdtemp(path.join(os.tmpdir(), 'dsh-ppt-personal-'));
  if (!storageRoot) t.after(() => rm(scratch, { recursive: true, force: true }));
  const workspace = path.join(scratch, 'workspace');
  await mkdir(workspace, { recursive: true });
  const ctx = new Context();
  const tools = new Map(), routes = new Map();
  let rpc;
  ctx.provide('connection', {
    rpc: { handle: (_channel, handler) => { rpc = handler; } },
    requestRejection: req => req.headers.authorization === 'Bearer test' ? undefined : 401,
  });
  ctx.provide('tools', { register: tool => tools.set(tool.name, tool) });
  ctx.provide('workingDirectory', { ensure: async () => workspace });
  ctx.provide('sandboxPolicy', { resolve: () => ({ mode: 'workspace-write', workspaceRoot: workspace }) });
  ctx.provide('webServer', { register: route => { routes.set(route.path, route); return () => routes.delete(route.path); } });
  const prompt = ctx.plugin(SystemPrompt, { includeHarnessIdentity: false });
  const skills = ctx.plugin(SkillRegistry);
  await prompt;
  await skills;
  const runtime = ctx.plugin(plugin, { root: path.join(scratch, 'storage') });
  await runtime;
  const child = ctx.plugin(preset);
  await child;
  t.after(async () => { await child.dispose(); await runtime.dispose(); await skills.dispose(); await prompt.dispose(); });
  const request = async (endpoint, payload = {}, sessionId = 'session-a') => {
    const response = await rpc(endpoint, { ...payload, sessionId });
    assert.equal(response.ok, true, JSON.stringify(response));
    if (response.value.status === 'error') throw new Error(response.value.error.message);
    return response.value.data;
  };
  const run = (name, args) => tools.get(name).execute(args, {
    agent: { id: 'session-b', session: { header: { cwd: workspace } } }, signal: new AbortController().signal,
  });
  return { scratch, workspace, routes, request, run };
}

test('custom PPTX templates upload, preview, persist, copy and export through the preset', { timeout: 180000 }, async t => {
  const f = await fixture(t);
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  for (let page = 1; page <= 2; page++) {
    const slide = pptx.addSlide();
    slide.background = { color: 'F4F2ED' };
    slide.addText(`Company ${page}`, { x: 0.6, y: 0.4, w: 11, h: 0.8, fontSize: 28, fontFace: 'Arial', bold: true });
    slide.addText('Quarterly performance and next steps', { x: 0.6, y: 1.8, w: 10, h: 2, fontSize: 18 });
  }
  const bytes = Buffer.from(await pptx.write({ outputType: 'nodebuffer' }));
  const upload = await f.request('template/upload-start', { input: { fileName: 'Company.pptx', size: bytes.length } });
  await assert.rejects(f.request('template/upload-chunk', { uploadId: upload.uploadId, offset: 0, base64: bytes.toString('base64') }, 'other-session'), /其他会话/);
  await f.request('template/upload-chunk', { uploadId: upload.uploadId, offset: 0, base64: bytes.toString('base64') });
  const requestId = randomUUID();
  const draft = await f.request('template/prepare', { input: { uploadId: upload.uploadId, requestId } });
  const preview = await sharp(Buffer.from(draft.preview.split(',')[1], 'base64')).metadata();
  assert.deepEqual([preview.width, preview.height], [1920, 1080]);
  assert.equal((await f.request('template/import-progress', { requestId })).completed, 2);
  assert.equal((await f.request('template/preview-page', { draftId: draft.draftId, page: 2 })).page, 2);
  await assert.rejects(f.request('template/save', { draftId: draft.draftId, name: 'Wrong owner' }, 'session-b'), /其他会话/);
  const saved = await f.request('template/save', { draftId: draft.draftId, name: '公司模板' });
  assert.equal(saved.origin, 'personal');
  const packageFile = await f.request('template/package-export', { templateId: saved.id });
  const receiving = await fixture(t);
  const imported = await receiving.request('template/package-import', packageFile);
  assert.equal(imported.outcome, 'imported');
  assert.deepEqual(imported.template.templateProfile, saved.templateProfile);
  assert.deepEqual(imported.template.templateSamples, saved.templateSamples);
  assert.equal(imported.template.authoringKind, undefined);
  assert.equal((await receiving.request('template/package-import', packageFile)).outcome, 'existing');
  await receiving.request('template/select', { templateId: imported.template.id, mode: 'ppt' }, 'session-b');
  await receiving.run('ppt_template_create_project', { template_id: imported.template.id, output_directory: 'shared-report' });
  assert.equal((await receiving.run('pptd_render', { project_path: 'shared-report', output_file: 'shared-report.pptx' })).status, 'exported');
  await assert.rejects(f.request('template/package-export', { templateId: 'dsh-blue-professional' }), /标识无效/);
  await f.request('template/select', { templateId: saved.id, mode: 'ppt' }, 'session-b');
  const restored = await fixture(t, f.scratch);
  const state = await restored.request('state', {}, 'session-b');
  assert.equal(state.selectedTemplateId, saved.id);
  assert(state.templates.some(template => template.id === saved.id));
  const metadataFile = path.join(f.scratch, 'storage/personal-templates/saved', saved.id, 'template.json');
  const legacy = JSON.parse(await readFile(metadataFile, 'utf8'));
  delete legacy.templateProfile;
  await writeFile(metadataFile, JSON.stringify(legacy));
  const duplicate = await restored.request('template/prepare', { input: { fileName: 'Same.pptx', base64: bytes.toString('base64') } });
  assert.equal(duplicate.duplicate, true);
  assert.equal(duplicate.template.id, saved.id);
  assert.notEqual(duplicate.template.version, legacy.version, 'The owner publishes profile enrichment returned by the import worker');
  assert.deepEqual(JSON.parse(await readFile(metadataFile, 'utf8')).templateProfile, duplicate.template.templateProfile);
  await restored.request('template/update', { templateId: saved.id, name: '年度报告', description: '公司报告模板' });
  assert.equal((await restored.request('state')).templates.find(template => template.id === saved.id).description, '公司报告模板');
  const index = await restored.run('ppt_get_template_pages', { template_id: saved.id });
  assert.equal(index.length, 2);
  await restored.run('ppt_template_create_project', { template_id: saved.id, output_directory: 'report' });
  await assert.rejects(restored.run('ppt_template_create_project', { template_id: saved.id, output_directory: '../escape' }), /inside|workspace|路径/);
  const output = await restored.run('pptd_render', { project_path: 'report', output_file: 'report.pptx' });
  assert.equal(output.status, 'exported', JSON.stringify(output.check));
  const exportedPath = path.join(f.workspace, output.outputPath);
  const zip = unzipSync(await readFile(exportedPath));
  assert.match(strFromU8(zip['ppt/slides/slide1.xml']), /<a:t>/);
  assert.deepEqual(await readFile(path.join(f.scratch, 'storage/personal-templates/saved', saved.id, 'source.pptx')), bytes);
  await restored.request('template/delete', { templateId: saved.id });
  assert(!(await restored.request('state')).templates.some(template => template.id === saved.id));
  assert.equal((await readFile(exportedPath)).subarray(0, 2).toString(), 'PK');
});

test('template upload rejects bad input and keeps HTTP authentication', async t => {
  const f = await fixture(t);
  await assert.rejects(f.request('template/upload-start', { input: { fileName: '../escape.pptx', size: 4 } }), /PPTX/);
  await assert.rejects(f.request('template/prepare', { input: { fileName: 'bad.pptx', base64: 'not base64' } }), /Base64/);
  const upload = await f.request('template/upload-start', { input: { fileName: 'cancel.pptx', size: 4 } });
  await f.request('template/upload-cancel', { uploadId: upload.uploadId });
  await f.request('template/upload-cancel', { uploadId: upload.uploadId });
  for (const endpoint of ['template/upload-start', 'template/package-import', 'template/package-export']) {
    const req = Readable.from([Buffer.from('{}')]);
    Object.assign(req, { method: 'POST', url: `/dsh-ppt/${endpoint}`, headers: {} });
    let status;
    await f.routes.get('/dsh-ppt').handler(req, { writeHead: code => { status = code; }, end() {} });
    assert.equal(status, 401);
  }
});
