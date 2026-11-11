import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile, cp } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { PersonalTemplateLibrary } from '../lib/personal-templates.js';

async function fixture(t) {
  const root = await mkdtemp(path.join(os.tmpdir(), 'dsh-template-catalog-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  const audits = [];
  const library = new PersonalTemplateLibrary({
    root, recordAudit: async (sessionId, activity, facts) => audits.push({ sessionId, ...activity, ...facts }),
    writeJson: async (target, value) => { await mkdir(path.dirname(target), { recursive: true }); await writeFile(target, JSON.stringify(value)); },
  });
  const save = async (digit, changes = {}) => {
    const id = `personal-${digit.repeat(64)}`;
    const record = {
      id, origin: 'personal', name: `Template ${digit}`, createdAt: '2026-11-10T00:00:00.000Z',
      slideCount: 1, pageIndex: [{ slideNumber: 1, file: 'pages/cover.page' }],
      previewImages: Array(3).fill('data:image/png;base64,eA=='), ...changes,
    };
    const directory = path.join(library.root, 'saved', id);
    await mkdir(directory, { recursive: true });
    const file = path.join(directory, 'template.json');
    await writeFile(file, JSON.stringify(record));
    return { id, record, file, directory };
  };
  return { root, library, audits, save };
}

test('catalog isolates broken records while preserving valid templates and strict reads', async t => {
  const { library, audits, save } = await fixture(t);
  const good = await save('a');
  const malformed = await save('b');
  const badDate = await save('c', { createdAt: null });
  const missing = await save('d');
  await writeFile(malformed.file, '{broken');
  await rm(missing.file);
  assert.deepEqual((await library.list()).map(record => record.id), [good.id]);
  assert.deepEqual(audits.map(activity => activity.templateId).sort(), [malformed.id, badDate.id, missing.id].sort());
  assert(audits.every(activity => activity.operation === 'skip-unavailable-template' && activity.status === 'failed'));
  assert.equal(await readFile(malformed.file, 'utf8'), '{broken', 'Catalog isolation preserves user files');
  await assert.rejects(library.readRecord(malformed.id), SyntaxError);
  await library.list();
  assert.equal(audits.length, 3, 'Repeated state reads do not duplicate unchanged diagnostics');
  await save('c');
  assert.deepEqual(new Set((await library.list()).map(record => record.id)), new Set([good.id, badDate.id]));
});

test('catalog reads and short library publication observe a consistent snapshot', async t => {
  const { library, save } = await fixture(t);
  const saved = await save('a');
  const unrelated = await save('b');
  const catalogLocked = library.catalogLocked.bind(library);
  let enter, release;
  const entered = new Promise(resolve => { enter = resolve; });
  const blocked = new Promise(resolve => { release = resolve; });
  let first = true;
  library.catalogLocked = work => catalogLocked(async () => {
    if (first) { first = false; enter(); await blocked; }
    return work();
  });
  const listing = library.list();
  await entered;
  const deletion = library.remove('session-a', saved.id);
  release();
  assert.deepEqual((await listing).map(record => record.id), [saved.id, unrelated.id]);
  assert.deepEqual(await deletion, { id: saved.id });
  assert.deepEqual((await library.list()).map(record => record.id), [unrelated.id]);
});

test('saving within the library queue does not reacquire its catalog lock', { timeout: 5000 }, async t => {
  const { library, save } = await fixture(t);
  const saved = await save('a');
  const draftId = '12345678-1234-1234-1234-123456789abc';
  const directory = path.join(library.root, 'drafts', draftId);
  await mkdir(directory, { recursive: true });
  await writeFile(path.join(directory, 'template.json'), JSON.stringify(saved.record));
  await writeFile(path.join(directory, 'owner.json'), JSON.stringify({ session: createHash('sha256').update('session-a').digest('hex') }));
  await mkdir(path.join(directory, 'project/pages'), { recursive: true });
  await writeFile(path.join(directory, 'project/deck.pptd'), 'version: v2\nsize: [960, 540]\npages: [pages/cover.page]\n');
  await writeFile(path.join(directory, 'project/pages/cover.page'), JSON.stringify({ elements: [] }));
  await cp(path.join(directory, 'project'), path.join(saved.directory, 'project'), { recursive: true });
  assert.deepEqual(await library.save('session-a', draftId, 'Already saved'), await library.readRecord(saved.id));
});

test('catalog keeps saved-root failures visible', async t => {
  const { root, library } = await fixture(t);
  await mkdir(library.root, { recursive: true });
  await symlink(root, path.join(library.root, 'saved'));
  await assert.rejects(library.list(), /目录无效/);
  await rm(path.join(library.root, 'saved'));
  await writeFile(path.join(library.root, 'saved'), 'invalid root');
  await assert.rejects(library.list(), { code: 'ENOTDIR' });
});

test('library audit preserves the original work failure and successful result', async t => {
  const { library, audits } = await fixture(t);
  const value = { id: 'saved-template' };
  assert.equal(await library.audit('session-a', 'save-personal-template', async () => value), value);
  const failure = new Error('original conversion failure');
  await assert.rejects(library.audit('session-a', 'prepare-personal-template', async () => { throw failure; }), error => error === failure);
  assert.deepEqual(audits.map(activity => activity.status), ['completed', 'failed']);
});
