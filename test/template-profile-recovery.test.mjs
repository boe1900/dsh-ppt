import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { PersonalTemplateLibrary } from '../lib/personal-templates.js';
import { extractPersonalTemplateProfile, validPersonalTemplateProfile } from '../lib/personal-template-profile.js';
import { parsePptdProject } from '../lib/pptd.js';

const fonts = { latin: 'Arial', ea: 'Microsoft YaHei', mac: 'Microsoft YaHei', win: 'Microsoft YaHei' };
const session = 'old-template-session';
const draftId = '12345678-1234-1234-1234-123456789abc';

async function fixture(t, kind = 'saved', changes = {}) {
  const root = await mkdtemp(path.join(os.tmpdir(), 'dsh-profile-recovery-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  const library = new PersonalTemplateLibrary({
    root, recordAudit: async () => {},
    writeJson: async (file, value) => { await mkdir(path.dirname(file), { recursive: true }); await writeFile(file, JSON.stringify(value)); },
  });
  const source = {
    entryName: 'deck.pptd',
    manifest: JSON.stringify({ version: 'v2', title: 'Legacy native template', size: [960, 540], pages: ['pages/cover.page'] }),
    pages: new Map([['pages/cover.page', JSON.stringify({ elements: [{
      elementId: 'title', elementType: 'text', bounds: [60, 50, 820, 90],
      content: { text: '品牌模板', fontFamily: fonts, fontSize: 36, color: '#123456' },
    }] })]]), assets: new Map(),
  };
  const profile = extractPersonalTemplateProfile(parsePptdProject(source));
  profile.typography.title.fontFace = fonts;
  profile.pages[0].slots[0].fontFace = fonts;
  const id = `personal-${'a'.repeat(64)}`;
  const record = {
    id, conversionVersion: 10, previewVersion: 11, origin: 'personal', authoringKind: 'native-pptd', name: '旧品牌模板', description: 'Keep the description',
    createdAt: '2026-11-10T00:00:00.000Z', slideCount: 1,
    pageIndex: [{ slideNumber: 1, file: 'pages/cover.page' }],
    previewImages: Array(3).fill('data:image/png;base64,eA=='), templateProfile: profile,
    palette: profile.palette, titleFontFace: fonts, bodyFontFace: profile.typography.body.fontFace,
    ...changes,
  };
  const directory = path.join(library.root, kind, kind === 'saved' ? id : draftId);
  await mkdir(path.join(directory, 'project/pages'), { recursive: true });
  const metadata = path.join(directory, 'template.json');
  await writeFile(metadata, JSON.stringify(record));
  await writeFile(path.join(directory, 'owner.json'), JSON.stringify({ session: createHash('sha256').update(session).digest('hex') }));
  await writeFile(path.join(directory, 'project/deck.pptd'), source.manifest);
  await writeFile(path.join(directory, 'project/pages/cover.page'), source.pages.get('pages/cover.page'));
  return { library, root, id, directory, metadata, source, record };
}

test('legacy native saved indexes recover without changing metadata, source or ownership', async t => {
  const f = await fixture(t);
  const before = await readFile(f.metadata);
  const owner = await readFile(path.join(f.directory, 'owner.json'));
  const record = await f.library.readRecord(f.id);
  assert(validPersonalTemplateProfile(record.templateProfile, record.pageIndex));
  assert.equal(record.templateProfile.typography.title.fontFace, 'Microsoft YaHei');
  assert.equal(record.titleFontFace, 'Microsoft YaHei');
  assert.equal(record.name, f.record.name);
  assert.equal(record.description, f.record.description);
  assert.deepEqual((await f.library.list()).map(item => item.id), [f.id]);
  const target = path.join(f.root, 'reused');
  await mkdir(target);
  await f.library.copyProject(f.id, target);
  assert.equal(await readFile(path.join(target, 'pages/cover.page'), 'utf8'), f.source.pages.get('pages/cover.page'));
  assert.deepEqual(await readFile(f.metadata), before, 'Reading and reusing never rewrites old metadata');
  assert.deepEqual(await readFile(path.join(f.directory, 'owner.json')), owner);
});

test('legacy native draft indexes are repaired and validated before save publication', async t => {
  const f = await fixture(t, 'drafts');
  await assert.rejects(f.library.save('foreign-session', draftId, 'Wrong owner'), /其他会话/);
  const saved = await f.library.save(session, draftId, '恢复后的品牌模板');
  assert(validPersonalTemplateProfile(saved.templateProfile, saved.pageIndex), 'Save must not publish another unreadable index');
  const metadata = JSON.parse(await readFile(path.join(f.library.root, 'saved', f.id, 'template.json'), 'utf8'));
  assert(validPersonalTemplateProfile(metadata.templateProfile, metadata.pageIndex));
  assert.equal(metadata.name, '恢复后的品牌模板');
  assert.equal((await f.library.readRecord(f.id)).id, f.id);
});

for (const [label, change] of [
  ['non-native record', f => { f.record.authoringKind = undefined; }],
  ['invalid independent metadata', f => { f.record.previewImages = []; }],
  ['invalid sample index', f => { f.record.templateSamples = { broken: true }; }],
  ['mismatched page index', f => { f.record.pageIndex[0].file = 'pages/other.page'; }],
]) test(`profile recovery does not bypass a ${label}`, async t => {
  const f = await fixture(t);
  change(f);
  await writeFile(f.metadata, JSON.stringify(f.record));
  const before = await readFile(f.metadata);
  await assert.rejects(f.library.readRecord(f.id));
  assert.deepEqual(await f.library.list(), []);
  assert.deepEqual(await readFile(f.metadata), before);
});

for (const issue of ['missing', 'invalid', 'symlink']) test(`a ${issue} native source stays isolated and is never published by recovery`, async t => {
  const f = await fixture(t, 'drafts');
  const file = path.join(f.directory, 'project/pages/cover.page');
  if (issue === 'invalid') await writeFile(file, '{broken');
  else {
    await rm(file);
    if (issue === 'symlink') {
      const other = path.join(f.directory, 'project/pages/other.page');
      await writeFile(other, f.source.pages.get('pages/cover.page'));
      await symlink(other, file);
    }
  }
  await assert.rejects(f.library.save(session, draftId, 'Should remain a draft'));
  await assert.rejects(readFile(path.join(f.library.root, 'saved', f.id, 'template.json')), { code: 'ENOENT' });
  assert.equal(await readFile(f.metadata, 'utf8'), JSON.stringify(f.record));
});
