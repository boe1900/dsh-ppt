import assert from 'node:assert/strict';
import { createHash, randomBytes, randomUUID } from 'node:crypto';
import { mkdtemp, mkdir, readFile, readdir, rm, writeFile, cp, symlink, rename } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import sharp from 'sharp';
import { zipSync, unzipSync } from 'fflate';
import { PersonalTemplateLibrary } from '../lib/personal-templates.js';
import { loadPptdProject, parsePptdProject, renderPptdProject } from '../lib/pptd.js';
import { extractPersonalTemplateProfile } from '../lib/personal-template-profile.js';
import { decodeTemplatePackage, encodeTemplatePackage, validateTemplatePackage } from '../lib/template-package.js';

const hash = bytes => createHash('sha256').update(bytes).digest('hex');
async function fixture(t) {
  const root = await mkdtemp(path.join(os.tmpdir(), 'dsh-package-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  const library = new PersonalTemplateLibrary({ root, recordAudit: async () => {},
    writeJson: async (file, value) => { const temporary = `${file}.tmp`; await mkdir(path.dirname(file), { recursive: true }); await writeFile(temporary, JSON.stringify(value)); await rename(temporary, file); } });
  return { root, library };
}
async function packageFixture({ legacy = false, absent = false, title = 'Shared template' } = {}) {
  const manifest = `version: v2\ntitle: ${title}\nsize: [960, 540]\npages: [pages/cover.page, pages/body.page]\n`;
  const page = text => `pageType: content\nbackground: { type: solid, color: '#FFFFFF' }\nelements:\n  - elementId: title\n    elementType: text\n    bounds: [40, 40, 700, 100]\n    content: { text: '${text}', fontSize: 28, color: '#222222' }\n  - elementId: logo\n    elementType: image\n    bounds: [800, 40, 80, 80]\n    src: assets/logo.png\n`;
  const logo = await sharp({ create: { width: 24, height: 24, channels: 3, background: '#2463aa' } }).png().toBuffer();
  const preview = await sharp({ create: { width: 960, height: 540, channels: 3, background: '#f4f4f4' } }).png().toBuffer();
  const files = new Map([
    ['project/deck.pptd', Buffer.from(manifest)],
    ['project/pages/cover.page', Buffer.from(page('Source title'))],
    ['project/pages/body.page', Buffer.from(page('Source body'))],
    ['project/assets/logo.png', logo],
    ['preview/pages/page-1.png', preview], ['preview/pages/page-2.png', preview],
  ]);
  const design = { version: 1, purpose: 'Reusable source and media', rules: Object.fromEntries(['palette', 'typography', 'spacing', 'imagery', 'data', 'overflow'].map(key => [key, 'Preserve hierarchy and readable source content.'])),
    layouts: ['cover', 'body'].map(name => ({ file: `pages/${name}.page`, name, purpose: 'Explain source', relationship: 'Independent', capacity: 'Short text', adaptation: 'Duplicate the page', preserve: 'Logo', replace: 'Text' })) };
  files.set('project/template-design.json', Buffer.from(JSON.stringify(design, null, 2)));
  const project = parsePptdProject({ entryName: 'deck.pptd', manifest,
    pages: new Map(['cover', 'body'].map(name => [`pages/${name}.page`, files.get(`project/pages/${name}.page`).toString()])),
    assets: new Map([['assets/logo.png', { path: 'assets/logo.png', bytes: logo, mediaType: 'image/png', sha256: hash(logo) }]]) });
  const profile = extractPersonalTemplateProfile(project);
  if (legacy) { profile.typography.title.fontFace = 'Legacy Original Face'; profile.pages[0].slots[0].textCapacity = 81; files.set('original/source.pptx', Buffer.from('opaque original PPTX bytes')); }
  const metadata = { name: title, description: 'Portable design', sourceKind: legacy ? 'converted-pptx' : 'native-pptd',
    ...absent ? {} : { templateProfile: profile },
    ...legacy && !absent ? { templateSamples: { version: 1, simplifiedTextCount: 12, layouts: [
      { slideNumber: 1, role: 'cover', content: 'sample', representative: true }, { slideNumber: 2, role: 'content', content: 'source', representative: false }
    ], representativePages: [1] } } : {} };
  const snapshot = await validateTemplatePackage({ metadata, files });
  return { snapshot, input: encodeTemplatePackage(snapshot) };
}
function repack(input, mutate) {
  const entries = unzipSync(Buffer.from(input.base64, 'base64'));
  const manifest = JSON.parse(Buffer.from(entries['manifest.json']).toString());
  mutate(entries, manifest);
  manifest.files = Object.entries(entries).filter(([file]) => file !== 'manifest.json').map(([file, bytes]) => ({ path: file, size: bytes.length, sha256: hash(bytes) }));
  entries['manifest.json'] = Buffer.from(JSON.stringify(manifest));
  return { fileName: input.fileName, base64: Buffer.from(zipSync(entries, { level: 0 })).toString('base64') };
}

for (const legacy of [false, true]) test(`${legacy ? 'converted PPTX' : 'native'} package preserves reusable bytes, metadata, previews and editable rendering`, async t => {
  const { library, root } = await fixture(t);
  const { snapshot, input } = await packageFixture({ legacy });
  const imported = await library.importPackage('session-a', input);
  assert.equal(imported.outcome, 'imported');
  assert.deepEqual(imported.template.templateProfile, snapshot.metadata.templateProfile);
  assert.deepEqual(imported.template.templateSamples, snapshot.metadata.templateSamples);
  assert.equal(imported.template.authoringKind, legacy ? undefined : 'native-pptd');
  assert.equal(imported.template.previewMetadata.source, 'template-package');
  const exported = decodeTemplatePackage(await library.exportPackage('session-b', imported.template.id));
  assert.deepEqual(exported.metadata.templateProfile, snapshot.metadata.templateProfile);
  assert.deepEqual(exported.metadata.templateSamples, snapshot.metadata.templateSamples);
  assert.equal(exported.metadata.sourceKind, snapshot.metadata.sourceKind);
  assert.equal(exported.files.size, snapshot.files.size);
  for (const [file, bytes] of snapshot.files) assert.deepEqual(Buffer.from(exported.files.get(file)), bytes, file);
  assert(!Object.hasOwn(exported.metadata, 'createdAt'));
  const output = path.join(root, 'copied'); await mkdir(output);
  await library.copyProject(imported.template.id, output);
  const rendered = await renderPptdProject(await loadPptdProject(output));
  assert.equal(rendered.check.status, 'pass');
  const pptx = unzipSync(rendered.bytes);
  assert.match(Buffer.from(pptx['ppt/slides/slide1.xml']).toString(), /Source title/);
  assert(Object.keys(pptx).some(file => file.startsWith('ppt/media/')));
});

test('repeat, concurrent, renamed and recompressed imports retain a single local template', async t => {
  const { library, root } = await fixture(t), { input } = await packageFixture();
  const [first, second] = await Promise.all([library.importPackage('a', input), library.importPackage('b', input)]);
  assert.deepEqual([first.outcome, second.outcome], ['imported', 'existing']);
  await library.update('a', first.template.id, { name: 'My local name', description: 'My local description' });
  const changed = repack(input, (entries, manifest) => { manifest.name = 'Foreign name'; manifest.sourceId = 'personal-forged';
    const reversed = Object.entries(entries).reverse(); for (const key of Object.keys(entries)) delete entries[key]; for (const [key, value] of reversed) entries[key] = value; });
  const again = await library.importPackage('a', changed);
  assert.equal(again.outcome, 'existing'); assert.equal(again.template.name, 'My local name');
  assert.equal(again.template.description, 'My local description');
  assert.deepEqual(await readdir(path.join(root, 'personal-templates/saved')), [first.template.id]);
});

test('supplied IDs and identical display names never overwrite different content', async t => {
  const { library } = await fixture(t), { input } = await packageFixture();
  const one = repack(input, (_entries, manifest) => { manifest.sourceId = 'same-id'; manifest.id = 'same-id'; });
  const two = repack(input, (entries, manifest) => { manifest.sourceId = 'same-id'; manifest.id = 'same-id'; entries['project/pages/body.page'] = Buffer.from(Buffer.from(entries['project/pages/body.page']).toString().replace('Source body', 'Different body')); });
  const a = await library.importPackage('a', one), b = await library.importPackage('a', two);
  assert.notEqual(a.template.id, b.template.id); assert.equal(a.template.name, b.template.name);
  assert.equal((await library.list()).length, 2);
});

test('legacy profile absence stays absent while display metadata is derived', async t => {
  const { library } = await fixture(t), { input } = await packageFixture({ legacy: true, absent: true });
  const result = await library.importPackage('a', input);
  assert.equal(result.template.templateProfile, undefined); assert.equal(result.template.templateSamples, undefined);
  assert(result.template.palette.background); assert(result.template.titleFontFace);
  const exported = decodeTemplatePackage(await library.exportPackage('a', result.template.id));
  assert.equal(exported.metadata.templateProfile, undefined); assert.equal(exported.metadata.templateSamples, undefined);
});

test('old local IDs deduplicate before capacity; broken unrelated templates do not block', async t => {
  const { library, root } = await fixture(t), { input } = await packageFixture();
  const first = await library.importPackage('a', input), saved = path.join(root, 'personal-templates/saved');
  const oldId = `personal-${'a'.repeat(64)}`, directory = path.join(saved, oldId);
  await cp(path.join(saved, first.template.id), directory, { recursive: true });
  await writeFile(path.join(directory, 'template.json'), JSON.stringify({ ...first.template, id: oldId }));
  await rm(path.join(saved, first.template.id), { recursive: true });
  for (let i = 0; i < 99; i++) {
    const id = `personal-${i.toString(16).padStart(64, '0')}`;
    await mkdir(path.join(saved, id));
    await writeFile(path.join(saved, id, 'template.json'), JSON.stringify({ ...first.template, id }));
  }
  assert.equal((await library.list()).length, 100);
  const duplicate = await library.importPackage('a', input);
  assert.equal(duplicate.outcome, 'existing'); assert.equal(duplicate.template.id, oldId);
  const different = (await packageFixture({ title: 'Different content' })).input;
  await assert.rejects(library.importPackage('a', different), /100/);
  await rm(path.join(saved, `personal-${'0'.repeat(64)}`), { recursive: true });
  assert.equal((await library.importPackage('a', different)).outcome, 'imported');
});

test('failure and cancellation clean staging; audit failure cannot undo committed imports', async t => {
  const { library, root } = await fixture(t), { input } = await packageFixture();
  const controller = new AbortController(), original = library.directory.bind(library);
  library.directory = async (...parts) => { const directory = await original(...parts); if (parts[0] === 'package-staging') controller.abort(); return directory; };
  await assert.rejects(library.importPackage('a', input, controller.signal), /abort/i);
  assert.deepEqual(await readdir(path.join(root, 'personal-templates/package-staging')), []);
  assert.deepEqual(await library.list(), []);
  library.directory = original;
  library.store.recordAudit = async () => { throw new Error('audit unavailable'); };
  const result = await library.importPackage('a', input);
  assert.equal(result.outcome, 'imported'); assert.equal((await library.importPackage('a', input)).outcome, 'existing');
  assert.deepEqual(await readdir(path.join(root, 'personal-templates/package-staging')), []);
});

test('existing unverified destination is an integrity conflict; built-in templates cannot export', async t => {
  const { library, root } = await fixture(t), { input, snapshot } = await packageFixture();
  const directory = path.join(root, 'personal-templates/saved', `personal-${snapshot.fingerprint}`);
  await mkdir(directory, { recursive: true }); await writeFile(path.join(directory, 'keep'), 'original');
  await assert.rejects(library.importPackage('a', input), /完整性冲突/);
  assert.equal(await readFile(path.join(directory, 'keep'), 'utf8'), 'original');
  await assert.rejects(library.exportPackage('a', 'dsh-blue-professional'), /标识无效/);
});

test('malformed archives, resources and metadata never publish a card', async t => {
  const { library } = await fixture(t), { input } = await packageFixture();
  const bad = [
    repack(input, (_entries, manifest) => { manifest.version = 2; }),
    repack(input, (_entries, manifest) => { manifest.templateProfile.version = 2; }),
    repack(input, entries => { entries['../escape'] = Buffer.from('bad'); }),
    repack(input, entries => { entries['project/assets/LOGO.png'] = entries['project/assets/logo.png']; }),
    repack(input, entries => { delete entries['project/assets/logo.png']; }),
    repack(input, entries => { delete entries['preview/pages/page-2.png']; }),
    repack(input, entries => { entries['project/pages/cover.page'] = Buffer.from(Buffer.from(entries['project/pages/cover.page']).toString().replace('assets/logo.png', '../outside.png')); }),
    repack(input, entries => { entries['preview/pages/page-2.png'] = entries['preview/pages/page-2.png'].subarray(0, 100); }),
    repack(input, (_entries, manifest) => { manifest.templateProfile.pages[0].file = 'pages/missing.page'; }),
    repack(input, (_entries, manifest) => { manifest.templateProfile.owner = '/private/owner'; }),
    { fileName: 'bad.dshtpl', base64: Buffer.from('not a zip').toString('base64') },
  ];
  const symlink = Buffer.from(input.base64, 'base64');
  const central = symlink.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02])); symlink.writeUInt32LE(0xa1ff0000, central + 38);
  bad.push({ ...input, base64: symlink.toString('base64') });
  const bomb = Buffer.from(input.base64, 'base64'); bomb.writeUInt32LE(0xffffffff, central + 24); bad.push({ ...input, base64: bomb.toString('base64') });
  for (const value of bad) await assert.rejects(library.importPackage('a', value));
  assert.deepEqual(await library.list(), []);
});

test('large valid base64 is checked without a regex stack overflow', () => {
  const bytes = randomBytes(8 * 1024 * 1024);
  assert.throws(() => decodeTemplatePackage({ fileName: 'large.dshtpl', base64: bytes.toString('base64') }), error => !(error instanceof RangeError));
});


test('export rejects symlinked preview ancestors in a corrupt saved record', async t => {
  const { library, root } = await fixture(t), { input } = await packageFixture();
  const imported = await library.importPackage('a', input);
  const directory = path.join(root, 'personal-templates/saved', imported.template.id);
  const outside = path.join(root, 'external-preview');
  await cp(path.join(directory, 'preview'), outside, { recursive: true });
  await rm(path.join(directory, 'preview'), { recursive: true });
  await symlink(outside, path.join(directory, 'preview'));
  await assert.rejects(library.exportPackage('a', imported.template.id), /本地模板目录无效/);
});


test('self-contained SVG text, CSS fragments and embedded raster images round trip and render', async t => {
  const { library, root } = await fixture(t), { input } = await packageFixture();
  const raster = await sharp({ create: { width: 24, height: 24, channels: 3, background: '#2463aa' } }).png().toBuffer();
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="100" height="100"><style>.brand {fill:url(#brand); font-family:Arial}</style><defs><linearGradient id="brand"><stop stop-color="red"/></linearGradient></defs><rect width="100" height="100" class="brand"/><image xlink:href="data:image/png;base64,${raster.toString('base64')}" width="24" height="24"/><text y="80" aria-label="sales@example.com">sales@example.com</text></svg>`;
  const shared = repack(input, entries => {
    delete entries['project/assets/logo.png'];
    entries['project/assets/logo.svg'] = Buffer.from(svg);
    for (const file of ['project/pages/cover.page', 'project/pages/body.page']) entries[file] = Buffer.from(Buffer.from(entries[file]).toString().replace('logo.png', 'logo.svg'));
  });
  const result = await library.importPackage('a', shared);
  const exported = decodeTemplatePackage(await library.exportPackage('a', result.template.id));
  assert.equal(exported.files.get('project/assets/logo.svg').toString(), svg);
  const destination = path.join(root, 'svg-project'); await mkdir(destination);
  await library.copyProject(result.template.id, destination);
  assert.equal((await renderPptdProject(await loadPptdProject(destination))).check.status, 'pass');
});

test('SVG external resources, CSS escapes, executable elements and invalid embedded images are rejected', async t => {
  const { library } = await fixture(t), { input } = await packageFixture();
  const fragments = [
    '<image href="https://example.com/image.png"/>',
    '<style>@import "https://example.com/style.css";</style>',
    String.raw`<style>rect{fill:u\72l("https://example.com/image.svg")}</style>`,
    '<rect style="fill:url(file:///private/image.svg)"/>',
    '<image href="data:image/svg+xml;base64,PHN2Zy8+"/>',
    '<image href="data:image/png;base64,aW52YWxpZA=="/>',
    '<script>alert(1)</script>',
    '<style><![CDATA[rect{fill:url(https://example.com/image.svg)}]]></style>',
  ];
  for (const fragment of fragments) {
    const shared = repack(input, entries => {
      delete entries['project/assets/logo.png'];
      entries['project/assets/logo.svg'] = Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">${fragment}</svg>`);
      for (const file of ['project/pages/cover.page', 'project/pages/body.page']) entries[file] = Buffer.from(Buffer.from(entries[file]).toString().replace('logo.png', 'logo.svg'));
    });
    await assert.rejects(library.importPackage('a', shared), undefined, fragment);
  }
  assert.deepEqual(await library.list(), []);
});


async function reviewDraft(library, snapshot, sessionId = 'editor') {
  const draftId = randomUUID(), directory = await library.directory('drafts', draftId);
  for (const [file, bytes] of snapshot.files) {
    const target = path.join(directory, file === 'original/source.pptx' ? 'source.pptx' : file);
    await mkdir(path.dirname(target), { recursive: true }); await writeFile(target, bytes);
  }
  const profile = snapshot.metadata.templateProfile;
  const record = { ...snapshot.metadata, id: `personal-${snapshot.fingerprint}`, sha256: snapshot.fingerprint,
    origin: 'personal', authoringKind: snapshot.metadata.sourceKind === 'native-pptd' ? 'native-pptd' : undefined,
    conversionVersion: 10, previewVersion: 11, createdAt: new Date().toISOString(), slideCount: snapshot.pageIndex.length, pageIndex: snapshot.pageIndex,
    palette: profile?.palette, titleFontFace: profile?.typography.title.fontFace, bodyFontFace: profile?.typography.body.fontFace,
    previewImages: Array(3).fill(`data:image/png;base64,${snapshot.files.get('preview/pages/page-1.png').toString('base64')}`) };
  await writeFile(path.join(directory, 'template.json'), JSON.stringify(record));
  await writeFile(path.join(directory, 'owner.json'), JSON.stringify({ session: hash(sessionId) }));
  return draftId;
}
const originalDestination = record => ({ kind: 'original', base: { templateId: record.id, version: record.version } });

test('content update keeps logical identity while source, design, preview and package readers follow its immutable revision', async t => {
  const { library } = await fixture(t), before = await packageFixture(), after = await packageFixture({ title: 'Updated source' });
  const png = await sharp({ create: { width: 960, height: 540, channels: 3, background: '#123456' } }).png().toBuffer();
  after.snapshot.files.set('preview/pages/page-1.png', png);
  const design = JSON.parse(after.snapshot.files.get('project/template-design.json')); design.purpose = 'Updated design';
  after.snapshot.files.set('project/template-design.json', Buffer.from(JSON.stringify(design)));
  const saved = (await library.importPackage('a', before.input)).template;
  const oldSource = await library.projectSource(saved.id);
  const draft = await reviewDraft(library, after.snapshot);
  const updated = await library.commit('editor', draft, { name: 'Updated original', destination: originalDestination(saved) });
  assert.equal(updated.template.id, saved.id); assert.notEqual(updated.base.version, saved.version);
  assert.equal(updated.template.createdAt, saved.createdAt); assert.equal((await library.list()).length, 1);
  assert.notEqual(await library.projectSource(saved.id), oldSource);
  const currentSnapshot = await library.readSnapshot(saved.id);
  assert.equal(currentSnapshot.record.version, updated.base.version);
  assert.equal(currentSnapshot.projectPath, await library.projectSource(saved.id));
  assert.equal(currentSnapshot.design.purpose, 'Updated design');
  assert.deepEqual(await readFile(path.join(oldSource, 'deck.pptd')), before.snapshot.files.get('project/deck.pptd'));
  assert.equal((await library.readDesign(saved.id)).purpose, 'Updated design');
  assert.equal((await library.previewSavedPage('a', saved.id, 1)).preview, `data:image/png;base64,${png.toString('base64')}`);
  const exported = decodeTemplatePackage(await library.exportPackage('a', saved.id));
  assert.deepEqual(exported.files.get('project/deck.pptd'), after.snapshot.files.get('project/deck.pptd'));
  assert.deepEqual(exported.files.get('project/template-design.json'), after.snapshot.files.get('project/template-design.json'));
  const importedOld = await library.importPackage('a', before.input);
  assert.equal(importedOld.outcome, 'imported'); assert.notEqual(importedOld.template.id, saved.id);
  assert.equal((await library.importPackage('a', await library.exportPackage('a', saved.id))).template.id, saved.id);
  const stage = await library.directory('test-checkout');
  await library.checkout(updated.base, stage);
  assert.deepEqual(await readFile(path.join(stage, 'deck.pptd')), after.snapshot.files.get('project/deck.pptd'));
  await assert.rejects(library.checkout({ templateId: saved.id, version: saved.version }, stage), /原模板已更新/);
});

test('explicit copies remain distinct even with identical bytes and updates still work at capacity', async t => {
  const { library, root } = await fixture(t), { input, snapshot } = await packageFixture();
  const saved = (await library.importPackage('a', input)).template;
  const copy = await library.commit('editor', await reviewDraft(library, snapshot), { name: saved.name, destination: { kind: 'copy' } });
  assert.notEqual(copy.template.id, saved.id); assert.equal((await library.list()).length, 2);
  const savedRoot = path.join(root, 'personal-templates/saved');
  for (let i = 0; i < 98; i++) {
    const id = `personal-${i.toString(16).padStart(64, '0')}`;
    await mkdir(path.join(savedRoot, id)); await writeFile(path.join(savedRoot, id, 'template.json'), JSON.stringify({ ...saved, id }));
  }
  assert.equal((await library.list()).length, 100);
  const draft = await reviewDraft(library, snapshot);
  await assert.rejects(library.commit('editor', draft, { name: 'Overflow copy', destination: { kind: 'copy' } }), /100/);
  assert(await library.draft('editor', draft));
  const updated = await library.commit('editor', draft, { name: 'Full-library update', destination: originalDestination(saved) });
  assert.equal(updated.template.id, saved.id); assert.equal((await library.list()).length, 100);
});

test('concurrent editors and metadata changes conflict; deleted originals can only be saved as a copy', async t => {
  const { library } = await fixture(t), { input, snapshot } = await packageFixture();
  const saved = (await library.importPackage('a', input)).template;
  const first = await reviewDraft(library, snapshot, 'first'), second = await reviewDraft(library, snapshot, 'second');
  const request = { name: 'Changed', destination: originalDestination(saved) };
  const results = await Promise.allSettled([library.commit('first', first, request), library.commit('second', second, request)]);
  assert.equal(results[0].status, 'fulfilled'); assert.equal(results[1].status, 'rejected'); assert.match(results[1].reason.message, /原模板已更新/);
  assert(await library.draft('second', second));
  const current = await library.readRecord(saved.id);
  await library.update('metadata-editor', saved.id, { name: 'Concurrent rename' });
  await assert.rejects(library.commit('second', second, { name: 'Stale', destination: originalDestination(current) }), /原模板已更新/);
  const latest = await library.readRecord(saved.id);
  await library.remove('a', saved.id);
  await assert.rejects(library.commit('second', second, { name: 'Deleted', destination: originalDestination(latest) }), /原模板已删除/);
  const copy = await library.commit('second', second, { name: 'Retained work', destination: { kind: 'copy' } });
  assert.notEqual(copy.template.id, saved.id); assert.equal((await library.list()).length, 1);
});

test('a failed manifest publication keeps the old source and preview visible and retry publishes the reserved revision', async t => {
  const { library, root } = await fixture(t), before = await packageFixture(), after = await packageFixture({ title: 'New revision' });
  const saved = (await library.importPackage('a', before.input)).template;
  const draft = await reviewDraft(library, after.snapshot), request = { name: 'Changed', destination: originalDestination(saved) };
  const writeJson = library.store.writeJson.bind(library.store);
  let fail = true;
  library.store.writeJson = async (file, value) => { if (fail && file === path.join(root, 'personal-templates/saved', saved.id, 'template.json')) { fail = false; throw new Error('manifest interrupted'); } return writeJson(file, value); };
  await assert.rejects(library.commit('editor', draft, request), /manifest interrupted/);
  assert.equal((await library.readRecord(saved.id)).version, saved.version);
  assert.deepEqual(await readFile(path.join(await library.projectSource(saved.id), 'deck.pptd')), before.snapshot.files.get('project/deck.pptd'));
  assert.equal((await library.previewPage('editor', draft, 1)).page, 1);
  const updated = await library.commit('editor', draft, request);
  assert.equal(updated.template.id, saved.id);
  assert.deepEqual(await library.commit('editor', draft, request), updated);
  assert.equal((await readdir(path.join(root, 'personal-templates/saved', saved.id, 'revisions'))).length, 1);
});

test('pending receipts reconcile before later metadata writes or deletion and exact retries return historical results', async t => {
  const { library } = await fixture(t), { snapshot } = await packageFixture();
  const draft = await reviewDraft(library, snapshot), request = { name: 'First committed name', destination: { kind: 'copy' } };
  const writeJson = library.store.writeJson.bind(library.store);
  let fail = true;
  library.store.writeJson = async (file, value) => { if (fail && file.includes('commit-receipts') && value.status === 'committed') { fail = false; throw new Error('receipt interrupted'); } return writeJson(file, value); };
  await assert.rejects(library.commit('editor', draft, request), /receipt interrupted/);
  const original = (await library.list())[0]; assert.equal(original.name, request.name);
  const restarted = new PersonalTemplateLibrary(library.store);
  await restarted.update('a', original.id, { name: 'Later name' });
  await restarted.remove('a', original.id);
  const recovered = await restarted.commit('editor', draft, request);
  assert.equal(recovered.template.name, request.name); assert.equal(recovered.template.id, original.id);
  assert.equal(recovered.base.version, original.version); assert.deepEqual(await restarted.list(), []);
  await assert.rejects(restarted.commit('editor', draft, { ...request, name: 'Changed retry' }), /其他名称或保存目标/);
  await assert.rejects(restarted.commit('editor', draft, { name: request.name, destination: originalDestination(original) }), /其他名称或保存目标/);
});

test('an unpublished original attempt can become a copy after the original changes', async t => {
  const { library, root } = await fixture(t), { input, snapshot } = await packageFixture();
  const saved = (await library.importPackage('a', input)).template;
  const draft = await reviewDraft(library, snapshot), request = { name: 'Unpublished original', destination: originalDestination(saved) };
  const writeJson = library.store.writeJson.bind(library.store);
  let fail = true;
  library.store.writeJson = async (file, value) => {
    if (fail && file === path.join(root, 'personal-templates/saved', saved.id, 'template.json')) { fail = false; throw new Error('manifest interrupted'); }
    return writeJson(file, value);
  };
  await assert.rejects(library.commit('editor', draft, request), /manifest interrupted/);
  await library.update('other', saved.id, { name: 'Another writer' });
  await assert.rejects(library.commit('editor', draft, request), /原模板已更新/);
  const copied = await library.commit('editor', draft, { name: 'Retained draft', destination: { kind: 'copy' } });
  assert.notEqual(copied.template.id, saved.id);
  assert.equal((await library.readRecord(saved.id)).name, 'Another writer');
  assert.equal((await library.list()).length, 2);
});

test('interruption immediately after pending receipt persistence reserves one copy identity across restart', async t => {
  const { library, root } = await fixture(t), { snapshot } = await packageFixture();
  const draft = await reviewDraft(library, snapshot), request = { name: 'Reserved copy', destination: { kind: 'copy' } };
  const writeJson = library.store.writeJson.bind(library.store); let fail = true;
  library.store.writeJson = async (file, value) => { await writeJson(file, value); if (fail && value.status === 'pending') { fail = false; throw new Error('process interrupted'); } };
  await assert.rejects(library.commit('editor', draft, request), /process interrupted/);
  assert.deepEqual(await library.list(), []);
  const receiptFile = (await readdir(path.join(root, 'personal-templates/commit-receipts')))[0];
  const pending = JSON.parse(await readFile(path.join(root, 'personal-templates/commit-receipts', receiptFile)));
  const restarted = new PersonalTemplateLibrary(library.store);
  const saved = await restarted.commit('editor', draft, request);
  assert.equal(saved.template.id, pending.template.id); assert.equal((await library.list()).length, 1);
});
