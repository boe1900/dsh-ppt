import assert from 'node:assert/strict';
import { afterEach, beforeEach, mock, test } from 'node:test';
import { access, mkdir, mkdtemp, readFile, rm, truncate, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

let factory;
mock.module('@deepseek-ai/libreoffice-kit', { namedExports: {
  createConverterFactory: async () => factory,
} });
mock.module('../lib/pptd.js', { namedExports: {
  renderPptdProject: async () => ({ bytes: Buffer.from('editable presentation') }),
} });
mock.module('../lib/office-preview-batches.js', { namedExports: {
  createOfficePreviewBatches: () => (first, count) => Buffer.from(JSON.stringify({ first, count })),
} });
const { renderOfficeTemplatePreview } = await import('../lib/office-template-preview.js');

let directory, events, imageArguments, disposed, renderPdf, renderImages;
const project = count => ({ width: 720, height: 405, pages: Array.from({ length: count }, () => ({})) });
beforeEach(async () => {
  directory = await mkdtemp(path.join(os.tmpdir(), 'dsh-ppt-preview-sdk-'));
  events = [];
  imageArguments = [];
  disposed = false;
  renderPdf = async ({ inputPath, outputPath }, signal) => {
    signal?.throwIfAborted();
    const batch = JSON.parse(await readFile(inputPath, 'utf8'));
    events.push(`office:${batch.first + 1}:${batch.count}`);
    await writeFile(outputPath, JSON.stringify(batch));
    return { missingFonts: ['Missing Face'] };
  };
  renderImages = async (...args) => {
    imageArguments.push(args.length);
    const [{ inputPath, outputDir, pages }, signal] = args;
    signal?.throwIfAborted();
    const batch = JSON.parse(await readFile(inputPath, 'utf8'));
    await mkdir(outputDir);
    const images = [];
    for (const page of pages) {
      const file = path.join(outputDir, `page-${page}.png`);
      await writeFile(file, String(batch.first + page));
      images.push({ page, path: file });
    }
    events.push(`returned:${batch.first + 1}`);
    return { pageCount: batch.count, images };
  };
  let slot = 0;
  factory = {
    create: async () => slot++ === 0
      ? { render: (...args) => renderPdf(...args) }
      : { renderImages: (...args) => renderImages(...args) },
    dispose: async () => { disposed = true; },
  };
});
afterEach(async () => { await rm(directory, { recursive: true, force: true }); });

test('publishes every checked image with the unpatched two-argument SDK', async () => {
  const published = [];
  const result = await renderOfficeTemplatePreview(project(10), directory, 1920, 200, async page => {
    assert.equal(await readFile(page.path, 'utf8'), String(page.page));
    assert.equal(page.total, 10);
    published.push(page.page);
    events.push(`published:${page.page}`);
  });
  assert.deepEqual(published, Array.from({ length: 10 }, (_, index) => index + 1));
  assert.deepEqual(imageArguments, [2, 2, 2]);
  assert.ok(events.indexOf('returned:1') < events.indexOf('published:1'));
  assert.ok(events.indexOf('published:1') < events.indexOf('office:2:8'));
  assert.deepEqual(result.missingFonts, ['Missing Face']);
  assert.equal(await readFile(path.join(directory, 'preview-deck.pptx'), 'utf8'), 'editable presentation');
  assert.equal(disposed, true);
  await assert.rejects(access(path.join(directory, 'office-preview')));
});

test('rejects inconsistent image order before publishing a batch', async () => {
  const original = renderImages;
  renderImages = async (...args) => {
    const result = await original(...args);
    result.images[0].page = 2;
    return result;
  };
  const published = [];
  await assert.rejects(renderOfficeTemplatePreview(project(2), directory, 1920, 200,
    async page => published.push(page.page)), /page count mismatch/);
  assert.deepEqual(published, []);
  assert.equal(disposed, true);
  await assert.rejects(access(path.join(directory, 'office-preview')));
});

test('stops publication on cancellation and cleans in-flight batch files', async () => {
  const controller = new AbortController();
  const published = [];
  await assert.rejects(renderOfficeTemplatePreview(project(17), directory, 1920, 200, async page => {
    published.push(page.page);
    if (page.page === 2) controller.abort();
  }, controller.signal), { name: 'AbortError' });
  assert.deepEqual(published, [1, 2]);
  assert.equal(disposed, true);
  await assert.rejects(access(path.join(directory, 'office-preview')));
});

test('limits total PDF bytes across batches', async () => {
  const original = renderPdf;
  renderPdf = async (...args) => {
    const result = await original(...args);
    await truncate(args[0].outputPath, 64 * 1024 * 1024 + 1);
    return result;
  };
  renderImages = async ({ outputDir }) => {
    await mkdir(outputDir);
    const file = path.join(outputDir, 'page-1.png');
    await writeFile(file, '1');
    return { pageCount: 1, images: [{ page: 1, path: file }] };
  };
  const published = [];
  await assert.rejects(renderOfficeTemplatePreview(project(10), directory, 1920, 200,
    async page => published.push(page.page)), /PDF byte budget/);
  assert.deepEqual(published, [1]);
  assert.equal(disposed, true);
  await assert.rejects(access(path.join(directory, 'office-preview')));
});
