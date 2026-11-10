import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { access, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { unzipSync } from 'fflate';
import { runCli } from '../lib/bin.js';
import { definitions } from '../lib/catalog.js';
import { checkPptdProject, loadPptdProject, parsePptdProject, renderPptdProject } from '../lib/pptd.js';

const root = fileURLToPath(new URL('../', import.meta.url));
const executable = path.join(root, 'lib/bin.js');
const text = () => ({
  elementId: 'caption', elementType: 'text', bounds: [40, 40, 800, 80],
  content: { text: 'One shared contract 中文', fontSize: 24 },
});
const chart = () => ({
  elementId: 'sales', elementType: 'chart', bounds: [40, 150, 700, 320],
  data: { cols: ['month', 'sales'], rows: [['Jan', 3], ['Feb', 8]] },
  series: [{ type: 'bar', encode: { x: 'month', y: 'sales' }, fill: '#2563EB' }],
  xAxis: { label: { fontSize: 12 }, line: false },
  yAxis: [{ min: 0, max: 10, majorUnit: 2, label: { position: 'low' } }],
  legend: { show: false },
  plotLayout: { x: 0.17, y: 0.12, w: 0.7, h: 0.68 },
  plotBorder: { color: '#EF4444', width: 2 },
});

async function fixture(t, elements, extra = {}) {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'pptd-shared-core-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const source = {
    entryName: 'deck.pptd',
    manifest: JSON.stringify({ version: 'v2', title: 'Shared core', size: [960, 540], pages: ['pages/1.page'], ...extra }),
    pages: new Map([['pages/1.page', JSON.stringify({ elements })]]), assets: new Map(),
  };
  await mkdir(path.join(directory, 'pages'));
  await writeFile(path.join(directory, source.entryName), source.manifest);
  await writeFile(path.join(directory, 'pages/1.page'), source.pages.get('pages/1.page'));
  return { directory, source };
}

async function cli(args) {
  let stdout = '', stderr = '';
  const code = await runCli(args, {
    stdout: { write: chunk => { stdout += chunk; } },
    stderr: { write: chunk => { stderr += chunk; } },
  });
  assert.equal(stderr, '');
  return { code, result: JSON.parse(stdout) };
}

function spawnedCli(args) {
  const run = spawnSync(process.execPath, [executable, ...args], { encoding: 'utf8' });
  assert.equal(run.stderr, '');
  return { code: run.status, result: JSON.parse(run.stdout) };
}

test('API and CLI reject nested typos with the same file, page and element diagnostics', async t => {
  const cases = [
    ['content.fontSzie', text, element => { element.content.fontSzie = 24; }],
    ['content.style.fontSzie', text, element => { element.content.style = { fontSzie: 24 }; }],
    ['series[0].encode.yl', chart, element => { element.series[0].encode.yl = 'sales'; }],
    ['series[0].dataLabels.fontSzie', chart, element => { element.series[0].dataLabels = { fontSzie: 12 }; }],
    ['seriesDefaults.bar.widht', chart, element => { element.seriesDefaults = { bar: { widht: 2 } }; }],
    ['xAxis.label.fontSzie', chart, element => { element.xAxis.label.fontSzie = 12; }],
    ['yAxis[0].gridLine.widht', chart, element => { element.yAxis[0].gridLine = { widht: 2 }; }],
    ['plotBorder.widht', chart, element => { element.plotBorder.widht = 2; }],
    ['data.colz', chart, element => { element.data.colz = ['wrong']; }],
    ['series[0].fill.colour', chart, element => { element.series[0].fill = { type: 'solid', colour: '#2563EB' }; }],
  ];
  for (const [field, create, mutate] of cases) {
    await t.test(field, async t => {
      const element = create();
      mutate(element);
      const { directory, source } = await fixture(t, [element]);
      const checked = checkPptdProject(parsePptdProject(source));
      const result = await cli(['check', directory, '--json']);
      assert.equal(result.code, 1);
      assert.equal(checked.status, 'fail');
      assert.deepEqual(result.result.issues, checked.issues);
      assert.equal(result.result.digest, checked.digest);
      assert(checked.issues.some(issue => issue.code === 'unknown-field' && issue.message.includes(field)
        && issue.file === 'pages/1.page' && issue.page === 1 && issue.elementId === element.elementId));
    });
  }
});

test('nested style fields reject typos while named theme entries and bilingual font maps remain valid', async t => {
  const theme = {
    colors: { customAccent: '#2563EB' },
    textStyles: { customHeading: { fontSize: 24, color: '$customAccent', fontFamily: { latin: 'Arial', ea: 'MiSans', mac: 'PingFang SC', win: 'Microsoft YaHei' } } },
  };
  const element = text();
  element.content.style = '$customHeading';
  const { directory, source } = await fixture(t, [element], { theme });
  assert.equal(checkPptdProject(parsePptdProject(source)).errorCount, 0);
  assert.equal((await cli(['check', directory, '--json'])).code, 0);
  theme.textStyles.customHeading.fontSzie = 24;
  const checked = checkPptdProject(parsePptdProject({ ...source, manifest: JSON.stringify({ ...JSON.parse(source.manifest), theme }) }));
  assert(checked.issues.some(issue => issue.code === 'unknown-field' && issue.message.includes('theme.textStyles.customHeading.fontSzie')));
});

test('API and executable CLI both block rendering a nested typo before writing output', async t => {
  const element = text();
  element.content.fontSzie = 24;
  const { directory, source } = await fixture(t, [element]);
  await assert.rejects(renderPptdProject(parsePptdProject(source)), /unknown-field/);
  const output = path.join(directory, 'blocked.pptx');
  const rendered = spawnedCli(['render', directory, '-o', output, '--json']);
  assert.equal(rendered.code, 1);
  assert.equal(rendered.result.exported, false);
  assert.equal(rendered.result.status, 'needs_revision');
  await assert.rejects(access(output), { code: 'ENOENT' });
});

test('API and executable CLI preserve identical chart plot layout, border and axis XML', async t => {
  const { directory } = await fixture(t, [text(), chart()]);
  const project = await loadPptdProject(directory);
  const checked = checkPptdProject(project);
  assert.equal(checked.errorCount, 0);
  const api = await renderPptdProject(project);
  const output = path.join(directory, 'shared.pptx');
  const rendered = spawnedCli(['render', directory, '-o', output, '--json']);
  assert.equal(rendered.code, 0);
  assert.equal(rendered.result.digest, checked.digest);
  const apiZip = unzipSync(api.bytes), cliZip = unzipSync(await readFile(output));
  const chartFile = Object.keys(apiZip).find(file => /^ppt\/charts\/chart\d+\.xml$/.test(file));
  assert(chartFile);
  const xml = Buffer.from(apiZip[chartFile]).toString();
  assert.equal(Buffer.from(cliZip[chartFile]).toString(), xml);
  assert.match(xml, /<c:manualLayout>/);
  assert.match(xml, /<c:x val="0\.17"\s*\/>/);
  assert.match(xml, /<c:w val="0\.7"\s*\/>/);
  assert.match(xml, /<a:srgbClr val="EF4444"/);
  assert.match(xml, /<c:majorUnit val="2"\/>/);
  const preview = path.join(directory, 'quick-preview');
  assert.equal(spawnedCli(['screenshot', directory, '-o', preview, '--json']).code, 0);
  const png = await readFile(path.join(preview, 'pages/page-1.png'));
  assert.equal(png.subarray(1, 4).toString(), 'PNG');
});

test('CLI auto repair preserves template provenance and shared chart properties', async t => {
  const element = chart();
  element.bounds[0] = '40';
  const template = { id: 'personal-example', name: 'Example', sourceFile: 'original.pptx', sourceSha256: 'a'.repeat(64) };
  const { directory } = await fixture(t, [element], { template });
  const templateReference = { sourceSlideNumber: 2, rationale: 'Preserve the source data layout', contentRelationship: 'data' };
  const pagePath = path.join(directory, 'pages/1.page');
  await writeFile(pagePath, JSON.stringify({ templateReference, elements: [element] }));
  const repaired = await cli(['check', directory, '--level', 'auto', '--json']);
  assert.equal(repaired.code, 0);
  const project = await loadPptdProject(directory);
  assert.deepEqual(project.template, template);
  assert.deepEqual(project.pages[0].templateReference, templateReference);
  assert.deepEqual(project.pages[0].elements[0].plotLayout, element.plotLayout);
  assert.deepEqual(project.pages[0].elements[0].plotBorder, element.plotBorder);
});

test('all bundled English and Chinese templates pass the shared API and CLI contracts', async () => {
  let count = 0;
  for (const definition of definitions) {
    for (const source of ['source', 'source-zh']) {
      const directory = path.join(root, 'skills/dsh-ppt/references', definition.referenceDirectory, source);
      const checked = checkPptdProject(await loadPptdProject(directory));
      const result = await cli(['check', directory, '--json']);
      assert.equal(checked.errorCount, 0, `${definition.id}/${source}: ${JSON.stringify(checked.issues)}`);
      assert.equal(result.code, 0, `${definition.id}/${source}`);
      assert.deepEqual(result.result.issues, checked.issues);
      assert.equal(result.result.digest, checked.digest);
      count++;
    }
  }
  assert.equal(count, definitions.length * 2);
});
