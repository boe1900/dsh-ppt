import assert from 'node:assert/strict';
import test from 'node:test';
import { parsePptdProject, checkPptdProject } from '../lib/pptd.js';
import { extractPersonalTemplateProfile, validPersonalTemplateProfile } from '../lib/personal-template-profile.js';

const fonts = { latin: 'Arial', ea: 'Microsoft YaHei', mac: 'Microsoft YaHei', win: 'Microsoft YaHei' };
const project = content => parsePptdProject({
  entryName: 'deck.pptd',
  manifest: JSON.stringify({ version: 'v2', title: 'Bilingual template', size: [960, 540], pages: ['pages/cover.page'] }),
  pages: new Map([['pages/cover.page', JSON.stringify({ elements: [{
    elementId: 'title', elementType: 'text', bounds: [60, 50, 820, 90], content,
  }] })]]), assets: new Map(),
});

for (const [name, content, expected] of [
  ['Chinese font map', { text: '品牌模板', fontFamily: fonts, fontSize: 36 }, 'Microsoft YaHei'],
  ['Latin font map', { text: 'Brand template', fontFamily: fonts, fontSize: 36 }, 'Arial'],
  ['rich text font map', { text: '品牌模板', fontSize: 36, paragraphs: [{ runs: [{ text: '品牌模板', options: { fontFamily: fonts, fontSize: 36 } }] }] }, 'Microsoft YaHei'],
]) test(`a checked ${name} produces a valid, scalar template profile`, () => {
  const input = project(content);
  assert.equal(checkPptdProject(input).errorCount, 0);
  const profile = extractPersonalTemplateProfile(input);
  assert.equal(validPersonalTemplateProfile(profile, [{ file: 'pages/cover.page' }]), true, 'A valid bilingual PPTD must produce a readable template index');
  assert.equal(profile.typography.title.fontFace, expected);
  assert.equal(profile.pages[0].slots[0].fontFace, expected);
  assert.deepEqual(input.pages[0].elements[0].content, content, 'Profile extraction preserves authored font maps');
});
